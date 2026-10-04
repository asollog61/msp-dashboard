"""
Property inspections for the MSP dashboard.

A phone-first checklist (one section per screen, Next/Back), with a comment and
photos on any item, a history of past inspections, and a punch list where the
fix for each flagged item is recorded.

Storage
  Google Sheets  'Inspections'       one row per inspection
                 'Inspection Items'  one row per checklist item per inspection
  Dropbox        photos, under <root>/<building>/<date>_<inspection id>/

Secrets (all optional; the tab works without them, minus the feature)
  [dropbox]
  app_key = "..."
  app_secret = "..."
  refresh_token = "..."
  root = "/MSP Inspections"

  inspector_password = "..."   # asked for on the ?view=inspect link
  dashboard_password = "..."   # asked for on the full dashboard
"""
import hmac
import json
from datetime import datetime
from io import BytesIO

import gspread
import requests
import streamlit as st
import streamlit.components.v1 as components

try:
    from zoneinfo import ZoneInfo
    _TZ = ZoneInfo("America/New_York")
except Exception:  # pragma: no cover
    _TZ = None

try:
    from PIL import Image, ImageOps
    HAS_PIL = True
except ImportError:  # pragma: no cover
    HAS_PIL = False


# --------------------------------------------------------------------------
# The checklist. Item numbers are stored with every inspection, so add new
# items at the end with a new number rather than renumbering existing ones.
# --------------------------------------------------------------------------
CHECKLIST = [
    ("Exterior and grounds", [
        (1, "Overall appearance: facade, signage, awnings"),
        (2, "Sidewalks, walkways and parking: cracks, holes, trip hazards"),
        (3, "Trash and dumpster area: tidy, no debris or odor"),
        (4, "Exterior lighting working"),
        (5, "Landscaping maintained (snow and ice in season)"),
        (6, "Exterior doors, locks and entry hardware working"),
    ]),
    ("Roof and envelope", [
        (7, "Roof clear of debris, no ponding"),
        (8, "Drains, gutters and downspouts clear"),
        (9, "Roof access door or hatch secure"),
        (10, "Windows and storefront glass intact, no leaks"),
    ]),
    ("Common areas", [
        (11, "Lobby and entry clean, no clutter or unauthorized signage"),
        (12, "Hallways and stairs clear of storage and obstructions"),
        (13, "Floors, walls and ceilings: damage, stains, broken tiles"),
        (14, "Common-area lighting working"),
        (15, "Restrooms clean, stocked, fixtures working"),
        (16, "Elevator clean, lit and operating (if applicable)"),
    ]),
    ("Life safety", [
        (17, "Exit signs lit"),
        (18, "Emergency lights working"),
        (19, "Fire extinguishers present and tagged"),
        (20, "Exit doors and panic hardware work; egress and fire escape clear"),
        (21, "Smoke and CO detectors and alarm panel normal"),
        (22, "Sprinkler and standpipe connections accessible"),
    ]),
    ("Building systems", [
        (23, "HVAC operating, no unusual noise or leaks"),
        (24, "Plumbing and water heater: no leaks"),
        (25, "Electrical: panels accessible, no exposed wiring"),
        (26, "No signs of water damage, mold or odors"),
        (27, "Basement and mechanical rooms clean, lit, no stored items"),
    ]),
    ("Tenant and lease compliance", [
        (28, "No unauthorized storage, hazardous materials or alterations"),
        (29, "Tenant upkeep acceptable; no pest evidence"),
    ]),
]

OK, NEEDS, NA = "OK", "Needs attention", "N/A"
RESULTS = [OK, NEEDS, NA]

WS_INSP = "Inspections"
WS_ITEMS = "Inspection Items"
INSP_HEADERS = ["Inspection ID", "Date", "Building", "Inspector", "OK",
                "Needs Attention", "N/A", "Notes", "Submitted At"]
ITEM_HEADERS = ["Inspection ID", "Date", "Building", "Item No", "Section", "Item",
                "Result", "Comment", "Photos", "Status", "Resolution",
                "Resolved By", "Resolved Date"]

PHOTO_MAX_PX = 1600
PHOTO_TYPES = ["jpg", "jpeg", "png", "webp"]


def _now():
    return datetime.now(_TZ) if _TZ else datetime.now()


# --------------------------------------------------------------------------
# Passwords
# --------------------------------------------------------------------------
def _secret(name):
    try:
        return str(st.secrets.get(name, "") or "")
    except Exception:
        return ""


def password_gate(kind):
    """Return True when this session may proceed.

    kind is 'inspector' (the inspections-only link) or 'dashboard' (everything).
    With no password configured for that kind the gate is open. The dashboard
    password also opens the inspector link.
    """
    own = _secret("inspector_password" if kind == "inspector" else "dashboard_password")
    if not own:
        return True
    if st.session_state.get("_auth_dashboard") or st.session_state.get(f"_auth_{kind}"):
        return True

    title = "Property Inspections" if kind == "inspector" else "MSP Property Dashboard"
    st.markdown(f"## 🏢 {title}")
    with st.form(f"_gate_{kind}"):
        entered = st.text_input("Password", type="password")
        go = st.form_submit_button("Enter", type="primary")
    if go:
        master = _secret("dashboard_password")
        if master and hmac.compare_digest(entered, master):
            st.session_state["_auth_dashboard"] = True
            st.rerun()
        elif hmac.compare_digest(entered, own):
            st.session_state[f"_auth_{kind}"] = True
            st.rerun()
        else:
            st.error("That password isn't right.")
    return False


# --------------------------------------------------------------------------
# Dropbox (photos)
# --------------------------------------------------------------------------
def _dbx_cfg():
    try:
        cfg = dict(st.secrets["dropbox"])
    except Exception:
        return None
    if all(cfg.get(k) for k in ("app_key", "app_secret", "refresh_token")):
        return cfg
    return None


@st.cache_data(ttl=3 * 3600, show_spinner=False)
def _dbx_token(app_key, app_secret, refresh_token):
    r = requests.post(
        "https://api.dropbox.com/oauth2/token",
        data={"grant_type": "refresh_token", "refresh_token": refresh_token},
        auth=(app_key, app_secret), timeout=20,
    )
    r.raise_for_status()
    return r.json()["access_token"]


def _dbx_headers(arg):
    cfg = _dbx_cfg()
    token = _dbx_token(cfg["app_key"], cfg["app_secret"], cfg["refresh_token"])
    return {"Authorization": f"Bearer {token}", "Dropbox-API-Arg": json.dumps(arg)}


def dbx_upload(path, data):
    """Upload bytes; returns the path Dropbox actually stored it at."""
    headers = _dbx_headers({"path": path, "mode": "add", "autorename": True, "mute": True})
    headers["Content-Type"] = "application/octet-stream"
    r = requests.post("https://content.dropboxapi.com/2/files/upload",
                      headers=headers, data=data, timeout=90)
    r.raise_for_status()
    return r.json().get("path_display", path)


@st.cache_data(ttl=3600, max_entries=200, show_spinner=False)
def dbx_download(path):
    r = requests.post("https://content.dropboxapi.com/2/files/download",
                      headers=_dbx_headers({"path": path}), timeout=60)
    r.raise_for_status()
    return r.content


def _shrink(raw):
    """Rotate per EXIF and cap the long edge so phone photos upload quickly."""
    if not HAS_PIL:
        return raw
    try:
        img = ImageOps.exif_transpose(Image.open(BytesIO(raw)))
        img = img.convert("RGB")
        img.thumbnail((PHOTO_MAX_PX, PHOTO_MAX_PX))
        out = BytesIO()
        img.save(out, format="JPEG", quality=82, optimize=True)
        return out.getvalue()
    except Exception:
        return raw


def _safe(s):
    return "".join(c if c.isalnum() or c in " -_" else "_" for c in str(s)).strip()


# --------------------------------------------------------------------------
# Google Sheets
# --------------------------------------------------------------------------
def _ws(sheet, name, headers):
    try:
        return sheet.worksheet(name)
    except gspread.exceptions.WorksheetNotFound:
        ws = sheet.add_worksheet(title=name, rows=200, cols=len(headers))
        ws.update(values=[headers], range_name="A1")
        return ws


@st.cache_data(ttl=60, show_spinner=False)
def _load(_sheet, name, headers):
    try:
        return _ws(_sheet, name, list(headers)).get_all_records()
    except Exception as exc:
        st.error(f"Couldn't read '{name}' from Google Sheets: {exc}")
        return []


def _load_all(sheet):
    insps = _load(sheet, WS_INSP, tuple(INSP_HEADERS))
    items = _load(sheet, WS_ITEMS, tuple(ITEM_HEADERS))
    return insps, items


def _save_inspection(sheet, meta, answers, photo_paths, notes):
    insp_id = meta["id"]
    counts = {r: 0 for r in RESULTS}
    rows = []
    for section, items in CHECKLIST:
        for no, text in items:
            a = answers.get(no, {})
            res = a.get("result") or NA
            counts[res] += 1
            rows.append([
                insp_id, meta["date"], meta["building"], no, section, text, res,
                a.get("comment", "") if res == NEEDS else "",
                "\n".join(photo_paths.get(no, [])),
                "Open" if res == NEEDS else "", "", "", "",
            ])
    _ws(sheet, WS_ITEMS, ITEM_HEADERS).append_rows(rows, value_input_option="RAW")
    _ws(sheet, WS_INSP, INSP_HEADERS).append_row([
        insp_id, meta["date"], meta["building"], meta["inspector"],
        counts[OK], counts[NEEDS], counts[NA], notes,
        _now().strftime("%Y-%m-%d %H:%M:%S"),
    ], value_input_option="RAW")
    _load.clear()


def _save_resolution(sheet, insp_id, item_no, status, resolution, by, when):
    ws = _ws(sheet, WS_ITEMS, ITEM_HEADERS)
    values = ws.get_all_values()
    for i, row in enumerate(values[1:], start=2):
        if len(row) >= 4 and row[0] == str(insp_id) and str(row[3]) == str(item_no):
            ws.update(values=[[status, resolution, by, when]], range_name=f"J{i}:M{i}")
            _load.clear()
            return True
    return False


# --------------------------------------------------------------------------
# UI helpers
# --------------------------------------------------------------------------
_CSS = """
<style>
.st-key-insp_root div[role="radiogroup"] { gap: 8px; flex-wrap: wrap; }
.st-key-insp_root div[role="radiogroup"] > label {
    border: 1px solid #2d333b; border-radius: 10px; padding: 10px 14px;
    background: #161b22; margin: 0; min-height: 44px;
}
.st-key-insp_root div[role="radiogroup"] > label:has(input:checked) {
    border-color: #58a6ff; background: #1a2332;
}
.st-key-insp_root .insp-item { font-weight: 600; margin: 4px 0 6px; line-height: 1.35; }
.st-key-insp_root .insp-chip {
    display: inline-block; padding: 2px 10px; border-radius: 10px;
    font-size: 12px; font-weight: 600;
}
.st-key-insp_root .insp-open { background: #f8514926; color: #f85149; }
.st-key-insp_root .insp-done { background: #23883826; color: #3fb950; }
.st-key-insp_root .stButton button { min-height: 44px; }
</style>
"""


def _scroll_top():
    """Streamlit keeps the scroll position across reruns; a new step should start at the top."""
    html = f"""<script>
        try {{
          const d = window.parent.document;
          const el = d.querySelector('[data-testid="stMain"]') || d.querySelector('section.main');
          if (el) el.scrollTo({{top: 0}});
          window.parent.scrollTo(0, 0);
        }} catch (e) {{}}
        </script><!-- {st.session_state.get('insp_step', 0)} -->"""
    try:
        if hasattr(st, "iframe"):
            st.iframe(html, height=1)
        else:
            components.html(html, height=0)
    except Exception:
        pass


def _reset_wizard():
    for k in list(st.session_state.keys()):
        if k.startswith("insp_") and k != "insp_view":
            del st.session_state[k]


def _init_wizard():
    ss = st.session_state
    ss.setdefault("insp_step", 0)
    ss.setdefault("insp_meta", {})
    ss.setdefault("insp_answers", {})
    ss.setdefault("insp_photos", {})
    ss.setdefault("insp_seen", {})
    ss.setdefault("insp_notes", "")


def _go(step):
    st.session_state.insp_step = step
    st.session_state.insp_scroll = True


def _render_item(no, text):
    ss = st.session_state
    ans = ss.insp_answers.setdefault(no, {"result": None, "comment": ""})
    st.markdown(f'<div class="insp-item">{no}. {text}</div>', unsafe_allow_html=True)

    idx = RESULTS.index(ans["result"]) if ans["result"] in RESULTS else None
    ans["result"] = st.radio("Result", RESULTS, index=idx, horizontal=True,
                             key=f"insp_r_{no}", label_visibility="collapsed")

    if ans["result"] == NEEDS:
        ans["comment"] = st.text_input(
            "What needs attention?", value=ans["comment"], key=f"insp_c_{no}",
            placeholder="Describe the issue and where it is",
        )

    photos = ss.insp_photos.setdefault(no, [])
    seen = ss.insp_seen.setdefault(no, set())
    label = f"📷 Photos ({len(photos)})" if photos else "📷 Add photo"
    with st.expander(label, expanded=False):
        files = st.file_uploader(
            "Take or choose a photo", type=PHOTO_TYPES, accept_multiple_files=True,
            key=f"insp_u_{no}", label_visibility="collapsed",
        )
        added = False
        for f in files or []:
            sig = (f.name, f.size)
            if sig in seen:
                continue
            seen.add(sig)
            photos.append({"name": f.name, "data": _shrink(f.getvalue())})
            added = True
        if added:
            st.rerun()
        for i, p in enumerate(photos):
            c1, c2 = st.columns([3, 1])
            c1.image(p["data"], width=160)
            if c2.button("Remove", key=f"insp_rm_{no}_{i}"):
                photos.pop(i)
                st.rerun()
        if not photos:
            st.caption("On a phone this opens the camera or your photo library.")


def _section_problems(items):
    ans = st.session_state.insp_answers
    missing = [no for no, _ in items if not ans.get(no, {}).get("result")]
    no_comment = [no for no, _ in items
                  if ans.get(no, {}).get("result") == NEEDS
                  and not (ans.get(no, {}).get("comment") or "").strip()]
    return missing, no_comment


# --------------------------------------------------------------------------
# New inspection (wizard)
# --------------------------------------------------------------------------
def _render_new(sheet, buildings):
    _init_wizard()
    ss = st.session_state
    n_sections = len(CHECKLIST)
    step = ss.insp_step

    if ss.pop("insp_scroll", False):
        _scroll_top()

    if ss.get("insp_done"):
        d = ss.insp_done
        st.success(f"Inspection saved: {d['building']}, {d['date']}.")
        if d["flagged"]:
            st.write(f"{d['flagged']} item(s) need attention. They're now on the Open issues list.")
        else:
            st.write("No items flagged.")
        if d.get("photo_errors"):
            st.warning(f"{d['photo_errors']} photo(s) could not be saved.")
        if st.button("Start another inspection", type="primary"):
            _reset_wizard()
            st.rerun()
        return

    # ---- Step 0: who / where / when
    if step == 0:
        meta = ss.insp_meta
        st.markdown("#### New inspection")
        b_idx = buildings.index(meta["building"]) if meta.get("building") in buildings else None
        building = st.selectbox("Building", buildings, index=b_idx, placeholder="Choose a building")
        inspector = st.text_input("Your name", value=meta.get("inspector", ss.get("inspector_name", "")))
        when = st.date_input("Inspection date", value=meta.get("date_obj", _now().date()))
        if not _dbx_cfg():
            st.info("Photo storage isn't connected yet, so photos can't be saved. "
                    "The checklist and comments still work.")
        if st.button("Start", type="primary", use_container_width=True):
            if not building or not inspector.strip():
                st.error("Choose a building and enter your name.")
            else:
                ss.inspector_name = inspector.strip()
                ss.insp_meta = {
                    "building": building, "inspector": inspector.strip(),
                    "date_obj": when, "date": when.strftime("%Y-%m-%d"),
                }
                _go(1)
                st.rerun()
        return

    meta = ss.insp_meta
    st.caption(f"{meta['building']} · {meta['date']} · {meta['inspector']}")

    # ---- Steps 1..n: one section per screen
    if 1 <= step <= n_sections:
        section, items = CHECKLIST[step - 1]
        st.progress(step / (n_sections + 1), text=f"Section {step} of {n_sections}")
        st.markdown(f"#### {section}")
        for no, text in items:
            _render_item(no, text)
            st.divider()

        back, nxt = st.columns(2)
        if back.button("← Back", use_container_width=True, key=f"insp_back_{step}"):
            _go(step - 1)
            st.rerun()
        nxt_label = "Review →" if step == n_sections else "Next →"
        if nxt.button(nxt_label, type="primary", use_container_width=True, key=f"insp_next_{step}"):
            missing, no_comment = _section_problems(items)
            if missing:
                st.error("Answer every item before continuing. Missing: "
                         + ", ".join(f"#{n}" for n in missing))
            elif no_comment:
                st.error("Add a comment for: " + ", ".join(f"#{n}" for n in no_comment))
            else:
                _go(step + 1)
                st.rerun()
        return

    # ---- Final step: review and submit
    st.progress(1.0, text="Review")
    st.markdown("#### Review and submit")
    ans = ss.insp_answers
    flagged = [(sec, no, text) for sec, items in CHECKLIST for no, text in items
               if ans.get(no, {}).get("result") == NEEDS]
    total = sum(len(items) for _, items in CHECKLIST)
    n_ok = sum(1 for a in ans.values() if a.get("result") == OK)
    n_na = sum(1 for a in ans.values() if a.get("result") == NA)
    c1, c2, c3 = st.columns(3)
    c1.metric("OK", n_ok)
    c2.metric("Needs attention", len(flagged))
    c3.metric("N/A", n_na)

    if flagged:
        st.markdown("**Items needing attention**")
        for sec, no, text in flagged:
            n_ph = len(ss.insp_photos.get(no, []))
            ph = f" · 📷 {n_ph}" if n_ph else ""
            st.markdown(f"- **{no}. {text}**{ph}  \n  {ans[no]['comment']}")
    else:
        st.write("Nothing flagged.")

    ss.insp_notes = st.text_area("General notes (optional)", value=ss.insp_notes)

    back, sub = st.columns(2)
    if back.button("← Back", use_container_width=True, key="insp_back_review"):
        _go(n_sections)
        st.rerun()
    if sub.button("Submit inspection", type="primary", use_container_width=True):
        if n_ok + n_na + len(flagged) < total:
            st.error("Some items are unanswered. Go back and complete them.")
            return
        if not sheet:
            st.error("Google Sheets is unavailable, so the inspection can't be saved right now. "
                     "Your answers are still here; try Submit again in a minute.")
            return
        code = "".join(ch for ch in meta["building"] if ch.isalnum())[:12]
        meta["id"] = f"{code}-{_now().strftime('%Y%m%d-%H%M%S')}"

        photo_paths, errors = {}, 0
        all_photos = [(no, p) for no, lst in ss.insp_photos.items() for p in lst]
        if all_photos:
            cfg = _dbx_cfg()
            if not cfg:
                errors = len(all_photos)
            else:
                root = (cfg.get("root") or "/MSP Inspections").rstrip("/")
                folder = f"{root}/{_safe(meta['building'])}/{meta['date']}_{meta['id']}"
                bar = st.progress(0.0, text="Uploading photos…")
                counter = {}
                for i, (no, p) in enumerate(all_photos):
                    counter[no] = counter.get(no, 0) + 1
                    try:
                        path = dbx_upload(f"{folder}/item{no:02d}_{counter[no]}.jpg", p["data"])
                        photo_paths.setdefault(no, []).append(path)
                    except Exception:
                        errors += 1
                    bar.progress((i + 1) / len(all_photos), text="Uploading photos…")
                bar.empty()
        try:
            with st.spinner("Saving…"):
                _save_inspection(sheet, meta, ans, photo_paths, ss.insp_notes.strip())
        except Exception as exc:
            st.error(f"Couldn't save the inspection: {exc}. Your answers are still here; try again.")
            return
        ss.insp_done = {"building": meta["building"], "date": meta["date"],
                        "flagged": len(flagged), "photo_errors": errors}
        st.rerun()


# --------------------------------------------------------------------------
# Open issues and history
# --------------------------------------------------------------------------
def _photo_list(item):
    return [p for p in str(item.get("Photos") or "").split("\n") if p.strip()]


def _show_photos(item, key):
    paths = _photo_list(item)
    if not paths:
        return
    if not st.toggle(f"📷 Show {len(paths)} photo(s)", key=f"insp_ph_{key}"):
        return
    if not _dbx_cfg():
        st.caption("Photo storage isn't connected.")
        return
    cols = st.columns(min(len(paths), 3))
    for i, path in enumerate(paths):
        try:
            cols[i % len(cols)].image(dbx_download(path), use_container_width=True)
        except Exception:
            cols[i % len(cols)].caption(f"Couldn't load {path}")


def _building_filter(buildings, key):
    return st.selectbox("Building", ["All buildings"] + buildings, key=key)


def _inspector_of(insps):
    return {str(i.get("Inspection ID")): i.get("Inspector", "") for i in insps}


def _render_open(sheet, buildings):
    insps, items = _load_all(sheet)
    pick = _building_filter(buildings, "insp_open_bldg")
    who = _inspector_of(insps)
    open_items = [it for it in items if it.get("Status") == "Open"
                  and (pick == "All buildings" or it.get("Building") == pick)]
    open_items.sort(key=lambda it: (str(it.get("Date")), str(it.get("Building")), int(it.get("Item No") or 0)))

    if not open_items:
        st.success("No open issues.")
        return
    st.caption(f"{len(open_items)} open issue(s), oldest first")

    for it in open_items:
        iid, no = str(it["Inspection ID"]), it["Item No"]
        key = f"{iid}_{no}"
        with st.container(border=True):
            st.markdown(f"**{it['Building']}** · {it['Date']} · {who.get(iid, '')}")
            st.markdown(f"**{no}. {it['Item']}**")
            st.write(it.get("Comment") or "")
            _show_photos(it, key)
            with st.form(f"insp_res_{key}"):
                text = st.text_area("What was done", value=str(it.get("Resolution") or ""),
                                    placeholder="Work performed, vendor, cost, work order…")
                c1, c2 = st.columns(2)
                by = c1.text_input("By", value=str(it.get("Resolved By") or "")
                                   or st.session_state.get("inspector_name", ""))
                when = c2.date_input("Date", value=_now().date())
                b1, b2 = st.columns(2)
                save = b1.form_submit_button("Save update", use_container_width=True)
                done = b2.form_submit_button("Mark resolved", type="primary", use_container_width=True)
            if save or done:
                if done and not text.strip():
                    st.error("Say what was done before marking it resolved.")
                else:
                    try:
                        ok = _save_resolution(sheet, iid, no, "Resolved" if done else "Open",
                                              text.strip(), by.strip(),
                                              when.strftime("%Y-%m-%d") if done else "")
                    except Exception as exc:
                        ok = False
                        st.error(f"Couldn't save: {exc}")
                    if ok:
                        st.rerun()


def _render_history(sheet, buildings):
    insps, items = _load_all(sheet)
    pick = _building_filter(buildings, "insp_hist_bldg")
    rows = [i for i in insps if pick == "All buildings" or i.get("Building") == pick]
    rows.sort(key=lambda i: str(i.get("Submitted At")), reverse=True)
    if not rows:
        st.info("No inspections yet.")
        return

    by_insp = {}
    for it in items:
        by_insp.setdefault(str(it.get("Inspection ID")), []).append(it)

    for insp in rows:
        iid = str(insp["Inspection ID"])
        its = by_insp.get(iid, [])
        flagged = [it for it in its if it.get("Result") == NEEDS]
        n_open = sum(1 for it in flagged if it.get("Status") == "Open")
        if not flagged:
            tail = "no issues"
        elif n_open:
            tail = f"{len(flagged)} flagged, {n_open} open"
        else:
            tail = f"{len(flagged)} flagged, all resolved"
        with st.expander(f"{insp['Date']} · {insp['Building']} · {insp['Inspector']} · {tail}"):
            if insp.get("Notes"):
                st.markdown(f"**Notes:** {insp['Notes']}")
            if not flagged:
                st.write("Nothing was flagged.")
            for it in flagged:
                resolved = it.get("Status") == "Resolved"
                chip = ('<span class="insp-chip insp-done">Resolved</span>' if resolved
                        else '<span class="insp-chip insp-open">Open</span>')
                st.markdown(f"**{it['Item No']}. {it['Item']}** {chip}", unsafe_allow_html=True)
                st.write(it.get("Comment") or "")
                if it.get("Resolution"):
                    meta = " · ".join(str(x) for x in (it.get("Resolved By"), it.get("Resolved Date")) if x)
                    st.markdown(f"↳ **What was done:** {it['Resolution']}" + (f"  \n  _{meta}_" if meta else ""))
                _show_photos(it, f"h_{iid}_{it['Item No']}")
                st.divider()
            if st.toggle("Show all items", key=f"insp_all_{iid}"):
                for it in sorted(its, key=lambda x: int(x.get("Item No") or 0)):
                    st.markdown(f"{it['Item No']}. {it['Item']} — **{it['Result']}**")


# --------------------------------------------------------------------------
# Entry point
# --------------------------------------------------------------------------
def render_inspections_tab(get_gsheet, buildings):
    """get_gsheet: callable returning the dashboard spreadsheet (or None)."""
    sheet = get_gsheet()
    st.markdown(_CSS, unsafe_allow_html=True)
    with st.container(key="insp_root"):
        view = st.radio("View", ["New inspection", "Open issues", "History"],
                        horizontal=True, key="insp_view", label_visibility="collapsed")
        if view == "New inspection":
            _render_new(sheet, list(buildings))
            return
        if not sheet:
            st.warning("Google Sheets is unavailable, so past inspections can't be shown right now.")
            return
        if view == "Open issues":
            _render_open(sheet, list(buildings))
        else:
            _render_history(sheet, list(buildings))
