"""Render the master as a document on the left and its decisions on the right.

Two different questions were being answered by one view, badly:

    "what does this lease say?"      needs the text as it will print, with
                                     nothing else in the way
    "what could it say instead?"     needs every alternative, including the
                                     ones not chosen

So there are two renderers. `render_document` shows only what is in — fixed
text, the selected option from each set, the optional blocks switched on. It
reads like the lease because it is the lease. `render_options` shows one
section's decisions: every competing version, every block that can be dropped.

The index links carry a `?lb_section=` query parameter rather than a plain
anchor, because Streamlit reruns on a query change and a bare `#anchor` only
scrolls. That is what makes clicking a heading drive the right-hand pane.

Produces HTML. No Streamlit import, so it can be tested by reading the string.
"""

from __future__ import annotations

import html
import re
from typing import Any

import lease_blocks as lbk

STYLE = """
<style>
/* The document pane is a sheet of paper: white, serif, justified, with real
   margins. Anything else and it stops reading like a lease. */
.lv-page { background: #fdfdfb; color: #17171a; padding: 2.6rem 3rem;
           border-radius: 3px; box-shadow: 0 1px 14px rgba(0,0,0,.45);
           font-family: 'Times New Roman', Georgia, serif; font-size: 13.5px;
           line-height: 1.62; }
.lv-page p { margin: 0 0 .62rem 0; text-align: justify; }
.lv-page .lv-h { font-weight: 700; text-align: left;
                 margin: 1.5rem 0 .55rem 0; }
.lv-page .lv-h:first-child { margin-top: 0; }
.lv-lead { font-weight: 700; }
.lv-num { display: inline-block; min-width: 1.9rem; font-weight: 700; }
.lv-l1 { margin-left: 2.2rem; }
.lv-l2 { margin-left: 4.4rem; }
.lv-token { color: #c0392b; font-weight: 600; }
.lv-hit { background: #fff6c9; }

/* Key Provisions keeps its table, because that is what it is in Word. */
.lv-kps { width: 100%; border-collapse: collapse; margin: .3rem 0 1rem 0;
          font-size: 12.5px; }
.lv-kps td { border: 1px solid #b9b9b3; padding: .34rem .5rem;
             vertical-align: top; }
.lv-kps td:first-child { font-weight: 700; width: 31%; background: #f2f2ee; }
.lv-kps .lv-alt { display: block; padding-top: .25rem; color: #6b6b66;
                  font-size: 11.5px; font-style: italic; }

/* Left rail: navigation only. */
.lv-nav { font-family: system-ui, sans-serif; font-size: 12.5px; line-height: 1.45; }
.lv-nav a { text-decoration: none; display: block; padding: .3rem .5rem;
            border-radius: .35rem; opacity: .82; }
.lv-nav a:hover { background: rgba(128,128,128,.16); opacity: 1; }
.lv-nav a.on { background: rgba(88,166,255,.20); opacity: 1; font-weight: 600; }
.lv-nav .lv-sub { padding-left: 1rem; }
.lv-badge { font-size: 10.5px; opacity: .8; }
.lv-badge.set { color: #2f9bb5; }
.lv-badge.opt { color: #2e9e5b; }
.lv-nav-h { font-family: system-ui, sans-serif; font-size: 10.5px;
            letter-spacing: .07em; text-transform: uppercase; opacity: .55;
            margin: .9rem 0 .35rem .5rem; }

/* Options panel, shown under the document. */
.lv-opts { font-family: system-ui, sans-serif; font-size: 12.5px; }
.lv-card { border: 1px solid rgba(128,128,128,.32); border-radius: .5rem;
           padding: .6rem .75rem; margin-bottom: .7rem; }
.lv-card.set { border-left: 3px solid #2f9bb5; }
.lv-card.opt { border-left: 3px solid #2e9e5b; }
.lv-tag { font-size: 10.5px; letter-spacing: .05em; text-transform: uppercase;
          opacity: .72; display: block; margin-bottom: .4rem; }
.lv-choice { border-top: 1px solid rgba(128,128,128,.2); padding: .4rem 0 .1rem;
             font-family: 'Times New Roman', Georgia, serif; font-size: 13px;
             line-height: 1.45; }
.lv-choice:first-of-type { border-top: none; }
.lv-name { font-family: system-ui, sans-serif; font-size: 11px; font-weight: 600;
           opacity: .85; display: block; margin-bottom: .15rem; }
.lv-in  { color: #2e9e5b; }
.lv-out { opacity: .5; }
</style>
"""

# Word restarts numbering per level: 1. then a. then i.
_ROMAN = ["i", "ii", "iii", "iv", "v", "vi", "vii", "viii", "ix", "x", "xi", "xii"]


def _marker(level: int, index: int) -> str:
    if level <= 0:
        return f"{index}."
    if level == 1:
        return f"{chr(96 + index) if index <= 26 else index}."
    return f"{_ROMAN[index - 1] if index <= len(_ROMAN) else index}."


LEAD_IN_RE = re.compile(r"^([^.:\n]{3,70}?)([.:]\s)(.*)$", re.DOTALL)
TOKEN_RE = re.compile(r"\[\s*(?:Space|KP)\s*:[^\]]*\]", re.IGNORECASE)


def _escape(text: str) -> str:
    """Escaped, with unresolved tokens made obvious rather than invisible."""
    escaped = html.escape(str(text or ""))
    return TOKEN_RE.sub(lambda m: f'<span class="lv-token">{m.group(0)}</span>', escaped)


def _para(text: str, depth: int = 0) -> str:
    body = _escape(text)
    match = LEAD_IN_RE.match(body)
    if match:
        body = (f'<span class="lv-lead">{match.group(1)}{match.group(2).rstrip()}</span> '
                f'{match.group(3)}')
    klass = f' class="lv-i{min(depth, 2)}"' if depth else ""
    return f"<p{klass}>{body}</p>"


def _heading(container: dict[str, Any]) -> str:
    if container["kind"] == "section":
        return f"Section {container['label']}. {container.get('title','')}".strip().rstrip(".")
    if container.get("title"):
        return f"{container['label']} — {container['title']}"
    return str(container["label"])


def _anchor(container: dict[str, Any]) -> str:
    label = str(container.get("label", "")).replace(" ", "-").replace(".", "_")
    return f"lv-{container.get('kind','x')}-{label}"


# ---------------------------------------------------------------------------
# Decisions
# ---------------------------------------------------------------------------

def groups_in(container: dict[str, Any]) -> dict[int, list[dict]]:
    groups: dict[int, list[dict]] = {}
    for block in container.get("blocks", []):
        if block.get("choice_group"):
            groups.setdefault(block["choice_group"], []).append(block)
    return groups


def optional_in(container: dict[str, Any]) -> list[dict]:
    return [b for b in container.get("blocks", [])
            if not b.get("choice_group") and b.get("field") is None]


def provisions_in(container: dict[str, Any]) -> list[dict]:
    return [b for b in container.get("blocks", []) if b.get("field") is not None]


def decisions_in(container: dict[str, Any]) -> int:
    return len(groups_in(container)) + len(optional_in(container)) + sum(
        1 for p in provisions_in(container)
        if p.get("alternatives") or p.get("optional")
    )


def describe_decisions(container: dict[str, Any]) -> str:
    """"2 options · 3 optional" — what this section is asking of you."""
    parts = []
    sets_ = groups_in(container)
    if sets_:
        total = sum(len(members) for members in sets_.values())
        parts.append(f'<span class="lv-badge set">{total} options</span>')
    optional = optional_in(container)
    if optional:
        parts.append(f'<span class="lv-badge opt">{len(optional)} optional</span>')
    provisions = [p for p in provisions_in(container)
                  if p.get("alternatives") or p.get("optional")]
    if provisions:
        parts.append(f'<span class="lv-badge set">{len(provisions)} provisions</span>')
    return " · ".join(parts)


# ---------------------------------------------------------------------------
# Selection
# ---------------------------------------------------------------------------

def default_selection(master: dict[str, Any]) -> dict[str, Any]:
    """What the lease says before you change anything.

    A pick-one set takes its first option, because a lease has to say
    something. A green block stays out, because that is what green means.
    """
    selection: dict[str, Any] = {}
    for container in master["containers"]:
        label = str(container["label"])
        for group, members in groups_in(container).items():
            selection[f"{label}|set{group}"] = members[0]["name"]
        for block in optional_in(container):
            selection[f"{label}|opt|{block['name']}"] = False
    return selection


def _chosen(selection: dict, label: str, group: int, members: list[dict]) -> dict:
    name = (selection or {}).get(f"{label}|set{group}")
    for member in members:
        if member["name"] == name:
            return member
    return members[0]


def _is_on(selection: dict, label: str, block: dict) -> bool:
    return bool((selection or {}).get(f"{label}|opt|{block['name']}", False))


# ---------------------------------------------------------------------------
# Left pane: the lease as it reads
# ---------------------------------------------------------------------------

def _kps_table(provisions: list[dict], selection: dict, label: str) -> str:
    """The Key Provisions Summary, as the bordered table it is in Word."""
    rows = []
    for provision in provisions:
        value = provision.get("text", "")
        alternatives = provision.get("alternatives") or []
        if alternatives:
            picked = selection.get(f"{label}|kp|{provision['field']}")
            value = next((a for a in alternatives if a == picked), alternatives[0])
        note = ""
        if len(alternatives) > 1:
            note = (f'<span class="lv-alt">{len(alternatives)} versions available'
                    f'</span>')
        elif provision.get("optional"):
            note = '<span class="lv-alt">optional provision</span>'
        rows.append(
            f'<tr><td>{html.escape(provision.get("field", ""))}</td>'
            f'<td>{_escape(value).replace(chr(10), "<br>")}{note}</td></tr>'
        )
    return f'<table class="lv-kps">{"".join(rows)}</table>' if rows else ""


def render_document(master: dict[str, Any], selection: dict | None = None,
                    only: str | None = None) -> str:
    """The lease as it will print — every paragraph, in order, decisions applied.

    Walks the container's full paragraph list rather than its blocks, because
    the blocks are only the parts with choices in them. Reading the blocks
    alone produced a document missing most of its own text.

    A cyan paragraph is emitted only if it belongs to the option in use; a
    green paragraph only if that block is switched on. Everything else is
    fixed text and always prints.
    """
    selection = selection if selection is not None else default_selection(master)
    parts = ['<div class="lv-page">']

    for container in master["containers"]:
        label = str(container["label"])
        if only is not None and label != only:
            continue
        body = container.get("body", [])
        provisions = provisions_in(container)
        if not body and not provisions:
            continue

        parts.append(f'<p class="lv-h" id="{_anchor(container)}">'
                     f'{html.escape(_heading(container))}</p>')
        if provisions:
            parts.append(_kps_table(provisions, selection, label))

        # Which cyan option is in use, so its paragraphs can be kept and the
        # rest of the run dropped.
        groups = groups_in(container)
        keep, drop = set(), set()
        for group, members in groups.items():
            winner = _chosen(selection, label, group, members)
            for member in members:
                target = keep if member["name"] == winner["name"] else drop
                target.add(member["text"])
                target.update(member.get("children", []))
        for block in optional_in(container):
            target = keep if _is_on(selection, label, block) else drop
            target.add(block["text"])
            target.update(block.get("children", []))

        counters: dict[int, int] = {}
        for entry in body:
            text = entry.get("text", "")
            if entry.get("separator") or not text:
                continue
            if entry.get("colour") in lbk.SCAFFOLDING:
                continue
            if text in drop and text not in keep:
                continue

            level = int(entry.get("indent", 0) or 0)
            counters[level] = counters.get(level, 0) + 1
            for deeper in [k for k in counters if k > level]:
                counters.pop(deeper)
            marker = _marker(level, counters[level]) if level else ""

            body_html = _escape(text)
            match = LEAD_IN_RE.match(body_html)
            if match:
                body_html = (f'<span class="lv-lead">{match.group(1)}'
                             f'{match.group(2).rstrip()}</span> {match.group(3)}')
            klass = f" lv-l{min(level, 2)}" if level else ""
            prefix = f'<span class="lv-num">{marker}</span>' if marker else ""
            parts.append(f'<p class="{klass.strip()}">{prefix}{body_html}</p>')

    parts.append("</div>")
    return "".join(parts)


# ---------------------------------------------------------------------------
# Right pane: the index, and one section's decisions
# ---------------------------------------------------------------------------

def render_nav(master: dict[str, Any], active: str | None = None) -> str:
    """The left rail: every section, badged where it asks something of you."""
    rows = ['<div class="lv-nav">', '<div class="lv-nav-h">Sections</div>']
    for container in master["containers"]:
        if not container.get("body") and not provisions_in(container):
            continue
        label = str(container["label"])
        badge = describe_decisions(container) if decisions_in(container) else ""
        on = " on" if active == label else ""
        rows.append(
            f'<a class="{on.strip()}" href="?lb_section={html.escape(label)}">'
            f'{html.escape(_heading(container))}'
            + (f' {badge}' if badge else "")
            + "</a>"
        )
    rows.append("</div>")
    return "".join(rows)


def render_index(master: dict[str, Any], active: str | None = None) -> str:
    """Sections that ask something of you, each saying what and how many."""
    rows = []
    for container in master["containers"]:
        if not decisions_in(container):
            continue
        label = str(container["label"])
        on = " on" if active == label else ""
        rows.append(
            f'<div><a class="lv-link{on}" href="?lb_section={html.escape(label)}">'
            f'{html.escape(_heading(container))}</a> '
            f'{describe_decisions(container)}</div>'
        )
    if not rows:
        return '<div class="lv-index">No options marked in this master.</div>'
    return f'<div class="lv-index">{"".join(rows)}</div>'


def render_options(container: dict[str, Any], selection: dict | None = None) -> str:
    """Everything this section could say instead of what it says now."""
    label = str(container["label"])
    selection = selection or {}
    parts = ['<div class="lv-opts">']

    for provision in provisions_in(container):
        if not (provision.get("alternatives") or provision.get("optional")):
            continue
        parts.append('<div class="lv-card set">')
        head = html.escape(provision.get("field", ""))
        if provision.get("alternatives"):
            parts.append(f'<span class="lv-tag">{head} · choose one of '
                         f'{len(provision["alternatives"])}</span>')
            for alternative in provision["alternatives"]:
                parts.append(f'<div class="lv-choice">{_escape(alternative)}</div>')
        else:
            parts.append(f'<span class="lv-tag">{head} · optional</span>')
            parts.append(f'<div class="lv-choice">{_escape(provision.get("text",""))}</div>')
        parts.append("</div>")

    groups = groups_in(container)
    for group, members in sorted(groups.items()):
        winner = _chosen(selection, label, group, members)
        parts.append('<div class="lv-card set">')
        parts.append(f'<span class="lv-tag">choose one of {len(members)}</span>')
        for member in members:
            mark = ('<span class="lv-in">● in the lease</span>'
                    if member["name"] == winner["name"] else
                    '<span class="lv-out">○ not used</span>')
            parts.append(
                f'<div class="lv-choice"><span class="lv-name">'
                f'{html.escape(member["name"])} — {mark}</span>{_escape(member["text"])}'
                + "".join(f'<div class="lv-i1">{_escape(c)}</div>'
                          for c in member.get("children", []))
                + "</div>"
            )
        parts.append("</div>")

    for block in optional_in(container):
        on = _is_on(selection, label, block)
        parts.append('<div class="lv-card opt">')
        parts.append(
            f'<span class="lv-tag">optional · '
            + ('<span class="lv-in">included</span>' if on
               else '<span class="lv-out">not included</span>')
            + "</span>"
        )
        parts.append(
            f'<div class="lv-choice"><span class="lv-name">'
            f'{html.escape(block["name"])}</span>{_escape(block["text"])}'
            + "".join(f'<div class="lv-i1">{_escape(c)}</div>'
                      for c in block.get("children", []))
            + "</div>"
        )
        parts.append("</div>")

    if len(parts) == 1:
        parts.append('<div class="lv-card">Nothing to decide in this section.</div>')
    parts.append("</div>")
    return "".join(parts)


def find_container(master: dict[str, Any], label: str | None):
    if label is None:
        return None
    for container in master["containers"]:
        if str(container["label"]) == str(label):
            return container
    return None


def first_with_decisions(master: dict[str, Any]):
    for container in master["containers"]:
        if decisions_in(container):
            return container
    return None
