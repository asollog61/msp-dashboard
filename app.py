"""
MSP Property Dashboard — Streamlit App
Shared multi-user dashboard with Google Sheets backend.
"""
import streamlit as st
import json
import os
import re
from datetime import datetime, date, timedelta
from pathlib import Path

import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
import openpyxl
import json

try:
    import fitz
    HAS_PYMUPDF = True
except ImportError:
    HAS_PYMUPDF = False

from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode, JsCode

# --- CONFIG ---
SHEET_ID = "1Jqdnf9JFPLBoYirN7ZpBiicDo13meTcEdaI19usbZss"
SERVICE_ACCOUNT_FILE = None
SERVICE_ACCOUNT_INFO = None
COLUMN_CONFIG_FILE = os.path.join(os.path.dirname(__file__), "column_widths.json")
EXPENSE_CONFIG_FILE = os.path.join(os.path.dirname(__file__), "building_expenses.json")

# Try file first, then Streamlit secrets
for p in [
    os.path.join(os.path.dirname(__file__), "service_account.json"),
    "/home/node/.openclaw/workspace/config/gcp_service_account.json",
]:
    if os.path.exists(p):
        SERVICE_ACCOUNT_FILE = p
        break

if not SERVICE_ACCOUNT_FILE:
    try:
        SERVICE_ACCOUNT_INFO = dict(st.secrets["gcp_service_account"])
    except Exception:
        pass

TENANCY_FILE = None
for p in [
    os.path.join(os.path.dirname(__file__), "data", "MSP Tenancy.xlsx"),
    "/home/node/OpenClaw/Tenant Database/MSP Tenancy.xlsx",
    r"C:\Dropbox\OpenClaw\Tenant Database\MSP Tenancy.xlsx",
]:
    if os.path.exists(p):
        TENANCY_FILE = p
        break

SOP_PDF = None
for p in [
    os.path.join(os.path.dirname(__file__), "data", "Marion_St_SOP_Manual.pdf"),
    "/home/node/OpenClaw/Share Jason/General Procedures/Marion_St_SOP_Manual.pdf",
]:
    if os.path.exists(p):
        SOP_PDF = p
        break

BUILDING_MAP = {
    "114 Central": {"code": "MSP114", "dest_folder": "114 Central Westfield", "share": "0_114Share"},
    "15 South": {"code": "MSP15", "dest_folder": "15 South Street", "share": "0_15Share"},
    "36 South": {"code": "MSP36", "dest_folder": "36 South Street", "share": "0_36Share"},
    "1280 Springfield": {"code": "MSP1280", "dest_folder": "1280-86 Springfield Ave", "share": "0_1280Share"},
}

ACTIVE_PROPS_ROOT = None
for _p in [
    "/home/node/OpenClaw/Share Jason/ACTIVE PROPERTIES",
    r"C:\Dropbox\ASRA Investments\Marion St Properties\ACTIVE PROPERTIES",
]:
    if os.path.exists(_p):
        ACTIVE_PROPS_ROOT = Path(_p)
        break

TODAY = date.today()

# --- PAGE CONFIG ---
st.set_page_config(
    page_title="MSP Property Dashboard",
    page_icon="🏢",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# Custom CSS
st.markdown("""
<style>
    .stTabs [data-baseweb="tab-list"] { gap: 8px; }
    .stTabs [data-baseweb="tab"] { padding: 8px 20px; font-weight: 600; }
    .metric-card { background: #161b22; border: 1px solid #2d333b; border-radius: 8px; padding: 16px; text-align: center; }
    .metric-value { font-size: 28px; font-weight: 700; color: #58a6ff; font-family: monospace; }
    .metric-label { font-size: 12px; color: #8b949e; margin-top: 4px; }
    .building-header { background: #1a2332; padding: 8px 12px; border-radius: 6px; margin: 16px 0 8px; border-left: 3px solid #58a6ff; }
    .badge-mtm { background: #f8514926; color: #f85149; padding: 2px 8px; border-radius: 10px; font-size: 12px; font-weight: 600; }
    .badge-short { background: #d2992244; color: #d29922; padding: 2px 8px; border-radius: 10px; font-size: 12px; font-weight: 600; }
    .badge-ok { background: #23883826; color: #3fb950; padding: 2px 8px; border-radius: 10px; font-size: 12px; font-weight: 600; }
    .badge-long { background: #1f6feb26; color: #58a6ff; padding: 2px 8px; border-radius: 10px; font-size: 12px; font-weight: 600; }
    .sop-card { background: #161b22; border: 1px solid #2d333b; border-radius: 8px; padding: 20px; margin-bottom: 12px; cursor: pointer; transition: all 0.2s; }
    .sop-card:hover { border-color: #58a6ff; }
    div[data-testid="stDataFrame"] table { font-size: 13px !important; }
    .ag-theme-streamlit .ag-header-cell-label { font-size: 12px !important; }
    .ag-theme-streamlit .ag-cell { font-size: 12px !important; padding: 2px 6px !important; }
    .ag-theme-streamlit .ag-row { height: 28px !important; }
    .ag-theme-streamlit .ag-header-row { height: 32px !important; }
    .sop-section { border: 1px solid rgba(255,255,255,0.1); border-radius: 6px; padding: 8px 12px; background: rgba(255,255,255,0.02); margin-bottom: 10px; }
    .sop-heading { font-weight: 600; margin-top: 8px; text-transform: uppercase; letter-spacing: 0.5px; font-size: 0.85rem; color: #c8d4ff; }
    .sop-text { margin: 2px 0 4px; line-height: 1.2; }
    .sop-steps { margin: 4px 0 8px 18px; padding-left: 12px; line-height: 1.25; }
    .sop-steps li { margin-bottom: 2px; }
    .sop-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); gap: 8px; margin: 6px 0 10px; }
    .sop-grid-card { border: 1px solid rgba(255,255,255,0.08); border-radius: 6px; padding: 6px 10px; background: rgba(255,255,255,0.03); }
    .sop-grid-title { font-size: 0.8rem; text-transform: uppercase; letter-spacing: 0.4px; color: #9fb7ff; margin-bottom: 4px; }
    .sop-grid-card ul { margin: 0; padding-left: 16px; line-height: 1.2; }
</style>
""", unsafe_allow_html=True)


def ensure_json_file(path, default):
    if not os.path.exists(path):
        try:
            with open(path, 'w', encoding='utf-8') as f:
                json.dump(default, f)
        except Exception:
            return default
    try:
        with open(path, 'r', encoding='utf-8') as f:
            data = json.load(f)
            return data if isinstance(data, dict) else default
    except Exception:
        return default


def read_column_config():
    return ensure_json_file(COLUMN_CONFIG_FILE, {})


def write_column_config(data):
    try:
        with open(COLUMN_CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=2)
    except Exception as e:
        st.error(f"Unable to save column config: {e}")


def read_expense_config():
    default = {name: 0 for name in BUILDING_MAP.keys()}
    return ensure_json_file(EXPENSE_CONFIG_FILE, default)


def write_expense_config(data):
    try:
        with open(EXPENSE_CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=2)
    except Exception as e:
        st.error(f"Unable to save expense config: {e}")


def get_column_width_overrides(tab_key):
    if not tab_key:
        return {}
    data = read_column_config()
    cfg = data.get(tab_key, {})
    clean = {}
    for col, width in cfg.items():
        try:
            clean[col] = int(width)
        except (ValueError, TypeError):
            continue
    return clean


def save_column_width_overrides(tab_key, overrides):
    data = read_column_config()
    data[tab_key] = overrides
    write_column_config(data)


def render_column_config_editor(tab_key, columns):
    if not tab_key or not columns:
        return
    existing = read_column_config().get(tab_key, {})
    default_lines = []
    for col in columns:
        width = existing.get(col, '')
        default_lines.append(f"{col},{width}")
    with st.expander(f"⚙️ Column Width Config — {tab_key.title()}"):
        st.write("Enter `ColumnName,width` per line. Leave width blank to use automatic sizing.")
        text = st.text_area(
            "column_widths", value='\n'.join(default_lines), height=150,
            key=f"cfg_text_{tab_key}", label_visibility="collapsed"
        )
        if st.button("Save Column Widths", key=f"cfg_save_{tab_key}"):
            new_cfg = {}
            for line in text.splitlines():
                parts = [p.strip() for p in line.split(',') if p.strip()]
                if len(parts) >= 2:
                    try:
                        new_cfg[parts[0]] = int(parts[1])
                    except ValueError:
                        continue
            save_column_width_overrides(tab_key, new_cfg)
            st.success("Saved. Reloading with new widths...")
            st.rerun()


def render_expense_editor():
    cfg = read_expense_config()
    with st.expander("💸 Building Expense Config"):
        st.caption("Enter annual operating expenses per building. Used for CAM reimbursement and NOI calcs.")
        updated = {}
        for building in BUILDING_MAP.keys():
            val = cfg.get(building, 0) or 0
            updated[building] = st.number_input(
                f"{building} Annual Expenses", min_value=0.0, value=float(val), step=1000.0,
                key=f"expense_{building}"
            )
        if st.button("Save Expenses", key="save_expenses"):
            write_expense_config({k: float(v) for k, v in updated.items()})
            st.success("Building expenses saved. Reload the page to apply.")
            st.rerun()


def show_grid(df, key, height=None, fit_columns=True, pinned_bottom=None, column_configs=None, tab_key=None):
    """Render a DataFrame as an AG Grid with compact columns and dark theme."""
    import pandas as pd
    gb = GridOptionsBuilder.from_dataframe(df)
    gb.configure_default_column(
        sortable=True,
        filterable=True,
        resizable=True,
        suppressSizeToFit=False,
        autoWidth=True,
        wrapHeaderText=True,
        autoHeaderHeight=True,
    )

    # Set narrow column widths based on content type
    for col in df.columns:
        sample = df[col].astype(str)
        max_len = max(sample.str.len().max(), len(col))
        if max_len <= 4:
            width = 60
        elif max_len <= 8:
            width = 80
        elif max_len <= 15:
            width = 110
        elif max_len <= 25:
            width = 150
        else:
            width = 200
        gb.configure_column(col, width=width, minWidth=50)

    # Make Tenant column wider
    if 'Tenant' in df.columns:
        gb.configure_column('Tenant', width=160, minWidth=120)
    if 'Building' in df.columns:
        gb.configure_column('Building', width=130, minWidth=100)

    if column_configs:
        for col, cfg in column_configs.items():
            if col in df.columns:
                gb.configure_column(col, **cfg)

    overrides = get_column_width_overrides(tab_key)
    for col, width in overrides.items():
        if col in df.columns:
            gb.configure_column(col, width=width, minWidth=width)

    grid_options = gb.build()

    # Add pinned bottom row for totals
    if pinned_bottom is not None:
        grid_options['pinnedBottomRowData'] = pinned_bottom

    # Calculate height
    if height is None:
        row_count = len(df) + (1 if pinned_bottom else 0)
        height = min(max(row_count * 30 + 40, 100), 600)

    try:
        AgGrid(
            df,
            gridOptions=grid_options,
            height=height,
            theme='streamlit',
            update_mode=GridUpdateMode.NO_UPDATE,
            fit_columns_on_grid_load=fit_columns,
            key=key,
            allow_unsafe_jscode=True,
        )
    except Exception:
        # Fallback to native Streamlit dataframe if AG Grid fails
        st.dataframe(df, height=height, use_container_width=True)


# --- GOOGLE SHEETS CONNECTION ---
@st.cache_resource(ttl=300)
def get_gsheet():
    scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
    if SERVICE_ACCOUNT_FILE:
        creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=scopes)
    elif SERVICE_ACCOUNT_INFO:
        creds = Credentials.from_service_account_info(SERVICE_ACCOUNT_INFO, scopes=scopes)
    else:
        return None
    gc = gspread.authorize(creds)
    return gc.open_by_key(SHEET_ID)


def get_activity_data():
    sheet = get_gsheet()
    if not sheet:
        return []
    try:
        ws = sheet.worksheet('Vacancy Activity')
        records = ws.get_all_records()
        return records
    except Exception as e:
        st.error(f"Error reading activity data: {e}")
        return []


def add_activity_entry(entry):
    sheet = get_gsheet()
    if not sheet:
        return False
    try:
        ws = sheet.worksheet('Vacancy Activity')
        ws.append_row([
            entry['date'], entry['space'], entry['prospect'], entry['broker'],
            entry['type'], entry['feedback'], entry.get('added_by', ''),
            datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        ], value_input_option='USER_ENTERED')
        return True
    except Exception as e:
        st.error(f"Error adding entry: {e}")
        return False


def delete_activity_entry(row_idx):
    """Delete a row from Vacancy Activity (row_idx is 0-based from records, +2 for header+1-index)."""
    sheet = get_gsheet()
    if not sheet:
        return False
    try:
        ws = sheet.worksheet('Vacancy Activity')
        ws.delete_rows(row_idx + 2)
        return True
    except Exception as e:
        st.error(f"Error deleting: {e}")
        return False


def get_marketing_data():
    sheet = get_gsheet()
    if not sheet:
        return []
    try:
        ws = sheet.worksheet('Marketing')
        return ws.get_all_records()
    except Exception:
        return []


def add_marketing_entry(entry):
    sheet = get_gsheet()
    if not sheet:
        return False
    try:
        ws = sheet.worksheet('Marketing')
        ws.append_row([
            entry['space'], entry['description'], entry['url'],
            entry.get('added_by', ''), datetime.now().strftime('%Y-%m-%d'),
            'Active'
        ], value_input_option='USER_ENTERED')
        return True
    except Exception as e:
        st.error(f"Error: {e}")
        return False


def delete_marketing_entry(row_idx):
    sheet = get_gsheet()
    if not sheet:
        return False
    try:
        ws = sheet.worksheet('Marketing')
        ws.delete_rows(row_idx + 2)
        return True
    except Exception:
        return False


# --- TENANCY DATA ---
@st.cache_data(ttl=3600)
def load_tenancy():
    if not TENANCY_FILE:
        return [], {}, []
    wb = openpyxl.load_workbook(TENANCY_FILE, data_only=True)
    ws = wb['Sheet1']

    # Summary rows (10-40)
    summary_header = [ws.cell(10, c).value for c in range(1, 25)]
    summaries = []
    for r in range(11, 41):
        row = {}
        for c, h in enumerate(summary_header, 1):
            if h:
                row[h] = ws.cell(r, c).value
        if row.get('Buidling'):
            summaries.append(row)

    # Term detail rows (45+)
    detail_header = [ws.cell(45, c).value for c in range(1, 23)]
    details = {}
    for r in range(46, ws.max_row + 1):
        row = {}
        for c, h in enumerate(detail_header, 1):
            if h:
                row[h] = ws.cell(r, c).value
        tid = row.get('ID')
        if not tid or tid == '-':
            continue
        if tid not in details:
            details[tid] = []
        details[tid].append(row)

    wb.close()

    # Build tenant list
    tenants = []
    def to_date(val):
        if isinstance(val, datetime):
            return val.date()
        if isinstance(val, date):
            return val
        return None

    for tid, rows in details.items():
        first = rows[0]
        rows_sorted = sorted(rows, key=lambda r: to_date(r.get('Start Date')) or date.min)
        current_row = rows_sorted[0]
        next_row = None
        for r in rows_sorted:
            start = to_date(r.get('Start Date'))
            end = to_date(r.get('End Date'))
            if start and end and start <= TODAY <= end:
                current_row = r
            elif start and start > TODAY:
                if not next_row or start < to_date(next_row.get('Start Date')):
                    next_row = r

        last_orig_end = None
        for r in rows:
            if r.get('Term') == 'Original' and r.get('End Date'):
                ed = r['End Date']
                if isinstance(ed, (datetime, date)):
                    if last_orig_end is None or ed > last_orig_end:
                        last_orig_end = ed

        tte_days = None
        tte_months = 0
        tte_label = 'MTM'
        if last_orig_end:
            if isinstance(last_orig_end, datetime):
                last_orig_end = last_orig_end.date()
            tte_days = (last_orig_end - TODAY).days
            if tte_days < 0:
                tte_months = 0
                tte_label = 'MTM'
            else:
                tte_months = max(0, round(tte_days / 30.44))
                tte_label = str(tte_months)
        else:
            tte_days = -1
            tte_months = 0
            tte_label = 'MTM'

        option_years = len([r for r in rows if r.get('Term', '').startswith('Option')])
        monthly = current_row.get('Monthly', 0) or 0
        annual = current_row.get('Annual', 0) or 0
        sqft = first.get('Sqft', 0) or 0
        psf = (annual / sqft) if sqft and sqft > 1 else 0
        sec_dep = first.get('Sec Dep', 0) or 0

        exp_dt = first.get('Exp Dt')
        exp_str = exp_dt.strftime('%Y-%m-%d') if isinstance(exp_dt, (datetime, date)) else str(exp_dt or 'N/A')

        next_anniv = to_date(next_row.get('Start Date')) if next_row else None
        next_monthly = (next_row.get('Monthly') if next_row else None)
        delta_monthly = (next_monthly - monthly) if (next_row and next_monthly is not None) else None
        anniv_months = None
        if next_anniv:
            diff_days = (next_anniv - TODAY).days
            anniv_months = max(0, round(diff_days / 30.44))

        # Get tenant type (Retail/Office/etc) from summary rows
        tenant_type = ''
        tenant_name = first.get('Tenant', '')
        for s in summaries:
            if s.get('Tenant Name') == tenant_name and s.get('Buidling') == first.get('Building'):
                tenant_type = s.get('Type', '') or ''
                break

        tenants.append({
            'Building': first.get('Building', ''),
            'Space': first.get('Space', ''),
            'Tenant': first.get('Tenant', ''),
            'Type': tenant_type,
            'Floor': first.get('Floor', ''),
            'SF': sqft,
            'Lease': first.get('Type', ''),
            'Monthly': round(monthly, 2),
            'Annual': round(annual, 2),
            'PSF': round(psf, 2),
            'CAM_Pct': first.get('CAM', 0) or 0,
            'Is_NNN': True if str(first.get('Type', '')).upper().startswith('NNN') else False,
            'TTE': '0' if tte_label == 'MTM' else tte_label,
            'TTE_Label': tte_label,
            'TTE_Months': tte_months,
            'TTE_Days': tte_days if tte_days is not None else 9999,
            'Options': f"{option_years}yr" if option_years > 0 else '-',
            'Exp Date': exp_str,
            'Escalation': current_row.get('Escalation', 0) or 0,
            'Sec Dep': round(sec_dep, 2),
            'Next Anniv': next_anniv,
            'Anniv_Months': anniv_months,
            'Next Monthly': round(next_monthly, 2) if next_monthly is not None else None,
            'Delta Monthly': round(delta_monthly, 2) if delta_monthly is not None else None,
        })

    tenants.sort(key=lambda t: (t['Building'], t['TTE_Days']))
    return tenants, details, summaries


# --- COI SCANNING ---
def extract_cert_info(pdf_path):
    if not HAS_PYMUPDF:
        return None, None
    try:
        doc = fitz.open(pdf_path)
        page = doc[0]
        full_text = page.get_text("text")
        all_dates = []
        for m in re.finditer(r'(\d{1,2}/\d{1,2}/\d{4})', full_text):
            try:
                dt = datetime.strptime(m.group(1), '%m/%d/%Y')
                all_dates.append(dt)
            except ValueError:
                pass
        exp_date = max(all_dates) if all_dates else None

        insured_name = None
        blocks = page.get_text("blocks")
        text_blocks = sorted([b for b in blocks if b[6] == 0], key=lambda b: (b[1], b[0]))
        insured_label_y = None
        for b in text_blocks:
            first_line = b[4].strip().split('\n')[0].strip().upper()
            if b[0] < page.rect.width * 0.5 and first_line in ('INSURED', 'NAMED INSURED'):
                insured_label_y = b[1]
                break
        if insured_label_y is not None:
            for b in text_blocks:
                text = b[4].strip().split('\n')[0].strip()
                if b[1] <= insured_label_y or b[0] > page.rect.width * 0.4:
                    continue
                if b[1] > insured_label_y + 50:
                    break
                if len(text) < 3 or re.match(r'^\d+\s', text):
                    continue
                if re.match(r'^[A-Za-z\s]+,?\s*[A-Z]{2}\s+\d{5}', text):
                    continue
                if any(kw in text.upper() for kw in ('INSURER', 'CERTIFICATE', 'REVISION')):
                    continue
                insured_name = text
                break
        doc.close()
        return exp_date, insured_name
    except Exception:
        return None, None


def fuzzy_match_tenant(tenant_name, cert_name):
    if not tenant_name or not cert_name:
        return False
    t, c = tenant_name.upper().strip(), cert_name.upper().strip()
    if t in c or c in t:
        return True
    skip = {'LLC', 'INC', 'CORP', 'LTD', 'DBA', 'THE', 'OF', 'AND', '&', 'PDF', 'COI', 'CERTIFICATE', 'INSURANCE'}
    t_words = [w for w in re.split(r'[\s_\-\.]+', t) if w not in skip and len(w) > 2]
    c_words = [w for w in re.split(r'[\s_\-\.]+', c) if w not in skip and len(w) > 2]
    if t_words and c_words and t_words[0] == c_words[0]:
        return True
    return len(set(t_words) & set(c_words)) >= 1


@st.cache_data(ttl=3600)
def scan_coi_files():
    coi_data = {}
    if not ACTIVE_PROPS_ROOT:
        return coi_data
    for building_name, mapping in BUILDING_MAP.items():
        cert_dir = ACTIVE_PROPS_ROOT / mapping["dest_folder"] / mapping["share"] / "Certificates of Insurance"
        if not cert_dir.exists():
            parent = ACTIVE_PROPS_ROOT / mapping["dest_folder"] / mapping["share"]
            if parent.exists():
                for d in parent.iterdir():
                    if d.is_dir() and 'certificate' in d.name.lower() and 'insurance' in d.name.lower():
                        cert_dir = d
                        break
        certs = []
        if cert_dir.exists():
            for f in sorted(cert_dir.iterdir()):
                if not f.is_file() or f.suffix.lower() != '.pdf':
                    continue
                exp_date, insured_name = extract_cert_info(str(f))
                if not insured_name:
                    m = re.match(r'Exp_(\d{8})_MSP\d+_(.+)\.pdf', f.name, re.IGNORECASE)
                    if m:
                        if not exp_date:
                            try:
                                exp_date = datetime.strptime(m.group(1), '%Y%m%d')
                            except ValueError:
                                pass
                        insured_name = m.group(2).replace('_', ' ')
                certs.append({'filename': f.name, 'exp_date': exp_date, 'insured_name': insured_name})
        coi_data[building_name] = certs
    return coi_data


# --- SOP EXTRACTION ---
@st.cache_data(ttl=86400)
def _extract_sop_content(pdf_path, cache_token):
    if not pdf_path or not HAS_PYMUPDF:
        return []
    doc = fitz.open(pdf_path)
    full_text = ""
    for page in doc:
        full_text += page.get_text() + "\n"
    doc.close()

    section_defs = [
        ("Roles & Definitions", "👥"),
        ("SOP 100 — New Lease or Addendum", "📝"), ("SOP 200 — Insurance Reconciliation", "🛡️"),
        ("SOP 300 — Security Deposit Reconciliation", "💰"), ("SOP 400 — Tenant Move-Out", "🚪"),
        ("SOP 500 — Recurring Reports", "📊"), ("SOP 600 — Vacancy & Lead Management", "🏠"),
        ("Appendix — Numbering, Revision Log & Open Items", "📎"),
    ]

    positions = []
    for title, icon in section_defs:
        idx = full_text.find(title)
        if idx >= 0:
            last_idx = idx
            search_from = idx + 1
            while True:
                next_idx = full_text.find(title, search_from)
                if next_idx < 0:
                    break
                last_idx = next_idx
                search_from = next_idx + 1
            positions.append((last_idx, title, icon))

    positions.sort(key=lambda x: x[0])
    sections = []
    for i, (pos, title, icon) in enumerate(positions):
        end = positions[i + 1][0] if i + 1 < len(positions) else len(full_text)
        content = full_text[pos + len(title):end].strip()
        clean_lines = []
        for line in content.split('\n'):
            line = line.strip()
            if line.startswith('Marion St Properties') or line.startswith('Draft v1.0') or line.startswith('Page '):
                continue
            if '...' in line and len(line) > 40:
                continue
            clean_lines.append(line)
        sections.append({'title': title, 'icon': icon, 'content': '\n'.join(clean_lines).strip()})
    return sections


def extract_sop_content():
    if not SOP_PDF:
        return []
    try:
        cache_token = os.path.getmtime(SOP_PDF)
    except OSError:
        cache_token = datetime.now().timestamp()
    return _extract_sop_content(SOP_PDF, cache_token)


# =====================
# RENDER TABS
# =====================

def render_sop_tab():
    sections = extract_sop_content()
    if not sections:
        st.warning("SOP Manual PDF not found. Place Marion_St_SOP_Manual.pdf in the data/ folder.")
        return

    st.markdown("### 📋 Marion St Properties — Standard Operating Procedures")
    st.caption("v1.0 Draft — Click any section to expand")

    heading_pattern = re.compile(r'^(Purpose|Trigger|Systems & Files|Systems|Files|Persons|Owners|Frequency|Procedure|Process|Checklist|Intake|Monthly reconciliation|Standard reports|Vacancy Report|Lead Sheet|Numbering convention|Revision log|Open items)', re.IGNORECASE)
    step_headings = {"procedure", "process", "steps", "checklist"}

    def build_steps(step_buffer):
        if not step_buffer:
            return ""
        items = ''.join(f"<li>{text}</li>" for text in step_buffer)
        return f"<ol class='sop-steps'>{items}</ol>"

    def build_grid(grid_dict):
        if not grid_dict:
            return ""
        cards = []
        for title, items in grid_dict.items():
            if not items:
                continue
            li = ''.join(f"<li>{val}</li>" for val in items)
            cards.append(f"<div class='sop-grid-card'><div class='sop-grid-title'>{title}</div><ul>{li}</ul></div>")
        if not cards:
            return ""
        return f"<div class='sop-grid'>{''.join(cards)}</div>"

    def add_to_grid(grid_dict, heading, line):
        heading_title = heading.title()
        values = []
        colon_match = re.match(r'^([^:]+):\s*(.+)$', line)
        if colon_match:
            key = colon_match.group(1).strip()
            val = colon_match.group(2).strip()
            values.append(f"{key}: {val}")
        else:
            values.append(line.strip())
        for val in values:
            if val:
                grid_dict.setdefault(heading_title, []).append(val)

    for sec in sections:
        with st.expander(f"{sec['icon']} {sec['title']}", expanded=False):
            lines = [re.sub(r'\s+', ' ', line.strip()) for line in sec['content'].split('\n') if line.strip()]
            step_buffer = []
            step_mode = False
            html_parts = []
            grid_dict = {}
            current_heading = None

            def flush_steps():
                nonlocal step_buffer, html_parts
                if step_buffer:
                    html_parts.append(build_steps(step_buffer))
                    step_buffer = []

            def flush_grid():
                nonlocal grid_dict, html_parts
                grid_html = build_grid(grid_dict)
                if grid_html:
                    html_parts.append(grid_html)
                grid_dict = {}

            current_heading = None
            for line in lines:
                heading_match = heading_pattern.match(line)
                if heading_match:
                    flush_steps()
                    flush_grid()
                    html_parts.append(f"<div class='sop-heading'>{line}</div>")
                    current_heading = heading_match.group(1).lower()
                    step_mode = current_heading in step_headings
                    continue

                if current_heading == 'purpose':
                    grid_match = re.match(r'^(Systems?|Files?):\s*(.+)', line, re.IGNORECASE)
                    if grid_match:
                        key = grid_match.group(1).title()
                        values = [v.strip(' •-') for v in re.split(r'[,;]|\s•', grid_match.group(2)) if v.strip(' •-')]
                        if values:
                            grid_dict.setdefault(key, []).extend(values)
                        continue

                if current_heading and 'report' in current_heading:
                    cleaned = line.strip(' •-')
                    if cleaned:
                        add_to_grid(grid_dict, current_heading, cleaned)
                        continue

                is_step_line = bool(re.match(r'^(\d+[\).]?\s+|[•\-]\s+)', line)) or step_mode
                if is_step_line:
                    clean = re.sub(r'^(\d+[\).]?\s+|[•\-]\s+)', '', line).strip()
                    if clean:
                        step_buffer.append(clean)
                    continue

                flush_steps()
                flush_grid()
                html_parts.append(f"<div class='sop-text'>{line}</div>")

            flush_steps()
            flush_grid()

            st.markdown(f"<div class='sop-section'>{''.join(html_parts)}</div>", unsafe_allow_html=True)


def render_tenancy_tab():
    tenants, details, summaries = load_tenancy()
    if not tenants:
        st.warning("MSP Tenancy.xlsx not found.")
        return

    expense_cfg = read_expense_config()

    # Get vacant spaces to flag in tenancy view
    vacant_keys, _vacant_meta = build_vacancy_lookup(tenants, include_auto=True)

    # Summary metrics
    active = [t for t in tenants if t['Tenant'] != 'Easement']
    # Add vacancy status
    for t in active:
        key = f"{t['Building']}|{t['Space']}"
        t['Status'] = '🔴 VACANT' if key in vacant_keys else ''
    total_sf = sum(t['SF'] for t in active if t['SF'] > 1)
    total_annual = sum(t['Annual'] for t in active)
    total_monthly = sum(t['Monthly'] for t in active)
    mtm_count = sum(1 for t in active if t.get('TTE_Label') == 'MTM')

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Portfolio SF", f"{total_sf:,.0f}")
    c2.metric("Annual Revenue", f"${total_annual:,.0f}")
    c3.metric("Monthly Revenue", f"${total_monthly:,.0f}")
    c4.metric("Month-to-Month", str(mtm_count))

    # Filter
    buildings = ['All'] + sorted(set(t['Building'] for t in active))
    selected = st.selectbox("Filter by Building", buildings, key="tenancy_filter")

    filtered = active if selected == 'All' else [t for t in active if t['Building'] == selected]

    # Display by building
    current_bldg = None
    for t in filtered:
        if t['Building'] != current_bldg:
            current_bldg = t['Building']
            bldg_tenants = [x for x in filtered if x['Building'] == current_bldg]
            bldg_annual = sum(x['Annual'] for x in bldg_tenants)
            st.markdown(f'<div class="building-header"><strong>{current_bldg}</strong> — {len(bldg_tenants)} tenants — ${bldg_annual:,.0f}/yr</div>', unsafe_allow_html=True)

            import pandas as pd
            df = pd.DataFrame(bldg_tenants)
            building_expense = float(expense_cfg.get(current_bldg, 0) or 0)
            df['CAM_Reimb'] = df.apply(
                lambda row: (building_expense * row['CAM_Pct']) if row.get('Is_NNN') and isinstance(row.get('CAM_Pct'), (int, float)) else 0,
                axis=1
            )
            df['Gross_Annual'] = df['Annual'] + df['CAM_Reimb']
            df['CAM_PSF'] = df.apply(
                lambda row: (row['CAM_Reimb'] / row['SF']) if row['SF'] and row['SF'] > 0 else 0,
                axis=1
            )
            df['Gross_PSF'] = df['PSF'] + df['CAM_PSF']

            df['Gross_Annual'] = df['Annual'] + df['CAM_Reimb']
            df['Total_Expenses'] = ''
            df['NOI'] = ''

            display_cols = ['Space', 'Tenant', 'Type', 'SF', 'Lease', 'Monthly', 'Annual', 'Gross_Annual', 'PSF', 'Gross_PSF', 'CAM_Pct', 'CAM_Reimb', 'TTE_Months', 'Options', 'Escalation', 'Next Anniv', 'Anniv_Months', 'Next Monthly', 'Delta Monthly', 'Status', 'Total_Expenses', 'NOI', 'TTE_Label', 'Is_NNN']

            # Calculate building totals
            b_monthly = df['Monthly'].sum()
            b_annual = df['Annual'].sum()
            b_sf = df.loc[df['SF'] > 1, 'SF'].sum()
            b_wavg_psf = b_annual / b_sf if b_sf > 0 else 0
            b_gross_annual = df['Gross_Annual'].sum()
            b_cam_reimb = df['CAM_Reimb'].sum()
            b_noi = b_gross_annual - building_expense

            # Create display dataframe
            display_df = df[display_cols].copy()
            display_df.rename(columns={
                'CAM_Pct': 'CAM %',
                'CAM_Reimb': 'CAM Reimb',
                'Gross_Annual': 'Gross Annual',
                'Gross_PSF': 'Gross PSF',
                'Total_Expenses': 'Total Expenses',
                'NOI': 'NOI',
                'TTE_Months': 'TTE (mo)',
                'Next Anniv': 'Next Anniversary',
                'Anniv_Months': 'Anniv Δ (mo)',
                'Next Monthly': 'New Rent',
                'Delta Monthly': 'Δ Monthly',
                'TTE_Label': 'TTE Label',
                'Is_NNN': 'NNN',
            }, inplace=True)
            display_df['CAM %'] = display_df.apply(
                lambda row: f"{(row['CAM %']*100):.1f}%" if row['NNN'] and isinstance(row['CAM %'], (int, float)) and row['CAM %'] not in (None, 0)
                else '-'
                , axis=1)
            display_df['CAM Reimb'] = display_df['CAM Reimb'].apply(lambda x: f"${x:,.0f}" if x and x != 0 else '$0')
            display_df['Gross Annual'] = display_df['Gross Annual'].apply(lambda x: f"${x:,.0f}" if x else '$0')
            display_df['Gross PSF'] = display_df['Gross PSF'].apply(lambda x: f"${x:,.2f}" if x else '$0')
            display_df['Total Expenses'] = ''
            display_df['NOI'] = ''
            display_df['TTE (mo)'] = display_df['TTE (mo)'].fillna(0).astype(int)
            display_df['Anniv Δ (mo)'] = display_df['Anniv Δ (mo)'].apply(
                lambda x: int(x) if x is not None and not (isinstance(x, float) and pd.isna(x)) else None
            )
            display_df['Escalation'] = display_df['Escalation'].apply(lambda x: f"{x:.1%}" if x and x > 0 else '-')
            display_df['Monthly'] = display_df['Monthly'].apply(lambda x: f"${x:,.0f}")
            display_df['Annual'] = display_df['Annual'].apply(lambda x: f"${x:,.0f}")
            display_df['PSF'] = display_df['PSF'].apply(lambda x: f"${x:,.2f}")
            display_df['SF'] = display_df['SF'].apply(lambda x: f"{x:,.0f}" if x > 1 else str(x))
            display_df['Next Anniversary'] = display_df['Next Anniversary'].apply(
                lambda d: d.strftime('%m/%d/%Y') if isinstance(d, (datetime, date)) else ('-' if not d else str(d))
            )
            display_df['New Rent'] = display_df['New Rent'].apply(
                lambda x: f"${x:,.0f}" if x is not None else '$0'
            )
            display_df['Δ Monthly'] = display_df['Δ Monthly'].apply(
                lambda x: f"{'+' if x and x > 0 else ''}${x:,.0f}" if x is not None else '$0'
            )

            # Pinned totals row
            totals_row = [{
                'Space': '', 'Tenant': 'TOTAL', 'Type': '', 'SF': f"{b_sf:,.0f}",
                'Lease': '', 'Monthly': f"${b_monthly:,.0f}", 'Annual': f"${b_annual:,.0f}",
                'PSF': f"${b_wavg_psf:,.2f}", 'Gross Annual': f"${b_gross_annual:,.0f}", 'Gross PSF': '',
                'CAM %': '', 'CAM Reimb': f"${b_cam_reimb:,.0f}",
                'Total Expenses': f"${building_expense:,.0f}" if building_expense else '$0', 'NOI': f"${b_noi:,.0f}",
                'TTE (mo)': '', 'Options': '', 'Escalation': '', 'Next Anniversary': '', 'Anniv Δ (mo)': '',
                'New Rent': '', 'Δ Monthly': '', 'Status': '', 'TTE Label': '', 'NNN': '',
            }]

            tte_formatter = JsCode("""
                function(params) {
                    if (params.value === null || params.value === undefined) { return '-'; }
                    return params.value.toString();
                }
            """)
            anniv_formatter = JsCode("""
                function(params) {
                    if (params.value === null || params.value === undefined) { return '-'; }
                    return params.value.toString();
                }
            """)
            column_configs = {
                'TTE (mo)': {'type': ['numericColumn'], 'valueFormatter': tte_formatter},
                'Anniv Δ (mo)': {'type': ['numericColumn'], 'valueFormatter': anniv_formatter},
                'TTE Label': {'hide': True},
                'NNN': {'hide': True},
            }

            show_grid(display_df, key=f"tenancy_{current_bldg}", pinned_bottom=totals_row,
                      column_configs=column_configs, tab_key="tenancy")

    # Portfolio totals
    st.divider()
    port_sf = sum(t['SF'] for t in filtered if t['SF'] > 1 and t['Tenant'] != 'Easement')
    port_monthly = sum(t['Monthly'] for t in filtered if t['Tenant'] != 'Easement')
    port_annual = sum(t['Annual'] for t in filtered if t['Tenant'] != 'Easement')
    port_wavg_psf = port_annual / port_sf if port_sf > 0 else 0
    pc1, pc2, pc3, pc4 = st.columns(4)
    pc1.metric("Total SF", f"{port_sf:,.0f}")
    pc2.metric("Total Monthly", f"${port_monthly:,.0f}")
    pc3.metric("Total Annual", f"${port_annual:,.0f}")
    pc4.metric("Wtd Avg $/SF", f"${port_wavg_psf:,.2f}")

    render_column_config_editor(
        'tenancy',
        ['Space', 'Tenant', 'Type', 'SF', 'Lease', 'Monthly', 'Annual', 'PSF', 'Gross PSF', 'CAM %', 'CAM Reimb',
         'Total Expenses', 'NOI', 'TTE (mo)', 'Options', 'Escalation', 'Next Anniversary', 'Anniv Δ (mo)', 'New Rent', 'Δ Monthly', 'Status']
    )
    render_expense_editor()


def get_vacant_spaces():
    """Get manually marked vacant spaces from Google Sheets."""
    sheet = get_gsheet()
    if not sheet:
        return []
    try:
        try:
            ws = sheet.worksheet('Vacant Spaces')
        except gspread.exceptions.WorksheetNotFound:
            ws = sheet.add_worksheet('Vacant Spaces', rows=100, cols=8)
            ws.update(values=[['Building', 'Space', 'Tenant', 'SF', 'Vacancy Date', 'Marked By', 'Notes', 'Status']], range_name='A1:H1')
        return ws.get_all_records()
    except Exception:
        return []


def build_vacancy_lookup(tenants=None, include_auto=True):
    """Return (set(keys), metadata dict) for vacant spaces keyed by Building|Space."""
    records = get_vacant_spaces()
    keys = set()
    meta = {}
    for rec in records:
        b = (rec.get('Building') or '').strip()
        space = str(rec.get('Space') or '').strip()
        if not b or not space:
            continue
        key = f"{b}|{space}"
        keys.add(key)
        meta[key] = {
            'tenant': rec.get('Tenant', ''),
            'vacancy_date': rec.get('Vacancy Date') or rec.get('Date Marked')
        }

    if include_auto and tenants:
        for t in tenants:
            if t.get('Tenant') == 'Easement':
                continue
            auto_flag = (t.get('Monthly') == 0 and t.get('Annual') == 0) or t.get('Status') == '🔴 VACANT'
            if not auto_flag:
                continue
            b = (t.get('Building') or '').strip()
            space = str(t.get('Space') or '').strip()
            if not b or not space:
                continue
            key = f"{b}|{space}"
            keys.add(key)
            meta.setdefault(key, {'tenant': t.get('Tenant', '')})
    return keys, meta


def add_vacant_space(entry):
    sheet = get_gsheet()
    if not sheet:
        return False
    try:
        try:
            ws = sheet.worksheet('Vacant Spaces')
        except gspread.exceptions.WorksheetNotFound:
            ws = sheet.add_worksheet('Vacant Spaces', rows=100, cols=8)
            ws.update(values=[['Building', 'Space', 'Tenant', 'SF', 'Vacancy Date', 'Marked By', 'Notes', 'Status']], range_name='A1:H1')
        ws.append_row([
            entry['building'], entry['space'], entry['tenant'], entry['sf'],
            entry.get('vacancy_date', datetime.now().strftime('%Y-%m-%d')),
            entry.get('marked_by', ''), entry.get('notes', ''), 'Vacant'
        ], value_input_option='USER_ENTERED')
        return True
    except Exception as e:
        st.error(f"Error: {e}")
        return False


def remove_vacant_space(row_idx):
    sheet = get_gsheet()
    if not sheet:
        return False
    try:
        ws = sheet.worksheet('Vacant Spaces')
        ws.delete_rows(row_idx + 2)
        return True
    except Exception:
        return False


def render_vacancy_tab():
    tenants, _, _ = load_tenancy()
    active = [t for t in tenants if t['Tenant'] not in ('Easement',)]

    # Get vacant lookup (manual + auto)
    vacant_keys, vacant_meta = build_vacancy_lookup(active, include_auto=True)
    vacant_records = get_vacant_spaces()
    manual_keys = set(f"{v.get('Building', '')}|{v.get('Space', '')}" for v in vacant_records)
    auto_vacant = [t for t in active if t['Monthly'] == 0 and t['Annual'] == 0]
    all_vacant_keys = vacant_keys

    # --- Mark space vacant ---
    st.markdown("### ✏️ Mark a Space Vacant")
    all_spaces = sorted(set(f"{t['Building']} #{t['Space']} — {t['Tenant']}" for t in active))

    with st.form("mark_vacant", clear_on_submit=True):
        vc1, vc2, vc3 = st.columns(3)
        selected_space = vc1.selectbox("Select Space", [""] + all_spaces)
        vacancy_date = vc2.date_input("Vacancy Date", value=date.today())
        marked_by = vc3.text_input("Your Name")
        notes = st.text_input("Notes (optional)")
        mark_btn = st.form_submit_button("🔴 Mark as Vacant", type="primary")

        if mark_btn and selected_space:
            parts = selected_space.split(' — ')
            bldg_space = parts[0]
            tenant_name = parts[1] if len(parts) > 1 else ''
            bldg_parts = bldg_space.split(' #')
            building = bldg_parts[0]
            space = bldg_parts[1] if len(bldg_parts) > 1 else ''

            sf = 0
            for t in active:
                if t['Building'] == building and str(t['Space']) == space:
                    sf = t['SF']
                    break

            success = add_vacant_space({
                'building': building, 'space': space, 'tenant': tenant_name,
                'sf': sf, 'marked_by': marked_by, 'notes': notes,
                'vacancy_date': str(vacancy_date),
            })
            if success:
                st.success(f"Marked {building} #{space} as vacant!")
                st.cache_resource.clear()
                st.rerun()

    st.divider()

    # --- Currently Vacant ---
    st.markdown("### 🔴 Currently Vacant Spaces")

    if vacant_records or auto_vacant:
        import pandas as pd
        vacant_display = []
        for v in vacant_records:
            # Calculate days vacant
            vac_date_str = v.get('Vacancy Date', v.get('Date Marked', ''))
            days_vacant = '—'
            try:
                vac_date = datetime.strptime(vac_date_str, '%Y-%m-%d').date()
                days = (TODAY - vac_date).days
                if days < 0:
                    days_vacant = f"Starts in {-days}d"
                elif days < 30:
                    days_vacant = f"{days}d"
                elif days < 365:
                    days_vacant = f"{days // 30}mo {days % 30}d"
                else:
                    days_vacant = f"{days // 365}yr {(days % 365) // 30}mo"
            except (ValueError, TypeError):
                pass

            vacant_display.append({
                'Building': v.get('Building', ''),
                'Space': v.get('Space', ''),
                'Last Tenant': v.get('Tenant', ''),
                'SF': v.get('SF', ''),
                'Vacant Since': vac_date_str,
                'Days Vacant': days_vacant,
                'Notes': v.get('Notes', ''),
            })
        for t in auto_vacant:
            key = f"{t['Building']}|{t['Space']}"
            if key not in manual_keys:
                vacant_display.append({
                    'Building': t['Building'],
                    'Space': t['Space'],
                    'Last Tenant': t['Tenant'],
                    'SF': t['SF'],
                    'Vacant Since': '—',
                    'Days Vacant': '—',
                    'Notes': 'Auto-detected ($0 rent)',
                })

        if vacant_display:
            show_grid(pd.DataFrame(vacant_display), key="vacant_spaces", tab_key="vacancy_current")

            # Remove buttons for manual entries
            if vacant_records:
                st.caption("Remove a manually marked space:")
                for i, v in enumerate(vacant_records):
                    col1, col2 = st.columns([8, 1])
                    col1.text(f"{v.get('Building', '')} #{v.get('Space', '')} — {v.get('Tenant', '')}")
                    if col2.button("✅ Leased / Not Vacant", key=f"unvacant_{i}"):
                        remove_vacant_space(i)
                        st.cache_resource.clear()
                        st.rerun()
        else:
            st.success("✅ No currently vacant spaces — 100% occupied!")
    else:
        st.success("✅ No currently vacant spaces — 100% occupied!")

    # At risk — exclude spaces already marked vacant (manual only)
    st.markdown("### 🟡 At Risk — Expiring Within 12 Months")
    at_risk = [t for t in active if (t.get('TTE_Label') == 'MTM' or (0 < t['TTE_Days'] <= 365))
               and f"{t['Building']}|{t['Space']}" not in manual_keys]
    if at_risk:
        import pandas as pd
        risk_df = pd.DataFrame(at_risk)[['Building', 'Space', 'Tenant', 'SF', 'Monthly', 'TTE_Months', 'Exp Date']]
        risk_df.rename(columns={'TTE_Months': 'TTE (mo)'}, inplace=True)
        risk_df['Monthly'] = risk_df['Monthly'].apply(lambda x: f"${x:,.0f}")
        show_grid(
            risk_df,
            key="at_risk",
            column_configs={'TTE (mo)': {'type': ['numericColumn']}},
            tab_key="vacancy_risk"
        )
    else:
        st.success("✅ No tenants expiring within 12 months.")

    st.divider()

    # --- VACANCY ACTIVITY LOG (Google Sheets) ---
    st.markdown("### 📋 Vacancy Activity Log")
    st.caption("Shared across all users via Google Sheets")

    all_spaces = sorted(set(f"{t['Building']} #{t['Space']}" for t in active))

    with st.form("add_activity", clear_on_submit=True):
        st.markdown("**Add New Entry**")
        ac1, ac2 = st.columns(2)
        activity_date = ac1.date_input("Date", value=date.today())
        activity_space = ac2.selectbox("Space", [""] + all_spaces)
        ac3, ac4 = st.columns(2)
        prospect = ac3.text_input("Prospect Name")
        broker = ac4.text_input("Broker")
        ac5, ac6 = st.columns(2)
        activity_type = ac5.selectbox("Type", ["Showing", "Inquiry", "LOI", "Application", "Lease Signed", "Follow-up"])
        added_by = ac6.text_input("Your Name")
        feedback = st.text_area("Feedback / Notes")
        submitted = st.form_submit_button("➕ Add Entry", type="primary")

        if submitted and activity_space:
            success = add_activity_entry({
                'date': str(activity_date),
                'space': activity_space,
                'prospect': prospect,
                'broker': broker,
                'type': activity_type,
                'feedback': feedback,
                'added_by': added_by,
            })
            if success:
                st.success("Entry added!")
                st.cache_resource.clear()
                st.rerun()

    # Show existing entries
    activities = get_activity_data()
    if activities:
        import pandas as pd
        df = pd.DataFrame(activities)
        display_cols = ['Date', 'Building', 'Space', 'Prospect', 'Broker', 'Type', 'Feedback', 'Added By', 'Timestamp']
        for col in display_cols:
            if col not in df.columns:
                df[col] = ''
        df = df[display_cols]
        df.rename(columns={'Added By': 'Added_By'}, inplace=True)
        show_grid(df, key="vacancy_activity", tab_key="vacancy_activity")

        # Delete controls
        st.caption("Delete an entry:")
        for idx, a in enumerate(activities):
            entry_label = f"{a.get('Date', '')} — {a.get('Space', '')} ({a.get('Type', '')})"
            if st.button(f"🗑️ {entry_label}", key=f"del_act_{idx}"):
                delete_activity_entry(idx)
                st.cache_resource.clear()
                st.rerun()
    else:
        st.info("No activity entries yet. Add one above!")

    st.divider()

    # --- MARKETING ---
    st.markdown("### 📢 Current Marketing")

    # File upload
    st.markdown("**Upload a brochure / flyer (PDF, image, etc.)**")
    uploaded_file = st.file_uploader(
        "Drag & drop or browse", type=["pdf", "png", "jpg", "jpeg", "doc", "docx"],
        key="mktg_upload", label_visibility="collapsed"
    )

    with st.form("add_marketing", clear_on_submit=True):
        mc1, mc2 = st.columns(2)
        mktg_space = mc1.selectbox("Space", [""] + all_spaces, key="mktg_space")
        mktg_desc = mc2.text_input("Description (e.g. LoopNet Listing)")
        mc3, mc4 = st.columns(2)
        mktg_url = mc3.text_input("URL (optional if uploading a file)")
        mktg_by = mc4.text_input("Your Name", key="mktg_by")
        mktg_submit = st.form_submit_button("➕ Add Marketing Entry", type="primary")

        if mktg_submit and mktg_space:
            file_url = mktg_url
            file_name = ''

            # Handle uploaded file — save to Google Drive via Sheets note, or encode
            if uploaded_file is not None:
                import base64
                file_bytes = uploaded_file.read()
                file_name = uploaded_file.name

                # Save file to shared Dropbox marketing folder
                marketing_dir = None
                for mp in [
                    "/home/node/OpenClaw/Share Jason/General Procedures/marketing",
                    os.path.join(os.path.dirname(__file__), "marketing"),
                ]:
                    try:
                        os.makedirs(mp, exist_ok=True)
                        marketing_dir = mp
                        break
                    except Exception:
                        continue

                if marketing_dir:
                    # Add timestamp to avoid conflicts
                    safe_name = f"{datetime.now().strftime('%Y%m%d_%H%M')}_{file_name}"
                    file_path = os.path.join(marketing_dir, safe_name)
                    with open(file_path, 'wb') as f:
                        f.write(file_bytes)
                    file_url = f"[File saved: {safe_name}]"
                    if not mktg_desc:
                        mktg_desc = file_name

            if file_url or file_name:
                success = add_marketing_entry({
                    'space': mktg_space,
                    'description': mktg_desc or file_name,
                    'url': file_url,
                    'added_by': mktg_by,
                })
                if success:
                    st.success("Marketing entry added!")
                    st.cache_resource.clear()
                    st.rerun()
            else:
                st.warning("Please provide a URL or upload a file.")

    marketing = get_marketing_data()
    if marketing:
        import pandas as pd
        mdf = pd.DataFrame(marketing)
        show_grid(mdf, key="marketing", tab_key="marketing")

        # Show download links for uploaded files
        marketing_dir = None
        for mp in [
            "/home/node/OpenClaw/Share Jason/General Procedures/marketing",
            os.path.join(os.path.dirname(__file__), "marketing"),
        ]:
            if os.path.isdir(mp):
                marketing_dir = mp
                break

        if marketing_dir:
            files = sorted(os.listdir(marketing_dir))
            if files:
                st.caption("📎 Uploaded files:")
                for fname in files:
                    fpath = os.path.join(marketing_dir, fname)
                    with open(fpath, 'rb') as f:
                        st.download_button(
                            f"📄 {fname}", f.read(), file_name=fname,
                            key=f"dl_{fname}", mime="application/octet-stream"
                        )
    else:
        st.info("No marketing entries yet.")
        mdf = pd.DataFrame(columns=['Space', 'Description', 'URL', 'Added By', 'Date', 'Status'])

    render_column_config_editor('vacancy_current', ['Building', 'Space', 'Last Tenant', 'SF', 'Vacant Since', 'Days Vacant', 'Notes'])
    render_column_config_editor('vacancy_risk', ['Building', 'Space', 'Tenant', 'SF', 'Monthly', 'TTE (mo)', 'Exp Date'])
    render_column_config_editor('vacancy_activity', ['Date', 'Building', 'Space', 'Prospect', 'Broker', 'Type', 'Feedback', 'Added_By', 'Timestamp'])
    render_column_config_editor('marketing', list(mdf.columns))


def render_insurance_tab():
    tenants, _, summaries = load_tenancy()
    coi_data = scan_coi_files()
    today_dt = datetime.now()
    vacant_keys, vacant_meta = build_vacancy_lookup(tenants, include_auto=True)

    st.markdown("### 🛡️ Certificate of Insurance Reconciliation")

    # Build insurance data
    ins_rows = []
    summary = {'total': 0, 'covered': 0, 'expired': 0, 'missing': 0, 'expiring_soon': 0}

    for s in summaries:
        b = s.get('Buidling', '')
        name = s.get('Tenant Name', '')
        ttype = s.get('Type', '')
        unit = s.get('Unit', '')
        if not b or not name or name == 'Easement':
            continue
        if ttype and 'apartment' in str(ttype).lower():
            continue

        key = f"{b}|{str(unit).strip()}"
        if key in vacant_keys:
            last_name = vacant_meta.get(key, {}).get('tenant') or name
            ins_rows.append({
                'Building': b,
                'Tenant': f"VACANT — {last_name}",
                'Unit': str(unit),
                'Type': ttype,
                'COI': 'N/A',
                'Expiration': '—',
                'Days Left': '—',
                'Status': 'Vacant',
            })
            continue

        summary['total'] += 1
        b_certs = coi_data.get(b, [])
        matched = None
        for cert in b_certs:
            if fuzzy_match_tenant(name, cert.get('insured_name')):
                matched = cert
                break
        if not matched:
            for cert in b_certs:
                if fuzzy_match_tenant(name, cert['filename']):
                    matched = cert
                    break

        if matched:
            exp = matched['exp_date']
            if exp:
                days_left = (exp - today_dt).days
                if exp < today_dt:
                    status = 'EXPIRED'
                    summary['expired'] += 1
                elif days_left < 60:
                    status = f'{days_left}d left'
                    summary['expiring_soon'] += 1
                    summary['covered'] += 1
                else:
                    status = 'Active'
                    summary['covered'] += 1
                exp_str = exp.strftime('%m/%d/%Y')
            else:
                exp_str = 'Unknown'
                days_left = None
                status = 'No date'
                summary['covered'] += 1
            ins_rows.append({
                'Building': b, 'Tenant': name, 'Unit': str(unit), 'Type': ttype,
                'COI': '✅ YES', 'Expiration': exp_str,
                'Days Left': str(days_left) if days_left is not None else '—',
                'Status': status, 'File': matched['filename'],
            })
        else:
            summary['missing'] += 1
            ins_rows.append({
                'Building': b, 'Tenant': name, 'Unit': str(unit), 'Type': ttype,
                'COI': '❌ NO', 'Expiration': '—', 'Days Left': '—',
                'Status': 'MISSING', 'File': '',
            })

    # Summary metrics
    coverage_pct = (summary['covered'] / summary['total'] * 100) if summary['total'] else 0
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Coverage", f"{coverage_pct:.0f}%")
    c2.metric("Active", summary['covered'] - summary['expiring_soon'])
    c3.metric("Expiring Soon", summary['expiring_soon'])
    c4.metric("Expired", summary['expired'])
    c5.metric("Missing", summary['missing'])

    # Display by building
    for building_name in BUILDING_MAP:
        b_rows = [r for r in ins_rows if r['Building'] == building_name]
        if not b_rows:
            continue
        b_covered = sum(1 for r in b_rows if '✅' in r['COI'])
        n_certs = len(coi_data.get(building_name, []))
        code = BUILDING_MAP[building_name]['code']

        st.markdown(f'<div class="building-header"><strong>▎ {building_name} ({code})</strong> — {b_covered}/{len(b_rows)} covered · {n_certs} files on disk</div>', unsafe_allow_html=True)

        import pandas as pd
        df = pd.DataFrame(b_rows)
        display_cols = ['Tenant', 'Unit', 'Type', 'COI', 'Expiration', 'Days Left', 'Status']
        show_grid(df[display_cols], key=f"ins_{building_name}", tab_key="insurance")

    render_column_config_editor('insurance', ['Tenant', 'Unit', 'Type', 'COI', 'Expiration', 'Days Left', 'Status'])


def render_deposits_tab():
    tenants, details, summaries = load_tenancy()
    active = [t for t in tenants if t['Tenant'] not in ('Easement',)]
    today_dt = datetime.now()
    vacant_keys, vacant_meta = build_vacancy_lookup(tenants, include_auto=True)

    st.markdown("### 💰 Security Deposit Reconciliation")

    # Build detailed SD data from term rows
    import pandas as pd
    dep_data = []

    for tid, rows in details.items():
        first = rows[0]
        tenant_name = first.get('Tenant', '')
        building = first.get('Building', '')
        if tenant_name in ('Easement', '-') or not building:
            continue

        # Find current SD (the row whose date range includes today)
        current_sd = first.get('Sec Dep', 0) or 0
        next_sd = None
        next_sd_date = None
        current_year_row = None

        for r in rows:
            start = r.get('Start Date')
            end = r.get('End Date')
            sd = r.get('Sec Dep', 0) or 0
            if isinstance(start, (datetime, date)) and isinstance(end, (datetime, date)):
                start_d = start.date() if isinstance(start, datetime) else start
                end_d = end.date() if isinstance(end, datetime) else end
                if start_d <= TODAY <= end_d:
                    current_sd = sd
                    current_year_row = r
                elif start_d > TODAY and next_sd is None:
                    next_sd = sd
                    next_sd_date = start_d

        # If we didn't find current, use first row's SD
        if current_year_row is None:
            current_sd = first.get('Sec Dep', 0) or 0

        # Get tenant type from summary
        tenant_type = ''
        for s in summaries:
            if s.get('Tenant Name') == tenant_name and s.get('Buidling') == building:
                tenant_type = s.get('Type', '') or ''
                break

        # Find TTE from active tenants list
        tte_str = 'N/A'
        lease_type = first.get('Type', '')
        for t in active:
            if t['Tenant'] == tenant_name and t['Building'] == building:
                tte_str = t['TTE']
                lease_type = t['Lease']
                break

        sd_anniv_date = next_sd_date if next_sd_date and next_sd is not None and next_sd != current_sd else None
        sd_new_amount = next_sd if sd_anniv_date else None
        sd_delta = (next_sd - current_sd) if sd_new_amount is not None else None

        space_val = str(first.get('Space', '')).strip()
        key = f"{building}|{space_val}"
        is_vacant = key in vacant_keys
        display_tenant = f"VACANT — {vacant_meta.get(key, {}).get('tenant') or tenant_name}" if is_vacant else tenant_name

        dep_data.append({
            'Building': building,
            'Tenant': tenant_name,
            'Display Tenant': display_tenant,
            'Space': space_val,
            'Type': tenant_type,
            'Lease': lease_type,
            'Current SD': current_sd,
            'Current SD Fmt': f"${current_sd:,.2f}" if current_sd > 0 else '❌ $0.00',
            'SD Anniversary': sd_anniv_date,
            'New SD Amount': sd_new_amount,
            'SD Delta': sd_delta,
            'TTE': tte_str,
        })

    dep_data.sort(key=lambda d: (d['Building'], d['Tenant']))

    # Summary metrics
    total_deposits = sum(d['Current SD'] for d in dep_data)
    with_dep = sum(1 for d in dep_data if d['Current SD'] > 0)
    without_dep = sum(1 for d in dep_data if d['Current SD'] == 0)
    upcoming_changes = sum(1 for d in dep_data if d['SD Anniversary'])

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total Deposits Held", f"${total_deposits:,.2f}")
    c2.metric("Tenants w/ Deposit", str(with_dep))
    c3.metric("Tenants w/o Deposit", str(without_dep))
    c4.metric("Upcoming SD Changes", str(upcoming_changes))

    # Display by building
    for building_name in BUILDING_MAP:
        b_deps = [d for d in dep_data if d['Building'] == building_name]
        if not b_deps:
            continue
        b_total = sum(d['Current SD'] for d in b_deps)
        code = BUILDING_MAP[building_name]['code']
        st.markdown(f'<div class="building-header"><strong>▎ {building_name} ({code})</strong> — ${b_total:,.2f} total deposits held</div>', unsafe_allow_html=True)

        df = pd.DataFrame(b_deps)
        display_cols = ['Display Tenant', 'Space', 'Type', 'Lease', 'Current SD Fmt', 'SD Anniversary', 'New SD Amount', 'SD Delta', 'TTE']
        display_df = df[display_cols].copy()
        display_df.columns = ['Tenant', 'Space', 'Type', 'Lease', 'Current SD', 'SD Anniversary', 'New SD', 'Δ SD', 'TTE']
        display_df['SD Anniversary'] = display_df['SD Anniversary'].apply(
            lambda d: d.strftime('%m/%d/%Y') if isinstance(d, (datetime, date)) else ('-' if not d else str(d))
        )
        display_df['New SD'] = display_df['New SD'].apply(lambda x: f"${x:,.0f}" if x is not None else '$0')
        display_df['Δ SD'] = display_df['Δ SD'].apply(
            lambda x: f"{'+' if x and x > 0 else ''}${x:,.0f}" if x is not None else '$0'
        )
        show_grid(display_df, key=f"dep_{building_name}", tab_key="deposits")

    render_column_config_editor('deposits', ['Tenant', 'Space', 'Type', 'Lease', 'Current SD', 'SD Anniversary', 'New SD', 'Δ SD', 'TTE'])

    # Portfolio total
    st.divider()
    st.metric("Portfolio Total Security Deposits", f"${total_deposits:,.2f}")


# =====================
# MAIN APP
# =====================

st.markdown("## 🏢 MSP Property Dashboard")
st.caption(f"Marion Street Properties · {TODAY.strftime('%B %d, %Y')}")

tab_sop, tab_tenancy, tab_vacancy, tab_insurance, tab_deposits = st.tabs([
    "📋 SOPs", "🏠 Current Tenancy", "🏚️ Vacancy", "🛡️ Insurance", "💰 Security Deposits"
])

with tab_sop:
    render_sop_tab()

with tab_tenancy:
    render_tenancy_tab()

with tab_vacancy:
    render_vacancy_tab()

with tab_insurance:
    render_insurance_tab()

with tab_deposits:
    render_deposits_tab()
