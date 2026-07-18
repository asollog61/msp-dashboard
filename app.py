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
    "114 Central": {"code": "MSP114", "dest_folder": "114 Central", "share": "0_114Share"},
    "15 South": {"code": "MSP15", "dest_folder": "15 South", "share": "0_15Share"},
    "36 South": {"code": "MSP36", "dest_folder": "36 South", "share": "0_36Share"},
    "1280 Springfield": {"code": "MSP1280", "dest_folder": "1280 Springfield", "share": "0_1280Share"},
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
    .ag-right-aligned-header .ag-header-cell-label { justify-content: flex-end !important; }
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
    .sop-table { width: 100%; border-collapse: collapse; margin: 8px 0 12px; font-size: 0.85rem; }
    .sop-table th { background: rgba(100,140,255,0.15); color: #c8d4ff; text-align: left; padding: 6px 10px; border: 1px solid rgba(255,255,255,0.1); font-weight: 600; }
    .sop-table td { padding: 5px 10px; border: 1px solid rgba(255,255,255,0.08); }
    .sop-table tr:nth-child(even) { background: rgba(255,255,255,0.03); }
    .sop-table tr:hover { background: rgba(100,140,255,0.08); }
    .mobile-card { background: #161b22; border: 1px solid #2d333b; border-radius: 8px; padding: 12px 16px; margin-bottom: 8px; }
    .mobile-card-header { font-size: 15px; font-weight: 700; color: #58a6ff; margin-bottom: 6px; border-bottom: 1px solid #2d333b; padding-bottom: 6px; }
    .mobile-card-row { display: flex; justify-content: space-between; padding: 2px 0; font-size: 13px; }
    .mobile-card-label { color: #8b949e; }
    .mobile-card-value { color: #e6edf3; font-weight: 500; text-align: right; }
    .mobile-card-vacant { border-left: 3px solid #f85149; }
</style>
""", unsafe_allow_html=True)

# --- MOBILE VIEW TOGGLE ---
if 'mobile_view' not in st.session_state:
    st.session_state.mobile_view = False


def show_mobile_cards(df, key_col=None, header_col='Tenant', highlight_cols=None, skip_cols=None):
    """Render a DataFrame as stacked mobile-friendly cards."""
    if skip_cols is None:
        skip_cols = set()
    if highlight_cols is None:
        highlight_cols = ['Monthly', 'Annual', 'MTE', 'Status']
    for _, row in df.iterrows():
        header = str(row.get(header_col, ''))
        is_vacant = 'VACANT' in header.upper() if header else False
        card_class = 'mobile-card mobile-card-vacant' if is_vacant else 'mobile-card'
        rows_html = []
        for col in df.columns:
            if col == header_col or col in skip_cols:
                continue
            val = row[col]
            if val is None or (isinstance(val, float) and pd.isna(val)) or str(val).strip() in ('', '-', '—'):
                continue
            bold = 'font-weight:700;' if col in highlight_cols else ''
            rows_html.append(
                f'<div class="mobile-card-row">'
                f'<span class="mobile-card-label">{col}</span>'
                f'<span class="mobile-card-value" style="{bold}">{val}</span>'
                f'</div>'
            )
        html = f'<div class="{card_class}"><div class="mobile-card-header">{header}</div>{"".join(rows_html)}</div>'
        st.markdown(html, unsafe_allow_html=True)


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


def _read_gsheet_config(tab_name):
    """Read a JSON config blob stored in cell A1 of a Google Sheet tab. Uses session_state cache."""
    cache_key = f"_gsheet_cfg_{tab_name}"
    if cache_key in st.session_state:
        return st.session_state[cache_key]
    try:
        sheet = get_gsheet()
        if not sheet:
            return None
        ws = sheet.worksheet(tab_name)
        val = ws.acell('A1').value
        if val:
            data = json.loads(val)
            st.session_state[cache_key] = data
            return data
    except Exception:
        pass
    return None


def _write_gsheet_config(tab_name, data):
    """Write a JSON config blob to cell A1 of a Google Sheet tab. Creates tab if needed."""
    try:
        sheet = get_gsheet()
        if not sheet:
            return False
        try:
            ws = sheet.worksheet(tab_name)
        except Exception:
            ws = sheet.add_worksheet(title=tab_name, rows=1, cols=1)
        ws.update_acell('A1', json.dumps(data))
        # Clear session_state cache for this tab
        cache_key = f"_gsheet_cfg_{tab_name}"
        if cache_key in st.session_state:
            del st.session_state[cache_key]
        return True
    except Exception:
        return False


def read_column_config():
    # Try Google Sheets first (persistent across deploys), fall back to local JSON
    data = _read_gsheet_config('Config: Columns')
    if data and isinstance(data, dict):
        return data
    return ensure_json_file(COLUMN_CONFIG_FILE, {})


def write_column_config(data):
    # Write to Google Sheets (primary) and local JSON (backup)
    _write_gsheet_config('Config: Columns', data)
    try:
        with open(COLUMN_CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=2)
    except Exception:
        pass


def read_expense_config():
    default = {name: 0 for name in BUILDING_MAP.keys()}
    data = _read_gsheet_config('Config: Expenses')
    if data and isinstance(data, dict):
        return data
    return ensure_json_file(EXPENSE_CONFIG_FILE, default)


def write_expense_config(data):
    _write_gsheet_config('Config: Expenses', data)
    try:
        with open(EXPENSE_CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=2)
    except Exception:
        pass


def get_column_width_overrides(tab_key):
    if not tab_key:
        return {}
    data = read_column_config()
    cfg = data.get(tab_key, {})
    clean = {}
    for col, width in cfg.items():
        if width == '' or width is None:
            continue
        try:
            clean[col] = int(width)
        except (ValueError, TypeError):
            continue
    return clean


def get_column_order(tab_key, default_cols):
    """Return column list reordered per config. Columns in config come first (in config order),
    then any remaining columns not in config keep their original order."""
    if not tab_key:
        return default_cols
    data = read_column_config()
    cfg = data.get(tab_key, {})
    if not cfg:
        return default_cols
    config_cols = [c for c in cfg.keys() if c in default_cols]
    remaining = [c for c in default_cols if c not in config_cols]
    return config_cols + remaining


def save_column_width_overrides(tab_key, overrides):
    data = read_column_config()
    data[tab_key] = overrides
    write_column_config(data)


def render_column_config_editor(tab_key, columns):
    if not tab_key or not columns:
        return
    existing = read_column_config().get(tab_key, {})
    # Show columns in config order first (preserves reordering), then any new ones
    ordered_cols = [c for c in existing.keys() if c in columns]
    remaining = [c for c in columns if c not in ordered_cols]
    all_cols = ordered_cols + remaining
    default_lines = []
    for col in all_cols:
        width = existing.get(col, '')
        default_lines.append(f"{col},{width}")
    with st.expander(f"⚙️ Column Config — {tab_key.title()} (order + width)"):
        st.write("Reorder lines to change column order. Set width after comma (blank = auto). Save to apply.")
        text = st.text_area(
            "column_widths", value='\n'.join(default_lines), height=150,
            key=f"cfg_text_{tab_key}", label_visibility="collapsed"
        )
        if st.button("Save Column Config", key=f"cfg_save_{tab_key}"):
            from collections import OrderedDict
            new_cfg = OrderedDict()
            for line in text.splitlines():
                parts = [p.strip() for p in line.split(',')]
                if not parts or not parts[0]:
                    continue
                col_name = parts[0]
                width = None
                if len(parts) >= 2 and parts[1]:
                    try:
                        width = int(parts[1])
                    except ValueError:
                        width = None
                new_cfg[col_name] = width if width else ""
            save_column_width_overrides(tab_key, dict(new_cfg))
            st.success("Saved! Column order + widths updated.")
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
    # Reorder columns based on config
    if tab_key:
        ordered_cols = get_column_order(tab_key, list(df.columns))
        df = df[ordered_cols]

    # Mobile view: render cards instead of grid
    if st.session_state.get('mobile_view', False):
        header_col = 'Tenant' if 'Tenant' in df.columns else ('Last Tenant' if 'Last Tenant' in df.columns else df.columns[0])
        skip = {'TTE Label', 'NNN', 'Is_NNN'}
        show_mobile_cards(df, header_col=header_col, skip_cols=skip)
        return

    gb = GridOptionsBuilder.from_dataframe(df)
    gb.configure_default_column(
        sortable=True,
        filterable=True,
        resizable=True,
        suppressSizeToFit=False,
        autoWidth=True,
        wrapHeaderText=True,
        autoHeaderHeight=True,
        cellStyle={'text-align': 'right'},
        headerClass='ag-right-aligned-header',
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

    # Render 'View' column as a clickable link (opens PDF in new tab)
    if 'View' in df.columns:
        # Use HTML cell renderer with proper escaping for the URL
        view_renderer = JsCode("""
            class ViewCellRenderer {
                init(params) {
                    this.eGui = document.createElement('div');
                    this.eGui.style.textAlign = 'center';
                    if (params.value) {
                        this.eGui.innerHTML = '<a href="' + params.value + '" target="_blank" rel="noopener" style="color:#58a6ff;text-decoration:none;font-weight:600;">📄 View</a>';
                    } else {
                        this.eGui.innerHTML = '';
                    }
                }
                getGui() {
                    return this.eGui;
                }
            }
        """)
        gb.configure_column('View', cellRenderer=view_renderer, width=90, minWidth=70, suppressSizeToFit=True, filter=False, sortable=False)

    if column_configs:
        for col, cfg in column_configs.items():
            if col in df.columns:
                gb.configure_column(col, **cfg)

    overrides = get_column_width_overrides(tab_key)
    has_overrides = bool(overrides)
    for col, width in overrides.items():
        if col in df.columns:
            gb.configure_column(col, width=width, initialWidth=width, minWidth=width, maxWidth=width + 50, suppressSizeToFit=True, suppressAutoSize=True)

    grid_options = gb.build()

    # If user has configured widths, don't let fit_columns override them
    if has_overrides:
        fit_columns = False
        grid_options['suppressColumnVirtualisation'] = True

    # Add pinned bottom row for totals
    if pinned_bottom is not None:
        grid_options['pinnedBottomRowData'] = pinned_bottom
        grid_options['getRowStyle'] = JsCode("""
            function(params) {
                if (params.node.rowPinned === 'bottom') {
                    return {'background-color': '#1a3a5c', 'font-weight': '600', 'border-top': '2px solid #58a6ff'};
                }
            }
        """)

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
def get_gspread_client():
    scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
    if SERVICE_ACCOUNT_FILE:
        creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=scopes)
    elif SERVICE_ACCOUNT_INFO:
        creds = Credentials.from_service_account_info(SERVICE_ACCOUNT_INFO, scopes=scopes)
    else:
        return None
    return gspread.authorize(creds)

@st.cache_resource(ttl=300)
def get_gsheet():
    gc = get_gspread_client()
    if gc is None:
        return None
    return gc.open_by_key(SHEET_ID)


@st.cache_data(ttl=120)
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
        get_activity_data.clear()
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
        get_activity_data.clear()
        return True
    except Exception as e:
        st.error(f"Error deleting: {e}")
        return False


@st.cache_data(ttl=120)
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
        get_marketing_data.clear()
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
@st.cache_data(ttl=300)
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

    # Term detail rows: auto-detect the header row (the one starting with 'ID').
    # Historically row 45, but inserts can shift it (e.g. to 46). Search 43-50.
    detail_header_row = 45
    for hr in range(43, 51):
        if str(ws.cell(hr, 1).value).strip() == 'ID':
            detail_header_row = hr
            break
    detail_header = [ws.cell(detail_header_row, c).value for c in range(1, 24)]
    details = {}
    for r in range(detail_header_row + 1, ws.max_row + 1):
        row = {}
        for c, h in enumerate(detail_header, 1):
            if h:
                val = ws.cell(r, c).value
                row[h] = val.strip() if isinstance(val, str) else val
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
        # Find current period: tightest fit (start <= TODAY <= end with smallest end date)
        best_end = None
        has_current_period = False
        for r in rows_sorted:
            start = to_date(r.get('Start Date'))
            end = to_date(r.get('End Date'))
            if start and end and start <= TODAY <= end:
                has_current_period = True
                if best_end is None or end < best_end:
                    best_end = end
                    current_row = r

        # Future tenant: no period contains TODAY and the earliest start is in the future
        earliest_start = None
        for r in rows_sorted:
            s = to_date(r.get('Start Date'))
            if s and (earliest_start is None or s < earliest_start):
                earliest_start = s
        is_future = (not has_current_period) and (earliest_start is not None) and (earliest_start > TODAY)
        # For a future tenant, anchor display to the first (commencement) period
        if is_future:
            current_row = rows_sorted[0]
            commence_date = earliest_start
            # Use first period that actually carries rent (>0) for the display figures,
            # so a $0 free-rent first month doesn't make the row look empty.
            for r in rows_sorted:
                try:
                    m = float(r.get('Monthly', 0) or 0)
                except (TypeError, ValueError):
                    m = 0
                if m > 0:
                    current_row = r
                    break
        else:
            commence_date = None
        # Find next row: the period starting right after current_row's end
        current_end = to_date(current_row.get('End Date'))
        for r in rows_sorted:
            start = to_date(r.get('Start Date'))
            end = to_date(r.get('End Date'))
            if end and current_end and end > current_end:
                if start and start <= TODAY:
                    # This is a future year with start=commencement — use end date ordering
                    if not next_row or end < to_date(next_row.get('End Date')):
                        next_row = r
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

        # Find first option expiration date (if exercised, when would it expire)
        first_option_exp = None
        option_rows = [r for r in rows if r.get('Term', '').startswith('Option')]
        if option_rows:
            # Sort by Term (Option_1, Option_2, etc.) and get the first one's end date
            option_rows_sorted = sorted(option_rows, key=lambda r: r.get('Term', ''))
            first_opt = option_rows_sorted[0]
            opt_end = to_date(first_opt.get('End Date'))
            if opt_end:
                first_option_exp = opt_end

        try:
            gross_mod = float(current_row.get('Gross Mod', 0) or 0)
        except (TypeError, ValueError):
            gross_mod = 0
        try:
            monthly = float(current_row.get('Monthly', 0) or 0) + gross_mod
        except (TypeError, ValueError):
            monthly = gross_mod
        try:
            annual = float(current_row.get('Annual', 0) or 0) + (gross_mod * 12)
        except (TypeError, ValueError):
            annual = gross_mod * 12
        try:
            sqft = float(first.get('Sqft', 0) or 0)
        except (TypeError, ValueError):
            sqft = 0
        psf = (annual / sqft) if sqft and sqft > 1 else 0
        try:
            sec_dep = float(first.get('Sec Dep', 0) or 0)
        except (TypeError, ValueError):
            sec_dep = 0

        # Exp Date = last original period end date
        exp_dt = last_orig_end
        exp_str = exp_dt.strftime('%m/%d/%Y') if isinstance(exp_dt, (datetime, date)) else 'MTM'

        # Next Anniversary = end date of current period
        current_end = to_date(current_row.get('End Date'))
        next_anniv = current_end if current_end and current_end > TODAY else None

        # Next rent = from the period AFTER current
        next_gross_mod = (next_row.get('Gross Mod', 0) or 0) if next_row else 0
        next_monthly = ((next_row.get('Monthly', 0) or 0) + next_gross_mod) if next_row else None
        delta_monthly = (next_monthly - monthly) if (next_row and next_monthly is not None) else None
        anniv_months = None
        if next_anniv:
            diff_days = (next_anniv - TODAY).days
            anniv_months = max(0, round(diff_days / 30.44))

        # Find nearest future cancel date
        cancel_date = None
        for r in rows:
            cd = to_date(r.get('Cancel Date'))
            if cd and cd > TODAY:
                if cancel_date is None or cd < cancel_date:
                    cancel_date = cd
        cancel_str = cancel_date.strftime('%m/%d/%Y') if cancel_date else '—'

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
            'First_Option_Exp': first_option_exp.strftime('%m/%d/%Y') if first_option_exp else '-',
            'Exp Date': exp_str,
            'Cancel Date': cancel_str,
            'Escalation': current_row.get('Escalation', 0) or 0,
            'Sec Dep': round(sec_dep, 2),
            'Next Anniv': next_anniv,
            'Anniv_Months': anniv_months,
            'Next Monthly': round(next_monthly, 2) if next_monthly is not None else None,
            'Delta Monthly': round(delta_monthly, 2) if delta_monthly is not None else None,
            'Future': is_future,
            'Commence': commence_date.strftime('%m/%d/%Y') if commence_date else '',
            'Commence_Date': commence_date,
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


COI_RAW_BASE = "https://raw.githubusercontent.com/asollog61/msp-dashboard/master/data/coi"


def build_coi_url(building_name, filename):
    """Build a clickable GitHub raw URL for a COI PDF (URL-encoded)."""
    from urllib.parse import quote
    if not filename:
        return ''
    return f"{COI_RAW_BASE}/{quote(building_name)}/{quote(filename)}"


@st.cache_data(ttl=3600)
def scan_coi_files():
    """Returns (coi_data, building_coi_data, pm_coi_data): tenant certs, building-level certs, and PM certs per building."""
    coi_data = {}
    building_coi_data = {}
    pm_coi_data = {}
    # Check for COIs bundled in data/coi/{building}/ (Streamlit Cloud) or ACTIVE_PROPERTIES (local)
    bundled_coi = Path(os.path.dirname(__file__)) / "data" / "coi"
    for building_name, mapping in BUILDING_MAP.items():
        cert_dir = None
        # Priority 1: bundled COI folder (works on Streamlit Cloud)
        candidate = bundled_coi / building_name
        if candidate.exists():
            cert_dir = candidate
        # Priority 2: ACTIVE_PROPERTIES local path
        elif ACTIVE_PROPS_ROOT:
            cert_dir = ACTIVE_PROPS_ROOT / mapping["dest_folder"] / mapping["share"] / "Certificates of Insurance"
            if not cert_dir.exists():
                parent = ACTIVE_PROPS_ROOT / mapping["dest_folder"] / mapping["share"]
                cert_dir = None
                if parent.exists():
                    for d in parent.iterdir():
                        if d.is_dir() and 'certificate' in d.name.lower() and 'insurance' in d.name.lower():
                            cert_dir = d
                            break
        certs = []
        building_certs = []
        pm_certs = []
        if cert_dir and cert_dir.exists():
            for f in sorted(cert_dir.iterdir()):
                if not f.is_file() or f.suffix.lower() != '.pdf':
                    continue
                fname_lower = f.name.lower()
                # Building-level COI: filename contains _building (case-insensitive)
                is_building = '_building' in fname_lower
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
                # PM COI: insured name contains "Proventus" (property manager)
                is_pm = insured_name and 'proventus' in insured_name.lower()
                rec = {'filename': f.name, 'exp_date': exp_date, 'insured_name': insured_name}
                if is_pm:
                    pm_certs.append(rec)
                elif is_building:
                    building_certs.append(rec)
                else:
                    certs.append(rec)
        coi_data[building_name] = certs
        building_coi_data[building_name] = building_certs
        pm_coi_data[building_name] = pm_certs
    return coi_data, building_coi_data, pm_coi_data


# --- YARDI PDF PARSING ---
def normalize_space(space_str, building=None):
    """Normalize Yardi space format to dashboard format.
    For 114 Central: 102->2, 104->4, 106->6, 108->8 (first floor even units)
    For 2-X format: 2-1->201, 2-2->202, etc.
    For other formats: keep as-is (101, 201, 1286, A-1, etc.)
    """
    s = str(space_str).strip()
    
    # Handle 2-1 format (second floor units)
    if '-' in s:
        parts = s.split('-')
        if len(parts) == 2 and parts[0].isdigit() and parts[1].isdigit():
            return f"{parts[0]}{parts[1].zfill(2)}"
    
    # Keep 3-digit units as-is (102, 104, 106, 108, 201, etc.)
    # Previously stripped leading "10" but spreadsheet now uses full 3-digit format
    
    # For other formats (101, 201, 1286, A-1, O-1, etc.), keep as-is
    return s


@st.cache_data(ttl=3600)
def _get_yardi_dir():
    """Return the Yardi PDF directory path, or None."""
    for p in [
        os.path.join(os.path.dirname(__file__), "data", "Yardi"),
        "/home/node/OpenClaw/Share Jason/General Procedures/msp-dashboard-src/data/Yardi/",
    ]:
        if os.path.exists(p):
            return Path(p)
    return None


YARDI_FILENAME_MAP = {
    "1280-springfield": "1280 Springfield",
    "15-south": "15 South",
    "36-south": "36 South",
    "114-central": "114 Central",
}

MONTH_ORDER = {
    'jan': 1, 'feb': 2, 'mar': 3, 'apr': 4, 'may': 5, 'jun': 6,
    'jul': 7, 'aug': 8, 'sep': 9, 'oct': 10, 'nov': 11, 'dec': 12,
}


def _filter_latest_per_building(pdfs):
    """Given a list of PDF paths, return only the most recent per building.
    Filenames start with Mon-YYYY (e.g., May-2026_...). Returns list of Paths."""
    building_best = {}  # building_key -> (sort_key, path)
    for pdf_path in pdfs:
        fname = pdf_path.stem.lower()
        # Identify building
        building_key = None
        for key in YARDI_FILENAME_MAP:
            if key in fname:
                building_key = key
                break
        if not building_key:
            continue
        # Parse date from filename: Mon-YYYY
        match = re.match(r'^([a-z]{3})-(\d{4})', fname)
        if match:
            mon_str, year_str = match.group(1), match.group(2)
            sort_key = int(year_str) * 100 + MONTH_ORDER.get(mon_str, 0)
        else:
            sort_key = 0
        prev = building_best.get(building_key)
        if not prev or sort_key > prev[0]:
            building_best[building_key] = (sort_key, pdf_path)
    return [v[1] for v in building_best.values()]


def parse_yardi_deposit_activity(latest_only=True):
    """Parse Security Deposit Activity pages from Yardi PDFs.
    Returns dict keyed by 'Building|Space' with Deposits On Hand amount.
    If latest_only=True, only parses the most recent report per building."""
    if not HAS_PYMUPDF:
        return {}

    yardi_dir = _get_yardi_dir()
    if not yardi_dir or not yardi_dir.exists():
        return {}

    filename_map = YARDI_FILENAME_MAP

    deposit_data = {}

    all_pdfs = sorted(yardi_dir.glob("*.pdf"))
    pdf_list = _filter_latest_per_building(all_pdfs) if latest_only else all_pdfs

    for pdf_path in pdf_list:
        filename_lower = pdf_path.stem.lower()
        building = None
        for key, name in filename_map.items():
            if key in filename_lower:
                building = name
                break
        if not building:
            continue

        try:
            doc = fitz.open(str(pdf_path))
            for page in doc:
                text = page.get_text()
                if 'Security Deposit Activity' not in text:
                    continue

                lines = [l.strip() for l in text.split('\n') if l.strip()]
                i_line = 0
                while i_line < len(lines):
                    line = lines[i_line]
                    if '(Current)' in line or '(Past)' in line:
                        tenant_name = line
                        # Collect numeric values following the tenant line
                        nums = []
                        j = i_line + 1
                        while j < len(lines) and len(nums) < 7:
                            try:
                                val = float(lines[j].replace(',', ''))
                                nums.append(val)
                                j += 1
                            except (ValueError, AttributeError):
                                break

                        # Find unit number - first try inline (unit + tenant on same line)
                        unit = ''
                        inline_match = re.match(r'^([A-Za-z0-9_-]+)\s+', line)
                        if inline_match:
                            candidate = inline_match.group(1)
                            # Verify it's a unit-like token (has digits or is short)
                            if len(candidate) <= 8 and any(c.isdigit() for c in candidate):
                                unit = candidate

                        # Fallback: look backwards for short alphanumeric token on prior line
                        if not unit:
                            skip_words = {'avenue', 'st', 'st.', 'south', 'central', 'ave', 'ave.',
                                          'd', '36', '15', '1280', '114', 'springfiel', 'property',
                                          'unit', 'tenant', 'prior', 'current', 'deposits', 'total',
                                          'on', 'hand', 'billed', 'receipts', 'forfeited', 'deposit'}
                            for k in range(i_line - 1, max(i_line - 6, 0), -1):
                                candidate = lines[k].strip()
                                if (re.match(r'^[A-Za-z0-9_-]+$', candidate) and
                                        len(candidate) <= 8 and
                                        candidate.lower() not in skip_words):
                                    unit = candidate
                                    break

                        if unit and len(nums) >= 5:
                            deposits_on_hand = nums[4]
                            normalized = normalize_space(unit, building)
                            key = f"{building}|{normalized}"
                            # Only use (Current) entries, or add to existing if (Past)
                            if '(Current)' in tenant_name:
                                deposit_data[key] = deposits_on_hand
                            elif key not in deposit_data:
                                deposit_data[key] = deposits_on_hand

                        i_line = j
                    else:
                        i_line += 1

            doc.close()
        except Exception:
            pass

    return deposit_data


@st.cache_data(ttl=300)
def parse_yardi_rent_rolls(latest_only=True):
    """Parse Yardi monthly statement PDFs from data/Yardi/ folder.
    Returns dict keyed by "Building|Space" with {monthly, cam, expiration, tenant}.
    If latest_only=True, only parses the most recent report per building.
    """
    if not HAS_PYMUPDF:
        return {}
    
    yardi_dir = _get_yardi_dir()
    if not yardi_dir or not yardi_dir.exists():
        return {}
    
    filename_map = YARDI_FILENAME_MAP
    
    yardi_data = {}
    
    all_pdfs = sorted(yardi_dir.glob("*.pdf"))
    pdf_list = _filter_latest_per_building(all_pdfs) if latest_only else all_pdfs

    for pdf_path in pdf_list:
        # Extract building name from filename
        filename_lower = pdf_path.stem.lower()
        building = None
        for key, name in filename_map.items():
            if key in filename_lower:
                building = name
                break
        
        if not building:
            continue
        
        try:
            doc = fitz.open(str(pdf_path))
            
            # Page 3 (0-indexed page 2) is the Monthly Rent Roll
            if len(doc) < 3:
                doc.close()
                continue
            
            page = doc[2]
            
            # Check if this is the Monthly Rent Roll page
            # Look for key headers: "Unit", "Tenant Name", "Actual Rent"
            text = page.get_text()
            if not all(keyword in text for keyword in ["Unit", "Tenant Name", "Actual Rent"]):
                doc.close()
                continue
            
            # Extract text blocks - each tenant is a block
            blocks = page.get_text("blocks")
            
            for block in blocks:
                if block[6] != 0:  # Not a text block
                    continue
                
                content = block[4].strip()
                lines = [l.strip() for l in content.split('\n') if l.strip()]
                
                if len(lines) < 5:
                    continue
                
                # First line should be unit number
                unit = lines[0]
                
                # Skip header rows, totals, and summary sections
                if unit.upper() in ('UNIT', 'TOTAL', 'SUMMARY', 'CURRENT/NOTICE/VACANT', 'FUTURE', 'OCCUPIED', 'TOTALS:', 'GROUPS'):
                    continue
                
                # Skip if doesn't look like a unit number (allow named units like EXTERIOR, ROOF, PARKING)
                NAMED_UNITS = {'EXTERIOR', 'ROOF', 'PARKING', 'ATM', 'PAD', 'BASEMENT', 'STORAGE', 'SIGN', 'BILLBOARD', 'EASEMENT'}
                if not any(c.isdigit() for c in unit) and unit.upper() not in NAMED_UNITS:
                    continue
                
                # Parse the tenant line
                # Expected format: Unit, SqFt, Tenant Name (may be multi-line), Monthly, ..., CAM, ..., dates
                
                # Second item should be SqFt (numeric with .00)
                if len(lines) < 3:
                    continue
                
                # Verify second item looks like a number (SqFt)
                sqft_str = lines[1]
                try:
                    sqft = float(sqft_str.replace(',', ''))
                except (ValueError, AttributeError):
                    continue
                
                # Extract tenant name (may span multiple lines for d/b/a)
                # Start from index 2, stop when we hit monthly rent (a larger number)
                tenant_lines = []
                idx = 2
                while idx < len(lines):
                    line = lines[idx]
                    # Check if this looks like the monthly rent (larger number, usually 3+ digits before decimal)
                    if re.match(r'^[\d,]+\.\d{2}$', line):
                        try:
                            val = float(line.replace(',', ''))
                            # If it's > 100, likely monthly rent; if < 100, might be part of tenant name
                            if val >= 100 or (idx > 2 and val > 0):
                                break
                        except (ValueError, AttributeError):
                            pass
                    tenant_lines.append(line)
                    idx += 1
                
                tenant_name = ' '.join(tenant_lines).strip()
                
                # Skip VACANT units
                if tenant_name.upper() == 'VACANT':
                    continue
                
                # Skip if tenant name is empty or looks like garbage
                if not tenant_name or len(tenant_name) < 2:
                    continue
                
                # Now extract monthly rent (next item after tenant name)
                if idx >= len(lines):
                    continue
                
                monthly_str = lines[idx]
                try:
                    monthly = float(monthly_str.replace(',', ''))
                    # Sanity check - monthly rent should be > 0 (allow named units like easements with $0)
                    if monthly <= 0 and unit.upper() not in NAMED_UNITS:
                        continue
                except (ValueError, AttributeError):
                    continue
                
                # Find CAM amount - typically the 5th numeric field after monthly rent
                # Format: Monthly, PSF, Tenant_Deposit, Other_Deposit, CAM, CAM_PSF, Misc, dates
                # Look for a reasonable CAM value (skip PSF which is small, skip deposits which might be 0)
                cam = 0.0
                numeric_fields = []
                for i in range(idx + 1, min(idx + 10, len(lines))):
                    try:
                        val = float(lines[i].replace(',', ''))
                        numeric_fields.append(val)
                    except (ValueError, AttributeError):
                        continue
                
                # CAM is typically at index 3 or 4 in numeric_fields (after PSF, 2 deposits)
                # Look for the first value > 10 that's reasonable relative to monthly rent
                for val in numeric_fields[2:6]:  # Skip first 2 (PSF, first deposit), check next 4
                    if val > 10 and val < monthly * 3:
                        cam = val
                        break
                
                # Find lease expiration date (format: MM/DD/YYYY)
                exp_date = None
                for i in range(idx + 1, len(lines)):
                    match = re.search(r'(\d{2}/\d{2}/\d{4})', lines[i])
                    if match:
                        date_str = match.group(1)
                        try:
                            exp_date = datetime.strptime(date_str, '%m/%d/%Y').date()
                            # Take the LAST date found (lease expiration, not move-in)
                        except ValueError:
                            pass
                
                # Normalize space number
                normalized_space = normalize_space(unit, building)
                key = f"{building}|{normalized_space}"
                
                # Extract security deposit (numeric_fields[1] = Tenant Deposit, after PSF)
                deposit = 0.0
                if len(numeric_fields) >= 2:
                    deposit = numeric_fields[1]

                yardi_data[key] = {
                    'monthly': monthly,
                    'cam': cam,
                    'deposit': deposit,
                    'expiration': exp_date,
                    'tenant': tenant_name,
                    'raw_unit': unit,
                }
            
            doc.close()
            
        except Exception as e:
            # Silently skip problematic PDFs
            pass
    
    return yardi_data


def compute_yardi_diffs(tenants, yardi_data):
    """Compare dashboard tenants with Yardi data.
    Returns dict keyed by "Building|Space" with comma-separated diff string.
    """
    diffs = {}
    
    for tenant in tenants:
        building = tenant.get('Building', '').strip()
        space = str(tenant.get('Space', '')).strip()
        
        if not building or not space:
            continue
        
        key = f"{building}|{space}"
        
        # Look up in Yardi data
        yardi_entry = yardi_data.get(key)
        
        if not yardi_entry:
            # Try alternate space formats
            # If space is "2", try "102"
            if space.isdigit() and len(space) == 1:
                alt_key = f"{building}|10{space}"
                yardi_entry = yardi_data.get(alt_key)
            # If space is "201", try "2-1"
            elif space.isdigit() and len(space) == 3:
                first = space[0]
                rest = space[1:]
                alt_key = f"{building}|{first}-{rest}"
                yardi_entry = yardi_data.get(alt_key)
        
        if not yardi_entry:
            diffs[key] = "Not in Yardi"
            continue
        
        # Compare monthly rent, CAM, and expiration
        diff_parts = []
        
        # Monthly rent
        dash_monthly = tenant.get('Monthly', 0) or 0
        yardi_monthly = yardi_entry.get('monthly', 0) or 0
        if abs(dash_monthly - yardi_monthly) > 0.01:
            diff_parts.append(f"Monthly: ${dash_monthly:,.0f} vs ${yardi_monthly:,.0f}")
        
        # CAM amount
        # Dashboard stores CAM as percentage of building expenses, Yardi stores actual monthly CAM
        # For comparison, we need building expenses
        # For now, just show if there's a CAM difference
        # We'll calculate actual CAM reimbursement in render function
        
        # Expiration date
        dash_exp_str = tenant.get('Exp Date', '').strip()
        yardi_exp = yardi_entry.get('expiration')
        
        if dash_exp_str and dash_exp_str != 'MTM' and yardi_exp:
            try:
                dash_exp = datetime.strptime(dash_exp_str, '%m/%d/%Y').date()
                if dash_exp != yardi_exp:
                    diff_parts.append(f"Exp: {dash_exp.strftime('%m/%d/%Y')} vs {yardi_exp.strftime('%m/%d/%Y')}")
            except (ValueError, AttributeError):
                pass
        
        diffs[key] = ', '.join(diff_parts) if diff_parts else ''
    
    return diffs


# --- SOP EXTRACTION ---
@st.cache_data(ttl=86400)
def _extract_sop_content(pdf_path, cache_token):
    if not pdf_path or not HAS_PYMUPDF:
        return [], {}
    doc = fitz.open(pdf_path)

    # Extract tables from all pages
    all_tables = []
    for page in doc:
        try:
            tabs = page.find_tables()
            for tab in tabs:
                cells = tab.extract()
                if cells and len(cells) > 1:
                    all_tables.append({'page': page.number, 'y': tab.bbox[1], 'rows': cells})
        except Exception:
            pass

    # Build table HTML lookup keyed by header signature
    table_map = {}
    for tbl in all_tables:
        rows = tbl['rows']
        header = rows[0]
        body = rows[1:]
        th = ''.join(f"<th>{(c or '').strip()}</th>" for c in header)
        tr_list = []
        for row in body:
            td = ''.join(f"<td>{(c or '').strip()}</td>" for c in row)
            tr_list.append(f"<tr>{td}</tr>")
        html = f"<table class='sop-table'><thead><tr>{th}</tr></thead><tbody>{''.join(tr_list)}</tbody></table>"
        # Key by first header cell text for matching
        key = (header[0] or '').strip().lower()
        if key:
            table_map[key] = {'html': html, 'cell_texts': set()}
            for row in rows:
                for cell in row:
                    if cell and len(cell.strip()) > 2:
                        table_map[key]['cell_texts'].add(cell.strip())

    # Standard text extraction
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
    return sections, table_map


def _match_table_for_section(lines, table_map):
    """Check if any extracted table belongs in this section's content. Returns (table_key, set of lines to remove)."""
    matches = []
    for key, tbl in table_map.items():
        cell_texts = tbl['cell_texts']
        matched_indices = set()
        for i, line in enumerate(lines):
            stripped = line.strip()
            if stripped in cell_texts or any(ct in stripped for ct in cell_texts if len(ct) > 5):
                matched_indices.add(i)
        # If we matched enough cells, this table belongs here
        if len(matched_indices) >= 2:
            matches.append((key, matched_indices))
    return matches


def extract_sop_content():
    if not SOP_PDF:
        return [], {}
    try:
        cache_token = os.path.getmtime(SOP_PDF)
    except OSError:
        cache_token = datetime.now().timestamp()
    return _extract_sop_content(SOP_PDF, cache_token)


# =====================
# RENDER TABS
# =====================

def render_sop_tab():
    if not SOP_PDF:
        st.warning("SOP Manual PDF not found. Place Marion_St_SOP_Manual.pdf in the data/ folder.")
        return

    st.markdown("### 📋 Marion St Properties — Standard Operating Procedures")

    if st.session_state.get('mobile_view', False) and HAS_PYMUPDF:
        # Mobile: page-by-page image view with navigation
        import base64
        doc = fitz.open(SOP_PDF)
        total_pages = len(doc)

        if 'sop_page' not in st.session_state:
            st.session_state.sop_page = 0

        page_num = st.session_state.sop_page
        page_num = max(0, min(page_num, total_pages - 1))

        # Render current page as image
        page = doc[page_num]
        mat = fitz.Matrix(1.25, 1.25)  # 1.25x for mobile
        pix = page.get_pixmap(matrix=mat)
        img_bytes = pix.tobytes("png")
        b64_img = base64.b64encode(img_bytes).decode()
        doc.close()

        # Navigation buttons on top
        col_prev, col_num, col_next = st.columns([1, 2, 1])
        if col_prev.button("⬅️ Prev", disabled=(page_num == 0), key="sop_prev"):
            st.session_state.sop_page = page_num - 1
            st.rerun()
        col_num.markdown(f"<div style='text-align:center; padding-top:8px;'>{page_num + 1} / {total_pages}</div>", unsafe_allow_html=True)
        if col_next.button("Next ➡️", disabled=(page_num >= total_pages - 1), key="sop_next"):
            st.session_state.sop_page = page_num + 1
            st.rerun()

        st.markdown(f'<img src="data:image/png;base64,{b64_img}" style="width:100%; border-radius:6px; border:1px solid rgba(255,255,255,0.1);">', unsafe_allow_html=True)
    else:
        # Desktop: page-by-page image view (same as mobile, higher res)
        import base64
        if HAS_PYMUPDF:
            doc = fitz.open(SOP_PDF)
            total_pages = len(doc)

            if 'sop_page' not in st.session_state:
                st.session_state.sop_page = 0

            page_num = st.session_state.sop_page
            page_num = max(0, min(page_num, total_pages - 1))

            page = doc[page_num]
            mat = fitz.Matrix(2.0, 2.0)
            pix = page.get_pixmap(matrix=mat)
            # Crop whitespace/margins from the rendered page
            img_bytes = pix.tobytes("png")
            # Use fitz to detect content bbox and crop
            try:
                content_rect = page.get_text("dict", flags=0).get("width", 0)
                # Simpler: use CropBox or trim via pixmap
                # Get the actual content bounding box
                clip = page.rect
                # Try trimming margins: typical PDF has ~72pt margins on letter (612x792)
                pw, ph = page.rect.width, page.rect.height
                margin_x = pw * 0.06  # ~6% margins
                margin_top = ph * 0.04
                margin_bot = ph * 0.04
                trimmed = fitz.Rect(margin_x, margin_top, pw - margin_x, ph - margin_bot)
                pix2 = page.get_pixmap(matrix=mat, clip=trimmed)
                img_bytes = pix2.tobytes("png")
            except Exception:
                pass  # Fall back to full page render
            b64_img = base64.b64encode(img_bytes).decode()
            doc.close()

            col_prev, col_num, col_next, col_dl = st.columns([1, 2, 1, 2])
            if col_prev.button("⬅️ Prev", disabled=(page_num == 0), key="sop_prev_d"):
                st.session_state.sop_page = page_num - 1
                st.rerun()
            col_num.markdown(f"<div style='text-align:center; padding-top:8px;'>{page_num + 1} / {total_pages}</div>", unsafe_allow_html=True)
            if col_next.button("Next ➡️", disabled=(page_num >= total_pages - 1), key="sop_next_d"):
                st.session_state.sop_page = page_num + 1
                st.rerun()
            with open(SOP_PDF, "rb") as f:
                col_dl.download_button("⬇️ Download PDF", data=f.read(), file_name="Marion_St_SOP_Manual.pdf", mime="application/pdf", key="sop_dl")

            st.markdown(f'<img src="data:image/png;base64,{b64_img}" style="width:100%; max-width:900px; border-radius:6px; border:1px solid rgba(255,255,255,0.1); display:block; margin:0 auto;">', unsafe_allow_html=True)
        else:
            with open(SOP_PDF, "rb") as f:
                pdf_bytes = f.read()
            b64 = base64.b64encode(pdf_bytes).decode()
            st.markdown(f'<iframe src="data:application/pdf;base64,{b64}" width="100%" height="1600px" style="border:1px solid rgba(255,255,255,0.1);border-radius:6px;"></iframe>', unsafe_allow_html=True)


def render_tenancy_tab():
    tenants, details, summaries = load_tenancy()
    if not tenants:
        st.warning("MSP Tenancy.xlsx not found.")
        return

    expense_cfg = read_expense_config()

    # Get vacant spaces to flag in tenancy view
    vacant_keys, _vacant_meta = build_vacancy_lookup(tenants, include_auto=False)

    # Parse Yardi data and compute diffs
    yardi_data = parse_yardi_rent_rolls()
    yardi_diffs = compute_yardi_diffs(tenants, yardi_data)

    # Separate future (not-yet-commenced) tenants from current ones
    future_tenants = [t for t in tenants if t.get('Future')]
    # Summary metrics (current tenants only)
    active = [t for t in tenants if not t.get('Future')]
    # Add vacancy status and Yardi diff
    for t in active:
        key = f"{t['Building']}|{t['Space']}"
        t['Status'] = '🔴 VACANT' if key in vacant_keys else ''
        t['Yardi Diff'] = yardi_diffs.get(key, '')
    total_sf = sum(t['SF'] for t in active if t['SF'] > 1)
    total_annual = sum(t['Annual'] for t in active)
    total_monthly = sum(t['Monthly'] for t in active)
    mtm_count = sum(1 for t in active if t.get('TTE_Label') == 'MTM')

    # Calculate total CAM reimbursement across all buildings
    total_cam_reimb = 0
    for bldg_name in BUILDING_MAP.keys():
        bldg_expense = float(expense_cfg.get(bldg_name, 0) or 0)
        bldg_tenants = [t for t in active if t['Building'] == bldg_name]
        for t in bldg_tenants:
            if t.get('Is_NNN') and isinstance(t.get('CAM_Pct'), (int, float)):
                total_cam_reimb += bldg_expense * t['CAM_Pct']

    total_expenses = sum(float(expense_cfg.get(b, 0) or 0) for b in BUILDING_MAP.keys())
    total_gross_rev = total_annual + total_cam_reimb
    total_noi = total_gross_rev - total_expenses
    wavg_net_psf = total_annual / total_sf if total_sf > 0 else 0
    wavg_gross_psf = total_gross_rev / total_sf if total_sf > 0 else 0

    c1, c2, c3, c4, c5, c6 = st.columns(6)
    c1.metric("Portfolio SF", f"{total_sf:,.0f}")
    c2.metric("Gross Revenue", f"${total_gross_rev:,.0f}")
    c3.metric("Expenses", f"${total_expenses:,.0f}")
    c4.metric("NOI", f"${total_noi:,.0f}")
    c5.metric("Wtd Avg Gross $/SF", f"${wavg_gross_psf:,.2f}")
    c6.metric("Wtd Avg Net $/SF", f"${wavg_net_psf:,.2f}")

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
            _bldg_header_placeholder = st.empty()

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

            display_cols = ['Space', 'Tenant', 'Type', 'SF', 'Lease', 'Monthly', 'Annual', 'Gross_Annual', 'PSF', 'Gross_PSF', 'CAM_Pct', 'CAM_Reimb', 'TTE_Months', 'Exp Date', 'Cancel Date', 'Options', 'First_Option_Exp', 'Escalation', 'Next Anniv', 'Anniv_Months', 'Next Monthly', 'Delta Monthly', 'Status', 'Yardi Diff', 'TTE_Label', 'Is_NNN']

            # Calculate building totals
            b_monthly = df['Monthly'].sum()
            b_annual = df['Annual'].sum()
            b_sf = df.loc[df['SF'] > 1, 'SF'].sum()
            b_wavg_psf = b_annual / b_sf if b_sf > 0 else 0
            b_gross_annual = df['Gross_Annual'].sum()
            b_cam_reimb = df['CAM_Reimb'].sum()
            b_noi = b_gross_annual - building_expense

            _bldg_header_placeholder.markdown(
                f'<div class="building-header"><strong>{current_bldg}</strong> — '
                f'{len(bldg_tenants)} tenants · '
                f'Gross Rev: ${b_gross_annual:,.0f} · '
                f'Expenses: ${building_expense:,.0f} · '
                f'NOI: ${b_noi:,.0f}</div>',
                unsafe_allow_html=True
            )

            # Create display dataframe
            display_df = df[display_cols].copy()
            display_df.rename(columns={
                'CAM_Pct': 'CAM %',
                'CAM_Reimb': 'CAM Reimb',
                'Gross_Annual': 'Gross Annual',
                'Gross_PSF': 'Gross PSF',
                'TTE_Months': 'MTE',
                'First_Option_Exp': '1st Opt Exp',
                'Next Anniv': 'Next Anniversary',
                'Anniv_Months': 'Anniv Δ',
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
            display_df['MTE'] = display_df['MTE'].fillna(0).astype(int)
            display_df['Anniv Δ'] = display_df['Anniv Δ'].apply(
                lambda x: int(x) if x is not None and not (isinstance(x, float) and pd.isna(x)) else None
            )
            display_df['Escalation'] = display_df['Escalation'].apply(lambda x: f"{x:.1%}" if x and x > 0 else '-')
            display_df['Monthly'] = display_df['Monthly'].apply(lambda x: f"${x:,.0f}")
            display_df['Annual'] = display_df['Annual'].apply(lambda x: f"${x:,.0f}")
            display_df['PSF'] = display_df['PSF'].apply(lambda x: f"${x:,.2f}")
            display_df['Space'] = display_df['Space'].astype(str)
            display_df['SF'] = display_df['SF'].apply(lambda x: f"{x:,.0f}" if isinstance(x, (int, float)) and x > 1 else '')
            display_df['Next Anniversary'] = display_df['Next Anniversary'].apply(
                lambda d: d.strftime('%m/%d/%Y') if isinstance(d, (datetime, date)) else ('-' if not d else str(d))
            )
            display_df['New Rent'] = display_df['New Rent'].apply(
                lambda x: f"${x:,.0f}" if x is not None and not (isinstance(x, float) and pd.isna(x)) else '$0'
            )
            display_df['Δ Monthly'] = display_df['Δ Monthly'].apply(
                lambda x: f"{'+' if x and x > 0 else ''}${x:,.0f}" if x is not None and not (isinstance(x, float) and pd.isna(x)) else '$0'
            )

            # Pinned totals row
            totals_row = [{
                'Space': '', 'Tenant': 'TOTAL', 'Type': '', 'SF': f"{b_sf:,.0f}",
                'Lease': '', 'Monthly': f"${b_monthly:,.0f}", 'Annual': f"${b_annual:,.0f}",
                'PSF': f"${b_wavg_psf:,.2f}", 'Gross Annual': f"${b_gross_annual:,.0f}", 'Gross PSF': '',
                'CAM %': '', 'CAM Reimb': f"${b_cam_reimb:,.0f}",
                'MTE': '', 'Exp Date': '', 'Cancel Date': '', 'Options': '', '1st Opt Exp': '', 'Escalation': '', 'Next Anniversary': '', 'Anniv Δ': '',
                'New Rent': '', 'Δ Monthly': '', 'Status': '', 'Yardi Diff': '', 'TTE Label': '', 'NNN': '',
            }]

            tte_formatter = JsCode("""
                function(params) {
                    if (params.value === null || params.value === undefined) { return '-'; }
                    return params.value.toString();
                }
            """)
            mte_cell_style = JsCode("""
                function(params) {
                    if (params.value === null || params.value === undefined || params.node.rowPinned) {
                        return {'text-align': 'right'};
                    }
                    var v = parseInt(params.value);
                    if (isNaN(v)) return {'text-align': 'right'};
                    var maxMte = 120;
                    var ratio = Math.min(v / maxMte, 1.0);
                    var r, g, b;
                    if (ratio <= 0.5) {
                        var t = ratio / 0.5;
                        r = Math.round(200 - t * 100);
                        g = Math.round(50 + t * 100);
                        b = Math.round(50);
                    } else {
                        var t = (ratio - 0.5) / 0.5;
                        r = Math.round(100 - t * 70);
                        g = Math.round(150 + t * 55);
                        b = Math.round(50 + t * 20);
                    }
                    return {
                        'background-color': 'rgba(' + r + ',' + g + ',' + b + ',0.35)',
                        'color': '#e6edf3',
                        'font-weight': '600',
                        'text-align': 'right'
                    };
                }
            """)
            anniv_formatter = JsCode("""
                function(params) {
                    if (params.value === null || params.value === undefined) { return '-'; }
                    return params.value.toString();
                }
            """)
            column_configs = {
                'Space': {'type': ['textColumn']},
                'MTE': {'type': ['numericColumn'], 'valueFormatter': tte_formatter, 'cellStyle': mte_cell_style},
                'Anniv Δ': {'type': ['numericColumn'], 'valueFormatter': anniv_formatter},
                'TTE Label': {'hide': True},
                'NNN': {'hide': True},
            }

            show_grid(display_df, key=f"tenancy_{current_bldg}", pinned_bottom=totals_row,
                      column_configs=column_configs, tab_key="tenancy")

    # Portfolio totals
    st.divider()
    port_sf = sum(t['SF'] for t in filtered if t['SF'] > 1 and t['Tenant'] != 'Easement')
    port_annual = sum(t['Annual'] for t in filtered if t['Tenant'] != 'Easement')
    port_cam_reimb = 0
    for bldg_name in BUILDING_MAP.keys():
        bldg_expense = float(expense_cfg.get(bldg_name, 0) or 0)
        for t in filtered:
            if t['Building'] == bldg_name and t.get('Is_NNN') and isinstance(t.get('CAM_Pct'), (int, float)):
                port_cam_reimb += bldg_expense * t['CAM_Pct']
    port_gross_rev = port_annual + port_cam_reimb
    port_expenses = sum(float(expense_cfg.get(b, 0) or 0) for b in BUILDING_MAP.keys())
    port_noi = port_gross_rev - port_expenses
    port_wavg_net_psf = port_annual / port_sf if port_sf > 0 else 0
    port_wavg_gross_psf = port_gross_rev / port_sf if port_sf > 0 else 0
    pc1, pc2, pc3, pc4, pc5, pc6 = st.columns(6)
    pc1.metric("Total SF", f"{port_sf:,.0f}")
    pc2.metric("Gross Revenue", f"${port_gross_rev:,.0f}")
    pc3.metric("Expenses", f"${port_expenses:,.0f}")
    pc4.metric("NOI", f"${port_noi:,.0f}")
    pc5.metric("Wtd Avg Gross $/SF", f"${port_wavg_gross_psf:,.2f}")
    pc6.metric("Wtd Avg Net $/SF", f"${port_wavg_net_psf:,.2f}")

    # --- FUTURE TENANTS (leases not yet commenced) ---
    fut = future_tenants if selected == 'All' else [t for t in future_tenants if t['Building'] == selected]
    if fut:
        import pandas as pd
        st.divider()
        st.markdown(f"### 🔜 Future Tenants ({len(fut)})")
        st.caption("Executed leases that have not yet commenced. Excluded from current portfolio totals above.")
        # Sort by commencement date (soonest first)
        fut_sorted = sorted(fut, key=lambda t: t.get('Commence_Date') or date.max)
        fdf = pd.DataFrame(fut_sorted)
        fdf['Space'] = fdf['Space'].astype(str)
        fdf['SF'] = fdf['SF'].apply(lambda x: f"{x:,.0f}" if isinstance(x, (int, float)) and x > 1 else '')
        fdf['Monthly'] = fdf['Monthly'].apply(lambda x: f"${x:,.0f}")
        fdf['Annual'] = fdf['Annual'].apply(lambda x: f"${x:,.0f}")
        fdf['PSF'] = fdf['PSF'].apply(lambda x: f"${x:,.2f}")
        fut_cols = ['Building', 'Space', 'Tenant', 'Type', 'SF', 'Lease', 'Commence', 'Monthly', 'Annual', 'PSF', 'Exp Date', 'Options', 'First_Option_Exp']
        fut_cols = [c for c in fut_cols if c in fdf.columns]
        fdisplay = fdf[fut_cols].copy()
        fdisplay.rename(columns={'Commence': 'Commences', 'First_Option_Exp': '1st Opt Exp'}, inplace=True)
        show_grid(fdisplay, key="future_tenants", tab_key="tenancy")

    render_column_config_editor(
        'tenancy',
        ['Space', 'Tenant', 'Type', 'SF', 'Lease', 'Monthly', 'Annual', 'Gross Annual', 'PSF', 'Gross PSF',
         'CAM %', 'CAM Reimb', 'MTE', 'Exp Date', 'Cancel Date', 'Options', '1st Opt Exp', 'Escalation',
         'Next Anniversary', 'Anniv Δ', 'New Rent', 'Δ Monthly', 'Status', 'Yardi Diff']
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
            # Never flag easements as vacant
            if 'easement' in (t.get('Tenant') or '').lower():
                continue
            auto_flag = (t.get('Monthly', 0) <= 0.01 and t.get('Annual', 0) <= 0.01) or t.get('Status') == '🔴 VACANT'
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
    active = [t for t in tenants]

    # Get vacant lookup (manual only)
    vacant_keys, vacant_meta = build_vacancy_lookup(active, include_auto=False)
    vacant_records = get_vacant_spaces()
    manual_keys = set(f"{v.get('Building', '')}|{v.get('Space', '')}" for v in vacant_records)

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

    if vacant_records:
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
                'Space': str(v.get('Space', '')),
                'Last Tenant': v.get('Tenant', ''),
                'SF': v.get('SF', ''),
                'Vacant Since': vac_date_str,
                'Days Vacant': days_vacant,
                'Notes': v.get('Notes', ''),
            })

        if vacant_display:
            vdf = pd.DataFrame(vacant_display)
            vdf['Space'] = vdf['Space'].astype(str)
            show_grid(vdf, key="vacant_spaces", tab_key="vacancy")

            # Remove buttons for manual entries
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

    # At risk — exclude spaces already shown in Currently Vacant
    st.markdown("### 🟡 At Risk — Expiring Within 12 Months")
    at_risk = [t for t in active if (t.get('TTE_Label') == 'MTM' or (0 < t['TTE_Days'] <= 365))
               and f"{t['Building']}|{t['Space']}" not in manual_keys]
    if at_risk:
        import pandas as pd
        risk_df = pd.DataFrame(at_risk)[['Building', 'Space', 'Tenant', 'SF', 'Monthly', 'TTE_Months', 'Exp Date']]
        risk_df.rename(columns={'TTE_Months': 'MTE'}, inplace=True)
        risk_df['Monthly'] = risk_df['Monthly'].apply(lambda x: f"${x:,.0f}")
        show_grid(
            risk_df,
            key="at_risk",
            column_configs={'MTE': {'type': ['numericColumn']}},
            tab_key="vacancy"
        )
    else:
        st.success("✅ No tenants expiring within 12 months.")

    # all_spaces used by Marketing section below
    all_spaces = sorted(set(f"{t['Building']} #{t['Space']}" for t in active))

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

    render_column_config_editor('vacancy', ['Building', 'Space', 'Tenant', 'SF', 'Monthly', 'MTE', 'Exp Date', 'Last Tenant', 'Vacant Since', 'Days Vacant', 'Notes'])
    render_column_config_editor('marketing', list(mdf.columns))


# --- COI EMAIL REPORT ---
EMAIL_TEAM_RECIPIENTS = [
    "asollog@gmail.com",
    "richard.b.angel@gmail.com",
    "jason.forster@proventusproperties.com",
]


def build_coi_report_data():
    """Build a unified list of every COI line item (building, PM, tenant) across all
    properties with a normalized status and days_left, for use in the email report
    and the urgent / expiring dashboard sections.

    Returns (items, summary) where items is a list of dicts:
      {building, code, category, entity, unit, coi, expiration (date|None),
       exp_str, days_left (int|None), status}
    """
    tenants, _, summaries = load_tenancy()
    coi_data, building_coi_data, pm_coi_data = scan_coi_files()
    today_dt = datetime.now()
    vacant_keys, vacant_meta = build_vacancy_lookup(tenants, include_auto=False)

    items = []
    summary = {'total': 0, 'covered': 0, 'expired': 0, 'missing': 0, 'expiring_soon': 0}

    def _classify(exp):
        """Return (status, coi_label, days_left) for an expiration date."""
        if not exp:
            return ('No date', '✅ YES', None)
        dl = (exp - today_dt).days
        if exp < today_dt:
            return ('EXPIRED', '❌ EXP', dl)
        if dl < 90:
            return (f'{dl}d left', '✅ YES', dl)
        return ('Active', '✅ YES', dl)

    for building_name in BUILDING_MAP:
        code = BUILDING_MAP[building_name]['code']

        # Building-level certificate
        b_building_certs = building_coi_data.get(building_name, [])
        if b_building_certs:
            for cert in b_building_certs:
                exp = cert.get('exp_date')
                status, coi_label, dl = _classify(exp)
                items.append({
                    'building': building_name, 'code': code, 'category': 'Building',
                    'entity': cert.get('insured_name') or building_name, 'unit': '',
                    'coi': coi_label, 'expiration': exp,
                    'exp_str': exp.strftime('%m/%d/%Y') if exp else 'Unknown',
                    'days_left': dl, 'status': status,
                })
        else:
            items.append({
                'building': building_name, 'code': code, 'category': 'Building',
                'entity': building_name, 'unit': '', 'coi': '❌ NO',
                'expiration': None, 'exp_str': '—', 'days_left': None,
                'status': 'MISSING',
            })

        # Property Manager certificate
        b_pm_certs = pm_coi_data.get(building_name, [])
        if b_pm_certs:
            for cert in b_pm_certs:
                exp = cert.get('exp_date')
                status, coi_label, dl = _classify(exp)
                items.append({
                    'building': building_name, 'code': code, 'category': 'Property Manager',
                    'entity': cert.get('insured_name') or 'Property Manager', 'unit': '',
                    'coi': coi_label, 'expiration': exp,
                    'exp_str': exp.strftime('%m/%d/%Y') if exp else 'Unknown',
                    'days_left': dl, 'status': status,
                })
        else:
            items.append({
                'building': building_name, 'code': code, 'category': 'Property Manager',
                'entity': 'Property Manager', 'unit': '', 'coi': '❌ NO',
                'expiration': None, 'exp_str': '—', 'days_left': None,
                'status': 'MISSING',
            })

    # Tenant certificates
    for s in summaries:
        b = s.get('Buidling', '')
        name = s.get('Tenant Name', '')
        ttype = s.get('Type', '')
        unit = s.get('Unit', '')
        if not b or not name:
            continue
        if ttype and 'apartment' in str(ttype).lower():
            continue
        key = f"{b}|{str(unit).strip()}"
        if key in vacant_keys:
            continue  # vacant units don't need a COI

        summary['total'] += 1
        code = BUILDING_MAP.get(b, {}).get('code', '')
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
            status, coi_label, dl = _classify(exp)
            if status == 'EXPIRED':
                summary['expired'] += 1
            else:
                summary['covered'] += 1
                if dl is not None and dl < 90:
                    summary['expiring_soon'] += 1
            items.append({
                'building': b, 'code': code, 'category': 'Tenant',
                'entity': name, 'unit': str(unit), 'coi': coi_label,
                'expiration': exp,
                'exp_str': exp.strftime('%m/%d/%Y') if exp else 'Unknown',
                'days_left': dl, 'status': status,
            })
        else:
            summary['missing'] += 1
            items.append({
                'building': b, 'code': code, 'category': 'Tenant',
                'entity': name, 'unit': str(unit), 'coi': '❌ NO',
                'expiration': None, 'exp_str': '—', 'days_left': None,
                'status': 'MISSING',
            })

    return items, summary


def split_urgent_expiring(items):
    """Split items into (urgent, expiring). Urgent = expired OR missing.
    Expiring = active certs with days_left < 90, sorted soonest-first."""
    urgent = [it for it in items if it['status'] == 'EXPIRED' or it['status'] == 'MISSING']
    # urgent: expired (by most overdue) then missing
    urgent.sort(key=lambda it: (it['status'] != 'EXPIRED',
                                it['days_left'] if it['days_left'] is not None else 0))
    expiring = [it for it in items
                if it['status'] not in ('EXPIRED', 'MISSING', 'No date')
                and it['days_left'] is not None and 0 <= it['days_left'] < 90]
    expiring.sort(key=lambda it: it['days_left'])
    return urgent, expiring


def generate_coi_pdf(items, summary):
    """Generate a PDF report of the COI status. Returns bytes."""
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import letter
    from reportlab.lib.units import inch
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.platypus import (SimpleDocTemplate, Table, TableStyle, Paragraph,
                                    Spacer)
    import io

    urgent, expiring = split_urgent_expiring(items)
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=letter,
                            leftMargin=0.6 * inch, rightMargin=0.6 * inch,
                            topMargin=0.6 * inch, bottomMargin=0.6 * inch)
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle('T', parent=styles['Title'], fontSize=18,
                                 textColor=colors.HexColor('#1a3a5c'))
    h2 = ParagraphStyle('H2', parent=styles['Heading2'], fontSize=13,
                        spaceBefore=14, spaceAfter=6)
    small = ParagraphStyle('S', parent=styles['Normal'], fontSize=9,
                           textColor=colors.HexColor('#555555'))
    story = []
    today_str = datetime.now().strftime('%B %d, %Y')
    story.append(Paragraph("Marion Street Properties", title_style))
    story.append(Paragraph("Certificate of Insurance Status Report", h2))
    story.append(Paragraph(f"Generated {today_str}", small))
    story.append(Spacer(1, 10))

    cov = (summary['covered'] / summary['total'] * 100) if summary['total'] else 0
    story.append(Paragraph(
        f"<b>Coverage:</b> {cov:.0f}%  |  <b>Active:</b> {summary['covered'] - summary['expiring_soon']}  |  "
        f"<b>Expiring Soon:</b> {summary['expiring_soon']}  |  <b>Expired:</b> {summary['expired']}  |  "
        f"<b>Missing:</b> {summary['missing']}", styles['Normal']))
    story.append(Spacer(1, 6))

    cell_style = ParagraphStyle('cell', parent=styles['Normal'], fontSize=8, leading=9)
    hdr_cell = ParagraphStyle('hcell', parent=styles['Normal'], fontSize=8,
                              leading=9, textColor=colors.white,
                              fontName='Helvetica-Bold')

    def _table(rows, header_color):
        header = [Paragraph(h, hdr_cell) for h in
                  ['Property', 'Type', 'Entity', 'Unit', 'Expiration', 'Days', 'Status']]
        data = [header]
        for it in rows:
            days = '' if it['days_left'] is None else str(it['days_left'])
            data.append([it['code'], it['category'],
                         Paragraph(it['entity'], cell_style),
                         it['unit'], it['exp_str'], days, it['status']])
        t = Table(data, repeatRows=1,
                  colWidths=[0.7*inch, 0.95*inch, 2.25*inch, 0.55*inch, 0.8*inch, 0.55*inch, 0.85*inch])
        style = [
            ('BACKGROUND', (0, 0), (-1, 0), header_color),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 8),
            ('GRID', (0, 0), (-1, -1), 0.4, colors.HexColor('#cccccc')),
            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#f4f6f8')]),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('TOPPADDING', (0, 0), (-1, -1), 3),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
        ]
        for i, it in enumerate(rows, start=1):
            if it['status'] == 'EXPIRED':
                style.append(('TEXTCOLOR', (6, i), (6, i), colors.HexColor('#c0392b')))
                style.append(('FONTNAME', (6, i), (6, i), 'Helvetica-Bold'))
            elif it['status'] == 'MISSING':
                style.append(('TEXTCOLOR', (6, i), (6, i), colors.HexColor('#b9770e')))
                style.append(('FONTNAME', (6, i), (6, i), 'Helvetica-Bold'))
        t.setStyle(TableStyle(style))
        return t

    urgent_style = ParagraphStyle('HU', parent=h2, textColor=colors.HexColor('#c0392b'))
    exp_style = ParagraphStyle('HE', parent=h2, textColor=colors.HexColor('#b9770e'))
    story.append(Paragraph("URGENT — Expired &amp; Missing Certificates", urgent_style))
    if urgent:
        story.append(_table(urgent, colors.HexColor('#c0392b')))
    else:
        story.append(Paragraph("None — all certificates present and current.", styles['Normal']))

    story.append(Paragraph("Expiring in the Next 3 Months", exp_style))
    if expiring:
        story.append(_table(expiring, colors.HexColor('#b9770e')))
    else:
        story.append(Paragraph("None expiring in the next 90 days.", styles['Normal']))

    doc.build(story)
    return buf.getvalue()


def build_coi_email_text(items, summary):
    """Build the plain-text + HTML email body for the COI report."""
    urgent, expiring = split_urgent_expiring(items)
    today_str = datetime.now().strftime('%B %d, %Y')
    cov = (summary['covered'] / summary['total'] * 100) if summary['total'] else 0

    lines = []
    lines.append("Hi Addison, Richie, and Jason,")
    lines.append("")
    lines.append(f"Certificate of Insurance status as of {today_str}:")
    lines.append("")
    lines.append(f"Coverage: {cov:.0f}%  |  Active: {summary['covered'] - summary['expiring_soon']}  |  "
                 f"Expiring Soon: {summary['expiring_soon']}  |  Expired: {summary['expired']}  |  "
                 f"Missing: {summary['missing']}")
    lines.append("")
    lines.append("=== URGENT — EXPIRED & MISSING ===")
    if urgent:
        for it in urgent:
            if it['status'] == 'EXPIRED':
                od = abs(it['days_left']) if it['days_left'] is not None else '?'
                lines.append(f"  [EXPIRED {od}d ago] {it['code']} — {it['category']}: "
                             f"{it['entity']}{(' (Unit ' + it['unit'] + ')') if it['unit'] else ''} "
                             f"— exp. {it['exp_str']}")
            else:
                lines.append(f"  [MISSING] {it['code']} — {it['category']}: "
                             f"{it['entity']}{(' (Unit ' + it['unit'] + ')') if it['unit'] else ''}")
    else:
        lines.append("  None — all certificates present and current.")
    lines.append("")
    lines.append("=== EXPIRING IN NEXT 3 MONTHS (soonest first) ===")
    if expiring:
        for it in expiring:
            lines.append(f"  [{it['days_left']}d left] {it['code']} — {it['category']}: "
                         f"{it['entity']}{(' (Unit ' + it['unit'] + ')') if it['unit'] else ''} "
                         f"— exp. {it['exp_str']}")
    else:
        lines.append("  None expiring in the next 90 days.")
    lines.append("")
    lines.append("Full status report attached (PDF).")
    lines.append("")
    lines.append("— Sis")
    lines.append("Marion Street Properties")
    text = "\n".join(lines)

    def _rows_html(rows, kind):
        if not rows:
            note = ("None — all certificates present and current." if kind == 'urgent'
                    else "None expiring in the next 90 days.")
            return f'<tr><td colspan="5" style="padding:6px;color:#2e7d32;">{note}</td></tr>'
        out = []
        for it in rows:
            if it['status'] == 'EXPIRED':
                badge = f'<span style="color:#c0392b;font-weight:bold;">EXPIRED</span>'
            elif it['status'] == 'MISSING':
                badge = f'<span style="color:#b9770e;font-weight:bold;">MISSING</span>'
            else:
                badge = f'<span style="color:#b9770e;">{it["days_left"]}d left</span>'
            unit = f" (Unit {it['unit']})" if it['unit'] else ''
            out.append(
                f'<tr>'
                f'<td style="padding:5px 8px;border-bottom:1px solid #eee;">{it["code"]}</td>'
                f'<td style="padding:5px 8px;border-bottom:1px solid #eee;">{it["category"]}</td>'
                f'<td style="padding:5px 8px;border-bottom:1px solid #eee;">{it["entity"]}{unit}</td>'
                f'<td style="padding:5px 8px;border-bottom:1px solid #eee;">{it["exp_str"]}</td>'
                f'<td style="padding:5px 8px;border-bottom:1px solid #eee;">{badge}</td>'
                f'</tr>')
        return "".join(out)

    html = f"""\
<html><body style="font-family:Arial,Helvetica,sans-serif;color:#222;font-size:14px;">
<p>Hi Addison, Richie, and Jason,</p>
<p>Certificate of Insurance status as of <b>{today_str}</b>:</p>
<p style="background:#f4f6f8;padding:10px;border-radius:6px;">
<b>Coverage:</b> {cov:.0f}% &nbsp;|&nbsp; <b>Active:</b> {summary['covered'] - summary['expiring_soon']} &nbsp;|&nbsp;
<b>Expiring Soon:</b> {summary['expiring_soon']} &nbsp;|&nbsp; <b>Expired:</b> {summary['expired']} &nbsp;|&nbsp;
<b>Missing:</b> {summary['missing']}</p>
<h3 style="color:#c0392b;margin-bottom:4px;">⚠️ Urgent — Expired &amp; Missing</h3>
<table style="border-collapse:collapse;width:100%;font-size:13px;">
<tr style="background:#c0392b;color:#fff;">
<th style="padding:6px 8px;text-align:left;">Property</th><th style="padding:6px 8px;text-align:left;">Type</th>
<th style="padding:6px 8px;text-align:left;">Entity</th><th style="padding:6px 8px;text-align:left;">Expiration</th>
<th style="padding:6px 8px;text-align:left;">Status</th></tr>
{_rows_html(urgent, 'urgent')}
</table>
<h3 style="color:#b9770e;margin-bottom:4px;margin-top:18px;">🗓️ Expiring in Next 3 Months</h3>
<table style="border-collapse:collapse;width:100%;font-size:13px;">
<tr style="background:#b9770e;color:#fff;">
<th style="padding:6px 8px;text-align:left;">Property</th><th style="padding:6px 8px;text-align:left;">Type</th>
<th style="padding:6px 8px;text-align:left;">Entity</th><th style="padding:6px 8px;text-align:left;">Expiration</th>
<th style="padding:6px 8px;text-align:left;">Status</th></tr>
{_rows_html(expiring, 'expiring')}
</table>
<p style="margin-top:18px;">Full status report attached (PDF).</p>
<p>— Sis<br>Marion Street Properties</p>
</body></html>"""
    return text, html


def send_coi_email(items, summary):
    """Send the COI report to the team. Returns (ok, message).
    SMTP credentials are read from st.secrets['smtp'].
    Expected secrets:
      [smtp]
      host = "smtp.gmail.com"
      port = 587
      user = "Sis.MarionStreet@gmail.com"
      password = "<app password>"
      from = "Sis.MarionStreet@gmail.com"   # optional, defaults to user
    """
    import smtplib
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText
    from email.mime.application import MIMEApplication

    try:
        smtp_cfg = dict(st.secrets['smtp'])
    except Exception:
        return (False, "SMTP credentials not configured. Add an [smtp] section to "
                       "Streamlit Cloud secrets (host, port, user, password).")

    host = smtp_cfg.get('host', 'smtp.gmail.com')
    port = int(smtp_cfg.get('port', 587))
    user = smtp_cfg.get('user')
    password = smtp_cfg.get('password')
    sender = smtp_cfg.get('from', user)
    if not user or not password:
        return (False, "SMTP user/password missing from secrets.")

    text, html = build_coi_email_text(items, summary)
    subject = f"MSP Certificate of Insurance Report — {datetime.now().strftime('%m/%d/%Y')}"

    msg = MIMEMultipart('mixed')
    msg['Subject'] = subject
    msg['From'] = sender
    msg['To'] = ", ".join(EMAIL_TEAM_RECIPIENTS)
    alt = MIMEMultipart('alternative')
    alt.attach(MIMEText(text, 'plain'))
    alt.attach(MIMEText(html, 'html'))
    msg.attach(alt)

    try:
        pdf_bytes = generate_coi_pdf(items, summary)
        pdf_part = MIMEApplication(pdf_bytes, _subtype='pdf')
        pdf_part.add_header('Content-Disposition', 'attachment',
                            filename=f"MSP_COI_Report_{datetime.now().strftime('%Y-%m-%d')}.pdf")
        msg.attach(pdf_part)
    except Exception as e:
        return (False, f"Failed to generate PDF: {e}")

    try:
        if port == 465:
            with smtplib.SMTP_SSL(host, port, timeout=30) as server:
                server.login(user, password)
                server.sendmail(sender, EMAIL_TEAM_RECIPIENTS, msg.as_string())
        else:
            with smtplib.SMTP(host, port, timeout=30) as server:
                server.starttls()
                server.login(user, password)
                server.sendmail(sender, EMAIL_TEAM_RECIPIENTS, msg.as_string())
    except Exception as e:
        return (False, f"SMTP send failed: {e}")

    return (True, f"Report emailed to {', '.join(EMAIL_TEAM_RECIPIENTS)}.")


def render_insurance_tab():
    tenants, _, summaries = load_tenancy()
    coi_data, building_coi_data, pm_coi_data = scan_coi_files()
    today_dt = datetime.now()
    vacant_keys, vacant_meta = build_vacancy_lookup(tenants, include_auto=False)

    # --- Email Team button + Urgent / Expiring sections (top of tab) ---
    report_items, report_summary = build_coi_report_data()
    urgent_items, expiring_items = split_urgent_expiring(report_items)

    hdr_col, btn_col = st.columns([3, 1])
    with hdr_col:
        st.markdown("### 🛡️ Certificate of Insurance Reconciliation")
    with btn_col:
        if st.button("📧 Email Team", key="coi_email_team", use_container_width=True,
                     help="Email the COI report (PDF + summary) to Addison, Richie, and Jason."):
            with st.spinner("Sending report to the team…"):
                ok, msg = send_coi_email(report_items, report_summary)
            if ok:
                st.success(f"✅ {msg}")
            else:
                st.error(f"❌ {msg}")

    # Downloadable PDF (always available, even if email isn't configured)
    try:
        _pdf = generate_coi_pdf(report_items, report_summary)
        btn_col.download_button(
            "⬇️ Download PDF", data=_pdf,
            file_name=f"MSP_COI_Report_{datetime.now().strftime('%Y-%m-%d')}.pdf",
            mime="application/pdf", key="coi_pdf_dl", use_container_width=True)
    except Exception:
        pass

    # (Urgent / Expiring summary sections are rendered at the BOTTOM of this tab.)

    # Build insurance data
    ins_rows = []
    summary = {'total': 0, 'covered': 0, 'expired': 0, 'missing': 0, 'expiring_soon': 0}

    for s in summaries:
        b = s.get('Buidling', '')
        name = s.get('Tenant Name', '')
        ttype = s.get('Type', '')
        unit = s.get('Unit', '')
        if not b or not name:
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
                    coi_label = '❌ EXP'
                elif days_left < 60:
                    status = f'{days_left}d left'
                    summary['expiring_soon'] += 1
                    summary['covered'] += 1
                    coi_label = '✅ YES'
                else:
                    status = 'Active'
                    summary['covered'] += 1
                    coi_label = '✅ YES'
                exp_str = exp.strftime('%m/%d/%Y')
            else:
                exp_str = 'Unknown'
                days_left = None
                status = 'No date'
                summary['covered'] += 1
                coi_label = '✅ YES'
            ins_rows.append({
                'Building': b, 'Tenant': name, 'Unit': str(unit), 'Type': ttype,
                'COI': coi_label, 'Expiration': exp_str,
                'Days Left': str(days_left) if days_left is not None else '—',
                'Status': status, 'File': matched['filename'],
                'View': build_coi_url(b, matched['filename']),
            })
        else:
            summary['missing'] += 1
            ins_rows.append({
                'Building': b, 'Tenant': name, 'Unit': str(unit), 'Type': ttype,
                'COI': '❌ NO', 'Expiration': '—', 'Days Left': '—',
                'Status': 'MISSING', 'File': '', 'View': '',
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
    import pandas as pd
    for building_name in BUILDING_MAP:
        b_rows = [r for r in ins_rows if r['Building'] == building_name]
        if not b_rows:
            continue
        b_covered = sum(1 for r in b_rows if '✅' in r['COI'])
        n_certs = len(coi_data.get(building_name, []))
        code = BUILDING_MAP[building_name]['code']

        st.markdown(f'<div class="building-header"><strong>▎ {building_name} ({code})</strong> — {b_covered}/{len(b_rows)} covered · {n_certs} files on disk</div>', unsafe_allow_html=True)

        display_cols = ['Tenant', 'Unit', 'Type', 'COI', 'Expiration', 'Days Left', 'Status', 'View']

        # --- Building-level certificate section (shown BEFORE tenants) ---
        b_building_certs = building_coi_data.get(building_name, [])
        bldg_rows = []
        if b_building_certs:
            for cert in b_building_certs:
                exp = cert.get('exp_date')
                if exp:
                    days_left = (exp - today_dt).days
                    if exp < today_dt:
                        b_status = 'EXPIRED'
                        b_coi = '❌ EXP'
                    elif days_left < 60:
                        b_status = f'{days_left}d left'
                        b_coi = '✅ YES'
                    else:
                        b_status = 'Active'
                        b_coi = '✅ YES'
                    b_exp_str = exp.strftime('%m/%d/%Y')
                    b_days = str(days_left)
                else:
                    b_exp_str = 'Unknown'
                    b_days = '—'
                    b_status = 'No date'
                    b_coi = '✅ YES'
                bldg_rows.append({
                    'Tenant': cert.get('insured_name') or building_name,
                    'Unit': '', 'Type': 'Building',
                    'COI': b_coi, 'Expiration': b_exp_str,
                    'Days Left': b_days, 'Status': b_status,
                    'View': build_coi_url(building_name, cert.get('filename')),
                })
        else:
            bldg_rows.append({
                'Tenant': building_name, 'Unit': '', 'Type': 'Building',
                'COI': '❌ NO', 'Expiration': '—',
                'Days Left': '—', 'Status': 'MISSING', 'View': '',
            })
        st.markdown('<div style="color:#f0883e;font-size:0.8rem;font-weight:600;margin:4px 0 2px 8px;">🏢 BUILDING CERTIFICATE</div>', unsafe_allow_html=True)
        bldg_df = pd.DataFrame(bldg_rows)
        show_grid(bldg_df[display_cols], key=f"ins_bldg_{building_name}", tab_key="insurance")

        # --- Property Manager certificate section ---
        b_pm_certs = pm_coi_data.get(building_name, [])
        pm_rows = []
        if b_pm_certs:
            for cert in b_pm_certs:
                exp = cert.get('exp_date')
                if exp:
                    days_left = (exp - today_dt).days
                    if exp < today_dt:
                        pm_status = 'EXPIRED'
                        pm_coi = '❌ EXP'
                    elif days_left < 60:
                        pm_status = f'{days_left}d left'
                        pm_coi = '✅ YES'
                    else:
                        pm_status = 'Active'
                        pm_coi = '✅ YES'
                    pm_exp_str = exp.strftime('%m/%d/%Y')
                    pm_days = str(days_left)
                else:
                    pm_exp_str = 'Unknown'
                    pm_days = '—'
                    pm_status = 'No date'
                    pm_coi = '✅ YES'
                pm_rows.append({
                    'Tenant': cert.get('insured_name') or 'Property Manager',
                    'Unit': '', 'Type': 'PM',
                    'COI': pm_coi, 'Expiration': pm_exp_str,
                    'Days Left': pm_days, 'Status': pm_status,
                    'View': build_coi_url(building_name, cert.get('filename')),
                })
        else:
            pm_rows.append({
                'Tenant': 'Property Manager', 'Unit': '', 'Type': 'PM',
                'COI': '❌ NO', 'Expiration': '—',
                'Days Left': '—', 'Status': 'MISSING', 'View': '',
            })
        st.markdown('<div style="color:#9b59b6;font-size:0.8rem;font-weight:600;margin:8px 0 2px 8px;">👤 PROPERTY MANAGER CERTIFICATE</div>', unsafe_allow_html=True)
        pm_df = pd.DataFrame(pm_rows)
        show_grid(pm_df[display_cols], key=f"ins_pm_{building_name}", tab_key="insurance")

        # --- Tenant certificates ---
        st.markdown('<div style="color:#8b949e;font-size:0.8rem;font-weight:600;margin:8px 0 2px 8px;">👥 TENANTS</div>', unsafe_allow_html=True)
        df = pd.DataFrame(b_rows)
        show_grid(df[display_cols], key=f"ins_{building_name}", tab_key="insurance")

    render_column_config_editor('insurance', ['Tenant', 'Unit', 'Type', 'COI', 'Expiration', 'Days Left', 'Status', 'View'])

    # --- Summary sections (bottom of tab): Urgent + Expiring in 3 months ---
    st.divider()
    st.markdown("### 📋 COI Summary")
    _uc = len(urgent_items)
    _ec = len(expiring_items)
    with st.expander(f"⚠️ URGENT — Expired & Missing ({_uc})", expanded=_uc > 0):
        if urgent_items:
            urg_df = pd.DataFrame([{
                'Property': it['code'], 'Type': it['category'],
                'Entity': it['entity'], 'Unit': it['unit'],
                'Expiration': it['exp_str'],
                'Overdue/Days': (f"{abs(it['days_left'])}d ago" if it['status'] == 'EXPIRED'
                                 and it['days_left'] is not None else ''),
                'Status': it['status'],
            } for it in urgent_items])
            st.dataframe(urg_df, hide_index=True, use_container_width=True)
        else:
            st.success("All certificates present and current.")

    with st.expander(f"🗓️ Expiring in Next 3 Months ({_ec})", expanded=_ec > 0):
        if expiring_items:
            exp_df = pd.DataFrame([{
                'Property': it['code'], 'Type': it['category'],
                'Entity': it['entity'], 'Unit': it['unit'],
                'Expiration': it['exp_str'], 'Days Left': it['days_left'],
                'Status': it['status'],
            } for it in expiring_items])
            st.dataframe(exp_df, hide_index=True, use_container_width=True)
        else:
            st.info("Nothing expiring in the next 90 days.")


def render_deposits_tab():
    tenants, details, summaries = load_tenancy()
    active = [t for t in tenants if t['Tenant'] not in ('Easement',)]
    today_dt = datetime.now()
    vacant_keys, vacant_meta = build_vacancy_lookup(tenants, include_auto=False)

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

        # Find current SD using tightest-fit logic (smallest end date where start <= today <= end)
        current_sd = first.get('Sec Dep', 0) or 0
        next_sd = None
        next_sd_date = None
        current_year_row = None
        best_end = None

        rows_with_dates = []
        for r in rows:
            start = r.get('Start Date')
            end = r.get('End Date')
            sd = r.get('Sec Dep', 0) or 0
            if isinstance(start, (datetime, date)) and isinstance(end, (datetime, date)):
                start_d = start.date() if isinstance(start, datetime) else start
                end_d = end.date() if isinstance(end, datetime) else end
                rows_with_dates.append((start_d, end_d, sd, r))
                if start_d <= TODAY <= end_d:
                    if best_end is None or end_d < best_end:
                        best_end = end_d
                        current_sd = sd
                        current_year_row = r

        # Next SD: find the row with the smallest end date AFTER current period's end
        if current_year_row and best_end:
            next_sd_date = best_end  # SD Anniversary = end of current period
            for start_d, end_d, sd, r in sorted(rows_with_dates, key=lambda x: x[1]):
                if end_d > best_end:
                    next_sd = sd
                    break

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

        # SD Anniversary = end of current period (when next term starts), shown for all tenants with future periods
        sd_anniv_date = next_sd_date if next_sd_date and next_sd_date > TODAY else None
        sd_new_amount = next_sd if sd_anniv_date and next_sd is not None else 0
        sd_delta = (sd_new_amount - current_sd) if sd_anniv_date else 0

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
            'Current SD': current_sd if isinstance(current_sd, (int, float)) else 0,
            'Current SD Fmt': f"${float(current_sd):,.2f}" if isinstance(current_sd, (int, float)) and current_sd > 0 else '❌ $0.00',
            'SD Anniversary': sd_anniv_date,
            'New SD Amount': sd_new_amount,
            'SD Delta': sd_delta,
            'TTE': tte_str,
        })

    dep_data.sort(key=lambda d: (d['Building'], d['Tenant']))

    # Yardi SD comparison — use Security Deposit Activity page (Deposits On Hand)
    # Use the reconcile name mapping to bridge Yardi spaces to spreadsheet spaces
    import json as _json_dep
    yardi_deposits = parse_yardi_deposit_activity()
    yardi_rent_data = parse_yardi_rent_rolls()

    # Load name mapping from reconcile tab (Yardi key -> spreadsheet tenant name)
    name_map_dep = {}
    try:
        raw = _read_gsheet_config("Config: Yardi Names")
        if raw:
            name_map_dep = _json_dep.loads(raw)
    except Exception:
        pass

    # Build reverse lookup: "Building|SheetSpace" -> yardi deposit amount
    # Strategy: for each Yardi deposit entry, find which spreadsheet tenant it maps to
    sheet_sd_from_yardi = {}  # "Building|SheetSpace" -> deposit amount

    # First, build sheet tenant->space index
    sheet_tenant_to_space = {}
    for d in dep_data:
        sheet_tenant_to_space[f"{d['Building']}|{d['Tenant']}"] = d['Space']

    for yardi_key, deposit_amt in yardi_deposits.items():
        building = yardi_key.split('|', 1)[0] if '|' in yardi_key else ''
        yardi_space = yardi_key.split('|', 1)[1] if '|' in yardi_key else ''

        # Method 1: Check name mapping (reconcile lookup table)
        matched_tenant = name_map_dep.get(yardi_key, '')

        # Method 2: Auto-match via normalized space in rent roll
        if not matched_tenant and yardi_key in yardi_rent_data:
            raw_unit = yardi_rent_data[yardi_key].get('raw_unit', yardi_space)
            norm = normalize_space(raw_unit, building)
            # Find a dep_data entry with this building+space
            for d in dep_data:
                if d['Building'] == building and d['Space'] == norm:
                    matched_tenant = d['Tenant']
                    break

        # Method 3: Direct space match
        if not matched_tenant:
            for d in dep_data:
                if d['Building'] == building and d['Space'] == yardi_space:
                    matched_tenant = d['Tenant']
                    break

        if matched_tenant:
            sheet_space = sheet_tenant_to_space.get(f"{building}|{matched_tenant}")
            if sheet_space:
                sheet_sd_from_yardi[f"{building}|{sheet_space}"] = deposit_amt

    for d in dep_data:
        key = f"{d['Building']}|{d['Space']}"
        yardi_sd = sheet_sd_from_yardi.get(key)
        d['Yardi SD'] = yardi_sd
        dash_sd = d['Current SD']
        if yardi_sd is None:
            d['SD Diff'] = 'Not in Yardi'
        elif abs(dash_sd - yardi_sd) > 0.01:
            d['SD Diff'] = f"${dash_sd:,.0f} vs ${yardi_sd:,.0f}"
        else:
            d['SD Diff'] = ''

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
        display_cols = ['Display Tenant', 'Space', 'Type', 'Lease', 'Current SD Fmt', 'SD Anniversary', 'New SD Amount', 'SD Delta', 'Yardi SD', 'SD Diff', 'TTE']
        display_df = df[display_cols].copy()
        display_df.columns = ['Tenant', 'Space', 'Type', 'Lease', 'Current SD', 'SD Anniversary', 'New SD', 'Δ SD', 'Yardi SD', 'SD Diff', 'MTE']
        display_df['SD Anniversary'] = display_df['SD Anniversary'].apply(
            lambda d: d.strftime('%m/%d/%Y') if isinstance(d, (datetime, date)) else ('-' if not d else str(d))
        )
        display_df['New SD'] = display_df['New SD'].apply(lambda x: f"${x:,.0f}" if x is not None and not (isinstance(x, float) and pd.isna(x)) else '$0')
        display_df['Δ SD'] = display_df['Δ SD'].apply(
            lambda x: f"{'+' if x and x > 0 else ''}${x:,.0f}" if x is not None and not (isinstance(x, float) and pd.isna(x)) else '$0'
        )
        display_df['Yardi SD'] = display_df['Yardi SD'].apply(
            lambda x: f"${x:,.2f}" if x is not None and not (isinstance(x, float) and pd.isna(x)) else 'N/A'
        )
        show_grid(display_df, key=f"dep_{building_name}", tab_key="deposits")

    render_column_config_editor('deposits', ['Tenant', 'Space', 'Type', 'Lease', 'Current SD', 'SD Anniversary', 'New SD', 'Δ SD', 'Yardi SD', 'SD Diff', 'MTE'])

    # Portfolio total
    st.divider()
    st.metric("Portfolio Total Security Deposits", f"${total_deposits:,.2f}")


def render_yardi_tab():
    import base64
    yardi_dir = Path(os.path.dirname(__file__)) / "data" / "Yardi"

    st.markdown("### 📊 Yardi Reports")

    if not yardi_dir.exists():
        st.warning("No Yardi reports found. Place PDF files in data/Yardi/ folder.")
        return

    pdfs = sorted([f for f in yardi_dir.iterdir() if f.suffix.lower() == '.pdf'])
    if not pdfs:
        st.warning("No PDF files found in data/Yardi/.")
        return

    st.caption(f"{len(pdfs)} report(s) available")

    # Sort PDFs by date (newest first) using Mon-YYYY prefix
    def _pdf_sort_key(p):
        m = re.match(r'^([a-z]{3})-(\d{4})', p.stem.lower())
        if m:
            return int(m.group(2)) * 100 + MONTH_ORDER.get(m.group(1), 0)
        return 0

    pdfs_sorted = sorted(pdfs, key=_pdf_sort_key, reverse=True)

    # Group by building
    building_files = {}
    for pdf in pdfs_sorted:
        name = pdf.stem
        matched_building = None
        for bldg_name, mapping in BUILDING_MAP.items():
            dest = mapping['dest_folder'].lower().replace('-', '').replace(' ', '')
            fname = name.lower().replace('-', '').replace(' ', '').replace('_', '')
            if any(part in fname for part in [
                dest, bldg_name.lower().replace(' ', ''),
                mapping['code'].lower()
            ]):
                matched_building = bldg_name
                break
        if not matched_building:
            matched_building = "Other"
        building_files.setdefault(matched_building, []).append(pdf)

    for building, files in building_files.items():
        code = BUILDING_MAP.get(building, {}).get('code', '')
        label = f"{building} ({code})" if code else building
        st.markdown(f'<div style="margin:8px 0 2px 0;font-size:0.95em;"><strong>▎ {label}</strong> — {len(files)} report(s)</div>', unsafe_allow_html=True)

        for pdf in files:
            with open(pdf, "rb") as f:
                pdf_bytes = f.read()
            b64 = base64.b64encode(pdf_bytes).decode()

            # Clean display name: "May 2026" from "May-2026_10097-15-South-St_..."
            date_match = re.match(r'^([A-Za-z]{3})-(\d{4})', pdf.stem)
            if date_match:
                display_name = f"{date_match.group(1)} {date_match.group(2)}"
            else:
                display_name = pdf.stem.replace('_', ' ').replace('-', ' ')

            with st.expander(f"📄 {display_name}", expanded=False):
                col_dl_inner, col_spacer = st.columns([2, 5])
                col_dl_inner.download_button("⬇️ Download", data=pdf_bytes, file_name=pdf.name, mime="application/pdf", key=f"dl_{pdf.name}")
                if st.session_state.get('mobile_view', False) and HAS_PYMUPDF:
                    doc = fitz.open(str(pdf))
                    total_pages = len(doc)
                    page_key = f"yardi_page_{pdf.name}"
                    if page_key not in st.session_state:
                        st.session_state[page_key] = 0
                    pg = max(0, min(st.session_state[page_key], total_pages - 1))
                    cp, cn_num, cn = st.columns([1, 2, 1])
                    if cp.button("⬅️", disabled=(pg == 0), key=f"yp_{pdf.name}"):
                        st.session_state[page_key] = pg - 1
                        st.rerun()
                    cn_num.markdown(f"<div style='text-align:center;padding-top:8px;'>{pg+1} / {total_pages}</div>", unsafe_allow_html=True)
                    if cn.button("➡️", disabled=(pg >= total_pages - 1), key=f"yn_{pdf.name}"):
                        st.session_state[page_key] = pg + 1
                        st.rerun()
                    page = doc[pg]
                    pix = page.get_pixmap(matrix=fitz.Matrix(0.9, 0.9))
                    img_b64 = base64.b64encode(pix.tobytes("png")).decode()
                    doc.close()
                    st.markdown(f'<img src="data:image/png;base64,{img_b64}" style="width:100%;border-radius:6px;">', unsafe_allow_html=True)
                else:
                    st.markdown(f'<iframe src="data:application/pdf;base64,{b64}" width="100%" height="600px" style="border:1px solid rgba(255,255,255,0.1);border-radius:6px;"></iframe>', unsafe_allow_html=True)


def render_reconcile_tab():
    """Yardi Reconciliation — Yardi-driven with lookup table at top for name matching."""
    import json as _json
    st.markdown("### 🔄 Yardi Reconciliation")

    tenants, details, summaries = load_tenancy()
    yardi_data = parse_yardi_rent_rolls()
    if not yardi_data:
        st.warning("No Yardi rent roll data found. Place PDF files in data/Yardi/ folder.")
        return

    # --- Name mapping: Yardi key -> spreadsheet tenant name (persisted in Google Sheets) ---
    MAPPING_TAB = "Config: Yardi Names"
    name_map = {}  # key: "Building|YardiSpace" -> spreadsheet tenant name
    try:
        raw = _read_gsheet_config(MAPPING_TAB)
        if raw:
            name_map = _json.loads(raw)
    except Exception:
        pass

    # Build spreadsheet tenant index keyed by "Building|TenantName" for lookup
    active = [t for t in tenants if t['Tenant'] not in ('-',)] if tenants else []
    sheet_by_name = {}  # "Building|TenantName" -> tenant dict
    sheet_by_space = {}  # "Building|Space" -> tenant dict
    for t in active:
        b = t.get('Building', '').strip()
        n = t.get('Tenant', '').strip()
        s = str(t.get('Space', '')).strip()
        if b and n:
            sheet_by_name[f"{b}|{n}"] = t
        if b and s:
            sheet_by_space[f"{b}|{s}"] = t

    # Get all unique spreadsheet tenant names per building for dropdown
    building_tenants = {}
    for t in active:
        b = t.get('Building', '').strip()
        n = t.get('Tenant', '').strip()
        if b and n:
            building_tenants.setdefault(b, []).append(n)
    for b in building_tenants:
        building_tenants[b] = sorted(set(building_tenants[b]))

    # Helper: get spreadsheet SD for a tenant
    def _get_sheet_sd(building, tenant_name):
        sd = 0
        for dk, dv in details.items():
            if not dv:
                continue
            first = dv[0]
            if first.get('Building') == building and first.get('Tenant') == tenant_name:
                best_end = None
                for r in dv:
                    start = r.get('Start Date')
                    end = r.get('End Date')
                    s = r.get('Sec Dep', 0) or 0
                    if isinstance(start, (datetime, date)) and isinstance(end, (datetime, date)):
                        start_d = start.date() if isinstance(start, datetime) else start
                        end_d = end.date() if isinstance(end, datetime) else end
                        if start_d <= TODAY <= end_d:
                            if best_end is None or end_d < best_end:
                                best_end = end_d
                                sd = s
                break
        return sd

    # ===========================
    # SECTION 1: LOOKUP TABLE (compact)
    # ===========================
    with st.expander("🔗 Name Lookup Table", expanded=False):
        st.caption("Match Yardi names → spreadsheet names. Changes save automatically.")

        # Inject compact CSS for lookup table
        st.markdown("""<style>
            .lookup-table { width: 100%; border-collapse: collapse; font-size: 0.8rem; margin-bottom: 0.5rem; }
            .lookup-table th { text-align: left; padding: 2px 6px; border-bottom: 1px solid #444; color: #888; font-weight: 600; }
            .lookup-table td { padding: 1px 6px; border-bottom: 1px solid #2a2a2a; }
            .lookup-table .unit { color: #888; width: 50px; }
            .lookup-table .yname { white-space: nowrap; max-width: 200px; overflow: hidden; text-overflow: ellipsis; }
        </style>""", unsafe_allow_html=True)

        mapping_changed = False

        for building_name in BUILDING_MAP:
            b_yardi = {k: v for k, v in yardi_data.items() if k.startswith(f"{building_name}|")}
            if not b_yardi:
                continue

            code = BUILDING_MAP[building_name]['code']
            st.markdown(f"<small><strong>{building_name} ({code})</strong></small>", unsafe_allow_html=True)

            available_names = ["—"] + building_tenants.get(building_name, [])

            cols = st.columns([1, 3, 3])
            cols[0].markdown("<small>**Unit**</small>", unsafe_allow_html=True)
            cols[1].markdown("<small>**Yardi Name**</small>", unsafe_allow_html=True)
            cols[2].markdown("<small>**Sheet Match**</small>", unsafe_allow_html=True)

            for yardi_key in sorted(b_yardi.keys()):
                yardi_entry = b_yardi[yardi_key]
                yardi_name = yardi_entry.get('tenant', '')
                yardi_space = yardi_key.split('|', 1)[1] if '|' in yardi_key else ''

                current_match = name_map.get(yardi_key, '')
                if not current_match:
                    norm_space = normalize_space(yardi_entry.get('raw_unit', yardi_space), building_name)
                    auto_key = f"{building_name}|{norm_space}"
                    auto_tenant = sheet_by_space.get(auto_key)
                    if auto_tenant:
                        current_match = auto_tenant.get('Tenant', '')

                default_idx = 0
                if current_match in available_names:
                    default_idx = available_names.index(current_match)

                c1, c2, c3 = st.columns([1, 3, 3])
                c1.markdown(f"<small style='color:#888;padding-top:8px;display:block;'>{yardi_space}</small>", unsafe_allow_html=True)
                c2.markdown(f"<small style='padding-top:8px;display:block;'>{yardi_name}</small>", unsafe_allow_html=True)
                selected = c3.selectbox("m", available_names, index=default_idx,
                                        key=f"lu_{yardi_key}", label_visibility="collapsed")

                new_val = '' if selected == "—" else selected
                if new_val != name_map.get(yardi_key, ''):
                    name_map[yardi_key] = new_val
                    mapping_changed = True

        if mapping_changed:
            name_map = {k: v for k, v in name_map.items() if v}
            _write_gsheet_config(MAPPING_TAB, _json.dumps(name_map))
            st.rerun()

    # ===========================
    # SECTION 2: RECONCILIATION GRID
    # ===========================
    st.markdown("---")
    st.markdown("#### 📊 Data Comparison")

    # Get Yardi deposit activity (Deposits On Hand) for SD comparison
    yardi_deposits = parse_yardi_deposit_activity()

    recon_rows = []
    for yardi_key, yardi_entry in sorted(yardi_data.items()):
        building = yardi_key.split('|', 1)[0] if '|' in yardi_key else ''
        yardi_space = yardi_key.split('|', 1)[1] if '|' in yardi_key else ''
        yardi_name = yardi_entry.get('tenant', '')
        yardi_monthly = yardi_entry.get('monthly', 0) or 0
        yardi_exp = yardi_entry.get('expiration')
        yardi_sd = yardi_deposits.get(yardi_key, 0) or 0

        # Find matched spreadsheet tenant
        matched_name = name_map.get(yardi_key, '')
        if not matched_name:
            # Auto-match via normalized space
            norm_space = normalize_space(yardi_entry.get('raw_unit', yardi_space), building)
            auto_key = f"{building}|{norm_space}"
            auto_t = sheet_by_space.get(auto_key)
            if auto_t:
                matched_name = auto_t.get('Tenant', '')

        # Get spreadsheet data for matched tenant
        sheet_t = sheet_by_name.get(f"{building}|{matched_name}") if matched_name else None
        dash_monthly = sheet_t.get('Monthly', 0) if sheet_t else None
        dash_exp_str = sheet_t.get('Exp Date', '') if sheet_t else ''
        dash_sd = _get_sheet_sd(building, matched_name) if matched_name else None

        # Compute diffs
        rent_diff = ''
        if sheet_t and dash_monthly is not None:
            diff_val = dash_monthly - yardi_monthly
            if abs(diff_val) > 0.01:
                rent_diff = f"${diff_val:+,.0f}"

        exp_diff = ''
        yardi_exp_str = yardi_exp.strftime('%m/%d/%Y') if yardi_exp else 'N/A'
        if dash_exp_str and dash_exp_str != 'MTM' and yardi_exp:
            try:
                d_exp = datetime.strptime(dash_exp_str, '%m/%d/%Y').date()
                if d_exp != yardi_exp:
                    delta = (d_exp - yardi_exp).days
                    exp_diff = f"{delta:+d}d"
            except (ValueError, AttributeError):
                pass

        sd_diff = ''
        if dash_sd is not None and abs(dash_sd - yardi_sd) > 0.01:
            sd_diff = f"${dash_sd - yardi_sd:+,.0f}"

        # Status
        if not matched_name:
            status = '❓ Unmatched'
        else:
            diffs = []
            if rent_diff:
                diffs.append('Rent')
            if exp_diff:
                diffs.append('Exp')
            if sd_diff:
                diffs.append('SD')
            status = '⚠️ ' + ', '.join(diffs) if diffs else '✅ Match'

        recon_rows.append({
            'Building': building,
            'Space': yardi_space,
            'Yardi Name': yardi_name,
            'Sheet Name': matched_name or '—',
            'Rent (Yardi)': f"${yardi_monthly:,.0f}",
            'Rent (Sheet)': f"${dash_monthly:,.0f}" if dash_monthly is not None else '—',
            'Rent Δ': rent_diff,
            'Exp (Yardi)': yardi_exp_str,
            'Exp (Sheet)': dash_exp_str if dash_exp_str else '—',
            'Exp Δ': exp_diff,
            'SD (Yardi)': f"${yardi_sd:,.2f}",
            'SD (Sheet)': f"${dash_sd:,.2f}" if dash_sd is not None else '—',
            'SD Δ': sd_diff,
            'Status': status,
        })

    # Append sheet-only tenants (in spreadsheet but not in Yardi) into main table
    matched_sheet_keys_main = set()
    for r in recon_rows:
        if r['Sheet Name'] and r['Sheet Name'] != '—':
            matched_sheet_keys_main.add(f"{r['Building']}|{r['Sheet Name']}")
    for yk, sn in name_map.items():
        b = yk.split('|', 1)[0] if '|' in yk else ''
        if b and sn:
            matched_sheet_keys_main.add(f"{b}|{sn}")

    for t in active:
        t_key = f"{t.get('Building', '')}|{t.get('Tenant', '')}"
        if t_key not in matched_sheet_keys_main and t.get('Tenant', '') not in ('-', '') and 'VACANT' not in t.get('Tenant', '').upper():
            dash_monthly = t.get('Monthly', 0) or 0
            dash_sd = _get_sheet_sd(t.get('Building', ''), t.get('Tenant', ''))
            recon_rows.append({
                'Building': t.get('Building', ''),
                'Space': str(t.get('Space', '')),
                'Yardi Name': '—',
                'Sheet Name': t.get('Tenant', ''),
                'Rent (Yardi)': '—',
                'Rent (Sheet)': f"${dash_monthly:,.0f}",
                'Rent Δ': '',
                'Exp (Yardi)': '—',
                'Exp (Sheet)': t.get('Exp Date', ''),
                'Exp Δ': '',
                'SD (Yardi)': '—',
                'SD (Sheet)': f"${dash_sd:,.2f}" if dash_sd else '—',
                'SD Δ': '',
                'Status': '📋 Sheet Only',
            })

    recon_rows.sort(key=lambda r: (r['Building'], r['Space']))

    # Summary metrics
    total = len(recon_rows)
    matches = sum(1 for r in recon_rows if r['Status'] == '✅ Match')
    mismatches = sum(1 for r in recon_rows if r['Status'].startswith('⚠️'))
    unmatched = sum(1 for r in recon_rows if r['Status'] == '❓ Unmatched')

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total Yardi Units", str(total))
    c2.metric("✅ Match", str(matches))
    c3.metric("⚠️ Mismatch", str(mismatches))
    c4.metric("❓ Unmatched", str(unmatched))

    # Filter
    filter_opt = st.radio("Show:", ["All", "Mismatches Only", "Unmatched Only"], horizontal=True, key="recon_filter")
    filtered = recon_rows
    if filter_opt == "Mismatches Only":
        filtered = [r for r in recon_rows if r['Status'].startswith('⚠️')]
    elif filter_opt == "Unmatched Only":
        filtered = [r for r in recon_rows if r['Status'] == '❓ Unmatched']

    for building_name in BUILDING_MAP:
        b_rows = [r for r in filtered if r['Building'] == building_name]
        if not b_rows:
            continue
        b_match = sum(1 for r in b_rows if r['Status'] == '✅ Match')
        code = BUILDING_MAP[building_name]['code']
        st.markdown(f'<div class="building-header"><strong>▎ {building_name} ({code})</strong> — {b_match}/{len(b_rows)} matching</div>', unsafe_allow_html=True)

        df = pd.DataFrame(b_rows)
        display_cols = ['Space', 'Yardi Name', 'Sheet Name', 'Rent (Yardi)', 'Rent (Sheet)', 'Rent Δ',
                        'Exp (Yardi)', 'Exp (Sheet)', 'Exp Δ', 'SD (Yardi)', 'SD (Sheet)', 'SD Δ', 'Status']
        display_df = df[display_cols].copy()
        show_grid(display_df, key=f"recon_{building_name}", tab_key="reconcile")

    render_column_config_editor('reconcile', ['Space', 'Yardi Name', 'Sheet Name', 'Rent (Yardi)', 'Rent (Sheet)', 'Rent Δ',
                                               'Exp (Yardi)', 'Exp (Sheet)', 'Exp Δ', 'SD (Yardi)', 'SD (Sheet)', 'SD Δ', 'Status'])

    # ===========================
    # Sheet-only tenants are now included in the main recon_rows table above


# =====================
# LEAD SHEET
# =====================
LEAD_SHEET_ID = "1Gbl7FM-C4dN_BHVwT_1dRbkODEZW57MdymzBMDvPrCw"
LEAD_SHEET_GID = 1762616490

@st.cache_data(ttl=300)
def _load_lead_changelog():
    """Load the Lead Sheet Changelog and return dict of {(section, row): (timestamp, user)} for most recent edit per row."""
    try:
        sheet = get_gsheet()
        if not sheet:
            return {}
        ws = sheet.worksheet('Lead Sheet Changelog')
        records = ws.get_all_records()
    except Exception:
        # Try the lead sheet's own spreadsheet
        try:
            client = get_gspread_client()
            if not client:
                return {}
            ss = client.open_by_key(LEAD_SHEET_ID)
            ws = ss.worksheet('Lead Sheet Changelog')
            records = ws.get_all_records()
        except Exception:
            return {}

    latest = {}
    for r in records:
        section = r.get('Section', 'Retail')
        row = r.get('Row', 0)
        ts = r.get('Timestamp', '')
        user = r.get('User', '')
        key = (section, row)
        # Keep the latest timestamp per row
        if key not in latest or str(ts) > str(latest[key][0]):
            latest[key] = (ts, user)
    return latest


@st.cache_data(ttl=300)
def load_lead_sheet():
    """Load lead sheet from Google Sheets and return retail_df, office_df."""
    client = get_gspread_client()
    if client is None:
        st.error("No Google service account configured.")
        return pd.DataFrame(), pd.DataFrame()
    try:
        sheet = client.open_by_key(LEAD_SHEET_ID)
        ws = sheet.get_worksheet_by_id(LEAD_SHEET_GID)
        rows = ws.get_all_values()
    except Exception as e:
        st.error(f"Failed to load lead sheet: {e}")
        return pd.DataFrame(), pd.DataFrame()

    if not rows:
        return pd.DataFrame(), pd.DataFrame()

    # Dynamically read column layout from header row
    header = rows[0]

    # Find the blank separator column between retail and office
    sep_col = None
    for i, h in enumerate(header):
        if i > 0 and not h.strip():
            # Check if previous cols had data and next col looks like a header
            if i + 1 < len(header) and header[i + 1].strip():
                sep_col = i
                break

    if sep_col is None:
        # Fallback: assume no office section
        sep_col = len(header)

    retail_headers = [h.strip() for h in header[:sep_col]]
    office_headers = [h.strip() for h in header[sep_col + 1:]] if sep_col < len(header) - 1 else []

    # Normalize office headers to match retail (strip "Office "/"Retail " prefix from first col)
    # Use retail headers as canonical column names, replacing first col name
    col_names = list(retail_headers)
    col_names[0] = "Tenant"

    # Map date columns to age column names, preserving order
    _age_name_map = {"Contacted": "Contact", "Shown": "Show"}
    date_cols = {}
    for c in col_names:
        if 'date' in c.lower():
            raw = c.replace("Date ", "").replace("Date", "").strip()
            friendly = _age_name_map.get(raw, raw)
            date_cols[c] = f"{friendly} Age"

    def parse_rows(rows_data, start_col, num_cols):
        parsed = []
        for row in rows_data[1:]:
            cells = row[start_col:start_col + num_cols] if len(row) > start_col else []
            # Pad if short
            while len(cells) < num_cols:
                cells.append('')
            cells = [c.strip() for c in cells[:num_cols]]
            if any(c for c in cells):
                parsed.append(cells)
        return parsed

    def add_age_columns(df):
        today = date.today()
        for date_col, age_col in date_cols.items():
            ages = []
            for val in df[date_col]:
                if val:
                    try:
                        dt = pd.to_datetime(val, format='mixed', dayfirst=False).date()
                        delta = (today - dt).days
                        ages.append(f"{delta}d")
                    except Exception:
                        ages.append("")
                else:
                    ages.append("")
            # Insert age column right after its date column
            idx = df.columns.get_loc(date_col) + 1
            df.insert(idx, age_col, ages)
        return df

    retail_rows = parse_rows(rows, 0, len(retail_headers))
    office_cols_count = len(office_headers) if office_headers else len(retail_headers)
    office_rows = parse_rows(rows, sep_col + 1, office_cols_count)

    # Office headers may differ slightly; normalize to match retail col_names
    office_col_names = list(col_names)
    if office_headers:
        office_col_names = list(office_headers)
        office_col_names[0] = "Tenant"
    # Ensure same length
    while len(office_col_names) < office_cols_count:
        office_col_names.append(f"Col_{len(office_col_names)}")

    retail_df = pd.DataFrame(retail_rows, columns=col_names) if retail_rows else pd.DataFrame(columns=col_names)
    office_df = pd.DataFrame(office_rows, columns=office_col_names[:office_cols_count]) if office_rows else pd.DataFrame(columns=col_names)

    retail_df = add_age_columns(retail_df)
    office_df = add_age_columns(office_df)

    # Merge changelog data (last modified + modified by)
    changelog = _load_lead_changelog()
    for section, df in [('Retail', retail_df), ('Office', office_df)]:
        last_mod = []
        mod_by = []
        for i in range(len(df)):
            sheet_row = i + 2  # row 1 is header, data starts at row 2
            entry = changelog.get((section, sheet_row))
            if entry:
                ts_str = str(entry[0])
                # Shorten email to just the name part
                user = entry[1].split('@')[0] if '@' in entry[1] else entry[1]
                last_mod.append(ts_str)
                mod_by.append(user)
            else:
                last_mod.append('')
                mod_by.append('')
        df['Last Modified'] = last_mod
        df['Modified By'] = mod_by

    return retail_df, office_df


def render_covenants_tab():
    """Render the Lease Covenants tab directly from MSP Tenancy.xlsx (cols AI-AU, rows 10-40)."""
    st.markdown("### 📜 Lease Covenants")
    st.caption("Special lease provisions from MSP Tenancy.xlsx · Edit the spreadsheet and re-upload to update")

    if not TENANCY_FILE:
        st.warning("MSP Tenancy.xlsx not found.")
        return

    # Column mapping: spreadsheet column letter → display name
    COV_COLUMNS = {
        35: "Renewal Options",
        36: "Option Notice",
        37: "LL Termination",
        38: "Tenant Termination",
        39: "Non-Compete / Exclusive",
        40: "CAM Obligations",
        41: "ROFO/ROFR",
        42: "Assignment/Subletting",
        43: "Personal Guarantee",
        44: "Holdover Rate",
        45: "Late Fee",
        46: "Other Notable",
    }

    try:
        wb = openpyxl.load_workbook(TENANCY_FILE, data_only=True)
        ws = wb.active

        rows = []
        for r in range(11, 41):
            building = ws.cell(row=r, column=1).value
            tenant = ws.cell(row=r, column=3).value
            if not building or not tenant:
                continue
            tenant_str = str(tenant)
            is_vacant = tenant_str.upper().startswith("VACANT") or tenant_str.upper() == "EASEMENT"
            row_data = {
                "Building": str(building),
                "Unit": ws.cell(row=r, column=2).value or "",
                "Tenant": "VACANT" if is_vacant else tenant_str,
            }
            for col_idx, col_name in COV_COLUMNS.items():
                if is_vacant:
                    row_data[col_name] = ""
                else:
                    val = ws.cell(row=r, column=col_idx).value
                    if val is not None and isinstance(val, datetime) and col_idx == 36:
                        row_data[col_name] = val.strftime("%m/%d/%Y")
                    else:
                        row_data[col_name] = str(val) if val is not None else ""
            rows.append(row_data)
        wb.close()
    except Exception as e:
        st.error(f"Error reading tenancy file: {e}")
        return

    if not rows:
        st.info("No covenant data found in spreadsheet.")
        return

    df = pd.DataFrame(rows)

    # --- Building filter ---
    buildings = sorted(df["Building"].unique().tolist())
    selected = st.multiselect("Filter by Building", buildings, default=buildings, key="cov_building_filter")
    filtered = df[df["Building"].isin(selected)].copy()

    # Summary metrics
    total = len(filtered)

    def _count_has(col_name, keyword="YES"):
        if col_name not in filtered.columns:
            return 0
        return filtered[col_name].astype(str).str.upper().str.contains(keyword, na=False).sum()

    has_noncompete = _count_has("Non-Compete / Exclusive")
    has_rofr = _count_has("ROFO/ROFR")
    has_guarantee = _count_has("Personal Guarantee")

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total Tenants", total)
    c2.metric("Non-Compete", has_noncompete)
    c3.metric("ROFO/ROFR", has_rofr)
    c4.metric("Personal Guarantee", has_guarantee)

    # Display per building
    for bldg in selected:
        bldg_df = filtered[filtered["Building"] == bldg].copy()
        if bldg_df.empty:
            continue
        st.markdown(f"#### 🏢 {bldg}")
        display_df = bldg_df.drop(columns=["Building"]).reset_index(drop=True)
        # Auto-size: 42px header + 42px per row + buffer
        grid_h = 42 + 42 * len(display_df) + 12
        show_grid(display_df, key=f"cov_{bldg}", height=grid_h, tab_key="covenants")

    # Column width config at the bottom
    all_cov_cols = ["Unit", "Tenant"] + list(COV_COLUMNS.values())
    render_column_config_editor('covenants', all_cov_cols)


def render_lead_sheet_tab():
    """Render the Lead Sheet tab with Retail and Office sections."""
    st.markdown("### 📋 Lead Sheet")
    st.caption("Prospect leads from Google Sheets · Auto-refreshes every 5 minutes")

    retail_df, office_df = load_lead_sheet()

    if retail_df.empty and office_df.empty:
        st.warning("No lead data found. Check Google Sheets connection.")
        return

    # Summary metrics
    c1, c2, c3, c4 = st.columns(4)
    c1.markdown(f'<div class="metric-card"><div class="metric-value">{len(retail_df)}</div><div class="metric-label">Retail Leads</div></div>', unsafe_allow_html=True)
    c2.markdown(f'<div class="metric-card"><div class="metric-value">{len(office_df)}</div><div class="metric-label">Office Leads</div></div>', unsafe_allow_html=True)
    showing_count = sum(1 for _, r in pd.concat([retail_df, office_df]).iterrows() if r.get("Showing", "").strip().lower() == "yes")
    loi_count = sum(1 for _, r in pd.concat([retail_df, office_df]).iterrows() if r.get("LOI", "").strip().lower() == "yes")
    c3.markdown(f'<div class="metric-card"><div class="metric-value">{showing_count}</div><div class="metric-label">Showings Done</div></div>', unsafe_allow_html=True)
    c4.markdown(f'<div class="metric-card"><div class="metric-value">{loi_count}</div><div class="metric-label">LOIs Received</div></div>', unsafe_allow_html=True)

    st.markdown("")

    # Retail section
    st.markdown('<div class="building-header">🏪 Retail Leads</div>', unsafe_allow_html=True)
    if not retail_df.empty:
        show_grid(retail_df, key="leads_retail", tab_key="leads")
    else:
        st.info("No retail leads.")

    st.markdown("")

    # Office section
    st.markdown('<div class="building-header">🏢 Office Leads</div>', unsafe_allow_html=True)
    if not office_df.empty:
        show_grid(office_df, key="leads_office", tab_key="leads")
    else:
        st.info("No office leads.")

    # Use retail columns (with age cols) as the config source
    all_cols = list(retail_df.columns) if not retail_df.empty else list(office_df.columns)
    render_column_config_editor('leads', all_cols)

    if st.button("🔄 Refresh Lead Sheet", key="refresh_leads"):
        load_lead_sheet.clear()
        st.rerun()


# =====================
# MAIN APP
# =====================

col_title, col_toggle = st.columns([8, 2])
col_title.markdown("## 🏢 MSP Property Dashboard")
st.session_state.mobile_view = col_toggle.toggle("📱 Mobile", value=st.session_state.mobile_view)
st.caption(f"Marion Street Properties · {TODAY.strftime('%B %d, %Y')}")

tab_tenancy, tab_vacancy, tab_leads, tab_covenants, tab_insurance, tab_deposits, tab_reconcile, tab_yardi, tab_sop = st.tabs([
    "🏠 Current Tenancy", "🏚️ Vacancy", "📋 Lead Sheet", "📜 Lease Covenants", "🛡️ Insurance", "💰 Security Deposits", "🔄 Yardi Reconcile", "📊 Yardi Reports", "📋 SOPs"
])

with tab_sop:
    render_sop_tab()

with tab_tenancy:
    render_tenancy_tab()

with tab_vacancy:
    render_vacancy_tab()

with tab_leads:
    render_lead_sheet_tab()

with tab_covenants:
    render_covenants_tab()

with tab_insurance:
    render_insurance_tab()

with tab_deposits:
    render_deposits_tab()

with tab_reconcile:
    render_reconcile_tab()

with tab_yardi:
    render_yardi_tab()
