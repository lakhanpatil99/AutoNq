import streamlit as st
import pandas as pd
from datetime import date
import os
import logging
import base64
from PIL import Image
from io import BytesIO

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

from ui_styles import (
    THEME_CSS, SIDEBAR_LOGO, NAV_ITEMS, BOSCH_LOGO_B64,
    kpi_card, section_header, ai_output_card,
    status_badge, empty_state,
    structured_ai_card, mail_preview_card,
    tracker_issue_card, priority_badge, domain_tag,
    severity_badge, category_tag, observation_display_card
)

# ── Excel backend ─────────────────────────────────────────────────────────────
from excel_backend import (
    load_master_data,
    load_audit_entries,
    add_audit_entry,
    get_line_options,
    get_stations_for_line,
    get_station_no,
    get_severity_options,
    get_category_options,
)
# ─────────────────────────────────────────────────────────────────────────────

# ── Outlook mail integration (safe lazy import — never runs at startup) ────────
try:
    from mail_service import send_mail as _send_mail_fn
    _MAIL_SERVICE_AVAILABLE = True
except ImportError:
    _send_mail_fn = None
    _MAIL_SERVICE_AVAILABLE = False

try:
    from mail_service import search_bosch_users as _search_users_fn
except ImportError:
    _search_users_fn = None
# ──────────────────────────────────────────────────────────────────────────────

logger = logging.getLogger(__name__)


# ── Cached Graph API people search (avoids repeated calls on Streamlit reruns)
@st.cache_data(ttl=300, show_spinner=False)
def _cached_people_search(query: str) -> list:
    """Cached wrapper around Graph API /users search.
    TTL = 5 min — same query reuses cached results across reruns."""
    if _search_users_fn is None:
        return []
    try:
        return _search_users_fn(query)
    except Exception:
        return []

# ── Unified Line/Area options (shared across ALL pages) ────────────────────────
# -- Dynamic Line/Area options (LIVE from Excel line_master) ----------------
# get_line_options() reads from line_master sheet in audit_master_data.xlsx
# Returns: ['HVML', 'LVML', 'Line 1', 'Line 2', 'Line 4', 'Line 6', 'Test Line']

def _get_live_lines():
    """Return dynamic line options from Excel line_master. Refreshed each Streamlit rerun."""
    return get_line_options()


def clean_line_data(df, col="line"):
    """UI-layer data cleaning: filter to lines present in line_master.
    Does NOT modify the database -- operates on DataFrame copies."""
    if df is None or df.empty or col not in df.columns:
        return df
    df = df.copy()
    df[col] = df[col].astype(str).str.strip()
    valid = _get_live_lines()
    df = df[df[col].isin(valid)]
    return df


# ===============================
# PAGE CONFIG
# ===============================

st.set_page_config(page_title="AutoNQ AI", layout="wide", page_icon="🤖")
st.markdown(THEME_CSS, unsafe_allow_html=True)

# ── Table dark-wrapper helper ──────────────────────────────────────────────────
_TABLE_WRAP_CSS = """
<style>
.nq-table-wrap {
    background: #132D44;
    border: 1px solid rgba(129,195,215,0.16);
    border-top: 2.5px solid #81C3D7;
    border-radius: 12px;
    overflow: hidden;
    box-shadow: 0 4px 24px rgba(10,25,41,0.45);
    margin: 8px 0 16px;
}
.nq-table-wrap[data-testid="stDataFrame"],
.nq-table-wrap [data-testid="stDataEditor"],
.nq-table-wrap [data-testid="stDataFrameResizable"],
.nq-table-wrap[data-testid="stDataFrame"] > div,
.nq-table-wrap [data-testid="stDataFrame"] > div > div,
.nq-table-wrap [data-testid="stDataFrame"] > div > div > div,
.nq-table-wrap [data-testid="stDataEditor"] > div,
.nq-table-wrap [data-testid="stDataEditor"] > div > div,
.nq-table-wrap [data-testid="stDataEditor"] > div > div > div {
    background: #132D44 !important;
    color: #D9DCD6 !important;
}
.nq-table-wrap [data-testid="stDataFrame"] > div > div > div:first-child,
.nq-table-wrap [data-testid="stDataEditor"] > div > div > div:first-child {
    background: #16425B !important;
    border-bottom: 1px solid rgba(129,195,215,0.16) !important;
    padding: 4px 10px !important;
}
.nq-table-wrap [data-testid="stDataFrame"] iframe,
.nq-table-wrap[data-testid="stDataEditor"] iframe {
    background: #132D44 !important;
    color-scheme: dark !important;
}
.nq-table-wrap [role="columnheader"] {
    background: #16425B !important;
    color: #9BADB8 !important;
    font-size: 11px !important;
    font-weight: 700 !important;
    text-transform: uppercase !important;
    letter-spacing: 0.07em !important;
    border-bottom: 1px solid rgba(129,195,215,0.16) !important;
    font-family: 'Inter', sans-serif !important;
}
.nq-table-wrap [role="gridcell"] {
    background: #132D44 !important;
    color: #D9DCD6 !important;
    font-size: 13px !important;
    font-family: 'Inter', sans-serif !important;
    border-bottom: 1px solid rgba(129,195,215,0.06) !important;
}
.nq-table-wrap [role="row"]:hover[role="gridcell"] {
    background: #16425B !important;
}
.nq-table-wrap [role="gridcell"][aria-selected="true"] {
    background: rgba(129,195,215,0.1) !important;
    outline: 1px solid #81C3D7 !important;
    outline-offset: -1px !important;
}
.nq-table-wrap [role="rowheader"] {
    background: #16425B !important;
    color: #5E7A8A !important;
    font-size: 11.5px !important;
    font-weight: 600 !important;
    border-right: 1px solid rgba(129,195,215,0.16) !important;
    font-family: 'Inter', sans-serif !important;
}
.nq-table-wrap table {
    background: #132D44 !important;
    width: 100% !important;
    border-collapse: collapse !important;
}
.nq-table-wrap thead tr th {
    background: #16425B !important;
    color: #9BADB8 !important;
    font-size: 11px !important;
    font-weight: 700 !important;
    text-transform: uppercase !important;
    letter-spacing: 0.07em !important;
    border-bottom: 1.5px solid rgba(129,195,215,0.16) !important;
    padding: 12px 16px !important;
    font-family: 'Inter', sans-serif !important;
    white-space: nowrap !important;
}
.nq-table-wrap tbody tr td {
    color: #D9DCD6 !important;
    border-bottom: 1px solid rgba(129,195,215,0.06) !important;
    padding: 10px 16px !important;
    font-size: 13px !important;
    background: #132D44 !important;
    font-family: 'Inter', sans-serif !important;
}
.nq-table-wrap tbody tr:nth-child(even) td {
    background: rgba(22,66,91,0.3) !important;
}
.nq-table-wrap tbody tr:hover td {
    background: #16425B !important;
}
.nq-table-wrap ::-webkit-scrollbar { width: 5px; height: 5px; }
.nq-table-wrap ::-webkit-scrollbar-track { background: #132D44; }
.nq-table-wrap ::-webkit-scrollbar-thumb { background: #3A5060; border-radius: 99px; }
.nq-table-wrap ::-webkit-scrollbar-thumb:hover { background: #5E7A8A; }
</style>
"""

def _df_wrap_open(label: str = "") -> None:
    st.markdown(_TABLE_WRAP_CSS, unsafe_allow_html=True)

def _df_wrap_close() -> None:
    pass

# ===============================
# EXPORT / DOWNLOAD ENGINE
# ===============================

def _df_to_excel_bytes(df: pd.DataFrame) -> bytes:
    """Convert a DataFrame to .xlsx bytes in memory. Uses openpyxl (already a dependency)."""
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=True, sheet_name="Export")
    return buf.getvalue()


def _df_to_csv_bytes(df: pd.DataFrame) -> bytes:
    """Convert a DataFrame to UTF-8 CSV bytes."""
    return df.to_csv(index=True).encode("utf-8-sig")  # BOM for Excel compat


_DOWNLOAD_BAR_CSS = """
<style>
.nq-dl-bar {
    display: flex;
    align-items: center;
    gap: 10px;
    flex-wrap: wrap;
    margin: 10px 0 14px;
}
</style>
"""

# Themed CSS injected once per page to style Streamlit download buttons
_DOWNLOAD_BTN_CSS = """
<style>
/* ── AutoNQ themed download buttons ────────────────────────────────── */
div[data-testid="stDownloadButton"] > button {
    background: linear-gradient(135deg, rgba(19,45,68,0.95), rgba(22,66,91,0.90)) !important;
    border: 1.5px solid rgba(77,166,255,0.25) !important;
    border-radius: 12px !important;
    color: #D9DCD6 !important;
    font-family: 'Inter', sans-serif !important;
    font-size: 13px !important;
    font-weight: 600 !important;
    padding: 10px 22px !important;
    min-height: 44px !important;
    box-shadow: 0 2px 12px rgba(77,166,255,0.08), 0 1px 3px rgba(0,0,0,0.2) !important;
    transition: all 0.25s cubic-bezier(0.2,0.8,0.2,1) !important;
    letter-spacing: 0.02em !important;
}
div[data-testid="stDownloadButton"] > button:hover {
    border-color: rgba(77,166,255,0.5) !important;
    box-shadow: 0 4px 20px rgba(77,166,255,0.18), 0 2px 8px rgba(0,0,0,0.3) !important;
    transform: translateY(-1px) !important;
    color: #FFFFFF !important;
    background: linear-gradient(135deg, rgba(22,66,91,0.98), rgba(58,124,165,0.3)) !important;
}
div[data-testid="stDownloadButton"] > button:active {
    transform: translateY(0px) scale(0.98) !important;
    box-shadow: 0 1px 6px rgba(77,166,255,0.12) !important;
}
</style>
"""


def _render_download_bar(
    df: pd.DataFrame,
    base_name: str,
    *,
    excel: bool = True,
    csv: bool = True,
    key_prefix: str = "dl",
):
    """Render a row of themed download buttons directly above a table.
    
    Args:
        df: The DataFrame to export (must not be None or a string).
        base_name: File base name, e.g. "QCheck_2026-05-31"
        excel: Show Excel download button.
        csv: Show CSV download button.
        key_prefix: Unique key prefix to avoid Streamlit widget ID collisions.
    """
    if df is None or isinstance(df, str) or (hasattr(df, 'empty') and df.empty):
        return

    st.markdown(_DOWNLOAD_BTN_CSS, unsafe_allow_html=True)
    st.markdown(_DOWNLOAD_BAR_CSS, unsafe_allow_html=True)

    # Calculate how many columns we need
    cols_needed = int(excel) + int(csv)
    if cols_needed == 0:
        return

    # Create columns with spacer to keep buttons compact on the left
    col_widths = [1] * cols_needed + [max(1, 6 - cols_needed)]
    dl_cols = st.columns(col_widths)

    col_idx = 0
    if excel:
        with dl_cols[col_idx]:
            st.download_button(
                label="📥 Download Excel",
                data=_df_to_excel_bytes(df),
                file_name=f"{base_name}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"{key_prefix}_xlsx",
            )
        col_idx += 1

    if csv:
        with dl_cols[col_idx]:
            st.download_button(
                label="📥 Download CSV",
                data=_df_to_csv_bytes(df),
                file_name=f"{base_name}.csv",
                mime="text/csv",
                key=f"{key_prefix}_csv",
            )


# ===============================
# LLM CLIENT  (lazy — never runs at import time)
# ===============================

_llm_instance = None

def _get_llm():
    global _llm_instance
    if _llm_instance is None:
        from setup_environment import BoschLLMClient
        _llm_instance = BoschLLMClient()
    return _llm_instance

@st.cache_data(show_spinner=False)
def cached_llm(_messages_tuple):
    messages = [dict(m) for m in _messages_tuple]
    return _get_llm().chat(messages)

def call_llm(messages):
    key = tuple(tuple(sorted(m.items())) for m in messages)
    return cached_llm(key)

# ===============================
# AI DETECTION
# ===============================

def detect_principle(observation):

    principles = [
        "Stop Sign", "Andon Cord", "Instructions", "Process Parameters",
        "Measurement / Test Equipment", "Check the Checker",
        "Total Productive Maintenance (TPM)", "Tools", "Restart",
        "Labeling", "Rework / Scrap", "Dropped Parts", "Correct Product",
        "Remaining Items", "1C – Cleanliness"
    ]

    prompt = f"""
Classify this manufacturing audit observation into ONE principle.

Observation:
{observation}

Choose ONLY from:
{principles}

Return only the principle name.
"""

    try:
        response = call_llm([
            {"role": "system", "content": "You are a manufacturing audit expert."},
            {"role": "user", "content": prompt}
        ])

        result = response.strip()
        if result in principles:
            return result
        for p in principles:
            if p.lower() in result.lower():
                return p
        logger.warning(f"Unrecognized principle: {result}")
        return "UNCLASSIFIED"

    except Exception as e:
        logger.warning(f"Principle detection failed: {e}")
        return "UNCLASSIFIED"

# ===============================
# AGENT FUNCTIONS  (lazy — deferred imports, never run at startup)
# ===============================

_agents = {}

def _agent(name):
    if name not in _agents:
        import setup_environment as _se
        _agents.update({
            "generate_daily_brief":                    _se.generate_daily_brief,
            "generate_weekly_brief":                   _se.generate_weekly_brief,
            "generate_monthly_brief":                  _se.generate_monthly_brief,
            "generate_mail_from_summary":              _se.generate_mail_from_summary,
            "classify_by_percentile":                  _se.classify_by_percentile,
            "generate_governance_annual_plan":         _se.generate_governance_annual_plan,
            "generate_guided_audit_questions":         _se.generate_guided_audit_questions,
            "generate_qcheck_questions":               _se.generate_qcheck_questions,
            "generate_iatf_process_audit_sheet":       _se.generate_iatf_process_audit_sheet,
            "generate_followup_checklist":             _se.generate_followup_checklist,
            "generate_external_audit_tracker_with_ai": _se.generate_external_audit_tracker_with_ai,
            "map_deviation_category_ai":               _se.map_deviation_category_ai,
            "preprocess_audit_data":                   _se.preprocess_audit_data,
            "tag_domain":                              _se.tag_domain,
            "tag_severity":                            _se.tag_severity,
            "verify_data_integrity":                   _se.verify_data_integrity,
        })
    return _agents[name]

def generate_daily_brief(*a, **kw):                    return _agent("generate_daily_brief")(*a, **kw)
def generate_weekly_brief(*a, **kw):                   return _agent("generate_weekly_brief")(*a, **kw)
def generate_monthly_brief(*a, **kw):                  return _agent("generate_monthly_brief")(*a, **kw)
def generate_mail_from_summary(*a, **kw):              return _agent("generate_mail_from_summary")(*a, **kw)
def classify_by_percentile(*a, **kw):                  return _agent("classify_by_percentile")(*a, **kw)
def generate_governance_annual_plan(*a, **kw):         return _agent("generate_governance_annual_plan")(*a, **kw)
def generate_guided_audit_questions(*a, **kw):         return _agent("generate_guided_audit_questions")(*a, **kw)
def generate_qcheck_questions(*a, **kw):               return _agent("generate_qcheck_questions")(*a, **kw)
def generate_iatf_process_audit_sheet(*a, **kw):       return _agent("generate_iatf_process_audit_sheet")(*a, **kw)
def generate_followup_checklist(*a, **kw):             return _agent("generate_followup_checklist")(*a, **kw)
def generate_external_audit_tracker_with_ai(*a, **kw): return _agent("generate_external_audit_tracker_with_ai")(*a, **kw)
def map_deviation_category_ai(*a, **kw):               return _agent("map_deviation_category_ai")(*a, **kw)
def preprocess_audit_data(*a, **kw):                   return _agent("preprocess_audit_data")(*a, **kw)
def tag_domain(*a, **kw):                              return _agent("tag_domain")(*a, **kw)
def tag_severity(*a, **kw):                            return _agent("tag_severity")(*a, **kw)

# ===============================
# SVG ICON MAP
# ===============================

def _icon(name, size=18, color="#81C3D7"):
    icons = {
        "chart":    f'<svg width="{size}" height="{size}" viewBox="0 0 24 24" fill="none" stroke="{color}" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><line x1="18" y1="20" x2="18" y2="10"/><line x1="12" y1="20" x2="12" y2="4"/><line x1="6" y1="20" x2="6" y2="14"/></svg>',
        "calendar": f'<svg width="{size}" height="{size}" viewBox="0 0 24 24" fill="none" stroke="{color}" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="4" width="18" height="18" rx="2" ry="2"/><line x1="16" y1="2" x2="16" y2="6"/><line x1="8" y1="2" x2="8" y2="6"/><line x1="3" y1="10" x2="21" y2="10"/></svg>',
        "alert":    f'<svg width="{size}" height="{size}" viewBox="0 0 24 24" fill="none" stroke="{color}" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="m21.73 18-8-14a2 2 0 0 0-3.48 0l-8 14A2 2 0 0 0 4 21h16a2 2 0 0 0 1.73-3Z"/><line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12.01" y2="17"/></svg>',
        "trend":    f'<svg width="{size}" height="{size}" viewBox="0 0 24 24" fill="none" stroke="{color}" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="22 12 18 12 15 21 9 3 6 12 2 12"/></svg>',
        "edit":     f'<svg width="{size}" height="{size}" viewBox="0 0 24 24" fill="none" stroke="{color}" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>',
        "summary":  f'<svg width="{size}" height="{size}" viewBox="0 0 24 24" fill="none" stroke="{color}" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="16" y1="13" x2="8" y2="13"/><line x1="16" y1="17" x2="8" y2="17"/></svg>',
        "plan":     f'<svg width="{size}" height="{size}" viewBox="0 0 24 24" fill="none" stroke="{color}" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M2 3h6a4 4 0 0 1 4 4v14a3 3 0 0 0-3-3H2z"/><path d="M22 3h-6a4 4 0 0 0-4 4v14a3 3 0 0 1 3-3h7z"/></svg>',
        "check":    f'<svg width="{size}" height="{size}" viewBox="0 0 24 24" fill="none" stroke="{color}" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"/><polyline points="22 4 12 14.01 9 11.01"/></svg>',
        "shield":   f'<svg width="{size}" height="{size}" viewBox="0 0 24 24" fill="none" stroke="{color}" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z"/></svg>',
        "list":     f'<svg width="{size}" height="{size}" viewBox="0 0 24 24" fill="none" stroke="{color}" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><line x1="8" y1="6" x2="21" y2="6"/><line x1="8" y1="12" x2="21" y2="12"/><line x1="8" y1="18" x2="21" y2="18"/><line x1="3" y1="6" x2="3.01" y2="6"/><line x1="3" y1="12" x2="3.01" y2="12"/><line x1="3" y1="18" x2="3.01" y2="18"/></svg>',
        "globe":    f'<svg width="{size}" height="{size}" viewBox="0 0 24 24" fill="none" stroke="{color}" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"/><line x1="2" y1="12" x2="22" y2="12"/><path d="M12 2a15.3 15.3 0 0 1 4 10 15.3 15.3 0 0 1-4 10 15.3 15.3 0 0 1-4-10 15.3 15.3 0 0 1 4-10z"/></svg>',
        "repeat":   f'<svg width="{size}" height="{size}" viewBox="0 0 24 24" fill="none" stroke="{color}" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="17 1 21 5 17 9"/><path d="M3 11V9a4 4 0 0 1 4-4h14"/><polyline points="7 23 3 19 7 15"/><path d="M21 13v2a4 4 0 0 1-4 4H3"/></svg>',
        "camera":   f'<svg width="{size}" height="{size}" viewBox="0 0 24 24" fill="none" stroke="{color}" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M23 19a2 2 0 0 1-2 2H3a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h4l2-3h6l2 3h4a2 2 0 0 1 2 2z"/><circle cx="12" cy="13" r="4"/></svg>',
        "refresh":  f'<svg width="{size}" height="{size}" viewBox="0 0 24 24" fill="none" stroke="{color}" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M3 12a9 9 0 1 0 9-9 9.75 9.75 0 0 0-6.74 2.74L3 8"/><path d="M3 3v5h5"/></svg>',
        "clock":    f'<svg width="{size}" height="{size}" viewBox="0 0 24 24" fill="none" stroke="{color}" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"/><polyline points="12 6 12 12 16 14"/></svg>',
        "fire":     f'<svg width="{size}" height="{size}" viewBox="0 0 24 24" fill="none" stroke="{color}" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M8.5 14.5A2.5 2.5 0 0 0 11 12c0-1.38-.5-2-1-3-1.072-2.143-.224-4.054 2-6 .5 2.5 2 4.9 4 6.5 2 1.6 3 3.5 3 5.5a7 7 0 1 1-14 0c0-1.153.433-2.294 1-3a2.5 2.5 0 0 0 2.5 2.5z"/></svg>',
        "flag":     f'<svg width="{size}" height="{size}" viewBox="0 0 24 24" fill="none" stroke="{color}" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M4 15s1-1 4-1 5 2 8 2 4-1 4-1V3s-1 1-4 1-5-2-8-2-4 1-4 1z"/><line x1="4" y1="22" x2="4" y2="15"/></svg>',
        "pin":      f'<svg width="{size}" height="{size}" viewBox="0 0 24 24" fill="none" stroke="{color}" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><line x1="12" y1="17" x2="12" y2="22"/><path d="M5 17h14v-1.76a2 2 0 0 0-1.11-1.79l-1.78-.9A2 2 0 0 1 15 10.68V6a3 3 0 0 0-3-3 3 3 0 0 0-3 3v4.68a2 2 0 0 1-1.11 1.79l-1.78.9A2 2 0 0 0 5 15.24Z"/></svg>',
        "mail":     f'<svg width="{size}" height="{size}" viewBox="0 0 24 24" fill="none" stroke="{color}" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="2" y="4" width="20" height="16" rx="2"/><path d="m22 7-8.97 5.7a1.94 1.94 0 0 1-2.06 0L2 7"/></svg>',
    }
    return icons.get(name, f'<span style="font-size:{size}px;">{name}</span>')

# ===============================
# SESSION STATE
# ===============================

if "agent_outputs" not in st.session_state:
    st.session_state.agent_outputs = {}
if "observations" not in st.session_state:
    st.session_state.observations = []

if "station_typed" not in st.session_state:
    st.session_state.station_typed = ""
if "station_selected" not in st.session_state:
    st.session_state.station_selected = ""

# ===============================
# DATA LOADER  (Excel backend)
# ===============================

def get_data():
    """Load all audit data from Excel backend live. Streamlit's @st.cache_data handles performance."""
    master_data = load_master_data()
    full_data = load_audit_entries()

    if "audit_date" in full_data.columns:
        full_data["audit_date"] = pd.to_datetime(full_data["audit_date"], errors="coerce")

    # ── FIX-A: Normalise image column name ─────────────────────────────
    if "image_base64" not in full_data.columns:
        if "image" in full_data.columns:
            full_data["image_base64"] = full_data["image"].fillna("")
        elif "image_path" in full_data.columns:
            full_data["image_base64"] = full_data["image_path"].fillna("")
        else:
            full_data["image_base64"] = ""
    else:
        full_data["image_base64"] = full_data["image_base64"].fillna("")
    # ───────────────────────────────────────────────────────────────────

    # We return the full data. Slicing logic for Daily/Weekly/Monthly happens at usage time.
    cutoff = pd.Timestamp.today() - pd.Timedelta(days=30)
    current_data  = full_data[full_data["audit_date"] >= cutoff].copy() if not full_data.empty else full_data.copy()
    previous_data = full_data[full_data["audit_date"] <  cutoff].copy() if not full_data.empty else full_data.copy()

    iqis_df      = pd.DataFrame()
    iqis_risk_df = pd.DataFrame()
    lpc_df       = pd.DataFrame()

    # ── Canonical AI Preprocessing (Single Source of Truth) ────────────
    from excel_backend import get_audit_df_for_ai
    raw_ai_df = get_audit_df_for_ai()
    # Apply preprocessing EXACTLY ONCE globally
    master_canonical_df = _agent("preprocess_audit_data")(raw_ai_df)

    return (
        current_data,
        previous_data,
        full_data,
        iqis_df,
        iqis_risk_df,
        lpc_df,
        master_canonical_df
    )

# ===============================
# SIDEBAR NAVIGATION
# ===============================

with st.sidebar:
    st.markdown(SIDEBAR_LOGO, unsafe_allow_html=True)
    st.markdown("---")
    page = st.radio("Navigation", NAV_ITEMS, label_visibility="collapsed")
    st.markdown("---")
    st.markdown(
        '<div style="display:flex;align-items:center;gap:8px;padding:6px 0;">'
        '<span style="width:8px;height:8px;border-radius:50%;background:#6ECBA0;display:inline-block;"></span>'
        '<span style="font-size:12px;color:#9BADB8;font-family:\'Inter\',sans-serif;">Excel Backend · Online</span>'
        '</div>',
        unsafe_allow_html=True,
    )
    st.caption(f"{date.today().strftime('%d %b %Y')}")

# ===============================
# LOAD DATA + KPIs
# ===============================

current_data, previous_data, full_data, iqis_df, iqis_risk_df, lpc_df, master_canonical_df = get_data()

LINE_OPTIONS = _get_live_lines()  # Dynamic, refreshed each rerun

current_data  = clean_line_data(current_data,  col="line")
previous_data = clean_line_data(previous_data, col="line")
full_data     = clean_line_data(full_data,     col="line")
iqis_df       = clean_line_data(iqis_df,       col="LINE") if not iqis_df.empty else iqis_df
iqis_risk_df  = clean_line_data(iqis_risk_df,  col="LINE") if not iqis_risk_df.empty else iqis_risk_df
lpc_df        = clean_line_data(lpc_df,        col="LINE") if not lpc_df.empty else lpc_df

img_tag = f'<img src="data:image/png;base64,{BOSCH_LOGO_B64}" style="height: 28px; margin-right: 14px; margin-bottom: -4px; border-radius: 2px;" />' if BOSCH_LOGO_B64 else ''

st.markdown(f"""
<div style="text-align:center; padding: 6px 0 2px;">
  <h1 style="
    background: linear-gradient(135deg, #F5F7FA, #9BADB8);
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
    font-size: 2.2rem; font-weight: 800; margin-bottom: 0;
    font-family: 'Inter', sans-serif; letter-spacing: -0.03em;
    display: flex; align-items: center; justify-content: center;
  ">
    {img_tag}AutoNQ AI
  </h1>
  <p style="color: #9BADB8; font-size: 0.95rem; margin-top: 6px;
            font-family: 'Inter', sans-serif; font-weight: 400;">
    Powered by Bosch Quality Excellence
  </p>
</div>
""", unsafe_allow_html=True)

if not full_data.empty:
    kc1, kc2, kc3, kc4 = st.columns(4)
    top_issue = current_data["ai_principle"].value_counts().idxmax() if not current_data.empty and "ai_principle" in current_data.columns else "N/A"
    with kc1:
        st.markdown(kpi_card(_icon("chart", 22), "Total Observations", len(full_data), "cyan"), unsafe_allow_html=True)
    with kc2:
        st.markdown(kpi_card(_icon("calendar", 22), "Current Period", len(current_data), "blue"), unsafe_allow_html=True)
    with kc3:
        st.markdown(kpi_card(_icon("alert", 22, "#E0A84D"), "Top Issue", top_issue, "amber"), unsafe_allow_html=True)
    with kc4:
        lines_covered = current_data["line"].nunique() if not current_data.empty and "line" in current_data.columns else 0
        st.markdown(kpi_card(_icon("trend", 22, "#6ECBA0"), "Lines Covered", lines_covered, "emerald"), unsafe_allow_html=True)
else:
    st.markdown(empty_state("No data available yet. Start by adding observations."), unsafe_allow_html=True)

st.markdown("---")

# ===============================
# PAGE: AUDIT ENTRY
# ===============================

if page == "Audit Entry":

    st.markdown(section_header(_icon("edit", 20), "Audit Entry", "Record new audit observations with AI-powered classification"), unsafe_allow_html=True)

    col1, col2 = st.columns([1.1, 0.9])

    with col1:
        st.markdown("##### Audit Details")
        line = st.selectbox("Line", LINE_OPTIONS)

        # Dynamic station list: station_master (official) + audit_entries (ad-hoc)
        _master_stations = get_stations_for_line(line) if line else []
        _live_stations = []
        if not full_data.empty and "station" in full_data.columns:
            _live_stations = sorted({
                s.strip()
                for s in full_data["station"].dropna().astype(str)
                if s.strip()
            })
        _station_options = sorted(set(_master_stations + _live_stations))

        if st.session_state.get("station_selected") == "" and st.session_state.get("station_typed") == "":
            if "station_selectbox" in st.session_state:
                st.session_state.station_selectbox = None
            if "station_custom_input" in st.session_state:
                st.session_state.station_custom_input = ""
            st.session_state.station_selected = None
            st.session_state.station_typed = None

        st.markdown(
            '<div style="display: flex; align-items: center; margin-bottom: 4px; margin-top: 2px; font-size: 11.5px; font-weight: 600; color: #FAFAFA; font-family: \'Inter\', sans-serif; text-transform: uppercase; letter-spacing: 0.06em;">STATION</div>',
            unsafe_allow_html=True
        )

        _station_choice = st.selectbox(
            "Station",
            options=["➕ Add New Station..."] + _station_options,
            index=None,
            placeholder="Type 's' or 'station' for instant auto-suggestions...",
            key="station_selectbox",
            label_visibility="collapsed"
        )

        if _station_choice == "➕ Add New Station...":
            station = st.text_input(
                "Custom Station Name",
                placeholder="Type custom station name here...",
                key="station_custom_input",
                label_visibility="collapsed"
            )
        else:
            station = _station_choice or ""

        st.session_state.station_typed = station

        # ── Dynamic FLM: dropdown from live data, fallback to text input ──
        _live_flms = sorted(
            full_data["flm_name"].dropna().astype(str).str.strip().unique()
        ) if not full_data.empty and "flm_name" in full_data.columns else []
        _live_flms = [f for f in _live_flms if f]
        if _live_flms:
            supervisor = st.selectbox("FLM", ["➕ New FLM..."] + _live_flms, index=None, placeholder="Select FLM...")
            if supervisor == "➕ New FLM...":
                supervisor = st.text_input("New FLM Name", placeholder="Enter FLM / supervisor name", label_visibility="collapsed")
        else:
            supervisor = st.text_input("FLM", placeholder="Enter FLM / supervisor name")
        shift = st.selectbox("Shift", ["Shift 1", "Shift 2", "Shift 3"])
        audit_date = st.date_input("Date")

        # ── Dynamic Auditor: dropdown from live data, fallback to text input ──
        _live_auditors = sorted(
            full_data["auditor_name"].dropna().astype(str).str.strip().unique()
        ) if not full_data.empty and "auditor_name" in full_data.columns else []
        _live_auditors = [a for a in _live_auditors if a]  # remove blanks
        if _live_auditors:
            auditor = st.selectbox("Auditor", ["➕ New Auditor..."] + _live_auditors, index=None, placeholder="Select auditor...")
            if auditor == "➕ New Auditor...":
                auditor = st.text_input("New Auditor Name", placeholder="Enter auditor name", label_visibility="collapsed")
        else:
            auditor = st.text_input("Auditor", placeholder="Enter auditor name")

        # ── Dynamic Severity & Category dropdowns ──
        _sev_options = get_severity_options()
        if _sev_options:
            severity = st.selectbox("Severity", _sev_options)
        else:
            st.warning("Severity master sheet is empty — please populate it.")
            severity = st.text_input("Severity", placeholder="e.g. Low, Medium, High, Critical")

        _cat_options = get_category_options()
        if _cat_options:
            category = st.selectbox("Category", _cat_options)
        else:
            st.warning("Category master sheet is empty — please populate it.")
            category = st.text_input("Category", placeholder="e.g. Process, Safety, Quality")

        area = st.text_input("Area", placeholder="e.g. Assembly, Paint, Weld")

    with col2:
        st.markdown("##### Capture Evidence")
        obs_text = st.text_area(
            "Observation",
            key="obs_input_main",
            placeholder="Describe what you observed...\nInclude: location, specific issue, impact",
            height=160
        )
        st.caption("Include: location, issue, impact")
        remarks = st.text_input("Remarks", placeholder="Optional remarks or additional context")

        uploaded_file = st.file_uploader(
            "Upload Image",
            type=["jpg", "png", "jpeg"],
            key="upload_main",
            help="Attach a photo of the observation"
        )
        if uploaded_file is not None:
            st.image(uploaded_file, caption="Uploaded preview", width=160)

        if "show_camera" not in st.session_state:
            st.session_state.show_camera = False
        if st.checkbox("Enable Camera", value=st.session_state.show_camera, key="cam_toggle"):
            camera_image = st.camera_input("Capture Image", key="camera_main")
            if camera_image is not None:
                st.caption("Camera image captured")
        else:
            camera_image = None

        if st.button("Add Observation", use_container_width=True):
            if not obs_text.strip():
                st.warning("Please enter an observation description.")
            elif len(obs_text.strip()) < 10:
                st.warning("Observation too short — please be more specific.")
            else:
                image_bytes = None
                if camera_image is not None:
                    image_bytes = camera_image.getvalue()
                elif uploaded_file is not None:
                    image_bytes = uploaded_file.read()

                st.session_state.observations.append({
                    "text": obs_text.strip(),
                    "image": image_bytes
                })
                st.rerun()

        if st.session_state.observations:
            st.markdown(section_header(_icon("list", 18), f"Observations ({len(st.session_state.observations)})"), unsafe_allow_html=True)
            for i, obs in enumerate(st.session_state.observations):
                st.markdown(
                    observation_display_card(
                        index=i + 1,
                        text=obs['text'],
                        has_image=obs.get('image') is not None
                    ),
                    unsafe_allow_html=True
                )
                col_img, col_del = st.columns([10, 1])
                with col_img:
                    if obs.get("image") is not None:
                        st.image(obs["image"], width=100)
                with col_del:
                    # Sleek trash icon button to avoid vertical wrapping in small columns
                    if st.button("🗑️", key=f"delete_{i}", help="Remove this observation"):
                        st.session_state.observations.pop(i)
                        st.rerun()
        else:
            st.markdown(empty_state("No observations added yet. Use the form above to add your first observation."), unsafe_allow_html=True)

        st.markdown("---")
        col_left, col_center, col_right = st.columns([1, 2, 1])
        with col_center:
            if st.button("Submit Full Audit", type="primary", use_container_width=True, disabled=st.session_state.get("submitting", False)):
                st.session_state.submitting = True

                if not st.session_state.observations:
                    st.warning("Add observations first")
                    st.session_state.submitting = False
                else:
                    total = len(st.session_state.observations)
                    progress = st.progress(0)

                    with st.spinner("Processing..."):
                        for i, obs in enumerate(st.session_state.observations):
                            try:
                                principle = detect_principle(obs["text"])
                                image_data = obs.get("image", None)

                                # Encode image to base64 string if present (validate first)
                                if image_data is not None:
                                    try:
                                        from PIL import Image
                                        from io import BytesIO
                                        img = Image.open(BytesIO(image_data))
                                        if img.mode != 'RGB':
                                            img = img.convert('RGB')
                                        img.thumbnail((400, 400))
                                        buf = BytesIO()
                                        img.save(buf, format="JPEG", quality=65)
                                        compressed_data = buf.getvalue()
                                        image_b64 = base64.b64encode(compressed_data).decode("utf-8")
                                    except Exception as e:
                                        logger.warning(f"Image encoding failed: {e}")
                                        image_b64 = ""
                                else:
                                    image_b64 = ""

                                add_audit_entry({
                                    "line": line,
                                    "station": station,
                                    "station_no": get_station_no(line, station) if station else "",
                                    "area": area.strip() if area else "",
                                    "supervisor": supervisor,
                                    "auditor": auditor.strip() if auditor else "",
                                    "observation_text": obs["text"],
                                    "ai_principle": principle,
                                    "severity": severity if severity else "",
                                    "category": category if category else "",
                                    "remarks": remarks.strip() if remarks else "",
                                    "audit_date": str(audit_date),
                                    "shift": shift,
                                    "image_base64": image_b64,
                                })
                            except Exception as e:
                                st.error(f"Error saving observation: {e}")

                            progress.progress((i + 1) / total)

                    st.toast("Audit submitted successfully", icon="✅")
                    
                    # Invalidate summary caches to ensure real-time reporting
                    if "agent_outputs" in st.session_state:
                        for k in ["daily", "weekly", "monthly", "mail"]:
                            st.session_state.agent_outputs.pop(k, None)
                            
                    st.session_state.observations = []
                    st.session_state.station_selected = ""
                    st.session_state.station_typed = ""
                    st.session_state.submitting = False
                    st.rerun()

# ===============================
# PAGE: SUMMARY
# ===============================

elif page == "Summary":

    st.markdown(section_header(_icon("summary", 20), "AI Summary Agent", "Generate AI-powered daily, weekly, and monthly audit briefs"), unsafe_allow_html=True)

    # ── Strict Date Filters for Live Summary Engine ──
    now = pd.Timestamp.now()
    
    # 1. Daily: Strictly today
    daily_df = master_canonical_df[master_canonical_df["audit_date"].dt.date == now.date()].copy() if not master_canonical_df.empty else master_canonical_df.copy()
    
    # 2. Weekly: Strictly current ISO week
    iso_week = now.isocalendar().week
    iso_year = now.isocalendar().year
    weekly_df = master_canonical_df[
        (master_canonical_df["audit_date"].dt.isocalendar().week == iso_week) &
        (master_canonical_df["audit_date"].dt.isocalendar().year == iso_year)
    ].copy() if not master_canonical_df.empty else master_canonical_df.copy()
    
    # Previous Weekly: For trend calculations
    prev_week_date = now - pd.Timedelta(weeks=1)
    prev_iso_week = prev_week_date.isocalendar().week
    prev_iso_year = prev_week_date.isocalendar().year
    prev_weekly_df = master_canonical_df[
        (master_canonical_df["audit_date"].dt.isocalendar().week == prev_iso_week) &
        (master_canonical_df["audit_date"].dt.isocalendar().year == prev_iso_year)
    ].copy() if not master_canonical_df.empty else master_canonical_df.copy()

    # 3. Monthly: Strictly current calendar month
    monthly_df = master_canonical_df[
        (master_canonical_df["audit_date"].dt.month == now.month) &
        (master_canonical_df["audit_date"].dt.year == now.year)
    ].copy() if not master_canonical_df.empty else master_canonical_df.copy()

    # Previous Monthly: For trend calculations
    prev_month_date = now.replace(day=1) - pd.Timedelta(days=1)
    prev_monthly_df = master_canonical_df[
        (master_canonical_df["audit_date"].dt.month == prev_month_date.month) &
        (master_canonical_df["audit_date"].dt.year == prev_month_date.year)
    ].copy() if not master_canonical_df.empty else master_canonical_df.copy()

    col1, col2, col3 = st.columns(3)
    with col1:
        if st.button("Daily Summary"):
            with st.spinner("Generating Daily Summary..."):
                st.session_state.agent_outputs["daily"] = generate_daily_brief(daily_df)
    with col2:
        if st.button("Weekly Summary"):
            with st.spinner("Generating Weekly Summary..."):
                st.session_state.agent_outputs["weekly"] = generate_weekly_brief(weekly_df, prev_weekly_df)
    with col3:
        if st.button("Monthly Summary"):
            with st.spinner("Generating Monthly Summary..."):
                st.session_state.agent_outputs["monthly"] = generate_monthly_brief(monthly_df, prev_monthly_df)

    st.markdown("---")

    if "daily" in st.session_state.agent_outputs:
        st.markdown(structured_ai_card("Daily Summary", st.session_state.agent_outputs["daily"]), unsafe_allow_html=True)
        if st.button("Generate Daily Mail"):
            with st.spinner("Formatting email..."):
                st.session_state.agent_outputs["mail"] = generate_mail_from_summary(
                    st.session_state.agent_outputs["daily"], "Daily"
                )
                st.session_state.agent_outputs["mail_type"] = "Daily"

    if "weekly" in st.session_state.agent_outputs:
        st.markdown(structured_ai_card("Weekly Summary", st.session_state.agent_outputs["weekly"]), unsafe_allow_html=True)
        if st.button("Generate Weekly Mail"):
            with st.spinner("Formatting email..."):
                st.session_state.agent_outputs["mail"] = generate_mail_from_summary(
                    st.session_state.agent_outputs["weekly"], "Weekly"
                )
                st.session_state.agent_outputs["mail_type"] = "Weekly"

    if "monthly" in st.session_state.agent_outputs:
        st.markdown(structured_ai_card("Monthly Summary", st.session_state.agent_outputs["monthly"]), unsafe_allow_html=True)
        if st.button("Generate Monthly Mail"):
            with st.spinner("Formatting email..."):
                st.session_state.agent_outputs["mail"] = generate_mail_from_summary(
                    st.session_state.agent_outputs["monthly"], "Monthly"
                )
                st.session_state.agent_outputs["mail_type"] = "Monthly"

    if "mail" in st.session_state.agent_outputs:
        st.markdown("---")
        st.markdown(section_header(_icon("summary", 20), "Mail Preview", "Review and edit before sending"), unsafe_allow_html=True)

        mail_left, mail_right = st.columns([1, 1])
        with mail_left:
            st.markdown("##### Edit Mail")
            edited_mail = st.text_area(
                "Edit Mail Content",
                value=st.session_state.agent_outputs["mail"],
                height=350,
                label_visibility="collapsed"
            )
            st.session_state.agent_outputs["mail"] = edited_mail
        with mail_right:
            st.markdown("##### Preview")
            st.markdown(mail_preview_card(edited_mail), unsafe_allow_html=True)

        _active_type = st.session_state.agent_outputs.get("mail_type", "")
        _mail_subject = f"{_active_type} Manufacturing Audit Summary | {date.today().strftime('%d %b %Y')}"
        for _ml in edited_mail.strip().split("\n"):
            _ml_s = _ml.strip()
            if _ml_s.lower().startswith("**subject:**") or _ml_s.lower().startswith("subject:"):
                _extracted = _ml_s.split(":", 1)[-1].strip().strip("*").strip()
                if len(_extracted) > 5:
                    _mail_subject = _extracted
                break

        subject = _mail_subject
        eml_content = f"""Subject: {subject}\nTo: Quality Team\nFrom: AutoNQ AI <noreply@autonq.bosch.com>\nDate: {date.today().strftime('%a, %d %b %Y')}\nContent-Type: text/plain; charset=utf-8\n\n{edited_mail}\n"""
        file_name = f"audit_mail_{_active_type.lower()}_{date.today()}.eml"

        try:
            onedrive_path = os.path.expanduser("~/OneDrive - Bosch/Audit_Mails")
            os.makedirs(onedrive_path, exist_ok=True)
            file_path = os.path.join(onedrive_path, file_name)
            with open(file_path, "w") as f:
                f.write(eml_content)
            st.success(f"Mail auto-saved: {file_name}")
        except Exception as e:
            logger.warning(f"OneDrive save failed: {e}")
            st.info("OneDrive not available — use download button below.")

        st.download_button(
            label="Download Mail (.eml)",
            data=eml_content,
            file_name=file_name,
            mime="message/rfc822"
        )

        # ══════════════════════════════════════════════════════════════════════
        # OUTLOOK MAIL SENDER  —  Enterprise People Picker
        # ══════════════════════════════════════════════════════════════════════
        st.markdown("---")
        st.markdown(
            section_header(
                _icon("mail", 20),
                "Send via Outlook",
                "Deliver this audit summary directly through Microsoft Graph API"
            ),
            unsafe_allow_html=True
        )

        if _MAIL_SERVICE_AVAILABLE:

            # ── Scoped CSS ────────────────────────────────────────────────────
            st.markdown("""<style>
            .nq-rcpt-label{display:flex;align-items:center;gap:8px;margin-bottom:10px;
              font-size:11.5px;font-weight:700;color:#9BADB8;text-transform:uppercase;
              letter-spacing:0.08em;}
            .nq-rcpt-label svg{opacity:0.7;}
            .nq-rcpt-empty{background:rgba(129,195,215,0.04);
              border:1px dashed rgba(129,195,215,0.12);border-radius:10px;
              padding:14px 18px;text-align:center;color:#5E7A8A;
              font-size:13px;font-family:'Inter',sans-serif;margin:6px 0 10px;}
            .nq-rcpt-unavailable{background:rgba(224,168,77,0.06);
              border:1px dashed rgba(224,168,77,0.18);border-radius:10px;
              padding:14px 18px;text-align:center;color:#9A7A3A;
              font-size:13px;font-family:'Inter',sans-serif;margin:6px 0 10px;}
            .nq-rcpt-card{display:flex;align-items:center;gap:14px;
              background:rgba(129,195,215,0.06);border:1px solid rgba(129,195,215,0.18);
              border-radius:10px;padding:12px 16px;margin:8px 0 14px;
              transition:border-color 0.2s;}
            .nq-rcpt-card:hover{border-color:rgba(129,195,215,0.35);}
            .nq-rcpt-avatar{width:38px;height:38px;border-radius:50%;
              background:linear-gradient(135deg,#3A7CA5,#81C3D7);
              display:flex;align-items:center;justify-content:center;flex-shrink:0;
              font-size:14px;font-weight:700;color:#fff;letter-spacing:0.02em;
              font-family:'Inter',sans-serif;}
            .nq-rcpt-info{flex:1;min-width:0;}
            .nq-rcpt-name{font-size:14px;font-weight:600;color:#D9DCD6;
              font-family:'Inter',sans-serif;margin-bottom:2px;
              white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}
            .nq-rcpt-email{font-size:12.5px;color:#81C3D7;font-family:'Inter',sans-serif;
              white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}
            .nq-rcpt-role{font-size:11.5px;color:#5E7A8A;font-family:'Inter',sans-serif;
              margin-top:1px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}
            .nq-rcpt-badge{display:inline-block;background:rgba(110,203,160,0.12);
              border:1px solid rgba(110,203,160,0.25);border-radius:6px;
              padding:2px 8px;font-size:10.5px;font-weight:600;color:#6ECBA0;
              text-transform:uppercase;letter-spacing:0.06em;margin-left:auto;flex-shrink:0;}
            .nq-manual-label{font-size:11px;font-weight:600;color:#5E7A8A;
              text-transform:uppercase;letter-spacing:0.07em;margin-bottom:6px;
              font-family:'Inter',sans-serif;}
            </style>""", unsafe_allow_html=True)

            st.markdown(
                '<div class="nq-rcpt-label">'
                + _icon("mail", 15, "#9BADB8")
                + ' Recipient'
                + '</div>',
                unsafe_allow_html=True,
            )

            _ppl_q = st.text_input(
                "Search recipient",
                placeholder="Search by name or email (e.g. @A-Z Bosch.com)",
                key="people_search_q",
                label_visibility="collapsed",
            )

            _query = _ppl_q.strip() if _ppl_q else ""

            if len(_query) < 2:
                recipient = ""

            else:
                _graph_available = _search_users_fn is not None
                _ppl_results = []
                _graph_error = False

                if _graph_available:
                    try:
                        _ppl_results = _cached_people_search(_query)
                    except Exception:
                        _graph_error = True
                else:
                    _graph_error = True

                if _graph_error:
                    st.markdown(
                        '<div class="nq-rcpt-unavailable">'
                        '⚠️ &nbsp;Directory search unavailable — enter recipient email manually'
                        '</div>',
                        unsafe_allow_html=True,
                    )
                    st.markdown(
                        '<div class="nq-manual-label">Recipient Email</div>',
                        unsafe_allow_html=True,
                    )
                    recipient = st.text_input(
                        "Recipient Email",
                        value="",
                        placeholder="e.g. john.doe@bosch.com",
                        key="outlook_recipient_fallback_unavailable",
                        label_visibility="collapsed",
                    )

                elif not _ppl_results:
                    st.markdown(
                        '<div class="nq-rcpt-empty">'
                        '🔍 &nbsp;No matching Bosch employees found — enter recipient email manually'
                        '</div>',
                        unsafe_allow_html=True,
                    )
                    st.markdown(
                        '<div class="nq-manual-label">Recipient Email</div>',
                        unsafe_allow_html=True,
                    )
                    recipient = st.text_input(
                        "Recipient Email",
                        value="",
                        placeholder="e.g. john.doe@bosch.com",
                        key="outlook_recipient_fallback_noresult",
                        label_visibility="collapsed",
                    )

                else:
                    _ppl_labels = [
                        f"{r['displayName']}  ·  {r['mail']}"
                        + (f"  ·  {r['jobTitle']}" if r.get("jobTitle") else "")
                        for r in _ppl_results
                    ]
                    _ppl_idx = st.selectbox(
                        "Select recipient from directory",
                        range(len(_ppl_labels)),
                        format_func=lambda i: _ppl_labels[i],
                        key="people_result_select",
                        label_visibility="collapsed",
                    )
                    _sel = _ppl_results[_ppl_idx]
                    recipient = _sel["mail"]

                    _name_parts = _sel["displayName"].split()
                    _initials = (
                        _name_parts[0][0]
                        + (_name_parts[-1][0] if len(_name_parts) > 1 else "")
                    ).upper()

                    if _sel.get("jobTitle") and _sel.get("department"):
                        _role_html = f'<div class="nq-rcpt-role">{_sel["jobTitle"]}  ·  {_sel["department"]}</div>'
                    elif _sel.get("jobTitle"):
                        _role_html = f'<div class="nq-rcpt-role">{_sel["jobTitle"]}</div>'
                    elif _sel.get("department"):
                        _role_html = f'<div class="nq-rcpt-role">{_sel["department"]}</div>'
                    else:
                        _role_html = ""

                    st.markdown(
                        f'<div class="nq-rcpt-card">'
                        f'  <div class="nq-rcpt-avatar">{_initials}</div>'
                        f'  <div class="nq-rcpt-info">'
                        f'    <div class="nq-rcpt-name">{_sel["displayName"]}</div>'
                        f'    <div class="nq-rcpt-email">{recipient}</div>'
                        f'    {_role_html}'
                        f'  </div>'
                        f'  <div class="nq-rcpt-badge">Selected</div>'
                        f'</div>',
                        unsafe_allow_html=True,
                    )

            st.markdown('<div style="height:6px"></div>', unsafe_allow_html=True)
            _valid_recipient = bool(recipient and "@" in recipient)
            _bcol_l, _bcol_c, _bcol_r = st.columns([1, 2, 1])
            with _bcol_c:
                if st.button(
                    "📨  Send via Outlook" if _valid_recipient else "Enter a recipient to send",
                    key="outlook_send_btn",
                    use_container_width=True,
                    disabled=not _valid_recipient,
                    type="primary" if _valid_recipient else "secondary",
                ):
                    with st.spinner("Sending mail via Microsoft Outlook…"):
                        try:
                            status, response = _send_mail_fn(recipient, subject, edited_mail)
                            if status == 202:
                                st.success(f"✅ Mail sent successfully to {recipient}!")
                                logger.info(f"Outlook mail sent to {recipient} — HTTP {status}")
                            else:
                                st.error(f"❌ Mail delivery failed: {response}")
                                logger.warning(f"Outlook send failed — HTTP {status}: {response}")
                        except Exception as _mail_exc:
                            st.error(f"❌ Unexpected error while sending mail: {_mail_exc}")
                            logger.error(f"Outlook send exception: {_mail_exc}")

        else:
            st.markdown(
                '<div style="background:rgba(224,168,77,0.08);border:1px solid rgba(224,168,77,0.2);'
                'border-left:3px solid #E0A84D;border-radius:10px;padding:14px 18px;'
                'font-family:\'Inter\',sans-serif;font-size:13.5px;color:#D9DCD6;">'
                '⚠️ <strong style="color:#E0A84D;">mail_service.py not available.</strong> '
                'Ensure the file is present in your project root and the required Microsoft Graph '
                'credentials are configured in your <code>.env</code> file.</div>',
                unsafe_allow_html=True
            )

# ===============================
# PAGE: AUDIT PLAN
# ===============================

elif page == "Audit Plan":

    st.markdown(section_header(_icon("plan", 20), "AI Audit Planning Agent", "Risk-based audit scheduling powered by deviation data"), unsafe_allow_html=True)

    if st.button("Generate Smart Audit Plan"):
        with st.spinner("Analyzing risk and building audit plan..."):
            dev_c = pd.DataFrame(current_data).groupby("line").size().reset_index(name="count_dev_c") if not current_data.empty else pd.DataFrame(columns=["line", "count_dev_c"])
            dev_c.columns = ["LINE", "count_dev_c"]

            all_lines = pd.DataFrame({"LINE": _get_live_lines()})
            combined_line = pd.merge(all_lines, dev_c, on="LINE", how="left").fillna(0)
            combined_line["count_iqis_c"] = 0
            combined_line["risk_score"] = combined_line["count_iqis_c"] * 0.5 + combined_line["count_dev_c"] * 0.5
            classified = classify_by_percentile(combined_line)
            plan = generate_governance_annual_plan(classified)
            st.session_state.agent_outputs["plan"] = plan
            st.session_state.agent_outputs["classified"] = classified

    st.markdown("---")

    if "classified" in st.session_state.agent_outputs:
        df = st.session_state.agent_outputs["classified"]
        col1, col2, col3 = st.columns(3)
        with col1:
            st.markdown(kpi_card(_icon("alert", 20, "red"), "Very High Risk", len(df[df["risk_level"] == "Very High"]), "rose"), unsafe_allow_html=True)
        with col2:
            st.markdown(kpi_card(_icon("alert", 20, "orange"), "High + Medium", len(df[df["risk_level"].isin(["High", "Medium"])]), "amber"), unsafe_allow_html=True)
        with col3:
            st.markdown(kpi_card(_icon("check", 20, "green"), "Stable Lines", len(df[df["risk_level"] == "Stable"]), "emerald"), unsafe_allow_html=True)

    if "plan" in st.session_state.agent_outputs:
        st.markdown(section_header(_icon("calendar", 20), "Annual Audit Plan"), unsafe_allow_html=True)
        _plan_df = st.session_state.agent_outputs["plan"]
        if isinstance(_plan_df, pd.DataFrame) and not _plan_df.empty:
            _render_download_bar(
                _plan_df,
                f"AuditPlan_{date.today()}",
                excel=True, csv=True,
                key_prefix="dl_auditplan",
            )
        _df_wrap_open("Annual Audit Plan")
        st.dataframe(st.session_state.agent_outputs["plan"], use_container_width=True)
        _df_wrap_close()

# ===============================
# PAGE: DAILY Q-CHECK
# ===============================

elif page == "Daily Q-Check":

    st.markdown(section_header(_icon("check", 20), "AI-Based Daily Check", "Smart audit checkpoints generated from deviation history"), unsafe_allow_html=True)

    col1, col2, col3 = st.columns(3)
    with col1:
        selected_line = st.selectbox("Select Line", LINE_OPTIONS)
    with col2:
        selected_date = st.date_input("Select Date")
    with col3:
        _live_auditors = sorted(full_data["auditor_name"].dropna().astype(str).str.strip().unique()) if not full_data.empty and "auditor_name" in full_data.columns else []
        _live_auditors = [a for a in _live_auditors if a]  # remove blanks
        if _live_auditors:
            selected_auditor = st.selectbox("Select Auditor", _live_auditors)
        else:
            selected_auditor = st.text_input("Auditor Name", placeholder="Enter auditor name")

    st.markdown("---")

    if st.button("Generate Q-Check"):
        with st.spinner("Generating smart audit checkpoints..."):
            _target_date = str(selected_date).strip()
            _filtered_records = []
            if not full_data.empty and "audit_date" in full_data.columns:
                # Safely normalise excel datetime to string YYYY-MM-DD
                _df_dates = pd.to_datetime(full_data["audit_date"], errors="coerce").dt.strftime('%Y-%m-%d')
                _date_mask = _df_dates == _target_date
                _filtered_records = full_data[_date_mask].to_dict(orient="records")
            
            records = generate_qcheck_questions(selected_line, _filtered_records)
            if not records:
                st.warning("No deviation data found for selected line.")
            else:
                st.session_state.qcheck_data = records

    if "qcheck_data" in st.session_state:
        st.markdown(section_header(_icon("check", 20), "Fill Daily Q-Check"), unsafe_allow_html=True)

        # ══════════════════════════════════════════════════════════════════════
        # FULLSCREEN IMAGE VIEWER  (robust single-component approach)
        #
        # Architecture:
        #   - streamlit.components.v1.html() renders ONE unsandboxed iframe
        #   - That iframe contains ALL CSS, modal HTML, and JavaScript
        #   - JS injects the modal overlay into window.parent.document.body
        #   - A MutationObserver watches for .nq-qcheck-thumb thumbnails
        #     appearing anywhere in the parent DOM (even after Streamlit reruns)
        #   - Click listeners are attached directly to each thumbnail
        #
        # WHY THIS WORKS:
        #   - components.v1.html() is guaranteed unsandboxed with JS execution
        #   - Injecting the modal into parent.document.body means it persists
        #     across the Streamlit iframe boundary
        #   - MutationObserver ensures thumbnails rendered AFTER this component
        #     still get click handlers attached
        # ══════════════════════════════════════════════════════════════════════

        import streamlit.components.v1 as components

        _VIEWER_HTML = """
        <script>
        (function() {
            var pd;
            try { pd = window.parent.document; } catch(e) { console.warn('[NQ Viewer] Cannot access parent document'); return; }
            if (!pd) return;

            /* ── Prevent duplicate initialisation ── */
            if (pd.getElementById('nqImgViewer') && pd.getElementById('nqImgViewer').getAttribute('data-nq-ready') === '2') {
                /* Already fully initialised — just re-scan for new thumbs */
                attachThumbListeners();
                return;
            }

            /* ── Remove stale overlay from previous render ── */
            var old = pd.getElementById('nqImgViewer');
            if (old) old.remove();
            var oldStyle = pd.getElementById('nqViewerCSS');
            if (oldStyle) oldStyle.remove();

            /* ── Inject CSS into parent <head> ── */
            var css = pd.createElement('style');
            css.id = 'nqViewerCSS';
            css.textContent = `
                .nq-img-viewer-overlay {
                    position: fixed; top: 0; left: 0;
                    width: 100vw; height: 100vh;
                    background: rgba(0,0,0,0.96);
                    display: none;
                    align-items: center; justify-content: center;
                    z-index: 999999;
                    font-family: 'Inter', 'Segoe UI', sans-serif;
                }
                .nq-img-viewer-overlay.nq-active { display: flex !important; }
                .nq-img-viewer-img {
                    max-width: 92vw; max-height: 88vh;
                    object-fit: contain;
                    cursor: grab;
                    transition: transform 0.08s ease-out;
                    border-radius: 4px;
                    box-shadow: 0 8px 60px rgba(0,0,0,0.6);
                }
                .nq-img-viewer-img:active { cursor: grabbing; }
                .nq-img-viewer-toolbar {
                    position: absolute; bottom: 28px; left: 50%;
                    transform: translateX(-50%);
                    display: flex; align-items: center; gap: 4px;
                    background: rgba(19,45,68,0.92);
                    backdrop-filter: blur(20px); -webkit-backdrop-filter: blur(20px);
                    border: 1px solid rgba(129,195,215,0.25);
                    border-radius: 14px;
                    padding: 8px 14px;
                    box-shadow: 0 8px 32px rgba(0,0,0,0.5);
                    z-index: 1000001;
                }
                .nq-img-viewer-toolbar button {
                    background: rgba(129,195,215,0.08);
                    border: 1px solid rgba(129,195,215,0.15);
                    color: #D9DCD6;
                    font-size: 13px; font-weight: 600;
                    padding: 8px 14px;
                    border-radius: 8px;
                    cursor: pointer;
                    transition: all 0.2s;
                    font-family: 'Inter', sans-serif;
                    display: flex; align-items: center; gap: 6px;
                    white-space: nowrap;
                }
                .nq-img-viewer-toolbar button:hover {
                    background: rgba(129,195,215,0.2);
                    border-color: rgba(129,195,215,0.4);
                    color: #FFFFFF;
                }
                .nq-img-viewer-toolbar button:active {
                    background: rgba(129,195,215,0.3);
                    transform: scale(0.96);
                }
                .nq-img-viewer-toolbar .nq-zoom-display {
                    color: #81C3D7; font-size: 12px; font-weight: 700;
                    min-width: 52px; text-align: center;
                    padding: 0 6px;
                    letter-spacing: 0.03em;
                }
                .nq-img-viewer-toolbar .nq-divider {
                    width: 1px; height: 24px;
                    background: rgba(129,195,215,0.15);
                    margin: 0 6px;
                }
                .nq-img-viewer-close {
                    position: absolute; top: 20px; right: 28px;
                    width: 44px; height: 44px;
                    background: rgba(19,45,68,0.85);
                    backdrop-filter: blur(12px);
                    border: 1px solid rgba(129,195,215,0.2);
                    border-radius: 12px;
                    color: #D9DCD6;
                    font-size: 22px; font-weight: 300;
                    cursor: pointer;
                    display: flex; align-items: center; justify-content: center;
                    transition: all 0.2s;
                    z-index: 1000001;
                }
                .nq-img-viewer-close:hover {
                    background: rgba(217,96,90,0.25);
                    border-color: rgba(217,96,90,0.5);
                    color: #E8847E;
                }
                .nq-img-viewer-info {
                    position: absolute; top: 22px; left: 28px;
                    background: rgba(19,45,68,0.85);
                    backdrop-filter: blur(12px);
                    border: 1px solid rgba(129,195,215,0.15);
                    border-radius: 10px;
                    padding: 8px 16px;
                    color: #9BADB8; font-size: 11.5px; font-weight: 600;
                    letter-spacing: 0.04em;
                    z-index: 1000001;
                    font-family: 'Inter', sans-serif;
                }
                .nq-qcheck-thumb {
                    width: 120px; height: 80px;
                    object-fit: cover;
                    border-radius: 8px;
                    cursor: zoom-in;
                    border: 1.5px solid rgba(129,195,215,0.15);
                    transition: all 0.25s cubic-bezier(0.2,0.8,0.2,1);
                    box-shadow: 0 2px 8px rgba(0,0,0,0.2);
                }
                .nq-qcheck-thumb:hover {
                    border-color: rgba(129,195,215,0.45);
                    box-shadow: 0 4px 20px rgba(129,195,215,0.15);
                    transform: scale(1.05);
                }
            `;
            pd.head.appendChild(css);

            /* ── Inject modal overlay into parent body ── */
            var overlay = pd.createElement('div');
            overlay.id = 'nqImgViewer';
            overlay.className = 'nq-img-viewer-overlay';
            overlay.innerHTML = `
                <button id="nqBtnCloseX" class="nq-img-viewer-close" title="Close (ESC)">✕</button>
                <div class="nq-img-viewer-info" id="nqImgViewerInfo">Reference Photo</div>
                <img id="nqImgViewerImg" class="nq-img-viewer-img" src="" alt="Reference Photo" draggable="false" />
                <div class="nq-img-viewer-toolbar">
                    <button id="nqBtnZoomIn" title="Zoom In">
                        <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><circle cx="11" cy="11" r="7"/><line x1="21" y1="21" x2="16.65" y2="16.65"/><line x1="11" y1="8" x2="11" y2="14"/><line x1="8" y1="11" x2="14" y2="11"/></svg>
                        Zoom In
                    </button>
                    <button id="nqBtnZoomOut" title="Zoom Out">
                        <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><circle cx="11" cy="11" r="7"/><line x1="21" y1="21" x2="16.65" y2="16.65"/><line x1="8" y1="11" x2="14" y2="11"/></svg>
                        Zoom Out
                    </button>
                    <div class="nq-divider"></div>
                    <div class="nq-zoom-display" id="nqZoomDisplay">100%</div>
                    <div class="nq-divider"></div>
                    <button id="nqBtnFit" title="Fit to Screen">
                        <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M8 3H5a2 2 0 0 0-2 2v3m18 0V5a2 2 0 0 0-2-2h-3m0 18h3a2 2 0 0 0 2-2v-3M3 16v3a2 2 0 0 0 2 2h3"/></svg>
                        Fit
                    </button>
                    <button id="nqBtnReset" title="Reset Zoom">
                        <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M3 12a9 9 0 1 0 9-9 9.75 9.75 0 0 0-6.74 2.74L3 8"/><path d="M3 3v5h5"/></svg>
                        Reset
                    </button>
                    <div class="nq-divider"></div>
                    <button id="nqBtnCloseBar" title="Close">
                        <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>
                        Close
                    </button>
                </div>
            `;
            pd.body.appendChild(overlay);

            /* ── Viewer state ── */
            var img    = pd.getElementById('nqImgViewerImg');
            var zoomEl = pd.getElementById('nqZoomDisplay');
            var infoEl = pd.getElementById('nqImgViewerInfo');

            var scale = 1, px = 0, py = 0, panning = false, sx = 0, sy = 0;

            function updZoom() { zoomEl.textContent = Math.round(scale * 100) + '%'; }
            function setTx()  { img.style.transform = 'translate(' + px + 'px,' + py + 'px) scale(' + scale + ')'; updZoom(); }

            function openV(src, label) {
                img.src = src;
                if (infoEl) infoEl.textContent = label || 'Reference Photo';
                scale = 1; px = 0; py = 0; setTx();
                overlay.classList.add('nq-active');
                pd.body.style.overflow = 'hidden';
            }

            function closeV() {
                overlay.classList.remove('nq-active');
                pd.body.style.overflow = '';
                img.src = '';
            }

            function zoomBy(f) {
                scale *= f;
                scale = Math.max(0.1, Math.min(20, scale));
                setTx();
            }

            function resetV() { scale = 1; px = 0; py = 0; setTx(); }

            /* ── Button listeners ── */
            pd.getElementById('nqBtnZoomIn').addEventListener('click', function() { zoomBy(1.3); });
            pd.getElementById('nqBtnZoomOut').addEventListener('click', function() { zoomBy(0.75); });
            pd.getElementById('nqBtnFit').addEventListener('click', resetV);
            pd.getElementById('nqBtnReset').addEventListener('click', resetV);
            pd.getElementById('nqBtnCloseX').addEventListener('click', closeV);
            pd.getElementById('nqBtnCloseBar').addEventListener('click', closeV);

            /* ── Click overlay background to close ── */
            overlay.addEventListener('click', function(e) {
                if (e.target === overlay) closeV();
            });

            /* ── Keyboard shortcuts ── */
            pd.addEventListener('keydown', function(e) {
                if (!overlay.classList.contains('nq-active')) return;
                if (e.key === 'Escape')            { closeV(); e.preventDefault(); }
                else if (e.key === '+' || e.key === '=') { zoomBy(1.2); e.preventDefault(); }
                else if (e.key === '-')             { zoomBy(0.8); e.preventDefault(); }
                else if (e.key === '0')             { resetV(); e.preventDefault(); }
            });

            /* ── Mouse wheel zoom (cursor-centered) ── */
            overlay.addEventListener('wheel', function(e) {
                if (!overlay.classList.contains('nq-active')) return;
                e.preventDefault();
                var xs = (e.clientX - px) / scale;
                var ys = (e.clientY - py) / scale;
                if ((e.wheelDelta ? e.wheelDelta : -e.deltaY) > 0) scale *= 1.15;
                else scale /= 1.15;
                scale = Math.max(0.1, Math.min(20, scale));
                px = e.clientX - xs * scale;
                py = e.clientY - ys * scale;
                setTx();
            }, { passive: false });

            /* ── Drag / pan (mouse) ── */
            img.addEventListener('mousedown', function(e) {
                e.preventDefault();
                sx = e.clientX - px; sy = e.clientY - py;
                panning = true;
            });
            pd.addEventListener('mouseup',  function() { panning = false; });
            pd.addEventListener('mousemove', function(e) {
                if (!panning) return; e.preventDefault();
                px = e.clientX - sx; py = e.clientY - sy;
                setTx();
            });

            /* ── Touch support (mobile auditors) ── */
            var lastTD = 0;
            img.addEventListener('touchstart', function(e) {
                if (e.touches.length === 1) {
                    sx = e.touches[0].clientX - px;
                    sy = e.touches[0].clientY - py;
                    panning = true;
                } else if (e.touches.length === 2) {
                    panning = false;
                    lastTD = Math.hypot(
                        e.touches[0].clientX - e.touches[1].clientX,
                        e.touches[0].clientY - e.touches[1].clientY);
                }
            }, { passive: true });
            img.addEventListener('touchmove', function(e) {
                e.preventDefault();
                if (e.touches.length === 1 && panning) {
                    px = e.touches[0].clientX - sx;
                    py = e.touches[0].clientY - sy;
                    setTx();
                } else if (e.touches.length === 2) {
                    var d = Math.hypot(
                        e.touches[0].clientX - e.touches[1].clientX,
                        e.touches[0].clientY - e.touches[1].clientY);
                    if (lastTD > 0) { scale *= d / lastTD; scale = Math.max(0.1, Math.min(20, scale)); setTx(); }
                    lastTD = d;
                }
            }, { passive: false });
            img.addEventListener('touchend', function() { panning = false; lastTD = 0; });

            /* ══════════════════════════════════════════════════════════════
             * THUMBNAIL CLICK ATTACHMENT
             * - Scans ALL iframes + parent doc for .nq-qcheck-thumb images
             * - Uses MutationObserver to catch thumbnails rendered AFTER init
             * - Marks attached thumbs with data-nq-click to avoid duplicates
             * ══════════════════════════════════════════════════════════════ */
            function attachThumbListeners() {
                /* Scan parent document */
                var thumbs = pd.querySelectorAll('img.nq-qcheck-thumb:not([data-nq-click])');
                thumbs.forEach(function(t) {
                    t.setAttribute('data-nq-click', '1');
                    t.addEventListener('click', function(e) {
                        e.preventDefault();
                        e.stopPropagation();
                        var src   = t.getAttribute('src') || t.src;
                        var label = t.getAttribute('title') || 'Reference Photo';
                        if (src) openV(src, label);
                    });
                });

                /* Also scan inside all iframes (Streamlit renders st.markdown in iframes) */
                try {
                    var iframes = pd.querySelectorAll('iframe');
                    iframes.forEach(function(iframe) {
                        try {
                            var idoc = iframe.contentDocument || iframe.contentWindow.document;
                            if (!idoc) return;
                            var iframeThumbs = idoc.querySelectorAll('img.nq-qcheck-thumb:not([data-nq-click])');
                            iframeThumbs.forEach(function(t) {
                                t.setAttribute('data-nq-click', '1');
                                t.addEventListener('click', function(e) {
                                    e.preventDefault();
                                    e.stopPropagation();
                                    var src   = t.getAttribute('src') || t.src;
                                    var label = t.getAttribute('title') || 'Reference Photo';
                                    if (src) openV(src, label);
                                });
                            });
                        } catch(ex) { /* cross-origin iframe, skip */ }
                    });
                } catch(ex) {}
            }

            /* Initial scan */
            attachThumbListeners();

            /* Re-scan whenever DOM changes (Streamlit re-renders) */
            var observer = new MutationObserver(function() {
                attachThumbListeners();
            });
            observer.observe(pd.body, { childList: true, subtree: true });

            /* Periodic fallback scan (catches late-loading iframe content) */
            setInterval(attachThumbListeners, 1500);

            overlay.setAttribute('data-nq-ready', '2');
            console.log('[NQ Viewer] Fully initialized with MutationObserver — click any thumbnail to zoom');
        })();
        </script>
        """

        components.html(_VIEWER_HTML, height=0, scrolling=False)

        h1, h2, h3, h4, h5 = st.columns([1, 3, 2, 1, 2])
        h1.markdown("**Station**")
        h2.markdown("**Checkpoint**")
        h3.markdown("**Ref Photo**")
        h4.markdown("**Status**")
        h5.markdown("**Remark**")
        st.markdown("---")

        updated_records = []
        for i, row in enumerate(st.session_state.qcheck_data):
            c1, c2, c3, c4, c5 = st.columns([1, 3, 2, 1, 2])
            with c1:
                st.write(row["Station"])
            with c2:
                st.write(row["Checkpoint"])
            with c3:
                # ── Robust base64 decode for Q-Check reference photos ──────
                # Excel cell storage can inject newlines, spaces, and broken
                # padding into long base64 strings. We sanitise before decoding.
                _raw_b64 = str(row.get("image_base64") or "").strip()
                if _raw_b64 and _raw_b64.lower() != "nan":
                    try:
                        # Strip data URI prefix if present
                        if "," in _raw_b64:
                            _raw_b64 = _raw_b64.split(",")[-1]

                        # Remove ALL whitespace/newlines Excel may have injected
                        _raw_b64 = "".join(_raw_b64.split())

                        # Fix missing base64 padding
                        missing_padding = len(_raw_b64) % 4
                        if missing_padding:
                            _raw_b64 += "=" * (4 - missing_padding)

                        _img_src = f"data:image/jpeg;base64,{_raw_b64}"
                        _img_label = f"{row.get('Station', '')} — {row.get('Checkpoint', '')[:60]}"
                        # Render thumbnail — NO onclick needed.
                        # JS engine uses delegated click on .nq-qcheck-thumb class.
                        # The 'title' attribute carries the label (preserved by sanitizer).
                        st.markdown(
                            f'<img src="{_img_src}" class="nq-qcheck-thumb" '
                            f'title="{_img_label}" />',
                            unsafe_allow_html=True
                        )
                    except Exception as _img_err:
                        st.caption("No Image")
                        logger.warning(f"Q-Check image decode failed at row {i}: {_img_err}")
                else:
                    st.caption("No Image")
                # ────────────────────────────────────────────────────────────────
            with c4:
                status = st.selectbox("Status", ["NOK", "OK"], key=f"status_{i}", label_visibility="collapsed")
            with c5:
                remark = st.text_input("Remark", key=f"remark_{i}", label_visibility="collapsed")
            updated_records.append({"Station": row["Station"], "Checkpoint": row["Checkpoint"], "Status": status, "Remark": remark})

        st.markdown("---")

        # ── Q-Check Download Buttons ──────────────────────────────────────
        if updated_records:
            _qcheck_export_df = pd.DataFrame(updated_records)
            _render_download_bar(
                _qcheck_export_df,
                f"QCheck_{selected_line}_{selected_date}",
                excel=True, csv=True,
                key_prefix="dl_qcheck",
            )
            st.markdown("---")
        # ─────────────────────────────────────────────────────────────────

        if st.button("Submit Q-Check", type="primary"):
            # Save NOK findings to Excel backend as audit entries (single source of truth)
            saved = 0
            for rec in updated_records:
                if rec.get("Status", "OK") == "NOK":
                    add_audit_entry({
                        "line": selected_line,
                        "station": rec.get("Station", ""),
                        "station_no": get_station_no(selected_line, rec.get("Station", "")) if rec.get("Station") else "",
                        "area": "",
                        "supervisor": "",
                        "auditor": selected_auditor if selected_auditor else "",
                        "observation_text": f"Q-Check NOK: {rec.get('Checkpoint', '')}",
                        "ai_principle": "Quality",
                        "severity": "",
                        "category": "Quality",
                        "remarks": rec.get("Remark", ""),
                        "audit_date": str(selected_date),
                        "shift": "",
                        "image_base64": "",
                    })
                    saved += 1
            if saved > 0:
                st.success(f"Q-Check submitted — {saved} NOK finding(s) saved to audit database")
            else:
                st.success("Q-Check submitted — all items OK, no findings to record")
            del st.session_state.qcheck_data

# ===============================
# PAGE: PROCESS AUDIT
# ===============================

elif page == "Process Audit":

    st.markdown(section_header(_icon("list", 20), "AI IATF-Based Process Audit", "Generate IATF-based process audit sheet from deviations"), unsafe_allow_html=True)

    line_p = st.selectbox("Select Line", LINE_OPTIONS, key="line_process")
    st.markdown("---")

    if st.button("Generate Process Audit Sheet"):
        with st.spinner("Analyzing deviations and building process audit sheet..."):
            canonical_recent = master_canonical_df[master_canonical_df["audit_date"] >= (pd.Timestamp.today() - pd.Timedelta(days=30))] if not master_canonical_df.empty else master_canonical_df.copy()
            records_to_use = canonical_recent.to_dict(orient="records") if not canonical_recent.empty else []
            df_process = generate_iatf_process_audit_sheet(line_p, records_to_use, iqis_df, lpc_df)
            if isinstance(df_process, str):
                records_to_use = master_canonical_df.to_dict(orient="records") if not master_canonical_df.empty else []
                df_process = generate_iatf_process_audit_sheet(line_p, records_to_use, iqis_df, lpc_df)
                if not isinstance(df_process, str):
                    st.warning("No recent deviations found — using historical data")
            else:
                st.success("Generated using recent deviation data")
            st.session_state.agent_outputs["process"] = df_process
            st.session_state.agent_outputs["process_line"] = line_p

    if "process" in st.session_state.agent_outputs:
        st.markdown(section_header(_icon("list", 20), f"Process Audit Sheet – {line_p}"), unsafe_allow_html=True)
        df_display = st.session_state.agent_outputs["process"]
        if isinstance(df_display, str):
            st.warning(df_display)
        else:
            _render_download_bar(
                df_display,
                f"ProcessAudit_{line_p}_{date.today()}",
                excel=True, csv=True,
                key_prefix="dl_process",
            )
            _df_wrap_open(f"Process Audit Sheet {line_p}")
            st.dataframe(df_display, use_container_width=True)
            _df_wrap_close()

# ===============================
# PAGE: FOLLOW-UP
# ===============================

elif page == "Follow-up":

    st.markdown(section_header(_icon("flag", 20), "AI Follow-up Checklist", "Track closure of audit findings and prevent recurrence"), unsafe_allow_html=True)

    line_f = st.selectbox("Select Line", LINE_OPTIONS, key="line_followup")
    st.markdown("---")

    if st.button("Generate Follow-up Checklist"):
        if "process" not in st.session_state.agent_outputs:
            st.warning("Please generate Process Audit first (Process Audit page)")
        else:
            _process_src = st.session_state.agent_outputs.get("process_line", None)
            if _process_src and str(_process_src) != str(line_f):
                st.warning(
                    f"Process Audit was generated for **{_process_src}** — "
                    f"regenerate it for **{line_f}** on the Process Audit page first."
                )
            else:
                with st.spinner("Building follow-up checklist..."):
                    df_process = st.session_state.agent_outputs["process"]
                    df_followup = generate_followup_checklist(line_f, df_process)
                    st.session_state.agent_outputs["followup"] = df_followup
                    st.session_state.agent_outputs["followup_line"] = line_f
                    st.success("Follow-up checklist generated")

    if "followup" in st.session_state.agent_outputs:
        st.markdown(section_header(_icon("flag", 20), f"Follow-up Checklist – {line_f}"), unsafe_allow_html=True)
        df_follow = st.session_state.agent_outputs["followup"]
        if isinstance(df_follow, str):
            st.warning(df_follow)
        else:
            _render_download_bar(
                df_follow,
                f"FollowUp_{line_f}_{date.today()}",
                excel=True, csv=True,
                key_prefix="dl_followup",
            )
            _df_wrap_open(f"Follow-up Checklist {line_f}")
            st.dataframe(df_follow, use_container_width=True)
            _df_wrap_close()

        st.markdown("---")
        col1, col2, col3 = st.columns(3)
        with col1:
            st.markdown(kpi_card(_icon("refresh", 20), "Total Issues", len(df_follow) if not isinstance(df_follow, str) else 0, "blue"), unsafe_allow_html=True)
        with col2:
            st.markdown(kpi_card(_icon("clock", 20), "Pending", len(df_follow) if not isinstance(df_follow, str) else 0, "amber"), unsafe_allow_html=True)
        with col3:
            st.markdown(kpi_card(_icon("check", 20), "Closed", 0, "emerald"), unsafe_allow_html=True)

# ===============================
# PAGE: TOP RECURRING ISSUES
# ===============================

elif page == "Top Recurring Issues":

    st.markdown(section_header(_icon("globe", 20), "Top Recurring Issues Intelligence Center", "AI-powered identification of recurring manufacturing deviations, quality risks, and process weaknesses."), unsafe_allow_html=True)

    most_repeated = "N/A"
    critical_issues = 0
    affected_stations = 0
    risk_level = "LOW"
    
    if not full_data.empty:
        if "observation_text" in full_data.columns and not full_data["observation_text"].dropna().empty:
            raw_issue = str(full_data["observation_text"].mode()[0])
            
            def extract_issue_name(text):
                t = str(text).split(":")[-1].strip()
                t = t.replace("Are ", "").replace("Is ", "").replace("Q-Check", "").replace("NOK", "").replace("OK", "").replace("-", "").strip()
                words = [w for w in t.split(" ") if w]
                t = " ".join(words[:2]).title()
                return t if t else "Unknown"
                
            clean_issue = extract_issue_name(raw_issue)
            # Enforce 34px max font size and two-line truncation
            most_repeated = f'<div style="font-size:30px; line-height:1.2; text-overflow:ellipsis; overflow:hidden; white-space:normal; display:-webkit-box; -webkit-line-clamp:2; -webkit-box-orient:vertical;">{clean_issue}</div>'
                
        if "severity" in full_data.columns:
            critical_issues = len(full_data[full_data["severity"].str.upper() == "CRITICAL"])
        if "station" in full_data.columns:
            affected_stations = full_data["station"].nunique()
            
        if critical_issues > 5:
            risk_level = "CRITICAL"
        elif critical_issues > 0:
            risk_level = "HIGH"
        elif len(full_data) > 20:
            risk_level = "MEDIUM"

    kpi1, kpi2, kpi3, kpi4 = st.columns(4)
    with kpi1:
        st.markdown(kpi_card(_icon("refresh", 20), "Most Repeated Issue", most_repeated, "amber"), unsafe_allow_html=True)
    with kpi2:
        st.markdown(kpi_card(_icon("alert", 20), "Critical Issues", critical_issues, "rose"), unsafe_allow_html=True)
    with kpi3:
        st.markdown(kpi_card(_icon("pin", 20), "Affected Stations", affected_stations, "cyan"), unsafe_allow_html=True)
    with kpi4:
        st.markdown(kpi_card(_icon("fire", 20), "Risk Level", risk_level, "rose" if risk_level in ["HIGH", "CRITICAL"] else "emerald"), unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    col_scope, col_ai = st.columns([1, 1])
    
    with col_scope:
        st.markdown('<div id="scope-anchor"></div>', unsafe_allow_html=True)
        
        st.markdown("""
        <style>
        /* === CONTAINER STYLING === */
        /* Fallback for browsers without :has() */
        div[data-testid="stHorizontalBlock"] > div[data-testid="column"]:first-child:nth-last-child(2) {
            background: rgba(19,45,68,0.45) !important;
            backdrop-filter: blur(8px) !important;
            border: 1px solid rgba(129,195,215,0.15) !important;
            border-radius: 12px !important;
            padding: 24px !important;
            box-shadow: 0 8px 32px rgba(10,25,41,0.5) !important;
        }
        
        /* Primary targeted styling */
        div[data-testid="column"]:has(#scope-anchor) {
            background: rgba(19,45,68,0.45) !important;
            backdrop-filter: blur(8px) !important;
            border: 1px solid rgba(129,195,215,0.15) !important;
            border-radius: 12px !important;
            padding: 24px !important;
            box-shadow: 0 8px 32px rgba(10,25,41,0.5) !important;
            height: 100% !important;
        }
        
        /* === BUTTON RESET === */
        div[data-testid="column"]:has(#scope-anchor) button:focus,
        div[data-testid="stHorizontalBlock"] > div[data-testid="column"]:first-child:nth-last-child(2) button:focus {
            outline: none !important; box-shadow: none !important;
        }
        
        /* === BASE CARD STYLING === */
        div[data-testid="column"]:has(#scope-anchor) button,
        div[data-testid="stHorizontalBlock"] > div[data-testid="column"]:first-child:nth-last-child(2) button {
            border-radius: 12px !important;
            min-height: 60px !important;
            height: auto !important;
            display: flex !important;
            flex-direction: column !important;
            align-items: flex-start !important;
            justify-content: center !important;
            padding: 12px 20px !important;
            cursor: pointer !important;
            transition: all 0.2s ease !important;
            width: 100% !important;
            margin-bottom: 8px !important;
        }
        
        /* Fix internal button alignment */
        div[data-testid="column"]:has(#scope-anchor) button div[data-testid="stMarkdownContainer"],
        div[data-testid="stHorizontalBlock"] > div[data-testid="column"]:first-child:nth-last-child(2) button div[data-testid="stMarkdownContainer"] {
            display: flex !important;
            flex-direction: column !important;
            align-items: flex-start !important;
            text-align: left !important;
            width: 100% !important;
        }
        
        div[data-testid="column"]:has(#scope-anchor) button p,
        div[data-testid="stHorizontalBlock"] > div[data-testid="column"]:first-child:nth-last-child(2) button p {
            font-weight: 600 !important;
            letter-spacing: 0.05em !important;
            font-size: 13px !important;
            margin: 0 !important;
            text-align: left !important;
        }
        
        /* === UNSELECTED (SECONDARY) === */
        div[data-testid="column"]:has(#scope-anchor) button[kind="secondary"],
        div[data-testid="stHorizontalBlock"] > div[data-testid="column"]:first-child:nth-last-child(2) button[kind="secondary"] {
            background: rgba(10,25,41,0.5) !important;
            border: 1px solid rgba(129,195,215,0.2) !important;
        }
        
        div[data-testid="column"]:has(#scope-anchor) button[kind="secondary"]:hover,
        div[data-testid="stHorizontalBlock"] > div[data-testid="column"]:first-child:nth-last-child(2) button[kind="secondary"]:hover {
            background: rgba(40,120,255,0.05) !important;
            border-color: rgba(77,163,255,0.5) !important;
            transform: translateX(4px) !important;
        }
        
        div[data-testid="column"]:has(#scope-anchor) button[kind="secondary"] p,
        div[data-testid="stHorizontalBlock"] > div[data-testid="column"]:first-child:nth-last-child(2) button[kind="secondary"] p {
            color: #9BADB8 !important;
        }
        
        /* === SELECTED (PRIMARY) === */
        div[data-testid="column"]:has(#scope-anchor) button[kind="primary"],
        div[data-testid="stHorizontalBlock"] > div[data-testid="column"]:first-child:nth-last-child(2) button[kind="primary"] {
            background: rgba(40,120,255,0.15) !important;
            border: 2px solid #4DA3FF !important;
            box-shadow: 0 0 20px rgba(77,163,255,0.35) !important;
            transform: scale(1.02) !important;
        }
        
        div[data-testid="column"]:has(#scope-anchor) button[kind="primary"] p,
        div[data-testid="stHorizontalBlock"] > div[data-testid="column"]:first-child:nth-last-child(2) button[kind="primary"] p {
            color: #FFFFFF !important;
            font-weight: 700 !important;
            text-shadow: 0 0 8px rgba(255,255,255,0.3) !important;
        }
        
        /* INJECT "Recommended" TEXT */
        div[data-testid="column"]:has(#scope-anchor) button[kind="primary"]::after,
        div[data-testid="stHorizontalBlock"] > div[data-testid="column"]:first-child:nth-last-child(2) button[kind="primary"]::after {
            content: "Recommended";
            display: block !important;
            font-size: 11px !important;
            color: #81C3D7 !important;
            margin-top: 4px !important;
            font-weight: 500 !important;
            text-transform: uppercase !important;
            letter-spacing: 0.1em !important;
        }
        </style>
        
        <h4 style="margin:0; color:#81C3D7; font-size:13px; letter-spacing:0.05em; text-transform:uppercase; margin-bottom:6px; font-weight:700;">
            <i class="fi fi-rr-settings-sliders" style="margin-right:8px;"></i>ANALYSIS SCOPE
        </h4>
        <div style="color:#9BADB8; font-size:12px; margin-bottom:24px;">Choose the number of recurring issues to analyze</div>
        """, unsafe_allow_html=True)
        
        if "tracker_top_n" not in st.session_state:
            st.session_state.tracker_top_n = 10
            
        for val in [5, 10, 15, 20, 25]:
            is_active = st.session_state.tracker_top_n == val
            btn_label = f"✓   TOP {val} ISSUES" if is_active else f"TOP {val} ISSUES"
            btn_type = "primary" if is_active else "secondary"
            
            if st.button(btn_label, key=f"tracker_btn_{val}", type=btn_type, use_container_width=True):
                st.session_state.tracker_top_n = val
                st.rerun()
                
        top_n = st.session_state.tracker_top_n

    with col_ai:
        st.markdown("""
<div style="background:linear-gradient(145deg, rgba(19,45,68,0.7) 0%, rgba(10,25,41,0.9) 100%); backdrop-filter:blur(12px); border:1px solid rgba(110,203,160,0.3); border-radius:12px; padding:20px; box-shadow:0 8px 32px rgba(10,25,41,0.6); position:relative; overflow:hidden; min-height:240px;">
<div style="position:absolute; top:-20px; right:-20px; width:100px; height:100px; background:radial-gradient(circle, rgba(110,203,160,0.15) 0%, rgba(0,0,0,0) 70%);"></div>
<div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:12px;">
<h4 style="margin:0; color:#6ECBA0; font-size:13px; font-weight:700; letter-spacing:0.08em; text-transform:uppercase; display:flex; align-items:center; gap:8px;">
✨ AI Recommendation
</h4>
<span style="background:rgba(217,96,90,0.15); color:#D9605A; padding:2px 8px; border-radius:4px; font-size:10px; font-weight:800; border:1px solid rgba(217,96,90,0.3); letter-spacing:0.05em;">HIGH PRIORITY</span>
</div>
<div style="display:grid; grid-template-columns:1fr 1fr; gap:12px; margin-bottom:12px;">
<div>
<span style="color:#9BADB8; font-size:11px; text-transform:uppercase; letter-spacing:0.05em;">Most recurring issue</span>
<div style="color:#FFF; font-size:14px; font-weight:600; margin-top:2px;">Oil Leakage</div>
</div>
<div>
<span style="color:#9BADB8; font-size:11px; text-transform:uppercase; letter-spacing:0.05em;">Occurrence</span>
<div style="color:#FFF; font-size:14px; font-weight:600; margin-top:2px;">2 times</div>
</div>
</div>
<div style="margin-bottom:12px;">
<span style="color:#9BADB8; font-size:11px; text-transform:uppercase; letter-spacing:0.05em;">Affected station</span>
<div style="color:#FFF; font-size:14px; font-weight:600; margin-top:2px;">IC Stud Assembly</div>
</div>
<div style="background:rgba(0,0,0,0.2); padding:10px 12px; border-radius:6px; border-left:3px solid #6ECBA0;">
<span style="color:#6ECBA0; font-size:11px; text-transform:uppercase; font-weight:700; letter-spacing:0.05em; display:block; margin-bottom:4px;">Recommended action</span>
<span style="color:#D9DCD6; font-size:13px; line-height:1.4; display:block;">Inspect hydraulic hoses and preventive maintenance schedule.</span>
</div>
</div>
        """, unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    st.markdown("""
<style>
.nq-btn-premium > div > button {
    background: linear-gradient(135deg, #1A4E70 0%, #0F324D 100%) !important;
    border: 1px solid rgba(129,195,215,0.4) !important;
    box-shadow: 0 4px 20px rgba(10,25,41,0.6), inset 0 1px 1px rgba(255,255,255,0.1) !important;
    height: 56px !important;
    font-size: 16px !important;
    color: #fff !important;
    font-weight: 700 !important;
    letter-spacing: 0.05em !important;
    border-radius: 8px !important;
    transition: all 0.3s cubic-bezier(0.2,0.8,0.2,1) !important;
    text-transform: uppercase !important;
}
.nq-btn-premium > div > button:hover {
    background: linear-gradient(135deg, #205C85 0%, #133E60 100%) !important;
    box-shadow: 0 8px 25px rgba(129,195,215,0.4), inset 0 1px 1px rgba(255,255,255,0.2) !important;
    transform: translateY(-2px) !important;
    border-color: #81C3D7 !important;
}
.nq-btn-premium > div > button::before {
    content: "⚡ ";
    font-size: 18px;
}
</style>
<div class="nq-btn-premium">
    """, unsafe_allow_html=True)
    
    analyze_clicked = st.button("Analyze Recurring Issues", use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)

    if analyze_clicked:
        with st.spinner("Analyzing deviations with AI (single batch)..."):
            df_tracker = generate_external_audit_tracker_with_ai(
                full_data.to_dict(orient="records") if not full_data.empty else [], top_n=top_n
            )
            st.session_state.agent_outputs["tracker"] = df_tracker
            st.success("Recurring issue analysis generated successfully")

    if "tracker" in st.session_state.agent_outputs:
        df_tracker = st.session_state.agent_outputs["tracker"]

        if isinstance(df_tracker, str):
            st.warning(df_tracker)
        else:
            col1, col2, col3 = st.columns(3)
            with col1:
                st.markdown(kpi_card(_icon("refresh", 20), "Tracked Issues", len(df_tracker), "blue"), unsafe_allow_html=True)
            with col2:
                high_pri = len(df_tracker[df_tracker["Priority"] == "High"]) if "Priority" in df_tracker.columns else 0
                st.markdown(kpi_card(_icon("fire", 20), "High Priority", high_pri, "rose"), unsafe_allow_html=True)
            with col3:
                critical = len(df_tracker[df_tracker["Severity"] == "Critical"]) if "Severity" in df_tracker.columns else 0
                st.markdown(kpi_card(_icon("alert", 20), "Critical Severity", critical, "amber"), unsafe_allow_html=True)

            st.markdown(section_header(_icon("alert", 20), "Issue Cards"), unsafe_allow_html=True)
            for i, (_, row) in enumerate(df_tracker.iterrows()):
                st.markdown(
                    tracker_issue_card(
                        index=i + 1,
                        line=str(row.get("Line", "")),
                        station=str(row.get("Station", "")),
                        issue=str(row.get("Issue_Raised_Last_Audit", "")),
                        recurrence=int(row.get("Recurrence_Count", 1)),
                        action=str(row.get("Corrective_Action", "Pending")),
                        root_cause=str(row.get("Root_Cause", "")),
                        owner=str(row.get("Owner", "")),
                        priority=str(row.get("Priority", "Medium")),
                        domain_name=str(row.get("Domain", "Process")),
                    ),
                    unsafe_allow_html=True
                )

            st.markdown("---")
            st.markdown(section_header(_icon("list", 20), "Status Tracker"), unsafe_allow_html=True)
            if "Current_Status" not in df_tracker.columns:
                df_tracker["Current_Status"] = "Not Started"

            display_cols = ["Line", "Station", "Issue_Raised_Last_Audit", "Priority", "Owner", "Current_Status"]
            display_cols = [c for c in display_cols if c in df_tracker.columns]

            _render_download_bar(
                df_tracker[display_cols],
                f"ExternalTracker_{date.today()}",
                excel=True, csv=True,
                key_prefix="dl_tracker",
            )
            _df_wrap_open("Status Tracker")
            edited_df = st.data_editor(
                df_tracker[display_cols],
                use_container_width=True,
                column_config={
                    "Current_Status": st.column_config.SelectboxColumn(
                        "Status", options=["Not Started", "Ongoing", "Solved"]
                    ),
                    "Priority": st.column_config.SelectboxColumn(
                        "Priority", options=["High", "Medium", "Low"]
                    ),
                },
                key="tracker_editor"
            )
            _df_wrap_close()

            for col in edited_df.columns:
                df_tracker[col] = edited_df[col].values
            st.session_state.agent_outputs["tracker"] = df_tracker

            st.markdown("---")
            col4, col5, col6 = st.columns(3)
            with col4:
                st.markdown(kpi_card(_icon("pin", 20), "Not Started", len(df_tracker[df_tracker["Current_Status"] == "Not Started"]), "gray"), unsafe_allow_html=True)
            with col5:
                st.markdown(kpi_card(_icon("clock", 20), "Ongoing", len(df_tracker[df_tracker["Current_Status"] == "Ongoing"]), "amber"), unsafe_allow_html=True)
            with col6:
                st.markdown(kpi_card(_icon("check", 20), "Closed", len(df_tracker[df_tracker["Current_Status"] == "Solved"]), "emerald"), unsafe_allow_html=True)

        st.info("Focus on high recurrence issues to improve audit readiness and reduce external audit risks.")

# ===============================
# PAGE: REPEATABILITY
# ===============================

elif page == "Repeatability":

    st.markdown(section_header(_icon("chart", 20), "Deviation Repeatability Analysis", "AI-powered grouping and trend visualization of recurring deviations"), unsafe_allow_html=True)

    col1, col2 = st.columns(2)
    with col1:
        selected_line = st.selectbox("Select Line", LINE_OPTIONS, key="repeat_line_select")
    with col2:
        time_view = st.selectbox("Select View", ["Day", "Week", "Month"], key="repeat_view")

    st.markdown("---")

    df = clean_line_data(full_data.copy()) if not full_data.empty else pd.DataFrame()
    if not df.empty:
        df = df[df["line"].astype(str) == str(selected_line)]

    _btn_l, _btn_c, _btn_r = st.columns([1, 1, 1])
    with _btn_c:
        _run_analysis = st.button(
            "Analyze Repeatability",
            key="repeat_btn",
            use_container_width=True,
            disabled=df.empty,
        )

    if df.empty:
        st.markdown(empty_state("No data available for this line."), unsafe_allow_html=True)
    else:
        df["date"] = pd.to_datetime(df["date"], errors="coerce")
        observations = df["observation_text"].tolist() if "observation_text" in df.columns else []

        if _run_analysis:
            with st.spinner("AI grouping similar deviations..."):
                categories = map_deviation_category_ai(observations)
                st.session_state["repeat_categories"] = categories
                st.session_state["repeat_line_val"] = selected_line
                st.session_state["repeat_time_val"] = time_view

        if "repeat_categories" in st.session_state:
            st.markdown("---")
            df["Deviation_Category"] = st.session_state["repeat_categories"][:len(df)]

            if time_view == "Day":
                df["Time"] = df["date"].dt.day_name().str[:3]
                order = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]
            elif time_view == "Week":
                df["Time"] = "CW" + df["date"].dt.isocalendar().week.astype(str)
                # ── Build full week range (min→max date) so empty weeks show 0 ──
                _valid_dates = df["date"].dropna()
                if not _valid_dates.empty:
                    from datetime import timedelta as _td
                    _d_min = _valid_dates.min()
                    _d_max = _valid_dates.max()
                    # Expand range: 4 weeks before earliest, 4 weeks after latest
                    _range_start = _d_min - pd.Timedelta(weeks=4)
                    _range_end   = _d_max + pd.Timedelta(weeks=4)
                    # Generate all Mondays in the range to enumerate calendar weeks
                    _all_mondays = pd.date_range(
                        start=_range_start - pd.Timedelta(days=_range_start.weekday()),
                        end=_range_end,
                        freq="W-MON",
                    )
                    # Build ordered week labels with date-range annotations
                    _seen = set()
                    order = []
                    _week_date_map = {}
                    for _mon in _all_mondays:
                        _cw = f"CW{_mon.isocalendar()[1]}"
                        _key = f"{_mon.year}-{_cw}"
                        if _key not in _seen:
                            _seen.add(_key)
                            order.append(_cw)
                            _sun = _mon + pd.Timedelta(days=6)
                            _week_date_map[_cw] = f"{_mon.strftime('%d-%b')} – {_sun.strftime('%d-%b')}"
                    # Deduplicate while preserving order (cross-year may repeat CW1 etc.)
                    _final_order = list(dict.fromkeys(order))
                    order = _final_order
                else:
                    order = sorted(df["Time"].unique())
            elif time_view == "Month":
                df["Time"] = df["date"].dt.strftime("%b")
                order = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

            pivot = pd.pivot_table(df, index="Deviation_Category", columns="Time", aggfunc="size", fill_value=0)
            pivot = pivot.reindex(columns=order, fill_value=0)

            # ── Week View: annotate columns with date ranges for readability ──
            if time_view == "Week" and '_week_date_map' in dir():
                _rename_cols = {}
                for _c in pivot.columns:
                    if _c in _week_date_map:
                        _rename_cols[_c] = f"{_c} ({_week_date_map[_c]})"
                if _rename_cols:
                    pivot.rename(columns=_rename_cols, inplace=True)

            st.markdown(section_header(_icon("chart", 20), "Repeatability Table"), unsafe_allow_html=True)
            _render_download_bar(
                pivot,
                f"Repeatability_{time_view}_{selected_line}_{date.today()}",
                excel=True, csv=True,
                key_prefix="dl_repeat",
            )
            _df_wrap_open("Repeatability Table")
            st.dataframe(pivot, use_container_width=True)
            _df_wrap_close()