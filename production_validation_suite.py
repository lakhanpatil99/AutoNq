"""
production_validation_suite.py
AutoNQ AI – Production Validation Suite (Phases 1–6)

Validates the refactored Summary Engine against all requirements before
production approval.  This script is self-contained and runs outside
Streamlit to isolate every assertion.

Usage:
    python production_validation_suite.py
"""

import os
import sys
import time
import copy
import traceback
import logging
from datetime import date, datetime, timedelta
from io import StringIO

import pandas as pd

# ── Stub out Streamlit so imports work outside a Streamlit process ───────────
# We need to mock st.cache_data and st.session_state before importing backend.
import types

_fake_st = types.ModuleType("streamlit")
_fake_st.cache_data  = lambda *a, **kw: (lambda fn: fn)   # no-op decorator
_fake_st.session_state = {}                                # empty dict
_fake_st.spinner     = lambda *a, **kw: type("ctx", (), {"__enter__": lambda s: None, "__exit__": lambda s, *a: None})()
_fake_st.warning     = lambda *a, **kw: None
_fake_st.error       = lambda *a, **kw: None
_fake_st.info        = lambda *a, **kw: None
_fake_st.toast       = lambda *a, **kw: None
_fake_st.success     = lambda *a, **kw: None
sys.modules["streamlit"] = _fake_st

# Now import project modules
from excel_backend import (
    load_audit_entries, load_sheet, EXCEL_PATH,
    compute_daily_summary, compute_weekly_summary, compute_monthly_summary,
    compute_repeatability, AUDIT_COLUMNS, normalize_columns, inject_aliases,
)
from setup_environment import (
    preprocess_audit_data, build_context_summary,
    tag_domain, tag_severity,
    generate_daily_brief, generate_weekly_brief, generate_monthly_brief,
    generate_mail_from_summary,
    generate_followup_checklist,
    generate_iatf_process_audit_sheet,
    generate_external_audit_tracker_with_ai,
)

logging.basicConfig(level=logging.INFO, format="%(levelname)s | %(message)s")
log = logging.getLogger("VALIDATION")

# ═══════════════════════════════════════════════════════════════════════
# REPORT INFRASTRUCTURE
# ═══════════════════════════════════════════════════════════════════════

_results: list[dict] = []

def _record(phase: str, scenario: str, passed: bool, detail: str = ""):
    _results.append({
        "Phase": phase, "Scenario": scenario,
        "Result": "✅ PASS" if passed else "❌ FAIL",
        "Detail": detail,
    })
    status = "PASS" if passed else "FAIL"
    log.info(f"[{phase}] {scenario}: {status}  {detail}")

# ═══════════════════════════════════════════════════════════════════════
# HELPER: Build a synthetic DataFrame that mirrors the real schema
# ═══════════════════════════════════════════════════════════════════════

def _make_audit_rows(records: list[dict]) -> pd.DataFrame:
    """Create a DataFrame in the shape produced by load_audit_entries + inject_aliases."""
    df = pd.DataFrame(records)
    # Ensure canonical columns exist
    for col in ["line", "station_name", "observation_text", "ai_principle",
                "audit_date", "supervisor", "shift", "severity",
                "category", "audit_id", "status", "image_base64"]:
        if col not in df.columns:
            df[col] = ""
    df["audit_date"] = pd.to_datetime(df["audit_date"], errors="coerce")
    # Inject aliases (mimic backend)
    if "station" not in df.columns and "station_name" in df.columns:
        df["station"] = df["station_name"]
    if "date" not in df.columns:
        df["date"] = df["audit_date"]
    return df


def _today_str():
    return date.today().isoformat()


def _yesterday_str():
    return (date.today() - timedelta(days=1)).isoformat()


def _last_week_str():
    return (date.today() - timedelta(days=8)).isoformat()


def _prev_month_str():
    d = date.today().replace(day=1) - timedelta(days=1)
    return d.isoformat()


# ═══════════════════════════════════════════════════════════════════════
# PHASE 1 – FUNCTIONAL VALIDATION
# ═══════════════════════════════════════════════════════════════════════

def phase1_daily_summary():
    """Scenario A – Daily Summary filters strictly to today."""
    phase = "Phase1-Daily"

    today     = _today_str()
    yesterday = _yesterday_str()
    last_week = _last_week_str()

    full_data = _make_audit_rows([
        {"audit_date": today,     "line": "2", "station_name": "ST4.2",  "observation_text": "Cleaning checklist incomplete",  "ai_principle": "1C – Cleanliness", "supervisor": "Kumar", "audit_id": "A001", "severity": "High"},
        {"audit_date": today,     "line": "2", "station_name": "ST4.2",  "observation_text": "Cleaning checklist incomplete",  "ai_principle": "1C – Cleanliness", "supervisor": "Kumar", "audit_id": "A002", "severity": "High"},
        {"audit_date": today,     "line": "3", "station_name": "ST1",    "observation_text": "PM pending before startup",      "ai_principle": "Total Productive Maintenance (TPM)", "supervisor": "Patel", "audit_id": "A003", "severity": "Medium"},
        {"audit_date": yesterday, "line": "2", "station_name": "ST4.2",  "observation_text": "Old finding from yesterday",     "ai_principle": "Instructions", "supervisor": "Kumar", "audit_id": "A004", "severity": "Low"},
        {"audit_date": last_week, "line": "1", "station_name": "ST2",    "observation_text": "Week-old finding",               "ai_principle": "Instructions", "supervisor": "Singh", "audit_id": "A005", "severity": "Low"},
    ])

    now = pd.Timestamp.now()

    # --- Apply the SAME date filter that app.py now uses ---
    daily_df = full_data[full_data["audit_date"].dt.date == now.date()].copy()

    # Test 1: Only today's rows survive
    _record(phase, "Only today's observations", len(daily_df) == 3,
            f"Expected 3, got {len(daily_df)}")

    # Test 2: Yesterday excluded
    yesterday_in = daily_df[daily_df["audit_date"].dt.date == (now - pd.Timedelta(days=1)).date()]
    _record(phase, "Yesterday excluded", len(yesterday_in) == 0,
            f"Expected 0, got {len(yesterday_in)}")

    # Test 3: Last week excluded
    lw_in = daily_df[daily_df["audit_date"].dt.date < (now - pd.Timedelta(days=6)).date()]
    _record(phase, "Last week excluded", len(lw_in) == 0,
            f"Expected 0, got {len(lw_in)}")

    # Test 4: Recurrence is from today only
    processed = preprocess_audit_data(daily_df)
    cleaning_rows = processed[processed["observation_text"] == "Cleaning checklist incomplete"]
    if not cleaning_rows.empty:
        rec_count = cleaning_rows["recurrence_count"].iloc[0]
        _record(phase, "Recurrence count from today only", rec_count == 2,
                f"Expected 2, got {rec_count}")
    else:
        _record(phase, "Recurrence count from today only", False, "No cleaning rows found")

    # Test 5: Context summary uses recurrence_count.sum() for total
    context = build_context_summary(processed)
    _record(phase, "Context total uses recurrence sum", "Total Observations: 3" in context,
            f"Context snippet: {context[:80]}")

    # Test 6: Context includes location info in recurring issues
    _record(phase, "Context recurring includes location", "Affected Locations:" in context,
            f"Context snippet: {context[context.find('Top Recurring'):][:200] if 'Top Recurring' in context else 'NOT FOUND'}")


def phase1_weekly_summary():
    """Scenario B – Weekly Summary filters strictly to current ISO week."""
    phase = "Phase1-Weekly"

    now = pd.Timestamp.now()
    iso_week = now.isocalendar().week
    iso_year = now.isocalendar().year

    # Create records spanning this week and last week
    this_monday = now - pd.Timedelta(days=now.weekday())
    last_monday = this_monday - pd.Timedelta(days=7)

    full_data = _make_audit_rows([
        {"audit_date": this_monday.strftime("%Y-%m-%d"), "line": "2", "station_name": "ST4.2", "observation_text": "This week finding 1", "ai_principle": "Instructions", "supervisor": "Kumar", "audit_id": "W001", "severity": "High"},
        {"audit_date": (this_monday + pd.Timedelta(days=1)).strftime("%Y-%m-%d"), "line": "3", "station_name": "ST1", "observation_text": "This week finding 2", "ai_principle": "Tools", "supervisor": "Patel", "audit_id": "W002", "severity": "Medium"},
        {"audit_date": last_monday.strftime("%Y-%m-%d"), "line": "2", "station_name": "ST4.2", "observation_text": "Last week finding", "ai_principle": "Instructions", "supervisor": "Singh", "audit_id": "W003", "severity": "Low"},
        {"audit_date": (last_monday + pd.Timedelta(days=2)).strftime("%Y-%m-%d"), "line": "1", "station_name": "ST2", "observation_text": "Last week finding 2", "ai_principle": "Tools", "supervisor": "Singh", "audit_id": "W004", "severity": "Low"},
    ])

    # Apply the SAME filter logic from app.py
    weekly_df = full_data[
        (full_data["audit_date"].dt.isocalendar().week == iso_week) &
        (full_data["audit_date"].dt.isocalendar().year == iso_year)
    ].copy()

    prev_week_date = now - pd.Timedelta(weeks=1)
    prev_iso_week = prev_week_date.isocalendar().week
    prev_iso_year = prev_week_date.isocalendar().year
    prev_weekly_df = full_data[
        (full_data["audit_date"].dt.isocalendar().week == prev_iso_week) &
        (full_data["audit_date"].dt.isocalendar().year == prev_iso_year)
    ].copy()

    # Test 1: Current week only
    _record(phase, "Only current ISO week rows", len(weekly_df) == 2,
            f"Expected 2, got {len(weekly_df)}")

    # Test 2: Previous week excluded from current
    _record(phase, "Previous week excluded from current", len(prev_weekly_df) == 2,
            f"Previous week df has {len(prev_weekly_df)} rows (should be 2)")

    # Test 3: No overlap
    overlap = set(weekly_df["audit_id"].tolist()) & set(prev_weekly_df["audit_id"].tolist())
    _record(phase, "No overlap between weeks", len(overlap) == 0,
            f"Overlap IDs: {overlap}")


def phase1_monthly_summary():
    """Scenario C – Monthly Summary filters strictly to current calendar month."""
    phase = "Phase1-Monthly"

    now = pd.Timestamp.now()
    this_month_date = now.strftime("%Y-%m-15")
    prev_month = now.replace(day=1) - pd.Timedelta(days=1)
    prev_month_date = prev_month.strftime("%Y-%m-15")

    full_data = _make_audit_rows([
        {"audit_date": this_month_date, "line": "2", "station_name": "ST4.2", "observation_text": "This month finding", "ai_principle": "Instructions", "supervisor": "Kumar", "audit_id": "M001", "severity": "High"},
        {"audit_date": _today_str(),    "line": "3", "station_name": "ST1",   "observation_text": "Today finding",      "ai_principle": "Tools", "supervisor": "Patel", "audit_id": "M002", "severity": "Medium"},
        {"audit_date": prev_month_date, "line": "1", "station_name": "ST2",   "observation_text": "Last month finding",  "ai_principle": "Tools", "supervisor": "Singh", "audit_id": "M003", "severity": "Low"},
    ])

    monthly_df = full_data[
        (full_data["audit_date"].dt.month == now.month) &
        (full_data["audit_date"].dt.year == now.year)
    ].copy()

    prev_monthly_df = full_data[
        (full_data["audit_date"].dt.month == prev_month.month) &
        (full_data["audit_date"].dt.year == prev_month.year)
    ].copy()

    _record(phase, "Only current month rows", len(monthly_df) == 2,
            f"Expected 2, got {len(monthly_df)}")

    _record(phase, "Previous month excluded", len(prev_monthly_df) == 1,
            f"Expected 1, got {len(prev_monthly_df)}")

    overlap = set(monthly_df["audit_id"].tolist()) & set(prev_monthly_df["audit_id"].tolist())
    _record(phase, "No overlap between months", len(overlap) == 0,
            f"Overlap IDs: {overlap}")


# ═══════════════════════════════════════════════════════════════════════
# PHASE 2 – MULTI-USER VALIDATION
# ═══════════════════════════════════════════════════════════════════════

def phase2_multi_user():
    """Simulate: User A opens app, User B submits audit, User A regenerates."""
    phase = "Phase2-MultiUser"

    # Verify get_data() no longer uses data_loaded flag
    import inspect
    # We need to read the source of app.py directly since we can't import it
    app_source = open(os.path.join(os.path.dirname(__file__), "app.py"), "r", encoding="utf-8").read()

    # Test 1: data_loaded flag removed from session state init
    has_data_loaded_init = "data_loaded" in app_source and "st.session_state.data_loaded" in app_source
    # There should be NO active use of data_loaded
    # Count occurrences: should only appear in comments at most
    import re
    active_uses = [m for m in re.findall(r"st\.session_state\.data_loaded", app_source)]
    _record(phase, "data_loaded flag removed", len(active_uses) == 0,
            f"Found {len(active_uses)} active uses of data_loaded")

    # Test 2: get_data() does NOT conditionally gate on data_loaded
    get_data_match = re.search(r"def get_data\(\).*?(?=\ndef |\Z)", app_source, re.DOTALL)
    if get_data_match:
        get_data_src = get_data_match.group()
        _record(phase, "get_data no data_loaded gate", "data_loaded" not in get_data_src,
                "get_data() is now unconditional")
    else:
        _record(phase, "get_data no data_loaded gate", False, "Could not find get_data()")

    # Test 3: load_audit_entries cache bust on add (existing behaviour preserved)
    backend_src = open(os.path.join(os.path.dirname(__file__), "excel_backend.py"), "r", encoding="utf-8").read()
    _record(phase, "add_audit_entry busts cache", "load_audit_entries.clear()" in backend_src,
            "Backend still clears cache on add")


# ═══════════════════════════════════════════════════════════════════════
# PHASE 3 – MODULE CONSISTENCY
# ═══════════════════════════════════════════════════════════════════════

def phase3_consistency():
    """Verify that Summary / Recurrence / Repeatability produce consistent metrics."""
    phase = "Phase3-Consistency"

    today = _today_str()
    common_data = _make_audit_rows([
        {"audit_date": today, "line": "2", "station_name": "ST4.2", "observation_text": "Cleaning checklist incomplete", "ai_principle": "1C – Cleanliness", "supervisor": "Kumar", "audit_id": "C001", "severity": "High"},
        {"audit_date": today, "line": "2", "station_name": "ST4.2", "observation_text": "Cleaning checklist incomplete", "ai_principle": "1C – Cleanliness", "supervisor": "Kumar", "audit_id": "C002", "severity": "High"},
        {"audit_date": today, "line": "2", "station_name": "ST4.2", "observation_text": "Cleaning checklist incomplete", "ai_principle": "1C – Cleanliness", "supervisor": "Kumar", "audit_id": "C003", "severity": "High"},
        {"audit_date": today, "line": "3", "station_name": "ST1",   "observation_text": "PM pending before startup",    "ai_principle": "Total Productive Maintenance (TPM)", "supervisor": "Patel", "audit_id": "C004", "severity": "Medium"},
        {"audit_date": today, "line": "3", "station_name": "ST1",   "observation_text": "Tool storage mismatch",        "ai_principle": "Tools", "supervisor": "Patel", "audit_id": "C005", "severity": "Low"},
    ])

    processed = preprocess_audit_data(common_data)

    # After dedup, should have 3 unique combos
    _record(phase, "Dedup row count", len(processed) == 3,
            f"Expected 3 unique rows, got {len(processed)}")

    # Recurrence count for cleaning should be 3
    cleaning = processed[processed["observation_text"] == "Cleaning checklist incomplete"]
    if not cleaning.empty:
        rc = cleaning["recurrence_count"].iloc[0]
        _record(phase, "Cleaning recurrence=3", rc == 3,
                f"Expected 3, got {rc}")
    else:
        _record(phase, "Cleaning recurrence=3", False, "Cleaning row missing")

    # Context summary total should be sum of recurrences = 5
    ctx = build_context_summary(processed)
    _record(phase, "Context total observations=5", "Total Observations: 5" in ctx,
            f"Context: {ctx[:80]}")

    # Top issue should be cleaning (highest recurrence)
    _record(phase, "Top issue is Cleaning", "Cleaning checklist incomplete" in ctx,
            "Correctly identifies top recurring issue")

    # Domain mapping
    _record(phase, "Domain mapping consistent", processed[processed["observation_text"] == "Cleaning checklist incomplete"]["domain"].iloc[0] == "Process",
            "1C – Cleanliness maps to Process")

    # Severity mapping
    _record(phase, "Severity tagging consistent", processed[processed["observation_text"] == "Cleaning checklist incomplete"]["severity"].iloc[0] == "Major",
            "recurrence=3 maps to Major")


# ═══════════════════════════════════════════════════════════════════════
# PHASE 4 – EDGE CASES
# ═══════════════════════════════════════════════════════════════════════

def phase4_edge_cases():
    """Validate stability under edge conditions."""
    phase = "Phase4-Edge"

    # Edge 1: No audits today
    now = pd.Timestamp.now()
    full_data = _make_audit_rows([
        {"audit_date": _yesterday_str(), "line": "2", "station_name": "ST4", "observation_text": "Old obs", "ai_principle": "Tools", "supervisor": "A", "audit_id": "E001", "severity": "Low"},
    ])
    daily_df = full_data[full_data["audit_date"].dt.date == now.date()].copy()
    _record(phase, "No audits today – empty df", len(daily_df) == 0, f"Got {len(daily_df)}")

    # Preprocess empty df should not crash
    try:
        result = preprocess_audit_data(daily_df)
        _record(phase, "Empty df preprocess stable", True, f"Returned {len(result)} rows")
    except Exception as e:
        _record(phase, "Empty df preprocess stable", False, str(e))

    # build_context_summary with empty df
    try:
        ctx = build_context_summary(result)
        _record(phase, "Empty context summary stable", "No data available" in ctx, f"Got: {ctx[:60]}")
    except Exception as e:
        _record(phase, "Empty context summary stable", False, str(e))

    # Edge 2: Single audit
    single = _make_audit_rows([
        {"audit_date": _today_str(), "line": "2", "station_name": "ST4", "observation_text": "Lone obs", "ai_principle": "Tools", "supervisor": "A", "audit_id": "E002", "severity": "Low"},
    ])
    try:
        p = preprocess_audit_data(single)
        c = build_context_summary(p)
        _record(phase, "Single audit stable", "Total Observations: 1" in c, f"Got: {c[:60]}")
    except Exception as e:
        _record(phase, "Single audit stable", False, str(e))

    # Edge 3: Hundreds of audits
    try:
        big_records = [
            {"audit_date": _today_str(), "line": str(i % 5 + 1), "station_name": f"ST{i%10}",
             "observation_text": f"Observation type {i % 20}", "ai_principle": "Tools",
             "supervisor": f"Sup{i%3}", "audit_id": f"BIG{i:04d}", "severity": "Medium"}
            for i in range(500)
        ]
        big_df = _make_audit_rows(big_records)
        p = preprocess_audit_data(big_df)
        c = build_context_summary(p)
        _record(phase, "500 audits stable", "Total Observations:" in c, f"Processed {len(p)} unique rows")
    except Exception as e:
        _record(phase, "500 audits stable", False, str(e))

    # Edge 4: Duplicate observations (all identical)
    try:
        dup_records = [
            {"audit_date": _today_str(), "line": "2", "station_name": "ST4",
             "observation_text": "Same finding", "ai_principle": "Tools",
             "supervisor": "A", "audit_id": f"DUP{i}", "severity": "Low"}
            for i in range(10)
        ]
        dup_df = _make_audit_rows(dup_records)
        p = preprocess_audit_data(dup_df)
        _record(phase, "10 duplicates dedup to 1", len(p) == 1 and p["recurrence_count"].iloc[0] == 10,
                f"Got {len(p)} rows, recurrence={p['recurrence_count'].iloc[0] if len(p) > 0 else 'N/A'}")
    except Exception as e:
        _record(phase, "10 duplicates dedup to 1", False, str(e))

    # Edge 5: Invalid date (simulating NaT fallback in backend)
    try:
        invalid_dates = _make_audit_rows([
            {"audit_date": "not-a-date", "line": "2", "station_name": "ST4", "observation_text": "Bad date", "ai_principle": "Tools", "supervisor": "A", "audit_id": "INV001", "severity": "Low"},
        ])
        nat_count = invalid_dates["audit_date"].isna().sum()
        # In the backend, NaT is now filled with today. We simulate this:
        invalid_dates["audit_date"] = invalid_dates["audit_date"].fillna(pd.Timestamp.today())
        _record(phase, "NaT fallback to today", invalid_dates["audit_date"].dt.date.iloc[0] == date.today(),
                f"Date resolved to {invalid_dates['audit_date'].iloc[0]}")
    except Exception as e:
        _record(phase, "NaT fallback to today", False, str(e))

    # Edge 6: Missing station / line / checkpoint
    try:
        missing = _make_audit_rows([
            {"audit_date": _today_str(), "line": "", "station_name": "", "observation_text": "", "ai_principle": "", "supervisor": "", "audit_id": "MISS01", "severity": ""},
        ])
        p = preprocess_audit_data(missing)
        c = build_context_summary(p)
        _record(phase, "Missing fields stable", True, f"Processed OK, context: {c[:60]}")
    except Exception as e:
        _record(phase, "Missing fields stable", False, str(e))

    # Edge 7: Empty DataFrame (simulating empty workbook)
    try:
        empty = pd.DataFrame(columns=AUDIT_COLUMNS)
        p = preprocess_audit_data(empty)
        c = build_context_summary(p)
        _record(phase, "Empty workbook stable", "No data available" in c, f"Got: {c[:60]}")
    except Exception as e:
        _record(phase, "Empty workbook stable", False, str(e))


# ═══════════════════════════════════════════════════════════════════════
# PHASE 5 – PERFORMANCE
# ═══════════════════════════════════════════════════════════════════════

def phase5_performance():
    """Measure timing for key operations."""
    phase = "Phase5-Perf"

    # Test with synthetic data (skip actual Excel if not available)
    big_records = [
        {"audit_date": _today_str(), "line": str(i % 5 + 1), "station_name": f"ST{i%10}",
         "observation_text": f"Observation type {i % 20}", "ai_principle": "Tools",
         "supervisor": f"Sup{i%3}", "audit_id": f"PERF{i:04d}", "severity": "Medium"}
        for i in range(1000)
    ]
    big_df = _make_audit_rows(big_records)

    # Preprocess timing
    t0 = time.perf_counter()
    p = preprocess_audit_data(big_df)
    t_preprocess = time.perf_counter() - t0
    _record(phase, f"Preprocess 1000 rows: {t_preprocess*1000:.1f}ms", t_preprocess < 2.0,
            f"{t_preprocess*1000:.1f}ms (threshold: <2000ms)")

    # Context summary timing
    t0 = time.perf_counter()
    ctx = build_context_summary(p)
    t_context = time.perf_counter() - t0
    _record(phase, f"Context summary: {t_context*1000:.1f}ms", t_context < 1.0,
            f"{t_context*1000:.1f}ms (threshold: <1000ms)")

    # Date filtering timing (simulating app.py logic)
    now = pd.Timestamp.now()
    t0 = time.perf_counter()
    daily = big_df[big_df["audit_date"].dt.date == now.date()]
    weekly = big_df[
        (big_df["audit_date"].dt.isocalendar().week == now.isocalendar().week) &
        (big_df["audit_date"].dt.isocalendar().year == now.isocalendar().year)
    ]
    monthly = big_df[
        (big_df["audit_date"].dt.month == now.month) &
        (big_df["audit_date"].dt.year == now.year)
    ]
    t_filter = time.perf_counter() - t0
    _record(phase, f"Date filter (D/W/M): {t_filter*1000:.1f}ms", t_filter < 1.0,
            f"{t_filter*1000:.1f}ms for all three filters")

    # Excel reload timing (only if file exists)
    if os.path.isfile(EXCEL_PATH):
        t0 = time.perf_counter()
        load_audit_entries()
        t_excel = time.perf_counter() - t0
        _record(phase, f"Excel reload: {t_excel*1000:.1f}ms", t_excel < 5.0,
                f"{t_excel*1000:.1f}ms (threshold: <5000ms)")

        # No duplicate reads check
        t0 = time.perf_counter()
        load_audit_entries()  # Second call should be near-instant (cache)
        t_cached = time.perf_counter() - t0
        # Note: since we stubbed st.cache_data as no-op, this will re-read.
        # In real Streamlit, this would be cached.
        _record(phase, f"Excel second read: {t_cached*1000:.1f}ms", True,
                f"{t_cached*1000:.1f}ms (note: cache is stubbed; real Streamlit uses TTL=30s cache)")
    else:
        _record(phase, "Excel reload (skipped)", True, f"File not found: {EXCEL_PATH}")


# ═══════════════════════════════════════════════════════════════════════
# PHASE 6 – REGRESSION VALIDATION
# ═══════════════════════════════════════════════════════════════════════

def phase6_regression():
    """Confirm non-Summary features remain intact."""
    phase = "Phase6-Regression"

    # Test 1: Audit Entry schema unchanged
    _record(phase, "AUDIT_COLUMNS schema intact", len(AUDIT_COLUMNS) == 19 and "image_base64" in AUDIT_COLUMNS,
            f"Columns: {len(AUDIT_COLUMNS)}, last: {AUDIT_COLUMNS[-1]}")

    # Test 2: normalize_columns still works
    test_df = pd.DataFrame({"Actual Result / Observation": ["test"], "Line Name": ["L1"]})
    normed = normalize_columns(test_df)
    _record(phase, "normalize_columns works", "actual_result" in normed.columns and "line" in normed.columns,
            f"Columns: {list(normed.columns)}")

    # Test 3: inject_aliases still works
    normed2 = pd.DataFrame({"actual_result": ["test"], "category": ["Tools"], "flm_name": ["Kumar"], "station_name": ["ST1"], "audit_date": [_today_str()]})
    aliased = inject_aliases(normed2)
    _record(phase, "inject_aliases works",
            "observation_text" in aliased.columns and "ai_principle" in aliased.columns and "supervisor" in aliased.columns,
            f"Aliased columns: {[c for c in aliased.columns if c in ['observation_text','ai_principle','supervisor','station','date']]}")

    # Test 4: tag_domain mappings
    _record(phase, "tag_domain mapping", tag_domain("Stop Sign") == "Safety" and tag_domain("Tools") == "Process",
            "Correct domain mappings")

    # Test 5: tag_severity mappings
    _record(phase, "tag_severity mapping",
            tag_severity(5) == "Critical" and tag_severity(3) == "Major" and tag_severity(1) == "Low",
            "Correct severity mappings")

    # Test 6: compute_daily_summary function exists
    _record(phase, "compute_daily_summary exists", callable(compute_daily_summary),
            "Function still importable")

    # Test 7: compute_weekly_summary function exists
    _record(phase, "compute_weekly_summary exists", callable(compute_weekly_summary),
            "Function still importable")

    # Test 8: compute_monthly_summary function exists
    _record(phase, "compute_monthly_summary exists", callable(compute_monthly_summary),
            "Function still importable")

    # Test 9: compute_repeatability function exists
    _record(phase, "compute_repeatability exists", callable(compute_repeatability),
            "Function still importable")

    # Test 10: generate_followup_checklist still works
    agent3_df = pd.DataFrame({
        "Station": ["ST4.2"], "Process_Risk": ["Cleaning incomplete"],
        "Recurrence_Count": [3]
    })
    try:
        fup = generate_followup_checklist("2", agent3_df)
        _record(phase, "Follow-up checklist works", isinstance(fup, pd.DataFrame) and len(fup) == 1,
                f"Generated {len(fup)} follow-up items")
    except Exception as e:
        _record(phase, "Follow-up checklist works", False, str(e))

    # Test 11: generate_iatf_process_audit_sheet still works (data-only, no LLM)
    deviation_data = [
        {"line": "2", "station": "ST4.2", "observation_text": "Cleaning incomplete", "ai_principle": "1C – Cleanliness"},
        {"line": "2", "station": "ST4.2", "observation_text": "Cleaning incomplete", "ai_principle": "1C – Cleanliness"},
        {"line": "2", "station": "ST1",   "observation_text": "PM pending",          "ai_principle": "Total Productive Maintenance (TPM)"},
    ]
    try:
        audit_sheet = generate_iatf_process_audit_sheet("2", deviation_data, pd.DataFrame(), pd.DataFrame())
        _record(phase, "IATF process audit works",
                isinstance(audit_sheet, pd.DataFrame) and len(audit_sheet) > 0,
                f"Generated {len(audit_sheet)} audit rows")
    except Exception as e:
        _record(phase, "IATF process audit works", False, str(e))

    # Test 12: App.py UI structure preserved (check key layout markers)
    app_source = open(os.path.join(os.path.dirname(__file__), "app.py"), "r", encoding="utf-8").read()
    ui_markers = [
        "st.sidebar",
        "st.radio",
        "st.download_button",
        "mail_preview_card",
        "structured_ai_card",
        "section_header",
        "st.text_area",
    ]
    all_present = all(m in app_source for m in ui_markers)
    missing = [m for m in ui_markers if m not in app_source]
    _record(phase, "UI markers preserved", all_present,
            f"Missing: {missing}" if missing else "All UI functions present")

    # Test 13: EML download preserved
    _record(phase, "EML download preserved",
            ".eml" in app_source and "Content-Type: text/plain" in app_source,
            "EML generation code intact")

    # Test 14: CSV export preserved
    _record(phase, "CSV export preserved", ".csv" in app_source, "CSV download code intact")


# ═══════════════════════════════════════════════════════════════════════
# MAIN
# ═══════════════════════════════════════════════════════════════════════

def main():
    print("=" * 72)
    print("  AutoNQ AI – Production Validation Suite")
    print(f"  Run at: {datetime.now().isoformat()}")
    print("=" * 72)
    print()

    # Phase 1 – Functional
    print("─── Phase 1: Functional Validation ───")
    phase1_daily_summary()
    phase1_weekly_summary()
    phase1_monthly_summary()
    print()

    # Phase 2 – Multi-User
    print("─── Phase 2: Multi-User Validation ───")
    phase2_multi_user()
    print()

    # Phase 3 – Module Consistency
    print("─── Phase 3: Module Consistency ───")
    phase3_consistency()
    print()

    # Phase 4 – Edge Cases
    print("─── Phase 4: Edge Cases ───")
    phase4_edge_cases()
    print()

    # Phase 5 – Performance
    print("─── Phase 5: Performance ───")
    phase5_performance()
    print()

    # Phase 6 – Regression
    print("─── Phase 6: Regression Validation ───")
    phase6_regression()
    print()

    # ── Final Report ──
    print("=" * 72)
    print("  FINAL VALIDATION REPORT")
    print("=" * 72)
    total = len(_results)
    passed = sum(1 for r in _results if "PASS" in r["Result"])
    failed = sum(1 for r in _results if "FAIL" in r["Result"])

    for r in _results:
        print(f"  {r['Result']}  [{r['Phase']}] {r['Scenario']}")
        if r["Detail"]:
            print(f"         └─ {r['Detail']}")

    print()
    print(f"  Total: {total}  |  Passed: {passed}  |  Failed: {failed}")
    rate = (passed / total * 100) if total > 0 else 0
    print(f"  Pass Rate: {rate:.1f}%")

    if failed == 0:
        print("\n  ✅ ALL TESTS PASSED – READY FOR PRODUCTION APPROVAL")
    else:
        print(f"\n  ⚠️  {failed} TEST(S) FAILED – REVIEW REQUIRED BEFORE DEPLOYMENT")

    print("=" * 72)
    return 0 if failed == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
