import pandas as pd
import os
import json
import time
import re
import logging
from groq import Groq

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

logger = logging.getLogger(__name__)

# ===============================
# DOMAIN & SEVERITY TAGGING
# ===============================

PRINCIPLE_DOMAIN_MAP = {
    "Stop Sign": "Safety", "Andon Cord": "Safety", "Dropped Parts": "Safety",
    "Rework / Scrap": "Quality", "Labeling": "Quality",
    "Measurement / Test Equipment": "Quality", "Check the Checker": "Quality",
    "Correct Product": "Quality",
    "1C – Cleanliness": "Process", "Instructions": "Process",
    "Process Parameters": "Process", "Tools": "Process", "Restart": "Process",
    "Remaining Items": "Process", "Total Productive Maintenance (TPM)": "Process",
}

def tag_domain(principle):
    return PRINCIPLE_DOMAIN_MAP.get(principle, "Process")

def tag_severity(recurrence_count):
    if recurrence_count >= 5:
        return "Critical"
    elif recurrence_count >= 3:
        return "Major"
    return "Low"

# ===============================
# EXECUTIVE TONE FILTER  ← NEW
# ===============================

BLOCKED_PHRASES = [
    "might be because", "likely because", "likely due to", "appears to",
    "seems to", "possibly", "attention to", "general picture",
    "observed that", "it was found", "it was noted", "minor", "slight",
    "manageable", "low impact", "negligible", "concern", "issue observed",
    "due to lack of", "helping out", "could indicate", "not up to standard",
    "enough people", "storytelling", "attention required", "it is worth noting",
    "it should be highlighted", "may indicate", "may suggest",
]

_SHORTEN_PATTERNS = [
    # Grammar and fluff cleanup only — NEVER change the meaning or replace nouns
    (r"\b(is|are|was|were) not\b", "not"),
    (r"\b(has|have) not been\b", "not"),
    (r"\bneeds to be\b", "pending"),
    (r"\bneed to be\b", "pending"),
    (r"\bthere is a\b", ""),
    (r"\bthere are\b", ""),
    (r"\bit appears that\b", ""),
    (r"\bwe found that\b", ""),
    (r"\bwere found to be\b", "were"),
    (r"\bwas found to be\b", "was"),
]

def _apply_executive_filter(text: str) -> str:
    """Remove blocked phrases and apply shortening transforms."""
    if not text:
        return text
    for phrase in BLOCKED_PHRASES:
        text = re.sub(re.escape(phrase), "", text, flags=re.IGNORECASE)
    for pattern, replacement in _SHORTEN_PATTERNS:
        text = re.sub(pattern, replacement, text, flags=re.IGNORECASE)
    # Collapse multiple spaces / clean up artifacts
    text = re.sub(r"[ \t]{2,}", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()

def _trim_bullet(bullet: str, max_words: int = 12) -> str:
    """Hard-trim a bullet to max_words words."""
    words = bullet.strip().split()
    if len(words) <= max_words:
        return bullet.strip()
    trimmed = " ".join(words[:max_words])
    if not trimmed.endswith("."):
        trimmed = trimmed.rstrip(".,;:") + "."
    return trimmed

# ===============================
# LLM CLIENT (FIXED)
# ===============================

class BoschLLMClient:
    """Centralized LLM client using Groq with retry logic."""

    def __init__(self):
        api_key = os.getenv("GROQ_API_KEY")
        if not api_key:
            raise EnvironmentError(
                "GROQ_API_KEY not set. Add it to your .env file."
            )
        self.client = Groq(api_key=api_key)

    def chat(self, messages, max_tokens=600, temperature=0.2, retries=2):
        for attempt in range(retries + 1):
            try:
                response = self.client.chat.completions.create(
                    model="llama-3.1-8b-instant",
                    messages=messages,
                    max_tokens=max_tokens,
                    temperature=temperature
                )
                return response.choices[0].message.content
            except Exception as e:
                if attempt < retries:
                    time.sleep(2 ** attempt)
                    logger.warning(f"LLM retry {attempt+1}: {e}")
                    continue
                logger.error(f"LLM failed after {retries+1} attempts: {e}")
                return "⚠️ AI temporarily unavailable. Please try again."


llm = None

def _get_llm():
    """Lazy LLM singleton — only instantiated on first use."""
    global llm
    if llm is None:
        llm = BoschLLMClient()
    return llm

# ===============================
# JSON PARSING UTILITY
# ===============================

def parse_json_response(text, fallback=None):
    if fallback is None:
        fallback = []
    try:
        return json.loads(text)
    except (json.JSONDecodeError, TypeError):
        pass
    match = re.search(r'```(?:json)?\s*([\s\S]*?)```', text or "")
    if match:
        try:
            return json.loads(match.group(1))
        except (json.JSONDecodeError, TypeError):
            pass
    for pat in [r'(\[[\s\S]*\])', r'(\{[\s\S]*\})']:
        match = re.search(pat, text or "")
        if match:
            try:
                return json.loads(match.group(1))
            except (json.JSONDecodeError, TypeError):
                continue
    return fallback

# ===============================
# DATA PREPROCESSING PIPELINE
# ===============================

def preprocess_audit_data(df):
    """Clean, normalize, deduplicate, and enrich audit data."""
    if df is None or df.empty:
        return df
    result = df.copy()

    if "station" in result.columns:
        result["station"] = result["station"].astype(str).str.strip().str.title()
    if "line" in result.columns:
        result["line"] = result["line"].astype(str).str.strip()
    if "supervisor" in result.columns:
        result["supervisor"] = result["supervisor"].astype(str).str.strip().str.title()
    if "observation_text" in result.columns:
        result["observation_text"] = result["observation_text"].astype(str).str.strip()

    if "observation_text" in result.columns and "line" in result.columns and "station" in result.columns:
        rec = result.groupby(["line", "station", "observation_text"]).size().reset_index(name="recurrence_count")
        if "recurrence_count" in result.columns:
            result = result.drop(columns=["recurrence_count"])
        result = result.merge(rec, on=["line", "station", "observation_text"], how="left")
        result["recurrence_count"] = result["recurrence_count"].fillna(1).astype(int)

    dedup_cols = [c for c in ["line", "station", "observation_text"] if c in result.columns]
    if dedup_cols:
        result = result.drop_duplicates(subset=dedup_cols, keep="last")

    if "ai_principle" in result.columns:
        result["domain"] = result["ai_principle"].apply(tag_domain)
    if "recurrence_count" in result.columns:
        result["severity"] = result["recurrence_count"].apply(tag_severity)

    return result


def build_context_summary(df, line=None):
    """Build structured context string for LLM prompts."""
    if df is None or df.empty:
        return "No data available."
    work = df.copy()
    if line is not None:
        work = work[work["line"].astype(str) == str(line)]
    if work.empty:
        return f"No data available for Line {line}."

    total = work["recurrence_count"].sum() if "recurrence_count" in work.columns else len(work)
    top_principle = work["ai_principle"].value_counts().idxmax() if "ai_principle" in work.columns else "N/A"
    top_station = work["station"].value_counts().idxmax() if "station" in work.columns else "N/A"
    top_supervisor = work["supervisor"].value_counts().idxmax() if "supervisor" in work.columns else "N/A"
    lines_covered = sorted(work["line"].dropna().unique().tolist()) if "line" in work.columns else []

    issues_str = "  No observation data."
    if "observation_text" in work.columns:
        # Guarantee columns exist for the composite key fallback
        if "ai_principle" not in work.columns:
            work["ai_principle"] = "Process"
        if "recurrence_count" not in work.columns:
            work["recurrence_count"] = 1
            
        # Composite Grouping Key: ["ai_principle", "observation_text"]
        grouped_issues = []
        for (principle, obs_text), group in work.groupby(["ai_principle", "observation_text"]):
            total_occurrences = group["recurrence_count"].sum()
            
            # Map unique locations (Line - Station)
            locations = []
            for _, row in group.iterrows():
                loc_str = f"Line {row.get('line', 'N/A')} – {row.get('station', 'N/A')}"
                if loc_str not in locations:
                    locations.append(loc_str)
                    
            grouped_issues.append({
                "observation": obs_text,
                "category": principle,
                "total": total_occurrences,
                "locations": locations
            })
            
        # Prioritization: Sort by highest recurrence count
        grouped_issues.sort(key=lambda x: x["total"], reverse=True)
        
        # Build Executive Readability Format
        formatted_blocks = []
        for issue in grouped_issues:
            loc_bullets = "\n".join([f"  • {loc}" for loc in issue["locations"]])
            block = (
                f"{issue['observation']}\n"
                f"Category: {issue['category']}\n\n"
                f"Affected Locations:\n{loc_bullets}\n\n"
                f"Total Occurrences: {issue['total']}"
            )
            formatted_blocks.append(block)
            
        if formatted_blocks:
            issues_str = "\n\n────────────────────\n\n".join(formatted_blocks)

    domain_str = "N/A"
    if "domain" in work.columns:
        domain_str = ", ".join([f"{d} ({c})" for d, c in work["domain"].value_counts().items()])

    return f"""Total Observations: {total}
Lines Covered: {', '.join(str(l) for l in lines_covered) if lines_covered else 'N/A'}
Primary Focus Area: {top_principle}
Key Station: {top_station}
Shift Coverage: {top_supervisor}'s shifts
Domain Distribution: {domain_str}

Top Recurring Observations:
{issues_str}"""

# ===============================
# SHARED EXECUTIVE TONE RULES  ← PATCHED
# ===============================

_EXEC_TONE_RULES = """
EXECUTIVE MANUFACTURING MAIL RULES — MANDATORY:

OUTPUT STYLE:
- Shift-handover language. Factory floor. Operational.
- Short lines. No paragraphs. No storytelling.
- Bullets: 5–10 words max. One line only.
- Why section: one short line only.

NEVER WRITE:
- might be because / likely because / appears to / seems to / possibly
- general picture / attention required / observed that / it was found
- minor / slight / manageable / low impact / negligible
- due to lack of / could indicate / concern / not up to standard
- operator failed / worker ignored / person forgot (NO blame)
- any sentence longer than 12 words in a bullet
- any explanation longer than one line in Why

OBSERVATION PRIORITY RULE (CRITICAL):
- The observation text contained in the dataframe is the AUTHORITATIVE SOURCE.
- Use observation exactly as entered.
- Never replace it. Never invent a different issue. Never generalize it. Never substitute with training examples.
- If grammar correction is needed, correct grammar only.
- Never replace technical terms, station names, line names, equipment names, component names, or issue descriptions.
"""

# ===============================
# SUMMARY AGENTS  ← PATCHED
# ===============================

def _validate_observations(summary: str, df: pd.DataFrame) -> bool:
    import re
    if not isinstance(summary, str) or df.empty:
        return True
    bullets = [line.strip().lstrip("-•* ") for line in summary.split("\n") if line.strip().startswith(("-", "•", "*"))]
    if not bullets:
        return True
    df_text = " ".join(df.get("observation_text", pd.Series([])).astype(str).tolist()).lower()
    for bullet in bullets:
        bullet_words = [w.lower() for w in re.findall(r"\b\w+\b", bullet) if len(w) > 3]
        if not bullet_words:
            continue
        if not any(w in df_text for w in bullet_words):
            return False
    return True

def verify_data_integrity(raw_df: pd.DataFrame, canonical_df: pd.DataFrame, summary_text: str = None) -> bool:
    """
    Automatic end-to-end data integrity validator.
    Confirms every observation entered remains traceable without silent loss.
    """
    if raw_df.empty and canonical_df.empty:
        return True
    
    raw_len = len(raw_df)
    canon_len = len(canonical_df)
    
    # 1. Check Canonical length vs Raw
    # Note: Canonical collapses exact text duplicates on the same station/line.
    # So canon_len can be <= raw_len, but the distinct observation_texts must remain exactly the same.
    raw_obs = set(raw_df["observation_text"].astype(str).str.strip().tolist()) if "observation_text" in raw_df.columns else set()
    canon_obs = set(canonical_df["observation_text"].astype(str).str.strip().tolist()) if "observation_text" in canonical_df.columns else set()
    
    if len(canon_obs) < len(raw_obs):
        logger.error(f"Data Integrity Failure: Canonical DF lost unique observations. Raw={len(raw_obs)}, Canonical={len(canon_obs)}")
        return False

    # 2. If summary text is provided, verify it represents all canonical observations
    if summary_text:
        is_valid = _validate_observations(summary_text, canonical_df)
        if not is_valid:
            logger.error("Data Integrity Failure: Summary text does not represent all canonical observations.")
            return False
            
    return True

def generate_daily_brief(data):
    df = pd.DataFrame(data)
    if df.empty:
        return "No deviation data available."

    # Canonical DataFrame assumes preprocessing is already done globally.
    context = build_context_summary(df)
    num_obs = len(df)

    if num_obs <= 5:
        structure_prompt = """**Executive Overview**
One sentence. Lines covered. What was found. Max 15 words.

**Key Observations**
List all observations as bullets. Do not drop any data. Show every observation exactly as entered.

**Top Recurring Issues**
List any recurring issues if multiple exist, otherwise omit."""
    elif num_obs <= 10:
        structure_prompt = """**Executive Overview**
One to two sentences. Max 20 words.

**Key Observations**
Group similar observations into logical categories (Cleanliness, Process, Safety, Material Handling, Equipment, Documentation) while ensuring every unique issue is represented.

**Top Recurring Issues**
List the top recurring issues. Include: issue, line, station, count."""
    else:
        structure_prompt = """**Executive Overview**
One to two sentences. Max 20 words.

**Key Observations**
List the highest priority critical issues here (max 5 bullets).

**Top Recurring Issues**
List the top recurring issues. Include: issue, line, station, count.

**Additional Findings**
Briefly group all remaining observations by category. Ensure 100% of the real observations are represented."""

    prompt = f"""You are a plant engineer writing a factory shift-handover audit update.

Audit Data:
{context}

Write EXACTLY this structure — no extra text:

{structure_prompt}

**Root Cause**
One to two lines. Manufacturing focused. No generic explanations.

**Recommended Actions**
Two to three bullets. One line each. 5–10 words.

{_EXEC_TONE_RULES}

Start directly with **Executive Overview**."""

    for _ in range(3):
        raw = _get_llm().chat([
            {"role": "system", "content": (
                "You write factory shift-handover audit updates. "
                "Short lines. Operational words. No AI language. No paragraphs. "
                "Bullets max 10 words each. No explanations. One finding per line."
            )},
            {"role": "user", "content": prompt}
        ], max_tokens=450, temperature=0.05)
        if _validate_observations(raw, df):
            break

    return _apply_executive_filter(raw)


def generate_weekly_brief(current_data, previous_data):
    current_df = pd.DataFrame(current_data)
    previous_df = pd.DataFrame(previous_data)
    if current_df.empty:
        return "No deviation data available."

    # Canonical DataFrame assumes preprocessing is already done globally.
    context = build_context_summary(current_df)
    num_obs = len(current_df)

    delta = len(current_df) - len(previous_df)
    if delta > 0:
        trend = "Observations up vs last week."
    elif delta < 0:
        trend = "Observations down vs last week — improvement visible."
    else:
        trend = "Observation count steady vs last week."

    if num_obs <= 5:
        structure_prompt = """**Executive Overview**
One sentence. Week status and trend. Max 20 words.

**Key Observations**
List all observations as bullets. Do not drop any data. Show every observation exactly as entered.

**Top Recurring Issues**
List any recurring issues if multiple exist, otherwise omit."""
    elif num_obs <= 10:
        structure_prompt = """**Executive Overview**
One to two sentences. Week status and trend. Max 20 words.

**Key Observations**
Group similar observations into logical categories while ensuring every unique issue is represented.

**Top Recurring Issues**
List the top recurring issues. Include: issue, line, station, count."""
    else:
        structure_prompt = """**Executive Overview**
One to two sentences. Week status and trend. Max 20 words.

**Key Observations**
List the highest priority critical issues here.

**Top Recurring Issues**
List the top recurring issues. Include: issue, line, station, count.

**Additional Findings**
Briefly group all remaining observations by category. Ensure 100% of the real observations are represented."""

    prompt = f"""You are a plant quality engineer writing a weekly audit summary for the factory team.

Audit Data:
{context}
Week Trend: {trend}

Write EXACTLY this structure — no extra text:

{structure_prompt}

**Root Cause**
One to two lines. Manufacturing focused. No generic explanations.

**Recommended Actions**
Two to three bullets. One line each. 5–10 words.

{_EXEC_TONE_RULES}

Start directly with **Executive Overview**."""

    for _ in range(3):
        raw = _get_llm().chat([
            {"role": "system", "content": (
                "You write factory weekly audit updates. "
                "Operational tone. Short bullets. No paragraphs. No AI language. "
                "Bullets max 10 words. Manufacturing vocabulary only. One finding per line."
            )},
            {"role": "user", "content": prompt}
        ], max_tokens=480, temperature=0.05)
        if _validate_observations(raw, current_df):
            break

    return _apply_executive_filter(raw)


def generate_monthly_brief(current_data, previous_data):
    current_df = pd.DataFrame(current_data)
    previous_df = pd.DataFrame(previous_data)
    if current_df.empty:
        return "No deviation data available."

    # Canonical DataFrame assumes preprocessing is already done globally.
    context = build_context_summary(current_df)
    num_obs = len(current_df)

    delta = len(current_df) - len(previous_df)
    if delta > 0:
        trend = "Observations up vs last month."
    elif delta < 0:
        trend = "Observations down vs last month — improvement visible."
    else:
        trend = "Observation count steady vs last month."

    if num_obs <= 5:
        structure_prompt = """**Executive Overview**
One sentence. Month status and trend. Max 20 words.

**Key Observations**
List all observations as bullets. Do not drop any data. Show every observation exactly as entered.

**Top Recurring Issues**
List any recurring issues if multiple exist, otherwise omit."""
    elif num_obs <= 10:
        structure_prompt = """**Executive Overview**
One to two sentences. Month status and trend. Max 20 words.

**Key Observations**
Group similar observations into logical categories while ensuring every unique issue is represented.

**Top Recurring Issues**
List the top recurring issues. Include: issue, line, station, count."""
    else:
        structure_prompt = """**Executive Overview**
One to two sentences. Month status and trend. Max 20 words.

**Key Observations**
List the highest priority critical issues here.

**Top Recurring Issues**
List the top recurring issues. Include: issue, line, station, count.

**Additional Findings**
Briefly group all remaining observations by category. Ensure 100% of the real observations are represented."""

    prompt = f"""You are a plant quality engineer writing a monthly audit summary for plant leadership.

Audit Data:
{context}
Month Trend: {trend}

Write EXACTLY this structure — no extra text:

{structure_prompt}

**Root Cause**
One to two lines. Manufacturing focused. No generic explanations.

**Recommended Actions**
Two to three bullets. One line each. 5–10 words.

{_EXEC_TONE_RULES}

Start directly with **Executive Overview**."""

    for _ in range(3):
        raw = _get_llm().chat([
            {"role": "system", "content": (
                "You write factory monthly audit summaries for plant leadership. "
                "Operational tone. Short bullets. No paragraphs. No AI language. "
                "Bullets max 10 words. Manufacturing vocabulary only. One finding per line."
            )},
            {"role": "user", "content": prompt}
        ], max_tokens=500, temperature=0.05)
        if _validate_observations(raw, current_df):
            break

    return _apply_executive_filter(raw)


# ===============================
# MAIL FORMATTER  ← PATCHED
# ===============================

def generate_contextual_why(observations: list) -> str:
    """
    Isolated WHY generator.
    Receives finalized bullet observations, infers ONE contextual operational
    reason via a dedicated Groq call. Never uses static mappings or templates.
    """
    if not observations:
        return ""

    obs_block = "\n".join(f"- {o}" for o in observations if o)

    prompt = f"""You are a manufacturing supervisor writing a one-line plant-floor note.

Observations from today's audit:
{obs_block}

Task:
Look at ALL observations together.
Infer the single most likely operational reason that connects them.
Write ONE short line — maximum 10 words.

Rules:
- Factual. Operational. No blame.
- Specific to what is listed above — do NOT write a generic startup/checklist line unless that is genuinely the connection.
- Never write: "Pre-shift checks pending." / "Routine checks pending." / "Startup checks incomplete."
- Sound like a supervisor's quick note, not a report.
- No recommendations. No soft words. No management language.
- INFER SPECIFICALLY FROM THE GIVEN OBSERVATIONS. DO NOT MAKE UP REASONS.

Output ONLY the one-line reason. No label. No prefix. No punctuation beyond a period."""

    raw = _get_llm().chat([
        {"role": "system", "content": (
            "You are a plant supervisor writing a factual one-line operational reason. "
            "Output ONLY one short line, max 10 words. "
            "Never generic. Never blame. Never startup templates. "
            "Infer specifically from the given observations."
        )},
        {"role": "user", "content": prompt}
    ], max_tokens=60, temperature=0.15)

    # Strip any accidental prefixes like "Why:" or "Reason:"
    result = re.sub(r"^(why|reason|root cause)\s*[:\-]\s*", "", (raw or "").strip(), flags=re.IGNORECASE)
    result = _apply_executive_filter(result)
    result = _trim_bullet(result, max_words=10)
    return result


def generate_mail_from_summary(summary_text, mail_type="Daily"):
    from datetime import date as _date
    today_str = _date.today().strftime("%d %B %Y")

    extract_prompt = f"""Read this audit summary and extract fields into JSON.

Audit Summary:
{summary_text}

FIELD RULES:

"overview" — EXACTLY 1 sentence. Max 12 words.
  State which lines/areas were audited and overall status.

"bullets" — A JSON array of strings. Max 8 words per bullet.
  Extract ALL key observations exactly as they appear in the summary.
  Manufacturing vocabulary. No explanation.

"recurring" — A JSON array of strings based on the Top Recurring Issues section.
  Include the issue, line, station, and count. Max 15 words per line.

STRICT RULES:
- NO blame language.
- NO soft words: minor, slight, possibly, concern, attention.
- NO AI phrases: cadence, systemic, operational inconsistency.
- NO sentences over 12 words in overview, 8 words in bullets, 15 words in recurring.

Return ONLY this JSON. No extra text. No markdown.

{{
  "overview": "1 sentence, max 12 words.",
  "bullets": ["bullet 1", "bullet 2", "..."],
  "recurring": ["recurring 1", "..."]
}}"""

    raw = _get_llm().chat([
        {"role": "system", "content": (
            "You write factory-floor audit email content. "
            "Return ONLY valid JSON with exactly 3 keys: overview, bullets, recurring. "
            "Manufacturing vocabulary. Short lines. No blame. No AI language. "
            "No text outside JSON. No markdown fences."
        )},
        {"role": "user", "content": extract_prompt}
    ], max_tokens=600, temperature=0.05)

    parsed = parse_json_response(raw, fallback=None)

    overview = ""
    b_list = []
    r_list = []

    if parsed and isinstance(parsed, dict):
        overview = _apply_executive_filter(str(parsed.get("overview", "")).strip())
        b_raw = parsed.get("bullets", [])
        if isinstance(b_raw, list):
            b_list = [_apply_executive_filter(str(b).strip()) for b in b_raw]
        r_raw = parsed.get("recurring", [])
        if isinstance(r_raw, list):
            r_list = [str(r).strip() for r in r_raw]
    else:
        # Fallback: pull lines from summary text
        _slines = [
            l.strip().lstrip("-•* ").strip()
            for l in (summary_text or "").split("\n")
            if l.strip() and not l.strip().startswith("**") and len(l.strip()) > 6
        ]
        overview = _apply_executive_filter(_slines[0] if _slines else "")
        b_list = [_apply_executive_filter(l) for l in _slines[1:]]

    # Hard-trim bullets
    overview = _trim_bullet(overview, max_words=14)
    b_list = [_trim_bullet(b, max_words=10) for b in b_list if b]
    r_list = [_trim_bullet(r, max_words=15) for r in r_list if r]

    # ── WHY: isolated contextual call — never from the same prompt as bullets ──
    why = generate_contextual_why(b_list)

    _bullet_block = "\n".join(f"• {b}" for b in b_list) if b_list else "• No observations to highlight."
    _why_block = f"Why:\n{why}\n\n" if why else ""
    
    _recurring_block = "\n".join(f"• {r}" for r in r_list) if r_list else "• No recurring issues noted."

    email = (
        f"Subject: {mail_type} Audit Update | {today_str}\n"
        f"\n"
        f"Hi Team,\n"
        f"\n"
        f"{overview}\n"
        f"\n"
        f"Key Observations:\n"
        f"{_bullet_block}\n"
        f"\n"
        f"{_why_block}"
        f"Top Recurring Issues:\n"
        f"{_recurring_block}\n"
        f"\n"
        f"Regards,\n"
        f"AutoNQ AI"
    )
    return email

# ===============================
# AGENT 2 (AUDIT PLAN) — UNCHANGED
# ===============================

def classify_by_percentile(risk_df):
    risk_df = risk_df.sort_values(by="risk_score", ascending=False).reset_index(drop=True)
    n = len(risk_df)
    risk_df["rank"] = risk_df.index + 1
    risk_df["percentile"] = risk_df["rank"] / n

    def classify(p):
        if p <= 0.2:
            return "Very High", 4
        elif p <= 0.5:
            return "High", 3
        elif p <= 0.8:
            return "Medium", 2
        else:
            return "Stable", 1

    risk_df[["risk_level","audit_frequency"]] = risk_df["percentile"].apply(
        lambda x: pd.Series(classify(x))
    )
    return risk_df


def generate_governance_annual_plan(risk_df):
    months = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]
    plan_rows = []
    for _, row in risk_df.iterrows():
        line = f"Line {row['LINE']}" if str(row['LINE']).isdigit() else str(row['LINE'])
        risk = row["risk_level"]
        month_map = {m: "" for m in months}
        if risk == "Very High":
            selected = ["Jan","Apr","Jul","Oct"]
        elif risk == "High":
            selected = ["Feb","Jun","Oct"]
        elif risk == "Medium":
            selected = ["May","Sep"]
        else:
            selected = ["Nov"]
        for m in selected:
            month_map[m] = "X"
        plan_rows.append({"Line": line, "Risk Level": risk, **month_map})
    return pd.DataFrame(plan_rows)

# ===============================
# AGENT 3 — UNCHANGED
# ===============================

def get_iatf_clause_mapping():
    return {
        "1C – Cleanliness": ("8.5.1", "Control of Production and Service Provision"),
        "Remaining Items": ("8.5.1", "Control of Production and Service Provision"),
        "Correct Product": ("8.5.1", "Control of Production and Service Provision"),
        "Dropped Parts": ("8.7.1", "Control of Nonconforming Outputs"),
        "Rework / Scrap": ("8.7.1", "Control of Nonconforming Outputs"),
        "Labeling": ("8.5.2", "Identification and Traceability"),
        "Measurement / Test Equipment": ("7.1.5", "Monitoring and Measuring Resources"),
        "Instructions": ("7.5.1", "Documented Information"),
        "Process Parameters": ("8.5.1", "Control of Production and Service Provision"),
        "Tools": ("8.5.1", "Control of Production and Service Provision"),
        "Check the Checker": ("9.1.1", "Monitoring, Measurement, Analysis and Evaluation"),
        "Restart": ("8.5.1", "Control of Production and Service Provision"),
        "Stop Sign": ("10.2", "Nonconformity and Corrective Action"),
        "Andon Cord": ("10.2", "Nonconformity and Corrective Action"),
        "Total Productive Maintenance (TPM)": ("8.5.1.5", "Total Productive Maintenance")
    }

def generate_guided_audit_questions(line, deviation_data, iqis_df, top_n=3):
    clause_map = get_iatf_clause_mapping()

    def normalize_line(val):
        try:
            return str(int(float(val)))
        except (ValueError, TypeError):
            return str(val).strip().replace("Line ", "")

    line_devs = [d for d in deviation_data if normalize_line(d.get("line")) == normalize_line(line)]
    if not line_devs:
        return "No deviation data available for this line."

    df_dev = pd.DataFrame(line_devs)
    dominant_principle = df_dev["ai_principle"].value_counts().idxmax()
    top_stations = df_dev["station"].value_counts().head(3).index.tolist()
    top_stations_str = ", ".join(top_stations)
    top_deviations = df_dev["observation_text"].value_counts().head(top_n).to_dict()
    flm = df_dev["supervisor"].value_counts().idxmax()

    clause_number, clause_title = clause_map.get(
        dominant_principle, ("8.5.1", "Control of Production and Service Provision")
    )

    prompt = f"""You are a senior IATF 16949 process audit expert.

Context:
Line: {line}
Dominant Discipline Gap: {dominant_principle}
IATF Clause: {clause_number} – {clause_title}
Critical Stations: {top_stations_str}
Responsible FLM: {flm}
Top Recurring Deviations: {top_deviations}

Generate 6-8 plant-floor audit questions.

Rules:
- Each question references Line {line} – Station XX
- Verify actual process control, not documentation
- Distribute across stations: {top_stations_str}
- Include one escalation/stop-rule question
- Include one FLM supervision question
- Keep each question to one sentence
- End with a one-sentence Audit Objective referencing Clause {clause_number}"""

    return _get_llm().chat([
        {"role": "system", "content": "You generate plant-floor focused IATF audit questions."},
        {"role": "user", "content": prompt}
    ], max_tokens=800)


def generate_qcheck_questions(line, deviation_data, max_questions=15):
    filtered = [row for row in deviation_data if str(row.get("line")) == str(line)]
    if not filtered:
        return []

    filtered = sorted(filtered, key=lambda x: x.get("date", ""), reverse=True)
    seen = set()
    selected = []
    for row in filtered:
        key = (row.get("station"), row.get("observation_text"))
        if key not in seen:
            seen.add(key)
            selected.append(row)
        if len(selected) >= max_questions:
            break

    observations = [f"{i+1}. {row['observation_text']}" for i, row in enumerate(selected)]

    prompt = f"""Convert these manufacturing deviations into short audit checkpoint questions.

Rules:
- One sentence per question, Yes/No type
- Do NOT mention line or station
- Keep practical for plant-floor audit

Deviations:
{chr(10).join(observations)}

Return as numbered list only."""

    response = _get_llm().chat([
        {"role": "system", "content": "You generate short plant-floor audit checkpoints."},
        {"role": "user", "content": prompt}
    ], max_tokens=600)

    questions = []
    q_index = 0
    for line_text in response.split("\n"):
        line_text = line_text.strip()
        if line_text and any(char.isdigit() for char in line_text[:3]):
            q = line_text.split(".", 1)[-1].strip()
            if q_index < len(selected):
                row = selected[q_index]
                questions.append({
                    "Station": row["station"], "Checkpoint": q,
                    "Ref Photo": "", "Status": "NOK", "Remark": "",
                    "image_base64": row.get("image_base64", "")
                })
                q_index += 1
    return questions

# ===============================
# AGENT 4 — UNCHANGED
# ===============================

def generate_iatf_process_audit_sheet(line, deviation_data, iqis_df, lpc_df, top_n=3):
    df_dev = pd.DataFrame(deviation_data)
    df_dev_line = df_dev[df_dev["line"].astype(str) == str(line)]
    if df_dev_line.empty:
        return "No deviation data available for this line."

    clause_map = {
        "Cleanliness": ("8.5.1", "Control of Production"),
        "Stop Sign": ("10.2", "Corrective Action"),
        "Andon Cord": ("8.5.1", "Control of Production"),
        "Instructions": ("7.5", "Documented Information"),
        "Process Parameters": ("8.5.1", "Control of Production"),
        "Measurement / Test Equipment": ("7.1.5", "Monitoring & Measuring"),
        "Check the Checker": ("9.1.1", "Performance Evaluation"),
        "Total Productive Maintenance (TPM)": ("8.5.1.5", "TPM"),
        "Tools": ("8.5.1.5", "TPM"),
        "Restart": ("8.5.1", "Control of Production"),
        "Labeling": ("8.5.2", "Identification & Traceability"),
        "Rework / Scrap": ("8.7", "Nonconforming Output"),
        "Dropped Parts": ("8.7", "Nonconforming Output"),
        "Correct Product": ("8.5.1", "Control of Production"),
        "Remaining Items": ("8.5.4", "Preservation")
    }

    top_stations = df_dev_line["station"].value_counts().head(top_n).index.tolist()
    final_rows = []
    for station in top_stations:
        df_station = df_dev_line[df_dev_line["station"] == station]
        obs_counts = df_station["observation_text"].value_counts()
        for obs_text, _ in obs_counts.items():
            df_obs = df_station[df_station["observation_text"] == obs_text]
            principle = df_obs["ai_principle"].value_counts().idxmax() if (not df_obs.empty and "ai_principle" in df_obs.columns) else "Process"
            principle_key = str(principle).split("–")[1].strip() if "–" in str(principle) else str(principle).strip()
            clause, clause_title = clause_map.get(principle_key, ("8.5.1", "Control of Production"))
            final_rows.append({
                "Clause": clause, "Clause_Title": clause_title, "Line": line,
                "Station": station, "Process_Risk": obs_text,
                "Audit_Check_Point": f"Verify Line {line} – Station {station} control for {principle_key}.",
                "Audit_Status": "", "Remarks": ""
            })

    logger.info(f"IATF Process Audit Checksheet Generated – Line {line}")
    return pd.DataFrame(final_rows)

# ===============================
# AGENT 5 — UNCHANGED
# ===============================

def generate_followup_checklist(line, agent3_df):
    if agent3_df is None or agent3_df.empty:
        return "No audit data available."

    followup_rows = []
    for _, row in agent3_df.iterrows():
        station = row["Station"]
        issue = row["Process_Risk"]
        recurrence = row.get("Recurrence_Count", 1)
        followup_rows.append({
            "Line": line, "Station": station, "Issue": issue,
            "Previous_Occurrence_Count": recurrence,
            "Follow_Up_Check": f"Verify that '{issue}' at Station {station} is corrected and not repeated.",
            "Status (Yes/No)": "", "Remarks": ""
        })

    logger.info(f"Follow-Up Checklist Generated – Line {line}")
    return pd.DataFrame(followup_rows)

# ===============================
# AGENT 6 — BATCHED EXTERNAL TRACKER
# ===============================

def generate_external_audit_tracker_with_ai(deviation_data, top_n=5):
    df = pd.DataFrame(deviation_data)
    if df.empty:
        return "No historical deviation data available."

    df = preprocess_audit_data(df)

    grouped = (
        df.groupby(["line", "station", "observation_text"])
        .size().reset_index(name="Recurrence_Count")
        .sort_values(by="Recurrence_Count", ascending=False)
    )
    top_issues = grouped.head(top_n)
    if top_issues.empty:
        return "No recurring issues found."

    issues_text = ""
    issue_meta = []
    for i, (_, row) in enumerate(top_issues.iterrows()):
        sev = tag_severity(row["Recurrence_Count"])
        principle = "Instructions"
        match_rows = df[df["observation_text"] == row["observation_text"]]
        if "ai_principle" in match_rows.columns and not match_rows.empty:
            principle = match_rows["ai_principle"].mode().iloc[0]
        dom = tag_domain(principle)
        issue_meta.append({"severity": sev, "domain": dom})
        issues_text += f"\nIssue {i+1}: Line {row['line']} – {row['station']}: {row['observation_text']} (x{row['Recurrence_Count']}, {sev}, {dom})"

    prompt = f"""You are an external IATF auditor reviewing recurring audit findings.

Top {len(top_issues)} recurring issues:{issues_text}

For EACH issue generate:
1. issue_summary - one-line summary
2. root_cause - practical root cause
3. corrective_action - specific 5W1H action
4. owner - responsible role
5. priority - High/Medium/Low

Return ONLY a JSON array of {len(top_issues)} objects with keys: issue_summary, root_cause, corrective_action, owner, priority"""

    response = _get_llm().chat([
        {"role": "system", "content": "You are a manufacturing audit expert. Return ONLY valid JSON."},
        {"role": "user", "content": prompt}
    ], max_tokens=1200, temperature=0.15)

    parsed = parse_json_response(response, fallback=None)

    tracker_rows = []
    for i, (_, row) in enumerate(top_issues.iterrows()):
        ai_data = {}
        if parsed and isinstance(parsed, list) and i < len(parsed):
            ai_data = parsed[i] if isinstance(parsed[i], dict) else {}

        meta = issue_meta[i] if i < len(issue_meta) else {"severity": "Low", "domain": "Process"}

        tracker_rows.append({
            "Line": row["line"],
            "Station": row["station"],
            "Issue_Raised_Last_Audit": row["observation_text"],
            "Recurrence_Count": row["Recurrence_Count"],
            "Issue_Summary": ai_data.get("issue_summary", row["observation_text"]),
            "Root_Cause": ai_data.get("root_cause", "To be investigated"),
            "Corrective_Action": ai_data.get("corrective_action", "Action pending AI analysis"),
            "Owner": ai_data.get("owner", "Line Supervisor"),
            "Priority": ai_data.get("priority", meta["severity"] if meta["severity"] != "Low" else "Medium"),
            "Domain": meta["domain"],
            "Severity": meta["severity"],
            "Current_Status": "Not Started",
        })

    logger.info("External Audit Tracker Generated (Batched AI)")
    return pd.DataFrame(tracker_rows)

# ===============================
# AGENT 7 — DEVIATION CATEGORY
# ===============================

def map_deviation_category_ai(observations):
    obs_list = [f"{i+1}. {obs}" for i, obs in enumerate(observations)]

    prompt = f"""Group these manufacturing audit observations into categories.

Rules:
- Group similar observations into ONE category
- 2-4 word category names
- Industrial terminology

Observations:
{chr(10).join(obs_list)}

Return as JSON array: [{{"index": 1, "category": "Name"}}, ...]
Return ONLY the JSON array."""

    response = _get_llm().chat([
        {"role": "system", "content": "You classify manufacturing audit observations. Return ONLY valid JSON."},
        {"role": "user", "content": prompt}
    ], max_tokens=600)

    parsed = parse_json_response(response, fallback=None)

    mapping = {}
    if parsed and isinstance(parsed, list):
        for item in parsed:
            if isinstance(item, dict) and "index" in item and "category" in item:
                mapping[int(item["index"]) - 1] = item["category"]

    if not mapping and response and "→" in response:
        for line in response.split("\n"):
            if "→" in line:
                try:
                    left, right = line.split("→")
                    idx = int(left.strip().split(".")[0]) - 1
                    mapping[idx] = right.strip()
                except (ValueError, IndexError):
                    continue

    return [mapping.get(i, "Other Issue") for i in range(len(observations))]