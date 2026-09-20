"""
Agency Lookup
=============

For each agency Name in your Excel sheet, this:

  1. Builds a list of candidate search keywords from the name (most
     distinctive words first), each truncated to a short stem (e.g.
     "Judicial" -> "Judicia", "Commission" -> "Commiss") so short/abbreviated
     database forms still get caught.

  2. Tries ONE keyword at a time against mai.MasterAgency:

       SELECT ComputedId, OrganizationId, Name, Address1, City,
              StateAbbreviation, Zip, Phone, Website
       FROM organization_master
       WHERE StateAbbreviation = ? AND SoftDelete = 0 AND Name LIKE '%kw%'

     If that keyword returns nothing that verifies as the same agency (see
     step 3), it moves on to the NEXT keyword and tries again — it does NOT
     stop after just one keyword, since the real database record might not
     share every word with the input name. Only after every candidate
     keyword has been tried with no verified match does it report
     "Not Found". Same approach for dbo.organizations (AccountName).

  3. Name verification: a SQL hit only counts as a real match if the
     candidate's Name/AccountName actually corresponds to the input name,
     allowing for state abbreviations ("AL" <-> "Alabama") and short/
     abbreviated forms ("AL Judicial Inquiry Comm" <-> "Alabama Judicial
     Inquiry Commission"). This stops an unrelated agency that happens to
     share one generic word from being reported as a match.

  4. Writes the result back into the SAME Excel sheet, in new columns:
       - MasterAgency_OrganizationId  (or "Not Found" / "Multiple Found")
       - MasterAgency_MatchedName
       - MasterAgency_ComputedId
       - MasterAgency_KeywordUsed     (which keyword actually found it)
       - MasterAgency_Confidence      (High / Medium / Low)
       - Organizations_Id             (or "Not Found" / "Multiple Found")
       - Organizations_MatchedName
       - Organizations_ComputedId
       - Organizations_KeywordUsed
       - Organizations_Confidence     (High / Medium / Low)

     Confidence meaning:
       High   = essentially an exact match (all distinctive words lined up)
       Medium = matched, but not word-for-word — or multiple candidates
                verified and it's unclear which one is correct — worth a
                quick manual look
       Low    = nothing matched (same as "Not Found")

This script does not insert, update, or delete anything in the database.
It only reads and reports.


Setup
-----
pip install pandas pyodbc python-dotenv openpyxl

password.env file needs:
    SQL_CONN_STRING=<your pyodbc connection string>
(or fill in SQL_SERVER / SQL_DATABASE / etc. directly below instead)


Input file
----------
Your Excel file needs a "Name" column and a "State" column with the
2-letter state abbreviation used in the database (e.g. "AL", "CA").

Update DEFAULT_INPUT_PATH / DEFAULT_SHEET_NAME below to point at your file,
or pass --input / --sheet on the command line.


Usage
-----
    python Judicial_stategov_matching_audit.py
        # uses the defaults set below

    python Judicial_stategov_matching_audit.py --input "C:\path\to\file.xlsx" --sheet "Sheet1"
"""

import argparse
import os
import re
from datetime import datetime
from pathlib import Path

import pandas as pd
import pyodbc
from dotenv import load_dotenv

# Looks for password.env in the same folder as this script, no matter
# where PyCharm's working directory happens to be.
load_dotenv(Path(__file__).resolve().parent / "password.env")

# --- Configuration ---
# Option A: put the full connection string in password.env as SQL_CONN_STRING.
SQL_CONN_STRING = os.getenv("SQL_CONN_STRING")

# Database credentials are loaded from password.env.
# Do NOT commit password.env or hardcoded credentials to GitHub.
SQL_DRIVER = os.getenv("SQL_DRIVER", "ODBC Driver 17 for SQL Server")
SQL_SERVER = os.getenv("SQL_SERVER")
SQL_DATABASE = os.getenv("SQL_DATABASE")
SQL_USERNAME = os.getenv("SQL_USERNAME")
SQL_PASSWORD = os.getenv("SQL_PASSWORD")

# Update these to match your actual file / sheet name, or pass
# --input / --sheet on the command line instead.
DEFAULT_INPUT_PATH = "input.xlsx"
DEFAULT_SHEET_NAME = "Sheet1"

MIN_KEYWORD_LENGTH = 3     # ignore words shorter than this as keyword candidates
STEM_SEARCH_LENGTH = 7     # truncate each keyword to its first N letters for the LIKE search

NOISE_WORDS = {
    "the", "of", "and", "a", "an", "for", "at", "in", "on", "st", "de", "la",
}

# Used to treat a state abbreviation ("AL") and its full name ("Alabama") as
# the same word, and to exclude the state name from search keywords
# entirely (DB records often omit it since State is already a column).
US_STATES = {
    "AL": "Alabama", "AK": "Alaska", "AZ": "Arizona", "AR": "Arkansas", "CA": "California",
    "CO": "Colorado", "CT": "Connecticut", "DE": "Delaware", "FL": "Florida", "GA": "Georgia",
    "HI": "Hawaii", "ID": "Idaho", "IL": "Illinois", "IN": "Indiana", "IA": "Iowa",
    "KS": "Kansas", "KY": "Kentucky", "LA": "Louisiana", "ME": "Maine", "MD": "Maryland",
    "MA": "Massachusetts", "MI": "Michigan", "MN": "Minnesota", "MS": "Mississippi", "MO": "Missouri",
    "MT": "Montana", "NE": "Nebraska", "NV": "Nevada", "NH": "New Hampshire", "NJ": "New Jersey",
    "NM": "New Mexico", "NY": "New York", "NC": "North Carolina", "ND": "North Dakota", "OH": "Ohio",
    "OK": "Oklahoma", "OR": "Oregon", "PA": "Pennsylvania", "RI": "Rhode Island", "SC": "South Carolina",
    "SD": "South Dakota", "TN": "Tennessee", "TX": "Texas", "UT": "Utah", "VT": "Vermont",
    "VA": "Virginia", "WA": "Washington", "WV": "West Virginia", "WI": "Wisconsin", "WY": "Wyoming",
    "DC": "District of Columbia",
}

# How much of the shorter name's words need to show up in the other name
# (after normalizing abbreviations) before we accept it as the same agency.
NAME_MATCH_THRESHOLD = 0.85


# --------------------------------------------------------------------------
# DB helper
# --------------------------------------------------------------------------

def build_conn_string():
    if SQL_CONN_STRING:
        return SQL_CONN_STRING
    if SQL_SERVER and SQL_DATABASE:
        parts = [f"Driver={{{SQL_DRIVER}}}", f"Server={SQL_SERVER}", f"Database={SQL_DATABASE}"]
        if SQL_USERNAME:
            parts.append(f"UID={SQL_USERNAME}")
            parts.append(f"PWD={SQL_PASSWORD}")
        else:
            parts.append("Trusted_Connection=yes")
        return ";".join(parts) + ";"
    return None


def get_db_connection():
    conn_str = build_conn_string()
    if not conn_str:
        raise RuntimeError(
            "No SQL connection info found. Either set SQL_CONN_STRING in "
            "password.env, or fill in SQL_SERVER / SQL_DATABASE (and "
            "SQL_USERNAME / SQL_PASSWORD if not using Windows Auth) near "
            "the top of this script."
        )
    return pyodbc.connect(conn_str)


# --------------------------------------------------------------------------
# Keyword candidates (tried one at a time, not combined)
# --------------------------------------------------------------------------

def generate_keyword_candidates(name, state=None, min_length=MIN_KEYWORD_LENGTH):
    """Returns an ordered list of whole words to try as search keywords,
    most distinctive first. The state's name/abbreviation is excluded
    (it's already filtered via its own SQL column, and DB records often
    omit it from the name itself)."""
    name = str(name)

    state_words = set()
    if state:
        state_abbr = str(state).strip().lower()
        state_name = US_STATES.get(str(state).strip().upper(), "")
        state_words.add(state_abbr)
        state_words.update(state_name.lower().split())

    def tokens_from(s):
        toks = re.findall(r"[A-Za-z']+", s)
        return [
            t for t in toks
            if t.lower() not in NOISE_WORDS
            and t.lower() not in state_words
            and len(t) >= min_length
        ]

    # Names like "First Judicial Circuit Court of Alabama - Clarke County"
    # have their most distinctive word (the county/branch) after a dash —
    # try those first.
    if " - " in name:
        base, suffix = name.split(" - ", 1)
    else:
        base, suffix = name, ""

    suffix_tokens = sorted(tokens_from(suffix), key=len, reverse=True)
    base_tokens = sorted(tokens_from(base), key=len, reverse=True)

    ordered = suffix_tokens + base_tokens
    seen, result = set(), []
    for t in ordered:
        key = t.lower()
        if key not in seen:
            seen.add(key)
            result.append(t)
    return result


def to_search_term(word, stem_length=STEM_SEARCH_LENGTH):
    """Truncates a word to its first N letters for the LIKE search, e.g.
    "Judicial" -> "Judicia", "Commission" -> "Commiss". Short words are
    used as-is."""
    return word[:stem_length] if len(word) > stem_length else word


# --------------------------------------------------------------------------
# Name verification (confirms a SQL candidate is actually the same agency)
# --------------------------------------------------------------------------

def _normalize_tokens(name, state):
    """Tokenizes a name and drops filler words AND the state's own
    name/abbreviation entirely (it's already enforced via the SQL State
    column, so it shouldn't count for or against a name match)."""
    tokens = re.findall(r"[A-Za-z']+", str(name).lower())
    state_abbr = str(state).lower().strip()
    state_name_words = set(US_STATES.get(str(state).upper().strip(), "").lower().split())
    drop = {state_abbr} | state_name_words
    return [t for t in tokens if t not in NOISE_WORDS and t not in drop]


def _tokens_equivalent(t1, t2, min_prefix_len=4):
    """True if the words are the same, or one is a genuine PREFIX of the
    other (e.g. "Comm" of "Commission") — not just sharing the same first
    few letters, which would wrongly equate "Marin" with "Mariposa"."""
    if t1 == t2:
        return True
    shorter, longer = (t1, t2) if len(t1) <= len(t2) else (t2, t1)
    if len(shorter) < min_prefix_len:
        return False
    return longer.startswith(shorter)


def name_match_score(input_name, candidate_name, state):
    """Fraction (0.0-1.0) of the shorter name's distinctive words that are
    found (exactly or as a prefix) in the other name."""
    in_tokens = _normalize_tokens(input_name, state)
    cand_tokens = _normalize_tokens(candidate_name, state)

    if not in_tokens or not cand_tokens:
        return 0.0

    smaller, larger = (in_tokens, cand_tokens) if len(in_tokens) <= len(cand_tokens) else (cand_tokens, in_tokens)

    matched, used = 0, set()
    for t in smaller:
        for j, u in enumerate(larger):
            if j in used:
                continue
            if _tokens_equivalent(t, u):
                matched += 1
                used.add(j)
                break

    return round(matched / len(smaller), 2)


def name_matches(input_name, candidate_name, state, threshold=NAME_MATCH_THRESHOLD):
    """True if candidate_name is plausibly the same agency as input_name,
    allowing for state abbreviations and short/abbreviated forms."""
    return name_match_score(input_name, candidate_name, state) >= threshold


# --------------------------------------------------------------------------
# SQL searches — ONE keyword at a time
# --------------------------------------------------------------------------

def search_master_agency(conn, term, state):
    query = """
        SELECT ComputedId, OrganizationId, Name, Address1, City,
               StateAbbreviation, Zip, Phone, Website
        FROM organization_master
        WHERE StateAbbreviation = ? AND SoftDelete = 0 AND Name LIKE ?
    """
    return pd.read_sql(query, conn, params=[state, f"%{term}%"])


def search_organizations(conn, term, state):
    query = """
        SELECT Id, ComputedId, AccountName, Address1, City, State,
               PhoneNumber, WebSite
        FROM organization_records
        WHERE SoftDelete = 0 AND State = ? AND AccountName LIKE ?
    """
    return pd.read_sql(query, conn, params=[state, f"%{term}%"])


# --------------------------------------------------------------------------
# Result summarizing (shared by both tables)
# --------------------------------------------------------------------------

def _fmt_id(x):
    if pd.isna(x):
        return None
    try:
        return str(int(x))
    except (ValueError, TypeError):
        return str(x)


def confidence_label(score):
    """Converts a numeric match score into a plain-English confidence label:
    High = all distinctive tokens in the shorter name matched, Medium = matched but not
    word-for-word (worth a glance), Low = nothing matched."""
    try:
        s = float(score)
    except (TypeError, ValueError):
        return "Low"
    if s >= 0.99:
        return "High"
    if s >= NAME_MATCH_THRESHOLD:
        return "Medium"
    return "Low"


def summarize_matches(df, id_col, name_col, computed_col):
    """Returns (id or status text, matched name(s), ComputedId(s), confidence label)."""
    if df.empty:
        return "Not Found", "", "", "Low"

    if computed_col in df.columns:
        df = df.drop_duplicates(subset=[computed_col])

    if len(df) == 1:
        row = df.iloc[0]
        computed_id = _fmt_id(row[computed_col]) or ""
        confidence = confidence_label(row.get("MatchScore"))
        id_val = _fmt_id(row[id_col])
        if id_val is None:
            return f"Found ({id_col} not set)", row[name_col], computed_id, confidence
        return id_val, row[name_col], computed_id, confidence

    # More than one candidate verified — even if each individually scores
    # high, we don't know WHICH one is correct, so this always needs a
    # human look regardless of the individual scores.
    id_list = [_fmt_id(x) for x in df[id_col].tolist()]
    id_list_clean = [x for x in id_list if x is not None]
    computed_list = [_fmt_id(x) for x in df[computed_col].tolist()]
    names = df[name_col].tolist()
    shown_ids = ", ".join(id_list_clean[:5]) + (f", +{len(id_list_clean) - 5} more" if len(id_list_clean) > 5 else "")
    shown_names = " | ".join(names[:5]) + (f" | +{len(names) - 5} more" if len(names) > 5 else "")
    shown_computed = ", ".join(c for c in computed_list[:5] if c) + (f", +{len(computed_list) - 5} more" if len(computed_list) > 5 else "")
    return f"Multiple Found ({len(df)} candidates: {shown_ids})", shown_names, shown_computed, "Medium"


def find_match(conn, search_fn, name_col, id_col, computed_col, input_name, state, keyword_candidates):
    """Tries each keyword ONE AT A TIME (never combined) until a
    name-verified match turns up. Only after every candidate keyword has
    been tried with nothing verifying does it give up.
    Returns ((id, matched_name, computed_id, confidence_label), keyword_that_worked)."""
    for word in keyword_candidates:
        term = to_search_term(word)
        df = search_fn(conn, term, state)
        if df.empty:
            continue
        df = df.copy()
        df["MatchScore"] = df[name_col].apply(lambda n: name_match_score(input_name, n, state))
        verified = df[df["MatchScore"] >= NAME_MATCH_THRESHOLD]
        if not verified.empty:
            return summarize_matches(verified, id_col, name_col, computed_col), term
    return ("Not Found", "", "", "Low"), None


# --------------------------------------------------------------------------
# Write-back
# --------------------------------------------------------------------------

def write_results(input_path, sheet_name, total_rows,
                   ma_ids, ma_names, ma_computed_ids, ma_keywords, ma_scores,
                   org_ids, org_names, org_computed_ids, org_keywords, org_scores):
    """Writes results collected so far back into the same sheet. Safe to
    call multiple times (checkpoints) — it just overwrites the same columns."""
    import openpyxl
    wb = openpyxl.load_workbook(input_path)
    ws = wb[sheet_name] if sheet_name else wb[wb.sheetnames[0]]

    start_col = ws.max_column + 1
    for col in range(1, ws.max_column + 1):
        if ws.cell(row=1, column=col).value == "MasterAgency_OrganizationId":
            start_col = col
            break

    headers = [
        "MasterAgency_OrganizationId", "MasterAgency_MatchedName",
        "MasterAgency_ComputedId", "MasterAgency_KeywordUsed", "MasterAgency_Confidence",
        "Organizations_Id", "Organizations_MatchedName",
        "Organizations_ComputedId", "Organizations_KeywordUsed", "Organizations_Confidence",
    ]
    for j, header in enumerate(headers):
        ws.cell(row=1, column=start_col + j, value=header)

    n_done = len(ma_ids)
    for i in range(total_rows):
        excel_row = i + 2  # header is row 1, data starts row 2
        if i < n_done:
            ws.cell(row=excel_row, column=start_col, value=str(ma_ids[i]))
            ws.cell(row=excel_row, column=start_col + 1, value=str(ma_names[i]))
            ws.cell(row=excel_row, column=start_col + 2, value=str(ma_computed_ids[i]))
            ws.cell(row=excel_row, column=start_col + 3, value=str(ma_keywords[i]))
            ws.cell(row=excel_row, column=start_col + 4, value=str(ma_scores[i]))
            ws.cell(row=excel_row, column=start_col + 5, value=str(org_ids[i]))
            ws.cell(row=excel_row, column=start_col + 6, value=str(org_names[i]))
            ws.cell(row=excel_row, column=start_col + 7, value=str(org_computed_ids[i]))
            ws.cell(row=excel_row, column=start_col + 8, value=str(org_keywords[i]))
            ws.cell(row=excel_row, column=start_col + 9, value=str(org_scores[i]))
        else:
            ws.cell(row=excel_row, column=start_col, value="Not processed yet")

    wb.save(input_path)


# --------------------------------------------------------------------------
# Main
# --------------------------------------------------------------------------

CHECKPOINT_EVERY = 5  # save progress to the Excel file every N rows

def run(input_path, sheet_name):
    df = pd.read_excel(input_path, sheet_name=sheet_name if sheet_name else 0)

    lower_map = {c.lower().strip(): c for c in df.columns}
    if "name" not in lower_map or "state" not in lower_map:
        raise ValueError(
            f"Couldn't find both a 'Name' and 'State' column. Found columns: {list(df.columns)}"
        )
    name_col = lower_map["name"]
    state_col = lower_map["state"]

    print(f"Loaded {len(df)} rows from {input_path} (sheet: {sheet_name})")

    backup_path = input_path.rsplit(".", 1)[0] + f"_backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    import shutil
    shutil.copy(input_path, backup_path)
    print(f"Backed up original file to {backup_path}")

    ma_ids, ma_names, ma_computed_ids, ma_keywords, ma_scores = [], [], [], [], []
    org_ids, org_names, org_computed_ids, org_keywords, org_scores = [], [], [], [], []

    try:
        with get_db_connection() as conn:
            for i, row in df.iterrows():
                name = str(row[name_col]).strip()
                state = str(row[state_col]).strip()
                keyword_candidates = generate_keyword_candidates(name, state=state)

                print(f"[{i+1}/{len(df)}] {name} ({state}) -- candidate keywords: {keyword_candidates}")

                if not keyword_candidates or not state:
                    ma_ids.append("Skipped - missing name/state/keywords")
                    ma_names.append("")
                    ma_computed_ids.append("")
                    ma_keywords.append("")
                    ma_scores.append("")
                    org_ids.append("Skipped - missing name/state/keywords")
                    org_names.append("")
                    org_computed_ids.append("")
                    org_keywords.append("")
                    org_scores.append("")
                    continue

                (ma_id, ma_name, ma_computed_id, ma_score), ma_kw = find_match(
                    conn, search_master_agency, "Name", "OrganizationId", "ComputedId",
                    name, state, keyword_candidates
                )
                (org_id, org_name, org_computed_id, org_score), org_kw = find_match(
                    conn, search_organizations, "AccountName", "Id", "ComputedId",
                    name, state, keyword_candidates
                )

                print(f"  MasterAgency: {ma_id} (via '{ma_kw}', score {ma_score})  |  organizations: {org_id} (via '{org_kw}', score {org_score})")

                ma_ids.append(ma_id)
                ma_names.append(ma_name)
                ma_computed_ids.append(ma_computed_id)
                ma_keywords.append(ma_kw or "none matched")
                ma_scores.append(ma_score)
                org_ids.append(org_id)
                org_names.append(org_name)
                org_computed_ids.append(org_computed_id)
                org_keywords.append(org_kw or "none matched")
                org_scores.append(org_score)

                if (i + 1) % CHECKPOINT_EVERY == 0:
                    write_results(input_path, sheet_name, len(df),
                                  ma_ids, ma_names, ma_computed_ids, ma_keywords, ma_scores,
                                  org_ids, org_names, org_computed_ids, org_keywords, org_scores)
                    print(f"  [checkpoint saved: {i+1}/{len(df)} rows written to {input_path}]")
    finally:
        write_results(input_path, sheet_name, len(df),
                      ma_ids, ma_names, ma_computed_ids, ma_keywords, ma_scores,
                      org_ids, org_names, org_computed_ids, org_keywords, org_scores)
        print(f"\nSaved {len(ma_ids)}/{len(df)} processed rows into {input_path} (sheet: {sheet_name})")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Look up agencies from an Excel sheet against MasterAgency and organizations tables.")
    parser.add_argument("--input", default=DEFAULT_INPUT_PATH, help=f"Path to input Excel file (default: {DEFAULT_INPUT_PATH})")
    parser.add_argument("--sheet", default=DEFAULT_SHEET_NAME, help=f"Sheet name to read/write (default: {DEFAULT_SHEET_NAME})")
    args = parser.parse_args()

    run(args.input, args.sheet)
