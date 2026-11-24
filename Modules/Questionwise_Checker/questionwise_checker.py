"""
==================================================
QUESTIONWISE CHECKER (MULTI-FILE & GMR INTEGRATED)
==================================================

PURPOSE
--------------------------------------------------
This Streamlit tool helps exam staff:

1. Switch between two analysis modes via a top-level toggle:
   - Moderation Analysis: Legacy/standard workflow across all students, outputting
     Cleaned File, Formatted File, College ZIP, and Consolidated Multi-Sheet Excel.
   - Revaluation Analysis: Filters strictly to students having Reval 1 and/or
     Reval 2 rows (retaining their Examiner and Moderator baseline rows), applies
     the complete Moderator rule set dynamically to Reval 1 (vs Moderator, or vs
     Examiner if Moderator is absent) and Reval 2 (vs Reval 1), fills Reval 1 rows
     in soft pastel purple and Reval 2 rows in soft pastel orange, applies
     row-specific SAP ID color coding, suppresses 100% Moderation alerts, and
     outputs only the Formatted File and Consolidated Multi-Sheet Excel.
2. Open password-protected or unprotected Excel mark sheets exported from
   the examination system (supports multiple uploads simultaneously).
3. Optionally upload a VGMR Excel file to match records by PRN/Student Number
   and Subject Code/Modulecode to fetch 'Internal Marks' (Internal Total Obtained),
   'Total Max Marks' (from Composite AGR type), 'Semester Total Max Marks' (from
   Semester Total AGR type), 'Semester Marks' (Grades value from "Semester Total
   Marks" AGR type rows - shown only on Moderator rows, or on an Examiner row
   with no corresponding Moderator record), calculate 'Total Marks Obtained',
   and derive 'Grade' dynamically based on percentage criteria.
4. Clean the raw export (remove blank/footer rows, drop unused columns).
5. Sort records so each student's rows appear in a fixed order:
   Examiner -> Moderator -> Reval 1 -> Reval 2.
6. Calculate 'Change %':
   - Moderation Analysis: Moderator vs Examiner.
   - Revaluation Analysis: Moderator vs Examiner, Reval 1 vs Moderator (or Examiner
     if no Moderator row), Reval 2 vs Reval 1 (fallback: Moderator -> Examiner).
7. Check paper eligibility for 100% moderation (Moderation Analysis only) if >10%
   mark change occurs in >50% of moderated answer books.
8. Export processed output Excel files with full conditional formatting, thick
   group borders, sub-question best-k excluded marks, and validation comments.
==================================================
"""

from datetime import datetime
import io
import os
import re
import tempfile
import zipfile

import msoffcrypto
from openpyxl.comments import Comment
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd
import streamlit as st


# ==================================================
# CONFIGURATION
# ==================================================

# Candidate passwords tried (in order) against any password-protected upload.
# decrypt_and_load() stops at the first one that successfully opens the file.
PASSWORDS = [
    "Exam@108",
    "UPG@123",
    "Exam@105"
]

# Canonical set of "known" (non-question) columns that always appear in a
# processed sheet. Anything NOT in this list is treated as a question-mark
# column (NOTE: TotalObtainedScore is intentionally omitted here so that
# _apply_change_highlighting_and_group_borders highlights it alongside
# question columns; pure question columns exclude it explicitly).
FIXED_COLUMNS = [
    "SrNo",
    "ExamAssingment_ID",
    "Subject Name-Subject code",
    "OnScreenID",
    "SubjectName",
    "Semester",
    "RollNo",
    "PRNNumber",
    "Name",
    "CampusName",
    "UserType",
    "MarkAttendance",
    "Email",
    "Semester Total Max Marks",
    "Semester Marks",
    "Internal Marks",
    "Total Max Marks",
    "Total Marks Obtained",
    "Grade",
    "Change %",
    "Excluded Extra Marks",
]

# Columns to drop from the original standard clean downloaded Excel files.
# This is the "legacy" export: no GMR-derived columns, no Change %.
COLUMNS_TO_DROP_STANDARD = [
    "Semester",
    "ExamAssingment_ID",
    "OnScreenID",
    "RollNo",
    "Name",
    "UserType",
    "CampusName",
    "MarkAttendance",
    "Email",
    "Grade",
    "Internal Marks",
    "Total Max Marks",
    "Semester Total Max Marks",
    "Semester Marks",
    "Total Marks Obtained",
    "Change %",
]

# Columns to drop from the formatted downloaded Excel files.
# This export keeps UserType, GMR columns, Grade, Change %, etc. - it only
# drops the columns nobody needs to see (identifiers/admin metadata).
# NOTE: "ExamAssingment_ID" is intentionally NOT dropped here - the Formatted File and the
# Consolidated Multi-Sheet Excel show it right next to SAP ID (see process_data_for_download).
COLUMNS_TO_DROP_FORMATTED = [
    "Semester",
    "OnScreenID",
    "RollNo",
    "Name",
    "CampusName",
    "MarkAttendance",
]

# Columns hidden from the on-screen Streamlit preview tables (kept in the
# underlying DataFrame / exported Excel files - this only affects display).
PREVIEW_HIDDEN_COLUMNS = [
    "OnScreenID",
    "RollNo",
    "Name",
    "CampusName",
    "Subject Name-Subject code",
    "SubjectName",
    "Semester",
]

# Defines the fixed row order within each student's record block.
# Lower number = appears first. Anything not listed here (shouldn't happen)
# falls back to 99 via .fillna(99) in clean_and_sort().
USER_TYPE_ORDER = {
    "Examiner": 1,
    "Moderator": 2,
    "Reval 1": 3,
    "Reval 2": 4,
}

# Row background colours (CSS strings) used only in the Streamlit preview
# tables (see highlight_rows/style_preview) to visually distinguish
# Moderator/Reval rows from the Examiner row above them.
ROW_COLOURS = {
    "Moderator": "background-color:#143d22;color:#ffffff;",  # Explicit white text for high contrast in light & dark mode
    "Reval 1": "background-color:#162456;color:white;",
    "Reval 2": "background-color:#290245;color:white;",
}

INCREASE_STYLE = "color:#00E676;font-weight:bold;"  # Bright green: mark went up / grade improved
DECREASE_STYLE = "color:#FF4500;font-weight:bold;"  # Red: mark went down / grade degraded

# Ordinal rank of each grade letter, used purely for comparison (e.g. did a
# reviewer's Grade rank higher or lower than the baseline's Grade).
GRADE_RANKS = {
    "O": 8,
    "A+": 7,
    "A": 6,
    "B+": 5,
    "B": 4,
    "C": 3,
    "P": 2,
    "F": 1
}


# --- Extra Excel formatting settings (Formatted File + Consolidated Excel only) ---
QUESTION_COL_WIDTH = 10        # Fixed width for question-wise mark columns (~75 px).
EXCEL_ROW_HEIGHT = 21          # Fixed height in points for header + data rows (~28 px).


HEADER_ROW_HEIGHT = 45         # Header row height in points (~60 px).
TOTAL_SCORE_LABEL = "Total Obtained Score"   # Display label for 'TotalObtainedScore' in the two formatted exports.
ASSIGNMENT_ID_LABEL = "Assignment ID"         # Display label for 'ExamAssingment_ID' in the two formatted exports.

# Fixed widths (keyed by the column name as it appears in the formatted export).
FIXED_COLUMN_WIDTHS = {
    "SrNo": 5,
    "Subject code": 11.57,
    "SubjectName": 30,
    "SAP ID": 12,
    "UserType": 10,
    "ExamAssingment_ID": 13,     # Shown as "Assignment ID" (renamed last, see _apply_enhanced_formatting).
    "TotalObtainedScore": 9,
}

# Row fills by UserType (Excel hex codes).
MODERATOR_ROW_FILL = "E2EFDA"   # Soft pastel green (Moderation + Revaluation modes)
REVAL1_ROW_FILL = "E8DCF4"      # Soft pastel purple (Revaluation Analysis mode)
REVAL2_ROW_FILL = "FDE4CF"      # Soft pastel orange (Revaluation Analysis mode)

# Tan, Background 2, Darker 50% - fills rows of students that only have an Examiner row.
EXAMINER_ONLY_FILL = "948A54"

# Red fill on the SAP ID cell of the baseline + reviewer rows when the reviewer value is "NA"
# or 0 for any question the baseline awarded non-zero marks for (overrides the score-change colour).
NA_MISMATCH_SAP_FILL = "FF0000"

# SAP ID colour-coding by the change in Total Obtained Score (Baseline -> Reviewer).
# Magnitude tiers use the absolute Change % of the reviewer row:
#   <= first limit -> light, <= second limit -> medium, above -> dark.
CHANGE_TIER_LIMITS = (5.0, 10.0)
SAP_INCREASE_FILLS = ["C3E6C3", "8FD18F", "5DBB63"]   # pastel green: light / medium / darker
SAP_DECREASE_FILLS = ["F8CBCB", "F19999", "E86A6A"]   # pastel red:   light / medium / darker
SAP_NO_CHANGE_FILL = "BDD7EE"                         # pastel blue


# ==================================================
# SHARED COLUMN & BASELINE HELPERS
# ==================================================
def _get_pure_question_columns(df_or_cols):
    """Returns only actual question-mark columns, excluding TotalObtainedScore and metadata."""
    cols = df_or_cols.columns.tolist() if hasattr(df_or_cols, "columns") else list(df_or_cols)
    non_question = (
        set(FIXED_COLUMNS)
        | set(COLUMNS_TO_DROP_FORMATTED)
        | {"SAP ID", "Subject code", "TotalObtainedScore"}
    )
    return [c for c in cols if c not in non_question]


def _get_baseline_row(group, user_type, reval_mode=False):
    """
    Resolves the baseline row (index, Series) against which a reviewer row is compared.

    - In Moderation Analysis mode (reval_mode=False):
      Always compares against Examiner (frozen legacy behavior).
    - In Revaluation Analysis mode (reval_mode=True):
      * Moderator -> Examiner
      * Reval 1   -> Moderator if present, else Examiner
      * Reval 2   -> Reval 1 if present, else Moderator if present, else Examiner
    """
    u_type = str(user_type).strip()
    if u_type == "Examiner":
        return None, None

    ut_series = group["UserType"].astype(str).str.strip()

    if not reval_mode:
        candidates = ["Examiner"]
    elif u_type == "Moderator":
        candidates = ["Examiner"]
    elif u_type == "Reval 1":
        candidates = ["Moderator", "Examiner"]
    elif u_type == "Reval 2":
        candidates = ["Reval 1", "Moderator", "Examiner"]
    else:
        candidates = ["Examiner"]

    for cand in candidates:
        matched = group[ut_series == cand]
        if not matched.empty:
            return matched.index[0], matched.iloc[0]

    return None, None


def filter_reval_students(df):
    """
    Returns a DataFrame containing ONLY rows belonging to students (PRNNumber)
    who have at least one 'Reval 1' and/or 'Reval 2' row, preserving all of their
    rows (Examiner, Moderator if present, Reval 1, Reval 2) in sorted order.
    """
    if df.empty or "PRNNumber" not in df.columns or "UserType" not in df.columns:
        return df.iloc[0:0].copy()

    reval_mask = df["UserType"].astype(str).str.strip().str.lower().isin(["reval 1", "reval 2"])
    reval_prns = set(df.loc[reval_mask, "PRNNumber"].unique())
    if not reval_prns:
        return df.iloc[0:0].copy()

    return df[df["PRNNumber"].isin(reval_prns)].reset_index(drop=True)


# ==================================================
# SUB-QUESTION GROUPING, BEST-K INFERENCE & EXTRA MARKS
# ==================================================
def parse_question_column_info(col_name):
    """
    Parses a question column header across variations such as:
      - "1) Q1.a (5.00)", "Q1 (10)", "Q.1A (5)", "Q.1.A", "Q.1a", "Q.1.a", "Q.1.I", "Q.1.i"
      - "2) Q2A. (8.00)" (trailing dot tolerated)
      - "Q1A", "Q1B", "Q1C", "Q1.1", "Q1.2", "Q1.3", "Q1 1", "Q1 2", "Q1 3"
      - "Q1 A", "Q1 B", "Q1 C", "Q1a", "Q1b", "Q1c", "Q1 a", "Q1 b", "Q1 c"
      - "Q1 i", "Q1 ii", "Q1 iii", "Q1i", "Q1ii", "Q1iii"
      - Compound sub-parts like "1) Q1A.i (10.00)", "2) Q1B.i (10.00)", "7) Q4A (5.00)"
    Returns dict with col, main_q, sub_part, max_marks, or None if not a question column.
    """
    s = str(col_name).strip()
    m_max = re.search(r"[\(\[]\s*([\d.]+)\s*[\)\]]\s*$", s)
    max_marks = float(m_max.group(1)) if m_max else None

    s_body = re.sub(r"\s*[\(\[]\s*[\d.]+\s*[\)\]]\s*$", "", s).strip()
    s_body = re.sub(r"^\d+\s*[\)\.]\s*", "", s_body).strip()
    # Allow optional trailing dot(s) after question/sub-part label (e.g., "2) Q2A. (8.00)").
    s_body = re.sub(r"\.+\s*$", "", s_body).strip()

    m_q = re.match(
        r"^Q\s*\.?\s*(\d+)(?:(?:\s*\.\s*|\s*|-)?([A-Za-z]+|[IVXLCDMivxlcdm]+|\d+))?(?:(?:\s*\.\s*|\s*|-)?([A-Za-z]+|[IVXLCDMivxlcdm]+|\d+))?$",
        s_body,
    )
    if not m_q:
        return None

    part1 = m_q.group(2)
    part2 = m_q.group(3)
    if part1 and part2:
        sub_part = f"{part1}.{part2}"
    else:
        sub_part = part1 or part2

    return {
        "col": col_name,
        "main_q": f"Q{int(m_q.group(1))}",
        "sub_part": sub_part,
        "max_marks": max_marks,
    }


def _parse_numeric_mark(val):
    """Returns float mark if cell is non-blank, non-NA, and numeric; else None."""
    if val is None:
        return None
    s = str(val).strip()
    if s == "" or s.upper() in ("NA", "NAN", "NONE", "ABSENT", "UFM", "AB"):
        return None
    cleaned = re.sub(r"^\+\s*", "", s)
    try:
        return float(cleaned)
    except (ValueError, TypeError):
        return None


def infer_subquestion_best_k_map(df, question_columns):
    """
    Automatically infers how many sub-questions count (best_k) for each main
    question group (e.g. Q1, Q2, Q3) from the question headers and, when
    needed, by checking candidate k values against TotalObtainedScore.
    """
    groups = {}
    for col in question_columns:
        info = parse_question_column_info(col)
        if info is not None:
            groups.setdefault(info["main_q"], []).append(info)

    best_k_map = {}
    ambiguous_groups = {}

    for main_q, items in groups.items():
        sub_items = [x for x in items if x["sub_part"]]
        if not sub_items:
            best_k_map[main_q] = len(items)
            continue

        n_subs = len(sub_items)
        sub_maxes = [x["max_marks"] for x in sub_items if x["max_marks"] is not None and x["max_marks"] > 0]
        main_items = [x for x in items if not x["sub_part"]]
        main_maxes = [x["max_marks"] for x in main_items if x["max_marks"] is not None and x["max_marks"] > 0]

        if main_maxes and sub_maxes:
            k_inf = int(round(main_maxes[0] / sub_maxes[0]))
            best_k_map[main_q] = max(1, min(k_inf, n_subs))
            continue

        if n_subs >= 2 and len(sub_maxes) == n_subs and len(set(sub_maxes)) == 1:
            ambiguous_groups[main_q] = [x["col"] for x in sub_items]
            best_k_map[main_q] = max(1, n_subs - 1)
            continue

        best_k_map[main_q] = n_subs

    if ambiguous_groups and "TotalObtainedScore" in df.columns:
        max_n_subs = max(len(c) for c in ambiguous_groups.values())
        best_uniform_k = None
        best_uniform_matches = -1

        for u_k in range(1, max_n_subs + 1):
            matches = 0
            evaluated = 0
            for _, row in df.iterrows():
                tot = _parse_numeric_mark(row.get("TotalObtainedScore"))
                if tot is None:
                    continue
                evaluated += 1
                subtotal = 0.0
                for mq, items in groups.items():
                    q_cols = [x["col"] for x in items]
                    vals = sorted(
                        [v for v in (_parse_numeric_mark(row.get(c)) for c in q_cols) if v is not None],
                        reverse=True,
                    )
                    k_use = min(u_k, len(vals)) if mq in ambiguous_groups else best_k_map[mq]
                    subtotal += sum(vals[:k_use])
                if abs(subtotal - tot) <= 0.5001:
                    matches += 1
            if evaluated > 0 and matches > best_uniform_matches:
                best_uniform_matches = matches
                best_uniform_k = u_k

        if best_uniform_k is not None and best_uniform_matches > 0:
            for mq, q_cols in ambiguous_groups.items():
                best_k_map[mq] = min(best_uniform_k, len(q_cols))

        for mq, q_cols in ambiguous_groups.items():
            n_subs = len(q_cols)
            best_matches = -1
            chosen_k = best_k_map[mq]

            for cand_k in range(1, n_subs + 1):
                matches = 0
                evaluated = 0
                for _, row in df.iterrows():
                    tot = _parse_numeric_mark(row.get("TotalObtainedScore"))
                    if tot is None:
                        continue
                    evaluated += 1
                    subtotal = 0.0
                    for omq, oitems in groups.items():
                        if omq == mq:
                            continue
                        ocols = [x["col"] for x in oitems]
                        ovals = sorted(
                            [v for v in (_parse_numeric_mark(row.get(c)) for c in ocols) if v is not None],
                            reverse=True,
                        )
                        subtotal += sum(ovals[: best_k_map[omq]])
                    cvals = sorted(
                        [v for v in (_parse_numeric_mark(row.get(c)) for c in q_cols) if v is not None],
                        reverse=True,
                    )
                    subtotal += sum(cvals[:cand_k])
                    if abs(subtotal - tot) <= 0.5001:
                        matches += 1
                if evaluated > 0 and matches > best_matches:
                    best_matches = matches
                    chosen_k = cand_k

            best_k_map[mq] = chosen_k

    return best_k_map


def _fmt_num_clean(num):
    """Formats numeric mark cleanly without trailing .0 when whole."""
    r = round(float(num), 2)
    return str(int(r)) if float(r).is_integer() else f"{r:g}"


def predict_semester_total_max_marks(df, question_columns=None, best_k_map=None):
    """
    Predicts subject-wide Semester Total Max Marks from non-excluded question
    bracket max values (e.g. (5.00)) when VGMR is not provided.
    Returns float total or None if no bracket max values are found.
    """
    q_cols = (
        question_columns
        if question_columns is not None
        else _get_pure_question_columns(df)
    )
    if not q_cols:
        return None
    if best_k_map is None:
        best_k_map = infer_subquestion_best_k_map(df, q_cols)

    groups = {}
    ungrouped = []
    for col in q_cols:
        info = parse_question_column_info(col)
        if info is not None:
            groups.setdefault(info["main_q"], []).append(info)
        else:
            ungrouped.append(col)

    total = 0.0
    found_any = False

    for col in ungrouped:
        m = re.search(r"[\(\[]\s*([\d.]+)\s*[\)\]]\s*$", str(col))
        if m:
            total += float(m.group(1))
            found_any = True

    for main_q, items in groups.items():
        sub_items = [x for x in items if x["sub_part"] and x["max_marks"] is not None and x["max_marks"] > 0]
        main_items = [x for x in items if not x["sub_part"] and x["max_marks"] is not None and x["max_marks"] > 0]

        if sub_items:
            k = min(best_k_map.get(main_q, len(sub_items)), len(sub_items))
            top_sub_maxes = sorted([x["max_marks"] for x in sub_items], reverse=True)[:k]
            total += sum(top_sub_maxes)
            found_any = True
        elif main_items:
            total += sum(x["max_marks"] for x in main_items)
            found_any = True

    return round(total, 2) if found_any else None


def apply_predicted_semester_total_max_marks(df, question_columns=None, best_k_map=None):
    """
    When VGMR is not provided (or Semester Total Max Marks is blank on rows),
    fills a single subject-wide predicted value formatted with '!' on both
    sides (e.g., '!30!') across all rows in Semester Total Max Marks.
    """
    df = df.copy()
    if "Semester Total Max Marks" not in df.columns:
        return df

    q_cols = (
        question_columns
        if question_columns is not None
        else _get_pure_question_columns(df)
    )
    predicted = predict_semester_total_max_marks(df, q_cols, best_k_map=best_k_map)
    if predicted is None:
        return df

    pred_str = f"!{_fmt_num_clean(predicted)}!"
    existing_non_empty = df["Semester Total Max Marks"].astype(str).str.strip().replace("", pd.NA).dropna()
    if existing_non_empty.empty:
        df["Semester Total Max Marks"] = pred_str
    else:
        mask_blank = df["Semester Total Max Marks"].astype(str).str.strip() == ""
        df.loc[mask_blank, "Semester Total Max Marks"] = pred_str

    return df


def compute_row_excluded_extra_marks(row, df, question_columns, best_k_map=None):
    """
    Computes total excluded extra marks and the non-excluded (counted) sub-question
    sum for a single row across all main question groups.
    If non-excluded total is less by 0.5 than TotalObtainedScore (rounded up to
    nearest greater number), shows "(0.5)" in Excluded Extra Marks and treats it
    as valid rounding rather than an error.
    Returns (extra_marks_or_blank, excluded_col_names_set, counted_sum_or_none, is_half_round_up).
    """
    tot = _parse_numeric_mark(row.get("TotalObtainedScore"))
    if tot is None:
        return "", set(), None, False

    if best_k_map is None:
        best_k_map = infer_subquestion_best_k_map(df, question_columns)

    groups = {}
    ungrouped_cols = []
    for col in question_columns:
        info = parse_question_column_info(col)
        if info is not None:
            groups.setdefault(info["main_q"], []).append(col)
        else:
            ungrouped_cols.append(col)

    total_extra = 0.0
    counted_sum = 0.0
    excluded_cols = set()

    for col in ungrouped_cols:
        m_val = _parse_numeric_mark(row.get(col))
        if m_val is not None:
            counted_sum += m_val

    for main_q, q_cols in groups.items():
        k = best_k_map.get(main_q, len(q_cols))
        scored = []
        for col in q_cols:
            m_val = _parse_numeric_mark(row.get(col))
            if m_val is not None:
                scored.append((m_val, col))

        if len(scored) <= k:
            counted_sum += sum(v for v, _ in scored)
            continue

        scored_sorted = sorted(scored, key=lambda x: x[0], reverse=True)
        counted_sum += sum(val for val, _ in scored_sorted[:k])
        for _, ex_col in scored_sorted[k:]:
            excluded_cols.add(ex_col)
        total_extra += sum(val for val, _ in scored_sorted[k:])

    # Check fractional rounding between counted_sum and TotalObtainedScore:
    # - Rounded up (e.g., 18.75 -> 19 diff +0.25, or 18.5 -> 19 diff +0.5): valid rounding, not an error.
    # - Rounded down (e.g., 18.25 -> 18 diff -0.25): valid rounding (<0.5 dropped), not an error.
    diff_tot_counted = round(tot - counted_sum, 4)
    is_round_up = (diff_tot_counted > 0) and (diff_tot_counted <= 0.5 + 1e-6) and (abs((diff_tot_counted * 4) - round(diff_tot_counted * 4)) <= 1e-4)
    is_round_down = (diff_tot_counted < 0) and (abs(diff_tot_counted) < 0.5 - 1e-6) and (abs((abs(diff_tot_counted) * 4) - round(abs(diff_tot_counted) * 4)) <= 1e-4)
    is_valid_rounding = is_round_up or is_round_down

    if total_extra <= 1e-9:
        if is_round_up:
            return f"({_fmt_num_clean(diff_tot_counted)})", excluded_cols, counted_sum, True
        if is_round_down:
            return f"({_fmt_num_clean(abs(diff_tot_counted))})", excluded_cols, counted_sum, True
        return "", excluded_cols, counted_sum, False

    if abs(total_extra - 0.5) <= 1e-6:
        return "", set(), counted_sum, is_valid_rounding

    rounded = round(total_extra, 2)
    val_out = int(rounded) if float(rounded).is_integer() else rounded
    return val_out, excluded_cols, counted_sum, is_valid_rounding


def add_excluded_extra_marks_column(df, question_columns=None, best_k_map=None):
    """Adds 'Excluded Extra Marks' column right after 'Change %' (or 'TotalObtainedScore')."""
    df = df.copy()
    q_cols = (
        question_columns
        if question_columns is not None
        else _get_pure_question_columns(df)
    )
    if best_k_map is None:
        best_k_map = infer_subquestion_best_k_map(df, q_cols)

    extra_values = []
    for _, row in df.iterrows():
        val, _, _, _ = compute_row_excluded_extra_marks(row, df, q_cols, best_k_map=best_k_map)
        extra_values.append(val)

    df["Excluded Extra Marks"] = extra_values

    cols = df.columns.tolist()
    cols.remove("Excluded Extra Marks")
    if "Change %" in cols:
        cols.insert(cols.index("Change %") + 1, "Excluded Extra Marks")
    elif "TotalObtainedScore" in cols:
        cols.insert(cols.index("TotalObtainedScore") + 1, "Excluded Extra Marks")
    else:
        cols.append("Excluded Extra Marks")
    df = df[cols]

    return df


# ==================================================
# STEP 1: DECRYPT + LOAD UPLOADED FILES
# ==================================================
def decrypt_and_load(uploaded_file):
    """
    Attempts to load unprotected Excel files directly first.
    If protected, iterates through PASSWORDS list to decrypt.

    Returns:
        (DataFrame, label) on success, where `label` is either
        "Unprotected (Direct)" or the password that worked.
        (None, None) if the file could not be opened by any means.
    """
    # --- Attempt 1: file has no password at all ------------------------
    try:
        uploaded_file.seek(0)  # Streamlit's UploadedFile is a stream; rewind before every read attempt.
        df = pd.read_excel(uploaded_file, dtype=object, keep_default_na=False)
        return df, "Unprotected (Direct)"
    except Exception:
        # Not readable as a plain workbook - most likely it's encrypted. Fall through to password attempts.
        pass

    # --- Attempt 2: try each known password in turn ---------------------
    for password in PASSWORDS:
        decrypted_path = None
        try:
            uploaded_file.seek(0)

            office_file = msoffcrypto.OfficeFile(uploaded_file)
            office_file.load_key(password=password)  # Raises if the password is wrong.

            # msoffcrypto needs to write the decrypted bytes somewhere before pandas can read them,
            # so we decrypt into a throwaway temp file and clean it up in `finally`.
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                office_file.decrypt(tmp)
                decrypted_path = tmp.name

            df = pd.read_excel(decrypted_path, dtype=object, keep_default_na=False)
            return df, password

        except Exception:
            # Wrong password (or some other decrypt/read failure) - try the next candidate.
            continue

        finally:
            # Always scrub the temp file, whether decryption succeeded or failed.
            if decrypted_path and os.path.exists(decrypted_path):
                os.remove(decrypted_path)

    # Exhausted every option - unprotected read failed and no password matched.
    return None, None


def load_gmr_file(gmr_file):
    """Loads GMR file directly or decrypts if password protected."""
    if gmr_file is None:
        return None
    df, _ = decrypt_and_load(gmr_file)  # Reuse the same open/decrypt logic; we don't care which password worked.
    return df


# ==================================================
# STEP 2: CLEAN, SORT, & GMR LOOKUP MATCHING
# ==================================================
def clean_and_sort(df):
    """Clean up raw Questionwise export and sort rows."""
    df = df.copy()
    df.columns = df.columns.astype(str).str.strip()  # Guard against stray whitespace in exported headers.

    # Drop rows with a blank PRNNumber - these are usually spacer/footer rows in the raw export.
    if "PRNNumber" in df.columns:
        df = df[df["PRNNumber"].astype(str).str.strip() != ""]

    df = df.dropna(how="all")  # Drop fully-empty rows (e.g. trailing blank lines in the sheet).

    # The export appends a "Count: N" footer row - strip it out using whichever
    # column is available to check ("SrNo" preferred, else the first column).
    if "SrNo" in df.columns:
        df = df[~df["SrNo"].astype(str).str.contains("Count:", case=False, na=False)]
    else:
        first_col = df.columns[0]
        df = df[~df[first_col].astype(str).str.contains("Count:", case=False, na=False)]

    # PRNNumber should be numeric - coerce and drop anything that isn't
    # (defensive cleanup in case a stray text row slipped through above).
    if "PRNNumber" in df.columns:
        df["PRNNumber"] = pd.to_numeric(df["PRNNumber"], errors="coerce")
        df = df[df["PRNNumber"].notna()]

    # Sort so each student's rows appear in a fixed order: Examiner -> Moderator -> Reval 1 -> Reval 2.
    if "PRNNumber" in df.columns and "UserType" in df.columns:
        df["SortOrder"] = df["UserType"].astype(str).str.strip().map(USER_TYPE_ORDER).fillna(99)
        df = df.sort_values(by=["PRNNumber", "SortOrder"]).drop(columns=["SortOrder"])

    df = df.reset_index(drop=True)
    return df


def calculate_grade(total_obtained, total_max):
    """Calculates Grade based on percentage of Total Obtained vs Total Max Marks."""
    try:
        obtained = float(total_obtained)
        max_m = float(total_max)

        if max_m <= 0:
            return ""  # Avoid division by zero / nonsensical percentages.

        pct = (obtained / max_m) * 100.0

        # Standard percentage-to-grade ladder, checked from highest to lowest.
        if pct >= 90:
            return "O"
        elif pct >= 80:
            return "A+"
        elif pct >= 70:
            return "A"
        elif pct >= 60:
            return "B+"
        elif pct >= 55:
            return "B"
        elif pct >= 50:
            return "C"
        elif pct >= 40:
            return "P"
        else:
            return "F"
    except (ValueError, TypeError):
        # total_obtained/total_max were blank, "AB", "NA", etc. - no grade can be derived.
        return ""


def apply_gmr_lookup(df, gmr_df):
    """
    Performs lookup matching between Questionwise and GMR data:
    - Internal Marks: from rows where AGR type contains "Internal Total"
    - Total Max Marks: MAX Marks from rows where AGR type contains "Composite "
    - Semester Total Max Marks: MAX Marks from rows where AGR type contains "Semester Total"
    - Semester Marks: Grades from "Semester Total Marks" rows (kept on Moderator rows, or
      Examiner rows with no corresponding Moderator record - unchanged across modes)
    - Total Marks Obtained: TotalObtainedScore + Internal Marks
    - Grade: Calculated based on percentage criteria
    """
    df = df.copy()

    # Always create these columns (even with no GMR file) so downstream code
    # (reorder_gmr_columns, exports, previews) can assume they exist.
    df["Internal Marks"] = ""
    df["Total Max Marks"] = ""
    df["Semester Total Max Marks"] = ""
    df["Semester Marks"] = ""
    df["Total Marks Obtained"] = ""
    df["Grade"] = ""

    if gmr_df is None or gmr_df.empty:
        return df  # No GMR file uploaded - leave the GMR columns blank.

    gmr_clean = gmr_df.copy()
    gmr_clean.columns = gmr_clean.columns.astype(str).str.strip()

    # Bail out quietly if the GMR file doesn't have the columns we need to match on.
    required_gmr_cols = ["Student Number", "Modulecode", "AGR type", "Grades"]
    if not all(col in gmr_clean.columns for col in required_gmr_cols):
        return df

    # --- Build a composite lookup key (PRN + Subject code) on both sides ---
    df_key = (
        df["PRNNumber"].astype(str).str.strip() +
        df["Subject Name-Subject code"].astype(str).str.strip()
    )

    gmr_clean["LookupKey"] = (
        gmr_clean["Student Number"].astype(str).str.strip() +
        gmr_clean["Modulecode"].astype(str).str.strip()
    )

    # 1. Fetch Internal Marks (GMR rows whose AGR type mentions "Internal Total").
    internal_df = gmr_clean[
        gmr_clean["AGR type"].astype(str).str.contains("Internal Total", case=False, na=False)
    ]
    internal_obtained_map = dict(zip(internal_df["LookupKey"], internal_df["Grades"]))
    df["Internal Marks"] = df_key.map(internal_obtained_map).fillna("")

    # 2. Fetch Total Max Marks from Composite AGR type.
    composite_max_map = {}
    if "MAX Marks" in gmr_clean.columns:
        composite_df = gmr_clean[
            gmr_clean["AGR type"].astype(str).str.contains("Composite ", case=False, na=False)
        ]
        composite_max_map = dict(zip(composite_df["LookupKey"], composite_df["MAX Marks"]))
    df["Total Max Marks"] = df_key.map(composite_max_map).fillna("")

    # 3. Fetch Semester Total Max Marks from Semester Total AGR type.
    sem_total_max_map = {}
    if "MAX Marks" in gmr_clean.columns:
        sem_total_df = gmr_clean[
            gmr_clean["AGR type"].astype(str).str.contains("Semester Total", case=False, na=False)
        ]
        sem_total_max_map = dict(zip(sem_total_df["LookupKey"], sem_total_df["MAX Marks"]))
    df["Semester Total Max Marks"] = df_key.map(sem_total_max_map).fillna("")

    # 3b. Fetch Semester Marks (Grades column) from GMR rows whose AGR type
    # contains "Semester Total Marks" (e.g. "Semester Total Marks(60 )").
    # Kept on Moderator rows, or Examiner rows with no corresponding Moderator record.
    sem_marks_df = gmr_clean[
        gmr_clean["AGR type"].astype(str).str.contains("Semester Total Marks", case=False, na=False)
    ]
    sem_marks_map = dict(zip(sem_marks_df["LookupKey"], sem_marks_df["Grades"]))
    raw_semester_marks = df_key.map(sem_marks_map).fillna("")

    if "UserType" in df.columns and "PRNNumber" in df.columns and "Subject Name-Subject code" in df.columns:
        module_key = (
            df["PRNNumber"].astype(str).str.strip() +
            "||" +
            df["Subject Name-Subject code"].astype(str).str.strip()
        )
        moderator_module_keys = set(
            module_key[df["UserType"].astype(str).str.strip() == "Moderator"]
        )

        def keep_semester_marks(idx):
            u_type = str(df.loc[idx, "UserType"]).strip()
            if u_type == "Moderator":
                return raw_semester_marks.loc[idx]
            if u_type == "Examiner" and module_key.loc[idx] not in moderator_module_keys:
                return raw_semester_marks.loc[idx]
            return ""

        df["Semester Marks"] = [keep_semester_marks(i) for i in df.index]

    # 4. Calculate Total Marks Obtained = TotalObtainedScore (exam script) + Internal Marks (GMR).
    def compute_total_obtained(row):
        tot_score = row.get("TotalObtainedScore", "")
        int_marks = row.get("Internal Marks", "")

        try:
            val_tot = float(str(tot_score).strip())
        except (ValueError, TypeError):
            val_tot = 0.0

        try:
            cleaned_int_marks = re.sub(r"^\+\s*", "", str(int_marks).strip())
            val_int = float(cleaned_int_marks)
        except (ValueError, TypeError):
            val_int = 0.0

        if str(tot_score).strip() == "" and str(int_marks).strip() == "":
            return ""

        return val_tot + val_int

    df["Total Marks Obtained"] = df.apply(compute_total_obtained, axis=1)

    # 5. Calculate Grade from the newly computed Total Marks Obtained / Total Max Marks.
    df["Grade"] = df.apply(
        lambda r: calculate_grade(r["Total Marks Obtained"], r["Total Max Marks"]), axis=1
    )

    return df


def calculate_paper_max_marks(df):
    """Extracts sum of question max marks from question column headers."""
    question_columns = _get_pure_question_columns(df)
    total_q_max = 0.0
    for col in question_columns:
        match = re.search(r"\(([\d.]+)\)", col)
        if match:
            try:
                total_q_max += float(match.group(1))
            except ValueError:
                pass
    return total_q_max


def calculate_moderator_change_pct(df, reval_mode=False):
    """
    Calculates Change % using Semester Total Max Marks (with fallback chain):
    - Moderation Analysis (reval_mode=False):
      Only Moderator rows, comparing Moderator against Examiner (frozen behavior).
    - Revaluation Analysis (reval_mode=True):
      Moderator (vs Examiner), Reval 1 (vs Moderator if present else Examiner),
      and Reval 2 (vs Reval 1, fallback Moderator -> Examiner).
    """
    df = df.copy()
    df["Change %"] = pd.Series([""] * len(df), index=df.index, dtype="object")

    if "PRNNumber" not in df.columns or "UserType" not in df.columns or "TotalObtainedScore" not in df.columns:
        return df

    paper_q_max = calculate_paper_max_marks(df)
    target_types = ("Moderator", "Reval 1", "Reval 2") if reval_mode else ("Moderator",)

    for _, group in df.groupby("PRNNumber", dropna=False):
        ut_series = group["UserType"].astype(str).str.strip()

        for u_type in target_types:
            review_rows = group[ut_series == u_type]
            if review_rows.empty:
                continue

            _, base_row = _get_baseline_row(group, u_type, reval_mode=reval_mode)
            if base_row is None:
                continue

            review_idx = review_rows.index[0]
            review_row = review_rows.iloc[0]

            try:
                base_score = float(base_row["TotalObtainedScore"])
                rev_score = float(review_row["TotalObtainedScore"])

                # Denominator selection, in priority order:
                # 1. Semester Total Max Marks (VGMR or predicted !value!).
                sem_max_marks = 0.0
                sem_max_val = str(review_row.get("Semester Total Max Marks", "")).strip().strip("!")
                if sem_max_val != "":
                    try:
                        sem_max_marks = float(sem_max_val)
                    except ValueError:
                        sem_max_marks = 0.0

                # 2. Fallback to Total Max Marks.
                if sem_max_marks <= 0:
                    tot_max_val = str(review_row.get("Total Max Marks", "")).strip()
                    if tot_max_val != "":
                        try:
                            sem_max_marks = float(tot_max_val)
                        except ValueError:
                            sem_max_marks = 0.0

                # 3. Fallback to sum of question max marks from column headers.
                if sem_max_marks <= 0:
                    sem_max_marks = paper_q_max

                if sem_max_marks > 0:
                    change_pct = ((rev_score - base_score) / sem_max_marks) * 100.0
                    df.loc[review_idx, "Change %"] = round(change_pct, 2)
            except (ValueError, TypeError):
                continue

    return df


def reorder_gmr_columns(df):
    """
    Reorders GMR and calculated columns right after TotalObtainedScore in the exact sequence:
    Semester Total Max Marks, Semester Marks, Internal Marks, Total Max Marks, Total Marks Obtained, Grade, Change %, Email
    """
    df = df.copy()
    target_tail = [
        "Semester Total Max Marks",
        "Semester Marks",
        "Internal Marks",
        "Total Max Marks",
        "Total Marks Obtained",
        "Grade",
        "Change %",
        "Email",
    ]
    if "TotalObtainedScore" in df.columns:
        cols = df.columns.tolist()
        tot_idx = cols.index("TotalObtainedScore")
        head = cols[:tot_idx + 1]
        tail = [c for c in target_tail if c in cols]
        remaining = [c for c in cols if c not in head and c not in tail]
        df = df[head + tail + remaining]
    return df


# ==================================================
# STEP 2B: PROCESS DATA FOR OUTPUT EXCEL DOWNLOAD
# ==================================================
def process_data_for_download(df, include_all_usertypes=False):
    """Applies download-specific transformations to DataFrame."""
    download_df = df.copy()

    if not include_all_usertypes:
        # --- "Standard" export: only Examiner + (filtered) Moderator rows ---
        if "UserType" in download_df.columns:
            reval_mask = download_df["UserType"].astype(str).str.strip().str.lower().isin(["reval 1", "reval 2"])
            download_df = download_df[~reval_mask]

        if "PRNNumber" in download_df.columns and "UserType" in download_df.columns:
            moderator_prns = set(
                download_df[download_df["UserType"].astype(str).str.strip().str.lower() == "moderator"]["PRNNumber"]
            )

            def filter_rows(row):
                prn = row["PRNNumber"]
                u_type = str(row["UserType"]).strip().lower()
                if prn in moderator_prns:
                    return u_type == "moderator"
                return True

            download_df = download_df[download_df.apply(filter_rows, axis=1)]

    # Replace "AB" in TotalObtainedScore with a human-readable label based on attendance.
    if "TotalObtainedScore" in download_df.columns:
        def adjust_score(row):
            score = str(row["TotalObtainedScore"]).strip()
            attendance = str(row.get("MarkAttendance", "")).strip().upper()

            if score.upper() == "AB":
                if attendance == "UFM":
                    return "UFM"
                return "ABSENT"
            return score

        download_df["TotalObtainedScore"] = download_df.apply(adjust_score, axis=1)

    # Drop Unwanted Columns - which set depends on which export flavour this is.
    cols_to_remove = COLUMNS_TO_DROP_FORMATTED if include_all_usertypes else COLUMNS_TO_DROP_STANDARD
    download_df = download_df.drop(columns=cols_to_remove, errors="ignore")

    # Rename Columns to the labels expected in the exported file.
    if "Subject Name-Subject code" in download_df.columns:
        download_df = download_df.rename(columns={"Subject Name-Subject code": "Subject code"})

    if "PRNNumber" in download_df.columns:
        download_df = download_df.rename(columns={"PRNNumber": "SAP ID"})

    # Formatted / Consolidated exports only: place ExamAssingment_ID immediately after SAP ID
    if include_all_usertypes and "ExamAssingment_ID" in download_df.columns and "SAP ID" in download_df.columns:
        download_df["ExamAssingment_ID"] = download_df["ExamAssingment_ID"].apply(
            lambda v: int(float(str(v).strip())) if re.fullmatch(r"\s*\d+(\.0+)?\s*", str(v)) else v
        )
        ordered = [c for c in download_df.columns if c != "ExamAssingment_ID"]
        ordered.insert(ordered.index("SAP ID") + 1, "ExamAssingment_ID")
        download_df = download_df[ordered]

    if "Excluded Extra Marks" in download_df.columns:
        ordered = [c for c in download_df.columns if c != "Excluded Extra Marks"]
        if "Change %" in ordered:
            ordered.insert(ordered.index("Change %") + 1, "Excluded Extra Marks")
        elif "TotalObtainedScore" in ordered:
            ordered.insert(ordered.index("TotalObtainedScore") + 1, "Excluded Extra Marks")
        else:
            ordered.append("Excluded Extra Marks")
        download_df = download_df[ordered]

    download_df = download_df.reset_index(drop=True)

    # Make SrNo Continuous (1, 2, 3, ...) since rows may have been dropped/filtered above.
    if "SrNo" in download_df.columns:
        download_df["SrNo"] = download_df.index + 1

    return download_df


# ==================================================
# STEP 3: DISPLAY FORMATTING HELPERS
# ==================================================
def format_value(value):
    """Formats values cleanly for preview display."""
    text = str(value).strip()

    if text == "" or text.upper() == "NA" or text.upper() == "NAN":
        return "" if text == "" else "NA"

    if (text.startswith("{") and text.endswith("}")) or (text.startswith("(") and text.endswith(")")):
        return text
    cleaned_text = re.sub(r"^\+\s*", "", text)

    try:
        number = float(cleaned_text)
        if number.is_integer():
            return str(int(number))
        return f"{number:g}"
    except (TypeError, ValueError):
        return text


def highlight_rows(data, question_columns, reval_mode=False, best_k_map=None):
    """Applies Excel-aligned CSS styles for Streamlit preview table safely."""
    styles = pd.DataFrame("", index=data.index, columns=data.columns)

    if reval_mode:
        excel_row_styles = {
            "Moderator": f"background-color:#{MODERATOR_ROW_FILL};color:#000000;",
            "Reval 1": f"background-color:#{REVAL1_ROW_FILL};color:#000000;",
            "Reval 2": f"background-color:#{REVAL2_ROW_FILL};color:#000000;",
        }
    else:
        excel_row_styles = {
            "Moderator": f"background-color:#{MODERATOR_ROW_FILL};color:#000000;",
            "Reval 1": "background-color:#162456;color:#ffffff;",
            "Reval 2": "background-color:#290245;color:#ffffff;",
        }

    # Pass 1: Row background colour based on UserType + Examiner-only tan fill.
    for idx in data.index:
        if "UserType" in data.columns:
            user_type = str(data.loc[idx, "UserType"]).strip()
            if user_type in excel_row_styles:
                styles.loc[idx, :] += excel_row_styles[user_type]

    if "PRNNumber" in data.columns and "UserType" in data.columns:
        for _, group in data.groupby("PRNNumber", dropna=False):
            utypes = set(group["UserType"].astype(str).str.strip())
            if utypes == {"Examiner"}:
                for idx in group.index:
                    styles.loc[idx, :] += f"background-color:#{EXAMINER_ONLY_FILL};color:#ffffff;"

    # Pass 2: cell-level green/red highlighting against baseline row.
    if "PRNNumber" in data.columns and "UserType" in data.columns:
        for _, group in data.groupby("PRNNumber", dropna=False):
            for idx in group.index:
                u_type = str(data.loc[idx, "UserType"]).strip()
                if u_type == "Examiner":
                    continue

                _, base_row = _get_baseline_row(group, u_type, reval_mode=reval_mode)
                if base_row is None:
                    continue

                for col in question_columns:
                    if col not in data.columns:
                        continue

                    base_value = str(base_row[col]).strip()
                    current_value = str(data.loc[idx, col]).strip()

                    if base_value == "" or base_value.upper() == "NA":
                        if reval_mode and base_value.upper() == "NA" and current_value != "" and current_value.upper() != "NA":
                            styles.loc[idx, col] += INCREASE_STYLE
                        continue

                    if current_value == "" or current_value.upper() == "NA":
                        styles.loc[idx, col] += DECREASE_STYLE
                        continue

                    try:
                        base_mark = float(base_value)
                        cur_mark = float(current_value)
                    except (TypeError, ValueError):
                        continue

                    if cur_mark > base_mark:
                        styles.loc[idx, col] += INCREASE_STYLE
                    elif cur_mark < base_mark:
                        styles.loc[idx, col] += DECREASE_STYLE

                if "Grade" in data.columns:
                    base_grade = str(base_row.get("Grade", "")).strip().upper()
                    cur_grade = str(data.loc[idx, "Grade"]).strip().upper()

                    if base_grade in GRADE_RANKS and cur_grade in GRADE_RANKS:
                        if GRADE_RANKS[cur_grade] > GRADE_RANKS[base_grade]:
                            styles.loc[idx, "Grade"] += INCREASE_STYLE
                        elif GRADE_RANKS[cur_grade] < GRADE_RANKS[base_grade]:
                            styles.loc[idx, "Grade"] += DECREASE_STYLE

                if "Total Marks Obtained" in data.columns:
                    base_tot_str = str(base_row.get("Total Marks Obtained", "")).strip()
                    cur_tot_str = str(data.loc[idx, "Total Marks Obtained"]).strip()
                    if base_tot_str != "" and cur_tot_str != "":
                        try:
                            base_tot = float(base_tot_str)
                            cur_tot = float(cur_tot_str)
                            if cur_tot > base_tot:
                                styles.loc[idx, "Total Marks Obtained"] += INCREASE_STYLE
                            elif cur_tot < base_tot:
                                styles.loc[idx, "Total Marks Obtained"] += DECREASE_STYLE
                        except (ValueError, TypeError):
                            pass

                if "Change %" in data.columns:
                    chg_str = str(data.loc[idx, "Change %"]).strip()
                    if chg_str != "":
                        try:
                            chg_val = float(chg_str)
                            if chg_val > 0:
                                styles.loc[idx, "Change %"] += INCREASE_STYLE
                            elif chg_val < 0:
                                styles.loc[idx, "Change %"] += DECREASE_STYLE
                        except (ValueError, TypeError):
                            pass

    # Pass 3: Semester Marks vs TotalObtainedScore comparison.
    if "Semester Marks" in data.columns and "TotalObtainedScore" in data.columns:
        for idx in data.index:
            sem_val_str = str(data.loc[idx, "Semester Marks"]).strip()
            if sem_val_str == "" or sem_val_str.upper() == "NA":
                continue

            tot_val_str = str(data.loc[idx, "TotalObtainedScore"]).strip()
            try:
                sem_val = float(re.sub(r"^\+\s*", "", sem_val_str))
                tot_val = float(tot_val_str)
            except (ValueError, TypeError):
                continue

            if sem_val == tot_val:
                styles.loc[idx, "Semester Marks"] += "background-color:#C6EFCE;color:#000000;"
            else:
                styles.loc[idx, "Semester Marks"] += "background-color:#ff0000;color:#ffff00;"

    pure_q_cols = [c for c in question_columns if c != "TotalObtainedScore"]
    if pure_q_cols:
        resolved_best_k = best_k_map if best_k_map is not None else infer_subquestion_best_k_map(data, pure_q_cols)
        for idx, row in data.iterrows():
            _, ex_cols, counted_sum, is_half_round_up = compute_row_excluded_extra_marks(
                row, data, pure_q_cols, best_k_map=resolved_best_k
            )
            for ex_col in ex_cols:
                if ex_col in styles.columns:
                    styles.loc[idx, ex_col] += "color:#0000FF;font-weight:bold;"
            if "TotalObtainedScore" in styles.columns and counted_sum is not None and not is_half_round_up:
                tot_num = _parse_numeric_mark(row.get("TotalObtainedScore"))
                if tot_num is not None and abs(counted_sum - tot_num) > 1e-6:
                    styles.loc[idx, "TotalObtainedScore"] += "background-color:#FF0000;color:#FFFF00;font-weight:bold;"

    if "TotalObtainedScore" in data.columns:
        styles.loc[:, "TotalObtainedScore"] += "font-weight:bold;"

    return styles


def _format_preview_dataframe_for_excel_mimic(df, question_columns, best_k_map=None):
    """Formats preview strings to closely match the formatted Excel output."""
    out = df.copy()
    pure_q_cols = [c for c in question_columns if c != "TotalObtainedScore"]
    if not pure_q_cols:
        return out
    resolved_best_k = best_k_map if best_k_map is not None else infer_subquestion_best_k_map(out, pure_q_cols)
    for idx, row in out.iterrows():
        ex_val, ex_cols, _, _ = compute_row_excluded_extra_marks(
            row, out, pure_q_cols, best_k_map=resolved_best_k
        )
        for ex_col in ex_cols:
            if ex_col in out.columns:
                raw_str = str(out.loc[idx, ex_col]).strip()
                if raw_str and not (raw_str.startswith("{") and raw_str.endswith("}")):
                    out.loc[idx, ex_col] = f"{{{raw_str}}}"
        if "Excluded Extra Marks" in out.columns and ex_val != "":
            out.loc[idx, "Excluded Extra Marks"] = str(ex_val)
    return out


def style_preview(df, question_columns, hide_columns=PREVIEW_HIDDEN_COLUMNS, reval_mode=False, best_k_map=None):
    """Builds styled interactive preview table mimicking formatted Excel."""
    styled_source = _format_preview_dataframe_for_excel_mimic(df, question_columns, best_k_map=best_k_map)
    display_df = styled_source.drop(columns=hide_columns, errors="ignore").copy()
    display_df.index = range(1, len(display_df) + 1)
    return (
        display_df.style
        .format(format_value)
        .apply(
            lambda data: highlight_rows(data, question_columns, reval_mode=reval_mode, best_k_map=best_k_map),
            axis=None,
        )
    )


# ==================================================
# STEP 4: SUMMARY STATISTICS & METADATA
# ==================================================
def count_reviewers(df, user_type):
    """Number of distinct students that have a row of the given UserType."""
    if "UserType" not in df.columns or "PRNNumber" not in df.columns:
        return 0
    return (
        df[df["UserType"].astype(str).str.strip().eq(user_type)]["PRNNumber"]
        .nunique()
    )


def score_change_summary(df, review_type, reval_mode=False):
    """
    Compares review_type TotalObtainedScore against its baseline row's TotalObtainedScore:
    - Moderation Analysis (reval_mode=False): always vs Examiner.
    - Revaluation Analysis (reval_mode=True):
      * Moderator vs Examiner
      * Reval 1 vs Moderator (or Examiner if no Moderator)
      * Reval 2 vs Reval 1 (fallback Moderator -> Examiner)
    """
    increased = decreased = unchanged = 0

    if "TotalObtainedScore" not in df.columns or "UserType" not in df.columns:
        return increased, decreased, unchanged

    for _, group in df.groupby("PRNNumber", dropna=False):
        ut_series = group["UserType"].astype(str).str.strip()
        review_rows = group[ut_series == review_type]
        if review_rows.empty:
            continue

        _, base_row = _get_baseline_row(group, review_type, reval_mode=reval_mode)
        if base_row is None:
            continue

        try:
            base_score = float(base_row["TotalObtainedScore"])
            review_score = float(review_rows.iloc[0]["TotalObtainedScore"])
        except (TypeError, ValueError):
            continue

        if review_score > base_score:
            increased += 1
        elif review_score < base_score:
            decreased += 1
        else:
            unchanged += 1

    return increased, decreased, unchanged


def compute_moderation_10pct_metrics(df):
    """
    Computes:
    - Moderator Count
    - Count of Moderator records with |Change %| > 10%
    - Change % (>10%) ratio & percentage
    - Eligibility for 100% Moderation (>50% threshold)
    """
    if "Change %" not in df.columns or "UserType" not in df.columns:
        return 0, 0, 0.0, False

    mod_rows = df[df["UserType"].astype(str).str.strip() == "Moderator"]
    mod_count = len(mod_rows)

    if mod_count == 0:
        return 0, 0, 0.0, False

    count_gt_10pct = 0
    for val in mod_rows["Change %"]:
        try:
            if abs(float(val)) > 10.0:
                count_gt_10pct += 1
        except (ValueError, TypeError):
            continue

    ratio_gt_10pct = (count_gt_10pct / mod_count) if mod_count > 0 else 0.0
    is_eligible_100pct_moderation = ratio_gt_10pct > 0.50

    return mod_count, count_gt_10pct, ratio_gt_10pct, is_eligible_100pct_moderation


def build_paper_info(df):
    """Builds header info line for display."""
    if df.empty:
        return ""

    semester = str(df.iloc[0].get("Semester", "")).strip()
    subject_code = str(df.iloc[0].get("Subject Name-Subject code", "")).strip()
    subject_name = str(df.iloc[0].get("SubjectName", "")).strip()

    return f"{semester}  -  {subject_code}  -  {subject_name}"


def generate_output_filename(df, idx=1, prefix=""):
    """Generates dynamic Excel filename."""
    if df.empty:
        return f"{prefix}Questionwise_Output_{idx}_{datetime.now().strftime('%d-%m-%y')}.xlsx"

    subject_name = str(df.iloc[0].get("SubjectName", "Subject")).strip()
    semester_raw = str(df.iloc[0].get("Semester", "Semester")).strip()

    semester_short = re.sub(r"(?i)\bsemester\b", "Sem", semester_raw).strip()

    unique_prns = df["PRNNumber"].nunique() if "PRNNumber" in df.columns else len(df)
    date_str = datetime.now().strftime("%d-%m-%y")

    filename = f"{prefix}{subject_name}_{semester_short}_{unique_prns}_{date_str}.xlsx"
    return filename


# ==================================================
# STEP 5: EXCEL GENERATION & STYLING
# ==================================================
def sanitize_sheet_name(subject_identifier):
    """Sanitizes SubjectName / Subject Code into a valid Excel sheet name (max 31 chars)."""
    if not subject_identifier:
        return "Questionwise Data"

    name = str(subject_identifier)
    name = re.sub(r"\[[^\]]*\]", "", name)
    name = re.sub(r"[\\/?*\[\]:]", "", name)
    name = re.sub(r"\s+", " ", name).strip()

    if not name:
        name = "Questionwise Data"

    return name[:31]


def get_sheet_name_for_df(df, default_idx=1):
    """Retrieves sheet name based on 'Subject Name-Subject code' or 'SubjectName'."""
    if not df.empty:
        if "SubjectName" in df.columns and str(df.iloc[0]["SubjectName"]).strip():
            return sanitize_sheet_name(df.iloc[0]["SubjectName"])
        if "Subject Name-Subject code" in df.columns and str(df.iloc[0]["Subject Name-Subject code"]).strip():
            return sanitize_sheet_name(df.iloc[0]["Subject Name-Subject code"])
    return f"Sheet_{default_idx}"


def add_sheet_popup_validation(worksheet, is_eligible):
    """
    Applies Data Validation Input Message to cell A1.
    When a user clicks on/selects cell A1 in Excel, a dynamic floating notification pop-up appears.
    """
    if not is_eligible:
        return

    dv = DataValidation(
        type="custom",
        formula1="TRUE",
        allow_blank=True,
        showInputMessage=True,
        showErrorMessage=False,
        promptTitle="⚠️ 100% Moderation Alert",
        prompt="The 10% change in marks is observed in more than 50% of the moderated answer books. The paper is eligible for 100% Moderation."
    )
    worksheet.add_data_validation(dv)
    dv.add(worksheet["A1"])


# --------------------------------------------------
# Shared Excel-styling helpers
# --------------------------------------------------

def _thin_border():
    """Returns a Border with a thin black line on all four sides."""
    side = Side(style="thin", color="000000")
    return Border(left=side, right=side, top=side, bottom=side)


def _center_column_indices(cols, explicit_center_names):
    """
    Works out the 1-based column indices that should be center-aligned.

    Two rules combine:
      1. Any column whose name is in `explicit_center_names`.
      2. Every column from 'SAP ID' to 'TotalObtainedScore' inclusive - this
         is always the block of per-question mark columns.
    """
    indices = [i for i, name in enumerate(cols, start=1) if name in explicit_center_names]

    if "SAP ID" in cols and "TotalObtainedScore" in cols:
        sap_idx = cols.index("SAP ID") + 1
        tot_idx = cols.index("TotalObtainedScore") + 1
        indices.extend(range(sap_idx, tot_idx + 1))

    return set(indices)


def _autosize_or_fixed_width_columns(worksheet, fixed_width_cols=None, fixed_width=15):
    """
    Sets a fixed width for any column named in `fixed_width_cols`; every other
    column auto-sizes to fit its longest cell value (+3 padding, 10 minimum).
    """
    fixed_width_cols = fixed_width_cols or set()
    for col in worksheet.columns:
        col_letter = get_column_letter(col[0].column)
        col_name = str(col[0].value).strip()

        if col_name in fixed_width_cols:
            worksheet.column_dimensions[col_letter].width = fixed_width
        else:
            max_len = 0
            for cell in col:
                if cell.value is not None:
                    max_len = max(max_len, len(str(cell.value)))
            worksheet.column_dimensions[col_letter].width = max(max_len + 3, 10)


def _apply_grid_styling(worksheet, cols, center_col_indices, total_score_col_idx,
                         download_df=None, user_type_col_idx=None, row_fills_by_type=None):
    """
    Applies the shared per-cell formatting pass used by every exported sheet:
      - thin border on every cell
      - bold, grey, wrapped & centered header row (row 1)
      - bold font on the 'TotalObtainedScore' column
      - center alignment on any column flagged in `center_col_indices`
      - (formatted sheets only) row fill mapped by lowercase UserType from
        `row_fills_by_type` (e.g., Moderator light green; in Revaluation mode
        also Reval 1 pastel purple and Reval 2 pastel orange).
    """
    thin_border = _thin_border()
    grey_fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
    center_wrap_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    center_align = Alignment(horizontal="center", vertical="center")

    for row in worksheet.iter_rows(min_row=1, max_row=worksheet.max_row,
                                    min_col=1, max_col=worksheet.max_column):
        row_fill = None
        if row_fills_by_type and user_type_col_idx is not None and download_df is not None:
            row_idx = row[0].row - 2  # Excel row 2 == DataFrame index 0 (row 1 is the header).
            if 0 <= row_idx < len(download_df):
                u_type = str(download_df.iloc[row_idx].get("UserType", "")).strip().lower()
                row_fill = row_fills_by_type.get(u_type)

        for cell in row:
            cell.border = thin_border

            if cell.row == 1:
                cell.font = Font(bold=True)
                cell.fill = grey_fill
                cell.alignment = center_wrap_align
            else:
                if row_fill is not None:
                    cell.fill = row_fill
                if total_score_col_idx and cell.column == total_score_col_idx:
                    cell.font = Font(bold=True)

            if cell.row > 1 and cell.column in center_col_indices:
                cell.alignment = center_align


def _apply_header_highlight_fills(worksheet, cols, orange_header_cols, dark_grey_header_cols):
    """
    Overrides the default grey header-row fill (set by _apply_grid_styling)
    for specific header cells only:
      - `orange_header_cols` get a medium orange, 60% lighter fill (FCE4D6)
      - `dark_grey_header_cols` get a 35% darker grey fill (999999)
    """
    orange_fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
    dark_grey_fill = PatternFill(start_color="999999", end_color="999999", fill_type="solid")

    for col_name in orange_header_cols:
        if col_name in cols:
            col_idx = cols.index(col_name) + 1
            worksheet.cell(row=1, column=col_idx).fill = orange_fill

    for col_name in dark_grey_header_cols:
        if col_name in cols:
            col_idx = cols.index(col_name) + 1
            worksheet.cell(row=1, column=col_idx).fill = dark_grey_fill


def _apply_change_highlighting_and_group_borders(worksheet, download_df, cols, bright_green_font, red_font, reval_mode=False):
    """
    For every student (grouped by 'SAP ID'), compares each reviewer row
    (Moderator / Reval 1 / Reval 2) against its baseline row and:
      - colors question marks (including TotalObtainedScore), Grade,
        Total Marks Obtained and Change % cells green if improved, red if
        dropped or a mark disappeared
      - draws a medium-thickness bottom border under the student's last row
        so each student's block is visually separated on the sheet

    Baseline selection:
      - Moderation Analysis (reval_mode=False): always Examiner (frozen legacy).
      - Revaluation Analysis (reval_mode=True):
        * Moderator vs Examiner
        * Reval 1 vs Moderator (or Examiner if no Moderator row)
        * Reval 2 vs Reval 1 (fallback Moderator -> Examiner)
    """
    if "UserType" not in download_df.columns or "SAP ID" not in download_df.columns:
        return

    thin_side = Side(style="thin", color="000000")
    medium_side = Side(style="medium", color="000000")

    # Includes TotalObtainedScore intentionally so it receives green/red font highlighting too.
    question_cols = [c for c in download_df.columns if c not in FIXED_COLUMNS and c not in COLUMNS_TO_DROP_FORMATTED]

    for _, group in download_df.groupby("SAP ID", dropna=False):
        if not reval_mode:
            examiner_rows = group[group["UserType"].astype(str).str.strip() == "Examiner"]
            if examiner_rows.empty:
                continue

        for row_idx in group.index:
            u_type = str(download_df.loc[row_idx, "UserType"]).strip()
            if u_type == "Examiner":
                continue

            _, base_row = _get_baseline_row(group, u_type, reval_mode=reval_mode)
            if base_row is None:
                continue

            excel_row_num = row_idx + 2  # +1 for header row, +1 for 1-based Excel rows.

            # --- Question-wise marks (and TotalObtainedScore) ---
            for q_col in question_cols:
                if q_col not in download_df.columns:
                    continue
                col_idx = cols.index(q_col) + 1
                base_val_str = str(base_row[q_col]).strip()
                cur_val_str = str(download_df.loc[row_idx, q_col]).strip()

                if base_val_str == "" or base_val_str.upper() == "NA":
                    continue

                cell = worksheet.cell(row=excel_row_num, column=col_idx)

                if cur_val_str == "" or cur_val_str.upper() == "NA":
                    cell.font = red_font
                    continue

                try:
                    base_val = float(base_val_str)
                    cur_val = float(cur_val_str)
                    if cur_val > base_val:
                        cell.font = bright_green_font
                    elif cur_val < base_val:
                        cell.font = red_font
                except (ValueError, TypeError):
                    pass

            # --- Grade (rank comparison) ---
            if "Grade" in download_df.columns:
                grade_col_idx = cols.index("Grade") + 1
                base_grade = str(base_row.get("Grade", "")).strip().upper()
                cur_grade = str(download_df.loc[row_idx, "Grade"]).strip().upper()

                if base_grade in GRADE_RANKS and cur_grade in GRADE_RANKS:
                    cell = worksheet.cell(row=excel_row_num, column=grade_col_idx)
                    if GRADE_RANKS[cur_grade] > GRADE_RANKS[base_grade]:
                        cell.font = bright_green_font
                    elif GRADE_RANKS[cur_grade] < GRADE_RANKS[base_grade]:
                        cell.font = red_font

            # --- Total Marks Obtained ---
            if "Total Marks Obtained" in download_df.columns:
                tot_obt_col_idx = cols.index("Total Marks Obtained") + 1
                base_tot_str = str(base_row.get("Total Marks Obtained", "")).strip()
                cur_tot_str = str(download_df.loc[row_idx, "Total Marks Obtained"]).strip()

                if base_tot_str != "" and cur_tot_str != "":
                    cell = worksheet.cell(row=excel_row_num, column=tot_obt_col_idx)
                    try:
                        base_tot = float(base_tot_str)
                        cur_tot = float(cur_tot_str)
                        if cur_tot > base_tot:
                            cell.font = bright_green_font
                        elif cur_tot < base_tot:
                            cell.font = red_font
                    except (ValueError, TypeError):
                        pass

            # --- Change % (sign-based) ---
            if "Change %" in download_df.columns:
                chg_col_idx = cols.index("Change %") + 1
                chg_str = str(download_df.loc[row_idx, "Change %"]).strip()

                if chg_str != "":
                    cell = worksheet.cell(row=excel_row_num, column=chg_col_idx)
                    try:
                        chg_val = float(chg_str)
                        if chg_val > 0:
                            cell.font = bright_green_font
                        elif chg_val < 0:
                            cell.font = red_font
                    except (ValueError, TypeError):
                        pass

        # Thick bottom border under this student's last row, marking the end of their block.
        last_row_idx = group.index[-1] + 2
        for c_idx in range(1, worksheet.max_column + 1):
            cell = worksheet.cell(row=last_row_idx, column=c_idx)
            cell.border = Border(left=thin_side, right=thin_side, top=cell.border.top, bottom=medium_side)


def _apply_semester_marks_highlighting(worksheet, download_df, cols, green_fill, red_fill):
    """
    For every row where 'Semester Marks' is populated (Moderator rows, or an
    Examiner row with no corresponding Moderator record), compares it against
    that same row's TotalObtainedScore and flags mismatches in red/yellow with comment.
    """
    if "Semester Marks" not in download_df.columns or "TotalObtainedScore" not in download_df.columns:
        return

    sem_col_idx = cols.index("Semester Marks") + 1

    for row_idx in download_df.index:
        sem_val_str = str(download_df.loc[row_idx, "Semester Marks"]).strip()
        if sem_val_str == "" or sem_val_str.upper() == "NA":
            continue

        tot_val_str = str(download_df.loc[row_idx, "TotalObtainedScore"]).strip()
        try:
            sem_val = float(re.sub(r"^\+\s*", "", sem_val_str))
            tot_val = float(tot_val_str)
        except (ValueError, TypeError):
            continue

        excel_row_num = row_idx + 2
        cell = worksheet.cell(row=excel_row_num, column=sem_col_idx)
        if sem_val != tot_val:
            cell.fill = red_fill
            f = cell.font
            cell.font = Font(name=f.name, size=f.size, bold=True, italic=f.italic, color="FFFF00")
            cell.comment = Comment(
                f"Semester Marks ({_fmt_num_clean(sem_val)}) does not match Total Obtained Score ({_fmt_num_clean(tot_val)}).",
                "Questionwise Checker",
            )


def _to_float_or_none(value):
    """Parses a cell value (allowing a leading '+') into a float, or None if blank / non-numeric."""
    try:
        return float(re.sub(r"^\+\s*", "", str(value).strip()))
    except (ValueError, TypeError):
        return None


def _apply_total_and_grade_validation(worksheet, download_df, cols):
    """
    Excel-only validation highlights (red fill + bold yellow font):
      - Total Marks Obtained must equal Semester Marks + Internal Marks
      - Grade must equal the grade derived from Total Marks Obtained / Total Max Marks
    """
    red_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")

    def _flag(row_idx, col_name, note=""):
        cell = worksheet.cell(row=row_idx + 2, column=cols.index(col_name) + 1)
        cell.fill = red_fill
        f = cell.font
        cell.font = Font(name=f.name, size=f.size, bold=True, italic=f.italic, color="FFFF00")
        if note:
            cell.comment = Comment(note, "Questionwise Checker")

    has_total = "Total Marks Obtained" in cols
    can_check_total = has_total and "Semester Marks" in cols and "Internal Marks" in cols
    can_check_grade = has_total and "Grade" in cols and "Total Max Marks" in cols

    for row_idx in download_df.index:
        total_obtained = _to_float_or_none(download_df.loc[row_idx, "Total Marks Obtained"]) if has_total else None
        if total_obtained is None:
            continue

        if can_check_total:
            semester = _to_float_or_none(str(download_df.loc[row_idx, "Semester Marks"]).strip("!"))
            if semester is not None:
                internal = _to_float_or_none(download_df.loc[row_idx, "Internal Marks"])
                expected_total = semester + (internal or 0.0)
                if abs(expected_total - total_obtained) > 0.005:
                    _flag(
                        row_idx,
                        "Total Marks Obtained",
                        f"Total Marks Obtained ({_fmt_num_clean(total_obtained)}) != Semester Marks + Internal Marks ({_fmt_num_clean(expected_total)}).",
                    )

        if can_check_grade:
            shown_grade = str(download_df.loc[row_idx, "Grade"]).strip()
            expected_grade = calculate_grade(total_obtained, download_df.loc[row_idx, "Total Max Marks"])
            if shown_grade != "" and expected_grade != "" and shown_grade.upper() != expected_grade.upper():
                _flag(
                    row_idx,
                    "Grade",
                    f"Grade '{shown_grade}' does not match calculated grade '{expected_grade}' from marks percentage.",
                )


def _is_zero_after_marks(baseline_val, reviewer_val):
    """True when the reviewer awarded 0 for a question where the baseline awarded non-zero marks."""
    try:
        base_num = float(re.sub(r"^\+\s*", "", str(baseline_val).strip()))
        rev_num = float(re.sub(r"^\+\s*", "", str(reviewer_val).strip()))
    except (ValueError, TypeError):
        return False
    return rev_num == 0 and base_num != 0


def _change_fill(diff, change_pct):
    """
    Picks the SAP ID fill from the Baseline -> Reviewer score change.
    diff > 0  -> pastel green (light/medium/darker by |Change %|)
    diff < 0  -> pastel red   (light/medium/darker by |Change %|)
    diff == 0 -> pastel blue
    A missing/non-numeric Change % falls back to the lightest tier.
    """
    if diff == 0:
        return PatternFill(start_color=SAP_NO_CHANGE_FILL, end_color=SAP_NO_CHANGE_FILL, fill_type="solid")

    try:
        magnitude = abs(float(str(change_pct).strip()))
    except (ValueError, TypeError):
        magnitude = 0.0

    low, high = CHANGE_TIER_LIMITS
    tier = 0 if magnitude <= low else (1 if magnitude <= high else 2)
    colour = (SAP_INCREASE_FILLS if diff > 0 else SAP_DECREASE_FILLS)[tier]
    return PatternFill(start_color=colour, end_color=colour, fill_type="solid")


def _apply_enhanced_formatting(worksheet, download_df, cols, bright_green_font, reval_mode=False, best_k_map=None):
    """
    Extra formatting pass for the Formatted File / Consolidated Excel. Runs LAST,
    after all existing styling and column auto-sizing:
      - fixed column widths (question columns + SrNo/Subject code/SubjectName/SAP ID/UserType/score)
      - NA / 0 mismatch yellow fill and NA->marks bright green bold font:
          * Moderation Analysis (reval_mode=False): Moderator vs Examiner only (frozen legacy).
          * Revaluation Analysis (reval_mode=True): Moderator vs Examiner, Reval 1 vs Moderator
            (or Examiner), and Reval 2 vs Reval 1.
      - Tan fill on rows of students that only have an Examiner row
      - SAP ID fill by score change:
          * Moderation Analysis (reval_mode=False): Examiner->Moderator applied across student block.
          * Revaluation Analysis (reval_mode=True): Row-specific (Option A) SAP ID fill per stage.
      - Red SAP ID fill on the exact (baseline, reviewer) row pair where NA/0 mismatch occurred
      - thick vertical borders between main question groups and around the score column
      - row heights (header 45pt, others 21pt) and vertical-middle alignment (all cells)
      - excluded sub-question {values} (using precomputed paper-wide best_k_map when provided)
      - Total Marks Obtained / Grade validation highlights (red fill + bold yellow font)
      - predicted Semester Total Max Marks (!number!) bold gray styling
      - rename 'TotalObtainedScore' and 'ExamAssingment_ID' header labels last
    """
    question_cols = _get_pure_question_columns(cols)

    # 1. Fixed widths: question-wise mark columns, then the named columns.
    for q_col in question_cols:
        letter = get_column_letter(cols.index(q_col) + 1)
        worksheet.column_dimensions[letter].width = QUESTION_COL_WIDTH
    for col_name, width in FIXED_COLUMN_WIDTHS.items():
        if col_name in cols:
            worksheet.column_dimensions[get_column_letter(cols.index(col_name) + 1)].width = width

    # 2 & 3. Reviewer vs Baseline "NA" and 0-after-marks handling.
    red_sap_rows = []   # DataFrame row indices whose SAP ID cell gets the red fill (applied after step 5).
    if "UserType" in cols and "SAP ID" in cols and question_cols:
        yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
        checked_types = ("Moderator", "Reval 1", "Reval 2") if reval_mode else ("Moderator",)

        for _, group in download_df.groupby("SAP ID", dropna=False):
            for row_idx in group.index:
                u_type = str(download_df.loc[row_idx, "UserType"]).strip()
                if u_type not in checked_types:
                    continue

                base_idx, base_row = _get_baseline_row(group, u_type, reval_mode=reval_mode)
                if base_row is None:
                    continue

                excel_row_num = row_idx + 2
                na_hit = False

                for q_col in question_cols:
                    base_val = str(base_row[q_col]).strip()
                    rev_val = str(download_df.loc[row_idx, q_col]).strip()
                    cell = worksheet.cell(row=excel_row_num, column=cols.index(q_col) + 1)

                    if base_val not in ("",) and base_val.upper() != "NA" and (
                        rev_val.upper() == "NA" or _is_zero_after_marks(base_val, rev_val)
                    ):
                        cell.fill = yellow_fill                  # Baseline marked, Reviewer "NA" or 0.
                        na_hit = True
                    elif base_val.upper() == "NA" and rev_val != "" and rev_val.upper() != "NA":
                        cell.font = bright_green_font            # Baseline "NA", Reviewer marked.

                if na_hit:
                    red_sap_rows.extend([base_idx, row_idx])

    # 4. Examiner-only students (no Moderator / Reval rows) -> Tan fill on the whole row.
    if "UserType" in cols and "SAP ID" in cols:
        tan_fill = PatternFill(start_color=EXAMINER_ONLY_FILL, end_color=EXAMINER_ONLY_FILL, fill_type="solid")
        for _, group in download_df.groupby("SAP ID", dropna=False):
            if set(group["UserType"].astype(str).str.strip()) != {"Examiner"}:
                continue
            for row_idx in group.index:
                for col_idx in range(1, worksheet.max_column + 1):
                    cell = worksheet.cell(row=row_idx + 2, column=col_idx)
                    if cell.fill is None or cell.fill.fill_type is None:
                        cell.fill = tan_fill

    # 5. SAP ID colour by score change.
    if "SAP ID" in cols and "UserType" in cols and "TotalObtainedScore" in cols:
        sap_col_idx = cols.index("SAP ID") + 1

        if not reval_mode:
            # Moderation Analysis (frozen legacy): Examiner -> Moderator applied to every row of student.
            for _, group in download_df.groupby("SAP ID", dropna=False):
                user_types = group["UserType"].astype(str).str.strip()
                examiner_rows = group[user_types == "Examiner"]
                moderator_rows = group[user_types == "Moderator"]
                if examiner_rows.empty or moderator_rows.empty:
                    continue

                try:
                    ex_score = float(examiner_rows.iloc[0]["TotalObtainedScore"])
                    mod_score = float(moderator_rows.iloc[0]["TotalObtainedScore"])
                except (ValueError, TypeError):
                    continue

                fill = _change_fill(mod_score - ex_score, moderator_rows.iloc[0].get("Change %", ""))
                for row_idx in group.index:
                    worksheet.cell(row=row_idx + 2, column=sap_col_idx).fill = fill
        else:
            # Revaluation Analysis (Option A - Row-specific SAP ID colour):
            # Each reviewer row gets the colour for its own stage transition vs baseline.
            # Examiner's SAP ID cell gets the colour of its immediate next stage (Moderator if present,
            # otherwise Reval 1, otherwise Reval 2).
            for _, group in download_df.groupby("SAP ID", dropna=False):
                examiner_colored = False
                for row_idx in group.index:
                    u_type = str(download_df.loc[row_idx, "UserType"]).strip()
                    if u_type not in ("Moderator", "Reval 1", "Reval 2"):
                        continue

                    base_idx, base_row = _get_baseline_row(group, u_type, reval_mode=True)
                    if base_row is None:
                        continue

                    try:
                        base_score = float(base_row["TotalObtainedScore"])
                        cur_score = float(download_df.loc[row_idx, "TotalObtainedScore"])
                    except (ValueError, TypeError):
                        continue

                    fill = _change_fill(cur_score - base_score, download_df.loc[row_idx].get("Change %", ""))
                    worksheet.cell(row=row_idx + 2, column=sap_col_idx).fill = fill

                    if not examiner_colored and str(base_row["UserType"]).strip() == "Examiner":
                        worksheet.cell(row=base_idx + 2, column=sap_col_idx).fill = fill
                        examiner_colored = True

    # 5b. Red SAP ID on baseline + reviewer rows where reviewer is "NA" or 0 after baseline marks.
    if red_sap_rows and "SAP ID" in cols:
        red_fill = PatternFill(start_color=NA_MISMATCH_SAP_FILL, end_color=NA_MISMATCH_SAP_FILL, fill_type="solid")
        for row_idx in set(red_sap_rows):
            worksheet.cell(row=row_idx + 2, column=cols.index("SAP ID") + 1).fill = red_fill

    # 5c. Total Marks Obtained (= Semester Marks + Internal Marks) and Grade validation.
    _apply_total_and_grade_validation(worksheet, download_df, cols)

    # 6. Thick vertical borders (top/bottom of each cell are preserved).
    thick_side = Side(style="thick", color="000000")

    def _thick_vertical(col_idx, left=False, right=False):
        for r in range(1, worksheet.max_row + 1):
            cell = worksheet.cell(row=r, column=col_idx)
            b = cell.border
            cell.border = Border(
                left=thick_side if left else b.left,
                right=thick_side if right else b.right,
                top=b.top,
                bottom=b.bottom,
            )

    # 6a. Between main question groups (Q1 | Q2 | Q3 | ... across all numbering styles).
    groups = []
    prev_key = None
    for q_col in question_cols:
        col_idx = cols.index(q_col) + 1
        q_info = parse_question_column_info(q_col)
        key = q_info["main_q"] if q_info else q_col
        if groups and key == prev_key and col_idx == groups[-1][1] + 1:
            groups[-1][1] = col_idx
        else:
            groups.append([col_idx, col_idx])
        prev_key = key
    for g_i, (first_idx, last_idx) in enumerate(groups):
        if g_i > 0:
            _thick_vertical(first_idx, left=True)
        if g_i < len(groups) - 1:
            _thick_vertical(last_idx, right=True)

    # 6b. Both sides of the Total Obtained Score column.
    if "TotalObtainedScore" in cols:
        _thick_vertical(cols.index("TotalObtainedScore") + 1, left=True, right=True)

    # 7. Fixed row heights + vertical-middle alignment on every cell (header + data).
    for row in worksheet.iter_rows(min_row=1, max_row=worksheet.max_row,
                                    min_col=1, max_col=worksheet.max_column):
        worksheet.row_dimensions[row[0].row].height = HEADER_ROW_HEIGHT if row[0].row == 1 else EXCEL_ROW_HEIGHT
        for cell in row:
            a = cell.alignment
            cell.alignment = Alignment(
                horizontal=a.horizontal, vertical="center", textRotation=a.textRotation,
                wrap_text=a.wrap_text, shrink_to_fit=a.shrink_to_fit, indent=a.indent,
            )

    # 5d. Indicate excluded (uncounted) sub-question values in {} preserving original string format
    #     with blue font, show (0.25)/(0.5) when rounded, and validate non-excluded sum == TotalObtainedScore.
    if question_cols:
        error_red_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
        resolved_best_k = best_k_map if best_k_map is not None else infer_subquestion_best_k_map(download_df, question_cols)
        tot_col_idx = cols.index("TotalObtainedScore") + 1 if "TotalObtainedScore" in cols else None
        ex_extra_col_idx = cols.index("Excluded Extra Marks") + 1 if "Excluded Extra Marks" in cols else None

        for row_idx, row in download_df.iterrows():
            ex_val, ex_cols, counted_sum, is_half_round_up = compute_row_excluded_extra_marks(
                row, download_df, question_cols, best_k_map=resolved_best_k
            )
            for ex_col in ex_cols:
                if ex_col in cols:
                    c_idx = cols.index(ex_col) + 1
                    cell = worksheet.cell(row=row_idx + 2, column=c_idx)
                    raw_str = str(row.get(ex_col, "")).strip()
                    if raw_str == "":
                        raw_str = str(cell.value or "").strip()
                    if raw_str and not (raw_str.startswith("{") and raw_str.endswith("}")):
                        cell.value = f"{{{raw_str}}}"
                    f = cell.font
                    cell.font = Font(
                        name=f.name if f else None,
                        size=f.size if f else None,
                        bold=True,
                        italic=f.italic if f else False,
                        color="0000FF",
                    )

            if ex_extra_col_idx is not None and ex_val != "":
                worksheet.cell(row=row_idx + 2, column=ex_extra_col_idx).value = str(ex_val)

            if tot_col_idx is not None and counted_sum is not None and not is_half_round_up:
                tot_num = _parse_numeric_mark(row.get("TotalObtainedScore"))
                if tot_num is not None and abs(counted_sum - tot_num) > 1e-6:
                    tot_cell = worksheet.cell(row=row_idx + 2, column=tot_col_idx)
                    tot_cell.fill = error_red_fill
                    tf = tot_cell.font
                    tot_cell.font = Font(
                        name=tf.name if tf else None,
                        size=tf.size if tf else None,
                        bold=True,
                        italic=tf.italic if tf else False,
                        color="FFFF00",
                    )
                    tot_cell.comment = Comment(
                        f"Total Obtained Score ({_fmt_num_clean(tot_num)}) != Non-Excluded Sub-Question Total ({_fmt_num_clean(counted_sum)}).",
                        "Questionwise Checker",
                    )

    # 5e. Bold gray text on predicted Semester Total Max Marks (!number!).
    if "Semester Total Max Marks" in cols:
        sem_max_col_idx = cols.index("Semester Total Max Marks") + 1
        gray_bold_font = Font(bold=True, color="7F7F7F")
        for r_idx in range(2, worksheet.max_row + 1):
            c = worksheet.cell(row=r_idx, column=sem_max_col_idx)
            c_txt = str(c.value or "").strip()
            if c_txt.startswith("!") and c_txt.endswith("!"):
                c.font = gray_bold_font

    # 8. Rename the header labels last (all column lookups above use the original names).
    if "TotalObtainedScore" in cols:
        worksheet.cell(row=1, column=cols.index("TotalObtainedScore") + 1).value = TOTAL_SCORE_LABEL
    if "ExamAssingment_ID" in cols:
        worksheet.cell(row=1, column=cols.index("ExamAssingment_ID") + 1).value = ASSIGNMENT_ID_LABEL


def _style_formatted_worksheet(worksheet, df, download_df, reval_mode=False, best_k_map=None):
    """
    Shared styling pipeline for both individual Formatted Excel files and
    each sheet of the Consolidated Multi-Sheet Excel workbook.
    """
    worksheet.views.sheetView[0].showGridLines = False
    worksheet.freeze_panes = "G2"          # Freeze header row + first 6 identifying columns (incl. ExamAssingment_ID).
    worksheet.row_dimensions[1].height = 30

    row_fills_by_type = {
        "moderator": PatternFill(start_color=MODERATOR_ROW_FILL, end_color=MODERATOR_ROW_FILL, fill_type="solid"),
    }
    if reval_mode:
        row_fills_by_type["reval 1"] = PatternFill(
            start_color=REVAL1_ROW_FILL, end_color=REVAL1_ROW_FILL, fill_type="solid"
        )
        row_fills_by_type["reval 2"] = PatternFill(
            start_color=REVAL2_ROW_FILL, end_color=REVAL2_ROW_FILL, fill_type="solid"
        )

    bright_green_font = Font(color="00B050", bold=True)
    red_font = Font(color="FF4500", bold=True)
    sem_marks_green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    sem_marks_red_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")

    cols = download_df.columns.tolist()

    fixed_10_cols = {
        "Semester Total Max Marks", "Semester Marks", "Internal Marks",
        "Total Max Marks", "Total Marks Obtained", "Grade", "Change %", "Excluded Extra Marks"
    }
    align_center_col_names = {
        "SrNo", "Subject code", "Grade", "Internal Marks", "Semester Marks",
        "Total Max Marks", "Semester Total Max Marks", "Total Marks Obtained", "Change %", "Excluded Extra Marks"
    }

    center_col_indices = _center_column_indices(cols, align_center_col_names)
    total_score_col_idx = cols.index("TotalObtainedScore") + 1 if "TotalObtainedScore" in cols else None
    user_type_col_idx = cols.index("UserType") if "UserType" in cols else None

    # Basic styling + UserType row fills.
    _apply_grid_styling(
        worksheet, cols, center_col_indices, total_score_col_idx,
        download_df=download_df, user_type_col_idx=user_type_col_idx, row_fills_by_type=row_fills_by_type,
    )

    # Green/red change highlighting + thick per-student borders.
    _apply_change_highlighting_and_group_borders(
        worksheet, download_df, cols, bright_green_font, red_font, reval_mode=reval_mode
    )

    # Semester Marks vs TotalObtainedScore match/mismatch cell fill (unchanged across modes).
    _apply_semester_marks_highlighting(worksheet, download_df, cols, sem_marks_green_fill, sem_marks_red_fill)

    # Header cell fills: GMR/Change % columns -> medium orange (60% lighter);
    # TotalObtainedScore -> 35% darker grey.
    _apply_header_highlight_fills(worksheet, cols, fixed_10_cols, {"TotalObtainedScore"})

    # Interactive pop-up Data Validation message (Moderation Analysis only; suppressed in Revaluation Analysis).
    if not reval_mode:
        _, _, _, is_eligible = compute_moderation_10pct_metrics(df)
        add_sheet_popup_validation(worksheet, is_eligible)

    # Auto-size columns, except the GMR/Change % block which gets a fixed width of 10.
    _autosize_or_fixed_width_columns(worksheet, fixed_width_cols=fixed_10_cols, fixed_width=10)

    # Extra formatting (question widths, row height, vertical middle, NA/score-change highlights, excluded marks).
    _apply_enhanced_formatting(
        worksheet, download_df, cols, bright_green_font, reval_mode=reval_mode, best_k_map=best_k_map
    )


def to_excel_bytes(df):
    """ORIGINAL Standard clean output without GMR columns (Moderation Analysis only)."""
    download_df = process_data_for_download(df, include_all_usertypes=False)
    buffer = io.BytesIO()

    sheet_name = "Questionwise Data"
    if not df.empty and "SubjectName" in df.columns:
        subj = str(df.iloc[0]["SubjectName"]).strip()
        if subj:
            sheet_name = sanitize_sheet_name(subj)

    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        download_df.to_excel(writer, index=False, sheet_name=sheet_name)

        worksheet = writer.sheets[sheet_name]

        worksheet.views.sheetView[0].showGridLines = False
        worksheet.freeze_panes = "A2"
        worksheet.row_dimensions[1].height = 30

        cols = download_df.columns.tolist()
        center_col_indices = _center_column_indices(cols, {"SrNo", "Subject code"})
        total_score_col_idx = cols.index("TotalObtainedScore") + 1 if "TotalObtainedScore" in cols else None

        _apply_grid_styling(worksheet, cols, center_col_indices, total_score_col_idx)
        _autosize_or_fixed_width_columns(worksheet)

    return buffer.getvalue()


def to_excel_bytes_formatted(df, reval_mode=False, best_k_map=None):
    """
    Formatted output Excel containing all kept UserTypes, GMR columns, Change %,
    Excluded Extra Marks, conditional formatting, borders, and validation comments.
    """
    download_df = process_data_for_download(df, include_all_usertypes=True)
    buffer = io.BytesIO()

    sheet_name = get_sheet_name_for_df(df)

    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        download_df.to_excel(writer, index=False, sheet_name=sheet_name)
        worksheet = writer.sheets[sheet_name]
        _style_formatted_worksheet(worksheet, df, download_df, reval_mode=reval_mode, best_k_map=best_k_map)

    return buffer.getvalue()


def create_consolidated_multi_sheet_excel(dfs_list, reval_mode=False):
    """
    Clubs all processed dataframes into a single Excel workbook with separate
    formatted sheets named by SubjectName / Subject Code.
    Each entry in `dfs_list` may be `(df, orig_name)` or `(df, orig_name, best_k_map)`.
    """
    buffer = io.BytesIO()
    used_sheet_names = set()

    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        for idx, item in enumerate(dfs_list, start=1):
            if len(item) == 3:
                df, _, best_k_map = item
            else:
                df, _ = item
                best_k_map = None

            download_df = process_data_for_download(df, include_all_usertypes=True)

            # Ensure sheet names are unique within the workbook (Excel forbids duplicates).
            base_sheet_name = get_sheet_name_for_df(df, idx)
            sheet_name = base_sheet_name
            counter = 1
            while sheet_name in used_sheet_names:
                sheet_name = f"{base_sheet_name[:28]}_{counter}"
                counter += 1
            used_sheet_names.add(sheet_name)

            download_df.to_excel(writer, index=False, sheet_name=sheet_name)
            worksheet = writer.sheets[sheet_name]
            _style_formatted_worksheet(worksheet, df, download_df, reval_mode=reval_mode, best_k_map=best_k_map)

    return buffer.getvalue()


def create_zip_file(excel_files):
    """Creates a ZIP archive containing multiple Excel file bytes."""
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
        for filename, file_bytes in excel_files:
            zip_file.writestr(filename, file_bytes)
    return zip_buffer.getvalue()


# ==================================================
# MAIN ENTRY POINT
# ==================================================
def show():
    """Render the Questionwise Checker page in Streamlit."""
    st.title("📝 Questionwise Checker")
    st.caption("Upload password-protected or unprotected Excel files.")

    # Analysis Mode Toggle.
    analysis_mode = st.radio(
        "Analysis Mode",
        options=["Moderation Analysis", "Revaluation Analysis"],
        horizontal=True,
        help=(
            "Moderation Analysis: Standard workflow for all students (Cleaned, Formatted, ZIP & Consolidated outputs). "
            "Revaluation Analysis: Filters to students with Reval 1 and/or Reval 2 rows and generates Formatted & Consolidated outputs only."
        ),
    )
    is_reval_mode = (analysis_mode == "Revaluation Analysis")

    # Side-by-side Upload Controls.
    col_up1, col_up2 = st.columns(2)

    with col_up1:
        uploaded_files = st.file_uploader(
            "Upload Questionwise Excel/Excels",
            type=["xlsx"],
            accept_multiple_files=True
        )

    with col_up2:
        gmr_file = st.file_uploader(
            "Upload VGMR Excel",
            type=["xlsx", "xls"]
        )

    if not uploaded_files:
        st.markdown("---")
        render_help_section()
        return

    # Load GMR File if present (optional - tool still works without it).
    gmr_df = None
    if gmr_file is not None:
        with st.spinner("Loading VGMR Excel..."):
            gmr_df = load_gmr_file(gmr_file)
            if gmr_df is not None:
                st.caption("✅ VGMR File uploaded and loaded successfully.")
            else:
                st.warning("⚠️ Could not load VGMR Excel file.")

    processed_outputs_clean = []   # (filename, bytes) pairs for the bulk ZIP download (Moderation mode only).
    dfs_for_multi_sheet = []       # (DataFrame, original filename, best_k_map) tuples for the consolidated workbook.

    for idx, uploaded_file in enumerate(uploaded_files, start=1):
        st.markdown("---")
        st.subheader(f"📁 File {idx}: {uploaded_file.name}")

        try:
            with st.spinner(f"Decrypting and processing {uploaded_file.name}..."):
                df, used_password = decrypt_and_load(uploaded_file)

            if df is None:
                st.error(f"Unable to open {uploaded_file.name}. Password not recognised.")
                continue

            st.caption(f"Opened successfully using: {used_password}")

            # Clean and Sort on the full file first.
            df = clean_and_sort(df)

            if df.empty:
                st.warning(f"No valid student rows found in {uploaded_file.name}.")
                continue

            if "PRNNumber" not in df.columns or "UserType" not in df.columns:
                st.error(f"Missing required 'PRNNumber' or 'UserType' columns in {uploaded_file.name}.")
                continue

            # Infer best_k_map on the FULL paper BEFORE any Revaluation student filtering
            # so sub-question counting and predicted Semester Total Max Marks stay 100% accurate.
            pure_question_columns = _get_pure_question_columns(df)
            full_best_k_map = infer_subquestion_best_k_map(df, pure_question_columns)

            # Apply GMR Lookup Match, Predicted Semester Total Max Marks (using full paper best_k),
            # Change % Calculation, column reordering, and Excluded Extra Marks.
            df = apply_gmr_lookup(df, gmr_df)
            df = apply_predicted_semester_total_max_marks(df, pure_question_columns, best_k_map=full_best_k_map)
            df = calculate_moderator_change_pct(df, reval_mode=is_reval_mode)
            df = reorder_gmr_columns(df)
            df = add_excluded_extra_marks_column(df, pure_question_columns, best_k_map=full_best_k_map)

            total_students_all = df["PRNNumber"].nunique()

            if is_reval_mode:
                active_df = filter_reval_students(df)
                if active_df.empty:
                    st.subheader(build_paper_info(df))
                    st.info(
                        f"No students with Reval 1 or Reval 2 rows found in {uploaded_file.name} "
                        f"(out of {total_students_all} total students). Skipping formatted export."
                    )
                    continue
            else:
                active_df = df

            dfs_for_multi_sheet.append((active_df, uploaded_file.name, full_best_k_map))

            # --- Summary statistics ---
            total_students = active_df["PRNNumber"].nunique()
            moderation_count = count_reviewers(active_df, "Moderator")
            reval1_count = count_reviewers(active_df, "Reval 1")
            reval2_count = count_reviewers(active_df, "Reval 2")

            moderator_positive, moderator_negative, moderator_no_change = score_change_summary(
                active_df, "Moderator", reval_mode=is_reval_mode
            )
            reval1_positive, reval1_negative, reval1_no_change = score_change_summary(
                active_df, "Reval 1", reval_mode=is_reval_mode
            )
            reval2_positive, reval2_negative, reval2_no_change = score_change_summary(
                active_df, "Reval 2", reval_mode=is_reval_mode
            )

            st.subheader(build_paper_info(active_df))

            if not is_reval_mode:
                mod_cnt, cnt_gt_10, ratio_gt_10, is_eligible = compute_moderation_10pct_metrics(active_df)

                # --- 5 Summary Cards Grid (Moderation Analysis) ---
                col1, col2, col3, col4, col5 = st.columns(5)

                with col1:
                    st.info(
                        f"Total Students: {total_students}\n\n"
                        f"Moderation Count: {moderation_count}\n\n"
                        f"Reval 1 Count: {reval1_count}\n\n"
                        f"Reval 2 Count: {reval2_count}"
                    )

                with col2:
                    st.success(
                        f"Positive Change in Marks:\n\n"
                        f"Moderation: {moderator_positive}\n\n"
                        f"Reval 1: {reval1_positive}\n\n"
                        f"Reval 2: {reval2_positive}"
                    )

                with col3:
                    st.error(
                        f"Negative Change in Marks:\n\n"
                        f"Moderation: {moderator_negative}\n\n"
                        f"Reval 1: {reval1_negative}\n\n"
                        f"Reval 2: {reval2_negative}"
                    )

                with col4:
                    st.warning(
                        f"No Change in Marks:\n\n"
                        f"Moderation: {moderator_no_change}\n\n"
                        f"Reval 1: {reval1_no_change}\n\n"
                        f"Reval 2: {reval2_no_change}"
                    )

                with col5:
                    pct_str = f"{ratio_gt_10 * 100:.1f}%"
                    color_style = "color:#FF4500;" if is_eligible else "color:#9ACD32;"

                    st.markdown(
                        f"""
                        <div style="background-color:#586987;padding:12px;border-radius:8px;color:white;border:1px solid #334155;">
                            <span style="font-size:15px;font-weight:600;"> Require 100% Moderation?</span><br/>
                            <span style="font-size:15px;"> Change % > 10%: <b>{cnt_gt_10} / {mod_cnt}</b></span><br/>
                            <span style="font-size:60px;font-weight:bold;{color_style}">{pct_str}</span>
                        </div>
                        """,
                        unsafe_allow_html=True
                    )

                if is_eligible:
                    st.error("Above 10% change in marks is observed in more than 50% of the moderated answer books. This module is eligible for 100% Moderation.")

                # --- Generate individual Excel Outputs (Moderation Analysis: Cleaned + Formatted) ---
                out_bytes_clean = to_excel_bytes(active_df)
                out_bytes_fmt = to_excel_bytes_formatted(active_df, reval_mode=False, best_k_map=full_best_k_map)
                out_filename = generate_output_filename(active_df, idx)
                out_filename_fmt = generate_output_filename(active_df, idx, prefix="Formatted_")

                processed_outputs_clean.append((out_filename, out_bytes_clean))

                dl_col1, dl_col2 = st.columns(2)
                with dl_col1:
                    st.download_button(
                        label=f"📥 Download Cleaned File: {out_filename}",
                        data=out_bytes_clean,
                        file_name=out_filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=f"dl_single_clean_{idx}"
                    )
                with dl_col2:
                    st.download_button(
                        label=f"🎨 Download Formatted File: {out_filename_fmt}",
                        data=out_bytes_fmt,
                        file_name=out_filename_fmt,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=f"dl_single_fmt_{idx}"
                    )
            else:
                # --- 4 Summary Cards Grid (Revaluation Analysis - 100% Moderation Alert Suppressed) ---
                col1, col2, col3, col4 = st.columns(4)

                with col1:
                    st.info(
                        f"Reval Students: {total_students} (of {total_students_all})\n\n"
                        f"With Moderation: {moderation_count}\n\n"
                        f"Reval 1 Count: {reval1_count}\n\n"
                        f"Reval 2 Count: {reval2_count}"
                    )

                with col2:
                    st.success(
                        f"Positive Change in Marks:\n\n"
                        f"Moderation (vs Exam.): {moderator_positive}\n\n"
                        f"Reval 1 (vs Base): {reval1_positive}\n\n"
                        f"Reval 2 (vs Reval 1): {reval2_positive}"
                    )

                with col3:
                    st.error(
                        f"Negative Change in Marks:\n\n"
                        f"Moderation (vs Exam.): {moderator_negative}\n\n"
                        f"Reval 1 (vs Base): {reval1_negative}\n\n"
                        f"Reval 2 (vs Reval 1): {reval2_negative}"
                    )

                with col4:
                    st.warning(
                        f"No Change in Marks:\n\n"
                        f"Moderation (vs Exam.): {moderator_no_change}\n\n"
                        f"Reval 1 (vs Base): {reval1_no_change}\n\n"
                        f"Reval 2 (vs Reval 1): {reval2_no_change}"
                    )

                # --- Generate individual Formatted Excel Output ONLY (Revaluation Analysis) ---
                out_bytes_fmt = to_excel_bytes_formatted(active_df, reval_mode=True, best_k_map=full_best_k_map)
                out_filename_fmt = generate_output_filename(active_df, idx, prefix="Reval_Formatted_")

                st.download_button(
                    label=f"🎨 Download Formatted File: {out_filename_fmt}",
                    data=out_bytes_fmt,
                    file_name=out_filename_fmt,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"dl_single_reval_fmt_{idx}"
                )

        except Exception as error:
            st.error(f"Error processing file {uploaded_file.name}: {error}")

    # --- Bulk Download Options ---
    if not is_reval_mode and processed_outputs_clean:
        st.markdown("---")
        st.header("📦 Bulk Download Options")

        bulk_col1, bulk_col2 = st.columns(2)

        with bulk_col1:
            zip_filename = f"Questionwise_Processed_{datetime.now().strftime('%d-%m-%y_%H%M%S')}.zip"
            zip_bytes = create_zip_file(processed_outputs_clean)

            st.download_button(
                label="📦 Download Question wise Excels for College (ZIP)",
                data=zip_bytes,
                file_name=zip_filename,
                mime="application/zip",
                key="dl_zip_all"
            )

        with bulk_col2:
            if len(dfs_for_multi_sheet) > 1:
                multi_sheet_bytes = create_consolidated_multi_sheet_excel(dfs_for_multi_sheet, reval_mode=False)
                multi_sheet_filename = f"Consolidated_Questionwise_{datetime.now().strftime('%d-%m-%y_%H%M%S')}.xlsx"

                st.download_button(
                    label="📊 Download Consolidated Multi-Sheet Excel",
                    data=multi_sheet_bytes,
                    file_name=multi_sheet_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="dl_multi_sheet_all"
                )
    elif is_reval_mode and len(dfs_for_multi_sheet) > 1:
        st.markdown("---")
        st.header("📦 Bulk Download Options")

        multi_sheet_bytes = create_consolidated_multi_sheet_excel(dfs_for_multi_sheet, reval_mode=True)
        multi_sheet_filename = f"Consolidated_Revaluation_Questionwise_{datetime.now().strftime('%d-%m-%y_%H%M%S')}.xlsx"

        st.download_button(
            label="📊 Download Consolidated Multi-Sheet Excel (Revaluation)",
            data=multi_sheet_bytes,
            file_name=multi_sheet_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="dl_multi_sheet_reval_all"
        )

    # --- Tool Description / Help (collapsed by default) ---
    st.markdown("---")
    render_help_section()


def render_help_section():
    """
    Renders a collapsible 'Tool Description / Help' section at the bottom of
    the page in simple language with full color, border, and symbol details.
    """
    with st.expander("ℹ️ Tool Description / Help - click to expand", expanded=False):
        st.markdown(
            """
### What This Tool Does (In Simple Words)
The **Questionwise Checker** helps exam staff process question-wise mark sheets in two modes —
**Moderation Analysis** and **Revaluation Analysis** — match them with the **VGMR** file (when
provided), predict `Semester Total Max Marks` when VGMR is not uploaded, identify extra
sub-question marks that were not counted in the total, flag genuine score mismatches with
explanatory comments, and generate clean or formatted Excel files.

---
### Analysis Modes
1. **Moderation Analysis (Default)**:
   - Processes **all students** in each uploaded file using the standard moderation workflow.
   - Generates **Cleaned File**, **Formatted File**, **College ZIP**, and **Consolidated Multi-Sheet Excel**.
   - Checks **100% Moderation Eligibility** (>10% change in >50% of moderated scripts) and shows
     the alert banner and Excel `A1` pop-up when eligible.
2. **Revaluation Analysis**:
   - Computes sub-question `best_k` and predicted `Semester Total Max Marks` on the **full paper first**,
     then filters outputs strictly to students who have **Reval 1** and/or **Reval 2** rows (retaining
     their `Examiner` and `Moderator` rows as comparison baselines).
   - Applies the exact same Moderator comparison and formatting rules dynamically:
     - **Moderator** compares against **Examiner**.
     - **Reval 1** compares against **Moderator** (or against **Examiner** if no Moderator row exists).
     - **Reval 2** compares against **Reval 1** (fallback: Moderator → Examiner).
   - Suppresses the 100% Moderation Eligibility alert/pop-up while keeping VGMR `Semester Marks` validation unchanged.
   - Generates **only** the **Formatted File** and **Consolidated Multi-Sheet Excel** (skipping files
     with zero Reval students).

---
### How Processing Works
1. **Upload Questionwise & Optional VGMR Excel Files**: Upload one or more Questionwise Excel
   files (protected or unprotected) and optionally one VGMR Excel file.
2. **Cleaning & Ordering**: Blank rows and `Count:` footer rows are removed. Rows per student are
   sorted as **Examiner → Moderator → Reval 1 → Reval 2**.
3. **Predicted `Semester Total Max Marks` (When VGMR Is Not Provided)**:
   - If VGMR is not uploaded (or `Semester Total Max Marks` is empty), the tool infers a single
     subject-wide `Semester Total Max Marks` from the non-excluded question bracket marks (for
     example, `2 × 5 + 2 × 5 + 2 × 5 = 30`).
   - That predicted value is written identically across all rows of the subject with `!` on both
     sides (e.g., `!30!`) in **bold gray text**.
4. **Sub-Question Grouping, Best-Of Selection & Excluded Extra Marks**:
   - Sub-questions across formats such as `Q1.a`, `Q1.b`, `Q1.c`, `Q.1.A`, `Q2A.`, `Q.1.I`, `Q.1.i`,
     `Q1A`, `Q1.1`, `Q1 1`, `Q1 A`, `Q1a`, `Q1 a`, `Q1 i`, `Q1i` are grouped under their main
     question (`Q1`, `Q2`, `Q3`, …).
   - The tool automatically determines `best_k` counted sub-questions per main question group on
     the full paper, keeps the highest marks, and treats smaller extra attempted marks as **excluded extra marks**.
   - Excluded sub-question cells are shown in curly brackets `{...}` keeping their original number
     format from Excel (for example, `{2.00}` or `{4.00}`).
5. **Half-Mark / Quarter-Mark Rounding (`0.25` / `0.5`) vs. Error Highlighting**:
   - **Round down (`< 0.5`, e.g., `18.25` → `18`)** and **Round up (`0.25` or `0.5`, e.g., `18.75` → `19` or `18.5` → `19`)**
     are treated as valid rounding and shown as `(0.25)` or `(0.5)` in **Excluded Extra Marks**.
   - **Mismatch Error**: If the non-excluded (counted) total does not equal `TotalObtainedScore`
     outside these rounding cases, `Total Obtained Score` is highlighted as an error with an
     explanatory cell comment in Excel.

---
### Complete Guide to Colors, Fonts, Borders & Symbols

#### 1. Error & Mismatch Highlights (Red Fill + Bold Yellow Font + Comment in Excel)
- **Red background (`#FF0000`) + Bold Yellow text (`#FFFF00`)**: Highlights a genuine validation or
  score mismatch on:
  - **Total Obtained Score**: Non-excluded question total does not match `TotalObtainedScore`
    (excluding valid `0.25` / `0.5` rounding cases).
  - **Semester Marks**: VGMR `Semester Marks` does not match that row's `TotalObtainedScore`.
  - **Total Marks Obtained**: Does not equal `Semester Marks + Internal Marks`.
  - **Grade**: Displayed grade does not match the grade derived from percentage marks.

#### 2. Excluded Extra Sub-Question Marks & Rounding Indicators
- **Curly brackets `{...}` + Bold Blue font (`#0000FF`)**: Marks a sub-question whose mark was **not**
  counted in the best-of total (e.g., `{2.00}`), preserving the original number format.
- **Parentheses `(0.25)` or `(0.5)` in Excluded Extra Marks**: Indicates a valid `0.25` or `0.5`
  rounding difference between counted marks and `TotalObtainedScore` (not an error).

#### 3. Predicted `Semester Total Max Marks`
- **Bold Gray font (`#7F7F7F`) with `!` on both sides (e.g., `!30!`)**: Shows the predicted
  `Semester Total Max Marks` when VGMR is not provided.

#### 4. Reviewer vs. Baseline Score Changes
- **Bright Green bold font (`#00B050`)**: Reviewer mark, total, grade, or `Change %` is higher/better
  than its baseline (or reviewer awarded marks where the baseline had `NA`).
- **Red/Orange-Red bold font (`#FF4500`)**: Reviewer mark, total, grade, or `Change %` is lower/worse
  than its baseline, or a mark given by the baseline is missing (`NA` or blank) for the reviewer.

#### 5. Row, Question & SAP ID Colors in Formatted & Consolidated Excel
- **Soft Pastel Green row fill (`#E2EFDA`)**: **Moderator** rows.
- **Soft Pastel Purple row fill (`#E8DCF4`)**: **Reval 1** rows (Revaluation Analysis mode).
- **Soft Pastel Orange row fill (`#FDE4CF`)**: **Reval 2** rows (Revaluation Analysis mode).
- **Tan / Olive-Brown row fill (`#948A54`)**: **Examiner-only** students (no Moderator or Reval row).
- **Yellow question cell fill (`#FFFF00`) + Red SAP ID fill (`#FF0000`)**: Reviewer entered `NA` or
  `0` on a question where the baseline awarded non-zero marks; SAP ID cells of both the baseline
  and reviewer rows for that pair are filled red.
- **SAP ID Color Scale**:
  - **Moderation Analysis**: Colored across the student block based on `Examiner → Moderator` change.
  - **Revaluation Analysis (Row-specific)**: Each stage row (`Examiner`/`Moderator`, `Reval 1`,
    `Reval 2`) is colored by its own transition against its baseline:
    - **Pastel Green shades (`#C3E6C3`, `#8FD18F`, `#5DBB63`)**: Score **increased** (up to 5%, 5%–10%, above 10%).
    - **Pastel Red shades (`#F8CBCB`, `#F19999`, `#E86A6A`)**: Score **decreased** (up to 5%, 5%–10%, above 10%).
    - **Pastel Blue (`#BDD7EE`)**: **No change**.

#### 6. Headers & Borders in Formatted & Consolidated Excel
- **Soft Peach / Light Orange headers (`#FCE4D6`)**: GMR columns, `Change %`, and `Excluded Extra Marks`.
- **Darker Grey header (`#999999`)**: `Total Obtained Score` header cell.
- **Medium horizontal bottom border**: Separates each student's row block.
- **Thick vertical borders**: Separate main question groups (`Q1`, `Q2`, `Q3`, …) and frame `Total Obtained Score`.
            """
        )