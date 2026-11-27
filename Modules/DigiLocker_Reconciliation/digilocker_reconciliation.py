"""
DigiLocker Reconciliation Module — SVKM ExamAutomationSuite.

Single, self-contained module consolidating:
  Section 1: Normalization & ABC ID Validation
  Section 2: String-Safe Excel/CSV Loader & Preflight Inspector
  Section 3: Deduplication, Reconciliation, ABC Recovery, Anomaly & M-Status Engine
  Section 4: NAD-Ready CSV & Styled Multi-Sheet Excel Exporters
  Section 5: Interactive Streamlit UI (show() entry point + standalone fallback)
"""

from __future__ import annotations

import datetime as dt
import hashlib
import io
import re
from dataclasses import dataclass, field
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any, Iterable, Literal

import openpyxl
import pandas as pd
import streamlit as st
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

try:
    from python_calamine import CalamineWorkbook
    HAS_CALAMINE = True
except ImportError:
    HAS_CALAMINE = False

try:
    import xlsxwriter  # noqa: F401
    HAS_XLSXWRITER = True
except ImportError:
    HAS_XLSXWRITER = False


# ==============================================================================
# SECTION 1: Normalization & ABC ID Validation
# ==============================================================================

PROTECTED_ID_COLUMNS: set[str] = {
    "regn_no",
    "abc_account_id",
    "academic_course_id",
    "rroll",
    "cert_no",
    "org_id",
    "org_code",
    "mobile",
}

NAME_COLUMNS: set[str] = {"cname", "aadhaar_name", "fname", "mname"}

NUMERIC_COLUMNS: set[str] = {
    "sgpa",
    "cgpa",
    "tot",
    "tot_mrks",
    "tot_credit",
    "tot_credit_points",
    "grand_tot_credit",
    "grand_tot_credit_points",
    "ncrf_level",
}

DATE_COLUMNS: set[str] = {"dob"}

MONTH_COLUMNS: set[str] = {"month"}

YEAR_COLUMNS: set[str] = {"year", "admission_year"}

MONTH_MAP: dict[str, tuple[int, str]] = {
    "1": (1, "JANUARY"),
    "01": (1, "JANUARY"),
    "jan": (1, "JANUARY"),
    "january": (1, "JANUARY"),
    "2": (2, "FEBRUARY"),
    "02": (2, "FEBRUARY"),
    "feb": (2, "FEBRUARY"),
    "february": (2, "FEBRUARY"),
    "3": (3, "MARCH"),
    "03": (3, "MARCH"),
    "mar": (3, "MARCH"),
    "march": (3, "MARCH"),
    "4": (4, "APRIL"),
    "04": (4, "APRIL"),
    "apr": (4, "APRIL"),
    "april": (4, "APRIL"),
    "5": (5, "MAY"),
    "05": (5, "MAY"),
    "may": (5, "MAY"),
    "6": (6, "JUNE"),
    "06": (6, "JUNE"),
    "jun": (6, "JUNE"),
    "june": (6, "JUNE"),
    "7": (7, "JULY"),
    "07": (7, "JULY"),
    "jul": (7, "JULY"),
    "july": (7, "JULY"),
    "8": (8, "AUGUST"),
    "08": (8, "AUGUST"),
    "aug": (8, "AUGUST"),
    "august": (8, "AUGUST"),
    "9": (9, "SEPTEMBER"),
    "09": (9, "SEPTEMBER"),
    "sep": (9, "SEPTEMBER"),
    "sept": (9, "SEPTEMBER"),
    "september": (9, "SEPTEMBER"),
    "10": (10, "OCTOBER"),
    "oct": (10, "OCTOBER"),
    "october": (10, "OCTOBER"),
    "11": (11, "NOVEMBER"),
    "nov": (11, "NOVEMBER"),
    "november": (11, "NOVEMBER"),
    "12": (12, "DECEMBER"),
    "dec": (12, "DECEMBER"),
    "december": (12, "DECEMBER"),
}

SHORT_MONTH_MAP: dict[int, str] = {
    1: "JAN",
    2: "FEB",
    3: "MAR",
    4: "APR",
    5: "MAY",
    6: "JUN",
    7: "JUL",
    8: "AUG",
    9: "SEP",
    10: "OCT",
    11: "NOV",
    12: "DEC",
}

ROMAN_NUMERALS: dict[int, str] = {
    1: "I",
    2: "II",
    3: "III",
    4: "IV",
    5: "V",
    6: "VI",
    7: "VII",
    8: "VIII",
    9: "IX",
    10: "X",
    11: "XI",
    12: "XII",
}

ROMAN_SEM_ORDER: dict[str, int] = {
    "I": 1,
    "II": 2,
    "III": 3,
    "IV": 4,
    "V": 5,
    "VI": 6,
    "VII": 7,
    "VIII": 8,
    "IX": 9,
    "X": 10,
    "XI": 11,
    "XII": 12,
}

_ABC_MISSING_PLACEHOLDERS: set[str] = {
    "",
    "NA",
    "N/A",
    "N.A",
    "N.A.",
    "NIL",
    "NONE",
    "NULL",
    "-",
    "--",
    "---",
}

_WS_RE = re.compile(r"\s+")
_PUNCT_RE = re.compile(r"[^\w\s]")
_PARENS_OPEN_RE = re.compile(r"\s*\(\s*")
_PARENS_CLOSE_RE = re.compile(r"\s*\)\s*")


def clean_str(val: Any) -> str:
    """Convert raw cell value to stripped string, treating None/NaN as empty string."""
    if val is None:
        return ""
    s = str(val).strip()
    if s.lower() in {"nan", "none", "nat", "<na>"}:
        return ""
    return s


def norm_text(val: Any, case_insensitive: bool = True, strip_punctuation: bool = False) -> str:
    """Trim, optionally remove punctuation, collapse repeated spaces, and uppercase."""
    s = clean_str(val)
    if not s:
        return ""
    if strip_punctuation:
        s = _PUNCT_RE.sub("", s)
    s = _WS_RE.sub(" ", s).strip()
    return s.upper() if case_insensitive else s


def norm_course_name(val: Any) -> str:
    """Normalize course_name casing and parenthesis spacing for matching/comparison only."""
    s = norm_text(val, case_insensitive=True)
    if not s:
        return ""
    s = _PARENS_OPEN_RE.sub(" (", s)
    s = _PARENS_CLOSE_RE.sub(") ", s)
    return _WS_RE.sub(" ", s).strip()


def _natural_sort_key(val: str) -> tuple[tuple[tuple[int, int | str], ...], str]:
    """Sort strings naturally so numeric segments sort numerically (e.g., 2 before 10)."""
    parts = re.split(r"(\d+)", val)
    tokens = tuple((0, int(p)) if p.isdigit() else (1, p) for p in parts if p)
    return (tokens, val)


def parse_month(val: Any) -> tuple[int, str]:
    """Return (month_number, CANONICAL_MONTH_NAME); (0, norm_text) if unknown."""
    s = norm_text(val, case_insensitive=True)
    if not s:
        return (0, "")
    if re.fullmatch(r"\d+\.0+", s):
        s = s.split(".", 1)[0]
    key = s.lower()
    if key in MONTH_MAP:
        return MONTH_MAP[key]
    return (0, s)


def parse_year(val: Any) -> int:
    """Extract integer year from strings like '2024', '2024.0', etc., or 0."""
    s = clean_str(val)
    if not s:
        return 0
    m = re.search(r"(19\d{2}|20\d{2})", s)
    if m:
        return int(m.group(1))
    try:
        return int(Decimal(s))
    except (InvalidOperation, ValueError):
        return 0


def norm_numeric(val: Any) -> str:
    """Normalize numeric string via Decimal so '471.00' == '471' and '4.50' == '4.5'."""
    s = clean_str(val).replace(",", "")
    if not s:
        return ""
    try:
        d = Decimal(s)
        if d == 0:
            return "0"
        normalized = format(d.normalize(), "f")
        if "." in normalized:
            normalized = normalized.rstrip("0").rstrip(".")
        return normalized
    except (InvalidOperation, ValueError):
        return norm_text(s, case_insensitive=True)


def parse_numeric_float(val: Any) -> float:
    """Convert numeric string to float for sorting/ranking; returns -1.0 if unparseable."""
    s = clean_str(val).replace(",", "")
    if not s:
        return -1.0
    try:
        return float(s)
    except ValueError:
        return -1.0


def norm_date(val: Any) -> str:
    """Normalize common date strings, timestamps, or Excel serial numbers to YYYY-MM-DD."""
    s = clean_str(val)
    if not s:
        return ""
    s_date = s.split(" ", 1)[0] if " " in s and ("-" in s or "/" in s) else s
    if "T" in s_date and "-" in s_date:
        s_date = s_date.split("T", 1)[0]
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%d.%m.%Y", "%m/%d/%Y", "%d/%m/%y", "%d-%m-%y"):
        try:
            return dt.datetime.strptime(s_date, fmt).strftime("%Y-%m-%d")
        except ValueError:
            continue
    if re.fullmatch(r"\d{5}(\.0+)?", s_date):
        try:
            serial = int(Decimal(s_date))
            if 20000 <= serial <= 60000:
                base = dt.date(1899, 12, 30)
                return (base + dt.timedelta(days=serial)).strftime("%Y-%m-%d")
        except (InvalidOperation, ValueError):
            pass
    return norm_text(s, case_insensitive=True)


def norm_sem_key(val: Any) -> str:
    """Normalize semester labels ('1', '01', 'Sem I', 'Semester 1') -> Roman 'I'..'XII'."""
    s = norm_text(val, case_insensitive=True)
    if not s:
        return ""
    s = re.sub(r"^(SEMESTER|SEM)\s*[-:]?\s*", "", s).strip()
    if re.fullmatch(r"\d+(\.0+)?", s):
        try:
            num = int(Decimal(s))
            if num in ROMAN_NUMERALS:
                return ROMAN_NUMERALS[num]
        except (InvalidOperation, ValueError):
            pass
    return s


def sem_sort_key(sem: str) -> tuple[int, str]:
    """Sort semesters naturally: I, II, III, ..., XII, then alphabetically."""
    s = norm_sem_key(sem)
    return (ROMAN_SEM_ORDER.get(s, 999), s)


def validate_abc_id(val: Any) -> tuple[str, str]:
    """
    Validate an Academic Bank of Credits (ABC / APAAR) ID.

    Returns (status, cleaned_value) where status is one of:
      - 'VALID': exactly 12 numeric digits, not all identical (leading zeros preserved)
      - 'FIXABLE_FORMAT': 12 digits separated by spaces or hyphens (e.g., '0224-0937-4403');
                          cleaned_value is the stripped 12-digit string
      - 'MISSING': blank, null, or placeholder ('NA', 'N/A', '-')
      - 'DUMMY': '0', '000000000000', or 12 identical digits
      - 'INVALID_FORMAT': any other non-12-digit or non-numeric string
    """
    s = clean_str(val)
    if not s or s.upper() in _ABC_MISSING_PLACEHOLDERS:
        return ("MISSING", "")

    if s == "0" or (s.isdigit() and len(set(s)) == 1 and (set(s) == {"0"} or len(s) == 12)):
        return ("DUMMY", s)

    if len(s) == 12 and s.isdigit() and len(set(s)) > 1:
        return ("VALID", s)

    if re.fullmatch(r"[\d\s-]+", s):
        digits_only = re.sub(r"[\s-]+", "", s)
        if len(digits_only) == 12 and digits_only.isdigit():
            if len(set(digits_only)) == 1:
                return ("DUMMY", s)
            return ("FIXABLE_FORMAT", digits_only)

    return ("INVALID_FORMAT", s)


def normalize_for_comparison(
    col_lower: str,
    val: Any,
    *,
    normalize_text: bool = True,
    normalize_numbers: bool = True,
    normalize_months: bool = True,
    normalize_dates: bool = True,
    ignore_name_punctuation: bool = True,
) -> str:
    """Normalize a cell value for anomaly comparison while protecting ID leading zeros."""
    raw = clean_str(val)
    if not raw:
        return ""

    if col_lower == "abc_account_id":
        abc_status, abc_cleaned = validate_abc_id(raw)
        if abc_status in ("VALID", "FIXABLE_FORMAT"):
            return abc_cleaned
        return norm_text(raw, case_insensitive=True) if normalize_text else raw

    if col_lower in PROTECTED_ID_COLUMNS:
        # Never run numeric normalization on protected IDs (preserves leading zeros)
        return norm_text(raw, case_insensitive=True) if normalize_text else raw

    if col_lower == "sem":
        return norm_sem_key(raw)

    if col_lower == "course_name":
        return norm_course_name(raw) if normalize_text else raw

    if normalize_months and col_lower in MONTH_COLUMNS:
        _, cname = parse_month(raw)
        return cname

    if normalize_dates and col_lower in DATE_COLUMNS:
        return norm_date(raw)

    if normalize_numbers and (col_lower in NUMERIC_COLUMNS or col_lower in YEAR_COLUMNS):
        return norm_numeric(raw)

    if col_lower in NAME_COLUMNS:
        if normalize_text:
            return norm_text(raw, case_insensitive=True, strip_punctuation=ignore_name_punctuation)
        if ignore_name_punctuation:
            return norm_text(raw, case_insensitive=False, strip_punctuation=True)
        return raw

    if normalize_text:
        return norm_text(raw, case_insensitive=True)

    return raw


# ==============================================================================
# SECTION 2: String-Safe Excel/CSV Loader & Preflight Inspector
# ==============================================================================

CSV_ENCODINGS: tuple[str, ...] = ("utf-8-sig", "utf-8", "cp1252", "latin1")
CSV_DELIMITERS: tuple[str, ...] = (",", ";", "\t", "|")

EMAIL_SIGNATURE_COLUMNS: set[str] = {
    "_id",
    "issuer_id",
    "doc_type",
    "csv_filename",
    "created_ts",
    "r_insert_date",
    "org_id",
    "created_date_time",
    "source_created_ts",
    "appar_created_ts",
}

SAP_SIGNATURE_COLUMNS: set[str] = {"sub1", "sub1nm", "sub1_tot"}


@dataclass
class LoadedDataset:
    """Container holding raw DataFrame, header lookup, and source file metadata."""

    df: pd.DataFrame
    col_map: dict[str, str]
    original_columns: list[str]
    source_names: list[str]
    row_counts_per_file: dict[str, int]


def normalize_col_name(name: Any) -> str:
    """Normalize column header for matching across SAP (UPPER) and Email (lower)."""
    return clean_str(name).lower().replace(" ", "_")


def _decimal_to_plain_str(d: Decimal) -> str:
    """Format Decimal cleanly without scientific notation or trailing '.0'."""
    if d == 0:
        return "0"
    s = format(d.normalize(), "f")
    if "." in s:
        s = s.rstrip("0").rstrip(".")
    return s


def _format_cell_as_str(val: Any) -> str:
    """Convert cell value cleanly as text without scientific notation or lost leading zeros."""
    if val is None:
        return ""
    if isinstance(val, bool):
        return "TRUE" if val else "FALSE"
    if isinstance(val, int):
        return str(val)
    if isinstance(val, dt.datetime):
        if val.hour == 0 and val.minute == 0 and val.second == 0:
            return val.strftime("%Y-%m-%d")
        return val.strftime("%Y-%m-%d %H:%M:%S")
    if isinstance(val, dt.date):
        return val.strftime("%Y-%m-%d")
    if isinstance(val, Decimal):
        return "" if val.is_nan() else _decimal_to_plain_str(val)
    if isinstance(val, float):
        if pd.isna(val):
            return ""
        if val.is_integer():
            return str(int(val))
        try:
            return _decimal_to_plain_str(Decimal(str(val)))
        except InvalidOperation:
            return str(val)
    # String cells stay untouched (preserving leading zeros like '022409374403')
    return clean_str(val)


def _dedupe_headers(raw_headers: list[Any]) -> list[str]:
    headers: list[str] = []
    seen: dict[str, int] = {}
    for idx, h in enumerate(raw_headers):
        col = clean_str(h) or f"Unnamed_{idx + 1}"
        if col in seen:
            seen[col] += 1
            col = f"{col}_{seen[col]}"
        else:
            seen[col] = 0
        headers.append(col)
    return headers


def _concat_sheet_frames(frames: list[pd.DataFrame]) -> pd.DataFrame:
    """Concatenate sheet DataFrames preserving column order and skipping empty trailing sheets."""
    non_empty = [f for f in frames if not f.empty or len(f.columns) > 0]
    if not non_empty:
        return pd.DataFrame()
    if len(non_empty) == 1:
        return non_empty[0]

    with_rows = [f for f in non_empty if not f.empty]
    target = with_rows if with_rows else non_empty
    ordered_cols: list[str] = []
    seen: set[str] = set()
    for f in target:
        for c in f.columns:
            if c not in seen:
                seen.add(c)
                ordered_cols.append(c)

    combined = pd.concat(target, ignore_index=True).fillna("")
    return combined[[c for c in ordered_cols if c in combined.columns]]


def get_excel_sheet_names(source: str | Path | bytes, filename: str = "") -> list[str]:
    """Return list of sheet names in an Excel file, or [] for CSV/non-Excel inputs."""
    if filename and filename.lower().endswith(".csv"):
        return []
    if isinstance(source, (str, Path)) and str(source).lower().endswith(".csv"):
        return []

    if HAS_CALAMINE:
        try:
            if isinstance(source, (str, Path)):
                wb = CalamineWorkbook.from_path(str(source))
            else:
                wb = CalamineWorkbook.from_filelike(io.BytesIO(source))
            return list(wb.sheet_names)
        except Exception:
            pass

    try:
        if isinstance(source, (str, Path)):
            wb = openpyxl.load_workbook(source, read_only=True, data_only=True)
        else:
            wb = openpyxl.load_workbook(io.BytesIO(source), read_only=True, data_only=True)
        try:
            return list(wb.sheetnames)
        finally:
            wb.close()
    except Exception:
        return []


def _read_excel_calamine(
    source: str | Path | bytes,
    sheet_name: str | None = None,
    read_all_sheets: bool = False,
) -> pd.DataFrame:
    if isinstance(source, (str, Path)):
        wb = CalamineWorkbook.from_path(str(source))
    else:
        wb = CalamineWorkbook.from_filelike(io.BytesIO(source))

    if not wb.sheet_names:
        return pd.DataFrame()

    if read_all_sheets:
        target_sheets = list(wb.sheet_names)
    else:
        target_sheets = [
            sheet_name if sheet_name and sheet_name in wb.sheet_names else wb.sheet_names[0]
        ]

    sheet_frames: list[pd.DataFrame] = []
    for sname in target_sheets:
        rows = wb.get_sheet_by_name(sname).to_python(skip_empty_area=False)
        if not rows:
            continue
        headers = _dedupe_headers(rows[0])
        n_cols = len(headers)
        data_rows: list[list[str]] = []
        for row in rows[1:]:
            padded = [_format_cell_as_str(row[i]) if i < len(row) else "" for i in range(n_cols)]
            if any(cell != "" for cell in padded):
                data_rows.append(padded)
        sheet_frames.append(pd.DataFrame(data_rows, columns=headers, dtype=str))

    return _concat_sheet_frames(sheet_frames)


def _read_excel_openpyxl(
    source: str | Path | bytes,
    sheet_name: str | None = None,
    read_all_sheets: bool = False,
) -> pd.DataFrame:
    if isinstance(source, (str, Path)):
        wb = openpyxl.load_workbook(source, read_only=True, data_only=True)
    else:
        wb = openpyxl.load_workbook(io.BytesIO(source), read_only=True, data_only=True)

    try:
        if read_all_sheets:
            worksheets = [wb[s] for s in wb.sheetnames]
        elif sheet_name and sheet_name in wb.sheetnames:
            worksheets = [wb[sheet_name]]
        else:
            worksheets = [wb.active]

        sheet_frames: list[pd.DataFrame] = []
        for ws in worksheets:
            it = ws.iter_rows(values_only=True)
            try:
                first_row = next(it)
            except StopIteration:
                continue

            headers = _dedupe_headers(list(first_row))
            n_cols = len(headers)
            data_rows: list[list[str]] = []
            for row in it:
                padded = [_format_cell_as_str(row[i]) if i < len(row) else "" for i in range(n_cols)]
                if any(cell != "" for cell in padded):
                    data_rows.append(padded)
            sheet_frames.append(pd.DataFrame(data_rows, columns=headers, dtype=str))

        return _concat_sheet_frames(sheet_frames)
    finally:
        wb.close()


def _detect_csv_delimiter(text: str) -> str:
    first_line = text.lstrip("\ufeff").split("\n", 1)[0]
    counts = {d: first_line.count(d) for d in CSV_DELIMITERS}
    best = max(counts, key=lambda k: counts[k])
    return best if counts[best] > 0 else ","


def _read_csv_bytes(raw: bytes) -> pd.DataFrame:
    if not raw or not raw.strip():
        raise ValueError("The CSV file is empty.")

    encodings = ("utf-16",) + CSV_ENCODINGS if raw[:2] in (b"\xff\xfe", b"\xfe\xff") else CSV_ENCODINGS
    last_err: Exception | None = None
    for enc in encodings:
        try:
            text = raw.decode(enc)
        except (UnicodeDecodeError, UnicodeError) as exc:
            last_err = exc
            continue
        try:
            sep = _detect_csv_delimiter(text)
            df = pd.read_csv(
                io.StringIO(text),
                sep=sep,
                dtype=str,
                keep_default_na=False,
                na_filter=False,
            )
            return df.fillna("")
        except Exception as exc:
            last_err = exc
    raise ValueError(f"Could not parse CSV: {last_err}")


def read_tabular_file(
    source: str | Path | bytes,
    filename: str,
    sheet_name: str | None = None,
    read_all_sheets: bool = False,
) -> pd.DataFrame:
    """Load an .xlsx/.xlsm/.xls or .csv file into a string-typed DataFrame."""
    lower_name = filename.lower()
    if lower_name.endswith(".csv"):
        if isinstance(source, (str, Path)):
            raw = Path(source).read_bytes()
        else:
            raw = source
        return _read_csv_bytes(raw)

    if HAS_CALAMINE:
        try:
            return _read_excel_calamine(
                source, sheet_name=sheet_name, read_all_sheets=read_all_sheets
            )
        except Exception:
            pass
    return _read_excel_openpyxl(source, sheet_name=sheet_name, read_all_sheets=read_all_sheets)


def load_uploaded_files(
    files: Iterable[tuple[str, str | Path | bytes]],
    sheet_name: str | None = None,
    read_all_sheets: bool = False,
) -> LoadedDataset:
    """Load and concatenate one or more uploaded Excel/CSV files preserving all columns."""
    frames: list[pd.DataFrame] = []
    source_names: list[str] = []
    row_counts: dict[str, int] = {}
    ordered_cols: list[str] = []
    seen_cols: set[str] = set()

    for fname, payload in files:
        df = read_tabular_file(
            payload, filename=fname, sheet_name=sheet_name, read_all_sheets=read_all_sheets
        )
        df = df.copy()
        df["__uploaded_file__"] = fname
        df["__uploaded_row__"] = [str(i + 2) for i in range(len(df))]
        frames.append(df)
        source_names.append(fname)
        row_counts[fname] = len(df)
        for c in df.columns:
            if c not in ("__uploaded_file__", "__uploaded_row__") and c not in seen_cols:
                seen_cols.add(c)
                ordered_cols.append(c)

    if not frames:
        raise ValueError("No files provided.")

    combined = pd.concat(frames, ignore_index=True).fillna("")
    final_cols = ordered_cols + ["__uploaded_file__", "__uploaded_row__"]
    combined = combined[final_cols]

    col_map: dict[str, str] = {}
    for c in ordered_cols:
        norm = normalize_col_name(c)
        if norm not in col_map:
            col_map[norm] = c

    return LoadedDataset(
        df=combined,
        col_map=col_map,
        original_columns=ordered_cols,
        source_names=source_names,
        row_counts_per_file=row_counts,
    )


def detect_file_role(norm_cols: set[str], col_count: int) -> str:
    """Classify a file's likely role ('SAP', 'EMAIL', or 'UNKNOWN') from its schema."""
    has_sap_subs = any(
        c in SAP_SIGNATURE_COLUMNS or re.fullmatch(r"sub\d+(nm|_tot)?", c) for c in norm_cols
    )
    email_sig_hits = len(norm_cols & EMAIL_SIGNATURE_COLUMNS)
    is_sap = has_sap_subs or col_count >= 100
    is_email = email_sig_hits >= 2 and not is_sap
    if is_sap:
        return "SAP"
    if is_email:
        return "EMAIL"
    return "UNKNOWN"


def inspect_uploaded_batch(
    file_entries: tuple[tuple[str, bytes], ...],
    read_all_sheets: bool = False,
) -> dict[str, Any]:
    """Inspect uploaded files for preflight stats, multi-sheet Excel, collisions, and file swaps."""
    total_rows = 0
    max_cols = 0
    all_norm_cols: set[str] = set()
    swap_flags: list[tuple[str, str]] = []
    file_details: list[dict[str, Any]] = []
    multi_sheet_files: list[dict[str, Any]] = []
    courses_by_regn_sem: dict[tuple[str, str], set[str]] = {}

    for fname, raw_bytes in file_entries:
        if not fname.lower().endswith(".csv"):
            sheets = get_excel_sheet_names(raw_bytes, filename=fname)
            if len(sheets) > 1:
                multi_sheet_files.append({"filename": fname, "sheet_names": sheets})

        df = read_tabular_file(raw_bytes, filename=fname, read_all_sheets=read_all_sheets)
        rows_cnt = len(df)
        cols_cnt = len(df.columns)
        col_lookup: dict[str, str] = {}
        for c in df.columns:
            nc = normalize_col_name(c)
            if nc not in col_lookup:
                col_lookup[nc] = c
        norm_cols = set(col_lookup.keys())

        role = detect_file_role(norm_cols, cols_cnt)
        total_rows += rows_cnt
        max_cols = max(max_cols, cols_cnt)
        all_norm_cols |= norm_cols
        swap_flags.append((fname, role))
        file_details.append({"filename": fname, "rows": rows_cnt, "cols": cols_cnt, "detected_role": role})

        if {"regn_no", "sem", "course_name"}.issubset(norm_cols) and not df.empty:
            regn_vals = df[col_lookup["regn_no"]].tolist()
            sem_vals = df[col_lookup["sem"]].tolist()
            course_vals = df[col_lookup["course_name"]].tolist()
            for r_raw, s_raw, c_raw in zip(regn_vals, sem_vals, course_vals):
                r_norm = norm_text(r_raw, case_insensitive=True)
                s_norm = norm_sem_key(s_raw)
                c_norm = norm_course_name(c_raw)
                if r_norm and s_norm and c_norm:
                    courses_by_regn_sem.setdefault((r_norm, s_norm), set()).add(c_norm)

    collision_items = [
        {"regn_no": r, "sem": s, "courses": sorted(cs)}
        for (r, s), cs in sorted(courses_by_regn_sem.items())
        if len(cs) >= 2
    ]

    return {
        "file_count": len(file_entries),
        "total_rows": total_rows,
        "max_cols": max_cols,
        "has_regn_no": "regn_no" in all_norm_cols,
        "has_sem": "sem" in all_norm_cols,
        "file_roles": swap_flags,
        "file_details": file_details,
        "multi_sheet_files": multi_sheet_files,
        "course_collision_count": len(collision_items),
        "course_collision_samples": collision_items[:5],
    }


@st.cache_data(show_spinner=False, max_entries=16)
def _cached_preflight_inspect(
    file_entries: tuple[tuple[str, bytes], ...],
    read_all_sheets: bool = False,
) -> dict[str, Any]:
    return inspect_uploaded_batch(file_entries, read_all_sheets=read_all_sheets)


# ==============================================================================
# SECTION 3: Deduplication, Reconciliation, ABC Recovery, Anomaly & M-Status Engine
# ==============================================================================

DEFAULT_METADATA_COLUMNS: set[str] = {
    "source_file",
    "csv_filename",
    "created_date_time",
    "created_ts",
    "insert_date",
    "r_insert_date",
    "r_insert_time",
    "r_insert_date_time",
    "source_insert_date",
    "source_created_ts",
    "appar_created_ts",
    "_id",
    "issuer_id",
    "doc_type",
    "org_id",
    "org_code",
    "doi",
    "dor",
    "dov",
    "part",
}

HIGH_SIGNAL_COLUMNS: list[str] = [
    "cname",
    "dob",
    "gender",
    "abc_account_id",
    "rroll",
    "academic_course_id",
    "result",
    "grade",
    "tot_mrks",
    "sgpa",
    "cgpa",
    "tot_credit",
    "tot_credit_points",
    "grand_tot_credit_points",
    "year",
    "month",
]

CRITICAL_COLUMNS: set[str] = {
    "result",
    "dob",
    "gender",
    "cname",
    "aadhaar_name",
    "academic_course_id",
    "abc_account_id",
    "regn_no",
    "rroll",
    "sem",
}

ACADEMIC_COLUMNS: set[str] = {
    "tot",
    "tot_mrks",
    "sgpa",
    "cgpa",
    "grade",
    "tot_credit",
    "tot_credit_points",
    "grand_tot_credit_points",
}

PASS_CORRECTION_CRITICAL_COLUMNS: set[str] = {
    "tot_mrks",
    "sgpa",
    "cgpa",
    "grade",
    "tot_credit",
    "tot_credit_points",
    "grand_tot_credit_points",
    "dob",
    "gender",
    "cname",
    "abc_account_id",
    "academic_course_id",
}

COLUMN_PRESETS: dict[str, list[str]] = {
    "High-Signal Academic & Identity (Recommended)": HIGH_SIGNAL_COLUMNS,
    "Core Student Identity": [
        "cname",
        "dob",
        "gender",
        "abc_account_id",
        "rroll",
        "academic_course_id",
        "course_name",
        "admission_year",
    ],
    "Academic Results & Credits": [
        "result",
        "grade",
        "tot_mrks",
        "tot",
        "sgpa",
        "cgpa",
        "tot_credit",
        "tot_credit_points",
        "grand_tot_credit_points",
    ],
    "Exam Cycle & Status": [
        "year",
        "month",
        "session",
        "exam_type",
        "mrks_rec_status",
        "ncrf_level",
    ],
}


@dataclass
class ReconciliationConfig:
    """Configuration controlling key matching, normalization, and anomaly reporting."""

    key_columns: list[str] = field(default_factory=lambda: ["regn_no", "sem"])
    compare_columns: list[str] | None = field(default_factory=lambda: list(HIGH_SIGNAL_COLUMNS))
    normalize_text: bool = True
    normalize_numbers: bool = True
    normalize_months: bool = True
    normalize_dates: bool = True
    ignore_name_punctuation: bool = True
    include_one_side_null_anomalies: bool = False
    include_format_only_anomalies: bool = False


@dataclass
class ReconciliationResult:
    """Complete output of SAP vs Email Data reconciliation."""

    config: ReconciliationConfig
    sap_original_columns: list[str]
    sap_col_map: dict[str, str]
    email_col_map: dict[str, str]
    common_columns: list[str]
    compared_columns: list[str]

    pure_pending_sap_df: pd.DataFrame
    abc_backfill_audit_df: pd.DataFrame
    verified_abc_map: dict[str, tuple[str, list[str]]]
    fail_to_pass_upgrades_df: pd.DataFrame
    fail_to_pass_raw_sap_df: pd.DataFrame
    pass_corrections_df: pd.DataFrame
    pass_corrections_raw_sap_df: pd.DataFrame
    abc_only_updates_df: pd.DataFrame
    abc_only_updates_raw_sap_df: pd.DataFrame
    sap_conflicts_df: pd.DataFrame
    email_conflicts_df: pd.DataFrame
    anomalies_long_df: pd.DataFrame
    anomalies_wide_df: pd.DataFrame
    format_only_df: pd.DataFrame
    column_anomaly_summary_df: pd.DataFrame
    email_only_df: pd.DataFrame
    student_report_df: pd.DataFrame
    student_pivot_df: pd.DataFrame
    student_pivot_cell_status: dict[tuple[str, str], str]
    student_timeline_map: dict[str, list[dict[str, str]]]
    semester_columns: list[str]

    metrics: dict[str, int]


def _get_col_series(dataset: LoadedDataset, canonical_col: str) -> pd.Series:
    orig = dataset.col_map.get(canonical_col)
    if orig and orig in dataset.df.columns:
        return dataset.df[orig]
    return pd.Series([""] * len(dataset.df), index=dataset.df.index, dtype=str)


def _normalize_key_series(dataset: LoadedDataset, col_name: str) -> pd.Series:
    """Normalize a matching-key column uniformly across composite keys and deduplication."""
    col_lower = normalize_col_name(col_name)
    s = _get_col_series(dataset, col_lower)
    if col_lower == "sem":
        return s.map(norm_sem_key)
    if col_lower == "month":
        return s.map(lambda v: parse_month(v)[1])
    if col_lower in ("year", "admission_year"):
        return s.map(norm_numeric)
    if col_lower == "dob":
        return s.map(norm_date)
    if col_lower == "course_name":
        return s.map(norm_course_name)
    return s.map(lambda v: norm_text(v, case_insensitive=True))


def _build_composite_key(dataset: LoadedDataset, key_columns: list[str]) -> pd.Series:
    parts = [_normalize_key_series(dataset, k) for k in key_columns]
    if len(parts) == 1:
        return parts[0]
    combined = parts[0]
    for p in parts[1:]:
        combined = combined + "||" + p
    return combined


def _result_rank(val: str) -> int:
    u = norm_text(val, case_insensitive=True)
    if u == "PASS":
        return 2
    if u == "FAIL":
        return 0
    return 1


def _mrks_status_rank(val: str) -> int:
    u = norm_text(val, case_insensitive=True)
    if u == "M":
        return 2
    if u == "O":
        return 1
    return 0


def _deduplicate_dataset(
    dataset: LoadedDataset,
    key_columns: list[str],
    source_label: str,
) -> tuple[pd.DataFrame, pd.DataFrame, dict[str, str]]:
    """
    Deduplicate dataset on key_columns.

    Priority order:
      1. PASS > other > FAIL
      2. Latest (year, month)
      3. On identical period: mrks_rec_status = M > O > blank
      4. Higher tot_mrks -> higher sgpa -> more populated fields -> earlier row
    """
    df = dataset.df.copy()
    df["__key__"] = _build_composite_key(dataset, key_columns)

    for k in key_columns:
        df[f"__k_{k}__"] = _normalize_key_series(dataset, k)

    regn_s = _normalize_key_series(dataset, "regn_no")
    sem_s = _normalize_key_series(dataset, "sem")
    cname_s = _get_col_series(dataset, "cname").map(lambda v: norm_text(v, case_insensitive=True))
    result_s = _get_col_series(dataset, "result").map(lambda v: norm_text(v, case_insensitive=True))
    year_raw_s = _get_col_series(dataset, "year")
    month_raw_s = _get_col_series(dataset, "month")
    tot_mrks_s = _get_col_series(dataset, "tot_mrks")
    sgpa_s = _get_col_series(dataset, "sgpa")
    mrks_status_s = _get_col_series(dataset, "mrks_rec_status")

    df["__regn_norm__"] = regn_s
    df["__sem_norm__"] = sem_s
    df["__cname_norm__"] = cname_s
    df["__result_norm__"] = result_s
    df["__year_int__"] = year_raw_s.map(parse_year)
    month_parsed = month_raw_s.map(parse_month)
    df["__month_int__"] = month_parsed.map(lambda x: x[0])
    df["__month_norm__"] = month_parsed.map(lambda x: x[1])
    df["__mrks_status_rank__"] = mrks_status_s.map(_mrks_status_rank)
    df["__mrks_float__"] = tot_mrks_s.map(parse_numeric_float)
    df["__sgpa_float__"] = sgpa_s.map(parse_numeric_float)
    df["__result_rank__"] = result_s.map(_result_rank)
    df["__orig_order__"] = range(len(df))

    orig_cols = dataset.original_columns
    df["__non_empty__"] = (df[orig_cols] != "").sum(axis=1)

    counts = df["__key__"].value_counts()
    dup_keys = set(counts[counts > 1].index)

    if not dup_keys:
        empty_conflicts = pd.DataFrame(
            columns=[
                "source",
                "regn_no",
                "sem",
                "cname",
                "conflict_type",
                "action",
                "result",
                "year",
                "month",
                "tot_mrks",
                "sgpa",
                "mrks_rec_status",
                "uploaded_file",
                "excel_row",
            ]
        )
        return df.drop(columns=["__orig_order__", "__non_empty__"]), empty_conflicts, {}

    df_sorted = df.sort_values(
        by=[
            "__key__",
            "__result_rank__",
            "__year_int__",
            "__month_int__",
            "__mrks_status_rank__",
            "__mrks_float__",
            "__sgpa_float__",
            "__non_empty__",
            "__orig_order__",
        ],
        ascending=[True, False, False, False, False, False, False, False, True],
        kind="mergesort",
    )

    dedup_df = df_sorted.drop_duplicates(subset=["__key__"], keep="first").copy()
    dedup_df = dedup_df.sort_values("__orig_order__").reset_index(drop=True)

    dup_rows = df_sorted[df_sorted["__key__"].isin(dup_keys)]
    mrks_status_col = dataset.col_map.get("mrks_rec_status", "")
    tot_mrks_col = dataset.col_map.get("tot_mrks", "")
    sgpa_col = dataset.col_map.get("sgpa", "")
    n_dup = len(dup_rows)

    grouped_dup_rows: dict[str, list[tuple[Any, ...]]] = {}
    for item in zip(
        dup_rows["__key__"].tolist(),
        dup_rows["__result_norm__"].tolist(),
        dup_rows["__year_int__"].tolist(),
        dup_rows["__month_norm__"].tolist(),
        dup_rows["__mrks_float__"].tolist(),
        dup_rows["__sgpa_float__"].tolist(),
        dup_rows["__mrks_status_rank__"].tolist(),
        dup_rows["__regn_norm__"].tolist(),
        dup_rows["__sem_norm__"].tolist(),
        dup_rows["__cname_norm__"].tolist(),
        dup_rows[tot_mrks_col].tolist() if tot_mrks_col else [""] * n_dup,
        dup_rows[sgpa_col].tolist() if sgpa_col else [""] * n_dup,
        dup_rows[mrks_status_col].tolist() if mrks_status_col else [""] * n_dup,
        dup_rows["__uploaded_file__"].tolist(),
        dup_rows["__uploaded_row__"].tolist(),
    ):
        grouped_dup_rows.setdefault(item[0], []).append(item[1:])

    conflict_records: list[dict[str, str]] = []
    conflict_type_by_key: dict[str, str] = {}

    for key_val, items in grouped_dup_rows.items():
        pass_count = sum(1 for it in items if it[0] == "PASS")
        fail_count = sum(1 for it in items if it[0] == "FAIL")
        has_pass = pass_count > 0
        has_fail = fail_count > 0

        sigs = {(it[0], it[1], it[2], it[3], it[4], it[5]) for it in items}
        if len(sigs) == 1:
            ctype = "EXACT_OR_NEAR_DUPLICATE"
        elif has_pass and has_fail and pass_count == 1:
            ctype = "FAIL_AND_PASS"
        elif has_pass and has_fail and pass_count > 1:
            ctype = "MULTIPLE_PASS_AND_FAIL"
        elif pass_count > 1:
            ym_pairs = {(it[1], it[2]) for it in items if it[0] == "PASS"}
            ctype = "MULTIPLE_PASS_DIFF_MONTH_YEAR" if len(ym_pairs) > 1 else "MULTIPLE_PASS_SAME_PERIOD"
        elif fail_count > 1:
            ctype = "MULTIPLE_FAIL"
        else:
            ctype = "DUPLICATE_OTHER"

        conflict_type_by_key[key_val] = ctype

        for idx_in_grp, (
            r_norm,
            yr_int,
            mo_norm,
            _mrks_flt,
            _sgpa_flt,
            _mstat_rank,
            regn_val,
            sem_val,
            cname_val,
            tot_raw,
            sgpa_raw,
            mstat_raw,
            ufile,
            urow,
        ) in enumerate(items):
            is_winner = idx_in_grp == 0
            if is_winner:
                if ctype == "EXACT_OR_NEAR_DUPLICATE":
                    action = "KEPT (Canonical Copy)"
                elif ctype in ("FAIL_AND_PASS", "MULTIPLE_PASS_AND_FAIL"):
                    action = "KEPT (Latest PASS)"
                elif ctype.startswith("MULTIPLE_PASS"):
                    action = "KEPT (Latest Year/Month/M-Status PASS)"
                else:
                    action = "KEPT (Latest Record)"
            else:
                if ctype == "EXACT_OR_NEAR_DUPLICATE":
                    action = "REMOVED (Duplicate Row)"
                elif r_norm == "FAIL" and has_pass:
                    action = "SUPERSEDED (By PASS Record)"
                elif has_pass:
                    action = "SUPERSEDED (Older/Lower PASS)"
                else:
                    action = "SUPERSEDED (Older Record)"

            conflict_records.append(
                {
                    "source": source_label,
                    "regn_no": regn_val,
                    "sem": sem_val,
                    "cname": cname_val,
                    "conflict_type": ctype,
                    "action": action,
                    "result": r_norm,
                    "year": str(yr_int or ""),
                    "month": mo_norm,
                    "tot_mrks": clean_str(tot_raw),
                    "sgpa": clean_str(sgpa_raw),
                    "mrks_rec_status": clean_str(mstat_raw),
                    "uploaded_file": clean_str(ufile),
                    "excel_row": clean_str(urow),
                }
            )

    conflicts_df = pd.DataFrame(conflict_records)
    return dedup_df.drop(columns=["__orig_order__", "__non_empty__"]), conflicts_df, conflict_type_by_key


def _classify_column_severity(col_lower: str) -> str:
    if col_lower in CRITICAL_COLUMNS:
        return "CRITICAL"
    if col_lower in ACADEMIC_COLUMNS:
        return "ACADEMIC"
    return "METADATA"


def _build_verified_abc_map(
    sap_dataset: LoadedDataset,
    email_dataset: LoadedDataset,
) -> tuple[dict[str, tuple[str, list[str]]], set[str]]:
    """
    Collect unambiguous valid 12-digit ABC IDs per regn_no across all SAP & Email rows.

    Returns:
      verified_abc_map: regn_no -> (valid_12digit_abc_id, sorted list of source tags)
      ambiguous_regns: set of regn_no having >= 2 conflicting valid ABC IDs across records
    """
    candidates: dict[str, dict[str, set[str]]] = {}

    for src_name, ds in (("SAP", sap_dataset), ("Email", email_dataset)):
        abc_col = ds.col_map.get("abc_account_id")
        regn_col = ds.col_map.get("regn_no")
        sem_col = ds.col_map.get("sem")
        if not abc_col or not regn_col:
            continue

        regn_vals = ds.df[regn_col].tolist()
        abc_vals = ds.df[abc_col].tolist()
        sem_vals = ds.df[sem_col].tolist() if sem_col else [""] * len(ds.df)

        for r_raw, a_raw, s_raw in zip(regn_vals, abc_vals, sem_vals):
            regn = norm_text(r_raw, case_insensitive=True)
            if not regn:
                continue
            status, cleaned_abc = validate_abc_id(a_raw)
            if status in ("VALID", "FIXABLE_FORMAT"):
                sem_norm = norm_sem_key(s_raw) or "?"
                src_label = f"{src_name} Sem {sem_norm}"
                candidates.setdefault(regn, {}).setdefault(cleaned_abc, set()).add(src_label)

    verified_map: dict[str, tuple[str, list[str]]] = {}
    ambiguous_regns: set[str] = set()
    for regn, abc_dict in candidates.items():
        if len(abc_dict) == 1:
            only_abc, sources = next(iter(abc_dict.items()))
            verified_map[regn] = (only_abc, sorted(sources))
        elif len(abc_dict) > 1:
            ambiguous_regns.add(regn)

    return verified_map, ambiguous_regns


def _short_month_label(month_int: int, month_norm: str = "") -> str:
    if month_int in SHORT_MONTH_MAP:
        return SHORT_MONTH_MAP[month_int]
    if month_norm:
        return month_norm[:3].upper()
    return ""


def _format_period_short(year_int: int, month_int: int, month_short: str) -> str:
    if month_short and year_int > 0:
        return f"{month_short} {year_int}"
    if month_short:
        return month_short
    if year_int > 0:
        return str(year_int)
    return ""


def _collect_exam_periods(
    dataset: LoadedDataset,
) -> dict[tuple[str, str], set[tuple[int, int, str]]]:
    """Collect distinct (year_int, month_int, month_short) exam attempts per (regn_no, sem)."""
    periods: dict[tuple[str, str], set[tuple[int, int, str]]] = {}
    regns = _normalize_key_series(dataset, "regn_no").tolist()
    sems = _normalize_key_series(dataset, "sem").tolist()
    yrs = _get_col_series(dataset, "year").tolist()
    mos = _get_col_series(dataset, "month").tolist()

    for regn, sem, y_raw, m_raw in zip(regns, sems, yrs, mos):
        if not regn or not sem:
            continue
        yr_int = parse_year(y_raw)
        mo_int, mo_norm = parse_month(m_raw)
        mo_short = _short_month_label(mo_int, mo_norm)
        if yr_int > 0 or mo_int > 0 or mo_short:
            periods.setdefault((regn, sem), set()).add((yr_int, mo_int, mo_short))
    return periods


def _format_attempt_cell(
    canon_period: tuple[int, int, str],
    all_periods: set[tuple[int, int, str]],
    fallback_label: str = "",
) -> str:
    """Format canonical attempt first, with distinct earlier attempts newest-to-oldest in brackets."""
    canon_str = _format_period_short(*canon_period)
    other_periods = sorted(
        (p for p in all_periods if p != canon_period and _format_period_short(*p) != ""),
        key=lambda p: (p[0], p[1], p[2]),
        reverse=True,
    )
    other_strs: list[str] = []
    seen_labels: set[str] = {canon_str} if canon_str else set()
    for p in other_periods:
        lbl = _format_period_short(*p)
        if lbl and lbl not in seen_labels:
            seen_labels.add(lbl)
            other_strs.append(lbl)

    if canon_str and other_strs:
        return f"{canon_str} ({', '.join(other_strs)})"
    if canon_str:
        return canon_str
    if other_strs:
        if len(other_strs) == 1:
            return other_strs[0]
        return f"{other_strs[0]} ({', '.join(other_strs[1:])})"
    return fallback_label or "PRESENT"


def _ingest_dataset_for_student_map(
    clean_df: pd.DataFrame,
    ds: LoadedDataset,
    by_student_sem_primary: dict[str, dict[str, dict[str, str]]],
    by_student_sem_records: dict[tuple[str, str], list[dict[str, str]]],
    student_profile: dict[str, dict[str, str]],
    sap_conflict_by_key: dict[str, str] | None = None,
) -> None:
    """Populate student profile and per-(student, sem) snapshots using columnar zip iteration."""
    n = len(clean_df)
    if n == 0:
        return

    def _col_vals(canon_col: str) -> list[str]:
        orig = ds.col_map.get(canon_col)
        if orig and orig in clean_df.columns:
            return clean_df[orig].tolist()
        return [""] * n

    keys = clean_df["__key__"].tolist()
    regns = clean_df["__regn_norm__"].tolist()
    sems = clean_df["__sem_norm__"].tolist()
    results = clean_df["__result_norm__"].tolist()
    yr_ints = clean_df["__year_int__"].tolist()
    mo_ints = clean_df["__month_int__"].tolist()
    mo_norms = clean_df["__month_norm__"].tolist()

    cnames = _col_vals("cname")
    courses = _col_vals("course_name")
    course_ids = _col_vals("academic_course_id")
    adm_years = _col_vals("admission_year")
    abc_ids = _col_vals("abc_account_id")
    dobs = _col_vals("dob")
    genders = _col_vals("gender")
    tot_mrks_vals = _col_vals("tot_mrks")
    sgpa_vals = _col_vals("sgpa")
    credit_pts_vals = _col_vals("tot_credit_points")
    mrks_status_vals = _col_vals("mrks_rec_status")

    conflict_lookup = sap_conflict_by_key or {}

    for (
        k,
        regn,
        sem,
        res_v,
        yr_int,
        mo_int,
        mo_norm,
        cname_raw,
        course_raw,
        course_id_raw,
        adm_raw,
        abc_raw,
        dob_raw,
        gender_raw,
        tot_m_raw,
        sgpa_raw,
        cred_raw,
        mstat_raw,
    ) in zip(
        keys,
        regns,
        sems,
        results,
        yr_ints,
        mo_ints,
        mo_norms,
        cnames,
        courses,
        course_ids,
        adm_years,
        abc_ids,
        dobs,
        genders,
        tot_mrks_vals,
        sgpa_vals,
        credit_pts_vals,
        mrks_status_vals,
    ):
        prof = student_profile.setdefault(
            regn,
            {
                "regn_no": regn,
                "cname": "",
                "course_name": "",
                "academic_course_id": "",
                "admission_year": "",
                "abc_account_id": "",
                "dob": "",
                "gender": "",
            },
        )
        if not prof["cname"]:
            v = norm_text(cname_raw, case_insensitive=True)
            if v:
                prof["cname"] = v
        for fld, raw_v in (
            ("course_name", course_raw),
            ("academic_course_id", course_id_raw),
            ("admission_year", adm_raw),
            ("dob", dob_raw),
            ("gender", gender_raw),
        ):
            if not prof[fld]:
                v = clean_str(raw_v)
                if v:
                    prof[fld] = v

        abc_clean = clean_str(abc_raw)
        if abc_clean:
            cur_abc = prof["abc_account_id"]
            if not cur_abc:
                prof["abc_account_id"] = abc_clean
            else:
                cur_status, _ = validate_abc_id(cur_abc)
                new_status, new_cleaned = validate_abc_id(abc_clean)
                if cur_status not in ("VALID", "FIXABLE_FORMAT") and new_status in (
                    "VALID",
                    "FIXABLE_FORMAT",
                ):
                    prof["abc_account_id"] = new_cleaned

        yr_str = str(yr_int or "")
        mo_short = _short_month_label(int(mo_int), mo_norm)
        tot_m = clean_str(tot_m_raw)
        sgpa_v = clean_str(sgpa_raw)

        snap = {
            "key": k,
            "result": res_v,
            "year_int": str(int(yr_int)),
            "month_int": str(int(mo_int)),
            "month_short": mo_short,
            "period": f"{yr_str} {mo_norm}".strip(),
            "marks_sgpa": f"{tot_m or '-'} / {sgpa_v or '-'}",
            "credit_pts": clean_str(cred_raw),
            "abc_id": abc_clean,
            "course_id": clean_str(course_id_raw),
            "mrks_status": clean_str(mstat_raw),
            "conflict": conflict_lookup.get(k, ""),
        }
        by_student_sem_primary.setdefault(regn, {})[sem] = snap
        by_student_sem_records.setdefault((regn, sem), []).append(snap)


def run_reconciliation(
    sap_dataset: LoadedDataset,
    email_dataset: LoadedDataset,
    config: ReconciliationConfig | None = None,
) -> ReconciliationResult:
    """Execute deduplication, DigiLocker reconciliation, ABC recovery, anomaly, M-status & 360° report."""
    cfg = config or ReconciliationConfig()

    missing_sap_keys = [k for k in cfg.key_columns if k not in sap_dataset.col_map]
    missing_email_keys = [k for k in cfg.key_columns if k not in email_dataset.col_map]
    if missing_sap_keys or missing_email_keys:
        raise ValueError(
            f"Missing key column(s) — SAP missing: {missing_sap_keys}, Email missing: {missing_email_keys}"
        )

    common_cols: list[str] = []
    for orig_c in sap_dataset.original_columns:
        norm_c = next((k for k, v in sap_dataset.col_map.items() if v == orig_c), None)
        if norm_c and norm_c in email_dataset.col_map and norm_c not in common_cols:
            common_cols.append(norm_c)

    if cfg.compare_columns is not None:
        compared_cols = [c for c in cfg.compare_columns if c in common_cols and c not in cfg.key_columns]
    else:
        compared_cols = [
            c for c in common_cols if c not in DEFAULT_METADATA_COLUMNS and c not in cfg.key_columns
        ]

    # Collect raw exam attempt periods before deduplication discards superseded rows
    sap_raw_periods = _collect_exam_periods(sap_dataset)
    email_raw_periods = _collect_exam_periods(email_dataset)

    # Step 1: Deduplicate SAP and Email internally
    sap_clean, sap_conflicts_df, sap_conflict_by_key = _deduplicate_dataset(
        sap_dataset, cfg.key_columns, source_label="SAP"
    )
    email_clean, email_conflicts_df, _ = _deduplicate_dataset(
        email_dataset, cfg.key_columns, source_label="Email"
    )

    sap_keys_set = set(sap_clean["__key__"])
    email_keys_set = set(email_clean["__key__"])

    # Step 2: Build Verified ABC Recovery Map
    verified_abc_map, ambiguous_abc_regns = _build_verified_abc_map(sap_dataset, email_dataset)

    # Step 3: Pure Pending SAP Data + ABC Readiness Classification
    pending_mask = ~sap_clean["__key__"].isin(email_keys_set)
    pending_sap = sap_clean[pending_mask].copy()
    pending_sap["__pending_result__"] = pending_sap["__result_norm__"]
    pending_sap["__had_sap_conflict__"] = pending_sap["__key__"].isin(sap_conflict_by_key)
    pending_sap["__sap_conflict_type__"] = pending_sap["__key__"].map(
        lambda k: sap_conflict_by_key.get(k, "")
    )

    sap_abc_col = sap_dataset.col_map.get("abc_account_id", "")
    sap_course_col = sap_dataset.col_map.get("course_name", "")

    abc_readiness_list: list[str] = []
    abc_orig_status_list: list[str] = []
    recovered_abc_list: list[str] = []
    recovery_source_list: list[str] = []
    abc_audit_rows: list[dict[str, str]] = []

    p_regns = pending_sap["__regn_norm__"].tolist()
    p_sems = pending_sap["__sem_norm__"].tolist()
    p_cnames = pending_sap["__cname_norm__"].tolist()
    p_results = pending_sap["__result_norm__"].tolist()
    p_abcs = pending_sap[sap_abc_col].tolist() if sap_abc_col else [""] * len(pending_sap)
    p_courses = pending_sap[sap_course_col].tolist() if sap_course_col else [""] * len(pending_sap)

    for regn, sem, cname_val, res_val, orig_abc_raw, course_raw in zip(
        p_regns, p_sems, p_cnames, p_results, p_abcs, p_courses
    ):
        orig_abc = clean_str(orig_abc_raw)
        course_val = clean_str(course_raw)
        abc_status, cleaned_abc = validate_abc_id(orig_abc)
        abc_orig_status_list.append(abc_status)

        if abc_status == "VALID":
            abc_readiness_list.append("READY")
            recovered_abc_list.append("")
            recovery_source_list.append("")
        elif abc_status == "FIXABLE_FORMAT":
            rec_src_str = "Format cleaned (removed spaces/hyphens)"
            abc_readiness_list.append("RECOVERABLE")
            recovered_abc_list.append(cleaned_abc)
            recovery_source_list.append(rec_src_str)
            abc_audit_rows.append(
                {
                    "regn_no": regn,
                    "sem": sem,
                    "cname": cname_val,
                    "course_name": course_val,
                    "result": res_val,
                    "original_abc_id": orig_abc,
                    "issue_type": abc_status,
                    "ABC_READINESS": "RECOVERABLE",
                    "recovered_abc_id": cleaned_abc,
                    "recovery_source": rec_src_str,
                }
            )
        elif regn in verified_abc_map:
            rec_id, rec_srcs = verified_abc_map[regn]
            rec_src_str = ", ".join(rec_srcs)
            abc_readiness_list.append("RECOVERABLE")
            recovered_abc_list.append(rec_id)
            recovery_source_list.append(rec_src_str)
            abc_audit_rows.append(
                {
                    "regn_no": regn,
                    "sem": sem,
                    "cname": cname_val,
                    "course_name": course_val,
                    "result": res_val,
                    "original_abc_id": orig_abc,
                    "issue_type": abc_status,
                    "ABC_READINESS": "RECOVERABLE",
                    "recovered_abc_id": rec_id,
                    "recovery_source": rec_src_str,
                }
            )
        else:
            readiness = "MISSING" if abc_status == "MISSING" else "INVALID_OR_DUMMY"
            note = (
                "Ambiguous: multiple distinct valid ABC IDs exist for student"
                if regn in ambiguous_abc_regns
                else "No valid 12-digit ABC ID found across SAP or Email records"
            )
            abc_readiness_list.append(readiness)
            recovered_abc_list.append("")
            recovery_source_list.append(note)
            abc_audit_rows.append(
                {
                    "regn_no": regn,
                    "sem": sem,
                    "cname": cname_val,
                    "course_name": course_val,
                    "result": res_val,
                    "original_abc_id": orig_abc,
                    "issue_type": abc_status,
                    "ABC_READINESS": readiness,
                    "recovered_abc_id": "",
                    "recovery_source": note,
                }
            )

    pending_sap["__abc_orig_status__"] = abc_orig_status_list
    pending_sap["__abc_readiness__"] = abc_readiness_list
    pending_sap["__recovered_abc_id__"] = recovered_abc_list
    pending_sap["__recovery_source__"] = recovery_source_list

    helper_cols = [
        "__pending_result__",
        "__had_sap_conflict__",
        "__sap_conflict_type__",
        "__abc_orig_status__",
        "__abc_readiness__",
        "__recovered_abc_id__",
        "__recovery_source__",
        "__uploaded_file__",
        "__uploaded_row__",
    ]
    pure_pending_sap_df = pending_sap[sap_dataset.original_columns + helper_cols].reset_index(drop=True)
    abc_backfill_audit_df = (
        pd.DataFrame(abc_audit_rows)
        if abc_audit_rows
        else pd.DataFrame(
            columns=[
                "regn_no",
                "sem",
                "cname",
                "course_name",
                "result",
                "original_abc_id",
                "issue_type",
                "ABC_READINESS",
                "recovered_abc_id",
                "recovery_source",
            ]
        )
    )

    # Step 4: Enriched Email-Only Records (showing whether student exists in SAP for other sems)
    sap_sems_sets_by_regn: dict[str, set[str]] = {}
    for regn, sem_val in zip(
        sap_clean["__regn_norm__"].tolist(), sap_clean["__sem_norm__"].tolist()
    ):
        if sem_val:
            sap_sems_sets_by_regn.setdefault(regn, set()).add(sem_val)
    sap_sems_by_regn: dict[str, list[str]] = {
        regn: sorted(sems, key=sem_sort_key) for regn, sems in sap_sems_sets_by_regn.items()
    }

    email_only_mask = ~email_clean["__key__"].isin(sap_keys_set)
    email_only_raw = email_clean[email_only_mask]
    email_only_cols = [
        c
        for c in [
            "regn_no",
            "sem",
            "cname",
            "course_name",
            "academic_course_id",
            "result",
            "year",
            "month",
            "tot_mrks",
            "sgpa",
            "abc_account_id",
            "mrks_rec_status",
            "source_file",
            "csv_filename",
        ]
        if c in email_dataset.col_map
    ]
    email_only_records: dict[str, Any] = {
        c: email_only_raw[email_dataset.col_map[c]].tolist() for c in email_only_cols
    }
    eo_regns = email_only_raw["__regn_norm__"].tolist()
    email_only_records["exists_in_sap_other_sems"] = [
        "Yes" if sap_sems_by_regn.get(r) else "No" for r in eo_regns
    ]
    email_only_records["sap_other_sems"] = [", ".join(sap_sems_by_regn.get(r, [])) for r in eo_regns]
    email_only_records["uploaded_file"] = email_only_raw["__uploaded_file__"].tolist()
    email_only_records["excel_row"] = email_only_raw["__uploaded_row__"].tolist()
    email_only_df = pd.DataFrame(email_only_records).reset_index(drop=True)

    # Step 5: Cross-File Matching, FAIL->PASS Upgrades, Anomalies, PASS->PASS Critical Corrections & ABC-Only Updates
    matched_sap = sap_clean[sap_clean["__key__"].isin(email_keys_set)].set_index("__key__")
    matched_email = email_clean[email_clean["__key__"].isin(sap_keys_set)].set_index("__key__")
    matched_keys = matched_sap.index.tolist()
    matched_email = matched_email.loc[matched_keys]
    key_to_pos = {k: i for i, k in enumerate(matched_keys)}

    base_regn = matched_sap["__regn_norm__"].tolist()
    base_sem = matched_sap["__sem_norm__"].tolist()
    base_cname = matched_sap["__cname_norm__"].tolist()
    base_course = matched_sap[sap_course_col].tolist() if sap_course_col else [""] * len(matched_keys)
    base_sap_res = matched_sap["__result_norm__"].tolist()
    base_email_res = matched_email["__result_norm__"].tolist()
    base_sap_yr = matched_sap["__year_int__"].tolist()
    base_email_yr = matched_email["__year_int__"].tolist()
    base_sap_mo = matched_sap["__month_norm__"].tolist()
    base_email_mo = matched_email["__month_norm__"].tolist()

    sap_tot_col = sap_dataset.col_map.get("tot_mrks", "")
    email_tot_col = email_dataset.col_map.get("tot_mrks", "")
    sap_sgpa_col = sap_dataset.col_map.get("sgpa", "")
    email_sgpa_col = email_dataset.col_map.get("sgpa", "")
    email_mrks_col = email_dataset.col_map.get("mrks_rec_status", "")
    email_abc_col = email_dataset.col_map.get("abc_account_id", "")

    base_sap_tot = matched_sap[sap_tot_col].tolist() if sap_tot_col else [""] * len(matched_keys)
    base_email_tot = matched_email[email_tot_col].tolist() if email_tot_col else [""] * len(matched_keys)
    base_sap_sgpa = matched_sap[sap_sgpa_col].tolist() if sap_sgpa_col else [""] * len(matched_keys)
    base_email_sgpa = matched_email[email_sgpa_col].tolist() if email_sgpa_col else [""] * len(matched_keys)
    base_email_mstat = matched_email[email_mrks_col].tolist() if email_mrks_col else [""] * len(matched_keys)

    upgrades_mask = (matched_email["__result_norm__"] == "FAIL") & (
        matched_sap["__result_norm__"] == "PASS"
    )
    upgrade_keys = matched_sap.index[upgrades_mask].tolist()

    upgrade_rows: list[dict[str, str]] = []
    for k in upgrade_keys:
        idx = key_to_pos[k]
        upgrade_rows.append(
            {
                "regn_no": base_regn[idx],
                "sem": base_sem[idx],
                "cname": base_cname[idx],
                "course_name": clean_str(base_course[idx]),
                "email_result": base_email_res[idx],
                "email_year": str(base_email_yr[idx] or ""),
                "email_month": base_email_mo[idx],
                "email_tot_mrks": clean_str(base_email_tot[idx]),
                "email_sgpa": clean_str(base_email_sgpa[idx]),
                "sap_result": base_sap_res[idx],
                "sap_year": str(base_sap_yr[idx] or ""),
                "sap_month": base_sap_mo[idx],
                "sap_tot_mrks": clean_str(base_sap_tot[idx]),
                "sap_sgpa": clean_str(base_sap_sgpa[idx]),
                "recommended_action": "Upload to DigiLocker with mrks_rec_status = M (Modified)",
            }
        )
    fail_to_pass_upgrades_df = pd.DataFrame(upgrade_rows)
    fail_to_pass_raw_sap_df = (
        matched_sap.loc[upgrade_keys, sap_dataset.original_columns].reset_index(drop=True)
        if upgrade_keys
        else pd.DataFrame(columns=sap_dataset.original_columns)
    )

    # Field-level comparison across matched records
    anomaly_records: list[dict[str, str]] = []
    format_only_records: list[dict[str, str]] = []
    col_summary_records: list[dict[str, Any]] = []

    src_sap_col = sap_dataset.col_map.get("source_file", "")
    src_email_col = email_dataset.col_map.get("source_file", "")
    base_src_sap = (
        matched_sap[src_sap_col].tolist() if src_sap_col else matched_sap["__uploaded_file__"].tolist()
    )
    base_src_email = (
        matched_email[src_email_col].tolist() if src_email_col else matched_email["__uploaded_file__"].tolist()
    )

    anomaly_count_by_key: dict[str, int] = {}
    per_key_mismatches: dict[str, list[str]] = {}
    per_key_missing: dict[str, list[str]] = {}
    per_key_diff_details: dict[str, list[tuple[str, str, str, str]]] = {}

    pass_crit_mismatches_by_key: dict[str, list[tuple[str, str, str]]] = {}

    for col in compared_cols:
        s_col = sap_dataset.col_map[col]
        e_col = email_dataset.col_map[col]
        raw_s_list = matched_sap[s_col].tolist()
        raw_e_list = matched_email[e_col].tolist()

        norm_s_list = [
            normalize_for_comparison(
                col,
                v,
                normalize_text=cfg.normalize_text,
                normalize_numbers=cfg.normalize_numbers,
                normalize_months=cfg.normalize_months,
                normalize_dates=cfg.normalize_dates,
                ignore_name_punctuation=cfg.ignore_name_punctuation,
            )
            for v in raw_s_list
        ]
        norm_e_list = [
            normalize_for_comparison(
                col,
                v,
                normalize_text=cfg.normalize_text,
                normalize_numbers=cfg.normalize_numbers,
                normalize_months=cfg.normalize_months,
                normalize_dates=cfg.normalize_dates,
                ignore_name_punctuation=cfg.ignore_name_punctuation,
            )
            for v in raw_e_list
        ]

        cnt_val_mismatch = 0
        cnt_miss_sap = 0
        cnt_miss_email = 0
        cnt_fmt_only = 0
        severity = _classify_column_severity(col)

        for idx, k in enumerate(matched_keys):
            rs = clean_str(raw_s_list[idx])
            re_ = clean_str(raw_e_list[idx])
            ns = norm_s_list[idx]
            ne = norm_e_list[idx]

            if ns != ne:
                if ns == "" and ne != "":
                    atype = "MISSING_IN_SAP"
                    cnt_miss_sap += 1
                    include = cfg.include_one_side_null_anomalies
                elif ne == "" and ns != "":
                    atype = "MISSING_IN_EMAIL"
                    cnt_miss_email += 1
                    include = cfg.include_one_side_null_anomalies
                else:
                    atype = "VALUE_MISMATCH"
                    cnt_val_mismatch += 1
                    include = True

                if include:
                    anomaly_count_by_key[k] = anomaly_count_by_key.get(k, 0) + 1
                    if atype == "VALUE_MISMATCH":
                        per_key_mismatches.setdefault(k, []).append(col)
                    else:
                        per_key_missing.setdefault(k, []).append(f"{col} ({atype})")
                    per_key_diff_details.setdefault(k, []).append((col, re_, rs, atype))
                    anomaly_records.append(
                        {
                            "regn_no": base_regn[idx],
                            "sem": base_sem[idx],
                            "cname": base_cname[idx],
                            "course_name": base_course[idx],
                            "column": col,
                            "severity": severity,
                            "anomaly_type": atype,
                            "email_value": re_,
                            "sap_value": rs,
                            "email_source": base_src_email[idx],
                            "sap_source": base_src_sap[idx],
                        }
                    )
            elif rs != re_:
                cnt_fmt_only += 1
                if cfg.include_format_only_anomalies:
                    format_only_records.append(
                        {
                            "regn_no": base_regn[idx],
                            "sem": base_sem[idx],
                            "cname": base_cname[idx],
                            "column": col,
                            "email_value": re_,
                            "sap_value": rs,
                            "normalized_value": ns,
                        }
                    )

        col_summary_records.append(
            {
                "column": col,
                "severity": severity,
                "value_mismatches": cnt_val_mismatch,
                "missing_in_sap": cnt_miss_sap,
                "missing_in_email": cnt_miss_email,
                "total_actionable": cnt_val_mismatch
                + (cnt_miss_sap + cnt_miss_email if cfg.include_one_side_null_anomalies else 0),
                "format_only_diffs": cnt_fmt_only,
            }
        )

    # Evaluate PASS -> PASS critical VALUE_MISMATCHes across all common PASS_CORRECTION_CRITICAL_COLUMNS
    pass_both_mask = (matched_email["__result_norm__"] == "PASS") & (
        matched_sap["__result_norm__"] == "PASS"
    )
    pass_both_indices = [i for i, flag in enumerate(pass_both_mask.tolist()) if flag]
    if pass_both_indices:
        crit_common = [c for c in common_cols if c in PASS_CORRECTION_CRITICAL_COLUMNS]
        for col in crit_common:
            s_col = sap_dataset.col_map[col]
            e_col = email_dataset.col_map[col]
            raw_s_list = matched_sap[s_col].tolist()
            raw_e_list = matched_email[e_col].tolist()
            for idx in pass_both_indices:
                rs = clean_str(raw_s_list[idx])
                re_ = clean_str(raw_e_list[idx])
                if rs == "" or re_ == "":
                    continue
                ns = normalize_for_comparison(
                    col,
                    rs,
                    normalize_text=cfg.normalize_text,
                    normalize_numbers=cfg.normalize_numbers,
                    normalize_months=cfg.normalize_months,
                    normalize_dates=cfg.normalize_dates,
                    ignore_name_punctuation=cfg.ignore_name_punctuation,
                )
                ne = normalize_for_comparison(
                    col,
                    re_,
                    normalize_text=cfg.normalize_text,
                    normalize_numbers=cfg.normalize_numbers,
                    normalize_months=cfg.normalize_months,
                    normalize_dates=cfg.normalize_dates,
                    ignore_name_punctuation=cfg.ignore_name_punctuation,
                )
                if ns != ne:
                    k = matched_keys[idx]
                    pass_crit_mismatches_by_key.setdefault(k, []).append((col, re_, rs))

    pass_correction_keys = [k for k in matched_keys if k in pass_crit_mismatches_by_key]
    pass_correction_rows: list[dict[str, str]] = []
    for k in pass_correction_keys:
        idx = key_to_pos[k]
        diffs = pass_crit_mismatches_by_key[k]
        pass_correction_rows.append(
            {
                "regn_no": base_regn[idx],
                "sem": base_sem[idx],
                "cname": base_cname[idx],
                "course_name": clean_str(base_course[idx]),
                "mismatch_columns": ", ".join(c for c, _, _ in diffs),
                "email_values": " | ".join(f"{c}={ev}" for c, ev, _ in diffs),
                "sap_values": " | ".join(f"{c}={sv}" for c, _, sv in diffs),
                "email_tot_mrks": clean_str(base_email_tot[idx]),
                "sap_tot_mrks": clean_str(base_sap_tot[idx]),
                "email_sgpa": clean_str(base_email_sgpa[idx]),
                "sap_sgpa": clean_str(base_sap_sgpa[idx]),
                "email_mrks_rec_status": clean_str(base_email_mstat[idx]),
                "recommended_action": "Upload to DigiLocker with mrks_rec_status = M (Modified)",
            }
        )

    pass_corrections_df = pd.DataFrame(pass_correction_rows)
    pass_corrections_raw_sap_df = (
        matched_sap.loc[pass_correction_keys, sap_dataset.original_columns].reset_index(drop=True)
        if pass_correction_keys
        else pd.DataFrame(columns=sap_dataset.original_columns)
    )

    # Detect ABC-Only Updates (Email abc_account_id missing/dummy/invalid while SAP or verified map has valid 12-digit ID)
    upgrade_keys_set = set(upgrade_keys)
    pass_corr_keys_set = set(pass_correction_keys)
    abc_only_keys: list[str] = []
    abc_only_rows: list[dict[str, str]] = []
    abc_only_resolved_map: dict[str, str] = {}
    abc_only_email_status_map: dict[str, str] = {}

    if email_abc_col:
        matched_email_abcs = matched_email[email_abc_col].tolist()
        matched_sap_abcs = (
            matched_sap[sap_abc_col].tolist() if sap_abc_col else [""] * len(matched_keys)
        )
        for idx, k in enumerate(matched_keys):
            if k in upgrade_keys_set or k in pass_corr_keys_set:
                continue
            raw_e_abc = clean_str(matched_email_abcs[idx])
            e_status, _ = validate_abc_id(raw_e_abc)
            if e_status in ("VALID", "FIXABLE_FORMAT"):
                continue

            raw_s_abc = clean_str(matched_sap_abcs[idx])
            s_status, s_cleaned = validate_abc_id(raw_s_abc)
            regn = base_regn[idx]

            resolved_abc = ""
            abc_src = ""
            if s_status == "VALID":
                resolved_abc = s_cleaned
                abc_src = "SAP (Same Row)"
            elif s_status == "FIXABLE_FORMAT":
                resolved_abc = s_cleaned
                abc_src = "SAP (Format Cleaned)"
            elif regn in verified_abc_map:
                resolved_abc, srcs = verified_abc_map[regn]
                abc_src = f"Verified ABC Map ({', '.join(srcs)})"

            if not resolved_abc:
                continue

            abc_only_keys.append(k)
            abc_only_resolved_map[k] = resolved_abc
            abc_only_email_status_map[k] = e_status
            abc_only_rows.append(
                {
                    "regn_no": regn,
                    "sem": base_sem[idx],
                    "cname": base_cname[idx],
                    "course_name": clean_str(base_course[idx]),
                    "sap_result": base_sap_res[idx],
                    "email_result": base_email_res[idx],
                    "email_abc_id": raw_e_abc,
                    "email_abc_status": e_status,
                    "sap_abc_id": raw_s_abc,
                    "resolved_abc_id": resolved_abc,
                    "abc_source": abc_src,
                    "email_mrks_rec_status": clean_str(base_email_mstat[idx]),
                    "recommended_action": "Upload to DigiLocker with mrks_rec_status = M (Modified)",
                }
            )

    abc_only_updates_df = pd.DataFrame(abc_only_rows)
    if abc_only_keys:
        abc_only_updates_raw_sap_df = matched_sap.loc[
            abc_only_keys, sap_dataset.original_columns
        ].reset_index(drop=True)
        if sap_abc_col:
            abc_only_updates_raw_sap_df[sap_abc_col] = [
                abc_only_resolved_map[k] for k in abc_only_keys
            ]
    else:
        abc_only_updates_raw_sap_df = pd.DataFrame(columns=sap_dataset.original_columns)
    abc_only_keys_set = set(abc_only_keys)

    anomalies_long_df = pd.DataFrame(anomaly_records)
    format_only_df = pd.DataFrame(format_only_records)
    if col_summary_records:
        column_anomaly_summary_df = (
            pd.DataFrame(col_summary_records)
            .sort_values(["total_actionable", "value_mismatches"], ascending=False)
            .reset_index(drop=True)
        )
    else:
        column_anomaly_summary_df = pd.DataFrame(
            columns=[
                "column",
                "severity",
                "value_mismatches",
                "missing_in_sap",
                "missing_in_email",
                "total_actionable",
                "format_only_diffs",
            ]
        )

    wide_rows: list[dict[str, Any]] = []
    for idx, k in enumerate(matched_keys):
        acnt = anomaly_count_by_key.get(k, 0)
        if acnt > 0:
            mm = per_key_mismatches.get(k, [])
            ms = per_key_missing.get(k, [])
            wide_rows.append(
                {
                    "regn_no": base_regn[idx],
                    "sem": base_sem[idx],
                    "cname": base_cname[idx],
                    "course_name": base_course[idx],
                    "total_anomalies": acnt,
                    "value_mismatch_count": len(mm),
                    "missing_field_count": len(ms),
                    "value_mismatch_columns": ", ".join(mm),
                    "missing_field_columns": ", ".join(ms),
                }
            )
    anomalies_wide_df = pd.DataFrame(wide_rows)

    # Step 6: Complete Student-Wise Report, Student_Pivot_Status & Student 360° Timeline Map
    all_sems = sorted(
        (set(sap_clean["__sem_norm__"]) | set(email_clean["__sem_norm__"])) - {""},
        key=sem_sort_key,
    )
    semester_columns = [f"Sem {s}" for s in all_sems]

    sap_by_student: dict[str, dict[str, dict[str, str]]] = {}
    email_by_student: dict[str, dict[str, dict[str, str]]] = {}
    sap_records_by_student_sem: dict[tuple[str, str], list[dict[str, str]]] = {}
    email_records_by_student_sem: dict[tuple[str, str], list[dict[str, str]]] = {}
    student_profile: dict[str, dict[str, str]] = {}

    pending_readiness_by_key: dict[str, str] = dict(
        zip(pending_sap["__key__"].tolist(), pending_sap["__abc_readiness__"].tolist())
    )

    _ingest_dataset_for_student_map(
        sap_clean,
        sap_dataset,
        sap_by_student,
        sap_records_by_student_sem,
        student_profile,
        sap_conflict_by_key=sap_conflict_by_key,
    )
    _ingest_dataset_for_student_map(
        email_clean,
        email_dataset,
        email_by_student,
        email_records_by_student_sem,
        student_profile,
    )

    student_rows: list[dict[str, Any]] = []
    pivot_rows: list[dict[str, Any]] = []
    student_pivot_cell_status: dict[tuple[str, str], str] = {}
    student_timeline_map: dict[str, list[dict[str, str]]] = {}

    status_counts = {
        "Complete": 0,
        "Complete (With Anomalies)": 0,
        "Partial": 0,
        "Not Uploaded": 0,
        "Email Only": 0,
    }

    for regn in sorted(student_profile.keys(), key=_natural_sort_key):
        prof = student_profile[regn]
        s_map = sap_by_student.get(regn, {})
        e_map = email_by_student.get(regn, {})

        sap_sems = sorted(s_map.keys(), key=sem_sort_key)
        email_sems = sorted(e_map.keys(), key=sem_sort_key)
        pending_sems = [s for s in sap_sems if s not in e_map]
        email_only_sems = [s for s in email_sems if s not in s_map]
        conflict_sems = [s for s in sap_sems if s_map[s]["conflict"]]

        student_anomalies = 0
        has_upgrade_or_corr = False
        timeline_entries: list[dict[str, str]] = []

        sem_cells: dict[str, str] = {}
        pivot_sem_cells: dict[str, str] = {}
        pivot_remark_notes: list[str] = []

        for s in all_sems:
            col_label = f"Sem {s}"
            in_s = s in s_map
            in_e = s in e_map
            sd = s_map.get(s)
            ed = e_map.get(s)
            s_recs = sap_records_by_student_sem.get((regn, s), [])
            e_recs = email_records_by_student_sem.get((regn, s), [])
            s_keys_for_sem = [r["key"] for r in s_recs]
            e_keys_for_sem = [r["key"] for r in e_recs]
            s_keys_set_sem = set(s_keys_for_sem)
            e_keys_set_sem = set(e_keys_for_sem)
            matched_sem_keys = [k for k in s_keys_for_sem if k in e_keys_set_sem]
            pending_sem_recs = [r for r in s_recs if r["key"] not in e_keys_set_sem]
            email_only_sem_recs = [r for r in e_recs if r["key"] not in s_keys_set_sem]

            # Student_Wise_Report cell status
            if in_s and in_e:
                assert sd is not None and ed is not None
                k = sd["key"]
                acnt = anomaly_count_by_key.get(k, 0)
                student_anomalies += acnt
                if k in upgrade_keys_set:
                    has_upgrade_or_corr = True
                    cell_status = "UPGRADE_NEEDED (FAIL->PASS)"
                elif k in pass_corr_keys_set:
                    has_upgrade_or_corr = True
                    cell_status = (
                        f"M_CORRECTION_NEEDED ({acnt or len(pass_crit_mismatches_by_key.get(k, []))})"
                    )
                elif k in abc_only_keys_set:
                    has_upgrade_or_corr = True
                    cell_status = "M_ABC_UPDATE_NEEDED"
                elif acnt > 0:
                    cell_status = f"UPLOADED_ANOMALY ({acnt})"
                elif sd["conflict"]:
                    cell_status = "UPLOADED_OK (SAP_DEDUPED)"
                else:
                    cell_status = "UPLOADED_OK"
                sem_cells[col_label] = cell_status
                mm_cols = ", ".join(per_key_mismatches.get(k, []) + per_key_missing.get(k, []))
            elif in_s and not in_e:
                assert sd is not None
                res_tag = sd["result"] or "UNKNOWN"
                abc_tag = pending_readiness_by_key.get(sd["key"], "")
                if sd["conflict"]:
                    cell_status = f"PENDING_{res_tag} (SAP_CONFLICT)"
                elif abc_tag in {"RECOVERABLE", "MISSING", "INVALID_OR_DUMMY"}:
                    cell_status = f"PENDING_{res_tag} (ABC:{abc_tag})"
                else:
                    cell_status = f"PENDING_{res_tag}"
                sem_cells[col_label] = cell_status
                mm_cols = ""
            elif in_e and not in_s:
                assert ed is not None
                cell_status = f"EMAIL_ONLY ({ed['result']})"
                sem_cells[col_label] = cell_status
                mm_cols = ""
            else:
                cell_status = "-"
                sem_cells[col_label] = "-"
                mm_cols = ""

            # Student_Pivot_Status cell value, color & remark
            if not in_s and not in_e:
                pivot_sem_cells[col_label] = "-"
                student_pivot_cell_status[(regn, col_label)] = "EMPTY"
            else:
                if in_s and sd is not None:
                    canon_p = (
                        int(sd["year_int"] or 0),
                        int(sd["month_int"] or 0),
                        sd["month_short"],
                    )
                    attempt_pool = sap_raw_periods.get((regn, s), set())
                    fallback_lbl = sd["result"] or "SAP"
                else:
                    assert ed is not None
                    canon_p = (
                        int(ed["year_int"] or 0),
                        int(ed["month_int"] or 0),
                        ed["month_short"],
                    )
                    attempt_pool = email_raw_periods.get((regn, s), set())
                    fallback_lbl = ed["result"] or "EMAIL"

                pivot_sem_cells[col_label] = _format_attempt_cell(
                    canon_p, attempt_pool, fallback_label=fallback_lbl
                )

                sem_has_red = False
                sem_has_yellow = len(pending_sem_recs) > 0
                sem_has_orange = len(email_only_sem_recs) > 0

                for mk in matched_sem_keys:
                    mk_acnt = anomaly_count_by_key.get(mk, 0)
                    is_upg = mk in upgrade_keys_set
                    is_pcc = mk in pass_corr_keys_set
                    is_abc_only = mk in abc_only_keys_set
                    if is_upg or is_pcc or is_abc_only or mk_acnt > 0:
                        sem_has_red = True
                        reasons: list[str] = []
                        if is_upg:
                            reasons.append("FAIL->PASS upgrade needed")
                        else:
                            seen_diff_cols: set[str] = set()
                            for col_d, ev_d, sv_d in pass_crit_mismatches_by_key.get(mk, []):
                                if col_d in seen_diff_cols:
                                    continue
                                if col_d == "abc_account_id" and is_abc_only:
                                    continue
                                seen_diff_cols.add(col_d)
                                reasons.append(
                                    f"{col_d}: Email={ev_d or 'BLANK'} vs SAP={sv_d or 'BLANK'}"
                                )
                            for col_d, ev_d, sv_d, _ in per_key_diff_details.get(mk, []):
                                if col_d in seen_diff_cols:
                                    continue
                                if col_d == "abc_account_id" and is_abc_only:
                                    continue
                                seen_diff_cols.add(col_d)
                                reasons.append(
                                    f"{col_d}: Email={ev_d or 'BLANK'} vs SAP={sv_d or 'BLANK'}"
                                )
                        if is_abc_only:
                            e_abc_st = abc_only_email_status_map.get(mk, "MISSING")
                            if e_abc_st == "MISSING":
                                reasons.append("ABC ID missing in DigiLocker, valid ID in SAP")
                            else:
                                reasons.append(
                                    "ABC ID invalid/dummy in DigiLocker, valid ID in SAP"
                                )
                        if len(reasons) > 5:
                            reasons_str = "; ".join(reasons[:5]) + f"; (+{len(reasons) - 5} more)"
                        else:
                            reasons_str = "; ".join(reasons)
                        pivot_remark_notes.append(
                            f"{col_label}: Wrong data in DigiLocker ({reasons_str})"
                        )

                for pr in pending_sem_recs:
                    res_tag = pr["result"]
                    if res_tag:
                        pivot_remark_notes.append(
                            f"{col_label}: Pending in SAP (Not uploaded \u2014 {res_tag})"
                        )
                    else:
                        pivot_remark_notes.append(f"{col_label}: Pending in SAP (Not uploaded)")

                if email_only_sem_recs:
                    pivot_remark_notes.append(
                        f"{col_label}: Email-only (Present in DigiLocker, missing in SAP)"
                    )

                if sem_has_red:
                    student_pivot_cell_status[(regn, col_label)] = "RED"
                elif sem_has_yellow:
                    student_pivot_cell_status[(regn, col_label)] = "YELLOW"
                elif sem_has_orange:
                    student_pivot_cell_status[(regn, col_label)] = "ORANGE"
                else:
                    student_pivot_cell_status[(regn, col_label)] = "GREEN"

            if in_s or in_e:
                timeline_entries.append(
                    {
                        "Sem": s,
                        "Status": cell_status,
                        "SAP_Result": sd["result"] if sd else "-",
                        "Email_Result": ed["result"] if ed else "-",
                        "SAP_Period": sd["period"] if sd else "-",
                        "Email_Period": ed["period"] if ed else "-",
                        "SAP_Marks_SGPA": sd["marks_sgpa"] if sd else "-",
                        "Email_Marks_SGPA": ed["marks_sgpa"] if ed else "-",
                        "SAP_Credit_Pts": sd["credit_pts"] if sd else "-",
                        "Email_Credit_Pts": ed["credit_pts"] if ed else "-",
                        "SAP_ABC_ID": sd["abc_id"] if sd else "-",
                        "Email_ABC_ID": ed["abc_id"] if ed else "-",
                        "SAP_Course_ID": sd["course_id"] if sd else "-",
                        "Email_Course_ID": ed["course_id"] if ed else "-",
                        "Email_O_M": ed["mrks_status"] if ed else "-",
                        "Mismatch_Columns": mm_cols or "-",
                        "SAP_Conflict": sd["conflict"] if sd and sd["conflict"] else "-",
                    }
                )

        student_timeline_map[regn] = timeline_entries

        uploaded_sap_sems = len(sap_sems) - len(pending_sems)
        if not sap_sems and email_sems:
            overall = "Email Only"
        elif len(pending_sems) == len(sap_sems) and len(sap_sems) > 0:
            overall = "Not Uploaded"
        elif len(pending_sems) > 0:
            overall = "Partial"
        elif student_anomalies > 0 or has_upgrade_or_corr:
            overall = "Complete (With Anomalies)"
        else:
            overall = "Complete"

        status_counts[overall] += 1

        row_dict: dict[str, Any] = {
            "regn_no": prof["regn_no"],
            "cname": prof["cname"],
            "course_name": prof["course_name"],
            "academic_course_id": prof["academic_course_id"],
            "admission_year": prof["admission_year"],
            "abc_account_id": prof["abc_account_id"],
            "dob": prof["dob"],
            "gender": prof["gender"],
            "Overall_Status": overall,
            "Total_SAP_Sems": len(sap_sems),
            "Uploaded_SAP_Sems": uploaded_sap_sems,
            "Pending_SAP_Sems": len(pending_sems),
            "Email_Only_Sems": len(email_only_sems),
            "SAP_Conflict_Count": len(conflict_sems),
            "Cross_File_Anomalies": student_anomalies,
            "SAP_Semesters": ", ".join(sap_sems),
            "Email_Semesters": ", ".join(email_sems),
            "Pending_Semesters": ", ".join(pending_sems),
            "Email_Only_Semesters": ", ".join(email_only_sems),
            "SAP_Conflict_Semesters": ", ".join(conflict_sems),
        }
        row_dict.update(sem_cells)
        student_rows.append(row_dict)

        pivot_row: dict[str, Any] = {
            "regn_no": prof["regn_no"],
            "cname": prof["cname"],
            "course_name": prof["course_name"],
            "admission_year": prof["admission_year"],
            "abc_account_id": prof["abc_account_id"],
        }
        pivot_row.update(pivot_sem_cells)
        pivot_row["Remark"] = " | ".join(pivot_remark_notes) if pivot_remark_notes else "All OK"
        pivot_rows.append(pivot_row)

    student_report_df = pd.DataFrame(student_rows)
    pivot_columns = [
        "regn_no",
        "cname",
        "course_name",
        "admission_year",
        "abc_account_id",
        *semester_columns,
        "Remark",
    ]
    student_pivot_df = (
        pd.DataFrame(pivot_rows, columns=pivot_columns)
        if pivot_rows
        else pd.DataFrame(columns=pivot_columns)
    )

    sap_removed_count = len(sap_dataset.df) - len(sap_clean)
    pending_pass_count = int((pure_pending_sap_df["__pending_result__"] == "PASS").sum())
    pending_fail_count = int((pure_pending_sap_df["__pending_result__"] == "FAIL").sum())
    abc_ready_count = int((pure_pending_sap_df["__abc_readiness__"] == "READY").sum())
    abc_recoverable_count = int((pure_pending_sap_df["__abc_readiness__"] == "RECOVERABLE").sum())
    abc_missing_or_invalid_count = len(pure_pending_sap_df) - abc_ready_count - abc_recoverable_count

    metrics = {
        "sap_raw_rows": len(sap_dataset.df),
        "sap_dedup_rows": len(sap_clean),
        "sap_duplicates_or_superseded": sap_removed_count,
        "sap_conflict_groups": len(sap_conflict_by_key),
        "email_raw_rows": len(email_dataset.df),
        "email_dedup_rows": len(email_clean),
        "email_duplicates_removed": len(email_dataset.df) - len(email_clean),
        "matched_records": len(matched_keys),
        "pending_total": len(pure_pending_sap_df),
        "pending_pass": pending_pass_count,
        "pending_fail": pending_fail_count,
        "pending_other": len(pure_pending_sap_df) - pending_pass_count - pending_fail_count,
        "pending_abc_ready": abc_ready_count,
        "pending_abc_recoverable": abc_recoverable_count,
        "pending_abc_missing_or_invalid": abc_missing_or_invalid_count,
        "fail_to_pass_upgrades": len(fail_to_pass_upgrades_df),
        "pass_critical_corrections": len(pass_corrections_df),
        "abc_only_updates": len(abc_only_updates_df),
        "email_only_records": len(email_only_df),
        "email_only_in_sap_other_sems": int((email_only_df["exists_in_sap_other_sems"] == "Yes").sum())
        if not email_only_df.empty
        else 0,
        "matched_rows_with_anomalies": len(anomalies_wide_df),
        "total_field_anomalies": len(anomalies_long_df),
        "total_students": len(student_report_df),
        "students_complete": status_counts["Complete"],
        "students_complete_with_anomalies": status_counts["Complete (With Anomalies)"],
        "students_partial": status_counts["Partial"],
        "students_not_uploaded": status_counts["Not Uploaded"],
        "students_email_only": status_counts["Email Only"],
    }

    return ReconciliationResult(
        config=cfg,
        sap_original_columns=sap_dataset.original_columns,
        sap_col_map=sap_dataset.col_map,
        email_col_map=email_dataset.col_map,
        common_columns=common_cols,
        compared_columns=compared_cols,
        pure_pending_sap_df=pure_pending_sap_df,
        abc_backfill_audit_df=abc_backfill_audit_df,
        verified_abc_map=verified_abc_map,
        fail_to_pass_upgrades_df=fail_to_pass_upgrades_df,
        fail_to_pass_raw_sap_df=fail_to_pass_raw_sap_df,
        pass_corrections_df=pass_corrections_df,
        pass_corrections_raw_sap_df=pass_corrections_raw_sap_df,
        abc_only_updates_df=abc_only_updates_df,
        abc_only_updates_raw_sap_df=abc_only_updates_raw_sap_df,
        sap_conflicts_df=sap_conflicts_df,
        email_conflicts_df=email_conflicts_df,
        anomalies_long_df=anomalies_long_df,
        anomalies_wide_df=anomalies_wide_df,
        format_only_df=format_only_df,
        column_anomaly_summary_df=column_anomaly_summary_df,
        email_only_df=email_only_df,
        student_report_df=student_report_df,
        student_pivot_df=student_pivot_df,
        student_pivot_cell_status=student_pivot_cell_status,
        student_timeline_map=student_timeline_map,
        semester_columns=semester_columns,
        metrics=metrics,
    )


# ==============================================================================
# SECTION 4: NAD-Ready CSV & Styled Multi-Sheet Excel Exporters
# ==============================================================================

INTERNAL_MERGE_COLUMNS: set[str] = {"source_file", "csv_filename"}

EXPORT_TEXT_ID_COLUMNS: set[str] = PROTECTED_ID_COLUMNS | {
    "abc_id",
    "original_abc_id",
    "recovered_abc_id",
    "email_abc_id",
    "sap_abc_id",
    "resolved_abc_id",
    "email_value",
    "sap_value",
}

PIVOT_STYLE_MAP: dict[str, tuple[str, str]] = {
    "GREEN": ("#D1FAE5", "#065F46"),
    "RED": ("#FEE2E2", "#991B1B"),
    "YELLOW": ("#FEF3C7", "#92400E"),
    "ORANGE": ("#FFEDD5", "#9A3412"),
}


def infer_export_prefix(sap_filenames: Iterable[str]) -> str:
    """Derive a clean filename prefix from the first uploaded SAP filename."""
    names = list(sap_filenames)
    if not names:
        return "DigiLocker"
    stem = Path(names[0]).stem.strip()
    token = re.split(r"[\s_\-]+", stem)[0] if stem else "DigiLocker"
    cleaned = re.sub(r"[^A-Za-z0-9]", "", token)
    return cleaned or "DigiLocker"


def build_dynamic_filenames(
    sap_filenames: Iterable[str],
    result_filter: str = "ALL",
    abc_filter: str = "ALL",
    today_str: str | None = None,
) -> tuple[str, str, str]:
    """Return (pure_csv_name, pure_xlsx_name, report_xlsx_name) with date stamp."""
    prefix = infer_export_prefix(sap_filenames)
    date_tag = today_str or dt.date.today().strftime("%Y-%m-%d")
    abc_tag = "" if abc_filter == "ALL" else f"_{abc_filter}"
    pure_base = f"{prefix}_Pure_Pending_{result_filter}{abc_tag}_{date_tag}"
    report_name = f"{prefix}_DigiLocker_Report_{date_tag}.xlsx"
    return (f"{pure_base}.csv", f"{pure_base}.xlsx", report_name)


def prepare_m_status_sap_export(
    result: ReconciliationResult,
    raw_sap_df: pd.DataFrame,
    *,
    auto_backfill_abc: bool = True,
    drop_internal_cols: bool = True,
    trim_whitespace: bool = True,
) -> pd.DataFrame:
    """Prepare raw SAP rows stamped with MRKS_REC_STATUS = 'M' for direct DigiLocker upload."""
    if raw_sap_df.empty:
        cols = [
            c
            for c in result.sap_original_columns
            if not drop_internal_cols or normalize_col_name(c) not in INTERNAL_MERGE_COLUMNS
        ]
        return pd.DataFrame(columns=cols)

    m_df = raw_sap_df[result.sap_original_columns].copy()
    sap_mrks_col = result.sap_col_map.get("mrks_rec_status")
    sap_abc_col = result.sap_col_map.get("abc_account_id")
    sap_regn_col = result.sap_col_map.get("regn_no")

    if sap_mrks_col:
        m_df[sap_mrks_col] = "M"

    if auto_backfill_abc and sap_abc_col and not m_df.empty:
        abc_vals = m_df[sap_abc_col].tolist()
        regn_vals = m_df[sap_regn_col].tolist() if sap_regn_col else [""] * len(m_df)
        updated_abcs: list[str] = []
        for a_raw, r_raw in zip(abc_vals, regn_vals):
            status, cleaned_abc = validate_abc_id(a_raw)
            if status == "FIXABLE_FORMAT":
                updated_abcs.append(cleaned_abc)
            elif status != "VALID":
                regn_norm = norm_text(r_raw, case_insensitive=True) if r_raw else ""
                if regn_norm and regn_norm in result.verified_abc_map:
                    updated_abcs.append(result.verified_abc_map[regn_norm][0])
                else:
                    updated_abcs.append(clean_str(a_raw))
            else:
                updated_abcs.append(cleaned_abc)
        m_df[sap_abc_col] = updated_abcs

    if drop_internal_cols:
        keep_cols = [c for c in m_df.columns if normalize_col_name(c) not in INTERNAL_MERGE_COLUMNS]
        m_df = m_df[keep_cols]

    if trim_whitespace and not m_df.empty:
        for col in m_df.columns:
            m_df[col] = m_df[col].map(clean_str)

    return m_df.reset_index(drop=True)


def filter_pending_sap(
    result: ReconciliationResult,
    result_filter: Literal["ALL", "PASS_ONLY", "FAIL_ONLY"] = "ALL",
    abc_filter: Literal[
        "ALL", "READY_ONLY", "READY_AND_RECOVERABLE", "MISSING_OR_INVALID_ONLY"
    ] = "ALL",
    auto_backfill_abc: bool = False,
    include_fail_to_pass_upgrades: bool = False,
    include_pass_corrections: bool = False,
    include_abc_only_updates: bool = False,
    drop_internal_cols: bool = True,
    trim_whitespace: bool = True,
) -> pd.DataFrame:
    """Filter, backfill, stamp M-status, sanitize, and export exact SAP columns."""
    df = result.pure_pending_sap_df.copy()

    if result_filter == "PASS_ONLY":
        df = df[df["__pending_result__"] == "PASS"]
    elif result_filter == "FAIL_ONLY":
        df = df[df["__pending_result__"] == "FAIL"]

    if abc_filter == "READY_ONLY":
        df = df[df["__abc_readiness__"] == "READY"]
    elif abc_filter == "READY_AND_RECOVERABLE":
        df = df[df["__abc_readiness__"].isin(["READY", "RECOVERABLE"])]
    elif abc_filter == "MISSING_OR_INVALID_ONLY":
        df = df[df["__abc_readiness__"].isin(["MISSING", "INVALID_OR_DUMMY"])]

    sap_abc_col = result.sap_col_map.get("abc_account_id")
    if auto_backfill_abc and sap_abc_col and not df.empty:
        rec_mask = (df["__abc_readiness__"] == "RECOVERABLE") & (df["__recovered_abc_id__"] != "")
        df.loc[rec_mask, sap_abc_col] = df.loc[rec_mask, "__recovered_abc_id__"]

    out = df[result.sap_original_columns].copy()

    extra_frames: list[pd.DataFrame] = []
    for enabled, raw_m_df in (
        (include_fail_to_pass_upgrades, result.fail_to_pass_raw_sap_df),
        (include_pass_corrections, result.pass_corrections_raw_sap_df),
        (include_abc_only_updates, result.abc_only_updates_raw_sap_df),
    ):
        if enabled and not raw_m_df.empty:
            extra_frames.append(
                prepare_m_status_sap_export(
                    result,
                    raw_m_df,
                    auto_backfill_abc=auto_backfill_abc,
                    drop_internal_cols=False,
                    trim_whitespace=False,
                )
            )

    if extra_frames:
        out = pd.concat([out] + extra_frames, ignore_index=True)

    if drop_internal_cols:
        keep_cols = [c for c in out.columns if normalize_col_name(c) not in INTERNAL_MERGE_COLUMNS]
        out = out[keep_cols]

    if trim_whitespace and not out.empty:
        for col in out.columns:
            out[col] = out[col].map(clean_str)

    return out.reset_index(drop=True)


def build_pure_pending_csv(export_df: pd.DataFrame) -> bytes:
    """Export DataFrame as clean UTF-8 CSV ready for direct NAD / DigiLocker portal upload."""
    return export_df.to_csv(index=False).encode("utf-8")


def _apply_xlsxwriter_sheet(
    writer: pd.ExcelWriter,
    df: pd.DataFrame,
    sheet_name: str,
    header_fmt: Any,
    text_fmt: Any,
    max_width: int = 42,
) -> None:
    safe_df = df.copy()
    if safe_df.empty:
        safe_df = pd.DataFrame({"Status": ["No records found in this category."]})
    safe_df.to_excel(writer, sheet_name=sheet_name, index=False, startrow=1, header=False)
    ws = writer.sheets[sheet_name]
    ws.freeze_panes(1, 0)
    ws.hide_gridlines(2)
    ws.autofilter(0, 0, max(len(safe_df), 1), max(len(safe_df.columns) - 1, 0))

    for col_idx, col_name in enumerate(safe_df.columns):
        ws.write(0, col_idx, str(col_name), header_fmt)
        sample_vals = safe_df.iloc[:200, col_idx].astype(str).tolist()
        max_len = max([len(str(col_name))] + [len(v) for v in sample_vals], default=10)
        col_cap = 80 if str(col_name) == "Remark" else max_width
        width = min(max(max_len + 3, 12), col_cap)
        if normalize_col_name(col_name) in EXPORT_TEXT_ID_COLUMNS:
            ws.set_column(col_idx, col_idx, width, text_fmt)
        else:
            ws.set_column(col_idx, col_idx, width)


def _apply_openpyxl_sheet(
    writer: pd.ExcelWriter,
    df: pd.DataFrame,
    sheet_name: str,
    max_width: int = 42,
) -> None:
    safe_df = df.copy()
    if safe_df.empty:
        safe_df = pd.DataFrame({"Status": ["No records found in this category."]})
    safe_df.to_excel(writer, sheet_name=sheet_name, index=False)
    ws = writer.sheets[sheet_name]
    ws.freeze_panes = "A2"
    ws.views.sheetView[0].showGridLines = False
    if len(safe_df.columns) > 0:
        ws.auto_filter.ref = f"A1:{get_column_letter(len(safe_df.columns))}{max(len(safe_df) + 1, 2)}"

    hdr_fill = PatternFill(start_color="1E3A5F", end_color="1E3A5F", fill_type="solid")
    hdr_font = Font(bold=True, color="FFFFFF")
    max_row = len(safe_df) + 1
    for col_idx, col_name in enumerate(safe_df.columns, start=1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = hdr_fill
        cell.font = hdr_font
        if normalize_col_name(col_name) in EXPORT_TEXT_ID_COLUMNS:
            for r_i in range(2, max_row + 1):
                ws.cell(row=r_i, column=col_idx).number_format = "@"
        sample_vals = safe_df.iloc[:200, col_idx - 1].astype(str).tolist()
        max_len = max([len(str(col_name))] + [len(v) for v in sample_vals], default=10)
        col_cap = 80 if str(col_name) == "Remark" else max_width
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max(max_len + 3, 12), col_cap)


def _write_pivot_sheet_xlsxwriter(
    writer: pd.ExcelWriter,
    pivot_df: pd.DataFrame,
    cell_status: dict[tuple[Any, str], str],
    sheet_name: str,
    header_fmt: Any,
    text_fmt: Any,
    max_width: int = 42,
) -> None:
    """Write Student_Pivot_Status with cell-by-cell pastel colors, frozen Row 1, autofilter & text IDs."""
    if pivot_df.empty:
        _apply_xlsxwriter_sheet(writer, pivot_df, sheet_name, header_fmt, text_fmt, max_width=max_width)
        return

    wb = writer.book
    ws = wb.add_worksheet(sheet_name)
    writer.sheets[sheet_name] = ws
    ws.freeze_panes(1, 0)
    ws.hide_gridlines(2)

    cols = list(pivot_df.columns)
    ws.autofilter(0, 0, len(pivot_df), len(cols) - 1)

    status_fmts = {
        code: wb.add_format({"bg_color": bg, "font_color": fg, "align": "center", "valign": "vcenter"})
        for code, (bg, fg) in PIVOT_STYLE_MAP.items()
    }
    center_fmt = wb.add_format({"align": "center", "valign": "vcenter"})

    identity_or_remark = {
        "regn_no",
        "cname",
        "course_name",
        "admission_year",
        "abc_account_id",
        "Remark",
    }
    sem_col_indices = {i: c for i, c in enumerate(cols) if c not in identity_or_remark}
    id_col_indices = {i for i, c in enumerate(cols) if normalize_col_name(c) in EXPORT_TEXT_ID_COLUMNS}
    regn_idx = cols.index("regn_no") if "regn_no" in cols else 0

    for col_idx, col_name in enumerate(cols):
        ws.write(0, col_idx, str(col_name), header_fmt)

    col_lists = [pivot_df[c].tolist() for c in cols]
    for r_zero, row_vals in enumerate(zip(*col_lists)):
        r_excel = r_zero + 1
        regn_val = clean_str(row_vals[regn_idx])
        for c_idx, raw_val in enumerate(row_vals):
            sval = clean_str(raw_val)
            if c_idx in sem_col_indices:
                scol = sem_col_indices[c_idx]
                st_code = cell_status.get(
                    (regn_val, scol), cell_status.get((r_zero, scol), "EMPTY")
                )
                fmt = status_fmts.get(st_code)
                if fmt is not None:
                    ws.write_string(r_excel, c_idx, sval, fmt)
                elif sval == "-":
                    ws.write_string(r_excel, c_idx, sval, center_fmt)
                else:
                    ws.write_string(r_excel, c_idx, sval)
            elif c_idx in id_col_indices:
                ws.write_string(r_excel, c_idx, sval, text_fmt)
            else:
                ws.write_string(r_excel, c_idx, sval)

    for col_idx, col_name in enumerate(cols):
        sample_vals = pivot_df.iloc[:200, col_idx].astype(str).tolist()
        max_len = max([len(str(col_name))] + [len(v) for v in sample_vals], default=10)
        col_cap = 80 if str(col_name) == "Remark" else max_width
        width = min(max(max_len + 3, 12), col_cap)
        if col_idx in id_col_indices:
            ws.set_column(col_idx, col_idx, width, text_fmt)
        else:
            ws.set_column(col_idx, col_idx, width)


def _write_pivot_sheet_openpyxl(
    writer: pd.ExcelWriter,
    pivot_df: pd.DataFrame,
    cell_status: dict[tuple[Any, str], str],
    sheet_name: str,
    max_width: int = 42,
) -> None:
    """Write Student_Pivot_Status via openpyxl with cell-by-cell pastel colors and text IDs."""
    _apply_openpyxl_sheet(writer, pivot_df, sheet_name, max_width=max_width)
    if pivot_df.empty:
        return

    ws = writer.sheets[sheet_name]
    cols = list(pivot_df.columns)
    identity_or_remark = {
        "regn_no",
        "cname",
        "course_name",
        "admission_year",
        "abc_account_id",
        "Remark",
    }
    sem_col_map = {c: i + 1 for i, c in enumerate(cols) if c not in identity_or_remark}
    regn_vals = pivot_df["regn_no"].tolist() if "regn_no" in cols else [""] * len(pivot_df)

    fills = {
        code: PatternFill(start_color=bg.lstrip("#"), end_color=bg.lstrip("#"), fill_type="solid")
        for code, (bg, _) in PIVOT_STYLE_MAP.items()
    }
    fonts = {
        code: Font(color=fg.lstrip("#"))
        for code, (_, fg) in PIVOT_STYLE_MAP.items()
    }

    for r_zero, regn_raw in enumerate(regn_vals):
        r_excel = r_zero + 2
        regn_val = clean_str(regn_raw)
        for scol, col_excel in sem_col_map.items():
            st_code = cell_status.get((regn_val, scol), cell_status.get((r_zero, scol), "EMPTY"))
            if st_code in fills:
                cell = ws.cell(row=r_excel, column=col_excel)
                cell.fill = fills[st_code]
                cell.font = fonts[st_code]


def build_student_pivot_excel(
    pivot_df: pd.DataFrame,
    cell_status: dict[tuple[Any, str], str],
    sheet_name: str = "Student_Pivot_Status",
) -> bytes:
    """Export Student_Pivot_Status with pastel cell color coding and text-safe IDs."""
    safe_name = sheet_name[:31] or "Student_Pivot_Status"
    buf = io.BytesIO()
    if HAS_XLSXWRITER:
        with pd.ExcelWriter(
            buf,
            engine="xlsxwriter",
            engine_kwargs={
                "options": {
                    "strings_to_numbers": False,
                    "strings_to_urls": False,
                    "strings_to_formulas": False,
                }
            },
        ) as writer:
            wb = writer.book
            hdr_fmt = wb.add_format(
                {
                    "bold": True,
                    "bg_color": "#1E3A5F",
                    "font_color": "#FFFFFF",
                    "border": 1,
                    "align": "left",
                    "valign": "vcenter",
                }
            )
            txt_fmt = wb.add_format({"num_format": "@"})
            _write_pivot_sheet_xlsxwriter(writer, pivot_df, cell_status, safe_name, hdr_fmt, txt_fmt)
    else:
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            _write_pivot_sheet_openpyxl(writer, pivot_df, cell_status, safe_name)
    return buf.getvalue()


def build_single_sheet_excel(df: pd.DataFrame, sheet_name: str = "Sheet1") -> bytes:
    """Reusable exporter for filtered UI tables with text-protected IDs and frozen headers."""
    safe_name = sheet_name[:31] or "Sheet1"
    buf = io.BytesIO()
    if HAS_XLSXWRITER:
        with pd.ExcelWriter(
            buf,
            engine="xlsxwriter",
            engine_kwargs={
                "options": {
                    "strings_to_numbers": False,
                    "strings_to_urls": False,
                    "strings_to_formulas": False,
                }
            },
        ) as writer:
            wb = writer.book
            hdr_fmt = wb.add_format(
                {
                    "bold": True,
                    "bg_color": "#1E3A5F",
                    "font_color": "#FFFFFF",
                    "border": 1,
                    "align": "left",
                    "valign": "vcenter",
                }
            )
            txt_fmt = wb.add_format({"num_format": "@"})
            _apply_xlsxwriter_sheet(writer, df, safe_name, hdr_fmt, txt_fmt)
    else:
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            _apply_openpyxl_sheet(writer, df, safe_name)
    return buf.getvalue()


def build_pure_pending_excel(
    result: ReconciliationResult,
    result_filter: Literal["ALL", "PASS_ONLY", "FAIL_ONLY"] = "ALL",
    abc_filter: Literal[
        "ALL", "READY_ONLY", "READY_AND_RECOVERABLE", "MISSING_OR_INVALID_ONLY"
    ] = "ALL",
    auto_backfill_abc: bool = False,
    include_fail_to_pass_upgrades: bool = False,
    include_pass_corrections: bool = False,
    include_abc_only_updates: bool = False,
    drop_internal_cols: bool = True,
    trim_whitespace: bool = True,
) -> bytes:
    """Export pure pending SAP records in text-safe Excel preserving exact SAP layout."""
    export_df = filter_pending_sap(
        result,
        result_filter=result_filter,
        abc_filter=abc_filter,
        auto_backfill_abc=auto_backfill_abc,
        include_fail_to_pass_upgrades=include_fail_to_pass_upgrades,
        include_pass_corrections=include_pass_corrections,
        include_abc_only_updates=include_abc_only_updates,
        drop_internal_cols=drop_internal_cols,
        trim_whitespace=trim_whitespace,
    )
    return build_single_sheet_excel(export_df, sheet_name="Sheet1")


def _apply_conditional_badges(
    wb: Any,
    ws: Any,
    df: pd.DataFrame,
    col_name: str,
    rules: list[tuple[str, str, str]],
) -> None:
    if df.empty or col_name not in df.columns:
        return
    col_idx = list(df.columns).index(col_name)
    n_rows = len(df)
    for val, bg, fg in rules:
        fmt = wb.add_format({"bg_color": bg, "font_color": fg, "bold": True})
        ws.conditional_format(
            1,
            col_idx,
            n_rows,
            col_idx,
            {"type": "cell", "criteria": "==", "value": f'"{val}"', "format": fmt},
        )


def build_reconciliation_report_excel(result: ReconciliationResult) -> bytes:
    """Build styled multi-sheet reconciliation, ABC backfill, anomaly, M-status & student report."""
    buf = io.BytesIO()
    m = result.metrics

    summary_rows = [
        {"Category": "SAP Raw Input", "Metric": "Total SAP Raw Rows", "Count": m["sap_raw_rows"]},
        {"Category": "SAP Raw Input", "Metric": "SAP Duplicates / Superseded Rows Removed", "Count": m["sap_duplicates_or_superseded"]},
        {"Category": "SAP Raw Input", "Metric": "SAP Conflict Groups (Dup / Fail+Pass / Multi-Pass)", "Count": m["sap_conflict_groups"]},
        {"Category": "SAP Raw Input", "Metric": "SAP Canonical Unique Records", "Count": m["sap_dedup_rows"]},
        {"Category": "Email Reference", "Metric": "Total Email Data Rows", "Count": m["email_raw_rows"]},
        {"Category": "Email Reference", "Metric": "Email Duplicate Rows Removed", "Count": m["email_duplicates_removed"]},
        {"Category": "Email Reference", "Metric": "Email Canonical Unique Records", "Count": m["email_dedup_rows"]},
        {"Category": "DigiLocker Reconciliation", "Metric": "Matched Records (In Both SAP & Email)", "Count": m["matched_records"]},
        {"Category": "DigiLocker Reconciliation", "Metric": "Pure Pending Upload (In SAP, Not in Email)", "Count": m["pending_total"]},
        {"Category": "DigiLocker Reconciliation", "Metric": "  - Pending PASS Records", "Count": m["pending_pass"]},
        {"Category": "DigiLocker Reconciliation", "Metric": "  - Pending FAIL Records", "Count": m["pending_fail"]},
        {"Category": "ABC Readiness (Pending)", "Metric": "ABC Ready (Valid 12-Digit)", "Count": m["pending_abc_ready"]},
        {"Category": "ABC Readiness (Pending)", "Metric": "ABC Recoverable (Found in Other Semesters / Cleaned)", "Count": m["pending_abc_recoverable"]},
        {"Category": "ABC Readiness (Pending)", "Metric": "ABC Missing / Dummy / Invalid (Unrecoverable)", "Count": m["pending_abc_missing_or_invalid"]},
        {"Category": "M-Status Remediation", "Metric": "FAIL -> PASS Upgrades Needed (mrks_rec_status = M)", "Count": m["fail_to_pass_upgrades"]},
        {"Category": "M-Status Remediation", "Metric": "PASS -> PASS Critical Corrections (mrks_rec_status = M)", "Count": m["pass_critical_corrections"]},
        {"Category": "M-Status Remediation", "Metric": "ABC-Only Updates Needed (mrks_rec_status = M)", "Count": m["abc_only_updates"]},
        {"Category": "Email-Only Audit", "Metric": "Email-Only Records (In Email, Missing in SAP)", "Count": m["email_only_records"]},
        {"Category": "Email-Only Audit", "Metric": "  - Student Exists in SAP for Other Semesters", "Count": m["email_only_in_sap_other_sems"]},
        {"Category": "Cross-File Anomalies", "Metric": "Matched Rows with >= 1 Anomaly", "Count": m["matched_rows_with_anomalies"]},
        {"Category": "Cross-File Anomalies", "Metric": "Total Field-Level Anomalies", "Count": m["total_field_anomalies"]},
        {"Category": "Student-Wise Coverage", "Metric": "Total Distinct Students", "Count": m["total_students"]},
        {"Category": "Student-Wise Coverage", "Metric": "Complete (All SAP Sems Uploaded, 0 Anomalies)", "Count": m["students_complete"]},
        {"Category": "Student-Wise Coverage", "Metric": "Complete (With Anomalies / M-Corrections)", "Count": m["students_complete_with_anomalies"]},
        {"Category": "Student-Wise Coverage", "Metric": "Partial (Some Sems Uploaded, Some Pending)", "Count": m["students_partial"]},
        {"Category": "Student-Wise Coverage", "Metric": "Not Uploaded (0 SAP Sems in Email)", "Count": m["students_not_uploaded"]},
        {"Category": "Student-Wise Coverage", "Metric": "Email Only (No SAP Records Found)", "Count": m["students_email_only"]},
    ]
    summary_df = pd.DataFrame(summary_rows)

    desired_pending_keys = [
        "regn_no",
        "sem",
        "cname",
        "course_name",
        "admission_year",
        "academic_course_id",
        "result",
        "year",
        "month",
        "tot_mrks",
        "sgpa",
        "abc_account_id",
        "source_file",
    ]
    pending_preview_cols = [
        result.sap_col_map[k] for k in desired_pending_keys if k in result.sap_col_map
    ]
    if not pending_preview_cols:
        pending_preview_cols = result.sap_original_columns[:12]

    pending_summary_df = result.pure_pending_sap_df[
        pending_preview_cols
        + ["__abc_readiness__", "__recovered_abc_id__", "__had_sap_conflict__", "__sap_conflict_type__"]
    ].rename(
        columns={
            "__abc_readiness__": "ABC_READINESS",
            "__recovered_abc_id__": "RECOVERED_ABC_ID",
            "__had_sap_conflict__": "HAD_SAP_CONFLICT",
            "__sap_conflict_type__": "SAP_CONFLICT_TYPE",
        }
    )

    sheets: list[tuple[str, pd.DataFrame]] = [
        ("Executive_Summary", summary_df),
        ("Student_Wise_Report", result.student_report_df),
        ("Student_Pivot_Status", result.student_pivot_df),
        ("Pending_Upload_Summary", pending_summary_df),
        ("ABC_Backfill_Audit", result.abc_backfill_audit_df),
        ("SAP_Duplicates_Conflicts", result.sap_conflicts_df),
    ]
    if not result.email_conflicts_df.empty:
        sheets.append(("Email_Duplicates", result.email_conflicts_df))
    sheets.extend(
        [
            ("Anomaly_Column_Stats", result.column_anomaly_summary_df),
            ("Anomalies_By_Record", result.anomalies_wide_df),
            ("Anomalies_Detailed", result.anomalies_long_df),
            ("Fail_To_Pass_Upgrades", result.fail_to_pass_upgrades_df),
            ("Pass_Critical_Corrections", result.pass_corrections_df),
            ("ABC_Only_Updates", result.abc_only_updates_df),
            ("Email_Only_Records", result.email_only_df),
        ]
    )
    if result.config.include_format_only_anomalies and not result.format_only_df.empty:
        sheets.append(("Format_Only_Diffs", result.format_only_df))

    if HAS_XLSXWRITER:
        with pd.ExcelWriter(
            buf,
            engine="xlsxwriter",
            engine_kwargs={
                "options": {
                    "strings_to_numbers": False,
                    "strings_to_urls": False,
                    "strings_to_formulas": False,
                }
            },
        ) as writer:
            wb = writer.book
            header_fmt = wb.add_format(
                {
                    "bold": True,
                    "bg_color": "#1E3A5F",
                    "font_color": "#FFFFFF",
                    "border": 1,
                    "align": "left",
                    "valign": "vcenter",
                }
            )
            text_fmt = wb.add_format({"num_format": "@"})

            for sname, sdf in sheets:
                if sname == "Student_Pivot_Status":
                    _write_pivot_sheet_xlsxwriter(
                        writer,
                        sdf,
                        result.student_pivot_cell_status,
                        sname,
                        header_fmt,
                        text_fmt,
                    )
                else:
                    _apply_xlsxwriter_sheet(writer, sdf, sname, header_fmt, text_fmt)

            _apply_conditional_badges(
                wb,
                writer.sheets["Student_Wise_Report"],
                result.student_report_df,
                "Overall_Status",
                [
                    ("Complete", "#D1FAE5", "#065F46"),
                    ("Complete (With Anomalies)", "#FEF3C7", "#92400E"),
                    ("Partial", "#FEF3C7", "#92400E"),
                    ("Not Uploaded", "#FEE2E2", "#991B1B"),
                    ("Email Only", "#DBEAFE", "#1E40AF"),
                ],
            )
            abc_rules = [
                ("READY", "#D1FAE5", "#065F46"),
                ("RECOVERABLE", "#DBEAFE", "#1E40AF"),
                ("MISSING", "#FEE2E2", "#991B1B"),
                ("INVALID_OR_DUMMY", "#FEE2E2", "#991B1B"),
            ]
            _apply_conditional_badges(
                wb, writer.sheets["Pending_Upload_Summary"], pending_summary_df, "ABC_READINESS", abc_rules
            )
            _apply_conditional_badges(
                wb, writer.sheets["ABC_Backfill_Audit"], result.abc_backfill_audit_df, "ABC_READINESS", abc_rules
            )
            sev_rules = [
                ("CRITICAL", "#FEE2E2", "#991B1B"),
                ("ACADEMIC", "#FEF3C7", "#92400E"),
                ("METADATA", "#E2E8F0", "#334155"),
            ]
            _apply_conditional_badges(
                wb, writer.sheets["Anomalies_Detailed"], result.anomalies_long_df, "severity", sev_rules
            )
            _apply_conditional_badges(
                wb, writer.sheets["Anomaly_Column_Stats"], result.column_anomaly_summary_df, "severity", sev_rules
            )
    else:
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            for sname, sdf in sheets:
                if sname == "Student_Pivot_Status":
                    _write_pivot_sheet_openpyxl(
                        writer, sdf, result.student_pivot_cell_status, sname
                    )
                else:
                    _apply_openpyxl_sheet(writer, sdf, sname)

    return buf.getvalue()


# ==============================================================================
# SECTION 5: Interactive Streamlit UI
# ==============================================================================


def _clear_results() -> None:
    for k in (
        "dl_recon_result",
        "dl_recon_run_fingerprint",
        "dl_recon_pure_cache",
        "dl_recon_report_cache",
        "dl_recon_m_cache",
        "dl_recon_sap_filenames",
        "dl_recon_email_filenames",
        "dl_recon_run_timestamp",
    ):
        st.session_state.pop(k, None)


def _reset_all_session_state() -> None:
    """Clear all dl_recon_* session state keys for a completely fresh run."""
    for k in [key for key in list(st.session_state.keys()) if key.startswith("dl_recon_")]:
        st.session_state.pop(k, None)


def _compute_run_fingerprint(
    sap_uploads: list | None,
    email_uploads: list | None,
    config_tuple: tuple[Any, ...],
) -> str:
    h = hashlib.sha256()
    for f in sap_uploads or []:
        h.update(f"S:{f.name}:{f.size}|".encode("utf-8"))
    for f in email_uploads or []:
        h.update(f"E:{f.name}:{f.size}|".encode("utf-8"))
    h.update(repr(config_tuple).encode("utf-8"))
    return h.hexdigest()


def _render_preflight_box(
    uploads: list | None,
    expected_role: Literal["SAP", "EMAIL"],
    *,
    read_all_sheets: bool = False,
    key_cols: Iterable[str] = ("regn_no", "sem"),
) -> None:
    if not uploads:
        return
    try:
        entries = tuple((f.name, f.getvalue()) for f in uploads)
        info = _cached_preflight_inspect(entries, read_all_sheets=read_all_sheets)
        key_badges = (
            f"Keys: `regn_no` {'✓' if info['has_regn_no'] else '✗'}, "
            f"`sem` {'✓' if info['has_sem'] else '✗'}"
        )
        st.caption(
            f"**{info['file_count']} file(s)** • **{info['total_rows']:,} rows** • "
            f"**{info['max_cols']} cols** • {key_badges}"
        )
        if not info["has_regn_no"] or not info["has_sem"]:
            st.warning("Required default key column(s) (`regn_no` / `sem`) not found in uploaded headers.")

        multi_sheets = info.get("multi_sheet_files", [])
        if multi_sheets:
            details = "; ".join(
                f"`{item['filename']}` ({len(item['sheet_names'])} sheets: {', '.join(item['sheet_names'])})"
                for item in multi_sheets
            )
            if not read_all_sheets:
                st.warning(
                    f"Multi-sheet Excel file(s) detected: {details}. Only the first/active sheet is "
                    "currently being read — enable **Read all sheets in multi-sheet Excel files** "
                    "under Configuration if data spans multiple sheets."
                )
            else:
                st.info(f"Reading all sheets from multi-sheet Excel file(s): {details}.")

        key_set = set(key_cols)
        coll_cnt = info.get("course_collision_count", 0)
        if coll_cnt > 0 and "course_name" not in key_set and "academic_course_id" not in key_set:
            samples = ", ".join(
                f"({c['regn_no']}, Sem {c['sem']})" for c in info.get("course_collision_samples", [])
            )
            st.warning(
                f"Cross-program collision detected: **{coll_cnt:,}** `(regn_no, sem)` pair(s) appear "
                f"under multiple `course_name` values (e.g., {samples}). Consider adding `course_name` "
                "to **Matching Key Columns** under Configuration."
            )

        wrong_files = [
            fname
            for fname, role in info["file_roles"]
            if (expected_role == "SAP" and role == "EMAIL")
            or (expected_role == "EMAIL" and role == "SAP")
        ]
        if wrong_files:
            expected_label = "SAP Raw Data" if expected_role == "SAP" else "Email Reference Data"
            detected_label = "Email Reference" if expected_role == "SAP" else "SAP Raw Data"
            st.warning(
                f"Possible swapped file(s) in **{expected_label}** box: `{', '.join(wrong_files)}` "
                f"looks like **{detected_label}**."
            )
    except Exception as exc:
        st.warning(f"Preflight check could not read file preview: {exc}")


def _render_download_bar(df: pd.DataFrame, base_filename: str, key_prefix: str) -> None:
    if df.empty:
        return
    c1, c2, _ = st.columns([1.2, 1.3, 3.5])
    with c1:
        st.download_button(
            label="Download Filtered CSV",
            data=df.to_csv(index=False).encode("utf-8"),
            file_name=f"{base_filename}.csv",
            mime="text/csv",
            key=f"{key_prefix}_csv_dl",
            use_container_width=True,
        )
    with c2:
        st.download_button(
            label="Download Filtered Excel",
            data=build_single_sheet_excel(df, sheet_name=base_filename[:31]),
            file_name=f"{base_filename}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"{key_prefix}_xlsx_dl",
            use_container_width=True,
        )


def _render_m_status_upload_buttons(
    res: ReconciliationResult,
    raw_sap_df: pd.DataFrame,
    file_slug: str,
    key_prefix: str,
) -> None:
    if raw_sap_df.empty:
        return
    sap_names: list[str] = st.session_state.get("dl_recon_sap_filenames", [])
    prefix = infer_export_prefix(sap_names)
    date_tag = dt.date.today().strftime("%Y-%m-%d")
    base_name = f"{prefix}_M_Status_{file_slug}_{date_tag}"

    m_export_df = prepare_m_status_sap_export(res, raw_sap_df, auto_backfill_abc=True)
    m_cache = st.session_state.setdefault("dl_recon_m_cache", {})
    if key_prefix not in m_cache:
        m_cache[key_prefix] = build_single_sheet_excel(m_export_df, sheet_name="Sheet1")

    st.caption(
        f"Direct NAD / DigiLocker upload export (**{len(m_export_df):,} rows**, "
        "`MRKS_REC_STATUS = M`, recoverable ABC IDs backfilled):"
    )
    c1, c2, _ = st.columns([1.3, 1.4, 3.3])
    with c1:
        st.download_button(
            label="Download M-Status SAP CSV",
            data=build_pure_pending_csv(m_export_df),
            file_name=f"{base_name}.csv",
            mime="text/csv",
            key=f"{key_prefix}_m_csv_dl",
            type="primary",
            use_container_width=True,
        )
    with c2:
        st.download_button(
            label="Download M-Status SAP Excel",
            data=m_cache[key_prefix],
            file_name=f"{base_name}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"{key_prefix}_m_xlsx_dl",
            type="primary",
            use_container_width=True,
        )


def _render_overview_tab(
    res: ReconciliationResult,
    sap_filenames: list[str],
) -> None:
    m = res.metrics

    st.subheader("1. Upload & Deduplication Summary")
    a1, a2, a3, a4 = st.columns(4)
    a1.metric(
        "SAP Raw Rows",
        f"{m['sap_raw_rows']:,}",
        f"{m['sap_dedup_rows']:,} canonical",
        delta_color="off",
    )
    a2.metric(
        "SAP Conflicts Resolved",
        f"{m['sap_duplicates_or_superseded']:,}",
        f"{m['sap_conflict_groups']:,} key groups",
        delta_color="off",
    )
    a3.metric(
        "Email Reference Rows",
        f"{m['email_raw_rows']:,}",
        f"{m['matched_records']:,} matched | {m['email_only_records']:,} email-only",
        delta_color="off",
    )
    a4.metric(
        "Pure Pending Upload",
        f"{m['pending_total']:,}",
        f"PASS {m['pending_pass']:,} | FAIL {m['pending_fail']:,}",
        delta_color="off",
    )

    st.subheader("2. ABC Readiness & M-Status Remediation")
    b1, b2, b3, b4, b5 = st.columns(5)
    b1.metric("Pending ABC Ready", f"{m['pending_abc_ready']:,}", "Valid 12-digit ID", delta_color="off")
    b2.metric(
        "Pending ABC Recoverable",
        f"{m['pending_abc_recoverable']:,}",
        f"{m['pending_abc_missing_or_invalid']:,} unrecoverable",
        delta_color="off",
    )
    b3.metric(
        "FAIL → PASS Upgrades",
        f"{m['fail_to_pass_upgrades']:,}",
        "Need mrks_rec_status = M",
        delta_color="off",
    )
    b4.metric(
        "PASS Critical Corrections",
        f"{m['pass_critical_corrections']:,}",
        "Need mrks_rec_status = M",
        delta_color="off",
    )
    b5.metric(
        "ABC-Only Updates",
        f"{m.get('abc_only_updates', 0):,}",
        "Need mrks_rec_status = M",
        delta_color="off",
    )

    st.subheader("3. Student-Wise Coverage")
    s1, s2, s3, s4 = st.columns(4)
    s1.metric("Total Distinct Students", f"{m['total_students']:,}")
    s2.metric(
        "Complete (Clean)",
        f"{m['students_complete']:,}",
        f"+ {m['students_complete_with_anomalies']:,} with anomalies",
        delta_color="off",
    )
    s3.metric("Partial (Pending Sems)", f"{m['students_partial']:,}")
    s4.metric(
        "Not Uploaded (0 Sems)",
        f"{m['students_not_uploaded']:,}",
        f"{m['students_email_only']:,} email-only students",
        delta_color="off",
    )

    st.markdown("---")
    st.subheader("Download Reports & NAD-Ready SAP Export")

    col_dl1, col_dl2 = st.columns(2)
    with col_dl1:
        with st.container(border=True):
            st.markdown("#### Pure Pending SAP Upload Export")
            ec1, ec2 = st.columns(2)
            with ec1:
                result_choice = st.radio(
                    "Result Scope",
                    options=["Both PASS & FAIL (Default)", "PASS Only", "FAIL Only"],
                    index=0,
                    key="dl_recon_exp_res",
                )
            with ec2:
                abc_choice = st.selectbox(
                    "ABC Readiness Filter",
                    options=[
                        "All Pending Records",
                        "ABC Ready Only",
                        "ABC Ready + Recoverable",
                        "Missing / Invalid ABC Only",
                    ],
                    index=0,
                    key="dl_recon_exp_abc",
                )

            result_filter_code: Literal["ALL", "PASS_ONLY", "FAIL_ONLY"] = {
                "Both PASS & FAIL (Default)": "ALL",
                "PASS Only": "PASS_ONLY",
                "FAIL Only": "FAIL_ONLY",
            }[result_choice]
            abc_filter_code: Literal[
                "ALL", "READY_ONLY", "READY_AND_RECOVERABLE", "MISSING_OR_INVALID_ONLY"
            ] = {
                "All Pending Records": "ALL",
                "ABC Ready Only": "READY_ONLY",
                "ABC Ready + Recoverable": "READY_AND_RECOVERABLE",
                "Missing / Invalid ABC Only": "MISSING_OR_INVALID_ONLY",
            }[abc_choice]

            auto_backfill = st.checkbox(
                "Auto-backfill recoverable ABC IDs into ABC_ACCOUNT_ID",
                value=False,
                key="dl_recon_exp_backfill",
            )
            include_upgrades = st.checkbox(
                "Include FAIL→PASS upgrades (stamped MRKS_REC_STATUS = M)",
                value=False,
                key="dl_recon_exp_upgrades",
            )
            include_pass_corr = st.checkbox(
                "Include PASS→PASS critical corrections (stamped MRKS_REC_STATUS = M)",
                value=False,
                key="dl_recon_exp_pass_corr",
            )
            include_abc_only = st.checkbox(
                "Include ABC-only updates (stamped MRKS_REC_STATUS = M)",
                value=False,
                key="dl_recon_exp_abc_only",
            )
            drop_internal_cols = st.checkbox(
                "Drop internal merge columns (SOURCE_FILE, csv_filename)",
                value=True,
                key="dl_recon_exp_drop_internal",
            )
            trim_ws = st.checkbox(
                "Trim leading/trailing whitespace in exported cells",
                value=True,
                key="dl_recon_exp_trim_ws",
            )

            export_df = filter_pending_sap(
                res,
                result_filter=result_filter_code,
                abc_filter=abc_filter_code,
                auto_backfill_abc=auto_backfill,
                include_fail_to_pass_upgrades=include_upgrades,
                include_pass_corrections=include_pass_corr,
                include_abc_only_updates=include_abc_only,
                drop_internal_cols=drop_internal_cols,
                trim_whitespace=trim_ws,
            )
            pure_csv_name, pure_xlsx_name, report_xlsx_name = build_dynamic_filenames(
                sap_filenames, result_filter=result_filter_code, abc_filter=abc_filter_code
            )

            st.info(
                f"**Ready to export:** **{len(export_df):,} rows** × **{len(export_df.columns)} columns**"
            )

            pure_cache_key = (
                f"{result_filter_code}|{abc_filter_code}|{auto_backfill}|"
                f"{include_upgrades}|{include_pass_corr}|{include_abc_only}|{drop_internal_cols}|{trim_ws}"
            )
            cache_dict = st.session_state.setdefault("dl_recon_pure_cache", {})
            if pure_cache_key not in cache_dict:
                cache_dict[pure_cache_key] = build_single_sheet_excel(export_df, sheet_name="Sheet1")

            b_csv, b_xlsx = st.columns(2)
            with b_csv:
                st.download_button(
                    label="Download NAD-Ready CSV",
                    data=build_pure_pending_csv(export_df),
                    file_name=pure_csv_name,
                    mime="text/csv",
                    key="dl_recon_dl_pure_csv_btn",
                    type="primary",
                    use_container_width=True,
                )
            with b_xlsx:
                st.download_button(
                    label="Download Excel (.xlsx)",
                    data=cache_dict[pure_cache_key],
                    file_name=pure_xlsx_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="dl_recon_dl_pure_xlsx_btn",
                    type="primary",
                    use_container_width=True,
                )

    with col_dl2:
        with st.container(border=True):
            st.markdown("#### Full Reconciliation & Student-Wise Report")
            _, _, report_xlsx_name = build_dynamic_filenames(sap_filenames)
            st.info(
                f"**File:** `{report_xlsx_name}`\n\n"
                f"Includes sheets for:\n"
                f"- **Executive Summary**, **Student-Wise Matrix** & **Student Pivot Status** ({m['total_students']:,} students)\n"
                f"- **Pending Upload Summary** & **ABC Backfill Audit** ({len(res.abc_backfill_audit_df):,} flagged)\n"
                f"- **SAP Duplicates/Conflicts** & **Email Duplicates**\n"
                f"- **Column Stats**, **Anomalies (By Record & Detailed)**, "
                f"**FAIL→PASS Upgrades**, **PASS Critical Corrections**, **ABC-Only Updates**, & **Email-Only Records**"
            )
            if "dl_recon_report_cache" not in st.session_state:
                with st.spinner("Building styled multi-sheet Excel report..."):
                    st.session_state["dl_recon_report_cache"] = build_reconciliation_report_excel(res)
            st.download_button(
                label=f"Download {report_xlsx_name}",
                data=st.session_state["dl_recon_report_cache"],
                file_name=report_xlsx_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_recon_dl_rep_btn",
                type="primary",
                use_container_width=True,
            )


def _render_pending_tab(res: ReconciliationResult) -> None:
    st.subheader("Pure SAP Data Not Yet Uploaded on DigiLocker")
    df = res.pure_pending_sap_df.copy()

    regn_col = res.sap_col_map.get("regn_no")
    cname_col = res.sap_col_map.get("cname")
    course_col = res.sap_col_map.get("course_name")
    adm_col = res.sap_col_map.get("admission_year")
    year_col = res.sap_col_map.get("year")

    # Collapsible Pending Breakdown by Course & Admission Year triage table
    with st.expander("Pending Breakdown by Course & Admission Year", expanded=False):
        if course_col and adm_col and not df.empty:
            triage_df = pd.DataFrame(
                {
                    "Course Name": df[course_col],
                    "Admission Year": df[adm_col],
                    "Total Pending": 1,
                    "PASS": (df["__pending_result__"] == "PASS").astype(int),
                    "FAIL": (df["__pending_result__"] == "FAIL").astype(int),
                    "ABC Ready": (df["__abc_readiness__"] == "READY").astype(int),
                    "ABC Recoverable": (df["__abc_readiness__"] == "RECOVERABLE").astype(int),
                    "Missing / Dummy": df["__abc_readiness__"].isin(["MISSING", "INVALID_OR_DUMMY"]).astype(int),
                }
            )
            triage_grp = (
                triage_df.groupby(["Course Name", "Admission Year"], dropna=False)
                .sum()
                .reset_index()
                .sort_values(["Total Pending", "Course Name"], ascending=[False, True])
                .reset_index(drop=True)
            )
            st.dataframe(triage_grp, use_container_width=True, height=300)
        else:
            st.info("Course Name or Admission Year column not present in SAP Data.")

    f1, f2, f3, f4 = st.columns(4)
    with f1:
        sel_res = st.selectbox("Result", ["ALL", "PASS", "FAIL"], key="dl_recon_pend_res")
    with f2:
        abc_opts = ["ALL"] + sorted([a for a in df["__abc_readiness__"].unique() if a])
        sel_abc = st.selectbox("ABC Readiness", abc_opts, key="dl_recon_pend_abc")
    with f3:
        courses = ["ALL"] + sorted([c for c in df[course_col].unique() if c]) if course_col else ["ALL"]
        sel_course = st.selectbox("Course Name", courses, key="dl_recon_pend_course")
    with f4:
        adm_years = ["ALL"] + sorted([y for y in df[adm_col].unique() if y]) if adm_col else ["ALL"]
        sel_adm = st.selectbox("Admission Year", adm_years, key="dl_recon_pend_adm")

    g1, g2, g3 = st.columns([1.2, 1.5, 2.3])
    with g1:
        years = ["ALL"] + sorted([y for y in df[year_col].unique() if y]) if year_col else ["ALL"]
        sel_year = st.selectbox("Exam Year", years, key="dl_recon_pend_yr")
    with g2:
        only_conflicts = st.checkbox(
            "Only resolved SAP conflict rows", value=False, key="dl_recon_pend_conf"
        )
    with g3:
        search_q = st.text_input("Search Regn No or Name", key="dl_recon_pend_search")

    if sel_res != "ALL":
        df = df[df["__pending_result__"] == sel_res]
    if sel_abc != "ALL":
        df = df[df["__abc_readiness__"] == sel_abc]
    if course_col and sel_course != "ALL":
        df = df[df[course_col] == sel_course]
    if adm_col and sel_adm != "ALL":
        df = df[df[adm_col] == sel_adm]
    if year_col and sel_year != "ALL":
        df = df[df[year_col] == sel_year]
    if only_conflicts:
        df = df[df["__had_sap_conflict__"]]
    if search_q.strip():
        q = search_q.strip().lower()
        mask = pd.Series(False, index=df.index)
        if regn_col:
            mask |= df[regn_col].fillna("").astype(str).str.lower().str.contains(q, regex=False, na=False)
        if cname_col:
            mask |= df[cname_col].fillna("").astype(str).str.lower().str.contains(q, regex=False, na=False)
        df = df[mask]

    show_all_cols = st.checkbox(
        f"Show all original SAP columns ({len(res.sap_original_columns)} columns)",
        value=False,
        key="dl_recon_show_all_sap_cols",
    )
    st.caption(f"Showing **{min(len(df), 500):,}** of **{len(df):,}** matching pending records.")

    if show_all_cols:
        export_view = df[res.sap_original_columns]
    else:
        desired_keys = [
            "regn_no",
            "sem",
            "cname",
            "course_name",
            "admission_year",
            "result",
            "year",
            "month",
            "tot_mrks",
            "sgpa",
            "abc_account_id",
            "source_file",
        ]
        key_preview = [res.sap_col_map[k] for k in desired_keys if k in res.sap_col_map]
        if not key_preview:
            key_preview = res.sap_original_columns[:10]
        export_view = df[
            key_preview
            + ["__abc_readiness__", "__recovered_abc_id__", "__had_sap_conflict__", "__sap_conflict_type__"]
        ].rename(
            columns={
                "__abc_readiness__": "ABC_READINESS",
                "__recovered_abc_id__": "RECOVERED_ABC_ID",
                "__had_sap_conflict__": "HAD_SAP_CONFLICT",
                "__sap_conflict_type__": "SAP_CONFLICT_TYPE",
            }
        )

    st.dataframe(export_view.head(500), use_container_width=True, height=460)
    _render_download_bar(export_view, "Filtered_Pending_SAP", "dl_recon_tab2")


def _render_conflicts_tab(res: ReconciliationResult) -> None:
    st.subheader("Internal Duplicates, FAIL+PASS Pairs & Multiple PASS Conflicts")
    st.caption(
        "Automatic resolution priority: **PASS > FAIL**, then **latest (year, month)**, "
        "then **mrks_rec_status M > O** on identical periods."
    )

    has_email_dups = not res.email_conflicts_df.empty
    source_view = "SAP Conflicts"
    if has_email_dups:
        source_view = st.radio(
            "Select Conflict Source",
            options=[
                f"SAP Conflicts ({len(res.sap_conflicts_df):,})",
                f"Email Duplicates ({len(res.email_conflicts_df):,})",
            ],
            horizontal=True,
            key="dl_recon_conf_source_radio",
        )

    df = res.email_conflicts_df if source_view.startswith("Email Duplicates") else res.sap_conflicts_df
    if df.empty:
        st.success("No duplicate or conflicting keys found inside the uploaded SAP Data (and 0 Email duplicates).")
        return

    c1, c2, c3 = st.columns(3)
    with c1:
        ctypes = ["ALL"] + sorted(df["conflict_type"].fillna("").astype(str).unique().tolist())
        sel_ct = st.selectbox("Conflict Type", ctypes, key="dl_recon_conf_type")
    with c2:
        actions = ["ALL"] + sorted(df["action"].fillna("").astype(str).unique().tolist())
        sel_act = st.selectbox("Resolution Action", actions, key="dl_recon_conf_act")
    with c3:
        q = st.text_input("Search Regn No / Name", key="dl_recon_conf_search")

    view = df.copy()
    if sel_ct != "ALL":
        view = view[view["conflict_type"] == sel_ct]
    if sel_act != "ALL":
        view = view[view["action"] == sel_act]
    if q.strip():
        ql = q.strip().lower()
        view = view[
            view["regn_no"].fillna("").astype(str).str.lower().str.contains(ql, regex=False, na=False)
            | view["cname"].fillna("").astype(str).str.lower().str.contains(ql, regex=False, na=False)
        ]

    st.caption(f"Showing **{min(len(view), 1000):,}** of **{len(view):,}** conflict rows.")
    st.dataframe(view.head(1000), use_container_width=True, height=460)
    _render_download_bar(view, "Filtered_Conflicts", "dl_recon_tab3")


def _render_anomalies_tab(res: ReconciliationResult) -> None:
    st.subheader("Cross-File Anomalies, M-Status Remediation & Email-Only Audit")

    if not res.fail_to_pass_upgrades_df.empty:
        with st.expander(
            f"FAIL → PASS Upgrades Needed ({len(res.fail_to_pass_upgrades_df):,} records — upload with MRKS_REC_STATUS = M)",
            expanded=True,
        ):
            st.dataframe(res.fail_to_pass_upgrades_df, use_container_width=True)
            _render_download_bar(res.fail_to_pass_upgrades_df, "Fail_To_Pass_Upgrades", "dl_recon_ftp")
            _render_m_status_upload_buttons(
                res, res.fail_to_pass_raw_sap_df, "Fail_To_Pass", "dl_recon_ftp_m"
            )

    if not res.pass_corrections_df.empty:
        with st.expander(
            f"PASS → PASS Critical Corrections ({len(res.pass_corrections_df):,} records — upload with MRKS_REC_STATUS = M)",
            expanded=False,
        ):
            st.dataframe(res.pass_corrections_df, use_container_width=True)
            _render_download_bar(
                res.pass_corrections_df, "Pass_Critical_Corrections", "dl_recon_pcc"
            )
            _render_m_status_upload_buttons(
                res, res.pass_corrections_raw_sap_df, "Pass_Critical_Corrections", "dl_recon_pcc_m"
            )

    if not res.abc_only_updates_df.empty:
        with st.expander(
            f"ABC-Only Updates ({len(res.abc_only_updates_df):,} records — upload with MRKS_REC_STATUS = M)",
            expanded=False,
        ):
            st.dataframe(res.abc_only_updates_df, use_container_width=True)
            _render_download_bar(res.abc_only_updates_df, "ABC_Only_Updates", "dl_recon_abco")
            _render_m_status_upload_buttons(
                res, res.abc_only_updates_raw_sap_df, "ABC_Only_Updates", "dl_recon_abco_m"
            )

    col_left, col_right = st.columns([1.15, 1.85])
    with col_left:
        st.markdown("**Column-Wise Anomaly Breakdown**")
        st.dataframe(res.column_anomaly_summary_df, use_container_width=True, height=360)

    with col_right:
        st.markdown("**Email-Only Records (In DigiLocker Reference, Missing in SAP)**")
        in_other = res.metrics["email_only_in_sap_other_sems"]
        st.caption(
            f"Total Email-Only records: **{len(res.email_only_df):,}** "
            f"(**{in_other:,}** exist in SAP for other semesters; "
            f"**{len(res.email_only_df) - in_other:,}** completely absent from SAP)."
        )
        st.dataframe(res.email_only_df.head(500), use_container_width=True, height=330)

    st.markdown("---")
    st.markdown("**Detailed Field-Level Anomalies**")
    adf = res.anomalies_long_df
    if adf.empty:
        st.success("No cross-file anomalies detected for the selected columns and normalization settings.")
        return

    f1, f2, f3, f4 = st.columns(4)
    with f1:
        sel_col = st.selectbox(
            "Column",
            ["ALL"] + sorted(adf["column"].fillna("").astype(str).unique().tolist()),
            key="dl_recon_anom_col",
        )
    with f2:
        sel_sev = st.selectbox(
            "Severity",
            ["ALL"] + sorted(adf["severity"].fillna("").astype(str).unique().tolist()),
            key="dl_recon_anom_sev",
        )
    with f3:
        sel_type = st.selectbox(
            "Anomaly Type",
            ["ALL"] + sorted(adf["anomaly_type"].fillna("").astype(str).unique().tolist()),
            key="dl_recon_anom_type",
        )
    with f4:
        q = st.text_input("Search Regn No / Name", key="dl_recon_anom_search")

    view = adf.copy()
    if sel_col != "ALL":
        view = view[view["column"] == sel_col]
    if sel_sev != "ALL":
        view = view[view["severity"] == sel_sev]
    if sel_type != "ALL":
        view = view[view["anomaly_type"] == sel_type]
    if q.strip():
        ql = q.strip().lower()
        view = view[
            view["regn_no"].fillna("").astype(str).str.lower().str.contains(ql, regex=False, na=False)
            | view["cname"].fillna("").astype(str).str.lower().str.contains(ql, regex=False, na=False)
        ]

    st.caption(f"Showing **{min(len(view), 1000):,}** of **{len(view):,}** matching anomalies.")
    st.dataframe(view.head(1000), use_container_width=True, height=440)
    _render_download_bar(view, "Filtered_Anomalies", "dl_recon_tab4")


def _style_pivot_preview(df: pd.DataFrame, cell_status: dict[tuple[str, str], str]) -> pd.DataFrame:
    """Return CSS style matrix matching Student_Pivot_Status pastel color coding."""
    styles = pd.DataFrame("", index=df.index, columns=df.columns)
    sem_cols = [
        c
        for c in df.columns
        if c not in {"regn_no", "cname", "course_name", "admission_year", "abc_account_id", "Remark"}
    ]
    regns = df["regn_no"].tolist() if "regn_no" in df.columns else [""] * len(df)
    for row_pos, (idx, regn_raw) in enumerate(zip(df.index, regns)):
        regn_val = clean_str(regn_raw)
        for scol in sem_cols:
            st_code = cell_status.get((regn_val, scol), "EMPTY")
            if st_code in PIVOT_STYLE_MAP:
                bg, fg = PIVOT_STYLE_MAP[st_code]
                styles.at[idx, scol] = f"background-color: {bg}; color: {fg};"
    return styles


def _render_student_report_tab(res: ReconciliationResult) -> None:
    st.subheader("Complete Student-Wise Semester Coverage & 360° Inspector")
    sdf = res.student_report_df
    if sdf.empty:
        st.info("No student records found.")
        return

    view_mode = st.radio(
        "Report View",
        options=[
            "Student_Wise_Report (Coverage Status)",
            "Student_Pivot_Status (Exam Attempts & Pastel Color Matrix)",
        ],
        horizontal=True,
        key="dl_recon_stu_view_mode",
    )

    f1, f2, f3, f4 = st.columns([1.3, 1.5, 1.2, 2])
    with f1:
        if view_mode.startswith("Student_Wise_Report"):
            statuses = ["ALL"] + sorted(sdf["Overall_Status"].fillna("").astype(str).unique().tolist())
            sel_status = st.selectbox("Overall Status", statuses, key="dl_recon_stu_status")
            sel_pivot_remark = "ALL"
        else:
            sel_status = "ALL"
            sel_pivot_remark = st.selectbox(
                "Remark Filter",
                ["ALL", "All OK", "Has Discrepancies"],
                key="dl_recon_pivot_remark_flt",
            )
    with f2:
        courses = ["ALL"] + sorted([c for c in sdf["course_name"].fillna("").astype(str).unique().tolist() if c])
        sel_course = st.selectbox("Course Name", courses, key="dl_recon_stu_course")
    with f3:
        adm_years = ["ALL"] + sorted(
            [y for y in sdf["admission_year"].fillna("").astype(str).unique().tolist() if y]
        )
        sel_adm = st.selectbox("Admission Year", adm_years, key="dl_recon_stu_adm")
    with f4:
        q = st.text_input("Search Regn No / Student Name / ABC ID", key="dl_recon_stu_search")

    def _apply_common_filters(base_df: pd.DataFrame) -> pd.DataFrame:
        out = base_df.copy()
        if sel_course != "ALL":
            out = out[out["course_name"] == sel_course]
        if sel_adm != "ALL":
            out = out[out["admission_year"] == sel_adm]
        if q.strip():
            ql = q.strip().lower()
            out = out[
                out["regn_no"].fillna("").astype(str).str.lower().str.contains(ql, regex=False, na=False)
                | out["cname"].fillna("").astype(str).str.lower().str.contains(ql, regex=False, na=False)
                | out["abc_account_id"].fillna("").astype(str).str.lower().str.contains(ql, regex=False, na=False)
            ]
        return out

    view = _apply_common_filters(sdf)
    if sel_status != "ALL":
        view = view[view["Overall_Status"] == sel_status]

    if view_mode.startswith("Student_Wise_Report"):
        st.caption(f"Showing **{min(len(view), 1000):,}** of **{len(view):,}** students.")
        st.dataframe(view.head(1000), use_container_width=True, height=450)
        _render_download_bar(view, "Filtered_Student_Report", "dl_recon_tab5")
    else:
        pivot_view = _apply_common_filters(res.student_pivot_df)
        if sel_pivot_remark == "All OK":
            pivot_view = pivot_view[pivot_view["Remark"] == "All OK"]
        elif sel_pivot_remark == "Has Discrepancies":
            pivot_view = pivot_view[pivot_view["Remark"] != "All OK"]

        st.caption(
            "Color legend: **Green** = Uploaded OK | **Red** = Wrong/Outdated Data in DigiLocker | "
            "**Yellow** = Pending in SAP | **Orange** = Email-Only | `-` = Not applicable"
        )
        preview_limit = 500
        st.caption(
            f"Showing **{min(len(pivot_view), preview_limit):,}** of **{len(pivot_view):,}** students "
            "(full matching dataset is included in CSV / styled Excel downloads below)."
        )
        preview_slice = pivot_view.head(preview_limit)
        if preview_slice.empty:
            st.dataframe(preview_slice, use_container_width=True, height=450)
        else:
            st.dataframe(
                preview_slice.style.apply(
                    _style_pivot_preview, cell_status=res.student_pivot_cell_status, axis=None
                ),
                use_container_width=True,
                height=450,
            )

        if not pivot_view.empty:
            p1, p2, _ = st.columns([1.2, 1.4, 3.4])
            with p1:
                st.download_button(
                    label="Download Pivot CSV",
                    data=pivot_view.to_csv(index=False).encode("utf-8"),
                    file_name="Student_Pivot_Status.csv",
                    mime="text/csv",
                    key="dl_recon_pivot_csv_dl",
                    use_container_width=True,
                )
            with p2:
                st.download_button(
                    label="Download Styled Pivot Excel",
                    data=build_student_pivot_excel(pivot_view, res.student_pivot_cell_status),
                    file_name="Student_Pivot_Status.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="dl_recon_pivot_xlsx_dl",
                    type="primary",
                    use_container_width=True,
                )

    # Single-Student 360° Inspector (searches across ALL rows in res.student_report_df)
    st.markdown("---")
    st.markdown("#### Single-Student 360° Inspector")
    inspector_q = st.text_input(
        "Search Any Student Across Entire Dataset (Regn No, Name, or ABC ID)",
        key="dl_recon_360_search",
        placeholder="Type full or partial Regn No, Student Name, or ABC ID to search all students...",
    )
    if inspector_q.strip():
        iql = inspector_q.strip().lower()
        inspector_candidates = sdf[
            sdf["regn_no"].fillna("").astype(str).str.lower().str.contains(iql, regex=False, na=False)
            | sdf["cname"].fillna("").astype(str).str.lower().str.contains(iql, regex=False, na=False)
            | sdf["abc_account_id"].fillna("").astype(str).str.lower().str.contains(iql, regex=False, na=False)
        ]
        st.caption(
            f"Found **{len(inspector_candidates):,}** matching student(s) across all **{len(sdf):,}** records."
        )
    else:
        inspector_candidates = view if not view.empty else sdf
        st.caption(
            f"Showing up to 500 students from the current table view ({len(inspector_candidates):,} total). "
            "Use the search box above to jump to any student in the full dataset."
        )

    if inspector_candidates.empty:
        st.info("No students match the 360° Inspector search.")
        return

    candidates = inspector_candidates.head(500)
    options_labels = [
        f"{r} — {n}" if n else str(r)
        for r, n in zip(candidates["regn_no"].tolist(), candidates["cname"].tolist())
    ]
    sel_label = st.selectbox(
        "Select Student",
        options=options_labels,
        key="dl_recon_360_select",
    )
    sel_regn = sel_label.split(" — ", 1)[0].strip()

    stu_meta = sdf[sdf["regn_no"] == sel_regn].iloc[0]
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Overall Status", str(stu_meta["Overall_Status"]))
    m2.metric("Course", str(stu_meta["course_name"]) or "-")
    m3.metric("ABC Account ID", str(stu_meta["abc_account_id"]) or "Missing")
    m4.metric("Pending / Anomalies", f"{stu_meta['Pending_SAP_Sems']} / {stu_meta['Cross_File_Anomalies']}")

    timeline_rows = res.student_timeline_map.get(sel_regn, [])
    if timeline_rows:
        st.markdown("**Semester-by-Semester Timeline (SAP vs. Email)**")
        st.dataframe(pd.DataFrame(timeline_rows), use_container_width=True)

    if not res.anomalies_long_df.empty:
        stu_anom = res.anomalies_long_df[res.anomalies_long_df["regn_no"] == sel_regn]
        if not stu_anom.empty:
            st.markdown("**Field-Level Anomalies for This Student**")
            st.dataframe(stu_anom, use_container_width=True)

    if not res.sap_conflicts_df.empty:
        stu_conf = res.sap_conflicts_df[res.sap_conflicts_df["regn_no"] == sel_regn]
        if not stu_conf.empty:
            st.markdown("**SAP Duplicate / Conflict Audit for This Student**")
            st.dataframe(stu_conf, use_container_width=True)


def _render_tool_guide() -> None:
    st.markdown("---")
    with st.expander("Tool Guide & Help — How to Use This App (Simple Guide)", expanded=False):
        st.markdown(
            """
### What This Tool Does
This tool compares your college's **SAP Exam Data** against the **DigiLocker Email Reference Data** (the official uploaded records received from NAD/DigiLocker) so you can immediately see:
- Which student semester records are **still pending upload** (`Pure Pending`),
- Which uploaded records have **errors, outdated FAIL status, or missing ABC IDs** that need correction (`MRKS_REC_STATUS = M`),
- Which pending records have **missing ABC IDs that can be auto-filled** from other semesters, and
- A complete **semester-by-semester status matrix** for every student.

---

### 5-Step Usage Guide
1. **Upload SAP Raw Data File(s):** Drop one or more `.xlsx` / `.csv` files exported from SAP into Box 1. Check the preflight summary below the box for row counts, sheet alerts, or cross-program warnings.
2. **Upload Email Reference File(s):** Drop the merged DigiLocker Email reference `.xlsx` / `.csv` file(s) into Box 2.
3. **Review Configuration (Optional):** Open *Matching Key & Anomaly Configuration* if you have multi-sheet Excel workbooks or students enrolled in multiple programs (`course_name`).
4. **Click "Run DigiLocker Reconciliation":** Wait a few seconds while duplicates are resolved, ABC IDs are checked, and cross-file comparisons finish.
5. **Download & Inspect:** Use **Tab 1 (Overview & Exports)** for one-click Pure Pending CSV/Excel and the Full Multi-Sheet Report, **Tab 4** for direct `MRKS_REC_STATUS = M` upload files, and **Tab 5** to inspect individual students or download the color-coded `Student_Pivot_Status` sheet.

---

### Plain-English Glossary
- **Pure Pending:** Student semester records that exist in SAP Data but are **not yet present** in DigiLocker Email Reference Data. These are ready for fresh upload (`MRKS_REC_STATUS = O` / SAP default).
- **ABC Readiness (`READY` / `RECOVERABLE` / `MISSING`):**
  - **READY:** Has a valid 12-digit ABC (APAAR) ID on that row.
  - **RECOVERABLE:** Missing, dummy (`0` / `0000...`), or hyphen/space-formatted (`0224-0937-4403`) on that row, **but** a valid 12-digit ID was found for the same student in another semester/file or cleaned automatically. Enable *Auto-backfill recoverable ABC IDs* before exporting to fill them in!
  - **MISSING / INVALID_OR_DUMMY:** No valid 12-digit ABC ID could be found anywhere for that student in the uploaded batch.
- **`MRKS_REC_STATUS = M` (Modified Upload):** DigiLocker requires `MRKS_REC_STATUS = M` whenever you re-upload a record to replace existing DigiLocker data:
  - **FAIL → PASS Upgrades:** DigiLocker currently has `FAIL`, but SAP now has `PASS` (e.g., student cleared ATKT/re-exam).
  - **PASS → PASS Critical Corrections:** Both sides say `PASS`, but marks, SGPA, CGPA, grade, credits, name, DOB, gender, course ID, or ABC ID differ.
  - **ABC-Only Updates:** DigiLocker is missing a valid 12-digit ABC ID (or has a dummy/invalid ID), while SAP or another semester has the student's valid 12-digit ABC ID.
- **Email-Only:** Records found in DigiLocker Email Data that do **not** appear in the uploaded SAP files (often older batches or already-graduated semesters not included in the current SAP export).
- **Automatic Duplicate Resolution Priority:** When the same matching key (e.g., `regn_no + sem`) appears multiple times inside SAP (such as ATKT re-attempts or duplicate exports), the tool automatically keeps one canonical row in this exact priority order:
  1. **PASS** beats other results, which beat **FAIL**,
  2. **Latest Exam Year & Month** (e.g., `MAR 2026` beats `DEC 2025`),
  3. On the same exam period: **`MRKS_REC_STATUS = M`** beats **`O`**, which beats blank,
  4. Higher total marks → higher SGPA → more populated columns.

---

### `Student_Pivot_Status` Sheet Color Legend
- 🟩 **Pastel Green (`#D1FAE5`):** Uploaded OK — matched in both SAP and DigiLocker with zero anomalies and no `M`-status updates needed.
- 🟥 **Pastel Red (`#FEE2E2`):** Wrong / Outdated Data in DigiLocker — needs `MRKS_REC_STATUS = M` (`FAIL→PASS` upgrade, `PASS→PASS` critical correction, `ABC-Only` update) or has a field-level mismatch. Check the **Remark** column for exact details.
- 🟨 **Pastel Yellow (`#FEF3C7`):** Pending in SAP — available in SAP, not yet uploaded to DigiLocker.
- 🟧 **Pastel Orange (`#FFEDD5`):** Email-Only — present in DigiLocker Email Data, missing from the uploaded SAP files.
- **Uncolored (`-`):** No record for that semester in either file.
- *Note:* When a student has multiple attempts for a semester (e.g., ATKT), the canonical/latest attempt is shown first, with earlier attempts in brackets — e.g. `MAR 2026 (DEC 2025, MAY 2025)`.

---

### Practical Tips
- **Leading-Zero Safety:** IDs like `regn_no` and `abc_account_id` (e.g., `022409374403`) are read as strict text and exported with Excel's `@` text format (`strings_to_numbers=False`) so Excel never strips leading zeros or converts them to scientific notation (`2.24E+11`).
- **Multi-Sheet Excel Files:** By default, only the first sheet of an `.xlsx` file is loaded. If your workbook splits semesters or courses across tabs, tick **Read all sheets in multi-sheet Excel files** under Configuration.
- **UG / PG Same `regn_no` (`course_name` Key):** If the preflight check warns that the same `(regn_no, sem)` appears under multiple courses, add `course_name` to **Matching Key Columns** so UG and PG records don't collide. Parenthesis spacing differences like `B.M.S.(CBCGS)` vs `B.M.S. (CBCGS)` are automatically ignored during matching while preserving your original display text.
- **Fresh Start:** Click **Reset / Clear Results** next to the Run button anytime to clear all cached files, filters, and outputs before starting a new batch.
            """
        )


def show() -> None:
    """Entry point rendered inside SVKM ExamAutomationSuite (or standalone main)."""
    st.title("DigiLocker Reconciliation")
    st.caption(
        "Upload raw SAP Data and DigiLocker Email Reference files to deduplicate SAP conflicts "
        "(duplicates, FAIL+PASS pairs, and multiple PASS cycles keeping latest Year/Month/M-status), "
        "recover missing ABC IDs, export NAD-ready pending SAP files, and inspect a 360° student matrix."
    )
    st.markdown("---")

    active_read_all_sheets = bool(st.session_state.get("dl_recon_read_all_sheets", False))
    active_key_cols = st.session_state.get("dl_recon_key_cols", ["regn_no", "sem"])

    up_col1, up_col2 = st.columns(2)
    with up_col1:
        with st.container(border=True):
            sap_uploads = st.file_uploader(
                "1. Upload SAP Raw Data File(s) (.xlsx / .csv)",
                type=["xlsx", "xlsm", "xls", "csv"],
                accept_multiple_files=True,
                key="dl_recon_sap_uploader",
            )
            _render_preflight_box(
                sap_uploads,
                expected_role="SAP",
                read_all_sheets=active_read_all_sheets,
                key_cols=active_key_cols,
            )
    with up_col2:
        with st.container(border=True):
            email_uploads = st.file_uploader(
                "2. Upload Email Reference Data File(s) (.xlsx / .csv)",
                type=["xlsx", "xlsm", "xls", "csv"],
                accept_multiple_files=True,
                key="dl_recon_email_uploader",
            )
            _render_preflight_box(
                email_uploads,
                expected_role="EMAIL",
                read_all_sheets=active_read_all_sheets,
                key_cols=active_key_cols,
            )

    with st.expander("Matching Key & Anomaly Configuration", expanded=False):
        cfg_c1, cfg_c2 = st.columns(2)
        with cfg_c1:
            key_cols = st.multiselect(
                "Matching Key Columns",
                options=[
                    "regn_no",
                    "sem",
                    "course_name",
                    "academic_course_id",
                    "exam_type",
                    "year",
                    "month",
                    "rroll",
                ],
                default=["regn_no", "sem"],
                key="dl_recon_key_cols",
            )
            risky_keys = [k for k in key_cols if k in {"academic_course_id", "year", "month"}]
            if risky_keys:
                st.warning(
                    f"Including `{', '.join(risky_keys)}` in the matching key can split ATKT re-attempts, "
                    "multiple exam cycles, or CBCGS→NEP `academic_course_id` transitions across SAP and Email."
                )

            scope_mode = st.radio(
                "Columns to Compare for Anomalies",
                options=["Preset Group", "All Common Columns", "Custom Columns"],
                index=0,
                key="dl_recon_scope_mode",
            )
            selected_preset = None
            custom_cols_input = ""
            if scope_mode == "Preset Group":
                selected_preset = st.selectbox(
                    "Select Preset", list(COLUMN_PRESETS.keys()), index=0, key="dl_recon_preset_sel"
                )
            elif scope_mode == "Custom Columns":
                custom_cols_input = st.text_input(
                    "Comma-separated column names (lowercase)",
                    value=", ".join(HIGH_SIGNAL_COLUMNS),
                    key="dl_recon_custom_cols",
                )

            read_all_sheets_opt = st.checkbox(
                "Read all sheets in multi-sheet Excel files",
                value=False,
                key="dl_recon_read_all_sheets",
            )

        with cfg_c2:
            norm_text_opt = st.checkbox("Normalize whitespace & casing", value=True, key="dl_recon_norm_txt")
            ignore_punct_opt = st.checkbox(
                "Ignore name punctuation (MRS.SAROJ == MRSSAROJ)", value=True, key="dl_recon_norm_punct"
            )
            norm_num_opt = st.checkbox(
                "Normalize numeric precision (471.00 == 471)", value=True, key="dl_recon_norm_num"
            )
            norm_month_opt = st.checkbox(
                "Normalize months (12 == DECEMBER == December)", value=True, key="dl_recon_norm_mon"
            )
            norm_date_opt = st.checkbox(
                "Normalize dates (DD/MM/YYYY == YYYY-MM-DD)", value=True, key="dl_recon_norm_dt"
            )
            include_null_gaps = st.checkbox(
                "Flag one-side missing/blank fields (schema blanks)", value=False, key="dl_recon_inc_nulls"
            )
            include_fmt_only = st.checkbox(
                "Record format-only differences in report", value=False, key="dl_recon_inc_fmt"
            )

    current_fp = _compute_run_fingerprint(
        sap_uploads,
        email_uploads,
        (
            tuple(key_cols),
            scope_mode,
            selected_preset,
            custom_cols_input,
            read_all_sheets_opt,
            norm_text_opt,
            ignore_punct_opt,
            norm_num_opt,
            norm_month_opt,
            norm_date_opt,
            include_null_gaps,
            include_fmt_only,
        ),
    )

    btn_col1, btn_col2 = st.columns([3.5, 1.0])
    with btn_col1:
        run_clicked = st.button(
            "Run DigiLocker Reconciliation",
            type="primary",
            use_container_width=True,
            key="dl_recon_run_btn",
        )
    with btn_col2:
        reset_clicked = st.button(
            "Reset / Clear Results",
            use_container_width=True,
            key="dl_recon_reset_btn",
        )

    if reset_clicked:
        _reset_all_session_state()
        st.rerun()

    run_failed = False
    if run_clicked:
        if not sap_uploads or not email_uploads:
            st.error("Please upload at least one SAP Data file and one Email Reference Data file.")
            run_failed = True
        elif not key_cols:
            st.error("Please select at least one matching key column.")
            run_failed = True
        else:
            _clear_results()
            try:
                with st.spinner("Loading Excel/CSV files and running reconciliation..."):
                    sap_payloads = [(f.name, f.getvalue()) for f in sap_uploads]
                    email_payloads = [(f.name, f.getvalue()) for f in email_uploads]
                    sap_ds = load_uploaded_files(sap_payloads, read_all_sheets=read_all_sheets_opt)
                    email_ds = load_uploaded_files(email_payloads, read_all_sheets=read_all_sheets_opt)

                    if scope_mode == "Preset Group" and selected_preset:
                        compare_cols: list[str] | None = list(COLUMN_PRESETS[selected_preset])
                    elif scope_mode == "Custom Columns":
                        compare_cols = [
                            c.strip().lower().replace(" ", "_")
                            for c in custom_cols_input.split(",")
                            if c.strip()
                        ]
                    else:
                        compare_cols = [
                            c
                            for c in sap_ds.col_map
                            if c in email_ds.col_map and c not in DEFAULT_METADATA_COLUMNS
                        ]

                    cfg = ReconciliationConfig(
                        key_columns=key_cols,
                        compare_columns=compare_cols,
                        normalize_text=norm_text_opt,
                        normalize_numbers=norm_num_opt,
                        normalize_months=norm_month_opt,
                        normalize_dates=norm_date_opt,
                        ignore_name_punctuation=ignore_punct_opt,
                        include_one_side_null_anomalies=include_null_gaps,
                        include_format_only_anomalies=include_fmt_only,
                    )
                    result = run_reconciliation(sap_ds, email_ds, config=cfg)
                    st.session_state["dl_recon_result"] = result
                    st.session_state["dl_recon_run_fingerprint"] = current_fp
                    st.session_state["dl_recon_sap_filenames"] = [f.name for f in sap_uploads]
                    st.session_state["dl_recon_email_filenames"] = [f.name for f in email_uploads]
                    st.session_state["dl_recon_run_timestamp"] = dt.datetime.now().strftime(
                        "%Y-%m-%d %H:%M:%S"
                    )
            except Exception as exc:
                st.error(f"Reconciliation failed: {exc}")
                run_failed = True

    if not run_failed:
        if "dl_recon_result" not in st.session_state:
            st.info(
                "Upload both SAP Data and Email Reference Data files above, then click **Run DigiLocker Reconciliation**."
            )
        else:
            if st.session_state.get("dl_recon_run_fingerprint") != current_fp:
                st.warning(
                    "Uploaded files or configuration settings have changed since the last run. "
                    "Click **Run DigiLocker Reconciliation** to refresh results."
                )

            res: ReconciliationResult = st.session_state["dl_recon_result"]
            sap_names: list[str] = st.session_state.get("dl_recon_sap_filenames", [])
            email_names: list[str] = st.session_state.get("dl_recon_email_filenames", [])
            run_ts: str = st.session_state.get("dl_recon_run_timestamp", "")

            st.caption(
                f"**Active Batch:** SAP: `{', '.join(sap_names) or '-'}` | "
                f"Email: `{', '.join(email_names) or '-'}` | "
                f"Run at: `{run_ts or '-'}`"
            )

            tab1, tab2, tab3, tab4, tab5 = st.tabs(
                [
                    "Overview & Exports",
                    f"Pure Pending Upload ({res.metrics['pending_total']:,})",
                    f"Duplicates & Conflicts ({len(res.sap_conflicts_df) + len(res.email_conflicts_df):,})",
                    f"Anomalies & M-Status ({res.metrics['total_field_anomalies']:,})",
                    f"Student-Wise Report ({res.metrics['total_students']:,})",
                ]
            )

            with tab1:
                _render_overview_tab(res, sap_names)
            with tab2:
                _render_pending_tab(res)
            with tab3:
                _render_conflicts_tab(res)
            with tab4:
                _render_anomalies_tab(res)
            with tab5:
                _render_student_report_tab(res)

    _render_tool_guide()


if __name__ == "__main__":
    st.set_page_config(
        page_title="DigiLocker Reconciliation",
        page_icon="🛡️",
        layout="wide",
        initial_sidebar_state="collapsed",
    )
    show()
