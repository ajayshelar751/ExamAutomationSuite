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

    if col_lower in PROTECTED_ID_COLUMNS:
        # Never run numeric normalization on protected IDs (preserves leading zeros)
        return norm_text(raw, case_insensitive=True) if normalize_text else raw

    if col_lower == "sem":
        return norm_sem_key(raw)

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


def _read_excel_calamine(source: str | Path | bytes, sheet_name: str | None = None) -> pd.DataFrame:
    if isinstance(source, (str, Path)):
        wb = CalamineWorkbook.from_path(str(source))
    else:
        wb = CalamineWorkbook.from_filelike(io.BytesIO(source))

    target_sheet = sheet_name if sheet_name and sheet_name in wb.sheet_names else wb.sheet_names[0]
    rows = wb.get_sheet_by_name(target_sheet).to_python(skip_empty_area=False)
    if not rows:
        return pd.DataFrame()

    headers = _dedupe_headers(rows[0])
    n_cols = len(headers)
    data_rows: list[list[str]] = []
    for row in rows[1:]:
        padded = [_format_cell_as_str(row[i]) if i < len(row) else "" for i in range(n_cols)]
        if any(cell != "" for cell in padded):
            data_rows.append(padded)

    return pd.DataFrame(data_rows, columns=headers, dtype=str)


def _read_excel_openpyxl(source: str | Path | bytes, sheet_name: str | None = None) -> pd.DataFrame:
    if isinstance(source, (str, Path)):
        wb = openpyxl.load_workbook(source, read_only=True, data_only=True)
    else:
        wb = openpyxl.load_workbook(io.BytesIO(source), read_only=True, data_only=True)

    try:
        ws = wb[sheet_name] if sheet_name and sheet_name in wb.sheetnames else wb.active
        it = ws.iter_rows(values_only=True)
        try:
            first_row = next(it)
        except StopIteration:
            return pd.DataFrame()

        headers = _dedupe_headers(list(first_row))
        n_cols = len(headers)
        data_rows: list[list[str]] = []
        for row in it:
            padded = [_format_cell_as_str(row[i]) if i < len(row) else "" for i in range(n_cols)]
            if any(cell != "" for cell in padded):
                data_rows.append(padded)
        return pd.DataFrame(data_rows, columns=headers, dtype=str)
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
            return _read_excel_calamine(source, sheet_name=sheet_name)
        except Exception:
            pass
    return _read_excel_openpyxl(source, sheet_name=sheet_name)


def load_uploaded_files(
    files: Iterable[tuple[str, str | Path | bytes]],
    sheet_name: str | None = None,
) -> LoadedDataset:
    """Load and concatenate one or more uploaded Excel/CSV files preserving all columns."""
    frames: list[pd.DataFrame] = []
    source_names: list[str] = []
    row_counts: dict[str, int] = {}
    ordered_cols: list[str] = []
    seen_cols: set[str] = set()

    for fname, payload in files:
        df = read_tabular_file(payload, filename=fname, sheet_name=sheet_name)
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


def inspect_uploaded_batch(file_entries: tuple[tuple[str, bytes], ...]) -> dict[str, Any]:
    """Inspect uploaded files for preflight stats, key presence, and file-swap detection."""
    total_rows = 0
    max_cols = 0
    all_norm_cols: set[str] = set()
    swap_flags: list[tuple[str, str]] = []
    file_details: list[dict[str, Any]] = []

    for fname, raw_bytes in file_entries:
        df = read_tabular_file(raw_bytes, filename=fname)
        rows_cnt = len(df)
        cols_cnt = len(df.columns)
        norm_cols = {normalize_col_name(c) for c in df.columns}
        role = detect_file_role(norm_cols, cols_cnt)
        total_rows += rows_cnt
        max_cols = max(max_cols, cols_cnt)
        all_norm_cols |= norm_cols
        swap_flags.append((fname, role))
        file_details.append({"filename": fname, "rows": rows_cnt, "cols": cols_cnt, "detected_role": role})

    return {
        "file_count": len(file_entries),
        "total_rows": total_rows,
        "max_cols": max_cols,
        "has_regn_no": "regn_no" in all_norm_cols,
        "has_sem": "sem" in all_norm_cols,
        "file_roles": swap_flags,
        "file_details": file_details,
    }


@st.cache_data(show_spinner=False, max_entries=16)
def _cached_preflight_inspect(file_entries: tuple[tuple[str, bytes], ...]) -> dict[str, Any]:
    return inspect_uploaded_batch(file_entries)


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
    sap_conflicts_df: pd.DataFrame
    email_conflicts_df: pd.DataFrame
    anomalies_long_df: pd.DataFrame
    anomalies_wide_df: pd.DataFrame
    format_only_df: pd.DataFrame
    column_anomaly_summary_df: pd.DataFrame
    email_only_df: pd.DataFrame
    student_report_df: pd.DataFrame
    student_timeline_map: dict[str, list[dict[str, str]]]
    semester_columns: list[str]

    metrics: dict[str, int]


def _get_col_series(dataset: LoadedDataset, canonical_col: str) -> pd.Series:
    orig = dataset.col_map.get(canonical_col)
    if orig and orig in dataset.df.columns:
        return dataset.df[orig]
    return pd.Series([""] * len(dataset.df), index=dataset.df.index, dtype=str)


def _build_composite_key(dataset: LoadedDataset, key_columns: list[str]) -> pd.Series:
    parts: list[pd.Series] = []
    for k in key_columns:
        s = _get_col_series(dataset, k)
        if k == "sem":
            parts.append(s.map(norm_sem_key))
        elif k == "month":
            parts.append(s.map(lambda v: parse_month(v)[1]))
        elif k in ("year", "admission_year"):
            parts.append(s.map(norm_numeric))
        elif k == "dob":
            parts.append(s.map(norm_date))
        else:
            parts.append(s.map(lambda v: norm_text(v, case_insensitive=True)))
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
        s = _get_col_series(dataset, k)
        df[f"__k_{k}__"] = s.map(norm_sem_key) if k == "sem" else s.map(lambda v: norm_text(v, case_insensitive=True))

    regn_s = _get_col_series(dataset, "regn_no").map(lambda v: norm_text(v, case_insensitive=True))
    sem_s = _get_col_series(dataset, "sem").map(norm_sem_key)
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

    dup_rows = df_sorted[df_sorted["__key__"].isin(dup_keys)].copy()
    mrks_status_col = dataset.col_map.get("mrks_rec_status")

    conflict_records: list[dict[str, str]] = []
    conflict_type_by_key: dict[str, str] = {}

    for key_val, grp in dup_rows.groupby("__key__", sort=False):
        pass_rows = grp[grp["__result_norm__"] == "PASS"]
        fail_rows = grp[grp["__result_norm__"] == "FAIL"]

        has_pass = len(pass_rows) > 0
        has_fail = len(fail_rows) > 0

        sigs = set(
            zip(
                grp["__result_norm__"],
                grp["__year_int__"],
                grp["__month_norm__"],
                grp["__mrks_float__"],
                grp["__sgpa_float__"],
                grp["__mrks_status_rank__"],
            )
        )
        if len(sigs) == 1:
            ctype = "EXACT_OR_NEAR_DUPLICATE"
        elif has_pass and has_fail and len(pass_rows) == 1:
            ctype = "FAIL_AND_PASS"
        elif has_pass and has_fail and len(pass_rows) > 1:
            ctype = "MULTIPLE_PASS_AND_FAIL"
        elif len(pass_rows) > 1:
            ym_pairs = set(zip(pass_rows["__year_int__"], pass_rows["__month_norm__"]))
            ctype = "MULTIPLE_PASS_DIFF_MONTH_YEAR" if len(ym_pairs) > 1 else "MULTIPLE_PASS_SAME_PERIOD"
        elif len(fail_rows) > 1:
            ctype = "MULTIPLE_FAIL"
        else:
            ctype = "DUPLICATE_OTHER"

        conflict_type_by_key[key_val] = ctype

        for idx_in_grp, (_, row) in enumerate(grp.iterrows()):
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
                elif row["__result_norm__"] == "FAIL" and has_pass:
                    action = "SUPERSEDED (By PASS Record)"
                elif has_pass:
                    action = "SUPERSEDED (Older/Lower PASS)"
                else:
                    action = "SUPERSEDED (Older Record)"

            conflict_records.append(
                {
                    "source": source_label,
                    "regn_no": row["__regn_norm__"],
                    "sem": row["__sem_norm__"],
                    "cname": row["__cname_norm__"],
                    "conflict_type": ctype,
                    "action": action,
                    "result": row["__result_norm__"],
                    "year": str(row["__year_int__"] or ""),
                    "month": row["__month_norm__"],
                    "tot_mrks": clean_str(row.get(dataset.col_map.get("tot_mrks", ""), "")),
                    "sgpa": clean_str(row.get(dataset.col_map.get("sgpa", ""), "")),
                    "mrks_rec_status": clean_str(row.get(mrks_status_col, "")) if mrks_status_col else "",
                    "uploaded_file": clean_str(row.get("__uploaded_file__", "")),
                    "excel_row": clean_str(row.get("__uploaded_row__", "")),
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
            if status == "VALID":
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
    sap_adm_col = sap_dataset.col_map.get("admission_year", "")

    abc_readiness_list: list[str] = []
    abc_orig_status_list: list[str] = []
    recovered_abc_list: list[str] = []
    recovery_source_list: list[str] = []
    abc_audit_rows: list[dict[str, str]] = []

    for _, prow in pending_sap.iterrows():
        regn = prow["__regn_norm__"]
        sem = prow["__sem_norm__"]
        orig_abc = clean_str(prow.get(sap_abc_col, "")) if sap_abc_col else ""
        abc_status, _ = validate_abc_id(orig_abc)
        abc_orig_status_list.append(abc_status)

        if abc_status == "VALID":
            abc_readiness_list.append("READY")
            recovered_abc_list.append("")
            recovery_source_list.append("")
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
                    "cname": prow["__cname_norm__"],
                    "course_name": clean_str(prow.get(sap_course_col, "")) if sap_course_col else "",
                    "result": prow["__result_norm__"],
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
                    "cname": prow["__cname_norm__"],
                    "course_name": clean_str(prow.get(sap_course_col, "")) if sap_course_col else "",
                    "result": prow["__result_norm__"],
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
    sap_sems_by_regn: dict[str, list[str]] = {}
    for regn, grp in sap_clean.groupby("__regn_norm__", sort=False):
        sems = sorted({s for s in grp["__sem_norm__"] if s}, key=sem_sort_key)
        sap_sems_by_regn[regn] = sems

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

    # Step 5: Cross-File Matching, FAIL->PASS Upgrades, Anomalies, and PASS->PASS Critical Corrections
    matched_sap = sap_clean[sap_clean["__key__"].isin(email_keys_set)].set_index("__key__")
    matched_email = email_clean[email_clean["__key__"].isin(sap_keys_set)].set_index("__key__")
    matched_keys = matched_sap.index.tolist()
    matched_email = matched_email.loc[matched_keys]

    upgrades_mask = (matched_email["__result_norm__"] == "FAIL") & (
        matched_sap["__result_norm__"] == "PASS"
    )
    upgrade_keys = matched_sap.index[upgrades_mask].tolist()

    upgrade_rows: list[dict[str, str]] = []
    for k in upgrade_keys:
        sr = matched_sap.loc[k]
        er = matched_email.loc[k]
        upgrade_rows.append(
            {
                "regn_no": sr["__regn_norm__"],
                "sem": sr["__sem_norm__"],
                "cname": sr["__cname_norm__"],
                "course_name": clean_str(sr.get(sap_course_col, "")) if sap_course_col else "",
                "email_result": er["__result_norm__"],
                "email_year": str(er["__year_int__"] or ""),
                "email_month": er["__month_norm__"],
                "email_tot_mrks": clean_str(er.get(email_dataset.col_map.get("tot_mrks", ""), "")),
                "email_sgpa": clean_str(er.get(email_dataset.col_map.get("sgpa", ""), "")),
                "sap_result": sr["__result_norm__"],
                "sap_year": str(sr["__year_int__"] or ""),
                "sap_month": sr["__month_norm__"],
                "sap_tot_mrks": clean_str(sr.get(sap_dataset.col_map.get("tot_mrks", ""), "")),
                "sap_sgpa": clean_str(sr.get(sap_dataset.col_map.get("sgpa", ""), "")),
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

    base_regn = matched_sap["__regn_norm__"].tolist()
    base_sem = matched_sap["__sem_norm__"].tolist()
    base_cname = matched_sap["__cname_norm__"].tolist()
    base_course = matched_sap[sap_course_col].tolist() if sap_course_col else [""] * len(matched_keys)
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

    # Also track PASS -> PASS critical VALUE_MISMATCHes across PASS_CORRECTION_CRITICAL_COLUMNS
    # regardless of user's anomaly column selection if both columns exist, OR within common_cols!
    # We evaluate PASS_CORRECTION_CRITICAL_COLUMNS directly on all common PASS->PASS pairs.
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
    email_mrks_col = email_dataset.col_map.get("mrks_rec_status", "")
    for k in pass_correction_keys:
        sr = matched_sap.loc[k]
        er = matched_email.loc[k]
        diffs = pass_crit_mismatches_by_key[k]
        pass_correction_rows.append(
            {
                "regn_no": sr["__regn_norm__"],
                "sem": sr["__sem_norm__"],
                "cname": sr["__cname_norm__"],
                "course_name": clean_str(sr.get(sap_course_col, "")) if sap_course_col else "",
                "mismatch_columns": ", ".join(c for c, _, _ in diffs),
                "email_values": " | ".join(f"{c}={ev}" for c, ev, _ in diffs),
                "sap_values": " | ".join(f"{c}={sv}" for c, _, sv in diffs),
                "email_tot_mrks": clean_str(er.get(email_dataset.col_map.get("tot_mrks", ""), "")),
                "sap_tot_mrks": clean_str(sr.get(sap_dataset.col_map.get("tot_mrks", ""), "")),
                "email_sgpa": clean_str(er.get(email_dataset.col_map.get("sgpa", ""), "")),
                "sap_sgpa": clean_str(sr.get(sap_dataset.col_map.get("sgpa", ""), "")),
                "email_mrks_rec_status": clean_str(er.get(email_mrks_col, "")) if email_mrks_col else "",
                "recommended_action": "Upload to DigiLocker with mrks_rec_status = M (Modified)",
            }
        )

    pass_corrections_df = pd.DataFrame(pass_correction_rows)
    pass_corrections_raw_sap_df = (
        matched_sap.loc[pass_correction_keys, sap_dataset.original_columns].reset_index(drop=True)
        if pass_correction_keys
        else pd.DataFrame(columns=sap_dataset.original_columns)
    )

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

    # Step 6: Complete Student-Wise Report & Student 360° Timeline Map
    all_sems = sorted(
        (set(sap_clean["__sem_norm__"]) | set(email_clean["__sem_norm__"])) - {""},
        key=sem_sort_key,
    )
    semester_columns = [f"Sem {s}" for s in all_sems]

    sap_by_student: dict[str, dict[str, dict[str, str]]] = {}
    email_by_student: dict[str, dict[str, dict[str, str]]] = {}
    student_profile: dict[str, dict[str, str]] = {}

    def _update_profile(regn: str, row: pd.Series, ds: LoadedDataset) -> None:
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
        for field_name in (
            "cname",
            "course_name",
            "academic_course_id",
            "admission_year",
            "abc_account_id",
            "dob",
            "gender",
        ):
            if not prof[field_name]:
                c = ds.col_map.get(field_name)
                if c:
                    val = clean_str(row.get(c, ""))
                    if val:
                        prof[field_name] = norm_text(val, case_insensitive=True) if field_name == "cname" else val

    pending_readiness_by_key: dict[str, str] = dict(
        zip(pending_sap["__key__"].tolist(), pending_sap["__abc_readiness__"].tolist())
    )

    def _extract_snap(row: pd.Series, ds: LoadedDataset) -> dict[str, str]:
        tot_m = clean_str(row.get(ds.col_map.get("tot_mrks", ""), ""))
        sgpa_v = clean_str(row.get(ds.col_map.get("sgpa", ""), ""))
        yr = str(row["__year_int__"] or "")
        mo = row["__month_norm__"]
        return {
            "key": row["__key__"],
            "result": row["__result_norm__"],
            "period": f"{yr} {mo}".strip(),
            "marks_sgpa": f"{tot_m or '-'} / {sgpa_v or '-'}",
            "credit_pts": clean_str(row.get(ds.col_map.get("tot_credit_points", ""), "")),
            "abc_id": clean_str(row.get(ds.col_map.get("abc_account_id", ""), "")),
            "course_id": clean_str(row.get(ds.col_map.get("academic_course_id", ""), "")),
            "mrks_status": clean_str(row.get(ds.col_map.get("mrks_rec_status", ""), "")),
            "conflict": sap_conflict_by_key.get(row["__key__"], ""),
        }

    for _, row in sap_clean.iterrows():
        regn = row["__regn_norm__"]
        sem = row["__sem_norm__"]
        _update_profile(regn, row, sap_dataset)
        sap_by_student.setdefault(regn, {})[sem] = _extract_snap(row, sap_dataset)

    for _, row in email_clean.iterrows():
        regn = row["__regn_norm__"]
        sem = row["__sem_norm__"]
        _update_profile(regn, row, email_dataset)
        email_by_student.setdefault(regn, {})[sem] = _extract_snap(row, email_dataset)

    upgrade_keys_set = set(upgrade_keys)
    pass_corr_keys_set = set(pass_correction_keys)
    student_rows: list[dict[str, Any]] = []
    student_timeline_map: dict[str, list[dict[str, str]]] = {}

    status_counts = {
        "Complete": 0,
        "Complete (With Anomalies)": 0,
        "Partial": 0,
        "Not Uploaded": 0,
        "Email Only": 0,
    }

    for regn in sorted(student_profile.keys()):
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
        for s in all_sems:
            col_label = f"Sem {s}"
            in_s = s in s_map
            in_e = s in e_map
            sd = s_map.get(s)
            ed = e_map.get(s)

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
                    cell_status = f"M_CORRECTION_NEEDED ({acnt or len(pass_crit_mismatches_by_key.get(k, []))})"
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

    student_report_df = pd.DataFrame(student_rows)

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
        sap_conflicts_df=sap_conflicts_df,
        email_conflicts_df=email_conflicts_df,
        anomalies_long_df=anomalies_long_df,
        anomalies_wide_df=anomalies_wide_df,
        format_only_df=format_only_df,
        column_anomaly_summary_df=column_anomaly_summary_df,
        email_only_df=email_only_df,
        student_report_df=student_report_df,
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
    "email_value",
    "sap_value",
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


def filter_pending_sap(
    result: ReconciliationResult,
    result_filter: Literal["ALL", "PASS_ONLY", "FAIL_ONLY"] = "ALL",
    abc_filter: Literal[
        "ALL", "READY_ONLY", "READY_AND_RECOVERABLE", "MISSING_OR_INVALID_ONLY"
    ] = "ALL",
    auto_backfill_abc: bool = False,
    include_fail_to_pass_upgrades: bool = False,
    include_pass_corrections: bool = False,
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
    sap_mrks_col = result.sap_col_map.get("mrks_rec_status")
    sap_regn_col = result.sap_col_map.get("regn_no")

    def _prepare_m_rows(raw_df: pd.DataFrame) -> pd.DataFrame:
        m_df = raw_df[result.sap_original_columns].copy()
        if sap_mrks_col:
            m_df[sap_mrks_col] = "M"
        if auto_backfill_abc and sap_abc_col and sap_regn_col and not m_df.empty:
            for ridx in m_df.index:
                status, _ = validate_abc_id(m_df.at[ridx, sap_abc_col])
                if status != "VALID":
                    regn_norm = norm_text(m_df.at[ridx, sap_regn_col], case_insensitive=True)
                    if regn_norm in result.verified_abc_map:
                        m_df.at[ridx, sap_abc_col] = result.verified_abc_map[regn_norm][0]
        return m_df

    if include_fail_to_pass_upgrades and not result.fail_to_pass_raw_sap_df.empty:
        extra_frames.append(_prepare_m_rows(result.fail_to_pass_raw_sap_df))

    if include_pass_corrections and not result.pass_corrections_raw_sap_df.empty:
        extra_frames.append(_prepare_m_rows(result.pass_corrections_raw_sap_df))

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
        width = min(max(max_len + 3, 12), max_width)
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
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max(max_len + 3, 12), max_width)


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
        {"Category": "ABC Readiness (Pending)", "Metric": "ABC Recoverable (Found in Other Semesters)", "Count": m["pending_abc_recoverable"]},
        {"Category": "ABC Readiness (Pending)", "Metric": "ABC Missing / Dummy / Invalid (Unrecoverable)", "Count": m["pending_abc_missing_or_invalid"]},
        {"Category": "M-Status Remediation", "Metric": "FAIL -> PASS Upgrades Needed (mrks_rec_status = M)", "Count": m["fail_to_pass_upgrades"]},
        {"Category": "M-Status Remediation", "Metric": "PASS -> PASS Critical Corrections (mrks_rec_status = M)", "Count": m["pass_critical_corrections"]},
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
                _apply_openpyxl_sheet(writer, sdf, sname)

    return buf.getvalue()


# ==============================================================================
# SECTION 5: Interactive Streamlit UI
# ==============================================================================


def _clear_results() -> None:
    st.session_state.pop("dl_recon_result", None)
    st.session_state.pop("dl_recon_run_fingerprint", None)
    st.session_state.pop("dl_recon_pure_cache", None)
    st.session_state.pop("dl_recon_report_cache", None)


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


def _render_preflight_box(uploads: list | None, expected_role: Literal["SAP", "EMAIL"]) -> None:
    if not uploads:
        return
    try:
        entries = tuple((f.name, f.getvalue()) for f in uploads)
        info = _cached_preflight_inspect(entries)
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
    b1, b2, b3, b4 = st.columns(4)
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
                drop_internal_cols=drop_internal_cols,
                trim_whitespace=trim_ws,
            )
            pure_csv_name, pure_xlsx_name, report_xlsx_name = build_dynamic_filenames(
                sap_filenames, result_filter=result_filter_code, abc_filter=abc_filter_code
            )

            st.info(
                f"**Ready to export:** **{len(export_df):,} rows** × **{len(export_df.columns)} columns**"
            )

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

            pure_cache_key = (
                f"{result_filter_code}|{abc_filter_code}|{auto_backfill}|"
                f"{include_upgrades}|{include_pass_corr}|{drop_internal_cols}|{trim_ws}"
            )
            cache_dict = st.session_state.setdefault("dl_recon_pure_cache", {})
            with b_xlsx:
                if pure_cache_key in cache_dict:
                    st.download_button(
                        label="Download Excel (.xlsx)",
                        data=cache_dict[pure_cache_key],
                        file_name=pure_xlsx_name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="dl_recon_dl_pure_xlsx_btn",
                        type="primary",
                        use_container_width=True,
                    )
                else:
                    if st.button(
                        "Prepare Pure SAP Excel",
                        key="dl_recon_prep_pure_btn",
                        use_container_width=True,
                    ):
                        with st.spinner("Generating text-safe Pure Pending SAP Excel..."):
                            cache_dict[pure_cache_key] = build_pure_pending_excel(
                                res,
                                result_filter=result_filter_code,
                                abc_filter=abc_filter_code,
                                auto_backfill_abc=auto_backfill,
                                include_fail_to_pass_upgrades=include_upgrades,
                                include_pass_corrections=include_pass_corr,
                                drop_internal_cols=drop_internal_cols,
                                trim_whitespace=trim_ws,
                            )
                        st.rerun()

    with col_dl2:
        with st.container(border=True):
            st.markdown("#### Full Reconciliation & Student-Wise Report")
            _, _, report_xlsx_name = build_dynamic_filenames(sap_filenames)
            st.info(
                f"**File:** `{report_xlsx_name}`\n\n"
                f"Includes sheets for:\n"
                f"- **Executive Summary** & **Student-Wise Matrix** ({m['total_students']:,} students)\n"
                f"- **Pending Upload Summary** & **ABC Backfill Audit** ({len(res.abc_backfill_audit_df):,} flagged)\n"
                f"- **SAP Duplicates/Conflicts** & **Email Duplicates**\n"
                f"- **Column Stats**, **Anomalies (By Record & Detailed)**, "
                f"**FAIL→PASS Upgrades**, **PASS Critical Corrections**, & **Email-Only Records**"
            )
            if "dl_recon_report_cache" in st.session_state:
                st.download_button(
                    label=f"Download {report_xlsx_name}",
                    data=st.session_state["dl_recon_report_cache"],
                    file_name=report_xlsx_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="dl_recon_dl_rep_btn",
                    type="primary",
                    use_container_width=True,
                )
            else:
                if st.button(
                    "Prepare Full Reconciliation Report",
                    key="dl_recon_prep_rep_btn",
                    type="primary",
                    use_container_width=True,
                ):
                    with st.spinner("Building styled multi-sheet Excel report..."):
                        st.session_state["dl_recon_report_cache"] = build_reconciliation_report_excel(res)
                    st.rerun()


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

    if not res.pass_corrections_df.empty:
        with st.expander(
            f"PASS → PASS Critical Corrections ({len(res.pass_corrections_df):,} records — upload with MRKS_REC_STATUS = M)",
            expanded=False,
        ):
            st.dataframe(res.pass_corrections_df, use_container_width=True)
            _render_download_bar(
                res.pass_corrections_df, "Pass_Critical_Corrections", "dl_recon_pcc"
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


def _render_student_report_tab(res: ReconciliationResult) -> None:
    st.subheader("Complete Student-Wise Semester Coverage & 360° Inspector")
    sdf = res.student_report_df
    if sdf.empty:
        st.info("No student records found.")
        return

    f1, f2, f3, f4 = st.columns([1.3, 1.5, 1.2, 2])
    with f1:
        statuses = ["ALL"] + sorted(sdf["Overall_Status"].fillna("").astype(str).unique().tolist())
        sel_status = st.selectbox("Overall Status", statuses, key="dl_recon_stu_status")
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

    view = sdf.copy()
    if sel_status != "ALL":
        view = view[view["Overall_Status"] == sel_status]
    if sel_course != "ALL":
        view = view[view["course_name"] == sel_course]
    if sel_adm != "ALL":
        view = view[view["admission_year"] == sel_adm]
    if q.strip():
        ql = q.strip().lower()
        view = view[
            view["regn_no"].fillna("").astype(str).str.lower().str.contains(ql, regex=False, na=False)
            | view["cname"].fillna("").astype(str).str.lower().str.contains(ql, regex=False, na=False)
            | view["abc_account_id"].fillna("").astype(str).str.lower().str.contains(ql, regex=False, na=False)
        ]

    st.caption(f"Showing **{min(len(view), 1000):,}** of **{len(view):,}** students.")
    st.dataframe(view.head(1000), use_container_width=True, height=450)
    _render_download_bar(view, "Filtered_Student_Report", "dl_recon_tab5")

    # Single-Student 360° Inspector
    st.markdown("---")
    st.markdown("#### Single-Student 360° Inspector")
    if view.empty:
        st.info("No students match the current filter to inspect.")
        return

    candidates = view.head(500)
    options_labels = [
        f"{r} — {n}" if n else str(r)
        for r, n in zip(candidates["regn_no"].tolist(), candidates["cname"].tolist())
    ]
    sel_label = st.selectbox(
        "Select Student (from current filtered view, up to 500)",
        options=options_labels,
        key="dl_recon_360_select",
    )
    sel_regn = sel_label.split(" — ", 1)[0].strip()

    stu_meta = view[view["regn_no"] == sel_regn].iloc[0]
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


def show() -> None:
    """Entry point rendered inside SVKM ExamAutomationSuite (or standalone main)."""
    st.title("DigiLocker Reconciliation")
    st.caption(
        "Upload raw SAP Data and DigiLocker Email Reference files to deduplicate SAP conflicts "
        "(duplicates, FAIL+PASS pairs, and multiple PASS cycles keeping latest Year/Month/M-status), "
        "recover missing ABC IDs, export NAD-ready pending SAP files, and inspect a 360° student matrix."
    )
    st.markdown("---")

    up_col1, up_col2 = st.columns(2)
    with up_col1:
        with st.container(border=True):
            sap_uploads = st.file_uploader(
                "1. Upload SAP Raw Data File(s) (.xlsx / .csv)",
                type=["xlsx", "xlsm", "xls", "csv"],
                accept_multiple_files=True,
                key="dl_recon_sap_uploader",
            )
            _render_preflight_box(sap_uploads, expected_role="SAP")
    with up_col2:
        with st.container(border=True):
            email_uploads = st.file_uploader(
                "2. Upload Email Reference Data File(s) (.xlsx / .csv)",
                type=["xlsx", "xlsm", "xls", "csv"],
                accept_multiple_files=True,
                key="dl_recon_email_uploader",
            )
            _render_preflight_box(email_uploads, expected_role="EMAIL")

    with st.expander("Matching Key & Anomaly Configuration", expanded=False):
        cfg_c1, cfg_c2 = st.columns(2)
        with cfg_c1:
            key_cols = st.multiselect(
                "Matching Key Columns",
                options=["regn_no", "sem", "academic_course_id", "exam_type", "year", "month", "rroll"],
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
            norm_text_opt,
            ignore_punct_opt,
            norm_num_opt,
            norm_month_opt,
            norm_date_opt,
            include_null_gaps,
            include_fmt_only,
        ),
    )

    run_clicked = st.button(
        "Run DigiLocker Reconciliation",
        type="primary",
        use_container_width=True,
        key="dl_recon_run_btn",
    )

    if run_clicked:
        if not sap_uploads or not email_uploads:
            st.error("Please upload at least one SAP Data file and one Email Reference Data file.")
            return
        if not key_cols:
            st.error("Please select at least one matching key column.")
            return

        _clear_results()
        try:
            with st.spinner("Loading Excel/CSV files and running reconciliation..."):
                sap_payloads = [(f.name, f.getvalue()) for f in sap_uploads]
                email_payloads = [(f.name, f.getvalue()) for f in email_uploads]
                sap_ds = load_uploaded_files(sap_payloads)
                email_ds = load_uploaded_files(email_payloads)

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
        except Exception as exc:
            st.error(f"Reconciliation failed: {exc}")
            return

    if "dl_recon_result" not in st.session_state:
        st.info("Upload both SAP Data and Email Reference Data files above, then click **Run DigiLocker Reconciliation**.")
        return

    if st.session_state.get("dl_recon_run_fingerprint") != current_fp:
        st.warning(
            "Uploaded files or configuration settings have changed since the last run. "
            "Click **Run DigiLocker Reconciliation** to refresh results."
        )

    res: ReconciliationResult = st.session_state["dl_recon_result"]
    sap_names: list[str] = st.session_state.get("dl_recon_sap_filenames", [])

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


if __name__ == "__main__":
    st.set_page_config(
        page_title="DigiLocker Reconciliation",
        page_icon="🛡️",
        layout="wide",
        initial_sidebar_state="collapsed",
    )
    show()
