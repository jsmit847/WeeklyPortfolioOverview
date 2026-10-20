from __future__ import annotations

import copy
import datetime as dt
import hashlib
import html
import io
import logging
import math
import numbers
import os
import posixpath
import re
import secrets
import sys
import time
import traceback
import urllib.parse
import zipfile
import xml.etree.ElementTree as ET
from contextlib import contextmanager
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Callable, Dict, Iterable, Iterator, List, Optional, Sequence, Tuple

import numpy as np
import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string, get_column_letter

import landgorilla_live as live_data
import meeting_workspace as meeting_work
import portfolio_cockpit as cockpit
import portfolio_data_views as data_views
import portfolio_exports as verified_exports
import portfolio_shell as shell


APP_TITLE = "Weekly Portfolio Meeting Deck"
APP_VERSION = "2.7.0"
APP_DIR = Path(__file__).resolve().parent

# Privacy-safe event log: key=value events with counts, sha prefixes and deal keys only
# (never deal names, borrowers or commentary).
LOG = logging.getLogger("wpo")

# An explicit workbook path (tests, local demos) wins over the files below.
DEFAULT_WORKBOOK_ENV = "WPO_DEFAULT_WORKBOOK"
DEFAULT_SAMPLE_FILES = [Path('portfolio_overview.xlsx'), APP_DIR / 'portfolio_overview.xlsx']

# Optional local file. Do not commit real deal-level history to a public repo.
STATUS_HISTORY_PATH = APP_DIR / "status_history.csv"
STATUS_HISTORY_SHEET_NAME = "_Status History"
STATUS_STALE_MEETING_THRESHOLD = 3
STATUS_HISTORY_COLUMNS = [
    "meeting_date",
    "sheet",
    "deal_number",
    "deal_key",
    "status",
    "owner",
    "commentary",
    "workbook_name",
    "saved_at",
]

DATE_HDR = re.compile(r"^\s*(\d{1,2})/(\d{1,2})(?:/(\d{2,4}))?\s+(.*\S)\s*$")

# -----------------------------------------------------------------------------
# Workbook ingest settings
# -----------------------------------------------------------------------------
PORTFOLIO_SHEETS = ("Bridge", "Term")
SHEET_ORDER = {"Bridge": 0, "Term": 1}
KEY_HEADER = "Deal Number"
HEADER_SEARCH_ROWS = 60  # the 'Deal Number' header must sit in the first 60 rows
TITLE_SEARCH_ROWS = 5  # 'as of MM/DD/YYYY' is looked for in the first 5 rows
BLANK_ROW_STOP = 500  # stop reading a sheet after this many fully blank rows
MAX_ISSUE_SAMPLES = 25  # cap on example rows kept per data issue

# Dates outside this window are typos or sentinels (3026-06-01, 12/31/9999).
DATE_MIN = pd.Timestamp("1900-01-01")
DATE_MAX_EXCLUSIVE = pd.Timestamp("2200-01-01")
EXCEL_EPOCH = pd.Timestamp("1899-12-30")
EXCEL_SERIAL_MIN = 20000  # 1954-10-04
EXCEL_SERIAL_MAX = 80000  # 2119-01-10

# Maturity within this many days of the as-of date is a watch item (amber).
MATURITY_WATCH_DAYS = 30

# Deal Number cells that are not deals: Excel error values and totals rows.
EXCEL_ERROR_RE = re.compile(
    r"^#(?:N/A|NULL!|DIV/0!|VALUE!|REF!|NAME\?|NUM!|GETTING_DATA|SPILL!|CALC!|FIELD!|"
    r"BLOCKED!|CONNECT!|BUSY!|UNKNOWN!|EXTERNAL!|PYTHON!)$",
    re.IGNORECASE,
)
TOTAL_ROW_RE = re.compile(
    r"^(?:(?:grand|portfolio|bridge|term)\s+)?(?:sub[\s-]*)?totals?\b",
    re.IGNORECASE,
)
TITLE_AS_OF_RE = re.compile(
    r"\bas\s+of\b[\s:]*(?:(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})|(\d{4})-(\d{1,2})-(\d{1,2}))",
    re.IGNORECASE,
)
TITLE_AS_OF_WORDS_RE = re.compile(
    r"\bas\s+of\b[\s:]*((?:jan|feb|mar|apr|may|jun|jul|aug|sep|sept|oct|nov|dec)[a-z]*\.?\s+\d{1,2}(?:st|nd|rd|th)?,?\s+\d{4})",
    re.IGNORECASE,
)

# Header candidates for loan size, in priority order. Term (DSCR) loans fund in full at
# closing, so when a Term sheet has no explicit commitment column the original loan
# amount is the commitment. Headers also match loosely (case, punctuation, '$' / '(USD)'
# and abbreviations such as 'Orig. Loan Amt ($)'), and every matching column is used in
# priority order, so a blank first column falls through to the next one.
COMMITMENT_HEADER_CANDIDATES = (
    "Loan Commitment",
    "Total Commitment",
    "Total Loan Commitment",
    "Commitment",
    "Commitment Amount",
    "Loan Commitment Amount",
    "Total Commitment Amount",
    "Current Commitment",
    "Current Loan Commitment",
)
LOAN_AMOUNT_HEADER_CANDIDATES = (
    "Loan Amount",
    "Original Loan Amount",
    "Original Balance",
    "Original Principal Balance",
    "Orig Loan Amount",
    "Note Amount",
    "Original Loan Balance",
    "Original Principal Amount",
    "Original Principal",
    "Original Note Amount",
    "Initial Loan Amount",
    "Total Loan Amount",
    "Loan Size",
    "Original UPB",
    "Funded Loan Amount",
    "Closing Loan Amount",
    "Loan Amount at Origination",
)
# Source labels for the 'Loan Commitment' KPI helper (commitment_source column).
COMMITMENT_SOURCE_EXPLICIT = "Total commitment"
COMMITMENT_SOURCE_LOAN_AMOUNT = "Original loan amount"
COMMITMENT_SOURCE_MISSING = "Not in workbook"
# Abbreviations expanded by loose_header_key ('Orig. Loan Amt' -> 'original loan amount').
HEADER_ABBREVIATIONS = {
    "orig": "original",
    "amt": "amount",
    "amnt": "amount",
    "bal": "balance",
    "princ": "principal",
    "commit": "commitment",
    "cmt": "commitment",
    "tot": "total",
    "ln": "loan",
    "init": "initial",
}

# Logical fields of a Bridge/Term sheet and the headers that can carry them, in
# priority order. The first matching column (left-most) wins; the loader and the
# workbook writer both resolve columns through this table (resolve_sheet_layout).
SHEET_ROLE_CANDIDATES: Dict[str, Tuple[str, ...]] = {
    "deal": (KEY_HEADER,),
    "name": ("Deal Name", "Loan Name", "Property Name"),
    "borrower": ("Borrower Name", "Borrower Entity", "Account Name"),
    "account": ("Account", "Account Name"),
    "status": ("Status",),
    "owner": ("Point Person", "Asset Manager", "Active RM"),
    "commentary": ("AM Commentary",),
    "servicer": ("Servicer",),
    "portfolio": ("Portfolio",),
    "segment": ("Segment",),
    "financing": ("Financing",),
    "loan_buyer": ("Loan Buyer",),
    "maturity": (
        "Current Maturity Date",
        "Maturity Date",
        "Original Maturity Date",
        "Next Advance Maturity Date",
    ),
    "next_payment": ("Next Payment Date",),
    "upb": ("UPB",),
    "npl": ("NPL", "Loan Level Delinquency", "DQ Status"),
    "funded": ("Active Funded Amount",),
    "remaining": ("Remaining Commitment",),
    "commitment": COMMITMENT_HEADER_CANDIDATES,
    "loan_amount": LOAN_AMOUNT_HEADER_CANDIDATES,
}
# Money roles whose headers also match loosely (see loose_header_key). For the loan-size
# roles every matching column is kept, in priority order, and coalesced per row.
LOOSE_MATCH_ROLES = ("commitment", "loan_amount", "funded", "remaining")
COALESCED_MONEY_ROLES = ("commitment", "loan_amount")
# Roles whose duplicate headers are reported (a second 'Status' column changes what the
# deck shows). 'Days Past Due' is deliberately multi-column and is never reported.
REPORTED_DUPLICATE_ROLES = (
    "deal",
    "name",
    "status",
    "owner",
    "commentary",
    "maturity",
    "next_payment",
    "upb",
    "npl",
    "commitment",
    "loan_amount",
)
DPD_HEADER = "Days Past Due"
OWNER_LABEL_FALLBACK = {"Bridge": "Point Person", "Term": "Asset Manager"}

COMMON_STATUS_SUGGESTIONS = [
    "Surveillance",
    "Upcoming Maturity",
    "Modification",
    "Foreclosure",
    "Litigation",
    "Receivership",
    "Probate Court",
    "Needs Follow-Up",
    "Resolved",
]

# -----------------------------------------------------------------------------
# Loan Modification sheet settings
# -----------------------------------------------------------------------------
LOAN_MOD_SHEET_NAME = "Loan Modifications"
LOAN_MOD_KEY_HEADERS = ("Deal Number", "Opportunity Deal Loan Number")

LOAN_MOD_DATE_COLUMNS = [
    "Cancellation Date",
    "Forbearance Termination Date",
    "Last Activity Date",
    "System Task Completed Date",
    "Mod Effective Date",
    "Modification Finalized Date",
    "Updated Maturity Date",
    "Previous Expiration Date",
    "Previous Maturity Date",
    "Prior Forbearance Termination Date",
    "Updated Expiration Date",
    "Created Date",
    "Last Modified Date",
    "System Modstamp",
    "Opportunity Close Date",
]

LOAN_MOD_MONEY_COLUMNS = [
    "Previous Loan Commitment",
    "Updated Loan Commitment",
]

LOAN_MOD_PERCENT_COLUMNS = [
    "Previous Exit Fee",
    "Previous Floor",
    "Previous Index Margin",
    "Previous Interest Rate",
    "Previous Max LTV %",
    "Previous Total ARV LTV",
    "Updated Exit Fee",
    "Updated Floor",
    "Updated Index Margin",
    "Updated Interest Rate",
    "Updated Max LTV %",
    "Updated Total ARV LTV",
]


# These are the columns the app needs to match loan mods back to the agenda.
# They stay lightweight and avoid loading the full Salesforce history table.
LOAN_MOD_REQUIRED_COLUMNS = [
    "Portfolio Overview Match",
    "Visible Overview Sheet(s)",
    "Deal Number",
    "Deal Name",
    "Opportunity Deal Loan Number",
    "Opportunity Deal Name",
]

# These are the highlighted fields in the Loan Modifications sheet.
# They get priority in Overview and Presentation.
LOAN_MOD_SPOTLIGHT_COLUMNS = [
    "Loan Mod Order Number",
    "Modification Type",
    "System Task Completed Date",
    "Mod Effective Date",
    "Previous Maturity Date",
    "Updated Maturity Date",
    "Previous Loan Commitment",
    "Updated Loan Commitment",
    "Comments",
]

# Small set of helper fields used only for sorting and labels.
LOAN_MOD_SUPPORT_COLUMNS = [
    "Loan Modification Name",
    "Status",
    "Loan Mod Type",
    "Mod Reporting Type",
    "Modification Finalized Date",
    "Deal Type",
    "Last Modified Date",
    "Created Date",
    "Cancellation Date",
    "Cancellation Reason",
]

LOAN_MOD_CORE_COLUMNS = list(
    dict.fromkeys(
        LOAN_MOD_REQUIRED_COLUMNS
        + LOAN_MOD_SPOTLIGHT_COLUMNS
        + LOAN_MOD_SUPPORT_COLUMNS
    )
)

# 'Portfolio Overview Match' values. The filter applies only when at least one row is
# recognised as a match; blank or unrecognised columns keep every row (fail open).
LOAN_MOD_MATCH_TRUE = frozenset({"YES", "Y", "TRUE", "T", "1", "X"})
LOAN_MOD_MATCH_FALSE = frozenset({"NO", "N", "FALSE", "F", "0"})

# A cancelled/withdrawn request is never the headline ("latest") mod of a deal.
LOAN_MOD_INACTIVE_RE = re.compile(
    r"cancel|withdr|declin|reject|denied|\bvoid|\bdead\b|expired|terminat|abandon|not\s+approved",
    re.IGNORECASE,
)
# Statuses whose updated terms are in force ("New maturity"); others are "Requested".
LOAN_MOD_DONE_RE = re.compile(r"complete|approved|executed|finali[sz]ed|booked|closed", re.IGNORECASE)
LOAN_MOD_TYPE_FIELDS = ("Modification Type", "Loan Mod Type", "Mod Reporting Type")

# -----------------------------------------------------------------------------
# Meeting edits, filters, dialogs and error boundaries
# -----------------------------------------------------------------------------
VIEW_OPTIONS = tuple(shell.VIEWS)
SHEET_FILTER_OPTIONS = ("All", "Bridge", "Term")
SEARCH_LABEL = "Search deal #, name, borrower, owner, status, commentary"
# Searched literally and case-insensitively (never as a regex).
SEARCH_COLUMNS = (
    "deal_number",
    "deal_name",
    "borrower",
    "owner",
    "status",
    "commentary",
    "servicer",
    "loan_mod_latest_summary",
    "loan_mod_latest_comments",
    "loan_mod_search_text",
)
# The three fields a meeting can edit; an override always holds all three.
EDITABLE_FIELDS = ("status", "owner", "commentary")
FIELD_LABELS = {"status": "Status", "owner": "Owner", "commentary": "AM Commentary"}
COMMENTARY_MAX_CHARS = 4000
# Review grid: editable display columns -> deal fields. Every other column, and the
# hidden deal_key index ('_index'), is read-only.
REVIEW_EDITABLE_COLUMNS = {"Status": "status", "Owner": "owner", "AM Commentary": "commentary"}
REVIEW_COLUMN_ORDER = (
    "Agenda #",
    "Type",
    "Deal Name",
    "Deal #",
    "Maturity",
    "Next Payment",
    "UPB",
    "Update Reminder",
    "Status",
    "Owner",
    "AM Commentary",
)
REVIEW_EDITOR_KEY_PREFIX = "review_editor::"
# Dialog kinds (active_dialog = {'kind', 'deal_key'}); only one dialog renders per run.
DIALOG_EDIT = "edit"
DIALOG_WORKBOOK_SWITCH = "wb_switch"
EDIT_DIALOG_VIEWS = ("Presentation", "Review")
# How an update reminder was addressed during this meeting.
REMINDER_CONFIRMED = "Confirmed"
REMINDER_REFRESHED = "Refreshed"
# Error boundaries: incident refs are 5 base32 characters; details only for local runs.
INCIDENT_REF_ALPHABET = "ABCDEFGHIJKLMNOPQRSTUVWXYZ234567"
DEBUG_ENV = "WPO_DEBUG"
LOCAL_HOSTS = frozenset({"localhost", "127.0.0.1", "::1"})
MAX_QUEUED_TOASTS = 8



# -----------------------------------------------------------------------------
# Text helpers: IDs, labels and free text
# -----------------------------------------------------------------------------
# Characters openpyxl refuses to write (plus U+FFFE/U+FFFF, which Excel cannot reload).
_XML_ILLEGAL_RE = re.compile("[\x00-\x08\x0b\x0c\x0e-\x1f￾￿]")
# Soft line breaks pasted from Word/PowerPoint/Outlook become real line breaks.
_SOFT_BREAK_RE = re.compile("[\x0b\x0c\x85  ]")
# Excel escapes control characters in shared strings as _x000D_; openpyxl keeps them.
_OOXML_CONTROL_ESCAPE_RE = re.compile(r"_x(00[01][0-9A-Fa-f])_")
_WHITESPACE_RE = re.compile(r"\s+")
_ID_FLOAT_TEXT_RE = re.compile(r"^\d+\.0$")
_BLANK_LINES_RE = re.compile(r"\n{3,}")


def is_missing(value: object) -> bool:
    """True for None, NaN, NaT and pd.NA (never raises, never True for text)."""
    if value is None or value is pd.NaT or value is pd.NA:
        return True
    if isinstance(value, (float, np.floating)):
        return math.isnan(value)
    if isinstance(value, np.datetime64):
        return bool(np.isnat(value))
    return False


def _number_text(value: object) -> str:
    """Text for a numeric cell. Integer-valued floats lose '.0' by type, never by suffix."""
    if isinstance(value, numbers.Integral):
        return str(int(value))
    if isinstance(value, numbers.Real):
        number = float(value)
        if math.isnan(number):
            return ""
        if math.isinf(number):
            return str(number)
        if number.is_integer() and abs(number) < 1e15:
            return str(int(number))
        return f"{number:.15g}"
    return str(value)


def _datetime_text(value: object) -> str:
    try:
        stamp = pd.Timestamp(value)
    except Exception:
        return str(value)
    if pd.isna(stamp):
        return ""
    if isinstance(value, dt.date) and not isinstance(value, dt.datetime):
        return stamp.strftime("%m/%d/%Y")
    if stamp.hour or stamp.minute or stamp.second:
        return stamp.strftime("%m/%d/%Y %H:%M")
    return stamp.strftime("%m/%d/%Y")


def _scalar_text(value: object) -> str:
    """Plain text for any cell value; '' for missing values."""
    if is_missing(value):
        return ""
    if isinstance(value, str):
        return value
    if isinstance(value, (bool, np.bool_)):
        return "TRUE" if value else "FALSE"
    if isinstance(value, (pd.Timestamp, dt.datetime, dt.date, np.datetime64)):
        return _datetime_text(value)
    if isinstance(value, dt.time):
        return value.strftime("%H:%M")
    if isinstance(value, numbers.Number):
        return _number_text(value)
    return str(value)


def _unescape_ooxml_controls(text: str) -> str:
    if "_x" not in text:
        return text
    return _OOXML_CONTROL_ESCAPE_RE.sub(lambda match: chr(int(match.group(1), 16)), text)


def norm_id(value: object) -> str:
    """Normalize an ID or key (deal numbers, deal keys, order numbers).

    Whitespace is collapsed. A trailing '.0' is dropped only from integer-valued floats
    (10001.0) or from text that is digits plus '.0' ('10001.0'); leading zeros stay.
    """
    if is_missing(value):
        return ""
    text = value if isinstance(value, str) else _scalar_text(value)
    text = _WHITESPACE_RE.sub(" ", text.replace("\xa0", " ")).strip()
    if _ID_FLOAT_TEXT_RE.match(text):
        text = text[:-2]
    return text


def norm_label(value: object) -> str:
    """Normalize a single-line label (status, owner, names, headers, sheet names).

    Only whitespace is normalized: line breaks and runs of spaces collapse to one space.
    Text is never otherwise rewritten ('Rate step-up to 9.0' stays as typed).
    """
    if is_missing(value):
        return ""
    text = _unescape_ooxml_controls(_scalar_text(value))
    text = _XML_ILLEGAL_RE.sub(" ", text)
    return _WHITESPACE_RE.sub(" ", text.replace("\xa0", " ")).strip()


def clean_free_text(value: object) -> str:
    """Normalize free text (commentary, loan-mod comments) while keeping line breaks.

    CRLF/CR become LF, soft breaks (vertical tab, form feed, U+2028/9) become LF, other
    characters Excel cannot store become spaces, trailing spaces are trimmed per line and
    three or more consecutive line breaks collapse to one blank line.
    """
    if is_missing(value):
        return ""
    text = _unescape_ooxml_controls(_scalar_text(value))
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = _SOFT_BREAK_RE.sub("\n", text)
    text = _XML_ILLEGAL_RE.sub(" ", text).replace("\xa0", " ")
    text = "\n".join(line.rstrip() for line in text.split("\n"))
    text = _BLANK_LINES_RE.sub("\n\n", text)
    return text.strip()


def _compare_text(value: object) -> str:
    return _WHITESPACE_RE.sub(" ", clean_free_text(value)).strip()


def _strip_legacy_zero(text: str) -> str:
    # Older app versions stored text through a normalizer that dropped a trailing '.0'.
    return text[:-2] if text.endswith(".0") else text


def text_equal(left: object, right: object) -> bool:
    """Whitespace-insensitive equality, for comparing against stored history only.

    Older history rows hold flattened commentary (and a dropped trailing '.0'), so a
    value that differs only in line breaks or that legacy '.0' still compares equal.
    """
    a = _compare_text(left)
    b = _compare_text(right)
    return a == b or _strip_legacy_zero(a) == _strip_legacy_zero(b)


def status_compare_key(value: object) -> str:
    """Key used to compare statuses across meetings (legacy-'.0' tolerant)."""
    return _strip_legacy_zero(norm_label(value))


def canon_header(value: object) -> str:
    return norm_label(value).upper()


_CURRENCY_MARKER_RE = re.compile(r"\(\s*(?:us\s*\$|\$\s*us|usd|\$)\s*\)|\$", re.IGNORECASE)
_CAMEL_BOUNDARY_RE = re.compile(r"(?<=[a-z])(?=[A-Z])")
_HEADER_WORD_RE = re.compile(r"[a-z0-9]+")


def loose_header_key(value: object) -> str:
    """Header key that ignores case, punctuation, '$' / '(USD)' and common abbreviations.

    'Orig. Loan Amt ($)', 'LoanAmount' and 'loan_amount' all give 'original loan amount'
    or 'loan amount'. Other parentheticals are kept, so 'Loan Amount ($000s)' (a value
    in thousands) never matches 'Loan Amount'.
    """
    text = _CAMEL_BOUNDARY_RE.sub(" ", norm_label(value))
    text = _CURRENCY_MARKER_RE.sub(" ", text.casefold().replace("&", " and "))
    words = [word for word in _HEADER_WORD_RE.findall(text) if word != "usd"]
    return " ".join(HEADER_ABBREVIATIONS.get(word, word) for word in words)


def norm_deal_key(value: object) -> str:
    """Normalize a 'Sheet::DealNumber' key (keeps any '#r<row>' duplicate suffix)."""
    text = norm_id(value)
    if "::" not in text:
        return text
    sheet, _sep, rest = text.partition("::")
    number, marker, row = rest.partition("#r")
    return f"{norm_label(sheet)}::{norm_id(number)}{marker}{row}"


def deal_match_key(value: object) -> str:
    """Join key between the deck and the loan-mod extract.

    All-digit numbers compare without leading zeros ('00010002' == 10002); other keys
    compare case-insensitively. Stored deal numbers and deal keys are never changed.
    """
    text = norm_id(value)
    if text.isdigit():
        return text.lstrip("0") or "0"
    return text.casefold()


def display_text(value: object, blank: str = "-") -> str:
    text = norm_label(value)
    return text if text else blank


_MARKDOWN_SPECIAL_RE = re.compile(r"([\\`*_{}\[\]<>#|~$:])")


def md_escape(value: object) -> str:
    """Workbook text made literal for Streamlit Markdown (toasts, captions, alerts):
    '*', '_', '[', '$' (LaTeX), ':' (icon and color directives) and friends are
    backslash-escaped."""
    return _MARKDOWN_SPECIAL_RE.sub(r"\\\1", norm_label(value))


def _plural(count: int, singular: str, plural: Optional[str] = None) -> str:
    """'1 deal' / '3 deals' (thousands separated)."""
    return f"{count:,} {singular if count == 1 else (plural or singular + 's')}"


def _rows_text(rows: Sequence[int], limit: int = 6) -> str:
    rows = list(rows)
    text = ", ".join(str(row) for row in rows[:limit])
    return text + (f" and {len(rows) - limit} more" if len(rows) > limit else "")


def _quoted_list(names: Sequence[str], limit: int = 6) -> str:
    names = list(names)
    text = ", ".join(f"'{name}'" for name in names[:limit])
    return text + (f" and {len(names) - limit} more" if len(names) > limit else "")


def get_override_key(sheet: str, deal_number: str) -> str:
    return f"{norm_label(sheet)}::{norm_id(deal_number)}"


def first_existing_file(paths: Iterable[Path]) -> Optional[Path]:
    for path in paths:
        if path.exists():
            return path
    return None


def value_is_blank(value: object) -> bool:
    return norm_label(value) == ""


def first_nonblank_value(*values: object) -> object:
    for value in values:
        if not value_is_blank(value):
            return value
    return None


def flag_array(values: Iterable[object]) -> np.ndarray:
    """Booleans from a column that may hold NaN/None (missing counts as False)."""
    return np.array([False if is_missing(value) else bool(value) for value in values], dtype=bool)


def bytes_sha256(data: bytes) -> str:
    return hashlib.sha256(data or b"").hexdigest()


def frame_sha256(frame: Optional[pd.DataFrame]) -> str:
    """Content hash of a DataFrame (column names + values), for explicit cache keys."""
    digest = hashlib.sha256()
    if frame is None or frame.empty:
        return digest.hexdigest()
    digest.update("\x1f".join(map(str, frame.columns)).encode("utf-8"))
    values = pd.util.hash_pandas_object(frame.astype(str), index=False)
    digest.update(values.to_numpy().tobytes())
    return digest.hexdigest()


# -----------------------------------------------------------------------------
# Number and date helpers
# -----------------------------------------------------------------------------
def _finite_number(value: object) -> Optional[float]:
    if is_missing(value) or isinstance(value, (bool, np.bool_)):
        return None
    try:
        number = float(value)
    except Exception:
        return None
    if not math.isfinite(number):
        return None
    return number


def fmt_money(value: object, decimals: int = 0, blank: str = "-") -> str:
    """Compact money: '$12.5MM' from 999,999.5 up, else '$250,000'; sign before '$'."""
    number = _finite_number(value)
    if number is None:
        return blank
    if round(abs(number), decimals) >= 1_000_000:
        text = f"{abs(number) / 1_000_000:,.1f}"
        sign = "-" if number < 0 else ""
        return f"{sign}${text}MM"
    return fmt_money_exact(number, decimals=decimals, blank=blank)


def fmt_money_exact(value: object, decimals: int = 0, blank: str = "-") -> str:
    """Exact money at any size ('$10,250,000', '-$250,000'); -0 shows as '$0'."""
    number = _finite_number(value)
    if number is None:
        return blank
    text = f"{abs(number):,.{decimals}f}"
    sign = "-" if number < 0 and any(ch in "123456789" for ch in text) else ""
    return f"{sign}${text}"


def fmt_int(value: object, blank: str = "-") -> str:
    number = _finite_number(value)
    if number is None:
        return blank
    return f"{int(round(number)):,}"


_NUMERIC_TEXT_RE = re.compile(r"^[+-]?\d+(?:\.\d+)?$")
_MONEY_TEXT_RE = re.compile(r"^\(?-?\$?\s*-?[\d,]*\.?\d+\)?$")
_YEARLESS_DATE_TEXT_RE = re.compile(r"^\d{1,2}[/-]\d{1,2}$")
_TZ_WORD_RE = r"(?<=\d)\s*(?:UTC|GMT|Z)$"


def coerce_number(value: object) -> float:
    """Number from a cell: numbers as-is, '$1,250,000' / '(1,000)' text parsed, else NaN."""
    if is_missing(value) or isinstance(value, (bool, np.bool_)):
        return np.nan
    if isinstance(value, numbers.Number):
        number = float(value)
        return number if math.isfinite(number) else np.nan
    if isinstance(value, str):
        text = value.strip().replace(" ", "")
        if not text or not _MONEY_TEXT_RE.match(text):
            return np.nan
        negative = text.startswith("(") and text.endswith(")")
        text = text.strip("()").replace("$", "").replace(",", "")
        try:
            number = float(text)
        except ValueError:
            return np.nan
        return -abs(number) if negative else number
    return np.nan


def coerce_number_series(values: object) -> pd.Series:
    series = values if isinstance(values, pd.Series) else pd.Series(values, dtype=object)
    if pd.api.types.is_numeric_dtype(series.dtype) and not pd.api.types.is_bool_dtype(series.dtype):
        return pd.to_numeric(series, errors="coerce").astype(float)
    return pd.Series([coerce_number(v) for v in series.astype(object)], index=series.index, dtype=float)


def _number_to_timestamp(number: float) -> Optional[pd.Timestamp]:
    """Excel serials (20000-80000) and YYYYMMDD integers; every other number is no date."""
    if EXCEL_SERIAL_MIN <= number <= EXCEL_SERIAL_MAX:
        return EXCEL_EPOCH + pd.Timedelta(days=number)
    if float(number).is_integer() and 19000101 <= number <= 21991231:
        try:
            return pd.Timestamp(dt.datetime.strptime(str(int(number)), "%Y%m%d"))
        except ValueError:
            return None
    return None


def _in_date_window(stamp: pd.Timestamp) -> bool:
    return DATE_MIN <= stamp < DATE_MAX_EXCLUSIVE


def _windowed_naive(parsed: pd.Series) -> np.ndarray:
    """tz-aware or naive datetimes -> naive datetime64[ns] array, NaT outside the window."""
    if getattr(parsed.dt, "tz", None) is not None:
        parsed = parsed.dt.tz_convert("UTC").dt.tz_localize(None)
    inside = (parsed >= DATE_MIN) & (parsed < DATE_MAX_EXCLUSIVE)
    # copy=True: under pandas 3 copy-on-write, to_numpy() can return a read-only view.
    return parsed.where(inside).astype("datetime64[ns]").to_numpy(copy=True)


def _parse_date_strings(texts: List[str]) -> np.ndarray:
    strings = pd.Series(texts, dtype=object)
    iso_ready = strings.str.replace(_TZ_WORD_RE, "+00:00", regex=True, case=False)
    parsed = pd.to_datetime(iso_ready, errors="coerce", format="ISO8601", utc=True)
    out = _windowed_naive(parsed)
    missing = np.isnat(out)
    if missing.any():
        rest_positions = np.flatnonzero(missing)
        rest = strings.iloc[rest_positions]
        usable = pd.Series([not _YEARLESS_DATE_TEXT_RE.match(text) for text in rest], index=rest.index, dtype=bool)
        if usable.any():
            candidates = rest[usable.to_numpy()]
            try:
                fallback = pd.to_datetime(candidates, errors="coerce", format="mixed", utc=True)
                out[rest_positions[usable.to_numpy()]] = _windowed_naive(fallback)
            except Exception:
                pass
    return out


def _coerce_datetime_core(values: object) -> Tuple[pd.Series, np.ndarray]:
    """Parse any column of date-like cells; never raises.

    Returns (naive datetime64[ns] Series, rejected mask). A value is rejected when it is
    not blank but did not become a date in 1900-2199: unreadable text, a number outside
    the Excel-serial window, a 3026 typo or a 12/31/9999 sentinel.
    """
    if isinstance(values, pd.Series):
        series = values
    else:
        series = pd.Series(list(values) if not isinstance(values, (list, tuple, np.ndarray)) else values, dtype=object)
    index = series.index
    count = len(series)
    if count == 0:
        return pd.Series([], index=index, dtype="datetime64[ns]"), np.zeros(0, dtype=bool)

    if pd.api.types.is_datetime64_any_dtype(series.dtype):
        present = series.notna().to_numpy()
        out = _windowed_naive(series)
        return pd.Series(out, index=index), present & np.isnat(out)

    objects = series.astype(object).to_numpy()
    result = np.full(count, np.datetime64("NaT", "ns"))
    blank = np.zeros(count, dtype=bool)
    text_pos: List[int] = []
    text_vals: List[str] = []
    stamp_pos: List[int] = []
    stamp_vals: List[object] = []

    for pos, value in enumerate(objects):
        if is_missing(value):
            blank[pos] = True
            continue
        if isinstance(value, str):
            text = value.strip()
            if not text:
                blank[pos] = True
                continue
            if _NUMERIC_TEXT_RE.match(text):
                stamp = _number_to_timestamp(float(text))
                if stamp is not None and _in_date_window(stamp):
                    result[pos] = stamp.to_datetime64()
                continue
            text_pos.append(pos)
            text_vals.append(text)
            continue
        if isinstance(value, (bool, np.bool_)):
            continue
        if isinstance(value, (pd.Timestamp, dt.datetime, dt.date, np.datetime64)):
            stamp_pos.append(pos)
            stamp_vals.append(value)
            continue
        if isinstance(value, numbers.Number):
            try:
                stamp = _number_to_timestamp(float(value))
            except (TypeError, ValueError, OverflowError):
                stamp = None
            if stamp is not None and _in_date_window(stamp):
                result[pos] = stamp.to_datetime64()
            continue
        # time-of-day, timedelta and other objects are not dates

    if stamp_pos:
        try:
            parsed = pd.to_datetime(pd.Series(stamp_vals, dtype=object), errors="coerce", utc=True)
            result[np.asarray(stamp_pos)] = _windowed_naive(parsed)
        except Exception:
            for pos, value in zip(stamp_pos, stamp_vals):
                stamp = _coerce_datetime_scalar(value)
                if not pd.isna(stamp):
                    result[pos] = stamp.to_datetime64()
    if text_pos:
        try:
            result[np.asarray(text_pos)] = _parse_date_strings(text_vals)
        except Exception:
            pass

    out = pd.Series(result, index=index)
    rejected = ~blank & np.isnat(result)
    return out, rejected


def _coerce_datetime_scalar(value: object):
    """Fast path for one datetime-like/number value (strings go through the series path)."""
    try:
        if isinstance(value, (pd.Timestamp, dt.datetime, dt.date, np.datetime64)):
            stamp = pd.Timestamp(value)
            if pd.isna(stamp):
                return pd.NaT
            if stamp.tzinfo is not None:
                stamp = stamp.tz_convert("UTC").tz_localize(None)
            return stamp if _in_date_window(stamp) else pd.NaT
        if isinstance(value, numbers.Number) and not isinstance(value, (bool, np.bool_)):
            number = float(value)
            if not math.isfinite(number):
                return pd.NaT
            stamp = _number_to_timestamp(number)
            return stamp if stamp is not None and _in_date_window(stamp) else pd.NaT
    except (TypeError, ValueError, OverflowError, pd.errors.OutOfBoundsDatetime):
        return pd.NaT
    return pd.NaT


def coerce_datetime_series(values: object) -> pd.Series:
    """Naive datetime64[ns] Series from mixed cells (dates, serials, text, tz strings)."""
    return _coerce_datetime_core(values)[0]


def coerce_datetime_value(value: object):
    if is_missing(value) or isinstance(value, (bool, np.bool_)):
        return pd.NaT
    if isinstance(value, str):
        return coerce_datetime_series(pd.Series([value], dtype=object)).iloc[0]
    return _coerce_datetime_scalar(value)


def fmt_date(value: object, blank: str = "-") -> str:
    timestamp = coerce_datetime_value(value)
    if pd.isna(timestamp):
        return blank
    return pd.Timestamp(timestamp).strftime("%m/%d/%Y")


def fmt_cell_text(value: object) -> str:
    """Display text for a raw workbook value: dates as MM/DD/YYYY, '' for blanks."""
    if is_missing(value):
        return ""
    if isinstance(value, (pd.Timestamp, dt.datetime, dt.date, np.datetime64)):
        return fmt_date(value, blank="")
    return norm_label(value)


def fmt_percent(value: object, decimals: int = 2, blank: str = "-") -> str:
    number = _finite_number(value)
    if number is None:
        return blank
    if abs(number) <= 1.5:
        number *= 100
    return f"{number:,.{decimals}f}%"


def fmt_day_delta(days: object, blank: str = "-") -> str:
    number = _finite_number(days)
    if number is None:
        return blank
    days_i = int(round(number))
    if days_i == 0:
        return "Today"
    if days_i > 0:
        return f"In {days_i:,}d"
    return f"{abs(days_i):,}d overdue"


def maturity_buckets(days: object, window: int = MATURITY_WATCH_DAYS) -> Tuple[int, int]:
    """(maturing within 0..window days, already past maturity) from days-to-maturity."""
    numbers_series = pd.to_numeric(pd.Series(days, dtype=object), errors="coerce")
    due = int(numbers_series.between(0, window).sum())
    past = int((numbers_series < 0).sum())
    return due, past


def safe_series(df: pd.DataFrame, column_name: Optional[str], default: object = "") -> pd.Series:
    if column_name and column_name in df.columns:
        return df[column_name]
    return pd.Series([default] * len(df), index=df.index, dtype=object)


def safe_numeric_series(df: pd.DataFrame, column_name: Optional[str]) -> pd.Series:
    if column_name and column_name in df.columns:
        return coerce_number_series(df[column_name])
    return pd.Series([np.nan] * len(df), index=df.index, dtype=float)


def safe_datetime_series(df: pd.DataFrame, column_name: Optional[str]) -> pd.Series:
    if column_name and column_name in df.columns:
        return coerce_datetime_series(df[column_name])
    return pd.Series([pd.NaT] * len(df), index=df.index, dtype="datetime64[ns]")


def resolve_column(columns: Iterable[object], *candidates: Optional[str]) -> Optional[str]:
    cols = [str(c) for c in columns if c is not None]
    for candidate in candidates:
        if not candidate:
            continue
        wanted = canon_header(candidate)
        for col in cols:
            if canon_header(col) == wanted:
                return col
    return None


def matching_columns(columns: Iterable[object], candidate: str) -> List[str]:
    wanted = canon_header(candidate)
    return [str(c) for c in columns if c is not None and canon_header(c) == wanted]


def parse_header_date(
    month: int,
    day: int,
    year_text: Optional[str],
    reference_date: dt.date,
) -> Optional[dt.date]:
    """Date of a dated header such as '10/15 UPB' or '10/15/26 UPB'.

    An explicit year always wins (two-digit years are 20xx). A year-less header takes the
    year that puts it closest to the reference date (ties go to the past), so '12/31'
    next to '1/7/27' is 2026-12-31 and '1/7' in late December is next January.
    """
    if year_text:
        year = int(year_text)
        if year < 100:
            year += 2000
        try:
            return dt.date(year, month, day)
        except ValueError:
            return None
    best: Optional[Tuple[Tuple[int, int], dt.date]] = None
    for year in (reference_date.year - 1, reference_date.year, reference_date.year + 1):
        try:
            candidate = dt.date(year, month, day)
        except ValueError:
            continue
        delta = (candidate - reference_date).days
        rank = (abs(delta), 1 if delta > 0 else 0)
        if best is None or rank < best[0]:
            best = (rank, candidate)
    return best[1] if best else None


def dated_header_parts(header: object) -> Optional[Tuple[int, int, Optional[str], str]]:
    match = DATE_HDR.match(norm_label(header))
    if not match:
        return None
    month, day = int(match.group(1)), int(match.group(2))
    if not (1 <= month <= 12 and 1 <= day <= 31):
        return None
    return month, day, match.group(3), canon_header(match.group(4))


def choose_dated_column(candidates: Sequence[Tuple[dt.date, str]], as_of: dt.date) -> Optional[str]:
    """Newest (date, label) on or before as_of; when none is, the earliest after it."""
    if not candidates:
        return None
    on_or_before = [item for item in candidates if item[0] <= as_of]
    if on_or_before:
        return max(on_or_before, key=lambda item: item[0])[1]
    return min(candidates, key=lambda item: item[0])[1]


def latest_dated_by_suffix(
    headers: Iterable[object],
    as_of_date: Optional[dt.date] = None,
    reference_date: Optional[dt.date] = None,
) -> Dict[str, str]:
    """{suffix: header} picking, per suffix, the newest dated column on or before as_of."""
    as_of_date = as_of_date or dt.date.today()
    reference_date = reference_date or as_of_date
    by_suffix: Dict[str, List[Tuple[dt.date, str]]] = {}
    for header in headers:
        parts = dated_header_parts(header)
        if parts is None:
            continue
        month, day, year_text, suffix = parts
        header_date = parse_header_date(month, day, year_text, reference_date)
        if header_date is None:
            continue
        by_suffix.setdefault(suffix, []).append((header_date, norm_label(header)))
    return {
        suffix: choose_dated_column(items, as_of_date)
        for suffix, items in by_suffix.items()
    }


def fallback_name(
    primary: pd.Series,
    secondary: pd.Series,
    tertiary: pd.Series,
    final_fallback: pd.Series,
) -> pd.Series:
    result = primary.map(norm_label)
    for candidate in [secondary, tertiary, final_fallback]:
        mask = result == ""
        if not mask.any():
            break
        result = result.where(~mask, candidate.map(norm_label))
    return result


# -----------------------------------------------------------------------------
# Status history handling
# -----------------------------------------------------------------------------
def empty_status_history() -> pd.DataFrame:
    return pd.DataFrame(columns=STATUS_HISTORY_COLUMNS)


def _history_saved_at_text(value: object) -> str:
    if isinstance(value, (pd.Timestamp, dt.datetime)) and not is_missing(value):
        return pd.Timestamp(value).strftime("%Y-%m-%d %H:%M:%S")
    return norm_label(value)


def normalize_status_history(history_df: Optional[pd.DataFrame]) -> pd.DataFrame:
    if history_df is None or history_df.empty:
        return empty_status_history()

    out = history_df.copy()
    out.columns = [norm_label(col) for col in out.columns]
    out = out.loc[:, ~pd.Index(out.columns).duplicated()]
    for col in STATUS_HISTORY_COLUMNS:
        if col not in out.columns:
            out[col] = ""

    out = out[STATUS_HISTORY_COLUMNS].copy()
    meeting_dates = coerce_datetime_series(out["meeting_date"].astype(object))
    out["sheet"] = out["sheet"].map(norm_label)
    out["deal_number"] = out["deal_number"].map(norm_id)
    out["deal_key"] = out["deal_key"].map(norm_deal_key)
    out["status"] = out["status"].map(norm_label)
    out["owner"] = out["owner"].map(norm_label)
    out["commentary"] = out["commentary"].map(clean_free_text)
    out["workbook_name"] = out["workbook_name"].map(norm_label)
    out["saved_at"] = out["saved_at"].map(_history_saved_at_text)

    keep = (
        meeting_dates.notna().to_numpy()
        & (out["deal_key"] != "").to_numpy()
        & (out["status"] != "").to_numpy()
    )
    out = out[keep].copy()
    if out.empty:
        return empty_status_history()

    out["meeting_date"] = meeting_dates[keep].dt.strftime("%Y-%m-%d").to_numpy()
    out = out.astype(object)
    out = out.drop_duplicates(subset=["meeting_date", "deal_key"], keep="last")
    out = out.sort_values(["meeting_date", "sheet", "deal_number"], kind="stable")
    return out.reset_index(drop=True)


def read_status_history_csv(file_obj) -> pd.DataFrame:
    try:
        df = pd.read_csv(file_obj, dtype=str, keep_default_na=False)
    except UnicodeDecodeError:
        file_obj.seek(0)
        df = pd.read_csv(file_obj, encoding="latin-1", dtype=str, keep_default_na=False)
    return normalize_status_history(df)


def store_status_history(history: pd.DataFrame, source_name: str) -> None:
    """Keep the normalized history plus its content hash (the streak cache key)."""
    st.session_state.status_history = history
    st.session_state.status_history_sha = frame_sha256(history)
    st.session_state["_status_history_obj_id"] = id(history)
    st.session_state.uploaded_status_history_name = source_name
    st.session_state.status_history_load_error = ""


def current_status_history_sha() -> str:
    """Content hash of the session's normalized history (recomputed only when the
    history object is replaced), used as the streak cache key."""
    history = st.session_state.get("status_history")
    sha = st.session_state.get("status_history_sha")
    if not sha or st.session_state.get("_status_history_obj_id") != id(history):
        sha = frame_sha256(history)
        st.session_state.status_history_sha = sha
        st.session_state["_status_history_obj_id"] = id(history)
    return sha


def load_default_status_history_once() -> None:
    if st.session_state.get("status_history_default_checked", False):
        return

    st.session_state.status_history_default_checked = True
    if not STATUS_HISTORY_PATH.exists():
        return

    try:
        history = pd.read_csv(STATUS_HISTORY_PATH, dtype=str, keep_default_na=False)
        store_status_history(normalize_status_history(history), STATUS_HISTORY_PATH.name)
    except Exception as exc:
        st.session_state.status_history_load_error = str(exc)


def current_status_history_snapshot(
    deck: pd.DataFrame,
    meeting_date: dt.date,
    workbook_name: str,
) -> pd.DataFrame:
    if deck.empty:
        return empty_status_history()

    source = deck
    if "is_duplicate" in source.columns:
        # Only the first occurrence of a duplicated deal number carries the deal's
        # history; '#r<row>' keys never reach _Status History.
        source = source[~flag_array(source["is_duplicate"])]
    snapshot = source[["sheet", "deal_number", "deal_key", "status", "owner", "commentary"]].copy()
    snapshot.insert(0, "meeting_date", meeting_date.isoformat())
    snapshot["workbook_name"] = workbook_name
    snapshot["saved_at"] = now_text()
    snapshot["status"] = snapshot["status"].map(norm_label)
    snapshot = snapshot[snapshot["status"] != ""].copy()
    return normalize_status_history(snapshot)


def merge_status_history(
    existing_history: pd.DataFrame,
    deck: pd.DataFrame,
    meeting_date: dt.date,
    workbook_name: str,
) -> pd.DataFrame:
    existing = normalize_status_history(existing_history)
    current = current_status_history_snapshot(deck, meeting_date, workbook_name)
    merged = pd.concat([existing, current], ignore_index=True)
    return normalize_status_history(merged)


def status_history_to_csv_bytes(history_df: pd.DataFrame) -> bytes:
    history = normalize_status_history(history_df)
    return history.to_csv(index=False).encode("utf-8")


def embedded_status_history(parsed: Optional[Dict[str, Any]]) -> pd.DataFrame:
    """Normalized _Status History from a parsed workbook (tolerant sheet name)."""
    if not parsed or parsed.get("error"):
        return empty_status_history()
    sheets = parsed.get("hidden_sheets") or {}
    tab, _how = resolve_sheet_name(list(sheets), STATUS_HISTORY_SHEET_NAME)
    if tab is None:
        return empty_status_history()
    table = sheets[tab].get("table")
    if table is None or table.empty:
        return empty_status_history()
    return normalize_status_history(table)


def read_embedded_status_history_from_workbook_bytes(file_bytes: bytes) -> pd.DataFrame:
    """Read hidden/embedded status history from the uploaded workbook, if present."""
    return embedded_status_history(parse_workbook_bytes(file_bytes))


def load_embedded_status_history_if_available(
    parsed: Optional[Dict[str, Any]],
    wb_sha: str,
    workbook_name: str,
) -> None:
    """
    Auto-load status history from the workbook itself.

    Manual CSV upload still wins. Embedded workbook history is used only when the user
    has not uploaded a separate CSV during this session. History is normalized once per
    workbook content (sha256), not on every rerun.
    """
    if st.session_state.get("uploaded_status_history") is not None:
        return

    signature = f"embedded:{wb_sha}"
    if signature == st.session_state.get("embedded_status_history_signature", ""):
        return

    embedded = embedded_status_history(parsed)
    st.session_state.embedded_status_history_signature = signature

    if embedded.empty:
        store_status_history(empty_status_history(), "")
        return

    store_status_history(embedded, f"Embedded in {workbook_name}")


def write_status_history_sheet_to_workbook_bytes(file_bytes: bytes, history_df: pd.DataFrame) -> bytes:
    from workbook_io import replace_hidden_sheets
    return replace_hidden_sheets(file_bytes, {STATUS_HISTORY_SHEET_NAME: normalize_status_history(history_df)})

def build_status_prompt_message(row: pd.Series) -> str:
    if not bool(row.get("status_needs_update_prompt", False)):
        return ""

    current_update = display_text(row.get("status"), blank="blank")
    count = int(row.get("status_same_meeting_count") or 0)
    since = display_text(row.get("status_same_since"), blank="unknown")
    return status_prompt_text(count, since, current_update)


def status_prompt_text(count: int, since: str, current_update: str) -> str:
    return (
        f"This deal has carried the same meeting update for {count} meetings since {since}. "
        f"Confirm whether '{current_update}' still reflects the latest story, or refresh the update/commentary."
    )


def leading_run_table(
    history: pd.DataFrame,
    current: Sequence[Tuple[str, str]],
    meeting_date: dt.date,
    value_col: str = "status",
) -> pd.DataFrame:
    """Per deal, how many consecutive meetings (ending at the meeting date) kept its value.

    history: normalized history (ISO meeting_date strings). current: (deal_key, value)
    pairs for this meeting; blank values are not recorded. The meeting calendar is every
    distinct meeting date on or before the meeting date; rows dated later are ignored,
    and a deal missing from a meeting in the calendar breaks its run.
    Returns a frame indexed by deal_key with 'count' (int) and 'since' (Timestamp).
    """
    meeting = pd.Timestamp(meeting_date).normalize()
    cur = pd.DataFrame(list(current), columns=["deal_key", "value"]).astype(object)
    cur["value"] = cur["value"].map(norm_label)
    cur = cur[(cur["deal_key"].astype(str) != "") & (cur["value"] != "")]
    cur = cur.drop_duplicates("deal_key", keep="last")
    result_index = pd.Index(cur["deal_key"].tolist(), name="deal_key")
    empty = pd.DataFrame(
        {"count": pd.Series(0, index=result_index, dtype="int64"),
         "since": pd.Series(pd.NaT, index=result_index, dtype="datetime64[ns]")}
    )
    if cur.empty:
        return empty

    if history is not None and not history.empty and value_col in history.columns:
        past = pd.DataFrame(
            {
                "deal_key": history["deal_key"].astype(object).to_numpy(),
                "value": history[value_col].astype(object).to_numpy(),
                "md": pd.to_datetime(history["meeting_date"].astype(str), format="%Y-%m-%d", errors="coerce").to_numpy(),
            }
        )
        past = past[past["md"].notna() & (past["md"] <= meeting)]
        replaced = (past["md"] == meeting) & past["deal_key"].isin(set(cur["deal_key"]))
        past = past[~replaced]
    else:
        past = pd.DataFrame({"deal_key": [], "value": [], "md": pd.Series([], dtype="datetime64[ns]")})

    now_rows = pd.DataFrame({"deal_key": cur["deal_key"].to_numpy(), "value": cur["value"].to_numpy(), "md": meeting})
    merged = pd.concat([past.astype({"deal_key": object, "value": object}), now_rows], ignore_index=True)
    merged["md"] = pd.to_datetime(merged["md"]).astype("datetime64[ns]")

    # Meeting calendar: every distinct meeting date (any deal) up to the meeting date,
    # newest first; rank 0 is this meeting.
    calendar_rank = merged["md"].rank(method="dense", ascending=False).astype("int64") - 1
    merged = merged.assign(_rank=calendar_rank.to_numpy())
    merged = merged[merged["deal_key"].isin(set(cur["deal_key"]))]
    merged = merged.sort_values(["deal_key", "md"], ascending=[True, False], kind="stable")
    position = merged.groupby("deal_key", sort=False).cumcount()
    target = merged["deal_key"].map(dict(zip(cur["deal_key"], cur["value"].map(status_compare_key))))
    same = merged["value"].map(status_compare_key).eq(target)
    consecutive = merged["_rank"].to_numpy() == position.to_numpy()
    ok = pd.Series(same.to_numpy() & consecutive, index=merged.index).astype("int64")
    run = ok.groupby(merged["deal_key"], sort=False).cumprod().astype(bool)

    counts = run.groupby(merged["deal_key"], sort=False).sum()
    since = merged["md"].where(run).groupby(merged["deal_key"], sort=False).min()
    out = empty.copy()
    out["count"] = counts.reindex(result_index).fillna(0).astype("int64").to_numpy()
    out["since"] = since.reindex(result_index).astype("datetime64[ns]").to_numpy()
    return out


def status_streak_table(
    history: pd.DataFrame,
    current_pairs: Sequence[Tuple[str, str]],
    meeting_date_iso: str,
    threshold: int = STATUS_STALE_MEETING_THRESHOLD,
) -> pd.DataFrame:
    """Streak columns per deal_key (count, since text, prompt flag, prompt message)."""
    meeting = dt.date.fromisoformat(meeting_date_iso)
    runs = leading_run_table(history, current_pairs, meeting, value_col="status")
    status_by_key = {key: norm_label(value) for key, value in current_pairs}
    since_text = [
        "" if pd.isna(value) else pd.Timestamp(value).strftime("%m/%d/%Y")
        for value in runs["since"]
    ]
    counts = runs["count"].astype("int64")
    needs = counts >= threshold
    messages = [
        status_prompt_text(int(count), since or "unknown", status_by_key.get(key) or "blank") if flag else ""
        for key, count, since, flag in zip(runs.index, counts, since_text, needs)
    ]
    return pd.DataFrame(
        {
            "status_same_meeting_count": counts.to_numpy(),
            "status_same_since": since_text,
            "status_needs_update_prompt": needs.to_numpy(dtype=bool),
            "status_prompt_message": messages,
        },
        index=runs.index,
    )


@st.cache_data(max_entries=16, show_spinner=False)
def compute_status_streaks(
    _history_norm: pd.DataFrame,
    history_sha: str,
    current_pairs: Tuple[Tuple[str, str], ...],
    meeting_date_iso: str,
    threshold: int = STATUS_STALE_MEETING_THRESHOLD,
) -> pd.DataFrame:
    """Cached streak table; keyed on the history content hash and the current statuses."""
    return status_streak_table(_history_norm, current_pairs, meeting_date_iso, threshold)


def add_status_staleness_columns(
    deck: pd.DataFrame,
    status_history: pd.DataFrame,
    meeting_date: dt.date,
    workbook_name: str = "",
    threshold: int = STATUS_STALE_MEETING_THRESHOLD,
    history_sha: Optional[str] = None,
) -> pd.DataFrame:
    if deck.empty:
        return deck.copy()

    out = deck.copy()
    eligible = out
    if "is_duplicate" in out.columns:
        eligible = out[~flag_array(out["is_duplicate"])]
    current_pairs = tuple(
        (str(key), norm_label(status))
        for key, status in zip(eligible["deal_key"], eligible["status"])
    )
    meeting_iso = meeting_date.isoformat()
    if history_sha is None:
        # Uncached path (scripts, tests): accept any history and normalize it here.
        table = status_streak_table(normalize_status_history(status_history), current_pairs, meeting_iso, threshold)
    else:
        # The app passes history already normalized once per content sha.
        table = compute_status_streaks(status_history, history_sha, current_pairs, meeting_iso, threshold)

    keys = out["deal_key"].astype(str).tolist()
    counts = table["status_same_meeting_count"].to_dict()
    since = table["status_same_since"].to_dict()
    needs = table["status_needs_update_prompt"].to_dict()
    messages = table["status_prompt_message"].to_dict()
    out["status_same_meeting_count"] = np.array([int(counts.get(key, 0)) for key in keys], dtype="int64")
    out["status_same_since"] = [str(since.get(key, "")) for key in keys]
    out["status_needs_update_prompt"] = np.array([bool(needs.get(key, False)) for key in keys], dtype=bool)
    out["status_prompt_message"] = [str(messages.get(key, "")) for key in keys]
    return out


# -----------------------------------------------------------------------------
# Workbook parse (cached per file) and deck derivation (cached per as-of/toggle)
# -----------------------------------------------------------------------------
def sheet_name_key(name: object) -> str:
    return norm_label(name).casefold()


def resolve_sheet_name(
    sheetnames: Sequence[str],
    canonical: str,
    aliases: Sequence[str] = (),
) -> Tuple[Optional[str], str]:
    """Find a tab case- and whitespace-insensitively ('TERM ', 'Loan modifications').

    Returns (actual tab name or None, how): how is 'exact', 'normalized', 'missing' or
    'ambiguous' (two non-exact spellings and no exact one: nothing is guessed).
    """
    names = [str(name) for name in sheetnames]
    if canonical in names:
        return canonical, "exact"
    wanted = {sheet_name_key(canonical), *(sheet_name_key(alias) for alias in aliases)}
    hits = [name for name in names if sheet_name_key(name) in wanted]
    if len(hits) == 1:
        return hits[0], "normalized"
    if len(hits) > 1:
        return None, "ambiguous"
    return None, "missing"


def find_header_in_rows(
    rows: Iterable[Sequence[object]],
    key_headers: Sequence[str] = (KEY_HEADER,),
    search_rows: int = HEADER_SEARCH_ROWS,
) -> Optional[Tuple[int, Tuple[object, ...]]]:
    """(1-based row, values) of the first row holding one of key_headers, if any."""
    keys = {canon_header(key) for key in key_headers}
    for row_number, values in enumerate(rows, start=1):
        if row_number > search_rows:
            break
        if values and any(v is not None and canon_header(v) in keys for v in values):
            return row_number, tuple(values)
    return None


def layout_from_header_values(header_values: Sequence[object], header_row: int) -> Dict[str, Any]:
    """Positional layout of a Bridge/Term header row.

    Columns are read by position. Labels are the header texts with exact duplicates
    renamed '<Header> (2)'; roles resolve through SHEET_ROLE_CANDIDATES (first candidate
    that matches, left-most column wins) and every 'Days Past Due' column is kept.
    """
    raw = list(header_values)
    max_col = 0
    for position, value in enumerate(raw, start=1):
        if norm_label(value):
            max_col = position
    raw = raw[:max_col]

    labels: List[Optional[str]] = []
    used: Dict[str, int] = {}
    renamed: Dict[str, List[int]] = {}
    for position, value in enumerate(raw, start=1):
        label = norm_label(value)
        if not label:
            labels.append(None)
            continue
        if label in used:
            used[label] += 1
            new_label = f"{label} ({used[label]})"
            while new_label in used:
                used[label] += 1
                new_label = f"{label} ({used[label]})"
            used[new_label] = 1
            renamed.setdefault(label, []).append(position)
            labels.append(new_label)
        else:
            used[label] = 1
            labels.append(label)

    canon = [canon_header(value) for value in raw]
    loose = [loose_header_key(value) for value in raw]
    role_cols: Dict[str, Optional[int]] = {}
    role_duplicates: List[Dict[str, Any]] = []
    for role, candidates in SHEET_ROLE_CANDIDATES.items():
        role_cols[role] = None
        for candidate in candidates:
            wanted = canon_header(candidate)
            hits = [pos for pos, text in enumerate(canon, start=1) if text and text == wanted]
            if not hits:
                continue
            role_cols[role] = hits[0]
            if len(hits) > 1 and role in REPORTED_DUPLICATE_ROLES:
                role_duplicates.append(
                    {
                        "role": role,
                        "header": norm_label(raw[hits[0] - 1]),
                        "columns": [get_column_letter(pos) for pos in hits],
                        "used": get_column_letter(hits[0]),
                        # loan-size roles take the first column with a value, per row
                        "coalesced": role in COALESCED_MONEY_ROLES,
                    }
                )
            break

    # Money roles: every matching column in priority order (exact header matches first,
    # then loose ones such as 'Orig. Loan Amt ($)'); the loan-size roles coalesce them
    # per row and the other roles use the first one.
    role_all_cols: Dict[str, List[int]] = {}
    for role in LOOSE_MATCH_ROLES:
        ordered: List[int] = []
        for candidate in SHEET_ROLE_CANDIDATES[role]:
            wanted = canon_header(candidate)
            ordered += [pos for pos, text in enumerate(canon, start=1) if text and text == wanted and pos not in ordered]
        for candidate in SHEET_ROLE_CANDIDATES[role]:
            wanted_loose = loose_header_key(candidate)
            ordered += [
                pos for pos, key in enumerate(loose, start=1) if key and key == wanted_loose and pos not in ordered
            ]
        role_all_cols[role] = ordered
        if role_cols.get(role) is None and ordered:
            role_cols[role] = ordered[0]

    dated: List[Dict[str, Any]] = []
    for position, value in enumerate(raw, start=1):
        parts = dated_header_parts(value)
        if parts is None or labels[position - 1] is None:
            continue
        month, day, year_text, suffix = parts
        dated.append(
            {
                "col": position,
                "label": labels[position - 1],
                "month": month,
                "day": day,
                "year": year_text,
                "suffix": suffix,
            }
        )

    dpd_cols = [pos for pos, text in enumerate(canon, start=1) if text == canon_header(DPD_HEADER)]
    duplicate_headers = [
        {
            "header": label,
            "columns": [get_column_letter(pos) for pos in [labels.index(label) + 1] + positions],
        }
        for label, positions in renamed.items()
    ]
    return {
        "header_row": header_row,
        "max_col": max_col,
        "raw_headers": [_scalar_text(value) for value in raw],
        "labels": labels,
        "role_cols": role_cols,
        "roles": {
            role: (labels[col - 1] if col else None) for role, col in role_cols.items()
        },
        "role_all_cols": role_all_cols,
        "role_all_labels": {
            role: [labels[col - 1] for col in cols if labels[col - 1]] for role, cols in role_all_cols.items()
        },
        "dpd_cols": dpd_cols,
        "dpd_labels": [labels[col - 1] for col in dpd_cols if labels[col - 1]],
        "dated": dated,
        "role_duplicates": role_duplicates,
        "duplicate_headers": duplicate_headers,
    }


def resolve_sheet_layout(
    ws,
    key_headers: Sequence[str] = (KEY_HEADER,),
    search_rows: int = HEADER_SEARCH_ROWS,
) -> Optional[Dict[str, Any]]:
    """Header row, positional columns and role map of a worksheet (None without a header).

    Shared by the loader (read-only worksheets) and the workbook writer (normal mode), so
    an edit lands in exactly the column the deck was read from.
    """
    max_col = None
    max_row = search_rows
    if not hasattr(ws, "reset_dimensions"):
        # Normal-mode worksheets create cells while iterating: stay inside the used range.
        max_col = min(int(ws.max_column or 1), 2048)
        max_row = min(search_rows, int(ws.max_row or 1))
    rows = ws.iter_rows(min_row=1, max_row=max_row, max_col=max_col, values_only=True)
    found = find_header_in_rows(rows, key_headers, search_rows)
    if found is None:
        return None
    header_row, values = found
    return layout_from_header_values(values, header_row)


def _cell_is_blank(value: object) -> bool:
    if value is None:
        return True
    if isinstance(value, str):
        return not value.strip()
    return is_missing(value)


def _title_as_of(rows: Sequence[Sequence[object]]) -> Optional[dt.date]:
    """'as of MM/DD/YYYY' (or 'As of' next to a date cell) in a sheet's title rows."""
    for values in rows[:TITLE_SEARCH_ROWS]:
        cells = list(values or ())
        for position, value in enumerate(cells):
            text = norm_label(value)
            if not text:
                continue
            match = TITLE_AS_OF_RE.search(text)
            if match:
                try:
                    if match.group(1):
                        year = int(match.group(3))
                        year = year + 2000 if year < 100 else year
                        return dt.date(year, int(match.group(1)), int(match.group(2)))
                    return dt.date(int(match.group(4)), int(match.group(5)), int(match.group(6)))
                except ValueError:
                    continue
            if re.fullmatch(r"as\s+of:?", text, flags=re.IGNORECASE):
                for neighbour in cells[position + 1:position + 3]:
                    stamp = coerce_datetime_value(neighbour)
                    if not pd.isna(stamp):
                        return pd.Timestamp(stamp).date()
                continue
            # 'as of October 15, 2026' / 'As of: Oct 15 2026' (a date written out in words).
            written = TITLE_AS_OF_WORDS_RE.search(text)
            if written:
                stamp = coerce_datetime_value(written.group(1))
                if not pd.isna(stamp):
                    return pd.Timestamp(stamp).date()
    return None


_ROW_TAG_RE = re.compile(rb"<(?:[A-Za-z_][\w.-]*:)?row\b([^>]*)>")
_ROW_NUMBER_RE = re.compile(rb"""\br\s*=\s*["'](\d+)["']""")
_ROW_HIDDEN_RE = re.compile(rb"""\bhidden\s*=\s*["'](?:1|true)["']""", re.IGNORECASE)
_CELL_END_RE = re.compile(rb"</(?:[A-Za-z_][\w.-]*:)?c>")
_FORMULA_TAG_RE = re.compile(rb"<(?:[A-Za-z_][\w.-]*:)?f\b")
_VALUE_TAG_RE = re.compile(rb"<(?:[A-Za-z_][\w.-]*:)?v>\s*[^<\s]")
_CELL_OPEN_RE = re.compile(rb"<(?:[A-Za-z_][\w.-]*:)?c\b([^>]*)>")
_CELL_REF_RE = re.compile(rb"""\br\s*=\s*["']([A-Za-z]{1,3})(\d+)["']""")
_CELL_TYPE_RE = re.compile(rb"""\bt\s*=\s*["']([A-Za-z]+)["']""")
_ANY_VALUE_RE = re.compile(rb"<(?:[A-Za-z_][\w.-]*:)?(?:v>\s*[^<\s]|t\b[^>]*>\s*[^<\s])")


def _sheet_part_paths(archive: zipfile.ZipFile) -> Dict[str, str]:
    """Sheet name -> worksheet part path, from xl/workbook.xml and its relationships."""
    try:
        workbook_xml = ET.fromstring(archive.read("xl/workbook.xml"))
        rels_xml = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    except Exception:
        return {}
    targets = {}
    for rel in rels_xml:
        if rel.get("Id"):
            targets[rel.get("Id")] = rel.get("Target") or ""
    paths: Dict[str, str] = {}
    for element in workbook_xml.iter():
        if not element.tag.endswith("}sheet") and element.tag != "sheet":
            continue
        rel_id = next((value for key, value in element.attrib.items() if key.endswith("}id")), None)
        target = targets.get(rel_id or "", "")
        if not target:
            continue
        if target.startswith("/"):
            path = target.lstrip("/")
        else:
            path = posixpath.normpath(posixpath.join("xl", target))
        paths[element.get("name") or ""] = path
    return paths


def _scan_sheet_xml(
    xml_bytes: bytes,
    header_row: int,
    last_row: Optional[int] = None,
) -> Dict[str, Any]:
    """Hidden rows and formula cells, from the raw sheet XML (read-only worksheets expose
    neither row_dimensions nor formulas next to their saved values).

    Covers rows up to last_row (the whole sheet when None). Formula cells below the header
    are tallied per column letter: 'cached' when the file holds a saved result, else the
    row is listed under 'uncached' (openpyxl-written formulas have none, so a data_only
    read sees blanks). 'values_below' tells whether any value sits after last_row.
    """
    hidden: set = set()
    counter = 0
    stop_at = len(xml_bytes)
    for match in _ROW_TAG_RE.finditer(xml_bytes):
        attrs = match.group(1)
        number = _ROW_NUMBER_RE.search(attrs)
        counter = int(number.group(1)) if number else counter + 1
        if last_row is not None and counter > last_row:
            stop_at = match.start()
            break
        if _ROW_HIDDEN_RE.search(attrs):
            hidden.add(counter)

    formulas: Dict[str, Dict[str, Any]] = {}
    region = xml_bytes[:stop_at]
    for formula in _FORMULA_TAG_RE.finditer(region):
        position = formula.start()
        # The enclosing <c ...> (or <x:c ...>) tag: the nearest cell start before the
        # formula. Text content cannot hold a literal '<', so rfind is safe here.
        plain = region.rfind(b"<c ", 0, position)
        prefixed = region.rfind(b":c ", 0, position)
        start = region.rfind(b"<", 0, prefixed) if prefixed > plain else plain
        opening = _CELL_OPEN_RE.match(region, start) if start >= 0 else None
        if opening is None or opening.end() > position:
            continue
        attrs = opening.group(1)
        reference = _CELL_REF_RE.search(attrs)
        if reference is None:
            continue
        row_number = int(reference.group(2))
        if row_number <= header_row:
            continue
        letter = reference.group(1).decode("ascii").upper()
        end = _CELL_END_RE.search(region, position)
        inner = region[position:end.start() if end else len(region)]
        cell_type = _CELL_TYPE_RE.search(attrs)
        # A saved result is a non-empty <v>, or any typed result (t="str" with an empty
        # <v> is Excel's saved empty string); openpyxl writes formulas untyped with <v/>.
        cached = bool(_VALUE_TAG_RE.search(inner)) or (
            cell_type is not None and cell_type.group(1).lower() in {b"str", b"e", b"b", b"inlinestr"}
        )
        entry = formulas.setdefault(letter, {"cached": 0, "uncached": []})
        if cached:
            entry["cached"] += 1
        else:
            entry["uncached"].append(row_number)

    values_below = bool(last_row is not None and stop_at < len(xml_bytes) and _ANY_VALUE_RE.search(xml_bytes, stop_at))
    return {"hidden_rows": hidden, "formulas": formulas, "values_below": values_below}


def _read_data_rows(ws, start_row: int, max_col: int) -> Tuple[List[Tuple[int, Tuple[object, ...]]], Optional[int]]:
    """Non-blank rows from start_row on (bounded by max_col), stopping after
    BLANK_ROW_STOP consecutive fully blank rows. Returns (rows, row where reading stopped)."""
    rows: List[Tuple[int, Tuple[object, ...]]] = []
    blank_run = 0
    for offset, values in enumerate(
        ws.iter_rows(min_row=start_row, max_col=max_col, values_only=True)
    ):
        excel_row = start_row + offset
        values = tuple(values or ())
        if not values or all(_cell_is_blank(value) for value in values):
            blank_run += 1
            if blank_run >= BLANK_ROW_STOP:
                return rows, excel_row
            continue
        blank_run = 0
        if len(values) < max_col:
            values = values + (None,) * (max_col - len(values))
        rows.append((excel_row, values[:max_col]))
    return rows, None


def _count_nonblank_rows(ws, start_row: int) -> int:
    count = 0
    blank_run = 0
    for values in ws.iter_rows(min_row=start_row, values_only=True):
        if not values or all(_cell_is_blank(value) for value in values):
            blank_run += 1
            if blank_run >= BLANK_ROW_STOP:
                break
            continue
        blank_run = 0
        count += 1
    return count


def _classify_deal_cell(value: object) -> Tuple[str, str]:
    """('deal' | 'blank' | 'error' | 'total', normalized deal number)."""
    text = norm_id(value)
    if not text:
        return "blank", ""
    if EXCEL_ERROR_RE.match(text):
        return "error", text
    if TOTAL_ROW_RE.match(text):
        return "total", text
    return "deal", text


def _issue_samples(items: List[Any]) -> List[Any]:
    return items[:MAX_ISSUE_SAMPLES]


def _record_rejected_dates(
    report: Dict[str, Any],
    raw_values: pd.Series,
    rejected: np.ndarray,
    excel_rows: Sequence[int],
    header: str,
) -> None:
    if not rejected.any():
        return
    report["rejected_date_count"] += int(rejected.sum())
    for position in np.flatnonzero(rejected):
        if len(report["rejected_dates"]) >= MAX_ISSUE_SAMPLES:
            break
        report["rejected_dates"].append(
            {
                "excel_row": int(excel_rows[position]),
                "header": header,
                "value": _scalar_text(raw_values.iloc[position])[:60],
            }
        )


def _empty_sheet_report(sheet: str) -> Dict[str, Any]:
    return {
        "sheet": sheet,
        "status": "missing",
        "tab": None,
        "tab_match": "missing",
        "note": "",
        "layout": None,
        "title_as_of": None,
        "raw": None,
        "base": None,
        "rows_read": 0,
        "blank_stop_row": None,
        "skipped_rows": [],
        "skipped_count": 0,
        "duplicate_deals": [],
        "hidden_rows": 0,
        "formula_deal_uncached": [],
        "formula_deal_cached": 0,
        "rejected_dates": [],
        "rejected_date_count": 0,
        "nonblank_rows_without_header": 0,
        "values_below_stop": False,
        "formula_uncached_columns": [],
        "loan_size": {},
    }


def _parse_portfolio_sheet(
    workbook,
    archive: Optional[zipfile.ZipFile],
    part_paths: Dict[str, str],
    sheet: str,
) -> Dict[str, Any]:
    report = _empty_sheet_report(sheet)
    tab, how = resolve_sheet_name(workbook.sheetnames, sheet)
    report["tab_match"] = how
    if tab is None:
        if how == "ambiguous":
            candidates = [n for n in workbook.sheetnames if sheet_name_key(n) == sheet_name_key(sheet)]
            report["status"] = "ambiguous"
            report["note"] = f"Several tabs could be {sheet} ({', '.join(repr(n) for n in candidates)}); none was loaded."
        return report

    report["tab"] = tab
    ws = workbook[tab]
    if hasattr(ws, "reset_dimensions"):
        ws.reset_dimensions()

    top_rows = [tuple(values or ()) for values in ws.iter_rows(min_row=1, max_row=HEADER_SEARCH_ROWS, values_only=True)]
    found = find_header_in_rows(top_rows)
    if found is None:
        nonblank = sum(1 for values in top_rows if any(not _cell_is_blank(v) for v in values))
        if nonblank == 0:
            nonblank = _count_nonblank_rows(ws, HEADER_SEARCH_ROWS + 1)
        if nonblank == 0:
            report["status"] = "blank"
            report["note"] = f"'{tab}' tab is empty, so there are no {sheet} deals this week."
        else:
            report["status"] = "no_header"
            report["nonblank_rows_without_header"] = nonblank
            report["note"] = (
                f"'{tab}' tab skipped: no '{KEY_HEADER}' header in its first {HEADER_SEARCH_ROWS} rows "
                f"({_plural(nonblank, 'non-empty row')} not loaded)."
            )
        return report

    header_row, header_values = found
    layout = layout_from_header_values(header_values, header_row)
    report["layout"] = layout
    report["title_as_of"] = _title_as_of(top_rows[: header_row - 1])
    max_col = max(layout["max_col"], 1)

    data_rows, stop_row = _read_data_rows(ws, header_row + 1, max_col)
    report["blank_stop_row"] = stop_row
    report["rows_read"] = len(data_rows)

    deal_col = layout["role_cols"].get("deal")
    scan: Dict[str, Any] = {"hidden_rows": set(), "formulas": {}, "values_below": False}
    part = part_paths.get(tab)
    if archive is not None and part:
        try:
            scan = _scan_sheet_xml(archive.read(part), header_row, stop_row)
        except Exception:
            pass
    hidden_rows = scan["hidden_rows"]
    report["values_below_stop"] = bool(scan["values_below"])
    deal_letter = get_column_letter(deal_col) if deal_col else ""
    deal_formulas = scan["formulas"].get(deal_letter, {"cached": 0, "uncached": []})
    report["formula_deal_uncached"] = sorted(deal_formulas["uncached"])
    report["formula_deal_cached"] = int(deal_formulas["cached"])

    kept: List[Tuple[int, Tuple[object, ...], str]] = []
    for excel_row, values in data_rows:
        kind, deal_text = _classify_deal_cell(values[deal_col - 1] if deal_col else None)
        if kind == "deal":
            kept.append((excel_row, values, deal_text))
        elif kind in {"error", "total"}:
            report["skipped_rows"].append({"excel_row": excel_row, "reason": kind, "value": deal_text[:40]})
    report["skipped_count"] = len(report["skipped_rows"])
    report["skipped_rows"] = _issue_samples(report["skipped_rows"])

    # Deal fields whose formulas have no saved result read as blank (U25, loader side);
    # totals and other non-deal rows are not counted.
    deal_rows = {row for row, _values, _deal in kept}
    uncached_columns: List[Dict[str, Any]] = []
    for letter, entry in scan["formulas"].items():
        rows = sorted(row for row in entry["uncached"] if row in deal_rows)
        if letter == deal_letter or not rows:
            continue
        try:
            position = column_index_from_string(letter)
        except ValueError:
            continue
        if position > max_col or not layout["labels"][position - 1]:
            continue
        uncached_columns.append(
            {
                "header": layout["labels"][position - 1],
                "column": letter,
                "count": len(rows),
                "rows": _issue_samples(rows),
            }
        )
    report["formula_uncached_columns"] = sorted(uncached_columns, key=lambda item: column_index_from_string(item["column"]))

    labels = layout["labels"]
    label_positions = [(pos, label) for pos, label in enumerate(labels) if label]
    raw = pd.DataFrame(
        [[values[pos] for pos, _label in label_positions] for _row, values, _deal in kept],
        columns=[label for _pos, label in label_positions],
        dtype=object,
    )
    excel_rows = [row for row, _values, _deal in kept]
    raw["_excel_row"] = pd.Series(excel_rows, dtype="int64")
    raw["_hidden"] = pd.Series([row in hidden_rows for row in excel_rows], dtype=bool)
    report["hidden_rows"] = int(raw["_hidden"].sum()) if len(raw) else 0
    report["raw"] = raw
    report["status"] = "ok" if len(raw) else "no_rows"
    report["base"] = _build_sheet_base(sheet, layout, raw, [deal for _r, _v, deal in kept], report)
    return report


def _role_values(raw: pd.DataFrame, layout: Dict[str, Any], role: str) -> pd.Series:
    label = layout["roles"].get(role)
    if label and label in raw.columns:
        return raw[label]
    return pd.Series([None] * len(raw), index=raw.index, dtype=object)


def _role_label_map(raw: pd.DataFrame, layout: Dict[str, Any], role: str, default: str = "") -> pd.Series:
    values = _role_values(raw, layout, role).map(norm_label)
    if default:
        values = values.where(values != "", default)
    return values


def _build_sheet_base(
    sheet: str,
    layout: Dict[str, Any],
    raw: pd.DataFrame,
    deal_numbers: List[str],
    report: Dict[str, Any],
) -> pd.DataFrame:
    """As-of-independent deck columns for one sheet (identity, labels, dates, money)."""
    count = len(raw)
    base = pd.DataFrame(index=raw.index)
    excel_rows = raw["_excel_row"].astype("int64")
    hidden = raw["_hidden"].astype(bool)

    deal_series = pd.Series(deal_numbers, index=raw.index, dtype=object)
    borrower = _role_values(raw, layout, "borrower")
    account = _role_values(raw, layout, "account")
    base["sheet"] = [sheet] * count
    base["sheet_order"] = SHEET_ORDER.get(sheet, len(SHEET_ORDER))
    base["original_order"] = excel_rows.to_numpy()
    base["_excel_row"] = excel_rows.to_numpy()
    base["_hidden"] = hidden.to_numpy()
    base["deal_number"] = deal_series.to_numpy()
    base["deal_name"] = fallback_name(_role_values(raw, layout, "name"), borrower, account, deal_series).to_numpy()
    base["borrower"] = borrower.map(norm_label).to_numpy()
    base["account_display"] = account.map(norm_label).to_numpy()
    base["servicer"] = _role_label_map(raw, layout, "servicer").to_numpy()
    base["portfolio"] = _role_label_map(raw, layout, "portfolio", default=sheet).to_numpy()
    base["segment"] = _role_label_map(raw, layout, "segment").to_numpy()
    base["financing"] = _role_label_map(raw, layout, "financing").to_numpy()
    base["loan_buyer"] = _role_label_map(raw, layout, "loan_buyer").to_numpy()
    base["owner"] = _role_label_map(raw, layout, "owner").to_numpy()
    base["status"] = _role_label_map(raw, layout, "status").to_numpy()
    base["commentary"] = _role_values(raw, layout, "commentary").map(clean_free_text).to_numpy()
    base["saved_at"] = ""

    # Identity: the first visible copy of a deal number keeps 'Sheet::Number'; any other
    # copy is 'Sheet::Number#r<excel_row>' so every agenda row has its own key.
    first_copy: Dict[str, int] = {}
    for position in sorted(range(count), key=lambda i: (bool(hidden.iloc[i]), int(excel_rows.iloc[i]))):
        first_copy.setdefault(deal_numbers[position], position)
    is_duplicate = [first_copy[deal_numbers[i]] != i for i in range(count)]
    base["is_duplicate"] = is_duplicate
    base["deal_key"] = [
        f"{sheet}::{number}" + (f"#r{int(excel_rows.iloc[i])}" if is_duplicate[i] else "")
        for i, number in enumerate(deal_numbers)
    ]
    groups: Dict[str, List[int]] = {}
    for i, number in enumerate(deal_numbers):
        groups.setdefault(number, []).append(i)
    report["duplicate_deals"] = [
        {
            "deal_number": number,
            "excel_rows": sorted(int(excel_rows.iloc[i]) for i in positions),
            "kept_row": int(excel_rows.iloc[first_copy[number]]),
            "hidden_rows": sorted(int(excel_rows.iloc[i]) for i in positions if bool(hidden.iloc[i])),
        }
        for number, positions in groups.items()
        if len(positions) > 1
    ]

    for role, column in (("maturity", "maturity_date"), ("next_payment", "next_payment_date")):
        values = _role_values(raw, layout, role)
        parsed, rejected = _coerce_datetime_core(values.astype(object))
        base[column] = parsed.to_numpy()
        if layout["roles"].get(role):
            _record_rejected_dates(report, values, rejected, excel_rows.tolist(), layout["roles"][role])

    # Days past due: the largest numeric value across every 'Days Past Due' column; NaN
    # (shown '-') when no DPD cell in the row holds a number.
    dpd_labels = [label for label in layout["dpd_labels"] if label in raw.columns]
    if dpd_labels and count:
        matrix = np.column_stack([coerce_number_series(raw[label]).to_numpy(dtype=float) for label in dpd_labels])
        best = np.where(np.isnan(matrix), -np.inf, matrix).max(axis=1)
        base["days_past_due"] = np.where(np.isneginf(best), np.nan, np.maximum(best, 0.0))
    else:
        base["days_past_due"] = np.nan

    # Loan size ('Loan Commitment' KPI): the explicit commitment column(s); for Term the
    # original loan amount when there is no commitment (DSCR loans fund in full at closing).
    commitment_labels = _role_labels_present(raw, layout, "commitment")
    explicit, explicit_header = _coalesce_amounts(raw, commitment_labels)
    size_labels = list(commitment_labels)
    if sheet == "Bridge":
        base["funded_amount"] = coerce_number_series(_role_values(raw, layout, "funded")).to_numpy()
        base["commitment"] = explicit
        base["commitment_source"] = np.where(np.isnan(explicit), "", COMMITMENT_SOURCE_EXPLICIT)
        base["commitment_header"] = explicit_header
        base["remaining_commitment"] = coerce_number_series(_role_values(raw, layout, "remaining")).to_numpy()
        base["loan_amount"] = np.nan
    else:
        loan_labels = _role_labels_present(raw, layout, "loan_amount")
        size_labels += loan_labels
        loan_amount, loan_header = _coalesce_amounts(raw, loan_labels)
        # A positive explicit commitment wins; a 0 / negative one only when there is no
        # positive loan amount to fall back to.
        use_explicit = (explicit > 0) | (~np.isnan(explicit) & ~(loan_amount > 0))
        has_loan = ~np.isnan(loan_amount)
        base["loan_amount"] = loan_amount
        base["funded_amount"] = np.nan
        base["commitment"] = np.where(use_explicit, explicit, loan_amount)
        base["commitment_source"] = np.where(
            use_explicit,
            COMMITMENT_SOURCE_EXPLICIT,
            np.where(has_loan, COMMITMENT_SOURCE_LOAN_AMOUNT, ""),
        )
        base["commitment_header"] = np.where(use_explicit, explicit_header, np.where(has_loan, loan_header, ""))
        base["remaining_commitment"] = np.nan

    missing = np.isnan(base["commitment"].to_numpy(dtype=float)) if count else np.zeros(0, dtype=bool)
    report["loan_size"] = {
        "headers": size_labels,
        "rows": int(count),
        "missing_count": int(missing.sum()),
        "missing_rows": _issue_samples([int(row) for row in excel_rows[missing].tolist()]),
    }
    return base.reset_index(drop=True)


def _role_labels_present(raw: pd.DataFrame, layout: Dict[str, Any], role: str) -> List[str]:
    """Every column label for a money role, in priority order (see layout role_all_labels)."""
    labels = (layout.get("role_all_labels") or {}).get(role)
    if labels is None:
        labels = [layout["roles"][role]] if layout["roles"].get(role) else []
    return [label for label in labels if label in raw.columns]


def _coalesce_amounts(raw: pd.DataFrame, labels: Sequence[str]) -> Tuple[np.ndarray, np.ndarray]:
    """(amount, source header) per row over money columns in priority order.

    The first positive amount wins; when no column holds one, the first present value (a
    0 or a negative) is kept; otherwise NaN and ''.
    """
    count = len(raw)
    values = np.full(count, np.nan)
    headers = np.full(count, "", dtype=object)
    if not labels or not count:
        return values, headers
    matrix = np.column_stack([coerce_number_series(raw[label]).to_numpy(dtype=float) for label in labels])
    with np.errstate(invalid="ignore"):
        positive = matrix > 0
    present = ~np.isnan(matrix)
    chosen = np.where(
        positive.any(axis=1),
        positive.argmax(axis=1),
        np.where(present.any(axis=1), present.argmax(axis=1), -1),
    )
    rows = np.flatnonzero(chosen >= 0)
    values[rows] = matrix[rows, chosen[rows]]
    headers[rows] = [labels[index] for index in chosen[rows]]
    return values, headers


def _parse_loan_mod_sheet(workbook) -> Tuple[pd.DataFrame, Dict[str, Any]]:
    info: Dict[str, Any] = {
        "status": "missing",
        "tab": None,
        "tab_match": "missing",
        "key_header": None,
        "rows_read": 0,
    }
    tab, how = resolve_sheet_name(workbook.sheetnames, LOAN_MOD_SHEET_NAME)
    info["tab_match"] = how
    if tab is None:
        info["status"] = how
        return pd.DataFrame(), info
    info["tab"] = tab
    ws = workbook[tab]
    if hasattr(ws, "reset_dimensions"):
        ws.reset_dimensions()

    found = find_header_in_rows(
        ws.iter_rows(min_row=1, max_row=HEADER_SEARCH_ROWS, values_only=True),
        key_headers=LOAN_MOD_KEY_HEADERS,
    )
    if found is None:
        info["status"] = "no_header"
        return pd.DataFrame(), info

    header_row, header_values = found
    wanted = {canon_header(col): col for col in LOAN_MOD_CORE_COLUMNS}
    selected: List[Tuple[int, str]] = []
    taken = set()
    for position, value in enumerate(header_values):
        key = canon_header(value)
        if key in wanted and key not in taken:
            taken.add(key)
            selected.append((position, wanted[key]))
    present_keys = [name for name in LOAN_MOD_KEY_HEADERS if canon_header(name) in taken]
    info["key_header"] = present_keys[0] if present_keys else None
    if not selected:
        info["status"] = "no_header"
        return pd.DataFrame(), info

    max_col = max(position for position, _name in selected) + 1
    data_rows, stop_row = _read_data_rows(ws, header_row + 1, max_col)
    records = []
    excel_rows = []
    for excel_row, values in data_rows:
        record = [values[position] if position < len(values) else None for position, _name in selected]
        if all(_cell_is_blank(value) for value in record):
            continue
        records.append(record)
        excel_rows.append(excel_row)
    raw = pd.DataFrame(records, columns=[name for _position, name in selected], dtype=object)
    raw["_excel_row"] = pd.Series(excel_rows, dtype="int64")
    info["status"] = "ok"
    info["rows_read"] = len(raw)
    info["blank_stop_row"] = stop_row
    return raw, info


def _parse_underscore_sheets(workbook) -> Dict[str, Dict[str, Any]]:
    """Every sheet whose name starts with '_' as a raw table (row 1 is the header)."""
    sheets: Dict[str, Dict[str, Any]] = {}
    for name in workbook.sheetnames:
        if not norm_label(name).startswith("_"):
            continue
        ws = workbook[name]
        if hasattr(ws, "reset_dimensions"):
            ws.reset_dimensions()
        header: Optional[Tuple[object, ...]] = None
        records: List[Tuple[object, ...]] = []
        blank_run = 0
        for values in ws.iter_rows(values_only=True):
            values = tuple(values or ())
            if header is None:
                header = values
                continue
            if not values or all(_cell_is_blank(value) for value in values):
                blank_run += 1
                if blank_run >= BLANK_ROW_STOP:
                    break
                continue
            blank_run = 0
            records.append(values)
        columns: List[Tuple[int, str]] = []
        seen = set()
        for position, value in enumerate(header or ()):
            label = norm_label(value)
            if not label or label in seen:
                continue
            seen.add(label)
            columns.append((position, label))
        table = pd.DataFrame(
            [[row[position] if position < len(row) else None for position, _label in columns] for row in records],
            columns=[label for _position, label in columns],
            dtype=object,
        )
        sheets[name] = {
            "tab": name,
            "state": getattr(ws, "sheet_state", "visible"),
            "headers": [label for _position, label in columns],
            "table": table,
        }
    return sheets


def _reference_date(sheets: Dict[str, Dict[str, Any]]) -> Tuple[Optional[dt.date], str]:
    """Workbook date: title-row 'as of', else the newest dated header that has a year."""
    title_dates = [report["title_as_of"] for report in sheets.values() if report.get("title_as_of")]
    if title_dates:
        return max(title_dates), "title"
    header_dates: List[dt.date] = []
    for report in sheets.values():
        layout = report.get("layout") or {}
        for column in layout.get("dated", []):
            if column["year"]:
                parsed = parse_header_date(column["month"], column["day"], column["year"], dt.date(2000, 1, 1))
                if parsed:
                    header_dates.append(parsed)
    if header_dates:
        return max(header_dates), "header"
    return None, ""


def parse_workbook_bytes(file_bytes: bytes) -> Dict[str, Any]:
    """One read-only openpyxl pass over a workbook; never raises.

    Reads the Bridge/Term layouts and rows (with Excel row numbers and hidden flags), the
    Loan Modifications extract and every '_' sheet. Nothing here depends on the as-of
    date or the hidden-rows toggle, so the result is cached per file content.
    """
    parsed: Dict[str, Any] = {
        "error": None,
        "sheetnames": [],
        "sheet_states": {},
        "sheets": {},
        "reference_date": None,
        "reference_source": "",
        "loan_mods": empty_loan_modifications(),
        "loan_mod_info": {"status": "missing"},
        "hidden_sheets": {},
        "rejected_dates": 0,
    }
    try:
        workbook = load_workbook(io.BytesIO(file_bytes), read_only=True, data_only=True)
    except Exception as exc:
        parsed["error"] = f"{type(exc).__name__}: {exc}"
        return parsed

    archive: Optional[zipfile.ZipFile] = None
    try:
        try:
            archive = zipfile.ZipFile(io.BytesIO(file_bytes))
            part_paths = _sheet_part_paths(archive)
        except Exception:
            archive, part_paths = None, {}

        parsed["sheetnames"] = list(workbook.sheetnames)
        parsed["sheet_states"] = {
            name: getattr(workbook[name], "sheet_state", "visible") for name in workbook.sheetnames
        }
        # Each part is isolated: a malformed sheet is reported and the others still load.
        for sheet in PORTFOLIO_SHEETS:
            try:
                parsed["sheets"][sheet] = _parse_portfolio_sheet(workbook, archive, part_paths, sheet)
            except Exception as exc:
                report = _empty_sheet_report(sheet)
                report["status"] = "error"
                report["note"] = f"{sheet} could not be read ({type(exc).__name__}); its deals are not in this deck."
                parsed["sheets"][sheet] = report

        mod_info: Dict[str, Any] = {"status": "missing"}
        try:
            raw_mods, mod_info = _parse_loan_mod_sheet(workbook)
            mods, normalize_info = normalize_loan_modifications_with_info(raw_mods)
            mod_info.update(normalize_info)
            parsed["loan_mods"] = mods
        except Exception as exc:
            mod_info = {"status": "error", "note": f"Loan modifications could not be read ({type(exc).__name__})."}
        parsed["loan_mod_info"] = mod_info

        try:
            parsed["supplemental"] = data_views.parse_supplemental(workbook)
        except Exception as exc:
            parsed["supplemental"] = {}
            LOG.warning("event=supplemental_read_failed type=%s", type(exc).__name__)

        try:
            parsed["hidden_sheets"] = _parse_underscore_sheets(workbook)
        except Exception:
            parsed["hidden_sheets"] = {}

        reference, source = _reference_date(parsed["sheets"])
        parsed["reference_date"] = reference.isoformat() if reference else None
        parsed["reference_source"] = source
        parsed["rejected_dates"] = int(
            sum(int(report.get("rejected_date_count") or 0) for report in parsed["sheets"].values())
            + int(mod_info.get("rejected_date_count", 0) or 0)
        )
    except Exception as exc:
        parsed["error"] = f"{type(exc).__name__}: {exc}"
    finally:
        workbook.close()
        if archive is not None:
            archive.close()
    return parsed


@st.cache_data(ttl="8h", max_entries=6, show_spinner="Reading workbook…", show_time=True)
def parse_workbook(_file_bytes: bytes, wb_sha: str) -> Dict[str, Any]:
    """Cached parse, keyed on the workbook's sha256 (the bytes themselves are not hashed)."""
    return parse_workbook_bytes(_file_bytes)


def _resolve_dated_choice(
    layout: Dict[str, Any],
    suffix: str,
    as_of: dt.date,
    reference: dt.date,
) -> Tuple[Optional[str], List[Tuple[dt.date, str]]]:
    columns: List[Tuple[dt.date, str]] = []
    for column in layout.get("dated", []):
        if column["suffix"] != suffix:
            continue
        header_date = parse_header_date(column["month"], column["day"], column["year"], reference)
        if header_date is not None:
            columns.append((header_date, column["label"]))
    columns.sort(key=lambda item: item[0])
    return choose_dated_column(columns, as_of), columns


def build_deck(
    parsed: Dict[str, Any],
    as_of: dt.date,
    include_hidden: bool,
) -> Tuple[pd.DataFrame, Dict[str, Any]]:
    """Deck rows for an as-of date: pick each sheet's dated UPB/NPL columns, filter hidden
    rows, add days-to-maturity/payment and the loan-mod columns."""
    info: Dict[str, Any] = {"as_of": as_of.isoformat(), "sheets": {}}
    reference = (
        dt.date.fromisoformat(parsed["reference_date"]) if parsed.get("reference_date") else as_of
    )
    frames: List[pd.DataFrame] = []
    for sheet in PORTFOLIO_SHEETS:
        report = parsed.get("sheets", {}).get(sheet) or {}
        base = report.get("base")
        raw = report.get("raw")
        layout = report.get("layout")
        if base is None or raw is None or layout is None or base.empty:
            continue
        keep = np.ones(len(base), dtype=bool) if include_hidden else ~base["_hidden"].to_numpy(dtype=bool)
        frame = base.loc[keep].copy()
        raw_kept = raw.loc[keep].reset_index(drop=True)
        frame = frame.reset_index(drop=True)

        upb_label, upb_columns = _resolve_dated_choice(layout, "UPB", as_of, reference)
        npl_label, npl_columns = _resolve_dated_choice(layout, "NPL", as_of, reference)
        roles = layout["roles"]
        upb_candidates = [upb_label, roles.get("upb")]
        if sheet == "Bridge":
            upb_candidates.append(roles.get("funded"))
        upb_candidates.append(roles.get("loan_amount"))
        upb_col = next((label for label in upb_candidates if label and label in raw_kept.columns), None)
        npl_col = next(
            (label for label in [npl_label, roles.get("npl")] if label and label in raw_kept.columns),
            None,
        )
        frame["upb"] = safe_numeric_series(raw_kept, upb_col).to_numpy()
        frame["npl_raw"] = (
            raw_kept[npl_col].map(fmt_cell_text).to_numpy() if npl_col else [""] * len(frame)
        )
        newest_upb = upb_columns[-1][1] if upb_columns else None
        newest_npl = npl_columns[-1][1] if npl_columns else None
        info["sheets"][sheet] = {
            "tab": report.get("tab"),
            "upb_header": upb_col or "",
            "npl_header": npl_col or "",
            "newest_upb_header": newest_upb or "",
            "newest_npl_header": newest_npl or "",
            "dated_columns": {
                suffix: [(day.isoformat(), label) for day, label in cols]
                for suffix, cols in (("UPB", upb_columns), ("NPL", npl_columns))
                if cols
            },
            "maturity_header": roles.get("maturity") or "",
            "next_payment_header": roles.get("next_payment") or "",
            "owner_header": roles.get("owner") or "",
            "status_header": roles.get("status") or "",
            "commentary_header": roles.get("commentary") or "",
            "uses_older_column": bool(
                (upb_label and newest_upb and upb_label != newest_upb)
                or (npl_label and newest_npl and npl_label != newest_npl)
            ),
        }
        frames.append(frame)

    if not frames:
        return pd.DataFrame(), info

    deck = pd.concat(frames, ignore_index=True)
    as_of_ts = pd.Timestamp(as_of)
    deck["maturity_date"] = pd.to_datetime(deck["maturity_date"]).astype("datetime64[ns]")
    deck["next_payment_date"] = pd.to_datetime(deck["next_payment_date"]).astype("datetime64[ns]")
    deck["days_to_maturity"] = (deck["maturity_date"] - as_of_ts).dt.days
    deck["days_to_next_payment"] = (deck["next_payment_date"] - as_of_ts).dt.days
    deck = deck.sort_values(["sheet_order", "original_order"], kind="stable").reset_index(drop=True)
    deck = add_loan_modification_columns(deck, parsed.get("loan_mods"))
    return deck, info


@st.cache_data(max_entries=12, show_spinner=False)
def derive_deck(
    _parsed: Dict[str, Any],
    wb_sha: str,
    as_of_iso: str,
    include_hidden: bool,
) -> Tuple[pd.DataFrame, Dict[str, Any]]:
    """Cached deck for (workbook sha, as-of, hidden toggle); never re-reads the file."""
    return build_deck(_parsed, dt.date.fromisoformat(as_of_iso), include_hidden)


def workbook_metadata(parsed: Dict[str, Any], deck_info: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
    """Flat summary of parse and derive notes (what the UI and later preflight read)."""
    meta: Dict[str, Any] = {
        "error": parsed.get("error"),
        "reference_date": parsed.get("reference_date"),
        "reference_source": parsed.get("reference_source"),
        "rejected_dates": int(parsed.get("rejected_dates") or 0),
        "sheetnames": list(parsed.get("sheetnames") or []),
        "loan_mods": dict(parsed.get("loan_mod_info") or {}),
    }
    for sheet, report in (parsed.get("sheets") or {}).items():
        layout = report.get("layout") or {}
        entry = {
            key: report.get(key)
            for key in (
                "status",
                "tab",
                "tab_match",
                "note",
                "rows_read",
                "blank_stop_row",
                "skipped_rows",
                "skipped_count",
                "duplicate_deals",
                "hidden_rows",
                "formula_deal_uncached",
                "formula_deal_cached",
                "formula_uncached_columns",
                "values_below_stop",
                "loan_size",
                "rejected_dates",
                "rejected_date_count",
            )
        }
        entry["title_as_of"] = report["title_as_of"].isoformat() if report.get("title_as_of") else None
        entry["header_row"] = layout.get("header_row")
        entry["role_duplicates"] = layout.get("role_duplicates", [])
        entry["duplicate_headers"] = layout.get("duplicate_headers", [])
        entry["writable_columns"] = {
            role: (get_column_letter(layout["role_cols"][role]) if layout.get("role_cols", {}).get(role) else None)
            for role in ("status", "owner", "commentary")
        } if layout else {}
        if deck_info and sheet in deck_info.get("sheets", {}):
            entry.update(deck_info["sheets"][sheet])
        meta[sheet] = entry
    return meta


def load_portfolio_workbook(
    file_bytes: bytes,
    as_of_date_iso: str,
    include_hidden: bool,
) -> Tuple[pd.DataFrame, Dict[str, Any]]:
    """Uncached parse + derive (tests, scripts). The app uses parse_workbook/derive_deck."""
    parsed = parse_workbook_bytes(file_bytes)
    if parsed.get("error"):
        raise ValueError(f"Could not read workbook: {parsed['error']}")
    deck, info = build_deck(parsed, dt.date.fromisoformat(as_of_date_iso), include_hidden)
    return deck, workbook_metadata(parsed, info)


# -----------------------------------------------------------------------------
# Loan Modification loading and presentation helpers
# -----------------------------------------------------------------------------
LOAN_MOD_DECK_DEFAULTS: Dict[str, object] = {
    "loan_mod_count": 0,
    "loan_mod_active_count": 0,
    "loan_mod_inactive_count": 0,
    "loan_mod_latest_summary": "",
    "loan_mod_latest_status": "",
    "loan_mod_latest_type": "",
    "loan_mod_latest_name": "",
    "loan_mod_latest_order_number": "",
    "loan_mod_latest_effective_date": pd.NaT,
    "loan_mod_latest_completed_date": pd.NaT,
    "loan_mod_latest_finalized_date": pd.NaT,
    "loan_mod_latest_previous_maturity_date": pd.NaT,
    "loan_mod_latest_updated_maturity_date": pd.NaT,
    "loan_mod_latest_previous_commitment": np.nan,
    "loan_mod_latest_updated_commitment": np.nan,
    "loan_mod_latest_comments": "",
    "loan_mod_latest_request_note": "",
    "loan_mod_search_text": "",
}


def empty_loan_modifications() -> pd.DataFrame:
    return pd.DataFrame(
        columns=[
            "deal_number_key",
            "Deal Number",
            "Deal Name",
            "Loan Modification Name",
            "Status",
            "Loan Mod Type",
            "Modification Type",
            "Mod Effective Date",
            "Modification Finalized Date",
            "Updated Maturity Date",
            "Comments",
            "Pay Pik Summary",
            "_mod_inactive",
        ]
    )


def loan_mod_type(record: Dict[str, object]) -> str:
    """One type for a mod everywhere: Modification Type, else Loan Mod Type, else Mod
    Reporting Type."""
    return norm_label(first_nonblank_value(*(record.get(field) for field in LOAN_MOD_TYPE_FIELDS)))


def loan_mod_is_inactive(status: object, cancellation_date: object = None) -> bool:
    if LOAN_MOD_INACTIVE_RE.search(norm_label(status)):
        return True
    return not pd.isna(coerce_datetime_value(cancellation_date))


def parse_loan_mod_match_flag(value: object) -> Optional[bool]:
    """'Portfolio Overview Match' cell -> True / False / None (blank or unrecognised)."""
    if is_missing(value):
        return None
    if isinstance(value, (bool, np.bool_)):
        return bool(value)
    if isinstance(value, numbers.Number):
        number = float(value)
        if number == 1:
            return True
        if number == 0:
            return False
        return None
    token = norm_label(value).upper()
    if token in LOAN_MOD_MATCH_TRUE:
        return True
    if token in LOAN_MOD_MATCH_FALSE:
        return False
    return None


def load_workbook_sheet_as_dataframe(
    workbook,
    sheet_name: str,
    key_header: str = "Deal Number",
    desired_columns: Optional[Iterable[str]] = None,
) -> pd.DataFrame:
    """Generic header-detecting reader (first HEADER_SEARCH_ROWS rows, tolerant name)."""
    tab, _how = resolve_sheet_name(workbook.sheetnames, sheet_name)
    if tab is None:
        return pd.DataFrame()

    ws = workbook[tab]
    found = find_header_in_rows(ws.iter_rows(min_row=1, max_row=HEADER_SEARCH_ROWS, values_only=True), (key_header,))
    if found is None:
        return pd.DataFrame(columns=list(desired_columns or []))
    header_row, header_values = found
    wanted = {canon_header(col) for col in desired_columns or [] if norm_label(col)}
    selected: List[Tuple[int, str]] = []
    seen = set()
    for idx, value in enumerate(header_values):
        header = norm_label(value)
        if not header or header in seen:
            continue
        if wanted and canon_header(header) not in wanted:
            continue
        seen.add(header)
        selected.append((idx, header))
    if not selected:
        return pd.DataFrame(columns=list(desired_columns or []))
    max_col = max(idx for idx, _header in selected) + 1
    data_rows, _stop = _read_data_rows(ws, header_row + 1, max_col)
    rows = []
    for _excel_row, values in data_rows:
        record = {header: values[idx] for idx, header in selected}
        if any(not _cell_is_blank(value) for value in record.values()):
            rows.append(record)
    if not rows:
        return pd.DataFrame(columns=[header for _, header in selected])
    return pd.DataFrame(rows)


def _loan_mod_sort(df: pd.DataFrame) -> pd.DataFrame:
    """Latest modification first per deal: highest order number, then newest dates.

    A mod with no order number that was created after the deal's newest numbered mod
    (a new or pending request) ranks above it; older unnumbered mods stay last.
    """
    if "Loan Mod Order Number" in df.columns:
        order = pd.to_numeric(df["Loan Mod Order Number"].map(norm_id), errors="coerce")
    else:
        order = pd.Series(np.nan, index=df.index, dtype=float)
    created = df["Created Date"] if "Created Date" in df.columns else pd.Series(pd.NaT, index=df.index, dtype="datetime64[ns]")
    activity_cols = [
        col
        for col in ["System Task Completed Date", "Modification Finalized Date", "Mod Effective Date"]
        if col in df.columns
    ]
    if activity_cols:
        activity = df[activity_cols].max(axis=1)
    else:
        activity = pd.Series(pd.NaT, index=df.index, dtype="datetime64[ns]")
    newness = pd.to_datetime(created.fillna(activity)).astype("datetime64[ns]")

    rank = order.copy()
    numbered = order.notna()
    if (~numbered).any() and numbered.any():
        keys = df["deal_number_key"]
        newest_numbered = newness.where(numbered).groupby(keys).transform("max")
        top_order = order.groupby(keys).transform("max")
        promote = (~numbered) & newness.notna() & (newest_numbered.isna() | (newness >= newest_numbered)) & top_order.notna()
        rank = rank.where(~promote, top_order + 0.5)

    work = df.assign(_mod_order_sort=rank.to_numpy(), _mod_newness=newness.to_numpy())
    sort_cols = ["_mod_order_sort"] + [
        col
        for col in [
            "System Task Completed Date",
            "Modification Finalized Date",
            "Mod Effective Date",
            "Last Modified Date",
            "Created Date",
        ]
        if col in work.columns
    ]
    work = work.sort_values(
        ["deal_number_key"] + sort_cols,
        ascending=[True] + [False] * len(sort_cols),
        kind="stable",
        na_position="last",
    )
    return work.drop(columns=["_mod_newness"])


def normalize_loan_modifications_with_info(raw_df: Optional[pd.DataFrame]) -> Tuple[pd.DataFrame, Dict[str, Any]]:
    info: Dict[str, Any] = {
        "rows_kept": 0,
        "excluded_by_match": 0,
        "unrecognized_match": 0,
        "match_filter_applied": False,
        "blank_key_rows": 0,
        "rejected_date_count": 0,
        "rejected_dates": [],
    }
    if raw_df is None or raw_df.empty:
        return empty_loan_modifications(), info

    df = raw_df.copy()
    canonical = {canon_header(col): col for col in LOAN_MOD_CORE_COLUMNS + LOAN_MOD_DATE_COLUMNS + LOAN_MOD_MONEY_COLUMNS}
    renamed = []
    for col in df.columns:
        key = canon_header(col)
        renamed.append(canonical.get(key, norm_label(col)) if col != "_excel_row" else col)
    df.columns = renamed
    df = df.loc[:, ~pd.Index(df.columns).duplicated()]

    primary = df["Deal Number"].map(norm_id) if "Deal Number" in df.columns else None
    fallback = (
        df["Opportunity Deal Loan Number"].map(norm_id)
        if "Opportunity Deal Loan Number" in df.columns
        else None
    )
    if primary is None and fallback is None:
        return empty_loan_modifications(), info
    if primary is None:
        primary = fallback
    elif fallback is not None:
        primary = primary.where(primary != "", fallback)
    df["Deal Number"] = primary.astype(object)

    blank_key = df["Deal Number"] == ""
    info["blank_key_rows"] = int(blank_key.sum())
    df = df[~blank_key].copy()

    # If the generator already marked which loan mods match the visible overview, keep
    # only those rows. The filter fails open: no recognised match leaves every row.
    if "Portfolio Overview Match" in df.columns:
        flags = df["Portfolio Overview Match"].map(parse_loan_mod_match_flag)
        recognised_true = flags.map(lambda flag: flag is True).astype(bool)
        info["unrecognized_match"] = int(
            (flags.isna() & df["Portfolio Overview Match"].map(lambda v: not _cell_is_blank(v)).astype(bool)).sum()
        )
        if flags.notna().any():
            info["match_filter_applied"] = True
            info["excluded_by_match"] = int((~recognised_true).sum())
            df = df[recognised_true].copy()

    excel_rows = df["_excel_row"].tolist() if "_excel_row" in df.columns else list(range(len(df)))
    for col in LOAN_MOD_DATE_COLUMNS:
        if col in df.columns:
            parsed, rejected = _coerce_datetime_core(df[col].astype(object))
            if rejected.any():
                info["rejected_date_count"] += int(rejected.sum())
                for position in np.flatnonzero(rejected)[: max(0, MAX_ISSUE_SAMPLES - len(info["rejected_dates"]))]:
                    info["rejected_dates"].append(
                        {
                            "excel_row": int(excel_rows[position]) if position < len(excel_rows) else None,
                            "header": col,
                            "value": _scalar_text(df[col].iloc[position])[:60],
                        }
                    )
            df[col] = parsed.to_numpy()

    for col in LOAN_MOD_MONEY_COLUMNS + LOAN_MOD_PERCENT_COLUMNS:
        if col in df.columns:
            df[col] = coerce_number_series(df[col]).to_numpy()

    for col in ["Loan Mod Order Number"]:
        if col in df.columns:
            df[col] = df[col].map(norm_id)
    for col in ["Comments", "Cancellation Reason"]:
        if col in df.columns:
            df[col] = df[col].map(clean_free_text)
    for col in [
        "Deal Name",
        "Opportunity Deal Name",
        "Loan Modification Name",
        "Status",
        "Loan Mod Type",
        "Modification Type",
        "Mod Reporting Type",
        "Deal Type",
        "Visible Overview Sheet(s)",
    ]:
        if col in df.columns:
            df[col] = df[col].map(norm_label)

    if "Deal Name" not in df.columns:
        df["Deal Name"] = ""
    if "Opportunity Deal Name" in df.columns:
        df["Deal Name"] = df["Deal Name"].where(df["Deal Name"] != "", df["Opportunity Deal Name"])

    df["deal_number_key"] = df["Deal Number"].map(deal_match_key)
    status = df["Status"] if "Status" in df.columns else pd.Series("", index=df.index)
    cancel_date = df["Cancellation Date"] if "Cancellation Date" in df.columns else pd.Series(None, index=df.index, dtype=object)
    df["_mod_inactive"] = [
        loan_mod_is_inactive(s, c) for s, c in zip(status.tolist(), cancel_date.tolist())
    ]
    df = _loan_mod_sort(df)
    info["rows_kept"] = len(df)
    return df.reset_index(drop=True), info


def normalize_loan_modifications(raw_df: pd.DataFrame) -> pd.DataFrame:
    return normalize_loan_modifications_with_info(raw_df)[0]


def load_loan_modifications_from_workbook(file_bytes: bytes) -> pd.DataFrame:
    """Normalized loan mods of a workbook (uncached; the app reads parse_workbook)."""
    return parse_workbook_bytes(file_bytes)["loan_mods"]


def _mod_maturity_phrase(record: Dict[str, object]) -> str:
    """'New maturity' for mods in force, 'Requested maturity' otherwise, nothing for a
    cancelled request or an unchanged date."""
    updated = coerce_datetime_value(record.get("Updated Maturity Date"))
    if pd.isna(updated):
        return ""
    previous = coerce_datetime_value(record.get("Previous Maturity Date"))
    if not pd.isna(previous) and pd.Timestamp(previous).normalize() == pd.Timestamp(updated).normalize():
        return ""
    if bool(record.get("_mod_inactive")) or loan_mod_is_inactive(record.get("Status"), record.get("Cancellation Date")):
        return ""
    done = bool(LOAN_MOD_DONE_RE.search(norm_label(record.get("Status")))) or not pd.isna(
        coerce_datetime_value(record.get("System Task Completed Date"))
    )
    label = "New maturity" if done else "Requested maturity"
    return f"{label} {fmt_date(updated)}"


def build_loan_mod_summary_from_record(record: Dict[str, object]) -> str:
    pieces = []

    mod_type = loan_mod_type(record)
    status = norm_label(record.get("Status"))
    effective = fmt_date(record.get("Mod Effective Date"))
    completed = fmt_date(record.get("System Task Completed Date"))

    if mod_type:
        pieces.append(mod_type)

    if status:
        pieces.append(f"Status: {status}")

    if effective != "-":
        pieces.append(f"Effective {effective}")

    if completed != "-":
        pieces.append(f"Completed {completed}")

    maturity_phrase = _mod_maturity_phrase(record)
    if maturity_phrase:
        pieces.append(maturity_phrase)

    return " | ".join(pieces)


def build_loan_mod_commentary_from_record(record: Dict[str, object]) -> str:
    sections = []

    for label in [
        "Comments",
    ]:
        value = clean_free_text(record.get(label))
        if value:
            sections.append(f"{label}: {value}")

    return "\n\n".join(sections)


def loan_mods_for_deal(loan_mods: pd.DataFrame, deal_number: object) -> pd.DataFrame:
    """Return every loan modification record for a single deal, latest first.

    This is what powers the presentation toggle, so the caller can step through
    each modification on a loan rather than seeing only the most recent one.
    """
    if loan_mods is None or loan_mods.empty or "deal_number_key" not in loan_mods.columns:
        return pd.DataFrame()

    key = deal_match_key(deal_number)
    if not key:
        return pd.DataFrame()

    subset = loan_mods[loan_mods["deal_number_key"] == key].copy()
    return subset.reset_index(drop=True)


def headline_mod_position(mods: pd.DataFrame) -> int:
    """Position of the deal's headline mod: the newest one that is not cancelled."""
    if mods is None or mods.empty or "_mod_inactive" not in mods.columns:
        return 0
    active = np.flatnonzero(~flag_array(mods["_mod_inactive"]))
    return int(active[0]) if len(active) else 0


def build_loan_mod_toggle_labels(mods: pd.DataFrame) -> List[str]:
    """Short, unique labels for toggling between each modification of one deal."""
    labels: List[str] = []
    seen: Dict[str, int] = {}

    for position, (_, record) in enumerate(mods.iterrows(), start=1):
        order = norm_id(record.get("Loan Mod Order Number"))
        base = f"Mod #{order}" if order else f"Mod (no #) {position}"
        if bool(record.get("_mod_inactive", False)):
            status = norm_label(record.get("Status")) or "Inactive"
            base = f"{base} · {status}"

        # Guarantee uniqueness so the toggle control always has distinct options.
        if base in seen:
            seen[base] += 1
            base = f"{base} ({seen[base]})"
        else:
            seen[base] = 1

        labels.append(base)

    return labels


def loan_mod_deal_summary(loan_mods: Optional[pd.DataFrame]) -> pd.DataFrame:
    """One row per deal_number_key with the loan_mod_* deck columns (headline = newest
    active mod; the count includes cancelled requests)."""
    columns = [col for col in LOAN_MOD_DECK_DEFAULTS]
    if loan_mods is None or loan_mods.empty or "deal_number_key" not in loan_mods.columns:
        return pd.DataFrame(columns=columns)

    mods = loan_mods.copy()
    if "_mod_inactive" not in mods.columns:
        mods["_mod_inactive"] = False
    inactive = pd.Series(flag_array(mods["_mod_inactive"]), index=mods.index)
    keys = mods["deal_number_key"]
    counts = keys.value_counts()
    active_counts = keys[~inactive].value_counts()

    ordered = mods.assign(_inactive_rank=inactive.astype(int).to_numpy(), _pos=np.arange(len(mods)))
    headline = ordered.sort_values(["deal_number_key", "_inactive_rank", "_pos"], kind="stable")
    headline = headline.drop_duplicates("deal_number_key", keep="first").set_index("deal_number_key")
    newest = mods.drop_duplicates("deal_number_key", keep="first").set_index("deal_number_key")

    records = headline.to_dict("index")
    newest_records = newest.to_dict("index")
    search_text = (
        pd.Series(
            [
                " ".join(
                    part
                    for part in (loan_mod_type(mod), norm_label(mod.get("Status")), norm_label(mod.get("Comments")))
                    if part
                )
                for mod in mods.to_dict("records")
            ],
            index=mods.index,
            dtype=object,
        )
        .groupby(keys, sort=False)
        .agg(" ".join)
        .to_dict()
    )
    rows = []
    for key, record in records.items():
        latest = newest_records.get(key, record)
        note = ""
        if bool(latest.get("_mod_inactive")) and (
            norm_id(latest.get("Loan Mod Order Number")) != norm_id(record.get("Loan Mod Order Number"))
            or int(active_counts.get(key, 0)) == 0
        ):
            order = norm_id(latest.get("Loan Mod Order Number"))
            label = f"Mod #{order}" if order else "Latest request"
            mod_kind = loan_mod_type(latest)
            note = f"{label}{' ' + mod_kind if mod_kind else ''}: {norm_label(latest.get('Status')) or 'inactive'}"
        rows.append(
            {
                "deal_number_key": key,
                "loan_mod_count": int(counts.get(key, 0)),
                "loan_mod_active_count": int(active_counts.get(key, 0)),
                "loan_mod_inactive_count": int(counts.get(key, 0)) - int(active_counts.get(key, 0)),
                "loan_mod_latest_summary": build_loan_mod_summary_from_record(record),
                "loan_mod_latest_status": norm_label(record.get("Status")),
                "loan_mod_latest_type": loan_mod_type(record),
                "loan_mod_latest_name": norm_label(record.get("Loan Modification Name")),
                "loan_mod_latest_order_number": norm_id(record.get("Loan Mod Order Number")),
                "loan_mod_latest_effective_date": coerce_datetime_value(record.get("Mod Effective Date")),
                "loan_mod_latest_completed_date": coerce_datetime_value(record.get("System Task Completed Date")),
                "loan_mod_latest_finalized_date": coerce_datetime_value(record.get("Modification Finalized Date")),
                "loan_mod_latest_previous_maturity_date": coerce_datetime_value(record.get("Previous Maturity Date")),
                "loan_mod_latest_updated_maturity_date": coerce_datetime_value(record.get("Updated Maturity Date")),
                "loan_mod_latest_previous_commitment": _finite_number(record.get("Previous Loan Commitment")),
                "loan_mod_latest_updated_commitment": _finite_number(record.get("Updated Loan Commitment")),
                "loan_mod_latest_comments": clean_free_text(record.get("Comments")),
                "loan_mod_latest_request_note": note,
                "loan_mod_search_text": search_text.get(key, ""),
            }
        )
    return pd.DataFrame(rows).set_index("deal_number_key")


def add_loan_modification_columns(deck: pd.DataFrame, loan_mods: pd.DataFrame) -> pd.DataFrame:
    """Attach the loan_mod_* columns with one join on the canonical deal key."""
    if deck.empty:
        return deck.copy()

    out = deck.copy()
    summary = loan_mod_deal_summary(loan_mods)
    keys = out["deal_number"].map(deal_match_key)
    for col, default in LOAN_MOD_DECK_DEFAULTS.items():
        if summary.empty or col not in summary.columns:
            mapped = pd.Series([default] * len(out), index=out.index, dtype=object)
        else:
            mapped = keys.map(summary[col])
        if col in {"loan_mod_count", "loan_mod_active_count", "loan_mod_inactive_count"}:
            out[col] = pd.to_numeric(mapped, errors="coerce").fillna(0).astype("int64").to_numpy()
        elif isinstance(default, type(pd.NaT)) or default is pd.NaT:
            out[col] = pd.to_datetime(mapped, errors="coerce").astype("datetime64[ns]").to_numpy()
        elif isinstance(default, float):
            out[col] = pd.to_numeric(mapped, errors="coerce").astype(float).to_numpy()
        else:
            out[col] = mapped.map(lambda value: "" if is_missing(value) else str(value)).to_numpy()
    return out


def has_loan_mods(row: pd.Series) -> bool:
    try:
        return int(row.get("loan_mod_count") or 0) > 0
    except Exception:
        return False


def compact_note(value: object, max_chars: int = 360) -> str:
    text = display_text(value, blank="")
    if not text:
        return ""
    text = re.sub(r"^Comments\s*:\s*", "", text, flags=re.IGNORECASE).strip()
    text = re.sub(r"\s+", " ", text).strip()
    if len(text) <= max_chars:
        return text
    return text[: max_chars - 1].rstrip() + "…"


def build_loan_mod_spotlight_dataframe(deck: pd.DataFrame) -> pd.DataFrame:
    if deck.empty or "loan_mod_count" not in deck.columns:
        return pd.DataFrame()

    source = deck[pd.to_numeric(deck["loan_mod_count"], errors="coerce").fillna(0).astype(int) > 0].copy()
    if source.empty:
        return pd.DataFrame()

    rows = []
    for _, row in source.iterrows():
        rows.append(
            {
                "Type": display_text(row.get("sheet")),
                "Deal Name": display_text(row.get("deal_name")),
                "Deal #": display_text(row.get("deal_number")),
                "Mod Count": fmt_int(row.get("loan_mod_count")),
                "Mod Order #": display_text(row.get("loan_mod_latest_order_number")),
                "Modification Type": display_text(row.get("loan_mod_latest_type")),
                "Status": display_text(row.get("loan_mod_latest_status")),
                "Completed": fmt_date(row.get("loan_mod_latest_completed_date")),
                "Effective": fmt_date(row.get("loan_mod_latest_effective_date")),
                "Previous Maturity": fmt_date(row.get("loan_mod_latest_previous_maturity_date")),
                "Updated Maturity": fmt_date(row.get("loan_mod_latest_updated_maturity_date")),
                "Previous Commitment": fmt_money_exact(row.get("loan_mod_latest_previous_commitment")),
                "Updated Commitment": fmt_money_exact(row.get("loan_mod_latest_updated_commitment")),
                "Comments": compact_note(row.get("loan_mod_latest_comments")),
                "Latest Request": display_text(row.get("loan_mod_latest_request_note"), blank=""),
            }
        )

    return pd.DataFrame(rows)


# -----------------------------------------------------------------------------
# Meeting state: pure cores (no st.* calls, no session_state)
# -----------------------------------------------------------------------------
def available_status_suggestions(deck: pd.DataFrame) -> List[str]:
    if "status" not in deck.columns:
        return COMMON_STATUS_SUGGESTIONS
    existing = sorted({norm_label(v) for v in deck["status"].tolist() if norm_label(v)})
    return sorted(set(COMMON_STATUS_SUGGESTIONS + existing))


# Characters that turn into spaces when commentary is cleaned (soft line breaks become
# real line breaks instead, so they are not counted as removed).
_HIDDEN_IN_FREE_TEXT_RE = re.compile("[\x00-\x08\x0e-\x1f￾￿]")


def _optional_int(value: object) -> Optional[int]:
    number = _finite_number(value)
    return None if number is None else int(number)


def _field_clean(field_name: str, value: object) -> str:
    return clean_free_text(value) if field_name == "commentary" else norm_label(value)


def sanitize_edit_value(field_name: str, value: object, current: str = "") -> Tuple[str, int, bool]:
    """(clean text, hidden characters removed, trimmed) for one edited field.

    Status and owner are single-line labels; commentary keeps its line breaks and is cut
    to COMMENTARY_MAX_CHARS, unless it equals the current value (a long workbook
    commentary saved unchanged is never shortened).
    """
    if is_missing(value):
        return "", 0, False
    raw = value if isinstance(value, str) else _scalar_text(value)
    pattern = _HIDDEN_IN_FREE_TEXT_RE if field_name == "commentary" else _XML_ILLEGAL_RE
    hidden = len(pattern.findall(raw))
    text = _field_clean(field_name, raw)
    trimmed = False
    if field_name == "commentary" and len(text) > COMMENTARY_MAX_CHARS and text != current:
        text = text[:COMMENTARY_MAX_CHARS].rstrip()
        trimmed = True
    return text, hidden, trimmed


def base_values(row: Dict[str, Any]) -> Dict[str, str]:
    """(status, owner, commentary) of a workbook row or an override, cleaned like an edit."""
    return {field_name: _field_clean(field_name, row.get(field_name)) for field_name in EDITABLE_FIELDS}


def effective_values(base: Dict[str, Any], override: Optional[Dict[str, Any]]) -> Dict[str, str]:
    """What the deck shows for a deal: an override is authoritative for all three fields
    (a blank is a deliberate clear); without one, the workbook's values."""
    return base_values(override if override else base)


def build_base_rows(parsed: Optional[Dict[str, Any]]) -> Dict[str, Dict[str, Any]]:
    """deal_key -> the workbook's own values for every parsed row, hidden rows included.

    apply_edit compares edits with these values and records the row's sheet, deal number
    and Excel row in the override; the workbook-switch guard matches edits against them.
    """
    rows: Dict[str, Dict[str, Any]] = {}
    for sheet in PORTFOLIO_SHEETS:
        base = (((parsed or {}).get("sheets") or {}).get(sheet) or {}).get("base")
        if base is None or base.empty:
            continue
        for record in base.to_dict("records"):
            key = str(record.get("deal_key"))
            rows[key] = {
                "sheet": norm_label(record.get("sheet")),
                "deal_number": norm_id(record.get("deal_number")),
                "deal_key": key,
                "deal_name": norm_label(record.get("deal_name")),
                "excel_row": _optional_int(record.get("_excel_row")),
                "hidden": bool(record.get("_hidden")),
                "status": norm_label(record.get("status")),
                "owner": norm_label(record.get("owner")),
                "commentary": clean_free_text(record.get("commentary")),
            }
    return rows


@dataclass
class EditOutcome:
    """What one apply_edit / edit_core call did."""

    deal_key: str
    deal_name: str = ""
    changed: bool = False  # the stored meeting state changed
    fields: Tuple[str, ...] = ()  # fields whose shown value changed
    removed: bool = False  # the override was removed (back to the workbook values)
    hidden_chars: int = 0
    trimmed: bool = False


def override_payload(
    base: Dict[str, Any],
    values: Dict[str, str],
    wb_sha: str,
    saved_at: str,
) -> Dict[str, Any]:
    """An override: the full (status, owner, commentary) triple plus the row it belongs
    to (sheet, deal number, deal key, Excel row) and the workbook it was made on."""
    return {
        "sheet": norm_label(base.get("sheet")),
        "deal_number": norm_id(base.get("deal_number")),
        "deal_key": str(base.get("deal_key")),
        "excel_row": _optional_int(base.get("excel_row")),
        "workbook_sha": wb_sha,
        "status": values["status"],
        "owner": values["owner"],
        "commentary": values["commentary"],
        "saved_at": saved_at,
    }


def _override_identity(item: Optional[Dict[str, Any]]) -> Optional[Tuple[Any, ...]]:
    # Everything that matters about a stored override except when it was saved.
    if not item:
        return None
    values = base_values(item)
    return (
        str(item.get("workbook_sha") or ""),
        _optional_int(item.get("excel_row")),
        *(values[field_name] for field_name in EDITABLE_FIELDS),
    )


def edit_core(
    overrides: Dict[str, Dict[str, Any]],
    base: Dict[str, Any],
    *,
    status: Optional[object] = None,
    owner: Optional[object] = None,
    commentary: Optional[object] = None,
    wb_sha: str = "",
    saved_at: str = "",
) -> Tuple[Dict[str, Dict[str, Any]], EditOutcome]:
    """Pure core of apply_edit. Fields left as None keep their current value.

    The edited triple is compared with the workbook row: equal means the override is
    removed, otherwise the full triple is stored. Returns the overrides (the same object
    when nothing changed, else a new dict) and the outcome.
    """
    deal_key = str(base.get("deal_key"))
    outcome = EditOutcome(
        deal_key=deal_key,
        deal_name=norm_label(base.get("deal_name")) or norm_label(base.get("deal_number")) or deal_key,
    )
    existing = overrides.get(deal_key)
    bound = existing if existing and str(existing.get("workbook_sha") or wb_sha) == wb_sha else None
    current = effective_values(base, bound)
    proposed = dict(current)
    for field_name, value in (("status", status), ("owner", owner), ("commentary", commentary)):
        if value is None:
            continue
        text, hidden, trimmed = sanitize_edit_value(field_name, value, current[field_name])
        proposed[field_name] = text
        outcome.hidden_chars += hidden
        outcome.trimmed = outcome.trimmed or trimmed
    outcome.fields = tuple(field_name for field_name in EDITABLE_FIELDS if proposed[field_name] != current[field_name])
    target = None if proposed == base_values(base) else override_payload(base, proposed, wb_sha, saved_at)
    if _override_identity(existing) == _override_identity(target):
        return overrides, outcome  # a no-op save keeps the stored override (and its saved_at)
    updated = dict(overrides)
    if target is None:
        updated.pop(deal_key, None)
        outcome.removed = existing is not None
    else:
        updated[deal_key] = target
    outcome.changed = True
    return updated, outcome


def rebind_override(item: Dict[str, Any], base: Dict[str, Any], wb_sha: str) -> Dict[str, Any]:
    """The same edit, bound to another workbook's row for the deal."""
    out = dict(item)
    out.update(
        sheet=norm_label(base.get("sheet")),
        deal_number=norm_id(base.get("deal_number")),
        deal_key=str(base.get("deal_key")),
        excel_row=_optional_int(base.get("excel_row")),
        workbook_sha=wb_sha,
    )
    return out


def reconcile_core(
    overrides: Dict[str, Dict[str, Any]],
    base_rows: Dict[str, Dict[str, Any]],
    wb_sha: str,
) -> Tuple[Dict[str, Dict[str, Any]], Dict[str, Dict[str, Any]], int]:
    """Bind meeting edits to the loaded workbook (sha256 of its bytes).

    Edits made on another workbook are never applied silently. Those whose deal holds
    exactly the edited values in this workbook (re-uploading the post-meeting export)
    are re-bound; edits saved without a workbook sha are bound here when the deal
    exists. Returns (overrides, foreign edits awaiting a decision, re-bound count).
    """
    updated: Optional[Dict[str, Dict[str, Any]]] = None
    foreign: Dict[str, Dict[str, Any]] = {}
    rebound = 0
    for key, item in overrides.items():
        sha = str(item.get("workbook_sha") or "")
        if sha == wb_sha:
            continue
        base = base_rows.get(key)
        if base is not None and (not sha or base_values(item) == base_values(base)):
            if updated is None:
                updated = dict(overrides)
            updated[key] = rebind_override(item, base, wb_sha)
            rebound += 1 if sha else 0
            continue
        foreign[key] = item
    return (overrides if updated is None else updated), foreign, rebound


def keep_matching_core(
    overrides: Dict[str, Dict[str, Any]],
    base_rows: Dict[str, Dict[str, Any]],
    wb_sha: str,
) -> Tuple[Dict[str, Dict[str, Any]], int, int]:
    """'Keep matching edits': re-bind other-workbook edits whose deal is in this workbook
    and drop the rest. Returns (overrides, kept, dropped)."""
    updated: Dict[str, Dict[str, Any]] = {}
    kept = dropped = 0
    for key, item in overrides.items():
        if str(item.get("workbook_sha") or "") == wb_sha:
            updated[key] = item
            continue
        base = base_rows.get(key)
        if base is None:
            dropped += 1
            continue
        updated[key] = rebind_override(item, base, wb_sha)
        kept += 1
    return updated, kept, dropped


def discard_foreign_core(overrides: Dict[str, Dict[str, Any]], wb_sha: str) -> Tuple[Dict[str, Dict[str, Any]], int]:
    """'Discard all edits': keep only the edits made on this workbook."""
    kept = {key: item for key, item in overrides.items() if str(item.get("workbook_sha") or "") == wb_sha}
    return kept, len(overrides) - len(kept)


def _position_order(value: object) -> Tuple[int, str]:
    try:
        return int(value), ""
    except (TypeError, ValueError):
        return -1, str(value)


def review_delta_to_changes(
    edited_rows: Dict[Any, Dict[str, Any]],
    row_keys: Sequence[str],
    current_lookup: Dict[str, Dict[str, str]],
) -> List[Tuple[str, Dict[str, str]]]:
    """Committed Review-grid cells -> [(deal_key, {field: text})] to save.

    Positions map through row_keys, the deal keys of the exact frame rendered, so the
    (hidden, read-only) index can never re-target a row. A cleared cell (None, NaN,
    pd.NA) is '' -- a deliberate clear, never the text 'nan'. Read-only columns and
    values equal to the deal's current value are skipped; positions outside the frame
    are ignored and logged.
    """
    changes: List[Tuple[str, Dict[str, str]]] = []
    for position_key, cells in sorted((edited_rows or {}).items(), key=lambda item: _position_order(item[0])):
        try:
            position = int(position_key)
        except (TypeError, ValueError):
            LOG.warning("event=review_edit_ignored reason=bad_position")
            continue
        if not 0 <= position < len(row_keys):
            LOG.warning("event=review_edit_ignored reason=out_of_range position=%d rows=%d", position, len(row_keys))
            continue
        deal_key = str(row_keys[position])
        current = current_lookup.get(deal_key)
        if current is None:
            LOG.warning("event=review_edit_ignored reason=unknown_deal deal_key=%s", deal_key)
            continue
        fields: Dict[str, str] = {}
        for column, value in (cells or {}).items():
            field_name = REVIEW_EDITABLE_COLUMNS.get(str(column))
            if field_name is None:
                continue
            text = "" if is_missing(value) else (value if isinstance(value, str) else _scalar_text(value))
            clean, _hidden, _trimmed = sanitize_edit_value(field_name, text, current.get(field_name, ""))
            if clean == current.get(field_name, ""):
                continue
            fields[field_name] = text
        if fields:
            changes.append((deal_key, fields))
    return changes


def search_mask(frame: pd.DataFrame, query: object) -> pd.Series:
    """Rows whose searchable text contains the query: literal and case-insensitive (never
    a regex, so '(', '*', '+', '$' or '\\' are plain characters)."""
    needle = norm_label(query).casefold()
    if frame.empty or not needle:
        return pd.Series(True, index=frame.index, dtype=bool)
    hits = np.zeros(len(frame), dtype=bool)
    for column in SEARCH_COLUMNS:
        if column not in frame.columns:
            continue
        hits |= np.fromiter(
            (needle in norm_label(value).casefold() for value in frame[column].tolist()),
            dtype=bool,
            count=len(frame),
        )
    return pd.Series(hits, index=frame.index, dtype=bool)


def filter_deck(
    deck: pd.DataFrame,
    *,
    sheet_filter: object = "All",
    query: object = "",
    stale_only: bool = False,
) -> pd.DataFrame:
    """Deck rows matching the filters, in workbook order. A missing or unknown portfolio
    type counts as 'All'; 'Update reminders only' keeps open reminders."""
    if deck.empty:
        return deck.copy()
    filtered = deck
    if sheet_filter in ("Bridge", "Term"):
        filtered = filtered[(filtered["sheet"] == sheet_filter).to_numpy()]
    if norm_label(query):
        filtered = filtered[search_mask(filtered, query).to_numpy()]
    if stale_only and "status_needs_update_prompt" in filtered.columns:
        filtered = filtered[flag_array(filtered["status_needs_update_prompt"])]
    return filtered.sort_values(["sheet_order", "original_order"], kind="stable").reset_index(drop=True)


def resolve_cursor(
    cursor: Optional[str],
    full_keys: Sequence[str],
    agenda_keys: Sequence[str],
    last_agenda: Sequence[str] = (),
) -> Optional[str]:
    """The deal the meeting is on.

    A cursor that still exists in the deck is kept even when the filters hide it
    (sticky: filters and edits never move the meeting). Otherwise the nearest deal after
    it in the agenda last shown that is still on the agenda, else the nearest before it,
    else the first agenda item.
    """
    if cursor is not None and cursor in set(full_keys):
        return cursor
    if not agenda_keys:
        return None
    available = set(agenda_keys)
    last = list(last_agenda or ())
    if cursor is not None and cursor in last:
        position = last.index(cursor)
        for key in last[position + 1:]:
            if key in available:
                return key
        for key in reversed(last[:position]):
            if key in available:
                return key
    return agenda_keys[0]


def displayed_agenda(
    agenda_keys: Sequence[str],
    cursor: Optional[str],
    order: Dict[str, Tuple[int, int]],
) -> Tuple[List[str], bool]:
    """The agenda Presentation steps through: the filtered keys, plus a sticky cursor the
    filters hide, at its workbook position. Returns (keys, cursor is outside the filter)."""
    keys = list(agenda_keys)
    if cursor is None or cursor in set(keys) or cursor not in order:
        return keys, False
    keys.append(cursor)
    keys.sort(key=lambda key: order.get(key, (len(SHEET_ORDER), 0)))
    return keys, True


def reminder_states(
    deal_keys: Sequence[str],
    streak_flags: Sequence[bool],
    overrides: Dict[str, Dict[str, Any]],
    base_rows: Dict[str, Dict[str, Any]],
    confirmed: Dict[str, str],
    meeting_iso: str,
) -> List[str]:
    """How each deal's update reminder was addressed in this meeting: 'Confirmed' (the
    status was confirmed as still accurate), 'Refreshed' (AM Commentary was rewritten),
    or '' (still open, or no reminder). An owner-only edit or a cosmetic re-spacing of
    the same commentary is not a refresh."""
    states: List[str] = []
    for key, flag in zip(deal_keys, streak_flags):
        if not flag:
            states.append("")
            continue
        if (confirmed or {}).get(key) == meeting_iso:
            states.append(REMINDER_CONFIRMED)
            continue
        item = overrides.get(key)
        base = base_rows.get(key)
        if item is not None and base is not None:
            new_commentary = clean_free_text(item.get("commentary"))
            if new_commentary and not text_equal(new_commentary, base.get("commentary")):
                states.append(REMINDER_REFRESHED)
                continue
        states.append("")
    return states


def mark_addressed_reminders(
    deck: pd.DataFrame,
    overrides: Dict[str, Dict[str, Any]],
    base_rows: Dict[str, Dict[str, Any]],
    confirmed: Dict[str, str],
    meeting_iso: str,
) -> pd.DataFrame:
    """Deck with reminders addressed this meeting closed: status_reminder_streak keeps the
    raw streak flag, reminder_addressed says how it was addressed, and
    status_needs_update_prompt (the open reminder every view and filter uses) is the
    streak flag minus the addressed ones."""
    if deck.empty or "status_needs_update_prompt" not in deck.columns:
        return deck
    out = deck.copy()
    streak = flag_array(out["status_needs_update_prompt"])
    states = reminder_states(out["deal_key"].astype(str).tolist(), streak, overrides, base_rows, confirmed, meeting_iso)
    out["status_reminder_streak"] = streak
    out["reminder_addressed"] = states
    out["status_needs_update_prompt"] = streak & np.array([state == "" for state in states], dtype=bool)
    return out


def reminder_labels(frame: pd.DataFrame) -> List[str]:
    """'Update Reminder' column text: 'Yes' while open, 'Confirmed' / 'Refreshed' once
    addressed this meeting, else ''."""
    count = len(frame)
    open_flags = (
        flag_array(frame["status_needs_update_prompt"])
        if "status_needs_update_prompt" in frame.columns
        else np.zeros(count, dtype=bool)
    )
    addressed = (
        [norm_label(value) for value in frame["reminder_addressed"].tolist()]
        if "reminder_addressed" in frame.columns
        else [""] * count
    )
    return ["Yes" if flag else state for flag, state in zip(open_flags, addressed)]


def dialog_to_render(
    active: object,
    *,
    view: str,
    deck_keys: Iterable[str],
    has_foreign: bool,
) -> Tuple[Optional[str], bool]:
    """(dialog kind to render or None, clear active_dialog).

    At most one dialog per run: the workbook-switch guard wins while edits from another
    workbook await a decision; an edit dialog renders only on Presentation and Review and
    only for a deal that is still in the deck. Anything else is cleared.
    """
    if has_foreign:
        return DIALOG_WORKBOOK_SWITCH, False
    if not isinstance(active, dict) or not active:
        return None, active is not None
    kind = active.get("kind")
    if kind == DIALOG_EDIT and view in EDIT_DIALOG_VIEWS and active.get("deal_key") in set(deck_keys):
        return DIALOG_EDIT, False
    return None, True


def _csv_cell(value: object) -> object:
    # Human-facing CSV: text that Excel would run as a formula gets a leading apostrophe.
    if isinstance(value, str) and value[:1] in ("=", "+", "-", "@", "\t", "\r"):
        return "'" + value
    return value


def edits_csv_bytes(items: Sequence[Dict[str, Any]], workbook_names: Dict[str, str]) -> bytes:
    """Meeting edits as a UTF-8 (BOM) CSV, e.g. 'Download my edits first'."""
    columns = ["Deal Key", "Type", "Deal #", "Status", "Owner", "AM Commentary", "Saved At", "Workbook"]
    rows = [
        {
            "Deal Key": str(item.get("deal_key") or get_override_key(item.get("sheet", ""), item.get("deal_number", ""))),
            "Type": norm_label(item.get("sheet")),
            "Deal #": norm_id(item.get("deal_number")),
            "Status": norm_label(item.get("status")),
            "Owner": norm_label(item.get("owner")),
            "AM Commentary": clean_free_text(item.get("commentary")),
            "Saved At": norm_label(item.get("saved_at")),
            "Workbook": norm_label(workbook_names.get(str(item.get("workbook_sha") or ""), "")),
        }
        for item in items
    ]
    frame = pd.DataFrame(rows, columns=columns).astype(object)
    frame = frame.map(_csv_cell)
    return ("﻿" + frame.to_csv(index=False)).encode("utf-8")


def apply_overrides(
    deck: pd.DataFrame,
    overrides: Optional[Dict[str, Dict[str, Any]]],
    wb_sha: Optional[str] = None,
) -> pd.DataFrame:
    """Deck with the meeting's edits applied by deal_key.

    An override is authoritative: its status, owner and commentary win even when blank
    (a deliberate clear), so the screen, the changes CSV, the history snapshot and the
    exported workbook agree. With wb_sha, edits bound to another workbook are left out.
    """
    if deck.empty:
        return deck.copy()
    out = deck.copy()
    if "saved_at" not in out.columns:
        out["saved_at"] = ""
    active = {
        str(key): item
        for key, item in (overrides or {}).items()
        if wb_sha is None or str(item.get("workbook_sha") or wb_sha) == wb_sha
    }
    if not active:
        return out
    keys = out["deal_key"].astype(str).tolist()
    hits = [key in active for key in keys]
    if not any(hits):
        return out
    for field_name in (*EDITABLE_FIELDS, "saved_at"):
        values = out[field_name].astype(object).tolist()
        for position, (key, hit) in enumerate(zip(keys, hits)):
            if hit:
                raw = active[key].get(field_name)
                values[position] = _field_clean(field_name, raw) if field_name in EDITABLE_FIELDS else norm_label(raw)
        out[field_name] = pd.Series(values, index=out.index, dtype=object)
    return out


# -----------------------------------------------------------------------------
# State model: seeded session keys, toasts, dialogs and the single edit path
# -----------------------------------------------------------------------------
# Every session key the app owns, seeded by init_state() (only missing keys; mutable
# defaults are deep-copied and callables are factories). Widgets bound to these keys
# never pass default=/value=/index=, and their keys are written only in callbacks or
# before the widget renders, so Streamlit never warns about a default plus a Session
# State value. Key names never shadow SessionStateProxy methods (items, keys, get, ...).
SESSION_DEFAULTS: Dict[str, Any] = {
    # Source: the loaded workbook {name, bytes, sha256, source, path, mtime_ns, size}.
    "wb": None,
    # Settings
    "as_of_date": dt.date.today,
    "include_hidden": False,
    # Filters (sidebar)
    "sheet_filter": "All",
    "search_query": "",
    "stale_only": False,
    # Meeting
    "meeting_overrides": {},
    "cursor": None,
    "reminder_confirmed": {},
    # Status history
    "status_history": empty_status_history,
    "status_history_sha": "",
    "status_history_default_checked": False,
    "uploaded_status_history_name": "",
    "uploaded_status_history_signature": "",
    "status_history_load_error": "",
    "embedded_status_history_signature": "",
    # UI
    "view_mode": "Presentation",
    "active_dialog": None,
    "_toasts": [],
    "jump_to": None,
    "review_editor_version": 0,
    "state_version": 0,
    "_last_agenda": [],
    # Workbook identity and per-workbook lookups
    "_wb_sha": "",
    "_wb_names": {},
    "_parsed_sha": "",
    "_base_rows": {},
    "_base_rows_sha": "",
    # The as-of date follows each new workbook's reference date unless the user picked a
    # date for that workbook.
    "_as_of_default_sha": "",
    "_as_of_manual_sha": "",
}


def init_state() -> None:
    """Seed every missing SESSION_DEFAULTS key (existing values are never touched)."""
    for key, default in SESSION_DEFAULTS.items():
        if key not in st.session_state:
            st.session_state[key] = default() if callable(default) else copy.deepcopy(default)


def _bump(key: str) -> None:
    st.session_state[key] = int(st.session_state.get(key) or 0) + 1


def now_text() -> str:
    return shell.viewer_now().strftime("%Y-%m-%d %H:%M:%S")


def current_meeting_date() -> dt.date:
    """The meeting date: P1 keeps a single date, the sidebar as-of date."""
    value = st.session_state.get("meeting_date") or st.session_state.get("as_of_date")
    return value if isinstance(value, dt.date) else dt.date.today()


def notify(body: str, icon: str = ":material/check_circle:", duration: str = "short") -> None:
    """Queue a toast. render_toasts() shows it at the top of the next run (or at the end
    of this one), so a confirmation survives st.rerun()."""
    queue = list(st.session_state.get("_toasts") or [])
    queue.append({"body": str(body), "icon": icon, "duration": duration})
    st.session_state["_toasts"] = queue[-MAX_QUEUED_TOASTS:]


def render_toasts() -> None:
    """Show and clear the queued toasts."""
    queue = list(st.session_state.get("_toasts") or [])
    if not queue:
        return
    st.session_state["_toasts"] = []
    for item in queue:
        st.toast(item.get("body", ""), icon=item.get("icon") or None, duration=item.get("duration") or "short")


def open_dialog_cb(kind: str, deal_key: Optional[str] = None) -> None:
    """Open a dialog for a deal captured at click time (only this and close_dialog_cb
    write active_dialog)."""
    st.session_state.active_dialog = {"kind": kind, "deal_key": deal_key}


def close_dialog_cb() -> None:
    """Close any dialog: on_dismiss (Esc, X, click-outside), Cancel, Save, navigation."""
    st.session_state.active_dialog = None


def current_wb_sha() -> str:
    return str((st.session_state.get("wb") or {}).get("sha256") or "")


def native_overrides(wb_sha: Optional[str] = None) -> Dict[str, Dict[str, Any]]:
    """Meeting edits bound to the loaded workbook. Edits made on another workbook wait for
    the workbook-switch decision and are never applied."""
    sha = current_wb_sha() if wb_sha is None else wb_sha
    overrides = st.session_state.get("meeting_overrides") or {}
    return {key: item for key, item in overrides.items() if str(item.get("workbook_sha") or "") == sha}


def apply_edit(
    deal_key: str,
    *,
    status: Optional[object] = None,
    owner: Optional[object] = None,
    commentary: Optional[object] = None,
    source: str = "dialog",
    announce: bool = True,
) -> Optional[EditOutcome]:
    """The single write path for Status / owner / AM Commentary (edit dialog, Review grid).

    Sanitizes the text (labels on one line; commentary keeps its line breaks, hidden
    control characters become spaces, at most 4,000 characters), stores the full triple
    bound to this workbook -- or removes the override when it equals the workbook --
    copy-on-write, bumps state_version only on an effective change and queues toasts.
    """
    base = (st.session_state.get("_base_rows") or {}).get(deal_key)
    if base is None:
        LOG.warning("event=edit_rejected deal_key=%s source=%s reason=unknown_deal", deal_key, source)
        notify("That deal is not in the loaded workbook, so the edit was not saved.", icon=":material/error:")
        return None
    overrides = st.session_state.get("meeting_overrides") or {}
    updated, outcome = edit_core(
        overrides,
        base,
        status=status,
        owner=owner,
        commentary=commentary,
        wb_sha=current_wb_sha(),
        saved_at=now_text(),
    )
    name = md_escape(outcome.deal_name)
    if not outcome.changed:
        if announce:
            notify(f"No changes · {name}", icon=":material/info:")
        return outcome
    st.session_state.meeting_overrides = updated
    _bump("state_version")
    if source != "review":
        # A new grid key, so cell edits pending in the Review grid never re-apply on top.
        _bump("review_editor_version")
    LOG.info(
        "event=edit deal_key=%s fields=%s source=%s removed=%s",
        deal_key,
        ",".join(outcome.fields) or "-",
        source,
        outcome.removed,
    )
    if announce:
        notify(f"Saved · {name}")
    if outcome.hidden_chars:
        notify(f"Removed {_plural(outcome.hidden_chars, 'hidden character')}", icon=":material/cleaning_services:")
    if outcome.trimmed:
        notify(f"Trimmed to {COMMENTARY_MAX_CHARS:,} characters", icon=":material/content_cut:")
    return outcome


def effective_lookup() -> Dict[str, Dict[str, str]]:
    """deal_key -> the values the deck shows now (workbook values or this workbook's edit)."""
    base_rows = st.session_state.get("_base_rows") or {}
    overrides = native_overrides()
    return {key: effective_values(base, overrides.get(key)) for key, base in base_rows.items()}


def review_saved_message(outcomes: Sequence[EditOutcome]) -> str:
    if len(outcomes) == 1:
        labels = ", ".join(FIELD_LABELS[field_name] for field_name in outcomes[0].fields)
        return f"Saved · {md_escape(outcomes[0].deal_name)}" + (f" · {labels}" if labels else "")
    changes = sum(max(len(outcome.fields), 1) for outcome in outcomes)
    return f"Saved {changes:,} changes"


def commit_review_cb(editor_key: str, row_keys: Sequence[str]) -> None:
    """Review grid on_change: every committed cell is saved through apply_edit at once.

    row_keys are the deal keys of the frame the user edited (captured at render time),
    so a filter or view change in the same rerun cannot re-target the edits.
    """
    state = st.session_state.get(editor_key)
    edited = state.get("edited_rows") if hasattr(state, "get") else None
    changes = review_delta_to_changes(edited or {}, list(row_keys), effective_lookup())
    saved: List[EditOutcome] = []
    for deal_key, fields in changes:
        outcome = apply_edit(deal_key, source="review", announce=False, **fields)
        if outcome is not None and outcome.changed:
            saved.append(outcome)
    if saved:
        notify(review_saved_message(saved))


def move_cursor_cb(agenda: Sequence[str], step: int) -> None:
    """Previous / Next through the agenda the presenter was looking at."""
    close_dialog_cb()
    keys = list(agenda)
    if not keys:
        return
    cursor = st.session_state.get("cursor")
    position = keys.index(cursor) if cursor in keys else 0
    st.session_state.cursor = keys[min(max(position + step, 0), len(keys) - 1)]


def jump_cb() -> None:
    """'Jump to deal': move the meeting to the chosen deal and clear the box (one run)."""
    close_dialog_cb()
    target = st.session_state.get("jump_to")
    if target:
        st.session_state.cursor = (st.session_state.get("_jump_key_map") or {}).get(target, target)
    st.session_state.jump_to = None


def reset_filters_cb() -> None:
    """'Reset filters' / 'Clear filters'. A callback runs before the widgets exist, so
    their keys can be written (assigning them after rendering crashed the app)."""
    st.session_state.sheet_filter = "All"
    st.session_state.search_query = ""
    st.session_state.stale_only = False
    st.session_state.owner_filter = []
    st.session_state.focus_filters = []
    close_dialog_cb()


def confirm_reminder_cb(deal_key: str, meeting_iso: str) -> None:
    """'Confirm still accurate': the update reminder is addressed for this meeting."""
    confirmed = dict(st.session_state.get("reminder_confirmed") or {})
    confirmed[deal_key] = meeting_iso
    st.session_state.reminder_confirmed = confirmed
    _bump("state_version")
    LOG.info("event=reminder_confirmed deal_key=%s", deal_key)
    notify("Confirmed: the status is still accurate for this meeting", icon=":material/verified:")


def unconfirm_reminder_cb(deal_key: str) -> None:
    confirmed = dict(st.session_state.get("reminder_confirmed") or {})
    if confirmed.pop(deal_key, None) is None:
        return
    st.session_state.reminder_confirmed = confirmed
    _bump("state_version")
    LOG.info("event=reminder_unconfirmed deal_key=%s", deal_key)


def active_filter_parts() -> List[str]:
    parts: List[str] = []
    sheet = st.session_state.get("sheet_filter")
    if sheet in ("Bridge", "Term"):
        parts.append(f"{sheet} only")
    query = norm_label(st.session_state.get("search_query"))
    if query:
        parts.append(f"search “{query}”")
    if st.session_state.get("stale_only"):
        parts.append("update reminders only")
    return parts


def apply_filters(deck: pd.DataFrame) -> pd.DataFrame:
    return filter_deck(
        deck,
        sheet_filter=st.session_state.get("sheet_filter"),
        query=st.session_state.get("search_query") or "",
        stale_only=bool(st.session_state.get("stale_only")),
    )


# Workbook source: st.session_state['wb'] holds the loaded file; change detection is by
# the sha256 of its bytes, never by name or size.
def workbook_record(name: str, data: bytes, source: str, path: Optional[Path] = None) -> Dict[str, Any]:
    record: Dict[str, Any] = {
        "name": str(name),
        "bytes": data,
        "sha256": bytes_sha256(data),
        "source": source,
        "path": "",
        "mtime_ns": 0,
        "size": len(data or b""),
    }
    if path is not None:
        try:
            stat = path.stat()
            record.update(path=str(path), mtime_ns=int(stat.st_mtime_ns), size=int(stat.st_size))
        except OSError:
            pass
    return record


def default_workbook_path() -> Optional[Path]:
    env_path = os.environ.get(DEFAULT_WORKBOOK_ENV, "").strip()
    if env_path:
        path = Path(env_path).expanduser()
        if path.is_file():
            return path
    return first_existing_file(DEFAULT_SAMPLE_FILES)


def set_workbook(record: Dict[str, Any]) -> None:
    st.session_state.wb = record
    names = dict(st.session_state.get("_wb_names") or {})
    names[record["sha256"]] = record["name"]
    st.session_state["_wb_names"] = names


def on_workbook_upload_cb() -> None:
    """Uploader on_change: load the new file. Clearing the uploader keeps the loaded one."""
    uploaded = st.session_state.get("uploaded_workbook")
    if uploaded is None:
        return
    set_workbook(workbook_record(uploaded.name, uploaded.getvalue(), "upload"))
    close_dialog_cb()


def current_workbook() -> Optional[Dict[str, Any]]:
    """The workbook for this run: the session's (an upload or a local default file, which
    is re-read when it changes on disk); on the first run the default file, if any."""
    wb = st.session_state.get("wb")
    if wb and wb.get("source") == "default" and wb.get("path"):
        path = Path(wb["path"])
        try:
            stat = path.stat()
            if int(stat.st_mtime_ns) != int(wb.get("mtime_ns") or 0) or int(stat.st_size) != int(wb.get("size") or 0):
                set_workbook(workbook_record(path.name, path.read_bytes(), "default", path))
        except OSError:
            pass
        return st.session_state.get("wb")
    if wb:
        return wb
    path = default_workbook_path()
    if path is None:
        return None
    try:
        record = workbook_record(path.name, path.read_bytes(), "default", path)
    except OSError:
        return None
    set_workbook(record)
    return record


def refresh_base_rows(parsed: Dict[str, Any], wb_sha: str) -> Dict[str, Dict[str, Any]]:
    """The workbook's own row values, rebuilt once per workbook sha."""
    if st.session_state.get("_base_rows_sha") != wb_sha:
        st.session_state["_base_rows"] = build_base_rows(parsed)
        st.session_state["_base_rows_sha"] = wb_sha
    return st.session_state.get("_base_rows") or {}


def reconcile_overrides(base_rows: Dict[str, Dict[str, Any]], wb_sha: str) -> Dict[str, Dict[str, Any]]:
    """Workbook-switch guard. Edits made on another workbook are not applied; while any
    await a decision the 'Different workbook loaded' dialog stays open. Returns them."""
    overrides = st.session_state.get("meeting_overrides") or {}
    if not overrides:
        return {}
    updated, foreign, rebound = reconcile_core(overrides, base_rows, wb_sha)
    if updated is not overrides:
        st.session_state.meeting_overrides = updated
        _bump("state_version")
        _bump("review_editor_version")
        if rebound:
            notify(
                f"{_plural(rebound, 'edit')} carried over: this workbook already has the edited values",
                icon=":material/sync:",
            )
    if foreign:
        active = st.session_state.get("active_dialog") or {}
        if active.get("kind") != DIALOG_WORKBOOK_SWITCH:
            open_dialog_cb(DIALOG_WORKBOOK_SWITCH)
            LOG.info("event=wb_switch_guard sha=%s foreign=%d", wb_sha[:8], len(foreign))
    return foreign


def keep_matching_edits_cb() -> None:
    """Guard: re-bind the edits whose deals are in this workbook; drop the others."""
    wb_sha = current_wb_sha()
    updated, kept, dropped = keep_matching_core(
        st.session_state.get("meeting_overrides") or {},
        st.session_state.get("_base_rows") or {},
        wb_sha,
    )
    st.session_state.meeting_overrides = updated
    _bump("state_version")
    _bump("review_editor_version")
    close_dialog_cb()
    LOG.info("event=wb_switch_decision decision=keep kept=%d dropped=%d", kept, dropped)
    message = f"Kept {_plural(kept, 'edit')}"
    if dropped:
        message += f"; dropped {dropped:,} for deals not in this workbook"
    notify(message)
    st.rerun()  # from inside the dialog: escalate to a full-app rerun


def discard_foreign_edits_cb() -> None:
    """Guard: drop every edit made on another workbook."""
    updated, dropped = discard_foreign_core(st.session_state.get("meeting_overrides") or {}, current_wb_sha())
    st.session_state.meeting_overrides = updated
    _bump("state_version")
    _bump("review_editor_version")
    close_dialog_cb()
    LOG.info("event=wb_switch_decision decision=discard dropped=%d", dropped)
    notify(f"Discarded {_plural(dropped, 'edit')}", icon=":material/delete:")
    st.rerun()  # from inside the dialog: escalate to a full-app rerun


@dataclass
class AppContext:
    """Everything one run derives, handed to the dialogs and views."""

    view: str
    wb: Dict[str, Any]
    parsed: Dict[str, Any]
    deck_info: Optional[Dict[str, Any]]
    raw_deck: pd.DataFrame
    deck: pd.DataFrame  # full deck: edits applied, streaks and reminders added
    filtered: pd.DataFrame  # the deck under the sidebar filters
    loan_mods: pd.DataFrame
    as_of: dt.date
    base_rows: Dict[str, Dict[str, Any]] = field(default_factory=dict)
    foreign_overrides: Dict[str, Dict[str, Any]] = field(default_factory=dict)
    agenda_keys: List[str] = field(default_factory=list)  # what Presentation steps through
    cursor: Optional[str] = None
    cursor_outside: bool = False
    deck_keys: frozenset = frozenset()

    @property
    def workbook_name(self) -> str:
        return str(self.wb.get("name") or "")

    @property
    def wb_sha(self) -> str:
        return str(self.wb.get("sha256") or "")

    def row(self, deal_key: Optional[str]) -> Optional[pd.Series]:
        if deal_key is None or deal_key not in self.deck_keys:
            return None
        match = self.deck[self.deck["deal_key"] == deal_key]
        return None if match.empty else match.iloc[0]


def build_presenter_prompts(row: pd.Series) -> List[str]:
    prompts: List[str] = []

    if bool(row.get("status_needs_update_prompt", False)):
        prompts.append(str(row.get("status_prompt_message") or "Review whether the meeting update should be refreshed."))

    if has_loan_mods(row):
        summary = display_text(row.get("loan_mod_latest_summary"), blank="latest modification on file")
        request_note = display_text(row.get("loan_mod_latest_request_note"), blank="")
        if int(row.get("loan_mod_active_count") or 0) > 0:
            prompt = f"Review loan modification details: {summary}."
            if request_note:
                prompt += f" Newest request ({request_note}) did not go ahead."
        else:
            prompt = f"No loan mod in force; the request on file did not go ahead ({request_note or summary})."
        prompts.append(prompt)

    maturity_text = fmt_date(row.get("maturity_date"))
    payment_text = fmt_date(row.get("next_payment_date"))

    if maturity_text != "-":
        prompts.append(f"Confirm maturity timing: {maturity_text} ({fmt_day_delta(row.get('days_to_maturity'))}).")

    if payment_text != "-":
        prompts.append(f"Confirm next payment timing: {payment_text} ({fmt_day_delta(row.get('days_to_next_payment'))}).")

    commentary_text = display_text(row.get("commentary"), blank="")
    if commentary_text:
        prompts.append("Decide whether AM commentary needs a live refresh before the meeting ends.")

    if not prompts:
        prompts.append("Confirm the latest story, next milestone, and any follow-up owner.")

    return prompts[:4]


# -----------------------------------------------------------------------------
# Dataframes for tables and exports
# -----------------------------------------------------------------------------
def build_agenda_dataframe(deck: pd.DataFrame) -> pd.DataFrame:
    queue = deck.copy()
    queue.insert(0, "Agenda #", range(1, len(queue) + 1))
    queue["UPB"] = queue["upb"].map(lambda x: fmt_money(x, decimals=0))
    queue["Maturity"] = queue["maturity_date"].map(fmt_date)
    queue["Next Payment"] = queue["next_payment_date"].map(fmt_date)
    queue["Status"] = queue["status"].map(display_text)
    queue["Owner"] = queue["owner"].map(display_text)
    queue["Update Reminder"] = reminder_labels(queue)

    if "loan_mod_count" not in queue.columns:
        queue["loan_mod_count"] = 0
    if "loan_mod_latest_summary" not in queue.columns:
        queue["loan_mod_latest_summary"] = ""

    queue["Loan Mods"] = queue["loan_mod_count"].map(lambda value: fmt_int(value) if int(value or 0) else "")
    queue["Latest Loan Mod"] = queue["loan_mod_latest_summary"].map(lambda value: display_text(value, blank=""))

    return queue[
        [
            "Agenda #",
            "sheet",
            "deal_name",
            "deal_number",
            "Update Reminder",
            "Loan Mods",
            "Latest Loan Mod",
            "Status",
            "Owner",
            "Maturity",
            "Next Payment",
            "UPB",
        ]
    ].rename(
        columns={
            "sheet": "Type",
            "deal_name": "Deal Name",
            "deal_number": "Deal #",
        }
    )


def build_review_dataframe(deck: pd.DataFrame) -> pd.DataFrame:
    """The Review grid's frame. The index is the deal key (hidden and read-only in the
    grid): with a keyed editor, a different row set therefore means a fresh grid, so
    pending cell edits can never land on another deal after a filter change."""
    review = deck.copy().set_index("deal_key")
    review.index = review.index.astype(str)
    review.insert(0, "Agenda #", range(1, len(review) + 1))
    review["Type"] = review["sheet"]
    review["Deal Name"] = review["deal_name"].map(display_text)
    review["Deal #"] = review["deal_number"].map(display_text)
    review["Maturity"] = review["maturity_date"].map(fmt_date)
    review["Next Payment"] = review["next_payment_date"].map(fmt_date)
    review["UPB"] = review["upb"].map(lambda x: fmt_money(x, decimals=0))
    review["Update Reminder"] = reminder_labels(review)
    # Plain Python strings (object dtype): a cleared cell comes back as None, never 'nan'.
    review["Status"] = pd.Series([norm_label(v) for v in review["status"].tolist()], index=review.index, dtype=object)
    review["Owner"] = pd.Series([norm_label(v) for v in review["owner"].tolist()], index=review.index, dtype=object)
    review["AM Commentary"] = pd.Series(
        [clean_free_text(v) for v in review["commentary"].tolist()], index=review.index, dtype=object
    )
    return review[list(REVIEW_COLUMN_ORDER)]


def build_updates_dataframe(raw_deck: pd.DataFrame, deck: pd.DataFrame) -> pd.DataFrame:
    """The changed-statuses CSV: one row per deal edited on the loaded workbook."""
    overrides = native_overrides()
    if not overrides or deck.empty:
        return pd.DataFrame()

    changed_keys = {
        str(item.get("deal_key") or get_override_key(item.get("sheet", ""), item.get("deal_number", "")))
        for item in overrides.values()
    }

    out = deck[deck["deal_key"].astype(str).isin(changed_keys)].copy()
    if out.empty:
        return pd.DataFrame()
    out["Update Reminder"] = reminder_labels(out)
    out = out.rename(
        columns={
            "sheet": "Type",
            "deal_name": "Deal Name",
            "deal_number": "Deal #",
            "status": "Status",
            "owner": "Owner",
            "commentary": "AM Commentary",
            "saved_at": "Last Saved",
        }
    )

    return out[
        [
            "Type",
            "Deal Name",
            "Deal #",
            "Update Reminder",
            "Status",
            "Owner",
            "AM Commentary",
            "Last Saved",
        ]
    ]


def export_overrides_csv(raw_deck: pd.DataFrame, deck: pd.DataFrame) -> bytes:
    export_df = build_updates_dataframe(raw_deck, deck)
    if export_df.empty:
        return b""
    return cockpit.safe_csv(export_df)


def update_workbook_bytes(file_bytes: bytes, overrides: Dict[str, Dict[str, Any]]) -> bytes:
    """Write edits to the exact parsed row without rewriting unrelated workbook cells."""
    if not overrides:
        return file_bytes
    from workbook_io import patch_text_cells
    parsed = parse_workbook(file_bytes, bytes_sha256(file_bytes))
    if parsed.get("error"):
        raise ValueError("Workbook could not be parsed for export")
    writes = {}
    for sheet in PORTFOLIO_SHEETS:
        report = parsed.get("sheets", {}).get(sheet) or {}
        base = report.get("base")
        layout = report.get("layout")
        if base is None or base.empty or not layout:
            continue
        by_key = {str(row["deal_key"]): row for row in base.to_dict("records")}
        by_row = {int(row["_excel_row"]): row for row in by_key.values()}
        for item in overrides.values():
            if norm_label(item.get("sheet")) != sheet:
                continue
            row_no = _optional_int(item.get("excel_row"))
            key = str(item.get("deal_key") or get_override_key(sheet, item.get("deal_number")))
            record = by_row.get(row_no) if row_no is not None else by_key.get(key)
            if record is None or norm_id(record.get("deal_number")) != norm_id(item.get("deal_number")):
                raise ValueError("An edited deal no longer matches its workbook row")
            if str(record["deal_key"]) != key:
                raise ValueError("An edited deal has a different duplicate-row identity")
            target_row = int(record["_excel_row"])
            for field_name, role in (("status", "status"), ("owner", "owner"), ("commentary", "commentary")):
                column = layout["role_cols"].get(role)
                value = _field_clean(field_name, item.get(field_name))
                if column is None:
                    if value != _field_clean(field_name, record.get(field_name)):
                        raise ValueError("An edited field has no corresponding workbook column")
                    continue
                writes.setdefault(report["tab"], {})[f"{get_column_letter(column)}{target_row}"] = value
    return patch_text_cells(file_bytes, writes)

def process_status_history_upload() -> None:
    uploaded_history = st.session_state.get("uploaded_status_history")
    if uploaded_history is None:
        return

    # Change detection by content (sha256): a corrected file with the same name and
    # size is still read.
    data = uploaded_history.getvalue()
    signature = bytes_sha256(data)
    if signature == st.session_state.get("uploaded_status_history_signature", ""):
        return

    try:
        history = read_status_history_csv(io.BytesIO(data))
        store_status_history(history, uploaded_history.name)
        st.session_state.uploaded_status_history_signature = signature
        LOG.info("event=history_upload sha=%s rows=%d", signature[:8], len(history))
    except Exception as exc:
        st.session_state.status_history_load_error = str(exc)
        LOG.warning("event=history_upload_failed sha=%s type=%s", signature[:8], type(exc).__name__)


# -----------------------------------------------------------------------------
# Standard Streamlit controls and overview
# -----------------------------------------------------------------------------
def workbook_reference_date(parsed: Optional[Dict[str, Any]]) -> Optional[dt.date]:
    if not parsed or not parsed.get("reference_date"):
        return None
    return dt.date.fromisoformat(parsed["reference_date"])


def sync_as_of_with_workbook(parsed: Optional[Dict[str, Any]], wb_sha: str) -> None:
    """On a new workbook sha, default the as-of date to the workbook's reference date.

    Runs before the date widget renders. A date the user picked for this workbook
    (mark_as_of_manual_cb) is left alone.
    """
    if not wb_sha or st.session_state.get("_as_of_default_sha") == wb_sha:
        return
    st.session_state["_as_of_default_sha"] = wb_sha
    if st.session_state.get("_as_of_manual_sha") == wb_sha:
        return
    reference = workbook_reference_date(parsed)
    if reference is not None:
        st.session_state.as_of_date = reference


def mark_as_of_manual_cb() -> None:
    st.session_state["_as_of_manual_sha"] = st.session_state.get("_wb_sha", "")


def set_as_of_cb(value_iso: str) -> None:
    st.session_state.as_of_date = dt.date.fromisoformat(value_iso)
    st.session_state["_as_of_manual_sha"] = ""


def as_of_target_date(parsed: Optional[Dict[str, Any]], deck_info: Optional[Dict[str, Any]]) -> Optional[dt.date]:
    """The date that makes every sheet use its newest dated column."""
    reference = workbook_reference_date(parsed)
    if reference is not None:
        return reference
    newest: List[dt.date] = []
    for info in (deck_info or {}).get("sheets", {}).values():
        for columns in info.get("dated_columns", {}).values():
            if columns:
                newest.append(dt.date.fromisoformat(columns[-1][0]))
    return max(newest) if newest else None


def render_as_of_notes(parsed: Optional[Dict[str, Any]], deck_info: Optional[Dict[str, Any]]) -> None:
    if not parsed or parsed.get("error"):
        return
    reference = workbook_reference_date(parsed)
    source = parsed.get("reference_source")
    if reference is not None and source == "title":
        st.caption(f"From workbook title row: {reference:%m/%d/%Y}")
    elif reference is not None:
        st.caption(f"From the newest dated column: {reference:%m/%d/%Y}")
    else:
        st.caption("The workbook has no as-of date; using the date above.")

    older = []
    for sheet, info in (deck_info or {}).get("sheets", {}).items():
        if not info.get("uses_older_column"):
            continue
        for kind in ("upb", "npl"):
            used = info.get(f"{kind}_header")
            newest = info.get(f"newest_{kind}_header")
            if used and newest and used != newest:
                older.append(f"{sheet} shows '{used}' (newest '{newest}')")
    target = as_of_target_date(parsed, deck_info)
    if older:
        st.warning("; ".join(older) + ".", title="As-of differs from the workbook")
        if target is not None:
            st.button(
                f"Use {target:%m/%d/%Y}",
                key="use_workbook_as_of",
                on_click=set_as_of_cb,
                args=(target.isoformat(),),
                width="stretch",
            )


def loan_size_notes(sheet: str, report: Dict[str, Any]) -> List[Tuple[str, str]]:
    """Why a sheet's 'Loan Commitment' tiles show '-' (no column, or blank cells)."""
    info = report.get("loan_size") or {}
    rows = int(info.get("rows") or 0)
    missing = int(info.get("missing_count") or 0)
    if not rows or not missing:
        return []
    formula_rows = {
        row
        for item in report.get("formula_uncached_columns") or []
        if item["header"] in (info.get("headers") or [])
        for row in item["rows"]
    }
    if info.get("headers") and set(info.get("missing_rows") or []) <= formula_rows:
        return []  # the formula warning already explains these blanks
    if not info.get("headers"):
        looked_for = list(COMMITMENT_HEADER_CANDIDATES[:2])
        if sheet != "Bridge":
            looked_for += list(LOAN_AMOUNT_HEADER_CANDIDATES[:2])
        return [
            (
                "warning",
                f"{sheet}: no loan size column found, so 'Loan Commitment' shows '-' for its "
                f"{_plural(rows, 'deal')}. Looked for headers such as {_quoted_list(looked_for)}.",
            )
        ]
    missing_rows = info.get("missing_rows") or []
    return [
        (
            "caption",
            f"{sheet}: no loan size for {missing:,} of {_plural(rows, 'deal')} "
            f"({'row' if len(missing_rows) == 1 else 'rows'} {_rows_text(missing_rows)}); "
            f"{_quoted_list(info['headers'])} {'is' if len(info['headers']) == 1 else 'are'} blank there.",
        )
    ]


def workbook_notes(parsed: Optional[Dict[str, Any]]) -> List[Tuple[str, str]]:
    """(level, message) data notes from the parse: 'warning' or 'caption'."""
    notes: List[Tuple[str, str]] = []
    if not parsed or parsed.get("error"):
        return notes
    skipped_labels: List[str] = []
    skipped_total = 0
    duplicate_parts: List[str] = []
    rejected_total = 0
    rejected_example = ""
    formula_parts: List[str] = []
    tab_names = [name for name in parsed.get("sheetnames") or [] if not norm_label(name).startswith("_")]
    for sheet in PORTFOLIO_SHEETS:
        report = (parsed.get("sheets") or {}).get(sheet) or {}
        status = report.get("status")
        if status in {"no_header", "ambiguous", "error", "blank"} and report.get("note"):
            notes.append(("warning", report["note"]))
        elif status == "missing":
            notes.append(
                (
                    "warning",
                    f"No '{sheet}' tab in this workbook, so there are no {sheet} deals in the deck. "
                    f"Tabs found: {_quoted_list(tab_names) or 'none'}.",
                )
            )
        if report.get("tab_match") == "normalized" and report.get("tab"):
            notes.append(("caption", f"Loaded tab '{report['tab']}' as {sheet}."))
        for item in report.get("skipped_rows") or []:
            label = item["value"] or ("Total" if item["reason"] == "total" else "error")
            if label not in skipped_labels:
                skipped_labels.append(label)
        skipped_total += int(report.get("skipped_count") or 0)
        for item in report.get("duplicate_deals") or []:
            duplicate_parts.append(f"{sheet} {item['deal_number']} (rows {_rows_text(item['excel_rows'])})")
        layout = report.get("layout") or {}
        for item in layout.get("role_duplicates") or []:
            used = (
                "using the first of them with a value"
                if item.get("coalesced")
                else f"using column {item['used']}"
            )
            notes.append(
                ("warning", f"{sheet}: '{item['header']}' appears in columns {', '.join(item['columns'])}; {used}.")
            )
        uncached = report.get("formula_deal_uncached") or []
        if uncached:
            one = len(uncached) == 1
            notes.append(
                (
                    "warning",
                    f"{sheet}: {_plural(len(uncached), 'Deal Number cell')} "
                    f"{'is a formula' if one else 'are formulas'} with no saved value "
                    f"({'row' if one else 'rows'} {_rows_text(uncached)}), so "
                    f"{'that row is' if one else 'those rows are'} not in the deck. Open and save the "
                    "workbook in Excel, then upload it again.",
                )
            )
        formula_columns = report.get("formula_uncached_columns") or []
        if formula_columns:
            cells = sum(int(item["count"]) for item in formula_columns)
            formula_parts.append(
                f"{sheet} {_quoted_list([item['header'] for item in formula_columns])} ({_plural(cells, 'cell')})"
            )
        notes.extend(loan_size_notes(sheet, report))
        if report.get("blank_stop_row") and report.get("values_below_stop"):
            notes.append(
                (
                    "warning",
                    f"'{report.get('tab')}' has values below row {report['blank_stop_row']}, after "
                    f"{BLANK_ROW_STOP} empty rows; those rows were not read.",
                )
            )
        rejected_total += int(report.get("rejected_date_count") or 0)
        if not rejected_example and report.get("rejected_dates"):
            first = report["rejected_dates"][0]
            rejected_example = f"{sheet} '{first['header']}' row {first['excel_row']}"
    mod_info = parsed.get("loan_mod_info") or {}
    rejected_total += int(mod_info.get("rejected_date_count") or 0)
    if not rejected_example and mod_info.get("rejected_dates"):
        first = mod_info["rejected_dates"][0]
        rejected_example = f"Loan Modifications '{first['header']}' row {first['excel_row']}"
    if formula_parts:
        notes.append(
            (
                "warning",
                "Formulas with no saved results: " + "; ".join(formula_parts) + ". Those fields show as blank "
                "('-') in the deck. Open the workbook in Excel and save it, then upload it again.",
            )
        )
    if skipped_total:
        notes.append(
            ("caption", f"{_plural(skipped_total, 'non-deal row')} skipped ({', '.join(skipped_labels)})")
        )
    if duplicate_parts:
        notes.append(
            (
                "warning",
                "Duplicate deal numbers: "
                + "; ".join(duplicate_parts)
                + ". The first visible row keeps the deal's history; other copies are separate agenda items.",
            )
        )
    if rejected_total:
        notes.append(
            (
                "caption",
                f"{_plural(rejected_total, 'unreadable or out-of-range date')} ignored"
                + (f" (e.g. {rejected_example})." if rejected_example else "."),
            )
        )
    if mod_info.get("status") == "no_header":
        notes.append(
            (
                "warning",
                f"'{mod_info.get('tab')}' has no 'Deal Number' or 'Opportunity Deal Loan Number' header in its "
                f"first {HEADER_SEARCH_ROWS} rows, so loan mods are not shown.",
            )
        )
    elif mod_info.get("status") == "error" and mod_info.get("note"):
        notes.append(("warning", mod_info["note"]))
    elif mod_info.get("tab_match") == "normalized" and mod_info.get("tab"):
        notes.append(("caption", f"Loaded tab '{mod_info['tab']}' as {LOAN_MOD_SHEET_NAME}."))
    return notes


def render_workbook_notes(parsed: Optional[Dict[str, Any]]) -> None:
    for level, message in workbook_notes(parsed):
        if level == "warning":
            st.warning(message, icon=":material/warning:")
        else:
            st.caption(message)


def render_controls_bar(
    workbook_name: str,
    parsed: Optional[Dict[str, Any]] = None,
    deck_info: Optional[Dict[str, Any]] = None,
) -> None:
    """The sidebar: view, workbook source and the filters (rendered on every run, so the
    filters apply to every view and never reset when the view changes)."""
    with st.sidebar:
        st.title(APP_TITLE)
        st.caption("Meeting navigation, workbook controls, and filters")
        st.caption(f"Version {APP_VERSION}")
        st.caption(f"Workbook: {md_escape(workbook_name)}")
        edited = len(native_overrides())
        if edited:
            st.caption(f"{_plural(edited, 'deal')} edited this meeting")

        st.divider()
        st.radio(
            "View",
            options=list(VIEW_OPTIONS),
            key="view_mode",
            horizontal=False,
            on_change=shell.request_view,
            persist_state="session",
        )

        st.divider()
        st.subheader("Workbook")
        st.file_uploader(
            "Upload workbook",
            type=["xlsx"],
            key="uploaded_workbook",
            on_change=on_workbook_upload_cb,
            max_upload_size=50,
        )
        st.file_uploader("Upload status history CSV backup", type=["csv"], key="uploaded_status_history")
        st.date_input(
            "As-of date",
            key="as_of_date",
            format="MM/DD/YYYY",
            on_change=mark_as_of_manual_cb,
            persist_state="session",
        )
        render_as_of_notes(parsed, deck_info)
        st.toggle("Include hidden rows", key="include_hidden", persist_state="session")
        render_workbook_notes(parsed)

        st.divider()
        st.subheader("Filters")
        st.segmented_control(
            "Portfolio type",
            options=list(SHEET_FILTER_OPTIONS),
            key="sheet_filter",
            required=True,
            persist_state="session",
            on_change=close_dialog_cb,
            width="stretch",
        )
        st.text_input(
            SEARCH_LABEL,
            key="search_query",
            persist_state="session",
            on_change=close_dialog_cb,
            placeholder="e.g. 10007, Harbor, Alice Chen",
        )
        st.toggle("Update reminders only", key="stale_only", persist_state="session", on_change=close_dialog_cb)
        # A callback: it runs before the widgets exist, so their keys can be written.
        st.button(
            "Reset filters",
            key="reset_filters",
            on_click=reset_filters_cb,
            width="stretch",
            icon=":material/filter_alt_off:",
        )


def render_status_history_notice(deck: pd.DataFrame) -> None:
    history = st.session_state.status_history
    if history is None:
        history = empty_status_history()
    history_name = display_text(st.session_state.get("uploaded_status_history_name", ""), blank="None loaded")
    open_count, addressed_count = reminder_counts(deck)

    c1, c2, c3 = st.columns(3)
    c1.metric("History rows", fmt_int(len(history)), border=True)
    c2.metric("History source", history_name, border=True)
    c3.metric("Update reminders", fmt_int(open_count), help=reminder_metric_help(addressed_count), border=True)

    if st.session_state.get("status_history_load_error"):
        st.error(f"Could not load status history: {st.session_state.status_history_load_error}")

    if history.empty:
        st.info(
            "No embedded status history is loaded yet. The app will still run; update reminders begin once "
            "you use a workbook generated by the weekly overview builder with the hidden history sheet."
        )


def reminder_counts(deck: pd.DataFrame) -> Tuple[int, int]:
    """(open update reminders, reminders addressed this meeting)."""
    if deck.empty or "status_needs_update_prompt" not in deck.columns:
        return 0, 0
    open_count = int(flag_array(deck["status_needs_update_prompt"]).sum())
    addressed = 0
    if "reminder_addressed" in deck.columns:
        addressed = int(sum(1 for value in deck["reminder_addressed"].tolist() if norm_label(value)))
    return open_count, addressed


def reminder_metric_help(addressed_count: int) -> str:
    text = (
        f"Deals whose status has stayed the same for {STATUS_STALE_MEETING_THRESHOLD}+ meetings. A reminder "
        "clears when the status changes, the AM commentary is refreshed, or the status is confirmed as still accurate."
    )
    if addressed_count:
        text += f" {_plural(addressed_count, 'reminder')} addressed this meeting."
    return text


def balances_caption(as_of_date: dt.date, deck_info: Optional[Dict[str, Any]] = None) -> str:
    parts = [f"As of {as_of_date:%m/%d/%Y}"]
    sources = [
        f"{sheet} '{info['upb_header']}'"
        for sheet, info in (deck_info or {}).get("sheets", {}).items()
        if info.get("upb_header")
    ]
    if sources:
        parts.append("UPB from " + " · ".join(sources))
    return " · ".join(parts)


def render_empty_filter_state(ctx: AppContext, where: str) -> None:
    """No deal matches the filters: say so and offer to clear them (never a dead end)."""
    unfiltered_reminders = filter_deck(
        ctx.deck,
        sheet_filter=st.session_state.get("sheet_filter"),
        query=st.session_state.get("search_query") or "",
        stale_only=False,
    )
    if st.session_state.get("stale_only") and not unfiltered_reminders.empty:
        st.success("All update reminders handled for the current filters.", icon=":material/task_alt:")
    st.info("No deals match the current filters", icon=":material/filter_alt_off:")
    parts = active_filter_parts()
    if parts:
        st.caption("Active filters: " + " · ".join(md_escape(part) for part in parts))
    st.button(
        "Clear filters",
        key=f"clear_filters::{where}",
        on_click=reset_filters_cb,
        icon=":material/filter_alt_off:",
    )


def render_overview_view(ctx: AppContext) -> None:
    deck = ctx.filtered
    if deck.empty:
        render_empty_filter_state(ctx, "overview")
        return
    total_upb = deck["upb"].fillna(0).sum()
    bridge_count = int((deck["sheet"] == "Bridge").sum())
    term_count = int((deck["sheet"] == "Term").sum())
    due_30, past_maturity = maturity_buckets(deck["days_to_maturity"])
    stale_count, addressed_count = reminder_counts(deck)

    if "loan_mod_count" in deck.columns:
        loan_mod_count_series = pd.to_numeric(deck["loan_mod_count"], errors="coerce").fillna(0)
    else:
        loan_mod_count_series = pd.Series([0] * len(deck), index=deck.index)
    loan_mod_deal_count = int((loan_mod_count_series.astype(int) > 0).sum())
    if "loan_mod_active_count" in deck.columns:
        cancelled_only = int(
            ((loan_mod_count_series > 0) & (pd.to_numeric(deck["loan_mod_active_count"], errors="coerce").fillna(0) == 0)).sum()
        )
    else:
        cancelled_only = 0
    loan_mod_help = "Deals with at least one loan modification record on file."
    if cancelled_only:
        loan_mod_help += f" {cancelled_only} of them only have cancelled or withdrawn requests."

    # Two rows of four, so values such as '$66.8MM' never truncate on a shared screen.
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Agenda items", fmt_int(len(deck)), border=True)
    m2.metric("Total UPB", fmt_money(total_upb, decimals=0), help=f"Exact: {fmt_money_exact(total_upb)}", border=True)
    m3.metric("Bridge / Term", f"{fmt_int(bridge_count)} / {fmt_int(term_count)}", border=True)
    m4.metric("Deals w/ loan mods", fmt_int(loan_mod_deal_count), help=loan_mod_help, border=True)
    m5, m6, m7, _spare = st.columns(4)
    m5.metric("Past maturity", fmt_int(past_maturity), help="Maturity date before the as-of date", border=True)
    m6.metric(
        "Maturing ≤30d",
        fmt_int(due_30),
        help=f"0–{MATURITY_WATCH_DAYS} days after the as-of date",
        border=True,
    )
    m7.metric("Update reminders", fmt_int(stale_count), help=reminder_metric_help(addressed_count), border=True)
    caption = balances_caption(ctx.as_of, ctx.deck_info)
    parts = active_filter_parts()
    if parts:
        caption += " · Filtered: " + ", ".join(md_escape(part) for part in parts)
    st.caption(caption)

    render_status_history_notice(deck)

    loan_mod_spotlight = build_loan_mod_spotlight_dataframe(deck)
    if not loan_mod_spotlight.empty:
        st.subheader("Loan modification spotlight")
        st.caption("Prioritizes the highlighted fields from the Loan Modifications sheet. Shows the latest mod per deal; use Presentation to toggle through every modification.")
        st.dataframe(
            loan_mod_spotlight,
            hide_index=True,
            width="stretch",
            column_config={
                "Comments": st.column_config.TextColumn("Comments", width="large"),
                "Modification Type": st.column_config.TextColumn("Modification Type", width="medium"),
            },
        )

    if stale_count:
        stale = deck[flag_array(deck["status_needs_update_prompt"])].copy()
        st.warning(f"{stale_count} deal(s) have an update reminder before the meeting is closed.")
        st.dataframe(
            stale[[
                "sheet",
                "deal_name",
                "deal_number",
                "status",
                "owner",
            ]].rename(
                columns={
                    "sheet": "Type",
                    "deal_name": "Deal Name",
                    "deal_number": "Deal #",
                    "status": "Status",
                    "owner": "Owner",
                }
            ),
            hide_index=True,
            width="stretch",
        )

    st.info(
        "Overview stays in original workbook order. Use Presentation for meeting discussion "
        "and Review for controlled update/commentary edits."
    )

    left, right = st.columns([1.15, 0.85], gap="large")
    with left:
        st.subheader("Agenda queue")
        st.dataframe(build_agenda_dataframe(deck), hide_index=True, width="stretch")

    with right:
        st.subheader("Upcoming timing")
        timing = deck[
            [
                "sheet",
                "deal_name",
                "maturity_date",
                "days_to_maturity",
                "next_payment_date",
                "days_to_next_payment",
            ]
        ].copy()
        timing["Maturity"] = timing["maturity_date"].map(fmt_date)
        timing["Maturity Timing"] = timing["days_to_maturity"].map(fmt_day_delta)
        timing["Next Payment"] = timing["next_payment_date"].map(fmt_date)
        timing["Payment Timing"] = timing["days_to_next_payment"].map(fmt_day_delta)
        st.dataframe(
            timing[["sheet", "deal_name", "Maturity", "Maturity Timing", "Next Payment", "Payment Timing"]].rename(
                columns={"sheet": "Type", "deal_name": "Deal Name"}
            ),
            hide_index=True,
            width="stretch",
        )


# -----------------------------------------------------------------------------
# Dialogs: one dispatcher, at most one dialog per run
# -----------------------------------------------------------------------------
def owner_header_label(deck_info: Optional[Dict[str, Any]], sheet: str) -> str:
    """The sheet's own owner header (Point Person, Asset Manager or Active RM)."""
    header = (((deck_info or {}).get("sheets") or {}).get(sheet) or {}).get("owner_header")
    return norm_label(header) or OWNER_LABEL_FALLBACK.get(sheet, "Owner")


def edit_dialog_payload(row: pd.Series, deck_info: Optional[Dict[str, Any]]) -> Dict[str, Any]:
    """Plain values the edit dialog shows (the deal on screen when Edit was clicked)."""
    sheet = norm_label(row.get("sheet"))
    reminder = ""
    if bool(row.get("status_needs_update_prompt", False)):
        reminder = str(row.get("status_prompt_message") or "")
    return {
        "deal_key": str(row.get("deal_key")),
        "deal_name": display_text(row.get("deal_name")),
        "deal_number": display_text(row.get("deal_number")),
        "sheet": sheet,
        "status": norm_label(row.get("status")),
        "owner": norm_label(row.get("owner")),
        "commentary": clean_free_text(row.get("commentary")),
        "owner_label": owner_header_label(deck_info, sheet),
        "reminder": reminder,
    }


@st.dialog("Update deal", width="large", icon=":material/edit_note:", on_dismiss=close_dialog_cb)
def edit_deal_dialog(deal: Dict[str, Any]) -> None:
    """Edit Status / owner / AM Commentary of the deal captured when 'Edit status' was
    clicked. Esc, X and click-outside run close_dialog_cb, so the dialog never reopens
    later for another deal; Enter in a text box does not submit."""
    deal_key = deal["deal_key"]
    active = st.session_state.get("active_dialog") or {}
    if active.get("kind") != DIALOG_EDIT or active.get("deal_key") != deal_key:
        st.rerun()  # stale: navigation already closed or retargeted this dialog
    st.caption(f"{md_escape(deal['deal_name'])} · {md_escape(deal['deal_number'])} · {md_escape(deal['sheet'])}")
    if deal.get("reminder"):
        st.warning(md_escape(deal["reminder"]), icon=":material/update:")
    with st.form(f"edit::{deal_key}", enter_to_submit=False, border=False):
        left, right = st.columns(2)
        with left:
            status = st.text_input("Status", value=deal["status"], key=f"edit_status::{deal_key}")
        with right:
            owner = st.text_input(md_escape(deal["owner_label"]), value=deal["owner"], key=f"edit_owner::{deal_key}")
        commentary = st.text_area(
            "AM Commentary",
            value=deal["commentary"],
            height="content",
            key=f"edit_commentary::{deal_key}",
        )
        submitted = st.form_submit_button("Save changes", type="primary", icon=":material/save:")
    if submitted:
        apply_edit(deal_key, status=status, owner=owner, commentary=commentary, source="dialog")
        close_dialog_cb()
        st.rerun()
    if st.button("Cancel", key=f"edit_cancel::{deal_key}"):
        close_dialog_cb()
        st.rerun()


def workbook_switch_summary(
    foreign: Dict[str, Dict[str, Any]],
    base_rows: Dict[str, Dict[str, Any]],
    workbook_names: Dict[str, str],
    new_name: str,
) -> Dict[str, Any]:
    """What the guard dialog says: how many edits, from which workbook(s), how many
    match deals in the new one, which do not, and the CSV of all of them."""
    items = [copy.deepcopy(item) for item in foreign.values()]
    names = dict(workbook_names)
    old_names = sorted(
        {names.get(str(item.get("workbook_sha") or ""), "") for item in items} - {""}
    )
    unmatched = [
        f"{norm_label(item.get('sheet'))} {norm_id(item.get('deal_number'))}".strip()
        for key, item in foreign.items()
        if key not in base_rows
    ]
    return {
        "total": len(items),
        "matched": sum(1 for key in foreign if key in base_rows),
        "unmatched": unmatched,
        "old_names": old_names,
        "new_name": new_name,
        # Built lazily on click, from copies captured now (download callables run on
        # another thread, without session state).
        "csv": lambda: edits_csv_bytes(items, names),
    }


@st.dialog("Different workbook loaded", width="medium", dismissible=False, icon=":material/swap_horiz:")
def workbook_switch_dialog(summary: Dict[str, Any]) -> None:
    """Edits made on another workbook are not applied until the user chooses here."""
    if not current_foreign_overrides():
        st.rerun()  # already decided (e.g. by another click)
    old = " and ".join(md_escape(name) for name in summary["old_names"]) or "an earlier workbook"
    st.markdown(
        f"You have {_plural(summary['total'], 'edit')} from {old}. "
        f"{summary['matched']:,} match deals in {md_escape(summary['new_name'])}."
    )
    if summary["unmatched"]:
        shown = summary["unmatched"][:12]
        more = len(summary["unmatched"]) - len(shown)
        st.markdown(
            "Not in this workbook (dropped if you keep the matching edits): "
            + ", ".join(md_escape(label) for label in shown)
            + (f" and {more} more" if more > 0 else "")
            + "."
        )
    st.caption("Until you choose, the deck shows this workbook's own values.")
    left, right = st.columns(2)
    with left:
        st.button(
            f"Keep {_plural(summary['matched'], 'matching edit')}",
            key="wb_switch_keep",
            type="primary",
            on_click=keep_matching_edits_cb,
            disabled=not summary["matched"],
            width="stretch",
        )
    with right:
        st.button("Discard all edits", key="wb_switch_discard", on_click=discard_foreign_edits_cb, width="stretch")
    st.download_button(
        "Download my edits first",
        data=summary["csv"],
        file_name=f"meeting_edits_{dt.date.today():%Y-%m-%d}.csv",
        mime="text/csv",
        key="wb_switch_download",
        on_click="ignore",
        icon=":material/download:",
        width="stretch",
    )


def current_foreign_overrides() -> Dict[str, Dict[str, Any]]:
    wb_sha = current_wb_sha()
    overrides = st.session_state.get("meeting_overrides") or {}
    return {key: item for key, item in overrides.items() if str(item.get("workbook_sha") or wb_sha) != wb_sha}


def render_active_dialog(ctx: AppContext) -> Optional[str]:
    """The single dialog dispatcher: renders at most one dialog per run and silently
    clears a target that no longer applies (deal gone, other view)."""
    active = st.session_state.get("active_dialog")
    kind, clear = dialog_to_render(
        active,
        view=ctx.view,
        deck_keys=ctx.deck_keys,
        has_foreign=bool(ctx.foreign_overrides),
    )
    if clear:
        close_dialog_cb()
    if kind == DIALOG_WORKBOOK_SWITCH:
        workbook_switch_dialog(
            workbook_switch_summary(
                ctx.foreign_overrides,
                ctx.base_rows,
                st.session_state.get("_wb_names") or {},
                ctx.workbook_name,
            )
        )
    elif kind == DIALOG_EDIT:
        row = ctx.row((active or {}).get("deal_key"))
        if row is None:
            close_dialog_cb()
            return None
        edit_deal_dialog(edit_dialog_payload(row, ctx.deck_info))
    return kind


# -----------------------------------------------------------------------------
# Error boundaries and privacy-safe logging
# -----------------------------------------------------------------------------
def configure_logging() -> None:
    """Send the 'wpo' INFO events to the server log (once per process)."""
    if getattr(LOG, "_wpo_configured", False):
        return
    if not LOG.handlers:
        handler = logging.StreamHandler()
        handler.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(name)s %(message)s"))
        LOG.addHandler(handler)
    LOG.setLevel(logging.INFO)
    LOG._wpo_configured = True  # type: ignore[attr-defined]


def incident_ref() -> str:
    """Five base32 characters to quote when reporting a failure."""
    return "".join(secrets.choice(INCIDENT_REF_ALPHABET) for _ in range(5))


def show_error_details() -> bool:
    """Tracebacks on screen only when running locally or with WPO_DEBUG=1."""
    if os.environ.get(DEBUG_ENV, "").strip() == "1":
        return True
    try:
        url = st.context.url
    except Exception:
        return False
    if not url:
        return False
    try:
        host = urllib.parse.urlparse(str(url)).hostname or ""
    except ValueError:
        return False
    return host.lower() in LOCAL_HOSTS


def traceback_frames(exc: BaseException, limit: int = 8) -> str:
    """'file:line:function' frames, innermost first. The exception message is left out:
    it can quote cell values."""
    frames = traceback.extract_tb(exc.__traceback__)[-limit:]
    return " < ".join(f"{Path(frame.filename).name}:{frame.lineno}:{frame.name}" for frame in reversed(frames))


def report_incident(where: str, exc: Exception) -> str:
    ref = incident_ref()
    LOG.error("event=error ref=%s where=%s type=%s at=%s", ref, where, type(exc).__name__, traceback_frames(exc))
    st.error(
        f"Ref {ref}. Your edits in this session are safe. Switch view or reload; send the ref to the app owner.",
        title=f"{where} could not render",
        icon=":material/error:",
    )
    if show_error_details():
        st.exception(exc)
    return ref


@contextmanager
def guard(where: str) -> Iterator[None]:
    """Error boundary around a view, dialog or panel: an Exception becomes a titled error
    with an incident ref (the rest of the page still renders). st.rerun, st.stop and
    st.switch_page raise BaseException subclasses and pass straight through."""
    try:
        yield
    except Exception as exc:  # noqa: BLE001 - this is the boundary
        report_incident(where, exc)




# -----------------------------------------------------------------------------
# Presentation HTML helpers
# -----------------------------------------------------------------------------
def html_safe(value: object, blank: str = "-") -> str:
    return html.escape(display_text(value, blank=blank), quote=True)


def html_multiline(value: object, blank: str = "-") -> str:
    """Escaped free text with line breaks as <br>, so the HTML stays on one line (a
    blank line inside st.markdown HTML would end the HTML block)."""
    text = clean_free_text(value)
    if not text:
        return html.escape(blank, quote=True)
    return "<br>".join(html.escape(line, quote=True) for line in text.split("\n"))


PRESENTATION_CSS = """
<style>
    .rt-hero {
        border: 1px solid rgba(148, 163, 184, 0.28);
        border-radius: 28px;
        padding: 1.28rem 1.38rem;
        background: linear-gradient(135deg, #ffffff 0%, #f7f9fc 48%, #eef4f8 100%);
        box-shadow: 0 18px 48px rgba(30, 41, 59, 0.10);
        margin: 0.70rem 0 0.88rem 0;
    }
    .rt-hero-top { display: flex; align-items: flex-start; justify-content: space-between; gap: 1.1rem; }
    .rt-eyebrow {
        display: inline-flex; border-radius: 999px; background: rgba(71, 85, 105, 0.08);
        color: #475569; border: 1px solid rgba(100, 116, 139, 0.16);
        padding: 0.34rem 0.72rem; font-size: 0.72rem; font-weight: 900;
        letter-spacing: 0.08em; text-transform: uppercase;
    }
    .rt-hero-title {
        max-width: 980px; font-size: clamp(1.45rem, 2.25vw, 2.75rem);
        line-height: 1.05; font-weight: 940; margin: 0.52rem 0 0.34rem 0;
        color: #111827; letter-spacing: -0.04em;
    }
    .rt-hero-subtitle { font-size: 1.02rem; color: #475569; font-weight: 720; line-height: 1.36; margin-bottom: 0.9rem; }
    .rt-chip-row { display: flex; flex-wrap: wrap; gap: 0.45rem; }
    .rt-chip {
        border: 1px solid rgba(100, 116, 139, 0.16); background: rgba(255, 255, 255, 0.72);
        color: #334155; border-radius: 999px; padding: 0.38rem 0.68rem; font-size: 0.82rem; font-weight: 820;
    }
    .rt-status-pill {
        border-radius: 22px; padding: 0.78rem 0.9rem; min-width: 190px; text-align: right;
        border: 1px solid rgba(100, 116, 139, 0.18); background: rgba(255, 255, 255, 0.76);
        box-shadow: 0 12px 24px rgba(30, 41, 59, 0.08);
    }
    .rt-status-pill .rt-label { font-size: 0.68rem; letter-spacing: 0.09em; text-transform: uppercase; color: #64748b; font-weight: 920; }
    .rt-status-pill .rt-value { margin-top: 0.18rem; font-size: 1.11rem; font-weight: 930; color: #111827; overflow-wrap: anywhere; }
    .rt-status-warning { background: rgba(255, 251, 235, 0.88); border-color: rgba(217, 119, 6, 0.28); }
    .rt-status-ok { background: rgba(255, 255, 255, 0.76); border-color: rgba(100, 116, 139, 0.18); }
    .rt-kpi-grid { display: grid; grid-template-columns: repeat(5, minmax(0, 1fr)); gap: 0.74rem; margin: 0.5rem 0 0.95rem 0; }
    .rt-kpi {
        border: 1px solid rgba(100, 116, 139, 0.16); border-radius: 22px; background: linear-gradient(180deg, #ffffff, #f8fafc);
        box-shadow: 0 12px 26px rgba(30, 41, 59, 0.07); padding: 1.02rem 1.04rem 0.94rem 1.04rem; min-height: 124px;
    }
    .rt-kpi-label { color: #64748b; font-size: 0.82rem; font-weight: 950; letter-spacing: 0.08em; text-transform: uppercase; margin-bottom: 0.40rem; }
    .rt-kpi-value { color: #111827; font-size: clamp(1.34rem, 1.85vw, 2.12rem); font-weight: 960; line-height: 1.02; overflow-wrap: anywhere; }
    .rt-kpi-helper { margin-top: 0.50rem; color: #64748b; font-size: 0.98rem; font-weight: 900; display: inline-flex; width: fit-content; border-radius: 999px; padding: 0.20rem 0.54rem; background: rgba(100, 116, 139, 0.08); }
    .rt-signal-green { color: #166534; background: #dcfce7; border: 1px solid rgba(22, 101, 52, 0.18); }
    .rt-signal-red { color: #991b1b; background: #fee2e2; border: 1px solid rgba(153, 27, 27, 0.18); }
    .rt-signal-amber { color: #92400e; background: #fef3c7; border: 1px solid rgba(146, 64, 14, 0.18); }
    .rt-signal-muted { color: #64748b; background: rgba(100, 116, 139, 0.08); border: 1px solid rgba(100, 116, 139, 0.12); }
    .rt-panel {
        border: 1px solid rgba(100, 116, 139, 0.16); border-radius: 26px; background: linear-gradient(180deg, #ffffff, #f8fafc);
        box-shadow: 0 14px 30px rgba(30, 41, 59, 0.065); padding: 1.02rem 1.05rem; margin-bottom: 0.85rem;
    }
    .rt-panel-header { display: flex; align-items: center; justify-content: space-between; gap: 0.75rem; margin-bottom: 0.82rem; padding-bottom: 0.68rem; border-bottom: 1px solid rgba(100, 116, 139, 0.14); }
    .rt-panel-title { font-size: 1.14rem; font-weight: 950; color: #111827; letter-spacing: -0.025em; }
    .rt-panel-note { border-radius: 999px; background: rgba(100, 116, 139, 0.09); padding: 0.28rem 0.6rem; color: #475569; font-size: 0.75rem; font-weight: 850; }
    .rt-field-grid { display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 0.72rem; }
    .rt-field { border-radius: 18px; background: rgba(241, 245, 249, 0.72); border: 1px solid rgba(100, 116, 139, 0.12); padding: 0.72rem 0.76rem; min-height: 78px; }
    .rt-field-label { font-size: 0.68rem; color: #64748b; font-weight: 920; text-transform: uppercase; letter-spacing: 0.07em; margin-bottom: 0.24rem; }
    .rt-field-value { font-size: 1.04rem; font-weight: 950; color: #111827; line-height: 1.14; overflow-wrap: anywhere; }
    .rt-status-card { border-radius: 24px; padding: 0.98rem 1.02rem; margin-bottom: 0.85rem; border: 1px solid rgba(217, 119, 6, 0.22); background: linear-gradient(135deg, #fffbeb, #ffffff); box-shadow: 0 12px 26px rgba(30, 41, 59, 0.055); }
    .rt-status-card-title { font-size: 1.04rem; font-weight: 950; color: #111827; margin-bottom: 0.30rem; }
    .rt-status-card-body { color: #475569; font-size: 0.93rem; font-weight: 700; line-height: 1.38; }
    .rt-prompt-list { display: flex; flex-direction: column; gap: 0.56rem; }
    .rt-prompt { display: grid; grid-template-columns: 34px minmax(0, 1fr); gap: 0.64rem; border-radius: 18px; border: 1px solid rgba(100, 116, 139, 0.14); background: rgba(248, 250, 252, 0.96); padding: 0.72rem 0.74rem; }
    .rt-prompt-num { width: 32px; height: 32px; display: inline-flex; align-items: center; justify-content: center; border-radius: 12px; background: #dbe6ee; color: #334155; font-weight: 950; font-size: 0.90rem; }
    .rt-prompt-text { color: #334155; font-weight: 760; line-height: 1.30; font-size: 0.95rem; }
    .rt-commentary { border-radius: 20px; border: 1px solid rgba(100, 116, 139, 0.14); background: rgba(255, 255, 255, 0.90); padding: 0.88rem 0.92rem; color: #334155; font-size: 0.96rem; font-weight: 700; line-height: 1.38; white-space: pre-wrap; }
    .rt-mod-notes { display: flex; flex-direction: column; gap: 0.66rem; }
    .rt-mod-note { border-radius: 18px; border: 1px solid rgba(100, 116, 139, 0.14); background: rgba(255, 255, 255, 0.94); padding: 0.78rem 0.84rem; }
    .rt-mod-note-label { color: #475569; font-size: 0.72rem; font-weight: 950; letter-spacing: 0.08em; text-transform: uppercase; margin-bottom: 0.36rem; }
    .rt-mod-note-body { color: #1f2937; font-size: 0.94rem; font-weight: 760; line-height: 1.42; overflow-wrap: anywhere; }
    .rt-mod-note-list { margin: 0.18rem 0 0 1.12rem; padding-left: 0.72rem; color: #1f2937; font-size: 0.94rem; font-weight: 760; line-height: 1.38; }
    .rt-mod-note-list li { margin-bottom: 0.34rem; padding-left: 0.08rem; }
    .rt-mod-banner { border-radius: 22px; border: 1px solid rgba(146, 64, 14, 0.20); background: linear-gradient(135deg, #fffbeb 0%, #ffffff 70%); box-shadow: 0 12px 26px rgba(30, 41, 59, 0.06); padding: 0.72rem 0.92rem; margin: 0.2rem 0 0.6rem 0; display: flex; align-items: center; justify-content: space-between; gap: 0.75rem; }
    .rt-mod-banner-title { font-size: 1.02rem; font-weight: 950; color: #111827; letter-spacing: -0.02em; }
    .rt-mod-banner-note { border-radius: 999px; background: rgba(146, 64, 14, 0.10); color: #92400e; padding: 0.26rem 0.62rem; font-size: 0.78rem; font-weight: 900; }
    @media (max-width: 1100px) { .rt-kpi-grid { grid-template-columns: repeat(2, minmax(0, 1fr)); } .rt-field-grid { grid-template-columns: 1fr; } .rt-hero-top { flex-direction: column; } .rt-status-pill { text-align: left; } }
</style>
"""


def apply_presentation_css() -> None:
    st.html(PRESENTATION_CSS)


def timing_signal(days: object, warn_within: int = 0) -> Tuple[str, str]:
    """(text, css class): past is red, 0..warn_within days is amber, later is green.

    Maturity uses warn_within=MATURITY_WATCH_DAYS; Next Payment keeps 0 (a monthly pay
    cycle would otherwise turn every current loan amber).
    """
    text_value = fmt_day_delta(days)
    if text_value == "-":
        return text_value, "rt-signal-muted"

    number = _finite_number(days)
    if number is None:
        return text_value, "rt-signal-muted"
    days_i = int(round(number))

    if days_i < 0:
        return text_value, "rt-signal-red"
    if days_i <= max(warn_within, 0):
        return text_value, "rt-signal-amber"
    return text_value, "rt-signal-green"


def render_html_kpi_strip(items: List[Tuple]) -> None:
    allowed_helper_classes = {
        "rt-signal-green",
        "rt-signal-red",
        "rt-signal-amber",
        "rt-signal-muted",
    }
    cards = []

    for item in items:
        tooltip: object = ""
        if len(item) >= 5:
            label, value, helper, helper_class, tooltip = item[:5]
        elif len(item) == 4:
            label, value, helper, helper_class = item
        else:
            label, value, helper = item
            helper_class = "rt-signal-muted"

        helper_text = html_safe(helper, blank="")
        helper_class = str(helper_class or "rt-signal-muted")
        if helper_class not in allowed_helper_classes:
            helper_class = "rt-signal-muted"

        helper_html = (
            f"<div class='rt-kpi-helper {helper_class}'>{helper_text}</div>"
            if helper_text
            else ""
        )
        tooltip_text = html_safe(tooltip, blank="")
        title_attr = f" title='{tooltip_text}'" if tooltip_text else ""
        cards.append(
            f"<div class='rt-kpi'{title_attr}>"
            f"<div class='rt-kpi-label'>{html_safe(label)}</div>"
            f"<div class='rt-kpi-value'>{html_safe(value)}</div>"
            f"{helper_html}"
            "</div>"
        )

    st.markdown(f"<div class='rt-kpi-grid'>{''.join(cards)}</div>", unsafe_allow_html=True)


def render_html_field_panel(title: str, items: List[Tuple[str, object]], note: object = "") -> None:
    fields = []
    for label, value in items:
        fields.append(
            "<div class='rt-field'>"
            f"<div class='rt-field-label'>{html_safe(label)}</div>"
            f"<div class='rt-field-value'>{html_safe(value)}</div>"
            "</div>"
        )
    note_text = html_safe(note, blank="")
    note_html = f"<div class='rt-panel-note'>{note_text}</div>" if note_text else ""
    st.markdown(
        "<section class='rt-panel'>"
        "<div class='rt-panel-header'>"
        f"<div class='rt-panel-title'>{html_safe(title)}</div>"
        f"{note_html}"
        "</div>"
        f"<div class='rt-field-grid'>{''.join(fields)}</div>"
        "</section>",
        unsafe_allow_html=True,
    )


def render_html_prompt_panel(prompts: List[str]) -> None:
    blocks = []
    for i, prompt in enumerate(prompts, start=1):
        blocks.append(
            "<div class='rt-prompt'>"
            f"<div class='rt-prompt-num'>{i}</div>"
            f"<div class='rt-prompt-text'>{html_safe(prompt)}</div>"
            "</div>"
        )
    st.markdown(
        "<section class='rt-panel'>"
        "<div class='rt-panel-header'><div class='rt-panel-title'>Presenter prompts</div></div>"
        f"<div class='rt-prompt-list'>{''.join(blocks)}</div>"
        "</section>",
        unsafe_allow_html=True,
    )


def render_html_commentary_panel(commentary: object) -> None:
    st.markdown(
        "<section class='rt-panel'>"
        "<div class='rt-panel-header'><div class='rt-panel-title'>AM commentary</div></div>"
        f"<div class='rt-commentary'>{html_multiline(commentary, blank='No commentary entered.')}</div>"
        "</section>",
        unsafe_allow_html=True,
    )


def render_agenda_jump(deck: pd.DataFrame, current_idx: int) -> None:
    labels = [f"{i + 1}. {row.sheet} | {row.deal_name} | {row.deal_number}" for i, row in enumerate(deck.itertuples(index=False))]
    st.session_state["_jump_key_map"] = dict(zip(labels, deck["deal_key"].astype(str)))
    st.selectbox("Jump to deal", options=labels, index=None, placeholder="Jump to deal # or name", filter_mode="contains", key="jump_to", on_change=jump_cb)

def render_html_status_check_panel(current_row: pd.Series) -> None:
    key = str(current_row.get("deal_key"))
    addressed = norm_label(current_row.get("reminder_addressed"))
    if addressed == REMINDER_CONFIRMED:
        st.caption("Status confirmed as still accurate for this meeting.")
        st.button("Undo confirmation", key=f"unconfirm_reminder::{key}", on_click=unconfirm_reminder_cb, args=(key,))
        return
    if not bool(current_row.get("status_needs_update_prompt", False)):
        return
    body = str(current_row.get("status_prompt_message") or "Review whether the meeting update should be refreshed.")
    st.markdown("<section class='rt-status-card rt-status-card-warning'><div class='rt-status-card-title'>Update reminder</div>" + f"<div class='rt-status-card-body'>{html_safe(body)}</div></section>", unsafe_allow_html=True)
    st.button("Confirm still accurate", key=f"confirm_reminder::{key}", on_click=confirm_reminder_cb, args=(key, current_meeting_date().isoformat()), icon=":material/verified:")

def build_capital_items(current_row: pd.Series) -> List[Tuple[str, object]]:
    # Exact dollars here so Commitment - Funded = Remaining reconciles on screen; the KPI
    # strip above keeps the compact '$12.5MM' style.
    if current_row.get("sheet") == "Bridge":
        return [
            ("UPB", fmt_money_exact(current_row.get("upb"))),
            ("Funded Amount", fmt_money_exact(current_row.get("funded_amount"))),
            ("Loan Commitment", fmt_money_exact(current_row.get("commitment"))),
            ("Remaining Commitment", fmt_money_exact(current_row.get("remaining_commitment"))),
            ("Maturity Date", fmt_date(current_row.get("maturity_date"))),
            ("Next Payment Date", fmt_date(current_row.get("next_payment_date"))),
            ("Financing", current_row.get("financing")),
            ("Loan Buyer", current_row.get("loan_buyer")),
        ]

    # Term: no 'Funded Amount' (the sheet has none; it would just repeat Loan Amount).
    return [
        ("UPB", fmt_money_exact(current_row.get("upb"))),
        ("Loan Commitment", fmt_money_exact(current_row.get("commitment"))),
        ("Loan Amount", fmt_money_exact(current_row.get("loan_amount"))),
        ("Maturity Date", fmt_date(current_row.get("maturity_date"))),
        ("Next Payment Date", fmt_date(current_row.get("next_payment_date"))),
        ("Financing", current_row.get("financing")),
        ("Loan Buyer", current_row.get("loan_buyer")),
        ("Segment", current_row.get("segment")),
    ]


def render_html_hero(current_row: pd.Series, current_idx: int, total_count: int) -> None:
    status_class = "rt-status-warning" if bool(current_row.get("status_needs_update_prompt", False)) else "rt-status-ok"
    st.markdown(
        "<section class='rt-hero'>"
        "<div class='rt-hero-top'>"
        "<div>"
        f"<div class='rt-eyebrow'>{html_safe(current_row.get('sheet'))} &bull; Agenda {current_idx + 1} of {total_count}</div>"
        f"<div class='rt-hero-title'>{html_safe(current_row.get('deal_name'))}</div>"
        f"<div class='rt-hero-subtitle'>Deal {html_safe(current_row.get('deal_number'))} &bull; Borrower: {html_safe(current_row.get('borrower'))}</div>"
        "<div class='rt-chip-row'>"
        f"<span class='rt-chip'>Owner: {html_safe(current_row.get('owner'))}</span>"
        f"<span class='rt-chip'>Servicer: {html_safe(current_row.get('servicer'))}</span>"
        f"<span class='rt-chip'>Portfolio: {html_safe(current_row.get('portfolio'))}</span>"
        f"<span class='rt-chip'>Segment: {html_safe(current_row.get('segment'))}</span>"
        "</div>"
        "</div>"
        f"<div class='rt-status-pill {status_class}'>"
        "<div class='rt-label'>Current update</div>"
        f"<div class='rt-value'>{html_safe(current_row.get('status'), blank='No status')}</div>"
        "</div>"
        "</div>"
        "</section>",
        unsafe_allow_html=True,
    )


_NOTE_NUMBER_MARKER_RE = re.compile(r"(?:(?<=\s)|^)(\d{1,2})([.)])[ \t]+(?=[A-Za-z$(\"'])", re.MULTILINE)
_NOTE_ROMAN_MARKERS = ("i", "ii", "iii", "iv", "v", "vi", "vii", "viii", "ix", "x", "xi", "xii")
_NOTE_ROMAN_MARKER_RE = re.compile(
    r"(?:(?<=\s)|^)\((" + "|".join(_NOTE_ROMAN_MARKERS) + r")\)[ \t]*", re.MULTILINE
)
_NOTE_BULLET_SPLIT_RE = re.compile(r"[ \t]*•[ \t]*")
_NOTE_SPACED_DASH_RE = re.compile(r"[ \t]+[-–—][ \t]+(?=[A-Za-z])")
_NOTE_LEADING_BULLET_RE = re.compile(r"^(?:[•·*][ \t]*|[-–—](?:[ \t]+|(?=[A-Za-z])))")


def _ascending_marker_run(matches: List[re.Match], number_of) -> List[re.Match]:
    """Markers that form the run 1, 2, 3, ... in text order (empty when shorter than 2)."""
    run: List[re.Match] = []
    expected = 1
    for match in matches:
        if number_of(match) == expected:
            run.append(match)
            expected += 1
    return run if len(run) >= 2 else []


def split_loan_mod_note_lines(label: str, value: object) -> List[str]:
    """Split a loan-mod note into display lines without changing any of its text.

    Splits only on line breaks, '•' bullets, dashes spaced on both sides and followed by
    a letter, numbered markers ('1. ' / '1) ') that form an ascending run from 1, and
    roman markers that form a run from '(i)'. Signs, digits, '$' and parentheticals such
    as '(V)' or '(LLC)' are never removed.
    """
    text = clean_free_text(value)
    if not text:
        return []

    # Remove duplicated field labels from Salesforce text, for example "Comments: ...".
    text = re.sub(rf"^{re.escape(label)}\s*:\s*", "", text, flags=re.IGNORECASE).strip()
    if not text:
        return []

    numbered = list(_NOTE_NUMBER_MARKER_RE.finditer(text))
    runs = [
        _ascending_marker_run([m for m in numbered if m.group(2) == style], lambda m: int(m.group(1)))
        for style in (".", ")")
    ]
    roman = _ascending_marker_run(
        list(_NOTE_ROMAN_MARKER_RE.finditer(text)),
        lambda m: _NOTE_ROMAN_MARKERS.index(m.group(1)) + 1,
    )
    cuts = sorted(
        [(m.start(), m.end()) for m in max(runs, key=len)] + [(m.start(), m.end()) for m in roman]
    )

    segments: List[str] = []
    position = 0
    for start, end in cuts:
        if start < position:
            continue
        segments.append(text[position:start])
        position = end
    segments.append(text[position:])

    lines: List[str] = []
    for segment in segments:
        for line in segment.split("\n"):
            for piece in _NOTE_BULLET_SPLIT_RE.split(line):
                for part in _NOTE_SPACED_DASH_RE.split(piece):
                    part = _NOTE_LEADING_BULLET_RE.sub("", part.strip()).strip().rstrip(";").strip()
                    if part:
                        lines.append(part)

    return lines or [text]


def loan_mod_note_section_html(label: str, value: object) -> str:
    lines = split_loan_mod_note_lines(label, value)
    if not lines:
        return ""

    if len(lines) == 1:
        body_html = f"<div class='rt-mod-note-body'>{html.escape(lines[0], quote=True)}</div>"
    else:
        items_html = "".join(f"<li>{html.escape(line, quote=True)}</li>" for line in lines)
        body_html = f"<ul class='rt-mod-note-list'>{items_html}</ul>"

    return (
        "<div class='rt-mod-note'>"
        f"<div class='rt-mod-note-label'>{html_safe(label)}</div>"
        f"{body_html}"
        "</div>"
    )


def render_html_loan_mod_notes_panel(current_row: pd.Series) -> None:
    html_block = loan_mod_note_section_html("Comments", current_row.get("loan_mod_latest_comments"))
    if not html_block:
        return

    st.markdown(
        "<section class='rt-panel'>"
        "<div class='rt-panel-header'><div class='rt-panel-title'>Highlighted loan mod comments</div></div>"
        f"<div class='rt-mod-notes'>{html_block}</div>"
        "</section>",
        unsafe_allow_html=True,
    )


def render_html_loan_modification_panel(current_row: pd.Series) -> None:
    if not has_loan_mods(current_row):
        return

    render_html_field_panel(
        "Loan modification spotlight",
        [
            ("Modification Type", current_row.get("loan_mod_latest_type")),
            ("Mod Status", current_row.get("loan_mod_latest_status")),
            ("Mod Order #", current_row.get("loan_mod_latest_order_number")),
            ("Completed Date", fmt_date(current_row.get("loan_mod_latest_completed_date"))),
            ("Effective Date", fmt_date(current_row.get("loan_mod_latest_effective_date"))),
            ("Previous Maturity", fmt_date(current_row.get("loan_mod_latest_previous_maturity_date"))),
            ("Updated Maturity", fmt_date(current_row.get("loan_mod_latest_updated_maturity_date"))),
            ("Previous Commitment", fmt_money_exact(current_row.get("loan_mod_latest_previous_commitment"))),
            ("Updated Commitment", fmt_money_exact(current_row.get("loan_mod_latest_updated_commitment"))),
        ],
        note=f"{fmt_int(current_row.get('loan_mod_count'))} mod record(s)",
    )

    render_html_loan_mod_notes_panel(current_row)


def render_html_loan_modification_record(
    record: Dict[str, object],
    position: int,
    total: int,
) -> None:
    """Render one specific modification record (used by the presentation toggle)."""
    note = f"Showing modification {position} of {total}"
    if bool(record.get("_mod_inactive", False)):
        note += f" · {display_text(record.get('Status'), blank='Inactive')} request, not in force"

    render_html_field_panel(
        "Loan modification detail",
        [
            ("Modification Type", loan_mod_type(record)),
            ("Mod Status", record.get("Status")),
            ("Mod Order #", record.get("Loan Mod Order Number")),
            ("Completed Date", fmt_date(record.get("System Task Completed Date"))),
            ("Effective Date", fmt_date(record.get("Mod Effective Date"))),
            ("Finalized Date", fmt_date(record.get("Modification Finalized Date"))),
            ("Previous Maturity", fmt_date(record.get("Previous Maturity Date"))),
            ("Updated Maturity", fmt_date(record.get("Updated Maturity Date"))),
            ("Previous Commitment", fmt_money_exact(record.get("Previous Loan Commitment"))),
            ("Updated Commitment", fmt_money_exact(record.get("Updated Loan Commitment"))),
        ],
        note=note,
    )

    comments_html = loan_mod_note_section_html("Comments", record.get("Comments"))
    if comments_html:
        st.markdown(
            "<section class='rt-panel'>"
            "<div class='rt-panel-header'><div class='rt-panel-title'>Highlighted loan mod comments</div></div>"
            f"<div class='rt-mod-notes'>{comments_html}</div>"
            "</section>",
            unsafe_allow_html=True,
        )


def render_presentation_loan_mods_top(
    current_row: pd.Series,
    loan_mods: pd.DataFrame,
    deal_key: str,
) -> None:
    """Loan modifications section, shown just below the KPI strip.

    The whole section is a collapsible dropdown (expander). When a loan carries
    more than one modification, the presenter can switch between them using either
    the button toggle or the dropdown selector; both stay in sync. The detail
    panel below always reflects whichever modification is selected.
    """
    mods = loan_mods_for_deal(loan_mods, current_row.get("deal_number"))
    if mods.empty:
        return

    total = len(mods)
    labels = build_loan_mod_toggle_labels(mods)

    with st.expander(f"Loan modifications ({fmt_int(total)} on this loan)", expanded=True):
        if total > 1:
            # A single shared position drives both the button toggle and the
            # dropdown, so changing one immediately updates the other.
            canon_key = f"loanmod_pos::{deal_key}"
            seg_key = f"loanmod_seg::{deal_key}"
            sel_key = f"loanmod_sel::{deal_key}"

            # Open on the headline mod (newest one in force) so the detail panel agrees
            # with the KPI and prompts; a presenter's later choice is kept.
            if canon_key not in st.session_state:
                st.session_state[canon_key] = headline_mod_position(mods)
            canonical = min(max(int(st.session_state.get(canon_key, 0) or 0), 0), total - 1)
            st.session_state[canon_key] = canonical

            def _sync_from_seg() -> None:
                value = st.session_state.get(seg_key)
                if value in labels:
                    st.session_state[canon_key] = labels.index(value)

            def _sync_from_sel() -> None:
                value = st.session_state.get(sel_key)
                if value in labels:
                    st.session_state[canon_key] = labels.index(value)

            # Preseed both widgets so they always reflect the shared selection.
            st.session_state[seg_key] = labels[canonical]
            st.session_state[sel_key] = labels[canonical]

            switch_left, switch_right = st.columns([0.62, 0.38], vertical_alignment="center")
            with switch_left:
                st.segmented_control(
                    "Toggle between modifications",
                    options=labels,
                    key=seg_key,
                    on_change=_sync_from_seg,
                )
            with switch_right:
                st.selectbox(
                    "Or pick from dropdown",
                    options=labels,
                    key=sel_key,
                    on_change=_sync_from_sel,
                )

            selected_pos = min(max(int(st.session_state.get(canon_key, 0) or 0), 0), total - 1)
        else:
            st.caption(f"Single modification on file: {labels[0]}")
            selected_pos = 0

        record = mods.iloc[selected_pos].to_dict()
        render_html_loan_modification_record(record, selected_pos + 1, total)


def loan_mod_kpi_helper(row: pd.Series) -> Tuple[str, str]:
    """(helper text, css class) for the Loan Mods KPI: the headline mod's type, amber
    only when a mod is in force; 'Cancelled request' when every request fell through."""
    count = int(row.get("loan_mod_count") or 0)
    if count <= 0:
        return "No mods", "rt-signal-muted"
    active = int(row.get("loan_mod_active_count", count) or 0)
    if active <= 0:
        return "Cancelled request", "rt-signal-muted"
    return display_text(row.get("loan_mod_latest_type"), blank="Type not set"), "rt-signal-amber"


def loan_size_kpi(row: pd.Series) -> Tuple[str, str, str, str, str]:
    """The 'Loan Commitment' tile: the explicit commitment, else (Term) the original loan
    amount; the helper names the source and the tooltip the workbook column."""
    value = fmt_money(row.get("commitment"), decimals=0)
    source = norm_label(row.get("commitment_source"))
    header = norm_label(row.get("commitment_header"))
    sheet = display_text(row.get("sheet"), blank="the")
    if source and header:
        tooltip = f"From the {sheet} column '{header}'"
    elif source:
        tooltip = f"From the {sheet} sheet"
    else:
        tooltip = f"No loan commitment or loan amount for this deal on the {sheet} sheet"
    return ("Loan Commitment", value, source or COMMITMENT_SOURCE_MISSING, "rt-signal-muted", tooltip)


def build_presentation_kpis(row: pd.Series) -> List[Tuple]:
    """The five KPI tiles of a Presentation slide (label, value, helper, class[, tooltip])."""
    maturity_timing, maturity_signal_class = timing_signal(row.get("days_to_maturity"), warn_within=MATURITY_WATCH_DAYS)
    payment_timing, payment_signal_class = timing_signal(row.get("days_to_next_payment"))
    from portfolio_cockpit import consistency_checks
    if any(check['code'] == 'STALE_PAYMENT_DATE' for check in consistency_checks(row)):
        payment_timing, payment_signal_class = 'Date not current?', 'rt-signal-amber'
    loan_mod_helper, loan_mod_signal = loan_mod_kpi_helper(row)
    return [
        ("UPB", fmt_money(row.get("upb"), decimals=0), "Current balance", "rt-signal-muted"),
        loan_size_kpi(row),
        ("Maturity", fmt_date(row.get("maturity_date")), maturity_timing, maturity_signal_class),
        ("Next Payment", fmt_date(row.get("next_payment_date")), payment_timing, payment_signal_class),
        ("Loan Mods", fmt_int(int(row.get("loan_mod_count") or 0)), loan_mod_helper, loan_mod_signal),
    ]


def render_presentation_view(ctx: AppContext) -> None:
    apply_presentation_css()
    if ctx.filtered.empty:
        render_empty_filter_state(ctx, "presentation")
    if not ctx.agenda_keys or ctx.cursor is None:
        return
    deck = ctx.deck.set_index("deal_key", drop=False).loc[ctx.agenda_keys].reset_index(drop=True)
    current_idx = ctx.agenda_keys.index(ctx.cursor)
    current_row = deck.iloc[current_idx]
    total_count = len(deck)
    progress_value = (current_idx + 1) / total_count
    deal_key = ctx.cursor
    loan_mods = ctx.loan_mods
    with st.container(border=True, key="presenter_bar"):
        top_left, top_mid, top_jump, top_prompt, top_edit = st.columns([0.8, 0.8, 3.2, 1.1, 1.2], vertical_alignment="bottom")
        with top_left:
            st.button("Previous", width="stretch", disabled=current_idx == 0, on_click=move_cursor_cb, args=(ctx.agenda_keys, -1), shortcut="Left")
        with top_mid:
            st.button("Next", width="stretch", disabled=current_idx >= total_count - 1, on_click=move_cursor_cb, args=(ctx.agenda_keys, 1), shortcut="Right")
        with top_jump:
            render_agenda_jump(deck, current_idx)
        with top_prompt:
            with st.popover("Prompts", icon=":material/lightbulb:", width="stretch"):
                render_html_prompt_panel((live_data.lg_presenter_prompts(current_row, (st.session_state.get("_data_context") or {}).get("lg")) + build_presenter_prompts(current_row))[:6])
        with top_edit:
            st.button("Edit status", type="primary", width="stretch", on_click=open_dialog_cb, args=(DIALOG_EDIT, deal_key), shortcut="E")
        st.progress(progress_value, text=f"Agenda item {current_idx + 1} of {total_count}")
        if ctx.cursor_outside:
            st.badge("Outside current filter", color="gray", icon=":material/filter_alt_off:")
    render_html_hero(current_row, current_idx, total_count)

    render_html_kpi_strip(build_presentation_kpis(current_row))
    with guard("Current loan picture"):
        data_views.render_current(current_row, ctx.deck, ctx.parsed, st.session_state.get("_data_context"))
        cockpit.render_deal(current_row, ctx, sys.modules[__name__])
    with guard("Meeting capture"):
        meeting_work.render_deal(ctx, current_row, sys.modules[__name__])

    # KPIs first, then the collapsible loan modifications section with the
    # button/dropdown toggle to switch between each modification on the loan.
    render_presentation_loan_mods_top(current_row, loan_mods, deal_key)

    left, right = st.columns([1.34, 0.86], gap="large", vertical_alignment="top")

    with left:
        render_html_field_panel(
            "Deal snapshot",
            [
                ("Borrower", current_row.get("borrower")),
                ("Account", current_row.get("account_display")),
                ("Deal #", current_row.get("deal_number")),
                ("Servicer", current_row.get("servicer")),
                ("Portfolio", current_row.get("portfolio")),
                ("Segment", current_row.get("segment")),
                ("Owner", current_row.get("owner")),
                ("NPL / Delinquency", current_row.get("npl_raw")),
            ],
            note=f"Days Past Due: {fmt_int(current_row.get('days_past_due'))}",
        )
        render_html_field_panel("Capital profile", build_capital_items(current_row))

    with right:
        render_html_status_check_panel(current_row)
        render_html_commentary_panel(current_row.get("commentary"))

        with st.expander("Nearby agenda", expanded=False):
            start = max(current_idx - 3, 0)
            end = min(current_idx + 4, total_count)
            agenda_slice = deck.iloc[start:end].copy()
            agenda_slice.insert(0, "Agenda #", range(start + 1, end + 1))
            agenda_slice["Current"] = agenda_slice["deal_key"].map(lambda value: "Current" if value == deal_key else "")
            agenda_slice["Update Reminder"] = agenda_slice["status_needs_update_prompt"].map(lambda value: "Yes" if bool(value) else "")
            agenda_slice["Loan Mods"] = agenda_slice["loan_mod_count"].map(lambda value: fmt_int(value) if int(value or 0) else "")
            agenda_slice["UPB"] = agenda_slice["upb"].map(lambda value: fmt_money(value, decimals=0))
            st.dataframe(
                agenda_slice[
                    [
                        "Agenda #",
                        "Current",
                        "sheet",
                        "deal_name",
                        "deal_number",
                        "Update Reminder",
                        "Loan Mods",
                        "status",
                        "owner",
                        "UPB",
                    ]
                ].rename(
                    columns={
                        "sheet": "Type",
                        "deal_name": "Deal Name",
                        "deal_number": "Deal #",
                        "status": "Status",
                        "owner": "Owner",
                    }
                ),
                hide_index=True,
                width="stretch",
            )


# -----------------------------------------------------------------------------
# Review and exports
# -----------------------------------------------------------------------------
def render_review_view(ctx: AppContext) -> None:
    st.subheader("Review and update")
    if ctx.filtered.empty:
        render_empty_filter_state(ctx, "review")
        return
    st.caption(f"Edits save when you leave a cell · {len(native_overrides()):,} deals edited this meeting")
    review_df = build_review_dataframe(ctx.filtered)
    editor_key = f"review_editor::{st.session_state.review_editor_version}"
    row_keys = ctx.filtered["deal_key"].astype(str).tolist()
    st.data_editor(review_df, key=editor_key, width="stretch", hide_index=True, column_order=list(REVIEW_COLUMN_ORDER), disabled=[column for column in REVIEW_COLUMN_ORDER if column not in REVIEW_EDITABLE_COLUMNS], on_change=commit_review_cb, args=(editor_key, row_keys), column_config={"AM Commentary": st.column_config.TextColumn("AM Commentary", width="large", max_chars=COMMENTARY_MAX_CHARS), "Status": st.column_config.TextColumn("Status", width="medium"), "Owner": st.column_config.TextColumn("Owner", width="medium")})

def render_exports_view(
    file_bytes: bytes,
    workbook_name: str,
    raw_deck: pd.DataFrame,
    deck: pd.DataFrame,
    as_of_date: dt.date,
) -> None:
    st.subheader("Exports")

    updated_history = merge_status_history(
        existing_history=st.session_state.status_history,
        deck=deck,
        meeting_date=as_of_date,
        workbook_name=workbook_name,
    )
    update_bytes = export_overrides_csv(raw_deck, deck)
    workbook_bytes = update_workbook_bytes(file_bytes, native_overrides())
    workbook_bytes = write_status_history_sheet_to_workbook_bytes(workbook_bytes, updated_history)
    agenda_bytes = build_agenda_dataframe(deck).to_csv(index=False).encode("utf-8")

    c1, c2 = st.columns(2)
    with c1:
        st.download_button(
            "Download changed statuses/commentary CSV",
            data=update_bytes,
            file_name="weekly_portfolio_meeting_updates.csv",
            mime="text/csv",
            width="stretch",
            disabled=(len(update_bytes) == 0),
        )
        st.download_button(
            "Download agenda CSV",
            data=agenda_bytes,
            file_name="weekly_portfolio_agenda.csv",
            mime="text/csv",
            width="stretch",
        )

    with c2:
        st.download_button(
            "Download updated workbook with embedded history",
            data=workbook_bytes,
            file_name=f"updated_{workbook_name}",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            width="stretch",
        )
        st.download_button(
            "Download status history CSV backup",
            data=status_history_to_csv_bytes(updated_history),
            file_name="status_history_backup.csv",
            mime="text/csv",
            width="stretch",
        )

    st.info(
        "The updated workbook includes the hidden _Status History sheet. Use that workbook or the "
        "Jupyter weekly overview generator as the normal path; the CSV is only a backup/export."
    )

# -----------------------------------------------------------------------------
# Workbook source and main app
# -----------------------------------------------------------------------------
def get_workbook_source() -> Tuple[Optional[bytes], str]:
    uploaded = st.session_state.get("uploaded_workbook")
    if uploaded is not None:
        return uploaded.getvalue(), uploaded.name

    env_path = os.environ.get(DEFAULT_WORKBOOK_ENV, "").strip()
    if env_path:
        path = Path(env_path).expanduser()
        if path.is_file():
            return path.read_bytes(), path.name

    default_path = first_existing_file(DEFAULT_SAMPLE_FILES)
    if default_path is None:
        return None, "No workbook loaded"

    return default_path.read_bytes(), default_path.name


def read_parsed_workbook(file_bytes: bytes, wb_sha: str) -> Dict[str, Any]:
    """Cached parse for this run; a skeleton holds the page while a new file is read."""
    st.session_state["_wb_sha"] = wb_sha
    if st.session_state.get("_parsed_sha") == wb_sha:
        return parse_workbook(file_bytes, wb_sha)
    with st.skeleton(height=400):
        parsed = parse_workbook(file_bytes, wb_sha)
    st.session_state["_parsed_sha"] = wb_sha
    return parsed


def render_no_workbook_loaded() -> None:
    st.info("Upload a Portfolio Overview workbook from the sidebar to start the meeting deck.")
    st.write("The app expects workbook sheets named Bridge and/or Term with a Deal Number header row.")
    st.subheader("Status history")
    st.write(
        "Status history is automatic when the workbook contains the hidden _Status History sheet "
        "created by the weekly overview generator. A CSV upload remains available in the sidebar only as a backup."
    )
    st.subheader("Loan modifications")
    st.write(
        "Loan modification details load automatically when the uploaded workbook includes a visible "
        f"{LOAN_MOD_SHEET_NAME!r} sheet with a 'Deal Number' or 'Opportunity Deal Loan Number' column."
    )


def main() -> None:
    st.set_page_config(page_title=APP_TITLE, layout="wide", menu_items={"About": f"{APP_TITLE} · {APP_VERSION}"})
    configure_logging()
    init_state()
    app = sys.modules[__name__]
    if hasattr(meeting_work, "restore_session"):
        meeting_work.restore_session()
    shell.navigation()
    render_toasts()
    wb = current_workbook()
    parsed = None
    raw_deck = pd.DataFrame()
    deck_info = None
    workbook_name = (wb or {}).get("name", "No workbook loaded")
    if wb is not None:
        parsed = read_parsed_workbook(wb["bytes"], wb["sha256"])
        if not parsed.get("error"):
            sync_as_of_with_workbook(parsed, wb["sha256"])
            raw_deck, deck_info = derive_deck(parsed, wb["sha256"], st.session_state.as_of_date.isoformat(), bool(st.session_state.include_hidden))
    render_controls_bar(workbook_name, parsed=parsed, deck_info=deck_info)
    shell.render_meeting_date(st.session_state.as_of_date)
    process_status_history_upload()
    if wb is None:
        render_no_workbook_loaded()
        return
    if parsed is None or parsed.get("error"):
        st.error("Could not load this workbook. Upload a valid Excel workbook to continue.")
        return
    load_embedded_status_history_if_available(parsed, wb["sha256"], workbook_name)
    if raw_deck.empty:
        st.warning("No Bridge or Term rows were found in the workbook.")
        return
    base_rows = refresh_base_rows(parsed, wb["sha256"])
    foreign = reconcile_overrides(base_rows, wb["sha256"])
    deck = apply_overrides(raw_deck, native_overrides(), wb["sha256"])
    deck = add_status_staleness_columns(deck, st.session_state.status_history, current_meeting_date(), workbook_name, history_sha=current_status_history_sha())
    deck = cockpit.apply_reminder_rule(deck, st.session_state.status_history, current_meeting_date(), st.session_state.get("reminder_rule", "Status only"))
    deck = mark_addressed_reminders(deck, native_overrides(), base_rows, st.session_state.reminder_confirmed, current_meeting_date().isoformat())
    mods = parsed.get("loan_mods")
    if mods is None:
        mods = empty_loan_modifications()
    ctx = AppContext(view=st.session_state.view_mode, wb=wb, parsed=parsed, deck_info=deck_info, raw_deck=raw_deck, deck=deck, filtered=deck, loan_mods=mods, as_of=st.session_state.as_of_date, base_rows=base_rows, foreign_overrides=foreign)
    ctx.app_version = APP_VERSION
    cockpit.render_settings(ctx, app)
    if hasattr(meeting_work, "checkpoint_overrides"):
        meeting_work.checkpoint_overrides(ctx, app)
        meeting_work.render_settings(ctx, app)
    with guard("Data sources"):
        data_context = data_views.prepare_context(deck, parsed)
        st.session_state["_data_context"] = data_context
        data_views.render_sources(deck, parsed, data_context)
    filtered = cockpit.apply_focus(apply_filters(deck), native_overrides())
    full_keys = deck["deal_key"].astype(str).tolist()
    filtered_keys = filtered["deal_key"].astype(str).tolist()
    cursor = resolve_cursor(st.session_state.cursor, full_keys, filtered_keys, st.session_state.get("_last_agenda") or ())
    st.session_state.cursor = cursor
    ordered_full = cockpit.order_deck(deck, st.session_state.get("agenda_order", "Workbook"))
    order = {str(key): (0, i) for i, key in enumerate(ordered_full["deal_key"])}
    agenda, outside = displayed_agenda(filtered_keys, cursor, order)
    st.session_state["_last_agenda"] = agenda
    ctx.filtered, ctx.agenda_keys, ctx.cursor = filtered, agenda, cursor
    ctx.cursor_outside, ctx.deck_keys = outside, frozenset(full_keys)
    st.session_state["_app_context"] = ctx
    with guard(ctx.view):
        if ctx.view == "Overview":
            render_overview_view(ctx)
            live_data.render_landgorilla_overview(filtered, (st.session_state.get("_data_context") or {}).get("lg"))
            cockpit.render_dashboard(ctx, app)
        elif ctx.view == "Presentation":
            render_presentation_view(ctx)
        elif ctx.view == "Review":
            render_review_view(ctx)
        elif ctx.view == "Follow-ups":
            meeting_work.render_followups(ctx, app)
        elif ctx.view == "Wrap-up":
            meeting_work.render_wrapup(ctx, app)
            verified_exports.render_exports(ctx, app, meeting_work)
        elif ctx.view == "Exports":
            verified_exports.render_exports(ctx, app, meeting_work)
    with guard("Dialog"):
        render_active_dialog(ctx)
    if hasattr(meeting_work, "checkpoint_overrides"):
        meeting_work.checkpoint_overrides(ctx, app)


if __name__ == "__main__":
    main()
