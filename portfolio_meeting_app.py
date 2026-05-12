from __future__ import annotations

import datetime as dt
import io
import re
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple

import numpy as np
import pandas as pd
import streamlit as st
from openpyxl import load_workbook


APP_TITLE = "Weekly Portfolio Meeting Deck"
APP_DIR = Path(__file__).resolve().parent

DEFAULT_SAMPLE_FILES = [
    APP_DIR / "Portfolio Overview 3.31.26.xlsx",
    APP_DIR / "portfolio_overview.xlsx",
    Path("Portfolio Overview 3.31.26.xlsx"),
    Path("portfolio_overview.xlsx"),
]

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


def norm_text(value: object) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and np.isnan(value):
        return ""
    text = str(value).replace("\xa0", " ")
    text = re.sub(r"\s+", " ", text).strip()
    if text.endswith(".0"):
        text = text[:-2]
    return text


def canon_header(value: object) -> str:
    return norm_text(value).upper()


def display_text(value: object, blank: str = "-") -> str:
    text = norm_text(value)
    return text if text else blank


def get_override_key(sheet: str, deal_number: str) -> str:
    return f"{norm_text(sheet)}::{norm_text(deal_number)}"


def first_existing_file(paths: Iterable[Path]) -> Optional[Path]:
    for path in paths:
        if path.exists():
            return path
    return None


def fmt_money(value: object, decimals: int = 0, blank: str = "-") -> str:
    if value is None or (isinstance(value, float) and np.isnan(value)):
        return blank
    try:
        number = float(value)
    except Exception:
        return blank
    if np.isnan(number):
        return blank
    if abs(number) >= 1_000_000:
        return f"${number / 1_000_000:,.1f}MM"
    if decimals == 0:
        return f"${number:,.0f}"
    return f"${number:,.{decimals}f}"


def fmt_int(value: object, blank: str = "-") -> str:
    if value is None or (isinstance(value, float) and np.isnan(value)):
        return blank
    try:
        return f"{int(round(float(value))):,}"
    except Exception:
        return blank


def fmt_date(value: object, blank: str = "-") -> str:
    if value is None:
        return blank
    try:
        timestamp = pd.to_datetime(value, errors="coerce")
    except Exception:
        return blank
    if pd.isna(timestamp):
        return blank
    return timestamp.strftime("%m/%d/%Y")


def fmt_day_delta(days: object, blank: str = "-") -> str:
    if days is None or (isinstance(days, float) and np.isnan(days)):
        return blank
    try:
        days_i = int(round(float(days)))
    except Exception:
        return blank
    if days_i == 0:
        return "Today"
    if days_i > 0:
        return f"In {days_i:,}d"
    return f"{abs(days_i):,}d overdue"


def safe_series(df: pd.DataFrame, column_name: Optional[str], default: object = "") -> pd.Series:
    if column_name and column_name in df.columns:
        return df[column_name]
    return pd.Series([default] * len(df), index=df.index)


def safe_numeric_series(df: pd.DataFrame, column_name: Optional[str]) -> pd.Series:
    if column_name and column_name in df.columns:
        return pd.to_numeric(df[column_name], errors="coerce")
    return pd.Series([np.nan] * len(df), index=df.index, dtype=float)


def safe_datetime_series(df: pd.DataFrame, column_name: Optional[str]) -> pd.Series:
    if column_name and column_name in df.columns:
        return pd.to_datetime(df[column_name], errors="coerce")
    return pd.Series([pd.NaT] * len(df), index=df.index)


def resolve_column(columns: Iterable[object], *candidates: Optional[str]) -> Optional[str]:
    cols = [str(c) for c in columns if c is not None]
    for candidate in candidates:
        if not candidate:
            continue
        for col in cols:
            if canon_header(col) == canon_header(candidate):
                return col
    return None


def matching_columns(columns: Iterable[object], candidate: str) -> List[str]:
    return [
        str(c)
        for c in columns
        if c is not None and canon_header(c) == canon_header(candidate)
    ]


def find_header_row(ws, search_rows: int = 30, key_header: str = "Deal Number") -> Tuple[int, int]:
    max_row = min(search_rows, ws.max_row)
    for r in range(1, max_row + 1):
        values = [ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
        if any(canon_header(v) == canon_header(key_header) for v in values if v is not None):
            return r, ws.max_column
    raise ValueError(f"Could not find header row in sheet {ws.title!r}.")


def parse_header_date(
    month: int,
    day: int,
    year_text: Optional[str],
    as_of_date: dt.date,
) -> Optional[dt.date]:
    try:
        if year_text:
            year = int(year_text)
            if year < 100:
                year += 2000
            return dt.date(year, month, day)

        candidate = dt.date(as_of_date.year, month, day)
        if candidate > as_of_date:
            candidate = dt.date(as_of_date.year - 1, month, day)
        return candidate
    except ValueError:
        return None


def latest_dated_by_suffix(
    headers: Iterable[object],
    as_of_date: Optional[dt.date] = None,
) -> Dict[str, str]:
    as_of_date = as_of_date or dt.date.today()
    best: Dict[str, Tuple[dt.date, str]] = {}

    for header in headers:
        if header is None:
            continue
        text = norm_text(header)
        match = DATE_HDR.match(text)
        if not match:
            continue

        header_date = parse_header_date(
            month=int(match.group(1)),
            day=int(match.group(2)),
            year_text=match.group(3),
            as_of_date=as_of_date,
        )
        if header_date is None:
            continue

        suffix = canon_header(match.group(4))
        if suffix not in best or header_date > best[suffix][0]:
            best[suffix] = (header_date, text)

    return {suffix: full for suffix, (_date_value, full) in best.items()}


def fallback_name(
    primary: pd.Series,
    secondary: pd.Series,
    tertiary: pd.Series,
    final_fallback: pd.Series,
) -> pd.Series:
    result = primary.fillna("").astype(str)
    for candidate in [secondary, tertiary, final_fallback]:
        mask = result.map(norm_text) == ""
        result = result.where(~mask, candidate.fillna("").astype(str))
    return result.map(norm_text)


def empty_status_history() -> pd.DataFrame:
    return pd.DataFrame(columns=STATUS_HISTORY_COLUMNS)


def normalize_status_history(history_df: Optional[pd.DataFrame]) -> pd.DataFrame:
    if history_df is None or history_df.empty:
        return empty_status_history()

    out = history_df.copy()
    for col in STATUS_HISTORY_COLUMNS:
        if col not in out.columns:
            out[col] = ""

    out = out[STATUS_HISTORY_COLUMNS].copy()
    out["meeting_date"] = pd.to_datetime(out["meeting_date"], errors="coerce")
    out["sheet"] = out["sheet"].map(norm_text)
    out["deal_number"] = out["deal_number"].map(norm_text)
    out["deal_key"] = out["deal_key"].map(norm_text)
    out["status"] = out["status"].map(norm_text)
    out["owner"] = out["owner"].map(norm_text)
    out["commentary"] = out["commentary"].fillna("").astype(str)
    out["workbook_name"] = out["workbook_name"].map(norm_text)
    out["saved_at"] = out["saved_at"].map(norm_text)

    out = out[
        out["meeting_date"].notna()
        & (out["deal_key"] != "")
        & (out["status"] != "")
    ].copy()

    if out.empty:
        return empty_status_history()

    out["meeting_date"] = out["meeting_date"].dt.date.astype(str)
    out = out.drop_duplicates(subset=["meeting_date", "deal_key"], keep="last")
    out = out.sort_values(["meeting_date", "sheet", "deal_number"], kind="stable")
    return out.reset_index(drop=True)


def read_status_history_csv(file_obj) -> pd.DataFrame:
    try:
        df = pd.read_csv(file_obj)
    except UnicodeDecodeError:
        file_obj.seek(0)
        df = pd.read_csv(file_obj, encoding="latin-1")
    return normalize_status_history(df)


def load_default_status_history_once() -> None:
    if st.session_state.get("status_history_default_checked", False):
        return

    st.session_state.status_history_default_checked = True
    if not STATUS_HISTORY_PATH.exists():
        return

    try:
        history = pd.read_csv(STATUS_HISTORY_PATH)
        st.session_state.status_history = normalize_status_history(history)
        st.session_state.uploaded_status_history_name = STATUS_HISTORY_PATH.name
    except Exception as exc:
        st.session_state.status_history_load_error = str(exc)


def current_status_history_snapshot(
    deck: pd.DataFrame,
    meeting_date: dt.date,
    workbook_name: str,
) -> pd.DataFrame:
    if deck.empty:
        return empty_status_history()

    snapshot = deck[["sheet", "deal_number", "deal_key", "status", "owner", "commentary"]].copy()
    snapshot.insert(0, "meeting_date", meeting_date.isoformat())
    snapshot["workbook_name"] = workbook_name
    snapshot["saved_at"] = dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    snapshot["status"] = snapshot["status"].map(norm_text)
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



def read_embedded_status_history_from_workbook_bytes(file_bytes: bytes) -> pd.DataFrame:
    """Read hidden/embedded status history from the uploaded workbook, if present."""
    try:
        workbook = load_workbook(io.BytesIO(file_bytes), data_only=True, read_only=True)
    except Exception:
        return empty_status_history()

    try:
        if STATUS_HISTORY_SHEET_NAME not in workbook.sheetnames:
            return empty_status_history()

        ws = workbook[STATUS_HISTORY_SHEET_NAME]
        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            return empty_status_history()

        headers = [norm_text(value) for value in rows[0]]
        records = []
        for row in rows[1:]:
            if row is None or all(norm_text(value) == "" for value in row):
                continue
            record = {}
            for idx, header in enumerate(headers):
                if not header:
                    continue
                record[header] = row[idx] if idx < len(row) else ""
            records.append(record)

        if not records:
            return empty_status_history()

        return normalize_status_history(pd.DataFrame(records))

    finally:
        workbook.close()


def load_embedded_status_history_if_available(file_bytes: bytes, workbook_name: str) -> None:
    """
    Auto-load status history from the workbook itself.

    Manual CSV upload still wins. Embedded workbook history is used only when the user
    has not uploaded a separate CSV during this session.
    """
    if st.session_state.get("uploaded_status_history") is not None:
        return

    signature = f"embedded:{workbook_name}:{len(file_bytes)}"
    if signature == st.session_state.get("embedded_status_history_signature", ""):
        return

    embedded = read_embedded_status_history_from_workbook_bytes(file_bytes)
    st.session_state.embedded_status_history_signature = signature

    if embedded.empty:
        return

    st.session_state.status_history = embedded
    st.session_state.uploaded_status_history_name = f"Embedded in {workbook_name}"
    st.session_state.status_history_load_error = ""


def write_status_history_sheet_to_workbook_bytes(file_bytes: bytes, history_df: pd.DataFrame) -> bytes:
    """Embed status history directly inside the workbook as a hidden sheet."""
    history = normalize_status_history(history_df)
    workbook = load_workbook(io.BytesIO(file_bytes))

    if STATUS_HISTORY_SHEET_NAME in workbook.sheetnames:
        del workbook[STATUS_HISTORY_SHEET_NAME]

    ws = workbook.create_sheet(STATUS_HISTORY_SHEET_NAME)
    ws.sheet_state = "hidden"

    for col_idx, header in enumerate(STATUS_HISTORY_COLUMNS, start=1):
        ws.cell(1, col_idx).value = header

    for row_idx, record in enumerate(history.to_dict("records"), start=2):
        for col_idx, header in enumerate(STATUS_HISTORY_COLUMNS, start=1):
            ws.cell(row_idx, col_idx).value = record.get(header, "")

    out = io.BytesIO()
    workbook.save(out)
    workbook.close()
    out.seek(0)
    return out.getvalue()


def mark_current_deal_reviewed(deal_key: str) -> None:
    flags = ensure_review_flags()
    flags[deal_key] = True


def status_streak_for_deal(
    history: pd.DataFrame,
    deal_key: str,
    current_status: str,
) -> Tuple[int, str]:
    deal_key = norm_text(deal_key)
    current_status = norm_text(current_status)
    if history.empty or not deal_key or not current_status:
        return 0, ""

    rows = history[history["deal_key"].map(norm_text) == deal_key].copy()
    if rows.empty:
        return 0, ""

    rows["meeting_date_dt"] = pd.to_datetime(rows["meeting_date"], errors="coerce")
    rows = rows.dropna(subset=["meeting_date_dt"])
    rows = rows.sort_values("meeting_date_dt", ascending=False, kind="stable")

    same_count = 0
    oldest_same_date = ""
    for _, row in rows.iterrows():
        if norm_text(row.get("status")) != current_status:
            break
        same_count += 1
        oldest_same_date = row["meeting_date_dt"].strftime("%m/%d/%Y")

    return same_count, oldest_same_date


def add_status_staleness_columns(
    deck: pd.DataFrame,
    status_history: pd.DataFrame,
    meeting_date: dt.date,
    workbook_name: str,
    threshold: int = STATUS_STALE_MEETING_THRESHOLD,
) -> pd.DataFrame:
    if deck.empty:
        return deck.copy()

    combined_history = merge_status_history(status_history, deck, meeting_date, workbook_name)
    out = deck.copy()
    same_counts: List[int] = []
    same_since: List[str] = []

    for _, row in out.iterrows():
        same_count, oldest_date = status_streak_for_deal(
            combined_history,
            deal_key=str(row.get("deal_key", "")),
            current_status=str(row.get("status", "")),
        )
        same_counts.append(same_count)
        same_since.append(oldest_date)

    out["status_same_meeting_count"] = same_counts
    out["status_same_since"] = same_since
    out["status_needs_update_prompt"] = out["status_same_meeting_count"] >= threshold
    out["status_prompt_message"] = out.apply(build_status_prompt_message, axis=1)
    return out


def build_status_prompt_message(row: pd.Series) -> str:
    if not bool(row.get("status_needs_update_prompt", False)):
        return ""

    status = display_text(row.get("status"), blank="blank")
    count = int(row.get("status_same_meeting_count") or 0)
    since = display_text(row.get("status_same_since"), blank="unknown")
    return (
        f"Status has been unchanged for {count} meetings since {since}. "
        f"Confirm whether '{status}' is still accurate or update the status/commentary."
    )


@st.cache_data(show_spinner=False)
def load_portfolio_workbook(
    file_bytes: bytes,
    as_of_date_iso: str,
    include_hidden: bool,
) -> Tuple[pd.DataFrame, Dict[str, Dict[str, str]]]:
    as_of_date = pd.Timestamp(as_of_date_iso).normalize()
    workbook = load_workbook(io.BytesIO(file_bytes), data_only=True)
    frames: List[pd.DataFrame] = []
    metadata: Dict[str, Dict[str, str]] = {}

    for sheet_name in ["Bridge", "Term"]:
        if sheet_name not in workbook.sheetnames:
            continue

        ws = workbook[sheet_name]
        header_row, last_col = find_header_row(ws)
        headers = [ws.cell(header_row, c).value for c in range(1, last_col + 1)]
        rows: List[Dict[str, object]] = []

        for r in range(header_row + 1, ws.max_row + 1):
            if not include_hidden and ws.row_dimensions[r].hidden:
                continue

            row_dict: Dict[str, object] = {}
            for c in range(1, last_col + 1):
                header = headers[c - 1]
                if header is None:
                    continue
                row_dict[str(header)] = ws.cell(r, c).value

            deal_col_in_row = resolve_column(row_dict.keys(), "Deal Number")
            deal_value = row_dict.get(deal_col_in_row) if deal_col_in_row else None
            if norm_text(deal_value) == "":
                continue

            row_dict["_excel_row"] = r
            rows.append(row_dict)

        if not rows:
            continue

        df = pd.DataFrame(rows)
        latest_columns = latest_dated_by_suffix(headers, as_of_date=as_of_date.date())

        deal_col = resolve_column(df.columns, "Deal Number")
        deal_name_col = resolve_column(df.columns, "Deal Name", "Loan Name", "Property Name")
        status_col = resolve_column(df.columns, "Status")
        commentary_col = resolve_column(df.columns, "AM Commentary")
        owner_col = resolve_column(df.columns, "Point Person", "Asset Manager", "Active RM")
        borrower_col = resolve_column(df.columns, "Borrower Name", "Borrower Entity", "Account Name")
        account_col = resolve_column(df.columns, "Account", "Account Name")
        servicer_col = resolve_column(df.columns, "Servicer")
        portfolio_col = resolve_column(df.columns, "Portfolio")
        segment_col = resolve_column(df.columns, "Segment")
        financing_col = resolve_column(df.columns, "Financing")
        loan_buyer_col = resolve_column(df.columns, "Loan Buyer")
        maturity_col = resolve_column(
            df.columns,
            "Current Maturity Date",
            "Maturity Date",
            "Original Maturity Date",
            "Next Advance Maturity Date",
        )
        next_payment_col = resolve_column(df.columns, "Next Payment Date")

        if sheet_name == "Bridge":
            upb_col = resolve_column(
                df.columns,
                latest_columns.get("UPB"),
                "UPB",
                "Active Funded Amount",
                "Loan Amount",
            )
        else:
            upb_col = resolve_column(df.columns, latest_columns.get("UPB"), "UPB", "Loan Amount")

        npl_col = resolve_column(
            df.columns,
            latest_columns.get("NPL"),
            "NPL",
            "Loan Level Delinquency",
            "DQ Status",
        )
        dpd_cols = matching_columns(df.columns, "Days Past Due")

        deal_series = safe_series(df, deal_col, "").map(norm_text)
        borrower_series = safe_series(df, borrower_col, "")
        account_series = safe_series(df, account_col, "")
        deal_name_series = fallback_name(
            safe_series(df, deal_name_col, ""),
            borrower_series,
            account_series,
            deal_series,
        )

        df["sheet"] = sheet_name
        df["sheet_order"] = 0 if sheet_name == "Bridge" else 1
        df["original_order"] = pd.to_numeric(df["_excel_row"], errors="coerce").fillna(999999).astype(int)
        df["deal_number"] = deal_series
        df["deal_name"] = deal_name_series
        df["borrower"] = borrower_series.map(norm_text)
        df["account_display"] = account_series.map(norm_text)
        df["servicer"] = safe_series(df, servicer_col, "").map(norm_text)
        df["portfolio"] = safe_series(df, portfolio_col, sheet_name).map(norm_text)
        df["segment"] = safe_series(df, segment_col, "").map(norm_text)
        df["financing"] = safe_series(df, financing_col, "").map(norm_text)
        df["loan_buyer"] = safe_series(df, loan_buyer_col, "").map(norm_text)
        df["owner"] = safe_series(df, owner_col, "").map(norm_text)
        df["status"] = safe_series(df, status_col, "").map(norm_text)
        df["commentary"] = safe_series(df, commentary_col, "").map(norm_text)
        df["saved_at"] = ""
        df["deal_key"] = df.apply(
            lambda row: get_override_key(str(row["sheet"]), str(row["deal_number"])),
            axis=1,
        )
        df["upb"] = safe_numeric_series(df, upb_col)
        df["maturity_date"] = safe_datetime_series(df, maturity_col)
        df["next_payment_date"] = safe_datetime_series(df, next_payment_col)
        df["npl_raw"] = safe_series(df, npl_col, "").fillna("").astype(str)

        if dpd_cols:
            dpd_matrix = np.column_stack(
                [
                    pd.to_numeric(df[col], errors="coerce").fillna(0).to_numpy(dtype=float)
                    for col in dpd_cols
                ]
            )
            df["days_past_due"] = np.maximum(dpd_matrix.max(axis=1), 0)
        else:
            df["days_past_due"] = 0

        if sheet_name == "Bridge":
            funded_amount_col = resolve_column(df.columns, "Active Funded Amount")
            commitment_col = resolve_column(df.columns, "Loan Commitment")
            remaining_commitment_col = resolve_column(df.columns, "Remaining Commitment")
            df["funded_amount"] = safe_numeric_series(df, funded_amount_col)
            df["commitment"] = safe_numeric_series(df, commitment_col)
            df["remaining_commitment"] = safe_numeric_series(df, remaining_commitment_col)
            df["loan_amount"] = np.nan
        else:
            loan_amount_col = resolve_column(df.columns, "Loan Amount")
            df["loan_amount"] = safe_numeric_series(df, loan_amount_col)
            df["funded_amount"] = np.nan
            df["commitment"] = np.nan
            df["remaining_commitment"] = np.nan

        metadata[sheet_name] = {
            "upb_header": upb_col or "",
            "maturity_header": maturity_col or "",
            "next_payment_header": next_payment_col or "",
            "owner_header": owner_col or "",
            "status_header": status_col or "",
            "commentary_header": commentary_col or "",
        }
        frames.append(df)

    workbook.close()

    if not frames:
        return pd.DataFrame(), metadata

    deck = pd.concat(frames, ignore_index=True)
    deck["days_to_maturity"] = (deck["maturity_date"] - as_of_date).dt.days
    deck["days_to_next_payment"] = (deck["next_payment_date"] - as_of_date).dt.days
    deck = deck.sort_values(
        ["sheet_order", "original_order"],
        ascending=[True, True],
        kind="stable",
    ).reset_index(drop=True)
    return deck, metadata


def available_status_suggestions(deck: pd.DataFrame) -> List[str]:
    if "status" not in deck.columns:
        return COMMON_STATUS_SUGGESTIONS
    existing = sorted({norm_text(v) for v in deck["status"].dropna().tolist() if norm_text(v)})
    return sorted(set(COMMON_STATUS_SUGGESTIONS + existing))


def initialize_state() -> None:
    defaults = {
        "view_mode": "Presentation",
        "sheet_filter": "All",
        "search_query": "",
        "include_hidden": False,
        "stale_only": False,
        "as_of_date": dt.date.today(),
        "selected_deal_key": None,
        "meeting_overrides": {},
        "review_flags": {},
        "dialog_target": None,
        "status_history": empty_status_history(),
        "status_history_default_checked": False,
        "uploaded_status_history_name": "",
        "uploaded_status_history_signature": "",
        "status_history_load_error": "",
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value


def ensure_override_store() -> Dict[str, Dict[str, str]]:
    if "meeting_overrides" not in st.session_state:
        st.session_state.meeting_overrides = {}
    return st.session_state.meeting_overrides


def ensure_review_flags() -> Dict[str, bool]:
    if "review_flags" not in st.session_state:
        st.session_state.review_flags = {}
    return st.session_state.review_flags


def upsert_override(
    base_row: pd.Series,
    status: str,
    owner: str,
    commentary: str,
) -> None:
    overrides = ensure_override_store()
    key = get_override_key(str(base_row.get("sheet")), str(base_row.get("deal_number")))
    payload = {
        "sheet": str(base_row.get("sheet")),
        "deal_number": str(base_row.get("deal_number")),
        "status": norm_text(status),
        "owner": norm_text(owner),
        "commentary": str(commentary or "").strip(),
        "saved_at": dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }

    original_status = norm_text(base_row.get("status"))
    original_owner = norm_text(base_row.get("owner"))
    original_commentary = str(base_row.get("commentary") or "")

    if (
        payload["status"] == original_status
        and payload["owner"] == original_owner
        and payload["commentary"] == original_commentary
    ):
        overrides.pop(key, None)
    else:
        overrides[key] = payload


def apply_overrides(deck: pd.DataFrame, overrides: Dict[str, Dict[str, str]]) -> pd.DataFrame:
    if deck.empty:
        return deck.copy()

    out = deck.copy()
    if "saved_at" not in out.columns:
        out["saved_at"] = ""

    if not overrides:
        return out

    override_df = pd.DataFrame(overrides.values())
    if override_df.empty:
        return out

    for field in ["sheet", "deal_number", "status", "owner", "commentary", "saved_at"]:
        if field not in override_df.columns:
            override_df[field] = ""

    merged = out.merge(
        override_df[["sheet", "deal_number", "status", "owner", "commentary", "saved_at"]],
        on=["sheet", "deal_number"],
        how="left",
        suffixes=("", "_override"),
    )

    for field in ["status", "owner", "commentary", "saved_at"]:
        override_field = f"{field}_override"
        merged[field] = merged[override_field].where(
            merged[override_field].notna() & (merged[override_field].astype(str) != ""),
            merged[field],
        )

    drop_cols = [col for col in merged.columns if col.endswith("_override")]
    return merged.drop(columns=drop_cols)


def apply_filters(deck: pd.DataFrame) -> pd.DataFrame:
    filtered = deck.copy()

    if st.session_state.sheet_filter == "Bridge":
        filtered = filtered[filtered["sheet"] == "Bridge"].copy()
    elif st.session_state.sheet_filter == "Term":
        filtered = filtered[filtered["sheet"] == "Term"].copy()

    query = st.session_state.search_query.strip().upper()
    if query:
        filtered = filtered[
            filtered["deal_number"].astype(str).str.upper().str.contains(query, na=False)
            | filtered["deal_name"].astype(str).str.upper().str.contains(query, na=False)
            | filtered["borrower"].astype(str).str.upper().str.contains(query, na=False)
        ].copy()


    if st.session_state.stale_only and "status_needs_update_prompt" in filtered.columns:
        filtered = filtered[filtered["status_needs_update_prompt"] == True].copy()  # noqa: E712

    return filtered.sort_values(
        ["sheet_order", "original_order"],
        ascending=[True, True],
        kind="stable",
    ).reset_index(drop=True)


def sync_selected_deal(deck: pd.DataFrame) -> None:
    if deck.empty:
        st.session_state.selected_deal_key = None
        return
    valid_keys = deck["deal_key"].tolist()
    if st.session_state.selected_deal_key not in valid_keys:
        st.session_state.selected_deal_key = valid_keys[0]


def get_selected_row(deck: pd.DataFrame) -> pd.Series:
    sync_selected_deal(deck)
    selected_key = st.session_state.selected_deal_key
    match = deck[deck["deal_key"] == selected_key]
    if match.empty:
        return deck.iloc[0]
    return match.iloc[0]


def get_row_by_key(deck: pd.DataFrame, deal_key: str) -> Optional[pd.Series]:
    match = deck[deck["deal_key"] == deal_key]
    if match.empty:
        return None
    return match.iloc[0]


def move_selection(deck: pd.DataFrame, step: int) -> None:
    sync_selected_deal(deck)
    keys = deck["deal_key"].tolist()
    current_key = st.session_state.selected_deal_key
    idx = keys.index(current_key)
    new_idx = min(max(idx + step, 0), len(keys) - 1)
    st.session_state.selected_deal_key = keys[new_idx]


def build_presenter_prompts(row: pd.Series) -> List[str]:
    prompts: List[str] = []

    if bool(row.get("status_needs_update_prompt", False)):
        prompts.append(str(row.get("status_prompt_message") or "Confirm whether the status should be updated."))

    maturity_text = fmt_date(row.get("maturity_date"))
    payment_text = fmt_date(row.get("next_payment_date"))
    dpd_text = fmt_int(row.get("days_past_due"))
    npl_text = display_text(row.get("npl_raw"), blank="")
    status_text = display_text(row.get("status"), blank="")

    if status_text and not bool(row.get("status_needs_update_prompt", False)):
        prompts.append(f"Confirm the current status remains accurate: {status_text}.")

    if maturity_text != "-":
        prompts.append(f"Confirm maturity timing: {maturity_text} ({fmt_day_delta(row.get('days_to_maturity'))}).")

    if payment_text != "-":
        prompts.append(f"Confirm next payment timing: {payment_text} ({fmt_day_delta(row.get('days_to_next_payment'))}).")

    if dpd_text != "-" and dpd_text != "0":
        prompts.append(f"Review payment delinquency: {dpd_text} days past due.")

    if npl_text:
        prompts.append(f"Confirm delinquency/NPL readout: {npl_text}.")

    if row.get("sheet") == "Bridge":
        remaining = fmt_money(row.get("remaining_commitment"), decimals=0)
        funded = fmt_money(row.get("funded_amount"), decimals=0)
        commitment = fmt_money(row.get("commitment"), decimals=0)
        if remaining != "-" or funded != "-" or commitment != "-":
            prompts.append(f"Review bridge capital profile: funded {funded}, commitment {commitment}, remaining {remaining}.")
    else:
        loan_amount = fmt_money(row.get("loan_amount"), decimals=0)
        if loan_amount != "-":
            prompts.append(f"Confirm term loan amount / UPB relationship: loan amount {loan_amount}.")

    prompts.append(f"Confirm owner and next follow-up: {display_text(row.get('owner'))}.")

    return prompts[:6]

def render_field_grid(items: List[Tuple[str, object]]) -> None:
    rows = [items[i : i + 2] for i in range(0, len(items), 2)]
    for row_items in rows:
        cols = st.columns(len(row_items))
        for col, (label, value) in zip(cols, row_items):
            with col:
                st.caption(label)
                st.write(display_text(value))



def render_compact_kpi(label: str, value: object, note: object = "") -> None:
    with st.container(border=True):
        st.caption(label)
        st.write(f"**{display_text(value)}**")
        note_text = display_text(note, blank="")
        if note_text:
            st.caption(note_text)


def primary_funded_display(row: pd.Series) -> str:
    funded = row.get("funded_amount")
    funded_text = fmt_money(funded, decimals=0)
    if funded_text != "-":
        return funded_text

    loan_amount_text = fmt_money(row.get("loan_amount"), decimals=0)
    if loan_amount_text != "-":
        return loan_amount_text

    return fmt_money(row.get("upb"), decimals=0)

def build_agenda_dataframe(deck: pd.DataFrame) -> pd.DataFrame:
    queue = deck.copy()
    queue.insert(0, "Agenda #", range(1, len(queue) + 1))
    queue["UPB"] = queue["upb"].map(lambda x: fmt_money(x, decimals=0))
    queue["Maturity"] = queue["maturity_date"].map(fmt_date)
    queue["Next Payment"] = queue["next_payment_date"].map(fmt_date)
    queue["Status"] = queue["status"].map(display_text)
    queue["Owner"] = queue["owner"].map(display_text)
    queue["Status Warning"] = queue["status_needs_update_prompt"].map(lambda value: "Yes" if bool(value) else "")

    return queue[
        [
            "Agenda #",
            "sheet",
            "deal_name",
            "deal_number",
            "Status Warning",
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
    review = deck.copy().set_index("deal_key")
    review.insert(0, "Agenda #", range(1, len(review) + 1))
    review["Type"] = review["sheet"]
    review["Deal Name"] = review["deal_name"].map(display_text)
    review["Deal #"] = review["deal_number"].map(display_text)
    review["Maturity"] = review["maturity_date"].map(fmt_date)
    review["Next Payment"] = review["next_payment_date"].map(fmt_date)
    review["UPB"] = review["upb"].map(lambda x: fmt_money(x, decimals=0))
    review["Status Update Needed"] = review["status_needs_update_prompt"].map(lambda value: "Yes" if bool(value) else "")
    review["Status"] = review["status"].fillna("").astype(str)
    review["Owner"] = review["owner"].fillna("").astype(str)
    review["AM Commentary"] = review["commentary"].fillna("").astype(str)

    return review[
        [
            "Agenda #",
            "Type",
            "Deal Name",
            "Deal #",
            "Maturity",
            "Next Payment",
            "UPB",
            "Status Update Needed",
            "Status",
            "Owner",
            "AM Commentary",
        ]
    ]


def apply_review_edits(
    edited_df: pd.DataFrame,
    current_deck: pd.DataFrame,
    raw_deck: pd.DataFrame,
) -> None:
    current_lookup = current_deck.set_index("deal_key")
    raw_lookup = raw_deck.set_index("deal_key")

    for deal_key, edited_row in edited_df.iterrows():
        if deal_key not in current_lookup.index or deal_key not in raw_lookup.index:
            continue

        current_row = current_lookup.loc[deal_key]
        raw_row = raw_lookup.loc[deal_key]
        new_status = norm_text(edited_row.get("Status", ""))
        new_owner = norm_text(edited_row.get("Owner", ""))
        new_commentary = str(edited_row.get("AM Commentary", "") or "").strip()

        if (
            new_status != norm_text(current_row.get("status"))
            or new_owner != norm_text(current_row.get("owner"))
            or new_commentary != str(current_row.get("commentary") or "")
        ):
            upsert_override(raw_row, new_status, new_owner, new_commentary)


def build_updates_dataframe(raw_deck: pd.DataFrame, deck: pd.DataFrame) -> pd.DataFrame:
    overrides = ensure_override_store()
    if not overrides:
        return pd.DataFrame()

    changed_keys = {
        get_override_key(item["sheet"], item["deal_number"])
        for item in overrides.values()
    }

    columns = [
        "deal_key",
        "sheet",
        "deal_name",
        "deal_number",
        "status",
        "owner",
        "commentary",
        "saved_at",
        "status_needs_update_prompt",
    ]
    out = deck[columns].copy()
    out = out[out["deal_key"].isin(changed_keys)].copy()
    if out.empty:
        return pd.DataFrame()

    out = out.rename(
        columns={
            "deal_key": "Deal Key",
            "sheet": "Type",
            "deal_name": "Deal Name",
            "deal_number": "Deal #",
            "status": "Status",
            "owner": "Owner",
            "commentary": "AM Commentary",
            "saved_at": "Last Saved",
            "status_needs_update_prompt": "Status Update Needed",
        }
    )
    out["Status Update Needed"] = out["Status Update Needed"].map(lambda value: "Yes" if bool(value) else "")

    return out[
        [
            "Type",
            "Deal Name",
            "Deal #",
            "Status Update Needed",
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
    return export_df.to_csv(index=False).encode("utf-8")


def update_workbook_bytes(file_bytes: bytes, overrides: Dict[str, Dict[str, str]]) -> bytes:
    if not overrides:
        return file_bytes

    workbook = load_workbook(io.BytesIO(file_bytes))

    for sheet_name, ws in [(name, workbook[name]) for name in workbook.sheetnames if name in {"Bridge", "Term"}]:
        relevant = [item for item in overrides.values() if item.get("sheet") == sheet_name]
        if not relevant:
            continue

        header_row, last_col = find_header_row(ws)
        headers = [ws.cell(header_row, c).value for c in range(1, last_col + 1)]
        key_col = next(
            c
            for c, header in enumerate(headers, start=1)
            if canon_header(header) == canon_header("Deal Number")
        )
        status_col = next(
            (c for c, header in enumerate(headers, start=1) if canon_header(header) == canon_header("Status")),
            None,
        )
        owner_col = next(
            (
                c
                for c, header in enumerate(headers, start=1)
                if canon_header(header) in {canon_header("Point Person"), canon_header("Asset Manager")}
            ),
            None,
        )
        commentary_col = next(
            (c for c, header in enumerate(headers, start=1) if canon_header(header) == canon_header("AM Commentary")),
            None,
        )

        row_lookup: Dict[str, int] = {}
        for r in range(header_row + 1, ws.max_row + 1):
            deal_number = norm_text(ws.cell(r, key_col).value)
            if deal_number:
                row_lookup[deal_number] = r

        for item in relevant:
            target_row = row_lookup.get(norm_text(item.get("deal_number")))
            if not target_row:
                continue
            if status_col is not None:
                ws.cell(target_row, status_col).value = item.get("status", "")
            if owner_col is not None:
                ws.cell(target_row, owner_col).value = item.get("owner", "")
            if commentary_col is not None:
                ws.cell(target_row, commentary_col).value = item.get("commentary", "")

    out = io.BytesIO()
    workbook.save(out)
    workbook.close()
    out.seek(0)
    return out.getvalue()


def process_status_history_upload() -> None:
    uploaded_history = st.session_state.get("uploaded_status_history")
    if uploaded_history is None:
        return

    signature = f"{uploaded_history.name}:{uploaded_history.size}"
    if signature == st.session_state.get("uploaded_status_history_signature", ""):
        return

    try:
        uploaded_history.seek(0)
        history = read_status_history_csv(uploaded_history)
        st.session_state.status_history = history
        st.session_state.uploaded_status_history_name = uploaded_history.name
        st.session_state.uploaded_status_history_signature = signature
        st.session_state.status_history_load_error = ""
    except Exception as exc:
        st.session_state.status_history_load_error = str(exc)




def render_controls_bar(workbook_name: str) -> None:
    with st.sidebar:
        st.title(APP_TITLE)
        st.caption("Meeting navigation and workbook controls")
        st.caption(f"Workbook: {workbook_name}")

        st.divider()
        st.radio(
            "View",
            options=["Overview", "Presentation", "Review", "Exports"],
            key="view_mode",
            horizontal=False,
        )

        st.divider()
        st.subheader("Workbook")
        st.file_uploader("Upload workbook", type=["xlsx"], key="uploaded_workbook")
        st.file_uploader("Upload status history CSV backup", type=["csv"], key="uploaded_status_history")
        st.date_input("As-of date", key="as_of_date")
        st.toggle("Include hidden rows", key="include_hidden")

        st.divider()
        st.subheader("Sidebar filters")
        st.text_input("Search deal # / name / borrower", key="search_query")
        st.toggle("Status warnings only", key="stale_only")

        if st.button("Reset filters", use_container_width=True):
            st.session_state.sheet_filter = "All"
            st.session_state.search_query = ""
            st.session_state.stale_only = False
            st.rerun()

def render_status_history_notice(deck: pd.DataFrame) -> None:
    history = normalize_status_history(st.session_state.status_history)
    history_name = display_text(st.session_state.get("uploaded_status_history_name", ""), blank="None loaded")
    warning_count = int(deck["status_needs_update_prompt"].sum()) if "status_needs_update_prompt" in deck.columns else 0

    c1, c2, c3 = st.columns(3)
    c1.metric("Status history rows", fmt_int(len(history)))
    c2.metric("History source", history_name)
    c3.metric("Status warnings", fmt_int(warning_count))

    if st.session_state.get("status_history_load_error"):
        st.error(f"Could not load status history: {st.session_state.status_history_load_error}")

    if history.empty:
        st.info(
            "No embedded status history is loaded yet. The app will still run; status warnings begin once "
            "you use a workbook generated by the weekly overview builder with the hidden history sheet."
        )

def render_overview_view(deck: pd.DataFrame, as_of_date: dt.date) -> None:
    total_upb = deck["upb"].fillna(0).sum()
    bridge_count = int((deck["sheet"] == "Bridge").sum())
    term_count = int((deck["sheet"] == "Term").sum())
    next_30 = int((deck["days_to_maturity"].fillna(9999) <= 30).sum())
    stale_count = int(deck["status_needs_update_prompt"].sum())

    m1, m2, m3, m4, m5 = st.columns(5)
    m1.metric("Agenda items", fmt_int(len(deck)))
    m2.metric("Total UPB", fmt_money(total_upb, decimals=0))
    m3.metric("Bridge / Term", f"{fmt_int(bridge_count)} / {fmt_int(term_count)}")
    m4.metric("Maturing in 30d", fmt_int(next_30), as_of_date.strftime("%m/%d/%Y"))
    m5.metric("Status warnings", fmt_int(stale_count))

    render_status_history_notice(deck)

    if stale_count:
        stale = deck[deck["status_needs_update_prompt"] == True].copy()  # noqa: E712
        st.warning(f"{stale_count} deal(s) need a status check before the meeting is closed.")
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
            use_container_width=True,
        )

    st.info(
        "Overview stays in original workbook order. Use Presentation for meeting discussion "
        "and Review for controlled status/commentary edits."
    )

    left, right = st.columns([1.15, 0.85], gap="large")
    with left:
        st.subheader("Agenda queue")
        st.dataframe(build_agenda_dataframe(deck), hide_index=True, use_container_width=True)

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
            use_container_width=True,
        )


def maybe_open_edit_dialog(current_deck: pd.DataFrame, raw_deck: pd.DataFrame) -> None:
    target = st.session_state.get("dialog_target")
    if not target:
        return

    current_row = get_row_by_key(current_deck, target)
    raw_row = get_row_by_key(raw_deck, target)
    if current_row is None or raw_row is None:
        st.session_state.dialog_target = None
        return

    @st.dialog("Edit status", width="large")
    def edit_dialog(current_row_dict: Dict[str, object], raw_row_dict: Dict[str, object]) -> None:
        current_row_local = pd.Series(current_row_dict)
        raw_row_local = pd.Series(raw_row_dict)
        deal_key = str(current_row_local.get("deal_key"))

        st.caption(
            f"{display_text(current_row_local.get('deal_name'))} | "
            f"Deal {display_text(current_row_local.get('deal_number'))}"
        )

        if bool(current_row_local.get("status_needs_update_prompt", False)):
            st.warning(str(current_row_local.get("status_prompt_message") or "Status review needed."))

        with st.form(f"edit-form::{deal_key}"):
            c1, c2 = st.columns(2)
            with c1:
                status = st.text_input("Status", value=str(current_row_local.get("status") or ""))
            with c2:
                owner_label = "Point Person" if current_row_local.get("sheet") == "Bridge" else "Asset Manager"
                owner = st.text_input(owner_label, value=str(current_row_local.get("owner") or ""))

            commentary = st.text_area(
                "AM Commentary",
                value=str(current_row_local.get("commentary") or ""),
                height=240,
            )

            save_clicked = st.form_submit_button("Save status", type="primary")

            if save_clicked:
                upsert_override(raw_row_local, status, owner, commentary)
                st.session_state.dialog_target = None
                st.rerun()

        if st.button("Cancel", use_container_width=True):
            st.session_state.dialog_target = None
            st.rerun()

    edit_dialog(current_row.to_dict(), raw_row.to_dict())


def render_presentation_view(deck: pd.DataFrame, raw_deck: pd.DataFrame) -> None:
    sync_selected_deal(deck)
    current_row = get_selected_row(deck)
    current_idx = int(deck.index[deck["deal_key"] == current_row["deal_key"]][0])
    total_count = len(deck)
    progress_value = 0.0 if total_count <= 1 else current_idx / max(total_count - 1, 1)
    deal_key = str(current_row.get("deal_key"))

    st.progress(progress_value, text=f"Agenda item {current_idx + 1} of {total_count}")

    nav_prev, nav_next, nav_filter, nav_jump, nav_prompt, nav_edit = st.columns(
        [0.7, 0.7, 1.25, 2.35, 1.0, 1.0],
        vertical_alignment="center",
    )
    with nav_prev:
        st.button(
            "Previous",
            use_container_width=True,
            disabled=current_idx == 0,
            on_click=move_selection,
            args=(deck, -1),
        )
    with nav_next:
        st.button(
            "Next",
            use_container_width=True,
            disabled=current_idx >= total_count - 1,
            on_click=move_selection,
            args=(deck, 1),
        )
    with nav_filter:
        st.segmented_control(
            "Portfolio type",
            options=["All", "Bridge", "Term"],
            key="sheet_filter",
            width="stretch",
        )
    with nav_jump:
        option_labels = [
            f"{i + 1}. {row.sheet} | {row.deal_name} | {row.deal_number}"
            for i, row in enumerate(
                deck[["sheet", "deal_name", "deal_number"]].itertuples(index=False, name="DealRow")
            )
        ]
        selected_label = st.selectbox("Jump to deal", options=option_labels, index=current_idx)
        selected_idx = option_labels.index(selected_label)
        if selected_idx != current_idx:
            st.session_state.selected_deal_key = deck.iloc[selected_idx]["deal_key"]
            st.rerun()
    with nav_prompt:
        with st.popover("Prompts", icon=":material/lightbulb:", width="stretch"):
            for prompt in build_presenter_prompts(current_row):
                st.write(f"- {prompt}")
    with nav_edit:
        if st.button("Edit status", type="primary", use_container_width=True):
            st.session_state.dialog_target = deal_key
            st.rerun()

    st.caption(f"{display_text(current_row.get('sheet'))} | Deal {display_text(current_row.get('deal_number'))}")
    st.title(display_text(current_row.get("deal_name")))
    st.write(
        f"Borrower: {display_text(current_row.get('borrower'))} | "
        f"Owner: {display_text(current_row.get('owner'))} | "
        f"Servicer: {display_text(current_row.get('servicer'))}"
    )

    if bool(current_row.get("status_needs_update_prompt", False)):
        st.warning(str(current_row.get("status_prompt_message") or "Status review needed."))
    else:
        st.success("Status check clear for this deal.")

    kpi_cols = st.columns(5)
    with kpi_cols[0]:
        render_compact_kpi("UPB", fmt_money(current_row.get("upb"), decimals=0))
    with kpi_cols[1]:
        render_compact_kpi("Funded Amount", primary_funded_display(current_row))
    with kpi_cols[2]:
        render_compact_kpi(
            "Maturity",
            fmt_date(current_row.get("maturity_date")),
            fmt_day_delta(current_row.get("days_to_maturity")),
        )
    with kpi_cols[3]:
        render_compact_kpi(
            "Next Payment",
            fmt_date(current_row.get("next_payment_date")),
            fmt_day_delta(current_row.get("days_to_next_payment")),
        )
    with kpi_cols[4]:
        render_compact_kpi("Status", display_text(current_row.get("status")))

    st.divider()

    st.subheader("Deal snapshot")
    render_field_grid(
        [
            ("Borrower", current_row.get("borrower")),
            ("Account", current_row.get("account_display")),
            ("Deal #", current_row.get("deal_number")),
            ("Owner", current_row.get("owner")),
            ("Servicer", current_row.get("servicer")),
            ("Portfolio", current_row.get("portfolio")),
            ("Segment", current_row.get("segment")),
            ("Financing", current_row.get("financing")),
            ("Loan Buyer", current_row.get("loan_buyer")),
            ("Days Past Due", fmt_int(current_row.get("days_past_due"))),
            ("NPL / Delinquency", current_row.get("npl_raw")),
        ]
    )

    st.subheader("Capital profile")
    if current_row.get("sheet") == "Bridge":
        render_field_grid(
            [
                ("UPB", fmt_money(current_row.get("upb"), decimals=0)),
                ("Funded Amount", fmt_money(current_row.get("funded_amount"), decimals=0)),
                ("Loan Commitment", fmt_money(current_row.get("commitment"), decimals=0)),
                ("Remaining Commitment", fmt_money(current_row.get("remaining_commitment"), decimals=0)),
                ("Maturity Date", fmt_date(current_row.get("maturity_date"))),
                ("Maturity Timing", fmt_day_delta(current_row.get("days_to_maturity"))),
                ("Next Payment", fmt_date(current_row.get("next_payment_date"))),
                ("Payment Timing", fmt_day_delta(current_row.get("days_to_next_payment"))),
            ]
        )
    else:
        render_field_grid(
            [
                ("UPB", fmt_money(current_row.get("upb"), decimals=0)),
                ("Funded Amount", primary_funded_display(current_row)),
                ("Loan Amount", fmt_money(current_row.get("loan_amount"), decimals=0)),
                ("Maturity Date", fmt_date(current_row.get("maturity_date"))),
                ("Maturity Timing", fmt_day_delta(current_row.get("days_to_maturity"))),
                ("Next Payment", fmt_date(current_row.get("next_payment_date"))),
                ("Payment Timing", fmt_day_delta(current_row.get("days_to_next_payment"))),
            ]
        )

    st.subheader("Nearby agenda items")
    start = max(current_idx - 3, 0)
    end = min(current_idx + 4, total_count)
    agenda_slice = deck.iloc[start:end].copy()
    agenda_slice.insert(0, "Agenda #", range(start + 1, end + 1))
    agenda_slice["Current"] = agenda_slice["deal_key"].map(lambda value: "Current" if value == deal_key else "")
    agenda_slice["Status Warning"] = agenda_slice["status_needs_update_prompt"].map(lambda value: "Yes" if bool(value) else "")
    agenda_slice["UPB"] = agenda_slice["upb"].map(lambda value: fmt_money(value, decimals=0))
    st.dataframe(
        agenda_slice[
            [
                "Agenda #",
                "Current",
                "sheet",
                "deal_name",
                "deal_number",
                "Status Warning",
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
        use_container_width=True,
    )

def render_review_view(deck: pd.DataFrame, raw_deck: pd.DataFrame) -> None:
    st.subheader("Review and update")
    st.caption("Edit Status, Owner, and AM Commentary in one place when you need a fast table view.")

    review_df = build_review_dataframe(deck)
    edited_df = st.data_editor(
        review_df,
        key="review_editor",
        use_container_width=True,
        hide_index=False,
        disabled=[
            "Agenda #",
            "Type",
            "Deal Name",
            "Deal #",
            "Maturity",
            "Next Payment",
            "UPB",
            "Status Update Needed",
        ],
        column_config={
            "AM Commentary": st.column_config.TextColumn("AM Commentary", width="large"),
            "Status": st.column_config.TextColumn("Status", width="medium"),
            "Owner": st.column_config.TextColumn("Owner", width="medium"),
        },
    )

    c1, c2 = st.columns([0.4, 0.6])
    with c1:
        if st.button("Apply review edits", type="primary", use_container_width=True):
            apply_review_edits(edited_df, deck, raw_deck)
            st.success("Review edits applied.")
            st.rerun()
    with c2:
        st.caption("Download the updated workbook from Exports after applying edits.")


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
    workbook_bytes = update_workbook_bytes(file_bytes, ensure_override_store())
    workbook_bytes = write_status_history_sheet_to_workbook_bytes(workbook_bytes, updated_history)
    agenda_bytes = build_agenda_dataframe(deck).to_csv(index=False).encode("utf-8")

    c1, c2 = st.columns(2)
    with c1:
        st.download_button(
            "Download changed statuses/commentary CSV",
            data=update_bytes,
            file_name="weekly_portfolio_meeting_updates.csv",
            mime="text/csv",
            use_container_width=True,
            disabled=(len(update_bytes) == 0),
        )
        st.download_button(
            "Download agenda CSV",
            data=agenda_bytes,
            file_name="weekly_portfolio_agenda.csv",
            mime="text/csv",
            use_container_width=True,
        )

    with c2:
        st.download_button(
            "Download updated workbook with embedded history",
            data=workbook_bytes,
            file_name=f"updated_{workbook_name}",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
        st.download_button(
            "Download status history CSV backup",
            data=status_history_to_csv_bytes(updated_history),
            file_name="status_history_backup.csv",
            mime="text/csv",
            use_container_width=True,
        )

    st.info(
        "The updated workbook includes the hidden _Status History sheet. Use that workbook or the "
        "Jupyter weekly overview generator as the normal path; the CSV is only a backup/export."
    )

    st.subheader("Recommended .gitignore entries")
    st.code(
        "status_history.csv\n"
        "*_status_history.csv\n"
        "weekly_portfolio_status_history*.csv\n"
        "*.xlsx\n",
        language="gitignore",
    )

def get_workbook_source() -> Tuple[Optional[bytes], str]:
    uploaded = st.session_state.get("uploaded_workbook")
    if uploaded is not None:
        return uploaded.getvalue(), uploaded.name

    default_path = first_existing_file(DEFAULT_SAMPLE_FILES)
    if default_path is None:
        return None, "No workbook loaded"

    return default_path.read_bytes(), default_path.name



def render_no_workbook_loaded() -> None:
    st.info("Upload a Portfolio Overview workbook from the sidebar to start the meeting deck.")
    st.write("The app expects workbook sheets named Bridge and/or Term with a Deal Number header row.")
    st.subheader("Status history")
    st.write(
        "Status history is automatic when the workbook contains the hidden _Status History sheet "
        "created by the weekly overview generator. A CSV upload remains available in the sidebar only as a backup."
    )

def main() -> None:
    st.set_page_config(page_title=APP_TITLE, layout="wide")
    initialize_state()
    load_default_status_history_once()

    file_bytes, workbook_name = get_workbook_source()
    render_controls_bar(workbook_name)
    process_status_history_upload()
    if file_bytes is not None:
        load_embedded_status_history_if_available(file_bytes, workbook_name)

    if file_bytes is None:
        render_no_workbook_loaded()
        return

    try:
        raw_deck, _metadata = load_portfolio_workbook(
            file_bytes=file_bytes,
            as_of_date_iso=st.session_state.as_of_date.isoformat(),
            include_hidden=bool(st.session_state.include_hidden),
        )
    except Exception as exc:
        st.error(f"Could not load workbook: {exc}")
        return

    if raw_deck.empty:
        st.warning("No Bridge or Term rows were found in the workbook.")
        return

    deck = apply_overrides(raw_deck, ensure_override_store())
    deck = add_status_staleness_columns(
        deck=deck,
        status_history=st.session_state.status_history,
        meeting_date=st.session_state.as_of_date,
        workbook_name=workbook_name,
    )
    filtered_deck = apply_filters(deck)

    if filtered_deck.empty:
        st.warning("No deals match the current filters.")
        return

    maybe_open_edit_dialog(deck, raw_deck)

    selected_view = st.session_state.view_mode
    if selected_view == "Overview":
        render_overview_view(filtered_deck, st.session_state.as_of_date)
    elif selected_view == "Presentation":
        render_presentation_view(filtered_deck, raw_deck)
    elif selected_view == "Review":
        render_review_view(filtered_deck, raw_deck)
    elif selected_view == "Exports":
        render_exports_view(
            file_bytes=file_bytes,
            workbook_name=workbook_name,
            raw_deck=raw_deck,
            deck=deck,
            as_of_date=st.session_state.as_of_date,
        )



if __name__ == "__main__":
    main()
