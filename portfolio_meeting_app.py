from __future__ import annotations

import datetime as dt
import html
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
DATE_HDR = re.compile(r"^\s*(\d{1,2})/(\d{1,2})\s+(.*\S)\s*$")

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
# Styling
# -----------------------------------------------------------------------------

def apply_app_css() -> None:
    st.markdown(
        """
        <style>
            .block-container {
                padding-top: 1.0rem;
                padding-bottom: 2rem;
                max-width: 1480px;
            }
            [data-testid="stSidebar"] {
                border-right: 1px solid rgba(15, 23, 42, 0.08);
            }
            .page-kicker {
                font-size: 0.8rem;
                font-weight: 800;
                letter-spacing: 0.08em;
                text-transform: uppercase;
                color: #0f766e;
                margin-bottom: 0.2rem;
            }
            .hero-card {
                border-radius: 28px;
                padding: 1.35rem 1.45rem 1.15rem 1.45rem;
                border: 1px solid rgba(15, 23, 42, 0.08);
                background:
                    radial-gradient(circle at top right, rgba(45, 212, 191, 0.18), transparent 35%),
                    linear-gradient(135deg, rgba(15, 23, 42, 0.03), rgba(15, 118, 110, 0.08));
                margin-bottom: 0.9rem;
            }
            .hero-kicker {
                display: inline-block;
                padding: 0.28rem 0.7rem;
                border-radius: 999px;
                font-size: 0.76rem;
                font-weight: 800;
                letter-spacing: 0.04em;
                text-transform: uppercase;
                background: rgba(15, 23, 42, 0.08);
                margin-bottom: 0.65rem;
            }
            .hero-title {
                font-size: 2.1rem;
                font-weight: 800;
                line-height: 1.08;
                margin-bottom: 0.35rem;
                color: #0f172a;
            }
            .hero-subtitle {
                color: rgba(15, 23, 42, 0.72);
                font-size: 1rem;
                margin-bottom: 0.7rem;
            }
            .chip-row {
                margin-top: 0.15rem;
            }
            .chip {
                display: inline-block;
                margin: 0.2rem 0.35rem 0 0;
                padding: 0.38rem 0.7rem;
                border-radius: 999px;
                background: rgba(15, 23, 42, 0.07);
                font-size: 0.82rem;
                font-weight: 700;
                color: #0f172a;
            }
            div[data-testid="stMetric"] {
                background: linear-gradient(180deg, rgba(255,255,255,0.98), rgba(248,250,252,0.98));
                border: 1px solid rgba(15, 23, 42, 0.08);
                border-radius: 18px;
                padding: 0.85rem 1rem;
            }
            .panel-card {
                border: 1px solid rgba(15, 23, 42, 0.08);
                border-radius: 22px;
                padding: 1rem 1rem 0.95rem 1rem;
                background: rgba(255,255,255,0.88);
                height: 100%;
            }
            .panel-title {
                font-size: 1rem;
                font-weight: 800;
                color: #0f172a;
                margin-bottom: 0.8rem;
            }
            .info-grid {
                display: grid;
                grid-template-columns: repeat(2, minmax(0, 1fr));
                gap: 0.75rem;
            }
            .info-item {
                border: 1px solid rgba(15, 23, 42, 0.07);
                border-radius: 16px;
                padding: 0.75rem 0.85rem;
                background: rgba(248,250,252,0.85);
            }
            .info-label {
                font-size: 0.72rem;
                font-weight: 800;
                letter-spacing: 0.04em;
                text-transform: uppercase;
                color: rgba(15, 23, 42, 0.55);
                margin-bottom: 0.32rem;
            }
            .info-value {
                font-size: 0.96rem;
                font-weight: 700;
                line-height: 1.3;
                color: #0f172a;
                word-break: break-word;
            }
            .timeline-grid {
                display: grid;
                grid-template-columns: repeat(3, minmax(0, 1fr));
                gap: 0.75rem;
            }
            .timeline-card {
                border: 1px solid rgba(15, 23, 42, 0.07);
                border-radius: 18px;
                padding: 0.85rem 0.95rem;
                background: rgba(248,250,252,0.9);
            }
            .timeline-label {
                font-size: 0.72rem;
                font-weight: 800;
                letter-spacing: 0.04em;
                text-transform: uppercase;
                color: rgba(15, 23, 42, 0.55);
                margin-bottom: 0.32rem;
            }
            .timeline-value {
                font-size: 1.02rem;
                font-weight: 800;
                color: #0f172a;
                margin-bottom: 0.18rem;
            }
            .timeline-subvalue {
                font-size: 0.86rem;
                color: rgba(15, 23, 42, 0.72);
            }
            .talk-list {
                margin: 0;
                padding-left: 1.15rem;
            }
            .talk-list li {
                margin-bottom: 0.46rem;
                color: #0f172a;
            }
            .small-muted {
                color: rgba(15, 23, 42, 0.72);
                font-size: 0.86rem;
            }
            .overview-banner {
                border: 1px solid rgba(15, 23, 42, 0.08);
                border-radius: 20px;
                padding: 0.9rem 1rem;
                background: linear-gradient(135deg, rgba(15, 118, 110, 0.06), rgba(15, 23, 42, 0.03));
                margin-bottom: 0.9rem;
            }
        </style>
        """,
        unsafe_allow_html=True,
    )


# -----------------------------------------------------------------------------
# Helpers
# -----------------------------------------------------------------------------

def norm_text(value: object) -> str:
    if value is None:
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


def find_header_row(ws, search_rows: int = 30, key_header: str = "Deal Number") -> Tuple[int, int]:
    max_row = min(search_rows, ws.max_row)
    for r in range(1, max_row + 1):
        values = [ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
        if any(canon_header(v) == canon_header(key_header) for v in values if v is not None):
            return r, ws.max_column
    raise ValueError(f"Could not find header row in sheet {ws.title!r}.")


def latest_dated_by_suffix(headers: Iterable[object]) -> Dict[str, str]:
    best: Dict[str, Tuple[Tuple[int, int], str]] = {}
    for header in headers:
        if header is None:
            continue
        text = norm_text(header)
        match = DATE_HDR.match(text)
        if not match:
            continue
        month_day = (int(match.group(1)), int(match.group(2)))
        suffix = canon_header(match.group(3))
        if suffix not in best or month_day > best[suffix][0]:
            best[suffix] = (month_day, text)
    return {suffix: full for suffix, (_month_day, full) in best.items()}


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
    return [str(c) for c in columns if c is not None and canon_header(c) == canon_header(candidate)]


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


def fmt_money(value: object, decimals: int = 0, blank: str = "-") -> str:
    if value is None or (isinstance(value, float) and np.isnan(value)):
        return blank
    try:
        number = float(value)
    except Exception:
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


def first_existing_file(paths: Iterable[Path]) -> Optional[Path]:
    for path in paths:
        if path.exists():
            return path
    return None


def fallback_name(primary: pd.Series, secondary: pd.Series, tertiary: pd.Series, final_fallback: pd.Series) -> pd.Series:
    result = primary.fillna("").astype(str)
    for candidate in [secondary, tertiary, final_fallback]:
        mask = result.map(norm_text) == ""
        result = result.where(~mask, candidate.fillna("").astype(str))
    return result.map(norm_text)


def agenda_label(row: pd.Series) -> str:
    return f"{int(row.get('agenda_number', 0)):02d} • {display_text(row.get('sheet'))} • {display_text(row.get('deal_name'))}"


def current_exposure_label(row: pd.Series) -> str:
    if display_text(row.get("sheet"), "") == "Bridge":
        return fmt_money(row.get("commitment"), decimals=0)
    return fmt_money(row.get("loan_amount"), decimals=0)


def current_exposure_title(row: pd.Series) -> str:
    if display_text(row.get("sheet"), "") == "Bridge":
        return "Commitment"
    return "Loan Amount"


# -----------------------------------------------------------------------------
# Workbook loading
# -----------------------------------------------------------------------------

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
        latest_columns = latest_dated_by_suffix(headers)

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
        upb_col = resolve_column(df.columns, latest_columns.get("UPB"), "UPB")
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

        df["upb"] = safe_numeric_series(df, upb_col)
        df["maturity_date"] = safe_datetime_series(df, maturity_col)
        df["next_payment_date"] = safe_datetime_series(df, next_payment_col)

        if dpd_cols:
            dpd_matrix = np.column_stack(
                [pd.to_numeric(df[col], errors="coerce").fillna(0).to_numpy(dtype=float) for col in dpd_cols]
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

    if not frames:
        return pd.DataFrame(), metadata

    deck = pd.concat(frames, ignore_index=True)
    deck["days_to_maturity"] = (deck["maturity_date"] - as_of_date).dt.days
    deck["days_to_next_payment"] = (deck["next_payment_date"] - as_of_date).dt.days
    deck = deck.sort_values(["sheet_order", "original_order"], ascending=[True, True], kind="stable").reset_index(drop=True)
    deck["agenda_number"] = range(1, len(deck) + 1)
    return deck, metadata


# -----------------------------------------------------------------------------
# Session / overrides
# -----------------------------------------------------------------------------

def get_override_key(sheet: str, deal_number: str) -> str:
    return f"{sheet}::{deal_number}"


def ensure_override_store() -> Dict[str, Dict[str, str]]:
    if "meeting_overrides" not in st.session_state:
        st.session_state.meeting_overrides = {}
    return st.session_state.meeting_overrides


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


# -----------------------------------------------------------------------------
# Presentation helpers
# -----------------------------------------------------------------------------

def build_meeting_prompts(row: pd.Series) -> List[str]:
    prompts: List[str] = []

    maturity_date = fmt_date(row.get("maturity_date"))
    next_payment_date = fmt_date(row.get("next_payment_date"))
    maturity_delta = fmt_day_delta(row.get("days_to_maturity"))
    payment_delta = fmt_day_delta(row.get("days_to_next_payment"))
    dpd_text = fmt_int(row.get("days_past_due"))
    status_text = display_text(row.get("status"))
    commentary_text = norm_text(row.get("commentary"))

    prompts.append(f"Maturity date: {maturity_date} ({maturity_delta}).")
    prompts.append(f"Next payment date: {next_payment_date} ({payment_delta}).")
    prompts.append(f"Days past due: {dpd_text}.")
    prompts.append(f"Status currently on file: {status_text}.")

    if commentary_text:
        prompts.append(f"Existing AM commentary: {commentary_text}.")
    else:
        prompts.append("Existing AM commentary is blank.")

    return prompts


def render_metric_row(row: pd.Series) -> None:
    cols = st.columns(5)
    cols[0].metric("UPB", fmt_money(row.get("upb"), decimals=0))
    cols[1].metric("Maturity", fmt_date(row.get("maturity_date")), fmt_day_delta(row.get("days_to_maturity")))
    cols[2].metric("Next Payment", fmt_date(row.get("next_payment_date")), fmt_day_delta(row.get("days_to_next_payment")))
    cols[3].metric("Days Past Due", fmt_int(row.get("days_past_due")))
    cols[4].metric(current_exposure_title(row), current_exposure_label(row))


def render_info_grid(items: List[Tuple[str, object]]) -> None:
    blocks = []
    for label, value in items:
        blocks.append(
            "<div class='info-item'>"
            f"<div class='info-label'>{html.escape(label)}</div>"
            f"<div class='info-value'>{html.escape(display_text(value))}</div>"
            "</div>"
        )
    st.markdown(f"<div class='info-grid'>{''.join(blocks)}</div>", unsafe_allow_html=True)


def render_timeline_strip(row: pd.Series) -> None:
    items = [
        ("Maturity Date", fmt_date(row.get("maturity_date")), fmt_day_delta(row.get("days_to_maturity"))),
        ("Next Payment", fmt_date(row.get("next_payment_date")), fmt_day_delta(row.get("days_to_next_payment"))),
        ("Days Past Due", fmt_int(row.get("days_past_due")), "Current servicing view"),
    ]
    cards = []
    for label, value, subvalue in items:
        cards.append(
            "<div class='timeline-card'>"
            f"<div class='timeline-label'>{html.escape(label)}</div>"
            f"<div class='timeline-value'>{html.escape(value)}</div>"
            f"<div class='timeline-subvalue'>{html.escape(subvalue)}</div>"
            "</div>"
        )
    st.markdown(f"<div class='timeline-grid'>{''.join(cards)}</div>", unsafe_allow_html=True)


def build_queue_dataframe(deck: pd.DataFrame) -> pd.DataFrame:
    queue = deck.copy()
    needed_defaults = {
        "agenda_number": 0,
        "sheet": "",
        "deal_number": "",
        "deal_name": "",
        "status": "",
        "owner": "",
        "upb": np.nan,
        "maturity_date": pd.NaT,
        "next_payment_date": pd.NaT,
        "days_past_due": np.nan,
        "saved_at": "",
        "sheet_order": 99,
        "original_order": 999999,
    }
    for col, default_value in needed_defaults.items():
        if col not in queue.columns:
            queue[col] = default_value

    queue = queue.sort_values(["sheet_order", "original_order"], ascending=[True, True], kind="stable").reset_index(drop=True)
    queue = queue[
        [
            "agenda_number",
            "sheet",
            "deal_number",
            "deal_name",
            "status",
            "owner",
            "upb",
            "maturity_date",
            "next_payment_date",
            "days_past_due",
            "saved_at",
        ]
    ].copy()

    queue = queue.rename(
        columns={
            "agenda_number": "Agenda #",
            "sheet": "Type",
            "deal_number": "Deal #",
            "deal_name": "Deal Name",
            "status": "Status",
            "owner": "Owner",
            "upb": "UPB",
            "maturity_date": "Maturity",
            "next_payment_date": "Next Payment",
            "days_past_due": "DPD",
            "saved_at": "Last Saved",
        }
    )

    queue["UPB"] = queue["UPB"].map(lambda x: fmt_money(x, decimals=0))
    queue["Maturity"] = queue["Maturity"].map(fmt_date)
    queue["Next Payment"] = queue["Next Payment"].map(fmt_date)
    queue["DPD"] = queue["DPD"].map(fmt_int)
    return queue


# -----------------------------------------------------------------------------
# Export helpers
# -----------------------------------------------------------------------------

def export_overrides_csv(overrides: Dict[str, Dict[str, str]]) -> bytes:
    if not overrides:
        return b""
    df = pd.DataFrame(overrides.values()).sort_values(["sheet", "deal_number"])
    return df.to_csv(index=False).encode("utf-8")


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
            c for c, header in enumerate(headers, start=1) if canon_header(header) == canon_header("Deal Number")
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
    out.seek(0)
    return out.getvalue()


# -----------------------------------------------------------------------------
# UI rendering
# -----------------------------------------------------------------------------

def render_overview_tab(deck: pd.DataFrame, as_of_date: dt.date, metadata: Dict[str, Dict[str, str]]) -> None:
    if deck.empty:
        st.warning("No deals match the current filters.")
        return

    total_upb = deck["upb"].fillna(0).sum()
    bridge_count = int((deck["sheet"] == "Bridge").sum())
    term_count = int((deck["sheet"] == "Term").sum())
    next_30 = int((deck["days_to_maturity"].fillna(9999) <= 30).sum())

    st.markdown(
        f"""
        <div class='overview-banner'>
            <div class='page-kicker'>Meeting Format</div>
            <div style='font-size:1.05rem; font-weight:700; color:#0f172a;'>
                Agenda is presented in original workbook order, with Bridge first and Term second.
            </div>
            <div class='small-muted' style='margin-top:0.25rem;'>
                As of {as_of_date.strftime('%m/%d/%Y')} • Bridge fields use {html.escape(metadata.get('Bridge', {}).get('maturity_header', 'Maturity Date') or 'Maturity Date')} • Term fields use {html.escape(metadata.get('Term', {}).get('maturity_header', 'Maturity Date') or 'Maturity Date')}
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Deals in agenda", f"{len(deck):,}")
    m2.metric("Total UPB", fmt_money(total_upb, decimals=0))
    m3.metric("Bridge / Term", f"{bridge_count} / {term_count}")
    m4.metric("Maturing in 30d", f"{next_30:,}")

    bridge_df = build_queue_dataframe(deck[deck["sheet"] == "Bridge"].copy())
    term_df = build_queue_dataframe(deck[deck["sheet"] == "Term"].copy())

    left, right = st.columns(2)
    with left:
        st.subheader("Bridge agenda")
        st.dataframe(bridge_df, use_container_width=True, hide_index=True)
    with right:
        st.subheader("Term agenda")
        st.dataframe(term_df, use_container_width=True, hide_index=True)


def render_presentation_tab(deck: pd.DataFrame, status_suggestions: List[str]) -> None:
    if deck.empty:
        st.warning("No deals match the current filters.")
        return

    if "deck_index" not in st.session_state:
        st.session_state.deck_index = 0
    st.session_state.deck_index = int(np.clip(st.session_state.deck_index, 0, len(deck) - 1))

    nav1, nav2, nav3, nav4 = st.columns([1, 1, 1.4, 2.8])
    with nav1:
        if st.button("⬅ Previous", use_container_width=True, disabled=st.session_state.deck_index == 0):
            st.session_state.deck_index -= 1
            st.rerun()
    with nav2:
        if st.button("Next ➡", use_container_width=True, disabled=st.session_state.deck_index >= len(deck) - 1):
            st.session_state.deck_index += 1
            st.rerun()
    with nav3:
        if st.button("Save + Next", use_container_width=True, key="top_save_next"):
            st.session_state["request_save_next"] = True
    with nav4:
        selected_position = st.selectbox(
            "Agenda item",
            options=list(range(len(deck))),
            index=st.session_state.deck_index,
            format_func=lambda i: agenda_label(deck.iloc[i]),
        )
        if selected_position != st.session_state.deck_index:
            st.session_state.deck_index = selected_position
            st.rerun()

    row = deck.iloc[st.session_state.deck_index]
    progress = (st.session_state.deck_index + 1) / len(deck)
    st.progress(progress, text=f"Agenda item {int(row['agenda_number'])} of {int(deck['agenda_number'].max())}")

    hero_html = f"""
        <div class='hero-card'>
            <div class='hero-kicker'>{html.escape(display_text(row.get('sheet')))} • Agenda {int(row.get('agenda_number', 0))}</div>
            <div class='hero-title'>{html.escape(display_text(row.get('deal_name')))}</div>
            <div class='hero-subtitle'>
                Deal {html.escape(display_text(row.get('deal_number')))} • Borrower {html.escape(display_text(row.get('borrower')))}
            </div>
            <div class='chip-row'>
                <span class='chip'>Status: {html.escape(display_text(row.get('status')))}</span>
                <span class='chip'>Owner: {html.escape(display_text(row.get('owner')))}</span>
                <span class='chip'>Servicer: {html.escape(display_text(row.get('servicer')))}</span>
                <span class='chip'>Portfolio: {html.escape(display_text(row.get('portfolio')))}</span>
            </div>
        </div>
    """
    st.markdown(hero_html, unsafe_allow_html=True)
    render_metric_row(row)

    left, right = st.columns([1.45, 1.0])

    with left:
        st.markdown("<div class='panel-card'><div class='panel-title'>Loan profile</div>", unsafe_allow_html=True)
        render_info_grid(
            [
                ("Deal Number", row.get("deal_number")),
                ("Borrower", row.get("borrower")),
                ("Servicer", row.get("servicer")),
                ("Owner / Point Person", row.get("owner")),
                ("Portfolio", row.get("portfolio")),
                ("Segment", row.get("segment")),
                ("Financing", row.get("financing")),
                ("Loan Buyer", row.get("loan_buyer")),
                ("Account", row.get("account_display")),
                (current_exposure_title(row), current_exposure_label(row)),
            ]
        )
        st.markdown("</div>", unsafe_allow_html=True)

        st.write("")
        st.markdown("<div class='panel-card'><div class='panel-title'>Dates at a glance</div>", unsafe_allow_html=True)
        render_timeline_strip(row)
        st.markdown("</div>", unsafe_allow_html=True)

        st.write("")
        prompts = build_meeting_prompts(row)
        prompt_html = "".join(f"<li>{html.escape(item)}</li>" for item in prompts)
        st.markdown(
            f"""
            <div class='panel-card'>
                <div class='panel-title'>Talking points</div>
                <ul class='talk-list'>{prompt_html}</ul>
                <div class='small-muted' style='margin-top:0.7rem;'>
                    Clean presentation mode with original order preserved and only the fields you need for the meeting.
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with right:
        st.markdown("<div class='panel-card'><div class='panel-title'>Meeting console</div>", unsafe_allow_html=True)
        st.caption("Update the current deal live during the meeting. Saved changes can be exported on the Meeting Log tab.")
        if status_suggestions:
            st.caption("Quick status ideas: " + " | ".join(status_suggestions[:7]))

        owner_label = "Point Person" if display_text(row.get("sheet"), "") == "Bridge" else "Asset Manager"
        status_value = st.text_input(
            "Status",
            value=str(row.get("status", "")),
            key=f"status_input::{row['sheet']}::{row['deal_number']}",
        )
        owner_value = st.text_input(
            owner_label,
            value=str(row.get("owner", "")),
            key=f"owner_input::{row['sheet']}::{row['deal_number']}",
        )
        commentary_value = st.text_area(
            "AM commentary / meeting notes",
            value=str(row.get("commentary", "")),
            height=260,
            key=f"commentary_input::{row['sheet']}::{row['deal_number']}",
        )

        button_col1, button_col2 = st.columns(2)
        save_clicked = button_col1.button("Save update", use_container_width=True, type="primary")
        save_next_clicked = button_col2.button("Save + next", use_container_width=True)

        if st.session_state.pop("request_save_next", False):
            save_next_clicked = True

        if save_clicked or save_next_clicked:
            overrides = ensure_override_store()
            key = get_override_key(str(row.get("sheet")), str(row.get("deal_number")))
            overrides[key] = {
                "sheet": str(row.get("sheet")),
                "deal_number": str(row.get("deal_number")),
                "status": status_value,
                "owner": owner_value,
                "commentary": commentary_value,
                "saved_at": dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            }
            if save_next_clicked and st.session_state.deck_index < len(deck) - 1:
                st.session_state.deck_index += 1
                st.rerun()
            st.success("Update saved.")

        st.markdown("</div>", unsafe_allow_html=True)


def render_meeting_log_tab(raw_deck: pd.DataFrame, displayed_deck: pd.DataFrame, file_bytes: bytes, workbook_name: str) -> None:
    overrides = ensure_override_store()
    st.metric("Saved meeting updates", f"{len(overrides):,}")

    if overrides:
        updates_df = pd.DataFrame(overrides.values()).sort_values(["sheet", "deal_number"])
        st.dataframe(updates_df, use_container_width=True, hide_index=True)
    else:
        st.info("No live updates saved yet. Use Presentation Mode to edit status, owner, or commentary.")

    dl1, dl2 = st.columns(2)
    with dl1:
        csv_bytes = export_overrides_csv(overrides)
        st.download_button(
            "Download meeting updates CSV",
            data=csv_bytes,
            file_name=f"{Path(workbook_name).stem}_meeting_updates.csv",
            mime="text/csv",
            disabled=not bool(overrides),
            use_container_width=True,
        )
    with dl2:
        workbook_bytes = update_workbook_bytes(file_bytes, overrides)
        st.download_button(
            "Download updated workbook",
            data=workbook_bytes,
            file_name=f"{Path(workbook_name).stem}_meeting_ready.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    st.subheader("Agenda queue")
    st.dataframe(build_queue_dataframe(displayed_deck), use_container_width=True, hide_index=True)

    st.subheader("Meeting-ready export view")
    export_df = displayed_deck.copy()
    needed_defaults = {
        "agenda_number": 0,
        "sheet": "",
        "deal_number": "",
        "deal_name": "",
        "status": "",
        "owner": "",
        "commentary": "",
        "upb": np.nan,
        "maturity_date": pd.NaT,
        "next_payment_date": pd.NaT,
        "days_past_due": np.nan,
        "sheet_order": 99,
        "original_order": 999999,
    }
    for col, default_value in needed_defaults.items():
        if col not in export_df.columns:
            export_df[col] = default_value

    export_df = export_df.sort_values(["sheet_order", "original_order"], ascending=[True, True], kind="stable")
    export_df["UPB"] = export_df["upb"].map(lambda x: fmt_money(x, decimals=0))
    export_df["Maturity Date"] = export_df["maturity_date"].map(fmt_date)
    export_df["Next Payment Date"] = export_df["next_payment_date"].map(fmt_date)
    export_df["Days Past Due"] = export_df["days_past_due"].map(fmt_int)

    st.dataframe(
        export_df[
            [
                "agenda_number",
                "sheet",
                "deal_number",
                "deal_name",
                "status",
                "owner",
                "UPB",
                "Maturity Date",
                "Next Payment Date",
                "Days Past Due",
                "commentary",
            ]
        ].rename(
            columns={
                "agenda_number": "Agenda #",
                "sheet": "Type",
                "deal_number": "Deal #",
                "deal_name": "Deal Name",
                "status": "Status",
                "owner": "Owner",
                "commentary": "AM Commentary",
            }
        ),
        use_container_width=True,
        hide_index=True,
    )


# -----------------------------------------------------------------------------
# Sidebar and main
# -----------------------------------------------------------------------------

def sidebar_file_picker() -> Tuple[Optional[bytes], str]:
    st.sidebar.header("Workbook")
    uploaded = st.sidebar.file_uploader("Upload Portfolio Overview workbook", type=["xlsx"])
    if uploaded is not None:
        return uploaded.getvalue(), uploaded.name

    sample_file = first_existing_file(DEFAULT_SAMPLE_FILES)
    if sample_file is not None:
        st.sidebar.caption(f"Using local workbook: {sample_file.name}")
        return sample_file.read_bytes(), sample_file.name

    return None, ""


def main() -> None:
    st.set_page_config(page_title=APP_TITLE, layout="wide")
    apply_app_css()

    st.markdown("<div class='page-kicker'>Redwood Weekly Review</div>", unsafe_allow_html=True)
    st.title(APP_TITLE)
    st.caption("Cleaner presentation mode built for weekly Bridge and Term review, preserving original workbook order.")

    file_bytes, workbook_name = sidebar_file_picker()
    if file_bytes is None:
        st.warning("Upload your Portfolio Overview workbook to begin.")
        st.stop()

    as_of_date = st.sidebar.date_input("As-of date", value=dt.date.today())
    include_hidden = st.sidebar.checkbox("Include hidden rows", value=False)

    deck, metadata = load_portfolio_workbook(
        file_bytes=file_bytes,
        as_of_date_iso=as_of_date.isoformat(),
        include_hidden=include_hidden,
    )
    if deck.empty:
        st.error("No usable Bridge / Term rows were found in this workbook.")
        st.stop()

    overrides = ensure_override_store()
    deck = apply_overrides(deck, overrides)
    status_suggestions = available_status_suggestions(deck)

    st.sidebar.header("Filters")
    sheet_choices = st.sidebar.multiselect("Loan type", options=["Bridge", "Term"], default=["Bridge", "Term"])
    search_text = st.sidebar.text_input("Search deal # / name / borrower")

    if sheet_choices:
        filtered = deck[deck["sheet"].isin(sheet_choices)].copy()
    else:
        filtered = deck.iloc[0:0].copy()

    if search_text and not filtered.empty:
        search_upper = search_text.strip().upper()
        filtered = filtered[
            filtered["deal_number"].astype(str).str.upper().str.contains(search_upper, na=False)
            | filtered["deal_name"].astype(str).str.upper().str.contains(search_upper, na=False)
            | filtered["borrower"].astype(str).str.upper().str.contains(search_upper, na=False)
        ].copy()

    filtered = filtered.sort_values(["sheet_order", "original_order"], ascending=[True, True], kind="stable").reset_index(drop=True)

    st.sidebar.header("Agenda")
    if filtered.empty:
        st.sidebar.info("No deals match the current filters.")
    else:
        if "deck_index" not in st.session_state:
            st.session_state.deck_index = 0
        st.session_state.deck_index = int(np.clip(st.session_state.deck_index, 0, len(filtered) - 1))

        sidebar_selection = st.sidebar.selectbox(
            "Jump to agenda item",
            options=list(range(len(filtered))),
            index=st.session_state.deck_index,
            format_func=lambda i: agenda_label(filtered.iloc[i]),
        )
        if sidebar_selection != st.session_state.deck_index:
            st.session_state.deck_index = sidebar_selection
            st.rerun()

        st.sidebar.caption(
            f"Showing {len(filtered)} deals in original workbook order. Bridge remains first, then Term."
        )

    tabs = st.tabs(["Overview", "Presentation Mode", "Meeting Log"])
    with tabs[0]:
        render_overview_tab(filtered, as_of_date, metadata)
    with tabs[1]:
        render_presentation_tab(filtered, status_suggestions)
    with tabs[2]:
        render_meeting_log_tab(deck, filtered, file_bytes, workbook_name)


if __name__ == "__main__":
    main()
