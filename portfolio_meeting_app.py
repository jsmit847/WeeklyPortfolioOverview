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
    "Watchlist",
    "Needs Follow-Up",
    "Resolved",
]


def apply_app_css() -> None:
    st.markdown(
        """
        <style>
            .block-container {
                padding-top: 1.0rem;
                padding-bottom: 2rem;
                max-width: 1460px;
            }
            [data-testid="stSidebar"] {
                border-right: 1px solid rgba(15, 23, 42, 0.08);
            }
            .hero-card {
                border-radius: 24px;
                padding: 1.25rem 1.4rem 1.15rem 1.4rem;
                border: 1px solid rgba(15, 23, 42, 0.08);
                background: linear-gradient(135deg, rgba(14,165,233,0.10), rgba(99,102,241,0.10));
                margin-bottom: 0.9rem;
            }
            .hero-kicker {
                display: inline-block;
                padding: 0.28rem 0.68rem;
                border-radius: 999px;
                font-size: 0.76rem;
                font-weight: 800;
                letter-spacing: 0.04em;
                text-transform: uppercase;
                background: rgba(15, 23, 42, 0.08);
                margin-bottom: 0.65rem;
            }
            .hero-title {
                font-size: 2rem;
                font-weight: 800;
                line-height: 1.1;
                margin-bottom: 0.3rem;
            }
            .hero-subtitle {
                color: rgba(15, 23, 42, 0.75);
                font-size: 1rem;
                margin-bottom: 0.65rem;
            }
            .chip-row {
                margin-top: 0.2rem;
            }
            .chip {
                display: inline-block;
                margin: 0.2rem 0.35rem 0 0;
                padding: 0.35rem 0.65rem;
                border-radius: 999px;
                background: rgba(15, 23, 42, 0.07);
                font-size: 0.83rem;
                font-weight: 700;
            }
            div[data-testid="stMetric"] {
                background: linear-gradient(180deg, rgba(255,255,255,0.95), rgba(248,250,252,0.95));
                border: 1px solid rgba(15, 23, 42, 0.08);
                border-radius: 18px;
                padding: 0.85rem 1rem;
            }
            .panel-card {
                border: 1px solid rgba(15, 23, 42, 0.08);
                border-radius: 20px;
                padding: 1rem 1rem 0.9rem 1rem;
                background: rgba(255,255,255,0.82);
                height: 100%;
            }
            .panel-title {
                font-size: 1rem;
                font-weight: 800;
                margin-bottom: 0.8rem;
            }
            .info-grid {
                display: grid;
                grid-template-columns: repeat(2, minmax(0, 1fr));
                gap: 0.7rem;
            }
            .info-item {
                border: 1px solid rgba(15, 23, 42, 0.07);
                border-radius: 16px;
                padding: 0.75rem 0.85rem;
                background: rgba(248,250,252,0.8);
            }
            .info-label {
                font-size: 0.73rem;
                font-weight: 800;
                letter-spacing: 0.04em;
                text-transform: uppercase;
                color: rgba(15, 23, 42, 0.58);
                margin-bottom: 0.32rem;
            }
            .info-value {
                font-size: 0.96rem;
                font-weight: 700;
                color: rgba(15, 23, 42, 0.95);
                line-height: 1.3;
                word-break: break-word;
            }
            .signal-list {
                margin: 0;
                padding-left: 1.1rem;
            }
            .signal-list li {
                margin-bottom: 0.45rem;
            }
            .small-muted {
                color: rgba(15, 23, 42, 0.72);
                font-size: 0.86rem;
            }
            .agenda-note {
                font-size: 0.86rem;
                color: rgba(15, 23, 42, 0.72);
                margin-top: 0.2rem;
            }
        </style>
        """,
        unsafe_allow_html=True,
    )


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
        npl_col = resolve_column(df.columns, latest_columns.get("NPL"), "NPL", "Loan Level Delinquency", "DQ Status")
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
        df["npl_raw"] = safe_series(df, npl_col, "").fillna("").astype(str)

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
            "npl_header": npl_col or "",
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
    deck["watch_flag"] = (
        deck["npl_raw"].astype(str).str.upper().str.contains(r"\bY\b|90\+|NPL|DQ 90", regex=True, na=False)
        | (deck["days_to_maturity"].fillna(9999) <= 30)
        | (deck["days_past_due"].fillna(0) >= 30)
        | (deck["days_to_next_payment"].fillna(9999) < 0)
    )

    deck = deck.sort_values(["sheet_order", "original_order"], ascending=[True, True], kind="stable").reset_index(drop=True)
    return deck, metadata


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


def available_status_suggestions(deck: pd.DataFrame) -> List[str]:
    if "status" not in deck.columns:
        return COMMON_STATUS_SUGGESTIONS
    status_values = sorted({norm_text(v) for v in deck["status"].dropna().tolist() if norm_text(v)})
    return sorted(set(COMMON_STATUS_SUGGESTIONS + status_values))


def build_flag_messages(row: pd.Series) -> List[str]:
    flags: List[str] = []
    days_to_maturity = row.get("days_to_maturity")
    days_to_next_payment = row.get("days_to_next_payment")
    days_past_due = row.get("days_past_due")
    npl_raw = norm_text(row.get("npl_raw", ""))

    if pd.notna(days_to_maturity):
        if days_to_maturity < 0:
            flags.append(f"Maturity date {fmt_date(row.get('maturity_date'))} has already passed ({fmt_day_delta(days_to_maturity)}).")
        elif days_to_maturity <= 30:
            flags.append(f"Maturity date {fmt_date(row.get('maturity_date'))} is within 30 days ({fmt_day_delta(days_to_maturity)}).")

    if pd.notna(days_to_next_payment):
        if days_to_next_payment < 0:
            flags.append(f"Next payment date {fmt_date(row.get('next_payment_date'))} is overdue ({fmt_day_delta(days_to_next_payment)}).")
        elif days_to_next_payment <= 15:
            flags.append(f"Next payment date {fmt_date(row.get('next_payment_date'))} is approaching soon ({fmt_day_delta(days_to_next_payment)}).")

    if pd.notna(days_past_due) and float(days_past_due) > 0:
        flags.append(f"Days past due is currently {fmt_int(days_past_due)}.")

    if npl_raw and canon_header(npl_raw) not in {"", "N", "NO"}:
        flags.append(f"Watchlist / NPL signal present: {npl_raw}.")

    if not flags:
        flags.append("No immediate system flags. Use the commentary box for live qualitative updates.")

    return flags


def render_metric_row(row: pd.Series) -> None:
    cols = st.columns(6)
    cols[0].metric("UPB", fmt_money(row.get("upb"), decimals=0))
    cols[1].metric("Maturity", fmt_date(row.get("maturity_date")), fmt_day_delta(row.get("days_to_maturity")))
    cols[2].metric("Next Payment", fmt_date(row.get("next_payment_date")), fmt_day_delta(row.get("days_to_next_payment")))
    cols[3].metric("Days Past Due", fmt_int(row.get("days_past_due")))
    cols[4].metric("Owner", display_text(row.get("owner")))
    cols[5].metric("Watchlist", "Yes" if bool(row.get("watch_flag")) else "No")


def render_info_grid(items: List[Tuple[str, object]]) -> None:
    cards = []
    for label, value in items:
        cards.append(
            "<div class='info-item'>"
            f"<div class='info-label'>{html.escape(label)}</div>"
            f"<div class='info-value'>{html.escape(display_text(value))}</div>"
            "</div>"
        )
    st.markdown(f"<div class='info-grid'>{''.join(cards)}</div>", unsafe_allow_html=True)


def build_queue_dataframe(deck: pd.DataFrame) -> pd.DataFrame:
    queue = deck.copy()

    needed_defaults = {
        "sheet": "",
        "deal_number": "",
        "deal_name": "",
        "status": "",
        "owner": "",
        "upb": np.nan,
        "maturity_date": pd.NaT,
        "next_payment_date": pd.NaT,
        "days_past_due": np.nan,
        "watch_flag": False,
        "saved_at": "",
        "sheet_order": 99,
        "original_order": 999999,
    }
    for col, default_value in needed_defaults.items():
        if col not in queue.columns:
            queue[col] = default_value

    queue = queue.sort_values(["sheet_order", "original_order"], ascending=[True, True], kind="stable").reset_index(drop=True)
    queue.insert(0, "Agenda #", range(1, len(queue) + 1))

    queue = queue[
        [
            "Agenda #",
            "sheet",
            "deal_number",
            "deal_name",
            "status",
            "owner",
            "upb",
            "maturity_date",
            "next_payment_date",
            "days_past_due",
            "watch_flag",
            "saved_at",
        ]
    ].copy()

    queue = queue.rename(
        columns={
            "sheet": "Type",
            "deal_number": "Deal #",
            "deal_name": "Deal Name",
            "status": "Status",
            "owner": "Owner",
            "upb": "UPB",
            "maturity_date": "Maturity",
            "next_payment_date": "Next Payment",
            "days_past_due": "DPD",
            "watch_flag": "Watch",
            "saved_at": "Last Saved",
        }
    )

    queue["UPB"] = queue["UPB"].map(lambda x: fmt_money(x, decimals=0))
    queue["Maturity"] = queue["Maturity"].map(fmt_date)
    queue["Next Payment"] = queue["Next Payment"].map(fmt_date)
    queue["DPD"] = queue["DPD"].map(fmt_int)
    queue["Watch"] = queue["Watch"].map(lambda x: "Yes" if bool(x) else "No")
    return queue


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

    output = io.BytesIO()
    workbook.save(output)
    output.seek(0)
    return output.getvalue()


def render_overview_tab(deck: pd.DataFrame, as_of_date: dt.date, metadata: Dict[str, Dict[str, str]]) -> None:
    if deck.empty:
        st.warning("No deals match the current filters.")
        return

    total_upb = deck["upb"].fillna(0).sum()
    bridge_count = int((deck["sheet"] == "Bridge").sum())
    term_count = int((deck["sheet"] == "Term").sum())
    near_maturity = int((deck["days_to_maturity"].fillna(9999) <= 30).sum())
    overdue_payments = int((deck["days_to_next_payment"].fillna(9999) < 0).sum())

    m1, m2, m3, m4, m5 = st.columns(5)
    m1.metric("Bridge Deals", fmt_int(bridge_count))
    m2.metric("Term Deals", fmt_int(term_count))
    m3.metric("Total UPB", fmt_money(total_upb, decimals=0))
    m4.metric("Maturing in 30d", fmt_int(near_maturity))
    m5.metric("Past Due Next Pay", fmt_int(overdue_payments))

    st.caption(
        f"As of {as_of_date.strftime('%m/%d/%Y')} • "
        + " • ".join(
            f"{sheet}: {meta.get('upb_header', 'UPB') or 'UPB'}"
            + (f", {meta.get('maturity_header') or 'Maturity Date'}" if meta.get("maturity_header") else "")
            for sheet, meta in metadata.items()
        )
    )

    left, right = st.columns([1.35, 1.0])

    with left:
        st.subheader("Meeting agenda")
        st.caption("Presented in the same order the deals appear in the workbook.")
        st.dataframe(build_queue_dataframe(deck), use_container_width=True, hide_index=True)

    with right:
        st.subheader("Upcoming dates")
        upcoming = deck[
            ["sheet", "deal_number", "deal_name", "maturity_date", "days_to_maturity", "next_payment_date", "days_to_next_payment"]
        ].copy()
        upcoming["Maturity Date"] = upcoming["maturity_date"].map(fmt_date)
        upcoming["Maturity Timing"] = upcoming["days_to_maturity"].map(fmt_day_delta)
        upcoming["Next Payment Date"] = upcoming["next_payment_date"].map(fmt_date)
        upcoming["Payment Timing"] = upcoming["days_to_next_payment"].map(fmt_day_delta)
        upcoming = upcoming.sort_values(["maturity_date", "next_payment_date"], ascending=[True, True], kind="stable").head(12)
        st.dataframe(
            upcoming[
                [
                    "sheet",
                    "deal_number",
                    "deal_name",
                    "Maturity Date",
                    "Maturity Timing",
                    "Next Payment Date",
                    "Payment Timing",
                ]
            ].rename(
                columns={
                    "sheet": "Type",
                    "deal_number": "Deal #",
                    "deal_name": "Deal Name",
                }
            ),
            use_container_width=True,
            hide_index=True,
        )


def render_presentation_tab(deck: pd.DataFrame, status_suggestions: List[str]) -> None:
    if deck.empty:
        st.warning("No deals match the current filters.")
        return

    if "deck_index" not in st.session_state:
        st.session_state.deck_index = 0
    st.session_state.deck_index = int(np.clip(st.session_state.deck_index, 0, len(deck) - 1))

    nav_left, nav_mid, nav_right = st.columns([1, 1, 3.2])
    with nav_left:
        if st.button("⬅ Previous", use_container_width=True, disabled=st.session_state.deck_index == 0):
            st.session_state.deck_index -= 1
    with nav_mid:
        if st.button("Next ➡", use_container_width=True, disabled=st.session_state.deck_index >= len(deck) - 1):
            st.session_state.deck_index += 1
    with nav_right:
        current_idx = st.selectbox(
            "Jump to deal",
            options=list(range(len(deck))),
            index=st.session_state.deck_index,
            format_func=lambda i: f"{i + 1}. {deck.iloc[i]['sheet']} | {deck.iloc[i]['deal_name']} ({deck.iloc[i]['deal_number']})",
        )
        st.session_state.deck_index = current_idx

    row = deck.iloc[st.session_state.deck_index]
    progress = (st.session_state.deck_index + 1) / len(deck)
    st.progress(progress, text=f"Agenda item {st.session_state.deck_index + 1} of {len(deck)}")

    hero_html = f"""
        <div class='hero-card'>
            <div class='hero-kicker'>{html.escape(str(row.get('sheet', '')))} • Agenda Item {st.session_state.deck_index + 1}</div>
            <div class='hero-title'>{html.escape(display_text(row.get('deal_name')))}</div>
            <div class='hero-subtitle'>
                Deal {html.escape(display_text(row.get('deal_number')))} • Borrower {html.escape(display_text(row.get('borrower')))} • Status {html.escape(display_text(row.get('status')))}
            </div>
            <div class='chip-row'>
                <span class='chip'>Servicer: {html.escape(display_text(row.get('servicer')))}</span>
                <span class='chip'>Owner: {html.escape(display_text(row.get('owner')))}</span>
                <span class='chip'>Portfolio: {html.escape(display_text(row.get('portfolio')))}</span>
                <span class='chip'>Watchlist: {'Yes' if bool(row.get('watch_flag')) else 'No'}</span>
            </div>
        </div>
    """
    st.markdown(hero_html, unsafe_allow_html=True)
    render_metric_row(row)

    left, middle, right = st.columns([1.05, 1.0, 1.05])

    with left:
        st.markdown("<div class='panel-card'><div class='panel-title'>Deal snapshot</div>", unsafe_allow_html=True)
        snapshot_items = [
            ("Deal Number", row.get("deal_number")),
            ("Borrower", row.get("borrower")),
            ("Servicer", row.get("servicer")),
            ("Portfolio", row.get("portfolio")),
            ("Segment", row.get("segment")),
            ("Financing", row.get("financing")),
            ("Loan Buyer", row.get("loan_buyer")),
            ("Account", row.get("account_display")),
        ]
        render_info_grid(snapshot_items)
        st.markdown("</div>", unsafe_allow_html=True)

    with middle:
        st.markdown("<div class='panel-card'><div class='panel-title'>Key dates and signals</div>", unsafe_allow_html=True)
        timeline_items = [
            ("Maturity Date", f"{fmt_date(row.get('maturity_date'))} • {fmt_day_delta(row.get('days_to_maturity'))}"),
            ("Next Payment Date", f"{fmt_date(row.get('next_payment_date'))} • {fmt_day_delta(row.get('days_to_next_payment'))}"),
            ("Days Past Due", fmt_int(row.get("days_past_due"))),
            ("Status", row.get("status")),
        ]
        if row.get("sheet") == "Bridge":
            timeline_items.extend(
                [
                    ("Commitment", fmt_money(row.get("commitment"), decimals=0)),
                    ("Funded Amount", fmt_money(row.get("funded_amount"), decimals=0)),
                    ("Remaining Commitment", fmt_money(row.get("remaining_commitment"), decimals=0)),
                    ("UPB", fmt_money(row.get("upb"), decimals=0)),
                ]
            )
        else:
            timeline_items.extend(
                [
                    ("Loan Amount", fmt_money(row.get("loan_amount"), decimals=0)),
                    ("UPB", fmt_money(row.get("upb"), decimals=0)),
                ]
            )
        render_info_grid(timeline_items)
        signals = build_flag_messages(row)
        signal_html = "".join(f"<li>{html.escape(signal)}</li>" for signal in signals)
        st.markdown(
            f"""
            <div style='margin-top:0.9rem;'>
                <div class='panel-title' style='margin-bottom:0.45rem;'>Talking points</div>
                <ul class='signal-list'>{signal_html}</ul>
                <div class='small-muted' style='margin-top:0.7rem;'>
                    Signals are based on dates, delinquency fields, and existing watchlist / NPL indicators. Presentation order stays true to the workbook.
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        st.markdown("</div>", unsafe_allow_html=True)

    with right:
        st.markdown("<div class='panel-card'><div class='panel-title'>Live meeting updates</div>", unsafe_allow_html=True)
        st.caption("Update these during the meeting. CSV and updated workbook downloads are on the Meeting Log tab.")
        if status_suggestions:
            st.caption("Quick status ideas: " + " | ".join(status_suggestions[:8]))

        status_value = st.text_input(
            "Status",
            value=str(row.get("status", "")),
            key=f"status_input::{row['sheet']}::{row['deal_number']}",
        )
        owner_label = "Point Person" if row.get("sheet") == "Bridge" else "Asset Manager"
        owner_value = st.text_input(
            owner_label,
            value=str(row.get("owner", "")),
            key=f"owner_input::{row['sheet']}::{row['deal_number']}",
        )
        commentary_value = st.text_area(
            "AM Commentary / Meeting Notes",
            value=str(row.get("commentary", "")),
            height=260,
            key=f"commentary_input::{row['sheet']}::{row['deal_number']}",
        )

        if st.button("Save this update", use_container_width=True, type="primary"):
            overrides = ensure_override_store()
            override_key = get_override_key(str(row.get("sheet")), str(row.get("deal_number")))
            overrides[override_key] = {
                "sheet": str(row.get("sheet")),
                "deal_number": str(row.get("deal_number")),
                "status": status_value,
                "owner": owner_value,
                "commentary": commentary_value,
                "saved_at": dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            }
            st.toast("Saved to the meeting update log.")

        st.markdown("</div>", unsafe_allow_html=True)


def render_meeting_log_tab(displayed_deck: pd.DataFrame, file_bytes: bytes, workbook_name: str) -> None:
    overrides = ensure_override_store()
    st.metric("Saved meeting updates", fmt_int(len(overrides)))

    if overrides:
        updates_df = pd.DataFrame(overrides.values()).sort_values(["sheet", "deal_number"])
        st.dataframe(updates_df, use_container_width=True, hide_index=True)
    else:
        st.info("No live updates saved yet. Use Presentation Mode to update status, owner, or commentary during the meeting.")

    export_left, export_right = st.columns(2)
    with export_left:
        csv_bytes = export_overrides_csv(overrides)
        st.download_button(
            "Download meeting updates CSV",
            data=csv_bytes,
            file_name=f"{Path(workbook_name).stem}_meeting_updates.csv",
            mime="text/csv",
            disabled=not bool(overrides),
            use_container_width=True,
        )
    with export_right:
        workbook_bytes = update_workbook_bytes(file_bytes, overrides)
        st.download_button(
            "Download updated workbook",
            data=workbook_bytes,
            file_name=f"{Path(workbook_name).stem}_meeting_ready.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    st.subheader("Current meeting queue")
    st.dataframe(build_queue_dataframe(displayed_deck), use_container_width=True, hide_index=True)


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


def render_sidebar_agenda(filtered: pd.DataFrame) -> None:
    st.sidebar.markdown("---")
    st.sidebar.subheader("Agenda")
    if filtered.empty:
        st.sidebar.caption("No deals match the current filters.")
        return

    if "deck_index" not in st.session_state:
        st.session_state.deck_index = 0
    st.session_state.deck_index = int(np.clip(st.session_state.deck_index, 0, len(filtered) - 1))

    selected_index = st.sidebar.selectbox(
        "Select a deal",
        options=list(range(len(filtered))),
        index=st.session_state.deck_index,
        format_func=lambda i: f"{i + 1}. {filtered.iloc[i]['sheet']} | {filtered.iloc[i]['deal_name']} ({filtered.iloc[i]['deal_number']})",
    )
    st.session_state.deck_index = selected_index

    bridge_count = int((filtered["sheet"] == "Bridge").sum())
    term_count = int((filtered["sheet"] == "Term").sum())
    st.sidebar.markdown(
        f"<div class='agenda-note'>Bridge: {bridge_count} deals<br>Term: {term_count} deals</div>",
        unsafe_allow_html=True,
    )


def main() -> None:
    st.set_page_config(page_title=APP_TITLE, layout="wide")
    apply_app_css()

    st.title(APP_TITLE)
    st.caption("Cleaner meeting view with Bridge first, Term second, and presentation order preserved from the workbook.")

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

    st.sidebar.markdown("---")
    st.sidebar.header("Filters")
    sheet_choices = st.sidebar.multiselect("Loan type", options=["Bridge", "Term"], default=["Bridge", "Term"])
    watch_only = st.sidebar.checkbox("Watchlist only", value=False)
    search_text = st.sidebar.text_input("Search deal # / name / borrower")

    if not sheet_choices:
        filtered = deck.iloc[0:0].copy()
    else:
        filtered = deck[deck["sheet"].isin(sheet_choices)].copy()

    if watch_only and not filtered.empty:
        filtered = filtered[filtered["watch_flag"] == True].copy()

    if search_text and not filtered.empty:
        search_upper = search_text.strip().upper()
        filtered = filtered[
            filtered["deal_number"].astype(str).str.upper().str.contains(search_upper, na=False)
            | filtered["deal_name"].astype(str).str.upper().str.contains(search_upper, na=False)
            | filtered["borrower"].astype(str).str.upper().str.contains(search_upper, na=False)
        ].copy()

    filtered = filtered.sort_values(["sheet_order", "original_order"], ascending=[True, True], kind="stable").reset_index(drop=True)
    render_sidebar_agenda(filtered)

    if filtered.empty:
        st.warning("No deals match the current filters.")
        st.stop()

    if "deck_index" not in st.session_state:
        st.session_state.deck_index = 0
    st.session_state.deck_index = int(np.clip(st.session_state.deck_index, 0, len(filtered) - 1))

    tabs = st.tabs(["Overview", "Presentation Mode", "Meeting Log"])
    with tabs[0]:
        render_overview_tab(filtered, as_of_date, metadata)
    with tabs[1]:
        render_presentation_tab(filtered, status_suggestions)
    with tabs[2]:
        render_meeting_log_tab(filtered, file_bytes, workbook_name)


if __name__ == "__main__":
    main()
