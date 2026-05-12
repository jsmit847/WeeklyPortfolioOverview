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


def apply_app_css() -> None:
    st.markdown(
        """
        <style>
            .block-container {
                max-width: 1520px;
                padding-top: 0.45rem;
                padding-bottom: 1rem;
            }
            .topline {
                display: flex;
                align-items: baseline;
                gap: 0.7rem;
                margin-bottom: 0.15rem;
            }
            .topline-title {
                font-size: 1.15rem;
                font-weight: 800;
                color: rgb(15,23,42);
                line-height: 1.05;
            }
            .topline-subtitle {
                font-size: 0.82rem;
                color: rgba(15,23,42,0.66);
                line-height: 1.1;
            }
            .toolbar-card {
                border: 1px solid rgba(15,23,42,0.08);
                border-radius: 18px;
                padding: 0.6rem 0.8rem;
                background: linear-gradient(180deg, rgba(255,255,255,0.95), rgba(248,250,252,0.95));
                margin-bottom: 0.6rem;
            }
            .hero-card {
                border: 1px solid rgba(15,23,42,0.08);
                border-radius: 22px;
                padding: 0.78rem 0.95rem;
                background:
                    radial-gradient(circle at top right, rgba(14,165,233,0.10), transparent 30%),
                    linear-gradient(135deg, rgba(255,255,255,0.98), rgba(248,250,252,0.98));
                margin-bottom: 0.55rem;
            }
            .hero-kicker {
                display: inline-block;
                border-radius: 999px;
                background: rgba(15,23,42,0.08);
                padding: 0.18rem 0.52rem;
                font-size: 0.68rem;
                font-weight: 800;
                letter-spacing: 0.04em;
                text-transform: uppercase;
                margin-bottom: 0.35rem;
            }
            .hero-title {
                font-size: 1.28rem;
                line-height: 1.1;
                font-weight: 800;
                margin-bottom: 0.18rem;
                color: rgb(15,23,42);
            }
            .hero-subtitle {
                font-size: 0.84rem;
                color: rgba(15,23,42,0.7);
                margin-bottom: 0.35rem;
            }
            .chip {
                display: inline-block;
                margin: 0.1rem 0.25rem 0 0;
                border-radius: 999px;
                border: 1px solid rgba(15,23,42,0.06);
                background: rgba(255,255,255,0.92);
                padding: 0.22rem 0.48rem;
                font-size: 0.74rem;
                font-weight: 700;
            }
            div[data-testid="stMetric"] {
                border: 1px solid rgba(15,23,42,0.08);
                border-radius: 18px;
                background: linear-gradient(180deg, rgba(255,255,255,0.98), rgba(248,250,252,0.98));
                padding: 0.52rem 0.68rem;
            }
            [data-testid="stMetricLabel"] {
                font-size: 0.76rem;
            }
            [data-testid="stMetricValue"] {
                font-size: 1.26rem;
                line-height: 1.02;
            }
            [data-testid="stMetricDelta"] {
                font-size: 0.78rem;
            }
            .panel-card {
                border: 1px solid rgba(15,23,42,0.08);
                border-radius: 20px;
                padding: 0.82rem 0.9rem;
                background: rgba(255,255,255,0.92);
                height: 100%;
            }
            .panel-title {
                font-size: 0.93rem;
                font-weight: 800;
                color: rgb(15,23,42);
                margin-bottom: 0.55rem;
            }
            .field-grid {
                display: grid;
                grid-template-columns: repeat(2, minmax(0, 1fr));
                gap: 0.5rem;
            }
            .field-card {
                border: 1px solid rgba(15,23,42,0.07);
                border-radius: 14px;
                background: rgba(248,250,252,0.8);
                padding: 0.55rem 0.62rem;
            }
            .field-label {
                font-size: 0.67rem;
                font-weight: 800;
                letter-spacing: 0.04em;
                text-transform: uppercase;
                color: rgba(15,23,42,0.56);
                margin-bottom: 0.18rem;
            }
            .field-value {
                font-size: 0.86rem;
                font-weight: 700;
                color: rgba(15,23,42,0.95);
                line-height: 1.18;
                word-break: break-word;
            }
            .notes-box {
                min-height: 145px;
                border: 1px solid rgba(15,23,42,0.07);
                border-radius: 14px;
                background: rgba(248,250,252,0.8);
                padding: 0.65rem 0.7rem;
                font-size: 0.9rem;
                line-height: 1.32;
                color: rgba(15,23,42,0.92);
            }
            .small-muted {
                font-size: 0.8rem;
                color: rgba(15,23,42,0.66);
            }
            .mini-chip {
                display: inline-block;
                margin-right: 0.3rem;
                border-radius: 999px;
                background: rgba(15,23,42,0.07);
                padding: 0.18rem 0.5rem;
                font-size: 0.72rem;
                font-weight: 700;
            }
            .agenda-note {
                font-size: 0.8rem;
                color: rgba(15,23,42,0.67);
                margin-bottom: 0.35rem;
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


def get_override_key(sheet: str, deal_number: str) -> str:
    return f"{sheet}::{deal_number}"


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
    return [str(c) for c in columns if c is not None and canon_header(c) == canon_header(candidate)]


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
        df["deal_key"] = df.apply(lambda row: get_override_key(str(row["sheet"]), str(row["deal_number"])), axis=1)

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
        "pending_only": False,
        "as_of_date": dt.date.today(),
        "selected_deal_key": None,
        "meeting_overrides": {},
        "review_flags": {},
        "dialog_target": None,
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
        "commentary": commentary.strip(),
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

    if st.session_state.pending_only:
        flags = ensure_review_flags()
        filtered = filtered[~filtered["deal_key"].map(lambda key: bool(flags.get(key, False)))].copy()

    return filtered.sort_values(["sheet_order", "original_order"], ascending=[True, True], kind="stable").reset_index(drop=True)


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

    maturity_text = fmt_date(row.get("maturity_date"))
    payment_text = fmt_date(row.get("next_payment_date"))
    if maturity_text != "-":
        prompts.append(f"Confirm maturity timing: {maturity_text} ({fmt_day_delta(row.get('days_to_maturity'))}).")
    if payment_text != "-":
        prompts.append(f"Confirm next payment timing: {payment_text} ({fmt_day_delta(row.get('days_to_next_payment'))}).")

    status_text = display_text(row.get("status"), blank="")
    if status_text:
        prompts.append(f"Check whether status should stay as '{status_text}'.")

    commentary_text = display_text(row.get("commentary"), blank="")
    if commentary_text:
        prompts.append("Decide whether AM commentary needs a live refresh before the meeting ends.")

    if not prompts:
        prompts.append("Confirm the latest story, next milestone, and any follow-up owner.")

    return prompts[:4]


def render_field_grid(items: List[Tuple[str, object]]) -> None:
    cards = []
    for label, value in items:
        cards.append(
            "<div class='field-card'>"
            f"<div class='field-label'>{html.escape(label)}</div>"
            f"<div class='field-value'>{html.escape(display_text(value))}</div>"
            "</div>"
        )
    st.markdown(f"<div class='field-grid'>{''.join(cards)}</div>", unsafe_allow_html=True)


def build_agenda_dataframe(deck: pd.DataFrame) -> pd.DataFrame:
    queue = deck.copy()
    queue.insert(0, "Agenda #", range(1, len(queue) + 1))
    queue["UPB"] = queue["upb"].map(lambda x: fmt_money(x, decimals=0))
    queue["Maturity"] = queue["maturity_date"].map(fmt_date)
    queue["Next Payment"] = queue["next_payment_date"].map(fmt_date)
    queue["Status"] = queue["status"].map(display_text)
    queue["Owner"] = queue["owner"].map(display_text)
    return queue[
        [
            "Agenda #",
            "sheet",
            "deal_name",
            "deal_number",
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
    flags = ensure_review_flags()
    review = deck.copy().set_index("deal_key")
    review.insert(0, "Agenda #", range(1, len(review) + 1))
    review["Type"] = review["sheet"]
    review["Deal Name"] = review["deal_name"].map(display_text)
    review["Deal #"] = review["deal_number"].map(display_text)
    review["Maturity"] = review["maturity_date"].map(fmt_date)
    review["Next Payment"] = review["next_payment_date"].map(fmt_date)
    review["UPB"] = review["upb"].map(lambda x: fmt_money(x, decimals=0))
    review["Reviewed"] = review.index.map(lambda key: bool(flags.get(key, False)))
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
            "Reviewed",
            "Status",
            "Owner",
            "AM Commentary",
        ]
    ]


def apply_review_edits(edited_df: pd.DataFrame, current_deck: pd.DataFrame, raw_deck: pd.DataFrame) -> None:
    current_lookup = current_deck.set_index("deal_key")
    raw_lookup = raw_deck.set_index("deal_key")
    flags = ensure_review_flags()

    for deal_key, edited_row in edited_df.iterrows():
        reviewed_value = bool(edited_row.get("Reviewed", False))
        flags[deal_key] = reviewed_value

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
    flags = ensure_review_flags()
    overrides = ensure_override_store()

    reviewed_df = pd.DataFrame(
        {
            "Deal Key": list(flags.keys()),
            "Reviewed": list(flags.values()),
        }
    )
    reviewed_df = reviewed_df[reviewed_df["Reviewed"] == True]  # noqa: E712

    if reviewed_df.empty and not overrides:
        return pd.DataFrame()

    out = deck[["deal_key", "sheet", "deal_name", "deal_number", "status", "owner", "commentary", "saved_at"]].copy()
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
        }
    )
    out["Reviewed"] = out["Deal Key"].map(lambda key: bool(flags.get(key, False)))

    if overrides:
        changed_keys = set(get_override_key(item["sheet"], item["deal_number"]) for item in overrides.values())
        out = out[(out["Reviewed"] == True) | (out["Deal Key"].isin(changed_keys))].copy()  # noqa: E712
    else:
        out = out[out["Reviewed"] == True].copy()  # noqa: E712

    return out[["Type", "Deal Name", "Deal #", "Reviewed", "Status", "Owner", "AM Commentary", "Last Saved"]]


def export_overrides_csv(raw_deck: pd.DataFrame) -> bytes:
    export_df = build_updates_dataframe(raw_deck, apply_overrides(raw_deck, ensure_override_store()))
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


def render_controls_bar(workbook_name: str) -> None:
    left, middle, right = st.columns([1.5, 1.6, 1.1], vertical_alignment="center")

    with left:
        st.markdown(
            "<div class='topline'><div class='topline-title'>Weekly Portfolio Meeting Deck</div>"
            "<div class='topline-subtitle'>Compact review flow built around original workbook order.</div></div>",
            unsafe_allow_html=True,
        )
        st.caption(f"Workbook: {workbook_name}")

    with middle:
        st.segmented_control(
            "View",
            options=["Overview", "Presentation", "Review", "Exports"],
            key="view_mode",
            label_visibility="collapsed",
            width="stretch",
        )

    with right:
        with st.popover("Controls", icon=":material/tune:", width="stretch"):
            st.file_uploader("Upload workbook", type=["xlsx"], key="uploaded_workbook")
            st.date_input("As-of date", key="as_of_date")
            st.toggle("Include hidden rows", key="include_hidden")
            st.segmented_control(
                "Sheet focus",
                options=["All", "Bridge", "Term"],
                key="sheet_filter",
                width="stretch",
            )
            st.text_input("Search deal # / name / borrower", key="search_query")
            st.toggle("Pending review only", key="pending_only")
            if st.button("Reset filters", use_container_width=True):
                st.session_state.sheet_filter = "All"
                st.session_state.search_query = ""
                st.session_state.pending_only = False
                st.rerun()


def render_overview_view(deck: pd.DataFrame, as_of_date: dt.date) -> None:
    total_upb = deck["upb"].fillna(0).sum()
    bridge_count = int((deck["sheet"] == "Bridge").sum())
    term_count = int((deck["sheet"] == "Term").sum())
    next_30 = int((deck["days_to_maturity"].fillna(9999) <= 30).sum())

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Agenda items", fmt_int(len(deck)))
    m2.metric("Total UPB", fmt_money(total_upb, decimals=0))
    m3.metric("Bridge / Term", f"{fmt_int(bridge_count)} / {fmt_int(term_count)}")
    m4.metric("Maturing in 30d", fmt_int(next_30), as_of_date.strftime("%m/%d/%Y"))

    st.markdown(
        "<div class='agenda-note'>Overview stays in original workbook order. Use Review for controlled editing and Presentation for full-screen discussion.</div>",
        unsafe_allow_html=True,
    )

    left, right = st.columns([1.15, 0.85], gap="large")

    with left:
        st.subheader("Agenda queue")
        st.dataframe(build_agenda_dataframe(deck), hide_index=True, use_container_width=True)

    with right:
        st.subheader("Upcoming timing")
        timing = deck[["sheet", "deal_name", "maturity_date", "days_to_maturity", "next_payment_date", "days_to_next_payment"]].copy()
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

    @st.dialog("Update current deal", width="large")
    def edit_dialog(current_row_dict: Dict[str, object], raw_row_dict: Dict[str, object]) -> None:
        current_row_local = pd.Series(current_row_dict)
        raw_row_local = pd.Series(raw_row_dict)
        flags = ensure_review_flags()
        deal_key = str(current_row_local.get("deal_key"))

        st.caption(
            f"{display_text(current_row_local.get('deal_name'))} • Deal {display_text(current_row_local.get('deal_number'))}"
        )

        with st.form(f"edit-form::{deal_key}"):
            reviewed = st.checkbox("Mark as reviewed", value=bool(flags.get(deal_key, False)))
            c1, c2 = st.columns(2)
            with c1:
                status = st.text_input("Status", value=str(current_row_local.get("status") or ""))
            owner_label = "Point Person" if current_row_local.get("sheet") == "Bridge" else "Asset Manager"
            with c2:
                owner = st.text_input(owner_label, value=str(current_row_local.get("owner") or ""))
            commentary = st.text_area(
                "AM Commentary",
                value=str(current_row_local.get("commentary") or ""),
                height=220,
            )
            save_clicked = st.form_submit_button("Save changes", type="primary")

        c_left, c_right = st.columns(2)
        with c_left:
            if save_clicked:
                flags[deal_key] = reviewed
                upsert_override(raw_row_local, status, owner, commentary)
                st.session_state.dialog_target = None
                st.rerun()
        with c_right:
            if st.button("Cancel", use_container_width=True):
                st.session_state.dialog_target = None
                st.rerun()

    edit_dialog(current_row.to_dict(), raw_row.to_dict())


def render_presentation_view(deck: pd.DataFrame, raw_deck: pd.DataFrame) -> None:
    sync_selected_deal(deck)
    current_row = get_selected_row(deck)
    current_idx = deck.index[deck["deal_key"] == current_row["deal_key"]][0]
    agenda_label = f"Agenda item {current_idx + 1} of {len(deck)}"

    nav_left, nav_mid, nav_right, nav_more = st.columns([0.8, 0.8, 1.4, 1.1], vertical_alignment="center")
    with nav_left:
        st.button(
            "Previous",
            use_container_width=True,
            disabled=current_idx == 0,
            on_click=move_selection,
            args=(deck, -1),
        )
    with nav_mid:
        st.button(
            "Next",
            use_container_width=True,
            disabled=current_idx >= len(deck) - 1,
            on_click=move_selection,
            args=(deck, 1),
        )
    with nav_right:
        with st.popover("Agenda jump", icon=":material/list:", width="stretch"):
            option_labels = [
                f"{i + 1}. {row.sheet} | {row.deal_name}"
                for i, row in enumerate(deck[["sheet", "deal_name"]].itertuples(index=False, name="DealRow"))
            ]
            selected_label = st.selectbox("Jump to", options=option_labels, index=current_idx)
            if st.button("Go to selected deal", use_container_width=True):
                new_idx = option_labels.index(selected_label)
                st.session_state.selected_deal_key = deck.iloc[new_idx]["deal_key"]
                st.rerun()
    with nav_more:
        st.markdown(
            f"<div class='small-muted' style='text-align:right;padding-top:0.35rem;'>{html.escape(agenda_label)}</div>",
            unsafe_allow_html=True,
        )

    hero_html = (
        "<div class='hero-card'>"
        f"<div class='hero-kicker'>{html.escape(display_text(current_row.get('sheet')))}</div>"
        f"<div class='hero-title'>{html.escape(display_text(current_row.get('deal_name')))}</div>"
        f"<div class='hero-subtitle'>Deal {html.escape(display_text(current_row.get('deal_number')))} • "
        f"{html.escape(display_text(current_row.get('borrower')))} • {html.escape(display_text(current_row.get('status')))}</div>"
        f"<span class='chip'>Servicer: {html.escape(display_text(current_row.get('servicer')))}</span>"
        f"<span class='chip'>Owner: {html.escape(display_text(current_row.get('owner')))}</span>"
        f"<span class='chip'>Portfolio: {html.escape(display_text(current_row.get('portfolio')))}</span>"
        f"<span class='chip'>Segment: {html.escape(display_text(current_row.get('segment')))}</span>"
        "</div>"
    )
    st.markdown(hero_html, unsafe_allow_html=True)

    metric_cols = st.columns(5)
    metric_cols[0].metric("UPB", fmt_money(current_row.get("upb"), decimals=0))
    metric_cols[1].metric("Maturity", fmt_date(current_row.get("maturity_date")), fmt_day_delta(current_row.get("days_to_maturity")))
    metric_cols[2].metric(
        "Next Payment",
        fmt_date(current_row.get("next_payment_date")),
        fmt_day_delta(current_row.get("days_to_next_payment")),
    )
    metric_cols[3].metric("Owner", display_text(current_row.get("owner")))
    metric_cols[4].metric("Status", display_text(current_row.get("status")))

    action_left, action_mid, action_right = st.columns([0.9, 1.0, 2.1], vertical_alignment="center")
    with action_left:
        if st.button("Edit current deal", use_container_width=True, type="primary"):
            st.session_state.dialog_target = current_row["deal_key"]
            st.rerun()
    with action_mid:
        with st.popover("Presenter prompts", icon=":material/lightbulb:", width="stretch"):
            for prompt in build_presenter_prompts(current_row):
                st.write(f"- {prompt}")
    with action_right:
        with st.popover("More deal details", icon=":material/info:", width="stretch"):
            st.markdown("**Timing**")
            st.write(f"Days Past Due: {fmt_int(current_row.get('days_past_due'))}")
            st.write(f"Maturity: {fmt_date(current_row.get('maturity_date'))} ({fmt_day_delta(current_row.get('days_to_maturity'))})")
            st.write(f"Next Payment: {fmt_date(current_row.get('next_payment_date'))} ({fmt_day_delta(current_row.get('days_to_next_payment'))})")
            st.markdown("**Additional fields**")
            st.write(f"Account: {display_text(current_row.get('account_display'))}")
            st.write(f"Financing: {display_text(current_row.get('financing'))}")
            st.write(f"Loan Buyer: {display_text(current_row.get('loan_buyer'))}")
            if current_row.get("sheet") == "Bridge":
                st.write(f"Commitment: {fmt_money(current_row.get('commitment'), decimals=0)}")
                st.write(f"Funded Amount: {fmt_money(current_row.get('funded_amount'), decimals=0)}")
                st.write(f"Remaining Commitment: {fmt_money(current_row.get('remaining_commitment'), decimals=0)}")
            else:
                st.write(f"Loan Amount: {fmt_money(current_row.get('loan_amount'), decimals=0)}")

    left, middle, right = st.columns([1.0, 1.0, 1.1], gap="large")

    with left:
        st.markdown("<div class='panel-card'><div class='panel-title'>Deal snapshot</div>", unsafe_allow_html=True)
        render_field_grid(
            [
                ("Borrower", current_row.get("borrower")),
                ("Account", current_row.get("account_display")),
                ("Servicer", current_row.get("servicer")),
                ("Portfolio", current_row.get("portfolio")),
                ("Segment", current_row.get("segment")),
                ("Deal #", current_row.get("deal_number")),
            ]
        )
        st.markdown("</div>", unsafe_allow_html=True)

    with middle:
        st.markdown("<div class='panel-card'><div class='panel-title'>Capital profile</div>", unsafe_allow_html=True)
        if current_row.get("sheet") == "Bridge":
            render_field_grid(
                [
                    ("UPB", fmt_money(current_row.get("upb"), decimals=0)),
                    ("Commitment", fmt_money(current_row.get("commitment"), decimals=0)),
                    ("Funded Amount", fmt_money(current_row.get("funded_amount"), decimals=0)),
                    ("Remaining Commitment", fmt_money(current_row.get("remaining_commitment"), decimals=0)),
                    ("Maturity Date", fmt_date(current_row.get("maturity_date"))),
                    ("Next Payment", fmt_date(current_row.get("next_payment_date"))),
                ]
            )
        else:
            render_field_grid(
                [
                    ("UPB", fmt_money(current_row.get("upb"), decimals=0)),
                    ("Loan Amount", fmt_money(current_row.get("loan_amount"), decimals=0)),
                    ("Maturity Date", fmt_date(current_row.get("maturity_date"))),
                    ("Next Payment", fmt_date(current_row.get("next_payment_date"))),
                    ("Owner", current_row.get("owner")),
                    ("Financing", current_row.get("financing")),
                ]
            )
        st.markdown("</div>", unsafe_allow_html=True)

    with right:
        st.markdown("<div class='panel-card'><div class='panel-title'>Meeting console</div>", unsafe_allow_html=True)
        reviewed = bool(ensure_review_flags().get(current_row["deal_key"], False))
        st.markdown(
            f"<span class='mini-chip'>Reviewed: {'Yes' if reviewed else 'No'}</span>"
            f"<span class='mini-chip'>Last saved: {html.escape(display_text(current_row.get('saved_at')))}</span>",
            unsafe_allow_html=True,
        )
        st.markdown("<div class='small-muted' style='margin:0.45rem 0 0.35rem 0;'>AM Commentary</div>", unsafe_allow_html=True)
        st.markdown(
            f"<div class='notes-box'>{html.escape(display_text(current_row.get('commentary'), blank='No commentary yet.'))}</div>",
            unsafe_allow_html=True,
        )
        st.markdown("<div class='small-muted' style='margin-top:0.55rem;'>Open the dialog when you need to update status, owner, commentary, or reviewed state without pushing the whole page downward.</div>", unsafe_allow_html=True)
        st.markdown("</div>", unsafe_allow_html=True)


def render_review_view(deck: pd.DataFrame, raw_deck: pd.DataFrame) -> None:
    st.markdown(
        "<div class='agenda-note'>Controlled review workflow: only <b>Reviewed</b>, <b>Status</b>, <b>Owner</b>, and <b>AM Commentary</b> are editable here.</div>",
        unsafe_allow_html=True,
    )

    review_scope = st.segmented_control(
        "Review scope",
        options=["All", "Pending", "Edited"],
        default="All",
        key="review_scope",
        width="content",
    )

    editor_df = build_review_dataframe(deck)
    flags = ensure_review_flags()
    overrides = ensure_override_store()

    if review_scope == "Pending":
        pending_keys = [key for key in editor_df.index if not bool(flags.get(key, False))]
        editor_df = editor_df.loc[pending_keys]
    elif review_scope == "Edited":
        changed_keys = {get_override_key(item["sheet"], item["deal_number"]) for item in overrides.values()}
        editor_df = editor_df.loc[[key for key in editor_df.index if key in changed_keys]]

    if editor_df.empty:
        st.info("No rows match the current review scope.")
        return

    disabled_columns = [
        "Agenda #",
        "Type",
        "Deal Name",
        "Deal #",
        "Maturity",
        "Next Payment",
        "UPB",
    ]

    edited_df = st.data_editor(
        editor_df,
        hide_index=True,
        width="stretch",
        num_rows="fixed",
        row_height=40,
        disabled=disabled_columns,
        column_config={
            "Agenda #": st.column_config.NumberColumn("Agenda #", width="small", format="%d"),
            "Type": st.column_config.TextColumn("Type", width="small"),
            "Deal Name": st.column_config.TextColumn("Deal Name", width="large"),
            "Deal #": st.column_config.TextColumn("Deal #", width="small"),
            "Maturity": st.column_config.TextColumn("Maturity", width="small"),
            "Next Payment": st.column_config.TextColumn("Next Payment", width="small"),
            "UPB": st.column_config.TextColumn("UPB", width="small"),
            "Reviewed": st.column_config.CheckboxColumn("Reviewed", width="small"),
            "Status": st.column_config.TextColumn("Status", width="medium"),
            "Owner": st.column_config.TextColumn("Owner / Point Person", width="medium"),
            "AM Commentary": st.column_config.TextColumn("AM Commentary", width="large"),
        },
        key="review_editor",
    )

    c1, c2, c3 = st.columns([1.1, 1.1, 2.2], vertical_alignment="center")
    with c1:
        if st.button("Apply review changes", use_container_width=True, type="primary"):
            apply_review_edits(edited_df, deck, raw_deck)
            st.toast("Review changes applied.")
            st.rerun()
    with c2:
        if st.button("Mark visible reviewed", use_container_width=True):
            visible_flags = ensure_review_flags()
            for deal_key in edited_df.index.tolist():
                visible_flags[deal_key] = True
            st.rerun()
    with c3:
        st.caption("Use Presentation for one-deal-at-a-time discussion. Use Review when you want a fast controlled pass down the agenda.")


def render_exports_view(raw_deck: pd.DataFrame, deck: pd.DataFrame, file_bytes: bytes, workbook_name: str) -> None:
    updates_df = build_updates_dataframe(raw_deck, deck)

    m1, m2 = st.columns(2)
    m1.metric("Reviewed deals", fmt_int(sum(1 for value in ensure_review_flags().values() if value)))
    m2.metric("Saved overrides", fmt_int(len(ensure_override_store())))

    if updates_df.empty:
        st.info("No reviewed rows or saved changes yet.")
    else:
        st.dataframe(updates_df, hide_index=True, use_container_width=True)

    c1, c2 = st.columns(2)
    with c1:
        st.download_button(
            "Download review log CSV",
            data=export_overrides_csv(raw_deck),
            file_name=f"{Path(workbook_name).stem}_meeting_updates.csv",
            mime="text/csv",
            disabled=updates_df.empty,
            use_container_width=True,
        )
    with c2:
        st.download_button(
            "Download updated workbook",
            data=update_workbook_bytes(file_bytes, ensure_override_store()),
            file_name=f"{Path(workbook_name).stem}_meeting_ready.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )


def get_workbook_bytes() -> Tuple[Optional[bytes], str]:
    uploaded = st.session_state.get("uploaded_workbook")
    if uploaded is not None:
        return uploaded.getvalue(), uploaded.name

    sample_file = first_existing_file(DEFAULT_SAMPLE_FILES)
    if sample_file is not None:
        return sample_file.read_bytes(), sample_file.name
    return None, ""


def main() -> None:
    st.set_page_config(page_title=APP_TITLE, layout="wide")
    initialize_state()
    apply_app_css()

    preview_bytes, preview_name = get_workbook_bytes()
    render_controls_bar(preview_name if preview_name else "No workbook loaded")
    file_bytes, workbook_name = get_workbook_bytes()

    if file_bytes is None:
        st.info("Open the Controls popover and upload your Portfolio Overview workbook to begin.")
        st.stop()

    raw_deck, _metadata = load_portfolio_workbook(
        file_bytes=file_bytes,
        as_of_date_iso=st.session_state.as_of_date.isoformat(),
        include_hidden=bool(st.session_state.include_hidden),
    )

    if raw_deck.empty:
        st.error("No usable Bridge / Term rows were found in this workbook.")
        st.stop()

    current_deck = apply_overrides(raw_deck, ensure_override_store())
    filtered_deck = apply_filters(current_deck)

    if filtered_deck.empty:
        st.warning("No deals match the current filters.")
        st.stop()

    sync_selected_deal(filtered_deck)
    maybe_open_edit_dialog(filtered_deck, raw_deck)

    if st.session_state.view_mode == "Overview":
        render_overview_view(filtered_deck, st.session_state.as_of_date)
    elif st.session_state.view_mode == "Presentation":
        render_presentation_view(filtered_deck, raw_deck)
    elif st.session_state.view_mode == "Review":
        render_review_view(filtered_deck, raw_deck)
    else:
        render_exports_view(raw_deck, current_deck, file_bytes, workbook_name)


if __name__ == "__main__":
    main()
