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
                padding-top: 1.1rem;
                padding-bottom: 2rem;
                max-width: 1450px;
            }
            div[data-testid="stMetric"] {
                background: linear-gradient(180deg, rgba(16,24,40,0.02), rgba(16,24,40,0.05));
                border: 1px solid rgba(15, 23, 42, 0.08);
                border-radius: 18px;
                padding: 0.9rem 1rem;
            }
            .hero-card {
                border-radius: 24px;
                padding: 1.2rem 1.35rem 1.1rem 1.35rem;
                border: 1px solid rgba(15, 23, 42, 0.08);
                background: linear-gradient(135deg, rgba(14,165,233,0.08), rgba(99,102,241,0.09));
                margin-bottom: 0.8rem;
            }
            .hero-pill {
                display: inline-block;
                padding: 0.25rem 0.65rem;
                border-radius: 999px;
                font-size: 0.78rem;
                font-weight: 700;
                letter-spacing: 0.04em;
                text-transform: uppercase;
                background: rgba(15, 23, 42, 0.08);
                margin-bottom: 0.6rem;
            }
            .hero-title {
                font-size: 2rem;
                font-weight: 800;
                line-height: 1.15;
                margin-bottom: 0.35rem;
            }
            .hero-subtitle {
                color: rgba(15, 23, 42, 0.75);
                font-size: 1rem;
            }
            .chip-row {
                margin-top: 0.7rem;
            }
            .chip {
                display: inline-block;
                margin: 0.2rem 0.35rem 0.15rem 0;
                padding: 0.35rem 0.65rem;
                border-radius: 999px;
                background: rgba(15, 23, 42, 0.07);
                font-size: 0.84rem;
                font-weight: 600;
            }
            .section-card {
                border: 1px solid rgba(15, 23, 42, 0.08);
                border-radius: 20px;
                padding: 1rem 1rem 0.75rem 1rem;
                background: rgba(255,255,255,0.7);
                height: 100%;
            }
            .section-title {
                font-size: 1rem;
                font-weight: 800;
                margin-bottom: 0.65rem;
            }
            .flag-list {
                margin: 0;
                padding-left: 1.1rem;
            }
            .flag-list li {
                margin-bottom: 0.35rem;
            }
            .small-muted {
                color: rgba(15, 23, 42, 0.72);
                font-size: 0.86rem;
            }
        </style>
        """,
        unsafe_allow_html=True,
    )


def norm_text(value) -> str:
    if value is None:
        return ""
    text = str(value).replace("\xa0", " ")
    text = re.sub(r"\s+", " ", text).strip()
    if text.endswith(".0"):
        text = text[:-2]
    return text


def canon_header(value) -> str:
    return norm_text(value).upper()


def find_header_row(ws, search_rows: int = 30, key_header: str = "Deal Number") -> Tuple[int, int]:
    max_row = min(search_rows, ws.max_row)
    for r in range(1, max_row + 1):
        vals = [ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
        if any(canon_header(v) == canon_header(key_header) for v in vals if v is not None):
            return r, ws.max_column
    raise ValueError(f"Could not find header row in sheet {ws.title!r}.")


def latest_dated_by_suffix(headers: Iterable) -> Dict[str, str]:
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
            best[suffix] = (month_day, str(header))
    return {suffix: full for suffix, (_md, full) in best.items()}


def resolve_column(columns: Iterable, *candidates: Optional[str]) -> Optional[str]:
    cols = [str(c) for c in columns if c is not None]
    for candidate in candidates:
        if not candidate:
            continue
        for col in cols:
            if canon_header(col) == canon_header(candidate):
                return col
    return None


def matching_columns(columns: Iterable, candidate: str) -> List[str]:
    return [str(c) for c in columns if c is not None and canon_header(c) == canon_header(candidate)]


def safe_series(df: pd.DataFrame, column_name: Optional[str], default="") -> pd.Series:
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


@st.cache_data(show_spinner=False)
def load_portfolio_workbook(
    file_bytes: bytes,
    as_of_date_iso: str,
    include_hidden: bool,
) -> Tuple[pd.DataFrame, Dict[str, Dict[str, str]]]:
    as_of_date = pd.Timestamp(as_of_date_iso)
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

            row_dict["_sheet"] = sheet_name
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

        df["sheet"] = sheet_name
        df["sheet_order"] = 0 if sheet_name == "Bridge" else 1
        df["deal_number"] = safe_series(df, deal_col, "").map(norm_text)
        df["deal_name"] = safe_series(df, deal_name_col, "").map(norm_text)
        df["borrower"] = safe_series(df, borrower_col, "").map(norm_text)
        df["account_display"] = safe_series(df, account_col, "").map(norm_text)
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
    deck["upb_mm"] = deck["upb"] / 1_000_000

    deck["watch_flag"] = (
        deck["npl_raw"].astype(str).str.upper().str.contains(r"\bY\b|90\+|NPL|DQ 90", regex=True, na=False)
        | (deck["days_to_maturity"].fillna(9999) <= 30)
        | (deck["days_past_due"].fillna(0) >= 30)
        | (deck["days_to_next_payment"].fillna(9999) < 0)
    )

    deck["sort_watch"] = deck["watch_flag"].astype(int)
    deck = deck.sort_values(
        ["sheet_order", "sort_watch", "maturity_date", "upb"],
        ascending=[True, False, True, False],
        kind="stable",
    ).reset_index(drop=True)

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

    drop_cols = [c for c in merged.columns if c.endswith("_override")]
    return merged.drop(columns=drop_cols)


def build_flag_messages(row: pd.Series) -> List[str]:
    flags: List[str] = []
    days_to_maturity = row.get("days_to_maturity")
    days_to_next_payment = row.get("days_to_next_payment")
    days_past_due = row.get("days_past_due")
    npl_raw = norm_text(row.get("npl_raw", ""))

    if pd.notna(days_to_maturity) and days_to_maturity < 0:
        flags.append(f"Maturity has already passed ({fmt_day_delta(days_to_maturity)}).")
    elif pd.notna(days_to_maturity) and days_to_maturity <= 30:
        flags.append(f"Maturity is within 30 days ({fmt_day_delta(days_to_maturity)}).")

    if pd.notna(days_to_next_payment) and days_to_next_payment < 0:
        flags.append(f"Next payment date is past due ({fmt_day_delta(days_to_next_payment)}).")
    elif pd.notna(days_to_next_payment) and days_to_next_payment <= 15:
        flags.append(f"Next payment date is approaching soon ({fmt_day_delta(days_to_next_payment)}).")

    if pd.notna(days_past_due) and float(days_past_due) > 0:
        flags.append(f"Days past due is currently {fmt_int(days_past_due)}.")

    if npl_raw and canon_header(npl_raw) not in {"N", "NO", ""}:
        flags.append(f"Watchlist / NPL signal present: {npl_raw}.")

    if not flags:
        flags.append("No immediate red flags from the automated screen. Use meeting notes for qualitative context.")

    return flags


def render_metric_row(row: pd.Series) -> None:
    cols = st.columns(5)
    cols[0].metric("UPB", fmt_money(row.get("upb"), decimals=0))
    cols[1].metric("Maturity", fmt_date(row.get("maturity_date")), fmt_day_delta(row.get("days_to_maturity")))
    cols[2].metric(
        "Next Payment",
        fmt_date(row.get("next_payment_date")),
        fmt_day_delta(row.get("days_to_next_payment")),
    )
    cols[3].metric("Days Past Due", fmt_int(row.get("days_past_due")))
    cols[4].metric("Watchlist", "Yes" if bool(row.get("watch_flag")) else "No")


def build_detail_table(row: pd.Series) -> pd.DataFrame:
    records = [
        ("Deal Number", row.get("deal_number", "")),
        ("Portfolio", row.get("portfolio", "")),
        ("Segment", row.get("segment", "")),
        ("Servicer", row.get("servicer", "")),
        ("Owner / Point Person", row.get("owner", "")),
        ("Borrower", row.get("borrower", "")),
        ("Account", row.get("account_display", "")),
        ("Financing", row.get("financing", "")),
        ("Loan Buyer", row.get("loan_buyer", "")),
        ("Status", row.get("status", "")),
        ("Maturity Date", fmt_date(row.get("maturity_date"))),
        ("Next Payment Date", fmt_date(row.get("next_payment_date"))),
    ]

    if row.get("sheet") == "Bridge":
        records.extend(
            [
                ("Commitment", fmt_money(row.get("commitment"), decimals=0)),
                ("Funded Amount", fmt_money(row.get("funded_amount"), decimals=0)),
                ("Remaining Commitment", fmt_money(row.get("remaining_commitment"), decimals=0)),
            ]
        )
    else:
        records.append(("Loan Amount", fmt_money(row.get("loan_amount"), decimals=0)))

    return pd.DataFrame(records, columns=["Field", "Value"])


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
    }
    for col, default_value in needed_defaults.items():
        if col not in queue.columns:
            queue[col] = default_value

    queue = queue[
        [
            "sheet_order",
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

    queue["Maturity"] = queue["Maturity"].map(fmt_date)
    queue["Next Payment"] = queue["Next Payment"].map(fmt_date)
    queue["UPB"] = queue["UPB"].map(lambda x: fmt_money(x, decimals=0))
    queue["DPD"] = queue["DPD"].map(fmt_int)
    queue["Watch"] = queue["Watch"].map(lambda x: "Yes" if bool(x) else "No")

    queue = queue.sort_values(["sheet_order", "Type", "Deal #"], ascending=[True, True, True], kind="stable")
    queue = queue.drop(columns=["sheet_order"])

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
            if owner_col is not None and "owner" in item:
                ws.cell(target_row, owner_col).value = item.get("owner", "")
            if commentary_col is not None:
                ws.cell(target_row, commentary_col).value = item.get("commentary", "")

    out = io.BytesIO()
    workbook.save(out)
    out.seek(0)
    return out.getvalue()


def available_status_suggestions(deck: pd.DataFrame) -> List[str]:
    if "status" not in deck.columns:
        return COMMON_STATUS_SUGGESTIONS
    status_values = sorted({norm_text(v) for v in deck["status"].dropna().tolist() if norm_text(v)})
    return sorted(set(COMMON_STATUS_SUGGESTIONS + status_values))


def render_overview_tab(deck: pd.DataFrame, as_of_date: dt.date, metadata: Dict[str, Dict[str, str]]) -> None:
    if deck.empty:
        st.warning("No deals match the current filters.")
        return

    total_upb = deck["upb"].fillna(0).sum()
    watch_count = int(deck["watch_flag"].fillna(False).sum())
    next_30 = int((deck["days_to_maturity"].fillna(9999) <= 30).sum())
    overdue_pay = int((deck["days_to_next_payment"].fillna(9999) < 0).sum())

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Deals in deck", fmt_int(len(deck)))
    m2.metric("Total UPB", fmt_money(total_upb, decimals=0))
    m3.metric("Maturing in 30d", fmt_int(next_30))
    m4.metric("Past due next pay", fmt_int(overdue_pay))

    st.caption(
        f"As of {as_of_date.strftime('%m/%d/%Y')} • "
        + " • ".join(
            f"{sheet}: {meta.get('upb_header', 'UPB') or 'UPB'}"
            + (
                f", {meta.get('maturity_header')}" if meta.get("maturity_header") else ""
            )
            for sheet, meta in metadata.items()
        )
    )

    col1, col2 = st.columns([1.1, 1.0])

    with col1:
        st.subheader("Nearest maturities")
        maturities = deck[
            ["sheet_order", "sheet", "deal_number", "deal_name", "maturity_date", "days_to_maturity", "upb"]
        ].copy()
        maturities = maturities.sort_values(
            ["sheet_order", "maturity_date", "upb"],
            ascending=[True, True, False],
            kind="stable",
        ).head(15)
        maturities["Maturity"] = maturities["maturity_date"].map(fmt_date)
        maturities["Timing"] = maturities["days_to_maturity"].map(fmt_day_delta)
        maturities["UPB"] = maturities["upb"].map(lambda x: fmt_money(x, decimals=0))
        st.dataframe(
            maturities[["sheet", "deal_number", "deal_name", "Maturity", "Timing", "UPB"]].rename(
                columns={"sheet": "Type", "deal_number": "Deal #", "deal_name": "Deal Name"}
            ),
            use_container_width=True,
            hide_index=True,
        )

    with col2:
        st.subheader("Watchlist-style queue")
        watch = deck[
            ["sheet_order", "sheet", "deal_number", "deal_name", "status", "days_past_due", "next_payment_date"]
        ].copy()
        watch["Next Payment"] = watch["next_payment_date"].map(fmt_date)
        watch["DPD"] = watch["days_past_due"].map(fmt_int)
        watch = watch.sort_values(
            ["sheet_order", "days_past_due", "deal_number"],
            ascending=[True, False, True],
            kind="stable",
        ).head(15)
        st.dataframe(
            watch[["sheet", "deal_number", "deal_name", "status", "DPD", "Next Payment"]].rename(
                columns={
                    "sheet": "Type",
                    "deal_number": "Deal #",
                    "deal_name": "Deal Name",
                    "status": "Status",
                }
            ),
            use_container_width=True,
            hide_index=True,
        )

    st.subheader("Full queue")
    st.dataframe(build_queue_dataframe(deck), use_container_width=True, hide_index=True)


def render_presentation_tab(deck: pd.DataFrame, status_suggestions: List[str]) -> None:
    if deck.empty:
        st.warning("No deals match the current filters.")
        return

    if "deck_index" not in st.session_state:
        st.session_state.deck_index = 0
    st.session_state.deck_index = int(np.clip(st.session_state.deck_index, 0, len(deck) - 1))

    nav_left, nav_mid, nav_right = st.columns([1, 1, 3])
    with nav_left:
        if st.button("⬅ Previous", use_container_width=True, disabled=st.session_state.deck_index == 0):
            st.session_state.deck_index -= 1
    with nav_mid:
        if st.button("Next ➡", use_container_width=True, disabled=st.session_state.deck_index >= len(deck) - 1):
            st.session_state.deck_index += 1
    with nav_right:
        jump_options = list(range(len(deck)))
        current_index = st.selectbox(
            "Jump to deal",
            options=jump_options,
            index=st.session_state.deck_index,
            format_func=lambda i: (
                f"{i + 1}. {deck.iloc[i]['sheet']} | "
                f"{deck.iloc[i]['deal_number']} | "
                f"{deck.iloc[i]['deal_name']}"
            ),
        )
        st.session_state.deck_index = current_index

    row = deck.iloc[st.session_state.deck_index]
    progress = (st.session_state.deck_index + 1) / len(deck)
    st.progress(progress, text=f"Deal {st.session_state.deck_index + 1} of {len(deck)}")

    hero_html = f"""
        <div class='hero-card'>
            <div class='hero-pill'>{html.escape(str(row.get('sheet', '')))}</div>
            <div class='hero-title'>{html.escape(str(row.get('deal_name', '')))}</div>
            <div class='hero-subtitle'>
                Deal {html.escape(str(row.get('deal_number', '')))} • {html.escape(str(row.get('borrower', '')))} •
                {html.escape(str(row.get('status', '')) or 'No status')}
            </div>
            <div class='chip-row'>
                <span class='chip'>Servicer: {html.escape(str(row.get('servicer', '')) or '-')}</span>
                <span class='chip'>Owner: {html.escape(str(row.get('owner', '')) or '-')}</span>
                <span class='chip'>Portfolio: {html.escape(str(row.get('portfolio', '')) or '-')}</span>
                <span class='chip'>Watchlist: {"Yes" if bool(row.get('watch_flag')) else "No"}</span>
            </div>
        </div>
    """
    st.markdown(hero_html, unsafe_allow_html=True)
    render_metric_row(row)

    info_col, flags_col, edit_col = st.columns([1.15, 0.85, 1.0])

    with info_col:
        st.markdown("<div class='section-card'><div class='section-title'>Core snapshot</div>", unsafe_allow_html=True)
        st.table(build_detail_table(row))
        st.markdown("</div>", unsafe_allow_html=True)

    with flags_col:
        flags = build_flag_messages(row)
        flag_html = "".join(f"<li>{html.escape(flag)}</li>" for flag in flags)
        st.markdown(
            f"""
            <div class='section-card'>
                <div class='section-title'>Talking points / prompts</div>
                <ul class='flag-list'>{flag_html}</ul>
                <div class='small-muted' style='margin-top:0.75rem;'>
                    Highlights are driven by maturity timing, payment timing, delinquency, and NPL / watchlist signals.
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with edit_col:
        st.markdown("<div class='section-card'><div class='section-title'>Live meeting updates</div>", unsafe_allow_html=True)
        st.caption("Type directly here during the meeting, then save. Exports are on the Meeting Log tab.")

        if status_suggestions:
            st.caption("Quick ideas: " + " | ".join(status_suggestions[:8]))

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
            "AM commentary / meeting notes",
            value=str(row.get("commentary", "")),
            height=185,
            key=f"commentary_input::{row['sheet']}::{row['deal_number']}",
        )

        if st.button("Save this update", use_container_width=True, type="primary"):
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
            st.toast("Saved to the meeting update log.")

        st.markdown("</div>", unsafe_allow_html=True)


def render_meeting_log_tab(
    raw_deck: pd.DataFrame,
    displayed_deck: pd.DataFrame,
    file_bytes: bytes,
    workbook_name: str,
) -> None:
    overrides = ensure_override_store()
    update_count = len(overrides)
    st.metric("Saved meeting updates", fmt_int(update_count))

    if update_count == 0:
        st.info("No live updates saved yet. Use the Presentation tab to update status, owner, or AM commentary.")
    else:
        updates_df = pd.DataFrame(overrides.values()).sort_values(["sheet", "deal_number"])
        st.dataframe(updates_df, use_container_width=True, hide_index=True)

    export_col1, export_col2 = st.columns(2)

    with export_col1:
        csv_bytes = export_overrides_csv(overrides)
        st.download_button(
            "Download meeting updates CSV",
            data=csv_bytes,
            file_name=f"{Path(workbook_name).stem}_meeting_updates.csv",
            mime="text/csv",
            disabled=not bool(overrides),
            use_container_width=True,
        )

    with export_col2:
        workbook_bytes = update_workbook_bytes(file_bytes, overrides)
        st.download_button(
            "Download updated workbook",
            data=workbook_bytes,
            file_name=f"{Path(workbook_name).stem}_meeting_ready.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    st.subheader("Current queue")
    st.dataframe(build_queue_dataframe(displayed_deck), use_container_width=True, hide_index=True)

    st.subheader("Deck-ready export view")
    export_table = displayed_deck.copy()

    needed_defaults = {
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
        "watch_flag": False,
        "sheet_order": 99,
    }
    for col, default_value in needed_defaults.items():
        if col not in export_table.columns:
            export_table[col] = default_value

    export_table["UPB"] = export_table["upb"].map(lambda x: fmt_money(x, decimals=0))
    export_table["Maturity Date"] = export_table["maturity_date"].map(fmt_date)
    export_table["Next Payment Date"] = export_table["next_payment_date"].map(fmt_date)
    export_table["Days Past Due"] = export_table["days_past_due"].map(fmt_int)
    export_table["Watchlist"] = export_table["watch_flag"].map(lambda x: "Yes" if bool(x) else "No")

    export_table = export_table.sort_values(
        ["sheet_order", "deal_number"],
        ascending=[True, True],
        kind="stable",
    )

    st.dataframe(
        export_table[
            [
                "sheet",
                "deal_number",
                "deal_name",
                "status",
                "owner",
                "UPB",
                "Maturity Date",
                "Next Payment Date",
                "Days Past Due",
                "Watchlist",
                "commentary",
            ]
        ].rename(
            columns={
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

    st.title(APP_TITLE)
    st.caption("Interactive weekly loan-by-loan meeting view for Bridge first, then Term.")

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
    sheet_choices = st.sidebar.multiselect(
        "Loan type",
        options=["Bridge", "Term"],
        default=["Bridge", "Term"],
    )

    if not sheet_choices:
        filtered = deck.iloc[0:0].copy()
    else:
        filtered = deck[deck["sheet"].isin(sheet_choices)].copy()

    watch_only = st.sidebar.checkbox("Watchlist only", value=False)
    if watch_only and not filtered.empty:
        filtered = filtered[filtered["watch_flag"] == True].copy()

    search_text = st.sidebar.text_input("Search deal # / name / borrower")
    if search_text and not filtered.empty:
        search_upper = search_text.strip().upper()
        filtered = filtered[
            filtered["deal_number"].astype(str).str.upper().str.contains(search_upper, na=False)
            | filtered["deal_name"].astype(str).str.upper().str.contains(search_upper, na=False)
            | filtered["borrower"].astype(str).str.upper().str.contains(search_upper, na=False)
        ].copy()

    filtered = filtered.sort_values(
        ["sheet_order", "sort_watch", "maturity_date", "upb"],
        ascending=[True, False, True, False],
        kind="stable",
    ).reset_index(drop=True)

    tabs = st.tabs(["Overview", "Presentation Mode", "Meeting Log"])

    with tabs[0]:
        render_overview_tab(filtered, as_of_date, metadata)

    with tabs[1]:
        render_presentation_tab(filtered, status_suggestions)

    with tabs[2]:
        render_meeting_log_tab(
            raw_deck=deck,
            displayed_deck=filtered,
            file_bytes=file_bytes,
            workbook_name=workbook_name,
        )


if __name__ == "__main__":
    main()
