from __future__ import annotations

import datetime as dt
import html
import io
import re
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple

import plotly.express as px
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
                padding-top: 1.25rem;
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
                padding: 1.25rem 1.4rem 1.15rem 1.4rem;
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


def fmt_number(value: object, decimals: int = 0, blank: str = "-") -> str:
    if value is None or (isinstance(value, float) and np.isnan(value)):
        return blank
    try:
        number = float(value)
    except Exception:
        return blank
    return f"{number:,.{decimals}f}"


def fmt_int(value: object, blank: str = "-") -> str:
    if value is None or (isinstance(value, float) and np.isnan(value)):
        return blank
    try:
        return f"{int(round(float(value))):,}"
    except Exception:
        return blank


def fmt_date(value: object, blank: str = "-") -> str:
    if value is None or (isinstance(value, float) and np.isnan(value)):
        return blank
    try:
        timestamp = pd.to_datetime(value)
    except Exception:
        return blank
    if pd.isna(timestamp):
        return blank
    return timestamp.strftime("%b %d, %Y")


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
def load_portfolio_workbook(file_bytes: bytes, as_of_date_iso: str, include_hidden: bool) -> Tuple[pd.DataFrame, Dict[str, Dict[str, str]]]:
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

            deal_value = row_dict.get(resolve_column(row_dict.keys(), "Deal Number") or "Deal Number")
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
        status_col = resolve_column(df.columns, "Status")
        commentary_col = resolve_column(df.columns, "AM Commentary")
        owner_col = resolve_column(df.columns, "Point Person", "Asset Manager", "Active RM")
        borrower_col = resolve_column(df.columns, "Borrower Name", "Borrower Entity", "Account Name")
        account_col = resolve_column(df.columns, "Account", "Account Name")
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
        df["deal_number"] = df[deal_col].map(norm_text) if deal_col else ""
        df["deal_name"] = df[resolve_column(df.columns, "Deal Name")].map(norm_text)
        df["borrower"] = df[borrower_col].map(norm_text) if borrower_col else ""
        df["account_display"] = df[account_col].map(norm_text) if account_col else ""
        df["servicer"] = df[resolve_column(df.columns, "Servicer")].map(norm_text)
        df["portfolio"] = (
            df[resolve_column(df.columns, "Portfolio")].map(norm_text)
            if resolve_column(df.columns, "Portfolio")
            else sheet_name
        )
        df["segment"] = (
            df[resolve_column(df.columns, "Segment")].map(norm_text)
            if resolve_column(df.columns, "Segment")
            else ""
        )
        df["financing"] = (
            df[resolve_column(df.columns, "Financing")].map(norm_text)
            if resolve_column(df.columns, "Financing")
            else ""
        )
        df["loan_buyer"] = (
            df[resolve_column(df.columns, "Loan Buyer")].map(norm_text)
            if resolve_column(df.columns, "Loan Buyer")
            else ""
        )
        df["owner"] = df[owner_col].map(norm_text) if owner_col else ""
        df["status"] = df[status_col].map(norm_text) if status_col else ""
        df["commentary"] = df[commentary_col].map(norm_text) if commentary_col else ""
        df["upb"] = pd.to_numeric(df[upb_col], errors="coerce") if upb_col else np.nan
        df["maturity_date"] = pd.to_datetime(df[maturity_col], errors="coerce") if maturity_col else pd.NaT
        df["next_payment_date"] = pd.to_datetime(df[next_payment_col], errors="coerce") if next_payment_col else pd.NaT
        df["npl_raw"] = df[npl_col].astype(str).fillna("") if npl_col else ""

        if dpd_cols:
            dpd_matrix = np.column_stack(
                [pd.to_numeric(df[col], errors="coerce").fillna(0).to_numpy(dtype=float) for col in dpd_cols]
            )
            df["days_past_due"] = np.maximum(dpd_matrix.max(axis=1), 0)
        else:
            df["days_past_due"] = 0

        if sheet_name == "Bridge":
            df["funded_amount"] = pd.to_numeric(
                df[resolve_column(df.columns, "Active Funded Amount")], errors="coerce"
            ) if resolve_column(df.columns, "Active Funded Amount") else np.nan
            df["commitment"] = pd.to_numeric(
                df[resolve_column(df.columns, "Loan Commitment")], errors="coerce"
            ) if resolve_column(df.columns, "Loan Commitment") else np.nan
            df["remaining_commitment"] = pd.to_numeric(
                df[resolve_column(df.columns, "Remaining Commitment")], errors="coerce"
            ) if resolve_column(df.columns, "Remaining Commitment") else np.nan
            df["loan_amount"] = np.nan
        else:
            df["loan_amount"] = pd.to_numeric(
                df[resolve_column(df.columns, "Loan Amount")], errors="coerce"
            ) if resolve_column(df.columns, "Loan Amount") else np.nan
            df["funded_amount"] = np.nan
            df["commitment"] = np.nan
            df["remaining_commitment"] = np.nan

        metadata[sheet_name] = {
            "upb_header": upb_col or "",
            "npl_header": npl_col or "",
            "maturity_header": maturity_col or "",
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

    maturity_days = deck["days_to_maturity"].fillna(9999).to_numpy(dtype=float)
    payment_days = deck["days_to_next_payment"].fillna(9999).to_numpy(dtype=float)
    dpd_days = deck["days_past_due"].fillna(0).to_numpy(dtype=float)
    upb_values = deck["upb"].fillna(0).to_numpy(dtype=float)
    status_upper = deck["status"].fillna("").str.upper()
    npl_upper = deck["npl_raw"].fillna("").astype(str).str.upper()

    maturity_urgency = np.clip((180 - maturity_days) / 180.0, 0, 1)
    payment_urgency = np.clip((45 - payment_days) / 45.0, 0, 1)
    delinquency_pressure = np.clip(dpd_days / 180.0, 0, 1)

    status_pressure = np.select(
        [
            status_upper.str.contains(r"FORECLOS|RECEIVERSHIP|LITIGATION|PROBATE|DEMAND", regex=True),
            status_upper.str.contains(r"UPCOMING MATURITY|CONSENT|SALE|MODIFICATION|RENT CONTROL", regex=True),
            status_upper.str.contains(r"SURVEILLANCE|POST CLOSING", regex=True),
        ],
        [1.0, 0.7, 0.3],
        default=0.1,
    )
    npl_pressure = np.where(npl_upper.str.contains(r"\bY\b|90\+|NPL|DQ 90", regex=True), 1.0, 0.0)
    size_component = np.log1p(upb_values)
    size_component = np.divide(
        size_component,
        size_component.max() if size_component.max() > 0 else 1,
        out=np.zeros_like(size_component),
        where=True,
    )

    raw_pressure = (
        0.28 * maturity_urgency
        + 0.20 * payment_urgency
        + 0.22 * delinquency_pressure
        + 0.15 * status_pressure
        + 0.10 * npl_pressure
        + 0.05 * size_component
    )
    pressure_score = np.rint(raw_pressure * 100).astype(int)

    order = np.argsort(pressure_score)
    ranks = np.empty_like(order)
    ranks[order] = np.arange(1, len(pressure_score) + 1)
    percentile = np.rint(100 * ranks / len(pressure_score)).astype(int)

    deck["pressure_score"] = pressure_score
    deck["pressure_percentile"] = percentile
    deck["pressure_bucket"] = np.select(
        [pressure_score >= 75, pressure_score >= 55, pressure_score >= 35],
        ["Critical", "High", "Moderate"],
        default="Low",
    )
    deck["watch_flag"] = (
        (deck["pressure_score"] >= 55)
        | (deck["days_to_maturity"].fillna(9999) <= 30)
        | (deck["days_past_due"].fillna(0) >= 30)
    )

    preferred_sort = deck["watch_flag"].astype(int) * 1000 + deck["pressure_score"]
    deck["deck_rank"] = (-preferred_sort).rank(method="first").astype(int)

    return deck, metadata


def get_override_key(sheet: str, deal_number: str) -> str:
    return f"{sheet}::{deal_number}"


def ensure_override_store() -> Dict[str, Dict[str, str]]:
    if "meeting_overrides" not in st.session_state:
        st.session_state.meeting_overrides = {}
    return st.session_state.meeting_overrides


def apply_overrides(deck: pd.DataFrame, overrides: Dict[str, Dict[str, str]]) -> pd.DataFrame:
    if deck.empty or not overrides:
        return deck.copy()

    out = deck.copy()
    override_df = pd.DataFrame(overrides.values())
    if override_df.empty:
        return out

    for field in ["status", "owner", "commentary"]:
        if field not in override_df.columns:
            override_df[field] = None

    merged = out.merge(
        override_df[["sheet", "deal_number", "status", "owner", "commentary", "saved_at"]],
        on=["sheet", "deal_number"],
        how="left",
        suffixes=("", "_override"),
    )

    for field in ["status", "owner", "commentary"]:
        override_field = f"{field}_override"
        merged[field] = merged[override_field].where(
            merged[override_field].notna() & (merged[override_field].astype(str) != ""),
            merged[field],
        )

    merged["saved_at"] = merged["saved_at"].fillna("")
    drop_cols = [c for c in merged.columns if c.endswith("_override")]
    return merged.drop(columns=drop_cols)


def build_flag_messages(row: pd.Series) -> List[str]:
    flags: List[str] = []
    days_to_maturity = row.get("days_to_maturity")
    days_to_next_payment = row.get("days_to_next_payment")
    days_past_due = row.get("days_past_due")
    pressure_bucket = row.get("pressure_bucket", "")
    npl_raw = norm_text(row.get("npl_raw", ""))

    if pd.notna(days_to_maturity) and days_to_maturity <= 30:
        flags.append(f"Maturity is within 30 days ({fmt_day_delta(days_to_maturity)}).")
    elif pd.notna(days_to_maturity) and days_to_maturity < 0:
        flags.append(f"Maturity has already passed ({fmt_day_delta(days_to_maturity)}).")

    if pd.notna(days_to_next_payment) and days_to_next_payment < 0:
        flags.append(f"Next payment date is past due ({fmt_day_delta(days_to_next_payment)}).")

    if pd.notna(days_past_due) and float(days_past_due) > 0:
        flags.append(f"Days past due is currently {fmt_int(days_past_due)}.")

    if npl_raw and canon_header(npl_raw) not in {"N", "NO", ""}:
        flags.append(f"Watchlist/NPL signal present: {html.escape(npl_raw)}.")

    if pressure_bucket in {"Critical", "High"}:
        flags.append(f"Portfolio pressure score is {row.get('pressure_score', '-')}, marked {pressure_bucket}.")

    if not flags:
        flags.append("No immediate red flags from the automated screen. Use meeting notes for qualitative context.")

    return flags


def render_metric_row(row: pd.Series) -> None:
    cols = st.columns(6)
    cols[0].metric("UPB", fmt_money(row.get("upb"), decimals=0))
    cols[1].metric("Maturity", fmt_date(row.get("maturity_date")), fmt_day_delta(row.get("days_to_maturity")))
    cols[2].metric("Next Payment", fmt_date(row.get("next_payment_date")), fmt_day_delta(row.get("days_to_next_payment")))
    cols[3].metric("Days Past Due", fmt_int(row.get("days_past_due")))
    cols[4].metric("Pressure Score", fmt_int(row.get("pressure_score")), f"P{fmt_int(row.get('pressure_percentile'))}")
    cols[5].metric("Watchlist", "Yes" if bool(row.get("watch_flag")) else "No", row.get("pressure_bucket", ""))


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
    queue = deck[
        [
            "sheet",
            "deal_number",
            "deal_name",
            "status",
            "owner",
            "upb",
            "maturity_date",
            "next_payment_date",
            "days_past_due",
            "pressure_score",
            "pressure_bucket",
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
            "pressure_score": "Pressure",
            "pressure_bucket": "Bucket",
            "watch_flag": "Watch",
            "saved_at": "Last Saved",
        }
    )
    return queue.sort_values(["Pressure", "UPB"], ascending=[False, False])


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
    status_values = sorted({norm_text(v) for v in deck["status"].dropna().tolist() if norm_text(v)})
    return sorted(set(COMMON_STATUS_SUGGESTIONS + status_values))


def render_overview_tab(deck: pd.DataFrame, as_of_date: dt.date, metadata: Dict[str, Dict[str, str]]) -> None:
    total_upb = deck["upb"].fillna(0).sum()
    watch_count = int(deck["watch_flag"].fillna(False).sum())
    next_30 = int((deck["days_to_maturity"].fillna(9999) <= 30).sum())
    overdue_pay = int((deck["days_to_next_payment"].fillna(9999) < 0).sum())

    m1, m2, m3, m4, m5 = st.columns(5)
    m1.metric("Deals in deck", fmt_int(len(deck)))
    m2.metric("Total UPB", fmt_money(total_upb, decimals=0))
    m3.metric("Watchlist", fmt_int(watch_count))
    m4.metric("Maturing in 30d", fmt_int(next_30))
    m5.metric("Past due next pay", fmt_int(overdue_pay))

    sheet_counts = deck.groupby("sheet")["deal_number"].count().reset_index(name="Deals")
    sheet_counts["UPB"] = deck.groupby("sheet")["upb"].sum().values
    st.caption(
        f"As of {as_of_date.strftime('%b %d, %Y')} • "
        + " • ".join(
            f"{sheet}: {meta.get('upb_header', 'UPB') or 'UPB'}"
            + (f", {meta.get('npl_header')}" if meta.get("npl_header") else "")
            for sheet, meta in metadata.items()
        )
    )

    left, right = st.columns([1.45, 1.0])

    with left:
        scatter_source = deck.copy()
        scatter_source["UPB ($MM)"] = scatter_source["upb_mm"].round(2)
        scatter_source["Days to Maturity"] = scatter_source["days_to_maturity"].fillna(9999)
        scatter_source["Tooltip"] = (
            scatter_source["sheet"]
            + " • "
            + scatter_source["deal_number"]
            + " • "
            + scatter_source["deal_name"]
        )

        scatter_fig = px.scatter(
            scatter_source,
            x="Days to Maturity",
            y="UPB ($MM)",
            color="pressure_bucket",
            size="pressure_score",
            hover_name="Tooltip",
            hover_data={
                "status": True,
                "Days to Maturity": True,
                "UPB ($MM)": ':.2f',
                "days_past_due": True,
                "pressure_score": True,
            },
            title="Portfolio radar: size vs. time-to-maturity",
        )
        scatter_fig.update_layout(height=420, margin=dict(l=10, r=10, t=50, b=10))
        st.plotly_chart(scatter_fig, use_container_width=True)

    with right:
        top_watch = (
            deck.sort_values(["pressure_score", "upb"], ascending=[False, False])
            [["sheet", "deal_number", "deal_name", "status", "pressure_score", "days_to_maturity"]]
            .head(10)
            .rename(
                columns={
                    "sheet": "Type",
                    "deal_number": "Deal #",
                    "deal_name": "Deal Name",
                    "status": "Status",
                    "pressure_score": "Pressure",
                    "days_to_maturity": "Days to Mat.",
                }
            )
        )
        st.dataframe(top_watch, use_container_width=True, hide_index=True)

    bottom_left, bottom_right = st.columns(2)
    with bottom_left:
        maturities = deck[
            ["sheet", "deal_number", "deal_name", "maturity_date", "days_to_maturity", "upb"]
        ].copy()
        maturities = maturities.sort_values(["maturity_date", "upb"], ascending=[True, False]).head(12)
        maturities["Maturity"] = maturities["maturity_date"].map(fmt_date)
        maturities["UPB"] = maturities["upb"].map(lambda x: fmt_money(x, decimals=0))
        maturities["Timing"] = maturities["days_to_maturity"].map(fmt_day_delta)
        st.subheader("Nearest maturities")
        st.dataframe(
            maturities[["sheet", "deal_number", "deal_name", "Maturity", "Timing", "UPB"]].rename(
                columns={"sheet": "Type", "deal_number": "Deal #", "deal_name": "Deal Name"}
            ),
            use_container_width=True,
            hide_index=True,
        )

    with bottom_right:
        status_summary = (
            deck.groupby(["sheet", "pressure_bucket"], dropna=False)["deal_number"]
            .count()
            .reset_index(name="Deals")
            .rename(columns={"sheet": "Type", "pressure_bucket": "Pressure Bucket"})
        )
        status_fig = px.bar(
            status_summary,
            x="Pressure Bucket",
            y="Deals",
            color="Type",
            category_orders={"Pressure Bucket": ["Low", "Moderate", "High", "Critical"]},
            title="Deck mix by pressure bucket",
            barmode="group",
        )
        status_fig.update_layout(height=320, margin=dict(l=10, r=10, t=50, b=10))
        st.plotly_chart(status_fig, use_container_width=True)


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
        if st.button(
            "Next ➡",
            use_container_width=True,
            disabled=st.session_state.deck_index >= len(deck) - 1,
        ):
            st.session_state.deck_index += 1
    with nav_right:
        jump_options = list(range(len(deck)))
        current_index = st.selectbox(
            "Jump to deal",
            options=jump_options,
            index=st.session_state.deck_index,
            format_func=lambda i: f"{i + 1}. {deck.iloc[i]['sheet']} | {deck.iloc[i]['deal_number']} | {deck.iloc[i]['deal_name']}",
        )
        st.session_state.deck_index = current_index

    row = deck.iloc[st.session_state.deck_index]
    progress = (st.session_state.deck_index + 1) / len(deck)
    st.progress(progress, text=f"Deal {st.session_state.deck_index + 1} of {len(deck)}")

    hero_html = f"""
        <div class='hero-card'>
            <div class='hero-pill'>{html.escape(str(row.get('sheet', '')))} • {html.escape(str(row.get('pressure_bucket', '')))}</div>
            <div class='hero-title'>{html.escape(str(row.get('deal_name', '')))}</div>
            <div class='hero-subtitle'>
                Deal {html.escape(str(row.get('deal_number', '')))} • {html.escape(str(row.get('borrower', '')))} •
                {html.escape(str(row.get('status', '')) or 'No status')}
            </div>
            <div class='chip-row'>
                <span class='chip'>Servicer: {html.escape(str(row.get('servicer', '')) or '-')}</span>
                <span class='chip'>Owner: {html.escape(str(row.get('owner', '')) or '-')}</span>
                <span class='chip'>Portfolio: {html.escape(str(row.get('portfolio', '')) or '-')}</span>
                <span class='chip'>Pressure: {html.escape(fmt_int(row.get('pressure_score')))}</span>
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
                    Auto-generated from maturity timing, payment timing, delinquency, status language, and UPB weighting.
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with edit_col:
        st.markdown("<div class='section-card'><div class='section-title'>Live meeting updates</div>", unsafe_allow_html=True)
        st.caption("Type directly here during the meeting, then save. Exports are available on the Meeting Log tab.")

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


def render_meeting_log_tab(raw_deck: pd.DataFrame, displayed_deck: pd.DataFrame, file_bytes: bytes, workbook_name: str) -> None:
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

    st.subheader("Deck-ready table")
    export_table = raw_deck.copy()
    export_table["UPB"] = export_table["upb"].map(lambda x: fmt_money(x, decimals=0))
    export_table["Maturity"] = export_table["maturity_date"].map(fmt_date)
    export_table["Next Payment"] = export_table["next_payment_date"].map(fmt_date)
    export_table["DPD"] = export_table["days_past_due"].map(fmt_int)
    export_table = export_table[
        [
            "sheet",
            "deal_number",
            "deal_name",
            "status",
            "owner",
            "UPB",
            "Maturity",
            "Next Payment",
            "DPD",
            "pressure_score",
            "pressure_bucket",
        ]
    ].rename(
        columns={
            "sheet": "Type",
            "deal_number": "Deal #",
            "deal_name": "Deal Name",
            "status": "Status",
            "owner": "Owner",
            "pressure_score": "Pressure",
            "pressure_bucket": "Bucket",
        }
    )
    st.dataframe(export_table.sort_values(["Pressure", "Deal Name"], ascending=[False, True]), use_container_width=True, hide_index=True)


def load_input_file() -> Tuple[Optional[bytes], str]:
    with st.sidebar:
        st.header("Workbook")
        uploaded = st.file_uploader("Upload Portfolio Overview workbook", type=["xlsx"])
        if uploaded is not None:
            return uploaded.getvalue(), uploaded.name

        sample_path = first_existing_file(DEFAULT_SAMPLE_FILES)
        if sample_path:
            st.caption(f"Using local workbook: {sample_path.name}")
            return sample_path.read_bytes(), sample_path.name

        st.info("Upload the weekly Portfolio Overview workbook to get started.")
        return None, "portfolio_overview.xlsx"


def main() -> None:
    st.set_page_config(page_title=APP_TITLE, page_icon="📊", layout="wide", initial_sidebar_state="expanded")
    apply_app_css()

    st.title("📊 Weekly Portfolio Meeting Deck")
    st.caption(
        "A slide-style Streamlit view for Bridge + Term reviews, with live status updates, watchlist cues, and workbook export."
    )

    file_bytes, workbook_name = load_input_file()
    if file_bytes is None:
        st.stop()

    with st.sidebar:
        st.header("Deck controls")
        as_of_date = st.date_input("As-of date", value=dt.date.today())
        include_hidden = st.toggle("Include hidden rows", value=False)

    raw_deck, metadata = load_portfolio_workbook(file_bytes, as_of_date.isoformat(), include_hidden)
    if raw_deck.empty:
        st.error("I could not find visible Bridge/Term deal rows in that workbook.")
        st.stop()

    overrides = ensure_override_store()
    deck = apply_overrides(raw_deck, overrides)

    with st.sidebar:
        st.header("Filters")
        type_options = sorted(deck["sheet"].dropna().unique().tolist())
        selected_types = st.multiselect("Loan type", options=type_options, default=type_options)

        owner_options = sorted([v for v in deck["owner"].dropna().unique().tolist() if norm_text(v)])
        selected_owners = st.multiselect("Owner / Point Person", options=owner_options, default=[])

        status_options = sorted([v for v in deck["status"].dropna().unique().tolist() if norm_text(v)])
        selected_statuses = st.multiselect("Status", options=status_options, default=[])

        watch_only = st.toggle("Watchlist only", value=False)
        min_pressure = st.slider("Minimum pressure score", min_value=0, max_value=100, value=0, step=5)
        sort_by = st.selectbox(
            "Sort order",
            options=["Meeting deck", "Highest pressure", "Nearest maturity", "Largest UPB", "Most delinquent"],
        )

    filtered = deck.copy()
    if selected_types:
        filtered = filtered[filtered["sheet"].isin(selected_types)]
    if selected_owners:
        filtered = filtered[filtered["owner"].isin(selected_owners)]
    if selected_statuses:
        filtered = filtered[filtered["status"].isin(selected_statuses)]
    if watch_only:
        filtered = filtered[filtered["watch_flag"]]
    filtered = filtered[filtered["pressure_score"] >= min_pressure]

    if sort_by == "Highest pressure":
        filtered = filtered.sort_values(["pressure_score", "upb"], ascending=[False, False])
    elif sort_by == "Nearest maturity":
        filtered = filtered.sort_values(["days_to_maturity", "upb"], ascending=[True, False])
    elif sort_by == "Largest UPB":
        filtered = filtered.sort_values(["upb", "pressure_score"], ascending=[False, False])
    elif sort_by == "Most delinquent":
        filtered = filtered.sort_values(["days_past_due", "pressure_score"], ascending=[False, False])
    else:
        filtered = filtered.sort_values(["deck_rank", "upb"], ascending=[True, False])

    filtered = filtered.reset_index(drop=True)
    status_suggestions = available_status_suggestions(deck)

    overview_tab, presentation_tab, log_tab = st.tabs(["Overview", "Presentation Mode", "Meeting Log"])

    with overview_tab:
        render_overview_tab(filtered, as_of_date, metadata)
    with presentation_tab:
        render_presentation_tab(filtered, status_suggestions)
    with log_tab:
        render_meeting_log_tab(raw_deck=deck, displayed_deck=filtered, file_bytes=file_bytes, workbook_name=workbook_name)


if __name__ == "__main__":
    main()
