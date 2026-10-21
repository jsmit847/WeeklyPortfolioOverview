"""Explainable portfolio attention rules, agenda ordering and readiness views."""
from __future__ import annotations

import datetime as dt
import io
import math
from typing import Any, Mapping

import pandas as pd
import streamlit as st


MATURITY_BUCKETS = ("Past maturity", "0–30 days", "31–60 days", "61–90 days", "91–180 days", "181–365 days", ">1 year", "Unknown")
WEIGHTS = {"30+ DPD": 100, "Past maturity": 80, "Payment overdue": 50, "Maturing soon": 40,
           "Update reminder": 20, "Active modification": 15, "Data check": 10}
ORDER_OPTIONS = ("Workbook", "By owner", "Attention first", "Maturity", "Largest balance")


def number(value: Any) -> float | None:
    try:
        parsed = float(value)
        return parsed if math.isfinite(parsed) else None
    except (TypeError, ValueError):
        return None


def text(value: Any) -> str:
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass
    return str(value).strip()


def sold_loan(row: Mapping) -> bool:
    labels = " ".join(text(row.get(k)).casefold() for k in ("portfolio", "status", "segment"))
    return any(term in labels for term in ("sold", "paid off", "paid-off", "liquidated"))


def consistency_checks(row: Mapping) -> list[dict]:
    """A stale payment date is not evidence of delinquency, especially after sale."""
    checks = []
    balance, dpd = number(row.get("upb")), number(row.get("days_past_due"))
    payment = number(row.get("days_to_next_payment"))
    if sold_loan(row):
        checks.append({"code": "SOLD_LOAN", "detail": "Sold or paid-off loan; servicing fields may lag the disposition."})
    if balance is not None and balance <= 0:
        checks.append({"code": "NO_BALANCE", "detail": "Workbook shows no outstanding balance."})
    if payment is not None and payment < -30 and (sold_loan(row) or dpd == 0 or (balance is not None and balance <= 0)):
        checks.append({"code": "STALE_PAYMENT_DATE", "detail": "Date not current? The payment date conflicts with the loan status, balance or reported DPD."})
    if number(row.get("maturity_date")) is None and not text(row.get("maturity_date")):
        checks.append({"code": "MISSING_MATURITY", "detail": "No usable maturity date in the workbook."})
    commitment = number(row.get("commitment"))
    if balance is not None and commitment is not None and commitment > 0 and balance > commitment * 1.01:
        checks.append({"code": "BALANCE_OVER_COMMITMENT", "detail": "Workbook UPB exceeds the commitment; confirm amounts and source dates."})
    return checks


def compute_flags(deck: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for row in deck.to_dict("records"):
        def flag(rule: str, detail: str):
            rows.append({"deal_key": row["deal_key"], "Deal": text(row.get("deal_name")),
                         "Deal #": text(row.get("deal_number")), "Type": text(row.get("sheet")),
                         "Owner": text(row.get("owner")), "Rule": rule, "Detail": detail,
                         "Weight": WEIGHTS.get(rule, 0), "UPB": number(row.get("upb"))})
        dpd, maturity, payment = (number(row.get(k)) for k in ("days_past_due", "days_to_maturity", "days_to_next_payment"))
        balance = number(row.get("upb"))
        checks = consistency_checks(row)
        stale = any(c["code"] == "STALE_PAYMENT_DATE" for c in checks)
        if dpd is not None and dpd >= 30 and not sold_loan(row) and (balance is None or balance > 0):
            flag("30+ DPD", f"Workbook reports {dpd:g} days past due.")
        if maturity is not None and maturity < 0:
            flag("Past maturity", f"Workbook maturity was {abs(int(maturity))} days before data as-of.")
        elif maturity is not None and maturity <= 30:
            flag("Maturing soon", f"Workbook maturity is in {int(maturity)} days.")
        if payment is not None and payment < 0 and not stale and not sold_loan(row) and dpd is not None and dpd > 0 and (balance is None or balance > 0):
            flag("Payment overdue", "Past payment date and reported DPD both indicate a servicing issue.")
        if row.get("status_needs_update_prompt") is True or str(row.get("status_needs_update_prompt")).lower() == "true":
            if not row.get("reminder_confirmed", False) and not row.get("reminder_addressed", False):
                flag("Update reminder", "Confirm that the repeated meeting update is still accurate.")
        if (number(row.get("loan_mod_active_count")) or 0) > 0:
            flag("Active modification", text(row.get("loan_mod_latest_summary")))
        for check in checks:
            if check["code"] not in ("SOLD_LOAN", "NO_BALANCE"):
                flag("Data check", check["detail"])
    return pd.DataFrame(rows, columns=["deal_key", "Deal", "Deal #", "Type", "Owner", "Rule", "Detail", "Weight", "UPB"])


def order_deck(deck: pd.DataFrame, mode: str, flags: pd.DataFrame | None = None) -> pd.DataFrame:
    if deck.empty or mode == "Workbook":
        return deck.copy()
    result = deck.copy()
    result["_stable_order"] = range(len(result))
    if mode == "By owner":
        result["_owner_order"] = result["owner"].map(lambda x: text(x).casefold() or "\uffff")
        result = result.sort_values(["_owner_order", "_stable_order"], kind="stable")
    elif mode == "Attention first":
        flags = compute_flags(result) if flags is None else flags
        scores = flags.groupby("deal_key")["Weight"].sum() if not flags.empty else pd.Series(dtype=float)
        result["_attention_score"] = result["deal_key"].map(scores).fillna(0)
        result = result.sort_values(["_attention_score", "_stable_order"], ascending=[False, True], kind="stable")
    elif mode == "Maturity":
        result = result.sort_values(["days_to_maturity", "_stable_order"], kind="stable", na_position="last")
    elif mode == "Largest balance":
        result = result.sort_values(["upb", "_stable_order"], ascending=[False, True], kind="stable", na_position="last")
    return result.drop(columns=[c for c in result if c.startswith("_owner_order") or c in ("_attention_score", "_stable_order")]).reset_index(drop=True)


def maturity_bucket(days: Any) -> str:
    value = number(days)
    if value is None:
        return "Unknown"
    if value < 0:
        return "Past maturity"
    for upper, label in zip((30, 60, 90, 180, 365), MATURITY_BUCKETS[1:6]):
        if value <= upper:
            return label
    return ">1 year"


def maturity_wall(deck: pd.DataFrame) -> pd.DataFrame:
    if deck.empty:
        return pd.DataFrame(columns=["Bucket", "Type", "UPB"])
    frame = deck.copy()
    frame["Bucket"] = pd.Categorical(frame["days_to_maturity"].map(maturity_bucket), categories=MATURITY_BUCKETS, ordered=True)
    return frame.groupby(["Bucket", "sheet"], observed=False)["upb"].sum().reset_index().rename(columns={"sheet": "Type", "upb": "UPB"})


def apply_reminder_rule(deck: pd.DataFrame, history: pd.DataFrame, meeting_date: dt.date, mode: str, threshold: int = 3) -> pd.DataFrame:
    if mode != "Status + commentary" or deck.empty:
        return deck
    result = deck.copy()
    eligible = set()
    if history is not None and not history.empty:
        prior = history.copy()
        dates = pd.to_datetime(prior["meeting_date"], errors="coerce", utc=True)
        prior = prior.loc[dates.dt.date < meeting_date].sort_values("meeting_date")
        for key, records in prior.groupby("deal_key", sort=False):
            records = records.drop_duplicates("meeting_date", keep="last").tail(threshold - 1)
            current = result[result["deal_key"] == key]
            if current.empty or len(records) < threshold - 1:
                continue
            commentary = " ".join(text(current.iloc[0].get("commentary")).split()).casefold()
            if all(" ".join(text(v).split()).casefold() == commentary for v in records["commentary"]):
                eligible.add(key)
    result["status_needs_update_prompt"] = result["status_needs_update_prompt"].fillna(False) & result["deal_key"].isin(eligible)
    return result


def preflight(deck: pd.DataFrame, parsed: dict, data_as_of: dt.date, today: dt.date) -> pd.DataFrame:
    notes = []
    def add(level, check, count, detail):
        notes.append({"Severity": level, "Check": check, "Count": int(count), "Detail": detail})
    add("Info", "Loaded agenda", len(deck), "Workbook rows after hidden-row selection; filters may narrow the displayed agenda.")
    age = (today - data_as_of).days
    if age > 8:
        add("Warning", "Workbook freshness", age, f"Data as-of is {age} days before the meeting date.")
    elif age < 0:
        add("Warning", "Future data date", abs(age), "Data as-of is after the meeting date.")
    for column, label in (("owner", "Missing owners"), ("status", "Missing status"), ("maturity_date", "Missing maturity")):
        if column in deck:
            missing = deck[column].map(lambda x: not text(x)).sum()
            if missing:
                add("Warning", label, missing, "Review missing source values before presenting.")
    if "is_duplicate" in deck and deck["is_duplicate"].sum():
        add("Warning", "Duplicate deal IDs", deck["is_duplicate"].sum(), "Rows have separate presentation keys; duplicates do not enter status-history snapshots.")
    for sheet, report in parsed.get("sheets", {}).items():
        if report.get("error") or not report.get("layout"):
            add("Warning", f"{sheet} sheet", 1, text(report.get("error")) or "No usable portfolio header found.")
    return pd.DataFrame(notes)


def verify_workbook(original: bytes, exported: bytes, edits: dict, app) -> dict:
    """Re-open generated XLSX, verify target edit cells and original portfolio rows."""
    source = app.parse_workbook_bytes(original)
    result = app.parse_workbook_bytes(exported)
    if result.get("error"):
        raise ValueError("The exported workbook could not be read back.")
    before, after = app.build_base_rows(source), app.build_base_rows(result)
    if set(before) != set(after):
        raise ValueError("Portfolio row identities changed during export.")
    for key, item in edits.items():
        if key not in after:
            raise ValueError("An edited row is missing from the exported workbook.")
        for field in ("status", "owner", "commentary"):
            if app._field_clean(field, after[key].get(field)) != app._field_clean(field, item.get(field)):
                raise ValueError(f"An exported {field} does not match the saved meeting edit.")
    return {"rows": len(after), "edits": len(edits), "size": len(exported)}


def render_settings(ctx, app=None):
    with st.sidebar.expander("Presentation settings", icon=":material/tune:"):
        st.selectbox("Agenda order", ORDER_OPTIONS, key="agenda_order", persist_state="session")
        owners = sorted({text(x) for x in ctx.deck.get("owner", []) if text(x)})
        # Prune options before the widget is created when a workbook changes.
        st.session_state["owner_filter"] = [x for x in st.session_state.get("owner_filter", []) if x in owners]
        st.multiselect("Owners", owners, key="owner_filter", persist_state="session")
        st.multiselect("Focus", ("Needs attention", "Modified loans", "Edited this meeting", "Past maturity", "Maturing within 30 days"), key="focus_filters", persist_state="session")
        st.toggle("Compact presentation", key="compact_presentation", persist_state="session")
        st.toggle("Show historical update reminders", key="show_update_reminders", persist_state="session",
                  help="Optional. Repeated workbook text does not mean a loan was not discussed.")
        st.caption("← / → move between deals · E opens the editor. Shortcuts pause while typing.")


def apply_focus(deck: pd.DataFrame, edits: dict) -> pd.DataFrame:
    result = deck.copy()
    owners = st.session_state.get("owner_filter", [])
    if owners:
        result = result[result["owner"].isin(owners)]
    focus = st.session_state.get("focus_filters", [])
    # Multiple focus selections form an intersection, consistent with owner/type/search filters.
    for mode in focus:
        if mode == "Needs attention":
            result = result[result["deal_key"].isin(compute_flags(result)["deal_key"])]
        elif mode == "Modified loans":
            result = result[result["loan_mod_count"].fillna(0) > 0]
        elif mode == "Edited this meeting":
            result = result[result["deal_key"].isin(edits)]
        elif mode == "Past maturity":
            result = result[result["days_to_maturity"] < 0]
        elif mode == "Maturing within 30 days":
            result = result[result["days_to_maturity"].between(0, 30)]
    return order_deck(result, st.session_state.get("agenda_order", "Workbook"))


def render_deal(row, ctx, app=None):
    if st.session_state.get("compact_presentation"):
        st.html("""<style>
        .rt-hero {padding: .75rem 1rem !important;}
        .rt-hero-title {font-size: 1.8rem !important;}
        .rt-panel {padding: .8rem !important;}
        .rt-field {min-height: 50px !important; padding: .5rem !important;}
        .rt-kpi-grid {gap: .4rem !important;}
        .rt-kpi {padding: .7rem !important;}
        </style>""")
    checks = consistency_checks(row)
    if checks:
        with st.expander(f"Source consistency · {len(checks)} checks", expanded=False):
            for check in checks:
                st.caption(check["detail"])


def render_dashboard(ctx, app=None):
    deck = ctx.filtered
    if deck.empty:
        return
    flags = compute_flags(deck)
    with st.expander("Portfolio intelligence", expanded=False, icon=":material/analytics:"):
        cols = st.columns(3)
        cols[0].metric("Needs attention", flags["deal_key"].nunique(), border=True)
        delinquent = flags.loc[flags["Rule"] == "30+ DPD", "deal_key"]
        cols[1].metric("30+ DPD balance", f"${deck.loc[deck['deal_key'].isin(delinquent), 'upb'].sum():,.0f}", border=True)
        cols[2].metric("Unassigned owners", int(deck["owner"].map(lambda x: not text(x)).sum()), border=True)
        tabs = st.tabs(["Attention", "Maturity wall", "Composition", "Readiness"])
        with tabs[0]:
            if flags.empty:
                st.info("No attention rules triggered for this scope.")
            else:
                rules = st.multiselect("Attention rules", sorted(flags["Rule"].unique()), key="attention_rules")
                shown = flags[flags["Rule"].isin(rules)] if rules else flags
                st.dataframe(shown.drop(columns=["Weight"]), hide_index=True, width="stretch", column_config={"deal_key": None, "UPB": st.column_config.NumberColumn(format="dollar")})
                st.download_button("Download attention list", data=safe_csv(shown), file_name="portfolio_attention.csv", mime="text/csv", on_click="ignore")
        with tabs[1]:
            wall = maturity_wall(deck)
            st.bar_chart(wall, x="Bucket", y="UPB", color="Type", stack=True, sort=False)
            buckets = st.multiselect("Maturity buckets", MATURITY_BUCKETS, key="maturity_buckets")
            details = deck[["deal_name", "deal_number", "sheet", "owner", "maturity_date", "upb", "days_to_maturity"]].copy()
            details["Bucket"] = details["days_to_maturity"].map(maturity_bucket)
            if buckets:
                details = details[details["Bucket"].isin(buckets)]
            st.dataframe(details.drop(columns=["days_to_maturity"]), hide_index=True, width="stretch", column_config={"maturity_date": st.column_config.DateColumn("Maturity", format="MM/DD/YYYY"), "upb": st.column_config.NumberColumn("UPB", format="dollar")})
        with tabs[2]:
            dimension = st.selectbox("Group balances by", ["owner", "servicer", "portfolio", "sheet"], key="balance_group")
            grouped = deck.assign(**{dimension: deck[dimension].map(lambda v: text(v) or "Unassigned")}).groupby(dimension)["upb"].sum().sort_values(ascending=False)
            st.bar_chart(grouped)
        with tabs[3]:
            meeting = st.session_state.get("meeting_date", ctx.as_of)
            st.dataframe(preflight(ctx.deck, ctx.parsed, ctx.as_of, meeting), hide_index=True, width="stretch")


def safe_csv(frame: pd.DataFrame) -> bytes:
    frame = frame.copy()
    for column in frame:
        frame[column] = frame[column].map(lambda value: "'" + value if isinstance(value, str) and value.lstrip().startswith(("=", "+", "-", "@")) else value)
    return frame.to_csv(index=False).encode("utf-8-sig")
