"""Current loan facts, workbook records, supplemental extracts and meeting changes."""
from __future__ import annotations

import datetime as dt
import difflib
import logging
import os
from collections import deque
from pathlib import Path
import re
from dataclasses import dataclass
from urllib.parse import quote, urlparse

import numpy as np
import pandas as pd
import streamlit as st

import landgorilla_live as lg
from portfolio_cockpit import number, sold_loan, text

LOGGER = logging.getLogger(__name__)


def canonical(value):
    value = lg.norm_text(value)
    return str(int(value)) if value.isdigit() else value.casefold()


def safe_date(value):
    if isinstance(value, (bool, int, float, np.number)) or not text(value):
        return None
    try:
        stamp = pd.to_datetime(value, errors="coerce", utc=True)
        return stamp.date() if pd.notna(stamp) and 1900 <= stamp.year <= 2199 else None
    except (TypeError, ValueError, OverflowError):
        return None


def money(value):
    n = number(value)
    if n is None:
        return "-"
    return f"${n / 1_000_000:,.1f}MM" if abs(n) >= 1_000_000 else f"${n:,.0f}"


def parse_supplemental(workbook):
    """Read visible deal-keyed extracts during the app's single workbook parse pass."""
    result = {}
    keys = {"deal number", "opportunity deal loan number", "loan number"}
    excluded = {"bridge", "term", "loan modifications"}
    for ws in workbook.worksheets:
        if ws.sheet_state != "visible" or ws.title.startswith("_") or ws.title.strip().casefold() in excluded:
            continue
        header = None
        width = min(ws.max_column, 4096)
        for rownum, values in enumerate(ws.iter_rows(max_row=min(ws.max_row, 60), max_col=width, values_only=True), 1):
            if any(text(v).casefold() in keys for v in values):
                header = (rownum, values)
                break
        if header is None:
            continue
        rownum, values = header
        last = max((i + 1 for i, value in enumerate(values) if text(value)), default=0)
        labels, seen = [], {}
        for value in values[:last]:
            label = text(value) or "Unnamed"
            seen[label] = seen.get(label, 0) + 1
            labels.append(label if seen[label] == 1 else f"{label} ({seen[label]})")
        key_col = next(i for i, label in enumerate(labels) if label.casefold() in keys)
        rows, blanks = [], 0
        for values in ws.iter_rows(min_row=rownum + 1, max_row=min(ws.max_row, rownum + 50_000), max_col=last, values_only=True):
            if not any(text(v) for v in values):
                blanks += 1
                if blanks >= 500:
                    break
                continue
            blanks = 0
            if not text(values[key_col]):
                continue
            rows.append(list(values))
        frame = pd.DataFrame(rows, columns=labels)
        for col in frame.columns:
            if re.search(r"date|modified|timestamp", col, re.I):
                frame[col] = pd.to_datetime(frame[col].map(safe_date), errors="coerce")
        frame["_deal_match"] = frame[labels[key_col]].map(canonical) if len(frame) else pd.Series(dtype=str)
        result[ws.title] = frame
    return result


def dated_headers(columns, reference):
    result = {"UPB": [], "NPL": []}
    for col in columns:
        match = re.fullmatch(r"(\d{1,2})/(\d{1,2})(?:/(\d{2,4}))?\s+(UPB|NPL)", str(col), re.I)
        if not match:
            continue
        month, day, year, suffix = match.groups()
        year = int(year) if year else reference.year
        year = year + 2000 if year < 100 else year
        try:
            date = dt.date(year, int(month), int(day))
        except ValueError:
            continue
        if date <= reference:
            result[suffix.upper()].append((date, col))
    return {key: sorted(items) for key, items in result.items()}


def comparison_deck(deck, parsed):
    """Attach dated source columns to effective rows by workbook row identity."""
    sources = []
    for sheet, report in parsed.get("sheets", {}).items():
        raw = report.get("raw")
        selected = deck[deck["sheet"] == sheet]
        if not isinstance(raw, pd.DataFrame) or selected.empty or "_excel_row" not in raw:
            continue
        attached = selected[["deal_key", "_excel_row"]].merge(raw, on="_excel_row", how="left")
        sources.append(attached.drop(columns=["_excel_row", "_hidden"], errors="ignore"))
    if not sources:
        return deck
    source = pd.concat(sources, ignore_index=True).set_index("deal_key")
    additions = [c for c in source.columns if c not in deck.columns]
    return deck.set_index("deal_key").join(source[additions]).reset_index()


def _same(a, b):
    return " ".join(text(a).split()) == " ".join(text(b).split())


@st.cache_data(max_entries=8, show_spinner=False)
def compute_changes(deck_effective, history_norm, meeting_date, dated_columns=None, loan_mods=None):
    """Compare each deal against the latest earlier meeting and its own dated columns."""
    as_of = safe_date(meeting_date) or dt.date.today()
    history = history_norm.copy() if isinstance(history_norm, pd.DataFrame) else pd.DataFrame()
    date_col = next((c for c in ("meeting_date", "Meeting Date") if c in history), None)
    prior_date, prior = None, pd.DataFrame()
    if date_col:
        history["_day"] = history[date_col].map(safe_date)
        previous = [d for d in history["_day"] if d is not None and d < as_of]
        if previous:
            prior_date = max(previous)
            prior = history[history["_day"] == prior_date]
    prior_by = {}
    for row in prior.to_dict("records"):
        key = text(row.get("deal_key")) or f"{text(row.get('sheet'))}::{lg.norm_text(row.get('deal_number'))}"
        prior_by[key] = row
    result = []
    for row in deck_effective.to_dict("records"):
        key = text(row.get("deal_key"))
        previous = prior_by.get(key)
        entry = {"deal_key": key, "prior_date": prior_date, "has_prior_meeting": prior_date is not None,
                 "no_status_last_meeting": prior_date is not None and previous is None,
                 "upb_prev": None, "upb_delta": None, "npl_prev": None, "npl_changed": False,
                 "commentary_same_count": 0, "mod_completed_since": False}
        for field in ("status", "owner", "commentary"):
            entry["prior_" + field] = text((previous or {}).get(field))
            entry[field + "_changed"] = previous is not None and not _same(row.get(field), previous.get(field))
        if date_col and not history.empty:
            lookup = history.get("deal_key", pd.Series("", index=history.index)).astype(str) == key
            streak = history[lookup].sort_values("_day", ascending=False)
            count = 0
            for old in streak.to_dict("records"):
                if old["_day"] is None or old["_day"] >= as_of:
                    continue
                if not _same(row.get("commentary"), old.get("commentary")):
                    break
                count += 1
            entry["commentary_same_count"] = count
        headers = dated_headers((dated_columns or {}).get(row.get("sheet"), row.keys()), as_of)
        for suffix in ("UPB", "NPL"):
            # Values for the other portfolio sheet are missing after concatenation.
            own = headers[suffix] if dated_columns and row.get("sheet") in dated_columns else [(day, col) for day, col in headers[suffix] if row.get(col) is not None and not pd.isna(row.get(col))]
            if len(own) < 2:
                continue
            _, prev_col = own[-2]
            old_value = row.get(prev_col)
            if suffix == "UPB":
                entry["upb_prev"] = number(old_value)
                current = number(row.get("upb"))
                if current is not None and entry["upb_prev"] is not None:
                    entry["upb_delta"] = current - entry["upb_prev"]
            else:
                entry["npl_prev"] = text(old_value)
                entry["npl_changed"] = not _same(old_value, row.get("npl_raw"))
        if prior_date is not None and isinstance(loan_mods, pd.DataFrame) and not loan_mods.empty:
            mod_key = "deal_number_key" if "deal_number_key" in loan_mods else "Deal Number"
            if mod_key in loan_mods:
                mods = loan_mods[loan_mods[mod_key].map(canonical) == canonical(row.get("deal_number"))]
                for mod in mods.to_dict("records"):
                    if mod.get("_mod_inactive") is True or not re.search(r"complet|effective", text(mod.get("Status")), re.I):
                        continue
                    for date_column in ("System Task Completed Date", "Mod Effective Date"):
                        day = safe_date(mod.get(date_column))
                        if day is not None and prior_date < day <= as_of:
                            entry["mod_completed_since"] = True
                            entry["mod_completed_date"] = day
        result.append(entry)
    frame = pd.DataFrame(result)
    current_keys = set(deck_effective["deal_key"]) if "deal_key" in deck_effective else set()
    frame.attrs["dropped_deals"] = [row for key, row in prior_by.items() if key not in current_keys]
    return frame


@dataclass
class LoanToday:
    sentences: list
    facts: dict


def build_loan_today(row, changes_row=None, mods_for_deal=None, checks=(), *, as_of, meeting_date=None):
    """Source-labeled sentences without interpreting stale servicing fields as delinquency."""
    sentences = []
    sold = sold_loan(row)
    dpd, upb, commitment = (number(row.get(k)) for k in ("days_past_due", "upb", "commitment"))
    if sold:
        sentences.append("Sold loan — servicing data may lag.")
    elif dpd is not None:
        sentences.append(f"{int(dpd)} DPD" + (f" (NPL {text(row.get('npl_raw'))})" if dpd >= 30 and text(row.get("npl_raw")) else "") + ".")
    if upb is not None:
        sentence = f"Workbook UPB {money(upb)}"
        if commitment is not None and commitment > 0:
            sentence += f" of {money(commitment)} commitment"
            funded = number(row.get("funded_amount"))
            if row.get("sheet") == "Bridge" and funded is not None:
                sentence += f" ({funded / commitment:.0%} funded)"
        sentences.append(sentence + ".")
    maturity, payment = safe_date(row.get("maturity_date")), safe_date(row.get("next_payment_date"))
    if maturity and not sold:
        sentences.append(f"Workbook maturity {maturity:%m/%d/%Y}.")
    if payment:
        days = (payment - as_of).days
        if days < -30 and (sold or dpd in (0, None)):
            sentences.append(f"Next payment date {payment:%m/%d/%Y} looks stale; confirm the servicer data.")
        elif days < 0 and dpd is not None and dpd > 0 and not sold:
            sentences.append(f"Payment {abs(days)} days overdue; workbook reports {int(dpd)} DPD.")
        elif days >= 0 and not sold:
            sentences.append(f"Next payment {payment:%m/%d/%Y} (in {days} days).")
    count = int(number(row.get("loan_mod_count")) or 0)
    sentences.append(f"{count} loan modification record(s) on file." if count else "No loan modifications on file.")
    status = text(row.get("status"))
    if status:
        sentences.append(f"Current meeting status: {status}.")
    return LoanToday(sentences[:6], {"upb": upb, "commitment": commitment, "sold": sold})


def raw_record(row, parsed):
    raw = ((parsed.get("sheets") or {}).get(text(row.get("sheet"))) or {}).get("raw")
    if not isinstance(raw, pd.DataFrame) or raw.empty:
        return {}
    excel_row = row.get("_excel_row")
    if "_excel_row" in raw and excel_row is not None:
        found = raw[raw["_excel_row"] == excel_row]
    else:
        found = raw.iloc[0:0]
    return {str(k): v for k, v in (found.iloc[0].to_dict() if not found.empty else {}).items() if not str(k).startswith("_")}


def format_field(header, value, column=None):
    if not text(value):
        return "-"
    if isinstance(value, (dt.date, dt.datetime, pd.Timestamp)):
        return value.strftime("%m/%d/%Y")
    n = number(value)
    if n is not None and re.search(r"upb|amount|commitment|balance|requested|approved", header, re.I):
        return lg.fmt_money_exact(n)
    if n is not None and re.search(r"%|rate|ltv", header, re.I):
        maximum = pd.to_numeric(column, errors="coerce").abs().max() if column is not None else abs(n)
        scale = 100 if re.search(r"%|rate", header, re.I) and pd.notna(maximum) and maximum <= 1.5 else 1
        return f"{n * scale:g}%"
    return text(value)


def quick_links(record, row, secrets=None):
    result = []
    for key, value in record.items():
        url = text(value)
        if re.search(r"link|url", key, re.I) and valid_url(url):
            result.append(("Open in Salesforce" if re.search(r"salesforce|force\.com", key + " " + url, re.I) else key, url))
    try:
        config = dict(st.secrets.get("links", {})) if secrets is None else dict(secrets)
    except Exception:
        config = {}
    template = config.get("salesforce_deal_url")
    if template:
        try:
            url = str(template).format(deal_number=quote(text(row.get("deal_number")), safe=""))
            if valid_url(url):
                result.append(("Open in Salesforce", url))
        except (ValueError, KeyError, IndexError):
            LOGGER.warning("Invalid Salesforce link template; link omitted.")
    servicers = config.get("servicers") or {}
    for name, url in servicers.items():
        if str(name).casefold() == text(row.get("servicer")).casefold() and valid_url(str(url)):
            result.append(("Open servicer portal", str(url)))
    return list(dict.fromkeys(result))


def valid_url(value):
    try:
        url = urlparse(value)
        return url.scheme in {"http", "https"} and bool(url.hostname) and not url.username and not url.password
    except (ValueError, TypeError):
        return False


def commentary_diff(previous, current):
    left, right = text(previous).split(), text(current).split()
    matcher = difflib.SequenceMatcher(None, left, right)
    parts = []
    for kind, a, b, c, d in matcher.get_opcodes():
        old, new = lg.md_escape(" ".join(left[a:b])), lg.md_escape(" ".join(right[c:d]))
        if kind == "equal":
            parts.append(new)
        else:
            if old:
                parts.append(f":red[~~{old}~~]")
            if new:
                parts.append(f":green[{new}]")
    return " ".join(parts)


def prepare_context(deck, parsed):
    as_of = st.session_state.get("as_of_date")
    reference = safe_date(as_of) or safe_date(parsed.get("reference_date")) or dt.date.today()
    history = st.session_state.get("status_history", pd.DataFrame())
    columns = {sheet: list(report["raw"].columns) for sheet, report in parsed.get("sheets", {}).items() if isinstance(report.get("raw"), pd.DataFrame)}
    source_deck = comparison_deck(deck, parsed)
    changes = compute_changes(source_deck, history, reference, columns, loan_mods=parsed.get("loan_mods"))
    return {"lg": lg.prepare_landgorilla_context(source_deck), "changes": changes, "source_deck": source_deck,
            "as_of": reference, "history": history, "supplemental": parsed.get("supplemental") or {}}


def render_current(row, deck, parsed, lg_ctx):
    """Render selected-row data only; workbook reads stay inside the cached parser."""
    ctx = lg_ctx or {}
    seen_links = set()
    as_of = ctx.get("as_of") or safe_date(parsed.get("reference_date")) or dt.date.today()
    changes = ctx.get("changes")
    change = {}
    if isinstance(changes, pd.DataFrame) and not changes.empty:
        found = changes[changes["deal_key"] == row.get("deal_key")]
        change = found.iloc[0].to_dict() if not found.empty else {}
    record = raw_record(row, parsed)
    raw = ((parsed.get("sheets") or {}).get(text(row.get("sheet"))) or {}).get("raw")
    def column(header):
        return raw[header] if isinstance(raw, pd.DataFrame) and header in raw else None
    today = build_loan_today(row, change, as_of=as_of)
    file_number, result = lg.lg_result_for_row(row, ctx.get("lg"))
    with st.container(border=True):
        st.markdown(f"##### As it stands today · {as_of:%m/%d/%Y}")
        if result and result.get("state") == "ok":
            overview = result["overview"]
            derived = lg.lg_derived(overview, lg.lg_today())
            facts = []
            if derived["project_pct"] is not None:
                facts.append(f"{derived['project_pct']:.0f}% built")
            if derived["duration_pct"] is not None:
                facts.append(f"{derived['duration_pct']:.0f}% of the schedule elapsed")
            if derived["to_finish"] is not None:
                facts.append(f"{money(derived['to_finish'])} left to fund")
            if facts:
                freshness = "last successful fetch" if result.get("stale") else "live"
                st.write(f"Land Gorilla ({freshness}): " + "; ".join(facts) + ".")
        st.write(" ".join(today.sentences))
        st.badge(f"Workbook as of {as_of:%m/%d}", color="gray")
        if change.get("prior_date"):
            st.badge(f"History: last meeting {change['prior_date']:%m/%d}", color="gray")
        mods = parsed.get("loan_mods")
        if isinstance(mods, pd.DataFrame) and not mods.empty:
            candidates = mods
            if "deal_number_key" in mods:
                matching = mods[mods["deal_number_key"].map(canonical) == canonical(row.get("deal_number"))]
                if not matching.empty:
                    candidates = matching
            modified = [safe_date(value) for field in ("Last Modified Date", "System Modstamp")
                        if field in candidates for value in candidates[field]]
            modified = [day for day in modified if day is not None]
            if modified:
                st.badge(f"Salesforce mods: modified {max(modified):%m/%d}", color="gray")
        documents = box_documents(row)
        if not documents.empty:
            latest = documents["modified"].max()
            st.badge(f"Box: latest file {latest:%m/%d}", color="gray")
        for header in st.session_state.get("pinned_fields", []):
            st.caption(f"{header}: {format_field(header, record.get(header), column(header))}")
        for label, url in quick_links(record, row):
            if (label, url) not in seen_links:
                seen_links.add((label, url))
                st.link_button(label, url)
        headers = dated_headers(record, as_of)["UPB"]
        values = [number(record.get(col)) for _, col in headers]
        if len(values) >= 3 and all(v is not None for v in values):
            st.metric("Workbook UPB history", money(values[-1]), chart_data=values, chart_type="line", border=True)
    lg.render_landgorilla_panel(row, ctx.get("lg"))
    render_changes(row, change)
    history = ctx.get("history")
    if isinstance(history, pd.DataFrame) and "deal_key" in history:
        records = history[history["deal_key"] == row.get("deal_key")]
        if not records.empty:
            with st.expander("Deal history"):
                cols = [c for c in ("meeting_date", "status", "owner", "commentary") if c in records]
                st.dataframe(records[cols].sort_values("meeting_date", ascending=False), hide_index=True, width="stretch")
    for name, frame in ctx.get("supplemental", {}).items():
        matches = frame[frame["_deal_match"] == canonical(row.get("deal_number"))].drop(columns="_deal_match")
        if matches.empty:
            continue
        date_cols = [c for c in matches if pd.api.types.is_datetime64_any_dtype(matches[c])]
        if date_cols:
            matches = matches.sort_values(date_cols[0], ascending=False)
        with st.expander(f"From '{name}' · {len(matches)} row(s)"):
            st.dataframe(matches, hide_index=True, width="stretch")
            for item in matches.to_dict("records"):
                for label, url in quick_links(item, row, secrets={}):
                    if (label, url) not in seen_links:
                        seen_links.add((label, url))
                        st.link_button(label, url)
    render_box_documents(row)
    with st.expander(f"All workbook fields ({len(record)})"):
        query = st.text_input("Search workbook fields", key="workbook_field_search")
        rows = [{"Field": key, "Value": format_field(key, value, column(key))} for key, value in record.items()
                if not query or query.casefold() in (key + " " + text(value)).casefold()]
        st.dataframe(pd.DataFrame(rows, columns=["Field", "Value"]), hide_index=True, width="stretch")


def render_changes(row, change):
    if not change.get("has_prior_meeting"):
        st.badge("No prior meeting", color="gray")
    elif change.get("no_status_last_meeting"):
        st.badge("No status recorded last meeting", color="violet")
    else:
        if change.get("status_changed"):
            st.badge(f"Status: {change['prior_status']} → {text(row.get('status'))}", color="blue")
        if change.get("owner_changed"):
            st.badge("Owner changed", color="violet")
        if change.get("commentary_same_count", 0) >= 3:
            st.badge(f"Commentary unchanged {change['commentary_same_count']} mtgs", color="orange")
    if number(change.get("upb_delta")) is not None:
        st.badge(f"UPB {lg.fmt_money_exact(change['upb_delta'])} vs prior column", color="gray")
    if change.get("mod_completed_since"):
        st.badge(f"Mod completed {change['mod_completed_date']:%m/%d}", color="blue")
    if change.get("npl_changed"):
        st.badge(f"NPL: {change.get('npl_prev') or 'blank'} → {text(row.get('npl_raw'))}", color="red")
    if change.get("prior_date"):
        with st.expander(f"Since last meeting ({change['prior_date']:%m/%d})"):
            st.caption(f"Prior status: {change.get('prior_status') or '-'} · Prior owner: {change.get('prior_owner') or '-'}")
            st.markdown(commentary_diff(change.get("prior_commentary"), row.get("commentary")) or "No commentary change.")
            if st.toggle("Show last week's full text", key=f"lastnotes::{row.get('deal_key')}"):
                st.write(change.get("prior_commentary") or "No commentary recorded.")


def render_sources(deck, parsed, ctx):
    with st.sidebar.expander("Data sources"):
        headers = sorted({c for report in parsed.get("sheets", {}).values()
                          for c in (report.get("raw").columns if isinstance(report.get("raw"), pd.DataFrame) else [])
                          if not str(c).startswith("_")})
        st.session_state["pinned_fields"] = [p for p in st.session_state.get("pinned_fields", []) if p in headers]
        st.multiselect("Pinned fields", headers, key="pinned_fields", persist_state="session")
        st.toggle("Show Box documents", key="box_enabled", persist_state="session")
        if st.session_state.get("box_enabled"):
            st.session_state.setdefault("box_root", os.environ.get("WPO_BOX_ROOT") or str(Path.home() / "Box"))
            st.text_input("Box folder", key="box_root", persist_state="session")
            root = Path(st.session_state["box_root"]).expanduser()
            if not root.is_dir():
                st.caption("Box Drive folder not found on this server — this works when the app runs on your PC with Box Drive.")
            else:
                try:
                    folders = sorted(entry.name for entry in root.iterdir() if entry.is_dir() and not entry.name.startswith("."))
                except OSError:
                    folders = []
                st.session_state.setdefault("box_subfolders", [name for name in ("Asset Management", "Bridge Loans", "Term Lending", "DSCR Lending") if name in folders])
                st.session_state["box_subfolders"] = [name for name in st.session_state["box_subfolders"] if name in folders]
                st.multiselect("Search in", folders, key="box_subfolders", persist_state="session")
                st.number_input("Max folder depth", min_value=0, max_value=12, value=4, key="box_depth")
                st.button("Refresh Box index", on_click=scan_box_index.clear)
        known = set(deck["deal_number"].map(canonical))
        for name, frame in ctx.get("supplemental", {}).items():
            unknown = sorted(set(frame["_deal_match"]) - known)
            st.caption(f"{name}: {len(frame)} rows · {frame['_deal_match'].nunique() - len(unknown)} matched deals")
            if unknown:
                st.caption("Unmatched deal numbers: " + ", ".join(unknown))
    lg.render_landgorilla_sidebar(ctx.get("lg"), ctx.get("source_deck", deck))

# Local Box Drive metadata is opt-in and stays in memory.
@st.cache_data(ttl=900, max_entries=4, show_spinner=False)
def scan_box_index(root, subfolders=(), max_depth=4, max_entries=50_000):
    """Bounded breadth-first metadata scan; file contents and symlinks are never opened."""
    root_path = Path(root).expanduser().resolve()
    columns = ["path", "rel_path", "name", "is_dir", "modified", "size", "ext"]
    rows, queue, errors = [], deque(), 0
    if not root_path.is_dir():
        return pd.DataFrame(columns=columns)
    selected = list(subfolders) or [""]
    for name in selected:
        target = (root_path / name).resolve()
        if target.is_relative_to(root_path) and target.is_dir():
            queue.append((target, 0))
    while queue and len(rows) < max_entries:
        directory, depth = queue.popleft()
        try:
            with os.scandir(directory) as entries:
                for entry in entries:
                    if len(rows) >= max_entries:
                        break
                    if entry.name.startswith((".", "~$")) or entry.name.casefold() == "desktop.ini" or entry.is_symlink():
                        continue
                    try:
                        stat = entry.stat(follow_symlinks=False)
                        if getattr(stat, "st_file_attributes", 0) & 6:
                            continue
                        is_dir = entry.is_dir(follow_symlinks=False)
                        path = Path(entry.path)
                        rows.append({"path": str(path), "rel_path": str(path.relative_to(root_path)),
                                     "name": entry.name, "is_dir": is_dir,
                                     "modified": pd.Timestamp(stat.st_mtime, unit="s"), "size": stat.st_size,
                                     "ext": "" if is_dir else path.suffix.lower()})
                        if is_dir and depth < max_depth:
                            queue.append((path, depth + 1))
                    except OSError:
                        errors += 1
        except OSError:
            errors += 1
    frame = pd.DataFrame(rows, columns=columns)
    frame.attrs["truncated"] = bool(queue) or len(rows) >= max_entries
    frame.attrs["unreadable_entries"] = errors
    return frame


def match_box_documents(index, *, deal_number, deal_name="", borrower="", account="", limit=15):
    if index.empty:
        return index.copy()
    number_key = canonical(deal_number)
    number_pattern = re.compile(r"(?<!\d)" + re.escape(number_key) + r"(?!\d)") if number_key.isdigit() else None
    names = [(label, re.sub(r"[^a-z0-9]", "", text(value).casefold()))
             for label, value in (("Deal name", deal_name), ("Borrower", borrower), ("Account", account))]
    rows = []
    for record in index.to_dict("records"):
        if record["is_dir"]:
            continue
        components = Path(record["rel_path"]).parts
        rule = None
        if number_pattern and any(number_pattern.search(part) for part in components):
            rule = "Deal number"
        else:
            normalized = [re.sub(r"[^a-z0-9]", "", part.casefold()) for part in components]
            rule = next((label for label, value in names if len(value) >= 6 and any(value in part for part in normalized)), None)
        if rule:
            rows.append({**record, "Matched by": rule})
    if not rows:
        return pd.DataFrame(columns=[*index.columns, "Matched by"])
    return pd.DataFrame(rows).sort_values("modified", ascending=False).head(limit)


def box_documents(row):
    if not st.session_state.get("box_enabled"):
        return pd.DataFrame()
    root = st.session_state.get("box_root") or os.environ.get("WPO_BOX_ROOT") or str(Path.home() / "Box")
    if not Path(root).expanduser().is_dir():
        return pd.DataFrame()
    index = scan_box_index(root, tuple(st.session_state.get("box_subfolders", [])),
                           int(st.session_state.get("box_depth", 4)))
    return match_box_documents(index, deal_number=row.get("deal_number"), deal_name=row.get("deal_name"),
                               borrower=row.get("borrower"), account=row.get("account_display"))


def render_box_documents(row):
    documents = box_documents(row)
    if documents.empty:
        return
    with st.expander(f"Box documents ({len(documents)})"):
        display = documents[["name", "rel_path", "modified", "size", "Matched by"]].rename(
            columns={"name": "File", "rel_path": "Folder / file", "modified": "Modified", "size": "Bytes"})
        st.dataframe(display, hide_index=True, width="stretch")
        for path in documents["path"]:
            st.code(path, language=None)
