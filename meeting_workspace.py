"""Meeting recovery, follow-ups and minutes. Disk journaling is opt-in."""
from __future__ import annotations
import copy
import csv
import datetime as dt
from email.message import EmailMessage
import hashlib
import html
import io
import json
import math
import os
from pathlib import Path
import re
import secrets
import tempfile
import threading
import time
import uuid
import zipfile
from typing import Any, Mapping
import pandas as pd
import streamlit as st

SCHEMA = "wpo-workspace/1"
SNAPSHOT_SCHEMA = "wpo-meeting-snapshot/1"
MAX_BYTES = 5 * 1024 * 1024
MAX_HISTORY = 60
ITEM_TYPES = ("Action", "Decision", "Parking lot", "Note")
OUTCOMES = ("Discussed", "No change", "Updated", "Parked", "Deferred", "Not reached")
ACTION_ITEM_COLUMNS = ["item_id", "created_meeting_date", "sheet", "deal_number", "deal_key", "deal_name", "type", "text", "owner", "due_date", "status", "created_at", "updated_at", "closed_at", "closed_meeting_date", "notes"]
MEETING_LOG_COLUMNS = ["meeting_date", "sheet", "deal_number", "deal_key", "agenda_position", "agenda_order", "outcome", "seconds_on_deal", "status_before", "status_after", "owner_before", "owner_after", "commentary_changed", "park_reason", "items_created", "saved_at"]
MEETING_SUMMARY_COLUMNS = ["meeting_date", "workbook_name", "workbook_sha256_prefix", "agenda_order", "started_at", "ended_at", "duration_min", "attendees", "deals_total", "deals_discussed", "no_change", "updated", "consent_confirmed", "parked", "deferred", "not_reached", "edits", "actions_created", "actions_closed", "open_actions", "overdue_actions", "total_upb", "past_maturity_count", "maturing_30d_count", "dpd30_upb", "reminders", "app_version", "saved_at"]
SNAPSHOT_FIELDS = ("status", "owner", "commentary", "upb", "commitment", "maturity_date", "next_payment_date")

def text(value):
    if value is None:
        return ""
    try:
        if bool(pd.isna(value)):
            return ""
    except (ValueError, TypeError):
        pass
    return str(value).strip()

def safe(value):
    if isinstance(value, Mapping):
        return {str(k): safe(v) for k, v in value.items()}
    if isinstance(value, (list, tuple)):
        return [safe(v) for v in value]
    if isinstance(value, (dt.datetime, dt.date)):
        return value.isoformat()
    if value is None or isinstance(value, (str, bool, int)):
        return value
    if isinstance(value, float):
        return value if math.isfinite(value) else None
    try:
        if bool(pd.isna(value)):
            return None
    except (ValueError, TypeError):
        pass
    if hasattr(value, "item"):
        return safe(value.item())
    return str(value)

def stamp(now=None):
    return dt.datetime.fromtimestamp(time.time() if now is None else now, dt.timezone.utc).isoformat()

def date_text(value):
    value = text(value)
    return dt.date.fromisoformat(value[:10]).isoformat() if value else ""

def new_workspace(workbook_sha, workbook_name, meeting_date):
    if not re.fullmatch(r"[a-fA-F0-9]{64}", str(workbook_sha)):
        raise ValueError("A workbook SHA-256 is required")
    date = date_text(meeting_date)
    if not date:
        raise ValueError("A meeting date is required")
    return {"schema": SCHEMA, "workbook_sha": workbook_sha.lower(), "workbook_name": str(workbook_name),
            "meeting_date": date, "next_meeting": (dt.date.fromisoformat(date) + dt.timedelta(days=7)).isoformat(),
            "overrides": {}, "undo": [], "redo": [], "changes": [], "items": [], "progress": {},
            "clock": {"started_at": None, "paused_at": None, "paused_total_s": 0.0, "ended_at": None},
            "attendees": [], "minutes_sent": False, "archives": [], "saved_at": "",
            "export_fingerprint": "", "exported_at": "", "baseline_ready": False}

def fingerprint(workspace):
    keys = ("workbook_sha", "meeting_date", "next_meeting", "overrides", "items", "changes", "progress", "clock", "attendees", "minutes_sent", "archives")
    content = {k: workspace.get(k) for k in keys}
    return hashlib.sha256(json.dumps(safe(content), sort_keys=True, ensure_ascii=False).encode()).hexdigest()

def checkpoint(workspace, overrides, *, base_rows=None, now=None):
    result = copy.deepcopy(dict(workspace))
    incoming, before = safe(dict(overrides)), result.get("overrides", {})
    if incoming == before:
        result["baseline_ready"] = True
        return result
    if result.get("baseline_ready"):
        result["undo"] = (result["undo"] + [copy.deepcopy(before)])[-MAX_HISTORY:]
        result["redo"] = []
        for key in sorted(set(before) | set(incoming)):
            base = (base_rows or {}).get(key, {})
            left, right = before.get(key, base), incoming.get(key, base)
            for field in ("status", "owner", "commentary"):
                if left.get(field) != right.get(field):
                    result["changes"].append({"at": stamp(now), "deal_key": key, "field": field,
                        "before": left.get(field), "after": right.get(field), "source": "edit"})
        result["changes"] = result["changes"][-5000:]
    result.update(overrides=incoming, baseline_ready=True, saved_at=stamp(now))
    return result

def undo_edit(workspace, *, redo=False, now=None):
    result = copy.deepcopy(dict(workspace))
    source, target = ("redo", "undo") if redo else ("undo", "redo")
    if not result.get(source):
        return result
    previous = result[source].pop()
    current = result["overrides"]
    result[target] = (result[target] + [copy.deepcopy(current)])[-MAX_HISTORY:]
    result["overrides"] = previous
    result["changes"].append({"at": stamp(now), "source": source, "deal_key": "", "field": "transaction", "before": current, "after": previous})
    result["changes"] = result["changes"][-5000:]
    result["saved_at"] = stamp(now)
    return result

def add_item(workspace, *, item_type, text, owner="", due_date="", row=None, notes="", now=None):
    if item_type not in ITEM_TYPES:
        raise ValueError("Unknown item type")
    item_text = globals()["text"](text)
    if not item_text or len(item_text) > 1000:
        raise ValueError("Enter an item of 1–1,000 characters")
    result, row = copy.deepcopy(dict(workspace)), {} if row is None else row
    item = {key: "" for key in ACTION_ITEM_COLUMNS}
    item.update(item_id=uuid.uuid4().hex, created_meeting_date=result["meeting_date"], type=item_type,
        text=item_text, owner=globals()["text"](owner), due_date=date_text(due_date) if item_type == "Action" else "",
        status="Open" if item_type in ("Action", "Parking lot") else "Logged", created_at=stamp(now),
        updated_at=stamp(now), notes=globals()["text"](notes)[:4000])
    for key in ("sheet", "deal_number", "deal_key", "deal_name"):
        item[key] = globals()["text"](row.get(key))
    result["items"].append(item)
    result["saved_at"] = stamp(now)
    return result

def update_item(workspace, item_id, *, status=None, owner=None, due_date=None, item_text=None, notes=None, now=None):
    result = copy.deepcopy(dict(workspace))
    item = next((i for i in result["items"] if i["item_id"] == item_id), None)
    if item is None:
        raise ValueError("Follow-up was not found")
    if status is not None and status not in ("Open", "Done", "Dropped", "Logged"):
        raise ValueError("Unknown follow-up status")
    if item_text is not None:
        if not text(item_text) or len(text(item_text)) > 1000:
            raise ValueError("Enter an item of 1–1,000 characters")
        item["text"] = text(item_text)
    if owner is not None:
        item["owner"] = text(owner)
    if due_date is not None:
        item["due_date"] = date_text(due_date)
    if notes is not None:
        item["notes"] = text(notes)[:4000]
    if status is not None:
        item["status"] = status
        item["closed_at"] = stamp(now) if status in ("Done", "Dropped") else ""
        item["closed_meeting_date"] = result["meeting_date"] if status in ("Done", "Dropped") else ""
    item["updated_at"] = stamp(now)
    result["saved_at"] = stamp(now)
    return result

def merge_items(existing, incoming):
    merged = {str(i["item_id"]): copy.deepcopy(i) for i in existing if i.get("item_id")}
    for item in incoming:
        key = text(item.get("item_id"))
        if key and (key not in merged or text(item.get("updated_at")) > text(merged[key].get("updated_at"))):
            if item.get("type") in ITEM_TYPES and item.get("status") in ("Open", "Done", "Dropped", "Logged"):
                merged[key] = {c: text(item.get(c)) for c in ACTION_ITEM_COLUMNS}
    return list(merged.values())

def followups(workspace, *, deal_key=None, mode="Open"):
    date = workspace["meeting_date"]
    end = (dt.date.fromisoformat(date) + dt.timedelta(days=7)).isoformat()
    found = []
    for item in workspace["items"]:
        if deal_key is not None and item.get("deal_key") != deal_key:
            continue
        opened, due = item.get("status") == "Open", item.get("due_date") or ""
        matches = {"All": True, "Open": opened, "Overdue": opened and bool(due) and due < date,
                   "Due ≤7d": opened and bool(due) and date <= due <= end,
                   "Closed this meeting": item.get("closed_meeting_date") == date}
        if matches.get(mode, True):
            found.append(copy.deepcopy(item))
    return sorted(found, key=lambda i: (i.get("owner", "").casefold(), i.get("due_date") or "9999", i.get("created_at", "")))

def set_outcome(workspace, deal_key, outcome, *, now=None):
    if outcome not in OUTCOMES:
        raise ValueError("Unknown meeting outcome")
    result = copy.deepcopy(dict(workspace))
    result["progress"][deal_key] = {**result["progress"].get(deal_key, {}), "outcome": outcome, "updated_at": stamp(now)}
    result["saved_at"] = stamp(now)
    return result

def clock_seconds(clock, *, now=None):
    if clock.get("started_at") is None:
        return 0.0
    end = clock.get("ended_at")
    if end is None:
        end = clock.get("paused_at")
    if end is None:
        end = time.time() if now is None else now
    return max(0.0, float(end) - float(clock["started_at"]) - float(clock.get("paused_total_s") or 0))

def control_clock(workspace, action, *, now=None):
    result = copy.deepcopy(dict(workspace))
    clock, now = result["clock"], time.time() if now is None else float(now)
    if action == "start" and clock["started_at"] is None:
        clock.update(started_at=now, paused_at=None, paused_total_s=0.0, ended_at=None)
    elif action == "pause" and clock["started_at"] is not None and clock["paused_at"] is None and clock["ended_at"] is None:
        clock["paused_at"] = now
    elif action == "resume" and clock["paused_at"] is not None and clock["ended_at"] is None:
        clock["paused_total_s"] += max(0, now - clock["paused_at"])
        clock["paused_at"] = None
    elif action == "end" and clock["started_at"] is not None and clock["ended_at"] is None:
        clock["ended_at"] = clock["paused_at"] if clock["paused_at"] is not None else now
        clock["paused_at"] = None
    elif action not in ("start", "pause", "resume", "end"):
        raise ValueError("Unknown clock action")
    return result


def validate_workspace(workspace, workbook_sha=None):
    if not isinstance(workspace, dict) or workspace.get("schema") != SCHEMA:
        raise ValueError("This is not a supported meeting recovery file")
    sha = workspace.get("workbook_sha")
    if not isinstance(sha, str) or not re.fullmatch(r"[a-fA-F0-9]{64}", sha):
        raise ValueError("Recovery workbook identity is missing")
    if workbook_sha and sha.lower() != workbook_sha.lower():
        raise ValueError("This recovery belongs to a different workbook. Load that workbook first")
    result = new_workspace(sha, text(workspace.get("workbook_name")), workspace.get("meeting_date"))
    for key in result:
        if key in workspace:
            result[key] = copy.deepcopy(workspace[key])
    for key in ("overrides", "progress", "clock"):
        if not isinstance(result[key], dict):
            raise ValueError("Malformed meeting state")
    for key in ("items", "undo", "redo", "changes", "attendees", "archives"):
        if not isinstance(result[key], list):
            raise ValueError("Malformed meeting state")
    if len(result["items"]) > 10000 or len(result["archives"]) > 104 or len(result["undo"]) > MAX_HISTORY or len(result["redo"]) > MAX_HISTORY:
        raise ValueError("Recovery contains too many records")
    for snapshot in [result["overrides"], *result["undo"], *result["redo"]]:
        if not isinstance(snapshot, dict) or any(not isinstance(v, dict) for v in snapshot.values()):
            raise ValueError("Malformed edit recovery")
        for payload in snapshot.values():
            if payload.get("workbook_sha") not in (None, "", sha):
                raise ValueError("Mismatched edit recovery")
            if any(not isinstance(payload.get(f, ""), str) for f in ("status", "owner", "commentary")):
                raise ValueError("Malformed edit recovery")
    seen = set()
    for item in result["items"]:
        if not isinstance(item, dict) or not isinstance(item.get("item_id"), str) or not item["item_id"] or item["item_id"] in seen:
            raise ValueError("Malformed or duplicate follow-up")
        seen.add(item["item_id"])
        if item.get("type") not in ITEM_TYPES or item.get("status") not in ("Open", "Done", "Dropped", "Logged") or not text(item.get("text")):
            raise ValueError("Malformed follow-up")
        date_text(item.get("due_date"))
    for value in result["progress"].values():
        if not isinstance(value, dict) or value.get("outcome", "") not in ("", *OUTCOMES):
            raise ValueError("Malformed meeting progress")
    for key in ("started_at", "paused_at", "paused_total_s", "ended_at"):
        value = result["clock"].get(key)
        if value is not None and (isinstance(value, bool) or not isinstance(value, (int, float)) or not math.isfinite(value)):
            raise ValueError("Malformed meeting clock")
    for archive in result["archives"]:
        if not isinstance(archive, dict) or archive.get("schema") != SNAPSHOT_SCHEMA or not isinstance(archive.get("deals"), dict):
            raise ValueError("Malformed meeting archive")
    if any(not isinstance(v, str) for v in result["attendees"]):
        raise ValueError("Malformed attendee list")
    date_text(result["next_meeting"])
    return result

def recovery_bytes(workspace):
    result = copy.deepcopy(dict(workspace))
    result["saved_at"] = stamp()
    data = json.dumps(safe(result), ensure_ascii=False, indent=2, allow_nan=False).encode("utf-8")
    if len(data) > MAX_BYTES:
        raise ValueError("Recovery exceeds the 5 MB limit")
    return data

def parse_recovery(data, workbook_sha=None):
    if not isinstance(data, bytes) or len(data) > MAX_BYTES:
        raise ValueError("Choose a recovery JSON smaller than 5 MB")
    def invalid_number(value):
        raise ValueError("Invalid JSON number")
    try:
        return validate_workspace(json.loads(data.decode("utf-8-sig"), parse_constant=invalid_number), workbook_sha)
    except (UnicodeDecodeError, json.JSONDecodeError, RecursionError, TypeError, KeyError) as exc:
        raise ValueError("The recovery JSON is malformed") from exc

def journal_path(folder, workbook_sha):
    if not re.fullmatch(r"[a-fA-F0-9]{64}", str(workbook_sha)):
        raise ValueError("A workbook SHA-256 is required")
    if not text(folder):
        raise ValueError("Choose a local journal folder")
    return Path(folder).expanduser().resolve() / ("wpo_recovery_" + workbook_sha.lower() + ".json")

def write_journal(folder, workspace):
    """Atomic write. UI calls this only after explicit local-disk opt-in."""
    data, path = recovery_bytes(workspace), journal_path(folder, workspace["workbook_sha"])
    path.parent.mkdir(parents=True, exist_ok=True)
    if path.is_symlink():
        raise ValueError("The journal target must be a regular file")
    temp_name = None
    try:
        with tempfile.NamedTemporaryFile(mode="wb", dir=path.parent, prefix=".wpo-", suffix=".tmp", delete=False) as handle:
            temp_name = handle.name
            handle.write(data)
            handle.flush()
            os.fsync(handle.fileno())
        os.replace(temp_name, path)
    finally:
        if temp_name and os.path.exists(temp_name):
            os.unlink(temp_name)
    return path

def read_journal(folder, workbook_sha):
    path = journal_path(folder, workbook_sha)
    if path.is_symlink() or path.stat().st_size > MAX_BYTES:
        raise ValueError("The journal file is not valid")
    return parse_recovery(path.read_bytes(), workbook_sha)

@st.cache_resource(show_spinner=False)
def journal_store():
    return {"lock": threading.RLock(), "entries": {}}

def memory_save(token, payload, *, store=None, now=None):
    store = journal_store() if store is None else store
    now = time.time() if now is None else now
    with store["lock"]:
        entries = store["entries"]
        for key in list(entries):
            if now - entries[key]["touched"] > 12 * 3600:
                del entries[key]
        entries[token] = {"touched": now, "payload": copy.deepcopy(payload)}
        while len(entries) > 8:
            del entries[min(entries, key=lambda key: entries[key]["touched"])]
        # At most 200 MB of workbook bytes process-wide.
        def size():
            return sum(len(e["payload"].get("workbook", {}).get("bytes") or b"") for e in entries.values())
        while len(entries) > 1 and size() > 200 * 1024 * 1024:
            del entries[min(entries, key=lambda key: entries[key]["touched"])]

def memory_load(token, *, store=None, now=None):
    store = journal_store() if store is None else store
    now = time.time() if now is None else now
    with store["lock"]:
        entry = store["entries"].get(token)
        if entry is None:
            return None
        if now - entry["touched"] > 12 * 3600:
            del store["entries"][token]
            return None
        entry["touched"] = now
        return copy.deepcopy(entry["payload"])

def restore_session(app=None):
    """Call after app.init_state(), before any widgets. Restores memory only."""
    if st.session_state.get("_mw_restored_once"):
        return False
    st.session_state["_mw_restored_once"] = True
    token = st.query_params.get("m")
    if not token or not re.fullmatch(r"[A-Za-z0-9_-]{20,80}", str(token)):
        return False
    recovered = memory_load(token)
    if recovered is None:
        st.session_state["_mw_restore_notice"] = "This autosave is no longer available after the app restarted. Resume from a recovery file."
        return False
    workbook = recovered.get("workbook")
    if workbook and workbook.get("bytes") and hashlib.sha256(workbook["bytes"]).hexdigest() != workbook.get("sha256"):
        return False
    for key, value in recovered.get("session", {}).items():
        st.session_state[key] = copy.deepcopy(value)
    if workbook and workbook.get("bytes"):
        # _current_workbook is the app's persistent workbook record.
        st.session_state["_mw_pending_workbook"] = workbook
        st.session_state["wb"] = copy.deepcopy(workbook)
    st.session_state["_mw_token"] = token
    st.session_state["_mw_restore_notice"] = "Restored meeting in progress from memory autosave"
    return True

def restored_workbook():
    """Return uploaded workbook record recovered before widget creation."""
    return copy.deepcopy(st.session_state.get("_mw_pending_workbook"))

def _token():
    token = st.session_state.get("_mw_token") or st.query_params.get("m")
    if not token or not re.fullmatch(r"[A-Za-z0-9_-]{20,80}", str(token)):
        token = secrets.token_urlsafe(24)
    st.session_state["_mw_token"] = token
    if st.query_params.get("m") != token:
        st.query_params["m"] = token
    return token

def build_snapshot(rows, workspace, *, now=None):
    if isinstance(rows, pd.DataFrame):
        rows = rows.to_dict("records")
    deals = {}
    for row in rows:
        key = text(row.get("deal_key"))
        if key:
            deals[key] = {f: safe(row.get(f)) for f in ("deal_name", "sheet", "deal_number", *SNAPSHOT_FIELDS)}
    return {"schema": SNAPSHOT_SCHEMA, "meeting_date": workspace["meeting_date"], "workbook_sha": workspace["workbook_sha"], "saved_at": stamp(now), "deals": deals}

def compare_snapshots(previous, current):
    if any(not isinstance(s, dict) or s.get("schema") != SNAPSHOT_SCHEMA or not isinstance(s.get("deals"), dict) for s in (previous, current)):
        raise ValueError("Choose valid meeting snapshots")
    old, new, changes = previous["deals"], current["deals"], []
    for key in sorted(set(old) | set(new)):
        if key not in old or key not in new:
            row = new.get(key, old.get(key))
            if not isinstance(row, dict):
                raise ValueError("Malformed snapshot deal")
            changes.append({"Deal Key": key, "Deal": row.get("deal_name", ""), "Field": "Deal", "Before": "Present" if key in old else "", "After": "Present" if key in new else ""})
            continue
        if not isinstance(old[key], dict) or not isinstance(new[key], dict):
            raise ValueError("Malformed snapshot deal")
        for field in SNAPSHOT_FIELDS:
            left, right = old[key].get(field), new[key].get(field)
            if left != right:
                changes.append({"Deal Key": key, "Deal": new[key].get("deal_name", ""), "Field": field, "Before": left, "After": right})
    return changes

def csv_bytes(rows, columns=None):
    columns = columns or list(dict.fromkeys(k for row in rows for k in row))
    stream = io.StringIO(newline="")
    writer = csv.DictWriter(stream, fieldnames=columns, extrasaction="ignore")
    writer.writeheader()
    for row in rows:
        result = {}
        for key in columns:
            value = row.get(key, "")
            if isinstance(value, (dict, list)):
                value = json.dumps(safe(value), ensure_ascii=False)
            if isinstance(value, str) and value.lstrip(" \t\r\n").startswith(("=", "+", "-", "@")):
                value = "'" + value
            result[key] = value
        writer.writerow(result)
    return stream.getvalue().encode("utf-8-sig")

def checklist(workspace, deal_keys):
    progress = workspace.get("progress", {})
    return {"pending": [k for k in deal_keys if not progress.get(k, {}).get("outcome")],
            "parked": [k for k in deal_keys if progress.get(k, {}).get("outcome") == "Parked"],
            "incomplete_actions": [i for i in workspace["items"] if i.get("type") == "Action" and i.get("status") == "Open" and (not i.get("owner") or not i.get("due_date"))],
            "open_actions": sum(i.get("type") == "Action" and i.get("status") == "Open" for i in workspace["items"]),
            "overdue": len(followups(workspace, mode="Overdue")),
            "exported": workspace.get("export_fingerprint") == fingerprint(workspace),
            "minutes_sent": bool(workspace.get("minutes_sent"))}

def _md(value):
    return html.escape(text(value)).replace("\n", "  \n")

def build_minutes(workspace, rows, *, data_as_of=""):
    if isinstance(rows, pd.DataFrame):
        rows = rows.to_dict("records")
    lookup = {text(r.get("deal_key")): r for r in rows}
    check = checklist(workspace, list(lookup))
    lines = [f"# Portfolio meeting · {workspace['meeting_date']}", "", "Confidential — internal use only", "",
             "Attendees: " + (_md(", ".join(workspace.get("attendees", []))) or "Not recorded"),
             "Data as of: " + (_md(data_as_of) or "Not recorded"),
             f"Duration: {clock_seconds(workspace['clock']) / 60:.1f} minutes",
             f"Coverage: {len(lookup) - len(check['pending'])} of {len(lookup)} deals have an outcome; {check['open_actions']} open actions.", ""]
    for kind in ("Decision", "Action", "Parking lot", "Note"):
        lines.extend(["## " + kind + ("s" if kind != "Parking lot" else ""), ""])
        items = [i for i in workspace["items"] if i.get("type") == kind and i.get("status") != "Dropped"]
        if not items:
            lines.append("None recorded.")
        for item in sorted(items, key=lambda i: (i.get("owner", "").casefold(), i.get("due_date", ""))):
            details = ""
            if kind == "Action":
                details = " · " + (_md(item.get("owner")) or "Owner unassigned") + " · due " + (_md(item.get("due_date")) or "not set") + " · " + item["status"]
            lines.append("- " + _md(item["text"]) + " · " + (_md(item.get("deal_name")) or "General") + details)
        lines.append("")
    lines.extend(["## Status and owner changes", ""])
    changes = [c for c in workspace["changes"] if c.get("field") in ("status", "owner")]
    if not changes:
        lines.append("None recorded.")
    for change in changes:
        name = lookup.get(change["deal_key"], {}).get("deal_name") or change["deal_key"]
        lines.append("- " + _md(name) + " · " + change["field"] + ": " + (_md(change.get("before")) or "Blank") + " → " + (_md(change.get("after")) or "Blank"))
    lines.extend(["", "## Deferred / not reached", ""])
    omitted = []
    for key, row in lookup.items():
        outcome = workspace["progress"].get(key, {}).get("outcome") or "Not reached"
        if outcome in ("Deferred", "Not reached", "Parked"):
            omitted.append("- " + (_md(row.get("deal_name")) or _md(key)) + " · " + outcome)
    lines.extend(omitted or ["None."])
    lines.extend(["", "Next meeting: " + _md(workspace.get("next_meeting")), ""])
    markdown = "\n".join(lines)
    plain = markdown.replace("## ", "").replace("# ", "")
    document = '<!doctype html><html><head><meta charset="utf-8"><title>Portfolio meeting minutes</title><style>body{font:15px/1.5 system-ui,sans-serif;max-width:900px;margin:2rem auto;padding:1rem;color:#172033}pre{white-space:pre-wrap;font:inherit}</style></head><body><pre>' + html.escape(plain) + '</pre></body></html>'
    return markdown, document, plain

def minutes_email(workspace, rows, *, data_as_of=""):
    _, document, plain = build_minutes(workspace, rows, data_as_of=data_as_of)
    message = EmailMessage()
    message["Subject"] = "Portfolio meeting minutes " + workspace["meeting_date"]
    message["X-Unsent"] = "1"
    message.set_content(plain)
    message.add_alternative(document, subtype="html")
    return message.as_bytes()

def meeting_pack(workspace, rows, *, data_as_of=""):
    snapshot = build_snapshot(rows, workspace)
    md, document, _ = build_minutes(workspace, rows, data_as_of=data_as_of)
    data = io.BytesIO()
    with zipfile.ZipFile(data, "w", zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("meeting_recovery.json", recovery_bytes(workspace))
        archive.writestr("meeting_snapshot.json", json.dumps(snapshot, ensure_ascii=False, indent=2))
        archive.writestr("minutes.md", md)
        archive.writestr("minutes.html", document)
        archive.writestr("outlook_draft.eml", minutes_email(workspace, rows, data_as_of=data_as_of))
        archive.writestr("actions.csv", csv_bytes(workspace["items"], ACTION_ITEM_COLUMNS))
        archive.writestr("changes.csv", csv_bytes(workspace["changes"], ["at", "deal_key", "field", "before", "after", "source"]))
    return data.getvalue()


def _sha(ctx):
    return text(getattr(ctx, "wb_sha", "") or getattr(ctx, "wb", {}).get("sha256"))

def _meeting_date(ctx):
    return st.session_state.get("meeting_date") or getattr(ctx, "as_of", dt.date.today())

def _records(ctx, raw=False):
    frame = getattr(ctx, "raw_deck" if raw else "deck", pd.DataFrame())
    return frame.to_dict("records") if isinstance(frame, pd.DataFrame) else list(frame or [])

def _embedded_rows(ctx, name):
    entry = getattr(ctx, "parsed", {}).get("hidden_sheets", getattr(ctx, "parsed", {}).get("underscore_sheets", {})).get(name, {})
    if isinstance(entry, pd.DataFrame):
        return entry.to_dict("records")
    if isinstance(entry, dict):
        for key in ("table", "dataframe", "data"):
            if isinstance(entry.get(key), pd.DataFrame):
                return entry[key].to_dict("records")
        rows, headers = entry.get("rows", []), entry.get("headers") or entry.get("columns") or []
        if rows and isinstance(rows[0], dict):
            return copy.deepcopy(rows)
        if headers:
            return [dict(zip(headers, row)) for row in rows]
    return []

def ensure_workspace(ctx):
    key = _sha(ctx) + ":" + date_text(_meeting_date(ctx))
    stored = st.session_state.get("_mw_workspaces", {})
    if key not in stored:
        workspace = new_workspace(_sha(ctx), text(getattr(ctx, "workbook_name", "")), _meeting_date(ctx))
        # Carry prior open follow-ups across a new meeting date in the same workbook.
        prior = next((v for k, v in reversed(list(stored.items())) if k.startswith(_sha(ctx) + ":")), {})
        workspace["items"] = merge_items(prior.get("items", []), _embedded_rows(ctx, "_Action Items"))
        workspace["archives"] = copy.deepcopy(prior.get("archives", []))
        stored = {**stored, key: workspace}
        if len(stored) > 8:
            stored = dict(list(stored.items())[-8:])
        st.session_state["_mw_workspaces"] = stored
    st.session_state["_mw_active"] = key
    return copy.deepcopy(stored[key])

def _memory_checkpoint(ctx):
    if st.session_state.get("_mw_memory_suspended"):
        return
    session_keys = ("_mw_workspaces", "_mw_active", "meeting_overrides", "meeting_date", "as_of_date", "data_as_of", "include_hidden", "cursor", "sheet_filter", "search_query", "stale_only", "owner_filter", "agenda_order", "focus_filters", "reminder_rule", "view_mode", "pinned_fields", "box_enabled", "box_root", "box_subfolders", "box_depth", "lg_loan_col", "reminder_confirmed", "state_version", "review_editor_version")
    session = {key: copy.deepcopy(st.session_state[key]) for key in session_keys if key in st.session_state}
    wb = copy.deepcopy(getattr(ctx, "wb", {}))
    if wb:
        wb["source"] = "journal"
        wb.pop("path", None)
    memory_save(_token(), {"workbook": wb, "session": session})

def _put(ctx, workspace):
    key = _sha(ctx) + ":" + date_text(_meeting_date(ctx))
    stored = dict(st.session_state.get("_mw_workspaces", {}))
    stored[key] = copy.deepcopy(workspace)
    st.session_state["_mw_workspaces"] = stored
    st.session_state["_mw_model"] = workspace  # Export fingerprint bridge for the app.
    _memory_checkpoint(ctx)
    if (st.session_state.get("_mw_local_enabled") and text(st.session_state.get("_mw_local_folder"))
            and st.session_state.get("_mw_local_authorized") == (str(Path(st.session_state["_mw_local_folder"]).expanduser().resolve()), _sha(ctx))):
        signature = (key, fingerprint(workspace))
        if st.session_state.get("_mw_journal_stamp") != signature:
            try:
                write_journal(st.session_state["_mw_local_folder"], workspace)
                st.session_state["_mw_journal_stamp"] = signature
                st.session_state["_mw_journal_error"] = ""
            except (OSError, ValueError, TypeError) as exc:
                st.session_state["_mw_journal_error"] = type(exc).__name__

def _native_overrides(ctx):
    keys = {text(r.get("deal_key")) for r in _records(ctx)}
    return {k: copy.deepcopy(v) for k, v in st.session_state.get("meeting_overrides", {}).items()
            if k in keys and v.get("workbook_sha", _sha(ctx)) in ("", _sha(ctx))}

def checkpoint_overrides(ctx, app=None):
    workspace = ensure_workspace(ctx)
    base = {text(r.get("deal_key")): r for r in _records(ctx, raw=True)}
    workspace = checkpoint(workspace, _native_overrides(ctx), base_rows=base)
    _put(ctx, workspace)
    return workspace

def _notify(message, app=None):
    if app is not None and callable(getattr(app, "notify", None)):
        app.notify(message)
    else:
        st.toast(message)

def _apply_overrides(ctx, workspace):
    keys = {text(r.get("deal_key")) for r in _records(ctx)}
    preserved = {k: v for k, v in st.session_state.get("meeting_overrides", {}).items() if k not in keys}
    preserved.update({k: copy.deepcopy(v) for k, v in workspace["overrides"].items() if k in keys})
    st.session_state["meeting_overrides"] = preserved
    for counter in ("state_version", "review_editor_version"):
        st.session_state[counter] = int(st.session_state.get(counter) or 0) + 1
    _put(ctx, workspace)

def _undo_cb(ctx, app=None, redo=False):
    _apply_overrides(ctx, undo_edit(ensure_workspace(ctx), redo=redo))
    _notify("Redid last edit" if redo else "Undid last edit", app)

def stamp_export(ctx):
    workspace = ensure_workspace(ctx)
    workspace["export_fingerprint"] = fingerprint(workspace)
    workspace["exported_at"] = stamp()
    _put(ctx, workspace)

def _restore_cb(ctx, recovered, app=None):
    current = ensure_workspace(ctx)
    recovered = validate_workspace(recovered, _sha(ctx))
    if recovered["meeting_date"] != date_text(_meeting_date(ctx)):
        st.session_state["meeting_date"] = dt.date.fromisoformat(recovered["meeting_date"])
    recovered["undo"] = (current["undo"] + [current["overrides"]])[-MAX_HISTORY:]
    recovered["redo"], recovered["baseline_ready"] = [], True
    _apply_overrides(ctx, recovered)
    st.session_state.pop("_mw_restore_preview", None)
    for key in list(st.session_state):
        if str(key).startswith(("mw_attendees_", "mw_next_", "mw_sent_", "mw_done_")):
            del st.session_state[key]
    _notify("Restored meeting recovery", app)

def render_settings(ctx, app=None):
    workspace = checkpoint_overrides(ctx, app)
    tag = workspace["workbook_sha"][:10]
    with st.expander("Meeting backup & undo", expanded=False):
        if st.session_state.get("_mw_memory_suspended"):
            st.info("Meeting memory autosave is cleared. Start a new meeting when ready.")
            st.button("Start new meeting", key="mw_new_meeting", on_click=_start_new_cb, args=(ctx,))
        notice = st.session_state.get("_mw_restore_notice")
        if notice:
            st.info(notice)
        left, right = st.columns(2)
        left.button("Undo edit", key="mw_undo_" + tag, disabled=not workspace["undo"], on_click=_undo_cb, args=(ctx, app, False), width="stretch")
        right.button("Redo edit", key="mw_redo_" + tag, disabled=not workspace["redo"], on_click=_undo_cb, args=(ctx, app, True), width="stretch")
        dirty = bool(workspace["overrides"] or workspace["items"] or workspace["progress"]) and workspace["export_fingerprint"] != fingerprint(workspace)
        st.caption("Meeting changes have not been downloaded" if dirty else "No pending meeting backup")
        st.download_button("Save recovery file", recovery_bytes(workspace), file_name="wpo_recovery_" + workspace["meeting_date"] + ".json", mime="application/json", key="mw_backup_" + tag, on_click=stamp_export, args=(ctx,), width="stretch")
        upload = st.file_uploader("Resume from recovery JSON", type=["json"], key="mw_restore_" + tag)
        if upload is not None:
            signature = (tag, hashlib.sha256(upload.getvalue()).hexdigest())
            if st.session_state.get("_mw_preview_signature") != signature:
                try:
                    st.session_state["_mw_restore_preview"] = parse_recovery(upload.getvalue(), workspace["workbook_sha"])
                    st.session_state["_mw_preview_signature"] = signature
                    st.session_state["_mw_restore_error"] = ""
                except ValueError as exc:
                    st.session_state["_mw_restore_preview"] = None
                    st.session_state["_mw_restore_error"] = str(exc)
            if st.session_state.get("_mw_restore_error"):
                st.error(st.session_state["_mw_restore_error"])
        recovered = st.session_state.get("_mw_restore_preview")
        if recovered and recovered.get("workbook_sha") == workspace["workbook_sha"]:
            st.info(f"Recovery preview: {len(recovered['overrides'])} edited deals, {len(recovered['items'])} captured items. Applying replaces this workbook's meeting state.")
            st.button("Apply recovery", key="mw_apply_restore_" + tag, on_click=_restore_cb, args=(ctx, recovered, app), width="stretch")
        enabled = st.toggle("Enable local disk autosave", value=False, key="_mw_local_enabled")
        if enabled:
            st.caption("Choose a private local folder. Recovery contains confidential meeting notes.")
            folder = st.text_input("Local journal folder", key="_mw_local_folder", placeholder="Choose a private folder")
            if folder:
                path = journal_path(folder, workspace["workbook_sha"])
                authorized = st.session_state.get("_mw_local_authorized") == (str(path.parent), _sha(ctx))
                if not authorized:
                    st.caption("Existing local recovery is kept until you explicitly start autosave.")
                    if path.exists() and st.button("Preview local journal", key="mw_local_preview_" + tag):
                        try:
                            st.session_state["_mw_restore_preview"] = read_journal(folder, workspace["workbook_sha"])
                            st.rerun()
                        except (OSError, ValueError) as exc:
                            st.warning("Could not read local journal (" + type(exc).__name__ + ")")
                    label = "Start local autosave (replace existing)" if path.exists() else "Start local autosave"
                    st.button(label, key="mw_local_start_" + tag, on_click=_enable_local_cb, args=(ctx, folder))
                else:
                    _put(ctx, workspace)
                    if st.session_state.get("_mw_journal_error"):
                        st.warning("Autosave failed (" + st.session_state["_mw_journal_error"] + "). Save a recovery file.")
                    else:
                        st.caption("Local journal saved atomically")

def _start_new_cb(ctx):
    st.session_state.pop("_mw_memory_suspended", None)
    st.session_state["_mw_restored_once"] = True
    _put(ctx, ensure_workspace(ctx))

def _enable_local_cb(ctx, folder):
    path = journal_path(folder, _sha(ctx))
    st.session_state["_mw_local_authorized"] = (str(path.parent), _sha(ctx))
    st.session_state.pop("_mw_journal_stamp", None)
    _put(ctx, ensure_workspace(ctx))

def _outcome_cb(ctx, key, outcome, app=None):
    _put(ctx, set_outcome(ensure_workspace(ctx), key, outcome))
    _notify("Marked " + outcome.lower(), app)

def _item_status_cb(ctx, item_id, widget_key, app=None):
    _put(ctx, update_item(ensure_workspace(ctx), item_id, status="Done" if st.session_state[widget_key] else "Open"))
    _notify("Follow-up updated", app)

def render_deal(ctx, row, app=None):
    workspace = ensure_workspace(ctx)
    key = text(row.get("deal_key"))
    if key not in workspace["progress"]:
        workspace["progress"][key] = {"visited_at": stamp(), "outcome": ""}
        _put(ctx, workspace)
    tag = hashlib.sha256((workspace["workbook_sha"] + key).encode()).hexdigest()[:12]
    with st.expander("Meeting outcome & capture", expanded=False):
        st.caption("Outcome: " + (workspace["progress"].get(key, {}).get("outcome") or "Not recorded"))
        for col, option in zip(st.columns(4), ("Discussed", "No change", "Deferred", "Parked")):
            col.button(option, key="mw_outcome_" + tag + option, on_click=_outcome_cb, args=(ctx, key, option, app), width="stretch")
        kind = st.selectbox("Capture type", ITEM_TYPES, key="mw_type_" + tag)
        with st.form("mw_capture_" + tag, clear_on_submit=True):
            item_text = st.text_input("What", max_chars=1000, key="mw_text_" + tag)
            owner, due = "", ""
            if kind == "Action":
                owner = st.text_input("Action owner", value=text(row.get("owner")), key="mw_owner_" + tag)
                due = st.date_input("Action due", value=dt.date.fromisoformat(workspace["meeting_date"]) + dt.timedelta(days=7), key="mw_due_" + tag)
            general = st.checkbox("General meeting item", key="mw_general_" + tag)
            if st.form_submit_button("Add item", type="primary"):
                try:
                    workspace = add_item(ensure_workspace(ctx), item_type=kind, text=item_text, owner=owner, due_date=due, row=None if general else row)
                    _put(ctx, workspace)
                    _notify("Captured " + kind.lower(), app)
                except ValueError as exc:
                    st.warning(str(exc))
        for item in [i for i in workspace["items"] if i.get("deal_key") == key and i.get("status") != "Dropped"]:
            if item["type"] in ("Action", "Parking lot"):
                widget_key = "mw_done_" + tag + item["item_id"]
                label = item["text"] + " · " + (item.get("owner") or "Unassigned")
                if item.get("due_date"):
                    label += " · due " + item["due_date"]
                st.checkbox(label, value=item["status"] == "Done", key=widget_key, on_change=_item_status_cb, args=(ctx, item["item_id"], widget_key, app))
                if item["status"] == "Open" and item.get("due_date") and item["due_date"] < workspace["meeting_date"]:
                    st.caption("Overdue follow-up")
            else:
                st.write(item["type"] + ": " + item["text"])

def _clock_cb(ctx, action):
    _put(ctx, control_clock(ensure_workspace(ctx), action))

@st.fragment(run_every="5s")
def _clock_face(clock):
    seconds = int(clock_seconds(clock))
    st.caption(f"Meeting elapsed {seconds // 60}:{seconds % 60:02d}" + (" · paused" if clock.get("paused_at") is not None else ""))

def render_clock(ctx):
    workspace = ensure_workspace(ctx)
    clock = workspace["clock"]
    _clock_face(clock)
    cols = st.columns(3)
    cols[0].button("Start meeting", key="mw_clock_start", disabled=clock["started_at"] is not None, on_click=_clock_cb, args=(ctx, "start"), width="stretch")
    action = "resume" if clock["paused_at"] is not None else "pause"
    cols[1].button("Resume" if action == "resume" else "Pause", key="mw_clock_pause", disabled=clock["started_at"] is None or clock["ended_at"] is not None, on_click=_clock_cb, args=(ctx, action), width="stretch")
    cols[2].button("End meeting", key="mw_clock_end", disabled=clock["started_at"] is None or clock["ended_at"] is not None, on_click=_clock_cb, args=(ctx, "end"), width="stretch")

def render_followups(ctx, app=None):
    workspace = checkpoint_overrides(ctx, app)
    st.subheader("Follow-ups")
    mode = st.selectbox("Show follow-ups", ["Open", "Overdue", "Due ≤7d", "Closed this meeting", "All"], key="mw_followup_mode")
    items = followups(workspace, mode=mode)
    owners = sorted({text(i.get("owner")) for i in items if text(i.get("owner"))})
    selected_owners = st.multiselect("Follow-up owners", owners, key="mw_followup_owners")
    if selected_owners:
        items = [i for i in items if i.get("owner") in selected_owners]
    if not items:
        st.info("No follow-ups match this view")
        return
    st.dataframe(pd.DataFrame(items)[["deal_name", "type", "text", "owner", "due_date", "status"]], hide_index=True, width="stretch")
    labels = {i["item_id"]: (i.get("deal_name") or "General") + " · " + i["text"] for i in items}
    selected = st.selectbox("Update follow-up", list(labels), format_func=labels.get, key="mw_followup_selected")
    item = next(i for i in items if i["item_id"] == selected)
    with st.form("mw_followup_edit_" + selected):
        statuses = ["Open", "Done", "Dropped", "Logged"]
        status = st.selectbox("Item status", statuses, index=statuses.index(item["status"]))
        owner = st.text_input("Item owner", value=item.get("owner", ""))
        due = st.date_input("Item due", value=dt.date.fromisoformat(item["due_date"]) if item.get("due_date") else None)
        item_text = st.text_area("Item text", value=item["text"], max_chars=1000)
        notes = st.text_area("Item notes", value=item.get("notes", ""), max_chars=4000)
        if st.form_submit_button("Save follow-up", type="primary"):
            try:
                _put(ctx, update_item(ensure_workspace(ctx), selected, status=status, owner=owner, due_date=due or "", item_text=item_text, notes=notes))
                _notify("Follow-up saved", app)
                st.rerun()
            except ValueError as exc:
                st.warning(str(exc))

def _remaining_cb(ctx, keys, outcome):
    workspace = ensure_workspace(ctx)
    for key in keys:
        workspace = set_outcome(workspace, key, outcome)
    _put(ctx, workspace)

def _archive_cb(ctx):
    workspace = ensure_workspace(ctx)
    snapshot = build_snapshot(_records(ctx), workspace)
    workspace["archives"] = ([s for s in workspace["archives"] if s.get("meeting_date") != workspace["meeting_date"]] + [snapshot])[-104:]
    _put(ctx, workspace)

def _metadata_cb(ctx, key, target):
    workspace = ensure_workspace(ctx)
    workspace[target] = safe(st.session_state[key])
    _put(ctx, workspace)

def render_wrapup(ctx, app=None):
    workspace = checkpoint_overrides(ctx, app)
    rows = _records(ctx)
    keys = [text(r.get("deal_key")) for r in rows]
    names = {text(r.get("deal_key")): text(r.get("deal_name")) for r in rows}
    check = checklist(workspace, keys)
    st.subheader("Meeting wrap-up")
    render_clock(ctx)
    st.write(f"Coverage: {len(keys) - len(check['pending'])} of {len(keys)} deals have an outcome")
    if check["pending"]:
        st.warning(f"{len(check['pending'])} deals still need an outcome")
        st.dataframe(pd.DataFrame({"Deal": [names.get(k, k) for k in check["pending"]]}), hide_index=True, width="stretch")
        left, right = st.columns(2)
        left.button("Mark remaining Not reached", key="mw_remaining_not_reached", on_click=_remaining_cb, args=(ctx, check["pending"], "Not reached"), width="stretch")
        right.button("Defer remaining", key="mw_remaining_defer", on_click=_remaining_cb, args=(ctx, check["pending"], "Deferred"), width="stretch")
    if check["parked"]:
        st.warning(f"{len(check['parked'])} parked deals need a final decision")
    if check["incomplete_actions"]:
        st.warning(f"{len(check['incomplete_actions'])} open actions need an owner or due date. Complete them in Follow-ups.")
    st.caption(f"{check['open_actions']} open actions · {check['overdue']} overdue follow-ups")
    tag = workspace["workbook_sha"][:10]
    attendee_key, next_key, sent_key = "mw_attendees_" + tag, "mw_next_" + tag, "mw_sent_" + tag
    for key, value in ((attendee_key, workspace["attendees"]), (next_key, dt.date.fromisoformat(workspace["next_meeting"])), (sent_key, bool(workspace["minutes_sent"]))):
        if key not in st.session_state:
            st.session_state[key] = value
    owners = sorted({text(r.get("owner")) for r in rows if text(r.get("owner"))} | set(workspace["attendees"]))
    st.multiselect("Attendees", owners, key=attendee_key, accept_new_options=True, on_change=_metadata_cb, args=(ctx, attendee_key, "attendees"))
    st.date_input("Next meeting", key=next_key, on_change=_metadata_cb, args=(ctx, next_key, "next_meeting"))
    st.checkbox("Minutes sent", key=sent_key, on_change=_metadata_cb, args=(ctx, sent_key, "minutes_sent"))
    workspace = ensure_workspace(ctx)
    md, document, _ = build_minutes(workspace, rows, data_as_of=getattr(ctx, "as_of", ""))
    with st.expander("Minutes preview", expanded=False):
        st.code(md, language="markdown", wrap_lines=True)
    snapshot, data_rows, data_as_of = copy.deepcopy(workspace), copy.deepcopy(rows), getattr(ctx, "as_of", "")
    for col, label, data, ext, mime in zip(st.columns(3), ["Minutes (.md)", "Minutes (.html)", "Outlook draft (.eml)"],
        [md.encode(), document.encode(), lambda: minutes_email(snapshot, data_rows, data_as_of=data_as_of)],
        ["md", "html", "eml"], ["text/markdown", "text/html", "message/rfc822"]):
        col.download_button(label, data, file_name="minutes_" + workspace["meeting_date"] + "." + ext, mime=mime, on_click="ignore", width="stretch")
    cols = st.columns(3)
    cols[0].download_button("Actions (.csv)", csv_bytes(workspace["items"], ACTION_ITEM_COLUMNS), file_name="actions_" + workspace["meeting_date"] + ".csv", mime="text/csv", on_click="ignore", width="stretch")
    cols[1].download_button("Meeting pack (.zip)", lambda: meeting_pack(snapshot, data_rows, data_as_of=data_as_of), file_name="meeting_pack_" + workspace["meeting_date"] + ".zip", mime="application/zip", on_click=stamp_export, args=(ctx,), width="stretch")
    cols[2].download_button("Recovery JSON", recovery_bytes(workspace), file_name="wpo_recovery_" + workspace["meeting_date"] + ".json", mime="application/json", on_click=stamp_export, args=(ctx,), width="stretch")
    with st.expander("Meeting snapshots", expanded=False):
        st.button("Archive current meeting snapshot", key="mw_archive", on_click=_archive_cb, args=(ctx,))
        archives = workspace["archives"]
        upload = st.file_uploader("Compare a prior meeting snapshot", type=["json"], key="mw_snapshot_upload")
        try:
            if upload is not None:
                if len(upload.getvalue()) > MAX_BYTES:
                    raise ValueError("Snapshot exceeds 5 MB")
                prior = json.loads(upload.getvalue().decode("utf-8-sig"))
                st.dataframe(pd.DataFrame(compare_snapshots(prior, build_snapshot(rows, workspace))), hide_index=True, width="stretch")
            elif archives:
                selected = st.selectbox("Archived meeting", list(range(len(archives))), format_func=lambda i: archives[i]["meeting_date"], key="mw_archive_choice")
                st.dataframe(pd.DataFrame(compare_snapshots(archives[selected], build_snapshot(rows, workspace))), hide_index=True, width="stretch")
        except (ValueError, TypeError, KeyError, UnicodeDecodeError, RecursionError):
            st.warning("Choose a valid meeting_snapshot.json from a meeting pack")


    with st.expander("End meeting & clear session data", expanded=False):
        st.caption("Download your workbook, meeting pack or recovery file first. This clears session edits and memory autosave. Local backup files remain in your chosen folder.")
        confirmed = st.checkbox("I have saved the meeting records and want to clear this session", key="mw_confirm_clear")
        st.button("Clear meeting data", key="mw_clear", disabled=not confirmed, on_click=clear_meeting, args=(ctx, app))

def _replace_meeting_rows(prior, incoming, date):
    records = [r for r in prior if text(r.get("meeting_date"))[:10] != date] + incoming
    dates = sorted({text(r.get("meeting_date"))[:10] for r in records})[-104:]
    return [r for r in records if text(r.get("meeting_date"))[:10] in dates]

def workbook_tables(ctx, workspace=None):
    workspace = workspace if workspace is not None else ensure_workspace(ctx)
    rows = _records(ctx)
    before = {text(r.get("deal_key")): r for r in _records(ctx, raw=True)}
    at, log = stamp(), []
    for index, row in enumerate(rows):
        key = text(row.get("deal_key"))
        old, progress = before.get(key, row), workspace["progress"].get(key, {})
        log.append({"meeting_date": workspace["meeting_date"], "sheet": row.get("sheet"), "deal_number": row.get("deal_number"),
            "deal_key": key, "agenda_position": index + 1, "agenda_order": st.session_state.get("agenda_order", "Workbook"),
            "outcome": progress.get("outcome") or "Not reached", "seconds_on_deal": progress.get("seconds", 0),
            "status_before": old.get("status"), "status_after": row.get("status"), "owner_before": old.get("owner"),
            "owner_after": row.get("owner"), "commentary_changed": text(old.get("commentary")) != text(row.get("commentary")),
            "park_reason": progress.get("reason", ""), "items_created": sum(i.get("deal_key") == key and i.get("created_meeting_date") == workspace["meeting_date"] for i in workspace["items"]), "saved_at": at})
    check = checklist(workspace, [text(r.get("deal_key")) for r in rows])
    outcomes = [r["outcome"] for r in log]
    def amount(row, field):
        try:
            value = float(row.get(field) or 0)
            return value if math.isfinite(value) else 0.0
        except (TypeError, ValueError):
            return 0.0
    date = dt.date.fromisoformat(workspace["meeting_date"])
    def maturity_days(row):
        try:
            return (dt.date.fromisoformat(text(row.get("maturity_date"))[:10]) - date).days
        except ValueError:
            return None
    maturity = [maturity_days(row) for row in rows]
    summary = {key: 0 for key in MEETING_SUMMARY_COLUMNS}
    summary.update(meeting_date=workspace["meeting_date"], workbook_name=workspace["workbook_name"],
        workbook_sha256_prefix=workspace["workbook_sha"][:12], agenda_order=st.session_state.get("agenda_order", "Workbook"),
        started_at=stamp(workspace["clock"]["started_at"]) if workspace["clock"]["started_at"] is not None else "",
        ended_at=stamp(workspace["clock"]["ended_at"]) if workspace["clock"]["ended_at"] is not None else "",
        duration_min=round(clock_seconds(workspace["clock"]) / 60, 2), attendees=", ".join(workspace["attendees"]),
        deals_total=len(rows), deals_discussed=sum(o in ("Discussed", "No change", "Updated") for o in outcomes),
        no_change=outcomes.count("No change"), updated=outcomes.count("Updated"), parked=outcomes.count("Parked"),
        deferred=outcomes.count("Deferred"), not_reached=outcomes.count("Not reached"), edits=len(workspace["overrides"]),
        actions_created=sum(i["type"] == "Action" and i.get("created_meeting_date") == workspace["meeting_date"] for i in workspace["items"]),
        actions_closed=sum(i["type"] == "Action" and i.get("closed_meeting_date") == workspace["meeting_date"] for i in workspace["items"]),
        open_actions=check["open_actions"], overdue_actions=check["overdue"], total_upb=sum(amount(r, "upb") for r in rows),
        past_maturity_count=sum(d is not None and d < 0 for d in maturity),
        maturing_30d_count=sum(d is not None and 0 <= d <= 30 for d in maturity),
        dpd30_upb=sum(amount(r, "upb") for r in rows if amount(r, "days_past_due") >= 30),
        reminders=sum(text(r.get("status_needs_update_prompt")).lower() in ("true", "1") and text(r.get("reminder_addressed")).lower() not in ("true", "1") for r in rows),
        app_version=getattr(ctx, "app_version", "2.0.0"), saved_at=at)
    actions = merge_items(_embedded_rows(ctx, "_Action Items"), workspace["items"])
    log = _replace_meeting_rows(_embedded_rows(ctx, "_Meeting Log"), log, workspace["meeting_date"])
    summaries = _replace_meeting_rows(_embedded_rows(ctx, "_Meeting Summary"), [summary], workspace["meeting_date"])
    return {"_Action Items": pd.DataFrame(actions, columns=ACTION_ITEM_COLUMNS),
            "_Meeting Log": pd.DataFrame(log, columns=MEETING_LOG_COLUMNS),
            "_Meeting Summary": pd.DataFrame(summaries, columns=MEETING_SUMMARY_COLUMNS)}

def enrich_workbook(ctx, file_bytes):
    from workbook_io import replace_hidden_sheets
    return replace_hidden_sheets(file_bytes, workbook_tables(ctx))

def clear_meeting(ctx, app=None):
    """Explicit action; remove this meeting's memory journal and session data."""
    token = st.session_state.get("_mw_token")
    store = journal_store()
    with store["lock"]:
        if token:
            store["entries"].pop(token, None)
    for key in list(st.session_state):
        if str(key).startswith(("_mw_", "mw_")):
            del st.session_state[key]
    st.session_state["_mw_memory_suspended"] = True
    st.session_state["_mw_restored_once"] = True
    st.session_state["meeting_overrides"] = {}
    st.session_state["wb"] = None
    for key in ("uploaded_workbook", "uploaded_status_history", "status_history", "prepared_export",
                "_app_context", "_data_context", "_base_rows", "_base_rows_sha", "_wb_names",
                "_wb_sha", "_parsed_sha", "cursor"):
        st.session_state.pop(key, None)
    if "m" in st.query_params:
        del st.query_params["m"]
    if app is not None and hasattr(getattr(app, "parse_workbook", None), "clear"):
        app.parse_workbook.clear()
        for name in ("derive_deck", "compute_status_streaks"):
            cached = getattr(app, name, None)
            if hasattr(cached, "clear"):
                cached.clear()
        changes = getattr(getattr(app, "data_views", None), "compute_changes", None)
        if hasattr(changes, "clear"):
            changes.clear()
    _notify("Meeting data cleared from this app", app)
