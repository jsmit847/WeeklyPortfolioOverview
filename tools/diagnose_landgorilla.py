"""Read one LG loan with the app's client; print statuses and field presence only.

Run from the repository: python tools/diagnose_landgorilla.py --loan 52406
Uses local Streamlit secrets when present, otherwise prompts without echoing secrets.
Never reads credentials from the saved notebooks, or writes anything to Land Gorilla.
"""
from __future__ import annotations

import argparse
from getpass import getpass
import json
import logging
from pathlib import Path
import re
import sys
import tomllib
from urllib.parse import urlsplit

import requests

REPO = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(REPO))
logging.getLogger("streamlit").setLevel(logging.ERROR)
import landgorilla_live as lg


def report(result, calls):
    overview = result.get("overview") or {}
    fields = ("project_pct", "duration_pct", "balance", "to_finish", "due_date", "last_draw", "last_approved_draw")
    return {"requests": calls, "state": result["state"], "error": result.get("error", ""),
            "partial_errors": result.get("partial_errors", []),
            "fields_present": {field: overview.get(field) is not None for field in fields}}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--loan", required=True, help="One workbook loan number, or LG file number")
    parser.add_argument("--secrets", type=Path, default=REPO / ".streamlit" / "secrets.toml")
    args = parser.parse_args()
    config = {}
    if args.secrets.is_file():
        with args.secrets.open("rb") as handle:
            config = tomllib.load(handle).get("landgorilla", {})
    user = config.get("user") or getpass("Land Gorilla user (hidden): ")
    password = config.get("password") or getpass("Land Gorilla password: ")
    settings = lg.lg_settings({**config, "user": user, "password": password})
    if settings is None:
        parser.error("Land Gorilla user and password are required.")
    number = lg.lg_file_number(args.loan, settings.file_prefix)
    if not number:
        parser.error("Use a numeric loan number or a prefixed LG file number.")
    calls = []

    class DiagnosticSession(requests.Session):
        def get(self, url, **kwargs):
            # Omit query strings, ids, credentials, response bodies and header values.
            path = re.sub(r"/\d+(?=/|$)", "/<id>", urlsplit(url).path)
            try:
                response = super().get(url, **kwargs)
            except requests.RequestException as exc:
                calls.append({"endpoint": path, "error_type": type(exc).__name__})
                raise
            calls.append({"endpoint": path, "status": response.status_code})
            return response

    result = lg.lg_fetch_loans(settings, [number], force=True, session_factory=DiagnosticSession)[number]
    print(json.dumps(report(result, calls), indent=2))


if __name__ == "__main__":
    main()
