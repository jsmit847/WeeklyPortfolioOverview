"""Optional browser smoke check; requires playwright and an installed Edge browser.

Run: python tests/browser_smoke.py
Uses a synthetic workbook, a temporary directory and a headless local server.
"""
from __future__ import annotations

import os
import re
from pathlib import Path
import socket
import subprocess
import sys
import tempfile
import time
import urllib.request

from playwright.sync_api import sync_playwright
from sample_workbook import build_sample_workbook_bytes


REPO = Path(__file__).resolve().parents[1]


def main():
    with tempfile.TemporaryDirectory(prefix="wpo-browser-", ignore_cleanup_errors=True) as folder:
        work = Path(folder)
        workbook = work / "portfolio_overview.xlsx"
        workbook.write_bytes(build_sample_workbook_bytes())
        with socket.socket() as sock:
            sock.bind(("127.0.0.1", 0))
            port = sock.getsockname()[1]
        env = {**os.environ, "WPO_DEFAULT_WORKBOOK": str(workbook), "PYTHONIOENCODING": "utf-8", "PYTHONDONTWRITEBYTECODE": "1"}
        with (work / "server.log").open("w", encoding="utf-8") as log:
            server = subprocess.Popen([sys.executable, "-B", "-m", "streamlit", "run", str(REPO / "portfolio_meeting_app.py"),
                                       "--server.headless=true", f"--server.port={port}", "--server.address=127.0.0.1"], cwd=REPO, env=env, stdout=log, stderr=log)
            try:
                url = f"http://127.0.0.1:{port}"
                for _ in range(60):
                    try:
                        urllib.request.urlopen(url + "/_stcore/health", timeout=1).close()
                        break
                    except OSError:
                        if server.poll() is not None:
                            raise RuntimeError("The smoke server exited before startup")
                        time.sleep(0.25)
                with sync_playwright() as p:
                    browser = p.chromium.launch(channel="msedge", headless=True)
                    page = browser.new_page(viewport={"width": 1600, "height": 1400})
                    page.goto(url, wait_until="networkidle")

                    def idle():
                        page.wait_for_function("document.querySelector('[data-testid=stApp]')?.getAttribute('data-test-script-state') === 'notRunning'", timeout=120000)
                        assert page.locator('[data-testid="stException"]').count() == 0, page.locator('[data-testid="stException"]').inner_text()

                    idle()
                    assert page.get_by_text("As it stands today", exact=False).count()
                    page.get_by_role("button", name=re.compile(r"^Edit status")).click()
                    page.get_by_role("dialog").wait_for()
                    page.keyboard.press("Escape")
                    page.get_by_role("dialog").wait_for(state="hidden")
                    idle()
                    page.get_by_role("button", name=re.compile(r"^Next(?:\s|$)")).click()
                    idle()
                    assert page.get_by_role("dialog").count() == 0
                    page.get_by_role("link", name="Overview", exact=True).click()
                    idle()
                    assert page.get_by_text("Portfolio intelligence", exact=True).count()
                    page.get_by_role("link", name="Follow-ups", exact=True).click()
                    idle()
                    page.get_by_role("link", name="Wrap-up", exact=True).click()
                    idle()
                    page.get_by_role("link", name="Exports", exact=True).click()
                    idle()
                    page.get_by_role("button", name="Prepare and verify export", exact=True).click()
                    idle()
                    try:
                        page.get_by_text("Workbook checked successfully", exact=False).wait_for(timeout=30000)
                    except Exception:
                        print("Export alerts:", page.locator('[data-testid="stAlert"]').all_text_contents())
                        page.screenshot(path=str(Path(tempfile.gettempdir()) / "wpo-smoke-failure.png"), full_page=True)
                        raise
                    screenshot = Path(tempfile.gettempdir()) / "wpo-export-smoke.png"
                    page.screenshot(path=str(screenshot), full_page=True)
                    browser.close()
                    print(f"Browser smoke passed; screenshot: {screenshot}")
            finally:
                if os.name == "nt":
                    subprocess.run(["taskkill", "/PID", str(server.pid), "/T", "/F"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, check=False)
                else:
                    server.terminate()
                try:
                    server.wait(timeout=10)
                except subprocess.TimeoutExpired:
                    server.kill()


if __name__ == "__main__":
    main()
