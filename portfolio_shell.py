"""Native page navigation and meeting date controls shared by every view."""
from __future__ import annotations

import datetime as dt
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError

import streamlit as st


VIEWS = {"Overview": "overview", "Presentation": "present", "Review": "review",
         "Follow-ups": "follow-ups", "Wrap-up": "wrap-up", "Exports": "exports"}
ICONS = ("dashboard", "present_to_all", "edit_note", "task_alt", "summarize", "download")


def _page_frame():
    # The app builds its shared context once, then dispatches the selected view.
    pass


def request_view():
    st.session_state["_requested_view"] = st.session_state.get("view_mode", "Presentation")
    st.session_state.active_dialog = None


def navigation() -> str:
    pages = {label: st.Page(_page_frame, title=label, url_path=path, icon=f":material/{icon}:", default=label == "Presentation")
             for (label, path), icon in zip(VIEWS.items(), ICONS)}
    selected = st.navigation(list(pages.values()), position="top")
    desired = st.session_state.pop("_requested_view", None)
    if desired in pages and desired != selected.title:
        st.switch_page(pages[desired], query_params=dict(st.query_params))
    selected.run()
    # This runs before the sidebar radio is instantiated.
    st.session_state.view_mode = selected.title
    return selected.title


def viewer_now() -> dt.datetime:
    try:
        zone = st.context.timezone or "America/New_York"
        return dt.datetime.now(ZoneInfo(zone))
    except (ZoneInfoNotFoundError, AttributeError, TypeError):
        return dt.datetime.now(dt.timezone.utc)


def viewer_today() -> dt.date:
    return viewer_now().date()


def render_meeting_date(as_of: dt.date):
    if "meeting_date" not in st.session_state:
        st.session_state.meeting_date = viewer_today()
    with st.sidebar:
        st.date_input("Meeting date", key="meeting_date", format="MM/DD/YYYY", persist_state="session")
        st.caption(f"Meeting: {st.session_state.meeting_date:%m/%d/%Y} · Workbook data: {as_of:%m/%d/%Y}")
