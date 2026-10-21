from streamlit.testing.v1 import AppTest


def shell_harness():
    import streamlit as st
    import portfolio_shell as shell
    view = shell.navigation()
    st.sidebar.radio("View", list(shell.VIEWS), key="view_mode", on_change=shell.request_view)
    st.write("Selected: " + view)


def test_sidebar_navigation_and_native_pages_stay_in_sync():
    at = AppTest.from_function(shell_harness).run()
    assert not at.exception
    assert at.session_state["view_mode"] == "Presentation"
    at.sidebar.radio(key="view_mode").set_value("Overview").run()
    assert not at.exception
    assert at.session_state["view_mode"] == "Overview"
    assert any(m.value == "Selected: Overview" for m in at.markdown)
    at.sidebar.radio(key="view_mode").set_value("Exports").run()
    assert not at.exception
    assert any(m.value == "Selected: Exports" for m in at.markdown)
