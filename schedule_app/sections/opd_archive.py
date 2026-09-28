"""Sidebar section: OPD Archive.

Extracted from the supplied app; this module performs no page rendering on import.
"""

from schedule_app.services.opd_archive import GitHubOPDArchive
from schedule_app.services.opd_archive import OPDArchiveError
from schedule_app.services.opd_archive import get_opd_archive_config
from schedule_app.services.opd_archive_ui import _opd_archive_picker
from schedule_app.services.opd_archive_ui import _opd_scope_archive_session
from schedule_app.services.opd_archive_ui import _opd_use_loaded_callback
import streamlit as st


def render():
    st.subheader("Encrypted OPD Archive")
    st.caption("One current original OPD per rotation. Only encrypted workbook bytes are stored in GitHub.")
    try:
        client = GitHubOPDArchive(get_opd_archive_config())
        _opd_scope_archive_session(client.config)
    except OPDArchiveError as exc:
        st.error(str(exc))
        return
    loaded = _opd_archive_picker(client, "archive_page")
    if loaded:
        st.button("Use this OPD to create student schedules", on_click=_opd_use_loaded_callback,
                  key="opd_use_loaded_in_schedule")
    st.info("A newer upload with the same first Monday replaces the current copy. Older encrypted "
            "versions remain in Git history. Rotation dates, file sizes and commit metadata are public.")
