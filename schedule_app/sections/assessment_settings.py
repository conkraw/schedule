"""PTS-only minimum-shifts control with verified automatic GitHub persistence."""
from __future__ import annotations
import streamlit as st
from schedule_app.services.assessment_settings import (
    DEFAULT_MINIMUM_SHIFTS, MAX_MINIMUM_SHIFTS,
    GitHubAssessmentSettings, validate_minimum_shifts,
)
from schedule_app.services.evaluation_access import evaluation_access_is_valid, lock_evaluation_records
from schedule_app.services.opd_archive import OPDArchiveError

P = "assessment_completion_settings_"


def _clear_downloads():
    for key in ("teaching_zip", "teaching_zip_signature", "teaching_report_issues"):
        st.session_state.pop(key, None)


def render_assessment_threshold(archive):
    """Return the verified chosen integer, or None on an unconfirmed load/save.

    A number-input edit triggers a rerun; the entry gate is checked BEFORE the
    edit is persisted. No network-writing callbacks run before authentication.
    Only an explicit new value/retry writes. Old sessions cannot auto-overwrite a
    concurrent edit; a revision conflict needs an explicit reload and review.
    """
    if not evaluation_access_is_valid(touch=True):
        lock_evaluation_records()
        return None
    signature = archive.config.signature()
    if st.session_state.get(P + "scope") != signature:
        for key in list(st.session_state):
            if str(key).startswith(P):
                st.session_state.pop(key, None)
        st.session_state[P + "scope"] = signature
        _clear_downloads()
    service = GitHubAssessmentSettings(archive)
    if st.button("Reload saved minimum shifts", key=P + "reload",
                 help="Reload the shared setting from GitHub, including changes made in another session."):
        for suffix in ("snapshot", "load_error", "save_error", "attempt", "value", "last_choice"):
            st.session_state.pop(P + suffix, None)
        _clear_downloads()
    if P + "snapshot" not in st.session_state and P + "load_error" not in st.session_state:
        try:
            st.session_state[P + "snapshot"] = service.load()
        except OPDArchiveError as exc:
            st.session_state[P + "load_error"] = str(exc)
    if P + "load_error" in st.session_state:
        _clear_downloads()
        st.warning("The saved minimum-shifts setting could not be loaded. " + st.session_state[P + "load_error"])
        st.caption("It has not been reset to 3. Use Reload saved minimum shifts to retry; "
                   "completion percentages stay Not checked until the setting is verified.")
        return None
    saved = st.session_state[P + "snapshot"]
    if P + "value" not in st.session_state:
        st.session_state[P + "value"] = saved["minimum_shifts"]
    chosen = st.number_input("Minimum shifts for assessment completion", min_value=1,
                             max_value=MAX_MINIMUM_SHIFTS, step=1, format="%d", key=P + "value",
                             help="Count a student in a preceptor's assessment-completion denominator only after "
                                  "this many distinct AM/PM shifts together in the selected dates. "
                                  "Changes save automatically, encrypted in GitHub. This is not the 3+ days continuity measure.")
    try:
        chosen = validate_minimum_shifts(chosen)
    except OPDArchiveError as exc:
        _clear_downloads()
        st.warning(str(exc))
        return None
    if st.session_state.get(P + "last_choice", saved["minimum_shifts"]) != chosen:
        _clear_downloads()
    st.session_state[P + "last_choice"] = chosen
    retry = False
    if P + "save_error" in st.session_state:
        retry = st.button("Retry saving minimum shifts", key=P + "retry")
    # A failed verification cannot be cleared merely by making another widget
    # interaction. It is cleared by explicit reload or a verified retry/change.
    pending = chosen != saved["minimum_shifts"] or P + "save_error" in st.session_state
    attempt = (saved.get("sha"), chosen)
    if pending and (retry or st.session_state.get(P + "attempt") != attempt):
        st.session_state[P + "attempt"] = attempt
        _clear_downloads()
        try:
            with st.spinner("Saving the minimum shifts encrypted in GitHub..."):
                saved = service.save(chosen, expected=saved)
            st.session_state[P + "snapshot"] = saved
            st.session_state.pop(P + "save_error", None)
        except OPDArchiveError as exc:
            st.session_state[P + "save_error"] = str(exc)
    if P + "save_error" in st.session_state:
        _clear_downloads()
        st.warning("Minimum-shifts change is NOT confirmed saved. " + st.session_state[P + "save_error"])
        st.caption("Use Retry saving minimum shifts, or reload the saved setting. "
                   "No completion percentages use this unconfirmed change.")
        return None
    if saved.get("sha"):
        st.success(f"Minimum shifts: {chosen}. Saved encrypted in GitHub and reused next time.")
    else:
        st.caption(f"Starting minimum: {DEFAULT_MINIMUM_SHIFTS} shifts. Your first change will be saved encrypted in GitHub automatically.")
    st.caption("This is one shared PTS setting across reporting dates and app users. "
               "Only the assessment-completion denominator changes; educational hours, Learner Reach, "
               "unique-student totals and the separate 3+ days measure do not change.")
    return chosen
