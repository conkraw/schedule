"""Reporting-date inputs with encrypted GitHub presets (no JSON file uploads).

All network writes are explicit button callbacks. Widget values are changed in
callbacks, before the next render, never after date widgets are instantiated.
"""
from datetime import date
import hashlib
import json
import streamlit as st
from schedule_app.services.evaluation_access import protected_evaluation_callback

from schedule_app.services.opd_archive import GitHubOPDArchive, OPDArchiveError, get_opd_archive_config
from schedule_app.services.reporting_periods import ReportingPeriod
from schedule_app.services.reporting_presets import (
    GitHubReportingPresets, ReportingPresetError, normalize_preset_name,
    preset_name_key, period_from_preset,
)

REPORTING_MODES = ("Custom dates", "Standard July-June academic years")
PERIOD_FIELDS = {"label": "teaching_period_label", "start_date": "teaching_period_start",
                 "end_date": "teaching_period_end", "mode": "teaching_reporting_mode"}
SNAPSHOT = "teaching_presets_snapshot"
SCOPE = "teaching_presets_scope"
ERROR = "teaching_presets_error"
NOTICE = "teaching_presets_notice"
CHOICE = "teaching_preset_choice"
NAME = "teaching_preset_name"


def clear_teaching_downloads():
    st.session_state.pop("teaching_zip", None)
    st.session_state.pop("teaching_zip_signature", None)
    st.session_state.pop("teaching_zip_payload", None)
    st.session_state.pop("teaching_build_seconds", None)


@protected_evaluation_callback
def remember_period_inputs():
    """Keep preferences across sidebar navigation without persisting every edit."""
    previous = dict(st.session_state.get("teaching_period_preferences", {}))
    preferences = dict(previous)
    for name, key in PERIOD_FIELDS.items():
        if key in st.session_state:
            preferences[name] = st.session_state[key]
    st.session_state["teaching_period_preferences"] = preferences
    if preferences != previous:
        clear_teaching_downloads()


def _service():
    return GitHubReportingPresets(GitHubOPDArchive(get_opd_archive_config()))


def _scope(service):
    signature = service.config.signature()
    if st.session_state.get(SCOPE) != signature:
        for key in list(st.session_state):
            if str(key).startswith(("teaching_preset",)):
                st.session_state.pop(key, None)
        st.session_state[SCOPE] = signature


def _error(exc):
    # Make a failure visible and require refresh, rather than treating it as no presets.
    st.session_state[ERROR] = str(exc)
    st.session_state.pop(SNAPSHOT, None)
    st.session_state.pop(NOTICE, None)


@protected_evaluation_callback
def refresh_saved_presets():
    try:
        service = _service()
        _scope(service)
        snapshot = service.load()
        st.session_state[SNAPSHOT] = snapshot
        st.session_state.pop(ERROR, None)
        ids = {row["id"] for row in snapshot["presets"]}
        if st.session_state.get(CHOICE) not in ids:
            st.session_state[CHOICE] = None
        st.session_state[NOTICE] = "Saved date presets refreshed from GitHub."
    except OPDArchiveError as exc:
        _error(exc)


def _snapshot_for(service):
    snapshot = st.session_state.get(SNAPSHOT)
    if not snapshot or snapshot.get("scope") != service.config.signature():
        raise ReportingPresetError("Refresh saved presets for this repository before using or changing them.")
    return snapshot


@protected_evaluation_callback
def load_selected_preset(preset_id):
    """Fetch the latest catalog before applying dates from the selected dropdown row."""
    try:
        service = _service()
        previous = _snapshot_for(service)
        snapshot = service.load()
        st.session_state[SNAPSHOT] = snapshot
        entry = next((row for row in snapshot["presets"] if row["id"] == preset_id), None)
        old = next((row for row in previous["presets"] if row["id"] == preset_id), None)
        if entry is None:
            st.session_state[CHOICE] = None
            raise ReportingPresetError("That preset was deleted in another session. Refresh saved presets.")
        if old is None or entry != old:
            st.session_state[NOTICE] = "That preset changed in GitHub. Review its current dates in the dropdown, then load it again."
            return
        period = period_from_preset(entry)
        values = {"label": period.label, "start_date": period.start, "end_date": period.end,
                  "mode": REPORTING_MODES[0]}
        for field, key in PERIOD_FIELDS.items():
            st.session_state[key] = values[field]
        st.session_state["teaching_period_preferences"] = values
        st.session_state[NAME] = entry["name"]
        st.session_state[CHOICE] = entry["id"]
        st.session_state.pop(ERROR, None)
        st.session_state[NOTICE] = "Saved dates and report label loaded. Create the reports again for this period."
        clear_teaching_downloads()  # Retain the already-decrypted OPD scan.
    except OPDArchiveError as exc:
        _error(exc)


@protected_evaluation_callback
def save_current_preset(replace_id=None, confirmation_key=None):
    try:
        service = _service()
        snapshot = _snapshot_for(service)
        if st.session_state.get("teaching_reporting_mode") != REPORTING_MODES[0]:
            raise ReportingPresetError("Switch to Custom dates to save a named date preset.")
        name = normalize_preset_name(st.session_state.get(NAME, ""))
        period = ReportingPeriod(st.session_state.get("teaching_period_label", ""),
                                 st.session_state.get("teaching_period_start"),
                                 st.session_state.get("teaching_period_end"))
        if replace_id is not None and (
            confirmation_key != _confirmation_key("replace", snapshot, replace_id, name, period.as_dict())
            or not st.session_state.get(confirmation_key, False)
        ):
            raise ReportingPresetError("Confirm replacement of the named preset before saving over it.")
        result = service.save(name, period, snapshot, replace_id=replace_id)
        st.session_state[SNAPSHOT] = result["snapshot"]
        st.session_state[CHOICE] = result["preset_id"]
        st.session_state[NAME] = name
        st.session_state.pop(ERROR, None)
        st.session_state[NOTICE] = {
            "created": "Date preset saved and verified in GitHub. It will be available in later sessions.",
            "updated": "Saved date preset updated and verified in GitHub.",
            "unchanged": "These exact date settings are already saved. No additional commit was needed.",
        }[result["action"]]
    except OPDArchiveError as exc:
        _error(exc)


@protected_evaluation_callback
def delete_selected_preset(preset_id, confirmation_key):
    try:
        service = _service()
        snapshot = _snapshot_for(service)
        if (confirmation_key != _confirmation_key("delete", snapshot, preset_id)
            or not st.session_state.get(confirmation_key, False)):
            raise ReportingPresetError("Confirm deletion of the selected date preset first.")
        result = service.delete(preset_id, snapshot)
        st.session_state[SNAPSHOT] = result["snapshot"]
        st.session_state[CHOICE] = None
        st.session_state[NAME] = ""
        st.session_state.pop(ERROR, None)
        st.session_state[NOTICE] = (
            "Date preset deleted from the current saved list. Your OPDs and current reporting dates were not changed."
        )
    except OPDArchiveError as exc:
        _error(exc)


def _confirmation_key(action, snapshot, *values):
    # Consent applies only to this exact catalog revision, preset and proposed values.
    material = json.dumps([snapshot.get("sha"), *values], ensure_ascii=True, sort_keys=True)
    return f"teaching_preset_confirm_{action}_" + hashlib.sha256(material.encode()).hexdigest()[:20]


def _render_preset_picker():
    try:
        service = _service()
        _scope(service)
        if SNAPSHOT not in st.session_state and ERROR not in st.session_state:
            with st.spinner("Loading saved date presets from GitHub..."):
                st.session_state[SNAPSHOT] = service.load()
    except OPDArchiveError as exc:
        _error(exc)
    if st.session_state.get(ERROR):
        st.warning(st.session_state[ERROR])
        st.caption("You can still enter dates manually. Saving/deleting presets is unavailable until the list refresh succeeds.")
    notice = st.session_state.pop(NOTICE, None)
    if notice:
        st.success(notice)
    snapshot = st.session_state.get(SNAPSHOT)
    entries = snapshot["presets"] if snapshot else []
    by_id = {row["id"]: row for row in entries}
    if st.session_state.get(CHOICE) not in by_id:
        st.session_state[CHOICE] = None

    def label(preset_id):
        if preset_id is None:
            return "Choose saved dates..."
        row = by_id[preset_id]
        period = period_from_preset(row)
        return f"{row['name']} | {period.start:%m/%d/%Y} - {period.end:%m/%d/%Y} | Label: {period.label}"

    selected = st.selectbox("Saved date presets", [None, *by_id], key=CHOICE,
                            format_func=label, disabled=not entries,
                            help="Shared presets stored in GitHub. Select one and click Load selected dates.")
    left, right = st.columns(2)
    left.button("Load selected dates", key="teaching_preset_load", disabled=selected is None,
                on_click=load_selected_preset, args=(selected,))
    right.button("Refresh saved presets", key="teaching_preset_refresh", on_click=refresh_saved_presets)
    if snapshot is not None and not entries:
        st.caption("No saved date presets yet. Enter dates and a report label below, then save a named preset.")
    return snapshot, selected


def _render_save_delete(period, snapshot, selected):
    with st.expander("Save or delete date presets in GitHub", expanded=False):
        st.caption("Presets are shared across sessions and use your existing encrypted archive repository and key. "
                   "Only dates, labels and preset names are saved here; no OPDs or student records are changed.")
        name = st.text_input("Preset name", key=NAME, max_chars=80,
                              placeholder="e.g., 26-27 teaching year or Chair review",
                              help="Use a unique name to add a preset. Reuse a saved name to update it after confirmation.")
        valid_name = None
        try:
            if name:
                valid_name = normalize_preset_name(name)
        except OPDArchiveError as exc:
            st.warning(str(exc))
        entries = snapshot["presets"] if snapshot else []
        existing = next((row for row in entries if valid_name and
                         preset_name_key(row["name"]) == preset_name_key(valid_name)), None)
        save_confirm = None
        confirmed = existing is None
        if existing and period is not None:
            save_confirm = _confirmation_key("replace", snapshot, existing["id"], valid_name, period.as_dict())
            old = period_from_preset(existing)
            st.caption(f"Existing preset: {old.start:%m/%d/%Y} - {old.end:%m/%d/%Y}; report label {old.label}.")
            confirmed = st.checkbox("Confirm replacing this saved preset with the dates and label above",
                                     key=save_confirm)
        st.button("Update saved preset in GitHub" if existing else "Save these dates to GitHub",
                  key="teaching_preset_save", disabled=(snapshot is None or period is None or not valid_name or not confirmed),
                  on_click=save_current_preset, args=(existing["id"] if existing else None, save_confirm))
        if period is None:
            st.caption("Choose Custom dates and enter both dates and a report label to save a preset.")
        if selected is not None and snapshot is not None:
            entry = next((row for row in entries if row["id"] == selected), None)
            if entry:
                st.markdown("**Delete the selected preset**")
                st.write(f"Selected preset: {entry['name']}")
                confirm_key = _confirmation_key("delete", snapshot, selected)
                confirmed_delete = st.checkbox("Confirm deletion of this date preset only (not OPDs)", key=confirm_key)
                st.button("Delete selected preset", key="teaching_preset_delete", disabled=not confirmed_delete,
                          on_click=delete_selected_preset, args=(selected, confirm_key))
                st.caption("Deletion removes it from the dropdown, not from older Git history. Your current date fields remain unchanged.")


def render_period_controls():
    preferences = st.session_state.get("teaching_period_preferences", {})
    for name, key in PERIOD_FIELDS.items():
        if key not in st.session_state and name in preferences:
            st.session_state[key] = preferences[name]
    if st.session_state.get("teaching_reporting_mode") not in REPORTING_MODES:
        st.session_state["teaching_reporting_mode"] = REPORTING_MODES[0]
    st.markdown("**Reporting period**")
    snapshot, selected = _render_preset_picker()
    mode = st.radio("Choose reporting dates", REPORTING_MODES, key="teaching_reporting_mode",
                    horizontal=True, on_change=remember_period_inputs)
    period, issue = None, None
    if mode == REPORTING_MODES[0]:
        left, right = st.columns(2)
        start = left.date_input("Start date (included)", value=None, min_value=date(1970, 1, 1),
                               max_value=date(2100, 12, 31), format="MM/DD/YYYY",
                               key="teaching_period_start", on_change=remember_period_inputs)
        end = right.date_input("End date (included)", value=None, min_value=date(1970, 1, 1),
                              max_value=date(2100, 12, 31), format="MM/DD/YYYY",
                              key="teaching_period_end", on_change=remember_period_inputs)
        label = st.text_input("Report label / academic year", value="", placeholder="e.g., 26-27",
                              max_chars=60, key="teaching_period_label", on_change=remember_period_inputs,
                              help="This label goes in the academic_year CSV column and every Word report. "
                                   "It does not determine the dates or force a July boundary.")
        try:
            period = ReportingPeriod(label, start, end)
        except OPDArchiveError as exc:
            issue = str(exc)
        st.caption("Choose the exact dates you need, even for a period longer than 12 months. "
                   "Both dates and weekends are included. July 1 does not split a custom report.")
    else:
        st.caption("Optional original mode: July 1 through June 30, with a separate section for each selected year. "
                   "Loading a saved date preset switches to Custom dates.")
    remember_period_inputs()
    _render_save_delete(period, snapshot, selected)
    return mode, period, issue
