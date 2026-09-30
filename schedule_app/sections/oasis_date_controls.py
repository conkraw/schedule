"""OASIS Submit Date periods; reuse the existing encrypted GitHub date presets.

Separate widget keys keep the OASIS selection independent of teaching reports.
Callbacks change date widgets before they are rendered (no double-click issue).
"""
from __future__ import annotations
import hashlib
import json
from datetime import date

import streamlit as st
from schedule_app.services.evaluation_access import protected_evaluation_callback
from schedule_app.services.opd_archive import OPDArchiveError
from schedule_app.services.reporting_periods import ReportingPeriod
from schedule_app.services.reporting_presets import (
    GitHubReportingPresets, period_from_preset, preset_name_key,
)

P = "oasis_combined_dates_"


def _invalidate_output():
    for suffix in ("receipt", "publish_signature", "publish_error", "attempt_signature"):
        st.session_state.pop("oasis_combined_" + suffix, None)


@protected_evaluation_callback
def _apply_period():
    try:
        period = ReportingPeriod(st.session_state.get(P+"label", ""),
                                 st.session_state.get(P+"start"), st.session_state.get(P+"end"))
        st.session_state[P+"applied"] = period
        st.session_state.pop(P+"error", None)
        _invalidate_output()
    except OPDArchiveError as exc:
        # Invalid inputs must not leave a previous period silently active.
        st.session_state.pop(P+"applied", None)
        st.session_state[P+"error"] = str(exc)
        _invalidate_output()


@protected_evaluation_callback
def _load_preset():
    try:
        selected = st.session_state.get(P+"choice")
        entry = next(e for e in st.session_state[P+"snapshot"]["presets"] if e["id"] == selected)
        period = period_from_preset(entry)
        st.session_state[P+"start"] = period.start
        st.session_state[P+"end"] = period.end
        st.session_state[P+"label"] = period.label
        st.session_state[P+"name"] = entry["name"]
        st.session_state[P+"applied"] = period
        st.session_state.pop(P+"error", None)
        st.session_state[P+"notice"] = "Saved dates loaded. OASIS uses Submit Date, including the full end date."
        _invalidate_output()
    except (KeyError, StopIteration, OPDArchiveError):
        st.session_state[P+"error"] = "Refresh saved presets and choose an existing reporting period."


@protected_evaluation_callback
def _refresh_presets(service):
    try:
        st.session_state[P+"snapshot"] = service.load()
        st.session_state.pop(P+"catalog_error", None)
    except OPDArchiveError as exc:
        st.session_state.pop(P+"snapshot", None)
        st.session_state[P+"catalog_error"] = str(exc)


def _confirmation(action, snapshot, *values):
    digest = hashlib.sha256(json.dumps([snapshot.get("sha"), *values], sort_keys=True, default=str).encode()).hexdigest()[:20]
    return P + action + "_" + digest


@protected_evaluation_callback
def _save_preset(service, replace_id, confirm_key):
    try:
        period = st.session_state.get(P+"applied")
        name = st.session_state.get(P+"name", "")
        snapshot = st.session_state[P+"snapshot"]
        if not isinstance(period, ReportingPeriod):
            raise OPDArchiveError("Apply the reporting dates before saving a preset.")
        if replace_id and (confirm_key != _confirmation("replace", snapshot, replace_id, name, period.as_dict())
                           or not st.session_state.get(confirm_key)):
            raise OPDArchiveError("Confirm replacement of that saved preset first.")
        result = service.save(name, period, snapshot, replace_id=replace_id)
        st.session_state[P+"snapshot"] = result["snapshot"]
        st.session_state[P+"choice"] = result["preset_id"]
        st.session_state[P+"notice"] = "Date preset saved encrypted in GitHub and verified."
        st.session_state.pop(P+"error", None)
    except OPDArchiveError as exc:
        st.session_state[P+"error"] = str(exc)


@protected_evaluation_callback
def _delete_preset(service, selected, confirm_key):
    try:
        snapshot = st.session_state[P+"snapshot"]
        if (confirm_key != _confirmation("delete", snapshot, selected)
                or not st.session_state.get(confirm_key)):
            raise OPDArchiveError("Confirm deletion of the selected shared preset first.")
        result = service.delete(selected, snapshot)
        st.session_state[P+"snapshot"] = result["snapshot"]
        st.session_state[P+"choice"] = None
        st.session_state[P+"name"] = ""
        st.session_state[P+"notice"] = "Preset deleted from the saved list. Active dates, archived sources, and saved output CSVs were not deleted."
        st.session_state.pop(P+"error", None)
    except OPDArchiveError as exc:
        st.session_state[P+"error"] = str(exc)


def render(archive) -> ReportingPeriod | None:
    service = GitHubReportingPresets(archive)
    if st.session_state.get(P+"scope") != archive.config.signature():
        for key in list(st.session_state):
            if key.startswith(P):
                st.session_state.pop(key, None)
        st.session_state[P+"scope"] = archive.config.signature()
    st.markdown("### Reporting period — Submit Date")
    if P+"snapshot" not in st.session_state and P+"catalog_error" not in st.session_state:
        _refresh_presets(service)
    if st.session_state.get(P+"catalog_error"):
        st.warning(st.session_state[P+"catalog_error"])
        st.caption("You may enter dates manually; preset changes wait until GitHub can be read.")
    snapshot = st.session_state.get(P+"snapshot")
    entries = snapshot["presets"] if snapshot else []
    lookup = {e["id"]: e for e in entries}
    if st.session_state.get(P+"choice") not in lookup:
        st.session_state[P+"choice"] = None

    def display(key):
        if key is None:
            return "Choose saved reporting dates..."
        item = lookup[key]
        p = item["period"]
        return f"{item['name']} | {p['start_date']} to {p['end_date']} | {p['label']}"

    choice = st.selectbox("Saved reporting dates", [None, *lookup], format_func=display, key=P+"choice")
    left, right = st.columns(2)
    left.button("Load selected dates", disabled=choice is None, on_click=_load_preset, key=P+"load")
    right.button("Refresh saved date presets", on_click=_refresh_presets, args=(service,), key=P+"refresh")
    notice = st.session_state.pop(P+"notice", None)
    if notice:
        st.success(notice)
    with st.form(P+"form"):
        left, right = st.columns(2)
        left.date_input("Start Submit Date (included)", value=None, min_value=date(1970,1,1),
                        max_value=date(2100,12,31), key=P+"start")
        right.date_input("End Submit Date (included)", value=None, min_value=date(1970,1,1),
                         max_value=date(2100,12,31), key=P+"end")
        st.text_input("Report label / academic year", key=P+"label", placeholder="For example: 26-27")
        st.form_submit_button("Apply reporting period", on_click=_apply_period)
    if st.session_state.get(P+"error"):
        st.error(st.session_state[P+"error"])
    active = st.session_state.get(P+"applied")
    if active:
        st.caption(f"Active: {active.label} — {active.start.isoformat()} through {active.end.isoformat()}. "
                   "Both full dates are included. Submit Date is used, never the rotation Start Date. "
                   "Date edits take effect after Apply reporting period.")
    else:
        st.info("Select and apply both dates and a label, or load a saved preset. The source CSV can still be archived before you choose dates.")
    with st.expander("Save or delete reporting-date presets in GitHub"):
        st.caption("These are the same saved date presets used by PTS. "
                   "The active OASIS dates are separate. Updating/deleting a saved preset changes the shared dropdown, "
                   "not any OPDs, evaluations, output CSVs, or another screen's active dates.")
        name = st.text_input("Preset name", key=P+"name", placeholder="For example: 26-27 evaluation year")
        existing = None
        if snapshot and name.strip():
            try:
                existing = next((e for e in entries if preset_name_key(e["name"]) == preset_name_key(name)), None)
            except OPDArchiveError:
                pass
        confirm_key = None
        confirmed = True
        if existing and active:
            confirm_key = _confirmation("replace", snapshot, existing["id"], name, active.as_dict())
            confirmed = st.checkbox("Replace this existing shared date preset", key=confirm_key)
        st.button("Update date preset in GitHub" if existing else "Save date preset to GitHub",
                  disabled=not (active and snapshot and name.strip() and confirmed),
                  on_click=_save_preset, args=(service, existing["id"] if existing else None, confirm_key), key=P+"save")
        if choice in lookup:
            delete_key = _confirmation("delete", snapshot, choice)
            confirmed_delete = st.checkbox("Delete the selected saved preset (shared with teaching summary)", key=delete_key)
            st.button("Delete selected date preset", disabled=not confirmed_delete,
                      on_click=_delete_preset, args=(service, choice, delete_key), key=P+"delete")
    return active
