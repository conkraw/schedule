"""Reversible, name-only OPD exclusions inside password-protected PTS.

The inventory is protected session data only. The catalog contains only the
explicitly excluded labels; no source records or student IDs are rewritten.
"""
from __future__ import annotations
import hashlib
import json
import streamlit as st
from schedule_app.services.evaluation_access import evaluation_access_is_valid, lock_evaluation_records
from schedule_app.services.ignored_student_entries import (
    GitHubIgnoredStudentEntries, exclusion_signature, ignored_entry_key,
)
from schedule_app.services.opd_archive import OPDArchiveError

P = "pts_ignored_students_"


def _invalidate():
    for key in ("teaching_scan", "teaching_zip", "teaching_zip_signature",
                "assessment_completion_inputs", "assessment_completion_scope",
                "assessment_completion_failure"):
        st.session_state.pop(key, None)


def _apply_saved(catalog, message, archive, order):
    old_scan = st.session_state.get("teaching_scan")
    if old_scan:
        st.session_state[P + "rescan"] = {
            "commit": old_scan["commit"], "scope": archive.config.signature(),
            "order": order,
        }
    st.session_state[P + "catalog"] = catalog
    st.session_state[P + "notice"] = message
    st.session_state[P + "reset_controls"] = True
    st.session_state.pop(P + "error", None)
    _invalidate()
    st.rerun()


def _control_token(names, sha):
    return hashlib.sha256(json.dumps([sorted(names), sha]).encode()).hexdigest()[:16]


def load_ignored_student_entries_ui(archive, order):
    """Return the verified catalog, or None; never silently use an empty list on failure."""
    if not evaluation_access_is_valid(touch=True):
        lock_evaluation_records()
        return None
    service = GitHubIgnoredStudentEntries(archive)
    scope = (archive.config.signature(), order)
    if st.session_state.get(P + "scope") != scope:
        for key in list(st.session_state):
            if str(key).startswith(P):
                st.session_state.pop(key, None)
        st.session_state[P + "scope"] = scope
    if st.session_state.pop(P + "reset_controls", False):
        for key in list(st.session_state):
            if str(key).startswith((P + "choose_", P + "confirm_", P + "restore_", P + "manual")):
                st.session_state.pop(key, None)
    if P + "catalog" not in st.session_state and P + "error" not in st.session_state:
        try:
            st.session_state[P + "catalog"] = service.load()
        except OPDArchiveError as exc:
            st.session_state[P + "error"] = str(exc)
            _invalidate()
    error = st.session_state.get(P + "error")
    if error:
        st.error("The saved ignored-student list could not be checked. " + error)
        st.info("PTS reports wait for a verified list; the app will not assume it is empty. Your source files are unchanged.")
        if st.button("Retry loading ignored student entries", key=P + "retry"):
            st.session_state.pop(P + "error", None)
            st.session_state.pop(P + "catalog", None)
            st.rerun()
        return None
    return st.session_state[P + "catalog"]


def render_ignored_student_entries(archive, order, *, period=None):
    """Display names after scanning, so the dropdown updates on the same run."""
    catalog = load_ignored_student_entries_ui(archive, order)
    if catalog is None:
        return None
    service = GitHubIgnoredStudentEntries(archive)
    catalog = st.session_state[P + "catalog"]
    entries = catalog["entries"]
    notice = st.session_state.pop(P + "notice", None)
    if notice:
        st.success(notice)
        st.caption("The teaching scan is recalculated after a change. Refresh evaluation completeness below before generating completion percentages.")
    inventory = st.session_state.get(P + "inventory", [])
    # The UI inventory contains original labels even after they are ignored, so
    # an excluded-only source never traps the user behind an empty-report return.
    shown = [row for row in inventory if period is None or any(
        period.start.isoformat() <= day <= period.end.isoformat() for day in row["dates"])]
    choices = {ignored_entry_key(row["student_entry"]): row["student_entry"] for row in shown
               if ignored_entry_key(row["student_entry"]) not in entries}
    with st.expander(f"Ignore or restore OPD student entries ({len(entries)} saved)"):
        st.caption("Use this for notes such as midcycle feedback or student entries you deliberately exclude from PTS. "
                   "Ignored entries will not count as students, teaching assignments, or assessment-eligible students. "
                   "The preceptor's recorded clinical availability is retained. Original OPD and OASIS files are never edited.")
        st.warning("This changes the reporting population, not just the warning display. Ignore only entries you intend to exclude—not a real student merely because their evaluation is missing.")
        st.caption("The list is shared across PTS users and all reporting dates. It matches the exact entry, ignoring capitalization and extra spacing; "
                   "different spellings or program designations are separate exclusions. Nothing is excluded automatically by a guessed keyword.")
        if st.button("Refresh ignored student entries from GitHub", key=P + "refresh"):
            try:
                current = service.load()
                if exclusion_signature(entries) != exclusion_signature(current["entries"]):
                    _apply_saved(current, "Ignored-student list refreshed from GitHub.", archive, order)
                st.session_state[P + "catalog"] = current
                st.session_state[P + "notice"] = "Ignored-student list refreshed from GitHub; no filtering change."
                st.session_state[P + "reset_controls"] = True
                st.rerun()
            except OPDArchiveError as exc:
                st.session_state[P + "error"] = str(exc)
                _invalidate()
                st.rerun()
        if not inventory:
            st.info("Load / refresh archived OPDs below to populate this name-only list. You can also enter an exact OPD label.")
        revision = _control_token([], catalog.get("sha"))
        select_key = P + "choose_" + revision
        if select_key in st.session_state:
            st.session_state[select_key] = [key for key in st.session_state[select_key] if key in choices]
        selected = st.multiselect("OPD student entries to ignore", sorted(choices), key=select_key,
                                  format_func=lambda key: choices[key])
        manual = st.text_input("Exact OPD entry not in the list (optional)", key=P + "manual", max_chars=256).strip()
        names = [choices[key] for key in selected if key in choices]
        if manual and ignored_entry_key(manual) not in entries:
            names.append(manual)
        names = list({ignored_entry_key(name): name for name in names}.values())
        checked = st.checkbox("Exclude these entries from PTS counts and name-match alerts for all reporting dates",
                              value=False, key=P + "confirm_" + _control_token(names, catalog.get("sha")))
        if st.button("Ignore selected entries and save to GitHub", key=P + "save", disabled=not (checked and names)):
            try:
                saved = service.ignore_entries(names, expected=catalog)
                _apply_saved(saved, "Ignored entries encrypted, saved, and verified in GitHub.", archive, order)
            except OPDArchiveError as exc:
                st.error(str(exc))
                st.info("No entries were dismissed. Refresh the saved list before retrying if another session changed it.")
        if entries:
            st.markdown("**Ignored entries — restore when needed**")
            st.dataframe([{"Ignored OPD student entry": row["student_entry"], "Saved (UTC)": row["updated_at"]}
                          for _, row in sorted(entries.items())], hide_index=True, use_container_width=True)
            restore_key = P + "restore_choose_" + revision
            restored = st.multiselect("Ignored student entries to restore", sorted(entries), key=restore_key,
                                      format_func=lambda key: entries[key]["student_entry"])
            restore_confirm = st.checkbox("Restore these entries to PTS counts and matching checks", value=False,
                key=P + "restore_confirm_" + _control_token(restored, catalog.get("sha")))
            if st.button("Restore selected entries", key=P + "restore", disabled=not (restore_confirm and restored)):
                try:
                    saved = service.restore_entries([entries[key]["student_entry"] for key in restored], expected=catalog)
                    _apply_saved(saved, "Selected entries restored. Saved student-name corrections were retained.", archive, order)
                except OPDArchiveError as exc:
                    st.error(str(exc))
    return catalog
