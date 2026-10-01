"""Missing-only student-name correction UI, called inside password-protected PTS."""
from __future__ import annotations
import hashlib
import json
import streamlit as st

from schedule_app.services.evaluation_access import evaluation_access_is_valid, lock_evaluation_records
from schedule_app.services.assessment_settings import DEFAULT_MINIMUM_SHIFTS, validate_minimum_shifts
from schedule_app.services.opd_archive import OPDArchiveError
from schedule_app.services.student_assessment_links import GitHubStudentAssessmentLinks
from schedule_app.services.student_name_review import (
    STUDENT_REVIEW_UI_VERSION, oasis_student_choices, student_name_review,
    review_table_rows, names_for_external_id, name_only_student_choices, selected_name_student_id,
)

P = "assessment_completion_name_review_"
INPUTS_KEY = "assessment_completion_inputs"
NAME_MATCH_MODE = "name_match"


def _clear_downloads():
    for key in ("teaching_zip", "teaching_zip_signature"):
        st.session_state.pop(key, None)


def _reset_controls():
    """Only call before drawing widgets; a save schedules this for the next run."""
    for key in list(st.session_state):
        if str(key).startswith(P) and key not in (P + "scope", P + "notice", P + "reset"):
            st.session_state.pop(key, None)


def _changed(inputs, saved, message):
    # A verified catalog is the source of truth, not a separate dismiss-flag list.
    # build_completion_bundle runs again on rerun before any report can be built.
    st.session_state[INPUTS_KEY] = {**inputs, "student_links": saved}
    st.session_state[P + "notice"] = message
    st.session_state[P + "reset"] = True
    _clear_downloads()
    st.rerun()


def _editor_suffix(key, catalog, editing=False):
    return hashlib.sha256(json.dumps([key, catalog.get("sha"), bool(editing)], sort_keys=True).encode()).hexdigest()[:16]


def _confirmation_key(suffix, selection, sid, chosen_name):
    """Every proposed name pair gets its own unchecked confirmation widget."""
    material = "\0".join((suffix, NAME_MATCH_MODE, selection, sid, chosen_name))
    return P + "confirm_" + hashlib.sha256(material.encode()).hexdigest()[:16]


def _render_editor(service, inputs, key, opd_name, choices, *, editing=False):
    """Ask the user about names only; attach the selected record's ID internally."""
    catalog = inputs["student_links"]
    old = catalog["entries"].get(key, {})
    suffix = _editor_suffix(key, catalog, editing)
    st.write("OPD student name: " + opd_name)
    if old:
        names = names_for_external_id(choices, old["external_id"])
        st.caption("Current saved name match: " + (" / ".join(names)
                   or "Saved student is not listed in the currently loaded OASIS exports"))

    name_choices = name_only_student_choices(choices)
    sid, chosen_name, selection = "", "", ""
    if name_choices:
        control = P + "oasis_choice_" + suffix
        options = [""] + list(name_choices)
        if st.session_state.get(control, "") not in options:
            st.session_state[control] = ""
        selection = st.selectbox("Matching OASIS student name", options, key=control,
            format_func=lambda token: ("— Select the correct OASIS student —" if not token else
                name_choices[token]["oasis_student_name"]
                + (" (duplicate name — source review needed)" if name_choices[token]["ambiguous_name"] else "")),
            help="Select the OASIS name that belongs to this OPD student. The app uses the associated OASIS record automatically; no ID entry is needed.")
        if selection:
            chosen = name_choices[selection]
            chosen_name = chosen["oasis_student_name"]
            st.write("Name match: " + opd_name + " → " + chosen_name)
            if chosen["ambiguous_name"]:
                st.warning("More than one OASIS student record uses this exact name. A name alone cannot distinguish them. "
                           "Review or correct the records in OER before saving; the app will not choose one automatically.")
            else:
                sid = selected_name_student_id(name_choices, selection)
    else:
        st.info("No usable OASIS student names are loaded. Upload the relevant student-assessment CSV in OER, "
                "then click Load / refresh evaluation completeness. Existing saved matches are kept.")

    # Do not ask for an external ID or prefill a likely person. The user confirms
    # the two displayed names; the selected OASIS record supplies the ID silently.
    confirmed = st.checkbox("I confirm these two names refer to the same student", value=False,
                            key=_confirmation_key(suffix, selection, sid, chosen_name))
    if sid and sid != old.get("external_id", ""):
        _clear_downloads()
        st.caption("This name match is not saved yet. The flag clears after the encrypted GitHub save is verified.")
    label = "Update saved student match in GitHub" if editing else "Save student match to GitHub"
    if st.button(label, key=P + ("update" if editing else "save"), disabled=not (confirmed and sid)):
        _clear_downloads()
        try:
            # Resolve again on this run, rather than trusting an older selection.
            sid = selected_name_student_id(name_choices, selection)
            saved = service.save_link(opd_name, sid, expected=catalog)
            _changed(inputs, saved, "Student-name match encrypted, saved, and verified in GitHub. "
                     "The name flags and affected percentages have been rechecked.")
        except OPDArchiveError as exc:
            st.error(str(exc))
            st.info("No flag was dismissed. After a competing change, use Refresh saved student matches and review again.")
    if editing and old:
        remove = st.checkbox("Confirm removal of this saved student match", value=False, key=P + "remove_confirm_" + suffix)
        if st.button("Remove saved student match", key=P + "remove", disabled=not remove):
            _clear_downloads()
            try:
                saved = service.remove_link(opd_name, expected=catalog)
                _changed(inputs, saved, "Saved match removed. Student names have been checked again; "
                         "an unresolved name returns to the correction list. Original files were not changed.")
            except OPDArchiveError as exc:
                st.error(str(exc))


def render_student_name_matches(archive, inputs, scan, years, unmatched, *, minimum_shifts=DEFAULT_MINIMUM_SHIFTS):
    """Flag unresolved names, retain confirmed links, and show a correction queue.

    Uses the existing encrypted ID catalog without changing its schema or any
    source CSV/OPD. Names used only for this UI never enter report data.
    """
    if not evaluation_access_is_valid(touch=True):
        lock_evaluation_records()
        return
    minimum_shifts = validate_minimum_shifts(minimum_shifts)
    scope = (STUDENT_REVIEW_UI_VERSION, archive.config.signature(),
             st.session_state.get("assessment_completion_scope"), inputs.get("commit"))
    if st.session_state.get(P + "scope") != scope:
        _reset_controls()
        st.session_state.pop(P + "notice", None)
        st.session_state[P + "scope"] = scope
        _clear_downloads()
    if st.session_state.pop(P + "reset", False):
        _reset_controls()
    notice = st.session_state.pop(P + "notice", None)
    if notice:
        st.success(notice)
    review = student_name_review(inputs, scan, years, unmatched)
    active, missing = review["active"], review["missing"]
    choices = oasis_student_choices(inputs["prepared"])
    service = GitHubStudentAssessmentLinks(archive)

    # Visible even when the correction panel is collapsed.
    if missing:
        st.warning(f"Student-name match needed: {len(missing):,} of {len(active):,} OPD student names in the selected dates "
                   "still need to be matched to the correct OASIS student name. "
                   f"{review['eligible_missing_count']:,} unresolved name(s) affect the {minimum_shifts}+ shift assessment check. "
                   "Match the two names below; no ID lookup is needed. All eligible students stay in the denominator, "
                   "and only affected completion results remain Not verified.")
    elif active:
        st.success(f"All {len(active):,} OPD student names in the selected dates have an exact or saved identity match. "
                   "This confirms identity matching, not that every assessment has been completed.")

    with st.expander("Resolve OPD student names to OASIS", expanded=bool(missing)):
        st.caption("Choose the OASIS student name that refers to the selected OPD student, then confirm the two names. "
                   "The app attaches the matching OASIS record automatically and saves your choice encrypted in GitHub. "
                   "Only unresolved OPD names appear in the first dropdown; you do not need to enter or verify an ID.")
        st.caption("Names that differ only by capitalization, spacing, or a trailing program/class label "
                   "such as (MD), (PA), (DO), or ; MD2028 match automatically when only one OASIS student fits. "
                   "Only actual name differences, missing records, or ambiguous matches need review; typos are not guessed.")
        st.caption("For a note or student entry you do not want included, use Ignore or restore OPD student entries near the top of PTS. "
                   "That removes the entry from PTS calculations and alerts; do not link a note to an actual student.")
        st.caption("Already matched names need no action. A saved name match does not mean an assessment exists. "
                   "If the correct student is absent, upload the relevant student-assessment CSV in OER and refresh this check; "
                   "do not choose someone else merely to clear the flag.")
        if st.button("Refresh saved student matches", key=P + "refresh"):
            _clear_downloads()
            try:
                saved = service.load()
                _changed(inputs, saved, "Saved student matches refreshed from GitHub and rechecked. "
                         "Use Load / refresh evaluation completeness after uploading new OASIS source data.")
            except OPDArchiveError as exc:
                st.error(str(exc))
                st.info("The existing matches were kept. A failed refresh is not treated as an empty catalog.")
        if missing:
            st.dataframe(review_table_rows(missing, minimum_shifts=minimum_shifts), hide_index=True, use_container_width=True)
            options = list(missing)
            if st.session_state.get(P + "student_choice") not in options:
                st.session_state[P + "student_choice"] = options[0]
            key = st.selectbox("OPD students needing an OASIS match", options, key=P + "student_choice",
                               format_func=lambda k: missing[k]["student_name"])
            _render_editor(service, inputs, key, missing[key]["student_name"], choices)
        elif active:
            st.info("No student-name corrections are needed for these dates. The missing-name dropdown is hidden.")
        else:
            st.info("No assigned OPD student names are available for these dates and named teaching preceptors.")

    entries = inputs["student_links"]["entries"]
    if entries:
        with st.expander("Review or correct saved student matches (optional)"):
            st.dataframe([{"OPD student name": row["student_name"],
                           "Matching OASIS student name": "; ".join(names_for_external_id(choices, row["external_id"]))
                               or "Not in the currently loaded exports",
                           "In current reporting dates": "YES" if key in active else "NO"}
                          for key, row in sorted(entries.items())], hide_index=True, use_container_width=True)
            editing = st.checkbox("Edit or remove a saved student match", key=P + "edit_saved", value=False)
            if editing:
                options = sorted(entries)
                if st.session_state.get(P + "saved_choice") not in options:
                    st.session_state[P + "saved_choice"] = options[0]
                key = st.selectbox("Saved student match to correct or remove", options, key=P + "saved_choice",
                                   format_func=lambda k: entries[k]["student_name"])
                _render_editor(service, inputs, key, entries[key]["student_name"], choices, editing=True)
