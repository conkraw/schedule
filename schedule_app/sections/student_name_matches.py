"""Missing-only student-name correction UI, called inside password-protected PTS."""
from __future__ import annotations
import hashlib
import json
import streamlit as st

from schedule_app.services.evaluation_access import evaluation_access_is_valid, lock_evaluation_records
from schedule_app.services.opd_archive import OPDArchiveError
from schedule_app.services.student_assessment_links import GitHubStudentAssessmentLinks
from schedule_app.services.student_name_review import (
    STUDENT_REVIEW_UI_VERSION, oasis_student_choices, student_name_review,
    review_table_rows, names_for_external_id, selected_student_id,
)

P = "assessment_completion_name_review_"
INPUTS_KEY = "assessment_completion_inputs"
CHOOSE = "Choose an OASIS student"
MANUAL = "Enter a verified Student External ID"


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


def _render_editor(service, inputs, key, opd_name, choices, *, editing=False):
    catalog = inputs["student_links"]
    old = catalog["entries"].get(key, {})
    suffix = _editor_suffix(key, catalog, editing)
    st.write("OPD student: " + opd_name)
    if old:
        names = names_for_external_id(choices, old["external_id"])
        st.caption("Current saved match: " + (" / ".join(names) or "ID not found in the currently loaded OASIS exports")
                   + " — Student External ID: " + old["external_id"])
    mode = st.radio("How to confirm the student", (CHOOSE, MANUAL), key=P + "method_" + suffix)
    sid, chosen_name = "", ""
    if mode == CHOOSE:
        if choices:
            control = P + "oasis_choice_" + suffix
            options = [""] + list(choices)
            if st.session_state.get(control, "") not in options:
                st.session_state[control] = ""
            selection = st.selectbox("Matching OASIS student name", options, key=control,
                format_func=lambda token: ("— Select the correct OASIS student —" if not token else
                    f"{choices[token]['oasis_student_name']} — {choices[token]['external_id']}"
                    + (" [same name, multiple IDs: verify]" if choices[token]["ambiguous_name"] else "")),
                help="Names/IDs come from the loaded student-assessment exports. No similar-name match is selected automatically.")
            if selection:
                sid = selected_student_id(choices, selection)
                chosen_name = choices[selection]["oasis_student_name"]
                st.write("Proposed match: " + opd_name + " → " + chosen_name + " (" + sid + ")")
                if choices[selection]["ambiguous_name"]:
                    st.warning("This OASIS name is attached to multiple IDs. Verify the actual student before choosing one.")
        else:
            st.info("No OASIS student names with usable external IDs are loaded. Upload the student-assessment CSV in OER, "
                    "then refresh evaluation completeness. You may enter an externally verified ID instead.")
    else:
        sid = st.text_input("Verified Student External ID", value=old.get("external_id", ""),
                            key=P + "external_id_" + suffix, max_chars=256).strip()
        st.caption("Use the actual Student External ID, not a guessed username or email. "
                   "A saved identity match cannot fill a missing ID inside an assessment form or create a missing evaluation.")
    # Tie confirmation to the selected identity. Changing a selection or typed ID
    # gets a new unchecked checkbox; no callback can bypass the PTS entry gate.
    confirm_suffix = hashlib.sha256((suffix + "\0" + mode + "\0" + sid + "\0" + chosen_name).encode()).hexdigest()[:16]
    confirmed = st.checkbox("I verified that these records refer to the same student", value=False,
                            key=P + "confirm_" + confirm_suffix)
    if sid and sid != old.get("external_id", ""):
        _clear_downloads()
        st.caption("The proposed match is not saved yet. The name stays flagged until GitHub verifies the save.")
    label = "Update saved student match in GitHub" if editing else "Save student match to GitHub"
    if st.button(label, key=P + ("update" if editing else "save"), disabled=not (confirmed and sid)):
        _clear_downloads()
        try:
            saved = service.save_link(opd_name, sid, expected=catalog)
            _changed(inputs, saved, "Student match encrypted, saved, and verified in GitHub. "
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


def render_student_name_matches(archive, inputs, scan, years, unmatched):
    """Flag unresolved names, retain confirmed links, and show a correction queue.

    Uses the existing encrypted ID catalog without changing its schema or any
    source CSV/OPD. Names used only for this UI never enter report data.
    """
    if not evaluation_access_is_valid(touch=True):
        lock_evaluation_records()
        return
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
                   "do not have a unique OASIS name/Student External ID match. "
                   f"{review['eligible_missing_count']:,} unresolved name(s) affect the 3+ shift assessment check. "
                   "The denominator is not reduced; only affected completion results remain Not verified.")
    elif active:
        st.success(f"All {len(active):,} OPD student names in the selected dates have an exact or saved identity match. "
                   "This confirms identity matching, not that every assessment has been completed.")

    with st.expander("Resolve OPD student names to OASIS", expanded=bool(missing)):
        st.caption("Choose the OASIS name belonging to the selected OPD student. The app saves OPD name → Student External ID, "
                   "not a new student username. Exact matching ignores capitalization, spacing and the '; MD2028' suffix. "
                   "No fuzzy matching is used. Only unresolved OPD names are in the first dropdown.")
        st.caption("OASIS choices come from the loaded assessment exports, which can cover more dates than this report. "
                   "Student names and IDs stay in this protected matching screen and encrypted GitHub catalog, not report tables.")
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
            st.dataframe(review_table_rows(missing), hide_index=True, use_container_width=True)
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
                           "OASIS name(s) for saved ID": "; ".join(names_for_external_id(choices, row["external_id"]))
                               or "Not in the currently loaded exports",
                           "Student External ID": row["external_id"],
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
