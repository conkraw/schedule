"""Sidebar section: Create Student Schedule.

Extracted from the supplied app; this module performs no page rendering on import.
"""

from collections import defaultdict
from datetime import timedelta
from io import BytesIO
from schedule_app.services.opd_archive import GitHubOPDArchive
from schedule_app.services.opd_archive import OPDArchiveError
from schedule_app.services.opd_archive import OPD_NAME_ORDERS
from schedule_app.services.opd_archive import OPD_XLSX_MIME
from schedule_app.services.opd_archive import get_opd_archive_config
from schedule_app.services.opd_archive_ui import _opd_archive_picker
from schedule_app.services.opd_archive_ui import _opd_archive_upload_ui
from schedule_app.services.opd_archive_ui import _opd_scope_archive_session
from schedule_app.services.opd_archive_ui import _opd_upload_changed
from schedule_app.services.student_schedules import _opd_name_key
from schedule_app.services.student_schedules import collect_opd_assignments
from schedule_app.services.student_schedules import create_ms_schedule_template
from schedule_app.services.student_schedules import populate_ms_schedule
import hashlib
import hmac
import pandas as pd
import streamlit as st


def render():
    """Render the Create Student Schedule sidebar section."""
    st.subheader("Create Student Schedule")


    # Archive the original OPD at this upload step, before producing schedules.
    try:
        archive_client = GitHubOPDArchive(get_opd_archive_config())
        _opd_scope_archive_session(archive_client.config)
    except OPDArchiveError as exc:
        st.error(str(exc))
        st.info("Complete the GitHub/Streamlit Secrets setup first. OPDs will not be processed without a verified archive.")
        st.stop()

    source_choice = st.radio("OPD source", ("Upload new/revised OPD", "Reload archived OPD"),
                             key="opd_source_choice", horizontal=True)
    raw_opd, opd_details, archive_ok = None, None, False
    if source_choice == "Upload new/revised OPD":
        opd_upload = st.file_uploader("Upload original OPD.xlsx", type=["xlsx"], key="opd_main",
                                      on_change=_opd_upload_changed)
        if opd_upload is not None:
            try:
                raw_opd, opd_details, archive_ok = _opd_archive_upload_ui(opd_upload, archive_client)
            except OPDArchiveError as exc:
                st.error(str(exc))
    else:
        loaded_opd = _opd_archive_picker(archive_client, "schedule_archive")
        if loaded_opd:
            raw_opd, opd_details, archive_ok = loaded_opd["raw"], loaded_opd["details"], True

    rot_upload = st.file_uploader("Upload Rotation Schedule (.xlsx or .csv)", type=["xlsx", "csv"],
                                  key="rot_main")
    df_rot = None
    if rot_upload is not None:
        try:
            rotation_bytes = rot_upload.getvalue()
            df_rot = (pd.read_csv(BytesIO(rotation_bytes)) if rot_upload.name.lower().endswith(".csv")
                      else pd.read_excel(BytesIO(rotation_bytes)))
        except Exception:
            st.error("The rotation list could not be read. Check the uploaded CSV or Excel file.")

    if raw_opd is None or df_rot is None:
        st.info("Upload or reload an OPD and upload its rotation list to create student schedules.")
        st.session_state.pop("opd_generated_master", None)
    else:
        if not {"legal_name", "start_date"}.issubset(df_rot.columns):
            st.error("The rotation list must contain legal_name and start_date columns.")
            st.stop()
        roster_df = df_rot.loc[df_rot["legal_name"].notna()].copy()
        roster_df["legal_name"] = roster_df["legal_name"].astype(str).str.strip()
        roster_df = roster_df.loc[roster_df["legal_name"].ne("")]
        roster_df["start_date"] = pd.to_datetime(roster_df["start_date"], errors="coerce")
        if roster_df.empty or roster_df["start_date"].isna().any():
            st.error("Each student in the rotation list must have a valid start_date.")
            st.stop()
        roster_mondays = {
            value.date() - timedelta(days=value.weekday()) for value in roster_df["start_date"]
        }
        if roster_mondays != {opd_details["rotation_start"]}:
            st.error(f"The rotation list does not match this OPD's first Monday "
                     f"({opd_details['rotation_start']:%m/%d/%Y}). Upload the matching rotation list. "
                     "The original OPD is archived under its own dates, not the roster's dates.")
            st.stop()
        if len(opd_details["week_mondays"]) != 4:
            st.error("The original has been archived, but the existing MS_Schedule template requires exactly four weeks.")
            st.stop()

        students = list({_opd_name_key(name): name for name in roster_df["legal_name"]}.values())
        name_order = st.selectbox("Names around '~' in the OPD", OPD_NAME_ORDERS,
                                  key="opd_name_order",
                                  help="Auto-detect matches the student to legal_name in the rotation list. "
                                       "It accepts either name order and ignores spaces around '~'.")
        signature = (hashlib.sha256(raw_opd).hexdigest(), hashlib.sha256(rotation_bytes).hexdigest(), name_order)
        if st.session_state.get("opd_master_signature") != signature:
            st.session_state.pop("opd_generated_master", None)
            st.session_state["opd_master_signature"] = signature
        try:
            assignments, unmatched = collect_opd_assignments(raw_opd, students, name_order)
        except OPDArchiveError as exc:
            st.error(str(exc))
            st.stop()
        if unmatched:
            st.warning(f"{len(unmatched)} filled assignment cell(s) do not match this rotation list/name order "
                       "and will not populate a student schedule. This can include students from another course.")
            with st.expander("Unmatched assignment cells"):
                st.write(", ".join(unmatched))
        assigned_students = {item["student"] for item in assignments}
        if any(name not in assigned_students for name in students):
            st.warning("Some students have no matching OPD assignments. Review the roster and name-order setting before building.")
        conflicts = defaultdict(list)
        for item in assignments:
            conflicts[(item["student"], item["week"], item["shift"], item["column"])].append(item)
        has_conflicts = False
        days = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
        for (student, week, shift, col), items in conflicts.items():
            if len(items) > 1:
                has_conflicts = True
                locations = "; ".join(f"{item['site']}@{item['coordinate']}" for item in items)
                st.warning(f"Week {week+1} {days[col-2]} {shift}: {student} double-booked ({locations}). "
                           "As in the previous version, the last assignment read fills that schedule cell.")
        if not has_conflicts:
            st.success("No AM/PM shift conflicts detected among matched students.")
        st.caption(f"OPD rotation: {opd_details['rotation_start']:%B %d, %Y}; "
                   f"{len(students)} student(s); {len(assignments)} matched assignment cell(s).")
        if not archive_ok:
            st.error("Schedule generation is disabled until the original OPD is successfully archived.")
            st.session_state.pop("opd_generated_master", None)
        if st.button("Create & Download Fully-Populated MS_Schedule", disabled=not archive_ok,
                     key="opd_build_master"):
            try:
                # Do not silently reuse or overwrite an older session's source.
                current_archive = archive_client.load(opd_details["rotation_start"])
                if not hmac.compare_digest(current_archive["raw"], raw_opd):
                    raise OPDArchiveError("A different OPD is now current in the archive for this rotation. "
                                          "Reload the latest archived OPD, or deliberately upload your revised file again.")
                dates = pd.date_range(start=opd_details["rotation_start"], periods=28, freq="D").tolist()
                blank_buf = create_ms_schedule_template(students, dates)
                st.session_state["opd_generated_master"] = populate_ms_schedule(blank_buf.getvalue(), assignments)
            except OPDArchiveError as exc:
                st.session_state.pop("opd_generated_master", None)
                st.error(str(exc))
        if archive_ok and st.session_state.get("opd_generated_master"):
            st.download_button("Download MS_Schedule.xlsx", data=st.session_state["opd_generated_master"],
                               file_name="MS_Schedule.xlsx", mime=OPD_XLSX_MIME, key="opd_master_download")
            st.caption("Next: open Create Individual Schedules and upload this MS_Schedule.xlsx. "
                       "The individual student ZIP and Power Automate preceptor workbook remain available there.")
