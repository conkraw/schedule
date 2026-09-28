"""Sidebar section: Preceptor Teaching Summary.

Extracted from the supplied app; this module performs no page rendering on import.
"""

from io import BytesIO
from schedule_app.reports.teaching_export import teaching_build_zip
from schedule_app.services.opd_archive import GitHubOPDArchive
from schedule_app.services.opd_archive import OPDArchiveError
from schedule_app.services.opd_archive import get_opd_archive_config
from schedule_app.services.teaching_analysis import teaching_academic_label
from schedule_app.services.teaching_analysis import teaching_academic_start
from schedule_app.services.teaching_analysis import teaching_annual_rows
from schedule_app.services.teaching_analysis import teaching_local_today
from schedule_app.services.teaching_analysis import teaching_require_work_type_data
from schedule_app.services.teaching_analysis import teaching_scan_archives
from schedule_app.services.teaching_analysis import teaching_work_type_rows
from schedule_app.settings import TEACHING_CHAIR_SUMMARY_FILENAME
from schedule_app.settings import TEACHING_CSV_COLUMNS
from schedule_app.settings import TEACHING_NAME_ORDERS
from schedule_app.settings import TEACHING_OPD_NAME_ORDER_OVERRIDES
from schedule_app.settings import TEACHING_PRECEPTOR_NAME_MAP
from schedule_app.settings import TEACHING_REPORT_VERSION
from schedule_app.settings import TEACHING_WORK_TYPE_CSV_COLUMNS
from schedule_app.settings import TEACHING_WORK_TYPE_MAP
from schedule_app.settings import TEACHING_WORK_TYPE_ORDER
from zipfile import ZipFile
import hashlib
import json
import pandas as pd
import streamlit as st


def render():
    st.subheader("Preceptor Teaching Summary")
    st.write("Read current encrypted OPDs from GitHub, then generate a chair-friendly Word summary, "
             "overall and work-type CSVs, and one Word teaching report per preceptor.")
    st.caption("Work types: HOPE_DRIVE + ETOWN + NYES = Academic Pediatrics. Ward A, PSHCH Nursery and Complex Care "
               "stay separate. Other OPD sites retain their own work-type labels; no extra weighting is applied.")
    st.caption("All OPD sites. One student assigned to one AM/PM session = one student-shift = four educational hours. "
               "Two students at once count twice. Hours are student-weighted scheduled hours, not unique clock hours.")
    current = teaching_academic_start(teaching_local_today())
    st.info(f"Current academic year: {teaching_academic_label(current)} (July 1, {current} through June 30, {current + 1}).")
    try:
        client = GitHubOPDArchive(get_opd_archive_config())
    except OPDArchiveError as exc:
        st.error(str(exc))
        return
    order = st.selectbox("Names around '~' in archived OPDs", TEACHING_NAME_ORDERS,
                          key="teaching_name_order",
                          help="The sample OPD uses Preceptor ~ Student. No rotation list is required. "
                               "Choose the actual name order in your archive. Multiple students may be on separate rows "
                               "or separated by semicolons/newlines on the student side; commas inside names are preserved.")
    options_signature = hashlib.sha256(json.dumps(
        [TEACHING_REPORT_VERSION, client.config.signature(), order, TEACHING_PRECEPTOR_NAME_MAP,
         TEACHING_OPD_NAME_ORDER_OVERRIDES, TEACHING_WORK_TYPE_MAP, TEACHING_WORK_TYPE_ORDER], sort_keys=True).encode()).hexdigest()
    if st.session_state.get("teaching_options_signature") != options_signature:
        for key in ("teaching_scan", "teaching_zip", "teaching_zip_signature", "teaching_selected_years"):
            st.session_state.pop(key, None)
        st.session_state["teaching_options_signature"] = options_signature
    if st.button("Load / refresh archived OPDs", key="teaching_load_archives", type="primary"):
        for key in ("teaching_scan", "teaching_zip", "teaching_zip_signature", "teaching_selected_years"):
            st.session_state.pop(key, None)
        bar = st.progress(0, text="Reading the current encrypted archive...")
        try:
            with st.spinner("Decrypting OPDs and counting student assignments..."):
                scan = teaching_scan_archives(
                    client, order,
                    progress=lambda n, total: bar.progress(n / total, text=f"Read {n} of {total} current OPD files"),
                )
            st.session_state["teaching_scan"] = scan
        except OPDArchiveError as exc:
            st.error(str(exc))
        except Exception:
            # Do not echo workbook contents, credentials or learner data in the UI.
            st.error("The teaching summary could not be completed. No partial ZIP was created. "
                     "Check the workbook layout and installed requirements, then retry.")
        finally:
            bar.empty()
    scan = st.session_state.get("teaching_scan")
    if scan is None:
        st.caption("Nothing is downloaded or decrypted until you click Load / refresh archived OPDs. "
                   "This page never saves a report or decrypted OPD to GitHub.")
        return
    try:
        teaching_require_work_type_data(scan)
    except OPDArchiveError as exc:
        for state_key in ("teaching_scan", "teaching_zip", "teaching_zip_signature"):
            st.session_state.pop(state_key, None)
        st.warning(str(exc))
        return
    if not scan["sources"]:
        st.warning("No current encrypted OPDs were found in the configured GitHub archive folder.")
        return
    st.success(f"Read and decrypted {len(scan['sources'])} current OPD files. Archive snapshot: {scan['generated_at']}.")
    st.caption("Upload revisions in Create Student Schedule as usual. Click Load / refresh again to include newer saves. "
               "Only current files are counted, not their older Git versions.")
    years = sorted({row["academic_start_year"] for row in scan["monthly"]} | {current}, reverse=True)
    selected = st.multiselect("Academic year(s) to include", years, default=[current],
                              format_func=teaching_academic_label, key="teaching_selected_years",
                              help="The current year is selected by default. Select other available years to add "
                                   "a separate section for each year in the chair summary and every preceptor's Word report.")
    signature = (options_signature, scan["commit"], tuple(sorted(selected)))
    if st.session_state.get("teaching_zip_signature") != signature:
        st.session_state.pop("teaching_zip", None)
        st.session_state["teaching_zip_signature"] = signature
    annual = teaching_annual_rows(scan, selected)
    if scan["duplicate_assignments_removed"]:
        st.warning(f"{scan['duplicate_assignments_removed']} identical student-shift duplicates were counted once. "
                   "Different students assigned in the same shift still count separately.")
    if scan["unresolved_preceptor_labels"]:
        st.warning("Some assigned providers are recorded as site/slot/combined labels rather than individual names. "
                   "They are retained literally and marked for review, not attributed to a guessed person.")
        with st.expander("Provider labels to review"):
            st.write(", ".join(scan["unresolved_preceptor_labels"]))
    if scan["warnings"]:
        st.warning("Some filled assignment cells have no identifiable provider and cannot be attributed. "
                   "See the data-quality details and Report_Notes.txt.")
    if scan.get("work_type_conflicts"):
        st.warning("Some identical assignments appear in different work types. They count once under Work type needs review, "
                   "not in both settings. Details are included in the report ZIP for selected years.")
    with st.expander("Archive coverage and data-quality details"):
        st.dataframe(pd.DataFrame(scan["sources"]), hide_index=True, use_container_width=True)
        if scan["warnings"]:
            st.dataframe(pd.DataFrame(scan["warnings"]), hide_index=True, use_container_width=True)
        if scan.get("work_type_conflicts"):
            st.dataframe(pd.DataFrame(scan["work_type_conflicts"]), hide_index=True, use_container_width=True)
        st.write("Assigned-site work-type mapping")
        st.dataframe(pd.DataFrame([{"opd_site": site, "work_type": work_type}
                                  for site, work_type in scan["site_work_type_mapping"].items()]),
                     hide_index=True, use_container_width=True)
    if not annual:
        st.info("No assigned student-shifts were found for the selected academic year(s). "
                "Choose another available year, check the '~' name order, or archive OPDs with completed student assignments. "
                "Provider availability without a student does not count.")
        return
    preview = pd.DataFrame(annual, columns=TEACHING_CSV_COLUMNS)
    a, b, c = st.columns(3)
    a.metric("Preceptors / provider labels", preview["preceptor_name"].nunique())
    b.metric("Assigned student-shifts", f"{preview['no_of_shifts'].sum():,}")
    c.metric("Student-weighted educational hours", f"{preview['educational_hours'].sum():,}")
    st.markdown("**Teaching by type of work**")
    type_preview = pd.DataFrame(teaching_work_type_rows(scan, selected), columns=TEACHING_WORK_TYPE_CSV_COLUMNS)
    st.dataframe(type_preview, hide_index=True, use_container_width=True)
    with st.expander("Overall annual totals (original CSV)"):
        st.dataframe(preview, hide_index=True, use_container_width=True)
    st.caption("The original CSV remains one row per preceptor per academic year. An additional CSV breaks this down by work type. "
               "Months follow the actual assignment dates. "
               "Student names are not included in the CSV, Word reports, or source notes.")
    if st.button("Create teaching reports ZIP", key="teaching_build_zip"):
        st.session_state.pop("teaching_zip", None)
        try:
            with st.spinner("Creating work-type chair summary, CSVs and individual Word reports..."):
                zip_bytes, _ = teaching_build_zip(scan, selected)
                st.session_state["teaching_zip"] = zip_bytes
                st.session_state["teaching_zip_signature"] = signature
        except OPDArchiveError as exc:
            st.error(str(exc))
        except Exception:
            st.error("The Word/CSV export could not be completed. No partial ZIP was retained. "
                     "Check that python-docx is installed and retry.")
    if st.session_state.get("teaching_zip"):
        year_part = teaching_academic_label(selected[0]) if len(selected) == 1 else "Multiple_Academic_Years"
        st.download_button("Download CSV + Word reports (ZIP)",
                           data=st.session_state["teaching_zip"], file_name=f"Preceptor_Teaching_{year_part}.zip",
                           mime="application/zip", key="teaching_download_zip")
        # Reuse the Word bytes already packaged in the ZIP; no second scan or decryption.
        with ZipFile(BytesIO(st.session_state["teaching_zip"])) as generated_zip:
            chair_bytes = generated_zip.read(TEACHING_CHAIR_SUMMARY_FILENAME)
        st.download_button("Download chair summary only (Word)", data=chair_bytes,
                           file_name=TEACHING_CHAIR_SUMMARY_FILENAME,
                           mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                           key="teaching_download_chair_summary")
        st.caption("The ZIP includes one chair summary by work type, the unchanged summary CSV, a new work-type CSV, individual Word reports "
                   "in Preceptor_Reports, Report_Notes.txt and Archive_Sources.csv. "
                   "Treat these downloads as unencrypted staff reports; do not commit them to the public repository.")
