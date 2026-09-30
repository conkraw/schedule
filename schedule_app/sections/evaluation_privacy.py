"""Explicit current-file cleanup, available only behind OER login."""
from __future__ import annotations
import streamlit as st
from schedule_app.services.opd_archive import OPDArchiveError
from schedule_app.services.oasis_evaluations import GitHubOASISEvaluations
from schedule_app.services.oasis_student_evaluations import GitHubOASISStudentEvaluations
from schedule_app.services.oasis_privacy import EDUCATOR_COLUMNS, STUDENT_COLUMNS
from schedule_app.services.evaluation_access import evaluation_access_is_valid

P = "evaluation_privacy_"


def render(archive):
    if not evaluation_access_is_valid():
        return
    scope = archive.config.signature()
    if st.session_state.get(P + "scope") != scope:
        for k in list(st.session_state):
            if str(k).startswith(P):
                st.session_state.pop(k, None)
        st.session_state[P + "scope"] = scope
    with st.expander("Stored columns and older-file privacy review"):
        st.write("New uploads keep only the reporting fields below, then encrypt the reduced CSV. "
                 "Unknown columns are dropped, not added to the allowlist automatically.")
        st.write("Educator feedback — " + ", ".join(EDUCATOR_COLUMNS))
        st.write("Student assessments — " + ", ".join(STUDENT_COLUMNS))
        st.warning("Older full exports may still exist in GitHub. This optional cleanup replaces their CURRENT files only. "
                   "Older encrypted versions can remain in Git history, clones, or other copies. It is not permanent erasure.")
        st.caption("Old exports are minimized in memory before review/download even before cleanup. "
                   "OPDs, date presets, username links and generated summary CSVs are not edited by this tool.")
        if st.button("Check previously saved evaluation files", key=P + "scan"):
            st.session_state.pop(P + "rows", None)
            st.session_state.pop(P + "error", None)
            rows = []
            try:
                commit = archive._head()
                with st.spinner("Checking current encrypted evaluation sources..."):
                    for kind, cls in (("educator", GitHubOASISEvaluations), ("student", GitHubOASISStudentEvaluations)):
                        service = cls(archive)
                        files = service.list_exports(commit=commit)["filenames"]
                        if len(files) > 100:
                            raise OPDArchiveError("More than 100 files in this evaluation archive. No partial cleanup list was kept.")
                        for filename in files:
                            item = service.load(filename, commit=commit)
                            policy = item["privacy"]
                            rows.append({"type": kind, "file": filename, "sha": item["sha"],
                                         "stored_columns": policy["source_column_count"],
                                         "needed_columns": policy["retained_column_count"],
                                         "columns_to_remove": policy["removed_column_count"],
                                         "replace_current_copy": policy["needs_minimization"]})
                st.session_state[P + "rows"] = rows
                st.session_state.pop(P + "done", None)
            except OPDArchiveError as exc:
                st.error(str(exc))
            except Exception:
                st.error("The privacy review could not finish. No partial cleanup list was kept; retry the check.")
        if st.session_state.get(P + "done"):
            st.success(st.session_state[P + "done"])
        if st.session_state.get(P + "error"):
            st.error(st.session_state[P + "error"])
        rows = st.session_state.get(P + "rows")
        if rows is None:
            return
        st.dataframe([{k: v for k, v in row.items() if k != "sha"} for row in rows],
                     hide_index=True, use_container_width=True)
        pending = [r for r in rows if r["replace_current_copy"]]
        if not pending:
            st.success("All checked current evaluation files already use the reduced format.")
            return
        st.info(f"{len(pending)} saved file(s) can be replaced with a reduced copy. "
                "After replacement, removed columns cannot be recovered through this app.")
        # Tie confirmation to this precise review, not to a different later scan.
        import hashlib
        tag = hashlib.sha256(repr([(r["file"], r["sha"]) for r in pending]).encode()).hexdigest()[:16]
        confirmed = st.checkbox("Replace these current files after verifying reduced copies. I understand Git history is not erased.",
                                key=P + "confirm_" + tag)
        if st.button("Minimize previously saved evaluation files", disabled=not confirmed, key=P + "apply"):
            completed = 0
            failure = None
            try:
                with st.spinner("Saving and verifying reduced copies before removing old current paths..."):
                    for row in pending:
                        cls = GitHubOASISStudentEvaluations if row["type"] == "student" else GitHubOASISEvaluations
                        cls(archive).minimize_saved_export(row["file"], expected_sha=row["sha"])
                        completed += 1
            except OPDArchiveError as exc:
                failure = str(exc)
            except Exception:
                failure = "A cleanup operation could not be confirmed. Check the current files before retrying."
            # No later widget has yet been created by the parent evaluation UI.
            # Clear its stale source/output snapshots; leave unrelated OPD state alone.
            for k in list(st.session_state):
                if str(k).startswith("oasis_"):
                    st.session_state.pop(k, None)
            st.session_state.pop(P + "rows", None)
            st.session_state[P + "done"] = (
                f"{completed} current file(s) minimized and verified. Previous Git history remains. "
                "Recheck the archive before any further cleanup."
            )
            if failure:
                st.session_state[P + "error"] = f"Cleanup stopped after {completed} confirmed file(s). {failure}"
                st.error(st.session_state[P + "error"])
                st.warning("Some replacements may have been written before the stop. No completed work was rolled back. Rescan the archive.")
            else:
                st.rerun()
