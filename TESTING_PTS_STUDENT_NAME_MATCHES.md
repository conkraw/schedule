# PTS student-name correction verification

## Environment and scope

All fixtures use invented student/provider identities and simulated GitHub and Streamlit services. The live application, archive repository, keys and institutional records were not accessed or changed. Streamlit itself is not installed in this runtime, so these are UI-state simulations, not a browser/Streamlit Cloud deployment test.

## New tests

`tests/test_student_name_matching_ui.py`: **41 passed**.

Coverage includes:
- Exact normalized names and existing mappings do not appear in the entry list.
- Only current OPD students appear; OASIS-only names are candidate choices, not report subjects.
- Repeated assignments and multiple preceptors produce one unresolved-name entry.
- Below-threshold names can be corrected without altering completion eligibility.
- Missing IDs, ambiguous OASIS names and leading-zero IDs are handled explicitly.
- The candidate starts unselected; no similarity match or automatic write occurs.
- Confirmation belongs to the proposed identity; changing the identity does not inherit the previous confirmation.
- A successful, verified encrypted save removes the identity flag, clears cached downloads and updates calculations on rerun.
- A save/refresh failure never falsely removes a flag or replaces a catalog with empty data.
- Concurrent edits require a refresh instead of silent overwrite.
- Existing saved matches can be updated or removed separately; removal can restore the flag.
- Verified IDs with no corresponding submitted assessment do not create completion credit.
- Name-link correction does not suppress independent assessment-source metadata issues.
- Original OPDs and OASIS sources remain byte-identical in simulated storage.
- Student names/IDs do not enter completion CSVs or Word tables after linking.
- A locked or expired PTS session cannot display the new controls or write matches.
- Package comparison confirms one changed existing runtime module and two new modules only.

## Regression runs

The 881 collected tests were covered by completed runs in separate batches. The main batch excluded the legacy teaching/OASIS-link file and completed with **829 passed, 3 skipped**. A subsequent batch covering that file plus the 41 new tests completed with **90 passed** (49 existing + 41 new). Unique total across the completed batches: **878 passed, 3 skipped**, with 139 successful subtests in the main batch.

A single-process full-suite attempt timed out near the final legacy document-equivalence checks; those checks passed in the separate final batch. No calculation or Word-rendering code was changed to make tests pass. Skips belong to existing environment/historical packaging checks. A new runtime-file comparison test covers this update's file boundary.

`compileall` succeeded for the launcher and all application modules. Both final ZIPs are integrity-checked, exclude caches and bytecode, and contain no real student data or credentials added by this update.

No new Word report layout was created or changed by this UI update. Report-content regression tests and the privacy checks above were used; no new visual-layout validation is claimed.
