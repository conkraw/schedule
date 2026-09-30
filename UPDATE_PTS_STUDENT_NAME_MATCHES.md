# PTS: student-name alerts and a missing-only correction dropdown

This update replaces the older manual-ID-only correction panel with a name-based selector. It reuses the existing encrypted student-ID catalog. It does not change the completion formula, hours, OER/PTS password boundary, archive encryption, or source CSV columns.

## Install the small update (recommended)

1. Extract `Schedule_App_PTS_Student_Name_Matches_Update.zip`.
2. Merge its `schedule_app` folder into your existing app repository. **Do not delete the existing folder.**
3. Replace this file:

   `schedule_app/sections/assessment_completion.py`

4. Add these two files:

   `schedule_app/sections/student_name_matches.py`

   `schedule_app/services/student_name_review.py`

5. Restart Streamlit and open **PTS** using your existing OER/PTS password.

Keep the launcher, settings.py, requirements.txt, all custom mappings, Streamlit Secrets, existing encryption key, saved date presets and source archives unchanged. No new dependencies or credentials are needed. The complete modular ZIP is an alternative for a clean install; use the small update to preserve deployed customizations.

## Normal workflow

1. Load your OPDs and select the reporting dates in PTS as usual.
2. Leave **Include assessment completion and missing-evaluation alerts** checked.
3. Click **Load / refresh evaluation completeness**. This loads the OASIS student-assessment metadata and the saved student matches.
4. A visible **Student-name match needed** alert gives the number of unresolved OPD names. Open **Resolve OPD student names to OASIS** (expanded automatically when names need attention).
5. The first dropdown, **OPD students needing an OASIS match**, contains only unresolved names. The table shows the issue, affected preceptors, periods, and whether it affects a 3+ shift completion percentage.
6. Select the OPD name. Under **Choose an OASIS student**, select the correct name in **Matching OASIS student name**. Each option includes the actual Student External ID from the loaded OASIS metadata. No candidate is preselected by similarity.
7. Confirm **I verified that these records refer to the same student** and click **Save student match to GitHub**.
8. Only after the encrypted save is read back and verified, the app rechecks the name list and assessment-completion results. The resolved name disappears, the next unresolved name is offered, and the selection/confirmation fields clear. Generate the teaching report ZIP again to use the new match.

The correction is **OPD student name → OASIS Student External ID**, not a username generated from the student's name. The selected OASIS name supplies the ID; original OPDs and OASIS source files are not edited.

## What is flagged and what is not

- The queue covers student names assigned to named teaching preceptors within the selected dates, after the existing OPD overlap/priority rules.
- Exact normalized name matches with one external ID are already resolved. Existing normalization ignores case, repeated spaces, comma spacing and the recognized `; MD2028`-style suffix. This has not changed.
- A name with no unique ID match or a name associated with multiple IDs is flagged. Different IDs for the same OASIS name remain separate choices; verify the actual student before selecting one.
- A confirmed saved match is reused, even when there is no assessment for that ID in this period. Missing identity and missing assessment are separate issues.
- Names below the 3-shift threshold may also be corrected proactively. Their table rows state that they do not currently affect the completion denominator. This does not change eligibility.
- Unassigned students, OASIS-only students and assignments to generic/nonindividual provider labels are not added to the OPD correction queue.
- When all current names are resolved, a success message replaces the name alert and the first dropdown is hidden.

Unresolved students are **not discarded** to improve a percentage. Applicable completion percentages remain Not verified until their identity requirements are met. Missing-record alerts remain nonblocking for teaching-report generation.

## Missing OASIS names or a verified external ID from another source

OASIS choices come from all currently loaded student-assessment exports; they may span more dates than the selected reporting period. The recognized MD-class suffix may be omitted from displayed names. Confirm the identity, not just a similar spelling.

If the student is absent from this list, upload the relevant student-assessment CSV in **OER → Evaluations of students**, then refresh evaluation completeness in PTS. Do not select a different student merely to remove a flag.

The alternative **Enter a verified Student External ID** remains available when you can verify the actual ID elsewhere. This can resolve the OPD identity match without claiming an evaluation exists. It cannot repair a missing ID or contradictory identity in an OASIS form itself; those source issues must be corrected separately.

Two genuinely different students with the same OPD name cannot safely share a global name link. Give the OPD records an unambiguous name/identifier before linking them. Explicit spelling aliases for the same person can share one verified external ID, using the existing catalog behavior.

## Review, update or remove saved matches

Open **Review or correct saved student matches (optional)**. It lists saved OPD names, their IDs, OASIS names associated with those IDs in the currently loaded metadata, and whether the OPD name is in the selected dates.

Enable **Edit or remove a saved student match**, select it, then choose the correct OASIS student (or enter a verified ID), confirm and update. Confirm removal to use **Remove saved student match**. A removed link returns to the first dropdown only if its OPD name cannot otherwise be matched exactly.

**Refresh saved student matches** retrieves changes saved by another session without downloading all OPDs again. Use **Load / refresh evaluation completeness** after new source uploads or changes to preceptor/summary links. Competing edits are rejected with a refresh/review instruction; a failed save does not dismiss a flag or erase the current catalog.

## Storage and privacy

The catalog path is unchanged:

```text
opd_archive/student_assessment_id_links.json.enc
```

The actual base folder follows Streamlit Secrets. The existing catalog stores the OPD name, confirmed external ID and update time. OASIS candidate names are taken from the loaded, already-minimized assessment metadata; no extra student fields, answers, grades, usernames, or plaintext catalog are uploaded by this feature.

Existing mappings remain compatible; there is no schema migration or new encryption key. The same verified, revision-checked encrypted save/load service is reused. The new controls are inside protected **PTS** and their session data is cleared by **Lock OER / PTS**.

Student names and IDs appear in the protected matching controls, not in the added completion tables or generated report ZIP. Existing free-text educator feedback can contain identifying content independently of this feature; keep treating downloads as evaluation records.

Changing a match clears cached report downloads. It does not change OPDs, OASIS originals, preceptor usernames, evaluation scores, date settings, or other student matches. Generate reports again after corrections. Removing a current catalog entry is not repository-history erasure.

## Verification

The package includes 41 new local tests for unresolved-only queues, explicit OASIS selections, ambiguity, encrypted persistence, stale edits, refresh/removal, clearing flags, keeping denominators, excluding student identifiers from reports, and password access. Runtime changes are limited to the three files listed above. See `TESTING_PTS_STUDENT_NAME_MATCHES.md` in the full ZIP.

Technical references used for the UI and unchanged persistence contract:
- Streamlit widget/session state: https://docs.streamlit.io/develop/api-reference/caching-and-state/st.session_state
- Streamlit rerun behavior: https://docs.streamlit.io/develop/api-reference/execution-flow/st.rerun
- GitHub Contents API: https://docs.github.com/en/rest/repos/contents
