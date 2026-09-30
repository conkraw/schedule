# Link preceptor teaching reports to saved OASIS evaluations

This update adds an optional username-linking function **inside Preceptor Teaching Summary**. For matched preceptors, it adds an OASIS evaluation section to the **existing individual Word report**. It does not add OASIS results to the chair summary, change clinical-hour calculations, or create reports for OASIS-only educators.

## Install the small update

Extract `Schedule_App_Teaching_OASIS_Links_Update.zip` and merge the `schedule_app` folder into your existing repository. **Do not delete the existing folder.**

Replace these four files:

```text
schedule_app/sections/preceptor_teaching_summary.py
schedule_app/reports/individual_teaching.py
schedule_app/reports/teaching_export.py
schedule_app/services/oasis_workflow.py
```

Add these four files:

```text
schedule_app/sections/preceptor_oasis_links.py
schedule_app/services/preceptor_oasis_links.py
schedule_app/services/teaching_evaluations.py
schedule_app/reports/preceptor_evaluations.py
```

Restart the app. Keep **app_sch_2026.py, settings.py, email/name mappings, requirements.txt, Streamlit Secrets, and your encryption key unchanged**. No new dependencies, credentials, repository, app password, or re-encryption is needed. Existing date presets, OASIS username corrections, OPDs, and archived OASIS exports remain intact.

The full modular-app ZIP is supplied for a clean installation; the small update is safer when preserving deployed customizations. The test files and this guide do not need to be uploaded to run the app.

## Normal workflow

1. In **OASIS Evaluations**, apply the intended reporting period. Let the app create/update and verify its encrypted summary CSV. No download is needed. Existing summaries of known questions can be used without rebuilding them.
2. In **Preceptor Teaching Summary**, load the archived OPDs and choose the same exact start and end dates (or use your saved date preset).
3. Near the report-generation button, expand **Link preceptors to OASIS evaluations (saved in GitHub)** and check **Include linked OASIS evaluations in individual Word reports**.
4. Choose the saved OASIS summary from the dropdown. Click **Save selected summary link to GitHub**. Each reporting period remembers its chosen summary. Different course/form scopes can have different summaries for the same dates; select the intended one explicitly.
5. Select a **Teaching preceptor**, enter the OASIS `record_id` (username before `@`), confirm that the username belongs to that preceptor, and click **Save username link to GitHub**. Repeat for the people you want linked. The reference expander shows the usernames and names in the selected OASIS summary; no matching is guessed.
6. Review the match-status table, then click the existing **Create teaching reports ZIP** button. Matched individual Word reports include the evaluation section. The usual ZIP and chair-only download remain available.

You only need to assign a preceptor's username once. Saved links are loaded in later sessions when this option is enabled. Different labels for the same person in OPD and OASIS are acceptable because the join uses your explicit username link, not similar names.

## What appears in a matched individual Word report

The existing hours, Learner Reach, work-type tables, unique-student counts, and three-distinct-day counts are preserved. A new **Learner feedback on teaching** section is appended for the applicable reporting period:

- Submitted evaluation count and linked username.
- Exact/full question wording, mean Multiple Choice Value, and scored response count for each question.
- A separate duration-category question section, clearly labeled as a mean of category codes rather than weeks or a quality rating.
- Combined **Please indicate this educator's strengths** comments.
- Combined **Areas for Improvement** comments.
- Exact OASIS Submit Date boundaries and source-summary identifier.

The document shows **Built on my knowledge and skill base.**, not `q588_mean`. Technical question-column names are not used as Word headings or question labels. Comments are not rewritten or summarized. A question without scored responses is shown as **Not scored**, with zero responses.

## Which people are included

The roster comes from teaching preceptors with student assignments in the selected period, not the OASIS educator list.

- Explicit username link **and** matching OASIS summary row: add that educator's feedback to the preceptor's individual teaching report.
- Teaching preceptor with no saved username or no matching OASIS row: keep the normal teaching-only report; the preview explains the missing match. A missing match is not reported as zero evaluations.
- OASIS educator absent from the teaching summary: ignore for this output. No additional Word document is created. The app does not infer anyone's faculty/resident status.
- Generic site/slot/combined labels: not linked as individual educators, even if a catalog entry were present for that label.

One username cannot be assigned to two distinct preceptor-name entries. Correct/remove the old link, or merge true spelling aliases in the existing `TEACHING_PRECEPTOR_NAME_MAP`, rather than duplicating someone's feedback. Names are normalized only for case, whitespace, and comma spacing; no fuzzy identity matching is performed.

## Dates: why the boundaries must match

Teaching effort is filtered by the **actual assignment date**. OASIS feedback is filtered by **Submit Date**. A saved OASIS summary is already aggregated, so it cannot be sliced into another date range or split across multiple years after the fact.

The dropdown lists only summaries with the **same exact start and end dates** as the teaching-report period. A shared label such as `26-27` is not sufficient to establish matching dates. Labels may differ when the actual boundaries match; the source OASIS label is identified in the evaluation section.

For multiple standard academic years, choose/save one matching summary per nonempty report year. The app does not attach a combined multi-year average to each separate year. No evaluation section is carried across unrelated periods.

Even with identical boundaries, evaluations are **not** claimed to be responses from precisely the students counted in the OPD's unique-student measure. The document explains that these are different datasets and date bases.

## GitHub storage

The app uses your existing token/key and configured archive folder. It creates one additional encrypted catalog:

```text
opd_archive/preceptor_oasis_links.json.enc
```

The actual folder follows your Streamlit Secrets setting. The catalog stores:

- Preceptor name -> explicitly assigned OASIS `record_id`.
- Reporting-period boundaries -> selected saved OASIS summary filename.

It is separate from `oasis_educator_usernames.json.enc`, which fixes usernames when generating the OASIS CSV. Editing a teaching link does not change the OASIS educator identity, the original CSV, or either source archive.

Every save is re-read, decrypted, and verified. A stale edit from another session causes a refresh/retry prompt; it is not silently overwritten. A damaged catalog or wrong key is not replaced with an empty catalog.

To remove a username link, select the preceptor, confirm removal, and click **Remove saved username link**. Saved preceptors remain selectable for management even when they are outside the current reporting period. To change a saved summary selection for the same dates, choose another matching file, confirm replacement, and save. These actions change current link settings only; no OPD, OASIS summary, or original export is deleted.

## Full question wording in new OASIS output CSVs

The combined OASIS workflow still produces **one encrypted output CSV**, not a separate question-key file. New/rebuilt outputs append `q<ID>_question` columns containing the exact source question wording, while preserving the existing identity fields, counts, means, and comments. The four period columns remain at the end.

Older summaries without those text columns remain compatible for the known questions explicitly defined from your supplied OASIS export. Unknown questions without a stored label stop linkage and direct you to rebuild the summary in **OASIS Evaluations** using the updated app; a label is never invented.

The OASIS workflow version changes so cached prepared data refreshes after installing this update. You do not need to re-upload the original evaluations. The original OASIS source contents and calculation rules are unchanged.

## Refresh and failure behavior

**Refresh links and OASIS summaries** reloads saved mappings, selections, and current summary contents. It does not redownload the OPDs. Refresh the OPDs separately when their source archive changed.

Before building linked reports, the app rechecks the mapping and selected summaries at one current GitHub commit. If a selected summary or mapping changed after the preview, report generation stops and asks you to refresh links. That prevents a preview for one result being used to generate a different result silently. A still-unrefreshed output CSV can remain an old OASIS snapshot; regenerate that period in **OASIS Evaluations** first when new evaluation exports exist.

Editing an unsaved username/summary selection, changing the reporting period, changing saved links, or turning linked evaluations on/off clears previous report downloads. The new options do not change the stored OPD scan. Missing links for individual preceptors simply leave those documents teaching-only, with a visible status; invalid CSVs, mismatched dates, or failed decryption are not silently treated as missing evaluations.

Turn off **Include linked OASIS evaluations in individual Word reports** to run the original teaching-only workflow. That option makes no OASIS catalog/report network requests. It does not delete your saved links.

## Privacy

Only ciphertext is written to GitHub by this feature. The individual Word documents and teaching ZIP downloads remain unencrypted and now may contain evaluation comments. Comments can identify learners, patients, or colleagues even when structured student columns are omitted. Handle the documents as evaluation data.

No app password or additional access restriction is introduced. Existing deployment access controls still determine who can view/modify links and download reports.

## Editing locations

| Change | Module |
|---|---|
| Username assignment, summary selection, preview UI | `sections/preceptor_oasis_links.py` |
| Encrypted link storage and collision checks | `services/preceptor_oasis_links.py` |
| CSV question validation, exact-date checks, username join | `services/teaching_evaluations.py` |
| Evaluation-section wording, tables, comments | `reports/preceptor_evaluations.py` |
| Existing individual teaching report plus optional attachment | `reports/individual_teaching.py` |
| Full-wording columns in new OASIS summaries | `services/oasis_workflow.py` |

## Validation

The existing suite plus the new linkage tests ran locally: **553 tests, 552 passed, 1 skipped**. The 49 new tests cover encrypted save/load/update/remove, stale edits, wrong keys, duplicate usernames, matching/missing usernames, OASIS-only exclusions, exact-date validation, legacy/new question text, averages/counts, comments, Word output, ZIP preservation, and simulated UI actions.

A separate local comparison used the OASIS CSV supplied in this conversation. Every original educator-summary column value and the missing-email issues were unchanged; appended question text matched the parsed source. No real username correction was guessed or saved.

The example and a long-comment report were rendered to page images and visually checked. The preview uses invented schedules and feedback. GitHub and Streamlit interactions were simulated: no live repository or deployment was accessed or changed.

Technical references for implementation: GitHub Contents API (https://docs.github.com/en/rest/repos/contents); Streamlit widget/session behavior (https://docs.streamlit.io/develop/api-reference/caching-and-state/st.session_state). Question wording comes from the supplied OASIS export, not outside guidance.
