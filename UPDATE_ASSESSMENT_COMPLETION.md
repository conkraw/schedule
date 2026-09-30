# Student assessment completion and evaluation-record alerts

## Install (small update recommended)

Extract `Schedule_App_Assessment_Completion_Update.zip`. Merge its `schedule_app/` folder into the existing application repository, replacing the seven matching Python files and adding four new modules. **Do not delete the existing folder.** Restart Streamlit.

Keep `app_sch_2026.py`, `schedule_app/settings.py`, email/name mappings, requirements, Streamlit Secrets, the encryption key and all saved date presets unchanged. No new dependency, credential, repository, app password or source-file migration is required.

Updated files:

- `schedule_app/services/teaching_analysis.py`
- `schedule_app/services/teaching_evaluations.py`
- `schedule_app/sections/preceptor_teaching_summary.py`
- `schedule_app/sections/preceptor_oasis_links.py`
- `schedule_app/reports/individual_teaching.py`
- `schedule_app/reports/chair_summary.py`
- `schedule_app/reports/teaching_export.py`

New files:

- `schedule_app/services/assessment_completion.py`
- `schedule_app/services/student_assessment_links.py`
- `schedule_app/sections/assessment_completion.py`
- `schedule_app/reports/assessment_completion.py`

## Normal workflow

1. Use **OASIS Evaluations → Evaluations of students** to archive original preceptor-to-student assessment CSVs, as before. No new upload is needed for already archived originals.
2. In **Preceptor Teaching Summary**, load OPDs and choose the reporting dates or a saved date preset. Use the existing saved preceptor username links. For the student-to-educator feedback check, save an educator-feedback summary link with the same exact date boundaries in the existing linking area.
3. Leave **Include assessment completion and missing-evaluation alerts** selected. Click **Load / refresh evaluation completeness**. This reads/decrypts archived student assessments, re-reads OPDs at the snapshot already displayed, and checks saved educator-feedback links. It does not change originals or save a generated report in GitHub.
4. If more than one course is represented in the target assessment forms, explicitly select the matching course(s). There is no automatic all-courses combination. With one course it is selected automatically.
5. Review the warning tables and, when needed, use **Resolve student external-ID matches (optional)**.
6. Click **Create teaching reports ZIP** as usual. Missing evaluation records do not prevent a teaching report. Unknown or unverified completion results are labeled rather than invented.

The check is a snapshot, not background monitoring. Refresh it after new OASIS uploads or username/summary-link changes. If linked-feedback controls were already open, refresh those links first and then this check. Dates, OPD snapshot or archive configuration changes invalidate the old assessment check. Turning the new checkbox off produces the existing teaching report without these new tables.

Before a first check, reports explicitly say **Not checked** for the new measures; this is not a zero. No further OPD refresh is mandatory for an already loaded current teaching scan; the completion check replays its exact source snapshot to obtain the learner-to-shift link that old aggregate-only scans cannot supply.

## Exactly what the percentage means

The label is **Student assessment completion**. The denominator is **unique students assigned to the same preceptor for at least three distinct AM/PM shifts within the selected reporting dates**. Eligible students are counted once across rotations, months and work types.

The numerator for a form is **how many of those same eligible students have at least one submitted form of that type from that preceptor, using Submit Date within the same reporting period**. The same student cannot contribute more than one to a form-specific numerator. A student evaluated but assigned fewer than three shifts is not in this percentage's numerator or denominator.

Separate columns/rows are shown for:

- `*Clinical Assessment of Student`
- `*PEDS History Taking & Physical Exam`
- At least one of the two forms (a union, not the sum of the first two counts)

Example: 10 eligible students; eight have a Clinical Assessment, four have a History & Physical assessment. The individual percentages are 8/10 = 80% and 4/10 = 40%. The combined rate depends on how those two sets overlap, not 12/10. No number is capped to hide double counting; membership sets prevent it.

**Three shifts is not three days.** Monday AM, Monday PM and Tuesday AM qualify for this denominator. They do not qualify for the existing **Students assigned on 3+ days** measure, which remains unchanged. Weekends count. Nonconsecutive shifts count. Repeated copies of the same student/date/AM-or-PM assignment count once. Two different students during one shift each gain one assignment; their shared shift still contributes only four educational hours to the preceptor's time report.

The existing Academic Pediatrics-over-PSHCH Nursery exception is applied **before** eligibility is calculated. Excluded nursery students are not transferred to clinic and do not add to the denominator. Other OPD work-type conflicts retain the existing blocking safety check. This update only makes **missing evaluation records** nonblocking; it does not disable schedule-conflict validation.

No eligible students is shown as **No eligible students**, not 0% or 100%. The measure is record completeness, not verified attendance, teaching quality or a policy stating both forms must be completed for every student.

## Interpreting the source CSV

Relevant columns are **Evaluator Email, Evaluation, Student External ID, Form Record, Course ID, Submit Date**. The educator is matched to an OPD preceptor using the local part before `@` in Evaluator Email and that preceptor's explicitly saved username link.

A completed form here means a submitted form record with valid identity metadata. The question values, grades, narrative answers and question IDs are not used to determine the completion percentage. This does not audit whether every question was answered.

Question-level rows are collapsed using **Course ID + requested form type + Form Record**. Distinct forms can be audited separately, but repeated submitted forms for the same evaluator and Student External ID still count once per student in the numerator. Repeated archive snapshots do not multiply the counts. Blank Submit Date means not submitted; a nonblank invalid date is a source issue. Both report date boundaries include the entire calendar day, with no time-zone conversion of OASIS timestamps.

`*PEDS Handoff` is not counted in the requested completion measures. Its identifying fields can help link an OPD student name to Student External ID. It remains stored in the original student-assessment archive, as before.

A consistent newer nonblank identity field may fill a blank field from another copy of the same form. Contradictory nonblank metadata is not resolved by choosing an arbitrary latest/first snapshot. It is reported for review. Adding another contradictory copy does not remove the earlier source problem.

## Matching students safely

OPDs contain student names, not Student External ID. The app first matches an OPD student name to the OASIS Student field and obtains that student's **Student External ID**. Matching ignores capitalization, repeated whitespace and comma spacing. It removes only the explicit trailing MD-class suffix seen in the supplied export, such as `; MD2028`. It does not fuzzy-match names or silently substitute Student Username for an external ID.

A name associated with multiple external IDs or no external ID is flagged. The affected preceptor's completion percentage is **Not verified** until the eligible students can be matched, rather than shrinking the denominator to just the matched students. Other preceptors' verified calculations can still be displayed and teaching reports remain available.

In **Resolve student external-ID matches**, select the unresolved OPD name, enter its verified external ID, confirm identity and save. The correction is encrypted and verified in the existing repository:

```
opd_archive/student_assessment_id_links.json.enc
```

The actual base folder follows Streamlit Secrets. Existing preceptor username maps and source CSVs are untouched. Saved student ID links can be reviewed, corrected or removed with confirmation. Explicit aliases can share an external ID. Do not use a single global name mapping for genuinely different students with the same name; the OPD itself needs an unambiguous name/identifier before those students can be distinguished reliably.

Student names and IDs appear only in the app's matching review and temporary session data. They are not included in the new report tables, chair document or report ZIP. Manual links are the only new persistent data, encrypted with the existing key. The API uses revision checks and verifies encrypted writes; a failed/competing save is not silently labeled successful.

## The two nonblocking alert tables

**Student → educator:** the linked exact-date educator-feedback summary does not include evaluations for the preceptor's saved username. If no summary or username is available, this says **Not checked**, not that the educator received zero evaluations. Extra OASIS educators absent from the teaching roster do not create new reports.

**Preceptor → student:** no submitted target forms were found for the preceptor within the selected Submit Dates. The table also identifies when only one of the two requested form types is present. Student assessments are not treated as educator-feedback scores and vice versa.

Each table shows the preceptor, username, period, issue and next action. Missing source fields are detailed separately with archive filename, form record, form type, source evaluator names/username hints and date where available—never student assessment answers. Exported Evaluator Username is a **review hint**, not a substitute for the email-derived username.

An unidentified/other evaluator's malformed form is not arbitrarily attributed to every preceptor. It appears as an unattributed-source alert; other rates describe **identifiable matched records** and may change after corrections. A source issue attributable to a particular preceptor marks that preceptor's rates unverified. Missing/invalid student identities never become guessed matches or fabricated zeroes.

These alerts continue to be shown in reports. They identify missing records, not proof that an evaluation never happened. A failed download or decryption produces a not-checked status, never an apparently complete partial count. Existing linked-feedback protections still prevent using a malformed or mismatched saved summary as actual learner feedback; missing exact-date summaries alone now permit teaching-only output.

## Report outputs

The chair document adds one **Student assessment completion by preceptor** table per reporting period, across all work types. Each cell shows the numerator/denominator and percentage. It also lists evaluation-record warnings. Individual documents add the same counts for that preceptor, alongside the unchanged teaching-time and continuity sections.

Additional files inside the existing teaching report ZIP:

- `preceptor_student_assessment_completion.csv` — one row per named teaching preceptor/report period.
- `Evaluation_Completeness_Alerts.csv` — provider/source-level warning rows, when present.

Existing teaching CSVs, OASIS output CSV, educational-hour calculations, Learner Reach percentages, charts, date presets and detailed educator-feedback sections are preserved. The optional completion CSV includes total submitted forms for each type as an audit field; these totals can exceed unique students and are **not** the completion numerator. Blank percentages mean not checked/unverified or no eligible denominator; see `assessment_status`.

## Access and confidentiality

No new password or access restriction is added. Keep the existing deployment's authorized-use arrangements in place. Original assessment CSVs remain encrypted in GitHub. Downloaded Word reports/CSVs are unencrypted; handle them as staff educational records. The new calculation does not export student scores, assessment comments or identifiers. Existing linked educator-feedback comments remain governed by the earlier feature's privacy notice.

## Validation

The release test run completed with **727 passed and 2 skipped**, including **54 new tests** for question-row collapse, repeated forms/snapshots, exact date limits, three shifts vs three days, eligibility/numerator membership, external IDs, encrypted mapping persistence, warnings, OPD priority, reports and simulated UI workflows. One existing test expectation was deliberately updated because a missing exact-date educator-feedback summary is now a nonblocking alert rather than preventing teaching reports.

The supplied student-assessment CSV was parsed separately: valid form metadata was counted without using answer content; missing/invalid evaluator emails produced review issues. Two synthetic individual reports and a synthetic chair report were generated; Word page layouts were rendered and inspected. No live repository, Streamlit Cloud account or institutional dataset beyond the provided files was accessed.

Technical references used for integration:
- GitHub Contents API (read-at-commit and revision-aware replacement): https://docs.github.com/en/rest/repos/contents
- Streamlit Session State and callback/widget behaviour: https://docs.streamlit.io/develop/api-reference/caching-and-state/st.session_state
