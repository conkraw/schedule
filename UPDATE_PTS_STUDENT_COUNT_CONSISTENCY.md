# PTS correction: consistent student identities and student-count dates

This update corrects the continuity-versus-assessment discrepancy. It supersedes the previous statement that student continuity always uses the full teaching-report period when assessment completion is included.

## Install the small update (recommended)

1. Extract **Schedule_App_Student_Count_Consistency_Update.zip**.
2. Merge its **schedule_app/** folder into the existing repository. Replace all 11 matching Python files and add the new **services/student_cohort.py** helper. **Do not delete the existing folder or upload only the ZIP as an application file.**
3. Restart Streamlit; the entrypoint remains **app_sch_2026.py**.
4. Unlock **PTS**, choose your existing reporting dates/date preset, and click **Load / refresh archived OPDs**.
5. Check the **Assessments as of** date and your saved **Minimum shifts for assessment completion**, then click **Load / refresh evaluation completeness**.
6. Click **Create teaching reports ZIP** and use the newly generated chair/individual documents. Previously downloaded files do not update retroactively.

The first two source refreshes are required: old OPD student groups and old assessment bundles are invalidated, rather than reused with changed identity rules. Later cutoff/threshold changes can recalculate from already loaded inputs as before.

**Keep the launcher, settings.py, requirements, Streamlit Secrets, encryption key, OER/PTS password, saved date presets, minimum shifts and existing student/preceptor mappings unchanged.** No OPD or OASIS source re-upload, re-encryption, or catalog migration is needed. The full modular-app ZIP is an alternative for a clean installation; the small ZIP is preferable for preserving deployed custom settings.

## What was wrong

Continuity was grouped by the original normalized OPD spelling and used the full report dates. Assessment eligibility used reconciled student identities and stopped at the assessment cutoff. A designation-only duplicate could therefore count twice in continuity and once in assessment eligibility, independently of the date issue.

A local reproduction with invented records, using identical dates in both sections, produced:

| Same learner listed as Alpha and Alpha (MD) on three dates | Before | After |
|---|---:|---:|
| Unique students | 2 | 1 |
| Students assigned on 3+ days | 2 | 1 |
| Eligible students at a three-shift minimum | 1 | 1 |
| Educational hours | 12 | 12 |
| Learner Reach | 100% | 100% |

These are test records, not a recalculation of any named preceptor in the user's live archive.

## One shared student cohort

When assessment inputs have been loaded, all three student measures are computed from the same per-preceptor, per-student set of distinct **date + AM/PM** assignments:

- **Unique students assigned:** all distinct students with a retained shift in the student-count window.
- **Students assigned on 3+ days:** those students with assignments on at least three different dates; AM and PM together count as one date.
- **Students assigned for N+ shifts:** those students with at least the saved minimum number of distinct date/AM-PM slots. This is the assessment denominator.

The same resolver applies confirmed saved name matches, unambiguous OASIS identities and designation normalization. A confirmed correction also carries across a designation-only sibling when no source or saved identifier contradicts it. Exact saved entries retain priority. Conflicting identities and genuine spelling discrepancies are not guessed; provisional counts remain identified for review.

Names absent from OASIS remain normalized OPD-derived members. They require no fictitious external ID and are not removed to improve the percentage. Only confirmed assessments enter the numerator. No new student catalog or persistent identity is created by this correction.

## Two clearly identified reporting windows

**Teaching time** continues to use the full selected reporting period: total scheduled availability, educational hours, Learner Reach, monthly time detail and clinical-experience charts remain unchanged. Weekends are included, and simultaneous students do not multiply hours.

**Student continuity and assessment eligibility** both use the report start through the earlier of report end and **Assessments as of**, inclusive. The documents display these exact student-count dates. Assessment submissions use the same cutoff, as before. The individual report also states the minimum-shift denominator beside the continuity explanation.

For example, the full teaching schedule can include five students, but only three have accumulated qualifying assignments by October 1. The continuity and completion sections now describe the same cutoff-based cohort rather than mixing the future schedule with completion-to-date.

The independent linked educator-feedback summary retains its own displayed dates. It is already averaged and is not silently sliced to the assessment cutoff.

## A mathematical consistency check

Three distinct assigned dates necessarily include at least three assigned AM/PM shifts. Consequently, with a minimum of **one, two or three shifts**, the eligible-student count cannot be lower than the three-day count. Both counts must also be no greater than the unique-student count.

The app validates those relationships before report output. At a minimum of **four or more shifts**, a student with only three dates/shifts may legitimately be in the three-day count but not be eligible; that is not rejected. The code does not force the two measures to be equal or change a denominator to match a printed value.

## Before assessment inputs are loaded

When assessment completion is enabled but its inputs have not been verified, student counts display **Not checked**, not an unrelated full-period fallback beside an unknown denominator. Teaching-hour reports remain available. Load / refresh evaluation completeness to obtain reconciled counts; no extra name confirmations are required simply to initialize the check.

When assessment completion is deliberately turned off, teaching-only continuity remains available from designation-normalized OPD names over the full period, explicitly labelled **OPD-name-only**. That report does not claim to have applied unloaded OASIS/saved identity links or contain a comparable assessment denominator. Generic provider labels are likewise not treated as verified individual preceptors.

## CSV and interface consistency

Chair and individual reports, the overall teaching CSV, and optional PTS previews use the same public student-cohort results. Optional previews now render after the completion bundle is attached, not before it.

The original overall teaching CSV columns remain first, with these appended fields:

```
student_counts_start_date
student_counts_end_date
student_counts_status
eligible_students
minimum_shifts
```

The assessment-completion CSV retains its existing fields and appends:

```
unique_students
unique_students_3plus_days
student_counts_status
```

Student counts now follow the explicit student-count dates when completion is included. Time/hour fields keep their full-period meaning. No student names, IDs, or individual assignment groups are added to CSVs or Word reports. The work-type and monthly teaching-time CSV schemas do not change.

## Files in the small update

| Action | Path |
|---|---|
| Add | schedule_app/services/student_cohort.py |
| Replace | schedule_app/services/student_name_matching.py |
| Replace | schedule_app/services/student_continuity.py |
| Replace | schedule_app/services/teaching_analysis.py |
| Replace | schedule_app/services/assessment_completion.py |
| Replace | schedule_app/services/educational_time.py |
| Replace | schedule_app/sections/assessment_completion.py |
| Replace | schedule_app/sections/preceptor_teaching_summary.py |
| Replace | schedule_app/reports/individual_teaching.py |
| Replace | schedule_app/reports/chair_summary.py |
| Replace | schedule_app/reports/assessment_completion.py |
| Replace | schedule_app/reports/teaching_export.py |

## Validation and limitations

See **TESTING_STUDENT_COUNT_CONSISTENCY.md** and the install manifest in the full package. All tests use local/simulated sources. No live repository, credentials or Streamlit deployment was accessed or changed. The supplied Word reports do not expose the underlying learner-level assignments, so they cannot establish revised actual totals for Ruth or Ben. Their totals will be recomputed from the user's archive after the refresh.

Existing password protection, encrypted catalogs, column-minimized OER storage, saved matches, ignore/restore behavior, Academic Pediatrics-over-PSHCH Nursery priority, clinical-conflict checks and four-hour teaching-time rule remain intact.

Technical reference for the unchanged session-local refresh pattern: https://docs.streamlit.io/develop/api-reference/caching-and-state/st.session_state
