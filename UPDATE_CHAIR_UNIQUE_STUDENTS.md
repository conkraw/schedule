# Chair report: unique students and students assigned on 3+ days

This update adds the existing individual-report student counts to the combined
chair Word report. It is built on `Schedule_App_Modular_Unique_Students.zip`.

## Install the small update

1. Extract `Schedule_App_Chair_Unique_Students_Update.zip`.
2. Merge its `schedule_app` folder into the existing folder in the app repository.
   Replace these three matching files; do not delete the existing folder:

   - `schedule_app/reports/chair_summary.py`
   - `schedule_app/reports/teaching_export.py`
   - `schedule_app/sections/preceptor_teaching_summary.py`

3. Restart the app so the revised modules are imported.
4. Open **Preceptor Teaching Summary**, select or load the reporting-date preset,
   and click **Create teaching reports ZIP**. Download either the complete ZIP or
   **Download chair summary only (Word)**.

Leave `app_sch_2026.py`, `schedule_app/settings.py`, requirements, email/name
mappings, Streamlit Secrets, encryption keys, and GitHub date presets unchanged.
This patch does not replace those files. It requires the preceding unique-students
update (which supplies `services/student_continuity.py` and its scanning changes).
For an older installation, use the full package and preserve your custom settings.

If a current unique-student scan is already loaded in this session, it can be
reused. Old report downloads are cleared automatically on the next run. When
starting a new session, or if the app asks for it, click **Load / refresh archived
OPDs** before generating reports. No original OPDs need to be edited, re-uploaded,
or re-encrypted.

## What the chair sees

Each selected reporting-period section now includes **Student continuity by
preceptor**, an alphabetical table with:

| Preceptor | Unique students | Students assigned on 3+ days |
|---|---:|---:|
| Example, Avery | 2 | 1 |

The row above is invented example data.

These are each preceptor's **overall counts across all work types** for the
selected dates, exactly matching their individual report. They are shown in a
separate table, not mixed with service-specific Learner Reach percentages.
The existing work-type effort tables, clinical-experience pies, and individual
reports remain included.

## Counting rules (unchanged)

- Monday PM, Tuesday AM, and Thursday PM for the same student = one unique
  student who qualifies for 3+ days.
- Monday AM, Monday PM, and Tuesday AM = one unique student, but only two days;
  that student does not qualify for 3+ days.
- Days need not be consecutive. Weekends count.
- Only dates inside the displayed reporting period count. A custom period is
  not split at July 1; standard academic-year sections each use their own bounds.
- A student assigned to the same preceptor in multiple settings, months, or
  rotations is still one unique student for that preceptor during that period.
- Academic Pediatrics takes priority over overlapping PSHCH Nursery assignments
  before student counts are calculated. Excluded nursery learners/dates are not
  transferred into clinic teaching credit.
- Unassigned-only preceptors and services remain omitted. Provider/site/slot
  labels needing attribution are kept in a separately labeled table.
- The same student can be assigned to several preceptors. **Do not sum these
  columns to infer a clerkship-wide number of unique students.** No misleading
  summed unique-student total is displayed.

The existing OPD name-matching rules still apply: capitalization, spacing, and
comma spacing are normalized; different spellings are not guessed to be the same
student. Two people sharing the exact same recorded name cannot be distinguished.
Student names, student identifiers, and anonymous date groups are not exported.
These counts describe scheduled assignments, not verified attendance.

## Other outputs and storage

No CSV column names or counting rules were changed. Shift counts, educational
hours, Learner Reach, pie-chart values, student schedules, primary-preceptor
reports, encrypted OPD storage, and saved date presets are unchanged. The ZIP's
`Report_Notes.txt` now explains that both chair and individual reports include the
two student-continuity measures. Nothing is uploaded to GitHub when reports are
created.

## Local checks

The updated package passed Python compilation and 272 local automated tests:
249 existing regression tests and 23 new chair-continuity tests. These cover
report-to-report agreement, distinct dates, AM/PM on the same date, simultaneous
students, weekends, custom and standard periods, multiple work types, outpatient
priority, conflicts, omitted unassigned providers, privacy, source immutability,
ZIP contents, and stale-download invalidation. GitHub and Streamlit interactions
were simulated; the live app/repository were not accessed or modified.

The supplied illustrative chair preview was generated from invented OPDs and
rendered for visual inspection. It is not a report from the live GitHub archive.
