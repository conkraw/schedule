# Simpler teaching hours — September 30, 2026

This update changes the **Preceptor Teaching Summary** calculations and presentation. It replaces student-weighted educational hours with the preceptor's scheduled time involving students. It applies to the chair report, individual reports, teaching CSVs, monthly detail, pie-chart labels/data, and on-screen summaries.

## The three measures

| Measure | Definition |
|---|---|
| **Total scheduled availability** | All recorded AM/PM shifts for the included preceptor within the selected dates, with or without a student, multiplied by four hours. Weekends are included. |
| **Educational hours** | Recorded AM/PM shifts with at least one student, multiplied by four hours. Two or more simultaneous students do **not** multiply hours. |
| **Learner Reach** | Educational hours divided by total scheduled availability, expressed as a percentage. |

A distinct shift means one preceptor + one calendar date + AM or PM. Duplicate listings are counted once. A preceptor teaching during both AM and PM on one date has eight educational hours. Two students present for that same whole day still produce eight hours, not sixteen.

**Example:** the same two students with a Ward A attending for AM and PM Monday–Friday produce 40 hours of total scheduled availability, 40 educational hours, 100% Learner Reach, two unique students, and two students assigned on three or more distinct dates.

These are four-hour scheduled equivalents from the saved OPDs, not a claim of actual attendance, payroll hours, all institutional clinical work, or teaching quality. Shifts absent from the archive cannot be counted. Future scheduled shifts within the selected dates remain included.

## What the reports look like

The individual report starts with a three-row **Teaching time** table. Unique students and students assigned on 3+ distinct dates appear separately under **Student continuity**. Clinical-experience and monthly tables use the same three time measures. The previous **student-weighted educational effort** table and student-shift column are removed.

The chair report uses the same availability, educational-hours, and Learner Reach columns in its overview and preceptor tables. Its student-continuity table and clinical-experience pies remain. Pie percentages are unchanged; their teaching portion is now labeled **Educational hours**.

Linked OASIS feedback remains in individual Word reports, using full question wording. This update does not change evaluation averages, comments, counts, username matching, or encrypted OASIS summaries.

## Rules that remain in place

- Both selected dates are included; all seven days of the week are read.
- Preceptors and clinical experiences without student assignments during the selected period remain omitted from the displayed reports.
- An included preceptor's overall availability still includes all their recorded shifts, including shifts in a hidden clinical experience with no students. A note explains when displayed experience subtotals are lower than overall availability. Such hidden experiences add no educational hours.
- For a displayed clinical experience, months without students still count in availability and appear in monthly detail.
- Academic Pediatrics combines HOPE_DRIVE, ETOWN, and NYES. Other clinical experiences stay separate.
- Academic Pediatrics takes priority over PSHCH Nursery for the same preceptor/date/AM-PM. The overlapping nursery listing and its student assignments are excluded. Nursery students are not transferred into outpatient counts; an unassigned outpatient listing remains unassigned. Nonoverlapping nursery shifts remain.
- Other conflicting simultaneous clinical experiences still block reports and show source/cell diagnostics. Invalid denominators are not replaced with invented hours or percentages.
- Unique students and the 3+ distinct-date measure are unchanged. AM and PM on the same day remain one day for continuity.

## Install the small update (recommended)

1. Extract **Schedule_App_Simple_Hours_Update.zip**.
2. Merge its `schedule_app` folder into the existing app repository. Replace the matching files and add `educational_time.py`. **Do not delete the existing folder.**
3. Restart the app. Open **Preceptor Teaching Summary**, load or refresh archived OPDs if needed, select/load your reporting period, and click **Create teaching reports ZIP**.

**Keep the launcher (`app_sch_2026.py`), `schedule_app/settings.py`, preceptor email/name mappings, requirements, Streamlit Secrets, encryption key, GitHub username links, and date presets unchanged.** No new dependencies or credentials are needed. No OPDs need to be re-uploaded, modified, or re-encrypted.

A complete recent scan can be reused because it already has distinct clinical-shift data. Old report downloads are invalidated. A new session still needs **Load / refresh archived OPDs**. Regenerate existing Word/CSV reports; previously downloaded files do not change by themselves.

The build marker near the report button is:

```text
Report builder: 2026-09-30-simple-educational-hours-1
```

### Files in the small update

| Action | File |
|---|---|
| Replace | `schedule_app/reports/chair_summary.py` |
| Replace | `schedule_app/reports/individual_teaching.py` |
| Replace | `schedule_app/reports/learner_reach_charts.py` |
| Replace | `schedule_app/reports/teaching_export.py` |
| Replace | `schedule_app/sections/preceptor_teaching_summary.py` |
| Add | `schedule_app/services/educational_time.py` |
| Replace | `schedule_app/services/learner_reach.py` |
| Replace | `schedule_app/services/report_diagnostics.py` |
| Replace | `schedule_app/services/teaching_analysis.py` |

The full modular ZIP is an alternative for a clean installation. Prefer the small update to preserve changes made only in your deployed settings or other modules.

## Teaching CSV column changes — intentional

The existing teaching CSV filenames remain. The columns are simpler and the meaning of `educational_hours` has changed. **Do not compare new educational hours with old student-weighted values as though they were the same measure.**

`preceptor_teaching_summary.csv`:

```text
preceptor_name, academic_year, total_scheduled_availability_hours, educational_hours,
learner_reach_pct, unique_students, unique_students_3plus_days,
months_with_students, months_scheduled, scheduled_shifts, teaching_shifts
```

`preceptor_teaching_by_work_type.csv`:

```text
preceptor_name, academic_year, work_type, total_scheduled_availability_hours,
educational_hours, learner_reach_pct, months_with_students, months_scheduled,
scheduled_shifts, teaching_shifts, source_sites
```

`preceptor_learner_reach_monthly.csv`:

```text
preceptor_name, academic_year, work_type, month, total_scheduled_availability_hours,
educational_hours, learner_reach_pct, scheduled_shifts, teaching_shifts, source_sites
```

The pie-chart data CSV uses `total_scheduled_availability_hours`, `educational_hours`, `hours_without_students`, and `learner_reach_pct` along with its existing period/experience/site fields.

- `scheduled_shifts` is a distinct preceptor/date/AM-PM count, not a student assignment count.
- `teaching_shifts` counts only those distinct shifts with at least one student.
- `educational_hours = teaching_shifts × 4`.
- `total_scheduled_availability_hours = scheduled_shifts × 4`.
- `learner_reach_pct` is a number on a 0–100 scale.
- Overall unique students are not copied into work-type rows and should not be added across preceptors.
- `months_with_students` replaces `months_worked` in these teaching CSVs. `months_scheduled` can include additional months without students.

Old redundant teaching-report columns such as `no_of_shifts`, `recorded_clinical_hours`, and `hours_with_students` are not in the simplified teaching CSVs. Update any external process that depended on those headers.

**The separate Power Automate preceptor-assignment Excel workbook and OASIS output CSV schemas are unchanged.** They are different workflows. Audit/source CSVs may still show student-assignment counts for tracing data; these are not educational hours and are not multiplied into teaching credit.

## Editing / technical details

- `services/educational_time.py` centralizes the published definitions, validation, display names, and CSV fields.
- Public annual/work-type teaching rows now obtain educational hours from validated distinct clinical counts. They no longer fall back to student-weighted totals when clinical data is missing.
- Raw scan assignment fields remain internally for backward-compatible duplicate and student-continuity validation. They are not displayed or exported as preceptor hours.
- The report-output version changed to clear stale downloads without changing the encrypted OPDs or forcing an archive migration.

## Validation

The local automated suite finished with **579 passed and one skipped**, plus 141 passing parameterized subtests. It includes **27 new tests** for this metric change, with simultaneous students, weekends, AM/PM, repeated rows, exact-date boundaries, hidden services, outpatient/nursery priority, other conflict blocking, CSV/chart agreement, source immutability, and report/username linkage compatibility.

Separate comparisons of the two supplied OPD samples verified 82 annual/work-type rows: scheduled availability, existing Learner Reach, and student-continuity counts stayed the same; new educational hours match the old distinct **hours with students**, not the old weighted hours.

The chair and linked individual Word previews and a twelve-month pagination fixture were rendered and visually checked. Preview documents use invented data. GitHub and Streamlit interactions were simulated. No live repository, credentials, or deployment were accessed or changed.
