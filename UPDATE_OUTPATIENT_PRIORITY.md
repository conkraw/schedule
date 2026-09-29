# Update: Academic Pediatrics takes priority over PSHCH Nursery

## Install this update

Use `Schedule_App_Outpatient_Priority_Update.zip` for the working modular app that already has strict Learner Reach validation and pie charts.

1. Extract the ZIP.
2. Merge its `schedule_app/` folder into the existing folder in your app repository. Replace the seven matching Python files and add `schedule_app/services/teaching_priority.py`. **Do not delete the existing folder**; this is a partial update.
3. Restart/reboot the Streamlit app after all eight Python files are uploaded.
4. Open **Preceptor Teaching Summary**, choose the reporting dates/label, and click **Load / refresh archived OPDs**.
5. Generate the report ZIP again. Old scans and downloads are invalidated because their numbers predate this rule.

Keep `app_sch_2026.py`, `schedule_app/settings.py`, email/name mappings, `requirements.txt`, Streamlit Secrets, and the encryption key unchanged. There are no new dependencies. Existing encrypted OPDs do not need to be changed, re-uploaded, or re-encrypted for these legitimate overlaps.

The complete app is also supplied as `Schedule_App_Modular_Outpatient_Priority.zip`. If using the complete package rather than the update, preserve your deployed `schedule_app/settings.py` and any other local customizations.

## Reporting rule

For the **same normalized preceptor, actual date, and AM/PM half-day**, a listing on **HOPE_DRIVE, ETOWN, or NYES** takes priority over **PSHCH_NURSERY**.

- Retain the Academic Pediatrics listing and its actual student assignments.
- Exclude the overlapping PSHCH Nursery listing from recorded clinical shifts/hours, shifts/hours with students, and shifts/hours without students.
- Exclude the nursery-side student assignments from student-shifts and student-weighted educational hours. Nursery learners are **not transferred to clinic**.
- Count the retained clinical half-day once overall, even if multiple rows or multiple clinic sites list the same preceptor.
- Keep all other nursery half-days. Nursery in the AM is not removed because clinic occurs in the PM.
- Apply the exception to AM or PM whenever that exact overlap occurs. It is not a whole-day or whole-rotation exclusion.
- Use existing name normalization and explicit name aliases; do not guess identities from similar names.

### Important blank-clinic case

If a clinic listing has no student, but the simultaneous nursery listing has one, the retained clinic half-day is **a clinical shift without a student**. The nursery learner does not make the clinic shift a teaching shift. This prevents inflated Learner Reach.

### Example

An invented schedule for one preceptor has:

| Half-day | OPD listings | Reporting treatment |
|---|---|---|
| Monday AM | PSHCH Nursery with one student | Nursery: one clinical shift and one student-shift |
| Monday PM | PSHCH Nursery with one student AND clinic with two students | Academic Pediatrics only: one clinical shift and two student-shifts |
| Tuesday AM | PSHCH Nursery without a student | Nursery: one clinical shift without students |
| Tuesday PM | Clinic without a student | Academic Pediatrics: one clinical shift without students |
| Wednesday PM | Clinic with one student | Academic Pediatrics: one clinical shift and one student-shift |

The resulting figures are:

| Category | Recorded hours | Hours with students | Learner Reach | Student-shifts | Educational hours |
|---|---:|---:|---:|---:|---:|
| Academic Pediatrics | 12 | 8 | 66.7% | 3 | 12 |
| PSHCH Nursery | 8 | 4 | 50.0% | 1 | 4 |
| Overall | 20 | 12 | 60.0% | 4 | 16 |

Clinical hours are four-hour equivalents of unique preceptor/date/AM-or-PM half-days. Student-weighted educational hours still count separate students independently. The Monday PM nursery learner earns no additional credit for that preceptor.

## What is updated automatically

The effective records are finalized before any aggregation. The same records feed:

- Chair summary totals, work-type overview, preceptor detail, and pie charts.
- Individual preceptor reports, including monthly detail.
- Overall teaching CSV, work-type CSV, monthly Learner Reach CSV, and clinical-experience chart-data CSV.
- On-screen totals and report previews.

Preceptors and services without retained student assignments in the selected period stay hidden. An included preceptor's other unassigned clinical shifts still count in the denominator. Pie charts continue to use exactly the same clinical-hour totals as their tables.

The rule is applied after reading all current rotations, so overlapping archive files and worksheet reading order do not change which site takes priority. Custom date filtering and standard academic-year selections remain available. Other clinical work types are not prioritized automatically.

## Other conflicts still stop the reports

Examples that remain blocking include Academic Pediatrics versus Ward A, Academic Pediatrics versus Complex Care, or PSHCH Nursery versus Complex Care without an Academic Pediatrics listing.

For a three-way listing such as Academic Pediatrics + PSHCH Nursery + Ward A, nursery is excluded under this explicit rule, but the remaining Academic Pediatrics/Ward A conflict still blocks all teaching reports and charts. The issue table retains the relevant source OPD, date, worksheet, and cell.

## Review what was adjusted

An **Outpatient priority adjustments (not conflicts)** expander appears when the selected period has affected half-days. It identifies each preceptor, date, AM/PM shift, rotation, archived OPD, worksheet, cell, and whether the source cell originally had a student. An adjustment ID groups the source cells belonging to one half-day.

Download **`Outpatient_Priority_Adjustments.csv`** from the page. It is also included in the teaching-report ZIP when adjustments exist. Student names are omitted. `student_assigned` describes the original cell; `report_action` indicates whether that source was retained or excluded. Adjustments are diagnostic information, not extra countable shifts. They may be reviewed even when an unrelated conflict is blocking report generation.

The Word reports include a brief note when this exception affected the selected period (and, for individual reports, that specific preceptor).

`Archive_Sources.csv` appends `nursery_student_assignment_listings_excluded`, which counts raw excluded student-assignment listings, including repeated rows. Its full-file counts reconcile as:

```
assigned_student_shifts_read
= assigned_student_shifts_counted
+ duplicate_student_shifts_removed
+ nursery_student_assignment_listings_excluded
```

The first credited source for an assignment can change when its nursery occurrence is excluded but a clinic occurrence is retained. The unique assignment is still credited once. These file-level figures describe whole archived OPDs, not just selected reporting dates.

## Files in the partial update

```
schedule_app/services/teaching_priority.py       NEW: explicit rule and adjustment log
schedule_app/services/learner_reach.py          Clinical-shift priority and denominators
schedule_app/services/teaching_analysis.py      Apply the same rule to student assignments
schedule_app/services/teaching_validation.py    Reject scans made before this rule
schedule_app/sections/preceptor_teaching_summary.py   UI, cache refresh, adjustment download
schedule_app/reports/chair_summary.py           Brief explanation in chair reports
schedule_app/reports/individual_teaching.py     Explanation and compact note spacing
schedule_app/reports/teaching_export.py         Adjustment CSV, source counts, report notes
```

The original OPDs are not modified. Archive encryption/save/reload code, student-schedule generation, the individual-schedule ZIP, and the Power Automate assignment workbook are unchanged. This rule applies to teaching-effort reports, not to the underlying clinical schedule or student calendars.

## Validation performed

The final package passes Python compilation and **163 local automated tests**, including **32 new outpatient-priority tests**. One older test that required blocking the nursery/clinic pair was updated to test an unrelated pair, since the requested rule now explicitly permits that overlap.

Tests cover all three clinic sites; both AM and PM; blank clinic fields; separate learners and repeated learners; duplicate rows; nonoverlapping nursery shifts; reverse name order; aliases; overlapping rotations in either reading order; exact-date filtering; July boundaries; hidden zero-teaching entries; strict blocking of unrelated conflicts; stale cached reports; CSV/Word/chart agreement; and preservation of archived bytes.

Both previously supplied OPD workbooks were also checked separately against an independent counting calculation. A synthetic chair report and individual report were rendered and visually inspected. GitHub/Streamlit interactions were simulated; no live repository, saved OPD, or deployment was changed.
