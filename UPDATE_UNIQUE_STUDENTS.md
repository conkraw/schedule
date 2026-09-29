# Individual preceptor reports: unique students and 3+ teaching days

This update builds on the modular app with GitHub date presets, Learner Reach,
clinical-experience pies, and the Academic Pediatrics-over-PSHCH Nursery rule.
It adds two numbers near the top of each individual Word report. It does not
change the existing CSV column names/order or add these metrics to the chair report.

## New numbers

**Unique students assigned:** the number of different students assigned to this
preceptor at least once within the report's selected dates, across all work types.

**Students assigned on 3+ days:** the number of those students assigned to this
preceptor on at least three distinct calendar dates within the same period.
This is a subset of the first number, not an additional group to add to it.

| One student's assignments to a preceptor | Distinct days | Included in 3+ days? |
| --- | ---: | --- |
| Monday PM, Tuesday AM, Thursday PM | 3 | Yes |
| Monday AM, Monday PM, Tuesday AM | 2 | No |
| Monday AM and Monday PM | 1 | No |
| Friday PM, Saturday AM, Sunday PM | 3 | Yes |

Days need not be consecutive or full days. The threshold does not reset each
week, month, rotation, or work type within a custom reporting period. The same
student is counted only once in a preceptor's period total, even after a return
visit in another rotation. AM and PM on one calendar date count as one day.
Two different students in one shift count as two unique students, with one day
credited to each. Three different students assigned on one day each do not
qualify anyone for the 3+ days count.

## Date controls and counting safeguards

Your current custom dates and GitHub date presets continue to work unchanged.
Both endpoints are included, and days outside the selected period do not help a
student reach three days. Custom dates can cross July, calendar years, and more
than twelve months without resetting the count. In the optional standard
academic-year mode, each year's Word section calculates its own counts.

Counts use the existing assignment parsing and duplicate handling. The
outpatient/nursery priority rule is applied first: an excluded nursery student
or excluded nursery date contributes neither a unique student nor an extra day.
A blank clinic student field does not inherit a student from the nursery.
Remaining cross-work-type conflicts still block reports as before.

Unique-student matching uses the student names in the OPDs, ignoring case,
extra spaces and comma spacing. It does not guess that different spellings
are the same person. People with identical recorded names cannot be separated
without an additional identifier in the source. Student names are still not
listed in any teaching report.

These are scheduled assignments, not verified attendance. This new 3+ **days**
metric does not change the primary-preceptor rule, which still uses **sessions**.
Learner Reach, OPD hours, student-shifts, educational hours and pies are unchanged.

## Install the small update in your working modular app

1. Extract `Schedule_App_Unique_Students_Update.zip` on your computer.
2. Merge its `schedule_app` folder into the existing folder in the repository.
   Replace matching files and add the new file. Do not delete the existing folder.
3. Restart the app, then open **Preceptor Teaching Summary**.
4. Click **Load / refresh archived OPDs** once. Old scans cannot reconstruct
   unique students from already-summed shift totals, so they must be refreshed.
5. Select or load your date preset, then click **Create teaching reports ZIP**.
   The individual Word files in `Preceptor_Reports` now contain both counts.

The update has five Python files:

- `schedule_app/services/student_continuity.py` — NEW: groups dates and computes counts.
- `schedule_app/services/teaching_analysis.py` — retains the minimum date-group detail after exclusions.
- `schedule_app/reports/individual_teaching.py` — displays both numbers and a short definition.
- `schedule_app/reports/teaching_export.py` — validates the new data before exporting; adds report notes.
- `schedule_app/sections/preceptor_teaching_summary.py` — invalidates old scans/downloads and prompts a refresh.

Leave `app_sch_2026.py`, `schedule_app/settings.py`, email/name mappings,
`requirements.txt`, Streamlit Secrets and the encryption key unchanged.
There are no new dependencies, credentials, or GitHub setup steps. No OPDs need
to be edited, re-uploaded or re-encrypted for this update. The full-app ZIP is
an alternative to the small update; preserve your own settings before using it.

## Storage and privacy

No new student data is saved to GitHub. During a scan, run-local keyed values
assemble each student's dates per preceptor. Those keys, values and student
names are discarded before the scan is returned. Only unlinked date groups
(one list entry per student) are held in that session to support changing the
report dates without another archive download. Two students with identical
dates remain two separate entries. The date groups are not exported in the
report ZIP; only the two aggregate counts appear in individual Word reports.

Normal archive/preset saves and reloads, the chair summary, clinical-experience
pies, student schedules, and the Power Automate workbook retain their existing
behavior. The same access warning still applies: an app with no access gate can
be used by anyone who can reach it. This update adds no authentication.

## Validation

The combined local suite passed **249 tests**, including **37 new unique-student
tests**. Tests use synthetic OPDs and simulated GitHub/Streamlit interactions.
They cover the exact three-day example, same-day AM/PM, multiple students,
weekends, cross-rotation/cross-year dates, precise cutoffs, name normalization,
nursery exclusion, stale scans, unchanged conflict blocking, privacy and exports.
The single-period and multi-year Word outputs were rendered and visually checked.
No live GitHub repository or Streamlit deployment was accessed or changed.
