# Teaching-report percentage checks and actionable errors

## Install into the working modular app

Use **Schedule_App_Report_Check_Update.zip** for an existing installation.

1. Extract the ZIP locally.
2. Merge its `schedule_app` folder into the existing repository folder. Replace
   all six matching Python files and add `services/report_diagnostics.py`.
   Do not delete the existing folder and do not upload the ZIP itself as code.
3. Restart the running Streamlit app after committing the seven Python files.
4. Open **Preceptor Teaching Summary**, load/refresh the archived OPDs, choose or
   load the same date preset, and click **Create teaching reports ZIP**.

Near that button, the updated section shows:

```text
Report builder: 2026-09-29-reach-diagnostics-2
```

No change is needed to `app_sch_2026.py`, `schedule_app/settings.py`,
`requirements.txt`, email/name mappings, Streamlit Secrets, the encryption key,
or saved GitHub date presets. No OPD re-upload or re-encryption is required by
this code update. A valid current scan can be reused; restarting may end the
session, in which case Load / refresh is needed.

The full **Schedule_App_Modular_Report_Check.zip** is an alternative for a complete
installation. It is based on the last full Chair Unique Students package supplied
in this conversation, not on a live repository checkout. Preserve any custom
mappings or source edits made only in your deployed repository before using it.

## What the screenshot establishes

The message "Learner Reach is undefined or invalid" comes from a formatting guard
that refuses to publish a missing, nonfinite or out-of-range percentage. It does
not identify which report entry failed, and does not by itself prove that an OPD
contains an error.

The original latest package successfully generated reports from both supplied OPD
examples locally. The failure in the live archive has therefore NOT been
reproduced or attributed to a specific preceptor/record. A missing derived
percentage reproduces the same formatter warning in a controlled test; this is
one handled case, not a confirmed diagnosis of the live failure.

## Changes in this update

- Report rows use complete validated clinical shift counts to derive their hours
  and percentage. If only the derived percentage/hours is missing (`None`), those
  display values are reconstructed exactly; no source assignments are changed.
- Contradictory counts, a real zero denominator, a nonfinite or out-of-range
  percentage, inconsistent existing hours/percentages and unresolved work-type
  conflicts still STOP the reports. Nothing invalid is replaced with 0% or N/A.
- A nonempty report row missing required clinical-count fields now produces a
  specific error, instead of silently treating the missing fields as zero.
- Optional empty preceptor subgroups do not create a table or a 0/0 subtotal.
- Report failures identify the stage (chair calculations, charts, chair Word,
  or individual Word), plus the preceptor, area and counts when known.
- On a report error, the section offers **Download report issue details (CSV)**.
  The file is named **Learner_Reach_Report_Issues.csv**. It includes only report
  labels, preceptor names, clinical area, aggregate counts, and the issue/action.
  It does not include student names, individual student date groups, OPD contents,
  GitHub tokens or encryption keys. Unknown fields are left blank, not guessed.
- Invalid or incomplete output is not kept as a downloadable report ZIP. The
  output version changes clear old ZIP/Word downloads without changing the scan
  schema or saved reporting dates.

If report generation still stops after installation and a refresh, share
**Learner_Reach_Report_Issues.csv** so the actual failing entry can be examined.
For true scheduling conflicts, the existing **OPD_Conflict_Review.csv** source-cell
workflow is unchanged.

## Existing rules preserved

- Custom inclusive reporting dates, GitHub date presets, and weekends.
- Academic Pediatrics = HOPE_DRIVE, ETOWN and NYES.
- The authorized Academic Pediatrics-over-PSHCH Nursery half-day priority rule.
  Nursery learners are not transferred to an unassigned outpatient shift.
- Other simultaneous cross-experience conflicts stop reports.
- Learner Reach = distinct recorded shifts with students / all distinct recorded
  shifts, using all blank shifts for included teaching contributors.
- Two students in one half-day count twice for student-weighted educational hours
  but once for clinical shifts and Learner Reach.
- Unique students and students assigned on 3+ distinct dates in both the chair
  and individual reports; AM and PM on the same date remain one day.
- Preceptors and areas with no student assignments stay hidden.
- All existing report filenames, CSV columns, clinical-experience pies, schedule
  outputs, Power Automate workbook, encrypted archives and preset storage.

## Files in the patch

```text
schedule_app/services/report_diagnostics.py          NEW
schedule_app/services/learner_reach.py
schedule_app/reports/chair_summary.py
schedule_app/reports/individual_teaching.py
schedule_app/reports/learner_reach_charts.py
schedule_app/reports/teaching_export.py
schedule_app/sections/preceptor_teaching_summary.py
```

## Local verification

- Original supplied package: 272 existing tests passed locally.
- Updated package: **297 tests passed**, including 25 new regression/diagnostic
  tests. One old output-signature assertion was adjusted for the new version
  element; its data-calculation assertions were unchanged.
- Both supplied OPD samples generated complete reports with the original and
  updated code in both custom-date and standard-year modes (four paired runs).
  **24 CSV files were byte-for-byte identical; the body/table text of 86 Word
  reports was identical.** ZIP member names matched as well.
- A separate invented-data chair report (two pages) and individual report (one
  page) were rendered and visually inspected, including the pies and continuity
  counts.
- Source compiled successfully. The GitHub transport and Streamlit UI were
  simulated. No live repository, credentials or deployment was used or changed.
