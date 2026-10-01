# Learner Reach: stop on conflicts + clinical-experience pie charts

This update is for the working modular app with custom dates, Learner Reach and
teaching-contributors-only reports. Keep your launcher, mappings and Secrets.

## Install the small update (recommended)

1. Extract `Schedule_App_Reach_Validation_Pies_Update.zip`.
2. Merge its `schedule_app/` folder into your existing repository. Replace the
   matching files and add the two new modules. Do not delete the existing folder.
3. Add this line to your existing `requirements.txt` if Matplotlib is not already
   included. Do not replace other dependencies:

   ```text
   matplotlib>=3.9,<4
   ```

4. Restart/redeploy the app. Keep the entrypoint `app_sch_2026.py`.
5. Open **Preceptor Teaching Summary**, select dates and a label, then click
   **Load / refresh archived OPDs**. This initial refresh is required: older
   scans did not retain the OPD source coordinates needed for conflict review.
6. Correct any listed conflicts, refresh again, and generate the report ZIP.

The patch deliberately does not contain `settings.py`, a launcher, Secrets,
original OPDs or a replacement `requirements.txt`. `REQUIREMENTS_ADDITION.txt`
contains the single new dependency line for reference. Your email mappings,
name aliases, work-type groupings and encryption key are unchanged.

A full package is also available as `Schedule_App_Modular_Reach_Validation_Pies.zip`.
It includes an updated complete `requirements.txt`. Use the small update to
avoid overwriting locally customized settings.

## New behavior

The app may scan/decrypt the archive to identify problems, but **no teaching
report, report totals/percentages, report ZIP or pie chart is generated when a
clinical work-type conflict affects the selected reporting dates**. Direct calls
to the Word/CSV report builders also enforce this rule. Old report downloads are
cleared when dates change, the source is refreshed, or validation fails.

A blocking conflict is the same normalized preceptor on the same calendar date
and AM/PM half-day listed in more than one clinical experience/work type. It is
blocked even when a student is listed at only one site, or neither site. All
providers in the selected date range are checked, including zero-teaching
providers omitted from the final tables. The app will not guess which source is
correct or reassign the hours automatically.

HOPE_DRIVE, ETOWN and NYES remain one experience, **Academic Pediatrics**.
Repeated entries within that combined experience are deduplicated as before.
Two different students with one preceptor in the same half-day are not a
conflict: educational effort counts two student-shifts, while clinical time
counts that half-day once. AM and PM on the same date are different half-days.

Conflicts on dates outside the selected report period do not block that period.
Unparseable/corrupt source workbooks still stop archive processing rather than
produce a partial scan, as before.

## What the issue table shows

Each conflict has an ID. Each source cell involved is a separate diagnostic row:

- Preceptor name, date, and AM/PM shift.
- Conflicting clinical experiences and the category for this source cell.
- Rotation start date, encrypted archive filename/path, worksheet and exact cell.
- YES/NO indicating a recorded student assignment; no student names are exported.

Download **OPD_Conflict_Review.csv** to keep the issue list. It is a diagnostic
CSV, not a teaching summary. It includes the source GitHub blob identifier.
Sources from two overlapping rotations are both listed rather than blaming
just one file. Diagnostics refer to the archive snapshot just scanned.

To correct a conflict, open **OPD Archive**, select the listed rotation, download
the original workbook and inspect the specified cells. Correct the underlying
schedule, then re-upload the revised OPD through **Create Student Schedule**.
The archive replaces that rotation's current encrypted file as usual. Return to
the summary section and click **Load / refresh archived OPDs** before reporting.
Do not delete a real assignment just to make validation pass; resolve which
clinical experience the actual scheduled half-day belongs to.

## Percentages

Once validation passes, all displayed Learner Reach values in the chair report,
individual reports, CSVs and app totals are numeric percentages. A missing,
invalid or zero denominator cannot be substituted with a fake percentage; a
report requiring such a value is blocked instead of printing N/A.

Learner Reach = recorded clinical half-days with at least one student divided
by all recorded clinical half-days, multiplied by 100. Group percentages use a
ratio of summed shifts, not an average of the individual percentages. The same
four-hour convention is used for both numerator and denominator. CSV percentage
values use a 0-100 scale (80.0 means 80%).

Preceptors and work types with no student assignments during the selected dates
remain omitted. Included preceptors retain their blank shifts in the denominator.
Each area pie/table represents the preceptors with student assignments in that
area, not every provider listed there. Overall totals also retain an included
preceptor's shifts in hidden, zero-teaching work types, so the displayed area
clinical-hour subtotals need not equal the overall denominator. The existing
explanatory note is retained. Education/student-shift totals still reconcile.

The metric describes recorded OPD assignments, not verified attendance, total
unlisted clinical work, teaching quality or a reason a student was not assigned.

## Pie charts

There is **one pie for each included clinical experience**, not one per preceptor.
For example, Academic Pediatrics is one pie combining HOPE_DRIVE, ETOWN and NYES;
Ward A, PSHCH Nursery and Complex Care have their own pies when they have student
assignments. No pie is generated for a wholly unassigned experience.

Each pie shows recorded OPD hours with students versus without students. It uses
the same category totals and percentage as the chair's table, not student-weighted
educational hours. Multiple simultaneous students do not inflate the percentage.
100% Reach shows the full circle with students and a 0-hour remainder.

Pies appear above their category's preceptor table in the combined chair Word
report. They are also viewable in the app after generating the report, under
**Learner Reach pies by clinical experience**. Individual Word reports retain
numeric overall and work-type Learner Reach tables without adding individual pies.

The report ZIP retains the prior Word/CSV/date/source outputs and adds:

```text
Learner_Reach_Charts/                   # A PNG per clinical experience/period
clinical_experience_learner_reach.csv    # Exact numbers behind those pies
```

Charts are embedded PNGs; report tables and wording remain editable Word content.
The app regenerates charts from the current selected data. No charts, reports,
decrypted OPDs or diagnostic CSVs are uploaded to GitHub by this section.

## Files changed

```text
schedule_app/services/learner_reach.py
schedule_app/services/teaching_analysis.py
schedule_app/services/teaching_validation.py        # NEW
schedule_app/reports/chair_summary.py
schedule_app/reports/individual_teaching.py
schedule_app/reports/teaching_export.py
schedule_app/reports/learner_reach_charts.py         # NEW
schedule_app/sections/preceptor_teaching_summary.py
```

Future chart layout edits belong in `reports/learner_reach_charts.py`. Conflict
rules and diagnostics belong in `services/teaching_validation.py`. Neither change
requires editing the main launcher or your Secrets.

## Local checks and limitations

The final package passed 131 automated tests, including simulated Streamlit and
GitHub flows, before delivery. Six previously incompatible legacy checks/fixtures
were updated for the deliberate replacement of warning/N/A behavior with strict
report blocking; remaining scheduling/date/participation checks were retained.
New checks cover source coordinates, overlapping source rotations, selected-date
boundaries, zero-teaching conflicts, blocking all export paths, clearing stale
ZIPs, correcting/reloading an OPD, numeric percentages and chart/table agreement.
Both earlier uploaded OPD examples were also scanned locally. No live repository
or Streamlit deployment was accessed or modified.

The chair preview (four pages), individual preview and 100%/long-label chart were
rendered and visually reviewed. Preview data is explicitly invented. Historical
update/testing documents describe earlier releases; this file and
`TESTING_STRICT_REACH_AND_PIES.md` describe this release.

## Technical references

- https://matplotlib.org/stable/api/_as_gen/matplotlib.axes.Axes.pie.html
- https://matplotlib.org/stable/gallery/user_interfaces/web_application_server_sgskip.html
- https://docs.streamlit.io/develop/api-reference/media/st.image
