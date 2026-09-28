# Which file should I edit?

## One file for each sidebar section

All section files are under **`schedule_app/sections/`**. Each contains a `render()` function that runs when that option is selected. The teaching-summary section now includes editable date controls; the other sections retain their existing workflows.

| Sidebar choice | File | What to edit here |
|---|---|---|
| Instructions | `instructions.py` | Onscreen QGenda instructions and the downloadable instruction document. |
| Format OPD + Summary | `format_opd_summary.py` | QGenda uploads, provider filtering, site/provider lists, assignment rules and OPD-plus-summary workflow. |
| Create Student Schedule | `create_student_schedule.py` | Original OPD upload/reload, rotation-list validation and master-schedule workflow. |
| OPD Check | `opd_check.py` | Baseline-versus-assigned OPD comparison and its Word change report. |
| Create Individual Schedules | `create_individual_schedules.py` | Master schedule upload, individual schedule ZIP and Power Automate report workflow. |
| OPD Archive | `opd_archive.py` | Archive screen and its reload/use-in-schedules navigation. |
| Preceptor Teaching Summary | `preceptor_teaching_summary.py` | Custom date/label controls, optional standard-year selection, saved-date JSON and report downloads. |
| OPD MD PA Conflict Detector | `opd_md_pa_conflict_detector.py` | MD/PA comparisons, available-preceptor suggestions and annotated-download controls. |
| Shift Availability Tracker | `shift_availability_tracker.py` | Availability grids, site selection and capacity displays. |

`format_opd_summary.py` is the largest section (under 800 lines). Its workbook-writing helpers have been separated, but its interdependent assignment sequence stays together so the original scheduling behavior remains intact.

## Common edits: settings

Open **`schedule_app/settings.py`** for:

| Setting | Purpose |
|---|---|
| `PRECEPTOR_EMAIL_MAP` | Preceptor names and email addresses for the weekly Power Automate report. |
| `TEACHING_PRECEPTOR_NAME_MAP` | Explicit spelling/name aliases for teaching totals. Do not assign rotating generic provider slots permanently to one person. |
| `TEACHING_OPD_NAME_ORDER_OVERRIDES` | Specific rotation dates whose `~` name order differs from the selected default. |
| `TEACHING_WORK_TYPE_MAP` | Combines HOPE_DRIVE, ETOWN and NYES as Academic Pediatrics; keeps Ward A, PSHCH Nursery and Complex Care separate. |
| `TEACHING_WORK_TYPE_ORDER` | Order of work-type sections in teaching reports. |
| `TEACHING_HOURS_PER_STUDENT_SHIFT` | Existing value of 4. No inpatient/outpatient multiplier is added. |
| `FOCUS_SITES` | Sites included in the weekly primary/fragmented preceptor report. Separate from the all-site teaching-summary scope. |
| `REPORT_COLUMNS`, `TEACHING_CSV_COLUMNS`, `TEACHING_WORK_TYPE_CSV_COLUMNS` | Output schemas. Keep unchanged unless deliberately updating downstream processing. |
| Report schema versions | Used to invalidate older generated results when their structure or rules change. |

Email example:

```python
PRECEPTOR_EMAIL_MAP = {
    "Example, Preceptor": "verified-address@example.org",
}
```

Use verified real addresses in your deployment. Names in the supplied source were not supplemented with guessed emails. **GitHub tokens and encryption keys do not belong in `settings.py`; they remain in Streamlit Secrets.**

Restart/reboot the running app after editing Python modules or settings so it imports the saved code. Refresh the archived OPDs after changing teaching-name or work-type mappings, then regenerate the teaching reports. Rebuild individual schedules after changing an email mapping; an already generated download is not retroactively edited.

## Report layout edits

These files are under **`schedule_app/reports/`**:

| File | Purpose |
|---|---|
| `chair_summary.py` | Combined chair Word report, overview totals, work-type sections and preceptor tables. |
| `individual_teaching.py` | Individual preceptor Word reports, reporting-period headings and monthly work-type tables. |
| `teaching_tables.py` | Shared table appearance used by the individual teaching report. The chair report retains its own table formatter. |
| `teaching_export.py` | CSV creation, ZIP contents, report filenames and source/quality notes. |

The weekly **Power Automate Excel report** is a different output: edit `services/primary_preceptors.py` for its logic and Excel appearance, not the teaching Word files.

## Calculation, Excel and archive helpers

These files are under **`schedule_app/services/`**:

| File | Purpose |
|---|---|
| `opd_archive.py` | Encryption/decryption, original workbook/date validation, GitHub reads/writes, replacement by rotation and archive listing. |
| `opd_archive_ui.py` | Shared archive upload/reload widgets, saved-file verification and navigation callbacks. |
| `student_schedules.py` | Parse student/preceptor assignments, create the master schedule template, and populate its worksheets. Includes Protected Self-Study Time wording. |
| `opd_workbooks.py` | OPD Excel template, CSV mapping into the template and hiding blank rows. |
| `primary_preceptors.py` | One primary per represented student/week, threshold/fallback/reuse flags, email lookup and `PreceptorAssignmentTable`. |
| `workbook_copy.py` | Copy a student worksheet to its individual workbook while retaining the existing formatting behavior. |
| `teaching_analysis.py` | Read-only archive analysis, daily counts, exact-date filtering, duplicate handling and work-type totals. |
| `reporting_periods.py` | NEW: custom date/label validation, saved JSON settings and report date/label helpers. Dates are entered in the app, not hard-coded here. |
| `md_pa_analysis.py` | MD/PA date parsing, booking/availability maps and annotated Excel copies. |
| `availability_analysis.py` | Shift parsing, Hope Drive grouping and weekly student-capacity calculations. |

Helpers use normal named imports. Importing a page does not render it; the launcher calls its `render()` function on every rerun. User uploads are not saved in persistent Python module globals.

## Adding another sidebar section later

Create `schedule_app/sections/new_section.py`:

```python
import streamlit as st


def render():
    st.subheader("My New Section")
    # Put this section's widgets and workflow here.
```

Add its label and filename without `.py` to `SECTIONS` in `app_sch_2026.py`:

```python
"My New Section": "new_section",
```

Keep this under `schedule_app/sections/`, not a top-level `pages/` folder. This app uses its original sidebar radio selector.

## What not to change during this installation

Do not regenerate encryption keys, move encrypted archives, introduce a second app-password setting, change Power Automate column names, or replace the archive repository. Deploy the package alongside its entrypoint and keep your existing Secrets.

## Changing reporting dates next year

Use the **Start date**, **End date** and **Report label / academic year** controls in Preceptor Teaching Summary. No Python edit is required. See **UPDATE_CUSTOM_DATES.md** for saving/reloading a small JSON date preset.
