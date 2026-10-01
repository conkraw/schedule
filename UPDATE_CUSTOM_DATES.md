> **Historical update note:** custom date filtering remains unchanged, but the
> local JSON save/reload controls described below have been replaced by named
> GitHub presets. Use `UPDATE_GITHUB_DATE_PRESETS.md` for the current interface.

# Editable dates for the Preceptor Teaching Summary

The teaching summary now defaults to **Custom dates**. Enter an exact start date, an exact end date, and a label such as `26-27`. That entire interval is treated as one reporting period. It can cross July 1, start/end in the middle of a month or rotation, or be longer than 12 months.

Both the first and last dates are included. When two reporting periods must not overlap, start the next one on the day after the previous one's end date. The example months discussed in chat have NOT been turned into guessed fixed dates: the controls initially wait for your choices.

## Install the small update ZIP (recommended for your already-working modular app)

1. Back up the current code. Extract `Schedule_App_Custom_Dates_Update.zip` locally.
2. Merge its **schedule_app folder** into your repository's existing **schedule_app folder**, preserving the paths below. Replace the matching files and add the new `reporting_periods.py` file. Do **not** delete the existing folder or its other files. Uploading the ZIP itself is not installation.
3. Leave `app_sch_2026.py`, `schedule_app/settings.py`, `requirements.txt`, and Streamlit Secrets unchanged. The small update ZIP intentionally does not contain them, so it will not overwrite your custom name/email/work-type mappings.
4. Restart/reboot the Streamlit app after all six Python files are updated.
5. Open **Preceptor Teaching Summary** and click **Load / refresh archived OPDs** once. The new scan retains provider-level daily counts needed for exact date filtering. The original encrypted OPDs require no changes or re-upload.

Files included in this update:

```text
schedule_app/
    sections/preceptor_teaching_summary.py
    services/teaching_analysis.py
    services/reporting_periods.py           NEW
    reports/chair_summary.py
    reports/individual_teaching.py
    reports/teaching_export.py
```

The complete modular package is also available as `Schedule_App_Modular_Custom_Dates.zip`. Use the small update when preserving local customizations. The full package includes the prior settings file as originally supplied; preserve any mappings you added to your deployed copy.

## Use the controls

In **Preceptor Teaching Summary**:

- **Choose reporting dates:** select **Custom dates** (the default).
- **Start date (included)** and **End date (included):** enter the exact dates to count.
- **Report label / academic year:** enter `26-27`, `27-28`, or another short meaningful label. It does not determine the date range.

Load/refresh archived OPDs if you have not done so, then click **Create teaching reports ZIP**. Download the ZIP, or use the separate chair-summary Word download.

After a scan, changing dates recalculates the totals from the current snapshot without another GitHub download. Changing either date or the label removes the old report download until you generate a fresh one. Click **Load / refresh archived OPDs** after later OPD uploads to use newer saved files.

### Reuse date settings later

Your choices stay selected while you use the current session, including switching to another sidebar section. This is not permanent server storage. To reuse them after the session ends:

1. Open **Save / reload these date settings (optional)** and click **Download reporting-date settings**.
2. Next time, upload that small JSON file in the same area and click **Apply saved reporting dates**.

A copy named `Reporting_Period.json` is also included inside each custom-date report ZIP. The file stores only your report label and dates. It contains no OPD, student names, encryption key or GitHub token. It is not automatically saved to GitHub.

### Outputs

Both existing teaching CSVs keep their existing column names and order. In custom mode, **academic_year contains exactly your entered label**. One row represents one preceptor in that reporting period (and one work type in the work-type CSV).

The chair and individual Word reports show the same label and exact date range. They do not break that period into July-June years. Monthly details include only selected dates, not the whole month merely because part of it overlaps the range.

The ZIP continues to contain both CSVs, the chair summary, individual preceptor reports, source notes and any work-type review. Two small additional files record the exact date definition: `Reporting_Period.json` and `Reporting_Period.csv`. The latter lists `academic_year`, `start_date`, `end_date`, and `both_dates_included`.

The original July-June reporting option remains available under **Standard July-June academic years**. It is optional; it does not control your custom period.

## What has not changed

- All other sidebar sections and the small launcher.
- Original OPD encryption, GitHub storage, latest-file-per-rotation logic and archive reloading.
- Individual student schedules and the weekly Power Automate preceptor workbook.
- Your site groupings: HOPE_DRIVE + ETOWN + NYES = Academic Pediatrics; Ward A, PSHCH Nursery, Complex Care and other services stay separate.
- Four educational hours per student-shift. Two different students in one half-day count twice; exact duplicate assignments count once.
- Scheduled effort is not proof of attendance or unique clock hours.
- Student names are not included in reports. No app password has been added.

## Coverage and diagnostics

Report totals are filtered by the actual assignment date, not the rotation's first Monday. An OPD beginning before the start date can therefore contribute later shifts within the selected dates.

Archive_Sources.csv and the source/data-quality details describe whole scanned files, including files outside the selected period. Their file-level counts are not the date-filtered teaching totals. Missing rotations are not assumed to mean zero teaching. Future scheduled assignments count only when they fall within the selected dates.

The scan still reads the current saved version of every archived OPD at one repository snapshot. It does not count superseded Git history and it does not modify any archive file.

## Verification

56 local automated tests passed: 26 existing modular-app checks and 30 new custom-date tests. Additional independent checks on both uploaded OPD examples matched exact-date counts. Representative chair and individual Word reports were rendered and visually inspected. Streamlit and GitHub calls were simulated; the live deployment was not changed or tested.
