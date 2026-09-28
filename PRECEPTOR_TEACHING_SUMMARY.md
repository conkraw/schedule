# Preceptor Teaching Summary

Current version: teaching totals by **custom reporting period** and type of work.

Open Preceptor Teaching Summary, choose **Custom dates**, enter the exact start/end dates and a label such as **26-27**. Click **Load / refresh archived OPDs**, then **Create teaching reports ZIP**. The app downloads/decrypts the current saved OPD for each rotation at one repository snapshot. It does not read superseded Git history or save decrypted reports to GitHub.

Both dates are included. The actual assignment date determines inclusion, including partial months/rotations. A custom period stays together even when it crosses July or exceeds twelve months. The academic_year CSV field uses your entered label. All Word reports display the exact dates. Future scheduled assignments within the selected dates are included; these are not attendance records.

The original July-June mode is still available as an optional choice for standard academic-year reports. No month/day is forced in custom mode. Your exact dates were not guessed from the example months discussed in chat.

To retain dates beyond the current browser session, use **Save / reload these date settings (optional)** to download and later import a small JSON file. See **UPDATE_CUSTOM_DATES.md** for detailed instructions.

HOPE_DRIVE, ETOWN and NYES are grouped as Academic Pediatrics. Ward A, PSHCH Nursery and Complex Care are separate. Other sites retain separate work-type labels. See EDITING_GUIDE.md for the settings and module locations; see README.md for installation instructions.

## Counting

One student assigned to one AM/PM session = one student-shift = four educational hours. Two different students in the same session count twice. Exact duplicate preceptor/student/date/AM-or-PM assignments count once. Provider availability without a student does not count. No setting-specific multiplier is applied. These are student-weighted scheduled hours, not unique clock hours.

## Names and data quality

Select the actual name order around ~ in archived OPDs: Preceptor ~ Student or Student ~ Preceptor. Spaces surrounding names are ignored. Separate multiple students with semicolons/newlines or use separate rows; commas inside Last, First names are preserved. Invalid repeated ~ entries fail visibly instead of being guessed. Optional per-rotation name-order overrides and preceptor spelling aliases remain in the Python code.

Student names are used only while counting and deduplicating; they are not included in the generated CSVs, Word reports or notes. Generic site/slot/combined provider labels remain flagged rather than attributed to a guessed individual. Assignments with no identifiable provider are reported as missing attribution and excluded from provider totals. An unreadable/decryption-failed file aborts the scan instead of generating misleading partial totals.

If the same exact assignment occurs in different work types, it is retained once under Work type needs review and listed without student names in Work_Type_Review.csv. Type subtotals must reconcile to the overall totals before export.

## Outputs

- One combined chair Word summary, divided into reporting periods and work types, with a separate Word-only download button.
- One Word report per preceptor, with reporting-period and monthly work-type breakdowns.
- preceptor_teaching_summary.csv: the original overall CSV with unchanged columns; academic_year is the custom label.
- preceptor_teaching_by_work_type.csv: preceptor/year/work type/months/shifts/hours/source sites.
- Report_Notes.txt and Archive_Sources.csv, plus Work_Type_Review.csv when necessary.
- Reporting_Period.json and Reporting_Period.csv in custom mode record the exact inclusive dates and label.

File-level source/quality diagnostics describe full scanned OPDs, not just the chosen date range. The teaching totals themselves use only assignments inside the chosen dates.

Reports remain a snapshot until Load / refresh archived OPDs is clicked again. Changing the grouping map or upgrading the report schema clears older cached results. Existing OPDs do not need to be re-uploaded. Keep your existing encryption key and other secrets unchanged.

## Access

The supplied app has no additional password gate, as requested. Access to the running app can expose decrypted files and generated reports unless restricted elsewhere. Encryption protects the stored GitHub contents, not who may operate the running app. Do not commit decrypted OPDs or report ZIPs to the public repository.
