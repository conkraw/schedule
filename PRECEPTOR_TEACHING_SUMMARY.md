# Preceptor Teaching Summary

Current version: teaching totals by academic year and type of work.

Open Preceptor Teaching Summary, click Load / refresh archived OPDs, select the academic year(s), and click Create teaching reports ZIP. The app downloads and decrypts the current saved OPD for each rotation at one repository snapshot. It does not read superseded Git history or save decrypted reports to GitHub.

The default academic year is July 1 through June 30 (for example, 26-27). Dates come from the actual scheduled assignments. Select other archived years to add separate sections to the same chair and individual Word documents. Future scheduled assignments in a selected year are included; these are not attendance records.

HOPE_DRIVE, ETOWN and NYES are grouped as Academic Pediatrics. Ward A, PSHCH Nursery and Complex Care are separate. Other sites retain separate work-type labels. See EDITING_GUIDE.md for the settings and module locations; see README.md for installation instructions.

## Counting

One student assigned to one AM/PM session = one student-shift = four educational hours. Two different students in the same session count twice. Exact duplicate preceptor/student/date/AM-or-PM assignments count once. Provider availability without a student does not count. No setting-specific multiplier is applied. These are student-weighted scheduled hours, not unique clock hours.

## Names and data quality

Select the actual name order around ~ in archived OPDs: Preceptor ~ Student or Student ~ Preceptor. Spaces surrounding names are ignored. Separate multiple students with semicolons/newlines or use separate rows; commas inside Last, First names are preserved. Invalid repeated ~ entries fail visibly instead of being guessed. Optional per-rotation name-order overrides and preceptor spelling aliases remain in the Python code.

Student names are used only while counting and deduplicating; they are not included in the generated CSVs, Word reports or notes. Generic site/slot/combined provider labels remain flagged rather than attributed to a guessed individual. Assignments with no identifiable provider are reported as missing attribution and excluded from provider totals. An unreadable/decryption-failed file aborts the scan instead of generating misleading partial totals.

If the same exact assignment occurs in different work types, it is retained once under Work type needs review and listed without student names in Work_Type_Review.csv. Type subtotals must reconcile to the overall totals before export.

## Outputs

- One combined chair Word summary, divided into academic years and work types, with a separate Word-only download button.
- One Word report per preceptor, with annual and monthly work-type breakdowns.
- preceptor_teaching_summary.csv: the original overall annual CSV with unchanged columns.
- preceptor_teaching_by_work_type.csv: preceptor/year/work type/months/shifts/hours/source sites.
- Report_Notes.txt and Archive_Sources.csv, plus Work_Type_Review.csv when necessary.

Reports remain a snapshot until Load / refresh archived OPDs is clicked again. Changing the grouping map or upgrading the report schema clears older cached results. Existing OPDs do not need to be re-uploaded. Keep your existing encryption key and other secrets unchanged.

## Access

The supplied app has no additional password gate, as requested. Access to the running app can expose decrypted files and generated reports unless restricted elsewhere. Encryption protects the stored GitHub contents, not who may operate the running app. Do not commit decrypted OPDs or report ZIPs to the public repository.
