Current Learner Reach validation: see **TESTING_LEARNER_REACH.md** (90 passing local tests).

---

# Validation: custom reporting dates

## Scope

Only six production Python files are new/modified relative to the previously delivered modular ZIP:

- sections/preceptor_teaching_summary.py
- services/teaching_analysis.py
- services/reporting_periods.py (new)
- reports/chair_summary.py
- reports/individual_teaching.py
- reports/teaching_export.py

All paths above are inside schedule_app/. The launcher, settings, requirements and every other production module are byte-identical to the prior modular package. Supporting documentation and test files were updated.

## Checks performed

**Compilation:** all delivered Python files compile in the local environment.

**56 offline automated tests passed**, comprising the 26 existing modular tests and 30 new reporting-date tests. The existing teaching-page test explicitly selects the retained standard-year mode; it otherwise retains its previous assertions.

```bash
python -m unittest discover -s tests -v
```

New tests cover inclusive boundaries; partial months and rotations; same-day ranges; periods longer than twelve months; July/year crossings without splitting; user-entered labels; leap-day validation; missing/reversed/invalid dates; old monthly-only cache rejection; daily-to-monthly reconciliation; unchanged source scans after filtering; multi-student counting; exact duplicates; conflicting work types; output CSV schemas; true date headings in Word; reporting-date JSON round trips; safe filenames; stale ZIP invalidation; empty periods; optional legacy multi-year reports; and date preference/import workflows.

**Supplied OPD checks:** both mounted OPD samples were encrypted/reloaded against a simulated GitHub service and scanned. Exact partial-date counts matched an independent assignment-level count: Updated_OPD.xlsx, 82 selected student-shifts; Copy of Updated_OPD.xlsx, 61. Full-range totals matched the unfiltered scans (220 and 158 respectively). These example workbooks and their data are not included in the code package.

**Word layout:** representative custom-period chair and individual reports were rendered and visually inspected. The chair sample was two pages and the individual sample one page, with true inclusive dates, reconciled counts, work-type tables and no student names.

## Limits

Streamlit and GitHub calls were simulated with offline test doubles. A native Streamlit server was not launched (Streamlit was unavailable in the working environment). No live repository, deployment, credential, secret or archived OPD was modified. Downloaded reports are still unencrypted staff records; do not commit them to the public repository.
