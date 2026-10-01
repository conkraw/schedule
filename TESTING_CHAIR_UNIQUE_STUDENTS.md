# Chair student-continuity update: local verification

Run from the extracted app folder:

```bash
python -m compileall -q app_sch_2026.py schedule_app
python -m unittest discover -s tests -q
```

272 tests passed locally, including 23 new tests in
`tests/test_chair_student_continuity.py`. The fixtures use invented student and
preceptor names. GitHub transport and Streamlit widgets/session state are
simulated; no live service credentials are required by these tests.

This is a report-presentation update. It reuses the original
`student_continuity_counts` function rather than estimating unique students from
shift totals or summing monthly/student counts. Existing individual-report and
CSV generation routines remain unchanged apart from shared ZIP notes.

The chair-report output version is included only in the generated-download
signature. Updating the report clears stale ZIP/Word downloads without changing
the scan signature, encrypted archive, or selected reporting preset.

A new table is intentionally overall per preceptor, not per clinical setting.
There is no sum across preceptors, because learners can appear with multiple
preceptors. All existing setting-specific effort tables and pies remain.
