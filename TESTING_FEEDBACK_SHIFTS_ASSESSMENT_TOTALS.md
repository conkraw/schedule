# Validation: feedback, shifts and total assessed students

Full local regression run: **1,196 tests passed, 5 skipped, 158 passing subtests**.
There were 87 existing dependency warnings from unchanged schedule modules. The
full suite completed in 52.23 seconds in this container. No live Streamlit or
GitHub instance was accessed.

The 29 tests in `tests/test_feedback_shift_totals.py` directly cover the new
behaviors. Existing historical output/default tests were changed only where the
request intentionally replaces those behaviors. Internal cohort, data-integrity,
password, encryption and calculation safeguards continue to run.

AST/syntax checks passed. Two fully integrated fictional Word reports were built,
rendered and all seven pages visually checked. The example educator-feedback data
and student assessments are invented. No real assessment CSV, workbook, encryption
key, token, downloaded source file or report is bundled in the app ZIP.

Installation manifest checks and ZIP integrity tests passed. The small update
contains only 14 changed runtime modules, its guide and the hash manifest. The
full package also includes source tests and prior versioned reference guides.
