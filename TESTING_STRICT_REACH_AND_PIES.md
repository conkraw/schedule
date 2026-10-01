# Validation of strict Learner Reach reporting and clinical-experience pies

Run from the extracted full package:

```bash
python -m unittest discover -s tests -p 'test_*.py'
```

Final result: **131 tests passed**. Tests use invented OPDs and in-memory GitHub
responses. Streamlit UI calls are simulated with the supplied test helper. No
user tokens, encryption keys, live archive data or external writes are required.
Existing warnings from unrelated pandas/openpyxl code do not fail these tests.

The new `tests/test_strict_reach_charts.py` includes 23 tests for source-level
conflict detection, exact date filtering, diagnostic privacy, blocked exports,
stale ZIP removal, correction-and-refresh workflows and chart/CSV/Word agreement.
Tests assert that every exported Learner Reach value is numeric and in 0-100;
invalid/missing values block output instead of appearing as N/A. Legitimate
multi-student shifts and same-experience duplicates still work as before.

Two original OPD examples provided earlier in this conversation were checked
separately using locally simulated archive encryption and transport:

| Example | Student-shifts | Distinct clinical shifts | Conflicting clinical half-days |
|---|---:|---:|---:|
| Updated_OPD.xlsx | 220 | 1,934 | 2 |
| Copy of Updated_OPD.xlsx | 158 | 1,277 | 0 |

All detected sample conflicts retained their filename, rotation, worksheet and
cell details. These are sample-file checks, not the user's full live archive.

QA documents used invented provider/student names. The four-page chair summary,
one-page individual report, and long-label 100% pie were rendered and visually
inspected. Student schedules, the Power Automate workbook, archive encryption,
settings and launcher files were not changed by this release.
