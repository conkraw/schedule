# Admin report editor — local verification

## Final test run

All packaged test files were run in four independent offline batches:

| Batch | Passed | Skipped | Passing subtests |
|---|---:|---:|---:|
| 1 | 271 | 2 | 41 |
| 2 | 330 | 0 | 19 |
| 3 | 386 | 2 | 48 |
| 4 | 250 | 2 | 31 |
| **Total** | **1,237** | **6** | **139** |

This includes 42 new Admin-specific tests. Skips are existing historical-package
checks whose earlier reference trees are unavailable; no runtime failure was
waived. An initial full test run found two old menu assertions that did not expect
the new Admin entry. Those expectations were updated while retaining checks for
OER/PTS last, correct module paths and access protection. A later single-process
run exceeded the execution limit; the complete suite was rerun in the four
independent batches above, without dropping any test file. Existing warnings from
unchanged scheduling modules remain visible in the test logs.

## New coverage

- Catalog fields/defaults/placeholder validation; no arbitrary execution.
- Missing catalog uses defaults without writing; no-op save avoids another commit.
- Encryption, verified readback, another session's reload, stale edit rejection,
  corrupt ciphertext, incorrect keys, unknown keys and unsupported style values.
- Restore one field/all settings leaves unrelated OPDs/catalogs untouched.
- Thread-local build contexts, no sensitive process-global caches, cache invalidation
  and failure handling.
- Password required before editor/network access; authorization rechecked before
  saves; lock clears Admin previews/custom settings along with protected data.
- Admin form saves only selected fields; draft preview does not save.
- New menu order; calculated placeholders still use actual dates/minimum shifts.
- Changed prose affects Word output, not CSV values or chart pixels; source comments
  and data are not editable targets.
- Optional font changes preserve text; default appearance is a strict no-op.

## Default-output comparison

The same invented fixture was passed to the previous release and this release in
separate processes, with no custom wording saved. Canonical archive/XML comparisons
(ignore document timestamps only) were identical for all five checks:

- Teaching reports ZIP, including individual/chair Word, CSVs, chart and notes.
- Individual report with linked OASIS feedback.
- QGenda instructions Word document.
- Student schedule Excel workbook.
- Power Automate preceptor report Excel workbook.

These are controlled example checks, not verification of every possible live
workbook. Existing calculation and identity services, settings.py and requirements
were compared byte for byte; the only service changes outside presentation are the
listed access integration, instructional-text insertion and treating editable OPD
note strings as text rather than auto-detected Excel formulas/URLs.

## Word visual checks

Generated and rendered four invented examples: an individual report (two pages),
a chair report (three pages), an Admin wording proof (one page), and QGenda
instructions (one page). All seven page images were inspected. The examples
exercise new opening text and changed wording with the existing calculations.
Longer future edits or different font choices may change pagination and should
be checked using the actual report. The simple Admin proof is not a full report
layout preview.

No live repository, token, encryption key or Streamlit deployment was accessed.
The code-update archives contain no real source assessments, source OPDs or QA
fixture pickles. No font files are included.
