# Student-evaluation archive — local validation

## Results

The final modular package was tested with:

```text
python -m compileall -q app_sch_2026.py schedule_app tests/test_oasis_student_evaluations.py
python -m unittest discover -s tests -p 'test_*.py'

Ran 675 tests
OK (skipped=2)
```

**673 passed, two skipped.** The two skips are existing historical/packaging checks whose comparison baseline is not supplied in this standalone package. The new module has 61 tests; all passed. The tests use simulated Streamlit and GitHub, not a running Cloud deployment. Existing tests emitted library deprecation warnings; none failed.

## New coverage

- All three form types in the supplied source (clinical assessment, handoff, history and physical), separately and together.
- Correct direction routing and rejection of mixed/unknown form types without silently omitting rows.
- Original bytes, encoding, multiline quotes/comments, HTML/question wording, and formula-like literal text preserved.
- Missing emails/usernames, missing form IDs, and blank submission dates do not destroy or drop original records.
- Source limits, malformed CSVs, and date uncertainty produce bounded or explicit behavior.
- Separate student vs educator archive folders and secret-keyed file identifiers.
- No student/preceptor names or assessment comments in public filenames or on-screen metadata.
- Exact encrypted save/load verification; repeat-file no-op; changed-file snapshot retention.
- Previous-key recovery, wrong-key/tampered-file failures, renamed ciphertext detection, source-path validation.
- Auth/write/verification failures; lost save confirmations and deliberate retry without duplicate snapshots.
- Read-only recovery, current vs historical reads, stale-selection protection, and directory-list completeness limits.
- No student source contribution to existing educator summaries; unchanged OPDs, date settings, and preceptor links.
- New nested OASIS selector; no date or username dependency on student upload; no student report generation.
- Separate session-state namespaces, no repeated saving on unrelated reruns, and immediate retry/status controls.

## Uploaded sample check

A separate local check used the supplied `oasis_eval_export(2).csv`:

- 5,231,178 bytes.
- 7,084 question-response records.
- 800 distinct Course ID / Evaluation / Form Record combinations.
- Recognized `*Clinical Assessment of Student`, `*PEDS Handoff`, and `*PEDS History Taking & Physical Exam`.

The sample was encrypted and reloaded using simulated GitHub transport, including the raw-content fallback used when inline contents are unavailable. Its recovered and optional-download bytes exactly matched the upload. Identical re-upload produced no duplicate, and the educator source list stayed empty when only the student file was stored.

Only aggregate test results are recorded here. The source CSV, names, grades, comments, generated encryption keys, and ciphertext from this check are not part of the software package. No live GitHub repository or Streamlit deployment was accessed or changed.
