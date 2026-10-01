# Verification: PTS student-cohort consistency correction

## Execution

All 29 packaged test modules were executed in three fresh-process batches to avoid cumulative test-run resource growth. Commands used `PYTHONPATH=.:tests pytest -q ... --disable-warnings`.

| Batch | Passed | Skipped | Additional passing subtests |
|---|---:|---:|---:|
| 1 | 385 | 1 | 59 |
| 2 | 381 | 3 | 40 |
| 3 | 400 | 2 | 40 |
| Total | **1,166** | **6** | **139** |

The six skips are existing historical release-packaging checks, not newly disabled runtime failures. An initial all-in-one attempt hit a tool time limit; the three batches together cover every test module. No runtime test was omitted to obtain the reported totals.

The new `test_student_cohort_consistency.py` module adds **39 passing tests**, including parameterized cases. It covers designation-only duplicates, saved spelling aliases, designation-sibling propagation, conflicting identities, genuinely different learners with identical dates, absent assessments, missing usernames, cutoff boundaries, future-only students, weekends, multiple settings, thresholds 1–5, report/table/CSV agreement, invalid/tampered cohorts, old schemas, unchanged hours, no GitHub writes during calculation, and no learner identifiers in exports.

Four existing tests were updated for deliberately changed behavior: assessment schema version, continuity marked Not checked until the enabled assessment inputs are loaded, and two tests that formerly compared entire time-row dictionaries rather than unchanged hour fields. The last two now assert that every hour/availability/Reach field stays identical while student-count metadata follows the cutoff. No failure was waived.

## Before/after reproduction

The same standalone synthetic reproduction was run against the supplied baseline package and the corrected package. One learner appeared with and without `(MD)` on three identical dates. Baseline output was two unique students/two three-day students/one eligible student. Corrected output was one/one/one. Both versions returned 12 educational hours, 12 scheduled hours and 100% Learner Reach.

A second scenario contains five real synthetic learners in the full schedule, but only three qualify before the selected cutoff. The corrected chair table, individual report and both overall/completion CSVs show the same three unique students, three three-day students and three eligible students. Two confirmed clinical assessments yield 66.7%; future assignments still contribute to the separate full-period teaching-time output.

## Documents and packaging

Two synthetic Word documents were generated using the delivered application code: the individual report (two pages) and chair report (three pages). Each was rendered with the DOCX skill renderer, and every rendered page was visually inspected for clipping, alignment, page breaks and matching student-count dates. The illustrative documents contain no actual faculty/student assessment data.

All runtime Python files compile. Archive-integrity checks and source-file comparisons confirm that the update contains only the 12 intended runtime files plus instructions/manifest. The launcher, settings.py, requirements and persistent catalog services remain unchanged. No font files, credentials, original user workbooks or evaluation CSVs are included in the update ZIP.

GitHub and Streamlit are simulated for tests. This does not verify the live deployment or supply recalculated actual totals for the user-uploaded faculty reports.
