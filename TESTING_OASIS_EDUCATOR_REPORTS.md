# OASIS Educator Reports: validation performed

## Automated suite

Command: `python -m unittest discover -s tests -q`

Result on this delivery: **438 tests run, 437 passed, 1 skipped**. No failures.
The skipped inherited OASIS archive test expects a prior packaging baseline at a specific local directory. A separate archive-byte comparison verified that the only pre-existing Python file changed in this delivery is the launcher.

There are **78 new tests** in `tests/test_oasis_educator_reports.py`, covering:
- evaluation counts versus question rows and same-student repeated evaluations;
- question numbering changes, ID/wording checks, different question denominators;
- exact duplicate and overlapping snapshots versus conflicting answers;
- numeric means, N/A/blank exclusions, invalid numeric inputs and duration-code labels;
- combining comments, preserving multiline quotes, literal n/a and repeated comments from separate forms;
- matching educator identities without substituting the exported username for missing email;
- manual usernames, duplicate record IDs, date filters and submitted-only forms;
- encrypted username saving/reloading/removal, wrong keys, corrupt data and concurrent edits;
- simulated UI source selection, missing-username blocking, manual correction, persistence and final downloads.

GitHub and Streamlit are **simulated**. No live repository or deployment was accessed. Real Streamlit is not installed in the execution environment, so this is not a live browser/Cloud deployment test.

## Supplied CSV checks

The conversation's OASIS CSV was parsed independently using Python's standard `csv` library and compared with the new report service:

- 4,162 question-response rows.
- 266 distinct submitted evaluations.
- 109 educators.
- 17 multiple-choice question IDs across old/new form versions.
- One educator without a source email; the final CSV remained blocked until an explicitly entered test username was supplied in a simulation. That dummy was not delivered as a real educator ID.
- All 1,853 educator/question means and all 1,853 scored-response counts matched the independent calculation.
- All 218 per-educator comment collections (strengths and improvement) matched the independently combined comments.
- Reading the same CSV twice removed 4,162 duplicated question rows and produced unchanged educator totals and comments.

The original CSV and any real comments/identifiers are **not bundled in the app package**. User-facing CSV previews use invented example data only.

## Regression scope

All previous runtime modules, settings, requirements, encryption/archive code, teaching-effort calculations, outpatient/nursery priority, reports and date presets are byte-identical to the supplied latest OASIS Archive package. The existing launcher changes only by adding one menu entry. Three new runtime modules contain the report, UI and optional encrypted username mapping logic.
