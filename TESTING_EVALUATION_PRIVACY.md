# Evaluation Records privacy update — local verification

## Release checks

The final local test suite completed with **781 passed, 2 skipped** (87 warnings; 139 subtests passed) on Python 3.13. The new `tests/test_evaluation_privacy.py` contributes **54 tests**, including parametrized checks. Runtime modules and the optional local password generator also passed Python compilation.

Tests use simulated GitHub and Streamlit. No live repository, credentials, or deployed application was accessed or changed. This is not an independent security audit or a certification of regulatory compliance. The skipped cases depend on optional historical baseline packages that are not part of this release. Warnings include existing dependency/deprecation warnings; tests were not run with warnings treated as errors.

## New coverage

- Both upload types use explicit allowlisted fields, before any encryption or write.
- Unknown fields and sentinel private values are absent from stored decrypted data.
- Supported text encodings, retained comments/quotes/newlines, deterministic serialization and idempotence.
- Existing educator averages/comments/identity issues and student assessment identities/counts remain unchanged.
- Changes only to discarded fields do not create new snapshots.
- Wrong-direction uploads fail before network writes.
- Legacy sources are verified before minimization; downloads expose only reduced contents.
- Optional old-source cleanup requires review and confirmation, verifies the replacement before deletion, respects source revisions, and retains safe retry behaviour.
- Missing/wrong password, timeouts, password changes, per-session attempt delays, session separation and lock/clear behaviour.
- Current and legacy evaluation screen entrypoints cannot render uploads or start network work while locked.
- Unrelated sections, including teaching reports, remain password-free and unchanged.

## Provided-source comparisons

The two source CSVs supplied in the conversation were processed locally. Their contents are **not included** in this package.

| Source | Original columns | Retained columns | Rows retained | Comparison |
|---|---:|---:|---:|---|
| Educator feedback | 33 | 15 | 4,162 | All educator summary rows and username issues unchanged; 109 educators. |
| Student assessments | 33 | 11 | 7,084 | Prepared assessment form metadata and student matching unchanged; 735 identified target forms. |

The student CSV becomes smaller because question text, answers, grades and assessment narratives are removed. Question-level rows are preserved in source order; the existing completion calculator deduplicates them at the form level.

## Access boundary deliberately preserved

The password protects **Evaluation Records administration only**. Existing evaluation information linked into Preceptor Teaching Summary and downloadable teaching reports remains available under that section's existing controls. A renamed menu, shared-password gate and encrypted storage do not replace institution-approved deployment access controls. See UPDATE_EVALUATION_PRIVACY.md.

Older full encrypted sources can remain in Git history and copies even after their current repository paths are replaced. No history rewrite was performed or included as an automatic operation.
