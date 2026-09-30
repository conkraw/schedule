# Verification: selectable PTS assessment minimum shifts

## Results

- Python compilation succeeded for all runtime modules.
- **55 new tests** cover the setting, its persistence, calculations, warnings,
  report text/CSV, and the release's limited runtime change set.
- The complete local suite finished with **933 passed, 3 skipped**, plus 158
  passing subtests. The three skips are existing environment/history-specific
  checks, not failures waived for this feature.
- Two complete synthetic Word reports (individual and chair) were rendered to
  page images and visually reviewed, including the **4+ shifts** headings,
  eligibility counts, fractions/percentages, and the unchanged **3+ days** counts.

The new tests use in-memory GitHub and Streamlit test doubles. Real Streamlit is
not installed in this runtime, and no live repository, password, encryption key,
or deployed app was accessed. No source uploads were altered.

## Coverage

Persistence tests cover the initial 3-shift default, encrypted round trips for
3/4/5 and boundary values, new-session reload, redundant-write suppression,
unchanged OPDs and mappings, wrong keys, invalid catalogs, duplicate JSON keys,
concurrent edits, missing save receipts, and mismatched verification results.

UI tests cover committed numeric edits, immediate autosave, successful values
returning in a new session, section changes, lock/expiry authorization checks,
clearing old report downloads, keeping loaded source data, and explicit
retry/reload paths. Load failures never become a silent 3-shift default; save
failures never claim that a choice was retained.

Calculation/report tests cover 3/4/5 thresholds, zero eligible students,
numerator membership, duplicate shifts, exact date boundaries, unchanged hours
and continuity, missing-name flags that follow the selected minimum, schema and
report-signature consistency, dynamic chair/individual/CSV/notes labels, no
student identifiers in report tables, and teaching-only generation with an
unavailable threshold marked Not checked.

The existing complete suite additionally covers OPD encryption, OER column
minimization, student/preceptor mappings, outpatient-over-nursery precedence,
weekend handling, saved dates, assessment matching, and other report workflows.

## Test maintenance

Earlier assessment tests now reference `eligible_students`, the intentional
replacement for the fixed `eligible_students_3plus_shifts` field. The shared
Streamlit test double gained a numeric-input method. One historical OASIS-only
packaging check now detects that its baseline already contains the combined
workflow by module presence, instead of an obsolete menu label. A new packaging
check verifies the exact eight runtime files changed in this release.

## Reproduce locally

Install the app requirements and pytest, then run:

```
python -m compileall -q schedule_app
python -m pytest -q tests
```

The source-baseline packaging check skips when its earlier-release comparison
directory is not available; that can change the pass/skip totals on another
computer. Tests do not require live GitHub access.
