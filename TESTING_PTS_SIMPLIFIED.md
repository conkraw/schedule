# Simplified PTS verification

## Scope

Tests use invented records, a simulated Streamlit interface, and simulated GitHub
requests. No live deployment, GitHub repository, encryption key, or evaluation
dataset was accessed. The native Streamlit server was not launched in this
container; existing repository test helpers exercise the rerun/callback paths.

## Automated checks

```
python -m pytest -q --disable-warnings --tb=short
1070 passed, 6 skipped, 87 warnings, 139 subtests passed in 41.80s
```

The six skips are pre-existing environment/historical-package checks for exact
old release files that are not available under the test harness's historical
paths. No failed behavioral check was reclassified as skipped. Existing warnings
were not suppressed as fixes; their display is summarized by the test command.

Thirty-five new checks live in `tests/test_pts_simplified_workspace.py`. They
cover the three protected matching tasks; gate-before-read behavior; same-password
navigation and lock cleanup; missing-only inputs; keeping sources/choices across
pages; no routine PTS tables or images; optional diagnostics; no extra ZIP build
on ordinary reruns; unchanged-build reuse with verification; invalidation after
date changes; failed-build clearing; progress stages; and once-per-build shared
tables. Canonical Word XML comparisons cover standalone versus batch generation.

Some existing UI expectations changed intentionally: matching actions are tested
on PTS Matching, charts/diagnostic tables require opt-in, and pressing Create on
an unchanged completed result reuses it. Tests that inject a new-build failure
now clear the old result first, so the failure path is still exercised.

Settings, requirements, source encryption/persistence services, the assessment
completion engine, and OPD calculation modules were checked unchanged from the
previous delivered package. No column lists or saved catalog schemas changed.

## Comparative output test

A shared, synthetic scan of 60 preceptors was processed by both the previous and
updated exporters. It generated the same files:

- 61 Word documents (60 individual reports and the chair summary)
- 5 CSV files
- 1 clinical-experience PNG
- 1 report-notes file

CSV, PNG, and note bytes matched exactly. Word XML and other Word package parts
matched after canonicalizing XML and excluding core-document creation/modified
timestamps. Report headings, tables, relationships, and embedded media matched.
The data fixtures are synthetic; none of the user's student files is packaged.

An individual Word report and a short chair report were rendered to page images
and inspected. This update changes app UI and repeated computation, not report
layouts or metric definitions.

## Local performance test

60 preceptors; same scan; fonts/imports warmed; three timed builds per version;
no network; no shared cross-build cache; every timed run produced the full ZIP.

| Version | Seconds per run | Median |
|---|---|---|
| Prior package | 4.763, 4.653, 4.673 | 4.673 |
| Updated package | 3.098, 2.730, 2.718 | 2.730 |

The updated median was about 42% lower in this specific local benchmark. This
is not a guaranteed user-facing speedup: larger data, OASIS comments, document
length, host resources, first-run fonts, and GitHub latency affect real times.
Shared full-archive teaching-time calculations are now performed three times per
ZIP (overall, by work type, monthly), not three additional times per preceptor.

All reuse is local to the build or the authenticated session. Existing exclusion
and OASIS link freshness checks still run before an unchanged ZIP can be reused.
