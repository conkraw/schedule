# Local verification: documented assessment completion

Release: 2026-10-01-documented-completion-1

The complete test-file set ran in four offline batches. GitHub and Streamlit were
simulated. No live repository or deployment was used.

| Batch | Passed tests | Skipped tests | Additional passing subtests |
|---|---:|---:|---:|
| Completion/matching/settings/new policy | 344 | 3 | 0 |
| Existing encryption/workflow group 1 | 314 | 1 | 35 |
| Existing encryption/workflow group 2 | 316 | 1 | 87 |
| Remaining reporting/workflow tests | 153 | 1 | 17 |
| Total | 1,127 | 6 | 139 |

All six skipped checks require historical release-packaging baselines that are
not part of this distribution. No runtime test failure was waived. The only
modified legacy expectations are those affected by the intended change in
absence/provisional statuses, cutoff, matching-list membership or report wording.

There are 32 new tests in tests/test_documented_completion.py. To run just those:

    python -m pytest tests/test_documented_completion.py -q

The regular runtime dependencies are unchanged. Running developer tests also
requires pytest (and pytest-subtests for the existing subtests). Some legacy test
files use historical external package paths; their comparison checks skip when
those optional baselines are unavailable. Runtime tests use the package itself.

Examples of covered scenarios: 6/8=75% with two absent OASIS names; all eligible
students absent -> 0% only after a valid source check; no sources/unreadable sources
or missing username -> unknown; genuine typo -> numeric provisional result;
explicit correction -> updated counts; later OASIS records auto-match an absent
student; inclusive cutoff and Submit Date limits; future shifts excluded only
from completion; unchanged teaching-time totals; course/form/duplicate handling;
unchanged minimum-shift persistence; no student identifiers in report bundles;
password/session clearing and optional correction controls.

Word reports were rendered using the document renderer. Both individual preview
pages, all three chair preview pages and both provisional QA pages were visually
inspected. The previews are fabricated examples, not live records.

All runtime Python modules compile. The launcher, settings.py, requirements.txt,
teaching_analysis.py, evaluation_access.py, student_assessment_links.py and
assessment_settings.py are byte-for-byte unchanged from the source release.
