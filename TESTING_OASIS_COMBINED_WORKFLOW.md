# Testing — combined OASIS workflow

## Results from this delivered version

- Python syntax compilation completed successfully.
- Full local automated suite: **504 tests passed**.
- New combined-workflow tests: **66 passed**.
- The service and GUI tests use in-memory simulated GitHub and Streamlit calls.
- Actual Streamlit is not installed in this environment; the running browser UI and a live Streamlit Cloud deployment were not exercised.
- No live GitHub repository, production encryption key, or deployment was accessed or changed.

## New test coverage

Cumulative union of old and new source exports; old educators retained when absent from a new-only export; repeated uploads counted once; recalculated question means and comment blocks; conflicting answer copies block rather than silently choose; changing question numbers do not change question identity; inclusive Submit Date endpoints including 23:59:59; cross-July and same-day custom periods; unsubmitted and invalid dates; period metadata; missing usernames and duplicate record IDs; encrypted username correction; original/summary encryption round trips; idempotent saves; same-scope replacement; distinct period/filter outputs; label-only replacement; public-path sanitization; invalid dates; malformed CSV, invalid counts and duplicate IDs; ciphertext tampering, wrong-key rejection and previous-key support; GitHub raw-content fallback; source-list/username revision checks before and after saves; racing saves; shared preset load/update/delete confirmation; source save with no selected dates; retry behavior; automatic saving after username correction; saved original/output recovery; and migration from either former menu label.

## Provided CSV check

A separate local run read the OASIS CSV supplied in the conversation:

- 266 distinct submitted evaluations, across 109 educators.
- One missing-email/username issue was correctly flagged.
- A synthetic manual username was used only in the test to exercise output generation. It is not an asserted real username and was not saved to a live repository or the delivered code.
- A revised in-memory source containing all prior rows plus one synthetic new evaluation yielded **267**, not double the earlier total.
- Original-file and summary-CSV encryption/decryption recovered identical bytes.
- Rebuilding the same period updated one current saved summary in the simulated repository.

The supplied CSV, its actual evaluation results, and source-derived comments are not included in either code package.

## Existing-feature regression coverage

All non-OASIS runtime modules, settings, and requirements are byte-for-byte unchanged from the supplied base package. The only existing service change is allowing the raw OASIS archive list to be read at an explicitly pinned commit.

Older GUI tests now invoke the retained legacy OASIS screens through `tests/legacy_oasis_entrypoint.py`; those screens are no longer menu options in the deployed launcher. The new unified screen and launcher migration are covered separately by `tests/test_oasis_combined_workflow.py`. The test-only router is not imported by the deployed app.

To rerun the suite locally from the complete package after installing its requirements:

```bash
python -m unittest discover -s tests -p 'test_*.py'
```

Some tests intentionally simulate service failures; their error messages are expected. Existing pandas warnings in unchanged legacy scheduling tests do not indicate a failed assertion.
