# GitHub date-preset validation

- 212 tests passed with `python -m unittest discover -s tests -v`.
- Includes 49 new storage/UI-callback tests in `tests/test_reporting_presets.py`.
- The old JSON settings UI callback test was replaced with a saved-GitHub-preset callback test.
- The conflict-only download test now expects the conflict CSV without the retired settings JSON button.
- All schedule, teaching, Learner Reach, conflict, priority, chart and report tests remain.
- All runtime modules compile.
- GitHub HTTP and Streamlit widget interactions are offline simulations; no live credentials were used.
- Real Streamlit launch was not tested: Streamlit was not installed, and the package index was unreachable.
- No document layouts or spreadsheet generators were modified in this update.
