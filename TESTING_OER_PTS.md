# OER / PTS release verification

All execution was local, with synthetic fixtures and simulated Streamlit/GitHub.

- Batch A: 548 passed, 2 skipped.
- Batch B: 289 passed, 1 skipped.
- Combined: 837 passed, 3 skipped (840 collected tests).
- The 57 new OER/PTS tests are included in Batch A.
- Skips concern optional historical release-comparison folders, not access tests.

New tests are in `tests/test_oer_pts_access.py`. Run them from the project root with `python -m pytest -q tests/test_oer_pts_access.py`.

Existing UI test fixtures explicitly sign in when testing allowed behavior; the test helper still defaults to no login. Runtime settings.py, requirements.txt and every report module are byte-for-byte unchanged from the supplied prior package. All Python files compile, and ZIP CRC checks pass. No live deployment or repository was used.
