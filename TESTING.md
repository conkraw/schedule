# Validation of this modular delivery

## Scope

The comparison baseline was the exact `app_sch_2026 (1)(1).py` uploaded for this refactor, not an older version. Source-to-module locations and the source SHA-256 are in SOURCE_MAP.json.

This delivery reorganizes code. It does not attempt unrelated fixes or changes to scheduling assumptions, parser rules, teaching credit, site groupings, or output schemas.

## Checks performed

**Python compilation:** all delivered Python files compiled successfully under Python 3.13.5.

**26 offline automated tests:** passed. Run the included tests with:

```bash
python -m unittest discover -s tests -v
```

The tests use invented OPD/roster data and test doubles for Streamlit and GitHub. They cover all nine sidebar choices, importing sections without rendering them, OPD template generation, upload and archived-source student schedules, individual schedule ZIPs, the named Excel report table, primary/fallback rules, stale report previews, encrypted round trips, identical/revised/different-rotation saves, wrong-key failures, teaching totals, double-student counting, duplicate handling, work-type grouping, the academic-year boundary, chair/individual report packaging, MD/PA annotations, and shift availability downloads. They do not require a live repository or real secrets.

**70 moved function/class comparisons:** passed against the original Python syntax trees. The comparison disregarded docstring indentation caused by moving nested definitions. Two original render functions were renamed to `render` for the uniform section interface. The preceptor-email lookup was moved inside its report function, preserving its calculation, and one explanatory Excel help message now points to `schedule_app/settings.py` instead of the former monolithic file. Original settings values were compared as well.

**Page control flow:** the executable UI flow for all seven extracted branches matched after removing relocated helper definitions, constants and imports; the remaining two sections were verified as the original render functions. Branch-level imports were placed at module level so wrapping page code in `render()` does not introduce Python local-variable/import errors.

**25 before-and-after scenarios, 30 downloadable payloads:** passed. These comparisons included all nine no-upload screens, populated workflows, both upload and archive-reload paths, and master schedules, individual outputs, and teaching reports generated from both OPD samples in the conversation. The sample workbooks themselves are not included in this package. Comparisons checked nested ZIP contents, CSV data, Excel XML and Word XML; ZIP metadata and document creation timestamps were ignored, as was the intentional email-mapping help-path change. All compared download filenames and widget labels/keys matched.

## Test environment

Python 3.13.5; pandas 2.2.3; NumPy 2.3.5; openpyxl 3.1.5; XlsxWriter 3.2.9; python-docx 1.2.0; requests 2.32.5; cryptography 46.0.4.

The requirements file is unchanged from the prior app package. No extra runtime package is introduced by modularization.

## Limits

Streamlit UI calls and GitHub responses were **simulated**. Streamlit was not installed in the working environment, and a native Streamlit server/Cloud session was not launched. No live repository, token, Secrets setting, encrypted archive, or deployed app was modified. These tests demonstrate local code/output parity on the tested inputs; they do not replace a check of your own deployment after uploading the full folder structure.
