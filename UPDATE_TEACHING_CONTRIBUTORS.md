# Teaching contributors only — report inclusion update

## What changes

The teaching reports now list only preceptors with at least one student assignment **during the selected reporting period**. Work-type entries also require at least one student assignment by that preceptor in that category during that period.

A service such as ADOLMED disappears automatically when it has no student assignments. It is not permanently excluded: it returns when assignments fall within the selected dates. A preceptor who has no assignments in that period receives no individual Word document or teaching CSV row, even if they taught outside the selected dates.

This applies to the chair summary, individual Word reports, overall and work-type CSVs, monthly Learner Reach CSV, and on-screen teaching tables. In standard multi-year mode the rule is applied separately for each academic year, and empty year sections are omitted.

## Install into the currently working Learner Reach app

Use **Schedule_App_Teaching_Contributors_Update.zip**. It contains only these five updated Python files:

```
schedule_app/
    services/learner_reach.py
    reports/chair_summary.py
    reports/individual_teaching.py
    reports/teaching_export.py
    sections/preceptor_teaching_summary.py
```

1. Extract the ZIP on your computer.
2. Merge its `schedule_app` folder into the existing folder in your app repository, replacing these matching files. Do not delete the existing folder and do not flatten its subfolders. Uploading the ZIP itself does not install the changes.
3. Restart the app. Open **Preceptor Teaching Summary**, select your dates and label, and click **Create teaching reports ZIP**.

Keep `app_sch_2026.py`, `schedule_app/settings.py`, email/name/work-type mappings, requirements, Streamlit Secrets, and the encryption key unchanged. There are no new dependencies or credentials. Student schedules, the Power Automate workbook, and archive upload/encryption/reload functions are unchanged.

The full package, **Schedule_App_Modular_Teaching_Contributors.zip**, is also supplied as an alternative. Prefer the smaller update for your existing working app so it cannot overwrite custom mappings in `settings.py`.

An existing complete Learner Reach scan can be reused. The update invalidates the old report ZIP automatically, but does not require another archive download. If the session has restarted or you need newer OPDs, first click **Load / refresh archived OPDs**. No OPD re-uploading or re-encryption is required.

## Learner Reach remains accurate

This is a **report visibility change**, not removal of non-teaching shifts from the source data. For a participating preceptor:

- Overall Learner Reach includes all of that preceptor's recorded clinical shifts within the selected dates, including shifts without students and shifts in services omitted from the detail tables.
- A displayed work-type percentage uses all recorded shifts for that preceptor in that category, not just the shifts that had students.
- The monthly Learner Reach CSV retains zero-teaching months for a participating preceptor/category. A whole period without assignments is excluded; a quiet month inside a period with teaching is not deleted from the denominator.

For example, eight teaching shifts out of ten recorded shifts still produces **80% Learner Reach**, not 100%. Two simultaneous students still receive two student-shifts/eight student-weighted educational hours, but only one clinical shift in Learner Reach.

The chair's overall percentage now describes **preceptors who had student assignments**, not all providers listed in the OPDs. Entirely nonparticipating providers are excluded from that report's totals.

### Why clinical detail subtotals may be lower than overall hours

A participating preceptor can have recorded clinical time in another service with no student assignments. That service is hidden in the detail tables, but those shifts remain part of the preceptor's overall clinical time. Consequently, the displayed work-type OPD-hour subtotals can be smaller than the overall OPD hours. A short explanatory note appears when this occurs. Student-shifts and student-weighted educational-hour subtotals still reconcile exactly.

Example: one teaching shift and one unassigned shift in Academic Pediatrics, plus one unassigned shift in ADOLMED. Only Academic Pediatrics is listed, with 50% category Learner Reach. Overall Learner Reach remains 1/3 = 33.3%; the hidden ADOLMED shift is not erased from overall clinical time.

Concurrent work-type conflicts remain flagged rather than guessed. Their clinical time remains counted once overall, affected category percentages are withheld, and **Clinical_Shift_Review.csv** retains the affected participating preceptor's details. A zero-teaching review-category summary row is omitted.

## Unchanged reporting scope

Custom dates remain inclusive, and your entered label stays in `academic_year`. Academic Pediatrics still combines HOPE_DRIVE, ETOWN, and NYES; other services retain their prior groupings. Four hours per student-shift is unchanged. Student names are not exported.

The app still reads all current encrypted OPDs. The raw scan and the source-file audit log remain complete; filtering does not delete provider availability or modify GitHub. Source logs and file-level warnings can reference full archived rotations, not just contributors or the selected dates.

If no student assignments are found, the app explains this and does not create empty reports. Blank student fields are not evidence of refusal to teach; scheduled OPD hours do not prove attendance or total clinical working hours.

## Validation

108 local automated tests passed, including 18 new report-inclusion tests. Existing zero-teaching-report expectations were revised for the requested visibility change while retaining tests for the underlying availability counts. Checks cover excluded people/services, automatic return of a service when it has students, both name orders, inclusive date boundaries, quiet months, cross-year reporting, simultaneous students, conflicting work types, CSV/Word consistency, unchanged raw scans, and stale-download invalidation without a network rescan.

Representative chair and individual documents were rendered and visually checked. GitHub and Streamlit calls in tests were simulated; no live archive or deployment was changed. Only the five listed runtime Python files differ from the previous Learner Reach package; the launcher, settings, archive services and other sections are byte-for-byte unchanged.
