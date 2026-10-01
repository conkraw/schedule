# PTS: ignore or restore OPD student entries

## Purpose

Explicitly exclude a nonstudent label (for example `midcycle feedback`) or a
student entry you deliberately do not want included in PTS. The choice is saved
encrypted in the configured GitHub archive and is reusable in later sessions.
This is not a change to the original OPD, an OASIS deletion, or a hidden warning
dismissal. It changes the PTS reporting population.

## Install the small update

1. Extract `Schedule_App_PTS_Ignored_Students_Update.zip`.
2. Merge its `schedule_app/` folder into your existing repository. Replace the
   five matching files and add the two new files listed below. Do not delete the
   existing folder, and do not upload just the ZIP as a runtime application file.
3. Restart the app and unlock **PTS** with the existing OER/PTS password.
4. Click **Load / refresh archived OPDs** once to populate the student-entry list.

Keep `app_sch_2026.py`, `settings.py`, requirements, all preceptor/username/name
mappings, Streamlit Secrets, encryption key, saved minimum shifts, and date
presets unchanged. No new dependency, credential, or source upload is needed.
The full modular ZIP is an alternative for a clean installation; the small ZIP
is preferable for preserving custom settings in an existing deployment.

## Ignore an entry

After loading the OPDs, open **Ignore or restore OPD student entries** near the
archive-load controls in PTS.

1. Choose one or more names/labels under **OPD student entries to ignore**. The
   list shows the student-side entries found in the loaded OPDs, excluding names
   already ignored. Custom-date mode limits these suggestions to those dates.
   It is a name-only list; no student external IDs are displayed.
2. Optionally type an exact label in **Exact OPD entry not in the list**.
3. Check **Exclude these entries from PTS counts and name-match alerts for all
   reporting dates**.
4. Click **Ignore selected entries and save to GitHub**.

Only a verified encrypted save applies the exclusion. A failed or competing
save does not dismiss any name or silently discard the existing catalog.
After a successful save, an already-loaded OPD snapshot is automatically reread
with the new list. If that read fails, use **Load / refresh archived OPDs** again;
old report downloads are not retained as current.

Click **Load / refresh evaluation completeness** afterward to recompute the
matching queue and assessment percentages, then generate a new teaching ZIP.
Old completeness inputs are cleared rather than reused with a changed denominator.
No already downloaded document changes retroactively.

## Restore an entry

The same panel has **Ignored entries — restore when needed**. Select entries
under **Ignored student entries to restore**, confirm, and click **Restore
selected entries**. The verified change removes them from the current ignore
list and starts a fresh teaching scan when one was previously loaded. Refresh
evaluation completeness afterward.

Previously saved student-name corrections remain intact. A restored name that
already matches OASIS will not need another confirmation. A restored genuine
mismatch will reappear in the usual missing-name correction dropdown.
Restoration is available even when every teaching entry has been excluded and
there are no reports to generate.

## Exactly what is excluded

For a matching student-side label, PTS excludes that entry before calculating:

- Student-name matching alerts and dropdowns.
- Unique students, the separate three-distinct-days continuity count, and
  eligibility at the saved minimum-shift threshold.
- Teaching assignments, educational hours, Learner Reach numerators, work-type
  and monthly totals, chair and individual reports, teaching CSVs, and pie charts.

The preceptor's valid clinical listing stays in **total scheduled availability**.
If no other retained student is in that half-day, it becomes a clinical shift
without a student. If a real student is also listed in that half-day, that
student and the four educational hours remain. Multiple retained students still
produce only four educational hours in the same preceptor/date/AM-or-PM shift.

A provider with no remaining student assignments is omitted from the teaching
reports under your existing contributors-only rule; their availability remains
in the internal OPD data. Existing outpatient-over-nursery priority and other
clinical-area conflict checks still apply. Ignoring a student does not hide a
preceptor's conflicting clinical availability.

This feature does NOT modify:

- Original archived OPDs or the student schedule generation tools.
- OASIS source CSVs, educator feedback, or OER summary calculations.
- Saved name corrections, preceptor usernames, date presets, minimum shifts,
  password protection, or encryption keys.

Do not ignore a real student solely to remove a missing-assessment warning.
The selection deliberately changes eligibility. Unmatched, non-ignored students
remain in the denominator and keep the existing Not verified safeguard.

## Matching scope

The ignore list is shared across users, rotations, dates, preceptors and work
types in this configured PTS archive. It matches the exact parsed OPD entry,
ignoring case, repeated whitespace and comma spacing. It does not use substring,
keyword, wildcard or fuzzy matching. For example, ignoring `midcycle feedback`
does not also ignore `midcycle feedback today`.

Program designations are deliberately retained in ignore keys: excluding
`Example, Jordan (PA)` does not automatically exclude `Example, Jordan (MD)`.
Select both labels to exclude both. This does not change the separate OASIS
identity-matching rule, which still automatically ignores recognized trailing
program/class designations when a student identity is unique.

## Saved location and security

```
opd_archive/pts_ignored_student_entries.json.enc
```

The actual base folder follows your existing Streamlit Secrets configuration.
Only the explicit labels and their update timestamps are in this new catalog;
no assessment answers, scores, student IDs, or source workbooks are copied into
it. All writes use the current catalog revision and verified encrypted read-back.
The app reuses the existing token/key and refuses an unreadable catalog rather
than assuming an empty list. Deleting/restoring an entry changes the current list;
it is not erasure of original sources or Git history.

The source-label dropdown inventory lives only in password-protected session
state and is cleared by **Lock OER / PTS**. It is not exported or uploaded. Report
notes disclose only the number of excluded source listings in the selected dates,
not their names. No ignored-name list is included in the report ZIP.

Use **Refresh ignored student entries from GitHub** to obtain another session's
changes. An explicit OPD reload also checks the current catalog. Before building
a ZIP, the app rechecks the exclusion list; a changed list requires refreshing
rather than combining old teaching counts with new eligibility rules.

## Runtime files

| Action | File |
|---|---|
| Add | `schedule_app/services/ignored_student_entries.py` |
| Add | `schedule_app/sections/ignored_student_entries.py` |
| Replace | `schedule_app/services/teaching_analysis.py` |
| Replace | `schedule_app/services/assessment_completion.py` |
| Replace | `schedule_app/sections/preceptor_teaching_summary.py` |
| Replace | `schedule_app/sections/student_name_matches.py` |
| Replace | `schedule_app/reports/teaching_export.py` |

The Word-layout modules are unchanged. Data values are recalculated by the same
existing report builders after applying explicit exclusions.

## Technical references

- GitHub revision-aware file updates: https://docs.github.com/en/rest/repos/contents
- Streamlit widget/session state: https://docs.streamlit.io/develop/api-reference/caching-and-state/st.session_state

See `TESTING_PTS_IGNORED_STUDENTS.md` in the full package for local validation.
No live GitHub repository or Streamlit deployment was accessed or changed.
