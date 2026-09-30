# PTS: choose and remember the minimum shifts for assessment completion

This update makes the student-assessment eligibility threshold adjustable in PTS.
It supersedes earlier instructions that fixed that threshold at three shifts.
It does not change educational hours, Learner Reach, primary-preceptor selection,
or the separate **Students assigned on 3+ days** continuity measure.

## Install the small update

1. Extract `Schedule_App_PTS_Minimum_Shifts_Update.zip`.
2. Merge its `schedule_app/` folder into your existing repository, replacing the
   six matching files and adding the two new files below. **Do not delete the
   existing folder and do not upload just the ZIP to the repository.**
3. Restart Streamlit. Unlock **PTS** with your existing OER/PTS password.
4. Load your OPDs and date preset as usual. Under **Student assessment completion
   and evaluation-record checks**, leave **Include assessment completion and
   missing-evaluation alerts** checked.
5. Set **Minimum shifts for assessment completion**. Click **Load / refresh
   evaluation completeness** once after this update so the current completion
   inputs are loaded. Then create the teaching reports ZIP as usual.

Keep `app_sch_2026.py`, `schedule_app/settings.py`, requirements, all username and
student-ID mappings, Streamlit Secrets, the encryption key, and GitHub settings
unchanged. No new dependencies, credentials, migrations, or source re-uploads
are needed. The full modular ZIP is an alternative for a clean installation;
use the small update to preserve your deployed settings and customizations.

## Normal use

**Minimum shifts for assessment completion** begins at **3** only when no saved
setting exists. Use the +/- buttons or type a whole number (for example 4 or 5)
and commit the edit by pressing Enter or leaving the field. The app automatically
saves the change, encrypted, in GitHub. There is no separate JSON file to upload
and no separate Save button.

A verified save displays:

> Minimum shifts: 4. Saved encrypted in GitHub and reused next time.

The next new session loads 4. Changing it to 5 makes 5 the next default. Returning
to 3 is also supported. Valid values are whole numbers from 1 through 10,000.

This is **one shared PTS setting for the configured archive**, not a per-user or
per-date-preset setting. Other users of the same archive get the last saved value
when they open a new session. An already open session can use **Reload saved
minimum shifts** to retrieve another session's change. A stale edit is not
silently used to overwrite a newer saved value.

Date presets continue to store dates and labels only. Loading a date preset does
not reset or change the saved minimum shifts.

## Exactly what changes in the calculation

For a selected minimum N:

- Denominator: unique students assigned to the same preceptor on **N or more
  distinct date + AM/PM shifts** within the report's dates.
- Numerator for a form: students in that same eligible group with at least one
  submitted assessment of that form type from the preceptor within the report's
  Submit Date boundaries.
- Clinical Assessment, History Taking & Physical Exam, and either-form measures
  keep their existing distinct-student counting rules.

A student with three shifts is eligible at N=3, but not at N=4 or N=5. A student
with five shifts is eligible at each of those settings. An assessment for a
student below the chosen minimum does not enter that percentage's numerator.

Monday AM and Monday PM are two shifts; they are still just one calendar date
for the independent 3+ days continuity measure. Weekends count. Exact duplicate
assignments do not add shifts. Selected reporting dates, confirmed student-ID
aliases, and the Academic Pediatrics-over-PSHCH Nursery priority rule continue
to apply before eligibility is determined.

Unmatched student names remain correctable even when they are below the minimum.
Their warning table now says whether the unmatched identity affects the chosen
threshold's completion percentage. An unresolved eligible student is never
removed from the denominator to improve the percentage.

## Reports and CSV

The current selection appears in the chair and individual assessment-completion
sections, on-screen result columns, student-name alerts, and `Report_Notes.txt`.
For a four-shift minimum, the reports say **Students assigned for 4+ shifts**.
No eligible students is labeled as such, not 0% or 100%.

The completion CSV uses two explicit fields instead of a fixed-three header:

- `eligible_students` replaces `eligible_students_3plus_shifts`.
- `minimum_shifts` is appended and records the selected number in every row.

This applies only to `preceptor_student_assessment_completion.csv`. Other
teaching CSVs, the Power Automate preceptor-assignment workbook, and the OASIS
summary CSV keep their existing schemas and calculations. Update any external
process that specifically used the former completion-CSV header.

Changing the minimum clears the previous downloadable teaching ZIP. Once inputs
are loaded, the app recalculates from them without downloading all OPDs or OASIS
sources again. Generate a fresh report ZIP after the change. Previously
downloaded documents are not retroactively changed.

## Encrypted storage and access

The app creates one small settings file automatically:

```
opd_archive/pts_assessment_settings.json.enc
```

The base folder follows your existing `[opd_archive]` settings. The file contains
only a format/version marker and the chosen number. It does not contain student
or preceptor data and does not replace date presets, username links, student-ID
links, OPDs, or OASIS files.

The existing encryption key and GitHub token are reused. Writes are checked
against the loaded file revision and re-read/decrypted at the save commit before
they are called verified. The field is inside protected PTS; locked/expired
sessions cannot load or save it. **Lock OER / PTS** clears the in-session value,
not the setting saved in GitHub.

## If a load or save fails

- An unavailable/corrupt existing settings file does **not** silently reset to 3.
- A failed or unverified save is clearly labeled as not confirmed saved. It is
  not retried on every unrelated widget interaction.
- Use **Retry saving minimum shifts** after a temporary save failure.
- Use **Reload saved minimum shifts** after a stale-edit conflict or uncertain
  verification, review the saved value, and then make the intended change.
- Until the setting is verified, completion percentages are marked **Not
  checked**. Teaching-hour reports remain available with that explanation;
  percentages are not invented using an unconfirmed threshold.

## Runtime files

| Action | File |
|---|---|
| Add | `schedule_app/services/assessment_settings.py` |
| Add | `schedule_app/sections/assessment_settings.py` |
| Replace | `schedule_app/services/assessment_completion.py` |
| Replace | `schedule_app/sections/assessment_completion.py` |
| Replace | `schedule_app/services/student_name_review.py` |
| Replace | `schedule_app/sections/student_name_matches.py` |
| Replace | `schedule_app/reports/assessment_completion.py` |
| Replace | `schedule_app/reports/teaching_export.py` |

## Verification and technical references

See `TESTING_PTS_MINIMUM_SHIFTS.md` in the full package. Testing uses invented
records and simulated GitHub/Streamlit; no live archive or deployment was changed.

Implementation references:

- Streamlit number input and edit/commit behavior:
  https://docs.streamlit.io/develop/api-reference/widgets/st.number_input
- Session state and widget lifecycle:
  https://docs.streamlit.io/develop/api-reference/caching-and-state/st.session_state
- GitHub repository contents and revision-aware updates:
  https://docs.github.com/en/rest/repos/contents
