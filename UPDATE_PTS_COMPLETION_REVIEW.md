# PTS assessment-completion review

## What this update addresses

A report can successfully read both archives yet withhold its assessment-completion
percentages because an eligible OPD student's identity is unresolved. The earlier
report named the number of unmatched students but did not identify them. The
simplified PTS screen hid the supporting table unless full diagnostics were opened.

This update adds an optional, focused explanation and a direct route to the
specific student matches. **It does not relax the matching safeguard or invent a
percentage. An unresolved student still remains in the eligible denominator.**

The supplied Fogel report specifically says eight students met the three-shift
minimum and two eligible student names required a match. His educator-feedback
username link was already working. These are separate checks.

A local read of the supplied Book1 workbook found 201 question rows, 17 Clinical
Assessment forms and five History & Physical forms, all attributed to bfogel.
All 22 submitted forms were within the displayed reporting period. Those forms
covered 17 distinct students; the five H&P students were also in the clinical set.
This confirms that these records can be parsed, not that all assessed students
met the OPD eligibility threshold. The two missing identities and exact eligible
subset cannot be inferred from this spreadsheet and the identifier-free report.
The live GitHub archive was not accessed.

## Install (small update recommended)

Extract `Schedule_App_Completion_Review_Update.zip` and merge `schedule_app/` into
the existing repository. Replace/add these four files at their matching paths:

| Action | File |
|---|---|
| Replace | `schedule_app/sections/assessment_completion.py` |
| Replace | `schedule_app/sections/student_name_matches.py` |
| Add | `schedule_app/sections/assessment_diagnostics.py` |
| Add | `schedule_app/services/assessment_diagnostics.py` |

Do not delete the existing folder. Do not upload only the ZIP. Restart Streamlit.
Leave `app_sch_2026.py`, `settings.py`, requirements, passwords, encryption key,
GitHub settings, date presets and saved mappings unchanged. No re-encryption or
new dependency is needed. No original source upload is required for this code
update. An incomplete source archive can still require the appropriate OER export.

## Investigate a missing percentage

1. Unlock **PTS**, load the OPDs and choose the correct reporting period.
2. Click **Load / refresh evaluation completeness** once after installing the
   update. This also reloads the currently saved identity mappings.
3. Enable **Show why an assessment percentage is unavailable (optional)**.
4. Choose the preceptor/reporting period, for example Fogel, Ben.
5. Review the eligible-student total, saved username, total submitted form counts,
   and the table of **only the eligible unresolved OPD student names** for that
   preceptor. Student IDs and assessment answers are not displayed in this table.
6. Click **Fix these student names in PTS Matching**. The correction queue is
   filtered to the selected preceptor's unresolved eligible names. Match each
   name to the correct OASIS name and use the existing confirmation/save button.
7. After the encrypted saves are verified, return to PTS and generate a fresh ZIP.
   Completion results are recomputed from the current in-session inputs. For
   changes made from another session, use **Recheck saved student matches from
   GitHub** or reload evaluation completeness.

The targeted queue updates after each save. **Show all unmatched student names**
restores the ordinary missing-only queue. Changing the report dates, OPD snapshot
or exclusions cancels an obsolete targeted filter. A filter does not exclude
anyone from the actual calculations, and lower-threshold unresolved students
remain reviewable in the normal matching screen.

## What to do with a name that cannot be selected

- A genuine spelling/name-order discrepancy: select the correct OASIS student
  and confirm it through the existing matching workflow.
- The student is absent from all loaded OASIS names: check which student-assessment
  exports have been uploaded in OER. A student without any recorded assessment
  may not supply an identity to that crosswalk. Do not choose another person or
  ignore a real eligible student merely to obtain a percentage.
- Duplicate OASIS identities: review the source records. The app still will not
  guess between different students sharing a name.
- A mapping was saved elsewhere: **Recheck saved student matches from GitHub**
  reloads that encrypted catalog without downloading every OPD again. A new OPD
  or OER upload requires **Load / refresh evaluation completeness**, not only
  the mapping recheck.

This update is not a new student-roster upload or an automatic identity merge.
If the correct identity is not available, the percentage remains Not verified
rather than treating an unverified name as no assessment or removing it.

## Why form counts are different from the numerator

The displayed submitted-form counts include forms for all students evaluated by
that preceptor within the Submit Dates, even students who had fewer than the
required OPD shifts. They are evidence that the parser found assessments, not
values to divide directly by the eligible-student denominator.

The unchanged calculation remains:

> eligible unique students with at least one target form from this preceptor
> divided by all unique students meeting the selected minimum OPD shift count.

The union measure counts each eligible student once even when both forms exist.
No changes are made to educational hours, Learner Reach, unique students,
three-distinct-day continuity, the saved minimum shifts, dates, weekend handling,
outpatient-over-nursery priority or the report layouts.

## Privacy and security

The detail table and targeted matching filter stay in the password-protected
session. No new catalog, plaintext GitHub file, student-name diagnostic download
or report field is created. The existing encrypted identity-save mechanism is
unchanged. Locking OER/PTS clears the new diagnostic state. Rechecking saved
matches is read-only and never silently replaces an unreadable catalog with an
empty one. Students' names, IDs and assessment answers are not added to reports.

## Validation

- 25 new offline tests cover targeted diagnostics, eight-eligible/two-unresolved
  scenarios, source-form counts versus numerator membership, verified correction,
  failed refreshes, cross-page routing, context changes, privacy and password gates.
- Regression suite ran in two batches: **1,095 passed; six skipped**, plus 139
  passing subtests. The skips require historical release-packaging baselines that
  are not present; runtime tests were not disabled.
- The actual completion engine, matching rules, report builders, launcher,
  settings and requirements are byte-for-byte unchanged from the supplied
  simplified-PTS package. Only the four modules listed above changed or were added.
- The user-supplied workbook was read locally and converted in the temporary test
  workspace to exercise the existing CSV parser. That temporary file, the user's
  report, spreadsheet, students and evaluation data are NOT in the update ZIP.
- GitHub and Streamlit interactions were simulated; no live deployment was changed.

Technical reference for callback order and session-local controls:
https://docs.streamlit.io/develop/concepts/architecture/widget-behavior
