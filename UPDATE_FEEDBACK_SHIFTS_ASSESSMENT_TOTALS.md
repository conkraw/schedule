# PTS: learner feedback, shift-based counts, and total students assessed

This update restores linked learner-feedback attachment as the default, removes the
three-day measure from report outputs, and adds distinct assessed-student totals.
It builds on the Student Count Consistency release and preserves its shared identity
and cutoff logic. It does not edit stored OPDs, evaluations, usernames, or date presets.

## Install the small update (recommended)

1. Extract `Schedule_App_Feedback_Shifts_Assessment_Totals_Update.zip`.
2. Merge its `schedule_app/` folder into the existing repository. Replace **all 14
   matching Python files** at the paths below. Do not delete the existing folder
   and do not upload only the ZIP as an application file.
3. Restart Streamlit. The entrypoint remains `app_sch_2026.py`.
4. Unlock **PTS**, load your OPDs as needed, and select your existing report dates.
5. Expand **Link preceptors to OASIS evaluations (saved in GitHub)** and click
   **Refresh links and OASIS summaries**. Existing saved usernames and exact-date
   summary links are reused. Confirm that the visible **Learner feedback ready**
   notice reports the expected number of linked preceptors.
6. Click **Load / refresh evaluation completeness** once after installing this
   update. Old completion bundles are invalidated because they lack the new totals.
7. Check **Assessments as of** and your saved minimum shifts, then click
   **Create teaching reports ZIP**. Use the newly generated documents.

Keep the launcher, `settings.py`, requirements, Streamlit Secrets, encryption key,
OER/PTS password, saved minimum shifts, date presets, and all mappings unchanged.
No new dependencies, credentials, source uploads, or re-encryption are needed.
A currently loaded OPD scan can be reused; refresh it when the source archive has
changed. Previously downloaded documents do not update retroactively.

## 1. Students' evaluations of the preceptor

The previous link control defaulted to OFF. A saved username/summary link alone
therefore did not ensure that the feedback section was appended. That code path
can omit feedback even when it remains in GitHub. This is a reproducible software
behavior, not proof of what was selected in a particular live browser session.

Linked educator feedback is now ON by default in a new session and at the first
use after this update. The user's explicit choice is held in a separate session
preference so hiding the widget while visiting PTS Matching does not reset it.
No new GitHub preference is written. Locking the protected sections clears the
in-session preference along with their existing protected state.

The default uses the existing, explicitly saved preceptor username and exact-date
OASIS summary binding. It does not guess a person or automatically choose a new
summary scope. When a new summary choice is unsaved, save that selection first.

A status message is visible even when the link controls are collapsed:

> Learner feedback ready: N of M preceptor-period reports will include students'
> evaluations, question averages, and comments.

Matched individual reports retain **Learner feedback on teaching**, submitted-form
count, full question wording (not q588_mean), averages, response counts, and the
existing strengths/improvement comments. The duration-code question stays separate.
Question values and comments are not recalculated or rewritten by this update.

For a requested attachment with no match, the Word report now states the reason,
such as no saved username, no exact-date summary, or no matching educator row.
This is not represented as a zero score. Invalid/decryption/date-mismatch safeguards
still apply. You can explicitly turn linked feedback OFF for teaching-only output;
the app then shows a visible OFF warning and makes no feedback-link network calls.

OASIS educator feedback uses the saved summary's full matching report dates. It is
already aggregated and is **not** sliced by the student-assessment as-of cutoff.
The source period remains printed in the feedback section.

## 2. Use shifts, not a competing three-day measure

The **Students assigned on 3+ days** measure and its definitions are removed from
chair/individual reports, public CSV schemas, previews, and related user-facing
captions. The student section now shows:

- Unique students assigned.
- Students assigned for the selected N+ AM/PM shifts, when completion inputs are
  available; this is the completion denominator.

AM and PM on one date are two shifts. Recognized designations and confirmed name
aliases continue to share one reconciled identity. The student-count window remains
report start through the earlier of report end and **Assessments as of** when the
completion check is enabled. Teaching hours and Learner Reach keep the full report
period. Explicit teaching-only output retains its labelled OPD-name-only counts.

The old three-day statistic can remain in private validation structures for
backward compatibility; it is not an exported measure and does not control eligibility.
No student names or identifiers are added to report files.

## 3. Total students assessed: one per student

Documented assessment completion now includes **Total students assessed (one per
student)**. This counts the distinct **Student External ID** values with a valid
submitted target form by this preceptor within the selected courses, reporting
period, and assessment cutoff.

Three totals are calculated:

- Clinical Assessment of Student: each assessed student once for that form type.
- History Taking & Physical Exam: each assessed student once for that form type.
- At least one of these forms: the union, counting each student once across both.

Multiple question rows, repeated archive exports, multiple submitted forms for the
same student, or both target form types never multiply the combined total. Handoff
forms do not enter these totals. Unsubmitted forms and submissions outside the
selected period/cutoff are excluded. The existing email-derived preceptor username
and saved matching rules remain in use.

**These all-student totals are separate from completion percentages.** They include
students below the minimum shift requirement or absent from the OPD. Percentages
continue to use only assessed members of the OPD-derived eligible group divided
by all eligible students. Do not divide the all-student total by that denominator.

Invented example:

| Measure | Result |
|---|---:|
| Eligible OPD students at 3 shifts | 2 |
| All distinct students with Clinical Assessment | 2 |
| All distinct students with History & Physical | 2 |
| All distinct students with either form | 3 |
| Eligible students with either form | 1 / 2 (50%) |

The clinical and H&P sets overlap, so 2 + 2 does not equal the combined unique total.
A valid check with no eligible students can still show the all-student total;
its completion percentage remains **No eligible students**, not 0% or 100%.
Source failures, absent usernames and contradictory source metadata retain unknown
statuses instead of becoming invented zeros. Unattributed-source warnings remain.

The individual report shows all three totals beside the existing form-specific
completion results. The chair's completion table adds the combined total; its
existing three percentage columns remain unchanged. The completion CSV includes
all three totals and their status.

These are preceptor-to-student assessment counts. The separate students-to-preceptor
feedback section still reports distinct submitted evaluation forms. The minimized
educator-feedback archive cannot establish unique anonymous respondents, and this
update makes no such claim.

## CSV changes

Removed public field: `unique_students_3plus_days` from the overall teaching CSV
and the assessment-completion CSV. Update any separate downstream process using
that retired column.

Appended to `preceptor_student_assessment_completion.csv`:

```
clinical_unique_students_assessed_total
hp_unique_students_assessed_total
either_unique_students_assessed_total
total_assessments_status
```

Existing `clinical_students_evaluated`, `hp_students_evaluated`, and
`either_students_evaluated` remain eligible-student numerator counts. Existing
`clinical_forms_submitted` and `hp_forms_submitted` remain submitted-form audit
counts; they are not the new unique-student totals.

The Power Automate preceptor-assignment workbook and OER output CSV are unchanged.
No source file, matching catalog, question average, or stored report is rewritten
by installing this update.

## Runtime files

| Action | File |
|---|---|
| Replace | `schedule_app/reports/assessment_completion.py` |
| Replace | `schedule_app/reports/chair_summary.py` |
| Replace | `schedule_app/reports/individual_teaching.py` |
| Replace | `schedule_app/reports/preceptor_evaluations.py` |
| Replace | `schedule_app/reports/teaching_export.py` |
| Replace | `schedule_app/sections/assessment_completion.py` |
| Replace | `schedule_app/sections/assessment_settings.py` |
| Replace | `schedule_app/sections/preceptor_oasis_links.py` |
| Replace | `schedule_app/sections/preceptor_teaching_summary.py` |
| Replace | `schedule_app/sections/pts_navigation.py` |
| Replace | `schedule_app/services/assessment_completion.py` |
| Replace | `schedule_app/services/educational_time.py` |
| Replace | `schedule_app/services/student_continuity.py` |
| Replace | `schedule_app/services/teaching_evaluations.py` |

## Validation and limits

- Python syntax/AST validation succeeded for all packaged Python files.
- Full local suite: **1,196 passed, 5 skipped, 158 passing subtests**. Existing
  warnings from unchanged scheduling dependencies were not hidden.
- 29 new tests cover distinct students versus forms/question rows, set union,
  dates/courses/cutoff, threshold independence, missing sources/usernames, genuine
  zero, invalid totals, feedback defaults/opt-out, missing-link explanations,
  navigation/widget-cleanup behavior, source immutability and privacy.
- Existing tests that required the removed three-day column or the previous OFF
  default were updated to the requested behavior. One historical release-diff
  test skips when its prerequisite historical baseline is unavailable; runtime
  failures were not waived.
- The fictional individual preview contains both assessment completion and two
  submitted educator-feedback forms with full questions and sample comments.
  All four individual pages and all three chair pages were rendered and inspected.
- Existing archive encryption, column minimization, password access, OPD parser,
  saved mapping services, settings, launcher, and requirements were preserved.

GitHub and Streamlit calls were simulated. No live repository, deployment, or
credentials were accessed or changed. The Word previews contain invented data;
they do not establish revised counts or actual current feedback for any real preceptor.

Technical reference for non-widget session preference persistence:
https://docs.streamlit.io/develop/concepts/architecture/widget-behavior
