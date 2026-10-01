# PTS: documented assessment completion and an assessment cutoff

## What changes

This release supersedes the older rule that required every eligible OPD student
to have an OASIS name/Student External ID match before displaying a percentage.
OPD assignments establish eligibility. OASIS establishes which assessments are
on file. An assessment file is not a complete student roster.

**Documented assessment completion = eligible students with a confirmed submitted
assessment / all eligible OPD students.** The saved minimum-shifts setting still
controls eligibility. The same student counts once per preceptor/form type;
question rows, duplicate exports and repeated forms do not multiply the numerator.

- **No assessment on file:** an eligible student with no matching OASIS record
  remains in the denominator, contributes no numerator credit, and needs no name
  confirmation simply because that record is absent.
- **Name review needed:** an ambiguous identity or a plausible spelling/name
  discrepancy remains available for correction. The numeric percentage uses
  confirmed matches only and is marked **provisional** with an asterisk. No
  similar-looking name is automatically linked or credited.
- **Assessment on file:** a verified link to a submitted target form contributes
  to that form's numerator, provided the same student meets the OPD threshold.
- Failed source downloads/decryption, no loaded assessment sources, a missing
  preceptor username or an unverified minimum-shift setting still mean **Not
  checked**, not zero. Contradictory or invalid assessment metadata attributable
  to a preceptor still produces **Not verified**, not a guessed percentage.

Example using invented records: eight eligible students, six with confirmed
Clinical Assessment forms, and two absent from OASIS -> **6 / 8 (75.0%)**. The two
absent records do not force name confirmations. A valid, successfully checked
source with no forms for the eligible group can produce **0%**; an unreadable or
unloaded source cannot. No eligible students is shown as **No eligible students**,
not 0% or 100%.

The two target forms remain Clinical Assessment of Student and History Taking &
Physical Exam. The either-form percentage uses their student-set union, not the
sum. Counts describe loaded records, not proof an assessment was never completed
or that one is overdue. New rotations and incomplete exports can both explain
absent records. This release does not implement an overdue policy or due-date
calculation.

## Install the small update (recommended)

1. Extract `Schedule_App_Documented_Completion_Update.zip`.
2. Merge the `schedule_app/` folder into the existing repository. Replace the nine
   matching Python files and add the new helper listed below. **Do not delete the
   existing folder or upload just the ZIP as an application file.**
3. Restart Streamlit. Keep the entrypoint `app_sch_2026.py`.
4. Unlock PTS, load the OPDs and choose your report dates/date preset as usual.
5. Click **Load / refresh evaluation completeness** once after this update. Old
   completion inputs are invalidated rather than reused with changed logic.
6. Generate a new teaching reports ZIP. Previously downloaded documents cannot
   change retroactively.

Keep the launcher, settings.py, requirements, all existing passwords, encryption
key, Streamlit Secrets, saved minimum shifts, date presets, username links and
student-name corrections unchanged. No new dependency, credential, repository,
OPD upload or re-encryption is needed. The full modular ZIP is an alternative for
a clean install; the small update avoids replacing deployed customizations.

## Assessments as of

The new **Assessments as of (included)** control starts at the local date when a
new session first opens these controls. Choose today or a past date. It stays
selected while switching between PTS and PTS Matching in the current session.
The existing date-preset catalog is unchanged; this cutoff is not automatically
written into the presets or saved as a new GitHub setting.

For assessment completion only, both OPD assignment dates and OASIS student-
assessment Submit Dates must be inside the report period **and on or before the
cutoff**. The reporting end date still caps that intersection. The entire cutoff
day is included; future dates cannot be chosen. A cutoff before a report period
starts produces no eligible students for that period. Clearing the control asks
you to select a date rather than silently applying one.

Example: the teaching report covers March 1, 2026 through March 1, 2027, and the
cutoff is October 1, 2026. Only qualifying shifts and assessment submissions from
March 1 through October 1 enter completion. Scheduled shifts later in the report
period remain in the teaching-time report but cannot prematurely qualify a
student for completion-to-date.

**Teaching hours, scheduled availability, Learner Reach and student-continuity
totals retain the full reporting period.** The linked student-to-educator feedback
summary also retains its own matching full-period Submit Dates: it is already
aggregated and is not sliced by this student-assessment cutoff. Its source dates
remain visible in the evaluation-feedback section.

After sources are loaded, changing the cutoff or minimum shifts recalculates the
completion results from those inputs without downloading all sources again.
Old report downloads are cleared. Generate a fresh ZIP afterward. The cutoff is
not a guarantee that the loaded OASIS export contains every submission through
that date; use Load / refresh evaluation completeness after new OER uploads.

## Matching without an impossible confirmation task

In **PTS Matching -> Student names**, the normal correction list now contains
only unresolved names with a potential spelling/name discrepancy or an ambiguous
OASIS identity. Existing exact matches, saved confirmations and recognized
MD/PA/DO/class-designation normalization remain in place.

Names absent from the loaded OASIS records are shown separately under
**Students not found in OASIS (optional — no action required)**. You do not need
to select a different student, create a fictitious ID, or ignore a genuine
student just to generate a percentage.

The suggested-review rule is deliberately limited: it can flag small one-token
typos, initials and reordered name tokens. It does not identify every typo and
does not prove identity. For a known discrepancy that is not suggested, enable
**I know one of these students has a different OASIS spelling** in the optional
area and use the existing explicit name-only confirmation. No candidate is
preselected and no matching correction is saved automatically.

Student External ID remains internal to confirmed assessment links. Absent
students use normalized OPD names for in-memory denominator grouping only;
no synthetic OASIS ID or extra student catalog is uploaded. If a new export later
includes the same student's matching name and an assessment, refresh the check:
the identity and numerator can update automatically. A real spelling difference
may still need the existing name correction. Confirming an identity is not proof
that a target assessment exists.

The focused diagnostic panel is now **Review documented completion or missing
records (optional)**. It can show partial/no-assessment results as well as source
issues. Its correction shortcut includes only eligible names needing review,
not students who are simply absent. The rest of PTS remains table-free by default.

## Report and CSV behavior

Chair and individual reports use the heading **Documented assessment completion**
and display the cutoff and the selected minimum-shift rule. A numeric result with
an asterisk is provisional because of name review. The explanation states that
all eligible students remain in the denominator and only confirmed assessments
enter the numerator. Missing-record notices are nonblocking. Genuine source and
OPD conflict safeguards remain active.

The completion CSV preserves its previous columns and appends:

```
assessments_as_of
assessment_end_date
student_names_needing_review
students_without_oasis_name_record
clinical_students_without_assessment
hp_students_without_assessment
either_students_without_assessment
```

`assessment_end_date` is the earlier of report end and cutoff. The three
without-assessment counts are eligible students minus confirmed numerator members
for that form/union, not total missing forms. They remain blank when checks cannot
be verified. `assessment_status` indicates Calculated, Provisional, Not checked,
Not verified or No eligible students. The existing completion alert CSV and ZIP
notes explain record limitations; student names and IDs are not added to report
outputs. The source/identity detail stays in the password-protected interface.

An absent student's eligible membership is retained using the normalized OPD name
until a unique OASIS identity or saved correction is available. Truly different
OPD spellings may remain separate before correction; this is one reason a
provisional result can change after review. No absent names are silently deleted
from the denominator.

## Unchanged safeguards and features

- OER, PTS and PTS Matching password access; current encrypted GitHub catalogs.
- Column-minimized OASIS uploads, source snapshots and verified encryption.
- Academic Pediatrics-over-PSHCH Nursery priority before counting; other genuine
  clinical-area conflicts still block the teaching report.
- Four educational hours per distinct preceptor/date/AM-PM with any retained
  student, including weekends. Simultaneous students do not multiply hours.
- Saved minimum shifts, ignored student entries, preceptor usernames and
  student-name correction/removal remain intact.
- No app-wide shared cache of assessment data, no new source uploads or network
  writes during completion calculation, and no automatic overdue judgment.
- The chair and individual evaluation feedback still uses full question wording.

## Runtime files

| Action | Path |
|---|---|
| Add | schedule_app/services/assessment_progress.py |
| Replace | schedule_app/services/assessment_completion.py |
| Replace | schedule_app/services/student_name_review.py |
| Replace | schedule_app/services/assessment_diagnostics.py |
| Replace | schedule_app/sections/assessment_completion.py |
| Replace | schedule_app/sections/student_name_matches.py |
| Replace | schedule_app/sections/assessment_diagnostics.py |
| Replace | schedule_app/sections/pts_navigation.py |
| Replace | schedule_app/reports/assessment_completion.py |
| Replace | schedule_app/reports/teaching_export.py |

See DOCUMENTED_COMPLETION_MANIFEST.json for hashes and install actions.

## Validation and limitations

The offline regression suite completed in four batches: **1,127 passed and six
historical release-packaging checks skipped**, plus 139 passing subtests. This
includes **32 new tests** for missing records, provisional name review, source
failures, per-form/union counts, exact cutoff boundaries, state clearing, optional
corrections, privacy and preservation of teaching-hour calculations. Existing
expectations that explicitly required percentages to be suppressed for missing
student records were updated for the new requested policy; runtime failures were
not waived. Tests cover the full packaged test-file set.

Individual, provisional and chair Word examples were rendered and every page was
visually checked. Example documents use invented records, not a recalculation of
Fogel's actual results. No live GitHub repository, credentials or Streamlit
Cloud deployment was accessed or changed. The available uploads do not establish
a new actual completion percentage for any named preceptor.

Technical references for the date control and existing session lifecycle:
- https://docs.streamlit.io/develop/api-reference/widgets/st.date_input
- https://docs.streamlit.io/develop/concepts/architecture/widget-behavior
