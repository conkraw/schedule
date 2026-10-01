## Latest correction — student-count consistency (October 1, 2026)

See [UPDATE_PTS_STUDENT_COUNT_CONSISTENCY.md](UPDATE_PTS_STUDENT_COUNT_CONSISTENCY.md).
Continuity and assessment eligibility now share reconciled identities and the assessment cutoff.
Teaching hours remain on the full report dates. After installing, refresh both OPDs and evaluation completeness.
This guide supersedes historical notes stating continuity stays full-period when completion is included.

> **Latest update: documented assessment completion.** Start with
> `UPDATE_DOCUMENTED_ASSESSMENT_COMPLETION.md`. Missing OASIS student records now
> remain in the denominator without mandatory confirmation; completion has its
> own cutoff. Older update guides below are historical and are superseded where
> they required a match for every eligible student.

# Pediatric Clerkship Schedule App

A modular Streamlit app. Run `streamlit run app_sch_2026.py` with the entire
`schedule_app/` folder beside the launcher. Keep the existing requirements and
Streamlit Secrets. Never commit secrets or downloaded unencrypted reports.

## Current update

Read **UPDATE_PTS_SIMPLIFIED.md** for this release's installation steps.

**PTS** is the reporting page, without routine preceptor/student tables.
**PTS Matching** is the separate password-protected matching/ignore page, with a
dropdown for preceptor usernames, student names, and ignored student entries.
**OER** and **PTS** remain the final two sidebar choices; PTS Matching precedes
them. All three use the existing `[evaluation_access] password`.

The report ZIP includes the existing chair and individual Word reports, teaching
CSVs, applicable diagnostics, and clinical-experience charts. Shared tables are
calculated once per build, progress is shown, and unchanged completed results
can be reused within the protected session after existing freshness checks.

## Current calculation and storage rules

- Total scheduled availability includes recorded AM/PM shifts, including weekends,
  at four hours each. Educational hours count a half-day once when at least one
  retained student is assigned; simultaneous students do not multiply time.
- Learner Reach is educational hours / scheduled availability. Academic Pediatrics
  combines HOPE_DRIVE, ETOWN, and NYES and takes priority over concurrent PSHCH
  Nursery listings. Other clinical-area conflicts still block reports.
- Only preceptors and work types with student assignments are listed. Included
  preceptors' blank shifts remain in the availability denominator.
- Saved custom reporting dates, unique-student counts, three-distinct-day
  continuity, and the persisted adjustable minimum-shifts threshold are retained.
- OER reduces source columns before encryption and saves the existing summaries.
  PTS uses verified username/name links and optional linked OASIS feedback and
  assessment completion. Missing records are not fabricated as completed forms.
- Student-name matching ignores recognized program/class designations but does not
  fuzzy-match typos. Explicit corrections and ignored labels are encrypted and
  reversible. Names-only matching controls do not require users to type IDs.

## Editing and installation

Use **EDITING_GUIDE.md** for module locations. `schedule_app/settings.py` contains
your editable mappings. Keep your deployed custom entries when using a complete
package. The small update is safer when updating an existing installation.

Historical `UPDATE_*.md` and `TESTING_*.md` files describe earlier releases and are
retained for reference. This README and UPDATE_PTS_SIMPLIFIED.md supersede their
older screen locations and current-release installation directions. In particular,
old names such as Evaluation Records/OASIS Evaluations refer to OER; older
student-weighted hours and app-password-free evaluation guidance are not current.

See TESTING_PTS_SIMPLIFIED.md for the latest local verification and its limits.


## Latest reporting update

See [Feedback, shifts and assessed-student totals](UPDATE_FEEDBACK_SHIFTS_ASSESSMENT_TOTALS.md) for the current
feedback default, retired three-day display, and distinct assessed-student totals.
