# Validation: PTS ignored student entries

## Local results

- Runtime Python compilation completed successfully.
- Full test suite: **1,035 passed, 6 skipped**, with **139 passing subtests**.
- New exclusion-specific test module: **39 passed**; no new test is skipped.
- The six existing skips are historical/local packaging-baseline checks whose
  comparison sources are not bundled or are no longer applicable.
- Existing warnings came from earlier code/dependencies; no failing assertion
  was suppressed.

All GitHub requests and Streamlit UI actions used the package's existing test
doubles. No live repository, credentials or deployment were accessed or changed.
Real Streamlit is not installed in this container, so browser/UI-server behavior
was not tested in a live deployment.

## New coverage

- Empty catalog loads without a write; encrypted create/read/restore round-trips.
- Duplicate/case-spacing equivalent exclusions avoid unnecessary writes.
- Current catalog revision checks reject competing save/restore attempts.
- Corrupt/wrongly encrypted stored contents are not treated as an empty list.
- Exact matching only: no substring/fuzzy matching or unintended designation removal.
- Invalid/empty/control-character entries do not produce GitHub writes.
- Ignored entries are removed before student-shift, teaching-time and continuity
  calculation; valid provider availability remains, including weekend shifts.
- A real student sharing the same half-day is preserved when a note is ignored.
- A provider with only ignored students retains internal availability but has
  no published teaching report under the existing contributors-only rules.
- Assessment replay uses the same verified ignore list; old scans or completion
  inputs cannot be combined with a different exclusion fingerprint.
- Name-match alerts, eligible denominators and form-specific numerators use the
  same retained students and the saved minimum-shift rule.
- Original OPD bytes, OASIS records and saved identity links remain unchanged.
- Restore reproduces original source counts and preserves existing name matches.
- Academic Pediatrics-over-PSHCH Nursery priority remains valid, including a
  clinic that becomes unassigned after a note is removed.
- Other cross-work-type conflicts still block reports even when a student entry
  is ignored; clinical availability conflicts are not hidden.
- Password gate prevents ignored-list reads from a locked UI; locking clears the
  session-only inventory and does not delete the GitHub catalog.
- Confirmation is required; failed saves do not dismiss names or clear valid scans.
- Full simulated PTS navigation: load, save, automatic recount, generate,
  restore, automatic recount and regenerate. Standard-year and custom dates remain.
- A concurrent exclusion change blocks publishing cached report counts.
- CSV/Word generation uses corrected values. Excluded student names and IDs are
  not written to report files; Report_Notes contains only a source-listing count.

The initially introduced extra PTS rerun was removed by separating catalog load
from panel rendering. This preserves the existing one-run scan/report workflow
and populates the name dropdown immediately after scanning. All original tests
then passed without removing or weakening their assertions.

## Preserved files

Byte-for-byte comparison against the supplied latest modular package confirmed:

- `app_sch_2026.py`
- `schedule_app/settings.py`
- `requirements.txt`
- `schedule_app/services/evaluation_access.py`
- `schedule_app/services/student_assessment_links.py`
- `schedule_app/services/student_name_matching.py`

All Word layout modules are unchanged; this update changes input filtering and
a text-only source audit note, not page layouts. No new Word preview is supplied.
