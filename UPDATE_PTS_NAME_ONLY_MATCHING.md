> **Later update:** designation-only differences now match automatically. See `UPDATE_PTS_AUTOMATIC_NAME_MATCHING.md`. The old statement below that `(MD)` requires confirmation no longer applies.

# PTS: confirm student names, not external IDs

## What changed

The student-matching screen now asks you to match **OPD student name → OASIS student name**.
The OASIS dropdown, proposed match, and saved-match table display names only. There is no
manual Student External ID entry field and no separate ID-verification step.

You still confirm that the two names refer to the same student. The app automatically
uses the existing record behind the selected OASIS name when calculating assessment
completion. The encrypted matching catalog and its internal identifiers are unchanged.
This is a user-interface correction, not a change to who is eligible or how percentages
are calculated.

## Install the small update (recommended)

Extract `Schedule_App_PTS_Name_Only_Matching_Update.zip` and merge its `schedule_app`
folder into the existing repository. Replace all three matching files:

```text
schedule_app/sections/student_name_matches.py
schedule_app/services/student_name_review.py
schedule_app/services/assessment_completion.py
```

Do not delete the existing folder. Do not upload only the ZIP as an application file.
Restart Streamlit and unlock PTS. Load your OPDs and selected reporting period, then
click **Load / refresh evaluation completeness** if its inputs are not already loaded.

Keep `app_sch_2026.py`, `settings.py`, requirements, OER/PTS password, Streamlit Secrets,
encryption key, preceptor usernames, saved minimum shifts, and date presets unchanged.
No source uploads or re-encryption are needed. Existing student matches do not need to
be entered again. The small update does not replace any settings or report files.

## Normal use

1. Open **Resolve OPD student names to OASIS** when its warning says matches are needed.
2. Select a name under **OPD students needing an OASIS match**. This list contains only
   unresolved names from the current reporting selection.
3. Select the corresponding name under **Matching OASIS student name**. The control
   starts unselected; no similar-name match is guessed.
4. Review the displayed **Name match: [OPD name] → [OASIS name]**.
5. Check **I confirm these two names refer to the same student** and click
   **Save student match to GitHub**.

A verified encrypted save removes the resolved name from the queue, recalculates the
completion results on the next app rerun, and clears old downloadable reports. Generate
fresh reports to include the correction. A failed save leaves the flag in place.

## Already saved or exact matches

Existing confirmed matches and unambiguous exact normalized matches do not require
another confirmation. Normalization is unchanged: capitalization, whitespace, comma
spacing, and the recognized trailing `; MD2028`-style class suffix are handled as before.
Other differences such as `(MD)` or a different name order may still require a one-time
name match. This update does not add fuzzy matching or silently combine people.

Use **Review or correct saved student matches (optional)** to change or remove an old
selection. Its table now shows the OPD name and available corresponding OASIS name(s),
not the internal ID. A previously saved match remains valid if its OASIS name is absent
from the currently loaded exports; the table notes that absence. Existing legacy
matches created by entering a verified ID are retained.

## Ambiguous or missing OASIS students

If OASIS contains more than one distinct student record under the exact same name,
the dropdown displays that name once with **duplicate name — source review needed**.
The screen explains the problem and does not save an arbitrary record based only on
the name. Review/correct the source records in OER. The matching screen no longer offers
a manual-ID workaround.

If the correct name is not listed, upload the relevant student-assessment CSV through
**OER → Evaluations of students** and refresh evaluation completeness. Do not select a
different person just to dismiss the flag. Underlying missing/contradictory assessment
metadata may still need source correction independently of name matching.

Unresolved students stay in the eligible denominator. Affected percentages remain
**Not verified**, while other valid results and teaching reports are still available.
A successful name match does not imply that an assessment was completed.

## Storage and privacy

The existing `student_assessment_id_links.json.enc` catalog is reused in the configured
GitHub archive folder. The selected OASIS record's ID remains internal because it
provides the stable join to assessment forms; it is not a username created from a name.
The source workbooks/CSVs are not edited. The report tables still omit student names
and IDs, and OER/PTS password protection is unchanged.

The catalog schema, encryption implementation, revision checks, and verified writes
are unchanged. The screen resets obsolete ID-entry controls once after this update,
without deleting stored matches or the saved minimum-shift threshold.

## Verification

Compilation succeeded for all runtime modules. The complete local suite was run in
two batches: **948 passed, 5 skipped**, plus 139 passing subtests. The skips are
existing environment or historical-baseline checks; no failed check was waived.
Seventeen additional name-only tests cover visible labels, no manual ID entry,
ambiguous names, old saved links, encrypted saves, recalculation, and threshold
preservation. Existing UI tests were updated for the intentionally removed manual-ID
control and the new name-based warning wording.

The completion-engine changes are limited to user-facing text; an AST comparison
confirmed that its non-text structure is unchanged. No report builder, setting,
archive service, or stored data was modified. GitHub/Streamlit were simulated; there
was no live repository or deployment access. This release does not include newly
rendered Word previews because it does not change document layouts.

Technical reference for the existing widget/session pattern:
https://docs.streamlit.io/develop/api-reference/caching-and-state/st.session_state
