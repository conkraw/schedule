# PTS: automatically ignore student program designations

## What changed

PTS no longer asks you to confirm a match merely because an OPD name ends in
`(MD)`, `(PA)`, or `(DO)`. It also ignores recognized class-year suffixes such as
`; MD2028`, `(MD2028)`, or `(PA 2028)`, as well as case and extra spacing.

The remaining name must match exactly and identify one OASIS student record.
For example:

| OPD | OASIS | Result |
|---|---|---|
| Adlassnig, Sarah (MD) | Adlassnig, Sarah | Automatic; no confirmation |
| Example, Jordan (PA) | Example, Jordan | Automatic; no confirmation |
| Example, Jordan; MD2028 | Example, Jordan | Automatic; no confirmation |
| Example, Jordann (MD) | Example, Jordan | Name difference; confirm once |

Existing confirmed corrections still take priority for the exact OPD name under
which they were saved. Nothing is fuzzy-matched, and different people are not
chosen merely because their names look similar.

## Install the small update (recommended)

1. Extract `Schedule_App_PTS_Automatic_Name_Matching_Update.zip`.
2. Merge its `schedule_app/` folder into your existing app repository. Replace
   the three matching files and add the new helper below. Do not delete the
   existing folder and do not upload only the ZIP as a runtime file.
3. Restart Streamlit. Unlock **PTS** with your existing password.
4. Load your OPDs and reporting dates as usual, then click **Load / refresh
   evaluation completeness** once. This clears earlier matching results and
   rebuilds them using the designation-aware rule.
5. Generate fresh teaching reports to use the recalculated results.

| Action | File |
|---|---|
| Replace | `schedule_app/services/assessment_completion.py` |
| Replace | `schedule_app/services/student_name_review.py` |
| Replace | `schedule_app/sections/student_name_matches.py` |
| Add | `schedule_app/services/student_name_matching.py` |

Keep `app_sch_2026.py`, `schedule_app/settings.py`, `requirements.txt`, Streamlit
Secrets, passwords, encryption key, date presets and saved minimum shifts
unchanged. No new dependency, source upload, migration or re-encryption is needed.
The full modular ZIP is an alternative for a clean installation; the small
update avoids replacing your deployed custom settings.

## What still needs confirmation

Only genuinely unresolved names appear in the routine correction dropdown:
spelling differences, different name order, initials versus a full name, a
missing OASIS record, or an ambiguous name shared by different records.

An unknown parenthetical note, a middle name, a hyphenated family name, or a
family-name suffix such as Jr. is not removed. The rule strips only the
recognized program/class labels at the end of a name. It is applied to both
OPD and OASIS names, not only to the OPD side.

If two distinct OASIS students share the same remaining name, their designations
are not used to guess which person is intended. The name remains flagged for
source review. The name-only selector still refuses an ambiguous choice.

Automatic matches need no confirmation checkbox or saved-link entry. They are
recomputed from the existing source data in later sessions without writing an
extra mapping to GitHub. Actual spelling corrections still use the existing
explicit, encrypted save workflow and disappear from the missing-name list
only after the save is verified.

## Existing saved links are preserved

The encrypted `student_assessment_id_links.json.enc` catalog format, keys and
save/load/remove implementation are unchanged. This matters because prior
saved keys may themselves end in `(MD)` or `(PA)`. They remain valid and editable;
this update does not silently re-key, delete, overwrite or merge them.

The internal matching helper is separate from the persistent catalog key.
You still select names only for a manual correction; the selected OASIS record
supplies its existing ID behind the scenes.

## Calculations and privacy

The same matching decision is used by both the missing-name alert and the
assessment-completion calculation. Once a designation-only difference resolves,
that student's retained shifts can be joined to their assessment records.
Different spellings of a designation for the same matched student do not create
extra eligible students or duplicate date/AM-PM shifts.

The saved minimum-shift setting, date filtering, eligible-group numerator rules,
and outpatient-over-nursery exception are unchanged. A missing assessment is
still different from a missing identity. Resolving a name does not assert that
an assessment exists. Unresolved eligible students are not dropped to improve
a percentage.

No teaching-time formula, Learner Reach rule, preceptor username, unique-days
reporting code, report layout, source file or archive content is modified.
The OER/PTS password boundary and column-minimization rules remain in place.
No names/IDs are added to the completion report tables or CSV by this update.

## Verification

- Python compilation passed for all runtime modules.
- The 49 new designation-specific tests passed.
- Full local regression coverage ran in two batches: **996 passed, 6 skipped**,
  with 139 additional passing subtests. Existing warnings occurred in unchanged
  scheduling modules. No failed test was waived.
- Tests cover designation-only auto-matches, genuine typos, multiple IDs sharing
  a name, old encrypted links, missing-only UI controls, selected thresholds,
  date boundaries, deduplication, and source/mapping preservation.
- One earlier test was changed because it explicitly required manual confirmation
  for `(MD)`, the exact behavior this release intentionally replaces.
- Launcher, settings, requirements and the encrypted student-link storage service
  were checked byte-for-byte against the previous release and are unchanged.

Tests use invented records and simulated GitHub/Streamlit. No live repository or
deployment was accessed or changed. No document preview is included because this
update changes matching behavior, not the Word report layouts.
