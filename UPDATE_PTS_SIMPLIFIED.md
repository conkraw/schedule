# Simplified PTS and a separate PTS Matching section

PTS is now the reporting screen. Preceptor username corrections, student-name
matching, and ignored student entries are in **PTS Matching**, a separate
password-protected main-menu choice. Routine tables are hidden unless you
explicitly request diagnostics or previews. The Word and CSV report contents
and calculation rules are preserved.

## Install the small update (recommended)

1. Extract `Schedule_App_PTS_Simplified_Update.zip` on your computer.
2. Replace **`app_sch_2026.py`** and merge the included **`schedule_app/`** folder
   into the matching folder in your existing repository. Replace matching files
   and add the four new modules. **Do not delete the existing folder and do not
   upload only the ZIP as a runtime file.** The new launcher needs every module
   in the existing app, not just the files in this update.
3. Restart Streamlit. Keep its entrypoint set to **`app_sch_2026.py`**.

Keep `schedule_app/settings.py`, requirements, all custom mappings, Streamlit
Secrets, your existing password and encryption key, date presets, and the saved
minimum-shift setting unchanged. No new dependency, credential, GitHub permission,
archive migration, or source re-upload is needed.

The small update has **14 runtime files: 10 replacements and 4 additions**.
The complete modular ZIP is an alternative for a clean installation; use the
small update to preserve deployed customizations. The manifest lists file paths,
actions, and SHA-256 checksums.

## Where the controls are now

The final menu entries are:

```
PTS Matching
OER
PTS
```

OER and PTS remain the final two choices. Other menu choices retain their order.

### PTS — create reports

Choose/load your reporting period, load the archived OPDs, and use the existing
report controls. The default view shows three totals (scheduled availability,
educational hours, Learner Reach), brief notices, and the report buttons rather
than long preceptor/student tables.

The linked OASIS-summary selector and assessment-completion settings/check stay
on this page because they determine report content. The user-editable minimum
shifts and its encrypted saved default remain unchanged.

Use **Create teaching reports ZIP**, then the ZIP or chair-only Word download.
The ZIP still includes the chair report, individual reports, teaching CSVs,
source/diagnostic files where applicable, and clinical-experience charts. Tables
inside the Word/CSV files are not removed.

- **Show diagnostic tables and report previews (optional)** is unchecked by
  default. It reveals the existing provider/source/assessment detail when needed.
- **Preview clinical-experience pie charts (optional)** is also unchecked by
  default. The charts are still in the chair report and ZIP regardless.
- Serious OPD conflicts still stop report generation. A short warning, optional
  details control, and the existing conflict CSV identify the source cells.
- Missing username/student matches or evaluations still produce short alerts
  rather than disappearing. Use **Open PTS Matching** to correct names. Missing
  assessment records remain nonblocking as before; unknown percentages are not
  invented.

### PTS Matching — fix names only when needed

Select **What needs updating?**:

- **Preceptor usernames:** the routine selector includes only teaching preceptors
  without a saved username. Verified saves remove names from that queue. Review,
  correction/removal of existing links, and OASIS username reference tables remain
  explicitly optional. Saved links are usable without having an OASIS summary
  selected first.
- **Student names:** load/refresh evaluation completeness, then use the existing
  unresolved-only name selector. Confirm the OPD name against the OASIS name;
  there is no routine ID-entry field. Recognized MD/PA/DO/class designations
  continue to match automatically when the remaining name is unambiguous.
- **Ignored student entries:** use the existing encrypted, reversible ignore and
  restore controls for labels such as midcycle feedback. The preceptor's
  scheduled availability remains; the ignored entry does not earn student credit.

The reporting dates and the loaded OPD snapshot are shared between PTS and PTS
Matching in the current session. You can load sources from either page. Switching
between them does not itself redownload the OPDs. Existing encrypted catalogs
and save/verification/concurrent-edit safeguards are unchanged.

After a preceptor username/summary link changes, return to PTS and refresh
**evaluation completeness** when prompted. Student-name saves recalculate their
matching inputs. Ignoring/restoring an entry keeps the existing rescan behavior
and requires the completeness refresh afterward. Generate a new ZIP after these
changes; a previously downloaded file cannot be changed retroactively.

## Password and data handling

PTS Matching checks the same **`[evaluation_access] password`** before displaying
controls, reading records, or changing links. Unlocking any of OER, PTS, or PTS
Matching uses the same authenticated session. **Lock OER / PTS** also locks PTS
Matching and clears its session state, including loaded names and report bytes.
The existing idle and maximum-session expiry rules remain.

No new password or encryption key is needed. No plaintext records are saved to
GitHub by this change. Other scheduling sections retain their existing access
behavior. This remains the existing shared-password control, not a new
institutional single-sign-on service.

## ZIP performance improvements

The older exporter recalculated complete teaching-time tables for every
individual report. The new exporter computes the three shared tables once per
ZIP, indexes them by preceptor, and reuses the same results. Report-generation
validation remains enabled. The chair summary data is also reused for its charts
and Word document within the same build.

A progress bar now describes the current stage and the number of individual
reports completed. It does not print preceptor/student names.

Clicking Create again without changing the loaded data, period, selections,
matching results, or settings reuses the existing ZIP in that authenticated
session after the existing GitHub freshness checks. A normal page interaction
or download does not rebuild it. Date, source, exclusion, matching, threshold,
or included-evaluation changes invalidate old outputs as before.

Reuse is confined to a single build or the current protected session. There is
no process-global shared cache of student data or evaluation comments. Chair
bytes are extracted from the ZIP once, and the page no longer extracts/displays
all chart images during every rerun. The underlying OPD and OASIS reads remain
snapshot-based; refresh their respective sources to include newer uploads.

The first build still takes time to write all documents and charts, and GitHub
connection speed still affects reads/checks. This is not background processing.
A failed build never leaves a partial ZIP available as a completed report.

## Runtime files

| Action | File |
|---|---|
| Replace | `app_sch_2026.py` |
| Replace | `schedule_app/services/evaluation_access.py` |
| Replace | `schedule_app/sections/preceptor_teaching_summary.py` |
| Replace | `schedule_app/sections/preceptor_oasis_links.py` |
| Replace | `schedule_app/sections/assessment_completion.py` |
| Replace | `schedule_app/sections/student_name_matches.py` |
| Replace | `schedule_app/sections/reporting_date_controls.py` |
| Replace | `schedule_app/reports/individual_teaching.py` |
| Replace | `schedule_app/reports/chair_summary.py` |
| Replace | `schedule_app/reports/teaching_export.py` |
| Add | `schedule_app/sections/pts_workspace.py` |
| Add | `schedule_app/sections/pts_navigation.py` |
| Add | `schedule_app/sections/pts_matching.py` |
| Add | `schedule_app/reports/teaching_batch.py` |

A customized launcher must import/call `preserve_pts_preferences()` as in the
supplied launcher and add `"PTS Matching": "pts_matching"` before OER and PTS.
Do not remove the other existing section entries.

## Verification

See `TESTING_PTS_SIMPLIFIED.md` in the complete package. The local suite completed
with **1,070 passed, 6 skipped, and 139 passing subtests**. Thirty-five new tests
cover the protected matching page, default table-free PTS, data reuse, stale-output
clearing, progress, errors, and unchanged document content. Existing UI tests were
updated to use the deliberately relocated controls and opt-in preview switches.

A synthetic 60-preceptor warm benchmark (three timed builds per version, no
GitHub/network time) had median build times of **4.67 seconds before and 2.73
seconds after**. This is a local example, not a Streamlit Cloud performance
promise. The same scenario's 61 Word documents, 5 CSVs, one chart PNG, and notes
matched the old version, excluding Word core-document timestamps. A chair report
and an individual report were rendered and visually checked.

GitHub and Streamlit interactions were simulated. The live repository,
credentials, and deployment have not been accessed or changed.

## Technical references

- Streamlit widget/session lifecycle:
  https://docs.streamlit.io/develop/api-reference/caching-and-state/st.session_state
- Streamlit caching scope and behavior:
  https://docs.streamlit.io/develop/concepts/architecture/caching

These references explain session/widget reuse; they are not sources for your
local clinical-hour or assessment-completion definitions.
