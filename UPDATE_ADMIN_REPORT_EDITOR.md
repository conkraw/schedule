# Admin: edit report wording without editing Python

## Purpose and unchanged defaults

This release adds **Admin** to the main menu, immediately before OER and PTS.
Admin uses the same existing `[evaluation_access]` password as OER, PTS and PTS
Matching. It allows editing report instructions, selected headings and explanatory
paragraphs, and basic Word appearance directly in the app. Customizations are saved
separately and encrypted in the configured GitHub archive.

All existing wording and appearance remain the defaults until you save an edit.
Installing this update does not automatically create a customization catalog, alter
previously generated reports, change calculations, or modify OPDs/OASIS sources.
The new defaults file is a bundled JSON resource containing only the existing
public instructional text. You do not edit or upload JSON settings during normal use.

## Install the small update (recommended)

1. Extract **Schedule_App_Admin_Report_Editor_Update.zip** locally.
2. Replace your deployed **app_sch_2026.py** with the included launcher.
3. Merge the included **schedule_app/** folder into your existing app repository.
   Replace matching files and add the new files. **Do not delete the existing
   folder and do not upload only the ZIP.**
4. Be sure this new resource is uploaded as well as the Python files:

   ```text
   schedule_app/resources/report_wording_defaults.json
   ```

5. Restart Streamlit. Open **Admin** and unlock it using your existing OER/PTS
   password. The main entrypoint is still `app_sch_2026.py`.

The update contains 16 replacement Python files, five new Python modules and one
new defaults JSON resource. The full modular ZIP is an alternative for a clean
installation; the small update is preferable for preserving deployed customizations.

**Keep settings.py, requirements.txt, Streamlit Secrets, the existing encryption
key and password, username/name mappings, date presets, exclusions and minimum
shifts unchanged.** No new dependency, token, repository, source upload, or
re-encryption is needed. Nothing is installed into your live app until you deploy it.

If you maintain a customized launcher, insert `"Admin": "admin",` in its SECTIONS
mapping before OER/PTS, rather than overwriting your custom menu entries. You must
still install all the other runtime files listed below. Admin enforces its own
password gate even when imported by a customized launcher.

## Normal editing workflow

1. In **Admin**, choose **Report or report component**.
2. Choose **Part to edit**. Friendly labels identify the heading or explanation.
3. Edit **Report wording**. Use ordinary plain text and line breaks; no Python,
   JSON, HTML, or templating language knowledge is required.
4. Optionally click **Preview this edit (Word)**, then **Download wording preview
   (Word)**. This is a proof of the selected wording, not a complete calculated
   report. Previewing does not save the change.
5. Click **Save wording to GitHub**. A success message appears only after the
   encrypted save has been retrieved and verified.
6. Return to the normal report section and generate a new output. For PTS, click
   **Create teaching reports ZIP**. The chair and individual documents use the
   saved wording. Already downloaded files do not change retroactively.

Example: choose **PTS — individual preceptor report**, then **Additional opening
instructions**, and add:

> Please review this report and contact the clerkship team with any scheduling corrections.

Saving this adds the instruction to newly generated individual reports. The blank
original default means no additional opening paragraph. The included example Word
file is illustrative only; no example wording is pre-saved in your archive.

Edits are shared by all app users using this archive. Use **Reload saved wording
from GitHub** after another session has made changes. Changing the selected field
without saving does not save its draft. The restore controls return saved text to
the original defaults; they are not undo for an unsaved draft.

## What is editable

There are 92 text fields organized into 11 components:

| Component | Editable material |
|---|---|
| PTS — shared explanations | Common availability/educational-hours explanations, inclusion notes and optional extra report-ZIP instructions |
| PTS — individual preceptor report | Title, headings, header/footer text, selected explanations, optional opening/closing instructions |
| PTS — chair summary | Title, headings, header/footer text, selected explanations, optional opening/closing instructions |
| PTS — documented assessment completion | Section headings and calculation/interpretation explanations |
| PTS — learner feedback section | Section headings and explanatory instructions, not source questions or comments |
| QGenda download instructions | Document/screen instructions, four report headings and their step-by-step instructions, optional opening/closing text |
| OPD change report | Report title and optional opening/closing instructions |
| OPD weekly assignment summary | Report title and optional opening/closing instructions |
| Student schedules — instructions | Protected Self-Study Time note written into new student schedule workbooks |
| OPD workbook — instructions | CRTS instruction note written into new OPD workbooks |
| Power Automate workbook — Definitions | Definitions-sheet explanations, not the machine-readable assignment table |

Shared explanations affect every report using them. To add wording only for the
chair or only for individual preceptors, use that specific component's opening or
closing instructions instead. The rest of the wording is populated from the exact
existing defaults; there is no new editorial rewrite.

This is a **wording and basic appearance editor**, not a drag-and-drop report designer.
New computed measures, new report types, table rearrangement or new data columns
still require a code change. The downloaded Word files remain normal editable
paragraphs and tables; charts remain embedded images. You can continue making
one-off changes in Word without changing the saved app settings.

## Automatic values and protected source data

A few fields contain automatic placeholders such as `{minimum_shifts}`,
`{shift_word}`, `{start_date}` or `{end_date}`. The editor shows the exact tokens
for the chosen field. Keep them intact so the actual selected values appear when
the report is generated. Removing them, adding unknown tokens or using code-like
format expressions is rejected. This is simple text substitution, not code execution.

Dates, calculated values, eligibility/identity rules, validation warnings, source
provenance and report filenames remain controlled by the existing app. The editor
does not change CSV column names, record_id, numeric scores or completion totals.
OASIS question wording and evaluator comments are source data: Admin cannot rewrite
them. Changing a calculation's explanation does not change its formula; keep any
new explanatory wording consistent with the actual calculation.

Required headings cannot be blank. Optional notes and additional opening/closing
paragraphs may be cleared where permitted. Each field allows up to 16,000 characters;
all encrypted presentation settings are limited to 256 KiB of JSON before encryption.

## Optional Word appearance

For the individual report, chair summary, QGenda instructions, OPD change report and
weekly assignment summary, **Word appearance (optional)** offers:

- Font: Keep existing, Calibri, Aptos, Arial, Cambria or Times New Roman.
- Body, table, title and note sizes within bounded ranges.

**Keep existing** preserves the present styles. You must click **Save appearance
to GitHub** to apply appearance edits. Header/footer geometry, automatic page-number
fields, table widths, calculation formats and page dimensions stay unchanged.
Larger fonts or longer notes can create additional pages. Generate the usual report
to check its full layout. The small wording preview is not a pagination guarantee.
Assessment and feedback sections inherit the individual/chair document appearance;
they do not have separate global style controls. Excel formatting is unchanged.

## Restore controls

- **Original wording / restore this part** shows the original text and restores
  only that selected part after confirmation.
- **Restore this report or all report settings** restores the selected component,
  or all wording and appearance, after explicit confirmation.

Restoring removes current customizations rather than deleting source records.
Restoring one component does not reset shared PTS explanations unless you select
that shared component or restore everything. Saved usernames, OPDs, evaluation
summaries, dates and other settings are not edited. Earlier encrypted revisions
can remain in Git history; restoring defaults does not erase history.

## Where changes are stored

```text
opd_archive/report_wording.json.enc
```

The base folder follows your existing `[opd_archive]` configuration, even if it is
not named opd_archive. The service reuses the existing token and encryption key.
Only explicitly customized text and appearance are stored in this new catalog.
Files remain encrypted in GitHub. You do not need to create the folder or catalog
manually. No additional CSV, Word template upload or auxiliary source file is stored.

The existing GitHub revision checks prevent a stale editing session from overwriting
another session's change. Reload, review and save again after a competing edit.
The encrypted file is read back at the save commit before success is reported.
A wrong key, corrupt catalog or failed read is not treated as empty settings.

## Password, access and privacy

Admin uses the same shared authenticated session as OER, PTS and PTS Matching.
It is not a separate administrator role: anyone with that password can edit these
shared settings. **Lock OER / PTS / Admin** locks all protected sections and clears
protected cached data and Admin previews. Existing inactivity and maximum-session
limits still apply. Installing the update requires unlocking once again.

Other scheduling sections remain password-free. Some editable text appears in
those public sections and downloaded instructions. **Do not place passwords,
tokens, student identities or evaluation comments into report instructions.**
Encryption of the saved settings does not make the rendered instructions private.
The authenticated editor never executes inserted Python, HTML or expressions.

## Fresh reports and performance

Saved edits clear affected cached downloads in the current session, but retain
loaded OPD and assessment inputs. PTS verifies the current wording again when
Create is pressed. One validated snapshot is used throughout the entire ZIP, not
one GitHub read per preceptor. A changed snapshot prevents stale ZIP reuse.
There is no process-wide cache of custom text or sensitive report content; only
the public bundled defaults are cached globally.

For the general scheduling/instruction screens, wording is loaded once for each
screen render so newly generated notes use the current settings. If no archive is
configured, those general screens explicitly use defaults as before. If an archive
is configured but its wording cannot be read, the affected output is withheld with
an error rather than silently using stale or default text. Retry after restoring
connectivity. The original sources are not changed by a wording error.

To use an updated Protected Self-Study note in student schedules, rebuild the
master student schedule first; splitting an older master copies the note that is
already in that master. Changing the CRTS note affects newly generated OPDs, not
archived originals. Custom Word instructions similarly affect new builds only.

## Runtime files

| Action | Path |
|---|---|
| Replace | `app_sch_2026.py` |
| Replace | `schedule_app/reports/assessment_completion.py` |
| Replace | `schedule_app/reports/chair_summary.py` |
| Replace | `schedule_app/reports/individual_teaching.py` |
| Replace | `schedule_app/reports/preceptor_evaluations.py` |
| Add | `schedule_app/reports/qgenda_instructions.py` |
| Add | `schedule_app/reports/report_appearance.py` |
| Replace | `schedule_app/reports/teaching_export.py` |
| Add | `schedule_app/sections/admin.py` |
| Replace | `schedule_app/sections/create_individual_schedules.py` |
| Replace | `schedule_app/sections/create_student_schedule.py` |
| Replace | `schedule_app/sections/format_opd_summary.py` |
| Replace | `schedule_app/sections/instructions.py` |
| Replace | `schedule_app/sections/opd_check.py` |
| Replace | `schedule_app/sections/preceptor_teaching_summary.py` |
| Add | `schedule_app/sections/report_wording_controls.py` |
| Replace | `schedule_app/services/evaluation_access.py` |
| Replace | `schedule_app/services/opd_workbooks.py` |
| Replace | `schedule_app/services/primary_preceptors.py` |
| Add | `schedule_app/services/report_wording.py` |
| Replace | `schedule_app/services/student_schedules.py` |
| Add | `schedule_app/resources/report_wording_defaults.json` |

See ADMIN_REPORT_EDITOR_MANIFEST.json for SHA-256 hashes. The resource JSON is part
of the app and must accompany the Python modules. The update does not contain a
filled-in secrets file, real student/evaluation records or fonts.

## Verification and technical references

See TESTING_ADMIN_REPORT_EDITOR.md for the test results and limits. GitHub and
Streamlit interactions were simulated. Your live deployment, repository and
stored reports have not been changed.

- Streamlit forms: https://docs.streamlit.io/develop/concepts/architecture/forms
- Streamlit session state: https://docs.streamlit.io/develop/api-reference/caching-and-state/st.session_state
- GitHub Contents API and revision-aware writes: https://docs.github.com/en/rest/repos/contents
- python-docx: https://python-docx.readthedocs.io/en/latest/
