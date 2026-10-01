> **Historical UI guide:** the two OASIS screens are now combined. Use `UPDATE_OASIS_COMBINED_WORKFLOW.md` for current workflow and installation instructions. Source schemas and calculation details below remain useful.

# OASIS Educator Reports — install and use

## Install the small update (recommended)

This adds a separate **OASIS Educator Reports** main-menu option. It does not replace the OASIS Evaluation Archive, OPD scheduling, teaching-effort reports, date presets, or Learner Reach code.

1. Extract `Schedule_App_OASIS_Educator_Reports_Update.zip`.
2. Replace `app_sch_2026.py` in your deployed repository, or add this entry to your existing `SECTIONS` dictionary when preserving a customized launcher:

   ```python
   "OASIS Educator Reports": "oasis_educator_reports",
   ```

3. Add these new files, keeping the folder paths:

   ```text
   schedule_app/sections/oasis_educator_reports.py
   schedule_app/services/oasis_educator_reports.py
   schedule_app/services/oasis_educator_usernames.py
   ```

4. Restart the app. Do not delete the existing `schedule_app` folder.

**Keep your existing settings.py, email mappings, requirements.txt, Streamlit Secrets, GitHub repository settings, and encryption key unchanged.** No new dependencies, credentials, app password, or archive migration is needed. The tests and this guide do not need to be deployed.

The complete modular-app ZIP is also supplied for a clean installation. The small update is safer for preserving deployed customizations.

## Normal workflow

1. Open **OASIS Educator Reports**.
2. Choose **Saved OASIS exports in GitHub**, select the saved export(s), and click **Read selected evaluations**. The app retrieves/decrypts them using the existing archive configuration. Refresh the saved-export list after archiving a new file in another section/session.
3. Choose the course and evaluation form. `*Clinical Teaching Eval` is selected by default when present. All submitted evaluations are included unless you enable a date range. An optional date filter uses your explicit choice of **Submit Date**, **Start Date**, or **End Date**, with both endpoints included. It does not silently assume your OPD teaching-date presets apply to evaluation submissions.
4. Correct any usernames under **Add / correct an educator username**.
5. Click **Create educator report CSV**, then **Download educator summary CSV**.

**Upload CSV for this report only** is an alternative. That path analyzes local files without archiving them; use the separate OASIS Evaluation Archive screen when you need an encrypted original saved.

Selecting multiple snapshots does not multiply evaluations. Identical responses for the same Course ID / Evaluation / Form Record / Question ID are counted once. If copies disagree, generation stops and identifies the source file, CSV record, form, and question. Choose the intended snapshot or correct the source—there is no silent first-copy/latest-copy choice.

The source list is read at one GitHub commit, so a changing archive cannot silently mix versions within one read. The originals are never rewritten by this feature.

## What the main CSV contains

`oasis_educator_summary.csv` has **one row per educator**, not one row per question.

| Column(s) | Meaning |
|---|---|
| `record_id` | Username before `@` in **Evaluator Email**, lowercased, or an explicitly entered username override. First column. |
| `educator_name` | Educator name from **Evaluator**. |
| `evaluator_email` | The available educator email address(es) from the selected export records; never fabricated. |
| `evaluation_count` | Distinct submitted **Course ID / Evaluation / Form Record** combinations for that educator. A form with 19 question rows counts as one evaluation. |
| `q587_mean`, `q587_n`, etc. | One mean and scored-response count per multiple-choice question, identified by its **Question ID**, not its changing **Question Number**. |
| `strengths_comments` | Nonblank **Answer text** for “Please indicate this educator's strengths,” numbered and combined. |
| `areas_for_improvement_comments` | Nonblank **Answer text** for “Areas for Improvement,” numbered and combined. |
| `email_missing` | `YES` when no valid source email is available for the educator in the selected data. Remains YES even after a manual username fixes record_id. |
| `record_id_source` | `evaluator_email` or `manual_username`. |

The supplied form has 17 numeric multiple-choice questions across its older/newer versions. The baseline CSV therefore has 42 columns: four identity/count columns, 17 mean/count pairs, two comment columns, and two identifier-status columns. Known questions absent from the chosen data keep a blank mean and zero response count. New numeric Question IDs append additional mean/count columns. A changed definition of an existing ID is flagged rather than silently combined.

### How means are calculated

The source is **Multiple Choice Value**, never Multiple Choice Order. Each question has its own denominator. Blank scores and non-scored labels such as N/A are excluded, not replaced by zero. Means are rounded to two decimal places. A question with no scored answers has a blank mean and `q<ID>_n = 0`.

**The duration question (ID 1286) is a category code.** For example, the source uses `3` for “1 week.” Its mean is included because it is a multiple-choice value, but it is not an average number of weeks and is not a teaching-quality rating. No composite average across unrelated questions is created.

`oasis_question_key.csv` maps each CSV column to its exact question wording and observed code/label pairs. The app shows the same information under **Question IDs, exact wording, and CSV columns**.

### How comments are combined

Comments are kept separately for strengths and improvement, in submission order. Nonblank comments are numbered and separated by blank lines in the same CSV cell. Internal text and line breaks are preserved; leading/trailing whitespace is trimmed. The app does not summarize, rewrite, or sentiment-score comments. Literal entries such as “n/a” remain. Repeated wording from different evaluations remains, while repeated copies of the same form/question do not duplicate it.

CSV quoting and UTF-8 BOM preserve commas, quotes, accents, and multiline comments. Spreadsheet-formula-like text is prefixed with an apostrophe for safety; numeric values remain numeric-looking text.

### How educators are matched

The source **Evaluator** is treated as the educator, per this workflow. Exported **Evaluator External ID**, **Evaluator Username**, and full email help link that educator across evaluations. A uniquely matching name may link a name-only record; ambiguous identifiers stop processing.

**Evaluator Username and Evaluator External ID are not automatically substituted for a missing email username.** The UI displays the exported username as a reference only, allowing you to confirm which username should become `record_id`. Different educators who would receive the same record_id are flagged and cannot be exported until resolved.

### Submitted evaluations only

An empty Submit Date is treated as an unsubmitted form and excluded with a visible count. An invalid nonblank Submit Date, missing required form/question identifiers, or a contradictory answer blocks the selected dataset rather than silently dropping records. Rows describe evaluations, not unique students, so two forms by the same student still count as two evaluations.

## Missing email / username workflow

1. The app displays an issue table and **OASIS_Username_Issues.csv**. No educator is silently dropped and the final CSV is blocked until every educator has a unique record_id.
2. Select the educator in **Add / correct an educator username**.
3. Type just the intended username, for example `jsmith`, not `jsmith@example.edu`.
4. Click **Apply username**. Leave **Save this username encrypted in GitHub for future reports** selected to persist it. Uncheck it for a session-only correction.
5. Create the CSV. The source email remains blank when absent, with `email_missing = YES` and `record_id_source = manual_username`.

The saved mapping uses the existing token/key and is stored separately as:

```text
opd_archive/oasis_educator_usernames.json.enc
```

No names or usernames are written in plaintext to GitHub by this feature. Every save is re-read and decrypted to verify the result. Concurrent edits cause a refresh/retry prompt instead of silently overwriting someone else's map. A damaged catalog or wrong key is never replaced with an empty catalog.

Use **Refresh saved usernames** to load another session's changes. A confirmation control also allows removing an override; this restores email-based identification and does not modify the source CSV. Removing a mapping changes the current catalog only; older encrypted Git revisions may remain in history. No original export, OPD, or date preset is deleted.

If GitHub cannot load the mapping, the app warns you. You may explicitly continue without saved mappings and use session-only corrections, rather than silently assuming existing overrides do not exist.

## Additional downloads

The main output is a directly downloadable CSV. An optional ZIP contains:

```text
oasis_educator_summary.csv
oasis_question_key.csv
oasis_educator_question_detail.csv
OASIS_Report_Notes.txt
```

The long-form detail file provides one row per educator/question with mean, scored-response count, and evaluation count. These are not new Word reports and are not added to the OPD chair or individual teaching-effort reports.

## Privacy and access

Structured Student, Student Email, Student Username, and Who Completed fields are not retained by the report calculation or exported. **Verbatim comments may themselves name a student, patient, or colleague; this is not automatic anonymization.** Review and handle the CSV as evaluation data.

CSV/ZIP downloads are unencrypted and are not automatically uploaded to GitHub. Do not commit them to the public app or archive repository. The username mapping and original archived exports remain encrypted.

No app password has been added. Anyone who can access this running app can use its existing archive and report features unless access is restricted elsewhere.

## Editing guide

- Buttons, source selection, previews, missing-username controls: `schedule_app/sections/oasis_educator_reports.py`
- Evaluation counts, question matching, averages, comments, CSV columns: `schedule_app/services/oasis_educator_reports.py`
- Encrypted username persistence: `schedule_app/services/oasis_educator_usernames.py`
- Original encrypted CSV archive behavior: existing `schedule_app/services/oasis_evaluations.py` (unchanged).

## Validation

See `TESTING_OASIS_EDUCATOR_REPORTS.md`. Tests use simulated GitHub/Streamlit; no live repository or deployment was accessed. A separate local check used the supplied OASIS CSV and did not guess its missing email username.

## Technical references

- GitHub Contents API (including replacement SHA/concurrency and Contents write permissions): https://docs.github.com/en/rest/repos/contents
- Streamlit widget/session behavior: https://docs.streamlit.io/develop/api-reference/caching-and-state/st.session_state
- Source schema and question wording: the OASIS evaluation CSV supplied in this conversation.
