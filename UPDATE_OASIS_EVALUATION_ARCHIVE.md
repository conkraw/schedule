# OASIS Evaluation Archive — update instructions

This update adds **OASIS Evaluation Archive** to the main sidebar menu. It saves
the original uploaded OASIS CSV using your existing OPD archive repository, branch,
GitHub token, and Fernet encryption key. The file is verified by downloading and
decrypting it after the save, then comparing the original bytes.

## Install the small update (recommended)

1. Back up your currently working code. Extract `Schedule_App_OASIS_Archive_Update.zip`.
2. Replace `app_sch_2026.py` in your app-code repository. This update **does change
   the launcher** because it adds a new menu option.
3. Add these two files, keeping their folders:

   ```text
   schedule_app/sections/oasis_evaluation_archive.py
   schedule_app/services/oasis_evaluations.py
   ```

   Merge into the existing `schedule_app` folder. **Do not delete that folder or
   upload the ZIP itself as a substitute for the Python files.**
4. Keep `schedule_app/settings.py`, your email/name mappings, `requirements.txt`,
   and every existing report/OPD/date-preset module unchanged.
5. Keep your existing Streamlit entrypoint (`app_sch_2026.py`) and Streamlit Secrets.
   Restart the app and open **OASIS Evaluation Archive**.

No new dependency, encryption key, repository, token, or app password is required.
Do not generate a new encryption key. Keep its existing secure backup.

### Existing launcher has local customizations?

Keep those customizations and add just this entry to the launcher's `SECTIONS`
dictionary, directly after the existing OPD Archive entry:

```python
"OASIS Evaluation Archive": "oasis_evaluation_archive",
```

Keep the accompanying `schedule_app` folder alongside the launcher. This avoids
removing any import-path fix or extra menu option you added in your deployed copy.

## Upload and automatic save

Open **OASIS Evaluation Archive**, then upload `oasis_eval_export.csv` (or the same
export under another filename). Saving starts automatically after file validation.
Wait for **OASIS export archived and verified**. If saving or verification fails,
the app does not claim success; use **Retry OASIS encrypted save**.

The entire original CSV is encrypted, not just selected columns. Quoting, text,
line breaks, encoding/BOM, response values, and all columns are retained. The app
shows the response-row count, column count, file size, and course-date coverage;
it does not display learner names, answers, or comments. A response row is a
question-level CSV row, not a unique evaluation.

The parser accepts the supplied OASIS export's headers, including Course ID,
Start Date, End Date, Evaluator, Evaluation, Form Record, Question ID and Question.
UTF-8 (with/without BOM), BOM-marked UTF-16, and Windows-1252 are supported. It
rejects broken CSV structure instead of dropping rows. Limit: 10 MiB original
file size, 250,000 response rows, and 2 MiB of text in any one CSV field.

Blank, unreadable, or reversed course dates do not discard the export: the original
is saved under an **undated** identifier and the UI explains the date issue.

## Storage destination and retention

With the default existing folder setting, files go here:

```text
opd_archive/
    OPD_YYYY-MM-DD.xlsx.enc
    reporting_date_presets.json.enc
    oasis_evaluations/
        OASIS_YYYY-MM-DD_to_YYYY-MM-DD_<opaque-id>.csv.enc
```

If your configured OPD folder has another name, that same folder is used instead.
The app creates the subfolder on the first successful save. **Never upload the
unencrypted OASIS CSV to the public GitHub repository yourself.**

Retention is intentionally non-destructive:

- An identical byte-for-byte CSV is recognized and not written again, including
  across sessions or after changing its local filename.
- Different contents create another saved export, even if the filename or course
  dates match. Existing exports are not overwritten or deleted.
- Each export is a separate snapshot. The app does **not** merge evaluations,
  combine overlapping exports, calculate scores, or add OASIS results to teaching
  summaries. Those would be separate analysis features.
- The course dates in a filename describe the minimum Start Date through maximum
  End Date inside that CSV, **not** the upload date or Submit Date range. They are
  for organization, not an overwrite key.
- The opaque ID is keyed; no unkeyed hash of the plaintext CSV, original local
  filename, evaluator name, learner name, answer, or comment is included in the
  public filename or commit message.

There is no separate index to update: a verified encrypted CSV is itself a saved
export. A retry after an interrupted confirmation detects the same file rather
than adding another copy. If two sessions race to create it, the unsuccessful
save is not reported as successful; retry explicitly.

## Reload and recover the original CSV

In the same section:

1. Choose an entry in **Saved OASIS exports**. The dropdown shows course-date
   coverage and a short export ID. It is ordered by course dates, not upload date.
2. Click **Load / decrypt selected OASIS export**.
3. Click **Download original OASIS CSV**.

The downloaded filename is a neutral archive filename; the **CSV bytes match the
original upload exactly**. Selecting a new entry clears any previously loaded
file. **Refresh saved OASIS exports** retrieves changes made by another session.
Reloading does not write to GitHub. Large-file retrieval uses GitHub's raw Contents
API response when the inline base64 content is omitted.

Previous encryption keys in your existing `previous_encryption_keys` setting are
supported for recovery and duplicate recognition. Removing the only correct key
prevents recovery. A new key alone does not make older files decryptable.

## Existing functions are unchanged

The original OPD archive service, OPD identifier rules, saved date presets, custom
reporting dates, student schedules, primary-preceptor workbook, Learner Reach,
outpatient-over-nursery rule, chair report, individual reports, unique-student
counts, and pie charts have not been edited. OASIS snapshots have their own folder
and `.csv.enc` filenames; the OPD parser only picks up `.xlsx.enc` OPD filenames.
No OPD refresh, upload, migration, or re-encryption is required for this update.

## Privacy and access

This section follows your current **no app-password** choice. Anyone who can
access the running app can upload exports and use the original-CSV download.
The encryption key is a server secret, but the app can decrypt files on a user's
behalf. Encryption at rest is not a replacement for access controls. Use your
institution's approved arrangements for identifiable evaluation data.

CSV contents are encrypted before being sent to GitHub, using Fernet authenticated
encryption. Public course-date filenames, file sizes, encryption timestamps and
Git commit metadata are still visible. Downloads from the running app are
**unencrypted** originals. Do not publish them or your secrets.

## Troubleshooting

**New menu option is missing:** update the launcher as well as adding both new
modules. The launcher is the only existing runtime file changed by this update.

**ModuleNotFoundError:** keep `schedule_app/sections/oasis_evaluation_archive.py`
and `schedule_app/services/oasis_evaluations.py` in those exact folders; do not
flatten or nest the application inside an extra ZIP-name folder.

**Archive setup incomplete:** the existing `[opd_archive]` settings must contain
owner, repo, github_token and encryption_key. The same settings that save OPDs
are used for OASIS; there is no separate `[oasis]` secrets section.

**Missing columns / malformed CSV:** upload the original OASIS CSV export. The
file is not reconstructed from an Excel workbook or a report screenshot.

**Save not confirmed:** check the token's Contents write permission, expiration,
repository/branch, and connection. Use the retry button. A wrong decryption key
or altered ciphertext does not cause an existing export to be overwritten.

**Many exports cover the same dates:** they contain different bytes and were
retained deliberately. Nothing is merged or counted together by this archive.

**Archive list limit:** this version stops visibly at the GitHub directory-list
limit (1,000 entries) rather than presenting a partial list as complete. Contact
your app maintainer for pagination/partitioning support before reaching that size.

## Verification and source references

See `TESTING_OASIS_ARCHIVE.md` in the complete package. The provided test suite uses
synthetic data and simulated GitHub/Streamlit calls. The actual uploaded CSV was
also tested locally for byte-identical recovery, including the larger-file raw
response path. The sample data and credentials are not in either code ZIP.

Official API documentation consulted for this implementation:

- GitHub Contents API: https://docs.github.com/en/rest/repos/contents
- Fernet / MultiFernet: https://cryptography.io/en/latest/fernet/
- Streamlit file uploader: https://docs.streamlit.io/develop/api-reference/widgets/st.file_uploader
