# Add student-evaluation CSV storage under OASIS Evaluations

## Install the small update (recommended)

This update is built on the modular app with the missing-preceptor-usernames update.
It changes one existing runtime file and adds two new modules:

| Action | File |
|---|---|
| Replace | `schedule_app/sections/oasis_workflow.py` |
| Add | `schedule_app/sections/oasis_student_evaluations.py` |
| Add | `schedule_app/services/oasis_student_evaluations.py` |

1. Extract `Schedule_App_OASIS_Student_Evaluations_Update.zip`.
2. Merge its `schedule_app` folder into your existing repository. Replace the matching file and add the two new files at the paths above. **Do not delete your existing folder.**
3. Restart the Streamlit app. Keep `app_sch_2026.py` as the entrypoint.
4. Open **OASIS Evaluations** in the main menu.

Keep the launcher, settings.py, requirements.txt, email/name mappings, Streamlit Secrets, encryption key, GitHub date presets, and teaching-to-OASIS links unchanged. No new credentials, dependencies, repository, or app password are required. Do not regenerate the encryption key. Existing archived files need no migration.

## Use the new upload

At the top of **OASIS Evaluations**, choose the evaluation type:

- **Evaluations of educators**: your existing learner-feedback upload, missing-username correction, Submit Date/date-presets workflow, and automatic encrypted educator-summary CSV.
- **Evaluations of students**: the new archive for preceptors' assessments of students. Upload the original CSV; no reporting dates or username corrections are required to preserve it.

The student-upload path supports every form type present in the sample supplied for this update:

- `*Clinical Assessment of Student`
- `*PEDS Handoff`
- `*PEDS History Taking & Physical Exam`

One CSV can contain any combination of these three forms. The file is stored in full; it is not separated into individual students, reduced to one form type, merged, or rewritten. Unknown form types or mixed educator/student feedback are rejected with a routing message; no rows are silently dropped. For a genuinely new student-assessment form, add its verified exact title to `STUDENT_FORM_TITLES` in the new service module.

After uploading, the app automatically:

1. Validates the CSV's structure and Evaluation form types.
2. Encrypts the **entire original file** using the existing archive key.
3. Saves it in the existing GitHub repository and branch, in a separate student folder.
4. Retrieves and decrypts the saved ciphertext, and checks a byte-for-byte match.
5. Displays **Student-evaluation CSV encrypted, archived, and verified in GitHub.**

No download or second save button is required. Errors are not retried repeatedly on unrelated app reruns. A failed attempt shows **Retry student-evaluation encrypted save**; a verification failure may occur after a write, so retry to verify rather than assuming no file was written.

## Storage location and update behavior

For the default archive folder:

```text
opd_archive/
    OPD_YYYY-MM-DD.xlsx.enc
    oasis_evaluations/                 # Existing educator-feedback source CSVs
    oasis_reports/                     # Existing educator-summary output CSVs
    oasis_student_evaluations/          # New, separate student-assessment originals
        OASIS_<course-dates>_<export-id>.csv.enc
```

The actual parent folder follows your existing Streamlit Secrets configuration.

**Each changed export is retained as a separate encrypted snapshot.** Re-uploading identical bytes does not save another copy, even if you rename the local file or use another session. An updated file with the same course dates is a new snapshot; the previous one remains available. This is the same snapshot convention as your other original OASIS archive, not the OPD's one-current-file-per-rotation convention.

The archive identifier is based on the file bytes using a secret-keyed identifier; a separate identifier domain is used for student files. Public filenames do not contain student/preceptor names or unkeyed plaintext hashes. Original local filenames are not archived. Course dates, encrypted file size, and Git metadata remain publicly visible. Fernet tokens also expose their encryption timestamp. Encryption is not an access-control or institutional-approval claim.

## Download an original later (optional)

Remain in **Evaluations of students**. In **Reload a saved student-evaluation CSV**:

1. Choose an item from **Saved student-evaluation exports**.
2. Click **Load / decrypt selected student-evaluation CSV**.
3. Click **Download original student-evaluation CSV (optional)**.

The download has a neutral archive filename, with contents identical to the uploaded CSV. It is **unencrypted**. Reloading does not write to GitHub or create a report. Use **Refresh saved student-evaluation exports** to see files uploaded from another session.

The dropdown is ordered by **course-date coverage**, not upload time. The export identifier distinguishes different snapshots covering the same dates. Do not assume the first item is the newest upload.

## Separation from educator reporting

The new folder is not read by the educator-feedback summary builder. Uploading student assessments does **not**:

- generate a new educator-summary CSV;
- add student grades/comments to educator evaluations or individual preceptor Word documents;
- alter the chair report, teaching-hours calculations, Learner Reach, or continuity counts;
- change preceptor usernames or saved links;
- use or change your active reporting-date presets.

The educator-feedback upload now also recognizes the three known student-assessment types and directs you to the correct option before saving. Previously archived files are not automatically moved or deleted; this update does not remediate a student file that was already placed in the educator archive before the update.

## What the page displays

Only aggregate file metadata is shown: forms in the file, question-response row count, original size, recognized form titles, and course-date coverage. It does not preview individual students, scores, free-text comments, or original local filenames.

“Forms in this file” counts distinct Course ID / Evaluation / Form Record combinations, not question-response rows. This count is informational, not a claim about completion, grading, or unique students. Rows with missing Form Record identifiers are still preserved and flagged. Rows with blank Submit Date are preserved as well. Start Date / End Date organize the original snapshot; no Submit Date, academic-year, or weekend filter is applied to storage.

Unknown/blank/unreadable course dates cause an `undated` filename rather than a guessed range. The date fields inside the CSV are never corrected or replaced.

## Access and privacy

No new login/password has been added. Anyone who can use the running app can access this original-file recovery function unless the deployment has access controls elsewhere. This file contains identifiable student assessments. Use institution-approved access and storage, and do not commit unencrypted CSV downloads to a public repository. Only encrypted originals are uploaded by this new feature; no evaluation data or encryption keys are included in this software update package.

## Limits and failure handling

The existing OASIS archive limits are retained: 10 MiB per original CSV; 250,000 response rows; bounded field lengths; no silent partial directory listing once the GitHub Contents directory limit is reached. CSV quoting, headers, and row widths are checked. UTF-8/BOM, UTF-16 with BOM, and Windows-1252 source text are recognized; the stored bytes still retain the original encoding.

Wrong keys, altered ciphertext, conflicting writes, and failed verification produce errors without replacing the existing snapshot. Retain your original key (or configured previous keys) for historical recovery. Missing token/key settings block archiving rather than saving plaintext.

## Editing locations

- Menu choices inside OASIS Evaluations and educator-upload guard: `sections/oasis_workflow.py`
- Student upload, status, retry, and recovery controls: `sections/oasis_student_evaluations.py`
- Student CSV form validation and archive separation: `services/oasis_student_evaluations.py`
- Shared, unchanged encryption/verified transport: `services/oasis_evaluations.py` and `services/opd_archive.py`

## Technical references

- Streamlit upload widget and original-byte access: https://docs.streamlit.io/develop/api-reference/widgets/st.file_uploader
- GitHub Contents API: https://docs.github.com/en/rest/repos/contents
- Fernet/MultiFernet encryption and recovery: https://cryptography.io/en/latest/fernet/

See `TESTING_OASIS_STUDENT_EVALUATIONS.md` in the full package for the local test scope. No live repository or deployment was accessed or changed while producing this update.
