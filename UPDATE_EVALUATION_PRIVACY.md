# Evaluation Records: reduced CSV storage and a section-only password

This update supersedes the older instructions that promised whole-original OASIS CSV storage and no evaluation-section password. It does not change the encryption key, OPD archive, teaching calculations, username links, date presets, or report formats.

## Install the update

1. Extract `Schedule_App_Evaluation_Privacy_Update.zip`.
2. Replace `app_sch_2026.py` and merge the included `schedule_app/` folder into your existing repository. Replace matching files and add the new files. **Do not delete the existing folder.**
3. Add this separate section at the end of **Streamlit Settings → Secrets**, replacing the placeholder with your own long password:

```toml
[evaluation_access]
password = "REPLACE_WITH_YOUR_OWN_LONG_PASSWORD"
```

Use a password or passphrase of at least 16 characters, different from the encryption key and GitHub token. The placeholder shown above is deliberately rejected. Keep the password out of Python source and GitHub. The included `generate_evaluation_password.py` is an optional local generator; run `python generate_evaluation_password.py` on your own computer to print a random password setting. It writes no file and makes no network calls.

4. Leave your entire existing `[opd_archive]` configuration unchanged, including `encryption_key`. Do not generate a new encryption key. Keep settings.py, requirements.txt and all custom mappings.
5. Restart Streamlit. **OASIS Evaluations** is now **Evaluation Records**. Open it, enter the new section password and click **Unlock Evaluation Records**.

No new dependencies, GitHub permissions, repository or encryption migration are needed. Missing/invalid password settings block Evaluation Records; they do not block the rest of the app.

### Customized launcher

Keep your custom non-evaluation entries. Change this dictionary entry:

```python
"OASIS Evaluations": "oasis_workflow",
```

to:

```python
"Evaluation Records": "oasis_workflow",
```

Before rendering the sidebar radio, migrate old selections:

```python
if st.session_state.get("schedule_app_mode") in (
    "OASIS Evaluation Archive", "OASIS Educator Reports", "OASIS Evaluations"
):
    st.session_state["schedule_app_mode"] = "Evaluation Records"
```

The updated evaluation section modules enforce the gate themselves, including the legacy archive/report modules. A customized launcher cannot bypass it merely by retaining an old menu entry.

## Which columns are saved

A strict allowlist is applied on the server **before encryption and before any GitHub write**. New/unknown source columns are not retained automatically. Retained cell values and record order are preserved; the reduced CSV is deterministically encoded as UTF-8 with a BOM. This is not a byte-for-byte copy of the full upload.

### Evaluations of educators — up to 15 columns

```text
Course ID
Start Date
End Date
Evaluator
Evaluator Username
Evaluator External ID
Evaluator Email
Evaluation
Form Record
Question ID
Question
Answer text
Multiple Choice Value
Multiple Choice Label
Submit Date
```

These support educator identity resolution (including existing username overrides), form/question deduplication, course/date checks, full question wording, numerical averages, and strengths/improvement comments. Structured Student fields, Who Completed, gender, location, department, classification, Question Number and Multiple Choice Order are not retained.

**Educator-feedback comments remain**, because they are part of your reports. Free text can still contain identifying information; removing structured columns is not automatic anonymization.

### Evaluations of students — up to 11 columns

```text
Course ID
Start Date
End Date
Student
Student External ID
Evaluator
Evaluator Username
Evaluator Email
Evaluation
Form Record
Submit Date
```

`Student` and `Student External ID` are required to match OPD names to eligible students. Evaluator name/username/email allow matching and source-issue review. Course/form/date fields support archive organization and unique-form/Submit Date counting. Exported Evaluator Username remains a review hint, not an automatic substitute for the email local-part.

**Question text, Question ID/Number, Answer text, scores, multiple-choice labels/order, student assessment comments, student email, other student identifiers, unused demographics, and Evaluator External ID are removed.** The assessment-completion measure already uses form metadata rather than answers, so these removals do not change its calculation.

Clinical Assessment, History & Physical, and Handoff form metadata remain supported. Handoff metadata can support the existing student-name/ID crosswalk; it still does not count as one of the two completion form types.

Missing optional columns are not fabricated. Existing report checks still flag missing or inconsistent identities rather than silently generating guessed usernames or zero counts. Question-level rows remain in source order after reducing columns; the existing form-counting logic collapses them, so they do not inflate evaluations.

## Upload and report workflow

The two choices under Evaluation Records remain **Evaluations of educators** and **Evaluations of students**.

- Educator upload: reduce columns → encrypt and verify the reduced source → resolve usernames as needed → rebuild and save the cumulative educator-summary CSV, using the applied Submit Date period.
- Student upload: reduce to metadata → encrypt and verify → use these records for the existing assessment-completion checks in Preceptor Teaching Summary.

Same retained data means the same snapshot, even when removed demographics/comments change. A genuinely changed retained form, date, answer or identity produces a new snapshot. In student-assessment uploads, changing a score or narrative alone does not produce a new stored version, because those fields are no longer stored.

The automatic educator output is still one encrypted summary CSV. Its existing fields, question labels, means, combined comments, Submit Date controls and saved presets are unchanged. No new report CSV/ZIP/Word output is added by this privacy feature.

The archive paths remain:

```text
opd_archive/oasis_evaluations/          # Reduced educator-feedback sources
opd_archive/oasis_student_evaluations/  # Reduced student-assessment metadata
opd_archive/oasis_reports/              # Existing generated educator-summary CSVs
```

The configured base folder may have a different name. Public filenames contain dates and keyed opaque identifiers, not names. Identifiers are calculated from reduced content, not from discarded columns. Existing username and date-setting catalogs remain separate.

## Reduced downloads, including older files

Optional source downloads now return the **reduced CSV**, not the full original. Stored-file verification occurs first, including decryption and filename/content checks. Older full exports are then reduced in memory before they reach a download button or downstream report calculation.

The upload itself necessarily reaches the Streamlit server for parsing. The app does not send discarded values to GitHub, put them into a second audit file, or keep a plaintext original on disk. Generated report downloads are unencrypted and must be handled as educational records.

## Optional cleanup of previously stored full exports

Future-upload filtering does **not** remove old encrypted full exports from GitHub. To reduce their current saved copies:

1. Unlock **Evaluation Records**.
2. Open **Stored columns and older-file privacy review**.
3. Click **Check previously saved evaluation files**. The table shows source type, neutral archive filename, stored/needed column counts and whether replacement is needed. It does not show answers or individual identities.
4. Review and select the explicit confirmation.
5. Click **Minimize previously saved evaluation files**.

For each file, the app writes and verifies a reduced encrypted copy before removing the superseded current path. SHA checks prevent silently overwriting a source that changed after review. The reduced file gets a new content identifier. Identical reduced sources may coalesce into one snapshot. Current OPDs, reports, saved date presets and username/ID catalogs are not edited by this cleanup.

A failed save never triggers deletion of the older file. A failed delete can leave both copies; a rescan/retry is safe, and the existing evaluation deduplication prevents double-counting. The UI reports confirmed progress and failures; there is no automatic rollback of completed replacements.

**This is NOT history erasure.** Earlier full encrypted copies may remain in Git history, clones, forks or independently retained copies. The app does not force-push/rewrite history or claim to revoke already downloaded data. Permanent repository-history cleanup requires a separate carefully coordinated procedure. The removed fields will no longer be recoverable through this app's reduced-download feature; retain an institution-approved original elsewhere when required.

## What the password protects—and does not protect

The password gates the **Evaluation Records administration section** before upload, listing, decryption, username/date management or downloads in that section. It also gates the old standalone OASIS screens if a custom launcher still exposes them.

- Authorization is session-local; it is not a public query parameter or shared process flag.
- Submitted passwords are removed from widget/session state after authentication. A fingerprint and timestamps are retained, not the password itself.
- Access expires after 30 minutes without interaction within a guarded evaluation screen, and at eight hours regardless. Expiration is checked on the next section interaction; it cannot erase a page or download already delivered to a browser.
- **Lock Evaluation Records** clears evaluation-administration session data and cached downloads; it does not delete stored files or clear unrelated teaching data.
- Changing the password invalidates old authorization on the next check. Changing the section password does not change the encryption key or require re-encryption.
- Wrong attempts receive a per-session delay. This is a basic shared-password gate, not institutional SSO, MFA, individual account permissions, or global brute-force protection. Other sessions can start independently.

**Other app sections remain password-free, as requested. In particular, existing OASIS-linked feedback and assessment-completion information can still appear in Preceptor Teaching Summary and its downloadable reports. This update is not an app-wide evaluation-data authorization boundary.** Choose appropriate deployment-level access controls for everyone who can reach those report features. A renamed menu is not itself a security control.

Minimization and encryption do not establish institutional/regulatory approval. Use your institution's approved arrangements for student/evaluation records, keep a secure backup of the existing encryption key, and share the section password only with intended users.

## Changed runtime files

| Action | File |
|---|---|
| Replace | app_sch_2026.py |
| Replace | schedule_app/services/oasis_evaluations.py |
| Replace | schedule_app/services/oasis_student_evaluations.py |
| Replace | schedule_app/services/oasis_workflow.py |
| Replace (navigation wording only) | schedule_app/services/teaching_evaluations.py |
| Replace | schedule_app/sections/oasis_workflow.py |
| Replace | schedule_app/sections/oasis_student_evaluations.py |
| Replace | schedule_app/sections/oasis_evaluation_archive.py |
| Replace | schedule_app/sections/oasis_educator_reports.py |
| Replace (navigation wording only) | schedule_app/sections/preceptor_oasis_links.py |
| Add | schedule_app/services/oasis_privacy.py |
| Add | schedule_app/services/evaluation_access.py |
| Add | schedule_app/sections/evaluation_privacy.py |

The optional password generator and instructions do not need to be deployed. All existing calculation/report modules, OPD storage, settings.py and requirements are preserved, apart from the two navigation-message changes identified above and OASIS source filtering/workflow invalidation.

## Verification

See `TESTING_EVALUATION_PRIVACY.md` in the full package. Tests use simulated GitHub/Streamlit, never the live repository or credentials. Local comparisons of both supplied CSV samples confirmed that educator summary rows/username issues and student form-completion metadata/matching results were unchanged after minimization.

## Technical sources

- Streamlit Secrets: https://docs.streamlit.io/deploy/streamlit-community-cloud/deploy-your-app/secrets-management
- Streamlit Session State and callbacks: https://docs.streamlit.io/develop/api-reference/caching-and-state/st.session_state
- GitHub verified reads/writes and SHA-guarded deletion: https://docs.github.com/en/rest/repos/contents
- Git history and sensitive-data removal limitations: https://docs.github.com/en/authentication/keeping-your-account-and-data-secure/removing-sensitive-data-from-a-repository
