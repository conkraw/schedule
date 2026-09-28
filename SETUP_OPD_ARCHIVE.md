# Encrypted OPD archive — modular app

## Existing archive: keep the current setup

This update only reorganizes code. **Keep your encryption key and GitHub settings unchanged.** The original `.xlsx.enc` archives can be loaded without conversion, migration or re-upload. The app still uses the same `[opd_archive]` section in Streamlit Secrets. There is no additional app-password requirement.

Upload `app_sch_2026.py` **and the entire `schedule_app` folder** as described in README.md. The archive code now lives in `schedule_app/services/opd_archive.py`; shared archive widgets live in `schedule_app/services/opd_archive_ui.py`.

## Secrets reference

Keep real values in Streamlit Secrets, not GitHub source:

```toml
[opd_archive]
owner = "YOUR_GITHUB_USERNAME"
repo = "YOUR_EXISTING_ARCHIVE_REPOSITORY"
branch = "main"
folder = "opd_archive"
github_token = "YOUR_EXISTING_GITHUB_TOKEN"
encryption_key = "YOUR_EXISTING_FERNET_KEY"
# Optional only when already used for a planned key migration:
# previous_encryption_keys = ["AN_OLD_FERNET_KEY"]
```

`owner` is an account/organization name, not a URL. Keep the actual branch and folder you already use. The token needs access to read and update contents in the configured repository. Keep a separate, secure backup of the encryption key; losing it prevents recovery of files encrypted with that key.

Do not place actual Secrets inside this package's `secrets.example.toml` and upload the filled file. The included template is for reference only. For local runs use `.streamlit/secrets.toml`, which is excluded by `.gitignore`.

## Only for a genuinely new archive

Use a separate public archive repository, initialized with a README, so the storage token does not need permission to write your app's code. Create a fine-grained GitHub token limited to that repository with **Contents: Read and write**. Store the token under `github_token` in Streamlit Secrets.

For a brand-new archive, generate a Fernet key locally using either included tool:

```bash
python generate_opd_secrets.py
```

Or open `generate_opd_secrets.html` locally. These utilities are not imported by the app. **Do not run a new key through Secrets when updating an archive that already contains files.** The new key would not decrypt those existing files.

No repository, token, key, or deployment was created or changed during this refactor.

## Normal workflow

In **Create Student Schedule**, upload the original OPD, wait for **OPD archived and verified**, and provide the matching rotation list. The original workbook is encrypted before storage; the archive identifier comes from its first scheduled Monday. A changed upload with that Monday replaces the rotation's current file. Other rotations remain separate. An identical upload avoids an unnecessary commit.

In **OPD Archive**, select a rotation and load/decrypt its original workbook. Reloading does not overwrite its current archived version.

In **Preceptor Teaching Summary**, refresh the archive scan to include newer saved files, select academic years, and generate the CSV/Word ZIP. This analysis does not write decrypted workbooks or reports back to GitHub.

## Recovery outside the app

The included `decrypt_opd.py` retains the offline recovery workflow:

```bash
python decrypt_opd.py OPD_2026-08-03.xlsx.enc recovered_OPD.xlsx
```

The script asks for the matching encryption key privately. It does not need Streamlit or a GitHub token once the encrypted file is on your computer. It refuses to overwrite an existing output file.

## Important limits retained from the original

The app saves one **current** file per rotation. Older encrypted versions remain in Git history. Public rotation-date filenames, file sizes and commit metadata remain visible. Uploading an older revision later can make that revision current again; the code cannot infer which document revision was intended.

Workbook validation, size limits, date validation, GitHub conflict handling, encryption-key errors and fail-visible behavior remain unchanged. Source labels and names around `~` are not rewritten in the archived workbook.

The app has no added sign-in gate. Access to the running app permits archive operations unless restricted elsewhere. Keep decrypted workbooks and report downloads out of the public repository, and follow your institution's storage/access requirements.
