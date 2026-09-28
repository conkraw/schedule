# Pediatric Clerkship Schedule App — Modular Version with Custom Reporting Dates

**Keep `app_sch_2026.py` as your Streamlit entrypoint.** It is now a 35-line launcher. Each sidebar section has its own file, and reusable Excel, archive, and Word-report code has been moved into clearly named modules.

This package keeps the existing modular scheduling app and adds editable dates to **Preceptor Teaching Summary**. Choose **Custom dates**, set the start and end dates, and enter a report label such as `26-27`. The chair summary, individual reports and both CSVs use that exact inclusive period, even when it crosses July or lasts longer than twelve months. The previous July-June mode remains optional. No app password has been added.

**Already using the working modular app?** Use the smaller six-file update and instructions in **UPDATE_CUSTOM_DATES.md**. It leaves the launcher, settings/mappings, requirements, archive service and all other sections untouched. A complete installation is described below.

## Update your existing deployment

1. Make a local backup of your currently working code and custom mappings.
2. Extract this ZIP on your computer. Upload its **contents** to the same repository directory that currently contains `app_sch_2026.py`. Replace that file and add the **entire `schedule_app` folder, including its subfolders and `__init__.py` files**. Do not flatten the folders or upload only the small launcher. Uploading the ZIP itself is not installation.
3. Keep your existing requirements, or merge the included `requirements.txt` entries into them. This refactor adds **no new third-party dependencies** relative to the previous full-app package. Do not remove extra packages your repository needs for other apps.
4. Leave the Streamlit entrypoint set to `app_sch_2026.py` at its existing location. Restart the app after updating the files.
5. **Keep your existing Streamlit Secrets unchanged**, including the encryption key, token, repository, branch, folder, and any old decryption keys. Do not regenerate a key, re-encrypt OPDs, or replace your archive repository.
6. Open each section as usual. On the teaching-summary screen, use **Load / refresh archived OPDs** before generating a new ZIP. Existing archived files remain usable.

Files must sit together like this:

```text
repository directory/
├── app_sch_2026.py
├── requirements.txt
└── schedule_app/
    ├── __init__.py
    ├── settings.py
    ├── sections/
    │   ├── __init__.py
    │   └── ...nine section files...
    ├── services/
    │   ├── __init__.py
    │   └── ...archive, schedule, and analysis helpers...
    └── reports/
        ├── __init__.py
        └── ...Word and teaching ZIP builders...
```

**Custom entries:** The uploaded source had an empty `PRECEPTOR_EMAIL_MAP`. It is still empty, now in `schedule_app/settings.py`. Copy any entries that exist only in your deployed version into that dictionary. The same file contains teaching-name aliases, per-rotation name-order overrides, and work-type grouping settings.

## Where to edit

Start with **EDITING_GUIDE.md**. It maps every sidebar choice to a file, and separately identifies the chair Word report, individual Word reports, weekly Power Automate report, encryption/archive logic, and editable settings.

For example:

- Chair Word summary: `schedule_app/reports/chair_summary.py`
- Individual teaching Word reports: `schedule_app/reports/individual_teaching.py`
- Preceptor email addresses and work-type grouping: `schedule_app/settings.py`

Sections use ordinary Python imports and `render()` functions. There is no second hidden monolithic app, code stored in strings, `exec`-based section loader, or shared global dictionary of uploaded workbooks. Runtime data stays local to the page call or in the existing Streamlit session state.

## Running locally

From this extracted directory, with your existing Python environment:

```bash
python -m pip install -r requirements.txt
streamlit run app_sch_2026.py
```

For archive features, keep local secrets in `.streamlit/secrets.toml` and never commit that file. `secrets.example.toml` is a placeholder-only reference. Setup and offline recovery details are in **SETUP_OPD_ARCHIVE.md**.

## Checks included

```bash
python -m unittest discover -s tests -v
```

These are offline tests with invented data and simulated Streamlit/GitHub calls. They do not use real credentials or change a live repository. See **TESTING.md** for the validation actually performed, including comparisons with the supplied monolithic app.

## Access and data handling

The no-password behavior remains as requested. Anyone who can reach the running app can use its archive functions unless access is restricted elsewhere. Encrypted GitHub storage does not restrict who may ask the running app to decrypt a workbook. Downloaded Excel, CSV, Word, and ZIP reports are unencrypted; do not put them in the public repository. No actual secrets, OPD workbooks, or generated reports are included in this code package.
