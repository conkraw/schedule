# Save and reuse reporting dates in GitHub

This update is for the working modular app with Learner Reach, custom reporting dates,
clinical-experience pie charts, teaching-contributor filtering, and the Academic
Pediatrics-over-PSHCH Nursery priority rule. Those calculations and reports are unchanged.

## Install the small update

1. Extract **Schedule_App_GitHub_Date_Presets_Update.zip**.
2. Merge its `schedule_app` folder into the existing folder in your app repository.
   Replace the matching section file and add the two new files below. **Do not
   delete the existing folder or upload the ZIP itself as the installation.**
3. Restart the app and open **Preceptor Teaching Summary**.

The update contains exactly three runtime Python files:

```text
schedule_app/
    sections/
        preceptor_teaching_summary.py    # Updated to use the new date controls
        reporting_date_controls.py      # New dropdown, save, load and delete UI
    services/
        reporting_presets.py            # New encrypted GitHub preset catalog
```

Leave `app_sch_2026.py`, `schedule_app/settings.py`, `requirements.txt`, the other
modules, and Streamlit Secrets unchanged. No new packages, tokens, passwords, or
keys are needed. Preserve any custom code you added to the replaced section.

The complete ZIP, **Schedule_App_Modular_GitHub_Date_Presets.zip**, is an alternative
for a full installation. Use the small update for your working deployment so your
custom settings and mappings are not replaced. The full package is based on the
last full outpatient-priority ZIP supplied in this conversation, not on a live
copy of your GitHub repository.

## Save dates the first time

In **Preceptor Teaching Summary**:

1. Choose **Custom dates** and enter the start date, end date, and report label.
2. Open **Save or delete date presets in GitHub**.
3. Enter a **Preset name**, such as `26-27 teaching year` or `Chair review`.
4. Click **Save these dates to GitHub**. Wait for the saved-and-verified message.

The **preset name** is the label in the dropdown. The **report label / academic
year** is the text printed in the Word reports and `academic_year` CSV column.
They may be the same or different. Both dates are inclusive, weekends remain
included, and custom periods are not split at July 1.

Example only (not automatic defaults):

| Preset name | Start | End | Report label |
|---|---|---|---|
| 26-27 teaching year | February 1, 2026 | March 31, 2027 | 26-27 |
| Chair review | January 1, 2026 | September 30, 2026 | Chair review |

No original OPD, student record, or generated report is saved by these controls.
The date preset contains only its name, report label, exact dates, identifier,
and creation/update timestamps.

## Reuse saved dates

The **Saved date presets** dropdown loads the small catalog when you first visit
this section in a new session. Select a preset and click **Load selected dates**.
The app fills both date fields and the report label, then switches to Custom dates.

A preset is not automatically applied just because the page opens. This prevents
an old period from replacing dates you have already entered. The dropdown shows
the saved name, exact dates, and report label before you load it.

Loading a preset clears any old teaching-report download, so you must generate
the reports for the newly selected dates. It keeps the OPD scan already in memory:
you do not have to download/decrypt all OPDs again just to change dates. Use the
existing **Load / refresh archived OPDs** button when there is no scan yet or the
archive itself has changed.

Use **Refresh saved presets** to see changes made from another browser/session.
Refreshing presets does not change your current date fields or reload the OPDs.

## Update an existing preset

Load the preset, edit the dates or report label, and keep its name in **Preset name**.
The app recognizes existing names without regard to capitalization or repeated
spaces. Review the old dates displayed below the name field, check the replacement
confirmation, and click **Update saved preset in GitHub**.

To retain the old preset and save another period, use a different preset name.
The dropdown entry is identified by its name, not just by the start date.

An unchanged save is verified without creating an unnecessary Git commit. The
app does not autosave every date edit or button rerun; changes are persisted only
when you explicitly click a save/update/delete button.

## Delete a saved date preset

Select the preset in the dropdown. In **Save or delete date presets in GitHub**,
review the selected name, check **Confirm deletion of this date preset only (not
OPDs)**, and click **Delete selected preset**.

Deletion removes that preset from the current dropdown. It does not delete any
OPD, report, other preset, or your currently entered dates. Deleting the final
preset leaves an empty encrypted catalog rather than deleting a folder.

Older encrypted versions remain in Git history; this is not a history-erasure tool.

## Where the presets are stored

The app creates this one file automatically on the first save, using your existing
`[opd_archive]` owner, repository, branch, folder, token, and encryption key:

```text
opd_archive/reporting_date_presets.json.enc
```

A custom `folder` setting replaces `opd_archive` in that path. No extra configuration
or manually created file is needed. The stored contents are encrypted with the same
Fernet/MultiFernet configuration already used for the OPDs. The names and dates do
not appear in the filename or commit message.

The OPD scanner only recognizes rotation files named `OPD_YYYY-MM-DD.xlsx.enc`;
it ignores this settings catalog. Preset saves do not replace any rotation file.

**Keep the same encryption key.** An unreadable or undecryptable catalog is an
error, not an empty list, and is never automatically overwritten. Previously
configured decryption keys remain supported.

Presets are shared by everyone using this app with the same archive settings.
There is no new app password, per-user preset ownership, or permission system.
Anyone with access to the running app can use these controls unless you restrict
access elsewhere. Saved presets survive session closure and app restarts because
they are stored in GitHub, not only in session state.

## Failure and concurrent edits

A GitHub error displays a message, not a success or a false empty catalog. You can
still enter dates manually and use an already-loaded OPD scan. Refresh the preset
list successfully before saving or deleting again.

Before changing the catalog, the service checks its current GitHub file revision.
If another session changed it, your operation stops and asks you to refresh and
review the latest entries. A race at the write itself also stops without an
automatic retry. Successful writes are read back at the returned commit and
verified by decryption before the app confirms them.

The existing archive token needs its usual repository **Contents: Read and write**
permission. A token that already saves OPDs to this same repository/branch normally
needs no change. Branch protections or token expiration can still prevent writes.

## What happened to the JSON upload controls?

They were removed. You do not have to download, keep, or upload a settings file.
The existing report ZIP can still include `Reporting_Period.json` and
`Reporting_Period.csv` as a record of the dates used; neither is required to reuse
a saved preset. Existing local settings files are not automatically imported.
Enter their dates once and save a GitHub preset instead.

## Maintainer notes

- Edit the preset interface in `schedule_app/sections/reporting_date_controls.py`.
- Edit encrypted preset persistence in `schedule_app/services/reporting_presets.py`.
- Existing date validation and report metadata stay in `services/reporting_periods.py`.
- The catalog supports up to 250 presets; labels are 1-60 characters and preset
  names are 1-80 characters. Date limits remain 1970 through 2100.
- Read/write calls use the existing GitHub transport and do not print tokens,
  plaintext workbooks, or raw network response bodies into user-facing errors.

## Validation of this update

Python compilation and all **212 offline automated tests passed**, including **49
new preset-specific tests**. Tests cover create/load/update/delete, fresh-session
reuse, encrypted round trips, duplicate names, confirmation checks, stale edits,
wrong keys, tampering, metadata validation, manual-date fallback, and preservation
of existing report/scheduling behavior. Two legacy UI tests were adapted to remove
the retired JSON-uploader/download behavior; their report/conflict checks remain.

Streamlit and GitHub interactions were simulated. A real Streamlit server was not
available in this environment, and no live GitHub repository or cloud deployment
was accessed or changed. The Word/Excel report builder modules were not edited;
this package contains code, not a new report or OPD.

## Technical references

- GitHub Contents API (file revision SHA and Contents write permission):
  https://docs.github.com/en/rest/repos/contents
- Streamlit callback order and widget session state:
  https://docs.streamlit.io/develop/api-reference/caching-and-state/st.session_state
- Fernet authenticated encryption:
  https://cryptography.io/en/latest/fernet/
- Git history and why deletion does not erase previous revisions:
  https://docs.github.com/en/authentication/keeping-your-account-and-data-secure/removing-sensitive-data-from-a-repository
