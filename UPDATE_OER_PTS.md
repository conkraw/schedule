# OER and PTS: renamed, moved, and password-protected

## Install the update

1. Extract `Schedule_App_OER_PTS_Update.zip`.
2. Replace `app_sch_2026.py` at the root of your existing app repository.
3. Merge the supplied `schedule_app/` folder into the existing folder, replacing the matching files. **Do not delete your existing folder.**
4. Restart the app. Keep Streamlit's entrypoint set to `app_sch_2026.py`.

This update contains 13 replacement Python files. Some change only menu-navigation wording. All report writers and calculation logic are unchanged. The complete modular ZIP is also available for a fresh installation; the smaller update is recommended for preserving deployed customizations.

## Main menu

The existing scheduling sections keep their relative order. Under **What do you want to do?**, the final choices are now:

```
OPD MD PA Conflict Detector
Shift Availability Tracker
OER
PTS
```

- **OER** replaces the former evaluation-administration label.
- **PTS** replaces the former teaching-summary label.
- An already-open session with either old label is redirected to its new label before the menu widget is created.
- Internal Python module names, saved filenames, GitHub folders, and report-document titles remain unchanged.

## Password: keep the one you already configured

Both OER and PTS use the existing secret:

```toml
[evaluation_access]
password = "REPLACE_WITH_YOUR_OWN_LONG_PASSWORD"
```

The displayed value is a placeholder, not a usable password. **When your password is already configured, do not change it or add a second section.** If it is not configured, set your own password/passphrase of at least 16 characters in Streamlit Secrets, different from the encryption key and GitHub token. The app rejects missing, short, placeholder, and reused encryption-key/token values.

One successful unlock gives access to both OER and PTS in that browser session. Other scheduling sections remain password-free. PTS checks the password before date-preset access, archive reads, usernames, feedback, completion tables, report generation, and download controls.

**Lock OER / PTS** locks both sections and clears their session data, including loaded teaching scans, evaluation data, cached link catalogs, and generated report downloads. It does not delete saved GitHub files, usernames, presets, or any original OPD. Other scheduling sections' session state is preserved. After unlocking again, reload the saved date preset and archived data as needed.

Access expires after 30 minutes without activity in either protected section, or after eight hours regardless of activity. Expiration is checked when a protected section or date-setting callback is used again. Changing the password invalidates existing authorization on its next check. This release also requires a fresh login instead of carrying over authorization issued by the older, narrower gate.

Date-preset callbacks validate authorization separately, because Streamlit executes callbacks before the section is rerendered. A stale date-setting button cannot save, delete, or reload presets after authorization has expired.

## Keep these unchanged

Keep your launcher filename, `schedule_app/settings.py`, custom usernames/email mappings, requirements, all existing `[opd_archive]` settings, and encryption key. No new dependency, repository, permission, credential, archive migration, or re-encryption is needed. OASIS column minimization remains in place.

The old standalone OASIS modules remain on disk for compatibility and are still gated. They are not offered as additional sidebar choices. Do not rename/delete service modules to match the new menu abbreviations.

## Access limits

This is a shared-password gate, not single sign-on, multifactor authentication, individual permissions, or an individual audit trail. It does not revoke documents already downloaded or erase information already delivered to the browser. Session cleanup does not erase Git history. Other scheduling and OPD-archive functions remain outside this gate, as requested.

## Updated runtime files

- `app_sch_2026.py`
- `schedule_app/sections/evaluation_privacy.py`
- `schedule_app/sections/oasis_date_controls.py`
- `schedule_app/sections/oasis_educator_reports.py`
- `schedule_app/sections/oasis_evaluation_archive.py`
- `schedule_app/sections/oasis_student_evaluations.py`
- `schedule_app/sections/oasis_workflow.py`
- `schedule_app/sections/preceptor_oasis_links.py`
- `schedule_app/sections/preceptor_teaching_summary.py`
- `schedule_app/sections/reporting_date_controls.py`
- `schedule_app/services/evaluation_access.py`
- `schedule_app/services/oasis_student_evaluations.py`
- `schedule_app/services/teaching_evaluations.py`

## Validation

The 840-test suite was run in two local batches: **837 passed and 3 skipped**, including **57 new access/navigation tests**. The three skips are historical packaging-comparison checks whose old baseline folders are not part of this release. Existing functional UI tests now explicitly authenticate when exercising a protected section; locked-access tests do not auto-authenticate.

Tests use simulated Streamlit and GitHub calls. The live repository and Streamlit deployment were not accessed or modified. Report modules, settings.py, requirements.txt, and archive/report calculations were checked against the supplied previous release. The protected date callbacks, locked entrypoints, shared login, timeout, password changes, stale-login invalidation, menu order, and old-label migration have explicit tests.

For Streamlit callback order and session behavior, see the official documentation: https://docs.streamlit.io/develop/api-reference/caching-and-state/st.session_state
