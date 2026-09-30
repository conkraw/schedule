# Missing preceptor usernames — focused entry queue

## Install the small update

This update changes **one runtime file**:

```text
schedule_app/sections/preceptor_oasis_links.py
```

1. Extract `Schedule_App_Missing_Usernames_Update.zip`.
2. Replace that one file at the matching path in your app repository. Merge the supplied folders; do not delete your existing `schedule_app` folder.
3. Restart the Streamlit app.

Keep `app_sch_2026.py`, `schedule_app/settings.py`, email/name mappings, `requirements.txt`, Streamlit Secrets, the encryption key, date presets, and all other files unchanged. No new dependencies, archive migration, re-encryption, or OPD uploads are needed.

The complete modular-app ZIP is also supplied for a clean installation. Use the small update to preserve any additional changes in your deployed app.

## What you will see

In **Preceptor Teaching Summary**, load the OPDs and select your reporting dates as usual. Open **Link preceptors to OASIS evaluations (saved in GitHub)** and enable **Include linked OASIS evaluations in individual Word reports**.

When the saved links are loaded, an alert above that panel states how many teaching preceptors still need a username. The alert remains visible when the panel is collapsed.

For example:

> Username needed: 3 of 12 teaching preceptors in the selected dates do not have a saved username.

Under **1. Add missing preceptor usernames**, the dropdown **Teaching preceptors needing a username** contains only those missing names, alphabetically. **Show names missing a username** provides the full pending list without requiring a download.

Choose a name, enter the username used as `record_id` in the OASIS summary, confirm the identity, and click **Save username link to GitHub**. The existing encryption/save/verification service is reused. After a successful save, the app redraws the screen, removes the completed name from the entry dropdown, updates the count, and clears the entry/confirmation for the next person. A failed, unverified, or conflicting save does not remove the name from the pending list.

Once all included preceptors have usernames, the alert becomes a success message and the empty entry dropdown is hidden.

Username entry is now shown before the OASIS summary selection. You may save missing usernames even when a matching summary has not been generated yet. The date-matched summary selection and report preview otherwise work as before.

## Which OPD names are checked

The check follows the existing report roster: **named preceptors with at least one retained student assignment during the selected reporting period**. It does not reintroduce preceptors with no student assignments, preceptors outside the selected dates, or generic site/slot labels. Normalization and the existing outpatient-over-nursery priority remain unchanged.

A preceptor with a saved username but no evaluation in the selected OASIS summary **does not** need a new username and therefore stays out of the missing-entry dropdown. The match-status table continues to distinguish that situation from an absent username. OASIS-only educators do not enter the OPD preceptor queue.

Missing usernames remain warnings, not a new mandatory block on all teaching reports. As before, unmatched preceptors can receive teaching-only reports, while linked preceptors receive evaluation content. An unsaved typed edit must still be saved or cleared before generating linked reports.

## Correct or remove an existing link

Existing links remain encrypted in the same catalog:

```text
<your configured archive folder>/preceptor_oasis_links.json.enc
```

To correct an existing mapping, expand **Review or correct saved username links (optional)** and check **Edit or remove an existing username link**. Only then is the separate maintenance selector shown. Choose the saved link, correct and confirm its username, then click **Update saved username link in GitHub**.

Removal still requires confirmation. Removing a mapping for a current teaching preceptor returns that name to the missing-entry dropdown. Saved links outside the current report remain manageable here without cluttering the routine entry list. No OPD, OASIS source, summary, or date preset is deleted.

**Refresh links and OASIS summaries** loads mappings changed in another session. It does not redownload the OPDs. If the saved catalog cannot be read or decrypted, the app reports that the username check is unavailable rather than assuming every name is missing or overwriting the catalog.

When the optional OASIS linkage is off, it still makes no OASIS link/catalog network requests. Enable it to check saved usernames. Username changes clear the old generated report download so the next report uses the updated links; already-loaded OPD data and your chosen reporting dates remain available.

## Unchanged

- Total scheduled availability, educational hours, and Learner Reach calculations.
- Four educational hours for a shift with one or more students; no multiplier for simultaneous students.
- Work-type groups, weekends, custom dates, and Academic Pediatrics-over-PSHCH Nursery priority.
- Unique students and students assigned on three or more distinct dates.
- Chair and individual Word-report layouts, OASIS full question text and comments.
- OASIS upload/summary workflow, encrypted archives, username collision checks, and date presets.
- Existing deployment access controls; no app password is added.

## Validation

The local suite completed with **613 passed and 1 skipped**, including **34 new missing-username tests**. The suite also reported 140 passing subtests. Two earlier UI tests were adjusted for the intentional immediate redraw after saving and the new opt-in maintenance controls. The prior immutable-file check was narrowed to exclude this intentionally changed UI module; all calculation, encryption, persistence, and report-writing runtime files remain byte-for-byte unchanged from the supplied Simple Hours package.

The tests cover missing-only selection, alerts, save/advance, completion, removal, correction, duplicate usernames, stale saves, failed catalog reads, failed summary listing, existing usernames without current evaluations, name normalization, retained OPD/date settings, and report-download invalidation.

GitHub and Streamlit interactions were simulated. A native Streamlit UI test was not run: Streamlit was not installed in this runtime and installation could not reach the package server. No live repository or deployment was accessed or changed.

Implementation references: Streamlit rerun and widget state behavior were checked against the official documentation:
- https://docs.streamlit.io/develop/api-reference/execution-flow/st.rerun
- https://docs.streamlit.io/develop/api-reference/caching-and-state/st.session_state
- https://docs.streamlit.io/develop/api-reference/widgets/st.selectbox
