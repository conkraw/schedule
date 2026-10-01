"""Optional saved OASIS linkage inside PTS.

Only explicit Save/Remove buttons write GitHub. Widget editing alone never writes.
The OPD scan and existing date presets are independent of this catalog.
"""
from __future__ import annotations

import hashlib
import streamlit as st
from schedule_app.services.opd_archive import OPDArchiveError
from schedule_app.services.preceptor_oasis_links import GitHubPreceptorOASISLinks, period_key, LINK_VERSION
from schedule_app.services.oasis_workflow import GitHubOASISSummaries, _filename_dates
from schedule_app.services.oasis_educator_reports import name_key
from schedule_app.services.teaching_evaluations import (
    active_periods, parse_saved_summary, join_feedback, TEACHING_OASIS_REPORT_VERSION,
)

P = "teaching_oasis_"
USERNAME_ENTRY_UI_VERSION = 3
FEEDBACK_PREFERENCE_KEY = P + "feedback_preference"


def _feedback_inclusion_control():
    """Default to saved feedback; keep explicit opt-out independent of widget life.

    This preference holds no evaluation data and is protected/cleared by the
    existing OER/PTS session gate. No new GitHub record or write is needed.
    """
    initial = bool(st.session_state.get(FEEDBACK_PREFERENCE_KEY, True))
    # Initialize the widget before creation. A legacy widget's default False
    # must not silently disable feedback after installing this correction.
    if FEEDBACK_PREFERENCE_KEY not in st.session_state:
        st.session_state.pop(P + "include", None)
    enabled = st.checkbox("Include linked OASIS evaluations in individual Word reports",
                          value=initial, key=P + "include")
    if enabled != initial:
        _clear_downloads()
    st.session_state[FEEDBACK_PREFERENCE_KEY] = bool(enabled)
    return bool(enabled)



def _clear_downloads():
    st.session_state.pop("teaching_zip", None)
    st.session_state.pop("teaching_zip_signature", None)


def _reset_username_controls():
    """Reset drafts before the username widgets are rendered, never afterward."""
    exact = {P + "preceptor_choice", P + "saved_preceptor_choice", P + "edit_saved"}
    prefixes = (P + "username_", P + "confirm_username_", P + "confirm_remove_")
    for key in list(st.session_state):
        if key in exact or key.startswith(prefixes):
            st.session_state.pop(key, None)


def _refresh():
    _reset_username_controls()
    for suffix in ("catalog", "listing", "summaries"):
        st.session_state.pop(P + suffix, None)
    _clear_downloads()


def _cached_summary(service, filename):
    cache = st.session_state.setdefault(P + "summaries", {})
    if filename not in cache:
        cache[filename] = parse_saved_summary(service.load(filename))
    return cache[filename]


def _username_candidates(scan, years, catalog):
    """Return current teaching preceptors and the missing-only entry queue.

    Only named preceptors with a retained student assignment in the selected
    reporting period are eligible, matching the individual-report roster.
    A saved username need not have OASIS evaluations in the current period.
    """
    selected = {int(year) for year in years}
    review = {name_key(name) for name in scan.get("unresolved_preceptor_labels", [])}
    active = {}
    for row in scan["monthly"]:
        name = str(row["preceptor_name"]).strip()
        key = name_key(name)
        if (key and row["academic_start_year"] in selected
                and int(row["no_of_shifts"]) > 0 and key not in review):
            active.setdefault(key, name)
    active = dict(sorted(active.items()))
    missing = {key: name for key, name in active.items()
               if not catalog["entries"].get(key, {}).get("record_id", "").strip()}
    return active, missing


def _username_edited(confirmation_key):
    """A different typed username needs its own confirmation."""
    st.session_state[confirmation_key] = False
    _clear_downloads()


def _username_saved(catalog, message):
    """Advance the entry queue only after GitHub verified the catalog write."""
    st.session_state[P + "catalog"] = catalog
    st.session_state[P + "reset_username_controls"] = True
    st.session_state[P + "notice"] = message
    _clear_downloads()
    # The new queue, count, and all-mapped state must be drawn on the same screen.
    # Widget state is cleared at the START of the next run to avoid changing an
    # already-instantiated Streamlit widget.
    st.rerun()


def _render_username_inputs(service, catalog, name, *, existing=None, editing=False):
    """One explicit username save, with the existing collision/identity checks."""
    existing = existing or {}
    key = name_key(name)
    suffix = hashlib.sha256((key + "\0" + existing.get("record_id", "")).encode()).hexdigest()[:16]
    confirmation_key = P + "confirm_username_" + suffix
    entered = st.text_input(
        "OASIS username / record_id (before @)", value=existing.get("record_id", ""),
        key=P + "username_" + suffix,
        on_change=_username_edited, args=(confirmation_key,),
    )
    confirmed = st.checkbox("I confirm this username belongs to the selected preceptor",
                            value=False, key=confirmation_key)
    ready = True
    label = "Update saved username link in GitHub" if editing else "Save username link to GitHub"
    button_key = P + ("update_username" if editing else "save_username")
    if st.button(label, key=button_key, disabled=not confirmed):
        try:
            saved = service.save_username(name, entered, expected=catalog)
            _username_saved(saved, f"Username for {name} encrypted, saved, and verified in GitHub.")
        except OPDArchiveError as exc:
            _clear_downloads()
            st.error(str(exc))
            ready = False
    if entered.strip().lower() != existing.get("record_id", ""):
        st.info("This username edit is not saved. Save it, or clear/restore the field, before generating linked reports.")
        ready = False
    if editing:
        remove = st.checkbox("Confirm removal of this preceptor's username link", value=False,
                             key=P + "confirm_remove_" + suffix)
        if st.button("Remove saved username link", key=P + "remove_username", disabled=not remove):
            try:
                saved = service.remove_username(name, expected=catalog)
                _username_saved(saved, "Username link removed. A current teaching preceptor will reappear "
                                       "in the missing-username list. OPDs and OASIS reports were not changed.")
            except OPDArchiveError as exc:
                _clear_downloads()
                st.error(str(exc))
                ready = False
    return ready


def _render_username_queue(service, catalog, active, missing, *, show_tables=True):
    """Keep the routine entry dropdown limited to names that still need a key."""
    st.markdown("**1. Add missing preceptor usernames**")
    st.caption("Checks named OPD preceptors with student assignments in the selected dates. "
               "Saved usernames are reused even when that person has no OASIS evaluations in this period. "
               "Preceptors with no student assignments and generic site/slot labels are not in this queue.")
    ready = True
    if missing:
        st.write(f"Usernames still needed: {len(missing):,}")
        if show_tables or st.checkbox("Show missing-username list (optional)", key=P + "show_missing", value=False):
            st.dataframe([{"preceptor_name": name} for name in missing.values()],
                         hide_index=True, use_container_width=True)
        options = list(missing)
        if st.session_state.get(P + "preceptor_choice") not in options:
            st.session_state[P + "preceptor_choice"] = options[0]
        key = st.selectbox("Teaching preceptors needing a username", options,
                           key=P + "preceptor_choice", format_func=lambda key: missing[key])
        ready = _render_username_inputs(service, catalog, missing[key])
    elif active:
        st.info("No usernames need to be entered for these dates. The entry dropdown is hidden because all are saved.")
    else:
        st.info("No individual teaching preceptor names are available to link in these dates.")

    # Existing links remain editable/removable without cluttering the entry queue.
    # Maintenance is deliberately opt-in; the routine dropdown stays missing-only.
    if catalog["entries"]:
        with st.expander("Review or correct saved username links (optional)"):
            ordered = sorted(catalog["entries"])
            if show_tables or st.checkbox("Show saved-username table (optional)", key=P + "show_saved", value=False):
                st.dataframe([
                    {"preceptor_name": catalog["entries"][key]["preceptor_name"],
                     "username": catalog["entries"][key]["record_id"],
                     "in_current_report": "YES" if key in active else "NO"}
                    for key in ordered
                ], hide_index=True, use_container_width=True)
            editing = st.checkbox("Edit or remove an existing username link", value=False, key=P + "edit_saved")
            if editing:
                if st.session_state.get(P + "saved_preceptor_choice") not in ordered:
                    st.session_state[P + "saved_preceptor_choice"] = ordered[0]
                key = st.selectbox("Saved link to correct or remove", ordered,
                                   key=P + "saved_preceptor_choice",
                                   format_func=lambda key: catalog["entries"][key]["preceptor_name"])
                existing = catalog["entries"][key]
                editor_ready = _render_username_inputs(
                    service, catalog, existing["preceptor_name"], existing=existing, editing=True)
                ready = ready and editor_ready
    return ready


def render_teaching_oasis_links(archive, scan, years, *, allow_missing_summaries=False,
                                manage_usernames=True, show_tables=True):
    """Return (plan, ready). With linkage disabled this performs NO network calls."""
    if st.session_state.pop(P + "reset_username_controls", False):
        _reset_username_controls()
    # Keep the alert visible even when the link editor's expander is collapsed.
    username_status = st.empty()
    feedback_status = st.empty()
    with st.expander("Link preceptors to OASIS evaluations (saved in GitHub)"):
        enabled = _feedback_inclusion_control()
        st.caption("Saved username links are applied automatically. Manage missing names or corrections in PTS Matching. "
                   "Only matched preceptors receive the feedback section; teaching-only reports remain available for the others.")
        if not enabled:
            feedback_status.warning("Learner feedback is OFF. Individual reports will omit students' evaluations of preceptors. "
                                    "Enable linked OASIS evaluations above to include saved feedback.")
            return None, True
        scope = (LINK_VERSION, TEACHING_OASIS_REPORT_VERSION, USERNAME_ENTRY_UI_VERSION, archive.config.signature())
        if st.session_state.get(P + "scope") != scope:
            st.session_state.pop(P + "notice", None)
            _refresh()
            st.session_state[P + "scope"] = scope
        if st.button("Refresh links and OASIS summaries", key=P + "refresh"):
            _refresh()
        service, reports = GitHubPreceptorOASISLinks(archive), GitHubOASISSummaries(archive)
        try:
            if P + "catalog" not in st.session_state:
                st.session_state[P + "catalog"] = service.load()
        except OPDArchiveError as exc:
            _clear_downloads()
            username_status.error("Username check unavailable: saved links could not be loaded. Nothing has been assumed missing or overwritten.")
            st.error(str(exc))
            st.info("Refresh to retry. Saved links are not assumed empty. Turn off linked evaluations to create teaching-only reports.")
            return None, False
        catalog = st.session_state[P + "catalog"]
        active, missing = _username_candidates(scan, years, catalog)
        if missing:
            username_status.warning(
                f"Username needed: {len(missing):,} of {len(active):,} teaching preceptors in the selected dates "
                "do not have a saved username. Open PTS Matching → Preceptor usernames to enter only the missing names.")
        elif active:
            username_status.success(f"All {len(active):,} teaching preceptors in the selected dates have a saved username.")
        flash = st.session_state.pop(P + "notice", None)
        if flash:
            st.success(flash)
        ready = _render_username_queue(service, catalog, active, missing, show_tables=show_tables) if manage_usernames else True
        # Username entry must not depend on an OASIS summary being available yet.
        try:
            if P + "listing" not in st.session_state:
                st.session_state[P + "listing"] = reports.list_outputs()
        except OPDArchiveError as exc:
            _clear_downloads()
            st.error(str(exc))
            st.info("Saved usernames remain available above. Refresh links and OASIS summaries to retry the summary list.")
            return None, False
        filenames = st.session_state[P + "listing"]["filenames"]
        summaries = {}
        st.markdown("**2. Choose the saved OASIS summary for this reporting period**")
        st.caption("OASIS is filtered by Submit Date. Teaching hours use assignment dates. The exact start and end dates "
                   "must match; an already-averaged summary cannot be filtered or split into new dates. "
                   "Use OER to generate a matching period when necessary.")
        for year, start, end, label in active_periods(scan, years):
            key = period_key(start, end)
            options = [f for f in filenames if _filename_dates(f) == (start, end)]
            saved = catalog["report_links"].get(key, {}).get("summary_filename")
            st.write(f"{label}: {start:%B %d, %Y} through {end:%B %d, %Y}")
            if not options:
                st.warning("No saved OASIS summary has these exact dates. Open OER, load the same date preset, "
                           "and let it save the summary. Then refresh links here. Username assignments can still be saved above.")
                if not allow_missing_summaries:
                    ready = False
                continue
            choice_key = P + "report_choice_" + key
            if st.session_state.get(choice_key) not in options:
                st.session_state[choice_key] = saved if saved in options else options[0]
            chosen = st.selectbox("Saved OASIS summary", options, key=choice_key,
                format_func=lambda f: f.removesuffix(".csv.enc"),
                help="Different suffixes on the same dates identify different course/evaluation-form selections.")
            confirm_key = P + "confirm_report_" + key
            confirmed = st.checkbox("Replace the previously linked summary for these dates", key=confirm_key) if saved and chosen != saved else True
            if st.button("Save selected summary link to GitHub", key=P + "save_report_" + key,
                         disabled=not confirmed):
                try:
                    # Verify the chosen report can be read before saving the link.
                    _cached_summary(reports, chosen)
                    catalog = service.save_report(start, end, chosen, expected=catalog)
                    st.session_state[P + "catalog"] = catalog
                    saved = chosen
                    _clear_downloads()
                    st.success("Summary link encrypted, saved, and verified. This period will remember its selection.")
                except OPDArchiveError as exc:
                    _clear_downloads(); st.error(str(exc)); ready = False
            if chosen != saved:
                st.info("Save this summary selection before creating linked reports.")
                ready = False
                continue
            try:
                summary = _cached_summary(reports, saved)
                summaries[year] = summary
                d = summary["details"]
                st.caption(f"Saved OASIS label: {d['label']} | {d['educator_count']:,} educators | "
                           f"{d['evaluation_count']:,} submitted evaluations. Only linked teaching preceptors will be used.")
            except OPDArchiveError as exc:
                st.error(str(exc)); ready = False
        if summaries and show_tables:
            with st.expander("Available OASIS usernames (reference only)"):
                reference = sorted({(r["record_id"], r["educator_name"]) for s in summaries.values()
                                    for r in s["rows_by_id"].values()})
                st.dataframe([{"record_id": rid, "oasis_educator_name": name} for rid, name in reference],
                             hide_index=True, use_container_width=True)
                st.caption("This is a reference list, not automatic matching. Extra OASIS-only educators do not receive a teaching report.")
        if not ready:
            _clear_downloads()
            return None, False
        try:
            bundle = join_feedback(scan, years, catalog, summaries, allow_missing_summaries=allow_missing_summaries)
        except OPDArchiveError as exc:
            _clear_downloads(); st.error(str(exc)); return None, False
        if show_tables:
            st.markdown("**Link preview for these teaching reports**")
            st.dataframe(bundle["status"], hide_index=True, use_container_width=True)
        linked = sum(len(group["preceptors"]) for group in bundle["periods"].values())
        total = len(bundle["status"])
        if linked:
            feedback_status.success(f"Learner feedback ready: {linked:,} of {total:,} preceptor-period reports will include "
                                    "students' evaluations, question averages, and comments.")
            if linked < total:
                st.warning(f"{total - linked:,} report(s) have no matching feedback. Their Word documents will explain "
                           "whether a username, matching-date summary, or evaluation row is missing.")
        else:
            feedback_status.warning("No learner feedback is attached yet. Check the saved username and exact-date OASIS summary "
                                    "links below. Each individual report will show the reason rather than silently omitting feedback.")
        st.caption("Saved summaries are snapshots, not a live recomputation of evaluations. Refresh links after OASIS is updated. "
                   "The app checks the saved summary again before generating Word reports. "
                   "Comments remain verbatim and can contain identifying information.")
        return {"catalog": catalog, "summaries": summaries, "bundle": bundle}, True


def render_preceptor_username_management(archive, scan, years):
    """Missing-only editor, independent of the report's feedback checkbox."""
    from schedule_app.services.evaluation_access import evaluation_access_is_valid, lock_evaluation_records
    if not evaluation_access_is_valid(touch=True):
        lock_evaluation_records()
        return
    if st.session_state.pop(P + "reset_username_controls", False):
        _reset_username_controls()
    scope = (LINK_VERSION, TEACHING_OASIS_REPORT_VERSION, USERNAME_ENTRY_UI_VERSION, archive.config.signature())
    if st.session_state.get(P + "scope") != scope:
        _refresh()
        st.session_state[P + "scope"] = scope
    if st.button("Refresh saved preceptor usernames", key=P + "matching_refresh"):
        _refresh()
    service = GitHubPreceptorOASISLinks(archive)
    try:
        if P + "catalog" not in st.session_state:
            st.session_state[P + "catalog"] = service.load()
    except OPDArchiveError as exc:
        st.error(str(exc))
        st.warning("Username checks are unavailable until the saved links can be read. No links were erased.")
        return
    catalog = st.session_state[P + "catalog"]
    active, missing = _username_candidates(scan, years, catalog)
    if missing:
        st.warning(f"Username needed: {len(missing):,} of {len(active):,} teaching preceptors. Only missing names appear below.")
    elif active:
        st.success(f"All {len(active):,} teaching preceptors have a saved username.")
    notice = st.session_state.pop(P + "notice", None)
    if notice:
        st.success(notice)
    if not _render_username_queue(service, catalog, active, missing, show_tables=False):
        _clear_downloads()
    if st.checkbox("Show OASIS usernames for reference (optional)", key=P + "matching_reference", value=False):
        # An explicit request, not a download on every edit. Existing verified cache is reused.
        try:
            service_reports = GitHubOASISSummaries(archive)
            reference = set()
            for year, start, end, _ in active_periods(scan, years):
                binding = catalog["report_links"].get(period_key(start, end))
                if binding:
                    summary = _cached_summary(service_reports, binding["summary_filename"])
                    reference.update((r["record_id"], r["educator_name"]) for r in summary["rows_by_id"].values())
            if reference:
                st.dataframe([{"username": rid, "OASIS educator": name} for rid, name in sorted(reference)],
                             hide_index=True, use_container_width=True)
            else:
                st.info("Choose and save an exact-date OASIS summary link in PTS to make its usernames available here.")
        except OPDArchiveError as exc:
            st.error(str(exc))
