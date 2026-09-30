"""Optional saved OASIS linkage inside Preceptor Teaching Summary.

Only explicit Save/Remove buttons write GitHub. Widget editing alone never writes.
The OPD scan and existing date presets are independent of this catalog.
"""
from __future__ import annotations

import hashlib
import streamlit as st
from schedule_app.services.opd_archive import OPDArchiveError
from schedule_app.services.preceptor_oasis_links import GitHubPreceptorOASISLinks, period_key, LINK_VERSION
from schedule_app.services.oasis_workflow import GitHubOASISSummaries, _filename_dates
from schedule_app.services.oasis_educator_reports import name_key, validate_username
from schedule_app.services.teaching_evaluations import (
    active_periods, parse_saved_summary, join_feedback, TEACHING_OASIS_REPORT_VERSION,
)

P = "teaching_oasis_"


def _clear_downloads():
    st.session_state.pop("teaching_zip", None)
    st.session_state.pop("teaching_zip_signature", None)


def _refresh():
    for suffix in ("catalog", "listing", "summaries"):
        st.session_state.pop(P + suffix, None)
    _clear_downloads()


def _cached_summary(service, filename):
    cache = st.session_state.setdefault(P + "summaries", {})
    if filename not in cache:
        cache[filename] = parse_saved_summary(service.load(filename))
    return cache[filename]


def render_teaching_oasis_links(archive, scan, years):
    """Return (plan, ready). With linkage disabled this performs NO network calls."""
    with st.expander("Link preceptors to OASIS evaluations (saved in GitHub)"):
        enabled = st.checkbox("Include linked OASIS evaluations in individual Word reports",
                              value=False, key=P + "include")
        st.caption("Assign each teaching preceptor the username used as record_id in your OASIS output. "
                   "Only explicit matches receive an evaluation section. Other teaching reports remain teaching-only; "
                   "OASIS-only educators are ignored. No names are matched automatically.")
        if not enabled:
            return None, True
        scope = (LINK_VERSION, TEACHING_OASIS_REPORT_VERSION, archive.config.signature())
        if st.session_state.get(P + "scope") != scope:
            _refresh()
            st.session_state[P + "scope"] = scope
        if st.button("Refresh links and OASIS summaries", key=P + "refresh"):
            _refresh()
        service, reports = GitHubPreceptorOASISLinks(archive), GitHubOASISSummaries(archive)
        try:
            if P + "catalog" not in st.session_state:
                st.session_state[P + "catalog"] = service.load()
            if P + "listing" not in st.session_state:
                st.session_state[P + "listing"] = reports.list_outputs()
        except OPDArchiveError as exc:
            _clear_downloads()
            st.error(str(exc))
            st.info("Refresh to retry. Saved links are not assumed empty. Turn off linked evaluations to create teaching-only reports.")
            return None, False
        catalog = st.session_state[P + "catalog"]
        filenames = st.session_state[P + "listing"]["filenames"]
        ready = True
        summaries = {}
        st.markdown("**1. Choose the saved OASIS summary for this reporting period**")
        st.caption("OASIS is filtered by Submit Date. Teaching hours use assignment dates. The exact start and end dates "
                   "must match; an already-averaged summary cannot be filtered or split into new dates. "
                   "Use OASIS Evaluations to generate a matching period when necessary.")
        for year, start, end, label in active_periods(scan, years):
            key = period_key(start, end)
            options = [f for f in filenames if _filename_dates(f) == (start, end)]
            saved = catalog["report_links"].get(key, {}).get("summary_filename")
            st.write(f"{label}: {start:%B %d, %Y} through {end:%B %d, %Y}")
            if not options:
                st.warning("No saved OASIS summary has these exact dates. Open OASIS Evaluations, load the same date preset, "
                           "and let it save the summary. Then refresh links here. Username assignments can still be saved below.")
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
        review = {name_key(n) for n in scan.get("unresolved_preceptor_labels", [])}
        active = {name_key(r["preceptor_name"]): r["preceptor_name"] for r in scan["monthly"]
                  if r["academic_start_year"] in {int(y) for y in years} and int(r["no_of_shifts"]) > 0
                  and name_key(r["preceptor_name"]) not in review}
        # Keep old saved preceptors editable/removable even when absent this year.
        names = {**{k: v["preceptor_name"] for k, v in catalog["entries"].items()}, **active}
        st.markdown("**2. Assign an individual preceptor's username**")
        if names:
            options = sorted(names, key=lambda k: (k not in active, k))
            if st.session_state.get(P + "preceptor_choice") not in options:
                st.session_state[P + "preceptor_choice"] = options[0]
            key = st.selectbox("Teaching preceptor", options, key=P + "preceptor_choice",
                format_func=lambda k: names[k] + (" (saved link; outside current report)" if k not in active else ""))
            name, existing = names[key], catalog["entries"].get(key, {})
            # A new key after a saved change avoids stale values without mutating
            # a widget after its instantiation. Merely editing does not save.
            suffix = hashlib.sha256((key + "\0" + existing.get("record_id", "")).encode()).hexdigest()[:16]
            entered = st.text_input("OASIS username / record_id (before @)", value=existing.get("record_id", ""),
                                     key=P + "username_" + suffix)
            confirmed = st.checkbox("I confirm this username belongs to the selected preceptor",
                                      value=False, key=P + "confirm_username_" + suffix)
            if st.button("Save username link to GitHub", key=P + "save_username", disabled=not confirmed):
                try:
                    catalog = service.save_username(name, entered, expected=catalog)
                    st.session_state[P + "catalog"] = catalog
                    existing = catalog["entries"][key]
                    _clear_downloads()
                    st.success("Preceptor username link encrypted, saved, and verified in GitHub.")
                except OPDArchiveError as exc:
                    _clear_downloads(); st.error(str(exc)); ready = False
            if entered.strip().lower() != existing.get("record_id", ""):
                st.info("This username edit is not saved. Save it, or restore the saved value, before generating linked reports.")
                ready = False
            if key in catalog["entries"]:
                remove = st.checkbox("Confirm removal of this preceptor's username link", value=False,
                                     key=P + "confirm_remove_" + suffix)
                if st.button("Remove saved username link", key=P + "remove_username", disabled=not remove):
                    try:
                        catalog = service.remove_username(name, expected=catalog)
                        st.session_state[P + "catalog"] = catalog
                        _clear_downloads()
                        st.success("Username link removed. OPDs and OASIS reports were not changed.")
                    except OPDArchiveError as exc:
                        _clear_downloads(); st.error(str(exc)); ready = False
        else:
            st.info("No individual teaching preceptor names are available to link in these dates.")
        if summaries:
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
            bundle = join_feedback(scan, years, catalog, summaries)
        except OPDArchiveError as exc:
            _clear_downloads(); st.error(str(exc)); return None, False
        st.markdown("**Link preview for these teaching reports**")
        st.dataframe(bundle["status"], hide_index=True, use_container_width=True)
        linked = sum(len(group["preceptors"]) for group in bundle["periods"].values())
        if linked:
            st.success(f"{linked:,} preceptor-period match(es). Their individual Word reports will include the linked OASIS evaluations.")
        else:
            st.warning("No teaching preceptors have a matching saved username in the selected OASIS summaries. Reports will contain teaching effort only until a match is saved.")
        st.caption("Saved summaries are snapshots, not a live recomputation of evaluations. Refresh links after OASIS is updated. "
                   "The app checks the saved summary again before generating Word reports. "
                   "Comments remain verbatim and can contain identifying information.")
        return {"catalog": catalog, "summaries": summaries, "bundle": bundle}, True
