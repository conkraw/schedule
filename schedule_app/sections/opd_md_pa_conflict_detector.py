"""Sidebar section: OPD MD PA Conflict Detector.

Extracted from the supplied app; this module performs no page rendering on import.
"""

from schedule_app.services.md_pa_analysis import DEFAULT_FOCUS
from schedule_app.services.md_pa_analysis import _annot_make_copy
from schedule_app.services.md_pa_analysis import build_maps_and_roster
from schedule_app.services.md_pa_analysis import load_sheet
from schedule_app.services.md_pa_analysis import read_sheet_names
import pandas as pd
import streamlit as st


def render():
    """Render the OPD MD PA Conflict Detector sidebar section."""

    st.title("OPD MD/PA Double-Booking & Availability")
    st.write(
        "Upload the MD and PA OPD Excel files to scan for double-booked preceptors and to list availability "
        "by site/date/period (including other sites)."
    )

    # -----------------------------
    # Config & Constants
    # -----------------------------


    # -----------------------------
    # Helpers
    # -----------------------------


    # -----------------------------
    # UI - File uploads
    # -----------------------------
    col1, col2 = st.columns(2)
    with col1:
        md_file = st.file_uploader("Upload MD OPD (xlsx)", type=["xlsx"], key="md")
    with col2:
        pa_file = st.file_uploader("Upload PA OPD (xlsx)", type=["xlsx"], key="pa")

    if md_file and pa_file:
        md_sheets = read_sheet_names(md_file)
        pa_sheets = read_sheet_names(pa_file)

        common_sheets = sorted([s for s in DEFAULT_FOCUS if s in md_sheets and s in pa_sheets])
        selected_sheets = st.multiselect(
            "Sites (tabs) to compare",
            options=sorted(list(set(md_sheets) & set(pa_sheets))),
            default=common_sheets or sorted(list(set(md_sheets) & set(pa_sheets)))
        )

        if not selected_sheets:
            st.warning("No common sheets selected.")
            st.stop()

        # Keep per-site context so we can search "other sites"
        site_ctx = {}

        conflict_rows = []
        diagnostics = []

        for sheet in selected_sheets:
            df_md = load_sheet(md_file, sheet)
            df_pa = load_sheet(pa_file, sheet)

            (md_map_wk, md_idx_date, md_roster_wk, md_day_roster, md_week_dates,
             md_occupied, md_diag) = build_maps_and_roster(df_md)
            (pa_map_wk, pa_idx_date, pa_roster_wk, pa_day_roster, pa_week_dates,
             pa_occupied, pa_diag) = build_maps_and_roster(df_pa)
            diagnostics.append({'site': sheet, 'md': md_diag, 'pa': pa_diag})

            # Save for cross-site availability
            site_ctx[sheet] = dict(
                md_idx_date=md_idx_date,
                pa_idx_date=pa_idx_date,
                md_week_dates=md_week_dates,
                pa_week_dates=pa_week_dates,
                md_day_roster=md_day_roster,
                pa_day_roster=pa_day_roster
            )

            # --------- CONFLICTS by actual date ---------
            md_keys = set(md_idx_date.keys())
            pa_keys = set(pa_idx_date.keys())
            for (date_obj, period, pre) in sorted(md_keys & pa_keys):
                md_entry = md_idx_date[(date_obj, period, pre)]
                pa_entry = pa_idx_date[(date_obj, period, pre)]
                conflict_rows.append({
                    'site': sheet,
                    'date': date_obj,
                    'day': pd.to_datetime(date_obj).strftime('%A'),
                    'period': period,
                    'preceptor': pre,
                    'md_student': md_entry['student'],
                    'pa_student': pa_entry['student']
                })

        # Conflicts dataframe
        conflicts_df = pd.DataFrame(conflict_rows)

        # --------- helpers to build pools ---------
        def pool_for_site_day(site, day_name, period, date_obj):
            """Union of preceptors present in THIS site for (day, period, date)."""
            ctx = site_ctx[site]
            pool = set()
            md_week_dates, pa_week_dates = ctx['md_week_dates'], ctx['pa_week_dates']
            md_day_roster, pa_day_roster = ctx['md_day_roster'], ctx['pa_day_roster']
            # MD
            for m in md_week_dates.keys():
                if md_week_dates[m].get(day_name) == date_obj:
                    pool |= (md_day_roster.get((m, period, day_name)) or set())
            # PA
            for m in pa_week_dates.keys():
                if pa_week_dates[m].get(day_name) == date_obj:
                    pool |= (pa_day_roster.get((m, period, day_name)) or set())
            return pool

        def pool_for_other_sites(current_site, day_name, period, date_obj):
            """Union of preceptors present in ALL OTHER sites for (day, period, date)."""
            pool = set()
            for site in site_ctx.keys():
                if site == current_site:
                    continue
                pool |= pool_for_site_day(site, day_name, period, date_obj)
            return pool

        def count_assigned_any_site(pre, date_obj, period):
            """How many students (MD+PA) does preceptor have across all sites at this date/period?"""
            total = 0
            for s, ctx in site_ctx.items():
                if (date_obj, period, pre) in ctx['md_idx_date']:
                    total += 1
                if (date_obj, period, pre) in ctx['pa_idx_date']:
                    total += 1
            return total

        # --------- AVAILABILITY (same-site & other-sites) for conflict slots ---------
        availability_same_rows = []
        availability_other_rows = []
        suggestions_rows = []

        if not conflicts_df.empty:
            for _, r in conflicts_df.iterrows():
                site   = r['site']
                date_o = r['date']
                day_nm = r['day']         # 'Monday'...'Sunday'
                period = r['period']

                # Pools
                same_pool  = pool_for_site_day(site, day_nm, period, date_o)
                other_pool = pool_for_other_sites(site, day_nm, period, date_o)

                # Build availability function
                def add_pool(pool, dest_list, pool_site_label):
                    for pre in sorted(pool):
                        total_assigned = count_assigned_any_site(pre, date_o, period)
                        is_acute = ("ACUTE" in str(pre).upper())
                        capacity = 2 if is_acute else 1
                        seats_left = max(0, capacity - total_assigned)
                        if seats_left > 0:
                            dest_list.append({
                                'site_of_conflict': site,
                                'candidate_site': pool_site_label,
                                'date': date_o,
                                'day': day_nm,
                                'period': period,
                                'conflict_preceptor': r['preceptor'],
                                'preceptor': pre,
                                'is_acute': is_acute,
                                'current_students': total_assigned,
                                'capacity': capacity,
                                'seats_left': seats_left,
                                'status': 'available'
                            })

                add_pool(same_pool, availability_same_rows, site)
                # For other sites, keep which site each candidate belongs to.
                for other_site in site_ctx.keys():
                    if other_site == site:
                        continue
                    pool = pool_for_site_day(other_site, day_nm, period, date_o)
                    add_pool(pool, availability_other_rows, other_site)

        avail_same_df = pd.DataFrame(availability_same_rows)
        avail_other_df = pd.DataFrame(availability_other_rows)

        # --------- SUGGESTIONS (top-3) prefer same-site, then other-sites ---------
        if not conflicts_df.empty:
            for _, r in conflicts_df.iterrows():
                in_slot_same  = avail_same_df[
                    (avail_same_df['site_of_conflict'] == r['site']) &
                    (avail_same_df['date'] == r['date']) &
                    (avail_same_df['period'] == r['period'])
                ].copy()

                in_slot_other = avail_other_df[
                    (avail_other_df['site_of_conflict'] == r['site']) &
                    (avail_other_df['date'] == r['date']) &
                    (avail_other_df['period'] == r['period'])
                ].copy()

                # Put same preceptor first if eligible (Acute w/ 1 student)
                def order(df):
                    df['_self'] = (df['preceptor'] == r['preceptor'])
                    return df.sort_values(['_self','candidate_site','preceptor'], ascending=[False, True, True]).drop(columns=['_self'])

                ordered = pd.concat([order(in_slot_same), order(in_slot_other)], ignore_index=True)

                if not ordered.empty:
                    for _, a in ordered.head(3).iterrows():
                        label = a['preceptor']
                        if a['preceptor'] == r['preceptor']:
                            label = f"{a['preceptor']} (currently assigned)"
                        suggestions_rows.append({
                            'conflict_site': r['site'],
                            'date': r['date'],
                            'day': r['day'],
                            'period': r['period'],
                            'conflict_preceptor': r['preceptor'],
                            'md_student': r['md_student'],
                            'pa_student': r['pa_student'],
                            'suggested_preceptor': label,
                            'suggested_site': a['candidate_site'],
                            'suggested_is_acute': bool(a['is_acute']),
                            'suggested_current_students': int(a['current_students']),
                            'suggested_capacity': int(a['capacity']),
                            'suggested_seats_left': int(a['seats_left'])
                        })
                else:
                    suggestions_rows.append({
                        'conflict_site': r['site'],
                        'date': r['date'],
                        'day': r['day'],
                        'period': r['period'],
                        'conflict_preceptor': r['preceptor'],
                        'md_student': r['md_student'],
                        'pa_student': r['pa_student'],
                        'suggested_preceptor': '⚠️ No alternative preceptors available',
                        'suggested_site': None,
                        'suggested_is_acute': None,
                        'suggested_current_students': None,
                        'suggested_capacity': None,
                        'suggested_seats_left': None
                    })

        suggestions_df = pd.DataFrame(suggestions_rows)

        # -----------------------------
        # Results UI (conflict-focused)
        # -----------------------------
        st.subheader("Results (conflict-focused)")
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            st.metric("Sites compared", len(selected_sheets))
        with c2:
            st.metric("Double bookings found", 0 if conflicts_df.empty else len(conflicts_df))
        with c3:
            st.metric("Avail. (same-site)", 0 if avail_same_df.empty else len(avail_same_df))
        with c4:
            st.metric("Avail. (other-sites)", 0 if avail_other_df.empty else len(avail_other_df))

        st.markdown("**Double-booked preceptors (MD & PA in same slot)**")
        if conflicts_df.empty:
            st.info("No double-bookings found for the selected sites.")
        else:
            st.dataframe(conflicts_df[['site','date','day','period','preceptor','md_student','pa_student']], use_container_width=True)
            st.download_button(
                label="Download double-bookings CSV",
                data=conflicts_df[['site','date','day','period','preceptor','md_student','pa_student']].to_csv(index=False).encode('utf-8'),
                file_name="opd_double_bookings.csv",
                mime="text/csv"
            )

        # Availability (same site)
        show_same  = st.toggle("Show available preceptors in the SAME site (Acutes can take 2)",value=False, key="tog_same_site")
        if show_same:
            if avail_same_df.empty:
                st.info("No same-site availability for the conflicted slots.")
            else:
                st.markdown("**Available preceptors (same site as conflict)** — Acutes shown if <2 students; others only if unbooked.")
                st.dataframe(
                    avail_same_df[['site_of_conflict','candidate_site','date','day','period',
                                   'conflict_preceptor','preceptor','is_acute',
                                   'current_students','capacity','seats_left','status']],
                    use_container_width=True
                )
                st.download_button(
                    label="Download same-site availability CSV",
                    data=avail_same_df.to_csv(index=False).encode('utf-8'),
                    file_name="opd_availability_same_site.csv",
                    mime="text/csv"
                )

        # Availability (other sites)
        show_other = st.toggle("Show available preceptors in OTHER sites (Acutes can take 2)",value=False, key="tog_other_site")
        if show_other:
            if avail_other_df.empty:
                st.info("No other-site availability for the conflicted slots.")
            else:
                st.markdown("**Available preceptors (other sites)** — same date & AM/PM, different site.")
                st.dataframe(
                    avail_other_df[['site_of_conflict','candidate_site','date','day','period',
                                    'conflict_preceptor','preceptor','is_acute',
                                    'current_students','capacity','seats_left','status']],
                    use_container_width=True
                )
                st.download_button(
                    label="Download other-site availability CSV",
                    data=avail_other_df.to_csv(index=False).encode('utf-8'),
                    file_name="opd_availability_other_sites.csv",
                    mime="text/csv"
                )

        # Suggestions
        show_sugg  = st.toggle("Show suggestions to resolve each conflict (prefers same site, then other sites)",
                       value=False, key="tog_suggestions")
        if show_sugg:
            if suggestions_df.empty:
                st.info("No suggestions available.")
            else:
                st.markdown("**Targeted suggestions** — same site first; if none, suggests from other sites on the same date & AM/PM.")
                st.dataframe(
                    suggestions_df[['conflict_site','date','day','period','conflict_preceptor',
                                    'md_student','pa_student','suggested_preceptor','suggested_site',
                                    'suggested_is_acute','suggested_current_students',
                                    'suggested_capacity','suggested_seats_left']],
                    use_container_width=True
                )
                st.download_button(
                    label="Download suggestions CSV",
                    data=suggestions_df.to_csv(index=False).encode('utf-8'),
                    file_name="opd_targeted_suggestions_cross_site.csv",
                    mime="text/csv"
                )

        # -----------------------------
        # Optional Annotated Downloads Toggle
        # -----------------------------
        show_annotated = st.toggle("Generate annotated OPD files (highlight conflicts in RED)",
                           value=False, key="tog_annotated_downloads")
        
        if show_annotated:
            st.markdown("---")
            st.subheader("Download annotated OPDs (conflicts highlighted in RED)")
            st.caption("Cells are red when the *other* OPD already has that preceptor booked for the same site, date, and AM/PM.")


            # Compare across files
            md_compare_against_pa = {s: site_ctx[s]['pa_idx_date'] for s in site_ctx}
            pa_compare_against_md = {s: site_ctx[s]['md_idx_date'] for s in site_ctx}
            
            col_md, col_pa = st.columns(2)
            with col_md:
                md_bytes = _annot_make_copy(md_file, md_compare_against_pa, selected_sheets)
                st.download_button("⬇️ MD annotated (RED = booked in PA)", md_bytes,
                                   "md_opd_annotated.xlsx",
                                   "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            with col_pa:
                pa_bytes = _annot_make_copy(pa_file, pa_compare_against_md, selected_sheets)
                st.download_button("⬇️ PA annotated (RED = booked in MD)", pa_bytes,
                                   "pa_opd_annotated.xlsx",
                                   "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    else:
        st.info("Upload both the MD and PA OPD files to begin.")
