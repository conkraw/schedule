"""Sidebar section: Shift Availability Tracker.

Extracted from the supplied app; this module performs no page rendering on import.
"""

from schedule_app.services.availability_analysis import build_segmented_name_map
from schedule_app.services.availability_analysis import daily_caps_three_sites
from schedule_app.services.availability_analysis import fold_hope_drive_rows
from schedule_app.services.availability_analysis import shift_order_for
from schedule_app.services.availability_analysis import weekly_student_capacity
import io
import pandas as pd
import streamlit as st


def render():
    """Render the Shift Availability Tracker sidebar section."""
    # --------------------
    # Helpers: parsing
    # --------------------


    # --------------------
    # UI: title + upload
    # --------------------
    st.title("Shift Availability Tracker")
    
    opd_file = st.file_uploader("Upload md_opd.xlsx", type=["xlsx"])
    if not opd_file:
        st.stop()
    
    excel = pd.ExcelFile(opd_file)
    
    # --------------------
    # Build name-level map + daily counts (segment-aware)
    # --------------------
    name_map = build_segmented_name_map(excel)
    
    rows = []
    for (site, dt, shift), names in name_map.items():
        rows.append({"Site": site, "Date": pd.to_datetime(dt), "Shift": shift, "Names": sorted(names)})
    
    raw = pd.DataFrame(rows)
    
    # Merge HOPE_DRIVE detailed labels into AM/PM for counts and names
    collapsed = []
    for (site, dt), sub in raw.groupby(["Site", "Date"]):
        if site == "HOPE_DRIVE":
            for r in fold_hope_drive_rows(sub):
                collapsed.append({"Site": site, "Date": dt, **r})
        else:
            for _, r in sub.iterrows():
                collapsed.append({
                    "Site": site,
                    "Date": r["Date"],
                    "Shift": r["Shift"],
                    "Names": r["Names"],
                    "Count": len(r["Names"]),
                })
    
    daily = pd.DataFrame(collapsed)
    if daily.empty:
        st.warning("No preceptors with '~' found. Check file/layout.")
        st.stop()
    
    daily["Weekday"] = daily["Date"].dt.weekday              # Mon=0..Sun=6
    daily["DayName"] = daily["Date"].dt.day_name()
    daily["WeekStart"] = daily["Date"] - pd.to_timedelta(daily["Date"].dt.weekday, unit="D")
    
    # --------------------
    # Weekly Grid (single table per site)
    # --------------------
    st.subheader("Weekly Grid (single table per site)")
    
    site_list = sorted(daily["Site"].unique().tolist())
    site_sel = st.selectbox("Site", site_list, index=(site_list.index("HOPE_DRIVE") if "HOPE_DRIVE" in site_list else 0))
    
    day_order = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
    
    if site_sel == "HOPE_DRIVE":
        raw_site = raw[raw["Site"] == "HOPE_DRIVE"].copy()
        raw_site["WeekStart"] = raw_site["Date"] - pd.to_timedelta(raw_site["Date"].dt.weekday, unit="D")
        raw_site["DayName"] = raw_site["Date"].dt.day_name()
        raw_site["DayCat"] = pd.Categorical(raw_site["DayName"], categories=day_order, ordered=True)
        raw_site["ShiftCat"] = pd.Categorical(raw_site["Shift"], categories=shift_order_for("HOPE_DRIVE"), ordered=True)
    
        blocks = []
        for wk, wkdf in raw_site.sort_values(["WeekStart", "ShiftCat", "DayCat"]).groupby("WeekStart"):
            grid = (
                wkdf.assign(Count=wkdf["Names"].apply(len))
                    .pivot_table(index="ShiftCat", columns="DayCat", values="Count", aggfunc="max")
                    .reindex(index=shift_order_for("HOPE_DRIVE"), columns=day_order)
                    .fillna(0).astype(int)
            )
            grid.index.name = "Shift"
            grid.insert(0, "Week of", f"Week of {wk:%Y-%m-%d}")
            blocks.append(grid.reset_index())
        weekly_single_table = pd.concat(blocks, axis=0, ignore_index=True) if blocks else pd.DataFrame()
    else:
        site_df = daily[daily["Site"] == site_sel].copy()
        site_df["DayCat"] = pd.Categorical(site_df["DayName"], categories=day_order, ordered=True)
        site_df["ShiftCat"] = pd.Categorical(site_df["Shift"], categories=shift_order_for(site_sel), ordered=True)
    
        blocks = []
        for wk, wkdf in site_df.sort_values(["WeekStart", "ShiftCat", "DayCat"]).groupby("WeekStart"):
            grid = (
                wkdf.pivot_table(index="ShiftCat", columns="DayCat", values="Count", aggfunc="max")
                    .reindex(index=shift_order_for(site_sel), columns=day_order)
                    .fillna(0).astype(int)
            )
            grid.index.name = "Shift"
            grid.insert(0, "Week of", f"Week of {wk:%Y-%m-%d}")
            blocks.append(grid.reset_index())
        weekly_single_table = pd.concat(blocks, axis=0, ignore_index=True) if blocks else pd.DataFrame()
    
    if weekly_single_table.empty:
        st.info("No data for this site.")
    else:
        st.dataframe(weekly_single_table, use_container_width=True)
        c1, c2 = st.columns(2)
        csv_bytes = weekly_single_table.to_csv(index=False).encode("utf-8")
        c1.download_button("Download Weekly Grid (CSV)", data=csv_bytes,
                           file_name=f"{site_sel}_weekly_grid.csv", mime="text/csv")
        xbuf = io.BytesIO()
        with pd.ExcelWriter(xbuf, engine="xlsxwriter") as writer:
            weekly_single_table.to_excel(writer, sheet_name="Weekly Grid", index=False)
        c2.download_button("Download Weekly Grid (Excel)", data=xbuf.getvalue(),
                           file_name=f"{site_sel}_weekly_grid.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    
    # --------------------
    # Weekly capacity across ETOWN + HOPE_DRIVE + NYES
    # AM = all 5 days (min across Mon–Fri)
    # PM = drop EXACTLY one day (second-smallest across Mon–Fri)
    # --------------------
    st.subheader("Weekly Max Students (ETOWN + HOPE_DRIVE + NYES)")
    
    
    daily_caps = daily_caps_three_sites(daily)
    
    
    weekly_capacity = daily_caps.groupby("WeekStart").apply(weekly_student_capacity).reset_index()
    st.dataframe(weekly_capacity, use_container_width=True)
