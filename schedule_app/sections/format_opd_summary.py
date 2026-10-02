"""Sidebar section: Format OPD + Summary.

Extracted from the supplied app; this module performs no page rendering on import.
"""

from schedule_app.services.report_wording import report_text as rt, with_report_wording

from docx import Document
from docx.enum.section import WD_ORIENT
from schedule_app.services.opd_workbooks import generate_opd_workbook
from schedule_app.services.opd_workbooks import hide_blank_rows_all_sheets
from schedule_app.services.opd_workbooks import update_excel_from_csv
import io
import pandas as pd
import random
import re
import streamlit as st
from schedule_app.sections.report_wording_controls import with_saved_report_wording
from schedule_app.reports.report_appearance import apply_report_appearance, add_report_message
import zipfile


@with_saved_report_wording
def render():
    """Render the Format OPD + Summary sidebar section."""
    # ─── Inputs ────────────────────────────────────────────────────────────────────
    # Required keywords to look for in the content
    #required_keywords = ["academic general pediatrics", "hospitalists", "complex care", "adol med"]
    required_keywords = ["academic general pediatrics", "hospitalists", "complex care"]
    found_keywords = set()
    
    schedule_files = st.file_uploader("1) Upload one or more QGenda calendar Excel(s)",type=["xlsx", "xls"],accept_multiple_files=True)
    
    #if schedule_files:
    #    for file in schedule_files:
    #        try:
    #            # Read the first sheet
    #            df = pd.read_excel(file, sheet_name=0, header=None)
    
    #            # Flatten all string values to a list of lowercase strings
    #            cell_values = df.astype(str).apply(lambda x: x.str.lower()).values.flatten().tolist()
    
    #            # Check if any keyword is found in cell values
    #            for keyword in required_keywords:
    #                if any(keyword in val for val in cell_values):
    #                    found_keywords.add(keyword)
    
    #       except Exception as e:
    #           st.error(f"Error reading {file.name}: {e}")

    if schedule_files:
        for file in schedule_files:
            try:
                df = pd.read_excel(file, sheet_name=0, header=None)
    
                cell_values = [
                    str(val).strip().lower()
                    for row in df.values
                    for val in row
                    if pd.notna(val)
                ]
    
                for keyword in required_keywords:
                    if any(keyword in val for val in cell_values):
                        found_keywords.add(keyword)
    
            except Exception as e:
                st.error(f"Error reading {file.name}: {e}")
    
        # Identify missing calendars
        missing_keywords = [k for k in required_keywords if k not in found_keywords]
    
        if missing_keywords:
            st.warning(f"Missing required calendar(s): {', '.join(missing_keywords)}. Please upload all required calendars.")
        else:
            st.success("All required calendars uploaded and verified by content.")
    
    student_file = st.file_uploader("2) Upload Redcap Rotation list CSV (must have a 'legal_name' and 'start_date' column)",type=["csv"])
    
    #record_id = st.text_input("3) Enter the REDCap record_id for this batch", "")
    
    record_id = "peds_clerkship"
    
    # ─── Guard ─────────────────────────────────────────────────────────────────────
    if not schedule_files or not student_file or not record_id:
        st.info("Please upload schedule Excel(s), student CSV")
        st.stop()
    
    # ─── Prep: Date regex & maps ───────────────────────────────────────────────────
    
    date_pat = re.compile(r'^[A-Za-z]+ \d{1,2}, \d{4}$')
    base_map = {
        "hope drive am continuity":    "hd_am_",
        "hope drive pm continuity":    "hd_pm_",
        
        "hope drive am acute precept": "hd_am_acute_",
        "hope drive pm acute precept": "hd_pm_acute_",
    
        "hope drive weekend acute 1": "hd_wknd_acute_1_", # Changed prefix
        "hope drive weekend acute 2": "hd_wknd_acute_2_", # Changed prefix
    
        "hope drive weekend continuity": "hd_wknd_am_",
        
        "etown am continuity":         "etown_am_",
        "etown pm continuity":         "etown_pm_",
        
        "nyes rd am continuity":       "nyes_am_",
        "nyes rd pm continuity":       "nyes_pm_",
        
        "nursery weekday 8a-6p":       ["nursery_am_", "nursery_pm_"],
        
        "rounder 1 7a-7p":             ["ward_a_am_","ward_a_pm_"],
        "rounder 2 7a-7p":             ["ward_a_am_","ward_a_pm_"],
        "rounder 3 7a-7p":             ["ward_a_am_","ward_a_pm_"],
    
        "hope drive clinic am":        "complex_am_",
        "hope drive clinic pm":        "complex_pm_",
        
        "briarcrest clinic am":       "adol_med_am_",
        "briarcrest clinic pm":       "adol_med_pm_",

        "lancaster am":       "lancaster_am_",
        "lancaster pm":       "lancaster_pm_",
    
    }
    
    # Which groups need at least 2 providers?
    min_required = {
        "hope drive am acute precept": 2,
        "hope drive pm acute precept": 2,
        
        "nursery weekday 8a-6p":       2,
        
        "rounder 1 7a-7p":             2,
        "rounder 2 7a-7p":             2,
        "rounder 3 7a-7p":             2,
    }
    
    file_configs = {"HAMPDEN_NURSERY.xlsx": {"title": "HAMPDEN_NURSERY","custom_text": "CUSTOM_PRINT","names": ["Folaranmi, Oluwamayoda","Alur, Pradeep","Nanda, Sharmilarani","HAMPDEN_NURSERY"]},
                    "SJR_HOSP.xlsx": {"title": "SJR_HOSPITALIST","custom_text": "CUSTOM_PRINT","names": ["Spangola, Haley","Gubitosi, Terry","SJR_1","SJR_2"]},
                    "AAC.xlsx": {"title": "AAC","custom_text": "CUSTOM_PRINT","names": ["Vaishnavi Harding","Abimbola Ajayi","Shilu Joshi","Desiree Webb","Amy Zisa","Abdullah Sakarcan","Anna Karasik","AAC_1","AAC_2","AAC_3",]},
                    "LANCASTER_CMG.xlsx": {"title": "LANCASTER_CMG","custom_text": "CUSTOM_PRINT","names": ["Ashleigh Sobotka","Susannah Christman"]},
                    "MAHOUSSI_AHOLOUKPE.xlsx": {"title": "MAHOUSSI_AHOLOUKPE","custom_text": "CUSTOM_PRINT","names": ["Mahoussi Aholoukpe"]},
                    #"REPLACE.xlsx": {"title": "REPLACE","custom_text": "CUSTOM_PRINT","names": ["ReplaceFirstName ReplaceLastName"]},
                   }
    
    # ─── HERE: generate sheet‐specific custom_print entries for the configss...  ────────────────────
    for cfg in file_configs.values():
        sheet = cfg["title"]              # e.g. "HAMPDEN_NURSERY"
        key   = sheet.lower() + "_print"  # e.g. "hampden_nursery_print"
        prefix = f"{cfg['custom_text'].lower()}_{sheet.lower()}_"
        base_map[key] = prefix
        
    # ─── 1. Aggregate schedule assignments by date ────────────────────────────────
    assignments_by_date = {}
    for file in schedule_files:
        df = pd.read_excel(file, header=None, dtype=str)
    
        # find all date cells
        date_positions = []
        for r in range(df.shape[0]):
            for c in range(df.shape[1]):
                val = str(df.iat[r,c]).replace("\xa0"," ").strip()
                if date_pat.match(val):
                    try:
                        d = pd.to_datetime(val).date()
                        date_positions.append((d,r,c))
                    except:
                        pass
    
        # dedupe to the topmost row per date
        unique = {}
        for d,r,c in date_positions:
            if d not in unique or r < unique[d][0]:
                unique[d] = (r,c)
    
        # before the loop, define:
        day_names = {"monday","tuesday","wednesday","thursday","friday","saturday","sunday"}
        
        # collect providers under each date
        for d, (row0,col0) in unique.items():
            grp = assignments_by_date.setdefault(d, {des:[] for des in base_map})
            
            for r in range(row0+1, df.shape[0]):
                raw = str(df.iat[r, col0]).replace("\xa0", " ").strip()
                # stop if we hit a blank row
                if raw == "":
                    break
                # stop if we hit another date header
                if date_pat.match(raw):
                    break
    
                desc = raw.lower()
                prov = str(df.iat[r, col0+1]).strip()
                if desc in grp and prov:
                    grp[desc].append(prov)
    
    # ─── Provider filter UI ──────────────────────────────────────────────────────
    all_providers = sorted({
        p.strip()
        for day in assignments_by_date.values()
        for provs in day.values()
        for p in provs
        if isinstance(p, str) and p.strip()
    })
    
    # Multiselect persists in session; start empty by design
    if "provider_filter" not in st.session_state:
        st.session_state["provider_filter"] = []
    
    col1, col2, col3 = st.columns(3)
    with col1:
        if st.button("Select All Providers", key="prov_select_all"):
            st.session_state["provider_filter"] = all_providers
    with col2:
        if st.button("Clear Providers", key="prov_clear_all"):
            st.session_state["provider_filter"] = []
    with col3:
        # Switch to actually apply the filter. Off = treat as 'All'
        apply_provider_filter = st.checkbox(
            "Apply provider filter",
            value=False,
            key="prov_apply_filter",
            help="When OFF, everyone is included even if the multiselect is blank."
        )
    
    allowed_providers = st.multiselect(
        "Limit providers included in OPD",
        options=all_providers,
        key="provider_filter",
        help="Only selected providers will be written when 'Apply provider filter' is ON.",
    )
    
    # Effective allow-list:
    effective_allowed = (
        set(allowed_providers) if (apply_provider_filter and allowed_providers) else set(all_providers)
    )

    # ─── 2. Read student list and prepare s1, s2, … ───────────────────────────────
    students_df = pd.read_csv(student_file, dtype=str)
    legal_names = students_df["legal_name"].dropna().tolist()
    
    # ─── 3. Build the single REDCap row ───────────────────────────────────────────
    redcap_row = {"record_id": record_id}
    sorted_dates = sorted(assignments_by_date.keys())
    
    for idx, date in enumerate(sorted_dates, start=1):
        redcap_row[f"hd_day_date{idx}"] = date
        suffix = f"d{idx}_"
    
        # build day‑specific prefixes
        des_map = {
            des: ([p + suffix for p in prefs] if isinstance(prefs, list)
                  else [prefs + suffix])
            for des, prefs in base_map.items()
        }
    
        # 3a) schedule providers (respect provider filter)
        for des, provs in assignments_by_date[date].items():
            # Do not mutate the original list
            filtered = [p for p in provs if p in effective_allowed]
        
            # If the group has a minimum requirement, pad by repeating the first allowed provider
            req = min_required.get(des, len(filtered))
            if filtered and len(filtered) < req:
                filtered = filtered + [filtered[0]] * (req - len(filtered))
        
            # If nothing allowed and no minimum → skip write
            if not filtered:
                continue
        
            if des.startswith("rounder"):
                # rounder N 7a-7p → slot math
                # NOTE: req here is the number of providers per team (usually 2)
                team_idx = int(des.split()[1]) - 1  # 0-based team index
                for i, name in enumerate(filtered, start=1):
                    slot = team_idx * req + i  # team1→1..req, team2→req+1..2*req, etc.
                    for prefix in des_map[des]:
                        redcap_row[f"{prefix}{i if prefix.endswith('_am_') or prefix.endswith('_pm_') else slot}"] = name
                        # ^ If your rounder prefixes are lists like ["ward_a_am_","ward_a_pm_"],
                        #   they'll be in des_map[des] already; the index logic above preserves slots.
            else:
                for i, name in enumerate(filtered, start=1):
                    for prefix in des_map[des]:
                        redcap_row[f"{prefix}{i}"] = name

        # 3b) custom_print names — once per date, using the SAME suffix
        for fname, cfg in file_configs.items():
            sheet = cfg["title"]          
            key   = sheet.lower() + "_print"
            prefix = base_map[key]       # e.g. "custom_print_hampden_nursery_"
            
            for i, person in enumerate(cfg["names"], start=1):
                # note the suffix goes BEFORE the slot index
                redcap_row[f"{prefix}{suffix}{i}"] = person
                
    # append student slots s1,s2,...
    for i,name in enumerate(legal_names, start=1):
        redcap_row[f"s{i}"] = name
    
    # ─── 4. Display & slice out dates/am/acute and students ─────────────────────
    out_df = pd.DataFrame([redcap_row])
    
    # 1) shuffle
    students = legal_names.copy()
    random.shuffle(students)
    
    # 2) define slot sequence
    slot_seq = [1, 3, 5, 2, 4, 6]
    
    # 3) assign
    ward_a_assignment = {}
    
    for idx, student in enumerate(students):
        slot_group = idx // 4                  # every 4 students move to next slot
        slot       = slot_seq[slot_group % len(slot_seq)]
        week_idx   = idx % 4                   # 0→week1,1→week2,2→week3,3→week4
    
        ward_a_assignment[student] = week_idx
        
        # for their week, each Mon–Fri (days 1–5 + 7*week_idx)
        for day in range(1, 6):
            day_num = day + 7 * week_idx
            for shift in ("am", "pm"):
                key  = f"ward_a_{shift}_d{day_num}_{slot}"
                orig = redcap_row.get(key, "")
                redcap_row[key] = f"{orig} ~ {student}" if orig else f"~ {student}"
                
    # ─── track who’s already grabbed a nursery slot ─────────────────────────────
    nursery_assigned = set()
    
    # ─── HAMPDEN_NURSERY: max 1 student for week1 and 1 for week3, into slot _4 ##FOCUSES ON SLOT 4!!! ─────
    for week_idx in (0, 2):  # 0→week1, 2→week3
        pool = [
            s for s in legal_names
            if s not in nursery_assigned
            and ward_a_assignment.get(s, -1) != week_idx
        ]
        if not pool:
            continue
        student = random.choice(pool)
        nursery_assigned.add(student)    # ← mark them as “used”!
    
        for day in range(1, 6):
            d   = day + 7 * week_idx
            key = f"custom_print_hampden_nursery_d{d}_4"        
            orig = redcap_row.get(key, "")
            redcap_row[key] = f"{orig} ~ {student}" if orig else f"~ {student}"
    
    # ─── 2) SJR_HOSPITALIST (max 2 students, any weeks ≠ their Ward A week) ─────
    for week_idx in range(4):  # 0→wk1,1→wk2,2→wk3,3→wk4
        # build pool excluding Hampden and anyone on Ward A that week
        pool = [
            s for s in legal_names
            if s not in nursery_assigned
            and ward_a_assignment.get(s, -1) != week_idx
        ]
        random.shuffle(pool)
        # assign up to two students: first to slot 3, next to slot 4
        for slot_idx in (3, 4):
            if not pool:
                break
            student = pool.pop()
            nursery_assigned.add(student)
            # Mon–Fri of this week
            for day in range(1, 6):
                d   = day + 7 * week_idx
                key = f"custom_print_sjr_hospitalist_d{d}_{slot_idx}"
                orig = redcap_row.get(key, "")
                redcap_row[key] = f"{orig} ~ {student}" if orig else f"~ {student}"
    
    
    # ─── 3) PSHCH_NURSERY (everyone else, up to 8 slots: slot1 weeks1–4, then slot2 wks1–4) ─────────
    leftovers = [s for s in legal_names if s not in nursery_assigned]
    # build (week_idx, slot) in the desired order
    psch_slots = [(wk,1) for wk in range(4)] + [(wk,2) for wk in range(4)]
    for student in leftovers:
        for wk, slot in psch_slots:
            # skip if conflicts with Ward A week
            if ward_a_assignment.get(student, -1) == wk:
                continue
            # build key once (AM & PM) to test existence and avoid duping
            key_am = f"nursery_am_d{day}_ {slot}"
            # assign across Mon–Fri
            for day in range(1, 6):
                d = day + wk * 7
                for prefix in ("nursery_am_","nursery_pm_"):
                    key  = f"{prefix}d{d}_{slot}"
                    orig = redcap_row.get(key, "")
                    redcap_row[key] = f"{orig} ~ {student}" if orig else f"~ {student}"
            # remove this slot so no one else uses it
            psch_slots.remove((wk,slot))
            nursery_assigned.add(student)
            break
        # if no slot left, the student remains unassigned in PSHCH_NURSERY
    
    # format date columns
    for c in out_df.columns:
        if c.startswith("hd_day_date"):
            out_df[c] = pd.to_datetime(out_df[c]).dt.strftime("%m-%d-%Y")
    
    
    out_df = pd.DataFrame([redcap_row])
    csv_full = out_df.to_csv(index=False).encode("utf-8")
    
    
    excel_bytes = generate_opd_workbook(out_df)
    #st.download_button(label="⬇️ Download OPD.xlsx",data=excel_bytes,file_name="OPD.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


    # --- MODIFIED update_excel_from_csv function to work with bytes ---
            
    # --- Configuration for update_excel_from_csv (your mappings) ---
    data_mappings        = []
    excel_column_letters = ['B','C','D','E','F','G','H']
    num_weeks            = 4
    
    # HOPE_DRIVE acute + continuity row offsets
    hd_row_defs = {
        'AM': {'acute_start': 6,  'cont_start': 8},
        'PM': {'acute_start': 16, 'cont_start': 18},
    }
    
    # Other sheets only need continuity (rows 6–13 for AM, 16–23 for PM)
    cont_row_defs = {
        'AM':  6,
        'PM': 16,
    }
    
    # your prefix map
    base_map = {
        "hope drive am continuity":      "hd_am_",
        "hope drive pm continuity":      "hd_pm_",
        "hope drive am acute precept":   "hd_am_acute_",
        "hope drive pm acute precept":   "hd_pm_acute_",
        "hope drive weekend acute 1":    "hd_wknd_acute_1_",
        "hope drive weekend acute 2":    "hd_wknd_acute_2_",
        "hope drive weekend continuity": "hd_wknd_am_",
        
        "etown am continuity":           "etown_am_",
        "etown pm continuity":           "etown_pm_",
        
        "nyes rd am continuity":         "nyes_am_",
        "nyes rd pm continuity":         "nyes_pm_",
        
        "nursery weekday 8a-6p":         ["nursery_am_","nursery_pm_"],
        
        "rounder 1 7a-7p":               ["ward_a_am_","ward_a_pm_"],
        "rounder 2 7a-7p":               ["ward_a_am_","ward_a_pm_"],
        "rounder 3 7a-7p":               ["ward_a_am_","ward_a_pm_"],
        
        "hope drive clinic am":          "complex_am_",
        "hope drive clinic pm":          "complex_pm_",
        
        "briarcrest clinic am":          "adol_med_am_",
        "briarcrest clinic pm":          "adol_med_pm_",

        "lancaster am":          "lancaster_am_",
        "lancaster pm":          "lancaster_pm_",
    
        'hampden_nursery_print':    'custom_print_hampden_nursery_',
        'sjr_hospitalist_print':    'custom_print_sjr_hospitalist_',
        'aac_print':                'custom_print_aac_',
        'lancaster_cmg_print':      'custom_print_lancaster_cmg_',
    
        'mahoussi_aholoukpe_print': 'custom_print_mahoussi_aholoukpe_',
        
    }
    
    # which keys from base_map for each sheet
    sheet_map = {
        'ETOWN':           ('etown am continuity','etown pm continuity'),
        'NYES':            ('nyes rd am continuity','nyes rd pm continuity'),
        'LANCASTER':            ('lancaster am','lancaster pm'),
        'LANCASTER_CMG':        ('lancaster_cmg_print',),
        
        'COMPLEX':         ('hope drive clinic am','hope drive clinic pm'),
        'WARD A':             ('rounder 1 7a-7p','rounder 2 7a-7p','rounder 3 7a-7p'),
        'PSHCH_NURSERY':    ("nursery weekday 8a-6p","nursery weekday 8a-6p"),
        
        'HAMPDEN_NURSERY': ('hampden_nursery_print',),
        'SJR_HOSP':        ('sjr_hospitalist_print',),
        'AAC':             ('aac_print',),
        'AHOLOUKPE':        ('mahoussi_aholoukpe_print',),
        
        'ADOLMED':             ('briarcrest clinic am','briarcrest clinic pm'),
    }
    
    worksheet_names = ['HOPE_DRIVE','ETOWN','NYES','LANCASTER', 'LANCASTER_CMG', 'COMPLEX','WARD A','PSHCH_NURSERY','HAMPDEN_NURSERY','SJR_HOSP','AAC','AHOLOUKPE','ADOLMED']
    
    for ws in worksheet_names:
        # ─── HOPE_DRIVE ───────────────────────────────────────────
        if ws == 'HOPE_DRIVE':
                    # ─── HOPE_DRIVE: exact same 4‑week AM/PM acute+cont logic ───
            for week_idx in range(1, num_weeks + 1):
                week_base  = (week_idx - 1) * 24
                day_offset = (week_idx - 1) * 7
    
                for day_idx, col in enumerate(excel_column_letters, start=1):
                    is_weekday = day_idx <= 5
                    day_num    = day_idx + day_offset
    
                    # AM acute + continuity
                    if is_weekday:
                        # acute (_1–2)
                        for prov in range(1, 3):
                            row = week_base + hd_row_defs['AM']['acute_start'] + (prov - 1)
                            data_mappings.append({
                                'csv_column':  f'hd_am_acute_d{day_num}_{prov}',
                                'excel_sheet': 'HOPE_DRIVE',
                                'excel_cell':  f'{col}{row}',
                            })
                        # continuity (_1–8)
                        for prov in range(1, 9):
                            row = week_base + hd_row_defs['AM']['cont_start'] + (prov - 1)
                            data_mappings.append({
                                'csv_column':  f'hd_am_d{day_num}_{prov}',
                                'excel_sheet': 'HOPE_DRIVE',
                                'excel_cell':  f'{col}{row}',
                            })
                    else:
                        # weekend acute 1 & 2
                        for acute_type in (1, 2):
                            row = week_base + hd_row_defs['AM']['acute_start'] + (acute_type - 1)
                            data_mappings.append({
                                'csv_column':  f'hd_wknd_acute_{acute_type}_d{day_num}_1',
                                'excel_sheet': 'HOPE_DRIVE',
                                'excel_cell':  f'{col}{row}',
                            })
                        # weekend continuity
                        for prov in range(1, 9):
                            row = week_base + hd_row_defs['AM']['cont_start'] + (prov - 1)
                            data_mappings.append({
                                'csv_column':  f'hd_wknd_am_d{day_num}_{prov}',
                                'excel_sheet': 'HOPE_DRIVE',
                                'excel_cell':  f'{col}{row}',
                            })
    
                    # PM acute + continuity
                    if is_weekday:
                        for prov in range(1, 3):
                            row = week_base + hd_row_defs['PM']['acute_start'] + (prov - 1)
                            data_mappings.append({
                                'csv_column':  f'hd_pm_acute_d{day_num}_{prov}',
                                'excel_sheet': 'HOPE_DRIVE',
                                'excel_cell':  f'{col}{row}',
                            })
                        for prov in range(1, 9):
                            row = week_base + hd_row_defs['PM']['cont_start'] + (prov - 1)
                            data_mappings.append({
                                'csv_column':  f'hd_pm_d{day_num}_{prov}',
                                'excel_sheet': 'HOPE_DRIVE',
                                'excel_cell':  f'{col}{row}',
                            })
                    else:
                        for acute_type in (1, 2):
                            row = week_base + hd_row_defs['PM']['acute_start'] + (acute_type - 1)
                            data_mappings.append({
                                'csv_column':  f'hd_wknd_pm_acute_{acute_type}_d{day_num}_1',
                                'excel_sheet': 'HOPE_DRIVE',
                                'excel_cell':  f'{col}{row}',
                            })
                        for prov in range(1, 9):
                            row = week_base + hd_row_defs['PM']['cont_start'] + (prov - 1)
                            data_mappings.append({
                                'csv_column':  f'hd_wknd_pm_d{day_num}_{prov}',
                                'excel_sheet': 'HOPE_DRIVE',
                                'excel_cell':  f'{col}{row}',
                            })
            # done with HOPE_DRIVE
            continue
    
    
        # ─── W_A (rounders) ───────────────────────────────────────
        if ws == 'W_A':
            mapping_keys = sheet_map[ws]  # ('rounder 1…','rounder 2…','rounder 3…')
            for week_idx in range(1, num_weeks+1):
                week_base  = (week_idx - 1) * 24
                day_offset = (week_idx - 1) * 7
    
                for day_idx, col in enumerate(excel_column_letters, start=1):
                    day_num = day_idx + day_offset
    
                    # AM block → rows 6–…
                    row = week_base + cont_row_defs['AM']
                    for team_idx, key in enumerate(mapping_keys):
                        am_pref = base_map[key][0]  # e.g. "ward_a_am_"
                        provs   = assignments_by_date[date][key]
                        req     = min_required.get(key, len(provs))
                        # pad to exactly 2 providers
                        while len(provs) < req:
                            provs.append(provs[0])
                        offset = team_idx * req
                        for i, name in enumerate(provs, start=1):
                            slot = offset + i     # team1→1,2; team2→3,4; team3→5,6
                            data_mappings.append({
                                'csv_column': f"{am_pref}d{day_num}_{slot}",
                                'excel_sheet': ws,
                                'excel_cell': f"{col}{row}",
                            })
                            row += 1
    
                    # PM block → rows 16–…
                    row = week_base + cont_row_defs['PM']
                    for team_idx, key in enumerate(mapping_keys):
                        pm_pref = base_map[key][1]  # e.g. "ward_a_pm_"
                        provs   = assignments_by_date[date][key]
                        req     = min_required.get(key, len(provs))
                        while len(provs) < req:
                            provs.append(provs[0])
                        offset = team_idx * req
                        for i, name in enumerate(provs, start=1):
                            slot = offset + i
                            data_mappings.append({
                                'csv_column': f"{pm_pref}d{day_num}_{slot}",
                                'excel_sheet': ws,
                                'excel_cell': f"{col}{row}",
                            })
                            row += 1
    
            continue  # skip the generic logic below
    
        # ─── ALL OTHER SHEETS ──────────────────────────────────────
        mapping_keys = sheet_map.get(ws, ())
        if not mapping_keys:
            continue
    
        for key in mapping_keys:
            val = base_map[key]
        
            # Decide which side(s) this key applies to
            am_prefix = pm_prefix = None
            if isinstance(val, list):
                am_prefix, pm_prefix = val
            else:
                k = key.lower()
                if " am " in k:
                    am_prefix = val
                elif " pm " in k:
                    pm_prefix = val
                else:
                    # keys that don't encode AM/PM (rare) write to both
                    am_prefix = pm_prefix = val
    
            for week_idx in range(1, num_weeks + 1):
                week_base  = (week_idx - 1) * 24
                day_offset = (week_idx - 1) * 7
    
                for day_idx, col in enumerate(excel_column_letters, start=1):
                    day_num = day_idx + day_offset
    
                    # AM continuity (_1–10)
                    for prov in range(1, 11):
                        row = week_base + cont_row_defs['AM'] + (prov - 1)
                        data_mappings.append({
                            'csv_column': f"{am_prefix}d{day_num}_{prov}",
                            'excel_sheet': ws,
                            'excel_cell': f"{col}{row}",
                        })
    
                    # PM continuity (_1-10)
                    for prov in range(1, 11):
                        row = week_base + cont_row_defs['PM'] + (prov - 1)
                        data_mappings.append({
                            'csv_column': f"{pm_prefix}d{day_num}_{prov}",
                            'excel_sheet': ws,
                            'excel_cell': f"{col}{row}",
                        })
    
                
    # --- Main execution flow for generating and then updating the workbook ---
    st.subheader("Generate & Update OPD.xlsx + Summary")
    
    if st.button("Generate OPD File For Sarah to Load Students"):
        # 1) Generate the initial OPD workbook
        excel_template_bytes = generate_opd_workbook(out_df)
        if not excel_template_bytes:
            st.error("Failed to generate OPD template.")
            st.stop()
    
        # 2) Update it with your CSV data
        updated_excel_bytes = update_excel_from_csv(excel_template_bytes, csv_full, data_mappings)
        if not updated_excel_bytes:
            st.error("Failed to update OPD.xlsx with data.")
            st.stop()

        cleaned_bytes, hidden_map, hidden_total = hide_blank_rows_all_sheets(updated_excel_bytes)
        st.success("✅ OPD.xlsx updated successfully!")
    
        # 3) Build your summary DataFrame (reuse your df_summary logic)
        summary = []
        for student in legal_names:
            entry = {"Student": student}
            for w in range(4):
                days = [d + w*7 for d in range(1,6)]
                assigns = []
                # Ward A
                ward_found = False
                for shift in ("am","pm"):
                    for slot in range(1,7):
                        for d in days:
                            key = f"ward_a_{shift}_d{d}_{slot}"
                            if student in redcap_row.get(key,""):
                                assigns.append("Ward A")
                                ward_found = True
                                break
                        if ward_found: break
                    if ward_found: break
                # Hampden
                if not ward_found:
                    for d in days:
                        key = f"custom_print_hampden_nursery_d{d}_4"
                        if student in redcap_row.get(key,""):
                            assigns.append("Hampden")
                            break
                # SJR
                sjr_found = False
                for slot in (3,4):
                    for d in days:
                        key = f"custom_print_sjr_hospitalist_d{d}_{slot}"
                        if student in redcap_row.get(key,""):
                            assigns.append("SJR")
                            sjr_found = True
                            break
                    if sjr_found: break
                # PSHCH
                pshch_found = False
                for slot in (1,2):
                    for d in days:
                        for pref in ("nursery_am_","nursery_pm_"):
                            key = f"{pref}d{d}_{slot}"
                            if student in redcap_row.get(key,""):
                                assigns.append("PSHCH")
                                pshch_found = True
                                break
                        if pshch_found: break
                    if pshch_found: break
    
                entry[f"Week {w+1}"] = ", ".join(assigns) or ""
            summary.append(entry)
        df_summary = pd.DataFrame(summary)
    
        # 4) Build a Word doc with the summary table
        doc = Document()
        # make landscape
        section = doc.sections[0]
        section.orientation = WD_ORIENT.LANDSCAPE
        section.page_width, section.page_height = section.page_height, section.page_width
        
        doc.add_heading(rt('assignment_summary.title'), level=1)
        add_report_message(doc, 'assignment_summary.opening')
        
        cols  = df_summary.columns.tolist()
        table = doc.add_table(rows=1, cols=len(cols), style="Table Grid")
        hdr_cells = table.rows[0].cells
        for i, c in enumerate(cols):
            hdr_cells[i].text = c
        
        for _, row in df_summary.iterrows():
            row_cells = table.add_row().cells
            for i, c in enumerate(cols):
                row_cells[i].text = str(row[c])
        
        # **Save** into bytes
        word_io = io.BytesIO()
        add_report_message(doc, 'assignment_summary.closing')
        apply_report_appearance(doc, 'assignment_summary')
        doc.save(word_io)
        word_io.seek(0)
        word_bytes = word_io.read()
        
        # 5) Package into a ZIP
        zip_io = io.BytesIO()
        with zipfile.ZipFile(zip_io, "w") as z:
            z.writestr("Updated_OPD.xlsx", cleaned_bytes)
            z.writestr("Assignment_Summary.docx", word_bytes)
        zip_io.seek(0)
        
        # 6) Single download
        st.download_button(label="⬇️ Download OPD.xlsx + Summary (zip)",data=zip_io.read(),file_name="Batch_Output.zip",mime="application/zip")
