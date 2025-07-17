import io
import streamlit as st
import pandas as pd
import re
import xlsxwriter
import random
from openpyxl import load_workbook # Ensure load_workbook is imported
import io, zipfile
from docx import Document
from docx.enum.section import WD_ORIENT
from datetime import timedelta
from xlsxwriter import Workbook as Workbook

st.set_page_config(page_title="Batch Preceptor → REDCap Import", layout="wide")
st.title("Batch Preceptor → REDCap Import Generator")

# ─── Sidebar mode selector ─────────────────────────────────────────────────────
mode = st.sidebar.radio(
    "What do you want to do?",
    ("Format OPD + Summary", "Create Student Schedule")
)

if mode == "Format OPD + Summary":
    # ─── Inputs ────────────────────────────────────────────────────────────────────
    # Required keywords to look for in the content
    required_keywords = ["academic general pediatrics", "hospitalists", "complex care", "adol med"]
    found_keywords = set()
    
    schedule_files = st.file_uploader("1) Upload one or more QGenda calendar Excel(s)",type=["xlsx", "xls"],accept_multiple_files=True)
    
    if schedule_files:
        for file in schedule_files:
            try:
                # Read the first sheet
                df = pd.read_excel(file, sheet_name=0, header=None)
    
                # Flatten all string values to a list of lowercase strings
                cell_values = df.astype(str).apply(lambda x: x.str.lower()).values.flatten().tolist()
    
                # Check if any keyword is found in cell values
                for keyword in required_keywords:
                    if any(keyword in val for val in cell_values):
                        found_keywords.add(keyword)
    
            except Exception as e:
                st.error(f"Error reading {file.name}: {e}")
    
        # Identify missing calendars
        missing_keywords = [k for k in required_keywords if k not in found_keywords]
    
        if missing_keywords:
            st.warning(f"Missing required calendar(s): {', '.join(missing_keywords)}. Please upload all four.")
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
    
    file_configs = {
        "HAMPDEN_NURSERY.xlsx": {
            "title":       "HAMPDEN_NURSERY",
            "custom_text": "CUSTOM_PRINT",
            "names": [
                "Folaranmi, Oluwamayoda",
                "Alur, Pradeep",
                "Nanda, Sharmilarani",
                "HAMPDEN_NURSERY"
            ]
        },
        "SJR_HOSP.xlsx": {
            "title":       "SJR_HOSPITALIST",
            "custom_text": "CUSTOM_PRINT",
            "names": [
                "Spangola, Haley",
                "Gubitosi, Terry",
                "SJR_1",
                "SJR_2"
            ]
        },
        "AAC.xlsx": {
            "title":       "AAC",
            "custom_text": "CUSTOM_PRINT",
            "names": [
                "Vaishnavi Harding",
                "Abimbola Ajayi",
                "Shilu Joshi",
                "Desiree Webb",
                "Amy Zisa",
                "Abdullah Sakarcan",
                "Anna Karasik",
                "AAC_1",
                "AAC_2",
                "AAC_3",
            ]
        },
        # New standalone sheet for Mahoussi
        "MAHOUSSI_AHOLOUKPE.xlsx": {
            "title":       "MAHOUSSI_AHOLOUKPE",
            "custom_text": "CUSTOM_PRINT",
            "names": [
                "Mahoussi Aholoukpe"
            ]
        },
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
    
        # 3a) schedule providers (your existing hd_am_, ward_a_am_, etc.)
        for des, provs in assignments_by_date[date].items():
            req = min_required.get(des, len(provs))
            while len(provs) < req and provs:
                provs.append(provs[0])
    
            if des.startswith("rounder"):
                team_idx = int(des.split()[1]) - 1
                for i, name in enumerate(provs, start=1):
                    slot = team_idx * req + i
                    for prefix in des_map[des]:
                        redcap_row[f"{prefix}{slot}"] = name
            else:
                for i, name in enumerate(provs, start=1):
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
    
    # ─── HAMPDEN_NURSERY: max 1 student for week1 and 1 for week3, into slot _4 ─────
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
    
    # ─── File to Check Column Assignments ─────────────────────────────────────────────────────────────────
    #st.subheader("✅ Full REDCap Import Preview")
    #st.dataframe(out_df)
    
    #st.download_button("⬇️ Download Full CSV", csv_full, "batch_import_full.csv", "text/csv")
    
    def generate_opd_workbook(full_df: pd.DataFrame) -> bytes:
        import io
        import xlsxwriter
    
        output = io.BytesIO()
        workbook = xlsxwriter.Workbook(output, {'in_memory': True})
    
        # ─── Formats ─────────────────────────────────────────────────────────────────
        format1     = workbook.add_format({
            'font_size':18,'bold':1,'align':'center','valign':'vcenter',
            'font_color':'black','bg_color':'#FEFFCC','border':1
        })
        format4     = workbook.add_format({
            'font_size':12,'bold':1,'align':'center','valign':'vcenter',
            'font_color':'black','bg_color':'#8ccf6f','border':1
        })
        format4a    = workbook.add_format({
            'font_size':12,'bold':1,'align':'center','valign':'vcenter',
            'font_color':'black','bg_color':'#9fc5e8','border':1
        })
        format5     = workbook.add_format({
            'font_size':12,'bold':1,'align':'center','valign':'vcenter',
            'font_color':'black','bg_color':'#FEFFCC','border':1
        })
        format5a    = workbook.add_format({
            'font_size':12,'bold':1,'align':'center','valign':'vcenter',
            'font_color':'black','bg_color':'#d0e9ff','border':1
        })
        format11    = workbook.add_format({
            'font_size':18,'bold':1,'align':'center','valign':'vcenter',
            'font_color':'black','bg_color':'#FEFFCC','border':1
        })
        formate     = workbook.add_format({
            'font_size':12,'bold':0,'align':'center','valign':'vcenter',
            'font_color':'white','border':0
        })
        format3     = workbook.add_format({
            'font_size':12,'bold':1,'align':'center','valign':'vcenter',
            'font_color':'black','bg_color':'#FFC7CE','border':1
        })
        format2     = workbook.add_format({'bg_color':'black'})
        format_date = workbook.add_format({
            'num_format':'m/d/yyyy','font_size':12,'bold':1,
            'align':'center','valign':'vcenter',
            'font_color':'black','bg_color':'#FFC7CE','border':1
        })
        format_label= workbook.add_format({
            'font_size':12,'bold':1,'align':'center','valign':'vcenter',
            'font_color':'black','bg_color':'#FFC7CE','border':1
        })
        merge_format= workbook.add_format({
            'bold':1,'align':'center','valign':'vcenter','text_wrap':True,
            'font_color':'red','bg_color':'#FEFFCC','border':1
        })
    
        # ─── Worksheets ─────────────────────────────────────────────────────────────
        worksheet_names = [
            'HOPE_DRIVE','ETOWN','NYES','COMPLEX',
            'W_A','PSHCH_NURSERY','HAMPDEN_NURSERY','SJR_HOSP','AAC','AHOLOUKPE','ADOLMED'
        ]
        sheets = {name: workbook.add_worksheet(name) for name in worksheet_names}
    
        # ─── Site headers ────────────────────────────────────────────────────────────
        site_list = [
            'Hope Drive','Elizabethtown','Nyes Road','Complex Care',
            'WARD A','PSHCH NURSERY','HAMPDEN NURSERY','SJR HOSPITALIST','AAC','AHOLOUKPE','ADOLMED'
        ]
        for ws, site in zip(sheets.values(), site_list):
            ws.write(0, 0, 'Site:', format1)
            ws.write(0, 1, site,   format1)
    
        # ─── HOPE_DRIVE specific ────────────────────────────────────────────────────
        hd = sheets['HOPE_DRIVE']
        for cr in ['A8:H15','A32:H39','A56:H63','A80:H87']:
            hd.conditional_format(cr, {'type':'cell','criteria':'>=','value':0,'format':format1})
        for cr in ['A18:H25','A42:H49','A66:H73','A90:H97']:
            hd.conditional_format(cr, {'type':'cell','criteria':'>=','value':0,'format':format5a})
        for cr in ['A6:H6','A7:H7','A30:H30','A31:H31','A54:H54','A55:H55','A78:H78','A79:H79']:
            hd.conditional_format(cr, {'type':'cell','criteria':'>=','value':0,'format':format4})
        for cr in ['A16:H16','A17:H17','A40:H40','A41:H41','A64:H64','A65:H65','A88:H88','A89:H89']:
            hd.conditional_format(cr, {'type':'cell','criteria':'>=','value':0,'format':format4a})
        acute_ranges = [(6,7),(16,17),(30,31),(40,41),(54,55),(64,65),(78,79),(88,89)]
        for r1, r2 in acute_ranges:
            fmt   = format4 if r1 % 2 == 0 else format4a
            label = 'AM - ACUTES' if r1 % 2 == 0 else 'PM - ACUTES'
            for r in range(r1, r2+1):
                hd.write(f'A{r}', label, fmt)
        cont_ranges = [(8,15),(18,25),(32,39),(42,49),(56,63),(66,73),(80,87),(90,97)]
        for r1, r2 in cont_ranges:
            for r in range(r1, r2+1):
                hd.write(
                    f'A{r}',
                    'AM - Continuity' if r1 % 2 == 0 else 'PM - Continuity',
                    format5a
                )
        labels = [f'H{i}' for i in range(20)]
        for start in [6, 30, 54, 78]:
            for i, lab in enumerate(labels):
                hd.write(f'I{start+i}', lab, formate)
    
        # ─── GENERIC SHEETS ─────────────────────────────────────────────────────────
        others       = [ws for name, ws in sheets.items() if name != 'HOPE_DRIVE']
        AM_COUNT     = 10
        PM_COUNT     = 10
        BLOCK_STARTS = [6, 30, 54, 78]
    
        for ws in others:
            # 1) conditional formats
            for cr in ['A6:H15','A30:H39','A54:H63','A78:H87']:
                ws.conditional_format(cr, {
                    'type':'cell','criteria':'>=','value':0,'format':format1
                })
            for cr in ['A16:H25','A40:H49','A64:H73','A88:H97']:
                ws.conditional_format(cr, {
                    'type':'cell','criteria':'>=','value':0,'format':format5a
                })
            for cr, fmt in [
                ('B6:H6',   format4), ('B16:H16', format4a),
                ('B30:H30', format4), ('B40:H40', format4a),
                ('B54:H54', format4), ('B64:H64', format4a),
                ('B78:H78', format4), ('B88:H88', format4a)
            ]:
                ws.conditional_format(cr, {
                    'type':'cell','criteria':'>=','value':0,'format':fmt
                })
    
            # 2) Write exactly 10 AM then 10 PM in column A
            for start in BLOCK_STARTS:
                zero_row = start - 1
                for i in range(AM_COUNT):
                    ws.write(zero_row + i, 0, 'AM', format5a)
                for i in range(PM_COUNT):
                    ws.write(zero_row + AM_COUNT + i, 0, 'PM', format5a)
                for i, lab in enumerate(labels):
                    ws.write(start + i, 8, lab, formate)
    
            # 3) Write H0…H19 in column I
            for start in BLOCK_STARTS:
                for i, lab in enumerate(labels):
                    ws.write(f'I{start + i}', lab, formate)
    
    
        # ─── Universal formatting & dates ────────────────────────────────────────────
        date_cols = [f"hd_day_date{i}" for i in range(1,29)]
        dates     = pd.to_datetime(full_df[date_cols].iloc[0]).tolist()
        weeks     = [dates[i*7:(i+1)*7] for i in range(4)]
        days      = ['Monday','Tuesday','Wednesday','Thursday','Friday','Saturday','Sunday']
    
        for ws in workbook.worksheets():
            ws.set_zoom(80)
            ws.set_column('A:A', 10)
            ws.set_column('B:H', 65)
            ws.set_row(0, 37.25)
    
            for idx, start in enumerate([2,26,50,74]):
                        # day names
                        for c, d in enumerate(days):
                            ws.write(start, 1+c, d, format3)
                        # dates
                        for c, val in enumerate(weeks[idx]):
                            ws.write(start+1, 1+c, val, format_date)
                            
                        # padding formula bars
                        ws.write_formula(f'A{start}',   '""', format_label)
                        
                        ws.conditional_format(
                            f'A{start+3}:H{start+3}',
                            {'type':'cell','criteria':'>=','value':0,'format':format_label}
                        )
    
    
            # black bars every 24 rows
            step = 24
            for row in range(2, 98, step):
                ws.merge_range(f'A{row}:H{row}', ' ', format2)
    
            # merge CRTS message on every sheet
            text1 = (
                'Students are to alert their preceptors when they have a Clinical '
                'Reasoning Teaching Session (CRTS).  Please allow the students to '
                'leave approximately 15 minutes prior to the start of their session '
                'so they can be prepared to actively participate.  ~ Thank you!'
            )
            ws.merge_range('C1:F1', text1, merge_format)
            ws.write('G1', '', merge_format)
            ws.write('H1', '', merge_format)
    
            #PAINT Empty White Spaces
            ws.write('A3', '', format_date)
            ws.write('A4', '', format_date)
            ws.write('A27', '', format_date)
            ws.write('A28', '', format_date)
            ws.write('A51', '', format_date)
            ws.write('A52', '', format_date)
            ws.write('A75', '', format_date)
            ws.write('A76', '', format_date)
    
    
        workbook.close()
        output.seek(0)
        return output.read()
    
    excel_bytes = generate_opd_workbook(out_df)
    #st.download_button(label="⬇️ Download OPD.xlsx",data=excel_bytes,file_name="OPD.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    
    
    import pandas as pd
    import io
    from openpyxl import load_workbook
    from openpyxl.styles import Alignment # <--- NEW IMPORT
    
    # --- MODIFIED update_excel_from_csv function to work with bytes ---
    def update_excel_from_csv(excel_template_bytes: bytes, csv_data_bytes: bytes, mappings: list) -> bytes | None:
        """
        Updates an Excel file (from bytes) with values from a CSV (from bytes)
        based on provided mappings, and returns the updated Excel as bytes.
    
        Args:
            excel_template_bytes (bytes): The bytes content of the Excel file to be updated.
            csv_data_bytes (bytes): The bytes content of the CSV file containing the data.
            mappings (list of dict): A list of dictionaries, where each dictionary
                                     defines a mapping:
                                     {'csv_column': 'name_of_csv_column',
                                      'excel_sheet': 'Sheet Name',
                                      'excel_cell': 'Cell Address (e.g., B8)'}
        Returns:
            bytes: The bytes content of the updated Excel workbook, or None if an error occurs.
        """
        try:
            # Load the CSV data from bytes using BytesIO
            df_csv = pd.read_csv(io.BytesIO(csv_data_bytes))
    
            if df_csv.empty:
                # Assuming st is available in the Streamlit environment
                # If running outside Streamlit, you might use print() or logging
                # st.error("Error: The CSV data is empty. Cannot update Excel.")
                return None
    
            # Load the Excel workbook from bytes using BytesIO
            wb = load_workbook(io.BytesIO(excel_template_bytes))
    
            # Iterate through the mappings and update the Excel file
            for mapping in mappings:
                csv_column = mapping['csv_column']
                excel_sheet_name = mapping['excel_sheet']
                excel_cell = mapping['excel_cell']
    
                if csv_column not in df_csv.columns:
                    # st.warning(f"Warning: CSV column '{csv_column}' not found in CSV data. Skipping this mapping.")
                    continue
    
                # Get the value from the first row of the specified CSV column
                value_to_transfer = df_csv.loc[0, csv_column]
    
                if excel_sheet_name not in wb.sheetnames:
                    # st.warning(f"Warning: Excel sheet '{excel_sheet_name}' not found in the Excel template. Skipping this mapping.")
                    continue
    
                ws = wb[excel_sheet_name]
    
                # --- APPLY THE REQUESTED FORMATTING HERE ---
                # 1. Convert to string and append " ~ "
                # This ensures that even if value_to_transfer is a number, it can be concatenated
                #formatted_value = str(value_to_transfer) + ' ~ '

                orig = str(value_to_transfer).strip()
                # if it already contains a student‐delimiter, don’t add another
                if ' ~ ' in orig:
                    formatted_value = orig
                else:
                    formatted_value = orig + ' ~ '
    
                # 2. Write the formatted value to the cell
                ws[excel_cell] = formatted_value
    
                # 3. Set the alignment for the cell using openpyxl's Alignment
                cell = ws[excel_cell] # Get the cell object
                cell.alignment = Alignment(horizontal='center', vertical='center') # Set horizontal and vertical centering
    
                # st.info(f"Successfully wrote '{formatted_value}' from CSV column '{csv_column}' to '{excel_sheet_name}'!{excel_cell}")
    
            # Save the modified Excel workbook to a BytesIO object
            output_excel_bytes_io = io.BytesIO()
            wb.save(output_excel_bytes_io)
            output_excel_bytes_io.seek(0) # Rewind the buffer to the beginning
    
            return output_excel_bytes_io.getvalue()
    
        except Exception as e:
            # st.error(f"An error occurred during Excel update: {e}")
            return None
            
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
    
        'hampden_nursery_print':    'custom_print_hampden_nursery_',
        'sjr_hospitalist_print':    'custom_print_sjr_hospitalist_',
        'aac_print':                'custom_print_aac_',
    
        'mahoussi_aholoukpe_print': 'custom_print_mahoussi_aholoukpe_',
        
    }
    
    # which keys from base_map for each sheet
    sheet_map = {
        'ETOWN':           ('etown am continuity','etown pm continuity'),
        'NYES':            ('nyes rd am continuity','nyes rd pm continuity'),
        'COMPLEX':         ('hope drive clinic am','hope drive clinic pm'),
        'W_A':             ('rounder 1 7a-7p','rounder 2 7a-7p','rounder 3 7a-7p'),
        'PSHCH_NURSERY':    ("nursery weekday 8a-6p","nursery weekday 8a-6p"),
        
        'HAMPDEN_NURSERY': ('hampden_nursery_print',),
        'SJR_HOSP':        ('sjr_hospitalist_print',),
        'AAC':             ('aac_print',),
        'AHOLOUKPE':        ('mahoussi_aholoukpe_print',),
        
        'ADOLMED':             ('briarcrest clinic am','briarcrest clinic pm'),
    }
    
    worksheet_names = ['HOPE_DRIVE','ETOWN','NYES','COMPLEX','W_A','PSHCH_NURSERY','HAMPDEN_NURSERY','SJR_HOSP','AAC','AHOLOUKPE','ADOLMED']
    
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
                                'csv_column':  f'hd_am_acute_d{day_idx}_{prov}',
                                'excel_sheet': 'HOPE_DRIVE',
                                'excel_cell':  f'{col}{row}',
                            })
                        # continuity (_1–8)
                        for prov in range(1, 9):
                            row = week_base + hd_row_defs['AM']['cont_start'] + (prov - 1)
                            data_mappings.append({
                                'csv_column':  f'hd_am_d{day_idx}_{prov}',
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
                                'csv_column':  f'hd_pm_acute_d{day_idx}_{prov}',
                                'excel_sheet': 'HOPE_DRIVE',
                                'excel_cell':  f'{col}{row}',
                            })
                        for prov in range(1, 9):
                            row = week_base + hd_row_defs['PM']['cont_start'] + (prov - 1)
                            data_mappings.append({
                                'csv_column':  f'hd_pm_d{day_idx}_{prov}',
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
            if isinstance(val, list):
                am_prefix, pm_prefix = val
            else:
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
    if st.button("Generate and Update Excel Files"):
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
        
        doc.add_heading("Assignment Summary by Week", level=1)
        
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
        doc.save(word_io)
        word_io.seek(0)
        word_bytes = word_io.read()
        
        # 5) Package into a ZIP
        zip_io = io.BytesIO()
        with zipfile.ZipFile(zip_io, "w") as z:
            z.writestr("Updated_OPD.xlsx", updated_excel_bytes)
            z.writestr("Assignment_Summary.docx", word_bytes)
        zip_io.seek(0)
        
        # 6) Single download
        st.download_button(label="⬇️ Download OPD.xlsx + Summary (zip)",data=zip_io.read(),file_name="Batch_Output.zip",mime="application/zip")

elif mode == "Create Student Schedule":
    st.subheader("Create Student Schedule")
    def save_to_session(filename, fileobj, namespace="uploaded_files"):
        st.session_state.setdefault(namespace, {})[filename] = fileobj

    def load_workbook_df(label, types, key):
        """
        Upload an .xlsx or .csv and return a DataFrame.
        Saves the raw upload into session_state.uploaded_files.
        """
        upload = st.file_uploader(label, type=types, key=key)
        if not upload:
            st.info(f"Please upload {label}.")
            return None
    
        name = upload.name
        try:
            if name.lower().endswith(".csv"):
                df = pd.read_csv(upload)
            else:
                df = pd.read_excel(upload)
            st.success(f"{name} loaded.")
            save_to_session(name, upload)
            return df
    
        except Exception as e:
            st.error(f"Error loading {name}: {e}")
            return None


    
    def create_ms_schedule_template(students, dates):
        buf = io.BytesIO()
        wb = xlsxwriter.Workbook(buf, {'in_memory': True})
    
        # — Formats —
        f1 = wb.add_format({'font_size':14,'bold':1,'align':'center','valign':'vcenter',
                            'font_color':'black','text_wrap':True,'bg_color':'#FEFFCC','border':1})
        f2 = wb.add_format({'font_size':10,'bold':1,'align':'center','valign':'vcenter',
                            'font_color':'yellow','bg_color':'black','border':1,'text_wrap':True})
        f3 = wb.add_format({'font_size':12,'bold':1,'align':'center','valign':'vcenter',
                            'font_color':'black','bg_color':'#FFC7CE','border':1})
        f4 = wb.add_format({'num_format':'mm/dd/yyyy','font_size':12,'bold':1,'align':'center',
                            'valign':'vcenter','font_color':'black','bg_color':'#F4F6F7','border':1})
        f5 = wb.add_format({'font_size':12,'bold':1,'align':'center','valign':'vcenter',
                            'font_color':'black','bg_color':'#F4F6F7','border':1})
        f6 = wb.add_format({'bg_color':'black','border':1})
        f7 = wb.add_format({'font_size':12,'bold':1,'align':'center','valign':'vcenter',
                            'font_color':'black','bg_color':'#90EE90','border':1})
        f8 = wb.add_format({'font_size':12,'bold':1,'align':'center','valign':'vcenter',
                            'font_color':'black','bg_color':'#89CFF0','border':1})
    
        days = ['Monday','Tuesday','Wednesday','Thursday','Friday','Saturday','Sunday']
        start_rows = [3, 11, 19, 27]
        weeks = ['Week 1','Week 2','Week 3','Week 4']
        due_texts = [
            'Quiz 1 Due',
            'Quiz 2, Pediatric Documentation #1, 1 Clinical Encounter Log Due',
            'Quiz 3 Due',
            'Quiz 4, Pediatric Documentation #2, Social Drivers of Health Assessment Form, Developmental Assessment of Pediatric Patient Form, All Clinical Encounter Logs are Due!'
        ]
    
        for name in students:
            title = name[:31].replace('/','-').replace('\\','-')
            ws = wb.add_worksheet(title)
            ws.set_zoom(70)
    
            # Header
            ws.merge_range('A1:A2','Student Name:', f1)
            ws.merge_range('B1:B2',      title,      f1)
            note = ("*Note* Asynchronous time is for coursework only. During this time period, "
                    "we expect students to do coursework, be available for any additional educational "
                    "activities, and any extra clinical time that may be available. If the student is not "
                    "available during this time period and has not made an absence request, the student "
                    "will be cited for unprofessionalism and will risk failing the course.")
            ws.merge_range('C1:H2', note, f2)
    
            # Column widths & row height
            ws.set_column('A:A', 20)
            ws.set_column('B:B', 30)
            ws.set_column('C:G', 40)
            ws.set_column('H:H',155)
            ws.set_row(0, 37.25)
    
            # Days headers and dates
            date_idx = 0
            for block, row in enumerate(start_rows):
                # 1) write the days on row `row`, cols B–H
                for col_offset, day in enumerate(days, start=1):
                    ws.write(row, col_offset, day, f3)
        
                # 2) write the dates on row `row+1`, cols C–I
                for col_offset in range(7):
                    if date_idx < len(dates):
                        ws.write(row+1, col_offset+2, dates[date_idx], f4)
                        date_idx += 1
    
            # Week labels
            for i, week in enumerate(weeks):
                row = 4 + (i * 8)
                ws.write(f'A{row}', week, f3)
    
            # AM / PM labels
            for i in range(4):
                ws.write(f'A{6 + i*8}', 'AM', f3)
                ws.write(f'A{7 + i*8}', 'PM', f3)
    
            # Fill AM/PM blocks with Asynchronous Time (cols C–J)
            for block in range(4):
                am_row = 6 + block*8
                pm_row = 7 + block*8
                for col in range(2, 10):
                    ws.write(am_row, col, "Asynchronous Time", f5)
                    ws.write(pm_row, col, "Asynchronous Time", f5)
    
            # Separators
            for sep in [10, 18, 26, 34]:
                ws.merge_range(f'A{sep}:H{sep}', '', f6)
    
            # Green filler rows
            for filler in [9, 17, 25, 33]:
                for col in range(8):
                    ws.write(filler, col, ' ', f7)
    
            # Assignment‑due rows
            for i, base in enumerate([8, 16, 24, 32]):
                ws.write(f'A{base}', 'ASSIGNMENT DUE:', f8)
                for col in range(1, 8):
                    if col == 5:
                        ws.write(base-1, col, 'Ask for Feedback!', f8)
                    elif col == 7:
                        ws.write(base-1, col, due_texts[i], f8)
                    else:
                        ws.write(base-1, col, ' ', f8)
    
        wb.close()
        buf.seek(0)
        return buf

            
    # 1️⃣ Load OPD.xlsx
    df_opd = load_workbook_df(
        label="Upload OPD.xlsx file",
        types=["xlsx"],
        key="opd_blank"
    )

    # 2️⃣ Load rotation schedule (if you still need it below)
    df_rot = load_workbook_df(
        label="Upload RedCap Rotation Schedule file (.xlsx or .csv)",
        types=["xlsx", "csv"],
        key="rot_blank"
    )

    # ───> INSERT MASTER SCHEDULE CREATION HERE <───
    if df_opd is not None:
        # 1) Load df_rot and df_opd earlier...
        df_rot['start_date'] = pd.to_datetime(df_rot['start_date'])
        min_start = df_rot['start_date'].min()
        monday    = min_start - timedelta(days=min_start.weekday())
        
        # 2) Build the 28 dates
        dates = pd.date_range(start=monday, periods=28, freq="D").tolist()
        
        # 3) Get your student list
        students = df_rot['legal_name'].dropna().unique().tolist()
        
        if st.button("Create Blank MS_Schedule.xlsx"):
            students = df_rot["legal_name"].dropna().unique()
            buf = create_ms_schedule_template(students, dates)
            st.download_button(
                "Download MS_Schedule.xlsx",
                data=buf.getvalue(),
                file_name="MS_Schedule.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    else:
        st.write("Upload both OPD.xlsx and the rotation schedule above to proceed.")
