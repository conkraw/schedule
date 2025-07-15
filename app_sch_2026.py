import io
import streamlit as st
import pandas as pd
import re
import xlsxwriter

st.set_page_config(page_title="Batch Preceptor → REDCap Import", layout="wide")
st.title("Batch Preceptor → REDCap Import Generator")

# ─── Inputs ────────────────────────────────────────────────────────────────────
schedule_files = st.file_uploader("1) Upload one or more AGP calendar Excel(s)",type=["xlsx","xls"],accept_multiple_files=True)

student_file = st.file_uploader(
    "2) Upload student list CSV (must have a 'legal_name' column)",
    type=["csv"]
)

record_id = st.text_input("3) Enter the REDCap record_id for this batch", "")

# ─── Guard ─────────────────────────────────────────────────────────────────────
if not schedule_files or not student_file or not record_id:
    st.info("Please upload schedule Excel(s), student CSV, and enter a record_id.")
    st.stop()

# ─── Prep: Date regex & maps ───────────────────────────────────────────────────

date_pat = re.compile(r'^[A-Za-z]+ \d{1,2}, \d{4}$')
base_map = {
    "hope drive am continuity":    "hd_am_",
    "hope drive pm continuity":    "hd_pm_",
    
    "hope drive am acute precept": "hd_am_acute_",
    "hope drive pm acute precept": "hd_pm_acute_",
    
    "etown am continuity":         "etown_am_",
    "etown pm continuity":         "etown_pm_",
    
    "nyes rd am continuity":       "nyes_am_",
    "nyes rd pm continuity":       "nyes_pm_",
    
    "nursery weekday 8a-6p":       ["nursery_am_", "nursery_pm_"],
    
    "rounder 1 7a-7p":             ["ward_a_am_team_1_","ward_a_pm_team_1_"],
    "rounder 2 7a-7p":             ["ward_a_am_team_2_","ward_a_pm_team_2_"],
    "rounder 3 7a-7p":             ["ward_a_am_team_3_","ward_a_pm_team_3_"],

    "hope drive clinic am":        "complex_am_1_",
    "hope drive clinic pm":        "complex_pm_1_",
    
    "briarcrest clinic am":       "adol_med_am_1_",
    "briarcrest clinic pm":       "adol_med_pm_1_",
    
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
            cell = raw.lower()
            
            if cell in day_names:
                break
                
            prov = str(df.iat[r,col0+1]).strip()
            if cell in grp and prov:
                grp[cell].append(prov)

# ─── 2. Read student list and prepare s1, s2, … ───────────────────────────────
students_df = pd.read_csv(student_file, dtype=str)
legal_names = students_df["legal_name"].dropna().tolist()

# ─── 3. Build the single REDCap row ───────────────────────────────────────────
redcap_row = {"record_id": record_id}

# sort dates chronologically
sorted_dates = sorted(assignments_by_date.keys())

# loop days for schedule fields
for idx, date in enumerate(sorted_dates, start=1):
    # date
    redcap_row[f"hd_day_date{idx}"] = date
    suffix = f"d{idx}_"
    # designation→ day-specific prefixes
    des_map = {
        des: ([p+suffix for p in prefs] if isinstance(prefs,list) else [prefs+suffix])
        for des,prefs in base_map.items()
    }
    # providers for this date
    for des, provs in assignments_by_date[date].items():
        req = min_required.get(des, len(provs))
        while len(provs) < req and provs:
            provs.append(provs[0])
        for i,name in enumerate(provs, start=1):
            for prefix in des_map[des]:
                redcap_row[f"{prefix}{i}"] = name

# append student slots s1,s2,...
for i,name in enumerate(legal_names, start=1):
    redcap_row[f"s{i}"] = name

# ─── 4. Display & slice out dates/am/acute and students ─────────────────────
out_df = pd.DataFrame([redcap_row])

# format date columns
for c in out_df.columns:
    if c.startswith("hd_day_date"):
        out_df[c] = pd.to_datetime(out_df[c]).dt.strftime("%m-%d-%Y")

st.subheader("✅ Full REDCap Import Preview")
st.dataframe(out_df)

# subset columns
date_cols     = [c for c in out_df.columns if c.startswith("hd_day_date")]
am_cont_cols  = [f"hd_am_d1_{i}" for i in range(1, 19)] + [f"hd_am_d2_{i}" for i in range(1, 19)] + [f"hd_am_d3_{i}" for i in range(1, 19)] + [f"hd_am_d4_{i}" for i in range(1, 19)] +[f"hd_am_d5_{i}" for i in range(1, 19)]
am_cont_cols  = am_cont_cols + [f"hd_pm_d1_{i}" for i in range(1, 19)] + [f"hd_pm_d2_{i}" for i in range(1, 19)] + [f"hd_pm_d3_{i}" for i in range(1, 19)] + [f"hd_pm_d4_{i}" for i in range(1, 19)] +[f"hd_pm_d5_{i}" for i in range(1, 19)]
am_acute_cols = [f"hd_am_acute_d1_{i}" for i in (1,2)] + [f"hd_am_acute_d2_{i}" for i in (1,2)] + [f"hd_am_acute_d3_{i}" for i in (1,2)] + [f"hd_am_acute_d4_{i}" for i in (1,2)] + [f"hd_am_acute_d5_{i}" for i in (1,2)] 
am_acute_cols = am_acute_cols + [f"hd_pm_acute_d1_{i}" for i in (1,2)] + [f"hd_pm_acute_d2_{i}" for i in (1,2)] + [f"hd_pm_acute_d3_{i}" for i in (1,2)] + [f"hd_pm_acute_d4_{i}" for i in (1,2)] + [f"hd_pm_acute_d5_{i}" for i in (1,2)] 
student_cols  = [f"s{i}" for i in range(1, len(legal_names)+1)]

subset = ["record_id"] + date_cols + am_cont_cols + am_acute_cols + student_cols
subset = [c for c in subset if c in out_df.columns]

st.subheader("📅 Dates, AM Continuity/Acute & Students")
st.dataframe(out_df[subset])

# downloads
csv_full = out_df.to_csv(index=False).encode("utf-8")
st.download_button("⬇️ Download Full CSV", csv_full, "batch_import_full.csv", "text/csv")

csv_sub  = out_df[subset].to_csv(index=False).encode("utf-8")
st.download_button("⬇️ Download Dates+AM+Students CSV", csv_sub, "dates_am_students.csv", "text/csv")

def generate_opd_workbook(full_df: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(output, {'in_memory': True})
    # Formats
    format1 = workbook.add_format({'font_size':18,'bold':1,'align':'center','valign':'vcenter','font_color':'black','bg_color':'#FEFFCC','border':1})
    format4 = workbook.add_format({'font_size':12,'bold':1,'align':'center','valign':'vcenter','font_color':'black','bg_color':'#8ccf6f','border':1})
    format4a= workbook.add_format({'font_size':12,'bold':1,'align':'center','valign':'vcenter','font_color':'black','bg_color':'#9fc5e8','border':1})
    format5 = workbook.add_format({'font_size':12,'bold':1,'align':'center','valign':'vcenter','font_color':'black','bg_color':'#FEFFCC','border':1})
    format5a= workbook.add_format({'font_size':12,'bold':1,'align':'center','valign':'vcenter','font_color':'black','bg_color':'#d0e9ff','border':1})
    format11= workbook.add_format({'font_size':18,'bold':1,'align':'center','valign':'vcenter','font_color':'black','bg_color':'#FEFFCC','border':1})
    formate= workbook.add_format({'font_size':12,'bold':0,'align':'center','valign':'vcenter','font_color':'white','border':0})
    format3 = workbook.add_format({'font_size':12,'bold':1,'align':'center','valign':'vcenter','font_color':'black','bg_color':'#FFC7CE','border':1})
    format2 = workbook.add_format({'bg_color':'black'})
    format_date = workbook.add_format({'num_format':'m/d/yyyy','font_size':12,'bold':1,'align':'center','valign':'vcenter','font_color':'black','bg_color':'#FFC7CE','border':1})
    format_label= workbook.add_format({'font_size':12,'bold':1,'align':'center','valign':'vcenter','font_color':'black','bg_color':'#FFC7CE','border':1})
    merge_format = workbook.add_format({'bold':1,'align':'center','valign':'vcenter','text_wrap':True,'font_color':'red','bg_color':'#FEFFCC','border':1})

    # Worksheets
    worksheet_names = ['HOPE_DRIVE','ETOWN','NYES','COMPLEX','W_A','PSHCH_NURSERY','HAMPDEN_NURSERY','SJR_HOSP','AAC']
    sheets = {name: workbook.add_worksheet(name) for name in worksheet_names}

    # Site headers
    site_map = dict(zip(sheets.values(), ['Hope Drive','Elizabethtown','Nyes Road','Complex Care','WARD A','PSHCH NURSERY','HAMPDEN NURSERY','SJR HOSPITALIST','AAC']))
    for ws, site in site_map.items():
        ws.write(0,0,'Site:',format1)
        ws.write(0,1,site,format1)

    # HOPE_DRIVE specific
    hd = sheets['HOPE_DRIVE']
    for cell_range in ['A8:H15','A32:H39','A56:H63','A80:H87']:
        hd.conditional_format(cell_range,{'type':'cell','criteria':'>=','value':0,'format':format1})
    for cell_range in ['A18:H25','A42:H49','A66:H73','A90:H97']:
        hd.conditional_format(cell_range,{'type':'cell','criteria':'>=','value':0,'format':format5a})
    for cell_range in ['A6:H6','A7:H7','A30:H30','A31:H31','A54:H54','A55:H55','A78:H78','A79:H79']:
        hd.conditional_format(cell_range,{'type':'cell','criteria':'>=','value':0,'format':format4})
    for cell_range in ['A16:H16','A17:H17','A40:H40','A41:H41','A64:H64','A65:H65','A88:H88','A89:H89']:
        hd.conditional_format(cell_range,{'type':'cell','criteria':'>=','value':0,'format':format4a})
    acute_ranges = [(6,7),(16,17),(30,31),(40,41),(54,55),(64,65),(78,79),(88,89)]
    for r1,r2 in acute_ranges:
        for r in range(r1,r2+1): hd.write(f'A{r}', 'AM - ACUTES' if r1%2==0 else 'PM - ACUTES', format4 if r1%2==0 else format4a)
    cont_ranges = [(8,15),(18,25),(32,39),(42,49),(56,63),(66,73),(80,87),(90,97)]
    for r1,r2 in cont_ranges:
        for r in range(r1,r2+1): hd.write(f'A{r}', 'AM - Continuity' if r1%2==0 else 'PM - Continuity', format5a)
    labels = [f'H{i}' for i in range(20)]
    for start in [6,30,54,78]:
        for i,lab in enumerate(labels): hd.write(f'I{start+i}', lab, formate)

    # Generic sheets
    others = [s for n,s in sheets.items() if n!='HOPE_DRIVE']
    for ws in others:
        for cell_range in ['A6:H15','A30:H39','A54:H63','A78:H87']:
            ws.conditional_format(cell_range,{'type':'cell','criteria':'>=','value':0,'format':format1})
        for cell_range in ['A16:H25','A40:H49','A64:H73','A88:H97']:
            ws.conditional_format(cell_range,{'type':'cell','criteria':'>=','value':0,'format':format5a})
        for cell_range,fmt in zip(['B6:H6','B16:H16','B30:H30','B40:H40','B54:H54','B64:H64','B78:H78','B88:H88'],[format4,format4a]*4):
            ws.conditional_format(cell_range,{'type':'cell','criteria':'>=','value':0,'format':fmt})
        am_pm = ['AM']*10+['PM']*10
        for block,start in enumerate([6,30,54,78]):
            for i,lab in enumerate(am_pm): ws.write(f'A{start+i}', lab, format5a)
        for start in [6,30,54,78]:
            for i,lab in enumerate(labels): ws.write(f'I{start+i}', lab, formate)

    # Universal formatting & dates
    date_cols = [f"hd_day_date{i}" for i in range(1,29)]
    dates = pd.to_datetime(full_df[date_cols].iloc[0]).tolist()
    weeks = [dates[i*7:(i+1)*7] for i in range(4)]
    for ws in workbook.worksheets():
        ws.set_zoom(80)
        ws.set_column('A:A',10)
        ws.set_column('B:H',65)
        ws.set_row(0,37.25)
        for idx, start in enumerate([2,26,50,74]):
            days = ['Monday','Tuesday','Wednesday','Thursday','Friday','Saturday','Sunday']
            for c,d in enumerate(days): ws.write(start,1+c,d,format3)
            for c,val in enumerate(weeks[idx]): ws.write(start+1,1+c,val,format_date)
            ws.write_formula(f'A{start}', '""', format_label)
            ws.write(f'A{start-1}', "", format_label)
            ws.write(f'A{start+1}', "", format_label)
            ws.conditional_format(f'A{start+3}:H{start+3}',{'type':'cell','criteria':'>=','value':0,'format':format_label})
        # black bars
        step = 24
        for row in range(2,98,step): ws.merge_range(f'A{row}:H{row}', ' ', format2)
        text1 = 'Students are to alert their preceptors when they have a Clinical Reasoning Teaching Session (CRTS).  Please allow the students to leave approximately 15 minutes prior to the start of their session so they can be prepared to actively participate.  ~ Thank you!'
        ws.merge_range('C1:F1', text1, merge_format)
        ws.write('G1','',merge_format); ws.write('H1','',merge_format)

    workbook.close()
    output.seek(0)
    return output.read()

excel_bytes = generate_opd_workbook(out_df)
st.download_button(label="⬇️ Download OPD.xlsx",data=excel_bytes,file_name="OPD.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


