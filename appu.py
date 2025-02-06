import streamlit as st
import csv
import datetime
import pandas as pd
import numpy as np
from datetime import timedelta
import xlsxwriter
import openpyxl
from io import BytesIO

# Initialize session state variables efficiently
session_defaults = {
    "page": "Home",
    "start_date": None,
    "uploaded_files": {},
    "uploaded_book4_file": {},
}
for key, value in session_defaults.items():
    st.session_state.setdefault(key, value)

# Function to change page and trigger rerun
def navigate_to(page):
    st.session_state.page = page
    st.rerun()

# Home Page
if st.session_state.page == "Home":
    st.title("Welcome to OPD Creator")
    st.write("Please choose what you'd like to do next.")

    if st.button("Go to Create OPD"):
        navigate_to("Create OPD")

    if st.button("Go to Create Student Schedule"):
        navigate_to("Create Student Schedule")

# Create OPD Page - Date Input
elif st.session_state.page == "Create OPD":
    st.title('Date Input for OPD')
    st.write('Enter start date in **m/d/yyyy format**, no leading zeros (e.g., 7/6/2021):')

    date_input = st.text_input('Start Date')

    if st.button('Submit Date') and date_input:
        try:
            test_date = datetime.datetime.strptime(date_input, "%m/%d/%Y")
            st.session_state.start_date = test_date
            st.success(f"Valid date entered: {test_date.strftime('%m/%d/%Y')}")
            navigate_to("Upload Files")
        except ValueError:
            st.error('Invalid date format. Please enter the date in **m/d/yyyy** format.')

# Upload Files Page
elif st.session_state.page == "Upload Files":
    st.title("File Upload Section")
    st.write("Upload the following required Excel files:")

    required_files = {
        "HOPE_DRIVE": "HOPE_DRIVE.xlsx",
        "ETOWN": "ETOWN.xlsx",
        "NYES": "NYES.xlsx",
        "WARD_A": "WARD_A.xlsx",
        "WARD_P": "WARD_P.xlsx",
        "COMPLEX": "COMPLEX.xlsx",
        "PICU": "PICU.xlsx",
    }

    uploaded_files = st.file_uploader("Choose your files", type="xlsx", accept_multiple_files=True)

    if uploaded_files:
        uploaded_files_dict = {fname: file for file in uploaded_files for key, fname in required_files.items() if key in file.name}

        st.session_state.uploaded_files = uploaded_files_dict

        missing_files = [fname for fname in required_files.values() if fname not in uploaded_files_dict]

        if not missing_files:
            st.success("All required files uploaded successfully!")
            navigate_to("OPD Creator")
        else:
            st.error(f"Missing files: {', '.join(missing_files)}. Please upload all required files.")

elif st.session_state.page == "OPD Creator":
	#test_date = datetime.datetime.strptime(x, "%m/%d/%Y")
	test_date = st.session_state.start_date
	uploaded_files = st.session_state.uploaded_files
	
	# initializing K
	K = 28
	 
	res = []
	 
	for day in range(K):
	    date = (test_date + datetime.timedelta(days = day)).strftime("%-m/%-d/%Y")
	    res.append(date)
	     
	#res
	
	dates = pd.DataFrame(res, columns =['dates'])
	
	dates['x'] = "y"
	
	dates['i'] = dates.index+1
	
	dates['x'] = dates['x'].astype(str)+dates['i'].astype(str)
	
	dates['x'] = dates['x'].astype(str) + "=" + "'"+dates['dates'].astype(str) + "'"
	
	dates = dates[['x']]
	
	dates.to_csv('dates.csv',index=False)
	
	import numpy as np
	
	datesdf = pd.read_csv('dates.csv')
	
	dates = datesdf['x'].astype(str)
	
	numpy_array=dates.to_numpy()
	np.savetxt("dates.py",numpy_array, fmt="%s")
	
	exec(open('dates.py').read())
	import xlsxwriter

	# Create workbook
	workbook = xlsxwriter.Workbook('OPD.xlsx')
	
	# Define worksheet names
	worksheet_names = [
	    'HOPE_DRIVE', 'ETOWN', 'NYES', 'COMPLEX', 'W_A', 'W_C',
	    'W_P', 'PICU', 'PSHCH_NURS', 'HAMPDEN_NURS',
	    'SJR_HOSP', 'AAC', 'ER_CONS','NF'
	]
	
	# Create worksheets and store them in a dictionary
	worksheets = {name: workbook.add_worksheet(name) for name in worksheet_names}
	
	# Assign worksheets to separate variables dynamically
	(
	    worksheet, worksheet2, worksheet3, worksheet4, worksheet5, 
	    worksheet6, worksheet7, worksheet8, worksheet9, worksheet10, 
	    worksheet11, worksheet12, worksheet13, worksheet14
	) = worksheets.values()
	
	# Define format
	format1 = workbook.add_format({
	    'font_size': 18, 'bold': 1, 'align': 'center',
	    'valign': 'vcenter', 'font_color': 'black',
	    'bg_color': '#FEFFCC', 'border': 1
	})
	
	# Define site names corresponding to worksheet names
	worksheet_sites = {
	    worksheet: 'Hope Drive',
	    worksheet2: 'Elizabethtown',
	    worksheet3: 'Nyes Road',
	    worksheet4: 'Complex Care',
	    worksheet5: 'WARD A',
	    worksheet6: 'WARD C',
	    worksheet7: 'WARD P',
	    worksheet8: 'PICU',
	    worksheet9: 'PSHCH NURSERY',
	    worksheet10: 'HAMPDEN NURSERY',
	    worksheet11: 'SJR HOSPITALIST',
	    worksheet12: 'AAC', worksheet13: 'ER CONSULTS', worksheet14: 'NIGHT FLOAT'
	}
	
	# Write "Site:" and corresponding site names in each worksheet
	for ws, site in worksheet_sites.items():
	    ws.write(0, 0, 'Site:', format1)
	    ws.write(0, 1, site, format1)
		
	#Color Coding
	format4 = workbook.add_format({'font_size':12,'bold': 1,'align': 'center','valign': 'vcenter','font_color':'black','bg_color':'#8ccf6f','border':1})
	format4a = workbook.add_format({'font_size':12,'bold': 1,'align': 'center','valign': 'vcenter','font_color':'black','bg_color':'#9fc5e8','border':1})    
	format5 = workbook.add_format({'font_size':12,'bold': 1,'align': 'center','valign': 'vcenter','font_color':'black','bg_color':'#FEFFCC','border':1})
	format5a = workbook.add_format({'font_size':12,'bold': 1,'align': 'center','valign': 'vcenter','font_color':'black','bg_color':'#d0e9ff','border':1})
	format11 = workbook.add_format({'font_size':18,'bold': 1,'align': 'center','valign': 'vcenter','font_color':'black','bg_color':'#FEFFCC','border':1})
	
	#H codes
	formate = workbook.add_format({'font_size':12,'bold': 0,'align': 'center','valign': 'vcenter','font_color':'white','border':0})
	
	# HOPE_DRIVE COLOR CODING AND IDENTIFYING ACUTE VERSUS CONTINUITY
	ranges_format1 = ['A8:H15', 'A32:H39', 'A56:H63', 'A80:H87']
	ranges_format5a = ['A18:H25', 'A42:H49', 'A66:H73', 'A90:H97']
	
	for cell_range in ranges_format1:
	    worksheet.conditional_format(cell_range, {'type': 'cell', 'criteria': '>=', 'value': 0, 'format': format1})
	
	for cell_range in ranges_format5a:
	    worksheet.conditional_format(cell_range, {'type': 'cell', 'criteria': '>=', 'value': 0, 'format': format5a})
	
	# HOPE_DRIVE CONDITIONAL FORMATTING
	ranges_format4 = ['A6:H6', 'A7:H7', 'A30:H30', 'A31:H31', 'A54:H54', 'A55:H55', 'A78:H78', 'A79:H79']
	ranges_format4a = ['A16:H16', 'A17:H17', 'A40:H40', 'A41:H41', 'A64:H64', 'A65:H65', 'A88:H88', 'A89:H89']
	
	for cell_range in ranges_format4:
	    worksheet.conditional_format(cell_range, {'type': 'cell', 'criteria': '>=', 'value': 0, 'format': format4})
	
	for cell_range in ranges_format4a:
	    worksheet.conditional_format(cell_range, {'type': 'cell', 'criteria': '>=', 'value': 0, 'format': format4a})
	
	# HOPE_DRIVE WRITING ACUTE AND CONTINUITY LABELS
	acute_format_ranges = [
	    (6, 7, 'AM - ACUTES', format4), (16, 17, 'PM - ACUTES', format4a), 
	    (30, 31, 'AM - ACUTES', format4), (40, 41, 'PM - ACUTES', format4a),
	    (54, 55, 'AM - ACUTES', format4), (64, 65, 'PM - ACUTES', format4a),
	    (78, 79, 'AM - ACUTES', format4), (88, 89, 'PM - ACUTES', format4a)
	]
	
	continuity_format_ranges = [
	    (8, 15, 'AM - Continuity', format5a), (18, 25, 'PM - Continuity', format5a),
	    (32, 39, 'AM - Continuity', format5a), (42, 49, 'PM - Continuity', format5a),
	    (56, 63, 'AM - Continuity', format5a), (66, 73, 'PM - Continuity', format5a),
	    (80, 87, 'AM - Continuity', format5a), (90, 97, 'PM - Continuity', format5a)
	]
	
	# Write Acute Labels
	for start_row, end_row, label, fmt in acute_format_ranges:
	    for row in range(start_row, end_row + 1):
	        worksheet.write(f'A{row}', label, fmt)
	
	# Write Continuity Labels
	for start_row, end_row, label, fmt in continuity_format_ranges:
	    for row in range(start_row, end_row + 1):
	        worksheet.write(f'A{row}', label, fmt)
	
	# Define the labels
	#labels = ['HX1', 'HX2', 'H2', 'H3', 'H4', 'H5', 'H6', 'H7', 'H8', 'H9', 'HXX1', 'HXX2', 'H12', 'H13', 'H14', 'H15', 'H16', 'H17', 'H18', 'H19']
	
	labels = ['H{}'.format(i) for i in range(20)]
	
	# Define the starting rows for each group
	start_rows = [6, 30, 54, 78]
	
	# Write the labels in each group
	for start_row in start_rows:
	    for i, label in enumerate(labels):
	        worksheet.write(f'I{start_row + i}', label, formate)
	
	# Simplify common formatting and label assignment for worksheets 2, 3, 4, 5
	worksheets = [worksheet2, worksheet3, worksheet4, worksheet5, worksheet6, worksheet7, worksheet8, worksheet9, worksheet10, worksheet11, worksheet12, worksheet13, worksheet14]
	
	ranges_format1 = ['A6:H15', 'A30:H39', 'A54:H63', 'A78:H87']
	ranges_format5a = ['A16:H25', 'A40:H49', 'A64:H73', 'A88:H97']
	specific_format_ranges = [
	    ('B6:H6', format4), ('B16:H16', format4a),
	    ('B30:H30', format4), ('B40:H40', format4a),
	    ('B54:H54', format4), ('B64:H64', format4a),
	    ('B78:H78', format4), ('B88:H88', format4a)
	]
	
	am_pm_labels = ['AM'] * 10 + ['PM'] * 10
	h_labels = ['H{}'.format(i) for i in range(20)]
	
	for worksheet in worksheets:
	    # Apply conditional formatting
	    for cell_range in ranges_format1:
	        worksheet.conditional_format(cell_range, {'type': 'cell', 'criteria': '>=', 'value': 0, 'format': format1})
	
	    for cell_range in ranges_format5a:
	        worksheet.conditional_format(cell_range, {'type': 'cell', 'criteria': '>=', 'value': 0, 'format': format5a})
	
	    for cell_range, fmt in specific_format_ranges:
	        worksheet.conditional_format(cell_range, {'type': 'cell', 'criteria': '>=', 'value': 0, 'format': fmt})
	
	    # Write AM/PM labels
	    sections = [(6, 25), (30, 49), (54, 73), (78, 97)]
	    for start_row, end_row in sections:
	        for i, label in enumerate(am_pm_labels):
	            worksheet.write(f'A{start_row + i}', label, format5a)
	
	    # Write H labels in column 'I'
	    start_rows = [6, 30, 54, 78]
	    for start_row in start_rows:
	        for i, label in enumerate(h_labels):
	            worksheet.write(f'I{start_row + i}', label, formate)
	
	# Loop through each worksheet in workbook
	for worksheet in workbook.worksheets():
	
	    # Set Zoom for all sheets
	    worksheet.set_zoom(80)
	
	    # Set Days
	    format3 = workbook.add_format({
	        'font_size': 12, 'bold': 1, 'align': 'center', 'valign': 'vcenter',
	        'font_color': 'black', 'bg_color': '#FFC7CE', 'border': 1
	    })
	    day_labels = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']
	    start_rows = [2, 26, 50, 74] #[3, 27, 51, 75]
	    for start_row in start_rows:
	        for i, day in enumerate(day_labels):
	            worksheet.write(start_row, 1 + i, day, format3)  # B=1, C=2, etc.
	
	    # Set Date Formats
	    format_date = workbook.add_format({
	        'num_format': 'm/d/yyyy', 'font_size': 12, 'bold': 1, 'align': 'center', 'valign': 'vcenter',
	        'font_color': 'black', 'bg_color': '#FFC7CE', 'border': 1
	    })
	
	    format_label = workbook.add_format({
	        'font_size': 12, 'bold': 1, 'align': 'center', 'valign': 'vcenter',
	        'font_color': 'black', 'bg_color': '#FFC7CE', 'border': 1
	    })
	
	    # Set Date Formulas
	    date_rows = [3, 27, 51, 75] #[4, 28, 52, 76]
	    for i, start_row in enumerate(date_rows):
	        worksheet.write(f'A{start_row - 1}', "", format_label)
	        worksheet.write_formula(f'A{start_row}', f'="Week of:"&" "&TEXT(B{start_row},"m/d/yy")', format_label)
	        worksheet.write(f'A{start_row + 1}', "", format_label)
	
	    # Set Pink Bars (Conditional Format)
	    pink_bar_rows = [5, 29, 53, 77]
	    for row in pink_bar_rows:
	        worksheet.conditional_format(f'A{row}:H{row}', {
	            'type': 'cell', 'criteria': '>=', 'value': 0, 'format': format_label
	        })
	
	    # Black Bars
	    format2 = workbook.add_format({'bg_color': 'black'})
	    black_bar_rows = [2, 26, 50, 74, 98]
	    for row in black_bar_rows:
	        worksheet.merge_range(f'A{row}:H{row}', " ", format2)
	        
	    # Write More Dates
	    date_values = [
	        [y1, y2, y3, y4, y5, y6, y7],
	        [y8, y9, y10, y11, y12, y13, y14],
	        [y15, y16, y17, y18, y19, y20, y21],
	        [y22, y23, y24, y25, y26, y27, y28]
	    ]
	    for i, start_row in enumerate(date_rows):
	        for j, value in enumerate(date_values[i]):
	            worksheet.write(start_row, 1 + j, value, format_date)  # B=1, C=2, etc.
	
	    # Set Column Widths
	    worksheet.set_column('A:A', 22)
	    worksheet.set_column('B:H', 40)
	    worksheet.set_row(0, 37.25)
	
	    # Merge Format for Text
	    merge_format = workbook.add_format({
	        'bold': 1, 'align': 'center', 'valign': 'vcenter', 'text_wrap': True,
	        'font_color': 'red', 'bg_color': '#FEFFCC', 'border': 1
	    })
	    text1 = 'Students are to alert their preceptors when they have a Clinical Reasoning Teaching Session (CRTS).  Please allow the students to leave approximately 15 minutes prior to the start of their session so they can be prepared to actively participate.  ~ Thank you!'
	
	    # Merge and Write Important Message
	    worksheet.merge_range('C1:F1', text1, merge_format)
	    worksheet.write('G1', "", merge_format)
	    worksheet.write('H1', "", merge_format)
	
	# Close Workbook
	workbook.close()
	
	####################################################################################################################################
	import pandas as pd
	import datetime
	from datetime import timedelta
	
	# Disable chained assignment warning
	pd.options.mode.chained_assignment = None
	
	# Parse the start date (assuming `test_date` is already defined)
	formatted_date = test_date.strftime('%m-%d-%Y')
	start_date = datetime.datetime.strptime(formatted_date, '%m-%d-%Y')
	
	# Calculate the end date (34 days from the start date)
	end_date = start_date + timedelta(days=34)
	
	# Generate the date range
	date_range = pd.date_range(start=start_date, end=end_date)
	
	# Create a DataFrame with the formatted dates
	xf201 = pd.DataFrame({'date': date_range})
	xf201['convert'] = xf201['date'].dt.strftime('%B %-d, %Y')
	
	xf201['t'] = "T"
	xf201['c'] = xf201.index+0
	xf201['T'] = xf201['t'].astype(str) + xf201 ['c'].astype(str)
	
	for i in range(35):
	    exec(f"day{i} = xf201['convert'][{i}]")

	column_pairs = [(0, 1), (2, 3), (4, 5), (6, 7), (8, 9), (10, 11), (12, 13)]
	
	days = [day0, day1, day2, day3, day4, day5, day6, day7, day8, day9, day10, day11, day12, day13,
	        day14, day15, day16, day17, day18, day19, day20, day21, day22, day23, day24, day25, day26, day27,
	        day28, day29, day30, day31, day32, day33, day34]
	
	# Function to process each file
	def process_file(file_key, clinic_name, replacements=None):
	    """Process an uploaded file and return cleaned DataFrame."""
	    
	    if file_key in uploaded_files:
	        df = pd.read_excel(uploaded_files[file_key], dtype=str)
	        df.rename(columns={col: str(i) for i, col in enumerate(df.columns)}, inplace=True)
	
	        D_dict = {}
	        for i in range(28):
	            col_idx = column_pairs[i % len(column_pairs)]
	            start_day = days[i]
	            end_day = days[i + 7]
	
	            start_idx = df.loc[df[str(col_idx[0])] == start_day].index[0]
	            end_idx = df.loc[df[str(col_idx[0])] == end_day].index[0]
	
	            extracted_data = df.iloc[start_idx + 1:end_idx, list(col_idx)].copy()
	            extracted_data.columns = ['type', 'provider']
	            extracted_data.insert(0, 'date', start_day)
	            extracted_data = extracted_data[:-1]
	
	            D_dict[f"D{i}"] = extracted_data
	
	        dfx = pd.concat(D_dict.values(), ignore_index=True)
	        dfx['clinic'] = clinic_name
	
	        if replacements:
	            dfx = dfx.replace(replacements, regex=True)
	
	        filename = f"{clinic_name.lower()}.csv"
	        dfx.to_csv(filename, index=False)
	
	        print(f"Processed {clinic_name} and saved to {filename}")
	        return dfx  # Return DataFrame for further processing
	    else:
	        return None  # Handle missing file case

	def duplicate_am_continuity(df, clinic_name):
	    if df is not None:
	        # Identify rows containing "AM - Continuity"
	        am_continuity_rows = df[df.eq("AM - Continuity").any(axis=1)].copy()
	
	        # Create corresponding "PM - Continuity" rows
	        pm_continuity_rows = am_continuity_rows.replace("AM - Continuity", "PM - Continuity")
	
	        # Append new rows to the original dataframe (first duplication)
	        df = pd.concat([df, pm_continuity_rows], ignore_index=True)
	
	        # Duplicate the AM - Continuity and PM - Continuity rows again
	        df = pd.concat([df, am_continuity_rows, pm_continuity_rows], ignore_index=True)
	
	        # Save the updated data
	        filename = f"{clinic_name.lower()}.csv"
	        df.to_csv(filename, index=False)
	        print(f"{clinic_name} updated with two AM - Continuity and two PM - Continuity entries and saved to {filename}.")
	    
	    return df

	def process_continuity_classes(df, clinic_name, am_csv, pm_csv):
	    if df is not None:
	        # AM - Continuity Processing
	        df[df['type'] == 'AM - Continuity '].assign(count=lambda x: x.groupby(['date'])['provider'].cumcount(),).assign(**{"class": lambda x: "H" + x['count'].astype(str)})[['date', 'type', 'provider', 'clinic', 'class']].to_csv(am_csv, index=False)
	        df[df['type'] == 'PM - Continuity '].assign(count=lambda x: x.groupby(['date'])['provider'].cumcount(),).assign(**{"class": lambda x: "H" + x['count'].astype(str)})[['date', 'type', 'provider', 'clinic', 'class']].to_csv(pm_csv, index=False)

	def process_hope_classes(df, clinic_name):
	    """
	    Processes Hope Drive's different continuity and acute types by assigning a count and class.
	    Saves the resulting DataFrame to separate CSV files based on type.
	    """
	    if df is not None:
	        hope_files = {
	            "AM - ACUTES": "5.csv",
	            "AM - ACUTES ": "6.csv",  # Handles potential trailing space issue
	            "PM - ACUTES ": "7.csv",
	            "AM - Continuity ": "8.csv",
	            "PM - Continuity ": "9.csv"
	        }
	
	        for type_key, filename in hope_files.items():
	            if type_key in df['type'].values:
	                subset_df = df[df['type'] == type_key].copy()
	
	                # Assign custom count logic based on type
	                if "AM - ACUTES" in type_key:
	                    subset_df['count'] = subset_df.groupby(['date'])['provider'].cumcount()
	                    subset_df['class'] = subset_df['count'].apply(
	                        lambda count: "H0" if count == 0 else ("H1" if count == 1 else "H" + str(count + 2))
	                    )
	                
	                elif "PM - ACUTES" in type_key:
	                    subset_df['count'] = subset_df.groupby(['date'])['provider'].cumcount()
	                    subset_df['class'] = subset_df['count'].apply(
	                        lambda count: "H10" if count == 0 else ("H11" if count == 1 else "H" + str(count + 12))
	                    )
	
	                elif "AM - Continuity" in type_key:
	                    subset_df['count'] = subset_df.groupby(['date'])['provider'].cumcount() + 2
	                    subset_df['class'] = "H" + subset_df['count'].astype(str)
	
	                elif "PM - Continuity" in type_key:
	                    subset_df['count'] = subset_df.groupby(['date'])['provider'].cumcount() + 12
	                    subset_df['class'] = "H" + subset_df['count'].astype(str)
	
	                # Keep only relevant columns
	                subset_df = subset_df[['date', 'type', 'provider', 'clinic', 'class']]
	
	                # Save to CSV
	                subset_df.to_csv(filename, index=False)
	                print(f"{clinic_name} {type_key} saved to {filename}.")

	# Define replacement rules for each clinic
	replacement_rules = {
	    "HOPE_DRIVE.xlsx": {
	        "Hope Drive AM Continuity": "AM - Continuity",
	        "Hope Drive PM Continuity": "PM - Continuity",
	        "Hope Drive\xa0AM Acute Precept ": "AM - ACUTES",  # Handles non-breaking space (\xa0)
	        "Hope Drive PM Acute Precept": "PM - ACUTES",
	        "Hope Drive Weekend Continuity": "AM - Continuity",
	        "Hope Drive Weekend Acute 1": "AM - ACUTES",
	        "Hope Drive Weekend Acute 2": "AM - ACUTES"
	    },
	    "PICU.xlsx": {
	        "2nd PICU Attending 7:45a-4p": "AM - Continuity",
	        "1st PICU Attending 7:30a-5p": "AM - Continuity"
	    },
	    "ETOWN.xlsx": {
	        "Etown AM Continuity": "AM - Continuity",
	        "Etown PM Continuity": "PM - Continuity"
	    },
	    "NYES.xlsx": {
	        "Nyes Rd AM Continuity": "AM - Continuity",
	        "Nyes Rd PM Continuity": "PM - Continuity"
	    },
	    "COMPLEX.xlsx": {
	        "Hope Drive Clinic AM": "AM - Continuity",
	        "Hope Drive Clinic PM": "PM - Continuity"
	    },
	    "WARD_A.xlsx": {
	        "Rounder 1 7a-7p": "AM - Continuity",
	        "Rounder 2 7a-7p": "AM - Continuity",
	        "Rounder 3 7a-7p": "AM - Continuity"
	    },
	    "WARD_P.xlsx": {
	        "On-Call 8a-8a": "AM - Continuity",
	        "On-Call": "AM - Continuity"
	    }
	}
	
	# Process each file

	hope_drive_df = process_file("HOPE_DRIVE.xlsx", "HOPE_DRIVE", replacement_rules.get("HOPE_DRIVE.xlsx"))
	etown_df = process_file("ETOWN.xlsx", "ETOWN", replacement_rules.get("ETOWN.xlsx"))
	nyes_df = process_file("NYES.xlsx", "NYES", replacement_rules.get("NYES.xlsx"))
	complex_df = process_file("COMPLEX.xlsx", "COMPLEX", replacement_rules.get("COMPLEX.xlsx"))
	
	warda_df = process_file("WARD_A.xlsx", "WARD_A", replacement_rules.get("WARD_A.xlsx"))
	wardp_df = process_file("WARD_P.xlsx", "WARD_P", replacement_rules.get("WARD_P.xlsx"))
	picu_df = process_file("PICU.xlsx", "PICU", replacement_rules.get("PICU.xlsx"))

	process_hope_classes(hope_drive_df, "HOPE_DRIVE")
	
	# Apply AM → PM Continuity Transformation for WARDA, WARDP, and PICU
	warda_df = duplicate_am_continuity(warda_df, "WARD_A")
	wardp_df = duplicate_am_continuity(wardp_df, "WARD_P")
	picu_df = duplicate_am_continuity(picu_df, "PICU")

	process_continuity_classes(etown_df, "ETOWN", "1.csv", "2.csv")
	process_continuity_classes(nyes_df, "NYES", "3.csv", "4.csv")
	process_continuity_classes(complex_df, "COMPLEX", "10.csv", "11.csv")
	
	process_continuity_classes(warda_df, "WARD_A", "12.csv", "13.csv")
	process_continuity_classes(wardp_df, "WARD_P", "14.csv", "15.csv")
	process_continuity_classes(picu_df, "PICU", "16.csv", "17.csv")
	
	############################################################################################################################
	tables = {f"t{i}": pd.read_csv(f"{i}.csv") for i in range(1, 18)} #Add +1 to 18... so if adding t18, t19... then add 2 to 18... and its 20. Or 1 plus the last t value... t17?... last number in range should be 18
	t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13, t14, t15, t16, t17 = tables.values()
	
	final2 = pd.DataFrame(columns=t1.columns)
	final2 = pd.concat([final2] + list(tables.values()), ignore_index=True)
	final2.to_csv('final2.csv',index=False)
	
	df=pd.read_csv('final2.csv',dtype=str) #MAP to Final2
	
	df['date'] = pd.to_datetime(df['date'])
	df['date'] = df['date'].dt.strftime('%m/%d/%Y')
	
	import csv
	
	dateMAP = xf201[['date','T']]
	
	dateMAP['date'] = pd.to_datetime(dateMAP['date'])
	dateMAP['date'] = dateMAP['date'].dt.strftime('%m/%d/%Y')
	
	dateMAP.to_csv('xxxDATEMAP.csv',index=False)
	
	mydict = {}
	with open('xxxDATEMAP.csv', mode='r')as inp:     #file is the objects you want to map. I want to map the IMP in this file to diagnosis.csv
		reader = csv.reader(inp)
		df1 = {rows[0]:rows[1] for rows in reader} 
	df['datecode'] = df.date.map(df1)               #'type' is the new column in the diagnosis file. 'encounter_id' is the key you are using to MAP 
	
	df['text'] = df['provider'] + " ~ "
	
	df['student'] = ""
	
	df = df.loc[:, ('date','type','provider','student','clinic','text','class','datecode')]
	
	df.to_csv('final.csv',index=False)
	df.to_excel('final.xlsx',index=False)

	########################################################################################################################################################################
	import openpyxl
	from openpyxl.styles import Alignment
	
	def generate_mapping(start_value):
	    """
	    Generates a mapping dictionary for H0 to H19 starting at a given start_value.
	    """
	    return {f"H{i}": start_value + i for i in range(20)}
	
	def create_t_mapping():
	    """
	    Creates the combined mapping for T0 to T27.
	    """
	    t_mappings = [
	        (0, 6),  # T0 to T6 starts at 6
	        (7, 30),  # T7 to T13 starts at 30
	        (14, 54),  # T14 to T20 starts at 54
	        (21, 78)   # T21 to T27 starts at 78
	    ]
	
	    combined_mapping = {}
	    for start_t, start_value in t_mappings:
	        common_mapping = generate_mapping(start_value)
	        combined_mapping.update({f"T{i}": common_mapping for i in range(start_t, start_t + 7)})
	
	    return combined_mapping
	
	def process_excel_mapping(location, sheet_name):
	    """
	    Processes an Excel sheet for a given location and writes data to the corresponding OPD sheet.
	    """
	    wb = openpyxl.load_workbook('final.xlsx')
	    ws = wb['Sheet1']
	    
	    wb1 = openpyxl.load_workbook('OPD.xlsx')
	    ws1 = wb1[sheet_name]
	
	    combined_t_mapping = create_t_mapping()
	
	    column_mapping = {f"T{i}": (i % 7) + 2 for i in range(28)}
	
	    for row in ws.iter_rows():
	        t_value = row[7].value  # Column H (index 7)
	        h_value = row[6].value  # Column G (index 6)
	        row_location = row[4].value  # Column E (index 4)
	
	        if row_location == location and t_value in combined_t_mapping and h_value in combined_t_mapping[t_value]:
	            target_row = combined_t_mapping[t_value][h_value]
	            target_column = column_mapping[t_value]
	
	            ws1.cell(row=target_row, column=target_column).value = row[5].value  # Column F (index 5)
	            ws1.cell(row=target_row, column=target_column).alignment = Alignment(horizontal='center')
	
	    wb1.save('OPD.xlsx')
	    print(f"Processed mapping for {location} in {sheet_name}.")

	# Process HOPE_DRIVE
	process_excel_mapping("HOPE_DRIVE", "HOPE_DRIVE")
	process_excel_mapping("ETOWN", "ETOWN")
	process_excel_mapping("NYES", "NYES")
	process_excel_mapping("COMPLEX", "COMPLEX")
	process_excel_mapping("WARD_A", "W_A")
	process_excel_mapping("WARD_P", "W_P")
	process_excel_mapping("PICU", "PICU")
	###############################################################################################

	# Button to trigger the download
	if st.button('Create OPD'):
	    # Path to the existing 'OPD.xlsx' workbook
	    file_path = 'OPD.xlsx'  # Replace with your file path if it's stored somewhere else
	
	    # Read the workbook into memory
	    with open(file_path, 'rb') as file:
	        file_data = file.read()
	
	    # Provide a download button for the existing OPD.xlsx file
	    st.download_button(
	        label="Download OPD.xlsx",
	        data=file_data,
	        file_name="OPD.xlsx",
	        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
	    )
		
elif st.session_state.page == "Create Student Schedule":
    st.title("Create Student Schedule")
    # Upload the OPD.xlsx file
    uploaded_opd_file = st.file_uploader("Upload OPD.xlsx file", type="xlsx")
    uploaded_book4_file = st.file_uploader("Upload Book4.xlsx file", type="xlsx")
    
    if uploaded_opd_file:
        try:
            # Read the uploaded OPD file into a pandas dataframe
            df_opd = pd.read_excel(uploaded_opd_file)
            st.write("File successfully uploaded and loaded.")
            
            # Store the uploaded file in session state for use later
            st.session_state.uploaded_files['OPD.xlsx'] = uploaded_opd_file
                
        except Exception as e:
            st.error(f"Error reading the uploaded file: {e}")
    else:
        st.write("Please upload the OPD.xlsx file to proceed.")

    if uploaded_book4_file:
        try:
            # Read the uploaded OPD file into a pandas dataframe
            df_opd = pd.read_excel(uploaded_book4_file)
            st.write("File successfully uploaded and loaded.")
            
            # Store the uploaded file in session state for use later
            st.session_state.uploaded_book4_file['Book4.xlsx'] = uploaded_book4_file
                
        except Exception as e:
            st.error(f"Error reading the uploaded file: {e}")
    else:
        st.write("Please upload the Book4.xlsx file to proceed.")

    # Button to go to the next page
    if st.button("Load Student Schedule"):
        st.session_state.page = "Create List"  # Update the session state to go to the next page
        st.rerun()  # Use st.rerun() instead of st.experimental_rerun() to force rerun and update the page

elif st.session_state.page == "Create List":
    st.title("Load Student Schedule")

    # Ensure the OPD.xlsx file exists in the session state before proceeding
    if 'OPD.xlsx' in st.session_state.uploaded_files:
        uploaded_opd_file = st.session_state.uploaded_files['OPD.xlsx']
        
        try:
            # Read the OPD file into a dataframe
            df_opd = pd.read_excel(uploaded_opd_file)
            
            # Display the first few rows of the OPD data for verification
            #st.dataframe(df_opd.head())
            
            # Save the OPD file again without the index column
            df_opd.to_excel('OPD.xlsx', index=False)
            st.write("OPD.xlsx file has been successfully saved.")
        
        except Exception as e:
            st.error(f"Error processing the OPD file: {e}")
    else:
        st.error("No OPD file found in session state.")


    if 'Book4.xlsx' in st.session_state.uploaded_book4_file:
        uploaded_book4_file = st.session_state.uploaded_book4_file['Book4.xlsx']
        
        try:
            # Read the OPD file into a dataframe
            book4 = pd.read_excel(uploaded_book4_file)
            
            # Display the first few rows of the OPD data for verification
            #st.dataframe(book4.head())
            
            # Save the OPD file again without the index column
            book4.to_excel('Book4.xlsx', index=False)
            st.write("Book4.xlsx file has been successfully saved.")
        
        except Exception as e:
            st.error(f"Error processing the Book4 file: {e}")
    else:
        st.error("No Book4 file found in session state.")
	    
    
    # Ensure the "HOPE_DRIVE" sheet exists in the uploaded Excel file
    try:
        df = pd.read_excel(uploaded_opd_file)

        # Extract the value from row 2, column 1
        test_date = df.iloc[2, 1]

        # Ensure that test_date is a valid datetime object
        # If it's a string, convert it into a datetime object using pd.to_datetime
        if isinstance(test_date, str):
            test_date = pd.to_datetime(test_date, errors='coerce')  # Handle invalid date gracefully

        # Check if the date is valid (not NaT)
        if pd.isna(test_date):
            print("Invalid date format in the cell.")
        else:
            # Format the date to mm-dd-yyyy
            formatted_date = test_date.strftime('%m-%d-%Y')

            # Calculate the start date (use test_date directly as it's already a datetime object)
            start_date = test_date

            # Calculate the end date (34 days from the start date)
            end_date = start_date + timedelta(days=34)

            # Generate the date range
            date_range = pd.date_range(start=start_date, end=end_date)

            # Create a DataFrame with the formatted dates
            xf201 = pd.DataFrame({'date': date_range})

            # Convert the date to the desired format (e.g., "November 25, 2024")
            xf201['convert'] = xf201['date'].dt.strftime('%B %-d, %Y')  # This works in Unix-like systems

            # Add additional columns
            xf201['t'] = "T"
            xf201['c'] = xf201.index + 0  # Simple index-based column
            xf201['T'] = xf201['t'].astype(str) + xf201['c'].astype(str)

        # Creating the dateMAP DataFrame
        dateMAP = xf201[['date', 'T']].copy()  # Use .copy() to avoid the SettingWithCopyWarning

        # Convert 'date' column to datetime and then format it
        dateMAP['date'] = pd.to_datetime(dateMAP['date'])
        dateMAP['date'] = dateMAP['date'].dt.strftime('%m/%d/%Y')

        # Save the result to a CSV file
        dateMAP.to_csv('xxxDATEMAP.csv', index=False)
        
        read_file = pd.read_excel(uploaded_opd_file, sheet_name='HOPE_DRIVE')
        read_file.to_csv ('hopedrive.csv', index = False, header=False)
        df=pd.read_csv('hopedrive.csv')
        clinictype=df.iloc[3:23, 0:1]
        a1 = pd.DataFrame(clinictype, columns = ['type'])
        a2 = pd.DataFrame(clinictype, columns = ['type'])
        a3 = pd.DataFrame(clinictype, columns = ['type'])
        a4 = pd.DataFrame(clinictype, columns = ['type'])
        a5 = pd.DataFrame(clinictype, columns = ['type'])
        a6 = pd.DataFrame(clinictype, columns = ['type'])
        a7 = pd.DataFrame(clinictype, columns = ['type'])

        a1['type']=clinictype
        a2['type']=clinictype
        a3['type']=clinictype
        a4['type']=clinictype
        a5['type']=clinictype
        a6['type']=clinictype
        a7['type']=clinictype

        week1day1=a1.replace(to_replace=r'- Continuity', value='', regex=True)
        week1day2=a2.replace(to_replace=r'- Continuity', value='', regex=True)
        week1day3=a3.replace(to_replace=r'- Continuity', value='', regex=True)
        week1day4=a4.replace(to_replace=r'- Continuity', value='', regex=True)
        week1day5=a5.replace(to_replace=r'- Continuity', value='', regex=True)
        week1day6=a6.replace(to_replace=r'- Continuity', value='', regex=True)
        week1day7=a7.replace(to_replace=r'- Continuity', value='', regex=True)
        
        day1=df.iloc[1,1]
        day2=df.iloc[1,2]
        day3=df.iloc[1,3]
        day4=df.iloc[1,4]
        day5=df.iloc[1,5]
        day6=df.iloc[1,6]
        day7=df.iloc[1,7]

        week1day1['date']=day1
        week1day2['date']=day2
        week1day3['date']=day3
        week1day4['date']=day4
        week1day5['date']=day5
        week1day6['date']=day6
        week1day7['date']=day7

        provider1=df.iloc[3:23,1]
        provider2=df.iloc[3:23,2]
        provider3=df.iloc[3:23,3]
        provider4=df.iloc[3:23,4]
        provider5=df.iloc[3:23,5]
        provider6=df.iloc[3:23,6]
        provider7=df.iloc[3:23,7]

        week1day1['provider']=provider1
        week1day2['provider']=provider2
        week1day3['provider']=provider3
        week1day4['provider']=provider4
        week1day5['provider']=provider5
        week1day6['provider']=provider6
        week1day7['provider']=provider7

        week1day1['clinic']="HOPE_DRIVE"
        week1day2['clinic']="HOPE_DRIVE"
        week1day3['clinic']="HOPE_DRIVE"
        week1day4['clinic']="HOPE_DRIVE"
        week1day5['clinic']="HOPE_DRIVE"
        week1day6['clinic']="HOPE_DRIVE"
        week1day7['clinic']="HOPE_DRIVE"

        week1=pd.DataFrame(columns=week1day1.columns)
        week1=pd.concat([week1,week1day1,week1day2,week1day3,week1day4,week1day5,week1day6,week1day7])
        week1.to_csv('week1.csv',index=False)

        clinictype=df.iloc[27:47, 0:1]
        b1 = pd.DataFrame(clinictype, columns = ['type'])
        b2 = pd.DataFrame(clinictype, columns = ['type'])
        b3 = pd.DataFrame(clinictype, columns = ['type'])
        b4 = pd.DataFrame(clinictype, columns = ['type'])
        b5 = pd.DataFrame(clinictype, columns = ['type'])
        b6 = pd.DataFrame(clinictype, columns = ['type'])
        b7 = pd.DataFrame(clinictype, columns = ['type'])

        b1['type']=clinictype
        b2['type']=clinictype
        b3['type']=clinictype
        b4['type']=clinictype
        b5['type']=clinictype
        b6['type']=clinictype
        b7['type']=clinictype

        week2day1=b1.replace(to_replace=r'- Continuity', value='', regex=True)
        week2day2=b2.replace(to_replace=r'- Continuity', value='', regex=True)
        week2day3=b3.replace(to_replace=r'- Continuity', value='', regex=True)
        week2day4=b4.replace(to_replace=r'- Continuity', value='', regex=True)
        week2day5=b5.replace(to_replace=r'- Continuity', value='', regex=True)
        week2day6=b6.replace(to_replace=r'- Continuity', value='', regex=True)
        week2day7=b7.replace(to_replace=r'- Continuity', value='', regex=True)

        day1=df.iloc[25,1]
        day2=df.iloc[25,2]
        day3=df.iloc[25,3]
        day4=df.iloc[25,4]
        day5=df.iloc[25,5]
        day6=df.iloc[25,6]
        day7=df.iloc[25,7]

        week2day1['date']=day1
        week2day2['date']=day2
        week2day3['date']=day3
        week2day4['date']=day4
        week2day5['date']=day5
        week2day6['date']=day6
        week2day7['date']=day7

        provider1=df.iloc[27:47,1]
        provider2=df.iloc[27:47,2]
        provider3=df.iloc[27:47,3]
        provider4=df.iloc[27:47,4]
        provider5=df.iloc[27:47,5]
        provider6=df.iloc[27:47,6]
        provider7=df.iloc[27:47,7]

        week2day1['provider']=provider1
        week2day2['provider']=provider2
        week2day3['provider']=provider3
        week2day4['provider']=provider4
        week2day5['provider']=provider5
        week2day6['provider']=provider6
        week2day7['provider']=provider7

        week2day1['clinic']="HOPE_DRIVE"
        week2day2['clinic']="HOPE_DRIVE"
        week2day3['clinic']="HOPE_DRIVE"
        week2day4['clinic']="HOPE_DRIVE"
        week2day5['clinic']="HOPE_DRIVE"
        week2day6['clinic']="HOPE_DRIVE"
        week2day7['clinic']="HOPE_DRIVE"

        week2=pd.DataFrame(columns=week2day1.columns)
        week2=pd.concat([week2,week2day1,week2day2,week2day3,week2day4,week2day5,week2day6,week2day7])
        week2.to_csv('week2.csv',index=False)

        clinictype=df.iloc[51:71, 0:1]
        c1 = pd.DataFrame(clinictype, columns = ['type'])
        c2 = pd.DataFrame(clinictype, columns = ['type'])
        c3 = pd.DataFrame(clinictype, columns = ['type'])
        c4 = pd.DataFrame(clinictype, columns = ['type'])
        c5 = pd.DataFrame(clinictype, columns = ['type'])
        c6 = pd.DataFrame(clinictype, columns = ['type'])
        c7 = pd.DataFrame(clinictype, columns = ['type'])

        c1['type']=clinictype
        c2['type']=clinictype
        c3['type']=clinictype
        c4['type']=clinictype
        c5['type']=clinictype
        c6['type']=clinictype
        c7['type']=clinictype

        week3day1=c1.replace(to_replace=r'- Continuity', value='', regex=True)
        week3day2=c2.replace(to_replace=r'- Continuity', value='', regex=True)
        week3day3=c3.replace(to_replace=r'- Continuity', value='', regex=True)
        week3day4=c4.replace(to_replace=r'- Continuity', value='', regex=True)
        week3day5=c5.replace(to_replace=r'- Continuity', value='', regex=True)
        week3day6=c6.replace(to_replace=r'- Continuity', value='', regex=True)
        week3day7=c7.replace(to_replace=r'- Continuity', value='', regex=True)

        day1=df.iloc[49,1]
        day2=df.iloc[49,2]
        day3=df.iloc[49,3]
        day4=df.iloc[49,4]
        day5=df.iloc[49,5]
        day6=df.iloc[49,6]
        day7=df.iloc[49,7]

        week3day1['date']=day1
        week3day2['date']=day2
        week3day3['date']=day3
        week3day4['date']=day4
        week3day5['date']=day5
        week3day6['date']=day6
        week3day7['date']=day7

        provider1=df.iloc[51:71,1]
        provider2=df.iloc[51:71,2]
        provider3=df.iloc[51:71,3]
        provider4=df.iloc[51:71,4]
        provider5=df.iloc[51:71,5]
        provider6=df.iloc[51:71,6]
        provider7=df.iloc[51:71,7]

        week3day1['provider']=provider1
        week3day2['provider']=provider2
        week3day3['provider']=provider3
        week3day4['provider']=provider4
        week3day5['provider']=provider5
        week3day6['provider']=provider6
        week3day7['provider']=provider7

        week3day1['clinic']="HOPE_DRIVE"
        week3day2['clinic']="HOPE_DRIVE"
        week3day3['clinic']="HOPE_DRIVE"
        week3day4['clinic']="HOPE_DRIVE"
        week3day5['clinic']="HOPE_DRIVE"
        week3day6['clinic']="HOPE_DRIVE"
        week3day7['clinic']="HOPE_DRIVE"

        week3=pd.DataFrame(columns=week3day1.columns)
        week3=pd.concat([week3,week3day1,week3day2,week3day3,week3day4,week3day5,week3day6,week3day7])
        week3.to_csv('week3.csv',index=False)

        clinictype=df.iloc[75:95, 0:1]
        d1 = pd.DataFrame(clinictype, columns = ['type'])
        d2 = pd.DataFrame(clinictype, columns = ['type'])
        d3 = pd.DataFrame(clinictype, columns = ['type'])
        d4 = pd.DataFrame(clinictype, columns = ['type'])
        d5 = pd.DataFrame(clinictype, columns = ['type'])
        d6 = pd.DataFrame(clinictype, columns = ['type'])
        d7 = pd.DataFrame(clinictype, columns = ['type'])

        d1['type']=clinictype
        d2['type']=clinictype
        d3['type']=clinictype
        d4['type']=clinictype
        d5['type']=clinictype
        d6['type']=clinictype
        d7['type']=clinictype

        week4day1=d1.replace(to_replace=r'- Continuity', value='', regex=True)
        week4day2=d2.replace(to_replace=r'- Continuity', value='', regex=True)
        week4day3=d3.replace(to_replace=r'- Continuity', value='', regex=True)
        week4day4=d4.replace(to_replace=r'- Continuity', value='', regex=True)
        week4day5=d5.replace(to_replace=r'- Continuity', value='', regex=True)
        week4day6=d6.replace(to_replace=r'- Continuity', value='', regex=True)
        week4day7=d7.replace(to_replace=r'- Continuity', value='', regex=True)


        day1=df.iloc[73,1]
        day2=df.iloc[73,2]
        day3=df.iloc[73,3]
        day4=df.iloc[73,4]
        day5=df.iloc[73,5]
        day6=df.iloc[73,6]
        day7=df.iloc[73,7]

        week4day1['date']=day1
        week4day2['date']=day2
        week4day3['date']=day3
        week4day4['date']=day4
        week4day5['date']=day5
        week4day6['date']=day6
        week4day7['date']=day7

        provider1=df.iloc[75:95,1]
        provider2=df.iloc[75:95,2]
        provider3=df.iloc[75:95,3]
        provider4=df.iloc[75:95,4]
        provider5=df.iloc[75:95,5]
        provider6=df.iloc[75:95,6]
        provider7=df.iloc[75:95,7]

        week4day1['provider']=provider1
        week4day2['provider']=provider2
        week4day3['provider']=provider3
        week4day4['provider']=provider4
        week4day5['provider']=provider5
        week4day6['provider']=provider6
        week4day7['provider']=provider7

        week4day1['clinic']="HOPE_DRIVE"
        week4day2['clinic']="HOPE_DRIVE"
        week4day3['clinic']="HOPE_DRIVE"
        week4day4['clinic']="HOPE_DRIVE"
        week4day5['clinic']="HOPE_DRIVE"
        week4day6['clinic']="HOPE_DRIVE"
        week4day7['clinic']="HOPE_DRIVE"

        week4=pd.DataFrame(columns=week4day1.columns)
        week4=pd.concat([week4,week4day1,week4day2,week4day3,week4day4,week4day5,week4day6,week4day7])
        week4.to_csv('week4.csv',index=False)

        hope=pd.DataFrame(columns=week1.columns)
        hope=pd.concat([hope,week1,week2,week3,week4])

        hope.to_csv('hope.csv',index=False)

        # Handle AM Continuity for hopei
        hope['H'] = "H"
        hopei = hope[hope['type'] == 'AM '].copy()  # Ensure we're working with a copy of the slice
        hopei['count'] = hopei.groupby(['date'])['provider'].cumcount() + 2  # Starts at H2 for AM
        hopei['class'] = "H" + hopei['count'].astype(str)
        hopei = hopei.loc[:, ('date', 'type', 'provider', 'clinic', 'class')]
        hopei.to_csv('5.csv', index=False)

        # Handle PM Continuity for hopeii
        hopeii = hope[hope['type'] == 'PM '].copy()  # Ensure we're working with a copy of the slice
        hopeii['count'] = hopeii.groupby(['date'])['provider'].cumcount() + 12  # Starts at H12 for PM
        hopeii['class'] = "H" + hopeii['count'].astype(str)
        hopeii = hopeii.loc[:, ('date', 'type', 'provider', 'clinic', 'class')]
        hopeii.to_csv('6.csv', index=False)

        # Handle AM - ACUTES for hopeiii
        hopeiii = hope[hope['type'] == 'AM - ACUTES'].copy()  # Ensure we're working with a copy of the slice
        hopeiii['count'] = hopeiii.groupby(['date'])['provider'].cumcount()  # Starts at H0 for AM-ACUTES
        hopeiii['class'] = "H" + hopeiii['count'].astype(str)
        hopeiii = hopeiii.loc[:, ('date', 'type', 'provider', 'clinic', 'class')]
        hopeiii.to_csv('7.csv', index=False)

        # Handle PM - ACUTES for hopeiiii
        hopeiiii = hope[hope['type'] == 'PM - ACUTES'].copy()  # Ensure we're working with a copy of the slice
        hopeiiii['count'] = hopeiiii.groupby(['date'])['provider'].cumcount() + 10  # Starts at H10 for PM-ACUTES
        hopeiiii['class'] = "H" + hopeiiii['count'].astype(str)
        hopeiiii = hopeiiii.loc[:, ('date', 'type', 'provider', 'clinic', 'class')]
        hopeiiii.to_csv('8.csv', index=False)

        # Combine all the data into one DataFrame
        hopes = pd.DataFrame(columns=hopei.columns)
        hopes = pd.concat([hopei, hopeii, hopeiii, hopeiiii])

        # Save the combined DataFrame to CSV
        hopes.to_csv('hopes.csv', index=False)
	    
        ####################################NYES#############################################################################
        import pandas as pd
        read_file = pd.read_excel (uploaded_opd_file, sheet_name='NYES')
        read_file.to_csv ('nyesroad.csv', index = False, header=False)
        df=pd.read_csv('nyesroad.csv')

        clinictype=df.iloc[3:23, 0:1]
        a1 = pd.DataFrame(clinictype, columns = ['type'])
        a2 = pd.DataFrame(clinictype, columns = ['type'])
        a3 = pd.DataFrame(clinictype, columns = ['type'])
        a4 = pd.DataFrame(clinictype, columns = ['type'])
        a5 = pd.DataFrame(clinictype, columns = ['type'])
        a6 = pd.DataFrame(clinictype, columns = ['type'])
        a7 = pd.DataFrame(clinictype, columns = ['type'])

        a1['type']=clinictype
        a2['type']=clinictype
        a3['type']=clinictype
        a4['type']=clinictype
        a5['type']=clinictype
        a6['type']=clinictype
        a7['type']=clinictype

        week1day1=a1.replace(to_replace=r'- Continuity', value='', regex=True)
        week1day2=a2.replace(to_replace=r'- Continuity', value='', regex=True)
        week1day3=a3.replace(to_replace=r'- Continuity', value='', regex=True)
        week1day4=a4.replace(to_replace=r'- Continuity', value='', regex=True)
        week1day5=a5.replace(to_replace=r'- Continuity', value='', regex=True)
        week1day6=a6.replace(to_replace=r'- Continuity', value='', regex=True)
        week1day7=a7.replace(to_replace=r'- Continuity', value='', regex=True)

        day1=df.iloc[1,1]
        day2=df.iloc[1,2]
        day3=df.iloc[1,3]
        day4=df.iloc[1,4]
        day5=df.iloc[1,5]
        day6=df.iloc[1,6]
        day7=df.iloc[1,7]

        week1day1['date']=day1
        week1day2['date']=day2
        week1day3['date']=day3
        week1day4['date']=day4
        week1day5['date']=day5
        week1day6['date']=day6
        week1day7['date']=day7

        provider1=df.iloc[3:23,1]
        provider2=df.iloc[3:23,2]
        provider3=df.iloc[3:23,3]
        provider4=df.iloc[3:23,4]
        provider5=df.iloc[3:23,5]
        provider6=df.iloc[3:23,6]
        provider7=df.iloc[3:23,7]

        week1day1['provider']=provider1
        week1day2['provider']=provider2
        week1day3['provider']=provider3
        week1day4['provider']=provider4
        week1day5['provider']=provider5
        week1day6['provider']=provider6
        week1day7['provider']=provider7

        week1day1['clinic']="NYES"
        week1day2['clinic']="NYES"
        week1day3['clinic']="NYES"
        week1day4['clinic']="NYES"
        week1day5['clinic']="NYES"
        week1day6['clinic']="NYES"
        week1day7['clinic']="NYES"

        week1=pd.DataFrame(columns=week1day1.columns)
        week1=pd.concat([week1,week1day1,week1day2,week1day3,week1day4,week1day5,week1day6,week1day7])
        week1.to_csv('week1.csv',index=False)

        clinictype=df.iloc[27:47, 0:1]
        b1 = pd.DataFrame(clinictype, columns = ['type'])
        b2 = pd.DataFrame(clinictype, columns = ['type'])
        b3 = pd.DataFrame(clinictype, columns = ['type'])
        b4 = pd.DataFrame(clinictype, columns = ['type'])
        b5 = pd.DataFrame(clinictype, columns = ['type'])
        b6 = pd.DataFrame(clinictype, columns = ['type'])
        b7 = pd.DataFrame(clinictype, columns = ['type'])

        b1['type']=clinictype
        b2['type']=clinictype
        b3['type']=clinictype
        b4['type']=clinictype
        b5['type']=clinictype
        b6['type']=clinictype
        b7['type']=clinictype

        week2day1=b1.replace(to_replace=r'- Continuity', value='', regex=True)
        week2day2=b2.replace(to_replace=r'- Continuity', value='', regex=True)
        week2day3=b3.replace(to_replace=r'- Continuity', value='', regex=True)
        week2day4=b4.replace(to_replace=r'- Continuity', value='', regex=True)
        week2day5=b5.replace(to_replace=r'- Continuity', value='', regex=True)
        week2day6=b6.replace(to_replace=r'- Continuity', value='', regex=True)
        week2day7=b7.replace(to_replace=r'- Continuity', value='', regex=True)

        day1=df.iloc[25,1]
        day2=df.iloc[25,2]
        day3=df.iloc[25,3]
        day4=df.iloc[25,4]
        day5=df.iloc[25,5]
        day6=df.iloc[25,6]
        day7=df.iloc[25,7]

        week2day1['date']=day1
        week2day2['date']=day2
        week2day3['date']=day3
        week2day4['date']=day4
        week2day5['date']=day5
        week2day6['date']=day6
        week2day7['date']=day7

        provider1=df.iloc[27:47,1]
        provider2=df.iloc[27:47,2]
        provider3=df.iloc[27:47,3]
        provider4=df.iloc[27:47,4]
        provider5=df.iloc[27:47,5]
        provider6=df.iloc[27:47,6]
        provider7=df.iloc[27:47,7]

        week2day1['provider']=provider1
        week2day2['provider']=provider2
        week2day3['provider']=provider3
        week2day4['provider']=provider4
        week2day5['provider']=provider5
        week2day6['provider']=provider6
        week2day7['provider']=provider7

        week2day1['clinic']="NYES"
        week2day2['clinic']="NYES"
        week2day3['clinic']="NYES"
        week2day4['clinic']="NYES"
        week2day5['clinic']="NYES"
        week2day6['clinic']="NYES"
        week2day7['clinic']="NYES"

        week2=pd.DataFrame(columns=week2day1.columns)
        week2=pd.concat([week2,week2day1,week2day2,week2day3,week2day4,week2day5,week2day6,week2day7])
        week2.to_csv('week2.csv',index=False)

        clinictype=df.iloc[51:71, 0:1]
        c1 = pd.DataFrame(clinictype, columns = ['type'])
        c2 = pd.DataFrame(clinictype, columns = ['type'])
        c3 = pd.DataFrame(clinictype, columns = ['type'])
        c4 = pd.DataFrame(clinictype, columns = ['type'])
        c5 = pd.DataFrame(clinictype, columns = ['type'])
        c6 = pd.DataFrame(clinictype, columns = ['type'])
        c7 = pd.DataFrame(clinictype, columns = ['type'])

        c1['type']=clinictype
        c2['type']=clinictype
        c3['type']=clinictype
        c4['type']=clinictype
        c5['type']=clinictype
        c6['type']=clinictype
        c7['type']=clinictype

        week3day1=c1.replace(to_replace=r'- Continuity', value='', regex=True)
        week3day2=c2.replace(to_replace=r'- Continuity', value='', regex=True)
        week3day3=c3.replace(to_replace=r'- Continuity', value='', regex=True)
        week3day4=c4.replace(to_replace=r'- Continuity', value='', regex=True)
        week3day5=c5.replace(to_replace=r'- Continuity', value='', regex=True)
        week3day6=c6.replace(to_replace=r'- Continuity', value='', regex=True)
        week3day7=c7.replace(to_replace=r'- Continuity', value='', regex=True)

        day1=df.iloc[49,1]
        day2=df.iloc[49,2]
        day3=df.iloc[49,3]
        day4=df.iloc[49,4]
        day5=df.iloc[49,5]
        day6=df.iloc[49,6]
        day7=df.iloc[49,7]

        week3day1['date']=day1
        week3day2['date']=day2
        week3day3['date']=day3
        week3day4['date']=day4
        week3day5['date']=day5
        week3day6['date']=day6
        week3day7['date']=day7

        provider1=df.iloc[51:71,1]
        provider2=df.iloc[51:71,2]
        provider3=df.iloc[51:71,3]
        provider4=df.iloc[51:71,4]
        provider5=df.iloc[51:71,5]
        provider6=df.iloc[51:71,6]
        provider7=df.iloc[51:71,7]

        week3day1['provider']=provider1
        week3day2['provider']=provider2
        week3day3['provider']=provider3
        week3day4['provider']=provider4
        week3day5['provider']=provider5
        week3day6['provider']=provider6
        week3day7['provider']=provider7

        week3day1['clinic']="NYES"
        week3day2['clinic']="NYES"
        week3day3['clinic']="NYES"
        week3day4['clinic']="NYES"
        week3day5['clinic']="NYES"
        week3day6['clinic']="NYES"
        week3day7['clinic']="NYES"

        week3=pd.DataFrame(columns=week3day1.columns)
        week3=pd.concat([week3,week3day1,week3day2,week3day3,week3day4,week3day5,week3day6,week3day7])
        week3.to_csv('week3.csv',index=False)

        clinictype=df.iloc[75:95, 0:1]
        d1 = pd.DataFrame(clinictype, columns = ['type'])
        d2 = pd.DataFrame(clinictype, columns = ['type'])
        d3 = pd.DataFrame(clinictype, columns = ['type'])
        d4 = pd.DataFrame(clinictype, columns = ['type'])
        d5 = pd.DataFrame(clinictype, columns = ['type'])
        d6 = pd.DataFrame(clinictype, columns = ['type'])
        d7 = pd.DataFrame(clinictype, columns = ['type'])

        d1['type']=clinictype
        d2['type']=clinictype
        d3['type']=clinictype
        d4['type']=clinictype
        d5['type']=clinictype
        d6['type']=clinictype
        d7['type']=clinictype

        week4day1=d1.replace(to_replace=r'- Continuity', value='', regex=True)
        week4day2=d2.replace(to_replace=r'- Continuity', value='', regex=True)
        week4day3=d3.replace(to_replace=r'- Continuity', value='', regex=True)
        week4day4=d4.replace(to_replace=r'- Continuity', value='', regex=True)
        week4day5=d5.replace(to_replace=r'- Continuity', value='', regex=True)
        week4day6=d6.replace(to_replace=r'- Continuity', value='', regex=True)
        week4day7=d7.replace(to_replace=r'- Continuity', value='', regex=True)


        day1=df.iloc[73,1]
        day2=df.iloc[73,2]
        day3=df.iloc[73,3]
        day4=df.iloc[73,4]
        day5=df.iloc[73,5]
        day6=df.iloc[73,6]
        day7=df.iloc[73,7]

        week4day1['date']=day1
        week4day2['date']=day2
        week4day3['date']=day3
        week4day4['date']=day4
        week4day5['date']=day5
        week4day6['date']=day6
        week4day7['date']=day7

        provider1=df.iloc[75:95,1]
        provider2=df.iloc[75:95,2]
        provider3=df.iloc[75:95,3]
        provider4=df.iloc[75:95,4]
        provider5=df.iloc[75:95,5]
        provider6=df.iloc[75:95,6]
        provider7=df.iloc[75:95,7]

        week4day1['provider']=provider1
        week4day2['provider']=provider2
        week4day3['provider']=provider3
        week4day4['provider']=provider4
        week4day5['provider']=provider5
        week4day6['provider']=provider6
        week4day7['provider']=provider7

        week4day1['clinic']="NYES"
        week4day2['clinic']="NYES"
        week4day3['clinic']="NYES"
        week4day4['clinic']="NYES"
        week4day5['clinic']="NYES"
        week4day6['clinic']="NYES"
        week4day7['clinic']="NYES"

        week4=pd.DataFrame(columns=week4day1.columns)
        week4=pd.concat([week4,week4day1,week4day2,week4day3,week4day4,week4day5,week4day6,week4day7])
        week4.to_csv('week4.csv',index=False)

        hope=pd.DataFrame(columns=week1.columns)
        hope=pd.concat([hope,week1,week2,week3,week4])
        hope.to_csv('nyes.csv',index=False)
	    
        # Handle AM Continuity for NYE (First set)
        hope['H'] = "H"
        NYEi = hope[hope['type'] == 'AM'].copy()  # Ensure we're working with a copy
        NYEi.loc[:, 'count'] = NYEi.groupby(['date'])['provider'].cumcount() + 0  # Starts at H0 for AM
        NYEi.loc[:, 'class'] = "H" + NYEi['count'].astype(str)
        NYEi = NYEi.loc[:, ('date', 'type', 'provider', 'clinic', 'class')]
        NYEi.to_csv('1.csv', index=False)

        #dfx1 = pd.read_csv('1.csv')
        #df = dfx1
        #import io
        #output = io.StringIO()
        #df.to_csv(output, index=False)
        #output.seek(0)

        # Streamlit download button
        #st.download_button(
        #    label="Download CSV File",
        #    data=output.getvalue(),
        #    file_name="1.csv",
        #    mime="text/csv"
        #)
        # Handle PM Continuity for NYE (Second set)
        hope['H'] = "H"
        NYEii = hope[hope['type'] == 'PM'].copy()  # Ensure we're working with a copy
        NYEii.loc[:, 'count'] = NYEii.groupby(['date'])['provider'].cumcount() + 10  # Starts at H10 for PM
        NYEii.loc[:, 'class'] = "H" + NYEii['count'].astype(str)
        NYEii = NYEii.loc[:, ('date', 'type', 'provider', 'clinic', 'class')]
        NYEii.to_csv('2.csv', index=False)

        # Combine AM and PM DataFrames
        nyess = pd.DataFrame(columns=NYEi.columns)
        nyess = pd.concat([NYEi, NYEii])
        nyess.to_csv('nyess.csv', index=False)
        

        ##############################ETOWN##############################################################################################
        import pandas as pd
        read_file = pd.read_excel (uploaded_opd_file, sheet_name='ETOWN')
        read_file.to_csv ('etownroad.csv', index = False, header=False)
        df=pd.read_csv('etownroad.csv')

        clinictype=df.iloc[3:23, 0:1]
        a1 = pd.DataFrame(clinictype, columns = ['type'])
        a2 = pd.DataFrame(clinictype, columns = ['type'])
        a3 = pd.DataFrame(clinictype, columns = ['type'])
        a4 = pd.DataFrame(clinictype, columns = ['type'])
        a5 = pd.DataFrame(clinictype, columns = ['type'])
        a6 = pd.DataFrame(clinictype, columns = ['type'])
        a7 = pd.DataFrame(clinictype, columns = ['type'])

        a1['type']=clinictype
        a2['type']=clinictype
        a3['type']=clinictype
        a4['type']=clinictype
        a5['type']=clinictype
        a6['type']=clinictype
        a7['type']=clinictype

        week1day1=a1.replace(to_replace=r'- Continuity', value='', regex=True)
        week1day2=a2.replace(to_replace=r'- Continuity', value='', regex=True)
        week1day3=a3.replace(to_replace=r'- Continuity', value='', regex=True)
        week1day4=a4.replace(to_replace=r'- Continuity', value='', regex=True)
        week1day5=a5.replace(to_replace=r'- Continuity', value='', regex=True)
        week1day6=a6.replace(to_replace=r'- Continuity', value='', regex=True)
        week1day7=a7.replace(to_replace=r'- Continuity', value='', regex=True)

        day1=df.iloc[1,1]
        day2=df.iloc[1,2]
        day3=df.iloc[1,3]
        day4=df.iloc[1,4]
        day5=df.iloc[1,5]
        day6=df.iloc[1,6]
        day7=df.iloc[1,7]

        week1day1['date']=day1
        week1day2['date']=day2
        week1day3['date']=day3
        week1day4['date']=day4
        week1day5['date']=day5
        week1day6['date']=day6
        week1day7['date']=day7

        provider1=df.iloc[3:23,1]
        provider2=df.iloc[3:23,2]
        provider3=df.iloc[3:23,3]
        provider4=df.iloc[3:23,4]
        provider5=df.iloc[3:23,5]
        provider6=df.iloc[3:23,6]
        provider7=df.iloc[3:23,7]

        week1day1['provider']=provider1
        week1day2['provider']=provider2
        week1day3['provider']=provider3
        week1day4['provider']=provider4
        week1day5['provider']=provider5
        week1day6['provider']=provider6
        week1day7['provider']=provider7

        week1day1['clinic']="ETOWN"
        week1day2['clinic']="ETOWN"
        week1day3['clinic']="ETOWN"
        week1day4['clinic']="ETOWN"
        week1day5['clinic']="ETOWN"
        week1day6['clinic']="ETOWN"
        week1day7['clinic']="ETOWN"

        week1=pd.DataFrame(columns=week1day1.columns)
        week1=pd.concat([week1,week1day1,week1day2,week1day3,week1day4,week1day5,week1day6,week1day7])
        week1.to_csv('week1.csv',index=False)

        clinictype=df.iloc[27:47, 0:1]
        b1 = pd.DataFrame(clinictype, columns = ['type'])
        b2 = pd.DataFrame(clinictype, columns = ['type'])
        b3 = pd.DataFrame(clinictype, columns = ['type'])
        b4 = pd.DataFrame(clinictype, columns = ['type'])
        b5 = pd.DataFrame(clinictype, columns = ['type'])
        b6 = pd.DataFrame(clinictype, columns = ['type'])
        b7 = pd.DataFrame(clinictype, columns = ['type'])

        b1['type']=clinictype
        b2['type']=clinictype
        b3['type']=clinictype
        b4['type']=clinictype
        b5['type']=clinictype
        b6['type']=clinictype
        b7['type']=clinictype

        week2day1=b1.replace(to_replace=r'- Continuity', value='', regex=True)
        week2day2=b2.replace(to_replace=r'- Continuity', value='', regex=True)
        week2day3=b3.replace(to_replace=r'- Continuity', value='', regex=True)
        week2day4=b4.replace(to_replace=r'- Continuity', value='', regex=True)
        week2day5=b5.replace(to_replace=r'- Continuity', value='', regex=True)
        week2day6=b6.replace(to_replace=r'- Continuity', value='', regex=True)
        week2day7=b7.replace(to_replace=r'- Continuity', value='', regex=True)

        day1=df.iloc[25,1]
        day2=df.iloc[25,2]
        day3=df.iloc[25,3]
        day4=df.iloc[25,4]
        day5=df.iloc[25,5]
        day6=df.iloc[25,6]
        day7=df.iloc[25,7]

        week2day1['date']=day1
        week2day2['date']=day2
        week2day3['date']=day3
        week2day4['date']=day4
        week2day5['date']=day5
        week2day6['date']=day6
        week2day7['date']=day7

        provider1=df.iloc[27:47,1]
        provider2=df.iloc[27:47,2]
        provider3=df.iloc[27:47,3]
        provider4=df.iloc[27:47,4]
        provider5=df.iloc[27:47,5]
        provider6=df.iloc[27:47,6]
        provider7=df.iloc[27:47,7]

        week2day1['provider']=provider1
        week2day2['provider']=provider2
        week2day3['provider']=provider3
        week2day4['provider']=provider4
        week2day5['provider']=provider5
        week2day6['provider']=provider6
        week2day7['provider']=provider7

        week2day1['clinic']="ETOWN"
        week2day2['clinic']="ETOWN"
        week2day3['clinic']="ETOWN"
        week2day4['clinic']="ETOWN"
        week2day5['clinic']="ETOWN"
        week2day6['clinic']="ETOWN"
        week2day7['clinic']="ETOWN"

        week2=pd.DataFrame(columns=week2day1.columns)
        week2=pd.concat([week2,week2day1,week2day2,week2day3,week2day4,week2day5,week2day6,week2day7])
        week2.to_csv('week2.csv',index=False)

        clinictype=df.iloc[51:71, 0:1]
        c1 = pd.DataFrame(clinictype, columns = ['type'])
        c2 = pd.DataFrame(clinictype, columns = ['type'])
        c3 = pd.DataFrame(clinictype, columns = ['type'])
        c4 = pd.DataFrame(clinictype, columns = ['type'])
        c5 = pd.DataFrame(clinictype, columns = ['type'])
        c6 = pd.DataFrame(clinictype, columns = ['type'])
        c7 = pd.DataFrame(clinictype, columns = ['type'])

        c1['type']=clinictype
        c2['type']=clinictype
        c3['type']=clinictype
        c4['type']=clinictype
        c5['type']=clinictype
        c6['type']=clinictype
        c7['type']=clinictype

        week3day1=c1.replace(to_replace=r'- Continuity', value='', regex=True)
        week3day2=c2.replace(to_replace=r'- Continuity', value='', regex=True)
        week3day3=c3.replace(to_replace=r'- Continuity', value='', regex=True)
        week3day4=c4.replace(to_replace=r'- Continuity', value='', regex=True)
        week3day5=c5.replace(to_replace=r'- Continuity', value='', regex=True)
        week3day6=c6.replace(to_replace=r'- Continuity', value='', regex=True)
        week3day7=c7.replace(to_replace=r'- Continuity', value='', regex=True)

        day1=df.iloc[49,1]
        day2=df.iloc[49,2]
        day3=df.iloc[49,3]
        day4=df.iloc[49,4]
        day5=df.iloc[49,5]
        day6=df.iloc[49,6]
        day7=df.iloc[49,7]

        week3day1['date']=day1
        week3day2['date']=day2
        week3day3['date']=day3
        week3day4['date']=day4
        week3day5['date']=day5
        week3day6['date']=day6
        week3day7['date']=day7

        provider1=df.iloc[51:71,1]
        provider2=df.iloc[51:71,2]
        provider3=df.iloc[51:71,3]
        provider4=df.iloc[51:71,4]
        provider5=df.iloc[51:71,5]
        provider6=df.iloc[51:71,6]
        provider7=df.iloc[51:71,7]

        week3day1['provider']=provider1
        week3day2['provider']=provider2
        week3day3['provider']=provider3
        week3day4['provider']=provider4
        week3day5['provider']=provider5
        week3day6['provider']=provider6
        week3day7['provider']=provider7

        week3day1['clinic']="ETOWN"
        week3day2['clinic']="ETOWN"
        week3day3['clinic']="ETOWN"
        week3day4['clinic']="ETOWN"
        week3day5['clinic']="ETOWN"
        week3day6['clinic']="ETOWN"
        week3day7['clinic']="ETOWN"

        week3=pd.DataFrame(columns=week3day1.columns)
        week3=pd.concat([week3,week3day1,week3day2,week3day3,week3day4,week3day5,week3day6,week3day7])
        week3.to_csv('week3.csv',index=False)

        clinictype=df.iloc[75:95, 0:1]
        d1 = pd.DataFrame(clinictype, columns = ['type'])
        d2 = pd.DataFrame(clinictype, columns = ['type'])
        d3 = pd.DataFrame(clinictype, columns = ['type'])
        d4 = pd.DataFrame(clinictype, columns = ['type'])
        d5 = pd.DataFrame(clinictype, columns = ['type'])
        d6 = pd.DataFrame(clinictype, columns = ['type'])
        d7 = pd.DataFrame(clinictype, columns = ['type'])

        d1['type']=clinictype
        d2['type']=clinictype
        d3['type']=clinictype
        d4['type']=clinictype
        d5['type']=clinictype
        d6['type']=clinictype
        d7['type']=clinictype

        week4day1=d1.replace(to_replace=r'- Continuity', value='', regex=True)
        week4day2=d2.replace(to_replace=r'- Continuity', value='', regex=True)
        week4day3=d3.replace(to_replace=r'- Continuity', value='', regex=True)
        week4day4=d4.replace(to_replace=r'- Continuity', value='', regex=True)
        week4day5=d5.replace(to_replace=r'- Continuity', value='', regex=True)
        week4day6=d6.replace(to_replace=r'- Continuity', value='', regex=True)
        week4day7=d7.replace(to_replace=r'- Continuity', value='', regex=True)


        day1=df.iloc[73,1]
        day2=df.iloc[73,2]
        day3=df.iloc[73,3]
        day4=df.iloc[73,4]
        day5=df.iloc[73,5]
        day6=df.iloc[73,6]
        day7=df.iloc[73,7]

        week4day1['date']=day1
        week4day2['date']=day2
        week4day3['date']=day3
        week4day4['date']=day4
        week4day5['date']=day5
        week4day6['date']=day6
        week4day7['date']=day7

        provider1=df.iloc[75:95,1]
        provider2=df.iloc[75:95,2]
        provider3=df.iloc[75:95,3]
        provider4=df.iloc[75:95,4]
        provider5=df.iloc[75:95,5]
        provider6=df.iloc[75:95,6]
        provider7=df.iloc[75:95,7]

        week4day1['provider']=provider1
        week4day2['provider']=provider2
        week4day3['provider']=provider3
        week4day4['provider']=provider4
        week4day5['provider']=provider5
        week4day6['provider']=provider6
        week4day7['provider']=provider7

        week4day1['clinic']="ETOWN"
        week4day2['clinic']="ETOWN"
        week4day3['clinic']="ETOWN"
        week4day4['clinic']="ETOWN"
        week4day5['clinic']="ETOWN"
        week4day6['clinic']="ETOWN"
        week4day7['clinic']="ETOWN"

        week4=pd.DataFrame(columns=week4day1.columns)
        week4=pd.concat([week4,week4day1,week4day2,week4day3,week4day4,week4day5,week4day6,week4day7])
        week4.to_csv('week4.csv',index=False)

        hope=pd.DataFrame(columns=week1.columns)
        hope=pd.concat([hope,week1,week2,week3,week4])
        hope.to_csv('etown.csv',index=False)

        # Handle AM Continuity for ETOWN (First set)
        hope['H'] = "H"
        ETOWNi = hope[hope['type'] == 'AM'].copy()  # Ensure we're working with a copy
        ETOWNi.loc[:, 'count'] = ETOWNi.groupby(['date'])['provider'].cumcount() + 0  # Starts at H0 for AM
        ETOWNi.loc[:, 'class'] = "H" + ETOWNi['count'].astype(str)
        ETOWNi = ETOWNi.loc[:, ('date', 'type', 'provider', 'clinic', 'class')]
        ETOWNi.to_csv('3.csv', index=False)

        # Handle PM Continuity for ETOWN (Second set)
        hope['H'] = "H"
        ETOWNii = hope[hope['type'] == 'PM'].copy()  # Ensure we're working with a copy
        ETOWNii.loc[:, 'count'] = ETOWNii.groupby(['date'])['provider'].cumcount() + 10  # Starts at H10 for PM
        ETOWNii.loc[:, 'class'] = "H" + ETOWNii['count'].astype(str)
        ETOWNii = ETOWNii.loc[:, ('date', 'type', 'provider', 'clinic', 'class')]
        ETOWNii.to_csv('4.csv', index=False)

        # Combine AM and PM DataFrames for ETOWN
        etowns = pd.DataFrame(columns=ETOWNi.columns)
        etowns = pd.concat([ETOWNi, ETOWNii])
        etowns.to_csv('etowns.csv', index=False)

	##############################WARD A##############################################################################################
        import pandas as pd
        read_file = pd.read_excel (uploaded_opd_file, sheet_name='W_A')
        read_file.to_csv ('extra.csv', index = False, header=False)
        df=pd.read_csv('extra.csv')

        clinictype=df.iloc[3:23, 0:1]
        a1 = pd.DataFrame(clinictype, columns = ['type'])
        a2 = pd.DataFrame(clinictype, columns = ['type'])
        a3 = pd.DataFrame(clinictype, columns = ['type'])
        a4 = pd.DataFrame(clinictype, columns = ['type'])
        a5 = pd.DataFrame(clinictype, columns = ['type'])
        a6 = pd.DataFrame(clinictype, columns = ['type'])
        a7 = pd.DataFrame(clinictype, columns = ['type'])

        a1['type']=clinictype
        a2['type']=clinictype
        a3['type']=clinictype
        a4['type']=clinictype
        a5['type']=clinictype
        a6['type']=clinictype
        a7['type']=clinictype

        week1day1=a1.replace(to_replace=r'- Continuity', value='', regex=True)
        week1day2=a2.replace(to_replace=r'- Continuity', value='', regex=True)
        week1day3=a3.replace(to_replace=r'- Continuity', value='', regex=True)
        week1day4=a4.replace(to_replace=r'- Continuity', value='', regex=True)
        week1day5=a5.replace(to_replace=r'- Continuity', value='', regex=True)
        week1day6=a6.replace(to_replace=r'- Continuity', value='', regex=True)
        week1day7=a7.replace(to_replace=r'- Continuity', value='', regex=True)

        day1=df.iloc[1,1]
        day2=df.iloc[1,2]
        day3=df.iloc[1,3]
        day4=df.iloc[1,4]
        day5=df.iloc[1,5]
        day6=df.iloc[1,6]
        day7=df.iloc[1,7]

        week1day1['date']=day1
        week1day2['date']=day2
        week1day3['date']=day3
        week1day4['date']=day4
        week1day5['date']=day5
        week1day6['date']=day6
        week1day7['date']=day7

        provider1=df.iloc[3:23,1]
        provider2=df.iloc[3:23,2]
        provider3=df.iloc[3:23,3]
        provider4=df.iloc[3:23,4]
        provider5=df.iloc[3:23,5]
        provider6=df.iloc[3:23,6]
        provider7=df.iloc[3:23,7]

        week1day1['provider']=provider1
        week1day2['provider']=provider2
        week1day3['provider']=provider3
        week1day4['provider']=provider4
        week1day5['provider']=provider5
        week1day6['provider']=provider6
        week1day7['provider']=provider7

        week1day1['clinic']="WARD_A"
        week1day2['clinic']="WARD_A"
        week1day3['clinic']="WARD_A"
        week1day4['clinic']="WARD_A"
        week1day5['clinic']="WARD_A"
        week1day6['clinic']="WARD_A"
        week1day7['clinic']="WARD_A"

        week1=pd.DataFrame(columns=week1day1.columns)
        week1=pd.concat([week1,week1day1,week1day2,week1day3,week1day4,week1day5,week1day6,week1day7])
        week1.to_csv('week1.csv',index=False)

        clinictype=df.iloc[27:47, 0:1]
        b1 = pd.DataFrame(clinictype, columns = ['type'])
        b2 = pd.DataFrame(clinictype, columns = ['type'])
        b3 = pd.DataFrame(clinictype, columns = ['type'])
        b4 = pd.DataFrame(clinictype, columns = ['type'])
        b5 = pd.DataFrame(clinictype, columns = ['type'])
        b6 = pd.DataFrame(clinictype, columns = ['type'])
        b7 = pd.DataFrame(clinictype, columns = ['type'])

        b1['type']=clinictype
        b2['type']=clinictype
        b3['type']=clinictype
        b4['type']=clinictype
        b5['type']=clinictype
        b6['type']=clinictype
        b7['type']=clinictype

        week2day1=b1.replace(to_replace=r'- Continuity', value='', regex=True)
        week2day2=b2.replace(to_replace=r'- Continuity', value='', regex=True)
        week2day3=b3.replace(to_replace=r'- Continuity', value='', regex=True)
        week2day4=b4.replace(to_replace=r'- Continuity', value='', regex=True)
        week2day5=b5.replace(to_replace=r'- Continuity', value='', regex=True)
        week2day6=b6.replace(to_replace=r'- Continuity', value='', regex=True)
        week2day7=b7.replace(to_replace=r'- Continuity', value='', regex=True)

        day1=df.iloc[25,1]
        day2=df.iloc[25,2]
        day3=df.iloc[25,3]
        day4=df.iloc[25,4]
        day5=df.iloc[25,5]
        day6=df.iloc[25,6]
        day7=df.iloc[25,7]

        week2day1['date']=day1
        week2day2['date']=day2
        week2day3['date']=day3
        week2day4['date']=day4
        week2day5['date']=day5
        week2day6['date']=day6
        week2day7['date']=day7

        provider1=df.iloc[27:47,1]
        provider2=df.iloc[27:47,2]
        provider3=df.iloc[27:47,3]
        provider4=df.iloc[27:47,4]
        provider5=df.iloc[27:47,5]
        provider6=df.iloc[27:47,6]
        provider7=df.iloc[27:47,7]

        week2day1['provider']=provider1
        week2day2['provider']=provider2
        week2day3['provider']=provider3
        week2day4['provider']=provider4
        week2day5['provider']=provider5
        week2day6['provider']=provider6
        week2day7['provider']=provider7

        week2day1['clinic']="WARD_A"
        week2day2['clinic']="WARD_A"
        week2day3['clinic']="WARD_A"
        week2day4['clinic']="WARD_A"
        week2day5['clinic']="WARD_A"
        week2day6['clinic']="WARD_A"
        week2day7['clinic']="WARD_A"

        week2=pd.DataFrame(columns=week2day1.columns)
        week2=pd.concat([week2,week2day1,week2day2,week2day3,week2day4,week2day5,week2day6,week2day7])
        week2.to_csv('week2.csv',index=False)

        clinictype=df.iloc[51:71, 0:1]
        c1 = pd.DataFrame(clinictype, columns = ['type'])
        c2 = pd.DataFrame(clinictype, columns = ['type'])
        c3 = pd.DataFrame(clinictype, columns = ['type'])
        c4 = pd.DataFrame(clinictype, columns = ['type'])
        c5 = pd.DataFrame(clinictype, columns = ['type'])
        c6 = pd.DataFrame(clinictype, columns = ['type'])
        c7 = pd.DataFrame(clinictype, columns = ['type'])

        c1['type']=clinictype
        c2['type']=clinictype
        c3['type']=clinictype
        c4['type']=clinictype
        c5['type']=clinictype
        c6['type']=clinictype
        c7['type']=clinictype

        week3day1=c1.replace(to_replace=r'- Continuity', value='', regex=True)
        week3day2=c2.replace(to_replace=r'- Continuity', value='', regex=True)
        week3day3=c3.replace(to_replace=r'- Continuity', value='', regex=True)
        week3day4=c4.replace(to_replace=r'- Continuity', value='', regex=True)
        week3day5=c5.replace(to_replace=r'- Continuity', value='', regex=True)
        week3day6=c6.replace(to_replace=r'- Continuity', value='', regex=True)
        week3day7=c7.replace(to_replace=r'- Continuity', value='', regex=True)

        day1=df.iloc[49,1]
        day2=df.iloc[49,2]
        day3=df.iloc[49,3]
        day4=df.iloc[49,4]
        day5=df.iloc[49,5]
        day6=df.iloc[49,6]
        day7=df.iloc[49,7]

        week3day1['date']=day1
        week3day2['date']=day2
        week3day3['date']=day3
        week3day4['date']=day4
        week3day5['date']=day5
        week3day6['date']=day6
        week3day7['date']=day7

        provider1=df.iloc[51:71,1]
        provider2=df.iloc[51:71,2]
        provider3=df.iloc[51:71,3]
        provider4=df.iloc[51:71,4]
        provider5=df.iloc[51:71,5]
        provider6=df.iloc[51:71,6]
        provider7=df.iloc[51:71,7]

        week3day1['provider']=provider1
        week3day2['provider']=provider2
        week3day3['provider']=provider3
        week3day4['provider']=provider4
        week3day5['provider']=provider5
        week3day6['provider']=provider6
        week3day7['provider']=provider7

        week3day1['clinic']="WARD_A"
        week3day2['clinic']="WARD_A"
        week3day3['clinic']="WARD_A"
        week3day4['clinic']="WARD_A"
        week3day5['clinic']="WARD_A"
        week3day6['clinic']="WARD_A"
        week3day7['clinic']="WARD_A"

        week3=pd.DataFrame(columns=week3day1.columns)
        week3=pd.concat([week3,week3day1,week3day2,week3day3,week3day4,week3day5,week3day6,week3day7])
        week3.to_csv('week3.csv',index=False)

        clinictype=df.iloc[75:95, 0:1]
        d1 = pd.DataFrame(clinictype, columns = ['type'])
        d2 = pd.DataFrame(clinictype, columns = ['type'])
        d3 = pd.DataFrame(clinictype, columns = ['type'])
        d4 = pd.DataFrame(clinictype, columns = ['type'])
        d5 = pd.DataFrame(clinictype, columns = ['type'])
        d6 = pd.DataFrame(clinictype, columns = ['type'])
        d7 = pd.DataFrame(clinictype, columns = ['type'])

        d1['type']=clinictype
        d2['type']=clinictype
        d3['type']=clinictype
        d4['type']=clinictype
        d5['type']=clinictype
        d6['type']=clinictype
        d7['type']=clinictype

        week4day1=d1.replace(to_replace=r'- Continuity', value='', regex=True)
        week4day2=d2.replace(to_replace=r'- Continuity', value='', regex=True)
        week4day3=d3.replace(to_replace=r'- Continuity', value='', regex=True)
        week4day4=d4.replace(to_replace=r'- Continuity', value='', regex=True)
        week4day5=d5.replace(to_replace=r'- Continuity', value='', regex=True)
        week4day6=d6.replace(to_replace=r'- Continuity', value='', regex=True)
        week4day7=d7.replace(to_replace=r'- Continuity', value='', regex=True)


        day1=df.iloc[73,1]
        day2=df.iloc[73,2]
        day3=df.iloc[73,3]
        day4=df.iloc[73,4]
        day5=df.iloc[73,5]
        day6=df.iloc[73,6]
        day7=df.iloc[73,7]

        week4day1['date']=day1
        week4day2['date']=day2
        week4day3['date']=day3
        week4day4['date']=day4
        week4day5['date']=day5
        week4day6['date']=day6
        week4day7['date']=day7

        provider1=df.iloc[75:95,1]
        provider2=df.iloc[75:95,2]
        provider3=df.iloc[75:95,3]
        provider4=df.iloc[75:95,4]
        provider5=df.iloc[75:95,5]
        provider6=df.iloc[75:95,6]
        provider7=df.iloc[75:95,7]

        week4day1['provider']=provider1
        week4day2['provider']=provider2
        week4day3['provider']=provider3
        week4day4['provider']=provider4
        week4day5['provider']=provider5
        week4day6['provider']=provider6
        week4day7['provider']=provider7

        week4day1['clinic']="WARD_A"
        week4day2['clinic']="WARD_A"
        week4day3['clinic']="WARD_A"
        week4day4['clinic']="WARD_A"
        week4day5['clinic']="WARD_A"
        week4day6['clinic']="WARD_A"
        week4day7['clinic']="WARD_A"

        week4=pd.DataFrame(columns=week4day1.columns)
        week4=pd.concat([week4,week4day1,week4day2,week4day3,week4day4,week4day5,week4day6,week4day7])
        week4.to_csv('week4.csv',index=False)

        hope=pd.DataFrame(columns=week1.columns)
        hope=pd.concat([hope,week1,week2,week3,week4])
        hope.to_csv('extra.csv',index=False)

        hope['H'] = "H"
        extrai = hope[hope['type'] == 'AM'].copy()  # Make sure we're working with a copy
        extrai.loc[:, 'count'] = extrai.groupby(['date'])['provider'].cumcount() + 0  # Starts at H0 for AM
        extrai.loc[:, 'class'] = "H" + extrai['count'].astype(str)
        extrai = extrai.loc[:, ('date', 'type', 'provider', 'clinic', 'class')]
        extrai.to_csv('3.csv', index=False)

        # Handle PM Continuity for EXTRAI
        hope['H'] = "H"
        extraii = hope[hope['type'] == 'PM'].copy()  # Make sure we're working with a copy
        extraii.loc[:, 'count'] = extraii.groupby(['date'])['provider'].cumcount() + 10  # Starts at H10 for PM
        extraii.loc[:, 'class'] = "H" + extraii['count'].astype(str)
        extraii = extraii.loc[:, ('date', 'type', 'provider', 'clinic', 'class')]
        extraii.to_csv('4.csv', index=False)

        # Combine AM and PM DataFrames for EXTRAI
        extras = pd.DataFrame(columns=extrai.columns)
        extras = pd.concat([extrai, extraii])
        extras.to_csv('extra_a.csv', index=False)

        ##############################MHS##############################################################################################
        import pandas as pd
        read_file = pd.read_excel (uploaded_opd_file, sheet_name='AAC')
        read_file.to_csv ('mhss.csv', index = False, header=False)
        df=pd.read_csv('mhss.csv')

        clinictype=df.iloc[3:23, 0:1]
        a1 = pd.DataFrame(clinictype, columns = ['type'])
        a2 = pd.DataFrame(clinictype, columns = ['type'])
        a3 = pd.DataFrame(clinictype, columns = ['type'])
        a4 = pd.DataFrame(clinictype, columns = ['type'])
        a5 = pd.DataFrame(clinictype, columns = ['type'])
        a6 = pd.DataFrame(clinictype, columns = ['type'])
        a7 = pd.DataFrame(clinictype, columns = ['type'])

        a1['type']=clinictype
        a2['type']=clinictype
        a3['type']=clinictype
        a4['type']=clinictype
        a5['type']=clinictype
        a6['type']=clinictype
        a7['type']=clinictype

        week1day1=a1.replace(to_replace=r'- Continuity', value='', regex=True)
        week1day2=a2.replace(to_replace=r'- Continuity', value='', regex=True)
        week1day3=a3.replace(to_replace=r'- Continuity', value='', regex=True)
        week1day4=a4.replace(to_replace=r'- Continuity', value='', regex=True)
        week1day5=a5.replace(to_replace=r'- Continuity', value='', regex=True)
        week1day6=a6.replace(to_replace=r'- Continuity', value='', regex=True)
        week1day7=a7.replace(to_replace=r'- Continuity', value='', regex=True)

        day1=df.iloc[1,1]
        day2=df.iloc[1,2]
        day3=df.iloc[1,3]
        day4=df.iloc[1,4]
        day5=df.iloc[1,5]
        day6=df.iloc[1,6]
        day7=df.iloc[1,7]

        week1day1['date']=day1
        week1day2['date']=day2
        week1day3['date']=day3
        week1day4['date']=day4
        week1day5['date']=day5
        week1day6['date']=day6
        week1day7['date']=day7

        provider1=df.iloc[3:23,1]
        provider2=df.iloc[3:23,2]
        provider3=df.iloc[3:23,3]
        provider4=df.iloc[3:23,4]
        provider5=df.iloc[3:23,5]
        provider6=df.iloc[3:23,6]
        provider7=df.iloc[3:23,7]

        week1day1['provider']=provider1
        week1day2['provider']=provider2
        week1day3['provider']=provider3
        week1day4['provider']=provider4
        week1day5['provider']=provider5
        week1day6['provider']=provider6
        week1day7['provider']=provider7

        week1day1['clinic']="AAC"
        week1day2['clinic']="AAC"
        week1day3['clinic']="AAC"
        week1day4['clinic']="AAC"
        week1day5['clinic']="AAC"
        week1day6['clinic']="AAC"
        week1day7['clinic']="AAC"

        week1=pd.DataFrame(columns=week1day1.columns)
        week1=pd.concat([week1,week1day1,week1day2,week1day3,week1day4,week1day5,week1day6,week1day7])
        week1.to_csv('week1.csv',index=False)

        clinictype=df.iloc[27:47, 0:1]
        b1 = pd.DataFrame(clinictype, columns = ['type'])
        b2 = pd.DataFrame(clinictype, columns = ['type'])
        b3 = pd.DataFrame(clinictype, columns = ['type'])
        b4 = pd.DataFrame(clinictype, columns = ['type'])
        b5 = pd.DataFrame(clinictype, columns = ['type'])
        b6 = pd.DataFrame(clinictype, columns = ['type'])
        b7 = pd.DataFrame(clinictype, columns = ['type'])

        b1['type']=clinictype
        b2['type']=clinictype
        b3['type']=clinictype
        b4['type']=clinictype
        b5['type']=clinictype
        b6['type']=clinictype
        b7['type']=clinictype

        week2day1=b1.replace(to_replace=r'- Continuity', value='', regex=True)
        week2day2=b2.replace(to_replace=r'- Continuity', value='', regex=True)
        week2day3=b3.replace(to_replace=r'- Continuity', value='', regex=True)
        week2day4=b4.replace(to_replace=r'- Continuity', value='', regex=True)
        week2day5=b5.replace(to_replace=r'- Continuity', value='', regex=True)
        week2day6=b6.replace(to_replace=r'- Continuity', value='', regex=True)
        week2day7=b7.replace(to_replace=r'- Continuity', value='', regex=True)

        day1=df.iloc[25,1]
        day2=df.iloc[25,2]
        day3=df.iloc[25,3]
        day4=df.iloc[25,4]
        day5=df.iloc[25,5]
        day6=df.iloc[25,6]
        day7=df.iloc[25,7]

        week2day1['date']=day1
        week2day2['date']=day2
        week2day3['date']=day3
        week2day4['date']=day4
        week2day5['date']=day5
        week2day6['date']=day6
        week2day7['date']=day7

        provider1=df.iloc[27:47,1]
        provider2=df.iloc[27:47,2]
        provider3=df.iloc[27:47,3]
        provider4=df.iloc[27:47,4]
        provider5=df.iloc[27:47,5]
        provider6=df.iloc[27:47,6]
        provider7=df.iloc[27:47,7]

        week2day1['provider']=provider1
        week2day2['provider']=provider2
        week2day3['provider']=provider3
        week2day4['provider']=provider4
        week2day5['provider']=provider5
        week2day6['provider']=provider6
        week2day7['provider']=provider7

        week2day1['clinic']="AAC"
        week2day2['clinic']="AAC"
        week2day3['clinic']="AAC"
        week2day4['clinic']="AAC"
        week2day5['clinic']="AAC"
        week2day6['clinic']="AAC"
        week2day7['clinic']="AAC"

        week2=pd.DataFrame(columns=week2day1.columns)
        week2=pd.concat([week2,week2day1,week2day2,week2day3,week2day4,week2day5,week2day6,week2day7])
        week2.to_csv('week2.csv',index=False)

        clinictype=df.iloc[51:71, 0:1]
        c1 = pd.DataFrame(clinictype, columns = ['type'])
        c2 = pd.DataFrame(clinictype, columns = ['type'])
        c3 = pd.DataFrame(clinictype, columns = ['type'])
        c4 = pd.DataFrame(clinictype, columns = ['type'])
        c5 = pd.DataFrame(clinictype, columns = ['type'])
        c6 = pd.DataFrame(clinictype, columns = ['type'])
        c7 = pd.DataFrame(clinictype, columns = ['type'])

        c1['type']=clinictype
        c2['type']=clinictype
        c3['type']=clinictype
        c4['type']=clinictype
        c5['type']=clinictype
        c6['type']=clinictype
        c7['type']=clinictype

        week3day1=c1.replace(to_replace=r'- Continuity', value='', regex=True)
        week3day2=c2.replace(to_replace=r'- Continuity', value='', regex=True)
        week3day3=c3.replace(to_replace=r'- Continuity', value='', regex=True)
        week3day4=c4.replace(to_replace=r'- Continuity', value='', regex=True)
        week3day5=c5.replace(to_replace=r'- Continuity', value='', regex=True)
        week3day6=c6.replace(to_replace=r'- Continuity', value='', regex=True)
        week3day7=c7.replace(to_replace=r'- Continuity', value='', regex=True)

        day1=df.iloc[49,1]
        day2=df.iloc[49,2]
        day3=df.iloc[49,3]
        day4=df.iloc[49,4]
        day5=df.iloc[49,5]
        day6=df.iloc[49,6]
        day7=df.iloc[49,7]

        week3day1['date']=day1
        week3day2['date']=day2
        week3day3['date']=day3
        week3day4['date']=day4
        week3day5['date']=day5
        week3day6['date']=day6
        week3day7['date']=day7

        provider1=df.iloc[51:71,1]
        provider2=df.iloc[51:71,2]
        provider3=df.iloc[51:71,3]
        provider4=df.iloc[51:71,4]
        provider5=df.iloc[51:71,5]
        provider6=df.iloc[51:71,6]
        provider7=df.iloc[51:71,7]

        week3day1['provider']=provider1
        week3day2['provider']=provider2
        week3day3['provider']=provider3
        week3day4['provider']=provider4
        week3day5['provider']=provider5
        week3day6['provider']=provider6
        week3day7['provider']=provider7

        week3day1['clinic']="AAC"
        week3day2['clinic']="AAC"
        week3day3['clinic']="AAC"
        week3day4['clinic']="AAC"
        week3day5['clinic']="AAC"
        week3day6['clinic']="AAC"
        week3day7['clinic']="AAC"

        week3=pd.DataFrame(columns=week3day1.columns)
        week3=pd.concat([week3,week3day1,week3day2,week3day3,week3day4,week3day5,week3day6,week3day7])
        week3.to_csv('week3.csv',index=False)

        clinictype=df.iloc[75:95, 0:1]
        d1 = pd.DataFrame(clinictype, columns = ['type'])
        d2 = pd.DataFrame(clinictype, columns = ['type'])
        d3 = pd.DataFrame(clinictype, columns = ['type'])
        d4 = pd.DataFrame(clinictype, columns = ['type'])
        d5 = pd.DataFrame(clinictype, columns = ['type'])
        d6 = pd.DataFrame(clinictype, columns = ['type'])
        d7 = pd.DataFrame(clinictype, columns = ['type'])

        d1['type']=clinictype
        d2['type']=clinictype
        d3['type']=clinictype
        d4['type']=clinictype
        d5['type']=clinictype
        d6['type']=clinictype
        d7['type']=clinictype

        week4day1=d1.replace(to_replace=r'- Continuity', value='', regex=True)
        week4day2=d2.replace(to_replace=r'- Continuity', value='', regex=True)
        week4day3=d3.replace(to_replace=r'- Continuity', value='', regex=True)
        week4day4=d4.replace(to_replace=r'- Continuity', value='', regex=True)
        week4day5=d5.replace(to_replace=r'- Continuity', value='', regex=True)
        week4day6=d6.replace(to_replace=r'- Continuity', value='', regex=True)
        week4day7=d7.replace(to_replace=r'- Continuity', value='', regex=True)


        day1=df.iloc[73,1]
        day2=df.iloc[73,2]
        day3=df.iloc[73,3]
        day4=df.iloc[73,4]
        day5=df.iloc[73,5]
        day6=df.iloc[73,6]
        day7=df.iloc[73,7]

        week4day1['date']=day1
        week4day2['date']=day2
        week4day3['date']=day3
        week4day4['date']=day4
        week4day5['date']=day5
        week4day6['date']=day6
        week4day7['date']=day7

        provider1=df.iloc[75:95,1]
        provider2=df.iloc[75:95,2]
        provider3=df.iloc[75:95,3]
        provider4=df.iloc[75:95,4]
        provider5=df.iloc[75:95,5]
        provider6=df.iloc[75:95,6]
        provider7=df.iloc[75:95,7]

        week4day1['provider']=provider1
        week4day2['provider']=provider2
        week4day3['provider']=provider3
        week4day4['provider']=provider4
        week4day5['provider']=provider5
        week4day6['provider']=provider6
        week4day7['provider']=provider7

        week4day1['clinic']="AAC"
        week4day2['clinic']="AAC"
        week4day3['clinic']="AAC"
        week4day4['clinic']="AAC"
        week4day5['clinic']="AAC"
        week4day6['clinic']="AAC"
        week4day7['clinic']="AAC"

        week4=pd.DataFrame(columns=week4day1.columns)
        week4=pd.concat([week4,week4day1,week4day2,week4day3,week4day4,week4day5,week4day6,week4day7])
        week4.to_csv('week4.csv',index=False)

        hope=pd.DataFrame(columns=week1.columns)
        hope=pd.concat([hope,week1,week2,week3,week4])
        hope.to_csv('mhss.csv',index=False)

        hope['H'] = "H"
        extrai = hope[hope['type'] == 'AM'].copy()  # Ensure we're working with a copy
        extrai.loc[:, 'count'] = extrai.groupby(['date'])['provider'].cumcount() + 0  # Starts at H0 for AM
        extrai.loc[:, 'class'] = "H" + extrai['count'].astype(str)
        extrai = extrai.loc[:, ('date', 'type', 'provider', 'clinic', 'class')]
        extrai.to_csv('3.csv', index=False)

        # Handle PM Continuity for EXTRAI
        hope['H'] = "H"
        extraii = hope[hope['type'] == 'PM'].copy()  # Ensure we're working with a copy
        extraii.loc[:, 'count'] = extraii.groupby(['date'])['provider'].cumcount() + 10  # Starts at H10 for PM
        extraii.loc[:, 'class'] = "H" + extraii['count'].astype(str)
        extraii = extraii.loc[:, ('date', 'type', 'provider', 'clinic', 'class')]
        extraii.to_csv('4.csv', index=False)

        extras=pd.DataFrame(columns=extrai.columns)
        extras=pd.concat([extrai,extraii])
        extras.to_csv('mhss.csv',index=False)
	##############################WARD C##############################################################################################
        import pandas as pd
        read_file = pd.read_excel (uploaded_opd_file, sheet_name='W_C')
        read_file.to_csv ('extra.csv', index = False, header=False)
        df=pd.read_csv('extra.csv')

        clinictype=df.iloc[3:23, 0:1]
        a1 = pd.DataFrame(clinictype, columns = ['type'])
        a2 = pd.DataFrame(clinictype, columns = ['type'])
        a3 = pd.DataFrame(clinictype, columns = ['type'])
        a4 = pd.DataFrame(clinictype, columns = ['type'])
        a5 = pd.DataFrame(clinictype, columns = ['type'])
        a6 = pd.DataFrame(clinictype, columns = ['type'])
        a7 = pd.DataFrame(clinictype, columns = ['type'])

        a1['type']=clinictype
        a2['type']=clinictype
        a3['type']=clinictype
        a4['type']=clinictype
        a5['type']=clinictype
        a6['type']=clinictype
        a7['type']=clinictype

        week1day1=a1.replace(to_replace=r'- Continuity', value='', regex=True)
        week1day2=a2.replace(to_replace=r'- Continuity', value='', regex=True)
        week1day3=a3.replace(to_replace=r'- Continuity', value='', regex=True)
        week1day4=a4.replace(to_replace=r'- Continuity', value='', regex=True)
        week1day5=a5.replace(to_replace=r'- Continuity', value='', regex=True)
        week1day6=a6.replace(to_replace=r'- Continuity', value='', regex=True)
        week1day7=a7.replace(to_replace=r'- Continuity', value='', regex=True)

        day1=df.iloc[1,1]
        day2=df.iloc[1,2]
        day3=df.iloc[1,3]
        day4=df.iloc[1,4]
        day5=df.iloc[1,5]
        day6=df.iloc[1,6]
        day7=df.iloc[1,7]

        week1day1['date']=day1
        week1day2['date']=day2
        week1day3['date']=day3
        week1day4['date']=day4
        week1day5['date']=day5
        week1day6['date']=day6
        week1day7['date']=day7

        provider1=df.iloc[3:23,1]
        provider2=df.iloc[3:23,2]
        provider3=df.iloc[3:23,3]
        provider4=df.iloc[3:23,4]
        provider5=df.iloc[3:23,5]
        provider6=df.iloc[3:23,6]
        provider7=df.iloc[3:23,7]

        week1day1['provider']=provider1
        week1day2['provider']=provider2
        week1day3['provider']=provider3
        week1day4['provider']=provider4
        week1day5['provider']=provider5
        week1day6['provider']=provider6
        week1day7['provider']=provider7

        week1day1['clinic']="WARD_C"
        week1day2['clinic']="WARD_C"
        week1day3['clinic']="WARD_C"
        week1day4['clinic']="WARD_C"
        week1day5['clinic']="WARD_C"
        week1day6['clinic']="WARD_C"
        week1day7['clinic']="WARD_C"

        week1=pd.DataFrame(columns=week1day1.columns)
        week1=pd.concat([week1,week1day1,week1day2,week1day3,week1day4,week1day5,week1day6,week1day7])
        week1.to_csv('week1.csv',index=False)

        clinictype=df.iloc[27:47, 0:1]
        b1 = pd.DataFrame(clinictype, columns = ['type'])
        b2 = pd.DataFrame(clinictype, columns = ['type'])
        b3 = pd.DataFrame(clinictype, columns = ['type'])
        b4 = pd.DataFrame(clinictype, columns = ['type'])
        b5 = pd.DataFrame(clinictype, columns = ['type'])
        b6 = pd.DataFrame(clinictype, columns = ['type'])
        b7 = pd.DataFrame(clinictype, columns = ['type'])

        b1['type']=clinictype
        b2['type']=clinictype
        b3['type']=clinictype
        b4['type']=clinictype
        b5['type']=clinictype
        b6['type']=clinictype
        b7['type']=clinictype

        week2day1=b1.replace(to_replace=r'- Continuity', value='', regex=True)
        week2day2=b2.replace(to_replace=r'- Continuity', value='', regex=True)
        week2day3=b3.replace(to_replace=r'- Continuity', value='', regex=True)
        week2day4=b4.replace(to_replace=r'- Continuity', value='', regex=True)
        week2day5=b5.replace(to_replace=r'- Continuity', value='', regex=True)
        week2day6=b6.replace(to_replace=r'- Continuity', value='', regex=True)
        week2day7=b7.replace(to_replace=r'- Continuity', value='', regex=True)

        day1=df.iloc[25,1]
        day2=df.iloc[25,2]
        day3=df.iloc[25,3]
        day4=df.iloc[25,4]
        day5=df.iloc[25,5]
        day6=df.iloc[25,6]
        day7=df.iloc[25,7]

        week2day1['date']=day1
        week2day2['date']=day2
        week2day3['date']=day3
        week2day4['date']=day4
        week2day5['date']=day5
        week2day6['date']=day6
        week2day7['date']=day7

        provider1=df.iloc[27:47,1]
        provider2=df.iloc[27:47,2]
        provider3=df.iloc[27:47,3]
        provider4=df.iloc[27:47,4]
        provider5=df.iloc[27:47,5]
        provider6=df.iloc[27:47,6]
        provider7=df.iloc[27:47,7]

        week2day1['provider']=provider1
        week2day2['provider']=provider2
        week2day3['provider']=provider3
        week2day4['provider']=provider4
        week2day5['provider']=provider5
        week2day6['provider']=provider6
        week2day7['provider']=provider7

        week2day1['clinic']="WARD_C"
        week2day2['clinic']="WARD_C"
        week2day3['clinic']="WARD_C"
        week2day4['clinic']="WARD_C"
        week2day5['clinic']="WARD_C"
        week2day6['clinic']="WARD_C"
        week2day7['clinic']="WARD_C"

        week2=pd.DataFrame(columns=week2day1.columns)
        week2=pd.concat([week2,week2day1,week2day2,week2day3,week2day4,week2day5,week2day6,week2day7])
        week2.to_csv('week2.csv',index=False)

        clinictype=df.iloc[51:71, 0:1]
        c1 = pd.DataFrame(clinictype, columns = ['type'])
        c2 = pd.DataFrame(clinictype, columns = ['type'])
        c3 = pd.DataFrame(clinictype, columns = ['type'])
        c4 = pd.DataFrame(clinictype, columns = ['type'])
        c5 = pd.DataFrame(clinictype, columns = ['type'])
        c6 = pd.DataFrame(clinictype, columns = ['type'])
        c7 = pd.DataFrame(clinictype, columns = ['type'])

        c1['type']=clinictype
        c2['type']=clinictype
        c3['type']=clinictype
        c4['type']=clinictype
        c5['type']=clinictype
        c6['type']=clinictype
        c7['type']=clinictype

        week3day1=c1.replace(to_replace=r'- Continuity', value='', regex=True)
        week3day2=c2.replace(to_replace=r'- Continuity', value='', regex=True)
        week3day3=c3.replace(to_replace=r'- Continuity', value='', regex=True)
        week3day4=c4.replace(to_replace=r'- Continuity', value='', regex=True)
        week3day5=c5.replace(to_replace=r'- Continuity', value='', regex=True)
        week3day6=c6.replace(to_replace=r'- Continuity', value='', regex=True)
        week3day7=c7.replace(to_replace=r'- Continuity', value='', regex=True)

        day1=df.iloc[49,1]
        day2=df.iloc[49,2]
        day3=df.iloc[49,3]
        day4=df.iloc[49,4]
        day5=df.iloc[49,5]
        day6=df.iloc[49,6]
        day7=df.iloc[49,7]

        week3day1['date']=day1
        week3day2['date']=day2
        week3day3['date']=day3
        week3day4['date']=day4
        week3day5['date']=day5
        week3day6['date']=day6
        week3day7['date']=day7

        provider1=df.iloc[51:71,1]
        provider2=df.iloc[51:71,2]
        provider3=df.iloc[51:71,3]
        provider4=df.iloc[51:71,4]
        provider5=df.iloc[51:71,5]
        provider6=df.iloc[51:71,6]
        provider7=df.iloc[51:71,7]

        week3day1['provider']=provider1
        week3day2['provider']=provider2
        week3day3['provider']=provider3
        week3day4['provider']=provider4
        week3day5['provider']=provider5
        week3day6['provider']=provider6
        week3day7['provider']=provider7

        week3day1['clinic']="WARD_C"
        week3day2['clinic']="WARD_C"
        week3day3['clinic']="WARD_C"
        week3day4['clinic']="WARD_C"
        week3day5['clinic']="WARD_C"
        week3day6['clinic']="WARD_C"
        week3day7['clinic']="WARD_C"

        week3=pd.DataFrame(columns=week3day1.columns)
        week3=pd.concat([week3,week3day1,week3day2,week3day3,week3day4,week3day5,week3day6,week3day7])
        week3.to_csv('week3.csv',index=False)

        clinictype=df.iloc[75:95, 0:1]
        d1 = pd.DataFrame(clinictype, columns = ['type'])
        d2 = pd.DataFrame(clinictype, columns = ['type'])
        d3 = pd.DataFrame(clinictype, columns = ['type'])
        d4 = pd.DataFrame(clinictype, columns = ['type'])
        d5 = pd.DataFrame(clinictype, columns = ['type'])
        d6 = pd.DataFrame(clinictype, columns = ['type'])
        d7 = pd.DataFrame(clinictype, columns = ['type'])

        d1['type']=clinictype
        d2['type']=clinictype
        d3['type']=clinictype
        d4['type']=clinictype
        d5['type']=clinictype
        d6['type']=clinictype
        d7['type']=clinictype

        week4day1=d1.replace(to_replace=r'- Continuity', value='', regex=True)
        week4day2=d2.replace(to_replace=r'- Continuity', value='', regex=True)
        week4day3=d3.replace(to_replace=r'- Continuity', value='', regex=True)
        week4day4=d4.replace(to_replace=r'- Continuity', value='', regex=True)
        week4day5=d5.replace(to_replace=r'- Continuity', value='', regex=True)
        week4day6=d6.replace(to_replace=r'- Continuity', value='', regex=True)
        week4day7=d7.replace(to_replace=r'- Continuity', value='', regex=True)


        day1=df.iloc[73,1]
        day2=df.iloc[73,2]
        day3=df.iloc[73,3]
        day4=df.iloc[73,4]
        day5=df.iloc[73,5]
        day6=df.iloc[73,6]
        day7=df.iloc[73,7]

        week4day1['date']=day1
        week4day2['date']=day2
        week4day3['date']=day3
        week4day4['date']=day4
        week4day5['date']=day5
        week4day6['date']=day6
        week4day7['date']=day7

        provider1=df.iloc[75:95,1]
        provider2=df.iloc[75:95,2]
        provider3=df.iloc[75:95,3]
        provider4=df.iloc[75:95,4]
        provider5=df.iloc[75:95,5]
        provider6=df.iloc[75:95,6]
        provider7=df.iloc[75:95,7]

        week4day1['provider']=provider1
        week4day2['provider']=provider2
        week4day3['provider']=provider3
        week4day4['provider']=provider4
        week4day5['provider']=provider5
        week4day6['provider']=provider6
        week4day7['provider']=provider7

        week4day1['clinic']="WARD_C"
        week4day2['clinic']="WARD_C"
        week4day3['clinic']="WARD_C"
        week4day4['clinic']="WARD_C"
        week4day5['clinic']="WARD_C"
        week4day6['clinic']="WARD_C"
        week4day7['clinic']="WARD_C"

        week4=pd.DataFrame(columns=week4day1.columns)
        week4=pd.concat([week4,week4day1,week4day2,week4day3,week4day4,week4day5,week4day6,week4day7])
        week4.to_csv('week4.csv',index=False)

        hope=pd.DataFrame(columns=week1.columns)
        hope=pd.concat([hope,week1,week2,week3,week4])
        hope.to_csv('extra.csv',index=False)

        hope['H'] = "H"
        extrai = hope[hope['type'] == 'AM'].copy()  # Make sure we're working with a copy
        extrai.loc[:, 'count'] = extrai.groupby(['date'])['provider'].cumcount() + 0  # Starts at H0 for AM
        extrai.loc[:, 'class'] = "H" + extrai['count'].astype(str)
        extrai = extrai.loc[:, ('date', 'type', 'provider', 'clinic', 'class')]
        extrai.to_csv('3.csv', index=False)

        # Handle PM Continuity for EXTRAI
        hope['H'] = "H"
        extraii = hope[hope['type'] == 'PM'].copy()  # Make sure we're working with a copy
        extraii.loc[:, 'count'] = extraii.groupby(['date'])['provider'].cumcount() + 10  # Starts at H10 for PM
        extraii.loc[:, 'class'] = "H" + extraii['count'].astype(str)
        extraii = extraii.loc[:, ('date', 'type', 'provider', 'clinic', 'class')]
        extraii.to_csv('4.csv', index=False)

        # Combine AM and PM DataFrames for EXTRAI
        extras = pd.DataFrame(columns=extrai.columns)
        extras = pd.concat([extrai, extraii])
        extras.to_csv('extra_c.csv', index=False)
	##############################WARD P##############################################################################################
        import pandas as pd
        read_file = pd.read_excel (uploaded_opd_file, sheet_name='W_P')
        read_file.to_csv ('extra.csv', index = False, header=False)
        df=pd.read_csv('extra.csv')

        clinictype=df.iloc[3:23, 0:1]
        a1 = pd.DataFrame(clinictype, columns = ['type'])
        a2 = pd.DataFrame(clinictype, columns = ['type'])
        a3 = pd.DataFrame(clinictype, columns = ['type'])
        a4 = pd.DataFrame(clinictype, columns = ['type'])
        a5 = pd.DataFrame(clinictype, columns = ['type'])
        a6 = pd.DataFrame(clinictype, columns = ['type'])
        a7 = pd.DataFrame(clinictype, columns = ['type'])

        a1['type']=clinictype
        a2['type']=clinictype
        a3['type']=clinictype
        a4['type']=clinictype
        a5['type']=clinictype
        a6['type']=clinictype
        a7['type']=clinictype

        week1day1=a1.replace(to_replace=r'- Continuity', value='', regex=True)
        week1day2=a2.replace(to_replace=r'- Continuity', value='', regex=True)
        week1day3=a3.replace(to_replace=r'- Continuity', value='', regex=True)
        week1day4=a4.replace(to_replace=r'- Continuity', value='', regex=True)
        week1day5=a5.replace(to_replace=r'- Continuity', value='', regex=True)
        week1day6=a6.replace(to_replace=r'- Continuity', value='', regex=True)
        week1day7=a7.replace(to_replace=r'- Continuity', value='', regex=True)

        day1=df.iloc[1,1]
        day2=df.iloc[1,2]
        day3=df.iloc[1,3]
        day4=df.iloc[1,4]
        day5=df.iloc[1,5]
        day6=df.iloc[1,6]
        day7=df.iloc[1,7]

        week1day1['date']=day1
        week1day2['date']=day2
        week1day3['date']=day3
        week1day4['date']=day4
        week1day5['date']=day5
        week1day6['date']=day6
        week1day7['date']=day7

        provider1=df.iloc[3:23,1]
        provider2=df.iloc[3:23,2]
        provider3=df.iloc[3:23,3]
        provider4=df.iloc[3:23,4]
        provider5=df.iloc[3:23,5]
        provider6=df.iloc[3:23,6]
        provider7=df.iloc[3:23,7]

        week1day1['provider']=provider1
        week1day2['provider']=provider2
        week1day3['provider']=provider3
        week1day4['provider']=provider4
        week1day5['provider']=provider5
        week1day6['provider']=provider6
        week1day7['provider']=provider7

        week1day1['clinic']="WARD_P"
        week1day2['clinic']="WARD_P"
        week1day3['clinic']="WARD_P"
        week1day4['clinic']="WARD_P"
        week1day5['clinic']="WARD_P"
        week1day6['clinic']="WARD_P"
        week1day7['clinic']="WARD_P"

        week1=pd.DataFrame(columns=week1day1.columns)
        week1=pd.concat([week1,week1day1,week1day2,week1day3,week1day4,week1day5,week1day6,week1day7])
        week1.to_csv('week1.csv',index=False)

        clinictype=df.iloc[27:47, 0:1]
        b1 = pd.DataFrame(clinictype, columns = ['type'])
        b2 = pd.DataFrame(clinictype, columns = ['type'])
        b3 = pd.DataFrame(clinictype, columns = ['type'])
        b4 = pd.DataFrame(clinictype, columns = ['type'])
        b5 = pd.DataFrame(clinictype, columns = ['type'])
        b6 = pd.DataFrame(clinictype, columns = ['type'])
        b7 = pd.DataFrame(clinictype, columns = ['type'])

        b1['type']=clinictype
        b2['type']=clinictype
        b3['type']=clinictype
        b4['type']=clinictype
        b5['type']=clinictype
        b6['type']=clinictype
        b7['type']=clinictype

        week2day1=b1.replace(to_replace=r'- Continuity', value='', regex=True)
        week2day2=b2.replace(to_replace=r'- Continuity', value='', regex=True)
        week2day3=b3.replace(to_replace=r'- Continuity', value='', regex=True)
        week2day4=b4.replace(to_replace=r'- Continuity', value='', regex=True)
        week2day5=b5.replace(to_replace=r'- Continuity', value='', regex=True)
        week2day6=b6.replace(to_replace=r'- Continuity', value='', regex=True)
        week2day7=b7.replace(to_replace=r'- Continuity', value='', regex=True)

        day1=df.iloc[25,1]
        day2=df.iloc[25,2]
        day3=df.iloc[25,3]
        day4=df.iloc[25,4]
        day5=df.iloc[25,5]
        day6=df.iloc[25,6]
        day7=df.iloc[25,7]

        week2day1['date']=day1
        week2day2['date']=day2
        week2day3['date']=day3
        week2day4['date']=day4
        week2day5['date']=day5
        week2day6['date']=day6
        week2day7['date']=day7

        provider1=df.iloc[27:47,1]
        provider2=df.iloc[27:47,2]
        provider3=df.iloc[27:47,3]
        provider4=df.iloc[27:47,4]
        provider5=df.iloc[27:47,5]
        provider6=df.iloc[27:47,6]
        provider7=df.iloc[27:47,7]

        week2day1['provider']=provider1
        week2day2['provider']=provider2
        week2day3['provider']=provider3
        week2day4['provider']=provider4
        week2day5['provider']=provider5
        week2day6['provider']=provider6
        week2day7['provider']=provider7

        week2day1['clinic']="WARD_P"
        week2day2['clinic']="WARD_P"
        week2day3['clinic']="WARD_P"
        week2day4['clinic']="WARD_P"
        week2day5['clinic']="WARD_P"
        week2day6['clinic']="WARD_P"
        week2day7['clinic']="WARD_P"

        week2=pd.DataFrame(columns=week2day1.columns)
        week2=pd.concat([week2,week2day1,week2day2,week2day3,week2day4,week2day5,week2day6,week2day7])
        week2.to_csv('week2.csv',index=False)

        clinictype=df.iloc[51:71, 0:1]
        c1 = pd.DataFrame(clinictype, columns = ['type'])
        c2 = pd.DataFrame(clinictype, columns = ['type'])
        c3 = pd.DataFrame(clinictype, columns = ['type'])
        c4 = pd.DataFrame(clinictype, columns = ['type'])
        c5 = pd.DataFrame(clinictype, columns = ['type'])
        c6 = pd.DataFrame(clinictype, columns = ['type'])
        c7 = pd.DataFrame(clinictype, columns = ['type'])

        c1['type']=clinictype
        c2['type']=clinictype
        c3['type']=clinictype
        c4['type']=clinictype
        c5['type']=clinictype
        c6['type']=clinictype
        c7['type']=clinictype

        week3day1=c1.replace(to_replace=r'- Continuity', value='', regex=True)
        week3day2=c2.replace(to_replace=r'- Continuity', value='', regex=True)
        week3day3=c3.replace(to_replace=r'- Continuity', value='', regex=True)
        week3day4=c4.replace(to_replace=r'- Continuity', value='', regex=True)
        week3day5=c5.replace(to_replace=r'- Continuity', value='', regex=True)
        week3day6=c6.replace(to_replace=r'- Continuity', value='', regex=True)
        week3day7=c7.replace(to_replace=r'- Continuity', value='', regex=True)

        day1=df.iloc[49,1]
        day2=df.iloc[49,2]
        day3=df.iloc[49,3]
        day4=df.iloc[49,4]
        day5=df.iloc[49,5]
        day6=df.iloc[49,6]
        day7=df.iloc[49,7]

        week3day1['date']=day1
        week3day2['date']=day2
        week3day3['date']=day3
        week3day4['date']=day4
        week3day5['date']=day5
        week3day6['date']=day6
        week3day7['date']=day7

        provider1=df.iloc[51:71,1]
        provider2=df.iloc[51:71,2]
        provider3=df.iloc[51:71,3]
        provider4=df.iloc[51:71,4]
        provider5=df.iloc[51:71,5]
        provider6=df.iloc[51:71,6]
        provider7=df.iloc[51:71,7]

        week3day1['provider']=provider1
        week3day2['provider']=provider2
        week3day3['provider']=provider3
        week3day4['provider']=provider4
        week3day5['provider']=provider5
        week3day6['provider']=provider6
        week3day7['provider']=provider7

        week3day1['clinic']="WARD_P"
        week3day2['clinic']="WARD_P"
        week3day3['clinic']="WARD_P"
        week3day4['clinic']="WARD_P"
        week3day5['clinic']="WARD_P"
        week3day6['clinic']="WARD_P"
        week3day7['clinic']="WARD_P"

        week3=pd.DataFrame(columns=week3day1.columns)
        week3=pd.concat([week3,week3day1,week3day2,week3day3,week3day4,week3day5,week3day6,week3day7])
        week3.to_csv('week3.csv',index=False)

        clinictype=df.iloc[75:95, 0:1]
        d1 = pd.DataFrame(clinictype, columns = ['type'])
        d2 = pd.DataFrame(clinictype, columns = ['type'])
        d3 = pd.DataFrame(clinictype, columns = ['type'])
        d4 = pd.DataFrame(clinictype, columns = ['type'])
        d5 = pd.DataFrame(clinictype, columns = ['type'])
        d6 = pd.DataFrame(clinictype, columns = ['type'])
        d7 = pd.DataFrame(clinictype, columns = ['type'])

        d1['type']=clinictype
        d2['type']=clinictype
        d3['type']=clinictype
        d4['type']=clinictype
        d5['type']=clinictype
        d6['type']=clinictype
        d7['type']=clinictype

        week4day1=d1.replace(to_replace=r'- Continuity', value='', regex=True)
        week4day2=d2.replace(to_replace=r'- Continuity', value='', regex=True)
        week4day3=d3.replace(to_replace=r'- Continuity', value='', regex=True)
        week4day4=d4.replace(to_replace=r'- Continuity', value='', regex=True)
        week4day5=d5.replace(to_replace=r'- Continuity', value='', regex=True)
        week4day6=d6.replace(to_replace=r'- Continuity', value='', regex=True)
        week4day7=d7.replace(to_replace=r'- Continuity', value='', regex=True)


        day1=df.iloc[73,1]
        day2=df.iloc[73,2]
        day3=df.iloc[73,3]
        day4=df.iloc[73,4]
        day5=df.iloc[73,5]
        day6=df.iloc[73,6]
        day7=df.iloc[73,7]

        week4day1['date']=day1
        week4day2['date']=day2
        week4day3['date']=day3
        week4day4['date']=day4
        week4day5['date']=day5
        week4day6['date']=day6
        week4day7['date']=day7

        provider1=df.iloc[75:95,1]
        provider2=df.iloc[75:95,2]
        provider3=df.iloc[75:95,3]
        provider4=df.iloc[75:95,4]
        provider5=df.iloc[75:95,5]
        provider6=df.iloc[75:95,6]
        provider7=df.iloc[75:95,7]

        week4day1['provider']=provider1
        week4day2['provider']=provider2
        week4day3['provider']=provider3
        week4day4['provider']=provider4
        week4day5['provider']=provider5
        week4day6['provider']=provider6
        week4day7['provider']=provider7

        week4day1['clinic']="WARD_P"
        week4day2['clinic']="WARD_P"
        week4day3['clinic']="WARD_P"
        week4day4['clinic']="WARD_P"
        week4day5['clinic']="WARD_P"
        week4day6['clinic']="WARD_P"
        week4day7['clinic']="WARD_P"

        week4=pd.DataFrame(columns=week4day1.columns)
        week4=pd.concat([week4,week4day1,week4day2,week4day3,week4day4,week4day5,week4day6,week4day7])
        week4.to_csv('week4.csv',index=False)

        hope=pd.DataFrame(columns=week1.columns)
        hope=pd.concat([hope,week1,week2,week3,week4])
        hope.to_csv('extra.csv',index=False)

        hope['H'] = "H"
        extrai = hope[hope['type'] == 'AM'].copy()  # Make sure we're working with a copy
        extrai.loc[:, 'count'] = extrai.groupby(['date'])['provider'].cumcount() + 0  # Starts at H0 for AM
        extrai.loc[:, 'class'] = "H" + extrai['count'].astype(str)
        extrai = extrai.loc[:, ('date', 'type', 'provider', 'clinic', 'class')]
        extrai.to_csv('3.csv', index=False)

        # Handle PM Continuity for EXTRAI
        hope['H'] = "H"
        extraii = hope[hope['type'] == 'PM'].copy()  # Make sure we're working with a copy
        extraii.loc[:, 'count'] = extraii.groupby(['date'])['provider'].cumcount() + 10  # Starts at H10 for PM
        extraii.loc[:, 'class'] = "H" + extraii['count'].astype(str)
        extraii = extraii.loc[:, ('date', 'type', 'provider', 'clinic', 'class')]
        extraii.to_csv('4.csv', index=False)

        # Combine AM and PM DataFrames for EXTRAI
        extras = pd.DataFrame(columns=extrai.columns)
        extras = pd.concat([extrai, extraii])
        extras.to_csv('extra_p.csv', index=False)
	#############################PICU##############################################################################################
        import pandas as pd
        read_file = pd.read_excel (uploaded_opd_file, sheet_name='PICU')
        read_file.to_csv ('extra.csv', index = False, header=False)
        df=pd.read_csv('extra.csv')

        clinictype=df.iloc[3:23, 0:1]
        a1 = pd.DataFrame(clinictype, columns = ['type'])
        a2 = pd.DataFrame(clinictype, columns = ['type'])
        a3 = pd.DataFrame(clinictype, columns = ['type'])
        a4 = pd.DataFrame(clinictype, columns = ['type'])
        a5 = pd.DataFrame(clinictype, columns = ['type'])
        a6 = pd.DataFrame(clinictype, columns = ['type'])
        a7 = pd.DataFrame(clinictype, columns = ['type'])

        a1['type']=clinictype
        a2['type']=clinictype
        a3['type']=clinictype
        a4['type']=clinictype
        a5['type']=clinictype
        a6['type']=clinictype
        a7['type']=clinictype

        week1day1=a1.replace(to_replace=r'- Continuity', value='', regex=True)
        week1day2=a2.replace(to_replace=r'- Continuity', value='', regex=True)
        week1day3=a3.replace(to_replace=r'- Continuity', value='', regex=True)
        week1day4=a4.replace(to_replace=r'- Continuity', value='', regex=True)
        week1day5=a5.replace(to_replace=r'- Continuity', value='', regex=True)
        week1day6=a6.replace(to_replace=r'- Continuity', value='', regex=True)
        week1day7=a7.replace(to_replace=r'- Continuity', value='', regex=True)

        day1=df.iloc[1,1]
        day2=df.iloc[1,2]
        day3=df.iloc[1,3]
        day4=df.iloc[1,4]
        day5=df.iloc[1,5]
        day6=df.iloc[1,6]
        day7=df.iloc[1,7]

        week1day1['date']=day1
        week1day2['date']=day2
        week1day3['date']=day3
        week1day4['date']=day4
        week1day5['date']=day5
        week1day6['date']=day6
        week1day7['date']=day7

        provider1=df.iloc[3:23,1]
        provider2=df.iloc[3:23,2]
        provider3=df.iloc[3:23,3]
        provider4=df.iloc[3:23,4]
        provider5=df.iloc[3:23,5]
        provider6=df.iloc[3:23,6]
        provider7=df.iloc[3:23,7]

        week1day1['provider']=provider1
        week1day2['provider']=provider2
        week1day3['provider']=provider3
        week1day4['provider']=provider4
        week1day5['provider']=provider5
        week1day6['provider']=provider6
        week1day7['provider']=provider7

        week1day1['clinic']="PICU"
        week1day2['clinic']="PICU"
        week1day3['clinic']="PICU"
        week1day4['clinic']="PICU"
        week1day5['clinic']="PICU"
        week1day6['clinic']="PICU"
        week1day7['clinic']="PICU"

        week1=pd.DataFrame(columns=week1day1.columns)
        week1=pd.concat([week1,week1day1,week1day2,week1day3,week1day4,week1day5,week1day6,week1day7])
        week1.to_csv('week1.csv',index=False)

        clinictype=df.iloc[27:47, 0:1]
        b1 = pd.DataFrame(clinictype, columns = ['type'])
        b2 = pd.DataFrame(clinictype, columns = ['type'])
        b3 = pd.DataFrame(clinictype, columns = ['type'])
        b4 = pd.DataFrame(clinictype, columns = ['type'])
        b5 = pd.DataFrame(clinictype, columns = ['type'])
        b6 = pd.DataFrame(clinictype, columns = ['type'])
        b7 = pd.DataFrame(clinictype, columns = ['type'])

        b1['type']=clinictype
        b2['type']=clinictype
        b3['type']=clinictype
        b4['type']=clinictype
        b5['type']=clinictype
        b6['type']=clinictype
        b7['type']=clinictype

        week2day1=b1.replace(to_replace=r'- Continuity', value='', regex=True)
        week2day2=b2.replace(to_replace=r'- Continuity', value='', regex=True)
        week2day3=b3.replace(to_replace=r'- Continuity', value='', regex=True)
        week2day4=b4.replace(to_replace=r'- Continuity', value='', regex=True)
        week2day5=b5.replace(to_replace=r'- Continuity', value='', regex=True)
        week2day6=b6.replace(to_replace=r'- Continuity', value='', regex=True)
        week2day7=b7.replace(to_replace=r'- Continuity', value='', regex=True)

        day1=df.iloc[25,1]
        day2=df.iloc[25,2]
        day3=df.iloc[25,3]
        day4=df.iloc[25,4]
        day5=df.iloc[25,5]
        day6=df.iloc[25,6]
        day7=df.iloc[25,7]

        week2day1['date']=day1
        week2day2['date']=day2
        week2day3['date']=day3
        week2day4['date']=day4
        week2day5['date']=day5
        week2day6['date']=day6
        week2day7['date']=day7

        provider1=df.iloc[27:47,1]
        provider2=df.iloc[27:47,2]
        provider3=df.iloc[27:47,3]
        provider4=df.iloc[27:47,4]
        provider5=df.iloc[27:47,5]
        provider6=df.iloc[27:47,6]
        provider7=df.iloc[27:47,7]

        week2day1['provider']=provider1
        week2day2['provider']=provider2
        week2day3['provider']=provider3
        week2day4['provider']=provider4
        week2day5['provider']=provider5
        week2day6['provider']=provider6
        week2day7['provider']=provider7

        week2day1['clinic']="PICU"
        week2day2['clinic']="PICU"
        week2day3['clinic']="PICU"
        week2day4['clinic']="PICU"
        week2day5['clinic']="PICU"
        week2day6['clinic']="PICU"
        week2day7['clinic']="PICU"

        week2=pd.DataFrame(columns=week2day1.columns)
        week2=pd.concat([week2,week2day1,week2day2,week2day3,week2day4,week2day5,week2day6,week2day7])
        week2.to_csv('week2.csv',index=False)

        clinictype=df.iloc[51:71, 0:1]
        c1 = pd.DataFrame(clinictype, columns = ['type'])
        c2 = pd.DataFrame(clinictype, columns = ['type'])
        c3 = pd.DataFrame(clinictype, columns = ['type'])
        c4 = pd.DataFrame(clinictype, columns = ['type'])
        c5 = pd.DataFrame(clinictype, columns = ['type'])
        c6 = pd.DataFrame(clinictype, columns = ['type'])
        c7 = pd.DataFrame(clinictype, columns = ['type'])

        c1['type']=clinictype
        c2['type']=clinictype
        c3['type']=clinictype
        c4['type']=clinictype
        c5['type']=clinictype
        c6['type']=clinictype
        c7['type']=clinictype

        week3day1=c1.replace(to_replace=r'- Continuity', value='', regex=True)
        week3day2=c2.replace(to_replace=r'- Continuity', value='', regex=True)
        week3day3=c3.replace(to_replace=r'- Continuity', value='', regex=True)
        week3day4=c4.replace(to_replace=r'- Continuity', value='', regex=True)
        week3day5=c5.replace(to_replace=r'- Continuity', value='', regex=True)
        week3day6=c6.replace(to_replace=r'- Continuity', value='', regex=True)
        week3day7=c7.replace(to_replace=r'- Continuity', value='', regex=True)

        day1=df.iloc[49,1]
        day2=df.iloc[49,2]
        day3=df.iloc[49,3]
        day4=df.iloc[49,4]
        day5=df.iloc[49,5]
        day6=df.iloc[49,6]
        day7=df.iloc[49,7]

        week3day1['date']=day1
        week3day2['date']=day2
        week3day3['date']=day3
        week3day4['date']=day4
        week3day5['date']=day5
        week3day6['date']=day6
        week3day7['date']=day7

        provider1=df.iloc[51:71,1]
        provider2=df.iloc[51:71,2]
        provider3=df.iloc[51:71,3]
        provider4=df.iloc[51:71,4]
        provider5=df.iloc[51:71,5]
        provider6=df.iloc[51:71,6]
        provider7=df.iloc[51:71,7]

        week3day1['provider']=provider1
        week3day2['provider']=provider2
        week3day3['provider']=provider3
        week3day4['provider']=provider4
        week3day5['provider']=provider5
        week3day6['provider']=provider6
        week3day7['provider']=provider7

        week3day1['clinic']="PICU"
        week3day2['clinic']="PICU"
        week3day3['clinic']="PICU"
        week3day4['clinic']="PICU"
        week3day5['clinic']="PICU"
        week3day6['clinic']="PICU"
        week3day7['clinic']="PICU"

        week3=pd.DataFrame(columns=week3day1.columns)
        week3=pd.concat([week3,week3day1,week3day2,week3day3,week3day4,week3day5,week3day6,week3day7])
        week3.to_csv('week3.csv',index=False)

        clinictype=df.iloc[75:95, 0:1]
        d1 = pd.DataFrame(clinictype, columns = ['type'])
        d2 = pd.DataFrame(clinictype, columns = ['type'])
        d3 = pd.DataFrame(clinictype, columns = ['type'])
        d4 = pd.DataFrame(clinictype, columns = ['type'])
        d5 = pd.DataFrame(clinictype, columns = ['type'])
        d6 = pd.DataFrame(clinictype, columns = ['type'])
        d7 = pd.DataFrame(clinictype, columns = ['type'])

        d1['type']=clinictype
        d2['type']=clinictype
        d3['type']=clinictype
        d4['type']=clinictype
        d5['type']=clinictype
        d6['type']=clinictype
        d7['type']=clinictype

        week4day1=d1.replace(to_replace=r'- Continuity', value='', regex=True)
        week4day2=d2.replace(to_replace=r'- Continuity', value='', regex=True)
        week4day3=d3.replace(to_replace=r'- Continuity', value='', regex=True)
        week4day4=d4.replace(to_replace=r'- Continuity', value='', regex=True)
        week4day5=d5.replace(to_replace=r'- Continuity', value='', regex=True)
        week4day6=d6.replace(to_replace=r'- Continuity', value='', regex=True)
        week4day7=d7.replace(to_replace=r'- Continuity', value='', regex=True)


        day1=df.iloc[73,1]
        day2=df.iloc[73,2]
        day3=df.iloc[73,3]
        day4=df.iloc[73,4]
        day5=df.iloc[73,5]
        day6=df.iloc[73,6]
        day7=df.iloc[73,7]

        week4day1['date']=day1
        week4day2['date']=day2
        week4day3['date']=day3
        week4day4['date']=day4
        week4day5['date']=day5
        week4day6['date']=day6
        week4day7['date']=day7

        provider1=df.iloc[75:95,1]
        provider2=df.iloc[75:95,2]
        provider3=df.iloc[75:95,3]
        provider4=df.iloc[75:95,4]
        provider5=df.iloc[75:95,5]
        provider6=df.iloc[75:95,6]
        provider7=df.iloc[75:95,7]

        week4day1['provider']=provider1
        week4day2['provider']=provider2
        week4day3['provider']=provider3
        week4day4['provider']=provider4
        week4day5['provider']=provider5
        week4day6['provider']=provider6
        week4day7['provider']=provider7

        week4day1['clinic']="PICU"
        week4day2['clinic']="PICU"
        week4day3['clinic']="PICU"
        week4day4['clinic']="PICU"
        week4day5['clinic']="PICU"
        week4day6['clinic']="PICU"
        week4day7['clinic']="PICU"

        week4=pd.DataFrame(columns=week4day1.columns)
        week4=pd.concat([week4,week4day1,week4day2,week4day3,week4day4,week4day5,week4day6,week4day7])
        week4.to_csv('week4.csv',index=False)

        hope=pd.DataFrame(columns=week1.columns)
        hope=pd.concat([hope,week1,week2,week3,week4])
        hope.to_csv('extra.csv',index=False)

        hope['H'] = "H"
        extrai = hope[hope['type'] == 'AM'].copy()  # Make sure we're working with a copy
        extrai.loc[:, 'count'] = extrai.groupby(['date'])['provider'].cumcount() + 0  # Starts at H0 for AM
        extrai.loc[:, 'class'] = "H" + extrai['count'].astype(str)
        extrai = extrai.loc[:, ('date', 'type', 'provider', 'clinic', 'class')]
        extrai.to_csv('3.csv', index=False)

        # Handle PM Continuity for EXTRAI
        hope['H'] = "H"
        extraii = hope[hope['type'] == 'PM'].copy()  # Make sure we're working with a copy
        extraii.loc[:, 'count'] = extraii.groupby(['date'])['provider'].cumcount() + 10  # Starts at H10 for PM
        extraii.loc[:, 'class'] = "H" + extraii['count'].astype(str)
        extraii = extraii.loc[:, ('date', 'type', 'provider', 'clinic', 'class')]
        extraii.to_csv('4.csv', index=False)

        # Combine AM and PM DataFrames for EXTRAI
        extras = pd.DataFrame(columns=extrai.columns)
        extras = pd.concat([extrai, extraii])
        extras.to_csv('extra_picu.csv', index=False)
	#############################PSHCH_NURS##############################################################################################
        import pandas as pd
        read_file = pd.read_excel (uploaded_opd_file, sheet_name='PSHCH_NURS')
        read_file.to_csv ('extra.csv', index = False, header=False)
        df=pd.read_csv('extra.csv')

        clinictype=df.iloc[3:23, 0:1]
        a1 = pd.DataFrame(clinictype, columns = ['type'])
        a2 = pd.DataFrame(clinictype, columns = ['type'])
        a3 = pd.DataFrame(clinictype, columns = ['type'])
        a4 = pd.DataFrame(clinictype, columns = ['type'])
        a5 = pd.DataFrame(clinictype, columns = ['type'])
        a6 = pd.DataFrame(clinictype, columns = ['type'])
        a7 = pd.DataFrame(clinictype, columns = ['type'])

        a1['type']=clinictype
        a2['type']=clinictype
        a3['type']=clinictype
        a4['type']=clinictype
        a5['type']=clinictype
        a6['type']=clinictype
        a7['type']=clinictype

        week1day1=a1.replace(to_replace=r'- Continuity', value='', regex=True)
        week1day2=a2.replace(to_replace=r'- Continuity', value='', regex=True)
        week1day3=a3.replace(to_replace=r'- Continuity', value='', regex=True)
        week1day4=a4.replace(to_replace=r'- Continuity', value='', regex=True)
        week1day5=a5.replace(to_replace=r'- Continuity', value='', regex=True)
        week1day6=a6.replace(to_replace=r'- Continuity', value='', regex=True)
        week1day7=a7.replace(to_replace=r'- Continuity', value='', regex=True)

        day1=df.iloc[1,1]
        day2=df.iloc[1,2]
        day3=df.iloc[1,3]
        day4=df.iloc[1,4]
        day5=df.iloc[1,5]
        day6=df.iloc[1,6]
        day7=df.iloc[1,7]

        week1day1['date']=day1
        week1day2['date']=day2
        week1day3['date']=day3
        week1day4['date']=day4
        week1day5['date']=day5
        week1day6['date']=day6
        week1day7['date']=day7

        provider1=df.iloc[3:23,1]
        provider2=df.iloc[3:23,2]
        provider3=df.iloc[3:23,3]
        provider4=df.iloc[3:23,4]
        provider5=df.iloc[3:23,5]
        provider6=df.iloc[3:23,6]
        provider7=df.iloc[3:23,7]

        week1day1['provider']=provider1
        week1day2['provider']=provider2
        week1day3['provider']=provider3
        week1day4['provider']=provider4
        week1day5['provider']=provider5
        week1day6['provider']=provider6
        week1day7['provider']=provider7

        week1day1['clinic']="PSHCH_NURS"
        week1day2['clinic']="PSHCH_NURS"
        week1day3['clinic']="PSHCH_NURS"
        week1day4['clinic']="PSHCH_NURS"
        week1day5['clinic']="PSHCH_NURS"
        week1day6['clinic']="PSHCH_NURS"
        week1day7['clinic']="PSHCH_NURS"

        week1=pd.DataFrame(columns=week1day1.columns)
        week1=pd.concat([week1,week1day1,week1day2,week1day3,week1day4,week1day5,week1day6,week1day7])
        week1.to_csv('week1.csv',index=False)

        clinictype=df.iloc[27:47, 0:1]
        b1 = pd.DataFrame(clinictype, columns = ['type'])
        b2 = pd.DataFrame(clinictype, columns = ['type'])
        b3 = pd.DataFrame(clinictype, columns = ['type'])
        b4 = pd.DataFrame(clinictype, columns = ['type'])
        b5 = pd.DataFrame(clinictype, columns = ['type'])
        b6 = pd.DataFrame(clinictype, columns = ['type'])
        b7 = pd.DataFrame(clinictype, columns = ['type'])

        b1['type']=clinictype
        b2['type']=clinictype
        b3['type']=clinictype
        b4['type']=clinictype
        b5['type']=clinictype
        b6['type']=clinictype
        b7['type']=clinictype

        week2day1=b1.replace(to_replace=r'- Continuity', value='', regex=True)
        week2day2=b2.replace(to_replace=r'- Continuity', value='', regex=True)
        week2day3=b3.replace(to_replace=r'- Continuity', value='', regex=True)
        week2day4=b4.replace(to_replace=r'- Continuity', value='', regex=True)
        week2day5=b5.replace(to_replace=r'- Continuity', value='', regex=True)
        week2day6=b6.replace(to_replace=r'- Continuity', value='', regex=True)
        week2day7=b7.replace(to_replace=r'- Continuity', value='', regex=True)

        day1=df.iloc[25,1]
        day2=df.iloc[25,2]
        day3=df.iloc[25,3]
        day4=df.iloc[25,4]
        day5=df.iloc[25,5]
        day6=df.iloc[25,6]
        day7=df.iloc[25,7]

        week2day1['date']=day1
        week2day2['date']=day2
        week2day3['date']=day3
        week2day4['date']=day4
        week2day5['date']=day5
        week2day6['date']=day6
        week2day7['date']=day7

        provider1=df.iloc[27:47,1]
        provider2=df.iloc[27:47,2]
        provider3=df.iloc[27:47,3]
        provider4=df.iloc[27:47,4]
        provider5=df.iloc[27:47,5]
        provider6=df.iloc[27:47,6]
        provider7=df.iloc[27:47,7]

        week2day1['provider']=provider1
        week2day2['provider']=provider2
        week2day3['provider']=provider3
        week2day4['provider']=provider4
        week2day5['provider']=provider5
        week2day6['provider']=provider6
        week2day7['provider']=provider7

        week2day1['clinic']="PSHCH_NURS"
        week2day2['clinic']="PSHCH_NURS"
        week2day3['clinic']="PSHCH_NURS"
        week2day4['clinic']="PSHCH_NURS"
        week2day5['clinic']="PSHCH_NURS"
        week2day6['clinic']="PSHCH_NURS"
        week2day7['clinic']="PSHCH_NURS"

        week2=pd.DataFrame(columns=week2day1.columns)
        week2=pd.concat([week2,week2day1,week2day2,week2day3,week2day4,week2day5,week2day6,week2day7])
        week2.to_csv('week2.csv',index=False)

        clinictype=df.iloc[51:71, 0:1]
        c1 = pd.DataFrame(clinictype, columns = ['type'])
        c2 = pd.DataFrame(clinictype, columns = ['type'])
        c3 = pd.DataFrame(clinictype, columns = ['type'])
        c4 = pd.DataFrame(clinictype, columns = ['type'])
        c5 = pd.DataFrame(clinictype, columns = ['type'])
        c6 = pd.DataFrame(clinictype, columns = ['type'])
        c7 = pd.DataFrame(clinictype, columns = ['type'])

        c1['type']=clinictype
        c2['type']=clinictype
        c3['type']=clinictype
        c4['type']=clinictype
        c5['type']=clinictype
        c6['type']=clinictype
        c7['type']=clinictype

        week3day1=c1.replace(to_replace=r'- Continuity', value='', regex=True)
        week3day2=c2.replace(to_replace=r'- Continuity', value='', regex=True)
        week3day3=c3.replace(to_replace=r'- Continuity', value='', regex=True)
        week3day4=c4.replace(to_replace=r'- Continuity', value='', regex=True)
        week3day5=c5.replace(to_replace=r'- Continuity', value='', regex=True)
        week3day6=c6.replace(to_replace=r'- Continuity', value='', regex=True)
        week3day7=c7.replace(to_replace=r'- Continuity', value='', regex=True)

        day1=df.iloc[49,1]
        day2=df.iloc[49,2]
        day3=df.iloc[49,3]
        day4=df.iloc[49,4]
        day5=df.iloc[49,5]
        day6=df.iloc[49,6]
        day7=df.iloc[49,7]

        week3day1['date']=day1
        week3day2['date']=day2
        week3day3['date']=day3
        week3day4['date']=day4
        week3day5['date']=day5
        week3day6['date']=day6
        week3day7['date']=day7

        provider1=df.iloc[51:71,1]
        provider2=df.iloc[51:71,2]
        provider3=df.iloc[51:71,3]
        provider4=df.iloc[51:71,4]
        provider5=df.iloc[51:71,5]
        provider6=df.iloc[51:71,6]
        provider7=df.iloc[51:71,7]

        week3day1['provider']=provider1
        week3day2['provider']=provider2
        week3day3['provider']=provider3
        week3day4['provider']=provider4
        week3day5['provider']=provider5
        week3day6['provider']=provider6
        week3day7['provider']=provider7

        week3day1['clinic']="PSHCH_NURS"
        week3day2['clinic']="PSHCH_NURS"
        week3day3['clinic']="PSHCH_NURS"
        week3day4['clinic']="PSHCH_NURS"
        week3day5['clinic']="PSHCH_NURS"
        week3day6['clinic']="PSHCH_NURS"
        week3day7['clinic']="PSHCH_NURS"

        week3=pd.DataFrame(columns=week3day1.columns)
        week3=pd.concat([week3,week3day1,week3day2,week3day3,week3day4,week3day5,week3day6,week3day7])
        week3.to_csv('week3.csv',index=False)

        clinictype=df.iloc[75:95, 0:1]
        d1 = pd.DataFrame(clinictype, columns = ['type'])
        d2 = pd.DataFrame(clinictype, columns = ['type'])
        d3 = pd.DataFrame(clinictype, columns = ['type'])
        d4 = pd.DataFrame(clinictype, columns = ['type'])
        d5 = pd.DataFrame(clinictype, columns = ['type'])
        d6 = pd.DataFrame(clinictype, columns = ['type'])
        d7 = pd.DataFrame(clinictype, columns = ['type'])

        d1['type']=clinictype
        d2['type']=clinictype
        d3['type']=clinictype
        d4['type']=clinictype
        d5['type']=clinictype
        d6['type']=clinictype
        d7['type']=clinictype

        week4day1=d1.replace(to_replace=r'- Continuity', value='', regex=True)
        week4day2=d2.replace(to_replace=r'- Continuity', value='', regex=True)
        week4day3=d3.replace(to_replace=r'- Continuity', value='', regex=True)
        week4day4=d4.replace(to_replace=r'- Continuity', value='', regex=True)
        week4day5=d5.replace(to_replace=r'- Continuity', value='', regex=True)
        week4day6=d6.replace(to_replace=r'- Continuity', value='', regex=True)
        week4day7=d7.replace(to_replace=r'- Continuity', value='', regex=True)


        day1=df.iloc[73,1]
        day2=df.iloc[73,2]
        day3=df.iloc[73,3]
        day4=df.iloc[73,4]
        day5=df.iloc[73,5]
        day6=df.iloc[73,6]
        day7=df.iloc[73,7]

        week4day1['date']=day1
        week4day2['date']=day2
        week4day3['date']=day3
        week4day4['date']=day4
        week4day5['date']=day5
        week4day6['date']=day6
        week4day7['date']=day7

        provider1=df.iloc[75:95,1]
        provider2=df.iloc[75:95,2]
        provider3=df.iloc[75:95,3]
        provider4=df.iloc[75:95,4]
        provider5=df.iloc[75:95,5]
        provider6=df.iloc[75:95,6]
        provider7=df.iloc[75:95,7]

        week4day1['provider']=provider1
        week4day2['provider']=provider2
        week4day3['provider']=provider3
        week4day4['provider']=provider4
        week4day5['provider']=provider5
        week4day6['provider']=provider6
        week4day7['provider']=provider7

        week4day1['clinic']="PSHCH_NURS"
        week4day2['clinic']="PSHCH_NURS"
        week4day3['clinic']="PSHCH_NURS"
        week4day4['clinic']="PSHCH_NURS"
        week4day5['clinic']="PSHCH_NURS"
        week4day6['clinic']="PSHCH_NURS"
        week4day7['clinic']="PSHCH_NURS"

        week4=pd.DataFrame(columns=week4day1.columns)
        week4=pd.concat([week4,week4day1,week4day2,week4day3,week4day4,week4day5,week4day6,week4day7])
        week4.to_csv('week4.csv',index=False)

        hope=pd.DataFrame(columns=week1.columns)
        hope=pd.concat([hope,week1,week2,week3,week4])
        hope.to_csv('extra.csv',index=False)

        hope['H'] = "H"
        extrai = hope[hope['type'] == 'AM'].copy()  # Make sure we're working with a copy
        extrai.loc[:, 'count'] = extrai.groupby(['date'])['provider'].cumcount() + 0  # Starts at H0 for AM
        extrai.loc[:, 'class'] = "H" + extrai['count'].astype(str)
        extrai = extrai.loc[:, ('date', 'type', 'provider', 'clinic', 'class')]
        extrai.to_csv('3.csv', index=False)

        # Handle PM Continuity for EXTRAI
        hope['H'] = "H"
        extraii = hope[hope['type'] == 'PM'].copy()  # Make sure we're working with a copy
        extraii.loc[:, 'count'] = extraii.groupby(['date'])['provider'].cumcount() + 10  # Starts at H10 for PM
        extraii.loc[:, 'class'] = "H" + extraii['count'].astype(str)
        extraii = extraii.loc[:, ('date', 'type', 'provider', 'clinic', 'class')]
        extraii.to_csv('4.csv', index=False)

        # Combine AM and PM DataFrames for EXTRAI
        extras = pd.DataFrame(columns=extrai.columns)
        extras = pd.concat([extrai, extraii])
        extras.to_csv('extra_nurs.csv', index=False)
	#############################HAMPDEN_NURS##############################################################################################
        import pandas as pd
        read_file = pd.read_excel (uploaded_opd_file, sheet_name='HAMPDEN_NURS')
        read_file.to_csv ('extra.csv', index = False, header=False)
        df=pd.read_csv('extra.csv')

        clinictype=df.iloc[3:23, 0:1]
        a1 = pd.DataFrame(clinictype, columns = ['type'])
        a2 = pd.DataFrame(clinictype, columns = ['type'])
        a3 = pd.DataFrame(clinictype, columns = ['type'])
        a4 = pd.DataFrame(clinictype, columns = ['type'])
        a5 = pd.DataFrame(clinictype, columns = ['type'])
        a6 = pd.DataFrame(clinictype, columns = ['type'])
        a7 = pd.DataFrame(clinictype, columns = ['type'])

        a1['type']=clinictype
        a2['type']=clinictype
        a3['type']=clinictype
        a4['type']=clinictype
        a5['type']=clinictype
        a6['type']=clinictype
        a7['type']=clinictype

        week1day1=a1.replace(to_replace=r'- Continuity', value='', regex=True)
        week1day2=a2.replace(to_replace=r'- Continuity', value='', regex=True)
        week1day3=a3.replace(to_replace=r'- Continuity', value='', regex=True)
        week1day4=a4.replace(to_replace=r'- Continuity', value='', regex=True)
        week1day5=a5.replace(to_replace=r'- Continuity', value='', regex=True)
        week1day6=a6.replace(to_replace=r'- Continuity', value='', regex=True)
        week1day7=a7.replace(to_replace=r'- Continuity', value='', regex=True)

        day1=df.iloc[1,1]
        day2=df.iloc[1,2]
        day3=df.iloc[1,3]
        day4=df.iloc[1,4]
        day5=df.iloc[1,5]
        day6=df.iloc[1,6]
        day7=df.iloc[1,7]

        week1day1['date']=day1
        week1day2['date']=day2
        week1day3['date']=day3
        week1day4['date']=day4
        week1day5['date']=day5
        week1day6['date']=day6
        week1day7['date']=day7

        provider1=df.iloc[3:23,1]
        provider2=df.iloc[3:23,2]
        provider3=df.iloc[3:23,3]
        provider4=df.iloc[3:23,4]
        provider5=df.iloc[3:23,5]
        provider6=df.iloc[3:23,6]
        provider7=df.iloc[3:23,7]

        week1day1['provider']=provider1
        week1day2['provider']=provider2
        week1day3['provider']=provider3
        week1day4['provider']=provider4
        week1day5['provider']=provider5
        week1day6['provider']=provider6
        week1day7['provider']=provider7

        week1day1['clinic']="HAMPDEN_NURS"
        week1day2['clinic']="HAMPDEN_NURS"
        week1day3['clinic']="HAMPDEN_NURS"
        week1day4['clinic']="HAMPDEN_NURS"
        week1day5['clinic']="HAMPDEN_NURS"
        week1day6['clinic']="HAMPDEN_NURS"
        week1day7['clinic']="HAMPDEN_NURS"

        week1=pd.DataFrame(columns=week1day1.columns)
        week1=pd.concat([week1,week1day1,week1day2,week1day3,week1day4,week1day5,week1day6,week1day7])
        week1.to_csv('week1.csv',index=False)

        clinictype=df.iloc[27:47, 0:1]
        b1 = pd.DataFrame(clinictype, columns = ['type'])
        b2 = pd.DataFrame(clinictype, columns = ['type'])
        b3 = pd.DataFrame(clinictype, columns = ['type'])
        b4 = pd.DataFrame(clinictype, columns = ['type'])
        b5 = pd.DataFrame(clinictype, columns = ['type'])
        b6 = pd.DataFrame(clinictype, columns = ['type'])
        b7 = pd.DataFrame(clinictype, columns = ['type'])

        b1['type']=clinictype
        b2['type']=clinictype
        b3['type']=clinictype
        b4['type']=clinictype
        b5['type']=clinictype
        b6['type']=clinictype
        b7['type']=clinictype

        week2day1=b1.replace(to_replace=r'- Continuity', value='', regex=True)
        week2day2=b2.replace(to_replace=r'- Continuity', value='', regex=True)
        week2day3=b3.replace(to_replace=r'- Continuity', value='', regex=True)
        week2day4=b4.replace(to_replace=r'- Continuity', value='', regex=True)
        week2day5=b5.replace(to_replace=r'- Continuity', value='', regex=True)
        week2day6=b6.replace(to_replace=r'- Continuity', value='', regex=True)
        week2day7=b7.replace(to_replace=r'- Continuity', value='', regex=True)

        day1=df.iloc[25,1]
        day2=df.iloc[25,2]
        day3=df.iloc[25,3]
        day4=df.iloc[25,4]
        day5=df.iloc[25,5]
        day6=df.iloc[25,6]
        day7=df.iloc[25,7]

        week2day1['date']=day1
        week2day2['date']=day2
        week2day3['date']=day3
        week2day4['date']=day4
        week2day5['date']=day5
        week2day6['date']=day6
        week2day7['date']=day7

        provider1=df.iloc[27:47,1]
        provider2=df.iloc[27:47,2]
        provider3=df.iloc[27:47,3]
        provider4=df.iloc[27:47,4]
        provider5=df.iloc[27:47,5]
        provider6=df.iloc[27:47,6]
        provider7=df.iloc[27:47,7]

        week2day1['provider']=provider1
        week2day2['provider']=provider2
        week2day3['provider']=provider3
        week2day4['provider']=provider4
        week2day5['provider']=provider5
        week2day6['provider']=provider6
        week2day7['provider']=provider7

        week2day1['clinic']="HAMPDEN_NURS"
        week2day2['clinic']="HAMPDEN_NURS"
        week2day3['clinic']="HAMPDEN_NURS"
        week2day4['clinic']="HAMPDEN_NURS"
        week2day5['clinic']="HAMPDEN_NURS"
        week2day6['clinic']="HAMPDEN_NURS"
        week2day7['clinic']="HAMPDEN_NURS"

        week2=pd.DataFrame(columns=week2day1.columns)
        week2=pd.concat([week2,week2day1,week2day2,week2day3,week2day4,week2day5,week2day6,week2day7])
        week2.to_csv('week2.csv',index=False)

        clinictype=df.iloc[51:71, 0:1]
        c1 = pd.DataFrame(clinictype, columns = ['type'])
        c2 = pd.DataFrame(clinictype, columns = ['type'])
        c3 = pd.DataFrame(clinictype, columns = ['type'])
        c4 = pd.DataFrame(clinictype, columns = ['type'])
        c5 = pd.DataFrame(clinictype, columns = ['type'])
        c6 = pd.DataFrame(clinictype, columns = ['type'])
        c7 = pd.DataFrame(clinictype, columns = ['type'])

        c1['type']=clinictype
        c2['type']=clinictype
        c3['type']=clinictype
        c4['type']=clinictype
        c5['type']=clinictype
        c6['type']=clinictype
        c7['type']=clinictype

        week3day1=c1.replace(to_replace=r'- Continuity', value='', regex=True)
        week3day2=c2.replace(to_replace=r'- Continuity', value='', regex=True)
        week3day3=c3.replace(to_replace=r'- Continuity', value='', regex=True)
        week3day4=c4.replace(to_replace=r'- Continuity', value='', regex=True)
        week3day5=c5.replace(to_replace=r'- Continuity', value='', regex=True)
        week3day6=c6.replace(to_replace=r'- Continuity', value='', regex=True)
        week3day7=c7.replace(to_replace=r'- Continuity', value='', regex=True)

        day1=df.iloc[49,1]
        day2=df.iloc[49,2]
        day3=df.iloc[49,3]
        day4=df.iloc[49,4]
        day5=df.iloc[49,5]
        day6=df.iloc[49,6]
        day7=df.iloc[49,7]

        week3day1['date']=day1
        week3day2['date']=day2
        week3day3['date']=day3
        week3day4['date']=day4
        week3day5['date']=day5
        week3day6['date']=day6
        week3day7['date']=day7

        provider1=df.iloc[51:71,1]
        provider2=df.iloc[51:71,2]
        provider3=df.iloc[51:71,3]
        provider4=df.iloc[51:71,4]
        provider5=df.iloc[51:71,5]
        provider6=df.iloc[51:71,6]
        provider7=df.iloc[51:71,7]

        week3day1['provider']=provider1
        week3day2['provider']=provider2
        week3day3['provider']=provider3
        week3day4['provider']=provider4
        week3day5['provider']=provider5
        week3day6['provider']=provider6
        week3day7['provider']=provider7

        week3day1['clinic']="HAMPDEN_NURS"
        week3day2['clinic']="HAMPDEN_NURS"
        week3day3['clinic']="HAMPDEN_NURS"
        week3day4['clinic']="HAMPDEN_NURS"
        week3day5['clinic']="HAMPDEN_NURS"
        week3day6['clinic']="HAMPDEN_NURS"
        week3day7['clinic']="HAMPDEN_NURS"

        week3=pd.DataFrame(columns=week3day1.columns)
        week3=pd.concat([week3,week3day1,week3day2,week3day3,week3day4,week3day5,week3day6,week3day7])
        week3.to_csv('week3.csv',index=False)

        clinictype=df.iloc[75:95, 0:1]
        d1 = pd.DataFrame(clinictype, columns = ['type'])
        d2 = pd.DataFrame(clinictype, columns = ['type'])
        d3 = pd.DataFrame(clinictype, columns = ['type'])
        d4 = pd.DataFrame(clinictype, columns = ['type'])
        d5 = pd.DataFrame(clinictype, columns = ['type'])
        d6 = pd.DataFrame(clinictype, columns = ['type'])
        d7 = pd.DataFrame(clinictype, columns = ['type'])

        d1['type']=clinictype
        d2['type']=clinictype
        d3['type']=clinictype
        d4['type']=clinictype
        d5['type']=clinictype
        d6['type']=clinictype
        d7['type']=clinictype

        week4day1=d1.replace(to_replace=r'- Continuity', value='', regex=True)
        week4day2=d2.replace(to_replace=r'- Continuity', value='', regex=True)
        week4day3=d3.replace(to_replace=r'- Continuity', value='', regex=True)
        week4day4=d4.replace(to_replace=r'- Continuity', value='', regex=True)
        week4day5=d5.replace(to_replace=r'- Continuity', value='', regex=True)
        week4day6=d6.replace(to_replace=r'- Continuity', value='', regex=True)
        week4day7=d7.replace(to_replace=r'- Continuity', value='', regex=True)


        day1=df.iloc[73,1]
        day2=df.iloc[73,2]
        day3=df.iloc[73,3]
        day4=df.iloc[73,4]
        day5=df.iloc[73,5]
        day6=df.iloc[73,6]
        day7=df.iloc[73,7]

        week4day1['date']=day1
        week4day2['date']=day2
        week4day3['date']=day3
        week4day4['date']=day4
        week4day5['date']=day5
        week4day6['date']=day6
        week4day7['date']=day7

        provider1=df.iloc[75:95,1]
        provider2=df.iloc[75:95,2]
        provider3=df.iloc[75:95,3]
        provider4=df.iloc[75:95,4]
        provider5=df.iloc[75:95,5]
        provider6=df.iloc[75:95,6]
        provider7=df.iloc[75:95,7]

        week4day1['provider']=provider1
        week4day2['provider']=provider2
        week4day3['provider']=provider3
        week4day4['provider']=provider4
        week4day5['provider']=provider5
        week4day6['provider']=provider6
        week4day7['provider']=provider7

        week4day1['clinic']="HAMPDEN_NURS"
        week4day2['clinic']="HAMPDEN_NURS"
        week4day3['clinic']="HAMPDEN_NURS"
        week4day4['clinic']="HAMPDEN_NURS"
        week4day5['clinic']="HAMPDEN_NURS"
        week4day6['clinic']="HAMPDEN_NURS"
        week4day7['clinic']="HAMPDEN_NURS"

        week4=pd.DataFrame(columns=week4day1.columns)
        week4=pd.concat([week4,week4day1,week4day2,week4day3,week4day4,week4day5,week4day6,week4day7])
        week4.to_csv('week4.csv',index=False)

        hope=pd.DataFrame(columns=week1.columns)
        hope=pd.concat([hope,week1,week2,week3,week4])
        hope.to_csv('extra.csv',index=False)

        hope['H'] = "H"
        extrai = hope[hope['type'] == 'AM'].copy()  # Make sure we're working with a copy
        extrai.loc[:, 'count'] = extrai.groupby(['date'])['provider'].cumcount() + 0  # Starts at H0 for AM
        extrai.loc[:, 'class'] = "H" + extrai['count'].astype(str)
        extrai = extrai.loc[:, ('date', 'type', 'provider', 'clinic', 'class')]
        extrai.to_csv('3.csv', index=False)

        # Handle PM Continuity for EXTRAI
        hope['H'] = "H"
        extraii = hope[hope['type'] == 'PM'].copy()  # Make sure we're working with a copy
        extraii.loc[:, 'count'] = extraii.groupby(['date'])['provider'].cumcount() + 10  # Starts at H10 for PM
        extraii.loc[:, 'class'] = "H" + extraii['count'].astype(str)
        extraii = extraii.loc[:, ('date', 'type', 'provider', 'clinic', 'class')]
        extraii.to_csv('4.csv', index=False)

        # Combine AM and PM DataFrames for EXTRAI
        extras = pd.DataFrame(columns=extrai.columns)
        extras = pd.concat([extrai, extraii])
        extras.to_csv('extra_hnurs.csv', index=False)

	#############################SJR_HOSP##############################################################################################
        import pandas as pd
        read_file = pd.read_excel (uploaded_opd_file, sheet_name='SJR_HOSP')
        read_file.to_csv ('extra.csv', index = False, header=False)
        df=pd.read_csv('extra.csv')

        clinictype=df.iloc[3:23, 0:1]
        a1 = pd.DataFrame(clinictype, columns = ['type'])
        a2 = pd.DataFrame(clinictype, columns = ['type'])
        a3 = pd.DataFrame(clinictype, columns = ['type'])
        a4 = pd.DataFrame(clinictype, columns = ['type'])
        a5 = pd.DataFrame(clinictype, columns = ['type'])
        a6 = pd.DataFrame(clinictype, columns = ['type'])
        a7 = pd.DataFrame(clinictype, columns = ['type'])

        a1['type']=clinictype
        a2['type']=clinictype
        a3['type']=clinictype
        a4['type']=clinictype
        a5['type']=clinictype
        a6['type']=clinictype
        a7['type']=clinictype

        week1day1=a1.replace(to_replace=r'- Continuity', value='', regex=True)
        week1day2=a2.replace(to_replace=r'- Continuity', value='', regex=True)
        week1day3=a3.replace(to_replace=r'- Continuity', value='', regex=True)
        week1day4=a4.replace(to_replace=r'- Continuity', value='', regex=True)
        week1day5=a5.replace(to_replace=r'- Continuity', value='', regex=True)
        week1day6=a6.replace(to_replace=r'- Continuity', value='', regex=True)
        week1day7=a7.replace(to_replace=r'- Continuity', value='', regex=True)

        day1=df.iloc[1,1]
        day2=df.iloc[1,2]
        day3=df.iloc[1,3]
        day4=df.iloc[1,4]
        day5=df.iloc[1,5]
        day6=df.iloc[1,6]
        day7=df.iloc[1,7]

        week1day1['date']=day1
        week1day2['date']=day2
        week1day3['date']=day3
        week1day4['date']=day4
        week1day5['date']=day5
        week1day6['date']=day6
        week1day7['date']=day7

        provider1=df.iloc[3:23,1]
        provider2=df.iloc[3:23,2]
        provider3=df.iloc[3:23,3]
        provider4=df.iloc[3:23,4]
        provider5=df.iloc[3:23,5]
        provider6=df.iloc[3:23,6]
        provider7=df.iloc[3:23,7]

        week1day1['provider']=provider1
        week1day2['provider']=provider2
        week1day3['provider']=provider3
        week1day4['provider']=provider4
        week1day5['provider']=provider5
        week1day6['provider']=provider6
        week1day7['provider']=provider7

        week1day1['clinic']="SJR_HOSP"
        week1day2['clinic']="SJR_HOSP"
        week1day3['clinic']="SJR_HOSP"
        week1day4['clinic']="SJR_HOSP"
        week1day5['clinic']="SJR_HOSP"
        week1day6['clinic']="SJR_HOSP"
        week1day7['clinic']="SJR_HOSP"

        week1=pd.DataFrame(columns=week1day1.columns)
        week1=pd.concat([week1,week1day1,week1day2,week1day3,week1day4,week1day5,week1day6,week1day7])
        week1.to_csv('week1.csv',index=False)

        clinictype=df.iloc[27:47, 0:1]
        b1 = pd.DataFrame(clinictype, columns = ['type'])
        b2 = pd.DataFrame(clinictype, columns = ['type'])
        b3 = pd.DataFrame(clinictype, columns = ['type'])
        b4 = pd.DataFrame(clinictype, columns = ['type'])
        b5 = pd.DataFrame(clinictype, columns = ['type'])
        b6 = pd.DataFrame(clinictype, columns = ['type'])
        b7 = pd.DataFrame(clinictype, columns = ['type'])

        b1['type']=clinictype
        b2['type']=clinictype
        b3['type']=clinictype
        b4['type']=clinictype
        b5['type']=clinictype
        b6['type']=clinictype
        b7['type']=clinictype

        week2day1=b1.replace(to_replace=r'- Continuity', value='', regex=True)
        week2day2=b2.replace(to_replace=r'- Continuity', value='', regex=True)
        week2day3=b3.replace(to_replace=r'- Continuity', value='', regex=True)
        week2day4=b4.replace(to_replace=r'- Continuity', value='', regex=True)
        week2day5=b5.replace(to_replace=r'- Continuity', value='', regex=True)
        week2day6=b6.replace(to_replace=r'- Continuity', value='', regex=True)
        week2day7=b7.replace(to_replace=r'- Continuity', value='', regex=True)

        day1=df.iloc[25,1]
        day2=df.iloc[25,2]
        day3=df.iloc[25,3]
        day4=df.iloc[25,4]
        day5=df.iloc[25,5]
        day6=df.iloc[25,6]
        day7=df.iloc[25,7]

        week2day1['date']=day1
        week2day2['date']=day2
        week2day3['date']=day3
        week2day4['date']=day4
        week2day5['date']=day5
        week2day6['date']=day6
        week2day7['date']=day7

        provider1=df.iloc[27:47,1]
        provider2=df.iloc[27:47,2]
        provider3=df.iloc[27:47,3]
        provider4=df.iloc[27:47,4]
        provider5=df.iloc[27:47,5]
        provider6=df.iloc[27:47,6]
        provider7=df.iloc[27:47,7]

        week2day1['provider']=provider1
        week2day2['provider']=provider2
        week2day3['provider']=provider3
        week2day4['provider']=provider4
        week2day5['provider']=provider5
        week2day6['provider']=provider6
        week2day7['provider']=provider7

        week2day1['clinic']="SJR_HOSP"
        week2day2['clinic']="SJR_HOSP"
        week2day3['clinic']="SJR_HOSP"
        week2day4['clinic']="SJR_HOSP"
        week2day5['clinic']="SJR_HOSP"
        week2day6['clinic']="SJR_HOSP"
        week2day7['clinic']="SJR_HOSP"

        week2=pd.DataFrame(columns=week2day1.columns)
        week2=pd.concat([week2,week2day1,week2day2,week2day3,week2day4,week2day5,week2day6,week2day7])
        week2.to_csv('week2.csv',index=False)

        clinictype=df.iloc[51:71, 0:1]
        c1 = pd.DataFrame(clinictype, columns = ['type'])
        c2 = pd.DataFrame(clinictype, columns = ['type'])
        c3 = pd.DataFrame(clinictype, columns = ['type'])
        c4 = pd.DataFrame(clinictype, columns = ['type'])
        c5 = pd.DataFrame(clinictype, columns = ['type'])
        c6 = pd.DataFrame(clinictype, columns = ['type'])
        c7 = pd.DataFrame(clinictype, columns = ['type'])

        c1['type']=clinictype
        c2['type']=clinictype
        c3['type']=clinictype
        c4['type']=clinictype
        c5['type']=clinictype
        c6['type']=clinictype
        c7['type']=clinictype

        week3day1=c1.replace(to_replace=r'- Continuity', value='', regex=True)
        week3day2=c2.replace(to_replace=r'- Continuity', value='', regex=True)
        week3day3=c3.replace(to_replace=r'- Continuity', value='', regex=True)
        week3day4=c4.replace(to_replace=r'- Continuity', value='', regex=True)
        week3day5=c5.replace(to_replace=r'- Continuity', value='', regex=True)
        week3day6=c6.replace(to_replace=r'- Continuity', value='', regex=True)
        week3day7=c7.replace(to_replace=r'- Continuity', value='', regex=True)

        day1=df.iloc[49,1]
        day2=df.iloc[49,2]
        day3=df.iloc[49,3]
        day4=df.iloc[49,4]
        day5=df.iloc[49,5]
        day6=df.iloc[49,6]
        day7=df.iloc[49,7]

        week3day1['date']=day1
        week3day2['date']=day2
        week3day3['date']=day3
        week3day4['date']=day4
        week3day5['date']=day5
        week3day6['date']=day6
        week3day7['date']=day7

        provider1=df.iloc[51:71,1]
        provider2=df.iloc[51:71,2]
        provider3=df.iloc[51:71,3]
        provider4=df.iloc[51:71,4]
        provider5=df.iloc[51:71,5]
        provider6=df.iloc[51:71,6]
        provider7=df.iloc[51:71,7]

        week3day1['provider']=provider1
        week3day2['provider']=provider2
        week3day3['provider']=provider3
        week3day4['provider']=provider4
        week3day5['provider']=provider5
        week3day6['provider']=provider6
        week3day7['provider']=provider7

        week3day1['clinic']="SJR_HOSP"
        week3day2['clinic']="SJR_HOSP"
        week3day3['clinic']="SJR_HOSP"
        week3day4['clinic']="SJR_HOSP"
        week3day5['clinic']="SJR_HOSP"
        week3day6['clinic']="SJR_HOSP"
        week3day7['clinic']="SJR_HOSP"

        week3=pd.DataFrame(columns=week3day1.columns)
        week3=pd.concat([week3,week3day1,week3day2,week3day3,week3day4,week3day5,week3day6,week3day7])
        week3.to_csv('week3.csv',index=False)

        clinictype=df.iloc[75:95, 0:1]
        d1 = pd.DataFrame(clinictype, columns = ['type'])
        d2 = pd.DataFrame(clinictype, columns = ['type'])
        d3 = pd.DataFrame(clinictype, columns = ['type'])
        d4 = pd.DataFrame(clinictype, columns = ['type'])
        d5 = pd.DataFrame(clinictype, columns = ['type'])
        d6 = pd.DataFrame(clinictype, columns = ['type'])
        d7 = pd.DataFrame(clinictype, columns = ['type'])

        d1['type']=clinictype
        d2['type']=clinictype
        d3['type']=clinictype
        d4['type']=clinictype
        d5['type']=clinictype
        d6['type']=clinictype
        d7['type']=clinictype

        week4day1=d1.replace(to_replace=r'- Continuity', value='', regex=True)
        week4day2=d2.replace(to_replace=r'- Continuity', value='', regex=True)
        week4day3=d3.replace(to_replace=r'- Continuity', value='', regex=True)
        week4day4=d4.replace(to_replace=r'- Continuity', value='', regex=True)
        week4day5=d5.replace(to_replace=r'- Continuity', value='', regex=True)
        week4day6=d6.replace(to_replace=r'- Continuity', value='', regex=True)
        week4day7=d7.replace(to_replace=r'- Continuity', value='', regex=True)


        day1=df.iloc[73,1]
        day2=df.iloc[73,2]
        day3=df.iloc[73,3]
        day4=df.iloc[73,4]
        day5=df.iloc[73,5]
        day6=df.iloc[73,6]
        day7=df.iloc[73,7]

        week4day1['date']=day1
        week4day2['date']=day2
        week4day3['date']=day3
        week4day4['date']=day4
        week4day5['date']=day5
        week4day6['date']=day6
        week4day7['date']=day7

        provider1=df.iloc[75:95,1]
        provider2=df.iloc[75:95,2]
        provider3=df.iloc[75:95,3]
        provider4=df.iloc[75:95,4]
        provider5=df.iloc[75:95,5]
        provider6=df.iloc[75:95,6]
        provider7=df.iloc[75:95,7]

        week4day1['provider']=provider1
        week4day2['provider']=provider2
        week4day3['provider']=provider3
        week4day4['provider']=provider4
        week4day5['provider']=provider5
        week4day6['provider']=provider6
        week4day7['provider']=provider7

        week4day1['clinic']="SJR_HOSP"
        week4day2['clinic']="SJR_HOSP"
        week4day3['clinic']="SJR_HOSP"
        week4day4['clinic']="SJR_HOSP"
        week4day5['clinic']="SJR_HOSP"
        week4day6['clinic']="SJR_HOSP"
        week4day7['clinic']="SJR_HOSP"

        week4=pd.DataFrame(columns=week4day1.columns)
        week4=pd.concat([week4,week4day1,week4day2,week4day3,week4day4,week4day5,week4day6,week4day7])
        week4.to_csv('week4.csv',index=False)

        hope=pd.DataFrame(columns=week1.columns)
        hope=pd.concat([hope,week1,week2,week3,week4])
        hope.to_csv('extra.csv',index=False)

        hope['H'] = "H"
        extrai = hope[hope['type'] == 'AM'].copy()  # Make sure we're working with a copy
        extrai.loc[:, 'count'] = extrai.groupby(['date'])['provider'].cumcount() + 0  # Starts at H0 for AM
        extrai.loc[:, 'class'] = "H" + extrai['count'].astype(str)
        extrai = extrai.loc[:, ('date', 'type', 'provider', 'clinic', 'class')]
        extrai.to_csv('3.csv', index=False)

        # Handle PM Continuity for EXTRAI
        hope['H'] = "H"
        extraii = hope[hope['type'] == 'PM'].copy()  # Make sure we're working with a copy
        extraii.loc[:, 'count'] = extraii.groupby(['date'])['provider'].cumcount() + 10  # Starts at H10 for PM
        extraii.loc[:, 'class'] = "H" + extraii['count'].astype(str)
        extraii = extraii.loc[:, ('date', 'type', 'provider', 'clinic', 'class')]
        extraii.to_csv('4.csv', index=False)

        # Combine AM and PM DataFrames for EXTRAI
        extras = pd.DataFrame(columns=extrai.columns)
        extras = pd.concat([extrai, extraii])
        extras.to_csv('extra_sjrhosp.csv', index=False)

	#############################ER_CONS##############################################################################################
        import pandas as pd
        read_file = pd.read_excel (uploaded_opd_file, sheet_name='ER_CONS')
        read_file.to_csv ('extra.csv', index = False, header=False)
        df=pd.read_csv('extra.csv')

        clinictype=df.iloc[3:23, 0:1]
        a1 = pd.DataFrame(clinictype, columns = ['type'])
        a2 = pd.DataFrame(clinictype, columns = ['type'])
        a3 = pd.DataFrame(clinictype, columns = ['type'])
        a4 = pd.DataFrame(clinictype, columns = ['type'])
        a5 = pd.DataFrame(clinictype, columns = ['type'])
        a6 = pd.DataFrame(clinictype, columns = ['type'])
        a7 = pd.DataFrame(clinictype, columns = ['type'])

        a1['type']=clinictype
        a2['type']=clinictype
        a3['type']=clinictype
        a4['type']=clinictype
        a5['type']=clinictype
        a6['type']=clinictype
        a7['type']=clinictype

        week1day1=a1.replace(to_replace=r'- Continuity', value='', regex=True)
        week1day2=a2.replace(to_replace=r'- Continuity', value='', regex=True)
        week1day3=a3.replace(to_replace=r'- Continuity', value='', regex=True)
        week1day4=a4.replace(to_replace=r'- Continuity', value='', regex=True)
        week1day5=a5.replace(to_replace=r'- Continuity', value='', regex=True)
        week1day6=a6.replace(to_replace=r'- Continuity', value='', regex=True)
        week1day7=a7.replace(to_replace=r'- Continuity', value='', regex=True)

        day1=df.iloc[1,1]
        day2=df.iloc[1,2]
        day3=df.iloc[1,3]
        day4=df.iloc[1,4]
        day5=df.iloc[1,5]
        day6=df.iloc[1,6]
        day7=df.iloc[1,7]

        week1day1['date']=day1
        week1day2['date']=day2
        week1day3['date']=day3
        week1day4['date']=day4
        week1day5['date']=day5
        week1day6['date']=day6
        week1day7['date']=day7

        provider1=df.iloc[3:23,1]
        provider2=df.iloc[3:23,2]
        provider3=df.iloc[3:23,3]
        provider4=df.iloc[3:23,4]
        provider5=df.iloc[3:23,5]
        provider6=df.iloc[3:23,6]
        provider7=df.iloc[3:23,7]

        week1day1['provider']=provider1
        week1day2['provider']=provider2
        week1day3['provider']=provider3
        week1day4['provider']=provider4
        week1day5['provider']=provider5
        week1day6['provider']=provider6
        week1day7['provider']=provider7

        week1day1['clinic']="ER_CONS"
        week1day2['clinic']="ER_CONS"
        week1day3['clinic']="ER_CONS"
        week1day4['clinic']="ER_CONS"
        week1day5['clinic']="ER_CONS"
        week1day6['clinic']="ER_CONS"
        week1day7['clinic']="ER_CONS"

        week1=pd.DataFrame(columns=week1day1.columns)
        week1=pd.concat([week1,week1day1,week1day2,week1day3,week1day4,week1day5,week1day6,week1day7])
        week1.to_csv('week1.csv',index=False)

        clinictype=df.iloc[27:47, 0:1]
        b1 = pd.DataFrame(clinictype, columns = ['type'])
        b2 = pd.DataFrame(clinictype, columns = ['type'])
        b3 = pd.DataFrame(clinictype, columns = ['type'])
        b4 = pd.DataFrame(clinictype, columns = ['type'])
        b5 = pd.DataFrame(clinictype, columns = ['type'])
        b6 = pd.DataFrame(clinictype, columns = ['type'])
        b7 = pd.DataFrame(clinictype, columns = ['type'])

        b1['type']=clinictype
        b2['type']=clinictype
        b3['type']=clinictype
        b4['type']=clinictype
        b5['type']=clinictype
        b6['type']=clinictype
        b7['type']=clinictype

        week2day1=b1.replace(to_replace=r'- Continuity', value='', regex=True)
        week2day2=b2.replace(to_replace=r'- Continuity', value='', regex=True)
        week2day3=b3.replace(to_replace=r'- Continuity', value='', regex=True)
        week2day4=b4.replace(to_replace=r'- Continuity', value='', regex=True)
        week2day5=b5.replace(to_replace=r'- Continuity', value='', regex=True)
        week2day6=b6.replace(to_replace=r'- Continuity', value='', regex=True)
        week2day7=b7.replace(to_replace=r'- Continuity', value='', regex=True)

        day1=df.iloc[25,1]
        day2=df.iloc[25,2]
        day3=df.iloc[25,3]
        day4=df.iloc[25,4]
        day5=df.iloc[25,5]
        day6=df.iloc[25,6]
        day7=df.iloc[25,7]

        week2day1['date']=day1
        week2day2['date']=day2
        week2day3['date']=day3
        week2day4['date']=day4
        week2day5['date']=day5
        week2day6['date']=day6
        week2day7['date']=day7

        provider1=df.iloc[27:47,1]
        provider2=df.iloc[27:47,2]
        provider3=df.iloc[27:47,3]
        provider4=df.iloc[27:47,4]
        provider5=df.iloc[27:47,5]
        provider6=df.iloc[27:47,6]
        provider7=df.iloc[27:47,7]

        week2day1['provider']=provider1
        week2day2['provider']=provider2
        week2day3['provider']=provider3
        week2day4['provider']=provider4
        week2day5['provider']=provider5
        week2day6['provider']=provider6
        week2day7['provider']=provider7

        week2day1['clinic']="ER_CONS"
        week2day2['clinic']="ER_CONS"
        week2day3['clinic']="ER_CONS"
        week2day4['clinic']="ER_CONS"
        week2day5['clinic']="ER_CONS"
        week2day6['clinic']="ER_CONS"
        week2day7['clinic']="ER_CONS"

        week2=pd.DataFrame(columns=week2day1.columns)
        week2=pd.concat([week2,week2day1,week2day2,week2day3,week2day4,week2day5,week2day6,week2day7])
        week2.to_csv('week2.csv',index=False)

        clinictype=df.iloc[51:71, 0:1]
        c1 = pd.DataFrame(clinictype, columns = ['type'])
        c2 = pd.DataFrame(clinictype, columns = ['type'])
        c3 = pd.DataFrame(clinictype, columns = ['type'])
        c4 = pd.DataFrame(clinictype, columns = ['type'])
        c5 = pd.DataFrame(clinictype, columns = ['type'])
        c6 = pd.DataFrame(clinictype, columns = ['type'])
        c7 = pd.DataFrame(clinictype, columns = ['type'])

        c1['type']=clinictype
        c2['type']=clinictype
        c3['type']=clinictype
        c4['type']=clinictype
        c5['type']=clinictype
        c6['type']=clinictype
        c7['type']=clinictype

        week3day1=c1.replace(to_replace=r'- Continuity', value='', regex=True)
        week3day2=c2.replace(to_replace=r'- Continuity', value='', regex=True)
        week3day3=c3.replace(to_replace=r'- Continuity', value='', regex=True)
        week3day4=c4.replace(to_replace=r'- Continuity', value='', regex=True)
        week3day5=c5.replace(to_replace=r'- Continuity', value='', regex=True)
        week3day6=c6.replace(to_replace=r'- Continuity', value='', regex=True)
        week3day7=c7.replace(to_replace=r'- Continuity', value='', regex=True)

        day1=df.iloc[49,1]
        day2=df.iloc[49,2]
        day3=df.iloc[49,3]
        day4=df.iloc[49,4]
        day5=df.iloc[49,5]
        day6=df.iloc[49,6]
        day7=df.iloc[49,7]

        week3day1['date']=day1
        week3day2['date']=day2
        week3day3['date']=day3
        week3day4['date']=day4
        week3day5['date']=day5
        week3day6['date']=day6
        week3day7['date']=day7

        provider1=df.iloc[51:71,1]
        provider2=df.iloc[51:71,2]
        provider3=df.iloc[51:71,3]
        provider4=df.iloc[51:71,4]
        provider5=df.iloc[51:71,5]
        provider6=df.iloc[51:71,6]
        provider7=df.iloc[51:71,7]

        week3day1['provider']=provider1
        week3day2['provider']=provider2
        week3day3['provider']=provider3
        week3day4['provider']=provider4
        week3day5['provider']=provider5
        week3day6['provider']=provider6
        week3day7['provider']=provider7

        week3day1['clinic']="ER_CONS"
        week3day2['clinic']="ER_CONS"
        week3day3['clinic']="ER_CONS"
        week3day4['clinic']="ER_CONS"
        week3day5['clinic']="ER_CONS"
        week3day6['clinic']="ER_CONS"
        week3day7['clinic']="ER_CONS"

        week3=pd.DataFrame(columns=week3day1.columns)
        week3=pd.concat([week3,week3day1,week3day2,week3day3,week3day4,week3day5,week3day6,week3day7])
        week3.to_csv('week3.csv',index=False)

        clinictype=df.iloc[75:95, 0:1]
        d1 = pd.DataFrame(clinictype, columns = ['type'])
        d2 = pd.DataFrame(clinictype, columns = ['type'])
        d3 = pd.DataFrame(clinictype, columns = ['type'])
        d4 = pd.DataFrame(clinictype, columns = ['type'])
        d5 = pd.DataFrame(clinictype, columns = ['type'])
        d6 = pd.DataFrame(clinictype, columns = ['type'])
        d7 = pd.DataFrame(clinictype, columns = ['type'])

        d1['type']=clinictype
        d2['type']=clinictype
        d3['type']=clinictype
        d4['type']=clinictype
        d5['type']=clinictype
        d6['type']=clinictype
        d7['type']=clinictype

        week4day1=d1.replace(to_replace=r'- Continuity', value='', regex=True)
        week4day2=d2.replace(to_replace=r'- Continuity', value='', regex=True)
        week4day3=d3.replace(to_replace=r'- Continuity', value='', regex=True)
        week4day4=d4.replace(to_replace=r'- Continuity', value='', regex=True)
        week4day5=d5.replace(to_replace=r'- Continuity', value='', regex=True)
        week4day6=d6.replace(to_replace=r'- Continuity', value='', regex=True)
        week4day7=d7.replace(to_replace=r'- Continuity', value='', regex=True)

        day1=df.iloc[73,1]
        day2=df.iloc[73,2]
        day3=df.iloc[73,3]
        day4=df.iloc[73,4]
        day5=df.iloc[73,5]
        day6=df.iloc[73,6]
        day7=df.iloc[73,7]

        week4day1['date']=day1
        week4day2['date']=day2
        week4day3['date']=day3
        week4day4['date']=day4
        week4day5['date']=day5
        week4day6['date']=day6
        week4day7['date']=day7

        provider1=df.iloc[75:95,1]
        provider2=df.iloc[75:95,2]
        provider3=df.iloc[75:95,3]
        provider4=df.iloc[75:95,4]
        provider5=df.iloc[75:95,5]
        provider6=df.iloc[75:95,6]
        provider7=df.iloc[75:95,7]

        week4day1['provider']=provider1
        week4day2['provider']=provider2
        week4day3['provider']=provider3
        week4day4['provider']=provider4
        week4day5['provider']=provider5
        week4day6['provider']=provider6
        week4day7['provider']=provider7

        week4day1['clinic']="ER_CONS"
        week4day2['clinic']="ER_CONS"
        week4day3['clinic']="ER_CONS"
        week4day4['clinic']="ER_CONS"
        week4day5['clinic']="ER_CONS"
        week4day6['clinic']="ER_CONS"
        week4day7['clinic']="ER_CONS"

        week4=pd.DataFrame(columns=week4day1.columns)
        week4=pd.concat([week4,week4day1,week4day2,week4day3,week4day4,week4day5,week4day6,week4day7])
        week4.to_csv('week4.csv',index=False)

        hope=pd.DataFrame(columns=week1.columns)
        hope=pd.concat([hope,week1,week2,week3,week4])
        hope.to_csv('extra.csv',index=False)

        hope['H'] = "H"
        extrai = hope[hope['type'] == 'AM'].copy()  # Make sure we're working with a copy
        extrai.loc[:, 'count'] = extrai.groupby(['date'])['provider'].cumcount() + 0  # Starts at H0 for AM
        extrai.loc[:, 'class'] = "H" + extrai['count'].astype(str)
        extrai = extrai.loc[:, ('date', 'type', 'provider', 'clinic', 'class')]
        extrai.to_csv('3.csv', index=False)

        # Handle PM Continuity for EXTRAI
        hope['H'] = "H"
        extraii = hope[hope['type'] == 'PM'].copy()  # Make sure we're working with a copy
        extraii.loc[:, 'count'] = extraii.groupby(['date'])['provider'].cumcount() + 10  # Starts at H10 for PM
        extraii.loc[:, 'class'] = "H" + extraii['count'].astype(str)
        extraii = extraii.loc[:, ('date', 'type', 'provider', 'clinic', 'class')]
        extraii.to_csv('4.csv', index=False)

        # Combine AM and PM DataFrames for EXTRAI
        extras = pd.DataFrame(columns=extrai.columns)
        extras = pd.concat([extrai, extraii])
        extras.to_csv('extra_ercons.csv', index=False)

	#############################NF##############################################################################################
        import pandas as pd
        read_file = pd.read_excel (uploaded_opd_file, sheet_name='NF')
        read_file.to_csv ('extra.csv', index = False, header=False)
        df=pd.read_csv('extra.csv')

        clinictype=df.iloc[3:23, 0:1]
        a1 = pd.DataFrame(clinictype, columns = ['type'])
        a2 = pd.DataFrame(clinictype, columns = ['type'])
        a3 = pd.DataFrame(clinictype, columns = ['type'])
        a4 = pd.DataFrame(clinictype, columns = ['type'])
        a5 = pd.DataFrame(clinictype, columns = ['type'])
        a6 = pd.DataFrame(clinictype, columns = ['type'])
        a7 = pd.DataFrame(clinictype, columns = ['type'])

        a1['type']=clinictype
        a2['type']=clinictype
        a3['type']=clinictype
        a4['type']=clinictype
        a5['type']=clinictype
        a6['type']=clinictype
        a7['type']=clinictype

        week1day1=a1.replace(to_replace=r'- Continuity', value='', regex=True)
        week1day2=a2.replace(to_replace=r'- Continuity', value='', regex=True)
        week1day3=a3.replace(to_replace=r'- Continuity', value='', regex=True)
        week1day4=a4.replace(to_replace=r'- Continuity', value='', regex=True)
        week1day5=a5.replace(to_replace=r'- Continuity', value='', regex=True)
        week1day6=a6.replace(to_replace=r'- Continuity', value='', regex=True)
        week1day7=a7.replace(to_replace=r'- Continuity', value='', regex=True)

        day1=df.iloc[1,1]
        day2=df.iloc[1,2]
        day3=df.iloc[1,3]
        day4=df.iloc[1,4]
        day5=df.iloc[1,5]
        day6=df.iloc[1,6]
        day7=df.iloc[1,7]

        week1day1['date']=day1
        week1day2['date']=day2
        week1day3['date']=day3
        week1day4['date']=day4
        week1day5['date']=day5
        week1day6['date']=day6
        week1day7['date']=day7

        provider1=df.iloc[3:23,1]
        provider2=df.iloc[3:23,2]
        provider3=df.iloc[3:23,3]
        provider4=df.iloc[3:23,4]
        provider5=df.iloc[3:23,5]
        provider6=df.iloc[3:23,6]
        provider7=df.iloc[3:23,7]

        week1day1['provider']=provider1
        week1day2['provider']=provider2
        week1day3['provider']=provider3
        week1day4['provider']=provider4
        week1day5['provider']=provider5
        week1day6['provider']=provider6
        week1day7['provider']=provider7

        week1day1['clinic']="NF"
        week1day2['clinic']="NF"
        week1day3['clinic']="NF"
        week1day4['clinic']="NF"
        week1day5['clinic']="NF"
        week1day6['clinic']="NF"
        week1day7['clinic']="NF"

        week1=pd.DataFrame(columns=week1day1.columns)
        week1=pd.concat([week1,week1day1,week1day2,week1day3,week1day4,week1day5,week1day6,week1day7])
        week1.to_csv('week1.csv',index=False)

        clinictype=df.iloc[27:47, 0:1]
        b1 = pd.DataFrame(clinictype, columns = ['type'])
        b2 = pd.DataFrame(clinictype, columns = ['type'])
        b3 = pd.DataFrame(clinictype, columns = ['type'])
        b4 = pd.DataFrame(clinictype, columns = ['type'])
        b5 = pd.DataFrame(clinictype, columns = ['type'])
        b6 = pd.DataFrame(clinictype, columns = ['type'])
        b7 = pd.DataFrame(clinictype, columns = ['type'])

        b1['type']=clinictype
        b2['type']=clinictype
        b3['type']=clinictype
        b4['type']=clinictype
        b5['type']=clinictype
        b6['type']=clinictype
        b7['type']=clinictype

        week2day1=b1.replace(to_replace=r'- Continuity', value='', regex=True)
        week2day2=b2.replace(to_replace=r'- Continuity', value='', regex=True)
        week2day3=b3.replace(to_replace=r'- Continuity', value='', regex=True)
        week2day4=b4.replace(to_replace=r'- Continuity', value='', regex=True)
        week2day5=b5.replace(to_replace=r'- Continuity', value='', regex=True)
        week2day6=b6.replace(to_replace=r'- Continuity', value='', regex=True)
        week2day7=b7.replace(to_replace=r'- Continuity', value='', regex=True)

        day1=df.iloc[25,1]
        day2=df.iloc[25,2]
        day3=df.iloc[25,3]
        day4=df.iloc[25,4]
        day5=df.iloc[25,5]
        day6=df.iloc[25,6]
        day7=df.iloc[25,7]

        week2day1['date']=day1
        week2day2['date']=day2
        week2day3['date']=day3
        week2day4['date']=day4
        week2day5['date']=day5
        week2day6['date']=day6
        week2day7['date']=day7

        provider1=df.iloc[27:47,1]
        provider2=df.iloc[27:47,2]
        provider3=df.iloc[27:47,3]
        provider4=df.iloc[27:47,4]
        provider5=df.iloc[27:47,5]
        provider6=df.iloc[27:47,6]
        provider7=df.iloc[27:47,7]

        week2day1['provider']=provider1
        week2day2['provider']=provider2
        week2day3['provider']=provider3
        week2day4['provider']=provider4
        week2day5['provider']=provider5
        week2day6['provider']=provider6
        week2day7['provider']=provider7

        week2day1['clinic']="NF"
        week2day2['clinic']="NF"
        week2day3['clinic']="NF"
        week2day4['clinic']="NF"
        week2day5['clinic']="NF"
        week2day6['clinic']="NF"
        week2day7['clinic']="NF"

        week2=pd.DataFrame(columns=week2day1.columns)
        week2=pd.concat([week2,week2day1,week2day2,week2day3,week2day4,week2day5,week2day6,week2day7])
        week2.to_csv('week2.csv',index=False)

        clinictype=df.iloc[51:71, 0:1]
        c1 = pd.DataFrame(clinictype, columns = ['type'])
        c2 = pd.DataFrame(clinictype, columns = ['type'])
        c3 = pd.DataFrame(clinictype, columns = ['type'])
        c4 = pd.DataFrame(clinictype, columns = ['type'])
        c5 = pd.DataFrame(clinictype, columns = ['type'])
        c6 = pd.DataFrame(clinictype, columns = ['type'])
        c7 = pd.DataFrame(clinictype, columns = ['type'])

        c1['type']=clinictype
        c2['type']=clinictype
        c3['type']=clinictype
        c4['type']=clinictype
        c5['type']=clinictype
        c6['type']=clinictype
        c7['type']=clinictype

        week3day1=c1.replace(to_replace=r'- Continuity', value='', regex=True)
        week3day2=c2.replace(to_replace=r'- Continuity', value='', regex=True)
        week3day3=c3.replace(to_replace=r'- Continuity', value='', regex=True)
        week3day4=c4.replace(to_replace=r'- Continuity', value='', regex=True)
        week3day5=c5.replace(to_replace=r'- Continuity', value='', regex=True)
        week3day6=c6.replace(to_replace=r'- Continuity', value='', regex=True)
        week3day7=c7.replace(to_replace=r'- Continuity', value='', regex=True)

        day1=df.iloc[49,1]
        day2=df.iloc[49,2]
        day3=df.iloc[49,3]
        day4=df.iloc[49,4]
        day5=df.iloc[49,5]
        day6=df.iloc[49,6]
        day7=df.iloc[49,7]

        week3day1['date']=day1
        week3day2['date']=day2
        week3day3['date']=day3
        week3day4['date']=day4
        week3day5['date']=day5
        week3day6['date']=day6
        week3day7['date']=day7

        provider1=df.iloc[51:71,1]
        provider2=df.iloc[51:71,2]
        provider3=df.iloc[51:71,3]
        provider4=df.iloc[51:71,4]
        provider5=df.iloc[51:71,5]
        provider6=df.iloc[51:71,6]
        provider7=df.iloc[51:71,7]

        week3day1['provider']=provider1
        week3day2['provider']=provider2
        week3day3['provider']=provider3
        week3day4['provider']=provider4
        week3day5['provider']=provider5
        week3day6['provider']=provider6
        week3day7['provider']=provider7

        week3day1['clinic']="NF"
        week3day2['clinic']="NF"
        week3day3['clinic']="NF"
        week3day4['clinic']="NF"
        week3day5['clinic']="NF"
        week3day6['clinic']="NF"
        week3day7['clinic']="NF"

        week3=pd.DataFrame(columns=week3day1.columns)
        week3=pd.concat([week3,week3day1,week3day2,week3day3,week3day4,week3day5,week3day6,week3day7])
        week3.to_csv('week3.csv',index=False)

        clinictype=df.iloc[75:95, 0:1]
        d1 = pd.DataFrame(clinictype, columns = ['type'])
        d2 = pd.DataFrame(clinictype, columns = ['type'])
        d3 = pd.DataFrame(clinictype, columns = ['type'])
        d4 = pd.DataFrame(clinictype, columns = ['type'])
        d5 = pd.DataFrame(clinictype, columns = ['type'])
        d6 = pd.DataFrame(clinictype, columns = ['type'])
        d7 = pd.DataFrame(clinictype, columns = ['type'])

        d1['type']=clinictype
        d2['type']=clinictype
        d3['type']=clinictype
        d4['type']=clinictype
        d5['type']=clinictype
        d6['type']=clinictype
        d7['type']=clinictype

        week4day1=d1.replace(to_replace=r'- Continuity', value='', regex=True)
        week4day2=d2.replace(to_replace=r'- Continuity', value='', regex=True)
        week4day3=d3.replace(to_replace=r'- Continuity', value='', regex=True)
        week4day4=d4.replace(to_replace=r'- Continuity', value='', regex=True)
        week4day5=d5.replace(to_replace=r'- Continuity', value='', regex=True)
        week4day6=d6.replace(to_replace=r'- Continuity', value='', regex=True)
        week4day7=d7.replace(to_replace=r'- Continuity', value='', regex=True)

        day1=df.iloc[73,1]
        day2=df.iloc[73,2]
        day3=df.iloc[73,3]
        day4=df.iloc[73,4]
        day5=df.iloc[73,5]
        day6=df.iloc[73,6]
        day7=df.iloc[73,7]

        week4day1['date']=day1
        week4day2['date']=day2
        week4day3['date']=day3
        week4day4['date']=day4
        week4day5['date']=day5
        week4day6['date']=day6
        week4day7['date']=day7

        provider1=df.iloc[75:95,1]
        provider2=df.iloc[75:95,2]
        provider3=df.iloc[75:95,3]
        provider4=df.iloc[75:95,4]
        provider5=df.iloc[75:95,5]
        provider6=df.iloc[75:95,6]
        provider7=df.iloc[75:95,7]

        week4day1['provider']=provider1
        week4day2['provider']=provider2
        week4day3['provider']=provider3
        week4day4['provider']=provider4
        week4day5['provider']=provider5
        week4day6['provider']=provider6
        week4day7['provider']=provider7

        week4day1['clinic']="NF"
        week4day2['clinic']="NF"
        week4day3['clinic']="NF"
        week4day4['clinic']="NF"
        week4day5['clinic']="NF"
        week4day6['clinic']="NF"
        week4day7['clinic']="NF"

        week4=pd.DataFrame(columns=week4day1.columns)
        week4=pd.concat([week4,week4day1,week4day2,week4day3,week4day4,week4day5,week4day6,week4day7])
        week4.to_csv('week4.csv',index=False)

        hope=pd.DataFrame(columns=week1.columns)
        hope=pd.concat([hope,week1,week2,week3,week4])
        hope.to_csv('extra.csv',index=False)

        hope['H'] = "H"
        extrai = hope[hope['type'] == 'AM'].copy()  # Make sure we're working with a copy
        extrai.loc[:, 'count'] = extrai.groupby(['date'])['provider'].cumcount() + 0  # Starts at H0 for AM
        extrai.loc[:, 'class'] = "H" + extrai['count'].astype(str)
        extrai = extrai.loc[:, ('date', 'type', 'provider', 'clinic', 'class')]
        extrai.to_csv('3.csv', index=False)

        # Handle PM Continuity for EXTRAI
        hope['H'] = "H"
        extraii = hope[hope['type'] == 'PM'].copy()  # Make sure we're working with a copy
        extraii.loc[:, 'count'] = extraii.groupby(['date'])['provider'].cumcount() + 10  # Starts at H10 for PM
        extraii.loc[:, 'class'] = "H" + extraii['count'].astype(str)
        extraii = extraii.loc[:, ('date', 'type', 'provider', 'clinic', 'class')]
        extraii.to_csv('4.csv', index=False)

        # Combine AM and PM DataFrames for EXTRAI
        extras = pd.DataFrame(columns=extrai.columns)
        extras = pd.concat([extrai, extraii])
        extras.to_csv('extra_nf.csv', index=False)
############################################################################
        df1=pd.read_csv('etowns.csv')
        df2=pd.read_csv('hopes.csv')
        df3=pd.read_csv('nyess.csv')
        df4=pd.read_csv('mhss.csv')
        df5=pd.read_csv('extra_a.csv')
        df6=pd.read_csv('extra_c.csv')
        df7=pd.read_csv('extra_p.csv')
        df8=pd.read_csv('extra_picu.csv')
        df9=pd.read_csv('extra_nurs.csv')
        df10=pd.read_csv('extra_hnurs.csv')
        df11=pd.read_csv('extra_sjrhosp.csv')
        df12=pd.read_csv('extra_ercons.csv')
        df13=pd.read_csv('extra_nf.csv')
	    
        dfx=pd.DataFrame(columns=df1.columns)
        dfx=pd.concat([dfx,df1,df2,df3,df4,df5,df6,df7,df8,df9,df10,df11,df12,df13])
        dfx['providers']=dfx['provider'].str.split('~').str[0]
        dfx['student']=dfx['provider'].str.split('~').str[1]
        dfx1=dfx[['date','type','providers','student','clinic','provider','class']]

        dfx1['date'] = pd.to_datetime(dfx1['date'])
        dfx1['date'] = dfx1['date'].dt.strftime('%m/%d/%Y')

        mydict = {}
        with open('xxxDATEMAP.csv', mode='r')as inp:     #file is the objects you want to map. I want to map the IMP in this file to diagnosis.csv
            reader = csv.reader(inp)
            df1 = {rows[0]:rows[1] for rows in reader} 
        dfx1['datecode'] = dfx1.date.map(df1)               #'type' is the new column in the diagnosis file. 'encounter_id' is the key you are using to MAP 

        dfx1.to_excel('STUDENTLIST.xlsx',index=False)
        dfx1.to_csv('STUDENTLIST.csv',index=False)

        dfx1['type'] = dfx1['type'].str.lstrip()
        dfx1['type'] = dfx1['type'].str.rstrip()

        dfx1['providers'] = dfx1['providers'].str.lstrip()
        dfx1['providers'] = dfx1['providers'].str.rstrip()

        dfx1['student'] = dfx1['student'].str.lstrip()
        dfx1['student'] = dfx1['student'].str.rstrip()

        dfs = [["AM","S1"],["PM","S2"],["AM - ACUTES","S11"],["PM - ACUTES","S12"]]
        dftype = pd.DataFrame(dfs, columns = ['type', 'datecode2'])
        dftype.to_csv('dftype.csv',index=False)

        mydict = {}
        with open('dftype.csv', mode='r')as inp:     #file is the objects you want to map. I want to map the IMP in this file to diagnosis.csv
            reader = csv.reader(inp)
            df1 = {rows[0]:rows[1] for rows in reader} 
        dfx1['datecode2'] = dfx1.type.map(df1)               #'type' is the new column in the diagnosis file. 'encounter_id' is the key you are using to MAP 

        dfx1.to_excel('PALIST.xlsx',index=False)

        dfx1.to_csv('PALIST.csv',index=False)

        dfx1 = pd.read_csv('PALIST.csv')
        #df = dfx1
        #df=pd.read_csv('PALIST.csv',dtype=str)
        #import io
        #output = io.StringIO()
        #df.to_csv(output, index=False)
        #output.seek(0)

        # Streamlit download button
        #st.download_button(
        #    label="Download CSV File",
        #    data=output.getvalue(),
        #    file_name="PALIST.csv",
        #    mime="text/csv"
        #)
 
 
        new_row = pd.DataFrame({'date':0, 'type':0, 'providers':0,
                                'student':0, 'clinic':0, 'provider':0,
                                'class':0, 'datecode':0, 'datecode2':0},
                                                                    index =[0])
        # simply concatenate both dataframes
        df = pd.concat([new_row, dfx1]).reset_index(drop = True)

        df.to_csv('PALIST.csv',index=False)
        
        #st.dataframe(df.head())
        
        df = pd.read_excel('Book4.xlsx')

        # Keep the first NaN in the first row as it is for the first column
        df.iloc[0, 0] = np.nan  # Ensure the first cell is NaN (or leave it as it is)

        # Initialize a counter for BLANK replacements
        blank_counter = 1

        # Replace NaN values in the first column with BLANK1, BLANK2, etc. starting from the second row
        for i in range(1, len(df)):
            if pd.isna(df.iloc[i, 0]):
                df.iloc[i, 0] = f'BLANK{blank_counter}'
                blank_counter += 1

        x = (df.loc[0, 'Week 1'])

        x = x.strftime("%m/%d/%Y")

        test_date = datetime.datetime.strptime(x, "%m/%d/%Y")

        # initializing K
        K = 28

        res = []

        for day in range(K):
            date = (test_date + datetime.timedelta(days = day)).strftime("%-m/%-d/%Y")
            res.append(date)

        #res

        dates = pd.DataFrame(res, columns =['dates'])

        dates['x'] = "y"

        dates['i'] = dates.index+1

        dates['i2']= dates.index+0

        dates['T'] = "T" + dates['i2'].astype(str)

        dates['x'] = dates['x'].astype(str)+dates['i'].astype(str)

        dates['x'] = dates['x'].astype(str) + "=" + "'"+dates['dates'].astype(str) + "'"

        dateT = dates[['dates','T']]

        dates = dates[['x']]

        dates.to_csv('dates.csv',index=False)

        dateT.to_csv('datesT.csv',index=False)

        import numpy as np

        datesdf = pd.read_csv('dates.csv')

        dates = datesdf['x'].astype(str)

        numpy_array=dates.to_numpy()
        np.savetxt("dates.py",numpy_array, fmt="%s")

        exec(open('dates.py').read())

        #################################STUDENT NAME and CLINICAL EXPERIENCES INPUT#################################
        column = "Student Name:"
        L=df[(column)].unique().tolist() #Must Use Unique
        xf200 = pd.DataFrame({'col':L})
        xf200['i'] = xf200.index

        column = "Week 1"
        L=df[(column)].tolist() #No Unique Required 
        xf201 = pd.DataFrame({'col':L})
        xf201['i'] = xf201.index

        column = "Week 2"
        L=df[(column)].tolist()
        xf202 = pd.DataFrame({'col':L})
        xf202['i'] = xf202.index

        column = "Week 3"
        L=df[(column)].tolist()
        xf203 = pd.DataFrame({'col':L})
        xf203['i'] = xf203.index

        column = "Week 4"
        L=df[(column)].tolist()
        xf204 = pd.DataFrame({'col':L})
        xf204['i'] = xf204.index

        ######################################################################################################################
        # Create a workbook and add the main worksheet
        # Create a workbook and add the main worksheet
        workbook = xlsxwriter.Workbook('Main_Schedule_MS.xlsx')

        # Define the formats
        format1 = workbook.add_format({
            'font_size': 14,
            'bold': 1,
            'align': 'center',
            'valign': 'vcenter',
            'font_color': 'black',
            'text_wrap': True,
            'bg_color': '#FEFFCC',
            'border': 1
        })

        format2 = workbook.add_format({
            'font_size': 10,
            'bold': 1,
            'align': 'center',
            'valign': 'vcenter',
            'font_color': 'yellow',
            'bg_color': 'black',
            'border': 1,
            'text_wrap': True
        })

        format3 = workbook.add_format({
            'font_size':12,
            'bold': 1,
            'align': 'center',
            'valign': 'vcenter',
            'font_color':'black',
            'bg_color':'#FFC7CE',
            'border':1
        })

        format4 = workbook.add_format({'num_format':'mm/dd/yyyy','font_size':12,'bold': 1,'align': 'center','valign': 'vcenter','font_color':'black','bg_color':'#F4F6F7','border':1})
        format5 = workbook.add_format({'font_size':12,'bold': 1,'align': 'center','valign': 'vcenter','font_color':'black','bg_color':'#F4F6F7','border':1})
        format6 = workbook.add_format({'bg_color':'black','border':1})
        format7 = workbook.add_format({'font_size':12,'bold': 1,'align': 'center','valign': 'vcenter','font_color':'black','bg_color':'#90EE90','border':1})
        format8 = workbook.add_format({'font_size':12,'bold': 1,'align': 'center','valign':'vcenter','font_color':'black','bg_color':'#89CFF0','border':1})
        # Initialize a list to store worksheets
        worksheets = []

        # Add a worksheet for each unique name in the DataFrame
        for name in xf200['col']:
            # Ensure that the name is a valid string (i.e., no NaN or floats)
            if isinstance(name, str):  # Check if it's a string
                name = name[:31]  # Truncate the name if it's longer than 31 characters
            else:
                continue  # Skip invalid entries (like NaN or numbers)

            # Ensure the name is not an empty string
            if len(name) > 0:
                worksheet = workbook.add_worksheet(name)
                worksheets.append(worksheet)  # Store the worksheet in the list

                # Merge range for the student name column
                worksheet.merge_range('A1:A2', 'Student Name:', format1)
                worksheet.merge_range('B1:B2', name, format1)

                # Merge range for the note
                note = '*Note** Asynchronous time is for coursework only. During this time period, we expect students to do coursework, be available for any additional educational activities, and any extra clinical time that may be available. If the student is not available during this time period and has not made an absence request, the student will be cited for unprofessionalism and will risk failing the course.'
                worksheet.merge_range('C1:H2', note, format2)

                # Set column widths (you can set them all at once or in a loop)
                worksheet.set_column('A:A', 20)
                worksheet.set_column('B:B', 30)
                worksheet.set_column('C:G', 30)
                worksheet.set_column('H:H', 155)
		    
                # Set row height for header row
                worksheet.set_row(0, 37.25)

                days_of_week = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']

                # Start row positions for each week
                start_rows = [3, 11, 19, 27]

                # Loop through each week's starting row and write the days
                for start_row in start_rows:
                    for col, day in enumerate(days_of_week, start=1):  
                        worksheet.write(chr(65 + col) + str(start_row), day, format3)

                # List of weeks
                weeks = ['Week 1', 'Week 2', 'Week 3', 'Week 4']

                for week_index, week in enumerate(weeks):
                    worksheet.write(f'A{4 + (week_index * 8)}', week, format3)

                for week_index in range(4):  # For each week (4 weeks)
                    row = 6 + (week_index * 8)  # Calculate row for AM (6, 14, 22, 30)
                    worksheet.write(f'A{row}', 'AM', format3)

                for week_index in range(4):  # For each week (4 weeks)
                    row = 7 + (week_index * 8)  # Calculate row for PM (7, 15, 23, 31)
                    worksheet.write(f'A{row}', 'PM', format3)

                column_idx = 1  # Starting at column B (Excel columns are 1-indexed)
                date_idx = 0   # Start from the first date in `res`

                # Loop through the weeks and the corresponding dates
                for week_row in start_rows:
                    for col in range(7):  # Loop through Monday to Sunday (7 days)
                        if date_idx < len(res):
                            date = res[date_idx]
                            worksheet.write(week_row, column_idx + col, date, format4)
                            date_idx += 1  # Move to the next date

                worksheet.set_zoom(70)
                worksheet.merge_range('A10:H10','',format6)
                worksheet.merge_range('A18:H18','',format6)
                worksheet.merge_range('A26:H26','',format6)
                worksheet.merge_range('A34:H34','',format6)

                worksheet.write('A9',' ', format7)
                worksheet.write('B9',' ', format7)
                worksheet.write('C9',' ', format7)
                worksheet.write('D9',' ', format7)
                worksheet.write('E9',' ', format7)
                worksheet.write('F9',' ', format7)
                worksheet.write('G9',' ', format7)
                worksheet.write('H9',' ', format7)

                worksheet.write('A17',' ', format7)
                worksheet.write('B17',' ', format7)
                worksheet.write('C17',' ', format7)
                worksheet.write('D17',' ', format7)
                worksheet.write('E17',' ', format7)
                worksheet.write('F17',' ', format7)
                worksheet.write('G17',' ', format7)
                worksheet.write('H17',' ', format7)

                worksheet.write('A25',' ', format7)
                worksheet.write('B25',' ', format7)
                worksheet.write('C25',' ', format7)
                worksheet.write('D25',' ', format7)
                worksheet.write('E25',' ', format7)
                worksheet.write('F25',' ', format7)
                worksheet.write('G25',' ', format7)
                worksheet.write('H25',' ', format7)

                worksheet.write('A33',' ', format7)
                worksheet.write('B33',' ', format7)
                worksheet.write('C33',' ', format7)
                worksheet.write('D33',' ', format7)
                worksheet.write('E33',' ', format7)
                worksheet.write('F33',' ', format7)
                worksheet.write('G33',' ', format7)
                worksheet.write('H33',' ', format7)

                # Writing to row 8
                worksheet.write('A8', 'ASSIGNMENT DUE:', format8)
                worksheet.write('B8', ' ', format8)
                worksheet.write('C8', ' ', format8)
                worksheet.write('D8', ' ', format8)
                worksheet.write('E8', ' ', format8)
                worksheet.write('F8', 'Ask for Feedback!', format8)
                worksheet.write('G8', ' ', format8)
                worksheet.write('H8', 'Quiz 1 Due', format8)

                # Writing to row 16
                worksheet.write('A16', 'ASSIGNMENT DUE:', format8)
                worksheet.write('B16', ' ', format8)
                worksheet.write('C16', ' ', format8)
                worksheet.write('D16', ' ', format8)
                worksheet.write('E16', ' ', format8)
                worksheet.write('F16', 'Ask for Feedback!', format8)
                worksheet.write('G16', ' ', format8)
                worksheet.write('H16', 'Quiz 2, Pediatric Documentation #1 Due', format8)

                # Writing to row 24
                worksheet.write('A24', 'ASSIGNMENT DUE:', format8)
                worksheet.write('B24', ' ', format8)
                worksheet.write('C24', ' ', format8)
                worksheet.write('D24', ' ', format8)
                worksheet.write('E24', ' ', format8)
                worksheet.write('F24', 'Ask for Feedback!', format8)
                worksheet.write('G24', ' ', format8)
                worksheet.write('H24', 'Quiz 3 Due', format8)

                # Writing to row 32
                worksheet.write('A32', 'ASSIGNMENT DUE:', format8)
                worksheet.write('B32', ' ', format8)
                worksheet.write('C32', ' ', format8)
                worksheet.write('D32', ' ', format8)
                worksheet.write('E32', ' ', format8)
                worksheet.write('F32', ' ', format8)
                worksheet.write('G32', ' ', format8)
                worksheet.write('H32', 'Quiz 4, Pediatric Documentation #2, Social Drivers of Health Assessment Form, Developmental Assessment of Pediatric Patient Form, Clinical Encounters are Due!', format8)

        # Now, we'll write the location data
        locations = xf201['col'].tolist()[1:]  # Assuming 'col' is a column in xf201 dataframe

        # Ensure locations match the number of worksheets
        if len(locations) != len(worksheets):
            raise ValueError(f"Number of locations ({len(locations)}) does not match number of worksheets ({len(worksheets)})")

        # Iterate over the names and locations and write the data
        for i, (name, location) in enumerate(zip(xf200['col'], locations)):
            worksheet = worksheets[i]  # Get the corresponding worksheet for this index

            # Write the location data to rows 6 and 7 for each worksheet
            for row in range(6, 8):  # Rows 6 and 7
                worksheet.write(f'B{row}', location, format5)
                worksheet.write(f'C{row}', location, format5)
                worksheet.write(f'D{row}', location, format5)
                worksheet.write(f'E{row}', location, format5)
                worksheet.write(f'F{row}', location, format5)
                worksheet.write(f'G{row}', "OFF", format5)
                worksheet.write(f'H{row}', "OFF", format5)

        # Now, we'll write the location data
        locations = xf202['col'].tolist()[1:]  # Assuming 'col' is a column in xf201 dataframe

        # Ensure locations match the number of worksheets
        if len(locations) != len(worksheets):
            raise ValueError(f"Number of locations ({len(locations)}) does not match number of worksheets ({len(worksheets)})")

        # Iterate over the names and locations and write the data
        for i, (name, location) in enumerate(zip(xf200['col'], locations)):
            worksheet = worksheets[i]  # Get the corresponding worksheet for this index

            for row in range(14, 16): 
                worksheet.write(f'B{row}', location, format5)
                worksheet.write(f'C{row}', location, format5)
                worksheet.write(f'D{row}', location, format5)
                worksheet.write(f'E{row}', location, format5)
                worksheet.write(f'F{row}', location, format5)
                worksheet.write(f'G{row}', "OFF", format5)
                worksheet.write(f'H{row}', "OFF", format5)

        # Now, we'll write the location data
        locations = xf203['col'].tolist()[1:]  # Assuming 'col' is a column in xf201 dataframe

        # Ensure locations match the number of worksheets
        if len(locations) != len(worksheets):
            raise ValueError(f"Number of locations ({len(locations)}) does not match number of worksheets ({len(worksheets)})")

        # Iterate over the names and locations and write the data
        for i, (name, location) in enumerate(zip(xf200['col'], locations)):
            worksheet = worksheets[i]  # Get the corresponding worksheet for this index

            for row in range(22, 24): 
                worksheet.write(f'B{row}', location, format5)
                worksheet.write(f'C{row}', location, format5)
                worksheet.write(f'D{row}', location, format5)
                worksheet.write(f'E{row}', location, format5)
                worksheet.write(f'F{row}', location, format5)
                worksheet.write(f'G{row}', "OFF", format5)
                worksheet.write(f'H{row}', "OFF", format5)

        # Now, we'll write the location data
        locations = xf204['col'].tolist()[1:]  # Assuming 'col' is a column in xf201 dataframe

        # Ensure locations match the number of worksheets
        if len(locations) != len(worksheets):
            raise ValueError(f"Number of locations ({len(locations)}) does not match number of worksheets ({len(worksheets)})")

        # Iterate over the names and locations and write the data
        for i, (name, location) in enumerate(zip(xf200['col'], locations)):
            worksheet = worksheets[i]  # Get the corresponding worksheet for this index

            for row in range(30, 32): 
                worksheet.write(f'B{row}', location, format5)
                worksheet.write(f'C{row}', location, format5)
                worksheet.write(f'D{row}', location, format5)
                worksheet.write(f'E{row}', location, format5)
                worksheet.write(f'F{row}', location, format5)
                worksheet.write(f'G{row}', "OFF", format5)
                worksheet.write(f'H{row}', "OFF", format5)


        # Close the workbook
        workbook.close()
        
        df = pd.read_csv('datesT.csv')
        dx = df

        df["dates"] = pd.to_datetime(df["dates"])
        df['dates'] = df['dates'].dt.strftime('%m/%d/%Y')

        df.to_csv('datesT.csv',index=False)

        df=pd.read_csv('PALIST.csv',dtype=str)
        #import io
        #output = io.StringIO()
        #df.to_csv(output, index=False)
        #output.seek(0)

        # Streamlit download button
        #st.download_button(
        #    label="Download CSV File",
        #    data=output.getvalue(),
        #    file_name="PALIST.csv",
        #    mime="text/csv"
        #)
 
	    
        df['text'] = df['providers'] + " - " + "[" + df['clinic'] + "]"
        df = df[['datecode','type','student','text','date','clinic']]


        mydict = {}
        with open('datesT.csv', mode='r')as inp:     #file is the objects you want to map. I want to map the IMP in this file to diagnosis.csv
            reader = csv.reader(inp)
            df1 = {rows[0]:rows[1] for rows in reader} 
        df['datecode'] = df.date.map(df1)

        df = df[~df['student'].isnull()]

        df = df.loc[df['student'] != "0"]

        df.to_excel('Source1.xlsx', index=False)
        #import io 
        #output = io.BytesIO()
        #with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        #    df.to_excel(writer, index=False, sheet_name='Sheet1')
        #    writer.close()
        #output.seek(0)
        #st.download_button(label="Download Excel File", data=output, file_name="Source1.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


        import openpyxl
        import numpy as np 

        column = "student"
        L=df[(column)].tolist() #Must Use Unique
        xf200 = pd.DataFrame({'col':L})
        xf200['i'] = xf200.index
        xf200['blank'] = ''
        
        wb = openpyxl.load_workbook('Source1.xlsx')
        ws = wb['Sheet1']
        wb1 = openpyxl.load_workbook('Main_Schedule_MS.xlsx')

        # Assuming xf200 is a pandas DataFrame loaded from somewhere, for example:
        # xf200 = pd.read_csv('your_dataframe.csv')  # Or load your DataFrame as needed

        # Iterate over the names in the xf200 DataFrame
        for name in xf200['col']:  # Assuming 'names' is the column with the names you want to look up
            # Check if the sheet with the name exists in the workbook
            if name in wb1.sheetnames:
                ws1 = wb1[name]  # Access the sheet directly since it exists
            else:
                print(f"Sheet for {name} not found, skipping.")
                continue  # Skip to the next name if the sheet does not exist

            for row in ws.iter_rows():
                if row[2].value == name:  # Check if the name matches (assuming name is in column 3)
                    # Handle T0 to T6
                    if row[0].value == "T0":
                        if row[1].value == "AM":
                            ws1.cell(row=6, column=2).value = row[3].value  # AM, T0
                        elif row[1].value == "PM":
                            ws1.cell(row=7, column=2).value = row[3].value  # PM, T0
                    elif row[0].value == "T1":
                        if row[1].value == "AM":
                            ws1.cell(row=6, column=3).value = row[3].value  # AM, T1
                        elif row[1].value == "PM":
                            ws1.cell(row=7, column=3).value = row[3].value  # PM, T1
                    elif row[0].value == "T2":
                        if row[1].value == "AM":
                            ws1.cell(row=6, column=4).value = row[3].value  # AM, T2
                        elif row[1].value == "PM":
                            ws1.cell(row=7, column=4).value = row[3].value  # PM, T2
                    elif row[0].value == "T3":
                        if row[1].value == "AM":
                            ws1.cell(row=6, column=5).value = row[3].value  # AM, T3
                        elif row[1].value == "PM":
                            ws1.cell(row=7, column=5).value = row[3].value  # PM, T3
                    elif row[0].value == "T4":
                        if row[1].value == "AM":
                            ws1.cell(row=6, column=6).value = row[3].value  # AM, T4
                        elif row[1].value == "PM":
                            ws1.cell(row=7, column=6).value = row[3].value  # PM, T4
                    elif row[0].value == "T5":
                        if row[1].value == "AM":
                            ws1.cell(row=6, column=7).value = row[3].value  # AM, T5
                        elif row[1].value == "PM":
                            ws1.cell(row=7, column=7).value = row[3].value  # PM, T5
                    elif row[0].value == "T6":
                        if row[1].value == "AM":
                            ws1.cell(row=6, column=8).value = row[3].value  # AM, T6
                        elif row[1].value == "PM":
                            ws1.cell(row=7, column=8).value = row[3].value  # PM, T6

                    # Handle T7 to T13
                    elif row[0].value == "T7":
                        if row[1].value == "AM":
                            ws1.cell(row=14, column=2).value = row[3].value  # AM, T7
                        elif row[1].value == "PM":
                            ws1.cell(row=15, column=2).value = row[3].value  # PM, T7
                    elif row[0].value == "T8":
                        if row[1].value == "AM":
                            ws1.cell(row=14, column=3).value = row[3].value  # AM, T8
                        elif row[1].value == "PM":
                            ws1.cell(row=15, column=3).value = row[3].value  # PM, T8
                    elif row[0].value == "T9":
                        if row[1].value == "AM":
                            ws1.cell(row=14, column=4).value = row[3].value  # AM, T9
                        elif row[1].value == "PM":
                            ws1.cell(row=15, column=4).value = row[3].value  # PM, T9
                    elif row[0].value == "T10":
                        if row[1].value == "AM":
                            ws1.cell(row=14, column=5).value = row[3].value  # AM, T10
                        elif row[1].value == "PM":
                            ws1.cell(row=15, column=5).value = row[3].value  # PM, T10
                    elif row[0].value == "T11":
                        if row[1].value == "AM":
                            ws1.cell(row=14, column=6).value = row[3].value  # AM, T11
                        elif row[1].value == "PM":
                            ws1.cell(row=15, column=6).value = row[3].value  # PM, T11
                    elif row[0].value == "T12":
                        if row[1].value == "AM":
                            ws1.cell(row=14, column=7).value = row[3].value  # AM, T12
                        elif row[1].value == "PM":
                            ws1.cell(row=15, column=7).value = row[3].value  # PM, T12
                    elif row[0].value == "T13":
                        if row[1].value == "AM":
                            ws1.cell(row=14, column=8).value = row[3].value  # AM, T13
                        elif row[1].value == "PM":
                            ws1.cell(row=15, column=8).value = row[3].value  # PM, T13

                    # Handle T14 to T20
                    elif row[0].value == "T14":
                        if row[1].value == "AM":
                            ws1.cell(row=22, column=2).value = row[3].value  # AM, T14
                        elif row[1].value == "PM":
                            ws1.cell(row=23, column=2).value = row[3].value  # PM, T14
                    elif row[0].value == "T15":
                        if row[1].value == "AM":
                            ws1.cell(row=22, column=3).value = row[3].value  # AM, T15
                        elif row[1].value == "PM":
                            ws1.cell(row=23, column=3).value = row[3].value  # PM, T15
                    elif row[0].value == "T16":
                        if row[1].value == "AM":
                            ws1.cell(row=22, column=4).value = row[3].value  # AM, T16
                        elif row[1].value == "PM":
                            ws1.cell(row=23, column=4).value = row[3].value  # PM, T16
                    elif row[0].value == "T17":
                        if row[1].value == "AM":
                            ws1.cell(row=22, column=5).value = row[3].value  # AM, T17
                        elif row[1].value == "PM":
                            ws1.cell(row=23, column=5).value = row[3].value  # PM, T17
                    elif row[0].value == "T18":
                        if row[1].value == "AM":
                            ws1.cell(row=22, column=6).value = row[3].value  # AM, T18
                        elif row[1].value == "PM":
                            ws1.cell(row=23, column=6).value = row[3].value  # PM, T18
                    elif row[0].value == "T19":
                        if row[1].value == "AM":
                            ws1.cell(row=22, column=7).value = row[3].value  # AM, T19
                        elif row[1].value == "PM":
                            ws1.cell(row=23, column=7).value = row[3].value  # PM, T19
                    elif row[0].value == "T20":
                        if row[1].value == "AM":
                            ws1.cell(row=22, column=8).value = row[3].value  # AM, T20
                        elif row[1].value == "PM":
                            ws1.cell(row=23, column=8).value = row[3].value  # PM, T20

                    # Handle T21 to T27
                    elif row[0].value == "T21":
                        if row[1].value == "AM":
                            ws1.cell(row=30, column=2).value = row[3].value  # AM, T21
                        elif row[1].value == "PM":
                            ws1.cell(row=31, column=2).value = row[3].value  # PM, T21
                    elif row[0].value == "T22":
                        if row[1].value == "AM":
                            ws1.cell(row=30, column=3).value = row[3].value  # AM, T22
                        elif row[1].value == "PM":
                            ws1.cell(row=31, column=3).value = row[3].value  # PM, T22
                    elif row[0].value == "T23":
                        if row[1].value == "AM":
                            ws1.cell(row=30, column=4).value = row[3].value  # AM, T23
                        elif row[1].value == "PM":
                            ws1.cell(row=31, column=4).value = row[3].value  # PM, T23
                    elif row[0].value == "T24":
                        if row[1].value == "AM":
                            ws1.cell(row=30, column=5).value = row[3].value  # AM, T24
                        elif row[1].value == "PM":
                            ws1.cell(row=31, column=5).value = row[3].value  # PM, T24
                    elif row[0].value == "T25":
                        if row[1].value == "AM":
                            ws1.cell(row=30, column=6).value = row[3].value  # AM, T25
                        elif row[1].value == "PM":
                            ws1.cell(row=31, column=6).value = row[3].value  # PM, T25
                    elif row[0].value == "T26":
                        if row[1].value == "AM":
                            ws1.cell(row=30, column=7).value = row[3].value  # AM, T26
                        elif row[1].value == "PM":
                            ws1.cell(row=31, column=7).value = row[3].value  # PM, T26
                    elif row[0].value == "T27":
                        if row[1].value == "AM":
                            ws1.cell(row=30, column=8).value = row[3].value  # AM, T27
                        elif row[1].value == "PM":
                            ws1.cell(row=31, column=8).value = row[3].value  # PM, T27

            for row in ws.iter_rows():
                if row[2].value == name:  # Check if the name matches (assuming name is in column 3)

                    # Handle T0 to T6
                    if row[0].value == "T0":
                        if row[1].value == "AM - ACUTES":
                            ws1.cell(row=6, column=2).value = row[3].value  # AM - ACUTES, T0
                        elif row[1].value == "PM - ACUTES":
                            ws1.cell(row=7, column=2).value = row[3].value  # PM - ACUTES, T0
                    elif row[0].value == "T1":
                        if row[1].value == "AM - ACUTES":
                            ws1.cell(row=6, column=3).value = row[3].value  # AM - ACUTES, T1
                        elif row[1].value == "PM - ACUTES":
                            ws1.cell(row=7, column=3).value = row[3].value  # PM - ACUTES, T1
                    elif row[0].value == "T2":
                        if row[1].value == "AM - ACUTES":
                            ws1.cell(row=6, column=4).value = row[3].value  # AM - ACUTES, T2
                        elif row[1].value == "PM - ACUTES":
                            ws1.cell(row=7, column=4).value = row[3].value  # PM - ACUTES, T2
                    elif row[0].value == "T3":
                        if row[1].value == "AM - ACUTES":
                            ws1.cell(row=6, column=5).value = row[3].value  # AM - ACUTES, T3
                        elif row[1].value == "PM - ACUTES":
                            ws1.cell(row=7, column=5).value = row[3].value  # PM - ACUTES, T3
                    elif row[0].value == "T4":
                        if row[1].value == "AM - ACUTES":
                            ws1.cell(row=6, column=6).value = row[3].value  # AM - ACUTES, T4
                        elif row[1].value == "PM - ACUTES":
                            ws1.cell(row=7, column=6).value = row[3].value  # PM - ACUTES, T4
                    elif row[0].value == "T5":
                        if row[1].value == "AM - ACUTES":
                            ws1.cell(row=6, column=7).value = row[3].value  # AM - ACUTES, T5
                        elif row[1].value == "PM - ACUTES":
                            ws1.cell(row=7, column=7).value = row[3].value  # PM - ACUTES, T5
                    elif row[0].value == "T6":
                        if row[1].value == "AM - ACUTES":
                            ws1.cell(row=6, column=8).value = row[3].value  # AM - ACUTES, T6
                        elif row[1].value == "PM - ACUTES":
                            ws1.cell(row=7, column=8).value = row[3].value  # PM - ACUTES, T6

                    # Handle T7 to T13
                    elif row[0].value == "T7":
                        if row[1].value == "AM - ACUTES":
                            ws1.cell(row=14, column=2).value = row[3].value  # AM - ACUTES, T7
                        elif row[1].value == "PM - ACUTES":
                            ws1.cell(row=15, column=2).value = row[3].value  # PM - ACUTES, T7
                    elif row[0].value == "T8":
                        if row[1].value == "AM - ACUTES":
                            ws1.cell(row=14, column=3).value = row[3].value  # AM - ACUTES, T8
                        elif row[1].value == "PM - ACUTES":
                            ws1.cell(row=15, column=3).value = row[3].value  # PM - ACUTES, T8
                    elif row[0].value == "T9":
                        if row[1].value == "AM - ACUTES":
                            ws1.cell(row=14, column=4).value = row[3].value  # AM - ACUTES, T9
                        elif row[1].value == "PM - ACUTES":
                            ws1.cell(row=15, column=4).value = row[3].value  # PM - ACUTES, T9
                    elif row[0].value == "T10":
                        if row[1].value == "AM - ACUTES":
                            ws1.cell(row=14, column=5).value = row[3].value  # AM - ACUTES, T10
                        elif row[1].value == "PM - ACUTES":
                            ws1.cell(row=15, column=5).value = row[3].value  # PM - ACUTES, T10
                    elif row[0].value == "T11":
                        if row[1].value == "AM - ACUTES":
                            ws1.cell(row=14, column=6).value = row[3].value  # AM - ACUTES, T11
                        elif row[1].value == "PM - ACUTES":
                            ws1.cell(row=15, column=6).value = row[3].value  # PM - ACUTES, T11
                    elif row[0].value == "T12":
                        if row[1].value == "AM - ACUTES":
                            ws1.cell(row=14, column=7).value = row[3].value  # AM - ACUTES, T12
                        elif row[1].value == "PM - ACUTES":
                            ws1.cell(row=15, column=7).value = row[3].value  # PM - ACUTES, T12
                    elif row[0].value == "T13":
                        if row[1].value == "AM - ACUTES":
                            ws1.cell(row=14, column=8).value = row[3].value  # AM - ACUTES, T13
                        elif row[1].value == "PM - ACUTES":
                            ws1.cell(row=15, column=8).value = row[3].value  # PM - ACUTES, T13

                    # Handle T14 to T20
                    elif row[0].value == "T14":
                        if row[1].value == "AM - ACUTES":
                            ws1.cell(row=22, column=2).value = row[3].value  # AM - ACUTES, T14
                        elif row[1].value == "PM - ACUTES":
                            ws1.cell(row=23, column=2).value = row[3].value  # PM - ACUTES, T14
                    elif row[0].value == "T15":
                        if row[1].value == "AM - ACUTES":
                            ws1.cell(row=22, column=3).value = row[3].value  # AM - ACUTES, T15
                        elif row[1].value == "PM - ACUTES":
                            ws1.cell(row=23, column=3).value = row[3].value  # PM - ACUTES, T15
                    elif row[0].value == "T16":
                        if row[1].value == "AM - ACUTES":
                            ws1.cell(row=22, column=4).value = row[3].value  # AM - ACUTES, T16
                        elif row[1].value == "PM - ACUTES":
                            ws1.cell(row=23, column=4).value = row[3].value  # PM - ACUTES, T16
                    elif row[0].value == "T17":
                        if row[1].value == "AM - ACUTES":
                            ws1.cell(row=22, column=5).value = row[3].value  # AM - ACUTES, T17
                        elif row[1].value == "PM - ACUTES":
                            ws1.cell(row=23, column=5).value = row[3].value  # PM - ACUTES, T17
                    elif row[0].value == "T18":
                        if row[1].value == "AM - ACUTES":
                            ws1.cell(row=22, column=6).value = row[3].value  # AM - ACUTES, T18
                        elif row[1].value == "PM - ACUTES":
                            ws1.cell(row=23, column=6).value = row[3].value  # PM - ACUTES, T18
                    elif row[0].value == "T19":
                        if row[1].value == "AM - ACUTES":
                            ws1.cell(row=22, column=7).value = row[3].value  # AM - ACUTES, T19
                        elif row[1].value == "PM - ACUTES":
                            ws1.cell(row=23, column=7).value = row[3].value  # PM - ACUTES, T19
                    elif row[0].value == "T20":
                        if row[1].value == "AM - ACUTES":
                            ws1.cell(row=22, column=8).value = row[3].value  # AM - ACUTES, T20
                        elif row[1].value == "PM - ACUTES":
                            ws1.cell(row=23, column=8).value = row[3].value  # PM - ACUTES, T20

                    # Handle T21 to T27
                    elif row[0].value == "T21":
                        if row[1].value == "AM - ACUTES":
                            ws1.cell(row=30, column=2).value = row[3].value  # AM - ACUTES, T21
                        elif row[1].value == "PM - ACUTES":
                            ws1.cell(row=31, column=2).value = row[3].value  # PM - ACUTES, T21
                    elif row[0].value == "T22":
                        if row[1].value == "AM - ACUTES":
                            ws1.cell(row=30, column=3).value = row[3].value  # AM - ACUTES, T22
                        elif row[1].value == "PM - ACUTES":
                            ws1.cell(row=31, column=3).value = row[3].value  # PM - ACUTES, T22
                    elif row[0].value == "T23":
                        if row[1].value == "AM - ACUTES":
                            ws1.cell(row=30, column=4).value = row[3].value  # AM - ACUTES, T23
                        elif row[1].value == "PM - ACUTES":
                            ws1.cell(row=31, column=4).value = row[3].value  # PM - ACUTES, T23
                    elif row[0].value == "T24":
                        if row[1].value == "AM - ACUTES":
                            ws1.cell(row=30, column=5).value = row[3].value  # AM - ACUTES, T24
                        elif row[1].value == "PM - ACUTES":
                            ws1.cell(row=31, column=5).value = row[3].value  # PM - ACUTES, T24
                    elif row[0].value == "T25":
                        if row[1].value == "AM - ACUTES":
                            ws1.cell(row=30, column=6).value = row[3].value  # AM - ACUTES, T25
                        elif row[1].value == "PM - ACUTES":
                            ws1.cell(row=31, column=6).value = row[3].value  # PM - ACUTES, T25
                    elif row[0].value == "T26":
                        if row[1].value == "AM - ACUTES":
                            ws1.cell(row=30, column=7).value = row[3].value  # AM - ACUTES, T26
                        elif row[1].value == "PM - ACUTES":
                            ws1.cell(row=31, column=7).value = row[3].value  # PM - ACUTES, T26
                    elif row[0].value == "T27":
                        if row[1].value == "AM - ACUTES":
                            ws1.cell(row=30, column=8).value = row[3].value  # AM - ACUTES, T27
                        elif row[1].value == "PM - ACUTES":
                            ws1.cell(row=31, column=8).value = row[3].value  # PM - ACUTES, T27

        # Save the modified workbook
        wb1.save('Main_Schedule_MS.xlsx')


            # Function to save the workbook to a BytesIO object
        def save_to_bytes(wb):
                output = BytesIO()
                wb.save(output)
                output.seek(0)  # Rewind the file pointer to the start
                return output

            # Prepare the workbook for download
        wb_bytes = save_to_bytes(wb1)
	#Creat	
        # Create a download button in Streamlit
        st.download_button(label="Download Medical Student Schedule",data=wb_bytes,file_name="Main_Schedule_MS.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    except Exception as e:
        st.error(f"Error processing the HOPE_DRIVE sheet: {e}")
