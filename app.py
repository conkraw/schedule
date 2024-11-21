import streamlit as st
import pandas as pd
import datetime
import requests
import subprocess
import sys
import os
import xlswriter 

# Initialize session state variables
if 'date_submitted' not in st.session_state:
    st.session_state['date_submitted'] = False
if 'files_uploaded' not in st.session_state:
    st.session_state['files_uploaded'] = False

# Display the title of the app
st.title('OPD Page')

# Allow the user to select "Create OPD"
page_option = st.selectbox('Select an option', ['Select Action', 'Create OPD'])

# If "Create OPD" is selected, show the date input form
if page_option == 'Create OPD':
    # Display instructions to the user
    st.write('Enter start date in m/d/yyyy format, no zeros in month or date (e.g., 7/6/2021):')

    # Create a text input field for the date
    x = st.text_input('Start Date')

    # Add a button to trigger the date parsing
    if st.button('Submit Date'):
        if x:
            try:
                # Try to parse the date entered by the user
                test_date = datetime.datetime.strptime(x, "%m/%d/%Y")
                
                # Store the date in session state
                st.session_state['date_submitted'] = True
                st.session_state['test_date'] = test_date.strftime('%m/%d/%Y')

                # Display the converted date
                st.write(f"Valid date entered: {test_date.strftime('%m/%d/%Y')}")
                st.write("You will now be redirected to the File Upload page.")
                st.rerun()  # Trigger the rerun to navigate to file upload

            except ValueError:
                # Display an error message if the date format is incorrect
                st.error('Invalid date format. Please enter the date in m/d/yyyy format.')
        else:
            # If no date is entered, display an error
            st.error('Please enter a date.')

# If the date is submitted, show the file upload page
if st.session_state['date_submitted'] and not st.session_state['files_uploaded']:
    st.write('Upload the following Excel files:')
    
    uploaded_files = {}
    uploaded_files['HOPE_DRIVE.xlsx'] = st.file_uploader('Upload HOPE_DRIVE.xlsx', type='xlsx')
    uploaded_files['ETOWN.xlsx'] = st.file_uploader('Upload ETOWN.xlsx', type='xlsx')
    uploaded_files['NYES.xlsx'] = st.file_uploader('Upload NYES.xlsx', type='xlsx')
    
    # Check if all files are uploaded
    if all(uploaded_files.values()):
        st.session_state['files_uploaded'] = True
        st.write("All files uploaded successfully!")
        st.write("You will now be redirected to the next page.")
        st.rerun()  # Trigger the rerun to proceed to the next step

# If files are uploaded, execute the next action
if st.session_state['files_uploaded']:
    st.write("Processing your files...")
    workbook = xlsxwriter.Workbook('OPD.xlsx')
    worksheet = workbook.add_worksheet('HOPE_DRIVE')
    worksheet2 = workbook.add_worksheet('ETOWN')
    worksheet3 = workbook.add_worksheet('NYES')
    worksheet4 = workbook.add_worksheet('EXTRA')
    worksheet5 = workbook.add_worksheet('MHS')
    
    format1 = workbook.add_format({'font_size':18,'bold': 1,'align': 'center','valign': 'vcenter','color':'black','bg_color':'#FEFFCC','border':1})
        
    worksheet.write(0, 0, 'Site:',format1)
    worksheet.write(0, 1, 'Hope Drive',format1)
                    
    worksheet2.write(0, 0, 'Site:',format1)
    worksheet2.write(0, 1, 'Elizabethtown',format1)
    
    worksheet3.write(0, 0, 'Site:',format1)
    worksheet3.write(0, 1, 'Nyes Road',format1)
    
    worksheet4.write(0, 0, 'Site:',format1)
    worksheet4.write(0, 1, 'EXTRA',format1)
    
    worksheet5.write(0, 0, 'Site:',format1)
    worksheet5.write(0, 1, 'MHS',format1)
    
    #Color Coding
    format4 = workbook.add_format({'font_size':12,'bold': 1,'align': 'center','valign': 'vcenter','color':'black','bg_color':'#8ccf6f','border':1})
    format4a = workbook.add_format({'font_size':12,'bold': 1,'align': 'center','valign': 'vcenter','color':'black','bg_color':'#9fc5e8','border':1})    
    format5 = workbook.add_format({'font_size':12,'bold': 1,'align': 'center','valign': 'vcenter','color':'black','bg_color':'#FEFFCC','border':1})
    format5a = workbook.add_format({'font_size':12,'bold': 1,'align': 'center','valign': 'vcenter','color':'black','bg_color':'#d0e9ff','border':1})
    format11 = workbook.add_format({'font_size':18,'bold': 1,'align': 'center','valign': 'vcenter','color':'black','bg_color':'#FEFFCC','border':1})
    #H codes
    formate = workbook.add_format({'font_size':12,'bold': 0,'align': 'center','valign': 'vcenter','color':'white','border':0})
    
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
    worksheets = [worksheet2, worksheet3, worksheet4, worksheet5]
    
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
            'color': 'black', 'bg_color': '#FFC7CE', 'border': 1
        })
        day_labels = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']
        start_rows = [2, 26, 50, 74] #[3, 27, 51, 75]
        for start_row in start_rows:
            for i, day in enumerate(day_labels):
                worksheet.write(start_row, 1 + i, day, format3)  # B=1, C=2, etc.
    
        # Set Date Formats
        format_date = workbook.add_format({
            'num_format': 'm/d/yyyy', 'font_size': 12, 'bold': 1, 'align': 'center', 'valign': 'vcenter',
            'color': 'black', 'bg_color': '#FFC7CE', 'border': 1
        })
    
        format_label = workbook.add_format({
            'font_size': 12, 'bold': 1, 'align': 'center', 'valign': 'vcenter',
            'color': 'black', 'bg_color': '#FFC7CE', 'border': 1
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
            'color': 'red', 'bg_color': '#FEFFCC', 'border': 1
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
        
    ####################################hopedrive###################################################################
    if uploaded_files['HOPE_DRIVE.xlsx']:
        df = pd.read_excel(uploaded_files['HOPE_DRIVE.xlsx'], dtype=str)
        #st.write("HOPE_DRIVE Data:")
        #st.dataframe(df)  # Show the dataframe
      
    #df=pd.read_excel('HOPE_DRIVE.xlsx',dtype=str)
    
    df.rename(columns={ df.columns[0]: "0" }, inplace = True)
    df.rename(columns={ df.columns[1]: "1" }, inplace = True)
    df.rename(columns={ df.columns[2]: "2" }, inplace = True)
    df.rename(columns={ df.columns[3]: "3" }, inplace = True)
    df.rename(columns={ df.columns[4]: "4" }, inplace = True)
    df.rename(columns={ df.columns[5]: "5" }, inplace = True)
    df.rename(columns={ df.columns[6]: "6" }, inplace = True)
    df.rename(columns={ df.columns[7]: "7" }, inplace = True)
    df.rename(columns={ df.columns[8]: "8" }, inplace = True)
    df.rename(columns={ df.columns[9]: "9" }, inplace = True)
    df.rename(columns={ df.columns[10]: "10" }, inplace = True)
    df.rename(columns={ df.columns[11]: "11" }, inplace = True)
    df.rename(columns={ df.columns[12]: "12" }, inplace = True)
    df.rename(columns={ df.columns[13]: "13" }, inplace = True)
    
    xf300 = pd.DataFrame({'no':['0','2','4','6','8','10','12','0','2','4','6','8','10','12','0','2','4','6','8','10','12','0','2','4','6','8','10','12']})
    
    xf300['no1'] = ['1','3','5','7','9','11','13','1','3','5','7','9','11','13','1','3','5','7','9','11','13','1','3','5','7','9','11','13']
    
    xf300['start']=['0','1','2','3','4','5','6','7','8','9','10','11','12','13','14','15','16','17','18','19','20','21','22','23','24','25','26','27']
    
    xf300['end'] = ['7','8','9','10','11','12','13','14','15','16','17','18','19','20','21','22','23','24','25','26','27','28','29','30','31','32','33','34']
    
    a = df.loc[df['0'] == day0].index[0]
    b = df.loc[df['0'] == day7].index[0]
    c = df.iloc[int(a)+1:int(b)]
    z0=c.loc[:, ('0','1')]
    z0['date']=day0
    z0.rename(columns={ df.columns[0]:"type"}, inplace = True)
    z0.rename(columns={ df.columns[1]:"provider"}, inplace = True)
    D0=z0[['date','type','provider']]
    D0=D0[:-1]
    
    a = df.loc[df['2'] == day1].index[0]
    b = df.loc[df['2'] == day8].index[0]
    c = df.iloc[int(a)+1:int(b)]
    z1=c.loc[:, ('2','3')]
    z1['date']=day1
    z1.rename(columns={ df.columns[2]:"type"}, inplace = True)
    z1.rename(columns={ df.columns[3]:"provider"}, inplace = True)
    D1=z1[['date','type','provider']]
    D1=D1[:-1]
    
    a = df.loc[df['4'] == day2].index[0]
    b = df.loc[df['4'] == day9].index[0]
    c = df.iloc[int(a)+1:int(b)]
    z2=c.loc[:, ('4','5')]
    z2['date']=day2
    z2.rename(columns={ df.columns[4]:"type"}, inplace = True)
    z2.rename(columns={ df.columns[5]:"provider"}, inplace = True)
    D2=z2[['date','type','provider']]
    D2=D2[:-1]
    
    a = df.loc[df['6'] == day3].index[0]
    b = df.loc[df['6'] == day10].index[0]
    c = df.iloc[int(a)+1:int(b)]
    z3=c.loc[:, ('6','7')]
    z3['date']=day3
    z3.rename(columns={ df.columns[6]:"type"}, inplace = True)
    z3.rename(columns={ df.columns[7]:"provider"}, inplace = True)
    D3=z3[['date','type','provider']]
    D3=D3[:-1]
    
    a = df.loc[df['8'] == day4].index[0]
    b = df.loc[df['8'] == day11].index[0]
    c = df.iloc[int(a)+1:int(b)]
    z4=c.loc[:, ('8','9')]
    z4['date']=day4
    z4.rename(columns={ df.columns[8]:"type"}, inplace = True)
    z4.rename(columns={ df.columns[9]:"provider"}, inplace = True)
    D4=z4[['date','type','provider']]
    D4=D4[:-1]
    
    a = df.loc[df['10'] == day5].index[0]
    b = df.loc[df['10'] == day12].index[0]
    c = df.iloc[int(a)+1:int(b)]
    z5=c.loc[:, ('10','11')]
    z5['date']=day5
    z5.rename(columns={ df.columns[10]:"type"}, inplace = True)
    z5.rename(columns={ df.columns[11]:"provider"}, inplace = True)
    D5=z5[['date','type','provider']]
    D5=D5[:-1]
    
    a = df.loc[df['12'] == day6].index[0]
    b = df.loc[df['12'] == day13].index[0]
    c = df.iloc[int(a)+1:int(b)]
    z6=c.loc[:, ('12','13')]
    z6['date']=day6
    z6.rename(columns={ df.columns[12]:"type"}, inplace = True)
    z6.rename(columns={ df.columns[13]:"provider"}, inplace = True)
    D6=z6[['date','type','provider']]
    D6=D6[:-1]
    
    a = df.loc[df['0'] == day7].index[0]
    b = df.loc[df['0'] == day14].index[0]
    c = df.iloc[int(a)+1:int(b)]
    z7=c.loc[:, ('0','1')]
    z7['date']=day7
    z7.rename(columns={ df.columns[0]:"type"}, inplace = True)
    z7.rename(columns={ df.columns[1]:"provider"}, inplace = True)
    D7=z7[['date','type','provider']]
    D7=D7[:-1]
    
    a = df.loc[df['2'] == day8].index[0]
    b = df.loc[df['2'] == day15].index[0]
    c = df.iloc[int(a)+1:int(b)]
    z8=c.loc[:, ('2','3')]
    z8['date']=day8
    z8.rename(columns={ df.columns[2]:"type"}, inplace = True)
    z8.rename(columns={ df.columns[3]:"provider"}, inplace = True)
    D8=z8[['date','type','provider']]
    D8=D8[:-1]
    
    a = df.loc[df['4'] == day9].index[0]
    b = df.loc[df['4'] == day16].index[0]
    c = df.iloc[int(a)+1:int(b)]
    z9=c.loc[:, ('4','5')]
    z9['date']=day9
    z9.rename(columns={ df.columns[4]:"type"}, inplace = True)
    z9.rename(columns={ df.columns[5]:"provider"}, inplace = True)
    D9=z9[['date','type','provider']]
    D9=D9[:-1]
    
    a = df.loc[df['6'] == day10].index[0]
    b = df.loc[df['6'] == day17].index[0]
    c = df.iloc[int(a)+1:int(b)]
    z10=c.loc[:, ('6','7')]
    z10['date']=day10
    z10.rename(columns={ df.columns[6]:"type"}, inplace = True)
    z10.rename(columns={ df.columns[7]:"provider"}, inplace = True)
    D10=z10[['date','type','provider']]
    D10=D10[:-1]
    
    a = df.loc[df['8'] == day11].index[0]
    b = df.loc[df['8'] == day18].index[0]
    c = df.iloc[int(a)+1:int(b)]
    z11=c.loc[:, ('8','9')]
    z11['date']=day11
    z11.rename(columns={ df.columns[8]:"type"}, inplace = True)
    z11.rename(columns={ df.columns[9]:"provider"}, inplace = True)
    D11=z11[['date','type','provider']]
    D11=D11[:-1]
    
    a = df.loc[df['10'] == day12].index[0]
    b = df.loc[df['10'] == day19].index[0]
    c = df.iloc[int(a)+1:int(b)]
    z12=c.loc[:, ('10','11')]
    z12['date']=day12
    z12.rename(columns={ df.columns[10]:"type"}, inplace = True)
    z12.rename(columns={ df.columns[11]:"provider"}, inplace = True)
    D12=z12[['date','type','provider']]
    D12=D12[:-1]
    
    a = df.loc[df['12'] == day13].index[0]
    b = df.loc[df['12'] == day20].index[0]
    c = df.iloc[int(a)+1:int(b)]
    z13=c.loc[:, ('12','13')]
    z13['date']=day13
    z13.rename(columns={ df.columns[12]:"type"}, inplace = True)
    z13.rename(columns={ df.columns[13]:"provider"}, inplace = True)
    D13=z13[['date','type','provider']]
    D13=D13[:-1]
    
    a = df.loc[df['0'] == day14].index[0]
    b = df.loc[df['0'] == day21].index[0]
    c = df.iloc[int(a)+1:int(b)]
    z14=c.loc[:, ('0','1')]
    z14['date']=day14
    z14.rename(columns={ df.columns[0]:"type"}, inplace = True)
    z14.rename(columns={ df.columns[1]:"provider"}, inplace = True)
    D14=z14[['date','type','provider']]
    D14=D14[:-1]
    
    a = df.loc[df['2'] == day15].index[0]
    b = df.loc[df['2'] == day22].index[0]
    c = df.iloc[int(a)+1:int(b)]
    z15=c.loc[:, ('2','3')]
    z15['date']=day15
    z15.rename(columns={ df.columns[2]:"type"}, inplace = True)
    z15.rename(columns={ df.columns[3]:"provider"}, inplace = True)
    D15=z15[['date','type','provider']]
    D15=D15[:-1]
    
    a = df.loc[df['4'] == day16].index[0]
    b = df.loc[df['4'] == day23].index[0]
    c = df.iloc[int(a)+1:int(b)]
    z16=c.loc[:, ('4','5')]
    z16['date']=day16
    z16.rename(columns={ df.columns[4]:"type"}, inplace = True)
    z16.rename(columns={ df.columns[5]:"provider"}, inplace = True)
    D16=z16[['date','type','provider']]
    D16=D16[:-1]
    
    a = df.loc[df['6'] == day17].index[0]
    b = df.loc[df['6'] == day24].index[0]
    c = df.iloc[int(a)+1:int(b)]
    z17=c.loc[:, ('6','7')]
    z17['date']=day17
    z17.rename(columns={ df.columns[6]:"type"}, inplace = True)
    z17.rename(columns={ df.columns[7]:"provider"}, inplace = True)
    D17=z17[['date','type','provider']]
    D17=D17[:-1]
    
    a = df.loc[df['8'] == day18].index[0]
    b = df.loc[df['8'] == day25].index[0]
    c = df.iloc[int(a)+1:int(b)]
    z18=c.loc[:, ('8','9')]
    z18['date']=day18
    z18.rename(columns={ df.columns[8]:"type"}, inplace = True)
    z18.rename(columns={ df.columns[9]:"provider"}, inplace = True)
    D18=z18[['date','type','provider']]
    D18=D18[:-1]
    
    a = df.loc[df['10'] == day19].index[0]
    b = df.loc[df['10'] == day26].index[0]
    c = df.iloc[int(a)+1:int(b)]
    z19=c.loc[:, ('10','11')]
    z19['date']=day19
    z19.rename(columns={ df.columns[10]:"type"}, inplace = True)
    z19.rename(columns={ df.columns[11]:"provider"}, inplace = True)
    D19=z19[['date','type','provider']]
    D19=D19[:-1]
    
    a = df.loc[df['12'] == day20].index[0]
    b = df.loc[df['12'] == day27].index[0]
    c = df.iloc[int(a)+1:int(b)]
    z20=c.loc[:, ('12','13')]
    z20['date']=day20
    z20.rename(columns={ df.columns[12]:"type"}, inplace = True)
    z20.rename(columns={ df.columns[13]:"provider"}, inplace = True)
    D20=z20[['date','type','provider']]
    D20=D20[:-1]
    
    a = df.loc[df['0'] == day21].index[0]
    b = df.loc[df['0'] == day28].index[0]
    c = df.iloc[int(a)+1:int(b)]
    z21=c.loc[:, ('0','1')]
    z21['date']=day21
    z21.rename(columns={ df.columns[0]:"type"}, inplace = True)
    z21.rename(columns={ df.columns[1]:"provider"}, inplace = True)
    D21=z21[['date','type','provider']]
    D21=D21[:-1]
    
    a = df.loc[df['2'] == day22].index[0]
    b = df.loc[df['2'] == day29].index[0]
    c = df.iloc[int(a)+1:int(b)]
    z22=c.loc[:, ('2','3')]
    z22['date']=day22
    z22.rename(columns={ df.columns[2]:"type"}, inplace = True)
    z22.rename(columns={ df.columns[3]:"provider"}, inplace = True)
    D22=z22[['date','type','provider']]
    D22=D22[:-1]
    
    a = df.loc[df['4'] == day23].index[0]
    b = df.loc[df['4'] == day30].index[0]
    c = df.iloc[int(a)+1:int(b)]
    z23=c.loc[:, ('4','5')]
    z23['date']=day23
    z23.rename(columns={ df.columns[4]:"type"}, inplace = True)
    z23.rename(columns={ df.columns[5]:"provider"}, inplace = True)
    D23=z23[['date','type','provider']]
    D23=D23[:-1]
    
    a = df.loc[df['6'] == day24].index[0]
    b = df.loc[df['6'] == day31].index[0]
    c = df.iloc[int(a)+1:int(b)]
    z24=c.loc[:, ('6','7')]
    z24['date']=day24
    z24.rename(columns={ df.columns[6]:"type"}, inplace = True)
    z24.rename(columns={ df.columns[7]:"provider"}, inplace = True)
    D24=z24[['date','type','provider']]
    D24=D24[:-1]
    
    a = df.loc[df['8'] == day25].index[0]
    b = df.loc[df['8'] == day32].index[0]
    c = df.iloc[int(a)+1:int(b)]
    z25=c.loc[:, ('8','9')]
    z25['date']=day25
    z25.rename(columns={ df.columns[8]:"type"}, inplace = True)
    z25.rename(columns={ df.columns[9]:"provider"}, inplace = True)
    D25=z25[['date','type','provider']]
    D25=D25[:-1]
    
    a = df.loc[df['10'] == day26].index[0]
    b = df.loc[df['10'] == day33].index[0]
    c = df.iloc[int(a)+1:int(b)]
    z26=c.loc[:, ('10','11')]
    z26['date']=day26
    z26.rename(columns={ df.columns[10]:"type"}, inplace = True)
    z26.rename(columns={ df.columns[11]:"provider"}, inplace = True)
    D26=z26[['date','type','provider']]
    D26=D26[:-1]
    
    a = df.loc[df['12'] == day27].index[0]
    b = df.loc[df['12'] == day34].index[0]
    c = df.iloc[int(a)+1:int(b)]
    z27=c.loc[:, ('12','13')]
    z27['date']=day27
    z27.rename(columns={ df.columns[12]:"type"}, inplace = True)
    z27.rename(columns={ df.columns[13]:"provider"}, inplace = True)
    D27=z27[['date','type','provider']]
    D27=D27[:-1]
    
    dfx=pd.DataFrame(columns=D0.columns)
    
    dfx=pd.concat([dfx,D0, D1, D2, D3, D4, D5, D6, D7, D8, D9, D10, D11, D12, D13, D14, D15, D16, D17, D18, D19, D20, D21, D22, D23, D24, D25, D26, D27])
    
    dfx['clinic'] = "HOPE_DRIVE"
    
    dfx.to_csv('hope.csv',index=False)
    hope=dfx.replace("Hope Drive AM Continuity", "AM - Continuity", regex=True)
    hope=hope.replace("Hope Drive PM Continuity", "PM - Continuity", regex=True)
    hope=hope.replace("Hope Drive\xa0AM Acute Precept ", "AM - ACUTES", regex=True)
    hope=hope.replace("Hope Drive PM Acute Precept", "PM - ACUTES", regex=True)
    hope=hope.replace("Hope Drive Weekend Continuity", "AM - Continuity", regex=True)
    hope=hope.replace("Hope Drive Weekend Acute 1", "AM - ACUTES", regex=True)
    hope=hope.replace("Hope Drive Weekend Acute 2", "AM - ACUTES", regex=True)
    
    hope.to_csv('hope.csv',index=False)
    
    ####################################################ETOWN#################################################
    if uploaded_files['ETOWN.xlsx']:
        df = pd.read_excel(uploaded_files['ETOWN.xlsx'], dtype=str)
        #st.write("ETOWN Data:")
        #st.dataframe(df_etown)  # Show the dataframe
      
    #df=pd.read_excel('ETOWN.xlsx',dtype=str)
    df.rename(columns={ df.columns[0]: "0" }, inplace = True)
    df.rename(columns={ df.columns[1]: "1" }, inplace = True)
    df.rename(columns={ df.columns[2]: "2" }, inplace = True)
    df.rename(columns={ df.columns[3]: "3" }, inplace = True)
    df.rename(columns={ df.columns[4]: "4" }, inplace = True)
    df.rename(columns={ df.columns[5]: "5" }, inplace = True)
    df.rename(columns={ df.columns[6]: "6" }, inplace = True)
    df.rename(columns={ df.columns[7]: "7" }, inplace = True)
    df.rename(columns={ df.columns[8]: "8" }, inplace = True)
    df.rename(columns={ df.columns[9]: "9" }, inplace = True)
    df.rename(columns={ df.columns[10]: "10" }, inplace = True)
    df.rename(columns={ df.columns[11]: "11" }, inplace = True)
    df.rename(columns={ df.columns[12]: "12" }, inplace = True)
    df.rename(columns={ df.columns[13]: "13" }, inplace = True)
    
    xf300 = pd.DataFrame({'no':['0','2','4','6','8','10','12','0','2','4','6','8','10','12','0','2','4','6','8','10','12','0','2','4','6','8','10','12']})
    
    xf300['no1'] = ['1','3','5','7','9','11','13','1','3','5','7','9','11','13','1','3','5','7','9','11','13','1','3','5','7','9','11','13']
    
    xf300['start']=['0','1','2','3','4','5','6','7','8','9','10','11','12','13','14','15','16','17','18','19','20','21','22','23','24','25','26','27']
    
    xf300['end'] = ['7','8','9','10','11','12','13','14','15','16','17','18','19','20','21','22','23','24','25','26','27','28','29','30','31','32','33','34']
    
    a = df.loc[df['0'] == day0].index[0]
    b = df.loc[df['0'] == day7].index[0]
    c = df.iloc[int(a)+1:int(b)]
    z0=c.loc[:, ('0','1')]
    z0['date']=day0
    z0.rename(columns={ df.columns[0]:"type"}, inplace = True)
    z0.rename(columns={ df.columns[1]:"provider"}, inplace = True)
    D0=z0[['date','type','provider']]
    D0=D0[:-1]
    
    a = df.loc[df['2'] == day1].index[0]
    b = df.loc[df['2'] == day8].index[0]
    c = df.iloc[int(a)+1:int(b)]
    z1=c.loc[:, ('2','3')]
    z1['date']=day1
    z1.rename(columns={ df.columns[2]:"type"}, inplace = True)
    z1.rename(columns={ df.columns[3]:"provider"}, inplace = True)
    D1=z1[['date','type','provider']]
    D1=D1[:-1]
    
    a = df.loc[df['4'] == day2].index[0]
    b = df.loc[df['4'] == day9].index[0]
    c = df.iloc[int(a)+1:int(b)]
    z2=c.loc[:, ('4','5')]
    z2['date']=day2
    z2.rename(columns={ df.columns[4]:"type"}, inplace = True)
    z2.rename(columns={ df.columns[5]:"provider"}, inplace = True)
    D2=z2[['date','type','provider']]
    D2=D2[:-1]
    
    a = df.loc[df['6'] == day3].index[0]
    b = df.loc[df['6'] == day10].index[0]
    c = df.iloc[int(a)+1:int(b)]
    z3=c.loc[:, ('6','7')]
    z3['date']=day3
    z3.rename(columns={ df.columns[6]:"type"}, inplace = True)
    z3.rename(columns={ df.columns[7]:"provider"}, inplace = True)
    D3=z3[['date','type','provider']]
    D3=D3[:-1]
    
    a = df.loc[df['8'] == day4].index[0]
    b = df.loc[df['8'] == day11].index[0]
    c = df.iloc[int(a)+1:int(b)]
    z4=c.loc[:, ('8','9')]
    z4['date']=day4
    z4.rename(columns={ df.columns[8]:"type"}, inplace = True)
    z4.rename(columns={ df.columns[9]:"provider"}, inplace = True)
    D4=z4[['date','type','provider']]
    D4=D4[:-1]
    
    a = df.loc[df['10'] == day5].index[0]
    b = df.loc[df['10'] == day12].index[0]
    c = df.iloc[int(a)+1:int(b)]
    z5=c.loc[:, ('10','11')]
    z5['date']=day5
    z5.rename(columns={ df.columns[10]:"type"}, inplace = True)
    z5.rename(columns={ df.columns[11]:"provider"}, inplace = True)
    D5=z5[['date','type','provider']]
    D5=D5[:-1]
    
    a = df.loc[df['12'] == day6].index[0]
    b = df.loc[df['12'] == day13].index[0]
    c = df.iloc[int(a)+1:int(b)]
    z6=c.loc[:, ('12','13')]
    z6['date']=day6
    z6.rename(columns={ df.columns[12]:"type"}, inplace = True)
    z6.rename(columns={ df.columns[13]:"provider"}, inplace = True)
    D6=z6[['date','type','provider']]
    D6=D6[:-1]
    
    a = df.loc[df['0'] == day7].index[0]
    b = df.loc[df['0'] == day14].index[0]
    c = df.iloc[int(a)+1:int(b)]
    z7=c.loc[:, ('0','1')]
    z7['date']=day7
    z7.rename(columns={ df.columns[0]:"type"}, inplace = True)
    z7.rename(columns={ df.columns[1]:"provider"}, inplace = True)
    D7=z7[['date','type','provider']]
    D7=D7[:-1]
    
    a = df.loc[df['2'] == day8].index[0]
    b = df.loc[df['2'] == day15].index[0]
    c = df.iloc[int(a)+1:int(b)]
    z8=c.loc[:, ('2','3')]
    z8['date']=day8
    z8.rename(columns={ df.columns[2]:"type"}, inplace = True)
    z8.rename(columns={ df.columns[3]:"provider"}, inplace = True)
    D8=z8[['date','type','provider']]
    D8=D8[:-1]
    
    a = df.loc[df['4'] == day9].index[0]
    b = df.loc[df['4'] == day16].index[0]
    c = df.iloc[int(a)+1:int(b)]
    z9=c.loc[:, ('4','5')]
    z9['date']=day9
    z9.rename(columns={ df.columns[4]:"type"}, inplace = True)
    z9.rename(columns={ df.columns[5]:"provider"}, inplace = True)
    D9=z9[['date','type','provider']]
    D9=D9[:-1]
    
    a = df.loc[df['6'] == day10].index[0]
    b = df.loc[df['6'] == day17].index[0]
    c = df.iloc[int(a)+1:int(b)]
    z10=c.loc[:, ('6','7')]
    z10['date']=day10
    z10.rename(columns={ df.columns[6]:"type"}, inplace = True)
    z10.rename(columns={ df.columns[7]:"provider"}, inplace = True)
    D10=z10[['date','type','provider']]
    D10=D10[:-1]
    
    a = df.loc[df['8'] == day11].index[0]
    b = df.loc[df['8'] == day18].index[0]
    c = df.iloc[int(a)+1:int(b)]
    z11=c.loc[:, ('8','9')]
    z11['date']=day11
    z11.rename(columns={ df.columns[8]:"type"}, inplace = True)
    z11.rename(columns={ df.columns[9]:"provider"}, inplace = True)
    D11=z11[['date','type','provider']]
    D11=D11[:-1]
    
    a = df.loc[df['10'] == day12].index[0]
    b = df.loc[df['10'] == day19].index[0]
    c = df.iloc[int(a)+1:int(b)]
    z12=c.loc[:, ('10','11')]
    z12['date']=day12
    z12.rename(columns={ df.columns[10]:"type"}, inplace = True)
    z12.rename(columns={ df.columns[11]:"provider"}, inplace = True)
    D12=z12[['date','type','provider']]
    D12=D12[:-1]
    
    a = df.loc[df['12'] == day13].index[0]
    b = df.loc[df['12'] == day20].index[0]
    c = df.iloc[int(a)+1:int(b)]
    z13=c.loc[:, ('12','13')]
    z13['date']=day13
    z13.rename(columns={ df.columns[12]:"type"}, inplace = True)
    z13.rename(columns={ df.columns[13]:"provider"}, inplace = True)
    D13=z13[['date','type','provider']]
    D13=D13[:-1]
    
    a = df.loc[df['0'] == day14].index[0]
    b = df.loc[df['0'] == day21].index[0]
    c = df.iloc[int(a)+1:int(b)]
    z14=c.loc[:, ('0','1')]
    z14['date']=day14
    z14.rename(columns={ df.columns[0]:"type"}, inplace = True)
    z14.rename(columns={ df.columns[1]:"provider"}, inplace = True)
    D14=z14[['date','type','provider']]
    D14=D14[:-1]
    
    a = df.loc[df['2'] == day15].index[0]
    b = df.loc[df['2'] == day22].index[0]
    c = df.iloc[int(a)+1:int(b)]
    z15=c.loc[:, ('2','3')]
    z15['date']=day15
    z15.rename(columns={ df.columns[2]:"type"}, inplace = True)
    z15.rename(columns={ df.columns[3]:"provider"}, inplace = True)
    D15=z15[['date','type','provider']]
    D15=D15[:-1]
    
    a = df.loc[df['4'] == day16].index[0]
    b = df.loc[df['4'] == day23].index[0]
    c = df.iloc[int(a)+1:int(b)]
    z16=c.loc[:, ('4','5')]
    z16['date']=day16
    z16.rename(columns={ df.columns[4]:"type"}, inplace = True)
    z16.rename(columns={ df.columns[5]:"provider"}, inplace = True)
    D16=z16[['date','type','provider']]
    D16=D16[:-1]
    
    a = df.loc[df['6'] == day17].index[0]
    b = df.loc[df['6'] == day24].index[0]
    c = df.iloc[int(a)+1:int(b)]
    z17=c.loc[:, ('6','7')]
    z17['date']=day17
    z17.rename(columns={ df.columns[6]:"type"}, inplace = True)
    z17.rename(columns={ df.columns[7]:"provider"}, inplace = True)
    D17=z17[['date','type','provider']]
    D17=D17[:-1]
    
    a = df.loc[df['8'] == day18].index[0]
    b = df.loc[df['8'] == day25].index[0]
    c = df.iloc[int(a)+1:int(b)]
    z18=c.loc[:, ('8','9')]
    z18['date']=day18
    z18.rename(columns={ df.columns[8]:"type"}, inplace = True)
    z18.rename(columns={ df.columns[9]:"provider"}, inplace = True)
    D18=z18[['date','type','provider']]
    D18=D18[:-1]
    
    a = df.loc[df['10'] == day19].index[0]
    b = df.loc[df['10'] == day26].index[0]
    c = df.iloc[int(a)+1:int(b)]
    z19=c.loc[:, ('10','11')]
    z19['date']=day19
    z19.rename(columns={ df.columns[10]:"type"}, inplace = True)
    z19.rename(columns={ df.columns[11]:"provider"}, inplace = True)
    D19=z19[['date','type','provider']]
    D19=D19[:-1]
    
    a = df.loc[df['12'] == day20].index[0]
    b = df.loc[df['12'] == day27].index[0]
    c = df.iloc[int(a)+1:int(b)]
    z20=c.loc[:, ('12','13')]
    z20['date']=day20
    z20.rename(columns={ df.columns[12]:"type"}, inplace = True)
    z20.rename(columns={ df.columns[13]:"provider"}, inplace = True)
    D20=z20[['date','type','provider']]
    D20=D20[:-1]
    
    a = df.loc[df['0'] == day21].index[0]
    b = df.loc[df['0'] == day28].index[0]
    c = df.iloc[int(a)+1:int(b)]
    z21=c.loc[:, ('0','1')]
    z21['date']=day21
    z21.rename(columns={ df.columns[0]:"type"}, inplace = True)
    z21.rename(columns={ df.columns[1]:"provider"}, inplace = True)
    D21=z21[['date','type','provider']]
    D21=D21[:-1]
    
    a = df.loc[df['2'] == day22].index[0]
    b = df.loc[df['2'] == day29].index[0]
    c = df.iloc[int(a)+1:int(b)]
    z22=c.loc[:, ('2','3')]
    z22['date']=day22
    z22.rename(columns={ df.columns[2]:"type"}, inplace = True)
    z22.rename(columns={ df.columns[3]:"provider"}, inplace = True)
    D22=z22[['date','type','provider']]
    D22=D22[:-1]
    
    a = df.loc[df['4'] == day23].index[0]
    b = df.loc[df['4'] == day30].index[0]
    c = df.iloc[int(a)+1:int(b)]
    z23=c.loc[:, ('4','5')]
    z23['date']=day23
    z23.rename(columns={ df.columns[4]:"type"}, inplace = True)
    z23.rename(columns={ df.columns[5]:"provider"}, inplace = True)
    D23=z23[['date','type','provider']]
    D23=D23[:-1]
    
    a = df.loc[df['6'] == day24].index[0]
    b = df.loc[df['6'] == day31].index[0]
    c = df.iloc[int(a)+1:int(b)]
    z24=c.loc[:, ('6','7')]
    z24['date']=day24
    z24.rename(columns={ df.columns[6]:"type"}, inplace = True)
    z24.rename(columns={ df.columns[7]:"provider"}, inplace = True)
    D24=z24[['date','type','provider']]
    D24=D24[:-1]
    
    a = df.loc[df['8'] == day25].index[0]
    b = df.loc[df['8'] == day32].index[0]
    c = df.iloc[int(a)+1:int(b)]
    z25=c.loc[:, ('8','9')]
    z25['date']=day25
    z25.rename(columns={ df.columns[8]:"type"}, inplace = True)
    z25.rename(columns={ df.columns[9]:"provider"}, inplace = True)
    D25=z25[['date','type','provider']]
    D25=D25[:-1]
    
    a = df.loc[df['10'] == day26].index[0]
    b = df.loc[df['10'] == day33].index[0]
    c = df.iloc[int(a)+1:int(b)]
    z26=c.loc[:, ('10','11')]
    z26['date']=day26
    z26.rename(columns={ df.columns[10]:"type"}, inplace = True)
    z26.rename(columns={ df.columns[11]:"provider"}, inplace = True)
    D26=z26[['date','type','provider']]
    D26=D26[:-1]
    
    a = df.loc[df['12'] == day27].index[0]
    b = df.loc[df['12'] == day34].index[0]
    c = df.iloc[int(a)+1:int(b)]
    z27=c.loc[:, ('12','13')]
    z27['date']=day27
    z27.rename(columns={ df.columns[12]:"type"}, inplace = True)
    z27.rename(columns={ df.columns[13]:"provider"}, inplace = True)
    D27=z27[['date','type','provider']]
    D27=D27[:-1]
    
    
    dfx=pd.DataFrame(columns=D0.columns)
    
    dfx=pd.concat([dfx,D0, D1, D2, D3, D4, D5, D6, D7, D8, D9, D10, D11, D12, D13, D14, D15, D16, D17, D18, D19, D20, D21, D22, D23, D24, D25, D26, D27])
    
    dfx['clinic'] = "ETOWN"
    
    
    dfx.to_csv('etown.csv',index=False)
    ETOWN=dfx.replace("Etown AM Continuity", "AM - Continuity", regex=True)
    ETOWN=ETOWN.replace("Etown PM Continuity", "PM - Continuity", regex=True)
    ETOWN.to_csv('etown.csv',index=False)
    
    #############################################################NYES################################################
    if uploaded_files['NYES.xlsx']:
        df = pd.read_excel(uploaded_files['NYES.xlsx'], dtype=str)
        #st.write("NYES Data:")
        #st.dataframe(df_nyes)
    #df=pd.read_excel('NYES.xlsx',dtype=str)
    
    df.rename(columns={ df.columns[0]: "0" }, inplace = True)
    df.rename(columns={ df.columns[1]: "1" }, inplace = True)
    df.rename(columns={ df.columns[2]: "2" }, inplace = True)
    df.rename(columns={ df.columns[3]: "3" }, inplace = True)
    df.rename(columns={ df.columns[4]: "4" }, inplace = True)
    df.rename(columns={ df.columns[5]: "5" }, inplace = True)
    df.rename(columns={ df.columns[6]: "6" }, inplace = True)
    df.rename(columns={ df.columns[7]: "7" }, inplace = True)
    df.rename(columns={ df.columns[8]: "8" }, inplace = True)
    df.rename(columns={ df.columns[9]: "9" }, inplace = True)
    df.rename(columns={ df.columns[10]: "10" }, inplace = True)
    df.rename(columns={ df.columns[11]: "11" }, inplace = True)
    df.rename(columns={ df.columns[12]: "12" }, inplace = True)
    df.rename(columns={ df.columns[13]: "13" }, inplace = True)
    
    xf300 = pd.DataFrame({'no':['0','2','4','6','8','10','12','0','2','4','6','8','10','12','0','2','4','6','8','10','12','0','2','4','6','8','10','12']})
    
    xf300['no1'] = ['1','3','5','7','9','11','13','1','3','5','7','9','11','13','1','3','5','7','9','11','13','1','3','5','7','9','11','13']
    
    xf300['start']=['0','1','2','3','4','5','6','7','8','9','10','11','12','13','14','15','16','17','18','19','20','21','22','23','24','25','26','27']
    
    xf300['end'] = ['7','8','9','10','11','12','13','14','15','16','17','18','19','20','21','22','23','24','25','26','27','28','29','30','31','32','33','34']
    
    a = df.loc[df['0'] == day0].index[0]
    b = df.loc[df['0'] == day7].index[0]
    c = df.iloc[int(a)+1:int(b)]
    z0=c.loc[:, ('0','1')]
    z0['date']=day0
    z0.rename(columns={ df.columns[0]:"type"}, inplace = True)
    z0.rename(columns={ df.columns[1]:"provider"}, inplace = True)
    D0=z0[['date','type','provider']]
    D0=D0[:-1]
    
    a = df.loc[df['2'] == day1].index[0]
    b = df.loc[df['2'] == day8].index[0]
    c = df.iloc[int(a)+1:int(b)]
    z1=c.loc[:, ('2','3')]
    z1['date']=day1
    z1.rename(columns={ df.columns[2]:"type"}, inplace = True)
    z1.rename(columns={ df.columns[3]:"provider"}, inplace = True)
    D1=z1[['date','type','provider']]
    D1=D1[:-1]
    
    a = df.loc[df['4'] == day2].index[0]
    b = df.loc[df['4'] == day9].index[0]
    c = df.iloc[int(a)+1:int(b)]
    z2=c.loc[:, ('4','5')]
    z2['date']=day2
    z2.rename(columns={ df.columns[4]:"type"}, inplace = True)
    z2.rename(columns={ df.columns[5]:"provider"}, inplace = True)
    D2=z2[['date','type','provider']]
    D2=D2[:-1]
    
    a = df.loc[df['6'] == day3].index[0]
    b = df.loc[df['6'] == day10].index[0]
    c = df.iloc[int(a)+1:int(b)]
    z3=c.loc[:, ('6','7')]
    z3['date']=day3
    z3.rename(columns={ df.columns[6]:"type"}, inplace = True)
    z3.rename(columns={ df.columns[7]:"provider"}, inplace = True)
    D3=z3[['date','type','provider']]
    D3=D3[:-1]
    
    a = df.loc[df['8'] == day4].index[0]
    b = df.loc[df['8'] == day11].index[0]
    c = df.iloc[int(a)+1:int(b)]
    z4=c.loc[:, ('8','9')]
    z4['date']=day4
    z4.rename(columns={ df.columns[8]:"type"}, inplace = True)
    z4.rename(columns={ df.columns[9]:"provider"}, inplace = True)
    D4=z4[['date','type','provider']]
    D4=D4[:-1]
    
    a = df.loc[df['10'] == day5].index[0]
    b = df.loc[df['10'] == day12].index[0]
    c = df.iloc[int(a)+1:int(b)]
    z5=c.loc[:, ('10','11')]
    z5['date']=day5
    z5.rename(columns={ df.columns[10]:"type"}, inplace = True)
    z5.rename(columns={ df.columns[11]:"provider"}, inplace = True)
    D5=z5[['date','type','provider']]
    D5=D5[:-1]
    
    a = df.loc[df['12'] == day6].index[0]
    b = df.loc[df['12'] == day13].index[0]
    c = df.iloc[int(a)+1:int(b)]
    z6=c.loc[:, ('12','13')]
    z6['date']=day6
    z6.rename(columns={ df.columns[12]:"type"}, inplace = True)
    z6.rename(columns={ df.columns[13]:"provider"}, inplace = True)
    D6=z6[['date','type','provider']]
    D6=D6[:-1]
    
    a = df.loc[df['0'] == day7].index[0]
    b = df.loc[df['0'] == day14].index[0]
    c = df.iloc[int(a)+1:int(b)]
    z7=c.loc[:, ('0','1')]
    z7['date']=day7
    z7.rename(columns={ df.columns[0]:"type"}, inplace = True)
    z7.rename(columns={ df.columns[1]:"provider"}, inplace = True)
    D7=z7[['date','type','provider']]
    D7=D7[:-1]
    
    a = df.loc[df['2'] == day8].index[0]
    b = df.loc[df['2'] == day15].index[0]
    c = df.iloc[int(a)+1:int(b)]
    z8=c.loc[:, ('2','3')]
    z8['date']=day8
    z8.rename(columns={ df.columns[2]:"type"}, inplace = True)
    z8.rename(columns={ df.columns[3]:"provider"}, inplace = True)
    D8=z8[['date','type','provider']]
    D8=D8[:-1]
    
    a = df.loc[df['4'] == day9].index[0]
    b = df.loc[df['4'] == day16].index[0]
    c = df.iloc[int(a)+1:int(b)]
    z9=c.loc[:, ('4','5')]
    z9['date']=day9
    z9.rename(columns={ df.columns[4]:"type"}, inplace = True)
    z9.rename(columns={ df.columns[5]:"provider"}, inplace = True)
    D9=z9[['date','type','provider']]
    D9=D9[:-1]
    
    a = df.loc[df['6'] == day10].index[0]
    b = df.loc[df['6'] == day17].index[0]
    c = df.iloc[int(a)+1:int(b)]
    z10=c.loc[:, ('6','7')]
    z10['date']=day10
    z10.rename(columns={ df.columns[6]:"type"}, inplace = True)
    z10.rename(columns={ df.columns[7]:"provider"}, inplace = True)
    D10=z10[['date','type','provider']]
    D10=D10[:-1]
    
    a = df.loc[df['8'] == day11].index[0]
    b = df.loc[df['8'] == day18].index[0]
    c = df.iloc[int(a)+1:int(b)]
    z11=c.loc[:, ('8','9')]
    z11['date']=day11
    z11.rename(columns={ df.columns[8]:"type"}, inplace = True)
    z11.rename(columns={ df.columns[9]:"provider"}, inplace = True)
    D11=z11[['date','type','provider']]
    D11=D11[:-1]
    
    a = df.loc[df['10'] == day12].index[0]
    b = df.loc[df['10'] == day19].index[0]
    c = df.iloc[int(a)+1:int(b)]
    z12=c.loc[:, ('10','11')]
    z12['date']=day12
    z12.rename(columns={ df.columns[10]:"type"}, inplace = True)
    z12.rename(columns={ df.columns[11]:"provider"}, inplace = True)
    D12=z12[['date','type','provider']]
    D12=D12[:-1]
    
    a = df.loc[df['12'] == day13].index[0]
    b = df.loc[df['12'] == day20].index[0]
    c = df.iloc[int(a)+1:int(b)]
    z13=c.loc[:, ('12','13')]
    z13['date']=day13
    z13.rename(columns={ df.columns[12]:"type"}, inplace = True)
    z13.rename(columns={ df.columns[13]:"provider"}, inplace = True)
    D13=z13[['date','type','provider']]
    D13=D13[:-1]
    
    a = df.loc[df['0'] == day14].index[0]
    b = df.loc[df['0'] == day21].index[0]
    c = df.iloc[int(a)+1:int(b)]
    z14=c.loc[:, ('0','1')]
    z14['date']=day14
    z14.rename(columns={ df.columns[0]:"type"}, inplace = True)
    z14.rename(columns={ df.columns[1]:"provider"}, inplace = True)
    D14=z14[['date','type','provider']]
    D14=D14[:-1]
    
    a = df.loc[df['2'] == day15].index[0]
    b = df.loc[df['2'] == day22].index[0]
    c = df.iloc[int(a)+1:int(b)]
    z15=c.loc[:, ('2','3')]
    z15['date']=day15
    z15.rename(columns={ df.columns[2]:"type"}, inplace = True)
    z15.rename(columns={ df.columns[3]:"provider"}, inplace = True)
    D15=z15[['date','type','provider']]
    D15=D15[:-1]
    
    a = df.loc[df['4'] == day16].index[0]
    b = df.loc[df['4'] == day23].index[0]
    c = df.iloc[int(a)+1:int(b)]
    z16=c.loc[:, ('4','5')]
    z16['date']=day16
    z16.rename(columns={ df.columns[4]:"type"}, inplace = True)
    z16.rename(columns={ df.columns[5]:"provider"}, inplace = True)
    D16=z16[['date','type','provider']]
    D16=D16[:-1]
    
    a = df.loc[df['6'] == day17].index[0]
    b = df.loc[df['6'] == day24].index[0]
    c = df.iloc[int(a)+1:int(b)]
    z17=c.loc[:, ('6','7')]
    z17['date']=day17
    z17.rename(columns={ df.columns[6]:"type"}, inplace = True)
    z17.rename(columns={ df.columns[7]:"provider"}, inplace = True)
    D17=z17[['date','type','provider']]
    D17=D17[:-1]
    
    a = df.loc[df['8'] == day18].index[0]
    b = df.loc[df['8'] == day25].index[0]
    c = df.iloc[int(a)+1:int(b)]
    z18=c.loc[:, ('8','9')]
    z18['date']=day18
    z18.rename(columns={ df.columns[8]:"type"}, inplace = True)
    z18.rename(columns={ df.columns[9]:"provider"}, inplace = True)
    D18=z18[['date','type','provider']]
    D18=D18[:-1]
    
    a = df.loc[df['10'] == day19].index[0]
    b = df.loc[df['10'] == day26].index[0]
    c = df.iloc[int(a)+1:int(b)]
    z19=c.loc[:, ('10','11')]
    z19['date']=day19
    z19.rename(columns={ df.columns[10]:"type"}, inplace = True)
    z19.rename(columns={ df.columns[11]:"provider"}, inplace = True)
    D19=z19[['date','type','provider']]
    D19=D19[:-1]
    
    a = df.loc[df['12'] == day20].index[0]
    b = df.loc[df['12'] == day27].index[0]
    c = df.iloc[int(a)+1:int(b)]
    z20=c.loc[:, ('12','13')]
    z20['date']=day20
    z20.rename(columns={ df.columns[12]:"type"}, inplace = True)
    z20.rename(columns={ df.columns[13]:"provider"}, inplace = True)
    D20=z20[['date','type','provider']]
    D20=D20[:-1]
    
    a = df.loc[df['0'] == day21].index[0]
    b = df.loc[df['0'] == day28].index[0]
    c = df.iloc[int(a)+1:int(b)]
    z21=c.loc[:, ('0','1')]
    z21['date']=day21
    z21.rename(columns={ df.columns[0]:"type"}, inplace = True)
    z21.rename(columns={ df.columns[1]:"provider"}, inplace = True)
    D21=z21[['date','type','provider']]
    D21=D21[:-1]
    
    a = df.loc[df['2'] == day22].index[0]
    b = df.loc[df['2'] == day29].index[0]
    c = df.iloc[int(a)+1:int(b)]
    z22=c.loc[:, ('2','3')]
    z22['date']=day22
    z22.rename(columns={ df.columns[2]:"type"}, inplace = True)
    z22.rename(columns={ df.columns[3]:"provider"}, inplace = True)
    D22=z22[['date','type','provider']]
    D22=D22[:-1]
    
    a = df.loc[df['4'] == day23].index[0]
    b = df.loc[df['4'] == day30].index[0]
    c = df.iloc[int(a)+1:int(b)]
    z23=c.loc[:, ('4','5')]
    z23['date']=day23
    z23.rename(columns={ df.columns[4]:"type"}, inplace = True)
    z23.rename(columns={ df.columns[5]:"provider"}, inplace = True)
    D23=z23[['date','type','provider']]
    D23=D23[:-1]
    
    a = df.loc[df['6'] == day24].index[0]
    b = df.loc[df['6'] == day31].index[0]
    c = df.iloc[int(a)+1:int(b)]
    z24=c.loc[:, ('6','7')]
    z24['date']=day24
    z24.rename(columns={ df.columns[6]:"type"}, inplace = True)
    z24.rename(columns={ df.columns[7]:"provider"}, inplace = True)
    D24=z24[['date','type','provider']]
    D24=D24[:-1]
    
    a = df.loc[df['8'] == day25].index[0]
    b = df.loc[df['8'] == day32].index[0]
    c = df.iloc[int(a)+1:int(b)]
    z25=c.loc[:, ('8','9')]
    z25['date']=day25
    z25.rename(columns={ df.columns[8]:"type"}, inplace = True)
    z25.rename(columns={ df.columns[9]:"provider"}, inplace = True)
    D25=z25[['date','type','provider']]
    D25=D25[:-1]
    
    a = df.loc[df['10'] == day26].index[0]
    b = df.loc[df['10'] == day33].index[0]
    c = df.iloc[int(a)+1:int(b)]
    z26=c.loc[:, ('10','11')]
    z26['date']=day26
    z26.rename(columns={ df.columns[10]:"type"}, inplace = True)
    z26.rename(columns={ df.columns[11]:"provider"}, inplace = True)
    D26=z26[['date','type','provider']]
    D26=D26[:-1]
    
    a = df.loc[df['12'] == day27].index[0]
    b = df.loc[df['12'] == day34].index[0]
    c = df.iloc[int(a)+1:int(b)]
    z27=c.loc[:, ('12','13')]
    z27['date']=day27
    z27.rename(columns={ df.columns[12]:"type"}, inplace = True)
    z27.rename(columns={ df.columns[13]:"provider"}, inplace = True)
    D27=z27[['date','type','provider']]
    D27=D27[:-1]
    
    dfx=pd.DataFrame(columns=D0.columns)
    
    dfx=pd.concat([dfx,D0, D1, D2, D3, D4, D5, D6, D7, D8, D9, D10, D11, D12, D13, D14, D15, D16, D17, D18, D19, D20, D21, D22, D23, D24, D25, D26, D27])
    
    dfx['clinic'] = "NYES"
    
    dfx.to_csv('nyes.csv',index=False)
    NYES=dfx.replace("Nyes Rd AM Continuity", "AM - Continuity", regex=True)
    NYES=NYES.replace("Nyes Rd PM Continuity", "PM - Continuity", regex=True)
    NYES.to_csv('nyes.csv',index=False)
    
    #############################################################################################################
    
    NYES['H'] = "H"
    NYEi = NYES[(NYES['type'] == 'AM - Continuity ')]
    NYEi['count'] = NYEi.groupby(['date'])['provider'].cumcount() + 0
    NYEi['class'] = "H" + NYEi['count'].astype(str)
    NYEi = NYEi.loc[:, ('date','type','provider','clinic','class')]
    NYEi.to_csv('1.csv',index=False)
    
    NYES['H'] = "H"
    NYEii = NYES[(NYES['type'] == 'PM - Continuity ')]
    NYEii['count'] = NYEii.groupby(['date'])['provider'].cumcount() + 10
    NYEii['class'] = "H" + NYEii['count'].astype(str)
    NYEii = NYEii.loc[:, ('date','type','provider','clinic','class')]
    NYEii.to_csv('2.csv',index=False)
    
    ETOWN['H'] = "H"
    ETOWNi = ETOWN[(ETOWN['type'] == 'AM - Continuity ')]
    ETOWNi['count'] = ETOWNi.groupby(['date'])['provider'].cumcount() + 0
    ETOWNi['class'] = "H" + ETOWNi['count'].astype(str)
    ETOWNi = ETOWNi.loc[:, ('date','type','provider','clinic','class')]
    ETOWNi.to_csv('3.csv',index=False)
    
    ETOWN['H'] = "H"
    ETOWNii = ETOWN[(ETOWN['type'] == 'PM - Continuity ')]
    ETOWNii['count'] = ETOWNii.groupby(['date'])['provider'].cumcount() + 10
    ETOWNii['class'] = "H" + ETOWNii['count'].astype(str)
    ETOWNii = ETOWNii.loc[:, ('date','type','provider','clinic','class')]
    ETOWNii.to_csv('4.csv',index=False)
    
    hope['class'] = "H"  # Top of Column
    hopeiii = hope[(hope['type'] == 'AM - ACUTES')]
    
    # Group by 'date' and count occurrences
    hopeiii['count'] = hopeiii.groupby(['date'])['provider'].cumcount()
    
    # Reserve H0 and H1 for the first two occurrences, subsequent start from H2
    hopeiii['class'] = hopeiii['count'].apply(
        lambda count: "H0" if count == 0 else ("H1" if count == 1 else "H" + str(count + 2))
    )
    
    hopeiii = hopeiii.loc[:, ('date', 'type', 'provider', 'clinic', 'class')]
    hopeiii.to_csv('7.csv', index=False)
    
    hope['class'] = "H"  # Top of Column
    hopeiii = hope[(hope['type'] == 'AM - ACUTES ')]
    
    # Group by 'date' and count occurrences
    hopeiii['count'] = hopeiii.groupby(['date'])['provider'].cumcount()
    
    # Reserve H0 and H1 for the first two occurrences, subsequent start from H2
    hopeiii['class'] = hopeiii['count'].apply(
        lambda count: "H0" if count == 0 else ("H1" if count == 1 else "H" + str(count + 2))
    )
    
    hopeiii = hopeiii.loc[:, ('date', 'type', 'provider', 'clinic', 'class')]
    hopeiii.to_csv('8.csv', index=False)
    
    ######STARTS AT H10 (PM - ACUTES)
    hope['class'] = "H"  # Top of Next Column in Hope Drive
    hopeiiiii = hope[(hope['type'] == 'PM - ACUTES ')]
    
    # Group by 'date' and count occurrences
    hopeiiiii['count'] = hopeiiiii.groupby(['date'])['provider'].cumcount()
    
    # Reserve H10 and H11 for the first two occurrences, subsequent start from H12
    hopeiiiii['class'] = hopeiiiii['count'].apply(
        lambda count: "H10" if count == 0 else ("H11" if count == 1 else "H" + str(count + 12))
    )
    
    hopeiiiii = hopeiiiii.loc[:, ('date', 'type', 'provider', 'clinic', 'class')]
    hopeiiiii.to_csv('9.csv', index=False)
    
    ######STARTS AT H2 (AM - Continuity)
    hope['H'] = "H"
    hopei = hope[(hope['type'] == 'AM - Continuity ')]
    
    # Group by 'date' and count occurrences
    hopei['count'] = hopei.groupby(['date'])['provider'].cumcount() + 2  # Start at H2
    hopei['class'] = "H" + hopei['count'].astype(str)
    hopei = hopei.loc[:, ('date', 'type', 'provider', 'clinic', 'class')]
    hopei.to_csv('5.csv', index=False)
    
    ######STARTS AT H12 (PM - Continuity)
    hope['H'] = "H"
    hopeii = hope[(hope['type'] == 'PM - Continuity ')]
    
    # Group by 'date' and count occurrences
    hopeii['count'] = hopeii.groupby(['date'])['provider'].cumcount() + 12  # Start at H12
    hopeii['class'] = "H" + hopeii['count'].astype(str)
    hopeii = hopeii.loc[:, ('date', 'type', 'provider', 'clinic', 'class')]
    hopeii.to_csv('6.csv', index=False)
    
    
    
    #############################################################################################################
    #MHS = pd.read_csv('MHS.csv')
    #MHS['H'] = "H"
    #MHSi = MHS[(MHS['type'] == 'AM - Continuity')]
    #MHSi['count'] = MHSi.groupby(['date'])['provider'].cumcount() + 0
    #MHSi['class'] = "H" + MHSi['count'].astype(str)
    #MHSi = MHSi.loc[:, ('date','type','provider','clinic','class')]
    #MHSi.to_csv('10.csv',index=False)
    
    #MHS['H'] = "H"
    #MHSii = MHS[(MHS['type'] == 'PM - Continuity')]
    #MHSii['count'] = MHSii.groupby(['date'])['provider'].cumcount() + 10
    #MHSii['class'] = "H" + MHSii['count'].astype(str)
    #MHSii = MHSii.loc[:, ('date','type','provider','clinic','class')]
    #MHSii.to_csv('11.csv',index=False)
    #############################################################################################################
    
    t1=pd.read_csv('1.csv')
    t2=pd.read_csv('2.csv')
    t3=pd.read_csv('3.csv')
    t4=pd.read_csv('4.csv')
    t5=pd.read_csv('5.csv')
    t6=pd.read_csv('6.csv')
    t7=pd.read_csv('7.csv')
    t8=pd.read_csv('8.csv')
    t9=pd.read_csv('9.csv')
    #t10=pd.read_csv('10.csv')
    #t11=pd.read_csv('11.csv')
    
    final2 = pd.DataFrame(columns=t1.columns)
    final2 = pd.concat([final2,t1,t2,t3,t4,t5,t6,t7,t8,t9]) #t10,t11])
    final2.to_csv('final2.csv',index=False)
    
    #final1 = pd.DataFrame(columns=NYEi.columns)
    #final1 = pd.concat([final1,hopei,hopeii,hopeiii,hopeiiii,ETOWNi,ETOWNii,NYEi,NYEii])
    #final1.to_csv('final1.csv',index=True)
    
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
    
    import openpyxl
    from openpyxl.styles import Alignment
    
    # Load workbooks and worksheets
    wb = openpyxl.load_workbook('final.xlsx')
    ws = wb['Sheet1']
    
    wb1 = openpyxl.load_workbook('OPD.xlsx')
    ws1 = wb1['HOPE_DRIVE']
    
    def generate_mapping(start_value):
        # Start with the special mappings
        mapping = {f"H{i}": start_value + i for i in range(0, 20)}
        
        return mapping
    
    # Generate the mappings starting from 6 for the first group
    common_mapping_1 = generate_mapping(6)
    # Construct the t_mapping dictionary with the common structure for T0 to T6
    t_mapping_1 = {f"T{i}": common_mapping_1 for i in range(7)}
    
    # Generate the mappings starting from 30 for the second group
    common_mapping_2 = generate_mapping(30)
    # Construct the t_mapping dictionary with the common structure for T7 to T13
    t_mapping_2 = {f"T{i}": common_mapping_2 for i in range(7, 14)}
    
    # Generate the mappings starting from 54 for the second group
    common_mapping_3 = generate_mapping(54)
    # Construct the t_mapping dictionary with the common structure for T14 to T20
    t_mapping_3 = {f"T{i}": common_mapping_3 for i in range(14, 21)}
    
    # Generate the mappings starting from 30 for the second group
    common_mapping_4 = generate_mapping(78)
    # Construct the t_mapping dictionary with the common structure for T21 to T28
    t_mapping_4 = {f"T{i}": common_mapping_4 for i in range(21, 28)}
    
    # Combine both mappings
    combined_t_mapping = {**t_mapping_1, **t_mapping_2, **t_mapping_3, **t_mapping_4}
    
    # Now the `combined_t_mapping` will have the correct structure
    #print(combined_t_mapping)
    
    
    column_mapping = {"T0": 2, "T1": 3, "T2": 4, "T3": 5, "T4": 6, "T5": 7, "T6": 8,
                      "T7": 2, "T8": 3, "T9": 4, "T10":5, "T11":6, "T12":7, "T13":8,
                      "T14":2, "T15":3, "T16":4, "T17":5, "T18":6, "T19":7, "T20":8,
                      "T21":2, "T22":3, "T23":4, "T24":5, "T25":6, "T26":7, "T27":8, 
                     }
    
    # Iterate through rows in the worksheet
    for row in ws.iter_rows():
        t_value = row[7].value  # Value in column H (index 7)
        h_value = row[6].value  # Value in column G (index 6)
        location = row[4].value  # Value in column E (index 4)
    
        # Check conditions and apply mapping
        if location == "HOPE_DRIVE" and t_value in combined_t_mapping and h_value in combined_t_mapping[t_value]:
            target_row = combined_t_mapping[t_value][h_value]
            target_column = column_mapping[t_value]
            ws1.cell(row=target_row, column=target_column).value = row[5].value  # Value in column F (index 5)
            ws1.cell(row=target_row, column=target_column).alignment = Alignment(horizontal='center')
    
    # Save updated workbook
    wb1.save('OPD.xlsx')
    
    ###############################################################################################
    
    import openpyxl
    from openpyxl.styles import Alignment
    
    # Load workbooks and worksheets
    wb = openpyxl.load_workbook('final.xlsx')
    ws = wb['Sheet1']
    
    wb1 = openpyxl.load_workbook('OPD.xlsx')
    ws1 = wb1['ETOWN']
    
    def generate_mapping(start_value):
        # Create a dictionary with H0 to H19 keys, with values starting from `start_value`
        mapping = {f"H{i}": start_value + i for i in range(0, 20)}
        
        return mapping
    
    # Generate the mappings starting from 6 for the first group
    common_mapping_1 = generate_mapping(6)
    # Construct the t_mapping dictionary with the common structure for T0 to T6
    t_mapping_1 = {f"T{i}": common_mapping_1 for i in range(7)}
    
    # Generate the mappings starting from 30 for the second group
    common_mapping_2 = generate_mapping(30)
    # Construct the t_mapping dictionary with the common structure for T7 to T13
    t_mapping_2 = {f"T{i}": common_mapping_2 for i in range(7, 14)}
    
    # Generate the mappings starting from 54 for the third group
    common_mapping_3 = generate_mapping(54)
    # Construct the t_mapping dictionary with the common structure for T14 to T20
    t_mapping_3 = {f"T{i}": common_mapping_3 for i in range(14, 21)}
    
    # Generate the mappings starting from 78 for the fourth group
    common_mapping_4 = generate_mapping(78)
    # Construct the t_mapping dictionary with the common structure for T21 to T28
    t_mapping_4 = {f"T{i}": common_mapping_4 for i in range(21, 28)}
    
    # Combine both mappings
    combined_t_mapping = {**t_mapping_1, **t_mapping_2, **t_mapping_3, **t_mapping_4}
    
    # Now the `combined_t_mapping` will have the correct structure
    #print(combined_t_mapping)
    
    
    column_mapping = {"T0": 2, "T1": 3, "T2": 4, "T3": 5, "T4": 6, "T5": 7, "T6": 8,
                      "T7": 2, "T8": 3, "T9": 4, "T10":5, "T11":6, "T12":7, "T13":8,
                      "T14":2, "T15":3, "T16":4, "T17":5, "T18":6, "T19":7, "T20":8,
                      "T21":2, "T22":3, "T23":4, "T24":5, "T25":6, "T26":7, "T27":8, 
                     }
    
    # Iterate through rows in the worksheet
    for row in ws.iter_rows():
        t_value = row[7].value  # Value in column H (index 7)
        h_value = row[6].value  # Value in column G (index 6)
        location = row[4].value  # Value in column E (index 4)
    
        # Check conditions and apply mapping
        if location == "ETOWN" and t_value in combined_t_mapping and h_value in combined_t_mapping[t_value]:
            target_row = combined_t_mapping[t_value][h_value]
            target_column = column_mapping[t_value]
            ws1.cell(row=target_row, column=target_column).value = row[5].value  # Value in column F (index 5)
            ws1.cell(row=target_row, column=target_column).alignment = Alignment(horizontal='center')
    
    # Save updated workbook
    wb1.save('OPD.xlsx')
    
    ###############################################################################################
    
    import openpyxl
    from openpyxl.styles import Alignment
    
    # Load workbooks and worksheets
    wb = openpyxl.load_workbook('final.xlsx')
    ws = wb['Sheet1']
    
    wb1 = openpyxl.load_workbook('OPD.xlsx')
    ws1 = wb1['NYES']
    
    def generate_mapping(start_value):
        # Create a dictionary with H0 to H19 keys, with values starting from `start_value`
        mapping = {f"H{i}": start_value + i for i in range(0, 20)}
        
        return mapping
    
    # Generate the mappings starting from 6 for the first group
    common_mapping_1 = generate_mapping(6)
    # Construct the t_mapping dictionary with the common structure for T0 to T6
    t_mapping_1 = {f"T{i}": common_mapping_1 for i in range(7)}
    
    # Generate the mappings starting from 30 for the second group
    common_mapping_2 = generate_mapping(30)
    # Construct the t_mapping dictionary with the common structure for T7 to T13
    t_mapping_2 = {f"T{i}": common_mapping_2 for i in range(7, 14)}
    
    # Generate the mappings starting from 54 for the third group
    common_mapping_3 = generate_mapping(54)
    # Construct the t_mapping dictionary with the common structure for T14 to T20
    t_mapping_3 = {f"T{i}": common_mapping_3 for i in range(14, 21)}
    
    # Generate the mappings starting from 78 for the fourth group
    common_mapping_4 = generate_mapping(78)
    # Construct the t_mapping dictionary with the common structure for T21 to T28
    t_mapping_4 = {f"T{i}": common_mapping_4 for i in range(21, 28)}
    
    # Combine both mappings
    combined_t_mapping = {**t_mapping_1, **t_mapping_2, **t_mapping_3, **t_mapping_4}
    
    # Now the `combined_t_mapping` will have the correct structure
    #print(combined_t_mapping)
    
    
    column_mapping = {"T0": 2, "T1": 3, "T2": 4, "T3": 5, "T4": 6, "T5": 7, "T6": 8,
                      "T7": 2, "T8": 3, "T9": 4, "T10":5, "T11":6, "T12":7, "T13":8,
                      "T14":2, "T15":3, "T16":4, "T17":5, "T18":6, "T19":7, "T20":8,
                      "T21":2, "T22":3, "T23":4, "T24":5, "T25":6, "T26":7, "T27":8, 
                     }
    
    # Iterate through rows in the worksheet
    for row in ws.iter_rows():
        t_value = row[7].value  # Value in column H (index 7)
        h_value = row[6].value  # Value in column G (index 6)
        location = row[4].value  # Value in column E (index 4)
    
        # Check conditions and apply mapping
        if location == "NYES" and t_value in combined_t_mapping and h_value in combined_t_mapping[t_value]:
            target_row = combined_t_mapping[t_value][h_value]
            target_column = column_mapping[t_value]
            ws1.cell(row=target_row, column=target_column).value = row[5].value  # Value in column F (index 5)
            ws1.cell(row=target_row, column=target_column).alignment = Alignment(horizontal='center')
    
    # Save updated workbook
    wb1.save('OPD.xlsx')
    
