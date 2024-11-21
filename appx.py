import streamlit as st
import pandas as pd
import datetime

# Display the title of the app
st.title('Date Input for OPD')

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
            
            # Display the converted date
            st.write(f"Valid date entered: {test_date.strftime('%m/%d/%Y')}")
        except ValueError:
            # Display an error message if the date format is incorrect
            st.error('Invalid date format. Please enter the date in m/d/yyyy format.')
    else:
        # If no date is entered, display an error
        st.error('Please enter a date.')

st.write('Upload the following Excel files:')
uploaded_files = {}
uploaded_files['HOPE_DRIVE.xlsx'] = st.file_uploader('Upload HOPE_DRIVE.xlsx', type='xlsx')
uploaded_files['ETOWN.xlsx'] = st.file_uploader('Upload ETOWN.xlsx', type='xlsx')
uploaded_files['NYES.xlsx'] = st.file_uploader('Upload NYES.xlsx', type='xlsx')

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


###############################################################################################
# Button to trigger the download
if st.button('Download OPD'):
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
