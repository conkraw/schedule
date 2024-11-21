import streamlit as st
import pandas as pd
import datetime
import requests
import subprocess
import sys
import os

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

    # URL to your GitHub raw file
    github_script_url = "https://raw.githubusercontent.com/yourusername/yourrepo/main/opd.py"

    # Download the opd.py script from GitHub
    response = requests.get(github_script_url)
    if response.status_code == 200:
        # Save the content of the script as a .py file locally
        script_path = "/tmp/opd.py"
        with open(script_path, 'wb') as f:
            f.write(response.content)

        # Ensure missing modules are installed before running the script
        required_modules = ['xlsxwriter']  # Add other required modules here if needed

        for module in required_modules:
            try:
                # Try importing the module
                __import__(module)
            except ImportError:
                # If module is missing, install it
                st.write(f"Installing missing module: {module}")
                subprocess.check_call([sys.executable, "-m", "pip", "install", module])

        # Execute the Python script using subprocess
        try:
            result = subprocess.run([sys.executable, script_path], capture_output=True, text=True)
            st.write("Script executed successfully!")
            st.text(result.stdout)  # Output from the script
            if result.stderr:
                st.error(result.stderr)  # Error messages from the script
        except Exception as e:
            st.error(f"Error executing the script: {e}")
    else:
        st.error("Failed to download the script from GitHub.")

