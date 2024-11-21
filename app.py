import streamlit as st
import pandas as pd
import datetime

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
                
                # Display the converted date
                st.write(f"Valid date entered: {test_date.strftime('%m/%d/%Y')}")
            except ValueError:
                # Display an error message if the date format is incorrect
                st.error('Invalid date format. Please enter the date in m/d/yyyy format.')
        else:
            # If no date is entered, display an error
            st.error('Please enter a date.')
