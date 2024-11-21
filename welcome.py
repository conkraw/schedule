import streamlit as st

def show_welcome_page():
    st.title("Welcome to the OPD Management System")

    # Display instructions or other welcome content
    st.write("This is the welcome page. To start managing OPD, click below.")

    # Button to navigate to the OPD creation page
    if st.button("Create OPD"):
        st.session_state.page = "opd"  # Update the session state to show the opd page
        st.rerun()  # Rerun the app to display the OPD page
