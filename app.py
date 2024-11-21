import streamlit as st

st.set_page_config(layout="wide")

# Importing the content of the pages
from welcome import show_welcome_page
from opd import show_opd_page  # Example of another page import

# Initialize session state variable to handle page routing
if "page" not in st.session_state:
    st.session_state.page = "welcome"  # Default page is welcome

def main():
    # Page routing based on session state
    if st.session_state.page == "welcome":
        show_welcome_page()
    #elif st.session_state.page == "opd":
    #    show_opd_page()
