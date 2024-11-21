import streamlit as st

# Importing the content of the pages
from welcome import show_welcome_page
from opd import show_opd_page  # Example of another page import

# Initialize session state variable to handle page routing
if "page" not in st.session_state:
    st.session_state.page = "welcome"  # Default page is welcome

def main():
    # Page routing based on session state
    if st.session_state.page == "welcome":
        show_welcome_page()  # Show the welcome page
    elif st.session_state.page == "opd":
        show_opd_page()  # Show the opd page

if __name__ == "__main__":
    main()
