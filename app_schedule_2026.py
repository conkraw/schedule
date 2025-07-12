import streamlit as st
import csv
import datetime
import pandas as pd
import numpy as np
from datetime import timedelta
import xlsxwriter
import openpyxl
from openpyxl import Workbook
import io
from io import BytesIO, StringIO
import os
import time 
import random
from openpyxl.styles import Font, Alignment

# Configure page
st.set_page_config(
    page_title="OPD Creator",      # shown in browser tab
    layout="wide",                 # full‑width
    initial_sidebar_state="expanded"
)

# Main title
st.title("Outpatient Department (OPD) Schedule Creator")

# Sidebar
st.sidebar.title("Navigation")
page = st.sidebar.radio("Go to:", [
    "Home",
    "Create OPD",
    "Upload Files",
    "Generate Schedule",
    "Download OPD"
])

# Example of reacting to the sidebar choice
if page == "Home":
    st.write("👋 Welcome! Use the sidebar to navigate through the app.")
elif page == "Create OPD":
    st.write("➡️ Here you’ll enter your start date and generate blank OPD templates.")
elif page == "Upload Files":
    st.write("⬆️ Upload your Excel/CSV files to be processed.")
elif page == "Generate Schedule":
    st.write("⚙️ Processing files and assigning students...")
elif page == "Download OPD":
    st.write("👇 Finally, download your completed OPD.xlsx here.")
