#########
# Sidebar
#########

import pandas as pd
import streamlit as st
from datetime import datetime
from io import BytesIO
import base64
from openpyxl import Workbook

# Define the Streamlit app
st.title("Invoice Data Analysis")

# Initialize session state
if 'raw_invoice' not in st.session_state:
    st.session_state.raw_invoice = None
if 'wsr_consolidated' not in st.session_state:
    st.session_state.wsr_consolidated = None
if 'onboarding_tracker' not in st.session_state:
    st.session_state.onboarding_tracker = None
if 'raw_invoice_copy' not in st.session_state:
    st.session_state.raw_invoice_copy = None
if 'wsr_consolidated_copy' not in st.session_state:
    st.session_state.wsr_consolidated_copy = None
if 'date_ranges' not in st.session_state:
    st.session_state.date_ranges = {}
if 'submit_button_pressed' not in st.session_state:
    st.session_state.submit_button_pressed = False

# Upload files in Streamlit
uploaded_raw_invoice = st.file_uploader("Upload Raw Invoice Excel File", type=["xlsx"])
uploaded_wsr_consolidated = st.file_uploader("Upload WSR Consolidated Excel File", type=["xlsb"])
uploaded_onboarding_tracker = st.file_uploader("Upload Onboarding Tracker Excel File", type=["xlsx"])

# Check if files are uploaded and load data if necessary
if uploaded_raw_invoice:
    st.session_state.raw_invoice = pd.read_excel(uploaded_raw_invoice, skiprows=1).drop("Unnamed: 0", axis=1)
    st.session_state.raw_invoice["Name"] = st.session_state.raw_invoice["Name"].str.replace(r' [A-Z]\b', '', regex=True)
    st.session_state.raw_invoice = st.session_state.raw_invoice[st.session_state.raw_invoice["Name"] != "Grand Total"]

if uploaded_wsr_consolidated:
    st.session_state.wsr_consolidated = pd.read_excel(uploaded_wsr_consolidated, sheet_name="Invoice Review", skiprows=5)
    st.session_state.wsr_consolidated['Reporting Week (MM/DD/YYYY)'] = pd.to_datetime(st.session_state.wsr_consolidated['Reporting Week (MM/DD/YYYY)'], unit='D', origin='1899-12-30')
    st.session_state.wsr_consolidated["Contractor (Last Name, First Name)2"] = st.session_state.wsr_consolidated["Contractor (Last Name, First Name)2"].str.replace(r' [A-Z]\b', '', regex=True)
    st.session_state.wsr_consolidated = st.session_state.wsr_consolidated.ffill()
    st.session_state.wsr_consolidated = st.session_state.wsr_consolidated[st.session_state.wsr_consolidated["Vendor Name"] != "Grand Total"]

if uploaded_onboarding_tracker:
    st.session_state.onboarding_tracker = pd.read_excel(uploaded_onboarding_tracker, sheet_name="Master List")
    st.session_state.onboarding_tracker["Candidate Name"] = st.session_state.onboarding_tracker["Candidate Name"].str.replace(r' [A-Z]\b', '', regex=True)

# Display the initial DataFrame (First)
if st.session_state.raw_invoice is not None:
    st.write(st.session_state.raw_invoice)

# Get unique combinations of Name and Reporting Week
try:
    unique_combinations = st.session_state.raw_invoice[['Name', 'Effective Bill Date']].drop_duplicates()
except TypeError:
    st.write("Error: 'raw_invoice' is not available or is None.")

for index, row in unique_combinations.iterrows():
    name = row['Name']
    effective_date = row['Effective Bill Date']

    # Sidebar for user input
    st.sidebar.header(f"User Input for {name} ({effective_date})")
    start_date_input = st.sidebar.text_input(f"Enter Start Date (MM/DD/YYYY) for {name} ({effective_date}):", key=f"{name}_{effective_date}_start_date")
    end_date_input = st.sidebar.text_input(f"Enter End Date (MM/DD/YYYY) for {name} ({effective_date}):", key=f"{name}_{effective_date}_end_date")

    if start_date_input and end_date_input:
        st.session_state.date_ranges[(name, effective_date)] = (start_date_input, end_date_input)

# "Submit All Date Ranges" button
if st.sidebar.button("Submit All Date Ranges"):
    st.session_state.submit_button_pressed = True

# Perform calculations when the button is pressed
if st.session_state.submit_button_pressed:
    if st.session_state.raw_invoice_copy is None:
        st.session_state.raw_invoice_copy = st.session_state.raw_invoice.copy()
        st.session_state.wsr_consolidated_copy = st.session_state.wsr_consolidated.copy()
    
    for (name, effective_date), (start_date, end_date) in st.session_state.date_ranges.items():
        start_date = pd.to_datetime(start_date, format='%m/%d/%Y')
        end_date = pd.to_datetime(end_date, format='%m/%d/%Y')

        filtered_wsr = st.session_state.wsr_consolidated_copy[st.session_state.wsr_consolidated_copy['Contractor (Last Name, First Name)2'] == name]
        filtered_wsr = filtered_wsr[(filtered_wsr['Reporting Week (MM/DD/YYYY)'] >= start_date) &
                                    (filtered_wsr['Reporting Week (MM/DD/YYYY)'] <= end_date)]

        total_hours = filtered_wsr['Sum of Time Spent (Hours) '].sum()
        contract_rate = filtered_wsr['Sum of Cost Calc'].sum() / total_hours
        contract_rate = round(contract_rate, 2)
        cost_check = contract_rate * total_hours

        st.session_state.raw_invoice_copy.loc[st.session_state.raw_invoice_copy['Name'] == name, 'WSR Hours'] = total_hours
        st.session_state.raw_invoice_copy.loc[st.session_state.raw_invoice_copy['Name'] == name, 'Contract Rate'] = contract_rate
        st.session_state.raw_invoice_copy.loc[st.session_state.raw_invoice_copy['Name'] == name, 'Cost Check'] = cost_check
        
    # Display the updated DataFrame in Streamlit
    st.write(st.session_state.raw_invoice_copy)  


    # Input field for Excel file name
    excel_filename = st.text_input("Enter Excel File Name (without extension)", "InvoiceReview")

    # Save to Excel button
    if st.button('Save Data to Excel'):
        # Save the filtered dataframe to an Excel file in memory
        excel_buffer = BytesIO()
        st.session_state.raw_invoice_copy.to_excel(excel_buffer, index=False)
        excel_data = excel_buffer.getvalue()

        # Generate a download link for the Excel file
        b64 = base64.b64encode(excel_data).decode('utf-8')
        excel_filename = f"{excel_filename}.xlsx"
        href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="{excel_filename}">Download Excel File</a>'
        st.markdown(href, unsafe_allow_html=True)
