# Invoice Data Analysis Streamlit App

## Overview
This Streamlit app is designed for analyzing invoice data. It allows you to upload and process raw invoice data, WSR Consolidated data, and Onboarding Tracker data. You can calculate contract rates, perform data analysis, and save the results to an Excel file.

## Getting Started
To use this app, follow these steps:

1. Make sure you have the required dependencies installed. You can install them using `pip`:

       pip install pandas streamlit openpyxl pyxlsb
       
2. Run the Streamlit app by executing the following command in your terminal:

        streamlit run invoiceReview.py
        
3. The app will launch in your web browser. You can start using it by following the instructions provided in the app.


## Features

    - Upload Raw Invoice, WSR Consolidated, and Onboarding Tracker Excel files.
    - Calculate contract rates and perform data analysis.
    - Save the analyzed data to an Excel file.
    - Customizable date ranges for analysis.
    - User-friendly interface.

## Usage

    - Upload Data Files:
        - Upload the Raw Invoice Excel file.
        - Upload the WSR Consolidated Excel file.
        - Upload the Onboarding Tracker Excel file.

    Configure Date Ranges:
        For each unique combination of "Name" and "Effective Bill Date" in the Labor Invoice data, you can configure a date range in the app's sidebar.

    Submit Date Ranges:
        After configuring date ranges, click the "Submit All Date Ranges" button to perform calculations.

    View and Save Results:
        View the calculated results in the Streamlit app interface.
        Enter an Excel file name (without the extension) and click "Save Data to Excel" to save the results to an Excel file.