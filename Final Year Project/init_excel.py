import pandas as pd
import os

# Excel file name
EXCEL_FILE = "expenses.xlsx"

# Initialize Excel file
def init_excel():
    try:
        # Check if the Excel file exists
        if not os.path.exists(EXCEL_FILE):
            print(f"File '{EXCEL_FILE}' not found. Creating a new file...")
            # Create a new DataFrame with required columns
            df = pd.DataFrame(columns=["Amount", "Category", "Date", "Notes"])
            # Save the DataFrame to an Excel file
            df.to_excel(EXCEL_FILE, index=False, engine="openpyxl")
            print(f"New Excel file '{EXCEL_FILE}' created successfully.")
        else:
            print(f"Excel file '{EXCEL_FILE}' already exists. No need to initialize.")
    except Exception as e:
        print(f"An error occurred while initializing the Excel file: {e}")