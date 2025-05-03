from flask import Flask, render_template, request, redirect, url_for
import pandas as pd
import os

app = Flask(__name__)

# Excel file to store data
EXCEL_FILE = "expenses.xlsx"

# Initialize the Excel file
def init_excel():
    if not os.path.exists(EXCEL_FILE):
        print(f"Creating a new Excel file: {EXCEL_FILE}")
        df = pd.DataFrame(columns=["Amount", "Category", "Date", "Notes"])
        df.to_excel(EXCEL_FILE, index=False, engine="openpyxl")
        print(f"Excel file '{EXCEL_FILE}' created successfully.")

# Route: Home
@app.route('/')
def index():
    # Read data from Excel
    if os.path.exists(EXCEL_FILE):
        df = pd.read_excel(EXCEL_FILE)
    else:
        df = pd.DataFrame(columns=["Amount", "Category", "Date", "Notes"])
    return render_template('index.html', expenses=df.to_dict(orient='records'))

# Route: Add Expense
@app.route('/add-expense', methods=['POST'])
def add_expense():
    # Get form data
    amount = request.form['amount']
    category = request.form['category']
    date = request.form['date']
    notes = request.form['notes']

    # Append data to Excel
    if os.path.exists(EXCEL_FILE):
        df = pd.read_excel(EXCEL_FILE)
    else:
        df = pd.DataFrame(columns=["Amount", "Category", "Date", "Notes"])
    
    new_row = {"Amount": float(amount), "Category": category, "Date": date, "Notes": notes}
    df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
    df.to_excel(EXCEL_FILE, index=False, engine="openpyxl")

    return redirect(url_for('index'))

# Route: Delete Expense
@app.route('/delete-expense/<int:row_id>', methods=['POST'])
def delete_expense(row_id):
    if os.path.exists(EXCEL_FILE):
        df = pd.read_excel(EXCEL_FILE)
        df = df.drop(index=row_id).reset_index(drop=True)
        df.to_excel(EXCEL_FILE, index=False, engine="openpyxl")
    return redirect(url_for('index'))

if __name__ == "__main__":
    init_excel()
    app.run(debug=True)