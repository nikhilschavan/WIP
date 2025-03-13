from flask import Flask, render_template, request, redirect, url_for
import pandas as pd
import os

app = Flask(__name__)

# Excel file path
EXCEL_FILE = "data.xlsx"

# Ensure the file exists with proper headers
if not os.path.exists(EXCEL_FILE):
    df = pd.DataFrame(columns=["Ring_code", "Description", "Ring_Type","Stage","Qty"])
    df.to_excel(EXCEL_FILE, index=False)

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        # Get form data
        Ring_code = request.form["Ring_code"]
        Description = request.form["Description"]
        Ring_Type = request.form["Ring_Type"]
        Stage = request.form["Stage"]
        Qty = request.form["Qty"]

        # Read existing data
        df = pd.read_excel(EXCEL_FILE)

        # Append new data
        new_data = pd.DataFrame([[Ring_code, Description, Ring_Type, Stage, Qty]], columns=["Ring_code", "Description", "Ring_Type","Stage","Qty"])
        df = pd.concat([df, new_data], ignore_index=True)

        # Save to Excel
        df.to_excel(EXCEL_FILE, index=False)

        return redirect(url_for("index"))

    return render_template("index.html")
if __name__ == "__main__":
    app.run(host="0.0.0.0",port=8000)