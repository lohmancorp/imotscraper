import os
import subprocess

# Directory where the setup is to be done
setup_dir = "/Users/bggidt01/Desktop/scraper"

# Check if the directory exists, create if it does not
if not os.path.exists(setup_dir):
    os.makedirs(setup_dir)

# Navigate to the directory
os.chdir(setup_dir)

# Creating a virtual environment
subprocess.run(["python3", "-m", "venv", "venv"])

# Installing Flask and other required packages
# Note: You might need to activate the virtual environment manually in some cases
subprocess.run([f"{setup_dir}/venv/bin/pip", "install", "flask", "requests", "pandas", "beautifulsoup4", "lxml", "openpyxl", "xlsxwriter"])

# Flask app script content
flask_app_content = """
from flask import Flask, request, render_template
import requests
from bs4 import BeautifulSoup
import pandas as pd
import numpy as np
import subprocess
import re
from datetime import datetime

app = Flask(__name__)

def extract_date_from_url(url):
    match = re.search(r'&date=(\\d{2}\\.\\d{2}\\.\\d{4})', url)
    if match:
        date_str = match.group(1)
        return datetime.strptime(date_str, '%d.%m.%Y').strftime('%Y-%m-%d')
    else:
        return None

def fetch_and_parse_table(url, table_id):
    response = requests.get(url)
    response.raise_for_status()
    soup = BeautifulSoup(response.content, 'html.parser')
    table = soup.find('table', id=table_id)
    if table:
        df = pd.read_html(str(table), header=0, encoding='utf-8')[0]
        return df
    else:
        return None

def post_process_dataframe(df, report_date):
    df = df.copy()
    note_indicator = '*Забележка:'
    df = df[~df[df.columns[0]].astype(str).str.contains(note_indicator, regex=False)]
    df.drop(index=[3], inplace=True)
    df.drop(columns=df.columns[[1, 4, 7, 10]], inplace=True)
    df.columns = ['Region', '1_Room_Price', '1_Room_Price_Sqm', '2_Room_Price', '2_Room_Price_Sqm',
                  '3_Room_Price', '3_Room_Price_Sqm', 'Avg_Price_Sqm']
    for col in df.columns[1:]:
        df[col] = df[col].str.replace(' ', '').replace('-', np.nan)
        df[col] = pd.to_numeric(df[col], errors='coerce')
    df.dropna(how='all', inplace=True)
    df['report_date'] = report_date
    return df

def open_in_excel(file_name):
    subprocess.run(["open", "-a", "Microsoft Excel", file_name])
    script = '''
        tell application "Microsoft Excel"
            activate
        end tell
    '''
    subprocess.run(["osascript", "-e", script])

def process_data(url, output, excel):
    report_date = extract_date_from_url(url)
    table_df = fetch_and_parse_table(url, 'tableStats')
    if table_df is not None:
        # Post-process DataFrame
        processed_df = post_process_dataframe(table_df, report_date)

        # Remove rows where 'Region' column is 'Район' or null/empty
        processed_df = processed_df[processed_df['Region'].notna()]
        processed_df = processed_df[~processed_df['Region'].str.strip().str.lower().eq('район')]

        # Continue with creating the Excel file
        if report_date:
            output_file_name = f"{report_date} - {output}.xlsx"
        else:
            output_file_name = f"{output}.xlsx"

        writer = pd.ExcelWriter(output_file_name, engine='xlsxwriter')
        processed_df.to_excel(writer, sheet_name='Sheet1', index=False)

        workbook = writer.book
        worksheet = writer.sheets['Sheet1']

        # Define header format
        header_format = workbook.add_format({
            'bold': True,
            'text_wrap': True,
            'valign': 'top',
            'fg_color': '#007ddf',  # blue
            'font_color': '#FFFFFF',  # white
            'border': 1
        })

        # Define the Euro accounting format
        euro_format = workbook.add_format({'num_format': '€ #,##0.00', 'align': 'right'})

        # Apply the Euro accounting format to the price columns
        price_columns = ['1_Room_Price', '1_Room_Price_Sqm', '2_Room_Price', '2_Room_Price_Sqm', '3_Room_Price', '3_Room_Price_Sqm', 'Avg_Price_Sqm']
        for column in price_columns:
            col_idx = processed_df.columns.get_loc(column)
            worksheet.set_column(col_idx, col_idx, 18, euro_format)  # Adjust the column width as needed

        # Apply the header format to the header row
        for col_num, value in enumerate(processed_df.columns):
            worksheet.write(0, col_num, value, header_format)

        # Apply the table format to the entire range of the DataFrame
        last_row = len(processed_df.index)
        last_col = len(processed_df.columns) - 1
        column_settings = [{'header': column_name, 'header_format': header_format} for column_name in processed_df.columns]
        worksheet.add_table(0, 0, last_row, last_col, {'columns': column_settings, 'style': 'Table Style Medium 9'})

        writer.close()

        if excel:
            open_in_excel(output_file_name)


@app.route('/')
def form():
    return render_template('form.html')

@app.route('/submit', methods=['POST'])
def submit():
    url = request.form['url']
    output = request.form['output']
    excel = 'excel' in request.form
    process_data(url, output, excel)
    return 'Data processed successfully!'

if __name__ == '__main__':
    app.run(debug=True)
"""

# Creating the Flask app script file
with open(os.path.join(setup_dir, "app.py"), "w") as file:
    file.write(flask_app_content)

# Creating the templates directory
templates_dir = os.path.join(setup_dir, "templates")
os.makedirs(templates_dir, exist_ok=True)

# HTML form content
html_form_content = """
<!DOCTYPE html>
<html lang="en">
<head>
    <!-- Meta tags and stylesheets from both documents -->
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Imot.bg Scrapper - Signup Form</title>
    <link href="https://stackpath.bootstrapcdn.com/bootstrap/5.2.0/css/bootstrap.min.css" rel="stylesheet">
    <link rel="stylesheet" href="/_next/static/css/635da5ebe4a3a30d.css">
    <!-- Additional styles and scripts -->
</head>
<body>
    <!-- Navbar from the first document -->
    <nav class="navbar navbar-expand-lg navbar-light bg-light">
        <div class="container-fluid">
            <a class="navbar-brand" href="#">Imot.bg Scrapper</a>
        </div>
    </nav>

    <!-- Main content combining elements from both documents -->
    <div class="container mt-5">
        <!-- Your form fields -->
        <form action="/submit" method="post" class="text-center">
            <div class="mb-3">
                <label for="url" class="form-label">URL:</label>
                <input type="text" class="form-control" id="url" name="url" maxlength="300">
            </div>
            <div class="mb-3">
                <label for="output" class="form-label">Output File Name:</label>
                <input type="text" class="form-control" id="output" name="output">
            </div>
            <div class="mb-3 form-check">
                <input type="checkbox" class="form-check-input" id="excel" name="excel">
                <label for="excel" class="form-check-label">Open in Excel</label>
            </div>
            <button type="submit" class="btn btn-primary">Submit</button>
        </form>
        <!-- Additional content from the second document -->
    </div>

    <!-- Footer from the first document -->
    <footer class="footer mt-auto py-3 bg-light">
        <div class="container text-center">
            <span class="text-muted">Please don't sue me Imot.bg!! xoxoxo.</span>
        </div>
    </footer>

    <!-- Scripts from both documents -->
    <script src="https://stackpath.bootstrapcdn.com/bootstrap/5.2.0/js/bootstrap.bundle.min.js"></script>
    <!-- Additional scripts -->
</body>
</html>

"""

# Creating the HTML form file
with open(os.path.join(templates_dir, "form.html"), "w") as file:
    file.write(html_form_content)

print("Setup complete. You can now run 'app.py' in your Flask environment.")

