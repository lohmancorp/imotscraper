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
<!doctype html>
<html lang="en" data-bs-theme="auto">
  <head><script src="https://getbootstrap.com/docs/5.3/assets/js/color-modes.js"></script>

    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <meta name="description" content="">
    <meta name="author" content="Taylor Giddens">
    <title>Imot.BG · Scrapper</title>


    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/@docsearch/css@3">

<link href="https://getbootstrap.com/docs/5.3/dist/css/bootstrap.min.css" rel="stylesheet" integrity="sha384-T3c6CoIi6uLrA9TneNEoa7RxnatzjcDSCmG1MXxSR1GAsXEV/Dwwykc2MPK8M2HN" crossorigin="anonymous">

    <!-- Favicons -->
<link rel="apple-touch-icon" href="https://www.imot.bg/favicon.ico" sizes="180x180">
<link rel="icon" href="https://www.imot.bg/favicon.ico" sizes="32x32" type="image/png">
<link rel="icon" href="https://www.imot.bg/favicon.ico" sizes="16x16" type="image/png">
<meta name="theme-color" content="#712cf9">


    <style>
      .bd-placeholder-img {
        font-size: 1.125rem;
        text-anchor: middle;
        -webkit-user-select: none;
        -moz-user-select: none;
        user-select: none;
      }

      @media (min-width: 768px) {
        .bd-placeholder-img-lg {
          font-size: 3.5rem;
        }
      }

      .b-example-divider {
        width: 100%;
        height: 3rem;
        background-color: rgba(0, 0, 0, .1);
        border: solid rgba(0, 0, 0, .15);
        border-width: 1px 0;
        box-shadow: inset 0 .5em 1.5em rgba(0, 0, 0, .1), inset 0 .125em .5em rgba(0, 0, 0, .15);
      }

      .b-example-vr {
        flex-shrink: 0;
        width: 1.5rem;
        height: 100vh;
      }

      .bi {
        vertical-align: -.125em;
        fill: currentColor;
      }

      .nav-scroller {
        position: relative;
        z-index: 2;
        height: 2.75rem;
        overflow-y: hidden;
      }

      .nav-scroller .nav {
        display: flex;
        flex-wrap: nowrap;
        padding-bottom: 1rem;
        margin-top: -1px;
        overflow-x: auto;
        text-align: center;
        white-space: nowrap;
        -webkit-overflow-scrolling: touch;
      }

      .btn-bd-primary {
        --bd-violet-bg: #712cf9;
        --bd-violet-rgb: 112.520718, 44.062154, 249.437846;

        --bs-btn-font-weight: 600;
        --bs-btn-color: var(--bs-white);
        --bs-btn-bg: var(--bd-violet-bg);
        --bs-btn-border-color: var(--bd-violet-bg);
        --bs-btn-hover-color: var(--bs-white);
        --bs-btn-hover-bg: #6528e0;
        --bs-btn-hover-border-color: #6528e0;
        --bs-btn-focus-shadow-rgb: var(--bd-violet-rgb);
        --bs-btn-active-color: var(--bs-btn-hover-color);
        --bs-btn-active-bg: #5a23c8;
        --bs-btn-active-border-color: #5a23c8;
      }

      .bd-mode-toggle {
        z-index: 1500;
      }

      .bd-mode-toggle .dropdown-menu .active .bi {
        display: block !important;
      }
    </style>

    
    <!-- Custom styles for this template -->
    <link href="sign-in.css" rel="stylesheet">
  </head>
  <body class="d-flex align-items-center py-4 bg-body-tertiary">

    <div class="dropdown position-fixed bottom-0 end-0 mb-3 me-3 bd-mode-toggle">
      <button class="btn btn-bd-primary py-2 dropdown-toggle d-flex align-items-center"
              id="bd-theme"
              type="button"
              aria-expanded="false"
              data-bs-toggle="dropdown"
              aria-label="Toggle theme (auto)">
        <svg class="bi my-1 theme-icon-active" width="1em" height="1em"><use href="#circle-half"></use></svg>
        <span class="visually-hidden" id="bd-theme-text">Toggle theme</span>
      </button>
      <ul class="dropdown-menu dropdown-menu-end shadow" aria-labelledby="bd-theme-text">
        <li>
          <button type="button" class="dropdown-item d-flex align-items-center" data-bs-theme-value="light" aria-pressed="false">
            <svg class="bi me-2 opacity-50 theme-icon" width="1em" height="1em"><use href="#sun-fill"></use></svg>
            Light
            <svg class="bi ms-auto d-none" width="1em" height="1em"><use href="#check2"></use></svg>
          </button>
        </li>
        <li>
          <button type="button" class="dropdown-item d-flex align-items-center" data-bs-theme-value="dark" aria-pressed="false">
            <svg class="bi me-2 opacity-50 theme-icon" width="1em" height="1em"><use href="#moon-stars-fill"></use></svg>
            Dark
            <svg class="bi ms-auto d-none" width="1em" height="1em"><use href="#check2"></use></svg>
          </button>
        </li>
        <li>
          <button type="button" class="dropdown-item d-flex align-items-center active" data-bs-theme-value="auto" aria-pressed="true">
            <svg class="bi me-2 opacity-50 theme-icon" width="1em" height="1em"><use href="#circle-half"></use></svg>
            Auto
            <svg class="bi ms-auto d-none" width="1em" height="1em"><use href="#check2"></use></svg>
          </button>
        </li>
      </ul>
    </div>

    
<main class="form-signin w-100 m-auto" style="max-width: 600px;">
   <form action="/submit" method="post" class="text-center">

    <img class="mb-4" src="https://www.imot.bg/images/picturess/logo.svg" alt="Imot.bg Logo">
    <h1 class="h3 mb-3 fw-normal">Please provide information.</h1>

    <div class="form-floating">
      <input type="text" class="form-control" id="url" name="url" placeholder="https://imog.bg/full/url">
       <label for="url" class="form-label">URL:</label>
    </div>
    <div class="form-floating">
      <input type="text" class="form-control" id="output" name="output" placeholder="File Name (Usually City Name)">
      <label for="output">File Name</label>
    </div>

    <div class="form-check text-start my-3">
      <input class="form-check-input" type="checkbox" value="Open In Excel" id="excel" name="excel">
      <label class="form-check-label" for="flexCheckDefault">
        Open In Excel
      </label>
    </div>
    <button class="btn btn-primary w-100 py-2" type="submit">Process</button>
    <p class="mt-5 mb-3 text-body-secondary">&copy; Taylor 2023</p>
  </form>
</main>
<script src="https://getbootstrap.com/docs/5.3/dist/js/bootstrap.bundle.min.js" integrity="sha384-C6RzsynM9kWDrMNeT87bh95OGNyZPhcTNXj1NW7RuBCsyN/o0jlpcV8Qyq46cDfL" crossorigin="anonymous"></script>

    </body>
</html>
"""

# Creating the HTML form file
with open(os.path.join(templates_dir, "form.html"), "w") as file:
    file.write(html_form_content)

print("Setup complete. You can now run 'app.py' in your Flask environment.")

