import requests
from bs4 import BeautifulSoup
import pandas as pd
import numpy as np
import argparse
import os
import subprocess
import re
from datetime import datetime

# Set up argument parsing
parser = argparse.ArgumentParser(description='Scrape data from a given URL into an Excel file.')
parser.add_argument('-l', '--url', required=True, help='Base URL of the page to scrape.')
parser.add_argument('-o', '--output', default='scrape', help='Output Excel file name (without extension).')
parser.add_argument('-e', '--excel', action='store_true', help='Open the output file in Excel after creation.')
args = parser.parse_args()

def extract_date_from_url(url):
    match = re.search(r'&date=(\d{2}\.\d{2}\.\d{4})', url)
    if match:
        date_str = match.group(1)
        return datetime.strptime(date_str, '%d.%m.%Y').strftime('%Y-%m-%d')
    else:
        return None

def fetch_and_parse_data(url, type):
    response = requests.get(url)
    response.raise_for_status()
    soup = BeautifulSoup(response.content, 'html.parser')
    table = soup.find('table', id='tableStats')
    if table:
        df = pd.read_html(str(table), header=0, encoding='utf-8')[0]
        df['type'] = type
        return df
    else:
        print(f"Table with id='tableStats' not found.")
        return None

def post_process_dataframe(df, report_date):
    df = df.copy()
    note_indicator = '*Забележка:'
    df = df[~df[df.columns[0]].astype(str).str.contains(note_indicator, regex=False)]
    df.drop(columns=['Unnamed: 0', 'Unnamed: 1', 'Unnamed: 4', 'Unnamed: 7', 'Unnamed: 10'], inplace=True)
    df.columns = ['Region', '1_Bed_Price', '1_Bed_Price_Sqm', '2_Bed_Price', '2_Bed_Price_Sqm',
                  '3_Bed_Price', '3_Bed_Price_Sqm', 'Avg_Price_Sqm']
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

# Main execution
if __name__ == "__main__":
    base_url = args.url
    report_date = extract_date_from_url(base_url)
    
    # Make two requests with different 'pn' values and 'type' values
    sales_url = f"{base_url}&pn=0"
    rent_url = f"{base_url}&pn=1"

    sales_df = fetch_and_parse_data(sales_url, 'sales')
    rent_df = fetch_and_parse_data(rent_url, 'rent')

    if sales_df is not None and rent_df is not None:
        combined_df = pd.concat([sales_df, rent_df], ignore_index=True)

        if report_date:
            output_file_name = f"{report_date} - {args.output}.xlsx"
        else:
            output_file_name = f"{args.output}.xlsx"
        
        writer = pd.ExcelWriter(output_file_name, engine='xlsxwriter')
        combined_df.to_excel(writer, sheet_name='Sheet1', startrow=0, header=True, index=False)
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']
        header_format = workbook.add_format({
            'bold': True,
            'text_wrap': True,
            'valign': 'top',
            'fg_color': '#007ddf',  # blue
            'font_color': '#FFFFFF',  # white
            'border': 1})
        accounting_format = workbook.add_format({'num_format': '€ #,##0.00', 'align': 'right'})
        for col_num in range(1, len(combined_df.columns)):
            worksheet.set_column(col_num, col_num, None, accounting_format)
        worksheet.add_table(0, 0, len(combined_df), len(combined_df.columns) - 1, {
            'columns': [{'header': col_name, 'header_format': header_format} for col_name in combined_df.columns],
            'style': 'Table Style Medium 2'
        })
        writer.close()

        print(combined_df)

        if args.excel:
            open_in_excel(output_file_name)
