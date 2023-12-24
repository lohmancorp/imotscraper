import requests
from bs4 import BeautifulSoup
import pandas as pd

def fetch_and_parse_table(url, table_id):
    # Send a GET request to the URL
    response = requests.get(url)
    response.raise_for_status()  # Raises an HTTPError if the HTTP request returned an unsuccessful status code

    # Use BeautifulSoup to parse the HTML content
    soup = BeautifulSoup(response.content, 'html.parser')

    # Find the table with the specified id in the HTML
    table = soup.find('table', id=table_id)

    # Parse the table with pandas
    if table:
        # Use pandas to read the table HTML into a DataFrame
        df = pd.read_html(str(table))[0]
        return df
    else:
        print(f"Table with id='{table_id}' not found.")
        return None

# URL from where to fetch the page content
url = 'https://www.imot.bg/pcgi/imot.cgi?act=14&pn=0&town=%D1%EE%F4%E8%FF&year=2023&date=21.11.2023'

# The ID of the table you're interested in
table_id = 'tableStats'

# Fetch the page content and parse the table
table_df = fetch_and_parse_table(url, table_id)

# Display the parsed DataFrame
if table_df is not None:
    print(table_df)

    # Optionally, save the DataFrame to a CSV file
    table_df.to_csv('table_data.csv', index=False, encoding='utf-8-sig')
