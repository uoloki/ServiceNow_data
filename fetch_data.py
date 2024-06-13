import requests
from requests.auth import HTTPBasicAuth
import json
import pandas as pd
from openpyxl import load_workbook
import argparse
import os

# Function to read credentials from a file
def read_credentials(filename):
    credentials = {}
    try:
        with open(filename, 'r') as file:
            for line in file:
                if '=' in line:
                    key, value = line.strip().split('=', 1)
                    credentials[key] = value
    except FileNotFoundError:
        print(f"Error: The file {filename} was not found.")
    except Exception as e:
        print(f"Error reading {filename}: {e}")
    return credentials

# Function to read categories from a file
def read_categories(filename):
    categories = {}
    try:
        with open(filename, 'r') as file:
            for line in file:
                if ':' in line:
                    category, tables = line.strip().split(':', 1)
                    categories[category] = tables.split(',')
    except FileNotFoundError:
        print(f"Error: The file {filename} was not found.")
    except Exception as e:
        print(f"Error reading {filename}: {e}")
    return categories

# Adjust column widths based on the content
def adjust_column_widths(ws):
    for column in ws.columns:
        max_length = 0
        column = list(column)
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = (max_length + 2)
        ws.column_dimensions[column[0].column_letter].width = adjusted_width

# Function to get all tables
def get_all_tables(base_url, auth, headers):
    url = base_url + 'sys_db_object'
    response = requests.get(url, auth=auth, headers=headers)
    if response.status_code == 200:
        data = response.json()
        return [table['name'] for table in data['result']]
    else:
        print(f"Failed to retrieve tables. Status code: {response.status_code}, Response: {response.text}")
        return []

# Function to get data from a table
def get_table_data(base_url, table_name, auth, headers):
    url = base_url + table_name
    response = requests.get(url, auth=auth, headers=headers)
    if response.status_code == 200:
        return response.json()['result']
    else:
        print(f"Failed to retrieve data from {table_name}. Status code: {response.status_code}, Response: {response.text}")
        return []

def main(credentials_file, categories_file):
    # Read credentials
    credentials = read_credentials(credentials_file)

    # Check if credentials are correctly loaded
    if 'instance' not in credentials or 'username' not in credentials or 'password' not in credentials:
        print("Error: Credentials are missing or improperly formatted in credentials.txt.")
        return

    # ServiceNow instance details
    instance = credentials['instance']
    username = credentials['username']
    password = credentials['password']

    # Base URL for ServiceNow API
    base_url = f'https://{instance}.service-now.com/api/now/table/'

    # Set up the request headers
    headers = {
        'Content-Type': 'application/json',
        'Accept': 'application/json'
    }

    auth = HTTPBasicAuth(username, password)

    # Read categories
    categories = read_categories(categories_file)

    # Check if categories are correctly loaded
    if not categories:
        print("Error: Categories are missing or improperly formatted in categories.txt.")
        return

    # Ensure the unfiltered directory exists
    output_dir = 'unfiltered'
    os.makedirs(output_dir, exist_ok=True)

    # Retrieve all tables
    all_tables = get_all_tables(base_url, auth, headers)

    # Categorize tables
    categorized_tables = {category: [] for category in categories}
    for table in all_tables:
        for category, table_list in categories.items():
            if table in table_list:
                categorized_tables[category].append(table)
                break

    # Retrieve data from each table and save to separate Excel files
    for category, tables in categorized_tables.items():
        output_path = os.path.join(output_dir, f'{category}.xlsx')
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            for table in tables:
                rows = get_table_data(base_url, table, auth, headers)
                df = pd.DataFrame(rows)

                # Create new columns with "_Y" suffix and 'Y' in the first row
                new_columns = {f'{col}_Y': ['Y'] + [''] * (len(df) - 1) for col in df.columns}
                new_df = pd.DataFrame(new_columns)

                # Concatenate the original dataframe with the new columns
                df = pd.concat([df, new_df], axis=1)

                df.to_excel(writer, sheet_name=table, index=False)

                # Load the worksheet to adjust column widths
                workbook = writer.book
                worksheet = workbook[table]
                adjust_column_widths(worksheet)

    print("Data retrieval, Excel file creation, and column adjustment complete.")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Retrieve data from ServiceNow and save to Excel files.")
    parser.add_argument('--credentials', default='credentials.txt', help="Path to the credentials.txt file")
    parser.add_argument('--categories', default='categories.txt', help="Path to the categories.txt file")
    args = parser.parse_args()
    
    main(args.credentials, args.categories)
