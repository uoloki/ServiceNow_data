import requests
from requests.auth import HTTPBasicAuth
import pandas as pd
import openpyxl
import argparse

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

# Function to get all table names and descriptions
def get_table_names_and_descriptions(base_url, auth, headers):
    url = base_url + 'sys_db_object'
    response = requests.get(url, auth=auth, headers=headers)
    if response.status_code == 200:
        data = response.json()
        table_data = [{'name': table['name'], 'description': table['label']} for table in data['result']]
        return table_data
    else:
        print(f"Failed to retrieve tables. Status code: {response.status_code}, Response: {response.text}")
        return []

def main(credentials_file, output_file):
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

    # Get table names and descriptions
    table_data = get_table_names_and_descriptions(base_url, auth, headers)

    # Create DataFrame and save to Excel
    df = pd.DataFrame(table_data)
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Tables', index=False)
        
        # Load the worksheet to adjust column widths
        workbook = writer.book
        worksheet = workbook['Tables']
        adjust_column_widths(worksheet)

    print("Table names and descriptions retrieval and Excel file creation complete.")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Retrieve table names and descriptions from ServiceNow and save to an Excel file.")
    parser.add_argument('--credentials', default='credentials.txt', help="Path to the credentials.txt file")
    parser.add_argument('--output_file', default='table_names.xlsx', help="Output Excel file name")
    args = parser.parse_args()
    
    main(args.credentials, args.output_file)
