import requests
from requests.auth import HTTPBasicAuth
import pandas as pd
import os
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

# Function to get all table names
def get_all_table_names(base_url, auth, headers):
    url = base_url + 'sys_db_object'
    response = requests.get(url, auth=auth, headers=headers)
    if response.status_code == 200:
        data = response.json()
        return [table['name'] for table in data['result']]
    else:
        print(f"Failed to retrieve tables. Status code: {response.status_code}, Response: {response.text}")
        return []

# Function to get data from a table
def get_table_data(base_url, table_name, auth, headers, limit=10000):
    url = f"{base_url}{table_name}?sysparm_limit={limit}"
    response = requests.get(url, auth=auth, headers=headers)
    if response.status_code == 200:
        return response.json()['result']
    else:
        print(f"Failed to retrieve data from {table_name}. Status code: {response.status_code}, Response: {response.text}")
        return []

def main(credentials_file, output_dir, table_limit):
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

    # Ensure the output directory exists
    os.makedirs(output_dir, exist_ok=True)

    # Retrieve all table names
    table_names = get_all_table_names(base_url, auth, headers)

    # Limit the number of tables to process if specified
    if table_limit:
        table_names = table_names[:table_limit]

    # Retrieve data from each table and save to separate CSV files
    for table_name in table_names:
        table_data = get_table_data(base_url, table_name, auth, headers)
        if table_data:
            df = pd.DataFrame(table_data)
            output_file = os.path.join(output_dir, f"{table_name}.csv")
            df.to_csv(output_file, index=False)
            print(f"Data for table '{table_name}' has been written to '{output_file}'.")

    print("Data retrieval and CSV file creation complete.")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Retrieve data from all ServiceNow tables and save to CSV files.")
    parser.add_argument('--credentials', default='credentials.txt', help="Path to the credentials.txt file")
    parser.add_argument('--output_dir', default='./servicenow_tables', help="Directory to save the CSV files")
    parser.add_argument('--table_limit', type=int, help="Limit the number of tables to process")
    args = parser.parse_args()
    
    main(args.credentials, args.output_dir, args.table_limit)
