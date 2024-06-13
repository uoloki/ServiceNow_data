### Instructions for Using the ServiceNow Scripts

#### Step 1: Get Your ServiceNow Credentials

1. **Log in to Your ServiceNow Instance**:
   - Open your web browser and navigate to your ServiceNow instance (e.g., `https://your_instance.service-now.com`).
   - Log in with your administrative credentials.

2. **Ensure You Have the Right Permissions**:
   - Make sure your account has permissions to access the tables and data you want to retrieve via the API.

#### Step 2: Create the Credentials File

1. Create a file named `credentials.txt`.
2. Add your credentials in the following format:

```
   instance=your_instance
   username=your_username
   password=your_password
```

   Replace `your_instance`, `your_username`, and `your_password` with your actual ServiceNow instance details.

#### Step 3: Create the Categories File

1. Create a file named `categories.txt`.
2. Add your table categories and tables in the following format:

```
   Incident Management:incident,incident_task
   Change Management:change_request,change_task
   Configuration Management:cmdb_ci,cmdb_rel_ci
```

   Customize the categories and tables according to your needs.

#### Running the Scripts

Note that if you want to run the scripts with default values, you do not have to include any arguments.

1. **Retrieve Data from ServiceNow and Save to Excel Files**:
   - Open a command prompt or terminal.
   - Run the script with:
     python servicenow_data_extraction.py

2. **Filter Excel Files Based on the Presence of 'Y' Columns**:
   - Ensure the `unfiltered` directory exists with Excel files generated from the first script.
   - Run the script with:
     python filter_excel_files.py --input_dir unfiltered --output_dir filtered

3. **Retrieve Table Names and Descriptions**:
   - Run the script with:
     python retrieve_table_names_and_descriptions.py --credentials credentials.txt --output_file table_names_and_descriptions.xlsx

4. **Download All Tables and Data**:
   - Run the script with:
     python download_servicenow_tables.py --credentials credentials.txt --output_dir servicenow_tables --table_limit 10

   - You can adjust the `--table_limit` argument to control the number of tables processed. If omitted, it will process all tables.

### Summary

1. **Credentials File**: `credentials.txt` for your ServiceNow instance, username, and password.
2. **Categories File**: `categories.txt` to define table categories.
3. **Scripts**:
   - `servicenow_data_extraction.py`: Retrieve data and save to Excel.
   - `filter_excel_files.py`: Filter Excel files based on 'Y' columns.
   - `retrieve_table_names_and_descriptions.py`: Retrieve table names and descriptions.
   - `download_servicenow_tables.py`: Download all tables and data to CSV files.
4. **Running the Scripts**: Use command line with the provided examples to execute each script. Ensure that your CLI is directed to the same folder where scripts are located


