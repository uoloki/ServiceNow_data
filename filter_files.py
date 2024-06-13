import pandas as pd
import openpyxl
import os

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

def filter_excel_files(input_dir, output_dir):
    # Ensure the output directory exists
    os.makedirs(output_dir, exist_ok=True)

    # Iterate over all Excel files in the input directory
    for filename in os.listdir(input_dir):
        if filename.endswith('.xlsx'):
            input_path = os.path.join(input_dir, filename)
            output_path = os.path.join(output_dir, f'filtered_{filename}')
            
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                # Load the workbook and iterate over all sheets
                workbook = openpyxl.load_workbook(input_path)
                for sheet_name in workbook.sheetnames:
                    df = pd.read_excel(input_path, sheet_name=sheet_name)

                    # Determine columns to keep based on presence of "Y" in the first row of *_Y columns
                    columns_to_keep = []
                    for col in df.columns:
                        if col.endswith('_Y') and df[col].iloc[0] == 'Y':
                            original_col = col[:-2]  # Remove '_Y' to get the original column name
                            columns_to_keep.append(original_col)

                    if columns_to_keep:
                        # Filter the DataFrame to keep only the relevant columns
                        filtered_df = df[columns_to_keep]

                        # Write the filtered DataFrame to the new Excel file
                        filtered_df.to_excel(writer, sheet_name=sheet_name, index=False)
                        
                        # Load the worksheet to adjust column widths
                        workbook = writer.book
                        worksheet = workbook[sheet_name]
                        adjust_column_widths(worksheet)

    print("Filtering and creation of new Excel files complete.")

if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Filter Excel files based on presence of 'Y' in *_Y columns.")
    parser.add_argument('--input_dir', default='unfiltered', help="Directory containing the original Excel files")
    parser.add_argument('--output_dir', default='./filtered', help="Directory to save the filtered Excel files")
    args = parser.parse_args()
    
    filter_excel_files(args.input_dir, args.output_dir)
