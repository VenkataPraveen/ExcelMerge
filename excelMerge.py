import os
import pandas as pd

def combine_excels(input_dir, output_file):
    # Get a list of all Excel files in the input directory
    excel_files = [f for f in os.listdir(input_dir) if f.endswith('.xlsx')]

    # Check if there are any Excel files in the input directory
    if not excel_files:
        print("No Excel files found in the input directory.")
        return

    # Create an empty list to hold all the dataframes
    all_dataframes = []

    # Iterate over each Excel file and read its contents into a dataframe
    for excel_file in excel_files:
        file_path = os.path.join(input_dir, excel_file)
        df = pd.read_excel(file_path, engine='openpyxl')
        all_dataframes.append(df)

    # Concatenate all dataframes into a single dataframe
    combined_df = pd.concat(all_dataframes, ignore_index=True)

    # Write the combined dataframe to a new Excel file
    combined_df.to_excel(output_file, index=False)

    print("Excel files successfully combined into a single Excel file.")

# Example usage:
input_directory = r'C:\Users\venkbat1\Downloads\PythonScripts\ExcelScript\Input'
output_file = r'C:\Users\venkbat1\Downloads\PythonScripts\ExcelScript\output\combined_file.xlsx'
combine_excels(input_directory, output_file)
