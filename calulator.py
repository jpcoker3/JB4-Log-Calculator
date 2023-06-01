import os
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment
import math

def round_rpm(rpm):
    if math.isnan(rpm):
        return rpm
    return int(rpm // 500) * 500

def average_duplicate_rows(df, unique_cols):
    # Group the dataframe by unique columns and calculate the mean for other columns
    grouped = df.groupby(unique_cols).mean().round(2) #rounds to 2 decimal places
    return grouped

def create_averages(file_path, sheet_name, file_format):
    # Read the file based on the format (CSV or XLSX)
    if file_format == 'csv':
        df = pd.read_csv(file_path, header=4)
    elif file_format == 'xlsx':
        df = pd.read_excel(file_path, sheet_name=sheet_name, header=4)
    else:
        raise ValueError("Invalid file format. Only 'csv' and 'xlsx' are supported." +
                         "\n Received: {}".format(file_format))

    # Convert the 'boost' and 'boost2' columns to numeric, skipping non-numeric values
    df['Chargepipe Boost'] = pd.to_numeric(df['boost'], errors='coerce')
    df['Manifold Boost'] = pd.to_numeric(df['boost2'], errors='coerce')

    # Round RPM values down to the nearest 500
    df['rounded_rpm'] = df['rpm'].apply(round_rpm)
    

    # Calculate the average values per rounded RPM per each gear
    averages = df.groupby(['map', 'gear', 'rounded_rpm']).mean().round(2)

    # Select the desired columns
    columns = ['ecu_psi', 'Chargepipe Boost', 'Manifold Boost', 'throttle', 'afr', 'mph']
    averages = averages[columns]

    return averages

def create_calcs(folder_path, calculation_file):
    # Create an empty dictionary to store the files and sheet names
    files_and_sheet_names = {}

    # Iterate over the files in the folder
    for filename in os.listdir(folder_path):
        if filename.endswith((".xlsx", ".csv")):
            file_path = os.path.join(folder_path, filename)
            if filename.endswith(".xlsx"):
                file_format = "xlsx"
                wb = load_workbook(file_path, read_only=True)
                sheet_names = wb.sheetnames
            else:
                file_format = "csv"
                sheet_names = [None]  # For CSV files, use a placeholder sheet name
            files_and_sheet_names[file_path] = sheet_names


    # Create an empty list to store all the averages
    all_averages = []

    # Iterate over the dictionary items
    for excel_file, sheet_names in files_and_sheet_names.items():
        for sheet_name in sheet_names:
            # Get the file format from the file extension
            file_format = excel_file.split('.')[-1]

            # Get the averages for the current file and sheet
            averages = create_averages(excel_file, sheet_name, file_format)

            # Append the averages to the overall list
            all_averages.append(averages)

    # Concatenate all the averages into a single DataFrame
    all_averages_df = pd.concat(all_averages)
    all_averages_df = average_duplicate_rows(all_averages_df, ['map', 'gear', 'rounded_rpm'])

    # Write the DataFrame to the calculation file
    with pd.ExcelWriter(calculation_file, mode='a', engine='openpyxl') as writer:
        # Check if the 'Averages' sheet already exists in the Excel file
        if 'Averages' in writer.book.sheetnames:
            # Delete the existing 'Averages' sheet
            writer.book.remove(writer.book['Averages'])
            writer.book.save(calculation_file)

        # Write the DataFrame to a new 'Averages' sheet
        all_averages_df.to_excel(writer, sheet_name='Averages')
        writer.book._sheets = [writer.book['Averages']] + [sheet for sheet in writer.book._sheets if sheet.title != 'Averages']

        worksheet = writer.sheets['Averages']
        for column in worksheet.columns:
            column_width = 15
            worksheet.column_dimensions[column[0].column_letter].width = column_width

        for row in worksheet.iter_rows():
            for cell in row:
                cell.alignment = Alignment(horizontal='center', vertical='center')

        writer.book.save(calculation_file)  # Save the workbook

    # Apply conditional formatting to the 'afr' column after saving the workbook
    book = load_workbook(calculation_file)
    sheet = book['Averages']

    column_location = 8  # 8 to account for the first 3 columns
    for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row, min_col=column_location, max_col=column_location):
        for cell in row:
            fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
            if cell.value is not None and 12.5 <= cell.value <= 13:
                fill = PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid")
            cell.fill = fill

    # Save the workbook again to apply the conditional formatting
    book.save(calculation_file)

def main():
    calculation_file = r"C:\Users\jpcok\Documents\CarStuff\Tiguan\JB4\Calculations.xlsx"
    # Load the log path
    log_path = r"C:\Users\jpcok\Documents\CarStuff\Tiguan\JB4\Logs"

    create_calcs(log_path, calculation_file)

if __name__ == "__main__":
    main()
