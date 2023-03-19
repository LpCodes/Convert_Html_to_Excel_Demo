import pandas as pd
import csv
from openpyxl import Workbook
from openpyxl.styles import Font

def read_html_file(html_file):
    # Read the HTML file and return a list of tables
    tables = pd.read_html(html_file)
    return tables

def write_csv_file(csv_file, tables):
    # Write the tables to a CSV file with the table name as the first cell of each row
    with open(csv_file, 'w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        for table in tables:
            # Add an empty row before each table
            writer.writerow([])
            # Add the table name as the first cell of the row
            writer.writerow([table.columns.name])
            # Write the table data to the CSV file
            writer.writerows(table.values)
    
def modify_excel_file(excel_file):
    # Create a new Workbook object
    wb = Workbook()
    # Get the active sheet
    sheet = wb.active
    # Open the CSV file and read its contents
    with open(excel_file, 'r', encoding='utf-8') as f:
        reader = csv.reader(f)
        # Add each row of the CSV file to the sheet
        for row in reader:
            sheet.append(row)
    # Loop through the cells in the first row
    for cell in sheet[1]:
        # If the cell contains "Test_Cases" or "Status", make the text bold
        if "Test_Cases" in str(cell.value) or "Status" in str(cell.value):
            cell.font = Font(bold=True)
    # Save the modified workbook to the same file
    wb.save(excel_file)

def main():
    # Set the file paths
    html_file = "./Test Report_2021-08-18_12-45-00.html"
    csv_file = "./your_csv_name.csv"
    excel_file = "name.xlsx"

    # Task 1: Convert the HTML file to CSV format
    print("Starting task one")
    tables = read_html_file(html_file)
    write_csv_file(csv_file, tables)
    print("Task one over")

    # Task 2: Modify the Excel file
    print("Starting task two")
    modify_excel_file(excel_file)
    print("Task two over")

if __name__ == '__main__':
    main()
