import openpyxl
import csv
import pandas as pd

# Define file paths
download_path = "C:/Users/Saifi/Downloads"
excel_path = f"{download_path}/cluster10.xlsx"
csv_path = f"{download_path}/output10.csv"

# Load the Excel workbook
excel = openpyxl.load_workbook(excel_path)

# Select the active sheet
sheet = excel.active

# Write data to CSV
with open(csv_path, 'w', newline="", encoding="utf-8") as file:
    writer = csv.writer(file)
    for row in sheet.iter_rows(values_only=True):  # Extract only values
        writer.writerow(row)

# Read CSV into a Pandas DataFrame
df = pd.read_csv(csv_path)

# Display DataFrame
print(df)
