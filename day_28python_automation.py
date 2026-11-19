import os
import pandas as pd

# Ask user for folder path
folder_path = input("Paste the folder path containing CSV files: ").strip()

# Check if folder exists
if not os.path.isdir(folder_path):
    print("❌ Folder not found. Please check the path.")
    exit()

# Output Excel file path
output_file = os.path.join(
    folder_path,
    "All_CSV_Converted_Into_Sheets.xlsx"
)

# Find all CSV files
csv_files = [
    file for file in os.listdir(folder_path)
    if file.lower().endswith(".csv")
]

# Check if CSV files exist
if not csv_files:
    print("❌ No CSV files found in the folder.")
    exit()

# Create Excel workbook
with pd.ExcelWriter(output_file, engine="openpyxl") as writer:

    for csv_file in csv_files:

        # Full CSV path
        csv_path = os.path.join(folder_path, csv_file)

        # Read CSV
        df = pd.read_csv(csv_path)

        # Sheet name = CSV filename without .csv
        sheet_name = os.path.splitext(csv_file)[0]

        # Excel sheet name maximum = 31 characters
        sheet_name = sheet_name[:31]

        # Write data to Excel sheet
        df.to_excel(
            writer,
            sheet_name=sheet_name,
            index=False
        )

        print(f"✅ Added: {csv_file} → Sheet: {sheet_name}")

print("\n🎉 All CSV files have been combined successfully!")
print(f"📁 Excel file saved at:\n{output_file}")