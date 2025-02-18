import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd

# Google Sheets API setup
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/spreadsheets",
         "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name('access-to-sheets-449017-376cfa18b986.json', scope)
client = gspread.authorize(creds)

# Open spreadsheet and worksheet
spreadsheet_name = "Colorado Springs (SKIPTRACED)"
sheet = client.open(spreadsheet_name)
worksheet = sheet.worksheet("testing")  # Change if needed

# Read Google Sheet data into a DataFrame
data = worksheet.get_all_values()
if not data:
    print("❌ No data found in the sheet.")
    exit()

df = pd.DataFrame(data[1:], columns=data[0])  # First row as headers

# Define updated structure
updated_headers = [
    "Full Name", "Address", "City", "State", "ZIP", "Email",
    "Mobile 1", "Mobile 2", "Mobile 3", "Mobile 4", "Mobile 5", "Mobile 6",
    "Landline 1", "Landline 2", "Landline 3", "Landline 4", "Landline 5", "Landline 6", "VOIP",
    "Status", "Date Added"
]

# Ensure all required columns exist in DataFrame
for col in updated_headers:
    if col not in df.columns:
        df[col] = ""  # Add missing columns

df = df[updated_headers]  # Reorder columns based on new structure

# Shift numbers left in phone columns
phone_columns = [
    "Mobile 1", "Mobile 2", "Mobile 3", "Mobile 4", "Mobile 5", "Mobile 6",
    "Landline 1", "Landline 2", "Landline 3", "Landline 4", "Landline 5", "Landline 6", "VOIP"
]

def shift_left(row):
    numbers = [num for num in row if num.strip()]  # Remove empty values
    return numbers + [""] * (len(row) - len(numbers))

df[phone_columns] = df[phone_columns].apply(shift_left, axis=1)

# Clear old data and update with the new structured DataFrame
worksheet.clear()
worksheet.update([df.columns.values.tolist()] + df.values.tolist())

# Apply formatting
worksheet.format('A1:Z1', {'backgroundColor': {'red': 1, 'green': 1, 'blue': 0}})  # Yellow header row

column_widths = {
    "A": 250, "B": 135, "C": 190, "D": 140, "E": 75, "F": 60, "G": 400,
    "H": 85, "I": 85, "J": 85, "K": 85, "L": 85, "M": 85, "N": 85, "O": 85,
    "P": 100, "Q": 120
}

for col, width in column_widths.items():
    worksheet.format(f"{col}:{col}", {"pixelSize": width})

print("✅ Spreadsheet updated with the new structure and cleaned data!")
