import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/spreadsheets",
         "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name('access-to-sheets-449017-376fa18b86.json', scope)
client = gspread.authorize(creds)
spreadsheet_name = "Colorado Springs (SKIPTRACED)"
sheet = client.open(spreadsheet_name)
worksheet = sheet.worksheet("testing")
data = worksheet.get_all_values()
if not data:
    print("No data found in the sheet.")
    exit()  # Exit the program
df = pd.DataFrame(data[1:], columns=data[0])
updated_headers = [
    "Full Name", "Address", "City", "State", "ZIP", "Email",
    "Mobile 1", "Mobile 2", "Mobile 3", "Mobile 4", "Mobile 5", "Mobile 6",
    "Landline 1", "Landline 2", "Landline 3", "Landline 4", "Landline 5", "Landline 6", "VOIP",
    "Status", "Date Added"
]
for col in updated_headers:
    if col not in df.columns:
        df[col] = ""

df = df[updated_headers]
phone_columns = [
    "Mobile 1", "Mobile 2", "Mobile 3", "Mobile 4", "Mobile 5", "Mobile 6",
    "Landline 1", "Landline 2", "Landline 3", "Landline 4", "Landline 5", "Landline 6", "VOIP"
]
def shift_left(row):
    numbers = [num for num in row if num.strip()]
    return numbers + [""] * (len(row) - len(numbers))

df[phone_columns] = df[phone_columns].apply(shift_left, axis=1)
worksheet.clear()
worksheet.update([df.columns.values.tolist()] + df.values.tolist())

worksheet.format('A1:Z1', {'backgroundColor': {'red': 1, 'green': 1, 'blue': 0}})
column_widths = {
    "A": 250, "B": 135, "C": 190, "D": 140, "E": 75, "F": 60, "G": 400,
    "H": 85, "I": 85, "J": 85, "K": 85, "L": 85, "M": 85, "N": 85, "O": 85,
    "P": 100, "Q": 120
}

for col, width in column_widths.items():
    worksheet.format(f"{col}:{col}", {"pixelSize": width})

print("Spreadsheet updated with the new structure and cleaned data!")
