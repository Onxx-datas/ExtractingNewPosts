import gspread  # Import gspread to interact with Google Sheets
from oauth2client.service_account import ServiceAccountCredentials  # Import for authentication
import pandas as pd  # Import pandas for data manipulation

# Define the scope for accessing Google Sheets and Google Drve
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/spreadsheets",
         "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]

# Load credentials from the JSON key file and authorize access
creds = ServiceAccountCredentials.from_json_keyfile_name('access-to-sheets-449017-376fa18b86.json', scope)
client = gspread.authorize(creds)  # Authenticate with Google Sheets

spreadsheet_name = "Colorado Springs (SKIPTRACED)"  # Name of the Google Spreadsheet
sheet = client.open(spreadsheet_name)  # Open the spreadsheet
worksheet = sheet.worksheet("testing")  # Select the "testing" worksheet

data = worksheet.get_all_values()  # Fetch all data from the sheet
if not data:  # Check if the sheet is empty
    print("❌No data found in the sheet.")  # Print an error message if no data
    exit()  # Exit the program

df = pd.DataFrame(data[1:], columns=data[0])  # Convert data into a DataFrame, using the first row as headers

# Define the expected column headers
updated_headers = [
    "Full Name", "Address", "City", "State", "ZIP", "Email",
    "Mobile 1", "Mobile 2", "Mobile 3", "Mobile 4", "Mobile 5", "Mobile 6",
    "Landline 1", "Landline 2", "Landline 3", "Landline 4", "Landline 5", "Landline 6", "VOIP",
    "Status", "Date Added"
]

# Ensure all expected columns exist in the DataFrame, adding empty ones if missing
for col in updated_headers:
    if col not in df.columns:
        df[col] = ""  # Add missing columns with empty values

df = df[updated_headers]  # Reorder DataFrame columns to match updated_headers

# Define the phone number columns
phone_columns = [
    "Mobile 1", "Mobile 2", "Mobile 3", "Mobile 4", "Mobile 5", "Mobile 6",
    "Landline 1", "Landline 2", "Landline 3", "Landline 4", "Landline 5", "Landline 6", "VOIP"
]

# Function to shift phone numbers left if there are empty spaces
def shift_left(row):
    numbers = [num for num in row if num.strip()]  # Keep only non-empty numbers
    return numbers + [""] * (len(row) - len(numbers))  # Fill the remaining spaces with empty values

df[phone_columns] = df[phone_columns].apply(shift_left, axis=1)  # Apply shift_left to each row in phone columns

worksheet.clear()  # Clear the worksheet before updating with new data
worksheet.update([df.columns.values.tolist()] + df.values.tolist())  # Update the sheet with cleaned data

worksheet.format('A1:Z1', {'backgroundColor': {'red': 1, 'green': 1, 'blue': 0}})  # Set header row background color to yellow

# Define column widths for better readability
column_widths = {
    "A": 250, "B": 135, "C": 190, "D": 140, "E": 75, "F": 60, "G": 400,
    "H": 85, "I": 85, "J": 85, "K": 85, "L": 85, "M": 85, "N": 85, "O": 85,
    "P": 100, "Q": 120
}

# Apply column width formatting
for col, width in column_widths.items():
    worksheet.format(f"{col}:{col}", {"pixelSize": width})  # Set column width

print("✅ Spreadsheet updated with the new structure and cleaned data!")  # Print success message
