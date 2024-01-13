from oauth2client.service_account import ServiceAccountCredentials
from googleapiclient.discovery import build

# Define your credentials and necessary IDs
GOOGLE_DRIVE_CREDENTIALS_FILE = 'sshbot-401810-507d88f6b018.json'
SPREADSHEET_ID = '1onQssJILUERje2KvacbM9rzSkCO5znoMqzpgzGRlLsY'
SHEET_NAME = 'Sheet1'  # Name of the sheet where you want to append data

# Define the data you want to append
values_to_append = [
    ['New Value A1', 'New Value B1', 'New Value C1'],
    ['New Value A2', 'New Value B2', 'New Value C2']
]

# Define the scope and credentials
scope = ['https://www.googleapis.com/auth/spreadsheets']
credentials = ServiceAccountCredentials.from_json_keyfile_name(GOOGLE_DRIVE_CREDENTIALS_FILE, scope)
service = build('sheets', 'v4', credentials=credentials)

try:
    # Retrieve the existing data to determine the next available row
    range_name = f'A:A'  # Adjust the range to the first column of your sheet
    result = service.spreadsheets().values().get(spreadsheetId=SPREADSHEET_ID, range=range_name).execute()
    values = result.get('values', [])
    
    # Calculate the next available row for appending
    next_row = len(values) + 1

    # Set the new range for appending
    append_range = f'A{next_row}'

    # Append the new data
    value_range_body = {
        'values': values_to_append
    }

    request = service.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=append_range,
        valueInputOption='RAW',
        body=value_range_body
    )
    response = request.execute()
    print('Data appended successfully.')
except Exception as e:
    print(f"Error: {e}")
