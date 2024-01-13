import gspread
from oauth2client.service_account import ServiceAccountCredentials
from apscheduler.schedulers.blocking import BlockingScheduler

# Google Sheets API credentials
GOOGLE_DRIVE_CREDENTIALS_FILE = 'sshbot-401810-507d88f6b018.json'
scope = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name(GOOGLE_DRIVE_CREDENTIALS_FILE, scope)
client = gspread.authorize(creds)

# Open the spreadsheet by title
spreadsheet_title = 'Applications'
spreadsheet = client.open(spreadsheet_title)

# Specify the sheet and cells to check
sheet_name = 'Лист1'
start_row = 2
end_row = 13
column = 'W'
cells_to_check = [f'{column}{row}' for row in range(start_row, end_row + 1)]

def check_cells():
    sheet = spreadsheet.worksheet(sheet_name)

    for cell in cells_to_check:
        cell_value = sheet.acell(cell).value

        if cell_value:
            print(f"Data exists in {cell}: {cell_value}")
        else:
            print(f"No data found in {cell}")

# Set up the scheduler
scheduler = BlockingScheduler()

# Schedule the job to run every minute
scheduler.add_job(check_cells, 'interval', minutes=1)

# Start the scheduler
try:
    scheduler.start()
except KeyboardInterrupt:
    pass
# import gspread
# from oauth2client.service_account import ServiceAccountCredentials

# # Google Sheets API credentials
# GOOGLE_DRIVE_CREDENTIALS_FILE = 'sshbot-401810-507d88f6b018.json'
# scope = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
# creds = ServiceAccountCredentials.from_json_keyfile_name(GOOGLE_DRIVE_CREDENTIALS_FILE, scope)
# client = gspread.authorize(creds)

# # Open the spreadsheet by title
# spreadsheet_title = 'Applications'
# spreadsheet = client.open(spreadsheet_title)

# # Print the names of all sheets in the spreadsheet
# print("Available sheets in the spreadsheet:")
# for sheet in spreadsheet.worksheets():
#     print(sheet.title)
