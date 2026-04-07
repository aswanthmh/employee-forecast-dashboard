import gspread
from google.oauth2.service_account import Credentials

scopes = [
    "https://www.googleapis.com/auth/spreadsheets"
]
creds = Credentials.from_service_account_file("credentials.json",scopes=scopes)
client = gspread.authorize(creds)

sheet_id="1akuaXxI5j4lzQTPj8F8fFNXQxThFjUeupSbNazCn25k"
Workbook =client.open_by_key(sheet_id)
sheets = map(lambda x: x .title, Workbook.worksheets())
print(list(sheets))
