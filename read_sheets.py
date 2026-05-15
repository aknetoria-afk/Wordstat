from google.oauth2 import service_account
from googleapiclient.discovery import build

creds = service_account.Credentials.from_service_account_file(
    'wordstat-google.json',
    scopes=['https://www.googleapis.com/auth/spreadsheets']
)
service = build('sheets', 'v4', credentials=creds)
meta = service.spreadsheets().get(spreadsheetId='198_LrEC4b04EuguRXT3MmGSx1E4ENQeQHgezC1YFby8').execute()
for s in meta['sheets']:
    title = s['properties']['title']
    result = service.spreadsheets().values().get(
        spreadsheetId='198_LrEC4b04EuguRXT3MmGSx1E4ENQeQHgezC1YFby8',
        range=f'{title}!1:1'
    ).execute()
    row = result.get('values', [[]])[0]
    print(f'{title}: {row}')
