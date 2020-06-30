from __future__ import print_function
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from PIL import Image

#
# Edit Here: The ID of your spreadsheet.
#
SHEET_ID = '{sheetId}'
IMAGE_FILENAME = 'test.png'

def main():
    #
    # Set up credentials and save in token.pickle
    #
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
    creds = None
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)
    #
    # Grab pixels from image
    #
    service = build('sheets', 'v4', credentials=creds)
    sheetId = 0
    im = Image.open(IMAGE_FILENAME)
    rgb_im = im.convert('RGB')
    rows, cols = im.size
    requests = []
    for x in range(0,rows):
        for y in range(0,cols):
            r, g, b = rgb_im.getpixel((x, y))
            #
            # Add each pixel colour to the request
            #
            requests.append(
                {
                  "repeatCell": {
                    "range": {
                      "sheetId": sheetId,
                      "startRowIndex": x,
                      "endRowIndex": x+1,
                      "startColumnIndex": y,
                      "endColumnIndex": y+1
                    },
                    "cell": {
                      "userEnteredFormat": {
                        "backgroundColor": {
                          "red": (float(r)/255),
                          "green": (float(g)/255),
                          "blue": (float(b)/255)
                        }
                      }
                    },
                    "fields": "userEnteredFormat(backgroundColor)"
                  }
                })
    #
    # Send the request via the Google API
    #
    body = {
        'requests': requests
    }

    if len(requests) > 0:
        response = service.spreadsheets().batchUpdate(spreadsheetId=SHEET_ID,body=body).execute()

if __name__ == '__main__':
    main()
