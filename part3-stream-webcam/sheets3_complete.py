from __future__ import print_function
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import cv2

#
# Edit Here: The ID of your spreadsheet.
#
SHEET_ID = '{sheetId}'

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
    # Iterate through gif using im.seek(),
    #
    service = build('sheets', 'v4', credentials=creds)
    sheetId = 0
    cam = cv2.VideoCapture(0)

    while(True):
        ret, frame = cam.read()
        x = 0
        requests = []

        for row in frame[0::6]:
            y = 0
            x += 1
            for cell in row[0::25]:
                y += 1
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
                              "red": float(cell[2])/255,
                              "green": float(cell[1])/255,
                              "blue": float(cell[0])/255
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
        response = service.spreadsheets().batchUpdate(
            spreadsheetId=SHEET_ID,
            body=body).execute()

        if cv2.waitKey(1) & 0xFF == ord('q'):
            break

if __name__ == '__main__':
    main()
