import os
import pickle
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
import cv2
import numpy as np
from pyzbar.pyzbar import decode
import pandas as pd

# fill in directory
file_name = 'google_sheets_credentials.json'
directory = 'place_holder'

file_path = os.path.join(directory, file_name)

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

creds = None
token_file = 'token.pickle'

if os.path.exists(token_file):
    with open(token_file, 'rb') as token:
        creds = pickle.load(token)

if not creds or not creds.valid:
    flow = InstalledAppFlow.from_client_secrets_file(
        file_path, SCOPES)
    creds = flow.run_local_server(port=0)

    with open(token_file, 'wb') as token:
        pickle.dump(creds, token)

service = build('sheets', 'v4', credentials=creds)

# fill in spreadsheet url
spreadsheet_url = 'place_holder'

spreadsheet_id = spreadsheet_url.split('/d/')[-1].split('/edit')[0]

sheet = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()


def decoder(image):
    gray_img = cv2.cvtColor(image, 0)
    barcode = decode(gray_img)
    for obj in barcode:
        points = obj.polygon
        (x, y, w, h) = obj.rect
        pts = np.array(points, np.int32)
        pts = pts.reshape((-1, 1, 2))
        cv2.polylines(image, [pts], True, (0, 255, 0), 3)

        barcode_data = obj.data.decode("utf-8")
        barcode_type = obj.type
        string = "Data " + str(barcode_data) + " | Type " + str(barcode_type)

        cv2.putText(frame, string, (x, y),
                    cv2.FONT_HERSHEY_SIMPLEX, 0.8, (255, 0, 0), 2)
        # print("Barcode: " + barcode_data + " | Type: " + barcode_type)

        new_data = [barcode_data[:8]]

        range_name = 'Sheet1!G:G'
        response = service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range=range_name
        ).execute()
        values = response.get('values', [])

        if new_data not in values:
            values.extend([[new_value] for new_value in new_data])

            request = service.spreadsheets().values().update(
                spreadsheetId=spreadsheet_id,
                range=range_name,
                valueInputOption='RAW',
                body={'values': values}
            )

            response = request.execute()

            row_id = int(new_data[0]) - 19623000 + 1
            range_name = "Sheet1!D" + str(row_id)
            print(range_name)

            request_value = service.spreadsheets().values().update(
                spreadsheetId=spreadsheet_id,
                range=range_name,
                valueInputOption='RAW',
                body={
                    "values": [["HADIR"]]
                }
            )

            color_request = service.spreadsheets().batchUpdate(
                spreadsheetId=spreadsheet_id,
                body={
                    "requests": [
                        {
                            "repeatCell": {
                                "range": {
                                    "sheetId": 0,
                                    "startRowIndex": row_id - 1,
                                    "endRowIndex": row_id,
                                    "startColumnIndex": 3,
                                    "endColumnIndex": 4
                                },
                                "cell": {
                                    "userEnteredFormat": {
                                        "backgroundColor": {
                                            "green": 1.0
                                        }
                                    }
                                },
                                "fields": "userEnteredFormat.backgroundColor"
                            }
                        }
                    ]
                }
            )

            value_response = request_value.execute()
            color_response = color_request.execute()


cap = cv2.VideoCapture(0)
while True:
    ret, frame = cap.read()
    decoder(frame)
    cv2.imshow('Image', frame)
    code = cv2.waitKey(10)
    if code == ord('q'):
        break
