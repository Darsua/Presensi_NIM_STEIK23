import cv2
import numpy as np
from pyzbar.pyzbar import decode
import pandas as pd

nimabsen = pd.read_excel('absen.xlsx')
nim = set(nimabsen['NIM'].tolist())
alist = pd.read_excel('list.xlsx')
print(nim)


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

        cv2.putText(frame, string, (x, y), cv2.FONT_HERSHEY_SIMPLEX, 0.8, (255, 0, 0), 2)
        print("Barcode: " + barcode_data + " | Type: " + barcode_type)
        nim.add(barcode_data[:8])
        print(nim)
        df = pd.DataFrame({'NIM': list(nim)})
        with pd.ExcelWriter('absen.xlsx', engine="openpyxl", mode='a', if_sheet_exists='replace') as writer:
            df.to_excel(writer)
        absen = pd.read_excel('absen.xlsx')
        listhadir = pd.DataFrame(
            {'Nama': alist['nama'], 'NIM': alist['NIM'], 'Absen': list(alist['NIM'].isin(absen['NIM']))})
        with pd.ExcelWriter('kehadiran.xlsx', engine="openpyxl", mode='a', if_sheet_exists='overlay') as writer:
            listhadir.to_excel(writer)


cap = cv2.VideoCapture(0)
while True:
    ret, frame = cap.read()
    decoder(frame)
    cv2.imshow('Image', frame)
    code = cv2.waitKey(10)
    if code == ord('q'):
        break
