import time
import win32com.client
import pandas as pd
import win32com.client as win32
from datetime import datetime, timedelta
import pyautogui as p


File = win32com.client.Dispatch("Excel.Application")
time.sleep(7)

File.Visible = True



Workbook = File.Workbooks.open("Pedidos MITIS_SAP - PRD.v8 atraso não reclamado(v_3) atual. 16.01 (4) (2).xlsx")
print("abriu")
time.sleep(40)

Workbook.RefreshAll()
time.sleep(60)
time.sleep(60)
time.sleep(60)
time.sleep(60)

print("abriu")
Workbook.Save()
print('salvo')
time.sleep(40)



File.Quit()
print('fechou')
time.sleep(40)
