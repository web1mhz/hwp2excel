import os

import win32com.client as win32
import pandas as pd

# hwp = win32.Dispatch('HWPFrame.HwpObject')

excel = pd.read_excel('test.xlsx')

print("현재 디렉토리")
print(os.getcwd())
print(1*1, 2**4)

print(excel.head())
print(type(excel))