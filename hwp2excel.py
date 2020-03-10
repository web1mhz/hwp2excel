import os
from time import sleep

import pandas as pd

import win32clipboard
import win32com.client as win32

# from io import BytesIO
# from PIL import Image # !pip install Pillow
# from openpyxl import Workbook

hwp_name = rf"C:\exe2hwp\test.hwp"

# # 한글 빈문서를 배경으로 불러오기
# hwp= win32.Dispatch("HWPFrame.HwpObject")

# # 한글 자동보안승인 요청
# hwp.RegisterModule("FilePathCheckDLL","SecurityModule")

hwp = win32.gencache.EnsureDispatch('HWPFrame.HwpObject')
hwp.RegisterModule("FilePathCheckDLL", "SecurityModule")

# hwp.Run("FileNew")
hwp.Open(hwp_name)

# 열린 한글문서를 지정하고자 할때 아래코드 두번째 열린 문서 지정
원본 = hwp.XHwpDocuments.Item(0)

# 한글문서 추가할때

# hwp.XHwpDocuments.Add(False)
# 추가창 = hwp.XHwpDocuments.Item(2)

# # 추가창에 포커스 추고
# 추가창.SetActive_XHwpDocument()

# # 추가창 달기
# 추가창.Close(True)

원본.SetActive_XHwpDocument()
hwp.InitScan()

original_full_text = []

stop_signal = True

while stop_signal:
    signal, text = hwp.GetText()
    original_full_text.append(text)
    if signal == 1:
        break


hwp.ReleaseScan()
길이 = len(original_full_text) / 19
Data = original_full_text  # .split("\r\n")[:-1]


df = pd.DataFrame([Data])
df = df.transpose()[3:-3]

print(len(df))

df.to_csv('test.csv', header=False, index=False, encoding='utf-8-sig')
