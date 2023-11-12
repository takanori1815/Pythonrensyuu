import pandas as pd
import tkinter as tk
import os
import matplotlib.pyplot as plt

# 1. datファイルを読み込む
dat_file = "/mnt/c/users/SugiyamaTakaki/Desktop/FEKO2/Nomalisemag10Nearfield.dat"  # ファイル名を適切に置き換えてください
df = pd.read_csv(dat_file, delimiter='\t')  # ファイルの区切り文字を適切に指定してください


excel_file = "output_data.xlsx"  # 出力Excelファイル名
sheet_name = "Sheet1"  # シート名を適切に指定してください
df.to_excel(excel_file, sheet_name=sheet_name, index=False)

df_excel = pd.read_excel(excel_file, sheet_name=sheet_name)



