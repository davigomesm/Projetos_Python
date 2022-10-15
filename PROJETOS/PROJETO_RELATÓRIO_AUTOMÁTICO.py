import pyautogui
import pandas as pd
import pyperclip
import time
import numpy
import openpyxl
import math

pyautogui.PAUSE = 2

pyautogui.press("win")
pyautogui.write("opera")
pyautogui.press("enter")
pyperclip.copy("https://drive.google.com/drive/u/1/my-drive")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")

time.sleep(5)

pyautogui.click(x=647, y=776)
pyautogui.click(x=1723, y=190)
pyautogui.click(x=1542, y=524)

time.sleep(10)

tabela = pd.read_excel(r"C:\Users\daviz\Downloads\Vendas - Dez.xlsx")
print(tabela)

faturamento = tabela["Valor Final"].sum()
quantidade = tabela["Quantidade"].sum()

print(f"O faturamento de vendas foi de {faturamento:.2f}")
print(f"E a quantidade de vendas doi de {quantidade:.2f}")

pyautogui.hotkey("ctrl", "t")
pyperclip.copy("https://mail.google.com/mail/u/0/#inbox")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")

time.sleep(8)

pyautogui.click(x=143, y=197)
pyautogui.write("davigomestestes@gmail.com")
pyautogui.press("tab")
pyautogui.press("tab")
pyautogui.write("Relatorios e indicadores de vendas")
pyautogui.press("tab")
pyautogui.write(f"""Prezados, segue o relatorio de vendas:
Total de vendas: {quantidade}
Total do faturamento: R${faturamento:.2f}
""")
pyperclip.copy(r"C:\Users\daviz\Downloads\Vendas - Dez.xlsx")
pyautogui.click(x=1421, y=1011)
pyautogui.click(x=1585, y=953)
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")

time.sleep(3)

pyautogui.click(x=1316, y=1004)
