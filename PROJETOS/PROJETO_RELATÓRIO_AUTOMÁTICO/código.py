#AUTOMAÇÃO DE RELATORIOS E INDICADORES DE VENDAS:

# PASSO - 1: ENTRAR NO SISTEMA DA EMPRESA (EMAIL)
# PASSO - 2: NAVEGAR NO SISTEMA (EMAIL) E ENCONTRAR A BASE DE COMPRAS E VENDAS
# PASSO - 3: FAZER O DOWNLOAD DO ARQUIVO
# PASSO - 4: IMPORTAR O ARQUIVO PARA O PYTHON
# PASSO - 5: CALCULAR OS INDICADORES DO ARQUIVO DA EMPRESA
# PASSO - 6: ENVIAR RELATORIO DOS INDICADORES PARA O EMAIL DA DIRETORIA DA EMPRESA

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

time.sleep(1)

pyperclip.copy("https://drive.google.com/drive/u/1/my-drive")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")

time.sleep(5)

pyautogui.click(x=471, y=622)
pyautogui.press("enter")
pyautogui.click(x=438, y=385)
pyautogui.click(x=1710, y=188)
pyautogui.click(x=1568, y=526)

time.sleep(6)

tabela = pd.read_excel(r"C:\Users\daviz\Downloads\compras e vendas.xlsx")
print(tabela)

faturamento = tabela["Faturamento"].sum()
quantidade = tabela["Vendas"].sum()

print(f"O faturamento de vendas foi de {faturamento:.2f}")
print(f"E a quantidade de vendas foi de {quantidade:.2f}")

pyautogui.hotkey("ctrl", "t")
pyperclip.copy("https://mail.google.com/mail/u/0/#inbox")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")

time.sleep(6)

pyautogui.click(x=143, y=197)
time.sleep(1)
pyautogui.write("davigomestestes@gmail.com")
pyautogui.press("tab")
pyautogui.press("tab")
pyautogui.write("Relatorios e indicadores de vendas")
pyautogui.press("tab")
pyautogui.write(f"""Prezados, segue o relatorio de vendas:
Total de vendas: {quantidade}
Total do faturamento: R${faturamento:.2f}
""")
pyperclip.copy(r"C:\Users\daviz\Downloads\compras e vendas.xlsx")
pyautogui.click(x=1421, y=1011)
pyautogui.click(x=1585, y=953)
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")

time.sleep(3)

pyautogui.click(x=1316, y=1004)
