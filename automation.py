# esse projeto tira dados de uma planilha e preenche qualquer formulário desejado de maneira automatizada.

import openpyxl
import pyautogui

workbook = openpyxl.load_workbook('nome do arquivo.xlsx')
page = workbook['página que deseja acessar']

for rows in page.iter_rows(min_row=2): # Nesse caso cada linha tem apenas 3 colunas
    pyautogui.click('coordenadas', duration=1)
    pyautogui.write(rows[0].value)
    pyautogui.click('coordenadas', duration=1)
    pyautogui.write(rows[1].value)
    pyautogui.click('coordenadas', duration=1)
    pyautogui.write(rows[2].value)