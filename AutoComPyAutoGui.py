import pyautogui
import pyperclip
import pandas as pd
from time import sleep

pyautogui.PAUSE = 1
# Abrir navegador
pyautogui.press('win')
pyautogui.write('opera')
pyautogui.press('enter')
# Entrar no sistema
sleep(3)
pyperclip.copy('https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga')
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('enter')
# Navegar no sistema até a base de dados
sleep(3)
pyautogui.click(x=417, y=302, clicks=2)
# Exportar a base de vendas
sleep(2)
pyautogui.rightClick(x=424, y=390)
pyautogui.click(x=623, y=862)
# Calcula indicadores (faturamento e quantidade vendidos)
sleep(2)
tabela = pd.read_excel(r'C:\Users\Matheus\Downloads\Vendas - Dez.xlsx')
faturamento = tabela['Valor Final'].sum()
quantidade = tabela['Quantidade'].sum()
# Entra no email
sleep(3)
pyautogui.hotkey('ctrl', 't')
pyperclip.copy('mail.google.com/mail/u/0/#inbox')
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('enter')
sleep(5)
# Clica no botao escrever
pyautogui.click(x=80, y=202)
# Escrever destinatario
sleep(2)
pyperclip.copy('matheusnetto2003@hotmail.com')
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('tab')
# Escrever o assunto
pyperclip.copy('Relatório de Vendas')
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('tab')
# Escrever o corpo do email
texto = (f"""Prezados, bom dia

O faturamento de ontem foi de: R${faturamento:,.2f}
A quantidade de produtos foi de: {quantidade:,}

Abs
Matheus Gomes
""").replace('.', ',')
pyperclip.copy(texto)
pyautogui.hotkey('ctrl', 'v')
# Enviar o email
sleep(1)
pyautogui.hotkey('ctrl', 'enter')
