import pyautogui,time,pyperclip
import pandas as pd

import pyautogui
import pyperclip
import time

pyautogui.PAUSE = 1

# Passo 1 - Entrar no nosso sistema 
pyautogui.hotkey("ctrl", "t")
link = "https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga?usp=sharing"
pyperclip.copy(link)
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")

# Passo 2 - Navegar no sistema (entrar na pasta Exportar)
time.sleep(5) # espera 5 segundos
pyautogui.click(x=478, y=317, clicks=2)

# Passo 3 - baixar o arquivo de vendas
time.sleep(2)
pyautogui.click(x=524, y=370)
pyautogui.click(x=1110, y=182)
pyautogui.click(x=968, y=724)
time.sleep(3) # esperar ele fazer o download

# Passo 4 - Calcular faturamento e quantidade de produtos vendidos

tabela = pd.read_excel(r"C:\Users\alonp\Downloads\Vendas - Dez.xlsx")
faturamento = tabela["Valor Final"].sum()
quantidade = tabela["Quantidade"].sum()
print(tabela)

# Passo 5 - Enviar o email para a diretoria
pyautogui.hotkey("ctrl", "t") # abre uma nova aba
link = "https://mail.google.com/"
pyperclip.copy(link)
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")

# clicar no botão escrever
time.sleep(5)
pyautogui.click(x=95, y=185)

# escrever pra quem eu to mandando o email
pyautogui.write("pythonimpressionador+diretoria@gmail.com")
pyautogui.press("tab") # escolher o email
pyautogui.press("tab") # passar pro campo de assunto

# escrever o assunto
assunto = "Relatório de Vendas"
pyperclip.copy(assunto)
pyautogui.hotkey("ctrl", "v")
pyautogui.press("tab") # passar pro corpo do email

# escrever o corpo do email
texto = f"""
Prezados, bom dia

O faturamento foi de R${faturamento:,.2f}
A quantidade de produtos foi de {quantidade:,}

Abs
Lira Python"""
pyautogui.write(texto)

# enviar o email
pyautogui.hotkey("ctrl", "enter")