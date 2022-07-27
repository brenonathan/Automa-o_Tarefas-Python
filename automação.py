import pyautogui as py
import pandas as pd
import time
import pyperclip as pyp

# abrindo navegador
py.hotkey('win')
py.write('google chrome')
py.press('enter')

time.sleep(5)

# baixando arquivo excel
py.click(x=351, y=55)
pyp.copy('https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga')
py.hotkey('ctrl', 'v')
py.press('enter')

time.sleep(10)

py.click(x=354, y=303, clicks= 2)

time.sleep(7)

py.click(x=363, y=365)
time.sleep(2)
py.click(x=1155, y=189)
time.sleep(2)
py.click(x=908, y=594)

time.sleep(7)

# lendo tabela

tabela = pd.read_excel(r'C:\Users\breno\Downloads\Vendas - Dez.xlsx')
print(tabela)
faturamento = tabela['Valor Final'].sum()
quantidade = tabela['Quantidade'].sum()

# abrindo email
py.hotkey('ctrl', 't')
py.write('gmail')
py.press('enter')
time.sleep(2)
py.click(x=265, y=314)
# obrigatório estar logado no email

time.sleep(10)

# Enviando email
py.click(x=39, y=200)
time.sleep(5)
pyp.copy('breno.nathanael07@aluno.ifce.edu.br') # email de destino
py.hotkey('ctrl', 'v')
py.press('tab')
pyp.copy('Relação de vendas')
py.hotkey('ctrl', 'v')
py.press('tab')
texto = (f'''
Prezados, bom dia

O faturamento de ontem foi de: R${faturamento:,.2f}
A quantidade de produtos foi de: {quantidade:,}

Abs
BrnPython
''')
pyp.copy(texto)
py.hotkey('ctrl', 'v')
py.hotkey('ctrl', 'enter')