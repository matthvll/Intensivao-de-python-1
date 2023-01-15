#1 Importando as bibliotecas
import pyautogui
import time
import pyperclip
import pandas as pd
#2 Abrindo o arquivo do relatório
pyautogui.FAILSAFE = True
pyautogui.PAUSE = 2
pyautogui.alert('Abra o seu navegador e aperte em OK')
time.sleep(3)
pyautogui.hotkey('ctrl', 't')
pyperclip.copy("https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga?usp=sharing")
pyautogui.hotkey("ctrl", "v")
pyautogui.press('enter')
time.sleep(5)
#3 Exportando o relatório
pyautogui.click(x=343, y=302, clicks= 2)
pyautogui.click(x=343, y=302, clicks=1, button='right')
pyautogui.click(x=426, y=778)
time.sleep(5)
#4 Calculando os indicadores  (faturamento e quantidade de produtos)
tabela = pd.read_excel(r"C:/Users/mathe/Downloads/Vendas - Dez.xlsx")
#salvando informações do relatorio em variáveis
faturamento = tabela["Valor Final"].sum()
quantidade = tabela["Quantidade"].sum()
#5 Enviando o resultado pelo email
#abrindo o email
pyautogui.hotkey('ctrl', 't')
pyperclip.copy("https://mail.google.com/mail/u/0/?hl=pt-BR#inbox")
pyautogui.hotkey("ctrl", "v")
pyautogui.press('enter')
time.sleep(5)
pyautogui.click(x=62, y=198)
time.sleep(3)
#Corpo do Email
#inserindo o destinatário do EMAIL:
pyautogui.write('sracarlakisuke@gmail.com')
pyautogui.press('enter')
time.sleep(2)
pyautogui.hotkey('tab')
#inserindo o Assunto do EMAIL:
pyautogui.write('Relatório de vendas')
pyautogui.hotkey('tab')
#Inserindo a mensagem do email
mensagem = f""" Aqui está o nosso relatório geral de vendas
O faturamento de ontem foi de R${faturamento:,.2f}
Quantidade de vendas foi de {quantidade:,}

Abs. Matheus"""
pyperclip.copy(mensagem)
pyautogui.hotkey('ctrl', 'v')
#Anexando Documento do relatório
caminho = (r"C:/Users/mathe/Downloads")
#Clicando em anexar
pyautogui.click(x=1434, y=1002)
#Clicando no caminho do diretório
pyautogui.click(x=657, y=114)
#Colando o caminho do arquivo do relatorio
pyautogui.write(caminho)
pyautogui.press('enter')
#Clicando em 'nome'
time.sleep(2)
pyautogui.click(x=612, y=547)
#Colando o nome do arquivo
pyautogui.write('Vendas - Dez')
#Pressionanto Enter
pyautogui.press('enter')
time.sleep(3)
#Enviar o Email
pyautogui.click(x=1308, y=1010)