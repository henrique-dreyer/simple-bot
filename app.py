#Ideia geral: o bot entrará na planilha excel do listado dos produtos, selecionará 1 produto por vez, e colará no modulo de movimentacao de estoque de produtos, 1 por vez. Se conseguir, seria interessante colocar a funcionalidade de pause/play no bot.
#Passo a passo
# 1 ir na planilha excel e copiar o nome do primeiro produto da lista
# 2 ir no sistema haupt e colar o nome do primeiro produto no listado
# 3 esperar alguns segundos para que a patricia possa olhar os movimentos
# 4 voltar na planilha excel e copiar o nome do segundo produto da lista
# 5 Repete tudo até que terminem as linhas da planilha (usar while)


import openpyxl
import pyautogui
import pyperclip
import time
planilha_produtos = openpyxl.load_workbook('PRECIOS.xlsx')
sheet_produtos = planilha_produtos['Hoja1']
def haupt():
    pyautogui.click(413,1057,duration=1)

def material():
    pyautogui.click(358,100,duration=1)

#1 copiar o nome do primeiro produto da lista da planilha excel
for coluna in sheet_produtos.iter_rows(min_row=4):
    codigo_produto = str(coluna[0].value)
    pyperclip.copy(codigo_produto)


# 2 ir no sistema haupt e colar o nome do primeiro produto no listado
    haupt()
    material()
    pyautogui.hotkey('esc')
    pyautogui.hotkey('ctrl','v')
    pyautogui.hotkey('enter') 
    time.sleep(0.5)
    pyautogui.hotkey('enter')
    time.sleep(0.5)
    pyautogui.hotkey('enter')

    pyautogui.hotkey('enter')
    time.sleep(3.0)
    confirmacao = pyautogui.confirm('Gata, quer seguir para o proximo produto? :)', buttons=['Sim', 'Nao'])
    if confirmacao == 'Sim':
        pass
    else:
        break