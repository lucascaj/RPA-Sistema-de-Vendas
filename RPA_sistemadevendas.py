import pyautogui
import os
import pandas as pd
from time import sleep
import mimetypes
import pyperclip
import sys

os.chdir(os.getcwd())

botao1 = 'entrar.PNG'
botao2 = 'sistemadevendas.PNG'
botao3 = "+produto.PNG"
botao4 = "enviar.PNG"
botao5 = "botao_ok.PNG"
cmp_produto = "cmp_produto.PNG"
cmp_qtd = "cmp_qtd.PNG"
arquivo = "Vendas.xlsx"

if mimetypes.guess_type(arquivo)[0] == 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet':
    dados = pd.read_excel(arquivo, engine='openpyxl')  # XLS

elif mimetypes.guess_type(arquivo)[0] == 'application/vnd.ms-excel':
    dados = pd.read_excel(arquivo)  # XLS

elif mimetypes.guess_type(arquivo)[0] == 'text/csv':
    dados = pd.read_csv(arquivo)  # CSV

else:
    raise Exception("File not supported")

login = 'lucasadm'
senha = 'lucas123'

screenWidth, screenHeight = pyautogui.size()
widthScale = 1920/screenWidth
heightScale = 1080/screenHeight

cmp_login = (915, 578)
cmp_senha = (915, 613)
cmp_codvenda = (796, 432)
cmp_data = (926, 432)
cmp_loja = (1078, 432)

pyautogui.moveTo(cmp_login, duration=0.25)
pyautogui.click()
sleep(1)
pyautogui.typewrite(login, interval=0.25)
pyautogui.moveTo(cmp_senha, duration=0.25)
pyautogui.click()
sleep(1)
pyautogui.typewrite(senha, interval=0.25)
pyautogui.moveTo(pyautogui.locateCenterOnScreen(
    botao1, confidence=0.9), duration=0.25)
pyautogui.click()

k = None

while k is None:

    k = pyautogui.locateCenterOnScreen(botao2)

pyautogui.moveTo(k, duration=0.5)
pyautogui.click()

i = 0

while i < len(dados.Data):

    dadoslist = dados.loc[dados['Código Venda'] == dados['Código Venda'][i]]
    count = len(dadoslist)

    data = dados.iloc[i]['Data']
    data_format1 = data.strftime("%d")
    data_format2 = data.strftime("%m")
    data_format3 = data.strftime("%Y")

    pyautogui.moveTo(cmp_codvenda, duration=0.5)
    pyautogui.click()
    pyautogui.typewrite(str(dados.iloc[i]['Código Venda']), interval=0.25)
    pyautogui.moveTo(cmp_data, duration=0.5)
    pyautogui.click()
    pyautogui.typewrite(data_format1, interval=0.25)
    pyautogui.typewrite(data_format2, interval=0.25)
    pyautogui.typewrite(data_format3, interval=0.25)
    pyautogui.moveTo(cmp_loja, duration=0.5)
    pyautogui.click()
    loja = dados.iloc[i]['ID Loja']
    pyperclip.copy(loja)
    pyautogui.hotkey('ctrl', 'v', interval=0.1)

    if(len(dados.loc[dados['Código Venda'] == dados['Código Venda'][i]]) > 1):

        if(count <= 4):

            for n in range(0, count):

                pyautogui.moveTo(pyautogui.locateCenterOnScreen(
                    cmp_produto, confidence=0.9), duration=0.5)
                pyautogui.click()
                produto = dadoslist.iloc[n]['Produto']
                pyperclip.copy(produto)
                pyautogui.hotkey('ctrl', 'v', interval=0.1)

                pyautogui.move(200, 0)
                pyautogui.click()
                pyautogui.typewrite(str(dadoslist.iloc[n]['Quantidade']))

                if(n+1 < count):

                    pyautogui.moveTo(pyautogui.locateCenterOnScreen(
                        botao3, confidence=0.7), duration=0.5)
                    pyautogui.click()

                else:

                    pyautogui.scroll(-220)
                    pyautogui.press('enter')
                    pyautogui.moveTo(pyautogui.locateCenterOnScreen(
                        botao4, confidence=0.7), duration=0.5)
                    pyautogui.click()
                    sleep(0.2)
                    pyautogui.press('enter')
                    pyautogui.press('f5')

                    print(f"{i}, {count}")

                    if(i >= 7):

                        sys.exit()

                    i += count

        else:

            for n in range(0, 4):

                pyautogui.moveTo(pyautogui.locateCenterOnScreen(
                    cmp_produto, confidence=0.9), duration=0.5)
                pyautogui.click()
                produto = dadoslist.iloc[n]['Produto']
                pyperclip.copy(produto)
                pyautogui.hotkey('ctrl', 'v', interval=0.1)

                pyautogui.move(200, 0)
                pyautogui.click()
                pyautogui.typewrite(str(dadoslist.iloc[n]['Quantidade']))

                pyautogui.moveTo(pyautogui.locateCenterOnScreen(
                    botao3, confidence=0.7), duration=0.5)
                pyautogui.click()

            for n in range(0, count-4):

                pyautogui.moveTo(pyautogui.locateCenterOnScreen(
                    cmp_produto, confidence=0.9), duration=0.5)
                pyautogui.click()
                produto = dadoslist.iloc[n+4]['Produto']
                pyperclip.copy(produto)
                pyautogui.hotkey('ctrl', 'v', interval=0.1)

                pyautogui.move(200, 0)
                pyautogui.click()
                pyautogui.typewrite(str(dadoslist.iloc[n+4]['Quantidade']))

                if(n+1 < count-4):

                    pyautogui.scroll(-220)
                    pyautogui.moveTo(pyautogui.locateCenterOnScreen(
                        botao3, confidence=0.7), duration=0.5)
                    pyautogui.click()

                else:

                    pyautogui.scroll(-220)
                    pyautogui.press('enter')
                    pyautogui.moveTo(pyautogui.locateCenterOnScreen(
                        botao4, confidence=0.7), duration=0.5)
                    pyautogui.click()
                    sleep(0.2)
                    pyautogui.press('enter')
                    pyautogui.press('f5')

                    if(i >= 7):

                        sys.exit()

                    i += count

    elif(len(dados.loc[dados['Código Venda'] == dados['Código Venda'][i]]) == 1):

        pyautogui.moveTo(pyautogui.locateCenterOnScreen(
            cmp_produto, confidence=0.9), duration=0.5)
        pyautogui.click()
        produto = dados.iloc[i]['Produto']
        pyperclip.copy(produto)
        pyautogui.hotkey('ctrl', 'v', interval=0.1)
        pyautogui.move(200, 0)
        pyautogui.click()
        pyautogui.typewrite(str(dados.iloc[i]['Quantidade']))

        pyautogui.scroll(-220)
        pyautogui.press('enter')
        pyautogui.moveTo(pyautogui.locateCenterOnScreen(
            botao4, confidence=0.7), duration=0.5)
        pyautogui.click()
        sleep(0.2)
        pyautogui.press('enter')
        pyautogui.press('f5')

        print(f"{i}, {count}")

        if(i >= 7):

            sys.exit()

        i += 1
