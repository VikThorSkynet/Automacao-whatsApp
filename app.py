import openpyxl
from urllib.parse import quote
import webbrowser 
from time import sleep
import pyautogui

workbook = openpyxl.load_workbook('clientes.xlsx')
pagina_clientes = workbook['pagina1']

for linha in pagina_clientes.iter_rows(min_row=2):
    nome = linha[0].value
    telefone = linha[1].value
    if nome and telefone:
        mensagem = f'Olá {nome}, consegui '
    
        link_msg_zap = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
        sleep(15)
        webbrowser.open(link_msg_zap)
        sleep(10)
        seta = pyautogui.locateOnScreen('image.png')
        if seta is not None:
            # Obtém as coordenadas e dimensões da imagem da seta
            x, y, largura, altura = seta
            
            # Calcula o centro da seta
            centro_x = x + largura / 2
            centro_y = y + altura / 2
            
            # Clica no centro da seta
            pyautogui.click(centro_x, centro_y)
        sleep(5)
        pyautogui.hotkey('Ctrl', 'w')
        sleep(5)
    else:
        print(f"Contato inválido na linha: {linha}")
