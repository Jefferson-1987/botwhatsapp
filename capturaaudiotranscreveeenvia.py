"""
PRECISO AUTOMATIZAR MINHAS MENSAGENS P/ MEUS CLIENTES GOSTARIA DE SABER VALORES, E GOSTARIA QUE ENTRASSEM EM CONTATO COMIGO P/ EXPLICAR MELHOR, QUERO PODER MANDAR MENSAGENS DE COBRANÇA EM DETERMINADO DIA COM CLIENTES COM VENCIMENTO DIFERENTE
"""
import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui
import os 
import numpy as np
from PIL import ImageGrab, Image
import openai
import base64
from dotenv import load_dotenv, find_dotenv

_ = load_dotenv(find_dotenv())

client = openai.Client()
sleep(3)

img=Image.open('C:/Users/Windows/OneDrive/Documentos/bot_whatsapp/bot_whatsapp/botao_de_audio.jpg') 
imgdl=Image.open('C:/Users/Windows/OneDrive/Documentos/bot_whatsapp/bot_whatsapp/nomearquivo.jpg') 

def find_first_ogg(directory):
    # Lista os arquivos no diretório
    files = os.listdir(directory)
    
    # Filtra apenas os arquivos que têm a extensão .ogg
    ogg_files = [f for f in files if f.endswith('.ogg') and os.path.isfile(os.path.join(directory, f))]
    
    # Verifica se há arquivos .ogg na pasta
    if not ogg_files:
        return None
    
    # Obtém o caminho completo do primeiro arquivo .ogg
    first_ogg_file = os.path.join(directory, ogg_files[0])

    # Retorna o caminho do arquivo apagado
    return first_ogg_file

def encode_image(caminho_imagem):
    with open(caminho_imagem, 'rb' ) as img:
        return base64.b64encode(img.read()).decode('utf-8')

def capture_screen_region(region):
    # Captura uma região específica da tela.
    screenshot = ImageGrab.grab(bbox=region)
    return np.array(screenshot)

def has_new_message(region, baseline_image, threshold=10000):
    """
    Verifica se há uma nova mensagem na região específica da tela comparando com uma imagem de referência.
    Args:
    region (tuple): Um tuple com a forma (left, top, width, height) que define a área a ser capturada.
    baseline_image (numpy array): Imagem de referência para comparação.
    threshold (int): Limite de diferença para considerar que há uma nova mensagem.
    Returns:
    bool: True se há uma nova mensagem, False caso contrário.
    """
    current_image = capture_screen_region(region)
    with Image.fromarray(current_image) as img_to_save:
        img_to_save.save('C:/Users/Windows/OneDrive/Documentos/bot_whatsapp/bot_whatsapp/printdavez.jpg')

    diff = np.sum(np.abs(current_image - baseline_image))
    sleep(5)
    return diff > threshold

def monitor_whatsapp(region, imgcomparacao):
    """
    Monitora a região especificada para detectar novas mensagens e responde automaticamente.
    
    Args:
    region (tuple): Um tuple com a forma (left, top, width, height) que define a área a ser monitorada.
    """
    comparacao = np.array(imgcomparacao)
    print("Iniciando a monitoração do WhatsApp...")
    #baseline_image = capture_screen_region(region)
    while True:
        if has_new_message(region, comparacao):
            print("Nova mensagem de áudio detectada!")
            return True
        sleep(5)

# Defina a região da tela para monitorar (left, top, width, height)
# Você precisará ajustar esses valores de acordo com sua tela e a posição da área de mensagens do WhatsApp Web
message_region = (550, 230, 1200, 360)

message_region_novamsg = (550, 250, 710, 450)
message_region_botaodeaudio = (760, 580, 910, 660)
message_region_fazerdownload = (900, 730, 1090, 780)
message_region_procuradl = (450, 220, 810, 320)

# Inicia a monitoração da região especificada
def procuratranscreveultimoaudio(regiaobotao,img1):
    monitor_whatsapp(regiaobotao, img1)
    pyautogui.click(900,632)
    sleep(0.5)
    pyautogui.click(850,580)
    sleep(0.5)
    #Procura o primeiro arquivo .ogg e manda pro whisper-1
    caminhoprapasta = 'C:/Users/Windows/Downloads'
    try:
        arqaudiostr= find_first_ogg(caminhoprapasta)
    except:
        print("não possui áudio nessa conversa")

    # Envia audio pro whisper
    with open(arqaudiostr, 'rb') as audio:
        transcricao = client.audio.transcriptions.create(
            model='whisper-1',
            file=audio,
            prompt='Esse é um áudio de um cliente do Salão da Sirlei.'
        )
    print(f'\n ---- Esse é o nome do arquivo de audio do wpp: {arqaudiostr}\n\n e a sua transcrição é a seguinte:\
        \n\n---- {transcricao.text}----\n\n')
    sleep(5)
    try:
        texto = transcricao.text
        os.remove(arqaudiostr)
        return texto
    except NameError:
        print("A variável não foi criada.")

        
   

    

def procuranovaconversa():
    
    try:
        nova = pyautogui.locateCenterOnScreen('C:/Users/Windows/OneDrive/Documentos/bot_whatsapp/bot_whatsapp/novamsg.jpg')
        sleep(2)
        pyautogui.click(nova[0],nova[1])
        procuratranscreveultimoaudio(message_region_botaodeaudio, img)
        if nova is None:
            return None
    except:
        i=''
        i=input()



while True:
    if procuranovaconversa() is not None:
        break


'''

# Ler planilha e guardar informações sobre nome, telefone e data de vencimento
workbook = openpyxl.load_workbook('clientes.xlsx')
pagina_clientes = workbook['Sheet1']

for linha in pagina_clientes.iter_rows(min_row=2):
    # nome, telefone, vencimento
    nome = linha[0].value
    telefone = linha[1].value
    vencimento = linha[2].value
    
    mensagem = f'Olá {nome} seu boleto vence no dia {vencimento.strftime('%d/%m/%Y')}. Favor pagar no link https://www.link_do_pagamento.com'

    # Criar links personalizados do whatsapp e enviar mensagens para cada cliente
    # com base nos dados da planilha
    try:
        link_mensagem_whatsapp = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
        webbrowser.open(link_mensagem_whatsapp)
        sleep(10)
        seta = pyautogui.locateCenterOnScreen('seta.png')
        sleep(2)
        pyautogui.click(seta[0],seta[1])
        sleep(2)
        pyautogui.hotkey('ctrl','w')
        sleep(2)
    except:
        print(f'Não foi possível enviar mensagem para {nome}')
        with open('erros.csv','a',newline='',encoding='utf-8') as arquivo:
            arquivo.write(f'{nome},{telefone}{os.linesep}')
    '''
