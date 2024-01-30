#restauração dos cursos EJA
import sys
import time
import pyautogui as pa
import pytesseract as pyts
import pymsgbox as pmsg
import os
from PIL import Image
import pyperclip as pclip
import re

programa_encerrado = 'Obrigado. Programa encerrado'
encerrar_programa = 'Encerrar programa'
login_proxy = 'Login Proxy - [PTFE v1.0]'
PTFE = 'PTFE v1.0'
desenho = """
⡏⠉⠉⠩⢭⣉⣩⣭⣉⠉⠉⠉⠉⢹⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿
⡇⠄⠄⣠⠆⣿⣿⢤⣽⢷⣠⣤⡇⢸⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿
⡇⠄⣼⠃⠄⣿⠉⢸⡇⣽⢃⡀⠣⢸⡇⠄⠄⠄⠄⠄⠄⠄⠄⠄⠄⠄⠄⢸
⡇⠈⢿⡄⠄⣿⠦⣼⡻⠶⣾⠄⠄⢸⡇⠄⠄⠄⠄⠄⠄⠄⠄⠄⠄⠄⠄⢸
⡇⠄⠈⠳⢦⣟⣠⢸⡱⡄⠋⠄⠄⢸⣷⣶⣶⣶⣶⣶⣶⣶⣶⣶⣶⣶⣶⣾
⡇⠄⠄⠄⠼⠻⠄⠘⠓⠙⠛⠄⠄⢸⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿
⡏⠉⠉⠉⠉⠉⠉⠉⠉⠉⠉⠉⠉⠉⠉⠉⠉⠉⠉⠉⠉⠉⠉⠉⠉⠉⠉⢹
⣧⠄⠄⠄⠄⠄⠄⠄⠄⠄⠄⠄⠄⠄⠄⠄⠄⠄⠄⠄⠄⠄⠄⠄⠄⠄⠄⡼
⢸⣶⣶⣶⣶⣶⣶⣶⣶⣶⣶⣶⣶⣶⣶⣶⣶⣶⣶⣶⣶⣶⣶⣶⣶⣶⣶⡇
⠈⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⠁
⠄⠸⡟⠛⠛⠛⠛⠛⠛⠛⠛⠛⠛⠛⠛⠛⠛⠛⠛⠛⠛⠛⠛⠛⠛⢻⠇⠄
⠄⠄⠹⡄⠄⠄⠄⠄⠄⠄⠄⠄⠄⠄⠄⠄⠄⠄⠄⠄⠄⠄⠄⠄⢠⠏⠄⠄
⠄⠄⠄⠙⣦⣤⣤⣤⣤⣤⣤⣤⣤⣤⣤⣤⣤⣤⣤⣤⣤⣤⣤⣴⠋⠄⠄⠄
⠄⠄⠄⠄⠈⢻⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⣿⡟⠁⠄⠄⠄⠄
⠄⠄⠄⠄⠄⠄⠙⠿⡛⠛⠛⠛⠛⠛⠛⠛⠛⠛⠛⢛⠟⠋⠄⠄⠄⠄⠄⠄
⠄⠄⠄⠄⠄⠄⠄⠄⠈⠳⢤⡀⠄⠄⠄⠄⢀⡤⠞⠁⠄⠄⠄⠄⠄⠄⠄⠄
⠄⠄⠄⠄⠄⠄⠄⠄⠄⠄⠄⠉⠲⢤⡤⠖⠉⠄⠄⠄⠄⠄⠄⠄⠄⠄⠄⠄
"""
#time.sleep(3) #apagar

def mult_img_selection(img_list, confidence_level):
    # Seleção entre diferentes imagens possíveis presentes na tela
    for i in img_list:
        try:
            pos_arq = pa.locateCenterOnScreen(i, confidence=confidence_level)
            return pos_arq  # Retorna a posição se a imagem for encontrada
        except pa.ImageNotFoundException:
            continue
    pmsg.alert(text='ALERTA. Nenhuma imagem foi encontrada: \nReveja o código.', title=PTFE)
    sys.exit()

def send_file(file_path):
    # envia um arquivo no menu de seleção

    # parte web
    pos_choose_file = pa.locateCenterOnScreen('media/escolha-um-arquivo.png', confidence=0.9)
    pa.click(pos_choose_file, duration=.2)
    time.sleep(.5)  # pausa para abrir o menu de seleção de arquivo

    # diferenciação entre as imagens. Se essa opção de envio tiver sido selecionada na exec. anterior, ela vai ficar azul (como na img2 "hover")
    pos_arq = mult_img_selection(['media/enviar-um-arquivo-1.png', 'media/enviar-um-arquivo-2.png'], 0.9)
    if pos_arq is not None:
        pa.click(pos_arq, duration=.2)
        time.sleep(.2)  # pausa para carregar o menu
        pos_escolher_arq = pa.locateCenterOnScreen('media/escolher-arquivo.png', confidence=0.7)
        pa.click(pos_escolher_arq, duration=.2)
    else:
        time.sleep(.2)  # pausa para carregar o menu
        pos_escolher_arq = pa.locateCenterOnScreen('media/escolher-arquivo.png', confidence=0.7)
        pa.click(pos_escolher_arq, duration=.2)

    # parte explorador de arquivos
    time.sleep(1)  # pausa para carregar o explorador de arquivos
    user_path = os.path.expanduser("~")
    search = os.path.join(user_path, file_path)
    write_special_char(search)
    pa.press('enter')

    # finalização do menu de seleção do arquivo
    pos_botao_enviar_arq = pa.locateCenterOnScreen('media/pos-enviar-arq.png', confidence=0.9)
    pa.click(pos_botao_enviar_arq, duration=.5)

def img_to_text(img_path):
    #extrai o texto de um print. ex.: img_path = 'media/nomedoarquivo.png'

    tesseract_path = r"programs\tesseract\tesseract.exe"
    pyts.pytesseract.tesseract_cmd = tesseract_path
    image = Image.open(img_path)
    texto = pyts.image_to_string(image)
    return texto

# #abrir configuracoes para coletar nome breve
# pos_screen = pa.locateCenterOnScreen('media/restauracao-botao-configuracoes.png', confidence = 0.9)
# pa.click(pos_screen, duration = .5)
# pos_screen = pa.locateCenterOnScreen('media/restauracao-opcao-configuracoes.png', confidence = 0.9)
# pa.click(pos_screen, duration = .5)
#
# #copiar nome breve
# time.sleep(2) #pausa para a página de configurações carregar
# pos_screen = pa.locateOnScreen('media/restauracao-nome-breve-do-curso.png', confidence = 0.95)
# pa.click(pos_screen[0] + pos_screen.width + 50, pos_screen[1] + 20, duration = .5)
# pa.hotkey('ctrl','a','ctrl', 'c')
# time.sleep(0.5) #pausa para garantir o sucesso da cópia
# nome_breve = pclip.paste()

import re

def identificar_informacoes(text):
    # Define padrões de expressões regulares para montar o respectivo nome do arquivo de backup
    padrao_turma = re.compile(r'EJA')
    padrao_etapa = re.compile(r'_([\d]+ª\sETAPA|[\d]+ªN)')
    padrao_tipo_curso = re.compile(r'_([^_]+)_INEP')
    padrao_ano_semestre = re.compile(r'_(\d+)/(\d+)|(\d{4})$')

    # Procura por correspondências nos padrões no texto fornecido
    turma = re.search(padrao_turma, text)
    etapa = re.search(padrao_etapa, text)
    tipo_curso_match = re.search(padrao_tipo_curso, text)
    ano_semestre = re.search(padrao_ano_semestre, text)

    # Verifica se as correspondências foram encontradas antes de acessar os grupos
    turma_grupo = turma.group() if turma else "EM"
    etapa_grupo = etapa.group(1) if etapa else None
    if etapa_grupo is None:
        pmsg.alert(text=f'ALERTA. ETAPA NÃO ENCONTRADA: \nReveja o código.\n{desenho}', title=PTFE)
        sys.exit() # Erro na etapa
    tipo_curso_grupo = tipo_curso_match.group(1) if tipo_curso_match else "EF"  # Padrão "EF" se não encontrar
    ano_semestre_grupo1 = ano_semestre.group(1) if ano_semestre and ano_semestre.group(1) else ano_semestre.group(3) if ano_semestre else None
    ano_semestre_grupo2 = ano_semestre.group(2) if ano_semestre and ano_semestre.group(2) else None

    # Cria o nome do arquivo de backup conforme as informações extraídas
    nome_arquivo = f"{turma_grupo}_{ano_semestre_grupo1}-{ano_semestre_grupo2}_{etapa_grupo}_{tipo_curso_grupo}"

    # Cria o nome do arquivo de backup conforme as informações extraídas
    if 'EJA' in text:
        nome_arquivo = f"{turma_grupo}_{ano_semestre_grupo1}-{ano_semestre_grupo2}_{etapa_grupo}_{tipo_curso_grupo}"
    else:
        # Aqui você pode definir como você quer que o nome do arquivo seja para strings que não contêm 'EJA'
        nome_arquivo = f"EM_{ano_semestre_grupo1}-_{etapa_grupo}_{tipo_curso_grupo}"

    return nome_arquivo

# Exemplo de uso
nome_breve = "EEEFFRANCISCOALVESMENDES_5ªE6ªMULTIN01EJA-EF_ª ETAPA_PIPAT_INEP32070853_281095_2023/1"
nome_backup = identificar_informacoes(nome_breve)

# Exibe as informações identificadas
print(nome_backup)
