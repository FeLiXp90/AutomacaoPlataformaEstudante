#restauração dos cursos EJA
import sys
import time
import pyautogui as pa
import pytesseract as pyts
import pymsgbox as pmsg
import os
from PIL import Image

programa_encerrado = 'Obrigado. Programa encerrado'
encerrar_programa = 'Encerrar programa'
login_proxy = 'Login Proxy - [PTFE v1.0]'
PTFE = 'PTFE v1.0'

time.sleep(3) #apagar

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

while True:
    #acessando a seção de categorias dos cursos
    pa.press('f6')
    pa.write('http://alunodchm.sedu.es.gov.br/ava/course/index.php') #TODO substituir o link para o de produção
    pa.press('enter')
    time.sleep(2) #pausa para a página carregar

    #abrindo sequência de cursos
    pos_screen = pa.locateOnScreen('media/restauracao-escolha-sala-sem-categoria.png', confidence = 0.7)
    pa.click(pos_screen[0] + 50, pos_screen[1] + pos_screen.height + 10, duration = .5)
    time.sleep(2) #pausa para a página carregar

        #verificação se deve encerrar ou não ao acessar o curso "automatização"
    try:
        pos_verific = pa.locateOnScreen('media/restauracao-imagem-para-encerramento.png', confidence = 0.95)
        if pos_verific is not None:
            pmsg.alert(text=f'Restaurações e categorizações concluídas. \n{programa_encerrado}.', title=PTFE)
            sys.exit() #TODO: antes de sair do programa, deve excluir o automatização
    except pa.ImageNotFoundException:
        pass

    # abrir configuracoes para coletar nome breve
    pos_screen = pa.locateCenterOnScreen('media/restauracao-botao-configuracoes.png', confidence=0.9)
    pa.click(pos_screen, duration=.5)
    pos_screen = pa.locateCenterOnScreen('media/restauracao-opcao-configuracoes.png', confidence=0.9)
    pa.click(pos_screen, duration=.5)

    # copiar nome breve
    time.sleep(2)  # pausa para a página de configurações carregar
    pos_screen = pa.locateOnScreen('media/restauracao-nome-breve-do-curso.png', confidence=0.95)
    pa.click(pos_screen[0] + pos_screen.width + 50, pos_screen[1] + 20, duration=.5)
    nome_breve = pa.hotkey('ctrl', 'a', 'ctrl', 'c')

    #abrindo menu de restauração
    pos_screen = pa.locateCenterOnScreen('media/restauracao-botao-configuracoes.png', confidence = 0.9)
    pa.click(pos_screen, duration = 0.5)
    pos_screen = pa.locateCenterOnScreen('media/restauracao-opcao-restaurar.png', confidence = 0.9)
    pa.click(pos_screen, duration = 0.5)
    time.sleep(2) #pausa para a nova guia carregar

    #envio do arquivo no menu de restauração
    #send_file('dsafsdsgfbdvbcvb')



    #break




#pesquisar por -EJA
#pesquisar pelo ano e semestre
#pesquisar pela etapa
#pesquisar pela variação

