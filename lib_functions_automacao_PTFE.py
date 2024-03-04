import tkinter as tk
import pymsgbox as pmsg
import sys
import os  # tratamento de dados
from tkinter import filedialog  # tratamento de dados
from openpyxl import load_workbook  # tratamento de dados
import psutil
import pyautogui as pa
import time
import pyperclip as pclip  # upload de arquivo .csv
import re  # restauração e categorização

programa_encerrado = 'Obrigado. Programa encerrado'
encerrar_programa = 'Encerrar programa'
login_proxy = 'Login Proxy - [PTFE v2.0]'
PTFE = 'PTFE v2.0'
imagem_escolher_arquivo = 'media/escolher-arquivo.png'

# tratamento de dados da planilha

def escolher_tipo_planilha():
    """
    Janela para escolha do tipo de planilha que será carregada para criação do dataframe.

    - Autor: Almir
    """
    while True:
        tipo_planilha = pmsg.confirm(text='Escolha o tipo de planilha:',
                                     title=PTFE, buttons=['Somente Usuários', 'Somente Professores',
                                                          'Usuários e Professores', encerrar_programa])

        if tipo_planilha == 'Somente Usuários':
            return 'usuarios'
        elif tipo_planilha == 'Somente Professores':
            return 'professores'
        elif tipo_planilha == 'Usuários e Professores':
            return 'ambos'
        elif tipo_planilha == encerrar_programa:
            pmsg.alert(text=f'{programa_encerrado}', title=PTFE)
            sys.exit()

def abrir_seletor_arquivos():
    """
    Seletor de arquivos. Renomeia a segunda aba das planilhas selecionadas

    - Autor: Almir
    """
    while True:
        # Cria uma instância do tkinter.Tk() para ser a janela principal
        root = tk.Tk()
        root.withdraw()  # Oculta a janela principal

        # Abre o seletor de arquivos e retorna o caminho do arquivo selecionado
        caminho_arquivo = filedialog.askopenfilename()
        workbook = load_workbook(caminho_arquivo)
        aba = workbook.worksheets[1]
        new_name = 'New_Enrolment'
        aba.title = new_name
        workbook.save(caminho_arquivo)


        # Verifica se o usuário selecionou um arquivo ou cancelou a operação
        if caminho_arquivo:
            confirmacao = pmsg.confirm(text=f'Você selecionou o arquivo:\n{caminho_arquivo}\nDeseja continuar?', title='Confirmação de arquivo - [PTFE v1.0]', buttons=['Confirmo', 'Escolher novamente', encerrar_programa])

            if confirmacao == 'Confirmo':
                return caminho_arquivo
            elif confirmacao == encerrar_programa:
                pmsg.alert(text=f'{programa_encerrado}', title=PTFE)
                sys.exit()
        else:
            escolha = pmsg.confirm(text='OPS, nenhum arquivo selecionado.', title='Erro - [PTFE v1.0]', buttons=['Escolher novamente', encerrar_programa])
            if escolha == encerrar_programa:
                pmsg.alert(text=f'{programa_encerrado}', title=PTFE)
                sys.exit()

def nome_curso(shortname):
    """
    Tratamento para definir o fullname com base no shortname.

    Parameters:
        - shortname (str): O nome abreviado do curso, que é uma das colunas da planilha recebida que contém todas as informações do curso.

    - Autor: Felipe
    """
    if "_PIP_" in shortname:
        return "PROJETO DE INTERVENÇÃO/PESQUISA APLICADA À COMUNIDADE"
    elif "_PROINPAC_" in shortname:
        return "PROJETO DE INTERVENÇÃO/PESQUISA APLICADA À COMUNIDADE"
    elif "_PIPAT_" in shortname:
        return "PROJETO INTEGRADOR DE PESQUISA E ARTICULAÇÃO COM O TERRITÓRIO"
    elif "_PVI1_" in shortname:
        return "PRÁTICAS E VIVÊNCIAS INTEGRADORAS I"
    elif "_PVI2_" in shortname:
        return "PRÁTICAS E VIVÊNCIAS INTEGRADORAS II"
    elif "_PV_" in shortname:
        return "PROJETO DE VIDA"
    elif "_EO_" in shortname:
        return "ESTUDO ORIENTADO"
    else:
        return ""



# upload do arquivo .csv



def is_chrome_running():
    """
    Verifica se o chrome está aberto com a utilização da lib psutil.

    Parameters:

    - Autor: Almir
    """
    for process in psutil.process_iter(['pid', 'name']): # pid solicita o processID
        if 'chrome.exe' in process.info['name'].lower():
            return True
    return False

def write_special_char(text):
    """
    Resolve a codificação do pa.write para escrever caracteres especiais, como ~ e ç.

    Parameters:
        - text (str): O texto a ser escrito, que contém os caracteres especiais.

    - Autor: Almir
    """
    pclip.copy(text)
    pa.hotkey('ctrl', 'v')

def mult_img_selection(img_list, confidence_level):
    """
    Seleção entre diferentes imagens possíveis presentes na tela.

    Parameters:
        - img_list (list): Lista de caminhos para as imagens a serem procuradas.
        - confidence_level (float): Nível de confiança para a correspondência da imagem (precisão da busca).

    - Autor: Almir
    """
    for i in img_list:
        try:
            pos_arq = pa.locateCenterOnScreen(i, confidence=confidence_level)
            return pos_arq  # Retorna a posição se a imagem for encontrada
        except pa.ImageNotFoundException:
            continue

def select_drop_down_list_exact(list_file_name, img_list, confidence_level_dropdown, confidence_level_options):
    """
    Clica e seleciona a opção correta das listas suspensas  se não houverem caixas de seleção de mesmo tamanho na tela

    Parameters:
        - list_file_name (str): Nome do arquivo da lista suspensa.
        - img_list (list): Lista de caminhos para as imagens a serem procuradas.
        - confidence_level_dropdown (float): Nível de confiança para identificar a lista suspensa (precisão da busca pelo print da tela).
        - confidence_level_options (float): Nível de confiança para identificar as opções na lista suspensa (precisão da busca pelo print da tela).

    - Autor: Almir
    """
    pos_list = pa.locateCenterOnScreen(list_file_name, confidence = confidence_level_dropdown)
    pa.click(pos_list[0], pos_list[1] - 10, duration = 0.5)
    pos_op = mult_img_selection(img_list, confidence_level_options)
    pa.click(pos_op, duration=.5)
    time.sleep(.1) #para garantir que a lista recolherá a tempo com o click e não ocultará a próxima exec.

def select_drop_down_list_desloc(list_file_name, img_list, confidence_level_dropdown, confidence_level_options):
    """
    Clica e seleciona a opção correta das listas suspensas com base na posição do título à esquerda apenas
    Essa parte do código é sensível ao tamanho da tela, pois localiza apenas o texto
    O print deve ser tirado da extremidade esquerda do título até imediatamente antes da caixa de seleção à direita

    Parameters:
        - list_file_name (str): Nome do arquivo da lista suspensa.
        - img_list (list): Lista de caminhos para as imagens a serem procuradas.
        - confidence_level_dropdown (float): Nível de confiança para identificar a lista suspensa (precisão da busca pelo print da tela).
        - confidence_level_options (float): Nível de confiança para identificar as opções na lista suspensa (precisão da busca pelo print da tela).

    - Autor: Almir
    """
    pos_list = pa.locateOnScreen(list_file_name, confidence = confidence_level_dropdown)
    pa.click(pos_list[0] + pos_list.width + 50, pos_list[1] + 20, duration = .5)
    pos_op = mult_img_selection(img_list, confidence_level_options)
    pa.click(pos_op, duration=.5)
    time.sleep(.1) #para garantir que a lista recolherá a tempo com o click e não ocultará a próxima exec.

def select_drop_down_list_dual(list_file_name, confidence_level_dropdown, direction):
    """
    Seleciona um item de uma lista suspensa que possui apenas duas opções.

    Parameters:
        - list_file_name (str): Nome do arquivo da lista suspensa.
        - confidence_level_dropdow (float): Nível de confiança para identificar a lista suspensa.
        - direction (str): Direção para clicar na lista suspensa ('up' ou 'down').

    - Autor: Almir
    """
    pos_list = pa.locateOnScreen(list_file_name, confidence=confidence_level_dropdown)
    pa.click(pos_list[0] + pos_list.width + 50, pos_list[1] + 20, duration=.5)
    if direction == 'up':
        pa.press('up')
    else:
        pa.press('down')
    pa.press('escape')
    time.sleep(.1)  # para garantir que a lista recolherá a tempo com o click e não ocultará a próxima exec.

def login_verification(nome_arquivo):
    """
    Verifica se as credenciais informadas em qualquer seção de login foram as corretas.

    Parameters:
        - nome_arquivo (str): O nome do arquivo de imagem para verificar a tela.

    - Autor: Almir
    """
    try:
        pos_screen = pa.locateOnScreen(nome_arquivo, confidence=0.7)
        if pos_screen is not None:
            pmsg.alert(text='Erro de digitação, tente novamente.', title='login', button = 'Tentar novamente')
            return True
    except pa.ImageNotFoundException:
        return False

def login_proxy():
    """
    Requisita dados para logar no proxy localmente
    Também cria um loop para verificar se o login foi digitado corretamente

    - Autor: Almir
    """
    time.sleep(1) #pausa para esperar o chrome carregar.
    pa.press('win')
    pa.write('Chrome')
    pa.press('enter')
    time.sleep(1.5) #pausa para o chrome abrir
    while True:
        proxy_user = pmsg.prompt(text='Digite seu nome de usuário', title=login_proxy)
        proxy_password = pmsg.password(text='Digite sua senha', title=login_proxy, mask='*')
        pa.write(proxy_user)
        pa.press('tab')
        pa.write(proxy_password)
        pa.press('enter')
        time.sleep(1.5)  # pausa para verificar se o login foi correto.
        logou = login_verification('media/login-proxy.png')
        if logou is False:
            break

def login_plataforma_estudante():
    """
    Loga na plataforma do estudante

    - Autor: Almir
    """

    #carrega a página da plataforma do estudante.
    pa.hotkey('ctrl', 't')
    pa.write('http://estudante.sedu.es.gov.br/ava/admin/tool/uploadcourse/index.php')
    pa.press('enter')
    time.sleep(2) #pausa para esperar o site carregar

    #verifica se já está logado e digita ou não a senha
    try:
        pos_verific_login = pa.locateCenterOnScreen('media/verificacao-login.png', confidence=0.7)
    except pa.ImageNotFoundException:
        pos_verific_login = None

    if pos_verific_login is None:
        while True:
            ava_user = pmsg.prompt(text='Digite seu nome de usuário. Se já estiver salvo, apenas aperte "enter"', title='Login AVA - [PTFE v1.0]')
            ava_password = pmsg.password(text='Digite sua senha. Se já estiver salvo, apenas aperte "enter"', title='Login AVA - [PTFE v1.0]', mask='*')
            pos_login = pa.locateCenterOnScreen('media/pos-login.png', confidence=0.7)
            pos_login = (pos_login[0], pos_login[1] - 120)
            pa.click(pos_login, duration=.5)
            pa.write(ava_user)
            pa.press('tab')
            pa.write(ava_password)
            pa.press('enter')
            time.sleep(3)  # pausa para verificar se o login foi correto.
            try:
                pos_screen= pa.locateCenterOnScreen('media/verificacao-login.png', confidence=0.7)
                if pos_screen is not None:
                    break
            except pa.ImageNotFoundException:
                pmsg.alert(text='Erro de digitação, tente novamente.', title='login', button='Tentar novamente')
                pa.press('f5')
                pa.press('enter')
                time.sleep(2) #pausa para a página recarregar



#categorização e restauração



def verifica_load(image_path, confidence_level):
    """
    Verifica se o carregamento da página ou recurso foi efetuado com sucesso.
    O objetivo desta função é eliminar as funções time.sleep() e melhorar a confiabilidade da execução.

    Parameters:
        - image_path (str): O caminho para a imagem a ser verificada.
        - confidence_level (float): Nível de confiança para a correspondência da imagem.

    Autor: Almir
    """
    time.sleep(.5) # pausa para o ícone começar a carregar
    while True:
        try:
            pos_screen = pa.locateOnScreen(image_path, confidence= confidence_level)
            if pos_screen is not None:
                time.sleep(.5) # tempo adicional para a página terminar de carregar
                break
        except pa.ImageNotFoundException:
            pass

def send_file(file_path):
    """
    Envia um arquivo no menu de seleção
    Também verifica o status de envio do arquivo de backup para restauração.

    Parameters:
        - file_path (str): O caminho para o arquivo a ser enviado.

    - Autor: Almir
    """

    # parte web
    pos_choose_file = pa.locateCenterOnScreen('media/escolha-um-arquivo.png', confidence=0.9)
    pa.click(pos_choose_file, duration=.2)
    time.sleep(.5)  # pausa para abrir o menu de seleção de arquivo

    # diferenciação entre as imagens. Se essa opção de envio tiver sido selecionada na exec. anterior, ela vai ficar azul (como na img2 "hover")
    pos_arq = mult_img_selection(['media/enviar-um-arquivo-1.png', 'media/enviar-um-arquivo-2.png'], 0.9)
    if pos_arq is not None:
        pa.click(pos_arq, duration=.2)
        time.sleep(.2)  # pausa para carregar o menu
        pos_escolher_arq = pa.locateCenterOnScreen(imagem_escolher_arquivo, confidence=0.7)
        pa.click(pos_escolher_arq, duration=.2)
    else:
        time.sleep(.2)  # pausa para carregar o menu
        pos_escolher_arq = pa.locateCenterOnScreen(imagem_escolher_arquivo, confidence=0.7)
        pa.click(pos_escolher_arq, duration=.2)

    # parte explorador de arquivos
    time.sleep(1)  # pausa para carregar o explorador de arquivos
    user_path = os.path.expanduser("~")
    search = os.path.join(user_path, file_path)
    write_special_char(search)
    pa.press('enter')

    # verificação de erro ao enviar o arquivo
    time.sleep(.5) #pausa para carregar possível alerta de erro
    try:
        pos_screen = pa.locateOnScreen('media/restauracao-erro-arquivo-nao-encontrado.png', confidence=0.8)
        if pos_screen is not None:
            status_restauracao = 'Sem categoria (manuais) / Restauração'
            pa.press('escape')
            pa.press('escape')
            pa.press('escape')
    except pa.ImageNotFoundException:
        status_restauracao = 'success'
        # finalização do menu de seleção do arquivo
        pos_botao_enviar_arq = pa.locateCenterOnScreen('media/pos-enviar-arq.png', confidence=0.9)
        pa.click(pos_botao_enviar_arq, duration=.5)

    return status_restauracao

def process_turma(turma):
    """
    Processamento do padrão de turma, referente às funções de identificar informações para categorização.

    Parameters:
        - turma (re.Match): O padrão correspondente à turma.

    - Autor: Felipe
    """
    if turma:
        return turma.group()
    else:
        return 'EM'

def process_etapa(etapa):
    """
    Processamento do padrão de etapa, referente às funções de identificar informações para categorização.

    Parameters:
        - etapa (re.Match): O padrão correspondente à etapa.

    - Autor: Felipe
    """
    if etapa:
        etapa_grupo = etapa.group(1)
    else:
        etapa_grupo = 'Erro: etapa não encontrada'

    return etapa_grupo

def process_tipo_curso(tipo_curso):
    """
    Processamento do padrão de tipo de curso, referente às funções de identificar informações para categorização.

    Parameters:
        - tipo_curso (re.Match): O padrão correspondente ao tipo de curso.

    - Autor: Felipe
    """
    if tipo_curso:
        tipo_curso = tipo_curso.group(1)
    else:
        tipo_curso = 'Erro: Tipo de curso (ex.: PIPAT) não encontrado'

    return tipo_curso

def process_ano_semestre(ano_semestre):
    """
    Processamento do padrão de período de data, referente às funções de identificar informações para categorização.

    Parameters:
        - ano_semestre (re.Match): O padrão correspondente ao ano e semestre.

    - Autor: Felipe
    """
    if ano_semestre:
        ano_semestre_grupo1 = ano_semestre.group(1) if ano_semestre.group(1) else ano_semestre.group(3)
        ano_semestre_grupo2 = ano_semestre.group(2) if ano_semestre.group(2) else None
    else:
        ano_semestre_grupo1 = 'Erro: Ano não encontrado'
        ano_semestre_grupo2 = 'Erro: Semestre não encontrado'

    return ano_semestre_grupo1, ano_semestre_grupo2

def process_codigo_inep(codigo_inep):
    """
    Processamento do padrão do código inep, referente às funções de identificar informações para categorização.

    Parameters:
        - codigo_inep (re.Match): O padrão correspondente ao código INEP.

    - Autor: Almir
    """
    if codigo_inep:
        codigo_inep = codigo_inep.group(1)
    else:
        codigo_inep = 'Erro: Código INEP não encontrado'

    return codigo_inep

def restauracao_file_info(text):
    """
    Extração das informações para montagem do nome do arquivo de backup (restauração), a partir do nome breve, pelo uso de expressões regulares

    Parameters:
        - text (str): O texto a ser analisado para extrair informações.

    - Autor: Felipe e Almir
    """
    padroes= {
        'turma': (re.compile(r'(EJA-\D{2})'), 'EM'),
        'etapaEJA': (re.compile(r'_(([\d]+ª\sETAPA))'), None),
        'etapaEM': (re.compile(r'_(([\d]+ªN))'), None),
        'tipo_curso': (re.compile(r'_([^_]+)_INEP'), None),
        'ano_semestre': (re.compile(r'(?:_(\d+)/(\d+)|(\d{4})$)'), None)
    }

    # Procura por correspondências nos padrões no texto fornecido
    turma = re.search(padroes['turma'][0], text)
    tipo_curso = re.search(padroes['tipo_curso'][0], text)
    ano_semestre = re.search(padroes['ano_semestre'][0], text)

    # Verifica se as correspondências foram encontradas antes de acessar os grupos
    turma_grupo = process_turma(turma)
    if 'EJA' in text:
        etapaEJA = re.search(padroes['etapaEJA'][0], text)
        etapa_grupo = process_etapa(etapaEJA)
    else:
        etapaEM = re.search(padroes['etapaEM'][0], text)
        etapa_grupo = process_etapa(etapaEM)

    tipo_curso_grupo = process_tipo_curso(tipo_curso)
    ano_semestre_grupo1, ano_semestre_grupo2 = process_ano_semestre(ano_semestre)

    # Cria o nome do arquivo de backup conforme as informações extraídas
    if 'EJA' in text:
        nome_arquivo = f'{turma_grupo}_{ano_semestre_grupo1}_{ano_semestre_grupo2}_{etapa_grupo}_{tipo_curso_grupo}'
    else:
        # Definir como você quer que o nome do arquivo seja para strings que não contêm 'EJA'
        nome_arquivo = f'{turma_grupo}_{ano_semestre_grupo1}_{etapa_grupo}_{tipo_curso_grupo}'

    chars = str.maketrans({' ': '', 'ª': 'a', '-': '_'})
    return nome_arquivo.translate(chars)

def categorizacao_file_info(text):
    """
    Extração das informações para montagem do nome da categoria (categorização), a partir do nome breve, pelo uso de expressões regulares

    Parameters:
        - text (str): O texto a ser analisado para extrair informações.

    - Autor: Almir
    """

    padroes = {
        'turma': (re.compile(r'(EJA-\D{2})'), 'EM'),
        'ano_semestre': (re.compile(r'(?:_(\d+)/(\d+)|(\d{4})$)'), None),
        'codigo_inep': (re.compile(r'INEP(\d+)_'), None)
    }

    # Procura por correspondências nos padrões no texto fornecido
    turma = re.search(padroes['turma'][0], text)
    codigo_inep = re.search(padroes['codigo_inep'][0], text)
    turma_grupo = process_turma(turma)
    ano_semestre = re.search(padroes['ano_semestre'][0], text)
    ano_semestre_grupo1, ano_semestre_grupo2 = process_ano_semestre(ano_semestre)
    codigo_inep = process_codigo_inep(codigo_inep)

    # Realiza a substituição da turma conforme necessário
    if turma_grupo == 'EJA-EF':
        turma_grupo = 'EJA - ENSINO FUNDAMENTAL'
    elif turma_grupo == 'EJA-EM':
        turma_grupo = 'EJA - ENSINO MÉDIO'
    elif turma_grupo == 'EJA-EP':
        turma_grupo = 'EJA - PROFISSIONALIZANTE EM'

    # Cria o nome do arquivo de backup conforme as informações extraídas
    if 'EJA' in text:
        nome_arquivo = rf'INEP{codigo_inep} / {turma_grupo} / {ano_semestre_grupo1}/{ano_semestre_grupo2}'
    else:
        nome_arquivo = rf'INEP{codigo_inep} / ENSINO MÉDIO / {ano_semestre_grupo1}'

    return nome_arquivo



#inserção das imagens do curso



def imagem_file_info(text):
    """
    Extração das informações para escolha do arquivo de imagem, a partir do nome breve, pelo uso de expressões regulares

    Parameters:
        - text (str): O texto a ser analisado para extrair informações.

    - Autor: Almir
    """

    padroes = {
        'turma': (re.compile(r'(EJA-\D{2})'), 'EM'),
        'tipo_curso': (re.compile(r'_([^_]+)_INEP'), None),
    }

    # Procura por correspondências nos padrões no texto fornecido
    turma = re.search(padroes['turma'][0], text)
    tipo_curso = re.search(padroes['tipo_curso'][0], text)

    # Verifica se as correspondências foram encontradas antes de acessar os grupos
    turma_grupo = process_turma(turma)
    tipo_curso_grupo = process_tipo_curso(tipo_curso)

    # Cria o nome do arquivo de backup conforme as informações extraídas
    nome_arquivo = f'{turma_grupo}_{tipo_curso_grupo}'

    chars = str.maketrans({' ': '', '-': '_'})
    return nome_arquivo.translate(chars)

def send_course_image(nome_imagem, dflogexecutados):
    """
    É uma adaptação da função anterior send_file(), para enviar as imagens do curso

    Parameters:
        - nome_imagem (str): O nome do arquivo de imagem a ser enviado.
        - dflogexecutados (DataFrame): O DataFrame de log de execuções.

    - Autor: Almir
    """
    time.sleep(.5) # pausa pro scroll
    pa.scroll(240)
    time.sleep(.5) # pausa pro scroll
    pos_screen = pa.locateCenterOnScreen('media/upimagem-botao-enviar.png', confidence = .9)
    pa.click(pos_screen, duration = .5)
    # diferenciação entre as imagens. Se essa opção de envio tiver sido selecionada na exec. anterior, ela vai ficar azul (como na img2 "hover")
    pos_arq = mult_img_selection(['media/enviar-um-arquivo-1.png', 'media/enviar-um-arquivo-2.png'], 0.7)
    if pos_arq is not None:
        pa.click(pos_arq, duration=.2)
        time.sleep(.2)  # pausa para carregar o menu
        pos_escolher_arq = pa.locateCenterOnScreen(imagem_escolher_arquivo, confidence=0.7)
        pa.click(pos_escolher_arq, duration=.2)
    else:
        time.sleep(.2)  # pausa para carregar o menu
        pos_escolher_arq = pa.locateCenterOnScreen(imagem_escolher_arquivo, confidence=0.7)
        pa.click(pos_escolher_arq, duration=.2)

    # parte explorador de arquivos
    time.sleep(1)  # pausa para carregar o explorador de arquivos
    user_path = os.path.expanduser("~")
    image_path = rf'Secretaria de Estado da Educação\Cefope - Equipe Tecnologia\Projetos Python\Automação Plataforma Estudante\backups\geral\{nome_imagem}.png'
    search = os.path.join(user_path, image_path)
    write_special_char(search)
    pa.press('enter')

    # verificação de erro ao enviar o arquivo
    time.sleep(1.5) #pausa para carregar possível alerta de erro
    try:
        pos_screen = pa.locateOnScreen('media/restauracao-erro-arquivo-nao-encontrado.png', confidence=0.8)
        if pos_screen is not None:
            pa.press('escape')
            pa.press('escape')
            pa.press('escape')
            dflogexecutados.at[dflogexecutados.index[-1], 'nome_imagem'] = 'erro_imagem'
    except pa.ImageNotFoundException:
        # finalização do menu de seleção do arquivo
        dflogexecutados.at[dflogexecutados.index[-1], 'nome_imagem'] = nome_imagem
        pos_botao_enviar_arq = pa.locateCenterOnScreen('media/pos-enviar-arq.png', confidence=0.9)
        pa.click(pos_botao_enviar_arq, duration=.5)


#GUI
