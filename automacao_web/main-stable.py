#TODO Pensar em formas de documentar o processo, criação de registros para mapeamentos e correções de erros.
#TODO replicar todas as alterações para o de produção.
#TODO reduzir o número de variáveis de posições para click
#TODO reduzir ao máximo o tempo dos time.sleep
#TODO testar o salvamento dos templates
#TODO renomear as imagens de forma que fiquem mais organizadas.: ex, 'partedoproceso_descricao-da-imagem'
#TODO na pre visualização dos dados, a opção de 100 linhas não é das melhores. Muda a posição da tela e estraga o resto do código, setar pra dez, dar um ctrl end e começar de baixo pra cima até encontrar a última configuração, daí ctrl end de novo e parte pra próxima etapa
#TODO documentar melhor as funções e seus argumentos
#TODO corrigir linhas com enter para quebra auto (warp)

import tkinter as tk, pandas as pd, pymsgbox as pmsg, sys, os
from tkinter import filedialog
from datetime import datetime

OBRIGADOPROGRAMAENCERRADO = 'Obrigado. Programa encerrado'
ENCERRARPROGRAMA = 'Encerrar programa'
LOGINPROXY = 'Login Proxy - [PTFE v1.0]'
PTFEE = 'PTFE v1.0'

def abrir_seletor_arquivos():
    while True:
        # Cria uma instância do tkinter.Tk() para ser a janela principal
        root = tk.Tk()
        root.withdraw()  # Oculta a janela principal

        # Abre o seletor de arquivos e retorna o caminho do arquivo selecionado
        caminho_arquivo = filedialog.askopenfilename()

        # Verifica se o usuário selecionou um arquivo ou cancelou a operação
        if caminho_arquivo:
            confirmacao = pmsg.confirm(text=f'Você selecionou o arquivo:\n{caminho_arquivo}\nDeseja continuar?',
                                       title='Confirmação de arquivo - [PTFE v1.0]',
                                       buttons=['Confirmo', 'Escolher novamente', ENCERRARPROGRAMA])
            if confirmacao == 'Confirmo':
                return caminho_arquivo
            elif confirmacao == ENCERRARPROGRAMA:
                pmsg.alert(text=OBRIGADOPROGRAMAENCERRADO,
                           title=PTFEE)
                sys.exit()
        else:
            escolha = pmsg.confirm(text='Oops, nenhum arquivo selecionado.',
                                   title='Erro - [PTFE v1.0]',
                                   buttons=['Escolher novamente', ENCERRARPROGRAMA])
            if escolha == ENCERRARPROGRAMA:
                pmsg.alert(text=OBRIGADOPROGRAMAENCERRADO,
                           title=PTFEE)
                sys.exit()

def nome_curso(shortname):
    if "_PIP_" in shortname:
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

# Chama a função para abrir o seletor de arquivos dentro de um loop
while True:
    df = pd.read_excel(abrir_seletor_arquivos(), sheet_name='Matriculas_Novas')

    if not df.empty:
        # Faça o que quiser com "arquivo_selecionado"
        break  # Sai do loop se o arquivo for selecionado

# Filtre as linhas onde 'Curso Existente?' é 'NÃO'
df_filtrado = df[df['Curso Existente?'] == 'NÃO']

# Crie uma cópia do DataFrame filtrado
df_filtrado_copia = df_filtrado.copy()

# Aplicar a função à cópia
df_filtrado_copia['fullname'] = df_filtrado_copia['IdentificacaoCurso'].apply(nome_curso)

# Renomeie as colunas
df_filtrado_copia = df_filtrado_copia.rename(columns={'IdentificacaoCurso': 'shortname'})

# Adicione a coluna 'category' com o valor 8827 para todas as linhas
df_filtrado_copia['category'] = 8827

#Verifica se já existe um templatefinal anterior e faz um backup para não sobrescrever
caminho_arquivo = 'uploadCSV/templatefinal.csv'

    # Verifica se o arquivo já existe
if os.path.exists(caminho_arquivo):
    # Informações de metadados do arquivo existente
    info_arquivo = os.stat(caminho_arquivo)

    # Extrai a data de criação e hora de criação do arquivo existente
    data_criacao = datetime.fromtimestamp(info_arquivo.st_ctime).strftime('%Y%m%d')
    hora_criacao = datetime.fromtimestamp(info_arquivo.st_ctime).strftime('%H%M%S')

    # Constrói o novo nome do arquivo com base nos metadados
    novo_nome_arquivo = f'templatefinal_{data_criacao}_{hora_criacao}.csv'

    # Renomeia o arquivo existente
    os.rename(caminho_arquivo, 'uploadCSV/' + novo_nome_arquivo)

# Salvar o DataFrame em um arquivo CSV
df_filtrado_copia[['shortname', 'fullname', 'category']].to_csv(caminho_arquivo, sep=';', index=False)













import psutil, pyautogui as pa, time

#abrindo chrome
def is_chrome_running():
    #verifica se o chrome está aberto com a utilização da lib psutil
    for process in psutil.process_iter(['pid', 'name']): #pid solicita o processID
        if 'chrome.exe' in process.info['name'].lower():
            return True
    return False

def write_special_char(text):
    #essa função resolve a codificação do pa.write para escrever caracteres especiais, como ~ e ç
    import pyperclip as pclip
    pclip.copy(text)
    pa.hotkey('ctrl', 'v')

def mult_img_selection(img_list, confidence_level):
    # Seleção entre diferentes imagens possíveis presentes na tela
    for i in img_list:
        try:
            pos_arq = pa.locateCenterOnScreen(i, confidence=confidence_level)
            return pos_arq  # Retorna a posição se a imagem for encontrada
        except pa.ImageNotFoundException:
            continue
    pmsg.alert(text='ALERTA. Nenhuma imagem foi encontrada: \nReveja o código.', title=PTFEE)
    sys.exit()

def select_drop_down_list_exact(list_file_name, img_list, confidence_level_dropdown, confidence_level_options):
    #clica e seleciona a opção correta das listas suspensas  se não houverem caixas de seleção de mesmo tamanho na tela
    pos_list = pa.locateCenterOnScreen(list_file_name, confidence = confidence_level_dropdown)
    pa.click(pos_list[0], pos_list[1] - 10, duration = 0.5)
    pos_op = mult_img_selection(img_list, confidence_level_options)
    pa.click(pos_op, duration=.5)
    time.sleep(.1) #para garantir que a lista recolherá a tempo com o click e não ocultará a próxima exec.

def select_drop_down_list_desloc(list_file_name, img_list, confidence_level_dropdown, confidence_level_options):
    #clica e seleciona a opção correta das listas suspensas com base na posição do título à esquerda apenas
    #essa parte do código é sensível ao tamanho da tela, pois localiza apenas o texto
    #o print deve ser tirado da extremidade esquerda do título até imediatamente antes da caixa de seleção à direita
    pos_list = pa.locateOnScreen(list_file_name, confidence = confidence_level_dropdown)
    pa.click(pos_list[0] + pos_list.width + 50, pos_list[1] + 20, duration = .5)
    pos_op = mult_img_selection(img_list, confidence_level_options)
    pa.click(pos_op, duration=.5)
    time.sleep(.1) #para garantir que a lista recolherá a tempo com o click e não ocultará a próxima exec.

def select_drop_down_list_dual(list_file_name, confidence_level_dropdown, direction):
    """
    Seleciona um item de uma lista suspensa que possui apenas duas opções.

    Parameters:
        - list_file_name (Any): Nome do arquivo da lista suspensa.
        - confidence_level_dropdow (Any): Nível de confiança para identificar a lista suspensa.
        - direction (str): Direção para clicar na lista suspensa ('up' ou 'down').
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
    #verifica se as credenciais informadas em qualquer seção de login foram as corretas
    try:
        time.sleep(.5)  # pausa para verificar se o login foi correto.
        pos_login_proxy = pa.locateOnScreen(nome_arquivo, confidence=0.7)
        if pos_login_proxy is not None:
            pmsg.alert(text='Erro de digitação, tente novamente.', title=LOGINPROXY, button = 'Tentar novamente')
            return True
    except pa.ImageNotFoundException:
        return False

def login_proxy():
    #requisita dados para logar no proxy localmente
    #também cria um loop para verificar se o login foi digitado corretamente
    time.sleep(1) #pausa para esperar o chrome carregar.
    while True:
        proxy_user = pmsg.prompt(text='Digite seu nome de usuário', title=LOGINPROXY)
        proxy_password = pmsg.password(text='Digite sua senha', title=LOGINPROXY, mask='*')
        pa.write(proxy_user)
        pa.press('tab')
        pa.write(proxy_password)
        pa.press('enter')
        if not login_verification('media/login-proxy.png'):
            break

def login_plataforma_estudante():
    # loga na plataforma do estudante

    #carrega a página da plataforma do estudante.
    pa.hotkey('ctrl', 't')
    pa.write('http://alunodchm.sedu.es.gov.br/ava/admin/tool/uploadcourse/index.php') #TODO trocar pelo ambiente de produção
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
            pa.click(pos_login, duration=1)
            pa.write(ava_user)
            pa.press('tab')
            pa.write(ava_password)
            pa.press('enter')
            if login_verification('media/pos-login.png') is False:
                break

if not is_chrome_running():
    pa.press('win')
    pa.write('Chrome')
    pa.press('enter')
    login_proxy()
    login_plataforma_estudante()
else:
    pos_icon_chrome = pa.locateCenterOnScreen('media/iconChrome.png', confidence = 0.9)
    pa.click(pos_icon_chrome, duration = 1)
    login_plataforma_estudante()


#carregamento do arquivo CSV
    #maximização da janela do chrome
pa.hotkey('alt', 'space')
pa.press('x')
    #parte web
time.sleep(1) #pausa para a plataforma do estudante carregar
pos_choose_file = pa.locateCenterOnScreen('media/escolha-um-arquivo.png', confidence = 0.9)
pa.click(pos_choose_file, duration = .2)
time.sleep(.5) #pausa para abrir o menu de seleção de arquivo

    #diferenciação entre as imagens. Se essa opção de envio tiver sido selecionada na exec. anterior, ela vai ficar azul (como na img2 "hover")
pos_arq = mult_img_selection(['media/enviar-um-arquivo-1.png','media/enviar-um-arquivo-2.png'], 0.9)
if pos_arq is not None:
    pa.click(pos_arq, duration = .2)
    time.sleep(.2) #pausa para carregar o menu
    pos_escolher_arq = pa.locateCenterOnScreen('media/escolher-arquivo.png', confidence=0.7)
    pa.click(pos_escolher_arq, duration=.2)
else:
    time.sleep(.2) #pausa para carregar o menu
    pos_escolher_arq = pa.locateCenterOnScreen('media/escolher-arquivo.png', confidence=0.7)
    pa.click(pos_escolher_arq, duration=.2)

    #parte explorador de arquivos
time.sleep(1) #pausa para carregar o explorador de arquivos
user_path = os.path.expanduser("~")
file_path = os.path.join(user_path, 'Secretaria de Estado da Educação', 'Cefope - Equipe Tecnologia', 'Projetos Python',
                                'Automação Plataforma Estudante', 'automacao_web', 'uploadCSV', 'templatefinal.CSV')
write_special_char(file_path)
pa.press('enter')

    #finalização do menu de seleção do arquivo
pos_botao_enviar_arq = pa.locateCenterOnScreen('media/pos-enviar-arq.png', confidence = 0.9)
pa.click(pos_botao_enviar_arq, duration = .5)

    #configuração das opções de envio das listas suspensas)
select_drop_down_list_exact('media/list-susp-delimitador.png', ['media/op-list-susp-delimitador-1.png',
                        'media/op-list-susp-delimitador-2.png'], 0.7, 0.9)

select_drop_down_list_exact('media/list-susp-codificacao.png', ['media/op-list-susp-codificacao-1.png',
                        'media/op-list-susp-codificacao-2.png'], 0.7, 0.9)

select_drop_down_list_exact('media/list-susp-linhas.png', ['media/op-list-susp-linhas-1.png',
                        'media/op-list-susp-linhas-2.png'], 0.7, 0.9)

pa.hotkey('ctrl', 'end')
time.sleep(.2)

select_drop_down_list_exact('media/list-susp-carregamento.png', ['media/op-list-susp-carregamento-1.png',
                        'media/op-list-susp-carregamento-2.png'], 0.95, 0.9)

pos_pre_visual = pa.locateCenterOnScreen('media/pre-visualizar.png', confidence = 0.8)
pa.click(pos_pre_visual)

#Pré-visualização de cursos carregados
time.sleep(1.5) #pausa para carregar a próxima guia com segurança
pos_screen = pa.locateCenterOnScreen('media/previsualizar_titulo-click-ctrl-end.png', confidence = 0.7)
pa.click(pos_screen)
pa.hotkey('ctrl','end')
time.sleep(.2) #pausa para ir ao final da tela sem complicações

select_drop_down_list_dual('media/previsualizar_forcar-modalidade-grupo.png', 0.8, 'down')

#Livro de notas = não
select_drop_down_list_dual('media/previsualizar_mostrar-livro-notas.png', 0.8, 'up')

#Formato de blocos
formato = pmsg.confirm(text='Selecione o formato do curso: Blocos ou Tópicos.', title='Formato de curso - [PTFE v1.0]', buttons=['Blocos', 'Tópicos', ENCERRARPROGRAMA])

if formato == 'Blocos':
    select_drop_down_list_desloc('media/previsualizar_formato-curso.png',['media/previsualizar_formato-curso-blocos-op-1.png', 'media/previsualizar_formato-curso-blocos-op-2.png'], 0.7, 0.9 )
elif formato == ENCERRARPROGRAMA:
    pmsg.alert(text=OBRIGADOPROGRAMAENCERRADO,
               title=PTFEE)
    sys.exit()
else:
    select_drop_down_list_desloc('media/previsualizar_formato-curso.png',['media/previsualizar_formato-curso-topicos-op-1.png', 'media/previsualizar_formato-curso-topicos-op-2.png'], 0.7, 0.9 )

#habilitar data
try:
    pos_screen = pa.locateCenterOnScreen('media/habilitar-data-termino-curso.png', confidence=0.9)
    pa.click(pos_screen, duration = .5)
except pa.ImageNotFoundException:
    pass

#Sem categoria
pa.scroll(1000)
time.sleep(.2) #para dar tempo do scroll exectuar tranquilamente

    #confirmação de que os cursos serão alocados no 'Sem Categoria', apenas para dar mais segurança
try:
    pos_screen = pa.locateCenterOnScreen('media/sem-categoria.png', confidence = 0.7)
except pa.ImageNotFoundException:
    pmsg.alert(text='ALERTA. Os cursos não seriam categorizados: \nReveja o código e soluções.',
               title=PTFEE)
    sys.exit()

#carregar cursos
select_drop_down_list_exact('media/list-susp-carregamento.png', ['media/op-list-susp-carregamento-1.png',
                        'media/op-list-susp-carregamento-2.png'], 0.95, 0.9)

#confirmar
pa.hotkey('ctrl', 'end')
time.sleep(.2) #pausa para ir ao final da tela sem problemas

pos_screen = pa.locateCenterOnScreen('media/previsualizar_carregar-cursos.png', confidence = 0.9)
pa.moveTo(pos_screen, duration = 1)