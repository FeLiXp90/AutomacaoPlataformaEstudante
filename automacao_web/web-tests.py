import psutil, pyautogui as pa, pymsgbox as pmsg, time, os, sys

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
    pmsg.alert(text='ALERTA. Nenhuma imagem foi encontrada: \nReveja o código.', title='PTFE v1.0')
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
            pmsg.alert(text='Erro de digitação, tente novamente.', title='Login Proxy - [PTFE v1.0]', button = 'Tentar novamente')
            return True
    except pa.ImageNotFoundException:
        return False

def login_proxy():
    #requisita dados para logar no proxy localmente
    #também cria um loop para verificar se o login foi digitado corretamente
    time.sleep(1) #pausa para esperar o chrome carregar.
    while True:
        proxyUser = pmsg.prompt(text='Digite seu nome de usuário', title='Login Proxy - [PTFE v1.0]')
        proxyPassword = pmsg.password(text='Digite sua senha', title='Login Proxy - [PTFE v1.0]', mask='*')
        pa.write(proxyUser)
        pa.press('tab')
        pa.write(proxyPassword)
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
        if pos_verific_login is not None:
            pass
    except pa.ImageNotFoundException:
        pos_verific_login = None

    if pos_verific_login is None:
        while True:
            AVAUser = pmsg.prompt(text='Digite seu nome de usuário. Se já estiver salvo, apenas aperte "enter"', title='Login AVA - [PTFE v1.0]')
            AVAPassword = pmsg.password(text='Digite sua senha. Se já estiver salvo, apenas aperte "enter"', title='Login AVA - [PTFE v1.0]', mask='*')
            pos_login = pa.locateCenterOnScreen('media/pos-login.png', confidence=0.7)
            pos_login = (pos_login[0], pos_login[1] - 120)
            pa.click(pos_login, duration=1)
            pa.write(AVAUser)
            pa.press('tab')
            pa.write(AVAPassword)
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
formato = pmsg.confirm(text=f'Selecione o formato do curso: Blocos ou Tópicos.', title='Formato de curso - [PTFE v1.0]', buttons=['Blocos', 'Tópicos', 'Encerrar programa'])

if formato == 'Blocos':
    select_drop_down_list_desloc('media/previsualizar_formato-curso.png',['media/previsualizar_formato-curso-blocos-op-1.png', 'media/previsualizar_formato-curso-blocos-op-2.png'], 0.7, 0.9 )
elif formato == 'Encerrar programa':
    pmsg.alert(text='Obrigado. Programa encerrado',
               title='PTFE v1.0')
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
    if pos_screen is not None:
        pass
except pa.ImageNotFoundException:
    pmsg.alert(text='ALERTA. Os cursos não seriam categorizados: \nReveja o código e soluções.',
               title='PTFE v1.0')
    sys.exit()

#carregar cursos
select_drop_down_list_exact('media/list-susp-carregamento.png', ['media/op-list-susp-carregamento-1.png',
                        'media/op-list-susp-carregamento-2.png'], 0.95, 0.9)

#confirmar
pa.hotkey('ctrl', 'end')
time.sleep(.2) #pausa para ir ao final da tela sem problemas

pos_screen = pa.locateCenterOnScreen('media/previsualizar_carregar-cursos.png', confidence = 0.9)
pa.moveTo(pos_screen, duration = 1)