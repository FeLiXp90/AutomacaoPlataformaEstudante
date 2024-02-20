import lib_functions_automacao_PTFE as lib_functions
import pandas as pd, pymsgbox as pmsg, sys, os  # tratamento de dados
from datetime import datetime  # upload de arquivo .csv
import pyautogui as pa, time, pyperclip as pclip  # upload de arquivo .csv

programa_encerrado = 'Obrigado. Programa encerrado'
encerrar_programa = 'Encerrar programa'
login_proxy = 'Login Proxy - [PTFE v1.0]'
PTFE = 'PTFE v1.0'
local_templatefinal = 'uploadCSV/templatefinal.csv'
local_imagem_botao_config = 'media/restauracao-botao-configuracoes.png'
local_carregamento_PTFE = 'media/verifica_carregamento_PTFE.png'
local_imagem_confimacao = 'media/restauracao-opcao-configuracoes.png'
local_imagem_lista_categoria = 'media/restauracao-lista-categoria-curso.png'
local_imagem_botao_salvar = 'media/restauracao-botao-salvaremostrar.png'


def tratamento_dados():
    '''
    Executa a parte de tratamento de dados das planilhas e faz o upload dos cursos do arquivo csv na plataforma do estudante
    Autor: Almir
    '''
    # Chama a função para abrir o seletor de arquivos dentro de um loop e para escolher o tipo de combinação de planilhas
    while True:
        tipo_planilha_escolhido = lib_functions.escolher_tipo_planilha()

        if tipo_planilha_escolhido == 'usuarios':
            pmsg.alert(text='Escolha a planilha de USUÁRIOS', title=PTFE)
            df_usuarios = pd.read_excel(lib_functions.abrir_seletor_arquivos(), sheet_name='New_Enrolment')
            if not df_usuarios.empty:
                break  # Sai do loop se o arquivo for selecionado

        elif tipo_planilha_escolhido == 'professores':
            pmsg.alert(text='Escolha a planilha de PROFESSORES', title=PTFE)
            df_professores = pd.read_excel(lib_functions.abrir_seletor_arquivos(), sheet_name='New_Enrolment')
            if not df_professores.empty:
                break  # Sai do loop se o arquivo for selecionado

        else:
            pmsg.alert(text='Escolha a planilha de USUÁRIOS', title=PTFE)
            df_usuarios = pd.read_excel(lib_functions.abrir_seletor_arquivos(), sheet_name='New_Enrolment')
            pmsg.alert(text='Escolha a planilha de PROFESSORES', title=PTFE)
            df_professores = pd.read_excel(lib_functions.abrir_seletor_arquivos(), sheet_name='New_Enrolment')

            if not df_usuarios.empty:
                break  # Sai do loop se o arquivo for selecionado
            if not df_professores.empty:
                break  # Sai do loop se o arquivo for selecionado

    if 'df_usuarios' in locals():
        # Filtre as linhas onde 'Curso Existente?' é 'NÃO'
        df_filtrado_usuarios = df_usuarios[df_usuarios['Curso Existente?'] == 'NÃO']
        # Crie uma cópia do DataFrame filtrado
        df_filtrado_copia_usuarios = df_filtrado_usuarios.copy()
        # Aplicar a função à cópia
        df_filtrado_copia_usuarios['fullname'] = df_filtrado_copia_usuarios['IdentificacaoCurso'].apply(lib_functions.nome_curso)
        # Renomeie as colunas
        df_filtrado_copia_usuarios = df_filtrado_copia_usuarios.rename(columns={'IdentificacaoCurso': 'shortname'})
        # Adicione a coluna 'category' com o valor 6622 para todas as linhas
        df_filtrado_copia_usuarios['category'] = 6622
    else:
        df_filtrado_copia_usuarios = pd.DataFrame()  # Criar um DataFrame vazio se df_usuarios não existir

    if 'df_professores' in locals():
        df_filtrado_professores = df_professores[df_professores['OBS:IdentificacaoCurso'] == 'NOVO']
        df_filtrado_copia_professores = df_filtrado_professores.copy()
        df_filtrado_copia_professores['fullname'] = df_filtrado_copia_professores['IdentificacaoCurso'].apply(lib_functions.nome_curso)
        df_filtrado_copia_professores = df_filtrado_copia_professores.rename(columns={'IdentificacaoCurso': 'shortname'})
        df_filtrado_copia_professores['category'] = 6622
    else:
        df_filtrado_copia_professores = pd.DataFrame()  # Criar um DataFrame vazio se df_professores não existir

    # Concatena os DataFrames
    df_final = pd.concat([df_filtrado_copia_usuarios, df_filtrado_copia_professores], ignore_index=True)

    # Verifica se já existe um templatefinal anterior e faz um backup para não sobrescrever
    caminho_arquivo = local_templatefinal

    # Adiciona a última linha ao DataFrame, para controle da restauração e categorização
    nova_linha = {'shortname': 'automatizacaoPTFE', 'fullname': 'AUTOMATIZACAO', 'category': 6622}
    df_final = pd.concat([pd.DataFrame([nova_linha]), df_final], ignore_index=True)

    # Salve o DataFrame em um arquivo CSV
    df_final[['shortname', 'fullname', 'category']].to_csv(local_templatefinal, sep=';', index=False)

    # Criação de log inicial, inclusive com cursos já existentes na base anterior (prevenção de erros durante a execução do código)
    data_hora_atual = datetime.now()
    data_log = data_hora_atual.strftime("%Y%m%d")  # Formato: AAAAMMDD
    hora_log = data_hora_atual.strftime("%H%M%S")  # Formato: HHMMSS

    df = pd.read_csv(local_templatefinal, encoding='utf-8', sep=';')
    df = df.assign(nome_backup=df['shortname'].apply(lib_functions.restauracao_file_info),nome_categoria=df['shortname'].apply(lib_functions.categorizacao_file_info),imagem_escolhida=df['shortname'].apply(lib_functions.imagem_file_info))
    nome_arquivo_log = f'log_geral_{data_log}_{hora_log}.csv'
    df.to_csv(f'logs/{nome_arquivo_log}', sep=';', index=False)

    # abrindo chrome para envio do arquivo csv
    chrome = lib_functions.is_chrome_running()
    if not chrome:
        lib_functions.login_proxy()
    else:
        pa.press('win')
        pa.write('Chrome')
        pa.press('enter')
        time.sleep(2)  # pausa para o chrome abrir

    lib_functions.login_plataforma_estudante()

    # carregamento do arquivo CSV
    # maximização da janela do chrome
    pa.hotkey('alt', 'space')
    pa.press('x')
    time.sleep(1)  # pausa para a plataforma do estudante carregar
    lib_functions.send_file(os.path.join('Secretaria de Estado da Educação', 'Cefope - Equipe Tecnologia', 'Projetos Python', 'Automação Plataforma Estudante', 'automacao_web', 'uploadCSV', 'templatefinal.CSV'))

    # configuração das opções de envio das listas suspensas)
    lib_functions.select_drop_down_list_exact('media/list-susp-delimitador.png',['media/op-list-susp-delimitador-1.png','media/op-list-susp-delimitador-2.png'], 0.7, 0.9)

    lib_functions.select_drop_down_list_exact('media/list-susp-codificacao.png',['media/op-list-susp-codificacao-1.png','media/op-list-susp-codificacao-2.png'], 0.7, 0.9)

    lib_functions.select_drop_down_list_exact('media/list-susp-linhas.png',['media/op-list-susp-linhas-1.png', 'media/op-list-susp-linhas-2.png'],0.7, 0.9)

    pa.hotkey('ctrl', 'end')
    time.sleep(.5)

    lib_functions.select_drop_down_list_exact('media/list-susp-carregamento.png', ['media/op-list-susp-carregamento-1.png', 'media/op-list-susp-carregamento-2.png'], 0.95, 0.9)

    pos_pre_visual = pa.locateCenterOnScreen('media/pre-visualizar.png', confidence=0.8)
    pa.click(pos_pre_visual)

    # Pré-visualização de cursos carregados
    time.sleep(2)  # pausa para carregar a próxima guia com segurança
    pos_screen = pa.locateCenterOnScreen('media/previsualizar_titulo-click-ctrl-end.png', confidence=0.7)
    pa.click(pos_screen)
    pa.hotkey('ctrl', 'end')
    time.sleep(.5)  # pausa para ir ao final da tela sem complicações

    lib_functions.select_drop_down_list_dual('media/previsualizar_forcar-modalidade-grupo.png', 0.8, 'down')

    # Livro de notas = não
    lib_functions.select_drop_down_list_dual('media/previsualizar_mostrar-livro-notas.png', 0.8, 'up')

    # Formato de blocos
    formato = pmsg.confirm(text='Selecione o formato do curso: Blocos ou Tópicos.', title='Formato de curso - [PTFE v1.0]', buttons=['Tópicos', 'Blocos', encerrar_programa])
    time.sleep(1)  # pausa pro scroll
    pa.scroll(240)
    time.sleep(1)  # pausa pro scrol
    if formato == 'Blocos':
        lib_functions.select_drop_down_list_desloc('media/previsualizar_formato-curso.png',['media/previsualizar_formato-curso-blocos-op-1.png','media/previsualizar_formato-curso-blocos-op-2.png'], 0.7, 0.9)
    elif formato == 'Tópicos':
        lib_functions.select_drop_down_list_desloc('media/previsualizar_formato-curso.png',['media/previsualizar_formato-curso-topicos-op-1.png','media/previsualizar_formato-curso-topicos-op-2.png'], 0.7, 0.9)
    elif formato == encerrar_programa:
        pmsg.alert(text=f'{programa_encerrado}', title=PTFE)
        sys.exit()

    time.sleep(1)

    # habilitar data
    try:
        pos_screen = pa.locateCenterOnScreen('media/habilitar-data-termino-curso.png', confidence=0.8)
        pa.click(pos_screen, duration=.5)
    except pa.ImageNotFoundException:
        pass

    # Sem categoria
    pa.scroll(1000)
    time.sleep(.5)  # para dar tempo do scroll exectuar tranquilamente

    # confirmação de que os cursos serão alocados no 'Sem Categoria', apenas para dar mais segurança
    try:
        pos_screen = pa.locateCenterOnScreen('media/sem-categoria.png', confidence=0.7)
    except pa.ImageNotFoundException:
        pmsg.alert(text='ALERTA. Os cursos não seriam categorizados: \nReveja o código e soluções.', title=PTFE)
        sys.exit()

    # carregar cursos
    lib_functions.select_drop_down_list_exact('media/list-susp-carregamento.png',['media/op-list-susp-carregamento-1.png','media/op-list-susp-carregamento-2.png'], 0.95, 0.9)

    # confirmar
    pa.hotkey('ctrl', 'end')
    time.sleep(.5)  # pausa para ir ao final da tela sem problemas

    pos_screen = pa.locateCenterOnScreen('media/previsualizar_carregar-cursos.png', confidence=0.9)
    confirmacao = pmsg.confirm(text='Carregar cursos?', title='Confirmação de carregamento - [PTFE v1.0]',
                               buttons=['Sim', 'Não'])
    if confirmacao == 'Sim':
        pa.click(pos_screen, duration=1)
    else:
        pmsg.alert(text=f'{programa_encerrado}', title=PTFE)
        sys.exit()

    time.sleep(5)
    continua = pmsg.confirm(text='Os cursos foram carregados?', title='Carregamento: Status - [PTFE v1.0]',
                            buttons=['Sim', 'Não'])
    if continua == 'Não':
        pmsg.alert(text=f'{programa_encerrado}', title=PTFE)
        sys.exit()


def web():
    '''
    Executa a parte do código de restauração, categorização e upload da imagem do curso
    Autor: Almir
    '''
    try:
        # restauração e categorização dos cursos
        # criação do dataframe para salvar os arquivos de log
        dflogexclusivos = pd.DataFrame(columns=['shortname', 'nome_backup', 'nome_categoria', 'nome_imagem'])
        dflogexecutados = pd.DataFrame(columns=['shortname', 'nome_backup', 'nome_categoria', 'nome_imagem'])

        pa.press('win')
        pa.write('chrome')
        pa.press('enter')
        time.sleep(2)  # pausa para o chrome iniciar novamente
        pa.write('http://estudante.sedu.es.gov.br/ava/course/index.php')
        pa.press('enter')
        lib_functions.verifica_load(local_carregamento_PTFE, .9) # pausa para a página carregar

        while True:
            # abrindo sequência de cursos
            pa.press('f6')
            pa.write('http://estudante.sedu.es.gov.br/ava/course/index.php')
            pa.press('enter')
            lib_functions.verifica_load(local_carregamento_PTFE, .9) # pausa para a página carregar
            time.sleep(1) #pausa adicional para terminar de carregar
            pos_screen = pa.locateOnScreen('media/restauracao-escolha-sala-sem-categoria.png', confidence=0.7)
            pa.click(pos_screen.left, pos_screen.top + pos_screen.height - 10, duration=.5)
            lib_functions.verifica_load('media/erro-verifica-lista-sem-categoria.png', 0.9)
            pa.click(pos_screen.left + 80, pos_screen.top + pos_screen.height + 35, duration=.5)
            lib_functions.verifica_load(local_carregamento_PTFE, .9) # pausa para a página carregar

            # verificação se deve encerrar ou não ao acessar o curso "automatização"
            try:
                pos_verific = pa.locateOnScreen('media/restauracao-imagem-para-encerramento.png', confidence=0.9)
                if pos_verific is not None:
                    # salva o log
                    data_hora_atual = datetime.now()
                    data_log = data_hora_atual.strftime("%Y%m%d")  # Formato: AAAAMMDD
                    hora_log = data_hora_atual.strftime("%H%M%S")  # Formato: HHMMSS
                    nome_arquivo_log = f'log_exclusivos_{data_log}_{hora_log}.csv'
                    dflogexclusivos.to_csv(f'logs/{nome_arquivo_log}', sep=';', index=False)
                    nome_arquivo_log = f'log_executados_{data_log}_{hora_log}.csv'
                    dflogexecutados.to_csv(f'logs/{nome_arquivo_log}', sep=';', index=False)
                    # abertura gerenciador de cursos para exclusão do automático
                    pa.press('f6')
                    pa.write('http://estudante.sedu.es.gov.br/ava/course/management.php')
                    pa.press('enter')
                    lib_functions.verifica_load(local_carregamento_PTFE, .9) # pausa para a página carregar
                    time.sleep(2) #tempo extra para a página carregar
                    pos_screen = pa.locateOnScreen('media/finalizacao-deletar-sala-automatizacao.png', confidence=0.8)
                    pa.click(pos_screen.left + pos_screen.width, pos_screen.height + pos_screen.top - 10, duration=5)
                    time.sleep(2)  # pausa para carregar a janela de confirmação
                    pos_screen = pa.locateCenterOnScreen('media/finalizacao-botao-excluir.png', confidence=0.9)
                    pa.click(pos_screen, duration=.5)
                    pmsg.alert(text=f'Restaurações e categorizações concluídas. \n{programa_encerrado}.', title=PTFE)
                    sys.exit()
            except pa.ImageNotFoundException:
                pass

            # abrir configuracoes para coletar nome breve
            pos_screen = pa.locateCenterOnScreen(local_imagem_botao_config, confidence=0.9)
            pa.click(pos_screen, duration=.5)
            pos_screen = pa.locateCenterOnScreen(local_imagem_confimacao, confidence=0.9)
            pa.click(pos_screen, duration=.5)

            # copiar nome breve
            lib_functions.verifica_load(local_carregamento_PTFE, .9) # pausa para a página carregar
            pos_screen = pa.locateOnScreen('media/restauracao-nome-breve-do-curso.png', confidence=0.95)
            pa.click(pos_screen[0] + pos_screen.width + 50, pos_screen[1] + 20, duration=.5)
            pa.hotkey('ctrl', 'a', 'ctrl', 'c')
            nome_breve = pclip.paste()
            nome_backup = lib_functions.restauracao_file_info(nome_breve)
            nome_categoria = lib_functions.categorizacao_file_info(nome_breve)
            nome_imagem = lib_functions.imagem_file_info(nome_breve)
            new_row = pd.DataFrame({'shortname': [nome_breve], 'nome_backup': [nome_backup], 'nome_categoria': [nome_categoria], 'nome_imagem': [nome_imagem]})
            dflogexclusivos = pd.concat([dflogexclusivos, new_row], ignore_index=True)
            new_row = pd.DataFrame({'shortname': [''], 'nome_backup': [''], 'nome_categoria': [''], 'nome_imagem': ['']})
            dflogexecutados = pd.concat([dflogexecutados, new_row], ignore_index=True)
            dflogexecutados.at[dflogexecutados.index[-1], 'shortname'] = nome_breve

            # abrindo menu de restauração
            pos_screen = pa.locateCenterOnScreen(local_imagem_botao_config, confidence=0.9)
            pa.click(pos_screen, duration=0.5)
            pos_screen = pa.locateCenterOnScreen('media/restauracao-opcao-restaurar.png', confidence=0.9)
            pa.click(pos_screen, duration=0.5)
            lib_functions.verifica_load(local_carregamento_PTFE, .9) # pausa para a página carregar

            # envio do arquivo no menu de restauração
            backup_path = rf'Secretaria de Estado da Educação\Cefope - Equipe Tecnologia\Projetos Python\Automação Plataforma Estudante\automacao_web\backups\geral\{nome_backup}.mbz'
            status_restauracao = lib_functions.send_file(backup_path)
            lib_functions.verifica_load('media/verifica_carregamento_upload_arquivo.png', 0.9)

            if status_restauracao == 'success':
                dflogexecutados.at[dflogexecutados.index[-1], 'nome_backup'] = nome_backup
                pos_screen = pa.locateCenterOnScreen('media/restauracao-botao-restaurar.png', confidence=0.9)
                pa.click(pos_screen, duration=0.5)

                # configuração das opções de restauração
                lib_functions.verifica_load(local_carregamento_PTFE, .9) # pausa para a página carregar
                # para o ctrl end funcionar, a rolagem tem que ser a interna da página(há dois scrolls)
                pos_screen = pa.locateCenterOnScreen('media/restauracao-detalhes-do-backup.png', confidence=0.9)
                pa.click(pos_screen)
                pa.hotkey('ctrl', 'end')
                time.sleep(.5)  # pausa para o programa rolar até o final da página com segurança
                pos_screen = pa.locateCenterOnScreen('media/restauracao-botao-continuar.png', confidence=0.9)
                pa.click(pos_screen, duration=0.5)

                lib_functions.verifica_load(local_carregamento_PTFE, .9) # pausa para a página carregar
                # para o scroll end funcionar, a rolagem tem que ser a interna da página(há dois scrolls)
                pos_screen = pa.locateCenterOnScreen('media/restauracao-restaurar-como-um-novo-curso.png',
                                                     confidence=0.9)
                pa.click(pos_screen)
                pa.scroll(-1200)
                time.sleep(1)  # pausa para scrolar
                pos_screen = lib_functions.mult_img_selection(['media/restauracao-opcao-restauracao-excluir-1.png', 'media/restauracao-opcao-restauracao-excluir-2.png'], confidence_level=0.95)
                pa.click(pos_screen, duration=0.5)
                pos_screen = pa.click(pos_screen[0] + 160, pos_screen[1] + 50, duration=0.5)
                pa.click(pos_screen, duration=0.5)
                lib_functions.verifica_load(local_carregamento_PTFE, .9) # pausa para a página carregar

                pos_screen = pa.locateCenterOnScreen('media/restauracao-texto-restaurar-configuracoes.png', confidence=0.95)
                pa.click(pos_screen, duration=.5)
                pa.hotkey('ctrl', 'end')
                time.sleep(.5)  # pausa para rolar até o final da página
                pos_screen = pa.locateCenterOnScreen('media/restauracao-botao-proximo.png', confidence=0.9)
                pa.click(pos_screen, duration=0.5)
                lib_functions.verifica_load(local_carregamento_PTFE, .9) # pausa para a página carregar

                pos_screen = pa.locateCenterOnScreen('media/restauracao-texto-configuracoes-do-curso.png', confidence=0.95)
                pa.click(pos_screen)
                pa.hotkey('ctrl', 'end')
                time.sleep(.5)  # pausa para rolar até o final da página
                pos_screen = pa.locateCenterOnScreen('media/restauracao-botao-proximo.png', confidence=0.9)
                pa.click(pos_screen, duration=0.5)
                lib_functions.verifica_load(local_carregamento_PTFE, .9) # pausa para a página carregar

                pos_screen = pa.locateCenterOnScreen('media/restauracao-texto-restaurar-configuracoes.png', confidence=0.95)
                pa.click(pos_screen)
                pa.hotkey('ctrl', 'end')
                time.sleep(.5)  # pausa para rolar até o final da página
                pos_screen = pa.locateCenterOnScreen('media/restauracao-botao-executar-restauracao.png', confidence=0.7)
                pa.click(pos_screen, duration=0.5)
                lib_functions.verifica_load(local_carregamento_PTFE, .9) # pausa para a página carregar

                pos_screen = pa.locateCenterOnScreen('media/restauracao-botao-continuar.png', confidence=0.9)
                pa.click(pos_screen, duration=.5)
                lib_functions.verifica_load(local_carregamento_PTFE, .9) # pausa para a página carregar

                # categorização

                # abertura do menu de configurações para fazer a categorização
                pos_screen = pa.locateCenterOnScreen(local_imagem_botao_config, confidence=0.9)
                pa.click(pos_screen, duration=.5)
                pos_screen = pa.locateCenterOnScreen(local_imagem_confimacao, confidence=0.9)
                pa.click(pos_screen, duration=.5)
                lib_functions.verifica_load(local_carregamento_PTFE, .9) # pausa para a página carregar

                # digitação da categoria correspondente
                pos_screen = pa.locateOnScreen(local_imagem_lista_categoria, confidence=0.9)
                pa.click(pos_screen.width + pos_screen.left, pos_screen.height + pos_screen.top, duration=.5)
                time.sleep(3)  # pausa para carregar as opções, está muito lagado
                lib_functions.write_special_char(nome_categoria)
                time.sleep(3)  # pausa para digitar
                pa.press('enter')
                pa.press('tab')
                pa.press('tab')
                time.sleep(2)  # pausa para a tela sofrer todas as alterações
                try:
                    pos_screen = pa.locateOnScreen('media/categorizacao-sem-categoria.png', confidence=0.9)
                    if pos_screen is not None:
                        pos_screen = pa.locateOnScreen('media/restauracao-texto-editar-configuracoes-curso.png', confidence=0.9)
                        pa.click(pos_screen)
                        pos_screen = pa.locateOnScreen(local_imagem_lista_categoria, confidence=0.7)
                        pa.click(pos_screen.width + pos_screen.left, pos_screen.height + pos_screen.top, duration=.5)
                        time.sleep(3)  # pausa para carregar as opções, está muito lagado
                        pa.hotkey('ctrl', 'a', 'backspace')
                        time.sleep(3)
                        lib_functions.write_special_char('Sem Categoria (manuais) / Categorização')
                        dflogexecutados.at[dflogexecutados.index[-1], 'nome_categoria'] = 'erro_categorização'
                        time.sleep(3)  # pausa para digitar
                        pa.press('enter')
                        pa.press('tab')
                        pa.hotkey('ctrl', 'end')
                        time.sleep(.5)  # pausa para executar o comando acima
                        lib_functions.send_course_image(nome_imagem, dflogexecutados)  # envio da imagem do curso. Espera-se que, como não houve erro de restauração, o erro de categorização não impacte no resultado do envio
                        lib_functions.verifica_load('media/verifica_carregamento_upload_imagem.png', 0.8)
                        dflogexecutados.at[dflogexecutados.index[-1], 'nome_imagem'] = nome_imagem
                        pos_screen = pa.locateCenterOnScreen(local_imagem_botao_salvar, confidence=0.9)
                        pa.click(pos_screen, duration=0.5)
                        lib_functions.verifica_load(local_carregamento_PTFE, .9) # pausa para a página carregar
                except pa.ImageNotFoundException:
                    dflogexecutados.at[dflogexecutados.index[-1], 'nome_categoria'] = nome_categoria
                    pa.hotkey('ctrl', 'end')
                    time.sleep(.5)  # pausa para executar o comando acima
                    lib_functions.send_course_image(nome_imagem, dflogexecutados)  # envio da imagem do curso.
                    lib_functions.verifica_load('media/verifica_carregamento_upload_imagem.png', 0.8)
                    dflogexecutados.at[dflogexecutados.index[-1], 'nome_imagem'] = nome_imagem
                    pos_screen = pa.locateCenterOnScreen(local_imagem_botao_salvar, confidence=0.9)
                    pa.click(pos_screen, duration=0.5)
                    lib_functions.verifica_load(local_carregamento_PTFE, .9) # pausa para a página carregar

            else:
                # abrir configuracoes para categorizar na categoria do erro de restauração, não faz o envio da imagem. Como a restauração já deu erro, espera-se que no envio da imagem também ocorra.
                dflogexecutados.at[dflogexecutados.index[-1], 'nome_backup'] = 'erro_restauracao'
                pos_screen = pa.locateCenterOnScreen(local_imagem_botao_config, confidence=0.9)
                pa.click(pos_screen, duration=.5)
                pos_screen = pa.locateCenterOnScreen(local_imagem_confimacao, confidence=0.9)
                pa.click(pos_screen, duration=.5)
                lib_functions.verifica_load(local_carregamento_PTFE, .9) # pausa para a página carregar

                # digitação da categoria correspondente
                pos_screen = pa.locateOnScreen(local_imagem_lista_categoria, confidence=0.9)
                pa.click(pos_screen.width + pos_screen.left, pos_screen.height + pos_screen.top, duration=.5)
                time.sleep(3)  # pausa para carregar as opções, está muito lagado
                lib_functions.write_special_char(status_restauracao)
                dflogexecutados.at[dflogexecutados.index[-1], 'nome_categoria'] = 'erro_restauracao'
                dflogexecutados.at[dflogexecutados.index[-1], 'nome_imagem'] = 'erro_restauracao'
                time.sleep(3)  # pausa para digitar
                pa.press('enter')
                # para o ctrl end funcionar, a rolagem tem que ser a interna da página(há dois scrolls)
                pa.press('tab')
                pa.hotkey('ctrl', 'end')
                time.sleep(.5)  # pausa para executar o comando acima
                pos_screen = pa.locateCenterOnScreen(local_imagem_botao_salvar, confidence=0.9)
                pa.click(pos_screen, duration=0.5)
                lib_functions.verifica_load(local_carregamento_PTFE, .9) # pausa para a página carregar
    finally:
        # Bloco de código para salvar os registros do log, independentemente de ocorrer um erro ou não
        data_hora_atual = datetime.now()
        data_log = data_hora_atual.strftime("%Y%m%d")  # Formato: AAAAMMDD
        hora_log = data_hora_atual.strftime("%H%M%S")  # Formato: HHMMSS
        nome_arquivo_log = f'log_exclusivos_{data_log}_{hora_log}.csv'
        dflogexclusivos.to_csv(f'logs/{nome_arquivo_log}', sep=';', index=False)
        nome_arquivo_log = f'log_executados_{data_log}_{hora_log}.csv'
        dflogexecutados.to_csv(f'logs/{nome_arquivo_log}', sep=';', index=False)

