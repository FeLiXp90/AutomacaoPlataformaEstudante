import pandas as pd
import pymsgbox as pmsg
import tkinter as tk
import sys
from datetime import datetime
from tkinter import filedialog

programa_encerrado = 'Obrigado. Programa encerrado'
encerrar_programa = 'Encerrar programa'
PTFE = 'PTFE v2.0'

def escolher_tipo_log():
    """
    Janela para escolha do tipo de arquivo a ser atualizado.

    - Autor: Almir
    """
    while True:
        tipo_planilha = pmsg.confirm(text='Escolha o tipo de log que deseja atualizar\n as informações de execução:', title=PTFE, buttons=['Log Executados', 'Log Exclusivos', encerrar_programa])

        if tipo_planilha == 'Log Executados':
            return 'executados'
        elif tipo_planilha == 'Log Exclusivos':
            return 'exclusivos'
        elif tipo_planilha == encerrar_programa:
            pmsg.alert(text=f'{programa_encerrado}', title=PTFE)
            sys.exit()

def abrir_seletor_arquivos():
    """
    Seletor de arquivo

    - Autor: Almir
    """
    while True:
        # Cria uma instância do tkinter.Tk() para ser a janela principal
        root = tk.Tk()
        root.withdraw()  # Oculta a janela principal

        # Abre o seletor de arquivos e retorna o caminho do arquivo selecionado
        caminho_arquivo = filedialog.askopenfilename()

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

def update_csv(tipo_log):
    """
    mescla os csv de diferentes execuções com base no nome breve. df1 deve ser o da primeira execução, df2 deve ser o da próxima execução e execução

    Esta função tem como objetivo mesclar múltiplos arquivos CSV com base em um campo comum (shortname) e atualizar as informações em um único arquivo CSV.
    Ela lê os dados de três arquivos CSV diferentes correspondentes a diferentes execuções e os mescla em um único DataFrame.
    Os arquivos são identificados pelo tipo de log (tipo_log) passado como parâmetro.
    Em seguida, substitui as informações relevantes com base nos dados do terceiro arquivo CSV.
    Por fim, salva o DataFrame resultante em um novo arquivo CSV.

    Parameters:
        - tipo_log: Uma string indicando o tipo de log a ser processado. Pode ser "executados" ou "exclusivos".

    - Autor: Almir
    """
    pmsg.alert(text='Escolha o log antigo.', title=PTFE)
    path_df1 = abrir_seletor_arquivos()
    pmsg.alert(text='Escolha o log novo.', title=PTFE)
    path_df2 = abrir_seletor_arquivos()


    df1 = pd.read_csv(rf'{path_df1}', encoding = 'utf-8', sep = ';') #substitua o nome do arquivo
    df2 = pd.read_csv(rf'{path_df2}', encoding = 'utf-8', sep = ';') #substitua o nome do arquivo

    df_final = pd.merge(df1, df2, on='shortname', how='left', validate='one_to_one')
    df_final['nome_backup'] = df_final['nome_backup_y'].fillna(df_final['nome_backup_x'])
    df_final['nome_categoria'] = df_final['nome_categoria_y'].fillna(df_final['nome_categoria_x'])
    df_final['nome_imagem'] = df_final['nome_imagem_y'].fillna(df_final['nome_imagem_x'])
    df_final = df_final.drop(['nome_backup_x', 'nome_backup_y', 'nome_categoria_x', 'nome_categoria_y', 'nome_imagem_x', 'nome_imagem_y'], axis=1)

    #data e hora
    data_hora_atual = datetime.now()
    data_log = data_hora_atual.strftime("%Y%m%d")  # Formato: AAAAMMDD
    hora_log = data_hora_atual.strftime("%H%M%S")  # Formato: HHMMSS

    # Salvar o DataFrame final como um novo CSV
    nome_arquivo_log = f'log_mesclado_{tipo_log}_{data_log}_{hora_log}.csv'
    df_final.to_csv(f'logs/{nome_arquivo_log}', sep=';', index=False)

tipo_log = escolher_tipo_log()

if tipo_log == 'executados':
    update_csv('executados')
else:
    update_csv('exclusivos')

