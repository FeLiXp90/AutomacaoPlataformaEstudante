import pandas as pd, os
from datetime import datetime
import pymsgbox as pmsg
import sys

programa_encerrado = 'Obrigado. Programa encerrado'
encerrar_programa = 'Encerrar programa'
PTFE = 'PTFE v2.0'

def escolher_tipo_log():
    """
    Janela para escolha do tipo de arquivo a ser atualizado.

    - Autor: Almir
    """
    while True:
        tipo_planilha = pmsg.confirm(text='Escolha o tipo de log que deseja unir\n as informações de execução:', title=PTFE, buttons=['Logs Executados', 'Logs Exclusivos', encerrar_programa])

        if tipo_planilha == 'Logs Executados':
            return 'executados'
        elif tipo_planilha == 'Logs Exclusivos':
            return 'exclusivos'
        elif tipo_planilha == encerrar_programa:
            pmsg.alert(text=f'{programa_encerrado}', title=PTFE)
            sys.exit()


def log_geral(tipo_log):

    """
    Esta função tem como objetivo unir vários arquivos CSV de logs gerados por erros em um único arquivo CSV.
    Ela procura por arquivos CSV no diretório especificado e identifica os arquivos com base no tipo de log (tipo_log) passado como parâmetro.
    Em seguida, lê e concatena os dados desses arquivos em um único DataFrame, removendo duplicatas com base no campo shortname.
    Por fim, salva o DataFrame resultante em um novo arquivo CSV.

    Parameters:
        - tipo_log: Uma string indicando o tipo de log a ser processado. Pode ser "executados" ou "exclusivos".

    - Autor: Almir
    """
    pmsg.alert(text='ATENÇÃO. Para esse tipo de execução \né importante que os arquivos que deseja unir estejam todos na pasta "logs/unir".', title='ATENÇÃO - [PTFE v1.0]')
    opcao = pmsg.confirm(text='Você atende às configurações \nmencionadas anteriormente?', title=PTFE, buttons=['Sim', 'Não'])
    if opcao == 'Não':
        pmsg.alert(text=f'Salve os arquivos corretamente.\nEntão, tente novamente.\n\n{programa_encerrado}', title=PTFE)
        sys.exit()

    dfs = []
    try:
        for arquivo in os.listdir('logs/01 - unir'):
            if arquivo.startswith(f'log_{tipo_log}_') and arquivo.endswith('.csv'):
                file_path = os.path.join('logs/01 - unir','', arquivo)
                dftemp = pd.read_csv(file_path, encoding='utf-8', sep=';')
                dfs.append(dftemp)

        df_geral = pd.concat(dfs, ignore_index = True)
        df_geral = df_geral.drop_duplicates(subset = 'shortname', keep = 'last')

        #salvando o arquivo de logs executados
        data_hora_atual = datetime.now()
        data_log = data_hora_atual.strftime("%Y%m%d")  # Formato: AAAAMMDD
        hora_log = data_hora_atual.strftime("%H%M%S")  # Formato: HHMMSS
        nome_arquivo_log = f'log_geral_{tipo_log}_{data_log}_{hora_log}.csv'
        df_geral.to_csv(f'logs/{nome_arquivo_log}', sep=';', index=False)
    except ValueError:
        pmsg.alert(text=f'Parece que não há nenhum arquivo do tipo selecionado\nna pasta "logs/unir".\n\nVerifique novamente a pasta referida com os arquivos atualizados.\n\n{programa_encerrado}', title = PTFE)
        sys.exit()

tipo_log = escolher_tipo_log()

if tipo_log == 'executados':
    log_geral('executados')
else:
    log_geral('exclusivos')

