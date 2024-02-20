import pandas as pd, os
from datetime import datetime

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

    user_path = os.path.expanduser("~")
    logs_path = r'Secretaria de Estado da Educação\Cefope - Equipe Tecnologia\Projetos Python\Automação Plataforma Estudante\automacao_web\logs'
    folder_path = os.path.join(user_path, logs_path)
    dfs = []

    for arquivo in os.listdir(folder_path):
        if arquivo.startswith(f'log_{tipo_log}_') and arquivo.endswith('.csv'):
            file_path = os.path.join(folder_path,'', arquivo)
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

log_geral('executados')
log_geral('exclusivos')

