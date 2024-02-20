import pandas as pd
from datetime import datetime

def update_csv(tipo_log):
    """
    adapte para a quantidade de arquivos csv que existirem
    mescla os csv de diferentes execuções com base no nome breve. df1 deve ser o da primeira execução, df3 deve ser o da última execução

    Esta função tem como objetivo mesclar múltiplos arquivos CSV com base em um campo comum (shortname) e atualizar as informações em um único arquivo CSV.
    Ela lê os dados de três arquivos CSV diferentes correspondentes a diferentes execuções e os mescla em um único DataFrame.
    Os arquivos são identificados pelo tipo de log (tipo_log) passado como parâmetro.
    Em seguida, substitui as informações relevantes com base nos dados do terceiro arquivo CSV.
    Por fim, salva o DataFrame resultante em um novo arquivo CSV.

    Parameters:
        - tipo_log: Uma string indicando o tipo de log a ser processado. Pode ser "executados" ou "exclusivos".

    - Autor: Almir
    """

    df1 = pd.read_csv('logs/log_geral_executados_20240215_170046.csv', encoding = 'utf-8', sep = ';') #substitua o nome do arquivo
    df2 = pd.read_csv('logs/log_geral_executados_recategorizacao_20240216_115839.csv', encoding = 'utf-8', sep = ';') #substitua o nome do arquivo
    df3 = pd.read_csv('logs/log_geral_executados_ensino_medio_20240216_175704.csv', encoding = 'utf-8', sep = ';') #substitua o nome do arquivo

    df_final = pd.merge(df1, df2, on='shortname', how='left', validate='one_to_one')
    df_final['nome_backup'] = df_final['nome_backup_y'].fillna(df_final['nome_backup_x'])
    df_final['nome_categoria'] = df_final['nome_categoria_y'].fillna(df_final['nome_categoria_x'])
    df_final['nome_imagem'] = df_final['nome_imagem_y'].fillna(df_final['nome_imagem_x'])
    df_final = df_final.drop(['nome_backup_x', 'nome_backup_y', 'nome_categoria_x', 'nome_categoria_y', 'nome_imagem_x', 'nome_imagem_y'], axis=1)

    # Substituir as informações resultantes com base no terceiro CSV
    df_final = pd.merge(df_final, df3, on='shortname', how='left', validate='one_to_one')
    df_final['nome_backup'] = df_final['nome_backup_y'].fillna(df_final['nome_backup_x'])
    df_final['nome_categoria'] = df_final['nome_categoria_y'].fillna(df_final['nome_categoria_x'])
    df_final['nome_imagem'] = df_final['nome_imagem_y'].fillna(df_final['nome_imagem_x'])
    df_final = df_final.drop(['nome_backup_y', 'nome_categoria_y', 'nome_imagem_y', 'nome_backup_x', 'nome_categoria_x', 'nome_imagem_x'], axis=1)

    #data e hora
    data_hora_atual = datetime.now()
    data_log = data_hora_atual.strftime("%Y%m%d")  # Formato: AAAAMMDD
    hora_log = data_hora_atual.strftime("%H%M%S")  # Formato: HHMMSS

    # Salvar o DataFrame final como um novo CSV
    nome_arquivo_log = f'log_mesclado_{tipo_log}_{data_log}_{hora_log}.csv'
    df_final.to_csv(f'logs/{nome_arquivo_log}', sep=';', index=False)


update_csv('executados') #substitua o nome do arquivo
#update_csv('exclusivos')

