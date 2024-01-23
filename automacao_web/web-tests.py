import pandas as pd

df = pd.read_csv('uploadCSV/templatefinal.csv', sep=';')

df['chave_backup'] = df['shortname'].apply(lambda x: 'Práticas inovadoras n sei o q' if '_PIP_' in x else '')

df.to_csv('uploadCSV/templatefinalchaves.csv', sep=';', index=False)