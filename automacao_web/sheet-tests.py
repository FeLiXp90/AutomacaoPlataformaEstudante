import tkinter as tk, pandas as pd, pymsgbox as pmsg, sys, os
from tkinter import filedialog
from datetime import datetime
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
                                       buttons=['Confirmo', 'Escolher novamente', 'Encerrar programa'])
            if confirmacao == 'Confirmo':
                return caminho_arquivo
            elif confirmacao == 'Encerrar programa':
                pmsg.alert(text='Obrigado. Programa encerrado',
                           title='PTFE v1.0')
                sys.exit()
        else:
            escolha = pmsg.confirm(text='Oops, nenhum arquivo selecionado.',
                                   title='Erro - [PTFE v1.0]',
                                   buttons=['Escolher novamente', 'Encerrar programa'])
            if escolha == 'Encerrar programa':
                pmsg.alert(text='Obrigado. Programa encerrado',
                           title='PTFE v1.0')
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