"""
Código:
    0 Melhorar o failsafe error do pyautogui durante as reexecuções web
    0.0 mult_img_select está cortando o programa, consertar a gambiarra
    1 Quando o notebook chegar, implementar as alterações para que o código rode localmente nele
    2 Fazer build executável
    2.2 Fazer função para reexecutar os cursos do "Sem_Categoria (manuais) / Categorização e Restauração"
    3 Documentar melhor as funções e seus argumentos (documentar como na função select_drop_down_list_dual)
    4 Na parte de restauracao, tem muito cóadigo repetido na parte de dar ctrl end e apertar o botão, fazer uma função para reduzir o código
    5 Caso a restauração dê erro na seleção de arquivo: código repetido, fazer função?
    7 Revisar o tempo dos time.sleep (revisão geral dos time.sleep, se alterar, alterar para mais tempo. Não reduzir!!!)
    8 Reduzir o número de variáveis de posições para click (colocar pos_screen onde der, deixa que eu (almir) faço isso
    9 Sonar lint
    10 renomear as imagens de forma que fiquem mais organizadas.: ex, 'partedoproceso_descricao-da-imagem' (pode deixar que eu (almir) faço isso

Processo:
    1 Avaliar o impacto das restaurações no armazenamento do servidor

Informações:
    2 Status dos arquivos de backup:
        EM 1a PROINPAC - Sala inteira
        EM 2a PROINPAC - Sala inteira
        EM 3a PROINPAC - 1o Trimestre (a sala inteira não estava modelada)
        EM EO - 1o Trimestre (a sala inteira excede o limite de upload)
        EM PV - Sala inteira
"""

import lib_code_automacao_PTFE as lib_code
import pymsgbox as pmsg, sys  # tratamento de dados


programa_encerrado = 'Obrigado. Programa encerrado'
encerrar_programa = 'Encerrar programa'
login_proxy = 'Login Proxy - [PTFE v1.0]'
PTFE = 'PTFE v2.0'

pmsg.alert(text=f'Verifique a atualização dos arquivos de backup\nVerifique a renomeação das abas das planilhas\nVerifique a qualidade dos dados, especialmente dos nomes breves', title=PTFE)

execution_type = pmsg.confirm(text='Qual tipo de execução você gostaria de fazer?', title='Tipo de execução- [PTFE v1.0]', buttons=['Apenas processamento de dados', 'Apenas parte web', 'Completa (dados + web)', encerrar_programa])
if execution_type == 'Apenas processamento de dados':
    lib_code.tratamento_dados()
    pmsg.alert(text=f'{programa_encerrado}', title=PTFE)

elif execution_type == 'Apenas parte web':
    pmsg.alert(text='ATENÇÃO. Para esse tipo de execução \né importante que o Chrome já esteja aberto\ne que já esteja logado no proxy e na plataforma do estudante.\n\nAlém disso, é importante que o chrome esteja\nconfigurado para não permitir abertura de caixas\nde diálogo adicionais.',title='ATENÇÃO - [PTFE v1.0]')
    opcao = pmsg.confirm(text='Você atende às configurações \nmencionadas anteriormente?', title='Tipo de execução- [PTFE v1.0]', buttons=['Sim', 'Não'])
    if opcao == 'Não':
        pmsg.alert(text=f'Abra o Chrome e logue no proxy.\nEntão, tente novamente.\n\n{programa_encerrado}', title=PTFE)
        sys.exit()
    else:
        while True:
            try:
                lib_code.web()
            except Exception as e:
                print(f'Ocorreu um erro: {e}')

elif execution_type == 'Completa (dados + web)':
    lib_code.tratamento_dados()
    while True:
        try:
            lib_code.web()
        except Exception as e:
            print(f'Ocorreu um erro: {e}')

else:
    pmsg.alert(text=f'{programa_encerrado}', title=PTFE)
    sys.exit()