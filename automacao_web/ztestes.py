import pyautogui as pa
import time
import sys
import pymsgbox as pmsg

def mult_img_selection(img_list, confidence_level):
    #seleção entre diferentes imagens possíveis presentes na tela
    for i in img_list:
        try:
            pos_arq = pa.locateCenterOnScreen(i, confidence = confidence_level)
            return pos_arq
            break
        except pa.ImageNotFoundException:
            pmsg.alert(text='ALERTA. Nenhuma imagem foi encontrada: \nReveja o código.',
                       title='PTFE v1.0')
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



#Pré-visualização de cursos carregados
time.sleep(1) #pausa para carregar a próxima guia com segurança
pos_screen = pa.locateCenterOnScreen('media/previsualizar_titulo-click-ctrl-end.png', confidence = 0.7)
pa.click(pos_screen)
pa.hotkey('ctrl','end')
time.sleep(.2) #pausa para ir ao final da tela sem complicações

select_drop_down_list_dual('media/previsualizar_forcar-modalidade-grupo.png', 0.8, 'down')

#Livro de notas = não
select_drop_down_list_dual('media/previsualizar_mostrar-livro-notas.png', 0.8, 'down')

#Formato de blocos
formato = pmsg.confirm(text=f'Selecione o formato do curso: Blocos ou Tópicos.', title='Formato de curso - [PTFE v1.0]', buttons=['Blocos', 'Tópicos', 'Encerrar programa'])

if formato == 'Blocos':
    pa.click(549,360, duration = .5)
    time.sleep(1)
    pa.click(mult_img_selection(['media/previsualizar_formato-curso-blocos-op-1.png', 'media/previsualizar_formato-curso-blocos-op-2.png'], confidence_level = 0.9), duration = .5)
# elif formato == 'Encerrar programa':
#     pmsg.alert(text='Obrigado. Programa encerrado',
#                title='PTFE v1.0')
#     sys.exit()
# else:
#     pa.click(541, 377, duration=.5)
#     mult_img_selection( ['media/previsualizar_formato-curso-topicos-op-1.png', 'media/previsualizar_formato-curso-topicos-op-2.png'], confidence_level=0.98)
#
# #habilitar data
# try:
#     pos_screen = pa.locateCenterOnScreen('media/habilitar-data-termino-curso.png', confidence=0.95)
#     pa.click(pos_screen, duration = .5)
# except pa.ImageNotFoundException:
#     pass

# #Sem categoria
# pa.scroll(1000)
# time.sleep(.2) #para dar tempo do scroll exectuar tranquilamente
#
#     #confirmação de que os cursos serão alocados no 'Sem Categoria', apenas para dar mais segurança
# try:
#     pos_screen = pa.locateCenterOnScreen('media/sem-categoria.png', confidence = 0.7)
#     if pos_screen is not None:
#         pass
# except pa.ImageNotFoundException:
#     pmsg.alert(text='ALERTA. Os cursos não seriam categorizados: \nReveja o código e soluções.',
#                title='PTFE v1.0')
#     sys.exit()
#
# #carregar cursos
# select_drop_down_list_exact('media/list-susp-carregamento.png', ['media/op-list-susp-carregamento-1.png',
#                         'media/op-list-susp-carregamento-2.png'], 0.95, 0.9)
#
# #confirmar
# pa.hotkey('ctrl', 'end')
# time.sleep(.2) #pausa para ir ao final da tela sem problemas
#
# pos_screen = pa.locateCenterOnScreen('media/previsualizar_carregar-cursos.png', confidence = 0.9)
# pa.moveTo(pos_screen, duration = 1)
#
