import pyautogui as pa
import time

time.sleep(2)

#habilitar data
try:
    pos_screen = pa.locateCenterOnScreen('media/habilitar-data-termino-curso.png', confidence=0.8)
    pa.click(pos_screen, duration = .5)
except pa.ImageNotFoundException:
    pass
    print('achei n boy')