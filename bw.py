import pyautogui as pg 
import time


def bw(initial_date, final_date):
    pg.hotkey('winleft')
    time.sleep(1)
    pg.typewrite('Analysis for Microsoft Excel\n', 0.2)

    time.sleep(15)

    pg.moveTo(1050, 70,duration=0.5)
    pg.doubleClick()

    pg.moveTo(50, 140,duration=0.5)
    pg.click()

    pg.moveTo(90, 242,duration=0.5)  # 242 -> 262
    pg.click()

    time.sleep(5)

    pg.typewrite('Mat.Mai$10\n')

    time.sleep(20)

    pg.moveTo(950, 410,duration=1)
    pg.click()
    pg.typewrite(final_date)
    pg.moveTo(1350, 410,duration=1)
    pg.click()
    pg.typewrite(initial_date)
    pg.moveTo(950, 710,duration=1)
    pg.click()
    pg.typewrite('CO; PER; AR; CL; MX')
    pg.moveTo(1610, 855,duration=1)
    pg.click()


