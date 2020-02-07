#-*- coding:utf-8 -*-

import pyautogui, time, pyperclip
from pynput import mouse

import pandas as pd
import numpy as np

def process_(Series, dbox, sbutton):
    # time.sleep(0.5)

    # pyautogui.moveTo('./date.PNG')
    pyautogui.moveTo(dbox)
    pyautogui.move(50,0)
    pyautogui.click()
    pyautogui.write(str(Series[0]))

    pyautogui.press('tab')
    pyautogui.write(str(Series[1]))
    pyautogui.press('tab', 3)
    pyautogui.press('up', 20)
    pyautogui.press('down', Series[2])
    pyautogui.press('tab')
    pyperclip.copy(Series[4])
    # print(Series[4].decode('EUC-KR'))
    # pyperclip.copy(unicode(str(Series[4]),"UTF-8"))
    pyautogui.hotkey('ctrl','v')
    # pyautogui.write('헬로우 월드')
    pyautogui.press('tab')
    if Series[5] == 1:
        pyautogui.press('up', 1)
        pyautogui.press('tab')
        pyautogui.press('up',29)
        pyautogui.press('down',29)

    # pyautogui.click('./saveButton.PNG')
    pyautogui.click(sbutton)

    pyautogui.hotkey('alt','y')

if __name__ == "__main__":

    dateBoxPoint = pyautogui.center(pyautogui.locateOnScreen('./date.PNG'))
    saveButtonPoint = pyautogui.center(pyautogui.locateOnScreen('./saveButton.PNG'))
    filepath = './data.xlsx'
    data = pd.read_excel(filepath, encoding='utf-8')

    print ('macro start!')

    for i in data.values:
        process_(i, dateBoxPoint, saveButtonPoint)

    print ('macro complete!')


