from ctypes.wintypes import HWND
import pygetwindow as gw
import win32gui, win32ui, win32con, win32api
import cv2 as cv
import numpy as np
import os
from time import time, sleep
import pyautogui as pag
import pywinauto
from pywinauto.application import Application
from pywinauto. keyboard import send_keys


def activateWindow(appName, detect):
    if pag.locateOnScreen(detect) == None:
        activate = gw.getWindowsWithTitle(appName)[0]
        activate.activate()
    else:
        print(pag.locateOnScreen(detect))





# def click_pg(x, y):
    # hwnd = win32gui.FindWindow(None, 'SAP Logon 750')

    # win = win32ui.CreateWindowFromHandle(hwnd)

    # hwnd = win32gui.FindWindowEx(hwnd, None, None, None)

    # print("Find window", win32gui.FindWindow(None, 'SAP Logon 750'))

    # click = win32api.MAKELONG(x, y)
    # win32gui.SendMessage(hwnd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, click)
    # win32gui.SendMessage(hwnd, win32con.WM_LBUTTONUP, None, click)

# def winEnumHandler( hwnd, ctx ):
#     if win32gui.IsWindowVisible( hwnd ):
#         print ( hex( hwnd ), win32gui.GetWindowText( hwnd ) )

# win32gui.EnumWindows( winEnumHandler, None )

# numpad = [[1268, 245], [1480, 150]]
# top left [960, 0], bottom right [1920, 1032]
# while True:
#     for i in numpad:
#         click_pg(i[0], i[1])
#         sleep(5)

# activateWindow("WhatsApp", './cimg/detect.png')
# greendot = pag.locateCenterOnScreen('./cimg/greendot.png', confidence=0.5)
# click_pg(greendot[0], greendot[1])
# print(greendot)


#############################################################


# def main(x, y):
#     window_name = "SAP Logon 750"
#     hwnd = win32gui.FindWindow(None, window_name)
    # hwnd = get_inner_windows(hwnd)['Khairat Production']
    # win = win32ui.CreateWindowFromHandle(hwnd)

    #win.SendMessage(win32con.WM_CHAR, ord('A'), 0)
    #win.SendMessage(win32con.WM_CHAR, ord('B'), 0)
    #win.SendMessage(win32con.WM_KEYDOWN, 0x1E, 0)
    #sleep(0.5)
    #win.SendMessage(win32con.WM_KEYUP, 0x1E, 0)
    # win32api.SendMessage(hwnd, win32con.WM_KEYDOWN, 0x41, 0)
    # sleep(0.5)
    # win32api.SendMessage(hwnd, win32con.WM_KEYUP, 0x41, 0)

    # click = win32api.MAKELONG(x, y)
    # win.SendMessage(win32con.WM_LBUTTONDOWN)
    # sleep(0.7)
    # win.SendMessage(win32con.WM_LBUTTONUP)
    # win32gui.SendMessage(hwnd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, click)
    # win32gui.SendMessage(hwnd, win32con.WM_LBUTTONUP, None, click)



def list_window_names():
    def winEnumHandler(hwnd, ctx):
        if win32gui.IsWindowVisible(hwnd):
            print(hex(hwnd), '"' + win32gui.GetWindowText(hwnd) + '"')
    win32gui.EnumWindows(winEnumHandler, None)


def get_inner_windows(whndl):
    def callback(hwnd, hwnds):
        if win32gui.IsWindowVisible(hwnd) and win32gui.IsWindowEnabled(hwnd):
            hwnds[win32gui.GetClassName(hwnd)] = hwnd
        return True
    hwnds = {}
    win32gui.EnumChildWindows(whndl, callback, hwnds)
    return hwnds

def main():
    # programPath = r"C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe"
    # filePath = r"C:\Users\HA\Desktop\Work\52414023.pdf"
    # app = Application().start(r'{} "{}"'.format(programPath, filePath))
    # sleep(1)
    # send_keys('^a^p')
    # w_handle = pywinauto.findwindows.find_windows(title="52414023.pdf - Adobe Acrobat Reader DC (64-bit)")[0]
    # w_handle = gw.getWindowsWithTitle('SAP')
    # print(w_handle)
    # window=app.window(handle=w_handle)
    # window.wait('ready', timeout=10)
    # window[u'Pri&nter:comboBox'].select(0)
    # window[u'&Properties'].click()
    # sleep(8)
    # print(w_handle)

    # w_handle = pywinauto.findwindows.find_windows(title=u'SAP Logon 750', class_name='#32770')[0]
    # pwa_app = pywinauto.application.Application().connect(title=u'SAP Logon 750')
    # print(pwa_app)
    # window = pwa_app.window(handle=w_handle)
    # print(window)
    # window.set_focus()
    # ctrl = window['ListView']
    # ctrl.select(0)
    # ctrl.click()
    # send_keys("{VK_SPACE}"
    # "{ENTER}"
    # "{ENTER}")

    app = Application(backend='uia').connect(title_re='.*SAP Logon*.')
    main = app.window(title_re='.*SAP Logon*.')
    main.set_focus()
    # main.print_control_identifiers()
    main.child_window(title="Khairat Production", control_type="ListItem").double_click_input()
    

if __name__ == '__main__':
    main()