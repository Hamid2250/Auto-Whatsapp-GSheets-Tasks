from ctypes.wintypes import HWND
from dataclasses import dataclass
import subprocess
import os
import re
import psutil
# import win32gui, win32api, win32con
import pywinauto
from pywinauto import Desktop
import pyautogui as pag
import pygetwindow as gw
from time import sleep
from pyrobogui import robo
import pyperclip as pc
import arabic_reshaper
from langdetect import detect_langs
import datetime
import gspread
from replays import process_message


now = datetime.datetime.now()
# connect to google sheets
service_account = gspread.service_account(filename='gs_account.json')
sheet = service_account.open('Ai workflow')
otl = sheet.worksheet('OrdersTaskList')


def isOpen(appName):
    running = any(appName in x for x in gw.getAllTitles())
    print(appName, "running")
    return running

def isNotOpen(appName):
    if isOpen(appName) == False:
        pag.hotkey('win')
        pag.write(appName)
        pag.hotkey('enter')
        sleep(2)
        while isOpen(appName) == False:
            print('Waiting to run',appName)
            sleep(2)
        print(appName,'running')

def activateWindow(appName):
    if any(appName in x for x in gw.getAllTitles()):
        activate = gw.getWindowsWithTitle(appName)[0]
        activate.activate()

def activeBrowserTab(tabNotActive, tabActive):
    if pag.locateCenterOnScreen(tabNotActive) is None:
        pag.click(pag.locateCenterOnScreen(tabNotActive, confidence=0.7))
    else:
        pag.click(pag.locateCenterOnScreen(tabActive, confidence=0.7))

######### WhatsApp Automation ########   
def navImage(image,clicks,off_x=0,off_y=0):
    position=pag.locateCenterOnScreen(image,confidence=0.8)
    if position is None:
        return 0
    else:
        pag.moveTo(position)
        pag.moveRel(off_x,off_y)
        pag.click(clicks=clicks)

def inspect(image):
    if pag.locateOnScreen(image, confidence=0.8) is None:
        pag.hotkey('f12')
    else:
        print(pag.locateOnScreen(image, confidence=0.8))

def scrollToLast(image):
    if pag.locateOnScreen(image, confidence=0.9) is not None:
        pag.click(pag.locateCenterOnScreen(image, confidence=0.9))

def get_sender():
    navImage('./cimg/chat_options.png', 0, off_x=-400)
    sleep(0.05)
    pag.click(button='right')
    robo.waitImageToAppear('./cimg/inspect_text.png')
    navImage('./cimg/inspect_text.png',1)
    robo.waitImageToAppear('./cimg/span.png')
    pag.click(pag.locateCenterOnScreen('./cimg/span.png'))
    pag.click(button='right')
    robo.waitImageToAppear('./cimg/copy.png')
    navImage('./cimg/copy.png', 1)
    robo.waitImageToAppear('./cimg/copy_element.png')
    navImage('./cimg/copy_element.png', 1)
    sender = pc.paste()
    sender = repr(remove_text_inside_brackets(sender))
    sender = sender.replace("'", "").replace("+", "")
    pag.hotkey('f12')
    return sender

def get_msg():
    navImage('./cimg/pyperclip.png', 1, off_x=-29, off_y=-59)
    sleep(0.05)
    pag.click(button='right')
    robo.waitImageToAppear('./cimg/inspect_text.png')
    navImage('./cimg/inspect_text.png',1)
    robo.waitImageToAppear('./cimg/div.png')
    pag.click(pag.locateCenterOnScreen('./cimg/div.png'))
    pag.click(button='right')
    robo.waitImageToAppear('./cimg/copy.png')
    navImage('./cimg/copy.png', 1)
    robo.waitImageToAppear('./cimg/copy_element.png')
    navImage('./cimg/copy_element.png', 1)
    message = pc.paste()
    message = repr(remove_text_inside_brackets(message))
    msgTime = now.strftime('%I')[0]
    if msgTime == '0':
        message = message[1:-8]
    else:
        message = message[1:-9]
    pag.hotkey('f12')
    return message

# Removes text inside brackets source code link 
# https://stackoverflow.com/questions/14596884/remove-text-between-and#:~:text=def-,remove_text_inside_brackets,-(text%2C%20brackets
def remove_text_inside_brackets(text, brackets="<>"):
    count = [0] * (len(brackets) // 2) # count open/close brackets
    saved_chars = []
    for character in text:
        for i, b in enumerate(brackets):
            if character == b: # found bracket
                kind, is_close = divmod(i, 2)
                count[kind] += (-1)**is_close # `+1`: open, `-1`: close
                if count[kind] < 0: # unbalanced bracket
                    count[kind] = 0  # keep it
                else:  # found bracket to remove
                    break
        else: # character is not a [balanced] bracket
            if not any(count): # outside brackets
                saved_chars.append(character)
    return ''.join(saved_chars)

def send_message(msg):
    navImage('./cimg/pyperclip.png', 1, off_x=100)
    pag.hotkey('ctrl', 'v')
    pag.typewrite('\n')
    navImage('./cimg/chat_options.png', 1, off_x=20)
    robo.waitImageToAppear('./cimg/close_chat.png')
    navImage('./cimg/close_chat.png', 1)


while True:
    if pag.locateOnScreen('./cimg/greendot.png', confidence=0.8) is not None:
        navImage('./cimg/greendot.png',1,off_x=-100)
        sleep(0.3)
        msg = get_msg()
        created_by = get_sender()
        send_message(process_message(msg, created_by))
    else:
        print('No new messages...!')
    sleep(5)


# cell = otl.find('52114551')
# x = otl.get_all_records()
# print(cell)

# def validate(date_text):
#     try:
#         datetime.datetime.strptime(date_text, '%d/%m/%Y')
#         print(date_text)
#     except:
#         print("Incorrect data format, should be DD/MM/YYYY")

# validate('24/09/2022')

