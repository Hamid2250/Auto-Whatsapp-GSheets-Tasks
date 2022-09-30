import ctypes
from itertools import islice
import code
from ctypes.wintypes import HWND
from dataclasses import dataclass
from posixpath import splitext
from turtle import title
import pywinauto
from pywinauto import Desktop
from pywinauto.application import Application
from pywinauto. keyboard import send_keys
import pyautogui as pag
import pygetwindow as gw
from time import sleep
from pyrobogui import robo
import pyperclip as pc
import arabic_reshaper
from langdetect import detect_langs
import datetime
import gspread
from openpyxl import load_workbook
import openpyxl
import xlwings as xw
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

# https://gist.github.com/jerblack/2b294916bd46eac13da7d8da48fcf4ab
def setWindowSizePosition():
    user32 = ctypes.windll.user32

    # get screen resolution of primary monitor
    res = (user32.GetSystemMetrics(0), user32.GetSystemMetrics(1))
    # res is (2293, 960) for 3440x1440 display at 150% scaling
    user32.SetProcessDPIAware()
    res = (user32.GetSystemMetrics(0), user32.GetSystemMetrics(1))
    # res is now (3440, 1440) for 3440x1440 display at 150% scaling

    # get handle for Notepad window
    # non-zero value for handle should mean it found a window that matches
    # handle = user32.FindWindowW(u'SAP', None)
    # or
    handle = user32.FindWindowW(None, u'SAP')

    # meaning of 2nd parameter defined here
    # https://msdn.microsoft.com/en-us/library/windows/desktop/ms633548(v=vs.85).aspx
    # minimize window using handle
    user32.ShowWindow(handle, 6)
    # maximize window using handle
    user32.ShowWindow(handle, 9)

    # move window using handle
    # MoveWindow(handle, x, y, width, height, repaint(bool))
    user32.MoveWindow(handle, 0, 0, 900, 1032, True)


##### Waiting SAP to finish processing #####
# x, y = 39, 37
# red, green, blue = 242, 242, 242
def waitPixelColor(x, y, red, green, blue):
    while True:
        pix = pag.pixel(x, y)
        if pix == (red, green, blue):
            break
        else:
            sleep(0.1)

######### WhatsApp Automation ########   
def navImage(image,clicks,off_x=0,off_y=0):
    position=pag.locateCenterOnScreen(image,confidence=0.8)
    if position is None:
        return 0
    else:
        pag.moveTo(position)
        pag.moveRel(off_x,off_y)
        sleep(0.5)
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




def sapLogon():
    ##### SAP Logon 750 > open Khairat Production #####
    app = Application(backend='uia').connect(title_re='.*SAP Logon*.')
    main = app.window(title_re='.*SAP Logon*.')
    main.set_focus()
    # main.print_control_identifiers()
    main.child_window(title="Khairat Production", control_type="ListItem").double_click_input()
    robo.waitImageToAppear('./sapImg/sapLoginWindow.png')
    print("SAP Started")

def sapLoginCreds():
    ##### SAP Login User #####
    app = Application(backend='uia').connect(title='SAP')
    main = app.window(title='SAP')
    main.set_focus()
    # main.print_control_identifiers()
    main.child_window(auto_id="100", control_type="Pane").click_input()
    send_keys(
        "M-Hamed{TAB}987951357{ENTER}"
    )
    robo.waitImageToAppear('./sapImg/sapEasyAccess.png')

def sapHomeCode(code):
    app = Application(backend='uia').connect(title_re='.*SAP Easy Access*.')
    main = app.window(title_re='.*SAP Easy Access*.')
    main.set_focus()
    # main.print_control_identifiers()
    main.child_window(auto_id="1001", control_type="Edit").click_input()
    pag.typewrite("/N {}\n".format(code))
    
def updateQuotationStatus():
    getOrders = otl.get_all_records()
    for i in getOrders:
        getRowCol = otl.find(str(i['Quotation']))
        if i['Need Approval'] == '':

            app = Application(backend='uia').connect(title_re='.*SAP Easy Access*.')
            main = app.window(title_re='.*SAP Easy Access*.')
            main.set_focus()
            # main.print_control_identifiers()
            main.child_window(auto_id="1001", control_type="Edit").click_input()
            # send_keys("/N{SPACE}VA22{ENTER}")
            pag.typewrite("/N VA22\n", interval=0.02)
            robo.waitImageToAppear('./sapImg/sapChangeQuotation.png')

            app = Application(backend='uia').connect(title_re='.*Change Quotation*.')
            main = app.window(title_re='.*Change Quotation*.')
            main.set_focus()
            # main.print_control_identifiers()
            main.child_window(auto_id="100", control_type="Pane").click_input()
            pag.typewrite(str(i['Quotation']), interval=0.02)
            robo.waitImageToAppear('./sapImg/qDoc.png')
            # sleep(0.5)
            main.child_window(title="Status overview", auto_id="149", control_type="Button").click_input()
            try:
                # if main.child_window(title="Information", control_type="Window") != None:
                main.child_window(title="Continue", auto_id="111", control_type="Button").click_input()
            except:
                pass
            robo.waitImageToAppear('./sapImg/sapStatusOverview.png')
            # sleep(0.5)
            navImage('./sapImg/sapStatusOverview.png', 1)
            pag.hotkey('shift', 'f8')
            # sleep(0.5)
            app = Application(backend='uia').connect(title_re='.*Status Overview*.')
            main = app.window(title_re='.*Status Overview*.')
            main.child_window(title="Continue", auto_id="111", control_type="Button").click_input()
            robo.waitImageToAppear('./sapImg/qDir.png')
            navImage('./sapImg/qDir.png', 1)
            # send_keys(
            #     "{TAB}E:\\Hamid\\alkhairat\\full-automation\\Auto-Whatsapp-GSheets-Tasks\{TAB}temp.txt{TAB}4110"
            # )
            pag.typewrite("\tE:\\Hamid\\alkhairat\\full-automation\\Auto-Whatsapp-GSheets-Tasks\ttemp.txt\t4110")
            main.child_window(title="Replace", auto_id="122", control_type="Button").click_input()
            main.child_window(title="Allow", auto_id="1018", control_type="Button").click_input()
            main.child_window(auto_id="1001", control_type="Edit").click_input()
            # send_keys("/N{ENTER}")
            pag.typewrite("/N\n")
            robo.waitImageToAppear('./sapImg/sapEasyAccess.png')

            with open(r'E:\\Hamid\\alkhairat\\full-automation\\Auto-Whatsapp-GSheets-Tasks\\temp.txt', encoding="utf8") as file:
                # read all content from a file using read()
                content = file.read()
                # check if string present or not
                if 'Q000' in content:
                    otl.update_cell(getRowCol.row, getRowCol.col+1, 'NO')
                elif 'Q002' or 'Q004' or 'Q003' in content:
                    otl.update_cell(getRowCol.row, getRowCol.col+1, 'YES')

            with open(r'E:\\Hamid\\alkhairat\\full-automation\\Auto-Whatsapp-GSheets-Tasks\\temp.txt', encoding="utf8") as file:
                # get Customer Number and Name
                lines = file.readlines()
                details = lines[1]
                customerNum = details[21:27]
                customerName = details[42:-1]
                otl.update_cell(getRowCol.row, getRowCol.col-1, customerNum)
                otl.update_cell(getRowCol.row, getRowCol.col-2, customerName)

    



def main():
    # sleep(2)
    # sapLogon()
    # setWindowSizePosition()
    # sapLoginCreds()
    # # sapHomeCode('VA22')
    # updateQuotationStatus()
    
    # wb1 = load_workbook('export.xlsx')
    # ws1 = wb1['Sheet1']
    
    # ws1.delete_rows(2) 
    # ws1.delete_cols(11, 2)
    # ws1.delete_cols(8, 2)
    # ws1.delete_cols(6)
    # ws1.delete_cols(1, 4)

    ##### Add Net price column in export.xlsx #####
    # for row in ws1.iter_rows(min_row=2, min_col=3):
    #     for cell in row:
    #         if cell.value[0] == "1" or "2" or "4" or "9":
    #             ws1.cell(1, 4).value="Net_price"
    #             ws1.cell(row=cell.row, column=4).value=f"=B{cell.row}/A{cell.row}"

    # wb.save('export.xlsx')

    # for row in ws1.iter_rows(min_row=2, min_col=3):
    #     for cell in row:
    #         if cell.internal_value != None:


    wbxl = xw.Book('export.xlsx')
    wbxl.sheets['Sheet1'].range('C2').value
    print(wbxl.sheets['Sheet1'].range('D2').value)
    
    

if __name__ == '__main__':
    main()

    
    



# while True:
#     if pag.locateOnScreen('./cimg/greendot.png', confidence=0.8) is not None:
#         navImage('./cimg/greendot.png',2,off_x=-100)
#         sleep(0.3)
#         msg = get_msg()
#         created_by = get_sender()
#         send_message(process_message(msg, created_by))
#     else:
#         print('No new messages...!')
#     sleep(5)


# def validate(date_text):
#     try:
#         datetime.datetime.strptime(date_text, '%d/%m/%Y')
#         print(date_text)
#     except:
#         print("Incorrect data format, should be DD/MM/YYYY")

# validate('24/09/2022')