# Made by Colin Jackson; January 2020; At Siemens under Reinhard Geisler
# Welcome to my code

from PIL import ImageGrab
import pyautogui as gui
import time
#import sys
#sys.path.insert(1, "C:\\Users\\z0040WBP\\Desktop\\Project")
#import readScreen as esther
import pytesseract as esther
import pyautogui as gui
import numpy as np

#global variable
lastText = ''
#this needs to be here for the image to text recognition to work
esther.pytesseract.tesseract_cmd = r'C:\\Program Files\\Tesseract-OCR\\tesseract.exe'

def LetsGetCrazy():
    print('LET\'S GET CRAZY')
    return LetsGetCrazy()

def readScreen(text, x, y, w, v):
    global lastText
    #I am now going to design a readScreen which takes textbox coordinates as parameters
    printscreen = np.array(ImageGrab.grab(bbox=(x,y,w,v)))
    textOnScreen = esther.image_to_string(printscreen)


    if (textOnScreen != lastText):
        if (text in textOnScreen):
            print('found "{}" in text below'.format(text))
            print('\n')
            print(textOnScreen)
            print('\n')
    lastText = textOnScreen

def readScreen2(text):
    global lastText
    #these coordinates refer to the email specifically. I am now going to design a readScreen which takes textbox coordinates as parameters
    printscreen = np.array(ImageGrab.grab(bbox=(350,353,761,490)))
    textOnScreen = esther.image_to_string(printscreen)


    if (textOnScreen != lastText):
        if (text in textOnScreen):
            print('found "{}" in text below'.format(text))
            print('\n')
            print(textOnScreen)
            print('\n')
            downloadPDF()
    lastText = textOnScreen

class SAP_PIN():
    pin = 'CENSORED'

def goDesktop():
    time.sleep(3)
    #gui.click(x = 1916, y = 1038)
    gui.hotkey('win', 'D')

def openFolder(foldername):
    time.sleep(1)
    gui.hotkey('win', 'r')
    time.sleep(2)
    if (foldername == 'New'):
        gui.write('"Y:\\ICC OM\\Data\\OrderData\\New"')
    gui.press('enter')
    time.sleep(6)

    #maximizes folder window
    gui.hotkey('alt', 'space')
    gui.press(['down','down','down','down', 'enter'])
    time.sleep(1)
    for i in range(34):
        gui.press(['up'])


def enterPKI():
    print('bypassing security...')
    security_ok = search('security_ok', 30)
    if security_ok == None:
        quit()
    print('found PKI window, inputting PIN...')
    gui.click(security_ok)

    PIN_image = search('PIN', 10)
    if PIN_image == None:
        quit()
    gui.click(PIN_image)

    gui.typewrite(SAP_PIN.pin)
    print('successfully entered pin')

    security_ok = search('security_ok', 15)
    if security_ok == None:
        quit()
    gui.click(security_ok, clicks = 2, interval = .3)

def OCR_Search(text, X, Y, distanceFromX, distanceFromY, time_to_wait):
    global lastText
    #I am now going to design a search which uses OCR to find windows or text on screen
    #sometimes the search function just plain fails. It, for no discernable reason, cannot find the image.
    #OCR is much more reliable. It will always detect clear text on screen.

    #base case
    if (time_to_wait == 0):
        print('OCR_Search could not find {}'.format(text))
        return None

    printscreen = np.array(ImageGrab.grab(bbox=(X,Y,distanceFromX,distanceFromY)))
    textOnScreen = esther.image_to_string(printscreen)

    #my success condition is if the parameter text is found within the text on screen
    if (textOnScreen != lastText):
        if (text in textOnScreen):
            print('found "{}" in text below'.format(text))
            print('\n')
            print(textOnScreen)
            print('\n')
            #maybe return the coordinates of the text?
            return

    lastText = textOnScreen

    return OCR_Search(text, X, Y, distanceFromX, distanceFromY, time_to_wait - 1)

#tries to find an object on screen for x time(s) and then gives up
#input must be string of a '___.png in the folder
def search(input_png, time_to_wait):
    input = None

    #base case
    if (time_to_wait == 0):
        #need this below if statement exception just to control how much text this thing spits out when looking through email
        if (input_png != 'yesterday'):
            print('Process failed. Could not find ' + input_png + '.png')
        return None
    #need this below if statement exception just to control how much text this thing spits out when looking through email
    if (input_png != 'yesterday'):
        print('searching for ' + input_png + '.png. Searching for T minus {} tries'.format(time_to_wait))
    #search for the object
    #this function takes a few seconds depending on the image size
    input = gui.locateCenterOnScreen(input_png + '.PNG')

    if (input != None):
        print(input_png + '.png was found on screen at ')
        print(input)
        #should return the coords of the object
        return input

    #recursive call
    return search(input_png, (time_to_wait - 1))

#exists for when one has to load reports for a long time_to_wait in seconds
#it will recursively search for the item but only every 60 seconds/1 minutes
#otherwise it will just wait. Example long_search(item, 900)
#at 900 sec, it searches. for 899-841 it doesn't. At 840 it searches. etc.
#searches twice per iteration because it has trouble finding it after only once
#when it finds the item, long_search will return, so there is no harm in overestimating time input

def long_search(input_png, time_to_wait):
    #base case
    if time_to_wait == 0:
        print('item not found using long search!')
        return None

    if time_to_wait%60 == 0:
        print('{} seconds left.'.format(time_to_wait))

        item = search(input_png, 2)
        if item != None:
            print('item found at {} seconds.'.format(time_to_wait))
            return item
        if item == None:
            print('item not found yet. Ignore Process Failed')
    #normal seconds elapsing
    time.sleep(1)

    return long_search(input_png, time_to_wait - 1)


def openAccess(database):
    time.sleep(1)
    gui.hotkey('win', 'r')
    time.sleep(2)
    if (database == 'DE'):
        #gui.write('"Y:\\ICC OM\\Data\\OrderData\\DE-Orders.accdb"')
        gui.write('"C:\\Users\\z0040WBP\\Desktop\\offlineDB\\OrderData\\DE-Orders.accdb"')

    if (database == 'ICC'):
        gui.write('"Y:\\ICC OM\\Data\\ICC_Orders.accdb"')

    gui.press('enter')
    #access can take a while to load
    time.sleep(15)

def deleteG50():
    print('checking if buyers desk has text')
    g50 = search('G50', 4)
    if g50 == None:
        print('No buyers desk text, move on')
        return
    gui.click(g50)
    for i in range(4):
        gui.press('right')
    gui.hotkey('ctrl', 'shift', 'left')
    gui.press('backspace')
    time.sleep(2)


def closeExcel():
    excel = search('xExcel', 25)
    if excel == None:
        print('could not close Excel')
        return
    gui.click(excel)
    print('closed excel')

def openReport2(reportname):
    if reportname == 'R1':
        for i in range(27):
            gui.press(['down'])
        gui.press('enter')
        time.sleep(3)
    if reportname == 'R2':
        for i in range(28):
            gui.press(['down'])
        gui.press('enter')
        time.sleep(3)
    if reportname == 'R3':
        for i in range(29):
            gui.press(['down'])
        gui.press('enter')
        time.sleep(3)
    if reportname == 'R4':
        for i in range(30):
            gui.press(['down'])
        gui.press('enter')
        time.sleep(3)
    if reportname == 'R5':
        #default 31
        for i in range(23):
            gui.press(['down'])
        gui.press('enter')
        time.sleep(3)
    print('opened ' + reportname)

def copyR_Reports():
    #not working!!
    confirm = search('topofEXCEL' , 40)
    if confirm == None:
        print('could not find Excel when copying report')
        quit()
    time.sleep(3)
    gui.press(['up', 'up', 'up', 'left', 'left'])
    gui.hotkey('ctrl', 'shift', 'down')
    gui.hotkey('ctrl', 'c')
    print('Copied info in report!')
    time.sleep(3)
    closeExcel()
    close = search('savetoClip', 10)
    if close == None:
        print('could not maintain info saved to clipboard')
        quit()
    gui.click(close)
    print('closed Excel')
    time.sleep(3)

def runQuery(name):
    gui.hotkey('ctrl','f')
    #max expected number of characters in transaction title
    for i in range(25):
        gui.press('backspace')
    gui.typewrite(name)
    time.sleep(1)
    gui.press('enter')

    print('Ran query ' + name)

def saveAs(reportname):
    print('Saving' + reportname + 'Now')
    saveAsWindow = search('SaveAsWindow', 45)
    if saveAsWindow == None:
        print('Save window never opened')
        quit()
    gui.typewrite(reportname)
    save = search('SAVE', 3)
    if save == None:
        print('Could not saveAs')
        quit()
    gui.click(save)
    saveYes = search('saveYes', 4)
    if saveYes == None:
        print('Could not saveAs')
        quit()
    gui.click(saveYes)
    print(reportname + ' saved')
    time.sleep(5)

def runReport(transName):
    if transName == 'ZMMH0010':
        gui.hotkey('shift', 'f5')
        time.sleep(1)
        gui.typewrite('/Geisler')
        gui.press(['down','down'])
        gui.hotkey('ctrl', 'shift', 'right')
        gui.press('backspace')
        gui.press('f8')
        time.sleep(1)
        gui.press('f8')
        time.sleep(3)
        print('Running ' + transName + '. Waiting for report to load. 15 minutes?')

        #boom have you ever seen a great function as this?
        list = long_search('LIST', 1320)
        #takes like 15 mins
        if list == None:
            print('PD2 Report didn\'t load')
            quit()
        gui.click(list)
        gui.move(0, 90)
        time.sleep(2)
        gui.move(400, 0)
        time.sleep(1)
        gui.move(0, 30)
        gui.click()
        time.sleep(4)
        confirmSpread = search('confirmSpreadsheet', 10)
        if confirmSpread == None:
            print('Could not save as spreadsheet')
            quit()
        gui.click(confirmSpread)
        time.sleep(10)
        saveAs('PD2-PO-Report-New.xlsx')

    if transName == 'SQ01':
        preq = search('PREQ-DATA', 10)
        if preq == None:
            print('could not find PREQ-Data query.')
            print('assuming it is already selected')
            #print('Cancelling SQ01')
            #eventually this will mean that you go to the next transaction, so it'll reopen PD2
            #return 'F'
            #quit()
        else:
            gui.click(preq)

        time.sleep(2)
        gui.press('f8')
        time.sleep(3)
        #confirm = search('PREQ-CONFIRM', 15)
        #if confirm == None:
        #    print('could not confirm PREQ-Data query selection')
        #    print('Cancelling SQ01')
        #    return 'F'
        #gui.click(confirm)
        deleteG50()
        #enterR5 = search('PREQ_MULTIPLE_SELECTION', 10)
        #if enterR5 == None:
        #    print('Could not enter R5 info')
        #    print('Cancelling SQ01')
        #gui.click(enterR5)
        gui.press('tab')
        gui.press('tab')
        gui.press('enter')
        time.sleep(3)
        #delete
        gui.hotkey('shift', 'f4')
        time.sleep(3)
        #paste
        gui.hotkey('shift','f12')
        time.sleep(3)
        #confirm
        gui.press('f8')
        time.sleep(3)
        #run transaction
        gui.press('f8')
        print('Successfully started running PREQ-DATA')
        reportDone = search('saveReportSQ01', 20)
        if reportDone == None:
           print('Report did not load')
           quit()
        print('report loaded. Save now.')
        gui.moveTo(reportDone)
        gui.move(220, -50)
        gui.click()
        spread = search('Spreadsheet', 10)
        if spread == None:
            print('Could not save as spreadsheet')
            quit()
        gui.click(spread)
        #confirmSpread = search('confirmSpreadsheet', 25)
        #if confirmSpread == None:
        #    print('Could not save as spreadsheet')
        #    quit()
        #gui.click(confirmSpread)
        saveAs('PD2-POReq-Data-New.xlsx')

    if transName == 'orderStatus':
        gui.hotkey('shift', 'f5')
        time.sleep(1)
        z00 = search('z00', 5)
        if z00 == None:
            print('could not change variant')
            quit()
        gui.click(z00)
        for i in range(10):
            gui.press(['right'])
        gui.hotkey('ctrl','shift','left')
        gui.press('backspace')
        gui.press(['up','up'])
        gui.typewrite('ICCOM-ORDER')
        time.sleep(1)
        gui.press('f8')
        multi = search('orderStatusMultiSelect', 10)
        if multi == None:
            print('could not enter R1 info. Could not select multiple selections')
            quit()
        gui.click(multi)
        time.sleep(3)
        gui.hotkey('shift','f4')
        time.sleep(3)
        gui.hotkey('shift','f12')
        time.sleep(3)
        gui.press('f8')
        time.sleep(3)
        gui.press('f8')
        saveExcel = long_search('downloadToExcel', 1500)
        if saveExcel == None:
            print('Status Report did not load')
            quit()
        gui.click(saveExcel)
        time.sleep(5)
        gui.press('enter')
        time.sleep(5)
        gui.press('enter')
        #waits until the saveas window says 'file already exists'
        window = long_search('order_status_window', 300)
        if window != None:
            print('report never finished loading')

        #left off here




def openSAP2():
    gui.click(x = 558, y = 1044)

def openSAP():
    SAP_icon = search('SAP icon', 5)
    if SAP_icon == None:
        SAP_icon = search('orangeSAP', 5)
        if SAP_icon == None:
            quit()
    gui.click(SAP_icon)

def openTransaction2(transName):
    gui.typewrite(transName)
    gui.press('enter')
    time.sleep(2)

def openTransaction(transName):
    trans = search(transName, 5)
    if trans == None:
        print('Failed. Could not open ' + transName)
        trans = search(transName + '2', 5)
        if trans == None:
            print('Failed. Could not open ' + transName)
            quit()
    gui.click(trans, clicks = 2, interval = .3)
    print('Opened ' + transName)

def stopTransaction():
    for i in range(5):
        gui.press('f3')
        time.sleep(2)

def open_KSP(mode):
    if (mode == 1):
        openSAP()
        time.sleep(15)
        gui.hotkey('ctrl', 'f')
        gui.typewrite('KSP')
        print('opening KSP')
        gui.press('enter')
        print('waiting for ksp to load')
        time.sleep(30)
        gui.press('f12')
        print('ksp opened successfully')

def open_PD2(mode):
    if (mode == 1):
        #open SAP
        openSAP()
        time.sleep(10)
        # open PD2
        PD2 = search('PD2', 20)
        if (PD2 == None):
            print('PD2 not found in SAP window, checking for the clicked variant (blue)')
            PD2_clicked = search('PD2_clicked', 20)
            gui.click(PD2_clicked, clicks = 2, interval = .3)
            print('opening PD2...')
            time.sleep(7)
            gui.press('f12')
        else:
            gui.click(PD2, clicks = 2, interval = .3)
            print('opening PD2...')
            time.sleep(15)
            print('waiting for confirm')
            gui.press('f12')

            #check if security is already bypassed
            PD2_window = search('PD2open', 20)
            if (PD2_window == None):

                security_ok = search('security_taskbar', 30)
                if security_ok != None:
                    gui.click(security_ok)
                    enterPKI()
                    time.sleep(10)
                    print('security bypassed')
                    return
                PD2_window = search('PD2open', 15)

            if PD2_window != None:
                print('PD2 security is already bypassed')
                time.sleep(7)
                gui.press('f12')
            elif security_ok != None:
                enterPKI()
            else:
                print('unkown error')

        time.sleep(10)

    if (mode == 2):
        print('maximizing PD2')
        openSAP()
        gui.move(130, -100)
        gui.click()
        print('re-opened PD2')

def checkMultipleLogon():
    multiple = search('mul', 10)
    if multiple == None:
        print('No other PD2 windows open')
        return


def openOutlook():
    icon = search('outlookIcon', 5)
    if icon == None:
        print('Outlook not found.')
        quit()
    gui.click(icon)
    time.sleep(20)
    #maximizes email window
    gui.hotkey('alt', 'space')
    gui.press(['down','down','down','down', 'enter'])

    time.sleep(1)
    print('opened outlook')

def downloadPDF():
    gui.click(x = 522, y = 406)
    time.sleep(1)
    gui.click(x = 1385, y = 500)
    time.sleep(1)
    gui.click(x = 1217, y = 703)
    time.sleep(1)
    gui.press('enter')
    time.sleep(3)
    gui.click(x = 522, y = 406)
    gui.press(['down','down','down','down'])

def searchToday(numberOfTimes, text):
    #searches numberOfTimes through your mailbox
    #hopefully stops when it reaches the yesterday section
    #if it never finds the yesterday, then it stops after numberOfTimes tries

    if numberOfTimes == 0:
        print('failed to recognize "yesterday" in inbox. Searched {} times.'.format(numberOfTimes))
        return

    yesterday = search('yesterday', 2)
    if yesterday != None:
        print('reached end of today section in mailbox')
        return

    gui.press('down')
    readScreen2(text)

    return searchToday(numberOfTimes - 1, text)



def DB1():
     #start PD2 PO Report New
     open_PD2(1)
     openTransaction2('ZMMH0010')
     runReport('ZMMH0010')

     closeExcel()

     goDesktop()
     openAccess('DE')
     runQuery('01a-PD2-PO-Report-Update')

def DB2():
    open_PD2(1)
    stopTransaction()
    goDesktop()
    openFolder('New')
    openReport2('R5')
    copyR_Reports()
    goDesktop()
    open_PD2(2)
    openTransaction2('SQ01')
    runReport('SQ01')
    closeExcel()
    openAccess('DE')
    runQuery('01b-PD2-PReq-Update')
    print('awaiting user input for save files')

    time.sleep(30)

def DB3():
    goDesktop()

    openFolder('New')
    openReport2('R1')
    copyR_Reports()
    goDesktop()
    open_PD2(2)
    stopTransaction()
    openTransaction('orderStatus')
    runReport('orderStatus')
    #unfinished but straightforward
    #openAccess('DE')
    #runQuery('01b-PD2-PReq-Update')


#begin the RPA

time.sleep(5)
#openTransaction('orderStatus')
DB2()
#openOutlook()
#searchToday(30, 'Dominik Linke')
#OCR_Search('hello', 350,353,761,490,50)
