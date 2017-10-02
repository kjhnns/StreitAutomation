from pywinauto import *
import ctypes  
import time


def openStreitInstance():
    app = Application().connect(path ="C:\Program Files (x86)\Streit\V1\Programme\SDTStart.exe")
    app_dialog = app.top_window_()
    try:
        app_dialog.Minimize()
        app_dialog.Restore()

        
        # this shows you all the usable elemets and 
        # app_dialog.print_control_identifiers()
        return (app_dialog, app)
        
        # app_dialog.SetFocus()
    except(WindowNotFoundError):
        print('Fenster nicht gefunden')
        pass
    except(WindowAmbiguousError):
        print('Zu viele Streit Fenster offen!')



def clickChkBox(app, no):
    children = app.Children()
    i = 0
    for c in children:
        if c.friendly_class_name() == 'CheckBox':
            if i==no:
                c.draw_outline()
                c.click_input()
            i+=1

def _findChkBox(app):
    children = app.Children()


    i = 0
    for c in children:

        if c.friendly_class_name() == 'CheckBox':
            c.draw_outline()
            print(i, c.Texts(), c)
        i+=1




def clickButton(app, no):
    children = app.Children()
    i = 0
    for c in children:
        if c.Class() == 'Button':
            if i==no:
                c.draw_outline()
                c.click_input()
        i+=1

def _findButton(app):
    children = app.Children()
    i = 0
    for c in children:

        if c.Class() == 'Button':
            c.draw_outline()
            print(i, c.Texts(), c)
            time.sleep(1) 
        i+=1


def setEditText(app, no, text):
    children = app.Children()
    i = 0
    for c in children:
        if c.Class() == 'Edit' and c.IsEnabled() and c.IsVisible():
            if i==no:
                c.draw_outline()
                c.type_keys(text, with_spaces = True)
                app.type_keys(r'{VK_TAB}')
                app.wait('ready', timeout=20)
            i+=1

def _findEdit(app):
    children = app.Children()
    i = 0
    for c in children:
        if c.Class() == 'Edit' and c.IsEnabled() and c.IsVisible():
            c.draw_outline()
            c.SetText(i)
            i+=1



def message(text):
    ctypes.windll.user32.MessageBoxW(0, text, "UGL Import", 0)



def FtpDatenAbholenZuLieferant(streit, UglWindow, LieferantenNr):
    setEditText(UglWindow, 2, LieferantenNr) 
    clickButton(UglWindow, 21)
    

    FtpWindow = streit.Window_(best_match=u'UGL-Schnittstelle FTP-Transfer...UGL-Schnittstelle')
    clickButton(FtpWindow, 20)


    Hinweis = streit.Window_(best_match=u'Hinweis...')

    Hinweis.Wait('ready', timeout=15)
    clickButton(Hinweis, 2)


        
def ImportFiles(streit):
    UglFileImportWindow = streit.Window_(best_match=u'Spezifikation für UGL-Import')
    loop = True
    while loop:
        try:
            UglFileImportWindow.Wait('ready', timeout=10)
            UglFileImportWindow.Ok.click()
        except TimeoutError as e:
            print("done")
            loop =False

if __name__ == "__main__":
    messages = []
    try:
        instances = openStreitInstance()
        app = instances[0]
        streit = instances[1]

        #UglWindow = streit.Window_(best_match=u'UGL-Schnittstelle...')

        #FtpDatenAbholenZuLieferant(streit, UglWindow, "00098295")  # uni elektro
        #FtpDatenAbholenZuLieferant(streit, UglWindow, "00091815")   # elektro braun

        
        #UglWindow.DatenEinlesen.click()

        ImportFiles(streit)



        #_findButton(UglFileImportWindow)

        #_findButton(UglWindow)


        # switchGutschriftTab(0)
        # metadata = getDataFromGutschrift(app)



        # # openNewGutschriftWindows(app)
        # # instead of openNewGutschriftWindows
        # rightWindow(app)
        # newGutschriftHotkey(app)
        # # ------


        # putDataToGutschrift(app,metadata)

        # switchGutschriftTab(1)
        # leftWindow(app)
        # switchGutschriftTab(1)

        # focusPositionsCopy(app)
        # copy(app)
        # rightWindow(app)
        # focusPositionsPaste(app)
        # paste(app)

        
        # focusPositionsCopy(app)
        
        # focusPositionsCopy(app)

        # leftWindow(app)
        # focusPositionsCopy(app)
        # deletePositions(streit,app)

    except TimeoutError:
        messages.append("Timeout Error")
    except MatchError:
        messages.append("Das Fenster für den nächsten Schritt konnte nicht gefunden werden.")
    else:
        messages.append("Blablabla")
    finally:
        #app.Minimize()
        if(len(messages) > 0): 
            print("; ".join(messages))
            message("; ".join(messages))
        else:
            print("; ".join(messages))
            message("Fehler! Bitte zwei Gutschriften-Tabs nebeneinander öffnen und in dem linken Fenster die Positionen auswählen.")


