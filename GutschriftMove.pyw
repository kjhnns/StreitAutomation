from pywinauto import *
import ctypes  

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


def _findButton(app):
    children = app.Children()
    i = 0
    for c in children:

        if c.Class() == 'Button':
            c.draw_outline()
            print(i, c.Texts(), c)
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




def getDataFromGutschrift(app):
    children = app.Vormaske.Children()


    bestellNr = children[3].Texts()
    bemerkung = children[4].Texts()
    kst = children[18].Texts()
    ktr = children[21].Texts()
    return (bestellNr[0],bemerkung[0],kst[0],ktr[0])


def newGutschriftHotkey(app):
    app.SetFocus()
    app.Edit.TypeKeys(r'^N')
    app.wait('ready', timeout=20)

def putDataToGutschrift(app, data):
    app.Vormaske.child_window(title="00000000", class_name="Edit").type_keys("00034940", with_spaces = True)
    app.Edit.TypeKeys(r'{ENTER}')
    app.wait('ready', timeout=20)



    clickChkBox(app,2)
    
    setEditText(app,5, data[0])
    setEditText(app,6, data[1])
    setEditText(app,13, data[2])
    setEditText(app,15, data[3])


def openNewGutschriftWindows(app):
    #keyboard.SendKeys(r'^{VK_TAB}')
    app.type_keys(r'^{VK_TAB}')
    app.wait('ready', timeout=20)
    app.Children()[25].click_input() 
    app.wait('ready', timeout=20)


def leftWindow(app):
    app.type_keys(r'^+{VK_TAB}')
    app.wait('ready', timeout=20)
def rightWindow(app):
    app.type_keys(r'^{VK_TAB}')
    app.wait('ready', timeout=20)



def switchGutschriftTab(tab):
    cs = app.children()
    for c in cs:
        if c.Class() == 'SysTabControl32':
            c.Select(tab)
            break


def copy(app):
    app.type_keys(r'^c')
    app.wait('ready', timeout=20)

def paste(app):
    app.type_keys(r'^v')
    app.wait('ready', timeout=20)

def focusPositionsCopy(app):
    listview = app.Position.children()[0]
    listview.draw_outline()
    rect = listview.Rectangle()
    app.ClickInput(coords=(rect.right-20,rect.bottom-20))


def focusPositionsPaste(app):
    listview = app.Position.children()[0]
    listview.draw_outline()
    rect = listview.Rectangle()
    #app.ClickInput(coords=(rect.left+5,rect.top+10))
    app.ClickInput(coords=(rect.left+20,rect.top+35))


def deletePositions(streit, app):
    app.type_keys(r'{DELETE}')
    Dialog = streit.Window_(best_match=u'Löschung')
    Dialog.child_window(title="&Ja", class_name="Button").click_input()

def message(text):
    ctypes.windll.user32.MessageBoxW(0, text, "Gutschriften", 0)


if __name__ == "__main__":
    messages = []
    try:
        instances = openStreitInstance()
        app = instances[0]
        streit = instances[1]


        switchGutschriftTab(0)
        metadata = getDataFromGutschrift(app)



        # openNewGutschriftWindows(app)
        # instead of openNewGutschriftWindows
        rightWindow(app)
        newGutschriftHotkey(app)
        # ------


        putDataToGutschrift(app,metadata)

        switchGutschriftTab(1)
        leftWindow(app)
        switchGutschriftTab(1)

        focusPositionsCopy(app)
        copy(app)
        rightWindow(app)
        focusPositionsPaste(app)
        paste(app)

        
        focusPositionsCopy(app)
        
        focusPositionsCopy(app)

        leftWindow(app)
        focusPositionsCopy(app)
        deletePositions(streit,app)

    except TimeoutError:
        messages.append("Timeout Error")
    except MatchError:
        messages.append("Das Fenster für den nächsten Schritt konnte nicht gefunden werden.")
    else:
        messages.append("Die Positionen wurden erfolgreich verschoben")
    finally:
        app.Minimize()
        if(len(messages) > 0): 
            message("; ".join(messages))
        else:
            message("Fehler! Bitte zwei Gutschriften-Tabs nebeneinander öffnen und in dem linken Fenster die Positionen auswählen.")


