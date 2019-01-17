import pywinauto
import ctypes  

def openStreitInstance():
    app = pywinauto.Application().connect(path ="C:\Program Files (x86)\Streit\V1\Programme\SDTStart.exe")
    app_dialog = app.top_window()
    try:
        app_dialog.minimize()
        app_dialog.restore()

        
        # this shows you all the usable elemets and 
        # app_dialog.print_control_identifiers()
        return (app_dialog, app)
        
        # app_dialog.set_focus()
    except(pywinauto.findwindows.WindowNotFoundError):
        print('Fenster nicht gefunden')
        pass
    except(pywinauto.WindowAmbiguousError):
        print('Zu viele Streit Fenster offen!')



def clickChkBox(app, no):
    children = app.children()
    i = 0
    for c in children:
        if c.friendly_class_name() == 'CheckBox':
            if i==no:
                c.draw_outline()
                c.click_input()
            i+=1

def _findChkBox(app):
    children = app.children()


    i = 0
    for c in children:

        if c.friendly_class_name() == 'CheckBox':
            c.draw_outline()
            print(i, c.texts(), c)
        i+=1


def _findButton(app):
    children = app.children()
    i = 0
    for c in children:

        if c.class_name() == 'Button':
            c.draw_outline()
            print(i, c.texts(), c)
        i+=1


def setEditText(app, no, text):
    children = app.children()
    i = 0
    for c in children:
        if c.class_name() == 'Edit' and c.is_enabled() and c.is_visible():
            if i==no:
                c.draw_outline()
                c.type_keys(text, with_spaces = True)
                app.type_keys(r'{VK_TAB}')
                app.wait('ready', timeout=20)
            i+=1

def _findEdit(app):
    children = app.children()
    i = 0
    for c in children:
        if c.class_name() == 'Edit' and c.is_enabled() and c.is_visible():
            c.draw_outline()
            c.SetText(i)
            i+=1




def getDataFromGutschrift(app):
    children = app.Vormaske.children()


    bestellNr = children[3].texts()
    bemerkung = children[4].texts()
    kst = children[18].texts()
    ktr = children[21].texts()
    return (bestellNr[0],bemerkung[0],kst[0],ktr[0])


def newGutschriftHotkey(app):
    app.set_focus()
    app.Edit.type_keys(r'^N')
    app.wait('ready', timeout=20)

def putDataToGutschrift(app, data):
    app.Vormaske.child_window(title="00000000", class_name="Edit").type_keys("00034940", with_spaces = True)
    app.Edit.type_keys(r'{ENTER}')
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
    app.children()[25].click_input() 
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
        if c.class_name() == 'SysTabControl32':
            c.select(tab)
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
    rect = listview.rectangle()
    
    # 1 wheel_dist eq 4 items in the listview 
    # est. 4* 99999 => 400,000 items
    wheeldist = -99999

    pywinauto.mouse.move(coords=(rect.left+60,rect.top+60))
    pywinauto.mouse.scroll(coords=(int((rect.left+rect.right)/2), int((rect.bottom+rect.top)/2)), wheel_dist=wheeldist)
    app.click_input(coords=(rect.right-60,rect.bottom-15))



def focusPositionsPaste(app):
    listview = app.Position.children()[0]
    listview.draw_outline()
    rect = listview.rectangle()
    app.click_input(coords=(rect.left+20,rect.top+35))


def deletePositions(streit, app):
    app.type_keys(r'{DELETE}')
    Dialog = streit.window(best_match=u'Löschung')
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
        
        ### Copy MetaInformationen 
        metadata = getDataFromGutschrift(app)
        rightWindow(app)
        newGutschriftHotkey(app)
        putDataToGutschrift(app,metadata)
        ### done



        switchGutschriftTab(1)
        leftWindow(app)
        switchGutschriftTab(1)

        ### Copy positions
        focusPositionsCopy(app)
        copy(app)
        rightWindow(app)
        focusPositionsPaste(app)
        paste(app)
        ### DONE


        focusPositionsCopy(app)
        focusPositionsCopy(app)

        ### Delete old 
        leftWindow(app)
        focusPositionsCopy(app)
        deletePositions(streit,app)
        ### Done

    except TimeoutError:
        messages.append("Timeout Error")
    except pywinauto.MatchError:
        messages.append("Das Fenster für den nächsten Schritt konnte nicht gefunden werden.")
    else:
        messages.append("Die Positionen wurden erfolgreich verschoben")
    finally:
        app.minimize()
        if(len(messages) > 0): 
            message("; ".join(messages))
        else:
            message("Fehler! Bitte zwei Gutschriften-Tabs nebeneinander öffnen und in dem linken Fenster die Positionen auswählen.")


