import os.path
import ctypes  
from os import listdir
from os.path import isfile, join
import time
import shutil

uglImportFolder = r'\\STREITSRV\StreitV1DatenZentral\Mandant_00013128_00_01\tempein\UGL'


def message(text):
    ctypes.windll.user32.MessageBoxW(0, text, "UGL Clutter Cleaning", 0)


# based on http://www.label-software.de/wp-content/uploads/2017/03/ugl_schnittstelle.pdf
def main():
    print(uglImportFolder)
    if os.path.isdir(uglImportFolder):
        allFiles = [f for f in listdir(uglImportFolder) if isfile(join(uglImportFolder, f))]
        lenAllFiles = len(allFiles)
        print("detected files:", lenAllFiles)
        welcomeMsg = "Es wurden %d Dateien im UGL-Ordner gefunden. Im nächsten Schritt werden alle Dateien, die keine Rechnungen oder Gutschriften sind archiviert." % (lenAllFiles)
        message(welcomeMsg)
        
        def isRechnungType(fileName):
            with open(uglImportFolder+'\\'+fileName) as fhandler:
                line = fhandler.readline()
                kop = line[0:3]
                anfrageArt = line[13:15]


                if not(kop == 'RGD' and (anfrageArt == 'RG' or anfrageArt == 'GS')):
                    return True
                else: 
                    print("Rechnung/Gutschrift: ",fileName)
            return False


        lFwithoutRechnung = [f for f in allFiles if isRechnungType(f)]
        lenLfWithoutRechnung = len(lFwithoutRechnung)
        noRechnungenInUglFolder = lenAllFiles - lenLfWithoutRechnung

        print("filtered files:", noRechnungenInUglFolder)


        if lenLfWithoutRechnung > 0:
            date = time.strftime("%Y%m%d")
            archiveName = date + " archive"
            archiveFolder = uglImportFolder + '\\' + archiveName

            if not(os.path.isdir(archiveFolder)):
                try:
                    print('creating archive folder: ', archiveFolder)
                    os.mkdir(archiveFolder)
                except:
                    message("Es konnte kein Ordner erstellt werden für das Archiv")

            for fileName in lFwithoutRechnung:
                src = uglImportFolder+'\\'+fileName
                dest = archiveFolder+'\\'+fileName
                print("moving "+ fileName+ " -> archive")
                shutil.move(src, dest) 


            doneMsg = "Es wurden %d Dateien archiviert. Es sind zurzeit noch %d Rechnungen im UGL-System zum importieren." % (lenLfWithoutRechnung, noRechnungenInUglFolder)
        else:
            doneMsg = "Es mussten keine Dateien archiviert werden. Es sind zurzeit noch %d Rechnungen im UGL-System zum importieren." % (noRechnungenInUglFolder)
        
        folderMsg = "\n\nUGL-Ordner: %s" % (uglImportFolder)
        doneMsgAndFolder = doneMsg + folderMsg

        print(doneMsgAndFolder)
        message(doneMsgAndFolder)



    else:
        message("Der UGL-Ordner konnte nicht aufgerufen werden.")




if __name__ == "__main__":
    try: 
        main()
    except:
        message("Es ist ein unerwarteter Fehler aufgetreten.")