

import win32com.client
import Log


__wordInstance = None

def OpenWord() :
    global __wordInstance
    if __wordInstance != None :
        Log.Verbose("[WORDCONTROLLER] Word is already open")
        return True
    try :
        __wordInstance = win32com.client.gencache.EnsureDispatch("Word.Application")
        __wordInstance.DisplayAlerts = False
        return True
    except Exception as e :
        Log.Error("[WORDCONTROLLER] : " + str(e))
        return False


def CloseWord() :
    global __wordInstance
    if __wordInstance == None :
        Log.Verbose("[WORDCONTROLLER] Word is not open, cannot close it")
        return True
    try :
        #CloseAllWorkbooks()
        __wordInstance.Quit()
        __wordInstance = None
        return True
    except Exception as e :
        Log.Error("[WORDCONTROLLER] : " + str(e))
        return False

def PrintDocument(path : str) :
    if __wordInstance == None :
        Log.Error("[WORDCONTROLLER] : Word n'est pas ouvert")
        return False
    
    try :
        __wordInstance.PrintOut(path)
        return True
    
    except Exception as e :
        Log.Error("[WORDCONTROLLER] : " + str(e))
        return False

def ConvertToPdfFromPath(path : str) :

    if __wordInstance == None :
        Log.Error("[WORDCONTROLLER] : Word n'est pas ouvert")
        return False

    try :
        path = path.replace("/", "\\")
        document = __wordInstance.Documents.Open(path)
        document.ExportAsFixedFormat(OutputFileName = path.removesuffix(".docx") + ".pdf", ExportFormat=17)
        document.Close(True)
        return True
    except Exception as e :
        Log.Error("[WORDCONTROLLER] : " + str(e))
        return False