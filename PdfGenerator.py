
import tkinter as tk
from tkinter import StringVar
from tkinter import IntVar
from tkinter import filedialog as fd
import os

import TkinterClasses
import Log
import ExcelController
import WordController
import Settings


def Init(tabBar : TkinterClasses.TabBar) :

    global tab

    tab = TkinterClasses.Tab(tabBar, "Conversion PDF")

    explanation = tk.Label(tab, text="ce programme converti tous les fichiers word et excel du répertoire choisi\nen pdf, si leurs noms contient le texte spécifié dans la première case\net ne contiennent pas celui de la deuxième case.\nLes fichiers prééxistants ne seront pas remplacés",
                           background=Settings.GetConfigValueString("theme", "background_color"), foreground=Settings.GetConfigValueString("theme", "text_color"))
    explanation.grid(row=0, column=0, columnspan=3, sticky="we")

    global pathDisplay, necessaryString, avoidString, convertExcels

    pathDisplay = TkinterClasses.StorageTextField(tab, "fields", "pdf_input_path", "répertoire")
    #pathDisplay = tk.Entry(tab, textvariable=pathText, background=Settings.entryColor, foreground=Settings.textColor)
    pathDisplay.grid(row=1, column=1, columnspan=6, sticky="we")

    choosePathButton = tk.Button(tab, text="chemin répertoire", command=ChoosePath, background=Settings.GetConfigValueString("theme", "unselected_color"), foreground=Settings.GetConfigValueString("theme", "text_color"))
    choosePathButton.grid(row=1, column=0)

    necessaryString = TkinterClasses.StorageTextField(tab, "fields", "pdf_necessary", "(optionnel) fichier doit contenir :")
    necessaryString.grid(row=2, column=0, sticky="we")


    avoidString = TkinterClasses.StorageTextField(tab, "theme", "pdf_avoid", "(optionnel) fichier ne doit pas contenir :")
    avoidString.grid(row=3, column=0, sticky="we")


    convertButton = tk.Button(tab, text="convertir en pdf", command=TryConvertToPdf, background=Settings.GetConfigValueString("theme", "unselected_color"), foreground=Settings.GetConfigValueString("theme", "text_color"))
    convertButton.grid(row=4, column=1, columnspan=6, sticky="we", padx=5)
    
    convertExcels = tk.IntVar()
    excelsButton = TkinterClasses.LabelledCheckbox(tab, convertExcels, "convert excel files ?")
    excelsButton.grid(column=0, row=4, sticky="news")

    messageWindow = TkinterClasses.OutputTextWindow(tab)
    messageWindow.grid(column=0, row=6, columnspan=8, rowspan=4, sticky="nwes", padx=5)



def ChoosePath() :
    val = fd.askdirectory()

    if not os.path.exists(val) :
        Log.Error("chemin invalide")
        return

    pathDisplay.SetValue(val)


def TryConvertToPdf() :
    tab.SetTabLock(True)
    try :
        ConvertToPdf()
    except Exception as e :
        Log.Error(str(e))
    tab.SetTabLock(False)


def ConvertToPdf() :

    files = os.listdir(pathDisplay.GetValue())

    suitableFiles = []
    for file in files :
        if necessaryString.GetValue() in file and ((not avoidString.GetValue() in file) or avoidString.GetValue() == "") :
            suitableFiles.append(file)
    
    ExcelController.OpenExcel()
    WordController.OpenWord()
    for file in suitableFiles :
        try :
            path = pathDisplay.GetValue()
            print(file)
            if file.endswith(".docx") :
                if os.path.isfile(path + "/" + file.removesuffix(".docx") + ".pdf") :
                    Log.Warning(f"Le pdf existe déjà pour {file}")
                    continue
                WordController.ConvertToPdfFromPath(path + "/" + file)

            elif (file.endswith(".xlsx") or file.endswith(".xlsm")) and convertExcels.get() == 1 :
                workbook = ExcelController.OpenWorkbook(path + "/" + file)
                sheets = ExcelController.GetSheetList(workbook)
                for sheet in sheets :
                    ExcelController.ConvertToPdf(workbook, sheet, path + "/" + sheet + ".pdf")
            
            else : continue

            Log.Message(f"Converted {file}")

            TkinterClasses.LoadingIcon.UpdateAll()
                    
        except Exception as e :
            Log.Error(str(e))
    ExcelController.CloseExcel()
    WordController.CloseWord()

    Log.Message("--Terminé")