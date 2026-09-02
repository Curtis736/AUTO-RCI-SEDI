import tkinter as tk
import tkinter.scrolledtext
from tkinter import StringVar
from tkinter import filedialog as fd
import os
import re  # Added for the new extract_lt_from_path function

import TkinterClasses
import Settings
import Log


# callback
onTemplatePathUpdate = None
directoryDisplay: tkinter.scrolledtext.ScrolledText | None = None

def Init(tabBar : TkinterClasses.TabBar) :
    global tab

    tab = TkinterClasses.Tab(tabBar, "Paths", OnOpen)
    
    # builds the panel
    global displayMeasurementPath, displayTemplatePath, displayWorkPath, displayPopplerPath

    chooseMeasurementPath = tk.Button(tab, text="choix chemin mesures", command=ChooseMeasurementPath)
    chooseMeasurementPath.grid(column=1, row=3, columnspan=2, rowspan=1)

    displayMeasurementPath = TkinterClasses.StorageLabel(tab, "paths", "measurement_file", wrapCount=100, wrapChar="/")
    displayMeasurementPath.grid(column=3, row=3, columnspan=5, rowspan=1)


    chooseTemplatePath = tk.Button(tab, text="choix chemin formulaire", command=ChooseTemplatePath)
    chooseTemplatePath.grid(column=1, row=4, columnspan=2, rowspan=1)

    displayTemplatePath = TkinterClasses.StorageLabel(tab, "paths", "template_file_path", wrapCount=100, wrapChar="/")
    displayTemplatePath.grid(column=3, row=4, columnspan=5, rowspan=1)


    chooseWorkPath = tk.Button(tab, text="choix chemin lancement", command=ChooseWorkDirectory)
    chooseWorkPath.grid(column=1, row=5, columnspan=2, rowspan=1)

    displayWorkPath = TkinterClasses.StorageLabel(tab, "paths", "root_work_dir", wrapCount=100, wrapChar="/")
    displayWorkPath.grid(column=3, row=5, columnspan=5, rowspan=1)

    choosePopplerPath = tk.Button(tab, text="choix chemin Poppler", command=ChoosePopplerPath)
    choosePopplerPath.grid(column=1, row=6, columnspan=2, rowspan=1)

    displayPopplerPath = TkinterClasses.StorageLabel(tab, "paths", "poppler_path", wrapCount=100, wrapChar="/")
    displayPopplerPath.grid(column=3, row=6, columnspan=5, rowspan=1)

    
    logWindow = TkinterClasses.OutputTextWindow(tab)
    logWindow.grid(row=7, column=1, rowspan=4, columnspan=5)


def OnOpen() :
    TkinterClasses.StorageField.LoadAll()


def FormatString(string) :
    if string == "" :
        return "N/A"
    return string

def UpdateDirectoryDisplay() :
    if directoryDisplay == None :
        return
    directoryDisplay.delete("1.0", tk.END)
    if not os.path.isdir(Settings.GetConfigValueString("paths", "root_work_dir")) :
        directoryDisplay.insert(tk.END, "pas de dossier de lancement sélectionné", "undefined")
        return
    
    directoryDisplay.insert(tk.END, "dossier de lancement :\n", "folder")
    directoryDisplay.insert(tk.END, Settings.GetConfigValueString("paths", "root_work_dir") + "\n\n", "file")
    
    directoryDisplay.insert(tk.END, "fichier excel :\n", "folder")
    directoryDisplay.insert(tk.END, Settings.GetConfigValueString("paths", "measurement_file") + "\n\n", "file")
    
    directoryDisplay.insert(tk.END, "fichier formulaire :\n", "folder")
    directoryDisplay.insert(tk.END, Settings.GetConfigValueString("paths", "template_file_path") + "\n\n", "file")

def ChooseMeasurementPath() :
    Log.Message("[DEBUG] Ouverture de la boîte de dialogue pour choisir le fichier Excel...")
    val = fd.askopenfilename(
        title="Sélectionner un fichier Excel",
        filetypes=[
            ("Fichiers Excel", "*.xlsx *.xlsm"),
            ("Classeur Excel", "*.xlsx"),
            ("Classeur Excel avec macros", "*.xlsm")
        ]
    )

    Log.Message(f"[DEBUG] Fichier sélectionné : {val}")

    if val == "":  # L'utilisateur a annulé la sélection
        Log.Message("[DEBUG] Sélection annulée par l'utilisateur.")
        return
        
    if not os.path.exists(val):
        Log.Error(f" Le fichier sélectionné n'existe pas : {val}")
        Log.Error("le fichier n'existe pas")
        displayMeasurementPath.SetValue("N/A")
        return

    if not val.lower().endswith((".xlsx", ".xlsm")) :
        Log.Error(f" Mauvaise extension pour le fichier sélectionné : {val}")
        Log.Error("le fichier n'a pas l'extension '.xlsx' ou '.xlsm'")
        displayMeasurementPath.SetValue("N/A")
        return
    
    Log.Message(f"[DEBUG] Chemin du fichier Excel validé : {val}")
    displayMeasurementPath.SetValue(val)

def ChooseTemplatePath() :
    val = fd.askopenfilename(
        title="Sélectionner le fichier modèle Word",
        filetypes=[
            ("Fichiers Word", "*.docx"),
        ]
    )

    if val == "":  # L'utilisateur a annulé la sélection
        return

    if not val.endswith(".docx") :
        Log.Error("le fichier n'a pas l'extension '.docx'")
        displayTemplatePath.SetValue("N/A")
        return
    displayTemplatePath.SetValue(val)



def extract_lt_from_path(path):
    """
    Extrait le numéro de lancement (LT) du chemin fourni.
    Retourne le LT trouvé ou None si aucun LT n'est trouvé.
    """
    if not path:
        return None
        
    try:
        # Recherche un pattern LT suivi de chiffres dans le chemin
        # Accepte les variations comme "LT ", "LT-", "LT_" suivies de chiffres
        lt_match = re.search(r'LT[-_ ]?\d+', path, re.IGNORECASE)
        if lt_match:
            lt_number = lt_match.group(0).upper()  # Convertir en majuscules
            # Nettoyer le numéro pour n'avoir que LT + chiffres
            lt_number = re.sub(r'[^LT0-9]', '', lt_number)
            Log.Message(f"LT extrait du chemin : {lt_number}")
            return lt_number
            
        # Si pas de match direct, chercher dans les sous-dossiers
        folders = path.split(os.path.sep)
        for folder in folders:
            lt_match = re.search(r'LT[-_ ]?\d+', folder, re.IGNORECASE)
            if lt_match:
                lt_number = lt_match.group(0).upper()
                lt_number = re.sub(r'[^LT0-9]', '', lt_number)
                Log.Message(f"LT extrait du nom de dossier : {lt_number}")
                return lt_number
                
        Log.Warning(f"Aucun numéro de lancement trouvé dans le chemin : {path}")
        return None
    except Exception as e:
        Log.Error(f"Erreur lors de l'extraction du LT : {str(e)}")
        return None

def ChooseWorkDirectory() :
    val = fd.askdirectory(
        title="Sélectionner le dossier de lancement"
    )
    
    if val == "":  # L'utilisateur a annulé la sélection
        return
        
    if not os.path.isdir(val):
        Log.Error("le chemin n'est pas un dossier")
        displayWorkPath.SetValue("N/A")
        UpdateDirectoryDisplay()
        return
    
    displayWorkPath.SetValue(val)


def ChoosePopplerPath() :
    val = fd.askdirectory(
        title="Sélectionner Poppler (dossier bin ou racine poppler-XX)"
    )

    if val == "":
        return

    import ImageReader
    resolved = ImageReader._resolve_poppler_bin(val)
    if not resolved:
        Log.Error(
            "Dossier Poppler invalide : pdftoppm.exe introuvable. "
            "Choisissez le dossier .../Library/bin ou la racine de l'archive Poppler."
        )
        displayPopplerPath.SetValue("N/A")
        return

    displayPopplerPath.SetValue(resolved)
    ImageReader.GetPopplerPath(force_refresh=True)
    Log.Message(f"[POPPLER] Chemin enregistré : {resolved}")