
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

    explanation = tk.Label(tab, text="ce programme converti tous les fichiers word et excel du répertoire choisi\nen pdf, si leurs noms contient le texte spécifié dans la première case\net ne contiennent pas celui de la deuxième case.",
                           background=Settings.GetConfigValueString("theme", "background_color"), foreground=Settings.GetConfigValueString("theme", "text_color"))
    explanation.grid(row=0, column=0, columnspan=3, sticky="we")

    global pathDisplay, necessaryString, avoidString, convertExcels, replacePdf

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
    
    replacePdf = tk.IntVar()
    # Récupérer la valeur sauvegardée si elle existe
    try:
        saved_replace = Settings.GetConfigValue("fields", "pdf_replace")
        if saved_replace is not None:
            replacePdf.set(saved_replace)
    except:
        pass
    
    replaceButton = TkinterClasses.LabelledCheckbox(tab, replacePdf, "Remplacer les PDF existants")
    replaceButton.grid(column=0, row=5, sticky="news")
    
    # Sauvegarder la valeur quand elle change
    def save_replace_option(*args):
        Settings.SetConfigValue("fields", "pdf_replace", replacePdf.get())
    replacePdf.trace_add("write", save_replace_option)

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
    import time
    from datetime import datetime
    
    path = pathDisplay.GetValue()
    
    # Afficher les informations de début
    Log.Message("="*60)
    Log.Message("=== DÉBUT DE LA CONVERSION PDF ===")
    Log.Message(f"Répertoire source: {path}")
    Log.Message(f"Filtre requis: '{necessaryString.GetValue()}'")
    if avoidString.GetValue():
        Log.Message(f"Filtre à éviter: '{avoidString.GetValue()}'")
    Log.Message(f"Convertir les fichiers Excel: {'Oui' if convertExcels.get() == 1 else 'Non'}")
    Log.Message(f"Remplacer les PDF existants: {'Oui' if replacePdf.get() == 1 else 'Non'}")
    Log.Message(f"Heure de début: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    Log.Message("="*60)

    files = os.listdir(path)
    Log.Message(f"\n[INFO] {len(files)} fichiers trouvés dans le répertoire")

    suitableFiles = []
    for file in files :
        if necessaryString.GetValue() in file and ((not avoidString.GetValue() in file) or avoidString.GetValue() == "") :
            suitableFiles.append(file)
    
    Log.Message(f"[INFO] {len(suitableFiles)} fichiers correspondent aux critères de filtrage")
    
    if len(suitableFiles) == 0:
        Log.Warning("Aucun fichier ne correspond aux critères. Conversion annulée.")
        return
    
    # Statistiques
    word_files = [f for f in suitableFiles if f.endswith(".docx")]
    excel_files = [f for f in suitableFiles if (f.endswith(".xlsx") or f.endswith(".xlsm")) and convertExcels.get() == 1]
    
    Log.Message(f"[INFO] Fichiers Word à convertir: {len(word_files)}")
    Log.Message(f"[INFO] Fichiers Excel à convertir: {len(excel_files)}")
    Log.Message("")
    
    # Initialiser les applications
    Log.Message("[INFO] Ouverture d'Excel...")
    excel_opened = ExcelController.OpenExcel()
    if excel_opened:
        Log.Message("[INFO] Excel ouvert avec succès")
    else:
        Log.Error("[ERREUR] Impossible d'ouvrir Excel")
    
    Log.Message("[INFO] Ouverture de Word...")
    word_opened = WordController.OpenWord()
    if word_opened:
        Log.Message("[INFO] Word ouvert avec succès")
    else:
        Log.Error("[ERREUR] Impossible d'ouvrir Word")
    
    Log.Message("")
    
    # Compteurs de conversion
    word_success = 0
    word_failed = 0
    word_skipped = 0
    excel_success = 0
    excel_failed = 0
    excel_skipped = 0
    total_pdfs_created = 0
    total_pdfs_replaced = 0
    
    start_time = time.time()
    
    for file in suitableFiles :
        try :
            file_path = os.path.join(path, file)
            Log.Message(f"\n{'='*60}")
            Log.Message(f"Traitement de: {file}")
            Log.Message(f"Chemin complet: {file_path}")
            
            if file.endswith(".docx") :
                pdf_name = file.removesuffix(".docx") + ".pdf"
                pdf_path = os.path.join(path, pdf_name)
                
                Log.Message(f"[WORD] Type: Fichier Word (.docx)")
                Log.Message(f"[WORD] PDF cible: {pdf_name}")
                Log.Message(f"[WORD] Chemin PDF: {pdf_path}")
                
                if os.path.isfile(pdf_path) :
                    if replacePdf.get() == 1 :
                        # Remplacer le PDF existant
                        try:
                            file_size = os.path.getsize(pdf_path)
                            os.remove(pdf_path)
                            Log.Message(f"[WORD] PDF existant supprimé (taille: {file_size} octets)")
                            Log.Message(f"[WORD] Remplacement en cours...")
                            total_pdfs_replaced += 1
                        except Exception as e:
                            Log.Error(f"[WORD] ERREUR: Impossible de supprimer le PDF existant: {str(e)}")
                            word_failed += 1
                            continue
                    else :
                        Log.Warning(f"[WORD] Le PDF existe déjà (non remplacé)")
                        word_skipped += 1
                        continue
                else:
                    Log.Message(f"[WORD] Nouveau PDF à créer")
                
                # Vérifier que le fichier source existe
                if not os.path.isfile(file_path):
                    Log.Error(f"[WORD] ERREUR: Le fichier source n'existe pas: {file_path}")
                    word_failed += 1
                    continue
                
                file_size = os.path.getsize(file_path)
                Log.Message(f"[WORD] Taille du fichier source: {file_size} octets")
                Log.Message(f"[WORD] Conversion en cours...")
                
                conversion_start = time.time()
                result = WordController.ConvertToPdfFromPath(file_path)
                conversion_time = time.time() - conversion_start
                
                if result:
                    if os.path.isfile(pdf_path):
                        pdf_size = os.path.getsize(pdf_path)
                        Log.Message(f"[WORD] ✓ Conversion réussie en {conversion_time:.2f} secondes")
                        Log.Message(f"[WORD] PDF créé: {pdf_name} ({pdf_size} octets)")
                        word_success += 1
                        total_pdfs_created += 1
                    else:
                        Log.Error(f"[WORD] ERREUR: Le PDF n'a pas été créé (fichier introuvable)")
                        word_failed += 1
                else:
                    Log.Error(f"[WORD] ERREUR: La conversion a échoué")
                    word_failed += 1

            elif (file.endswith(".xlsx") or file.endswith(".xlsm")) and convertExcels.get() == 1 :
                Log.Message(f"[EXCEL] Type: Fichier Excel ({file[-4:]})")
                file_size = os.path.getsize(file_path)
                Log.Message(f"[EXCEL] Taille du fichier: {file_size} octets")
                Log.Message(f"[EXCEL] Ouverture du classeur...")
                
                workbook_name = ExcelController.OpenWorkbook(file_path)
                if workbook_name is None:
                    Log.Error(f"[EXCEL] ERREUR: Impossible d'ouvrir le classeur")
                    excel_failed += 1
                    continue
                
                Log.Message(f"[EXCEL] Classeur ouvert: {workbook_name}")
                Log.Message(f"[EXCEL] Récupération des feuilles...")
                
                sheets = ExcelController.GetSheetNames(workbook_name)
                if sheets is None:
                    Log.Error(f"[EXCEL] ERREUR: Impossible d'obtenir les feuilles")
                    excel_failed += 1
                    continue
                
                Log.Message(f"[EXCEL] {len(sheets)} feuille(s) trouvée(s): {', '.join(sheets)}")
                
                for sheet in sheets :
                    sheet_pdf_name = sheet + ".pdf"
                    sheet_pdf_path = os.path.join(path, sheet_pdf_name)
                    
                    Log.Message(f"\n  [FEUILLE] {sheet}")
                    Log.Message(f"  [FEUILLE] PDF cible: {sheet_pdf_name}")
                    Log.Message(f"  [FEUILLE] Chemin PDF: {sheet_pdf_path}")
                    
                    if os.path.isfile(sheet_pdf_path) :
                        if replacePdf.get() == 1 :
                            # Remplacer le PDF existant
                            try:
                                old_size = os.path.getsize(sheet_pdf_path)
                                os.remove(sheet_pdf_path)
                                Log.Message(f"  [FEUILLE] PDF existant supprimé (taille: {old_size} octets)")
                                Log.Message(f"  [FEUILLE] Remplacement en cours...")
                                total_pdfs_replaced += 1
                            except Exception as e:
                                Log.Error(f"  [FEUILLE] ERREUR: Impossible de supprimer le PDF existant: {str(e)}")
                                excel_failed += 1
                                continue
                        else :
                            Log.Warning(f"  [FEUILLE] Le PDF existe déjà (non remplacé)")
                            excel_skipped += 1
                            continue
                    else:
                        Log.Message(f"  [FEUILLE] Nouveau PDF à créer")
                    
                    Log.Message(f"  [FEUILLE] Conversion en cours...")
                    conversion_start = time.time()
                    result = ExcelController.ConvertToPdf(workbook_name, sheet, sheet_pdf_path)
                    conversion_time = time.time() - conversion_start
                    
                    if result:
                        if os.path.isfile(sheet_pdf_path):
                            pdf_size = os.path.getsize(sheet_pdf_path)
                            Log.Message(f"  [FEUILLE] ✓ Conversion réussie en {conversion_time:.2f} secondes")
                            Log.Message(f"  [FEUILLE] PDF créé: {sheet_pdf_name} ({pdf_size} octets)")
                            excel_success += 1
                            total_pdfs_created += 1
                        else:
                            Log.Error(f"  [FEUILLE] ERREUR: Le PDF n'a pas été créé")
                            excel_failed += 1
                    else:
                        Log.Error(f"  [FEUILLE] ERREUR: La conversion a échoué")
                        excel_failed += 1
                
                Log.Message(f"[EXCEL] Classeur {workbook_name} terminé")
            
            else :
                Log.Message(f"[INFO] Fichier ignoré (format non supporté ou non demandé): {file}")
                continue

            TkinterClasses.LoadingIcon.UpdateAll()
                    
        except Exception as e :
            Log.Error(f"[ERREUR GLOBALE] Erreur lors du traitement de {file}: {str(e)}")
            import traceback
            Log.Error(traceback.format_exc())
            if file.endswith(".docx"):
                word_failed += 1
            elif (file.endswith(".xlsx") or file.endswith(".xlsm")):
                excel_failed += 1
    
    # Fermer les applications
    Log.Message("\n" + "="*60)
    Log.Message("[INFO] Fermeture des applications...")
    ExcelController.CloseExcel()
    WordController.CloseWord()
    Log.Message("[INFO] Applications fermées")
    
    # Afficher le résumé détaillé
    total_time = time.time() - start_time
    Log.Message("\n" + "="*60)
    Log.Message("=== RÉSUMÉ DE LA CONVERSION PDF ===")
    Log.Message(f"Heure de fin: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    Log.Message(f"Durée totale: {total_time:.2f} secondes")
    Log.Message("")
    Log.Message("--- FICHIERS WORD ---")
    Log.Message(f"  Réussis: {word_success}")
    Log.Message(f"  Échoués: {word_failed}")
    Log.Message(f"  Ignorés (PDF existant): {word_skipped}")
    Log.Message(f"  Total traités: {word_success + word_failed + word_skipped}")
    Log.Message("")
    Log.Message("--- FICHIERS EXCEL ---")
    Log.Message(f"  Feuilles converties avec succès: {excel_success}")
    Log.Message(f"  Feuilles échouées: {excel_failed}")
    Log.Message(f"  Feuilles ignorées (PDF existant): {excel_skipped}")
    Log.Message(f"  Total de feuilles traitées: {excel_success + excel_failed + excel_skipped}")
    Log.Message("")
    Log.Message("--- STATISTIQUES GLOBALES ---")
    Log.Message(f"  PDF créés: {total_pdfs_created}")
    Log.Message(f"  PDF remplacés: {total_pdfs_replaced}")
    Log.Message(f"  PDF ignorés (existants, non remplacés): {word_skipped + excel_skipped}")
    Log.Message(f"  Total de conversions réussies: {word_success + excel_success}")
    Log.Message(f"  Total d'échecs: {word_failed + excel_failed}")
    Log.Message("="*60)
    Log.Message("--Terminé")