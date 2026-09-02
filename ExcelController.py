"""
A program that controls MS Excel.
"""

import win32com.client
import os
import PIL.ImageGrab
import pythoncom
import time
import traceback
import uuid
import re

import Log

import ImageReader

__excelInstance = None

__openedWorkbooks = []
__openedWorkbooksNames = []
__openedWorkbooksPaths = []

# useless
busy = False

# variables pour stocker les résultats des lookups
class ExcelGraphLookupResult :
    def __init__(self) :
        self.x = ""
        self.y = ""
        self.graphIndex = ""

lastGraphResults = ExcelGraphLookupResult()

# Tenter d'importer win32timezone, sinon utiliser notre module helper
try:
    import win32timezone
    excel_com_available = True
    Log.Message("Module win32timezone chargé pour Excel")
except ImportError:
    # Essayer d'utiliser notre module helper
    try:
        import win32timezone_helper as win32timezone
        excel_com_available = True
        Log.Message("win32timezone non disponible, utilisation du module d'aide")
    except ImportError:
        # Si win32timezone_helper n'est pas disponible, essayer pythoncom
        try:
            from win32com.client import pythoncom
            excel_com_available = True
            Log.Message("win32timezone et helper non disponibles, utilisation d'une alternative pour Excel")
        except ImportError:
            excel_com_available = False
            Log.Error("win32com non disponible, certaines fonctionnalités Excel ne seront pas accessibles")

def OpenExcel(filePath = None):
    """
    Ouvre une nouvelle instance d'Excel avec une meilleure gestion des erreurs
    et support des chemins réseau
    """
    global __excelInstance, excel_com_available
    
    if __excelInstance is not None:
        Log.Verbose("[EXCELCONTROLLER] Excel is already open")
        return True
    
    # Si win32com n'est pas disponible, signaler l'erreur
    if not excel_com_available:
        Log.Error("Impossible d'ouvrir Excel : win32com n'est pas disponible")
        return False
    
    try:
        # Initialiser COM pour le thread actuel
        pythoncom.CoInitialize()
        
        # Tenter d'abord avec DispatchEx qui est plus stable pour les chemins réseau
        try:
            __excelInstance = win32com.client.DispatchEx("Excel.Application")
        except:
            # Si DispatchEx échoue, essayer Dispatch normal
            try:
                __excelInstance = win32com.client.Dispatch("Excel.Application")
            except:
                Log.Error("Impossible de créer l'instance Excel")
                return False
        
        # Définir les propriétés avec gestion d'erreurs
        # NB : on NE TOUCHE PLUS à la propriété .Visible pour éviter les erreurs
        # de type "Property 'Excel.Application.Visible' can not be set" sur certains postes.
        try:
            __excelInstance.DisplayAlerts = False
        except Exception:
            Log.Warning("Impossible de désactiver les alertes Excel")
            
        # On laisse Excel gérer tout seul sa visibilité pour éviter les crashs COM.
        # Si besoin d'afficher Excel, cela pourra être fait manuellement par l'utilisateur.
            
        Log.Message("Instance Excel créée avec succès")
        
        # Si un chemin de fichier est fourni, tenter de l'ouvrir
        if filePath is not None:
            try:
                # Vérifier si le chemin est accessible
                if not os.path.exists(filePath):
                    Log.Error(f"Le fichier Excel n'existe pas: {filePath}")
                    CloseExcel()
                    return False
                    
                # Ouvrir le fichier avec gestion du timeout
                wb = None
                try:
                    wb = __excelInstance.Workbooks.Open(filePath)
                    Log.Message(f"Fichier Excel ouvert: {filePath}")
                except Exception as e:
                    Log.Error(f"Erreur lors de l'ouverture du fichier: {str(e)}")
                    if wb:
                        try:
                            wb.Close(False)
                        except:
                            pass
                    CloseExcel()
                    return False
            except Exception as e:
                Log.Error(f"Erreur lors de l'accès au fichier: {str(e)}")
                CloseExcel()
                return False
        
        return True
        
    except Exception as e:
        Log.Error(f"Erreur lors de l'initialisation d'Excel: {str(e)}")
        try:
            if __excelInstance:
                __excelInstance.Quit()
        except:
            pass
        try:
            pythoncom.CoUninitialize()
        except:
            pass
        return False

def CloseExcel():
    """
    Ferme proprement l'instance d'Excel
    """
    global __excelInstance, excel_com_available
    
    if __excelInstance is None:
        Log.Verbose("[EXCELCONTROLLER] Excel is not open, cannot close it")
        return True
    
    try:
        # Fermer tous les classeurs
        try:
            for wb in __excelInstance.Workbooks:
                try:
                    wb.Close(SaveChanges=False)
                except:
                    pass
        except:
            pass
            
        # Quitter Excel
        try:
            __excelInstance.Quit()
        except:
            pass
            
    except Exception as e:
        Log.Error(f"Erreur lors de la fermeture d'Excel: {str(e)}")
        
    finally:
        __excelInstance = None
        excel_com_available = True  # Réinitialiser pour permettre de nouvelles tentatives
        
        # Forcer la fermeture des processus Excel
        try:
            os.system('taskkill /F /IM excel.exe /T >nul 2>&1')
        except:
            pass
            
        # Libérer COM
        try:
            pythoncom.CoUninitialize()
        except:
            pass
            
        return True

def CloseAllWorkbooks() :
    for workbook in __openedWorkbooks :
        workbook.Close(True)
    __openedWorkbooks.clear()
    __openedWorkbooksNames.clear()
    __openedWorkbooksPaths.clear()
    
def IsWorkbookOpenedFromName(name : str) :
    return name in __openedWorkbooks

def IsWorkbookOpenedFromPath(path : str) :
    return path in __openedWorkbooksPaths

def NewWorkbook(path : str) :
    """
    Creates a new workbook at the given path, adds it to the opened workbooks and return its name
    """
    path = path.replace("/", "\\")
    Log.Verbose("[EXCELCONTROLLER] creating new workbook at " + path)
    print(path)

    if __excelInstance == None :
        Log.Error("[EXCELCONTROLLER] excel n'est pas ouvert")
        return None

    if not path.endswith(".xlsx") and not path.endswith(".xlsm") :
        Log.Error("[EXCELCONTROLLER] le fichier " + path + " n'est pas un workbook")
        return None
    
    try :
        newWorkbook = __excelInstance.Workbooks.Add()
        newWorkbook.SaveAs(Filename=path)
        __openedWorkbooksNames.append(newWorkbook.Name)
        __openedWorkbooksPaths.append(path)
        __openedWorkbooks.append(newWorkbook)
        return newWorkbook.Name
    except Exception as e :
        Log.Error("[EXCELCONTROLLER] : " + str(e))
        return None

def __OpenWorkbookAndGiveObject(path : str) :
    """
    Opens the workbook at the given path
    """

    path = path.replace("/", "\\")
    file_name = os.path.basename(path)
    lower_path = path.lower()

    if not os.path.isfile(path) :
        Log.Error("[EXCELCONTROLLER] le chemin de workbook " + path + " est invalide")
        return None
    if file_name.startswith("~$"):
        Log.Warning("[EXCELCONTROLLER] fichier Excel temporaire ignoré : " + path)
        return None
    if not lower_path.endswith(".xlsx") and not lower_path.endswith(".xlsm") :
        Log.Error("[EXCELCONTROLLER] le fichier " + path + " n'est pas un workbook")
        return None
    
    Log.Verbose(f"Opening workbook {path}")
    
    try :

        if path in __openedWorkbooksPaths :
            return __openedWorkbooks[__openedWorkbooksPaths.index(path)]
        
        book = __excelInstance.Workbooks.Open(path)
        __openedWorkbooksNames.append(book.Name)
        __openedWorkbooksPaths.append(path)
        __openedWorkbooks.append(book)
        return book
    
    except Exception as e :
        Log.Error("[EXCELCONTROLLER] : " + str(e))
        return None
    
def OpenWorkbook(path : str) :
    """
    Open the workbook at the given absolute path.
    Returns the name of the workbook if successful, else returns None.
    """
    thing = __OpenWorkbookAndGiveObject(path)
    if thing == None :
        return None
    else :
        return thing.Name

def SaveWorkbook(workbookName : str) :
    """
    Tries to save an opened workbook.
    Returns False on failure and True on success.
    """
    if __excelInstance == None :
        return False
    try :
        if not workbookName in __openedWorkbooksNames :
            return False
        workbook = __openedWorkbooks[__openedWorkbooksNames.index(workbookName)]
        workbook.Save()
        return True
    
    except Exception as e :
        Log.Error("[EXCELCONTROLLER] : " + str(e))
        return False


def PrintWholeBook(path : str) :
    """
    Prints all the sheets in the workbook at the given absolute path.
    """
    workbook = OpenWorkbook(path)
    sheets = GetSheets(workbook)
    try :
        for sheet in sheets :
            sheet.PrintOut()
        return True
    except Exception as e :
        Log.Error("[EXCELCONTROLLER] : " + str(e))
        return None
        

def GetSheet(workbookName : str, sheetName : str) :
    """
    If possible, returns the worksheet object with the correct name from the given workbook.
    In case of failure, returns None.
    """
    if __excelInstance == None :
        return None
    try :
        if not workbookName in __openedWorkbooksNames :
            return None
        workbook = __openedWorkbooks[__openedWorkbooksNames.index(workbookName)]
        
        return workbook.Sheets(sheetName)
    
    except Exception as e :
        Log.Error("[EXCELCONTROLLER] : " + str(e))
        return None

def GetSheets(workbookName : str) :
    """
    Returns the Worksheets object associated to the workbook.
    In case of failure, returns None.
    """
    if __excelInstance == None :
        return None
    try :
        if not workbookName in __openedWorkbooksNames :
            Log.Error("[EXCELCONTROLLER] : Le workbook " + workbookName + " n'est pas ouvert")
            return None
        workbook = __openedWorkbooks[__openedWorkbooksNames.index(workbookName)]
        
        return workbook.Sheets
    
    except Exception as e :
        Log.Error("[EXCELCONTROLLER] : " + str(e))
        return None

def GetSheetNames(workbookName : str) :
    """
    If possible, returns a list containing the names of all the sheets in the workbook
    workbookName : the name of an opened workbook, as retuned by OpenWorkbook.
    Returns None on failure, and a list of strings on success.
    """
    if __excelInstance == None :
        return None
    try :
        if not workbookName in __openedWorkbooksNames :
            return None
        workbook = __openedWorkbooks[__openedWorkbooksNames.index(workbookName)]
        return [sheet.Name for sheet in workbook.Sheets]
    except Exception as e :
        Log.Error("[EXCELCONTROLLER] : " + str(e))
        return None

def CopySheet(workbookName : str, worksheetName : str) :
    """
    copy a sheet in the same workbook.
    returns the name of the copied sheet on success, and None on failure.
    """
    return CopySheetAt(workbookName, worksheetName, workbookName)

def CopySheetAt(startBook : str, worksheetName : str, endBook : str) :
    """
    Copies the given sheet from startBook and put the copy in endBook.
    returns the name of the copied sheet on success and None on failure.
    """
    if __excelInstance == None :
        return None
    try :
        if not startBook in __openedWorkbooksNames :
            return None
        if not endBook in __openedWorkbooksNames :
            return None
        workbookStart = __openedWorkbooks[__openedWorkbooksNames.index(startBook)]
        workbookEnd = __openedWorkbooks[__openedWorkbooksNames.index(endBook)]

        previousSheets = GetSheetNames(endBook)

        workbookStart.Sheets(worksheetName).Copy(Before=workbookEnd.Sheets.Item(1))

        newSheets = GetSheetNames(endBook)

        for sheet in newSheets :
            if not sheet in previousSheets :
                return sheet

        return None

    except Exception as e :
        Log.Error("[EXCELCONTROLLER] : " + str(e))
        return None

def FindAndReplaceInSheet(workbookName : str, sheetName : str, find : str, replace : str) :
    """
    Search for the given word all the cells of the sheet and replace it.
    """
    if not workbookName in __openedWorkbooksNames :
        Log.Error("[EXCELCONTROLLER] : no such workbook as " + workbookName)
        return False
    
    workbook = __openedWorkbooks[__openedWorkbooksNames.index(workbookName)]

    # get through all the cells that contains data
    try :
        worksheet = workbook.Sheets(sheetName)
        x = 0
        while x < worksheet.UsedRange.Rows.Count :
            y = 0
            while y < worksheet.UsedRange.Rows.Item(x + 1).Cells.Count :
                if worksheet.UsedRange.Rows.Item(x + 1).Cells.Item(y + 1).Value != None :
                    worksheet.UsedRange.Rows.Item(x + 1).Cells.Item(y + 1).Value = str(worksheet.UsedRange.Rows.Item(x + 1).Cells.Item(y + 1).Value).replace(find, replace)
                #print(worksheet.UsedRange.Rows.Item(x + 1).Cells.Item(y + 1).Value)
                y += 1
            x += 1
        workbook.Save()
        return True
    
    except Exception as e :
        Log.Error("[EXCELCONTROLLER] : " + str(e))
        return False

def GetCellValue(workbookName : str, sheetName : str, x : int, y : int) :
    if not workbookName in __openedWorkbooksNames :
        Log.Error("[EXCELCONTROLLER] : no such workbook as " + workbookName)
        return False
    workbook = __openedWorkbooks[__openedWorkbooksNames.index(workbookName)]

    try :
        Log.Verbose("Getting cell content in " + workbookName + "->" + sheetName + " at " + str(x) + ";" + str(y))
        return workbook.Sheets(sheetName).UsedRange.Columns.Item(x).Cells.Item(y).Value
    except Exception as e :
        Log.Error("n'a pas pu atteindre la cell à : " + str(x) + " " + str(y))
        Log.Error(str(e))
        return None

def SetCellValue(workbookName : str, sheetName : str, x : int, y : int, value) :
    if not workbookName in __openedWorkbooksNames :
        Log.Error("[EXCELCONTROLLER] : no such workbook as " + workbookName)
        return False
    workbook = __openedWorkbooks[__openedWorkbooksNames.index(workbookName)]

    try :
        Log.Verbose("Setting cell content in " + workbookName + "->" + sheetName + " at " + str(x) + ";" + str(y))
        workbook.Sheets(sheetName).UsedRange.Columns.Item(x).Cells.Item(y).Value = value
        return True
    except Exception as e :
        Log.Error("n'a pas pu atteindre la cell à : " + str(x) + " " + str(y))
        Log.Error(str(e))
        return False
    

def FindInSheet(workbookName : str, sheetName : str, find : str) :
    """
    Search for a specific string in the sheet.
    Returns True if the word is found, False otherwise or in case of errors
    """

    if not workbookName in __openedWorkbooksNames :
        Log.Error("[EXCELCONTROLLER] : no such workbook as " + workbookName)
        return False
    
    workbook = __openedWorkbooks[__openedWorkbooksNames.index(workbookName)]

    # get through all the cells that contains data
    try :
        worksheet = workbook.Sheets(sheetName)
        x = 0
        while x < worksheet.UsedRange.Rows.Count :
            y = 0
            while y < worksheet.UsedRange.Rows.Item(x + 1).Cells.Count :
                if find in worksheet.UsedRange.Rows.Item(x + 1).Cells.Item(y + 1).Value :
                    return True
                #print(worksheet.UsedRange.Rows.Item(x + 1).Cells.Item(y + 1).Value)
                y += 1
            x += 1
        workbook.Save()
    
    except Exception as e :
        Log.Error("[EXCELCONTROLLER] : " + str(e))
        return False

def GetAllNeededTags(workbookName : str, sheetName : str, separator = "*") :
    """
    Returns every tag in excel worksheet
    """

    if not workbookName in __openedWorkbooksNames :
        Log.Error("[EXCELCONTROLLER] : no such workbook as " + workbookName)
        return False
    
    workbook = __openedWorkbooks[__openedWorkbooksNames.index(workbookName)]

    # get through all the cells that contains data
    try :
        worksheet = workbook.Sheets(sheetName)
        tags = []
        y = 0
        while y < worksheet.UsedRange.Rows.Count :
            x = 0
            while x < worksheet.UsedRange.Rows.Item(y + 1).Cells.Count :
                if isinstance(worksheet.UsedRange.Rows.Item(y + 1).Cells.Item(x + 1).Value, str) :
                    result = FindTagInString(worksheet.UsedRange.Rows.Item(y + 1).Cells.Item(x + 1).Value, separator)
                    if result != None :
                        tags.append((result, x + 1, y + 1))
                #print(worksheet.UsedRange.Rows.Item(x + 1).Cells.Item(y + 1).Value)
                x += 1
            y += 1
        workbook.Save()
        return tags
    
    except Exception as e :
        Log.Error("[EXCELCONTROLLER] : " + str(e))
        return False
        
def FindTagInString(string : str, separator = "*") :
        finding = 0
        buffer = ""
        i = 0
        while i < len(string) :
            
            if string[i] == separator :
                finding += 1
                if finding == 4 :
                    return buffer
            else :
                if finding == 2 :
                    buffer += string[i]
                else :
                    finding = 0
                    buffer = ""
            i += 1
        
        return None

def _normalize_excel_path(path: str) -> str:
    if not path:
        return ""
    return os.path.normcase(os.path.normpath(path.replace("/", "\\")))


def _make_chart_export_path(prefix: str = "GRAPH") -> str:
    """Chemin PNG court (Excel Export échoue au-delà ~260 caractères)."""
    ImageReader.Init()
    return os.path.join(ImageReader.GetAbspathTempFolder(), f"{prefix}_{uuid.uuid4().hex}.png")


def _export_chart_to_png(chart_obj, file_path: str) -> bool:
    """Exporte une feuille graphique ou un Chart embarqué vers un PNG."""
    try:
        chart_obj.Export(file_path)
        return os.path.isfile(file_path) and os.path.getsize(file_path) > 0
    except Exception as export_err:
        Log.Verbose(f"[EXCELCONTROLLER] Export PNG échoué: {export_err}")
        return False


def _charts_from_worksheet(worksheet):
    if hasattr(worksheet, "ChartArea"):
        return [worksheet]
    charts = []
    try:
        for shape in worksheet.Shapes:
            try:
                if hasattr(shape, "Chart"):
                    charts.append(shape.Chart)
            except Exception:
                pass
    except Exception as e:
        Log.Verbose(f"[EXCELCONTROLLER] impossible d'énumérer les shapes : {e}")
    return charts


def _reference_search_fields(path, serie_name_raw, sheetName, chart_title_raw):
    fields = [serie_name_raw, sheetName, chart_title_raw, os.path.basename(path)]
    for part in re.split(r"[\\/]+", path):
        if part and part not in fields:
            fields.append(part)
    return fields


def ConvertToPdf(workbookName : str, sheetName : str, outfile : str) :
    if not workbookName in __openedWorkbooksNames :
        Log.Error("[EXCELCONTROLLER] : no such workbook as " + workbookName)
        return False
    
    workbook = __openedWorkbooks[__openedWorkbooksNames.index(workbookName)]

    try :
        # Type=0 means pdf file
        outfile = outfile.replace("/", "\\")
        workbook.Sheets(sheetName).ExportAsFixedFormat(Type=0, Filename=outfile)
        return True
    except Exception as e :
        Log.Error(str(e))
        return False
    
def FindGraphInExcelFile(path : str, legendKeywords : list = None, titleKeywords : list = None, sheetKeywords : list = None, pulseCallback = None) :
    """
    Recherche un graphique dans un fichier Excel dont une des séries a pour nom un des legendKeywords (ex : S/N 119, SN119, etc.), même si le nom contient un préfixe ou un suffixe.
    """
    if __excelInstance == None :
        Log.Error("[EXCELCONTROLLER] : excel n'est pas ouvert")
        return None

    path = os.path.normpath(path.replace("/", "\\"))
    path_key = _normalize_excel_path(path)
    print("path to the excel workbook :", path, "and currently opened : ", __openedWorkbooksPaths)
    opened_keys = [_normalize_excel_path(p) for p in __openedWorkbooksPaths]
    if path_key in opened_keys :
        print("reusing workbook")
        workbook = __openedWorkbooks[opened_keys.index(path_key)]
    else :
        try :
            print("reopening workbook")
            workbook = __OpenWorkbookAndGiveObject(path)
        except Exception as e :
            Log.Error("le chemin vers le fichier excel est peut-être invalide")
            Log.Error(str(e))
            return None
    
    # now we must iterate over every sheet in the workbook
    listOfSheetNames = [sheet.Name for sheet in workbook.Sheets]
    print(listOfSheetNames)

    # if there are keywords in sheetKeywords, we trim the sheet list
    validSheetNames = []
    if sheetKeywords != None :
        for name in listOfSheetNames :
            valid = True
            for keyword in sheetKeywords :
                valid = valid and keyword in name
            if valid :
                validSheetNames.append(name)
    else :
        validSheetNames = listOfSheetNames
    
    Log.Verbose("Valid sheet names are " + str(validSheetNames))
    result = ExcelGraphLookupResult()

    def _normalize(value: str) -> str:
        if value is None:
            return ""
        return re.sub(r"[^A-Z0-9]", "", str(value).upper())

    # Préparer les variantes de légende à chercher
    legend_variants = set()
    if legendKeywords:
        for key in legendKeywords:
            key = str(key).replace("S/N", "SN").replace(" ", "")
            legend_variants.add(_normalize(key))
            legend_variants.add(_normalize(key.replace("SN", "S/N ")))
            legend_variants.add(_normalize(key.replace("SN", "SN ")))
            legend_variants.add(_normalize(key.replace("SN", "S/N")))
            legend_variants.add(_normalize(key.replace("SN", "SN")))

    title_variants = set()
    if titleKeywords:
        for key in titleKeywords:
            normalized = _normalize(str(key))
            if normalized:
                title_variants.add(normalized)
    
    exported_paths = []
    
    for sheetName in validSheetNames :
        worksheet = workbook.Sheets(sheetName)
        Log.Verbose(f"checking shapes of worksheet {sheetName}, got {worksheet}")

        # it is possible to have Chart object in the sheet collection, since graphs can be put as worksheet in excel
        # However the COM interface doesn't allow to obtain a static type
        # So the only method yet is to test if the object has an attribute only a Chart object would have
        charts = _charts_from_worksheet(worksheet)
        if not charts:
            continue
        for chart in charts :
            try:
                seriesCollection = chart.SeriesCollection()
                Log.Verbose(f"checking chart legend {[serie.Name for serie in seriesCollection]}")
                chart_matched = False
                for serie in seriesCollection :
                    if chart_matched:
                        break
                    serie_name_raw = str(serie.Name)
                    serie_name = _normalize(serie_name_raw)
                    # Vérifier si une des variantes est CONTENUE dans le nom de la série
                    for variant in legend_variants:
                        if variant and variant in serie_name:
                            if title_variants:
                                chart_title_raw = ""
                                try:
                                    if hasattr(chart, "HasTitle") and chart.HasTitle:
                                        chart_title_raw = str(chart.ChartTitle.Text)
                                except Exception:
                                    chart_title_raw = ""

                                searchable_fields = _reference_search_fields(
                                    path, serie_name_raw, sheetName, chart_title_raw
                                )
                                searchable_normalized = [_normalize(field) for field in searchable_fields if field]
                                has_title_match = any(
                                    title_variant in field
                                    for title_variant in title_variants
                                    for field in searchable_normalized
                                )
                                if not has_title_match:
                                    continue

                            Log.Message(
                                "[EXCELCONTROLLER] Graphique retenu | "
                                f"fichier='{path}' | feuille='{sheetName}' | serie='{serie_name_raw}' | "
                                f"filtre_sn='{variant}' | filtres_ref='{list(title_variants) if title_variants else []}'"
                            )

                            # On a trouvé le graphique correspondant
                            ImageReader.Init()
                            filePath = _make_chart_export_path("GRAPH")
                            try:
                                if not _export_chart_to_png(chart, filePath):
                                    raise RuntimeError("export PNG vide ou échec")
                            except Exception as export_err:
                                Log.Error(
                                    f"[EXCELCONTROLLER] échec export graphique "
                                    f"(feuille='{sheetName}', série='{serie_name_raw}') : {export_err}"
                                )
                                continue
                            #image = PIL.ImageGrab.grabclipboard()
                            #Log.Verbose(f"got from clipboard {image}")
                            
                            #filePath = ImageReader.CacheFromData(image, "GRAPH")
                            result.filePath = filePath
                            global lastGraphResults
                            lastGraphResults = result
                            
                            exported_paths.append(filePath)
                            chart_matched = True
                            break
            except Exception as e:
                Log.Verbose(f"failed to read sheet : {str(e)} {traceback.format_exc()}")
                continue
            if pulseCallback != None :
                pulseCallback()
                
    if len(exported_paths) > 0:
        return exported_paths
    
    Log.Warning(
        "[EXCELCONTROLLER] Aucun graphique exporté dans "
        f"'{path}' | légendes={legendKeywords} | ref={titleKeywords} | feuilles={sheetKeywords}"
    )
    return None

def FindGraphInExcelFileFolder(folderPath : str, titleKeywords : list = None, legendKeywords : list = None, sheetKeywords : list = None, pulseCallback = None) :
    """
    search recursively in the given folder for an excel file containing the correct graph
    """
    if folderPath.endswith("/") :
        folderPath = folderPath[:-1]
    if not os.path.isdir(folderPath) :
        return False
    
    pathsToSearch = [folderPath]
    while len(pathsToSearch) > 0 :
        path = pathsToSearch.pop(0)
        elements = os.listdir(path)
        for element in elements :
            print("element :", element)
            if os.path.isdir(path + "/" + element) :
                pathsToSearch.append(path + "/" + element)
                continue
            if os.path.isfile(path + "/" + element) :
                # when an excel file is opened, excel will often create a temporary file in the same directory.
                # these files have the same name as the original file but start with ~$.
                # trying to open them would cause an error.
                element_lower = element.lower()
                if (element_lower.endswith(".xlsx") or element_lower.endswith(".xlsm")) and not element.startswith("~$") :
                    if FindGraphInExcelFile(path + "/" + element, titleKeywords, legendKeywords, sheetKeywords, pulseCallback) :
                        return True
                continue
    return False


def ExportFirstChartFromSheet(workbookPath : str, sheetName : str) :
    """
    Exporte le premier graphique trouvé sur la feuille demandée d'un classeur Excel.
    Retourne le chemin du PNG exporté, ou None en cas d'erreur.
    """
    try :
        if __excelInstance == None :
            Log.Error("[EXCELCONTROLLER] : excel n'est pas ouvert")
            return None
        workbookPath = workbookPath.replace("/", "\\")
        # Ouvrir ou réutiliser le classeur
        if workbookPath in __openedWorkbooksPaths :
            workbook = __openedWorkbooks[__openedWorkbooksPaths.index(workbookPath)]
        else :
            workbook = __OpenWorkbookAndGiveObject(workbookPath)
            if workbook is None :
                return None
        # Obtenir la feuille
        try :
            worksheet = workbook.Sheets(sheetName)
        except Exception as e :
            Log.Error(f"[EXCELCONTROLLER] : feuille '{sheetName}' introuvable dans {workbookPath}")
            Log.Error(str(e))
            return None
        # Récupérer les graphiques présents (feuille graphique ou shapes)
        charts = _charts_from_worksheet(worksheet)
        if len(charts) == 0 :
            Log.Error(f"[EXCELCONTROLLER] : aucun graphique trouvé dans la feuille '{sheetName}'")
            return None
        # Exporter le premier graphique
        try :
            # Générer un nom de fichier unique pour le graphique dichroïque
            filePath = _make_chart_export_path("DICHRO_GRAPH")
            if not _export_chart_to_png(charts[0], filePath):
                raise RuntimeError("export PNG vide ou échec")
            return filePath
        except Exception as e :
            Log.Error(f"[EXCELCONTROLLER] : échec export graphique : {str(e)}")
            return None
    except Exception as e :
        Log.Error(f"[EXCELCONTROLLER] : erreur ExportFirstChartFromSheet : {str(e)}")
        return None

def PutImageInSheet(workBookName : str, sheetName : str, pathToImage : str, column : int, row : int, imageSize : tuple, width : int = 500) :
    """
    Put the image at the requested column or row.
    Width is in points. The aspect ratio of the image is conserved.
    """
    if not workBookName in __openedWorkbooksNames :
        Log.Error("[EXCELCONTROLLER] : Le workbook " + workBookName + " n'est pas ouvert")
        return False
    workbook = __openedWorkbooks[__openedWorkbooksNames.index(workBookName)]
    
    try :
        return PutImageInSheetObject(workbook.Sheets(sheetName), pathToImage, column, row, imageSize, width)
    except Exception as e :
        Log.Error("[EXCELCONTROLLER] : " + str(e))
        return False

def PutImageInSheetObject(sheetObject, pathToImage : str, column : int, row : int, width : int = 500) :
    """
    Put the image at the requested column or row. Both of those should have their index taken from Worksheet.UsedRange.
    Width is in points. The aspect ratio of the image is conserved.
    """

    xpos = sheetObject.UsedRange.Columns.Item(column).Left
    ypos = sheetObject.UsedRange.Rows.Item(row).Top
    print(xpos, ypos, column, row)
    
    imageSize = ImageReader.GetDimensionsOfImageFromPath(pathToImage)
    resizeRatio = width / imageSize[0]
    height = imageSize[1] * resizeRatio

    sheetObject.Shapes.AddPicture(os.path.abspath(pathToImage), False, True, xpos, ypos, width, height)
    