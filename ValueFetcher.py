import Log
import Settings
import win32com.client
from win32com.client import Dispatch
from datetime import datetime
import os
import pythoncom
import traceback
import re
from math import isnan

# Import numpy de façon robuste
try:
    import numpy as np
    numpy_available = True
except ImportError:
    # Si numpy n'est pas disponible, on crée une implémentation minimale
    class FakeNumpy:
        @staticmethod
        def isnan(x):
            if isinstance(x, (int, float)):
                return False
            if isinstance(x, str) and x.strip() == "":
                return True
            try:
                float(x)
                return False
            except (ValueError, TypeError):
                return True
    np = FakeNumpy()
    print("Warning: numpy non installé, utilisation d'une implémentation minimale")

import Writer

# Vérifier si win32timezone est disponible
try:
    import win32timezone
    excel_com_available = True
    Log.Message("Module win32timezone chargé pour ValueFetcher")
except ImportError:
    # Essayer d'utiliser notre module helper
    try:
        import win32timezone_helper as win32timezone
        excel_com_available = True
        Log.Message("win32timezone non disponible, utilisation du module d'aide pour ValueFetcher")
    except ImportError:
        # Si win32timezone_helper n'est pas disponible, essayer pythoncom
        try:
            from win32com.client import pythoncom
            excel_com_available = True
            Log.Message("win32timezone et helper non disponibles, utilisation d'une alternative pour ValueFetcher")
        except ImportError:
            excel_com_available = False
            Log.Error("win32com non disponible pour ValueFetcher, certaines fonctionnalités ne seront pas accessibles")

# Déclarer la variable globale
valueCollector = None

"""
The list of all text tag values, found in the measurement excel file, associated with a specific SN (a row)
"""
class Container :

    def __init__(self, SN : str) :
        
        # the SN associated with this container
        self.SN = SN

        # the tags and their associated values
        self.tagAndValues = {}

        # the specs (forgot the details)
        self.tagAndSpecs = {}
        
        # every cell encountered, where the column had a valid tag but there was no data in the cell corresponding to our SN
        self.undefinedValues = []

    def TryVerifyAllSpecs(self) :
        try :
            return self.VerifyAllSpecs()
        except Exception as e :
            Log.Error(str(e))
            return False
    
    def VerifyAllSpecs(self) :
        # we consider all specs have the associated tag created, since it is verified in the valuecollector
        for key, value in self.tagAndSpecs.items() :
            if value == None :
                # this is normal and means this value do not have a spec
                continue
            
            # Vérification spéciale pour les mesures en dB
            if key in ["**RL1**", "**RL2**", "**IL_BEFORE**", "**IL_AFTER**",
            "**RL_1310_ELIO_SE**", "**RL_1310_ELIO_E**"]:
                if not self.VerifyDbMeasurement(key):
                    return False
            
            result = value.IsCorrect(self.tagAndValues[key])
            if result == False :
                Log.Error("la valeur de " + key + " pour le fichier SN" + self.SN + " n'est pas bonne")
                return False
        return True
    
    def VerifyDbMeasurement(self, key):
        """
        Vérifie spécifiquement les mesures en dB pour s'assurer qu'elles sont dans des plages raisonnables
        et qu'elles respectent les conventions de signe.
        """
        if key not in self.tagAndValues:
            return True  # Si la clé n'existe pas, on ne peut pas vérifier
            
        value_str = self.tagAndValues[key]
        if not value_str:
            return True  # Si la valeur est vide, on ne peut pas vérifier
            
        try:
            value = float(value_str)
            
            # Vérification des plages de valeurs typiques pour les mesures en dB
            if key in ["**RL1**", "**RL2**", "**RL_1310_ELIO_SE**", "**RL_1310_ELIO_E**"]:
                # Les valeurs RL (Return Loss) sont généralement négatives et comprises entre -10 et -70 dB
                if value > 0:
                    Log.Warning(f"Attention: La mesure {key} pour SN{self.SN} est positive ({value} dB). Les valeurs RL sont généralement négatives.")
                
                if value < -70 or value > -10:
                    Log.Warning(f"Attention: La mesure {key} pour SN{self.SN} ({value} dB) est en dehors de la plage typique pour RL (-70 dB à -10 dB).")
                    
            elif key in ["**IL_BEFORE**", "**IL_AFTER**"]:
                # Les valeurs IL (Insertion Loss) sont généralement négatives et comprises entre 0 et -3 dB
                if value > 0:
                    Log.Warning(f"Attention: La mesure {key} pour SN{self.SN} est positive ({value} dB). Les valeurs IL sont généralement négatives.")
                
                if value < -3 or value > 0:
                    Log.Warning(f"Attention: La mesure {key} pour SN{self.SN} ({value} dB) est en dehors de la plage typique pour IL (-3 dB à 0 dB).")
            
            # On retourne True car ce sont des avertissements, pas des erreurs bloquantes
            return True
            
        except ValueError:
            Log.Error(f"Impossible de convertir la valeur '{value_str}' en nombre pour la mesure {key} (SN{self.SN})")
            return False


class Spec :

    CONDITION_LESS_EQUAL = 0
    CONDITION_MORE_EQUAL = 1

    def __init__(self, condition : int, value : float) :

        self.condition = condition
        self.value = value
    
    def NewFromString(string : str) :
        """
        creates a new Spec object from the given string, that should have a sign and a number, like '<= 0.8'
        returns False in case of failure, True if spec is unneeded, and a Spec object if needed.
        """

        if string.strip() == "N/A" :
            # dans ce cas-là, il est normal de ne pas renvoyer d'object
            return True

        conditionString = ""
        valueString = ""
        condition = -1
        value = 0
        i = 0
        step = 0
        while i < len(string) :
            if string[i] in "=><" :
                conditionString += string[i]
                step = 1
            elif step == 1 and string[i] != " " :
                valueString += string[i]
            i += 1
        
        # now we need to have a valid sign
        if conditionString == "<=" :
            condition = Spec.CONDITION_LESS_EQUAL
        elif conditionString == ">=" :
            condition = Spec.CONDITION_MORE_EQUAL
        else :
            Log.Error("la condition " + conditionString +" n'est pas prise en charge.\n seules les conditions <= et >= sont valides")
            return False
        
        # excel utilise des virgules mais python utilise des points
        valueString = valueString.replace(",", ".")
        try :
            value = float(valueString)
        except :
            Log.Error("impossible de convertir " + valueString + " en float")
            return False
        
        return Spec(condition, value)
    
    def IsCorrect(self, value : str) :
        try :
            value = float(value)
        except :
            Log.Error("une valeur avec une spec devrait être convertible en float : " + value)
            return False
        
        if self.condition == Spec.CONDITION_LESS_EQUAL :
            return value <= self.value
        elif self.condition == Spec.CONDITION_MORE_EQUAL :
            return value >= self.value
        else :
            Log.Error("pas de condition de validité valide pour cette spec")
            return False


class ValueCollector :

    def __init__(self) :
        """Initialise le collecteur de valeurs"""
        # Initialisation des attributs de base
        self.containers = []
        self.tagAndValues = {}
        self.tagAndSpecs = {}
        self.excelInstance = None
        self.workbook = None
        
        # Valeurs par défaut importantes
        self.date = datetime.now().strftime("%d/%m/%Y")
        self.numPlan = "00.000"  # Valeur par défaut pour le numéro de plan
        self.LT = ""  # Pas de valeur par défaut pour le LT
        
        # Extraire le LT du chemin si possible
        try:
            current_path = Settings.GetConfigValueString("paths", "root_work_dir")
            lt_match = re.search(r'LT\d+', current_path)
            if lt_match:
                self.LT = lt_match.group(0)
                Log.Message(f"LT extrait du chemin : {self.LT}")
        except Exception as e:
            Log.Warning(f"Impossible d'extraire le LT du chemin : {str(e)}")
            
        Log.Message("ValueCollector initialisé avec les valeurs par défaut")
        
        self.foundSns = []
        self.isExcelOpen = False
        self.excelApp = None
        self.excel_com_available = excel_com_available

        self.formulaireNum = ""
        self.formulaireIndice = ""
        self.pathToPlan = ""
        self.UpdateDate()
    
    def UpdateDate(self) :
        self.date = datetime.today().strftime("%d/%m/%Y")

    def IsPlanNumberValid(self) :
        '''
        Nous acceptons toujours la valeur par défaut comme valide, 
        car le champ a été supprimé de l'interface
        '''
        return True
    
    def IsLTValid(self) :
        """
        Vérifie si le numéro de lancement est valide.
        Un LT valide doit :
        1. Ne pas être vide
        2. Commencer par 'LT' (insensible à la casse)
        3. Être suivi de chiffres
        """
        if not self.LT or self.LT == "":
            Log.Error("Le numéro de lancement est vide")
            return False
            
        # Nettoyer le LT (enlever les espaces et mettre en majuscules)
        lt_clean = self.LT.strip().upper()
        
        if not lt_clean.startswith("LT"):
            Log.Error("Le numéro de lancement doit commencer par 'LT'")
            return False
            
        # Vérifier que le reste est composé de chiffres
        if not lt_clean[2:].replace(" ", "").isdigit():
            Log.Error("Le numéro de lancement doit être suivi de chiffres")
            return False
            
        return True
    
    def OpenExcel(self) :
        """
        Ouvre l'application Excel avec une meilleure gestion des erreurs.
        Retourne:
            0: erreur
            1: réussite (nouvelle instance)
            2: Excel déjà ouvert
        """
        if self.isExcelOpen and self.excelApp is not None:
            Log.Message("Excel est déjà ouvert")
            return 2
        
        # Si win32com n'est pas disponible, signaler l'erreur
        if not self.excel_com_available:
            Log.Error("Impossible d'ouvrir Excel : win32com n'est pas disponible")
            return 0
        
        try:
            # Création de l'instance Excel avec gestion explicite des erreurs
            try:
                # Utiliser Dispatch au lieu de gencache.EnsureDispatch qui peut poser problème
                self.excelApp = win32com.client.Dispatch("Excel.Application")
                self.excelApp.Visible = 0  # False = 0, non visible
                self.excelApp.DisplayAlerts = 0  # False = 0, pas d'alertes
                self.isExcelOpen = True
                Log.Message("Instance Excel créée avec succès dans ValueFetcher")
                return 1
            except Exception as excel_err:
                Log.Error(f"Erreur lors de la création de l'instance Excel: {str(excel_err)}")
                return 0
        
        except Exception as e:
            Log.Error(f"Erreur générale lors de l'ouverture d'Excel: {str(e)}")
            Log.Error(traceback.format_exc())
            return 0

    def CloseExcel(self) :
        """
        Ferme l'application Excel et libère les ressources.
        """
        if not self.isExcelOpen or self.excelApp is None:
            Log.Verbose("Excel n'est pas ouvert, impossible de le fermer")
            return
        
        try:
            # Fermer Excel proprement
            try:
                self.excelApp.Quit()
                self.excelApp = None
                self.isExcelOpen = False
                Log.Message("Application Excel fermée dans ValueFetcher")
            except Exception as app_quit_err:
                Log.Error(f"Erreur lors de la fermeture d'Excel: {str(app_quit_err)}")
            
        except Exception as e:
            Log.Error(f"Erreur générale lors de la fermeture d'Excel: {str(e)}")
            self.excelApp = None
            self.isExcelOpen = False
    
    def FindTagInString(string : str, separator="*") :
        # Si le string ne contient pas de séparateur, retourner le string tel quel
        if separator not in string:
            return string.strip()
            
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
    
    def GetTagId(tagContent : str) :
        """
        Returns the ID of the tag.
        Note : this only applies to image tag, since they can have arguments.
        """
        return tagContent.split("{")[0]
    
    def GetTagArgs(tagContent : str) :
        """
        Returns the arguments of the tag, as a list.
        Note : this only applies to image tag, since they can have arguments.
        """
        split = tagContent.split("{")
        if len(split) != 2 :
            return []
        args = split[1]
        return args.split(";")
        
    
    def FetchValuesInExcel(self) :
        """
        Ouvre la feuille Excel dont le chemin est donné dans Settings. 
        Lit toutes les colonnes et trouve les tags et valeurs associés à chaque SN dans chaque ligne.
        Ouvre l'application Excel donc attention aux conflits.
        """
        # Les vérifications sont toujours vraies maintenant
        # Nous n'affichons plus les messages concernant les numéros supprimés
        self.UpdateDate()

        # Afficher le chemin du fichier Excel
        Log.Message(f"\n=== LECTURE DU FICHIER EXCEL ===")
        Log.Message(f"Fichier Excel: {Settings.GetConfigValueString('paths', 'measurement_file')}")
        Log.Message(f"[DEBUG] Chemin utilisé pour ouvrir le fichier Excel : {Settings.GetConfigValueString('paths', 'measurement_file')}")
        Log.Message(f"[DEBUG] Fichier existe ? {os.path.exists(Settings.GetConfigValueString('paths', 'measurement_file'))}")
        Log.Message(f"===========================\n")

        # Tentative d'ouverture d'Excel
        result = self.OpenExcel()
        if result == 0 :
            Log.Error("Impossible d'ouvrir Excel, fin de la recherche de données")
            return False
        
        # Ouvrir le classeur correct
        workbook = None
        try :
            Log.Message(f"[DEBUG] Tentative d'ouverture du fichier Excel avec win32com : {Settings.GetConfigValueString('paths', 'measurement_file')}")
            workbook = self.excelApp.Workbooks.Open(Settings.GetConfigValueString("paths", "measurement_file"))
            Log.Message(f"[DEBUG] Ouverture du fichier Excel réussie !")
        except Exception as e :
            Log.Error(f"Impossible d'ouvrir le classeur {Settings.GetConfigValueString('paths', 'measurement_file')}")
            Log.Error(str(e))
            Log.Error(f"[DEBUG] Exception lors de l'ouverture du fichier Excel : {e}")
            self.CloseExcel()  # S'assurer de fermer Excel en cas d'erreur
            return False
        
        try:
            # Trouver les feuilles disponibles sans se fier au numéro de plan
            worksheetNames = [sheet.Name for sheet in workbook.Sheets]
            Log.Message(f"Feuilles trouvées dans le classeur: {', '.join(worksheetNames)}")
    
            # Nous conservons l'ordre exact des feuilles tel que défini dans le fichier
            Log.Message(f"Utilisation de l'ordre original des feuilles Excel")
            
            # Si une seule feuille, l'utiliser directement
            if len(worksheetNames) == 1:
                worksheetName = worksheetNames[0]
                Log.Message(f"Utilisation de la seule feuille disponible: {worksheetName}")
            else:
                # Utiliser la première feuille dans l'ordre original
                worksheetName = worksheetNames[0]
                Log.Message(f"Utilisation de la première feuille dans l'ordre d'origine: {worksheetName}")
            
            # Récupérer l'objet worksheet
            worksheet = workbook.Worksheets(worksheetName)
            
            # Vérifier si la feuille contient des tableaux (ListObjects)
            listOfTables = worksheet.ListObjects
            if listOfTables.Count == 0:
                Log.Warning("Aucun tableau trouvé dans la feuille. Lecture directe des cellules...")
                # Lecture directe des cellules si aucun tableau n'est trouvé
                return self.ReadWorksheetDirectly(worksheet, workbook)
            
            Log.Message(f"Nombre de tableaux trouvés dans la feuille: {listOfTables.Count}")
            table = listOfTables.Item(1)
    
            columns = table.ListColumns
            Log.Message(f"Lecture des {columns.Count} colonnes du tableau...")
    
            # Réinitialiser les dictionnaires
            self.tagAndValues.clear()
            self.tagAndSpecs.clear()
    
            # La hauteur du tableau devrait être la même pour toutes les colonnes. Le -1 est dû aux titres
            tableHeight = len(columns.Item(1).Range.Value) - 1
            Log.Message(f"Nombre de lignes trouvées dans le tableau: {tableHeight}")
    
            # Parcourir toutes les colonnes
            for i in range(1, columns.Count + 1):
                column = columns.Item(i)
                title = column.Range.Cells(1, 1).Value
                
                if title is None:
                    Log.Warning(f"Colonne {i} sans titre, ignorée")
                    continue
                    
                tag = ValueCollector.FindTagInString(title)
                if tag != None :
                    Log.Message(f"Traitement de la colonne {i}: {title} (tag: {tag})")
                    value = column.Range.Value
                    value = value[1:]  # Ignorer la première ligne (titre)
    
                    # Convertir toutes les valeurs en chaînes de caractères
                    fixedValues = []
                    for val in value :
                        if val[0] is None:
                            newVal = ""
                        else:
                            # Forcer la conversion en string et traiter les cas particuliers
                            if isinstance(val[0], (int, float)):
                                newVal = str(val[0])
                                # Supprimer les décimales inutiles (.0)
                                if newVal.endswith(".0"):
                                    newVal = newVal.split(".")[0]
                            else:
                                newVal = str(val[0])
                                
                        # Un SN ne devrait pas avoir de décimale
                        if newVal.endswith(".0") and tag == "SN" :
                            newVal = newVal.removesuffix(".0")
                        
                        fixedValues.append(newVal)
                    
                    self.tagAndValues[tag] = fixedValues
                    Log.Message(f"Valeurs trouvées pour {tag}: {len(fixedValues)} entrées")
                    
                    # Vérification des valeurs vides ou NULL
                    empty_count = sum(1 for v in fixedValues if not v)
                    if empty_count > 0:
                        Log.Warning(f"Le tag {tag} contient {empty_count} valeurs vides sur {len(fixedValues)}")
    
                    # Gestion des spécifications (tags commençant par SPEC_)
                    if tag.startswith("SPEC_") :
                        if len(tag) <= len("SPEC_") :
                            Log.Error("Un tag 'SPEC_' doit avoir un identifiant")
                            continue  # Ne pas retourner False, continuer avec les autres colonnes
                        
                        originalTag = tag.replace("SPEC_", "")
                        Log.Message(f"Traitement des spécifications pour {originalTag}")
    
                        allSpecsForThisTag = []
                        for value in fixedValues :
                            result = Spec.NewFromString(value)
                            if result == False :
                                Log.Warning(f"Spécification invalide: {value} pour {originalTag}")
                                allSpecsForThisTag.append(None)  # On continue avec None au lieu d'échouer
                            elif result == True :
                                allSpecsForThisTag.append(None)
                            else :
                                allSpecsForThisTag.append(result)
                        self.tagAndSpecs[originalTag] = allSpecsForThisTag
                else:
                    Log.Warning(f"Colonne {i} avec titre '{title}' ne contient pas de tag valide, ignorée")
            
            # Afficher un résumé des tags trouvés
            Log.Message(f"\n=== RÉSUMÉ DES TAGS TROUVÉS ===")
            Log.Message(f"Nombre total de tags trouvés: {len(self.tagAndValues)}")
            for tag in sorted(self.tagAndValues.keys()):
                Log.Message(f"- {tag}: {len(self.tagAndValues[tag])} valeurs")
            Log.Message(f"=============================\n")
            
            # Ajouter les tags constants pour la cohérence
            self.tagAndValues["**NUMPLAN**"] = [self.numPlan] * tableHeight
            self.tagAndValues["**TITREPLAN**"] = [Settings.GetConfigValueString("fields", "TITREPLAN")] * tableHeight
            self.tagAndValues["**DATE**"] = [self.date] * tableHeight
            
            # Vérifier que chaque spec a son tag normal correspondant
            for normalTag in self.tagAndSpecs.keys() :
                if not normalTag in self.tagAndValues.keys() :
                    Log.Warning(f"La spec pour les valeurs marquées '{normalTag}' existe mais ce marquage n'existe pas dans les données")
                    # Créer un tag vide plutôt que d'échouer
                    self.tagAndValues[normalTag] = [""] * tableHeight
            
            # Vérifier les tags obligatoires
            if not "SN" in self.tagAndValues.keys() :
                Log.Error("Le fichier Excel ne contient pas de colonne avec un tag 'SN'")
                workbook.Close(False)
                self.CloseExcel()
                return False
            
            # Utiliser toutes les lignes qui ont un SN
            validRows = []
            i = 0
            for value in self.tagAndValues["SN"] :
                if value and value.strip():  # Si la valeur SN n'est pas vide
                    validRows.append(i)
                i += 1
            
            if len(validRows) == 0 :
                Log.Error(f"Aucun SN trouvé dans {Settings.GetConfigValueString('paths', 'measurement_file')}")
                workbook.Close(False)
                self.CloseExcel()
                return False
            
            Log.Message(f"Nombre de lignes valides trouvées: {len(validRows)}")
    
            # Créer les containers pour chaque SN
            self.containers.clear()
            self.foundSns.clear()
            
            # Utiliser les lignes valides dans leur ordre original
            Log.Message("Création des containers dans l'ordre original du fichier Excel")
            
            for row in validRows :
                SN = self.tagAndValues["SN"][row]
                self.foundSns.append(SN)
                container = Container(SN)
                
                # Copier toutes les valeurs dans le container
                for key in self.tagAndValues.keys() :
                    
                    value = self.tagAndValues[key][row] if row < len(self.tagAndValues[key]) else ""
                    if not value or value == "" or value is None:
                        container.undefinedValues.append(key)
                        Log.Message(f"Valeur manquante pour {key} dans SN{SN}")
                        # Fournir une valeur par défaut pour les tags critiques
                        if key == "**LT**":
                            value = self.LT
                        elif key == "**NUMPLAN**":
                            value = self.numPlan
                        elif key == "**DATE**":
                            value = self.date
                    
                    else :
                        Log.Verbose(f"added {key}:{value} to container {SN}")
                        container.tagAndValues[key] = value
                
                # Copier les spécifications
                for key, value in self.tagAndSpecs.items() :
                    if row < len(value):
                        container.tagAndSpecs[key] = value[row]
                    else:
                        container.tagAndSpecs[key] = None
                
                self.containers.append(container)
                Log.Message(f"Container créé pour SN{SN} avec {len(container.tagAndValues)} valeurs")
    
            workbook.Close(True)
            self.CloseExcel()
    
            # IMPORTANT: Ne pas trier les containers pour conserver l'ordre original
            Log.Message("Conservation de l'ordre original des données du fichier Excel")
            
            return True
            
        except Exception as e:
            Log.Error(f"Erreur lors de la lecture du fichier Excel: {str(e)}")
            import traceback
            Log.Error(traceback.format_exc())
            # Assurez-vous de fermer Excel en cas d'erreur
            try:
                if workbook:
                    workbook.Close(False)
                self.CloseExcel()
            except:
                pass
            return False
    
    def ReadWorksheetDirectly(self, worksheet, workbook):
        """
        Lit directement les cellules de la feuille Excel sans utiliser de tableau.
        Cette méthode est utilisée quand aucun tableau n'est trouvé dans la feuille.
        """
        try:
            Log.Message("\n=== DÉBUT DE LA LECTURE DIRECTE DU FICHIER EXCEL ===")
            
            # Réinitialiser les dictionnaires
            self.tagAndValues.clear()
            self.tagAndSpecs.clear()
            
            # Trouver la dernière ligne et colonne utilisées
            last_row = worksheet.UsedRange.Rows.Count
            last_col = worksheet.UsedRange.Columns.Count
            Log.Message(f"Dimensions de la feuille: {last_row} lignes x {last_col} colonnes")
            
            # Dictionnaire de correspondance pour les titres de colonnes spéciaux
            special_column_titles = {
                # IL measurements
                "IL 850 nm (≤ 0,5 dB) avant cyclage thermique": "IL_850_A",
                "IL 850 nm (≤ 0,5 dB) après cyclage thermique": "IL_850_B",
                "IL 940": "IL_940",
                "**IL_940**": "IL_940",
                "VOIE_940": "IL_940",
                "IL 1310 BC": "IL_1310_BC",
                "IL 1310 CB": "IL_1310_CB",
                
                # RL measurements
                "RL 1310nm (≥ 45 dB) avant cyclage thermique 1er connecteur": "RL_1310_1_A",
                "RL 1310nm (≥ 45 dB) avant cyclage thermique 2ème connecteur": "RL_1310_2_A",
                "RL 1310nm (≥ 45 dB) apres cyclage thermique 1er connecteur": "RL_1310_1_B",
                "RL 1310nm (≥ 45 dB) apres cyclage thermique 2ème connecteur": "RL_1310_2_B",
                "RL COEUR V940": "RL_COEUR_V940",
                "RL V940": "RL_V940",
                "RL Vligne": "RL_Vligne",
                "RL V1310": "RL_V1310",
                
                # Specifications
                "SPEC V940": "Spec_V940",
                "SPEC Vligne": "Spec_Vligne",
                "SPEC V1310": "Spec_V1310",
                
                # Measurements
                "mesure V940": "mesure_V940",
                "mesure Vligne": "mesure_Vligne",
                "mesure V1310": "mesure_V1310",
                "Mesure Junction 1": "len_1",
                "Mesure Junction 2": "len_2",
                "Longueur Totale": "len_tot",
                
                # Client and order info
                "N° commande Client": "CDE_CLIENT",
                "Référence Client": "REF_CLIENT",
                "N° commande SEDI – ATI": "CDE_SEDI",
                "Commande SEDI": "CDE_SEDI",
                "BL": "BL",
                
                # Other measurements
                "Dérogation": "derog",
                "DEROGATION": "derog",
                "CIT": "fin_cit",
                "FIN CIT": "fin_cit",
                "FIN_CIT": "fin_cit",
                
                # Personnel
                "Controleur": "Controleur",
                "Verificateur": "Verificateur",
                "Commentaire": "Commentaire",
                
                # Screening info
                "Screening puissance": "screen_pow",
                "Screening impulsion": "screen_impulse",
                "Screening nombre tirs": "screen_nb_tirs"
            }
            
            # Ajouter aussi les colonnes avec des tags standards
            standard_column_titles = {
                "N° de S/N": "SN",
                "Numéro de série": "SN",
                "S/N": "SN",
                "Référence produit": "REF_SEDI",
                "Référence SEDI": "REF_SEDI",
                "Désignation": "Des",
                "Description": "Des",
                "Lancement": "LT",
                "N° Lancement": "LT",
                "AGS": "Cde_AGS",
                "Commande AGS": "Cde_AGS"
            }
            
            # Lire les titres des colonnes (première ligne)
            Log.Message("\n=== ANALYSE DES TITRES DE COLONNES ===")
            column_tags = {}  # Pour stocker les associations colonne -> tag
            for col in range(1, last_col + 1):
                title = worksheet.Cells(1, col).Value
                if title:
                    # Vérifier d'abord si c'est un titre spécial
                    title_str = str(title).strip()
                    if title_str in special_column_titles:
                        tag = special_column_titles[title_str]
                        column_tags[col] = tag
                        Log.Message(f"Colonne {col}: Titre spécial '{title_str}' -> Tag '{tag}'")
                    elif title_str in standard_column_titles:
                        tag = standard_column_titles[title_str]
                        column_tags[col] = tag
                        Log.Message(f"Colonne {col}: Titre '{title_str}' -> Tag '{tag}'")
                    else:
                        # Sinon, chercher un tag normal
                        tag = ValueCollector.FindTagInString(str(title))
                        if tag:
                            column_tags[col] = tag
                            Log.Message(f"Colonne {col}: Titre '{title}' -> Tag '{tag}'")
                        else:
                            Log.Message(f"Colonne {col}: Titre '{title}' -> Pas de tag trouvé")
            
            Log.Message(f"\nNombre de colonnes avec tags valides: {len(column_tags)}")
            
            # Lire les valeurs pour chaque colonne identifiée
            Log.Message("\n=== LECTURE DES VALEURS ===")
            for col, tag in column_tags.items():
                values = []
                empty_count = 0
                for row in range(2, last_row + 1):  # Commencer à la ligne 2 (après les titres)
                            cell_value = worksheet.Cells(row, col).Value
                            
                    # Convertir la valeur en string et la nettoyer
                            if cell_value is None:
                                str_value = ""
                            empty_count += 1
                else:
                        # Forcer la conversion en string et traiter les cas particuliers
                        if isinstance(cell_value, (int, float)):
                            str_value = str(cell_value)
                            # Supprimer les décimales inutiles (.0)
                            if str_value.endswith(".0"):
                                str_value = str_value.split(".")[0]
                            else:
                                str_value = str(cell_value)
                                        
                                    # Un SN ne devrait pas avoir de décimale
                        if str_value.endswith(".0") and tag == "SN":
                            str_value = str_value.removesuffix(".0")
                    
                        values.append(str_value)
                        Log.Verbose(f"Colonne {col} (tag {tag}), ligne {row}: '{str_value}'")
                
                # Stocker les valeurs
                self.tagAndValues[tag] = values
                Log.Message(f"Tag {tag}: {len(values)} valeurs lues, dont {empty_count} vides")
                
                # Gestion des spécifications (tags commençant par SPEC_)
                if tag.startswith("SPEC_"):
                    if len(tag) <= len("SPEC_"):
                        Log.Error("Un tag 'SPEC_' doit avoir un identifiant")
                        continue
                    
                    originalTag = tag.replace("SPEC_", "")
                    Log.Message(f"\nTraitement des spécifications pour {originalTag}")
                    
                    allSpecsForThisTag = []
                    for value in values:
                        result = Spec.NewFromString(value)
                        if result == False:
                            Log.Warning(f"Spécification invalide: {value}")
                            allSpecsForThisTag.append(None)
                        elif result == True:
                            allSpecsForThisTag.append(None)
                        else:
                            allSpecsForThisTag.append(result)
                            Log.Message(f"Spécification valide: {value}")
                    
                            self.tagAndSpecs[originalTag] = allSpecsForThisTag
            
            # Ajouter les tags constants
            Log.Message("\n=== AJOUT DES TAGS CONSTANTS ===")
            num_rows = last_row - 1  # Soustraire 1 pour la ligne de titre
            self.tagAndValues["**NUMPLAN**"] = [self.numPlan] * num_rows
            self.tagAndValues["**TITREPLAN**"] = [Settings.GetConfigValueString("fields", "TITREPLAN")] * num_rows
            self.tagAndValues["**DATE**"] = [self.date] * num_rows
            Log.Message(f"Tags constants ajoutés pour {num_rows} lignes:")
            Log.Message(f"- **NUMPLAN**: {self.numPlan}")
            Log.Message(f"- **TITREPLAN**: {Settings.GetConfigValueString('fields', 'TITREPLAN')}")
            Log.Message(f"- **DATE**: {self.date}")
            
            # Vérifier que chaque spec a son tag normal correspondant
            Log.Message("\n=== VÉRIFICATION DES SPÉCIFICATIONS ===")
            for normalTag in self.tagAndSpecs.keys():
                if normalTag not in self.tagAndValues.keys():
                    Log.Warning(f"La spec pour les valeurs marquées '{normalTag}' existe mais ce marquage n'existe pas dans les données")
                    # Créer un tag vide plutôt que d'échouer
                    self.tagAndValues[normalTag] = [""] * num_rows
                    Log.Message(f"Tag vide créé pour {normalTag}")
            
            # Afficher un résumé des tags trouvés
            Log.Message("\n=== RÉSUMÉ DES TAGS TROUVÉS ===")
            Log.Message(f"Nombre total de tags trouvés: {len(self.tagAndValues)}")
            for tag in sorted(self.tagAndValues.keys()):
                empty_count = sum(1 for v in self.tagAndValues[tag] if not v)
                if empty_count > 0:
                    Log.Message(f"- {tag}: {len(self.tagAndValues[tag])} valeurs, dont {empty_count} vides")
                else:
                    Log.Message(f"- {tag}: {len(self.tagAndValues[tag])} valeurs, toutes non vides")
            
            Log.Message("\n=== FIN DE LA LECTURE DU FICHIER EXCEL ===")
            Log.Message(f"Lecture directe du fichier Excel terminée avec succès. {len(self.tagAndValues)} tags traités.")
            return True
            
        except Exception as e:
            Log.Error(f"Erreur lors de la lecture directe du fichier Excel: {str(e)}")
            import traceback
            Log.Error(traceback.format_exc())
            
            # S'assurer de fermer Excel proprement en cas d'erreur
            try:
                if workbook:
                    workbook.Close(False)
                self.CloseExcel()
            except:
                pass
            
            return False
    
    def SetWriterTagLists(self, container):
        """Configure les remplacements de tags pour le document Word"""
        # Vérifier que le container est valide
        if not container:
            Log.Error("Container invalide pour la configuration des tags")
            return False
        
        Writer.columnContentMap.clear()
        Writer.replacementMap.clear()

        self.AddTags(container, Writer.SetReplacementMapping)
        self.AddTags(container, Writer.SetColumnFillMapping)
        
        # also the date
        Writer.SetReplacementMapping("DATE", self.date)
        Writer.SetColumnFillMapping("DATE", self.date)
        
        return True
    
    def GenerateDbMeasurementsSummary(self, container):
        """
        Génère un résumé des mesures en dB pour le SN spécifié.
        Ce résumé sera inclus dans le document généré.
        """
        # Rechercher tous les tags qui pourraient contenir des mesures en dB
        db_tags = []
        
        # Tags connus pour les mesures en dB
        known_db_tags = ["**RL1**", "**RL2**", "**IL_BEFORE**", "**IL_AFTER**",
        "**RL_1310_ELIO_SE**", "**RL_1310_ELIO_E**"]
        
        # Ajouter les tags connus
        for tag in known_db_tags:
            if tag in container.tagAndValues:
                db_tags.append(tag)
        
        # Rechercher d'autres tags qui pourraient contenir des mesures en dB
        for tag, value in container.tagAndValues.items():
            # Si le tag contient "dB", "RL", "IL" ou "LOSS" et n'est pas déjà dans la liste
            if tag not in db_tags and (
                "DB" in tag.upper() or 
                "RL" in tag.upper() or 
                "IL" in tag.upper() or 
                "LOSS" in tag.upper()
            ):
                db_tags.append(tag)
        
        # Trier les tags pour une présentation cohérente
        db_tags.sort()
        
        Log.Message(f"Génération du résumé des mesures en dB pour SN{container.SN}")
        Log.Message(f"Tags de mesures en dB trouvés: {len(db_tags)}")
        
        summary = f"=== RÉSUMÉ DES MESURES EN dB POUR SN{container.SN} ===\n\n"
        has_measurements = False
        
        # Regrouper les mesures par type
        rl_measurements = []
        il_measurements = []
        other_measurements = []
        
        for tag in db_tags:
            if tag in container.tagAndValues and container.tagAndValues[tag]:
                has_measurements = True
                value = container.tagAndValues[tag]
                
                # Déterminer le type de mesure
                measure_type = ""
                if "RL" in tag.upper():
                    measure_type = "Return Loss (RL)"
                    rl_measurements.append((tag, value, measure_type))
                elif "IL" in tag.upper():
                    measure_type = "Insertion Loss (IL)"
                    il_measurements.append((tag, value, measure_type))
                else:
                    measure_type = "Autre mesure en dB"
                    other_measurements.append((tag, value, measure_type))
        
        # Ajouter les mesures de Return Loss
        if rl_measurements:
            summary += "MESURES DE RETURN LOSS (RL):\n"
            summary += "--------------------------\n"
            for tag, value, measure_type in rl_measurements:
                summary += f"• {tag}: {value} dB\n"
                
                # Vérifier si la valeur est dans la plage attendue
                try:
                    value_float = float(value)
                    if value_float > 0:
                        summary += "  ⚠️ Attention: Valeur positive (devrait être négative)\n"
                    if value_float < -70 or value_float > -10:
                        summary += f"  ⚠️ Attention: Valeur hors plage typique (-70 dB à -10 dB)\n"
                except ValueError:
                    summary += f"  ⚠️ Attention: Valeur non numérique\n"
            summary += "\n"
        
        # Ajouter les mesures d'Insertion Loss
        if il_measurements:
            summary += "MESURES D'INSERTION LOSS (IL):\n"
            summary += "----------------------------\n"
            for tag, value, measure_type in il_measurements:
                summary += f"• {tag}: {value} dB\n"
                
                # Vérifier si la valeur est dans la plage attendue
                try:
                    value_float = float(value)
                    if value_float > 0:
                        summary += "  ⚠️ Attention: Valeur positive (devrait être négative)\n"
                    if value_float < -3 or value_float > 0:
                        summary += f"  ⚠️ Attention: Valeur hors plage typique (-3 dB à 0 dB)\n"
                except ValueError:
                    summary += f"  ⚠️ Attention: Valeur non numérique\n"
            summary += "\n"
        
        # Ajouter les autres mesures en dB
        if other_measurements:
            summary += "AUTRES MESURES EN dB:\n"
            summary += "-------------------\n"
            for tag, value, measure_type in other_measurements:
                summary += f"• {tag}: {value} dB\n"
            summary += "\n"
        
        if not has_measurements:
            summary += "Aucune mesure en dB trouvée pour ce SN.\n"
        
        summary += "===========================================\n"
            
        Log.Message(f"Résumé des mesures en dB généré pour SN{container.SN}: {len(rl_measurements) + len(il_measurements) + len(other_measurements)} mesures trouvées")
        return summary
    
    def AddTags(self, container : Container, func) :
        
        for key, value in container.tagAndValues.items() :
            func(key, value)
    
    def GetNewFilename(self, container) :
        return self.formulaireNum + " Ind " + self.formulaireIndice + " " + self.LT + " " + self.numPlan + " SN" + container.SN
    
    def extract_numplan_from_excel_filename(self):
        """
        Extrait le numéro de plan depuis le nom du fichier Excel (ex: Mesure 24.131.xlsx -> 24.131)
        """
        import os, re
        excel_path = Settings.GetConfigValueString("paths", "measurement_file")
        filename = os.path.basename(excel_path)
        match = re.search(r'(\d+\.\d+)', filename)
        if match:
            return match.group(1)
        return ""

    def CreateNewFilenameFromString(self, formatString : str) :
        string = Writer.ReplaceTagsInString(formatString)
        string = string.replace("*", "")
        return string
    
    def GetOutputFolder(self) :
        """
        Returns the configured output path from Settings, ensuring it is relative
        """
        # Importer la fonction depuis DocumentGenerator
        from DocumentGenerator import ensure_relative_path
        return ensure_relative_path(Settings.GetConfigValueString("paths", "generator_out_path"))
    
    def GetFormulaireNumber(self, string : str) :

        i = 0
        record = False
        buffer = ""
        while i < len(string) :
            if string[i] == "F" :
                record = True
                buffer += string[i]
            elif record :
                if string[i].isnumeric() :
                    buffer += string[i]
                else :
                    self.formulaireNum = buffer
                    Log.Verbose("indice du formulaire : " + self.formulaireNum)
                    return True
            i += 1
        self.formulaireNum = "-"
        return False
    
    def GetFormulaireIndice(self, string : str) :

        index = string.find("Ind ")
        if index == -1 :
            self.formulaireIndice = "-"
            return True
        if index + 4 >= len(string) :
            return False
        self.formulaireIndice = string[index + 4]
        Log.Verbose("indice du formulaire : " + self.formulaireIndice)
        return True
    
    def GetPlan(self) :
        # Comme le numéro de plan n'est plus demandé à l'utilisateur,
        # cette fonction est simplifiée pour toujours réussir
        
        Log.Message("La recherche automatique de plan est désactivée.")
        return True

    def FindSNFromLT(self, launch_number) :
        """
        Cherche tous les SN associés à un numéro de lancement.
        Args:
            launch_number (str): Le numéro de lancement à rechercher
        Returns:
            list: Liste des SN trouvés pour ce LT
        """
        # Vérifier que le LT est valide
        if not launch_number or launch_number == "":
            Log.Error("Le numéro de lancement est vide")
            return []
        if not launch_number.replace(' ', '').startswith("LT"):
            Log.Error("Le numéro de lancement doit commencer par 'LT'")
            return []
        if not launch_number.replace(' ', '')[2:].isdigit():
            Log.Error("Le numéro de lancement doit être suivi de chiffres")
            return []
        # S'assurer que les données sont chargées
        if not self.containers:
            Log.Warning("Aucune donnée chargée, tentative de lecture du fichier Excel")
            if not self.FetchValuesInExcel():
                return []
        # Chercher dans les containers existants
        sns = []
        for container in self.containers:
            container_lt = None
            if hasattr(container, 'tagAndValues') and 'LT' in container.tagAndValues:
                container_lt = container.tagAndValues['LT']
            elif hasattr(container, 'LT'):
                container_lt = container.LT
            # Nettoyer les espaces
            container_lt_clean = str(container_lt).replace(' ', '').strip() if container_lt else ''
            launch_number_clean = str(launch_number).replace(' ', '').strip()
            # Ajouter "LT" si manquant
            if container_lt_clean and not container_lt_clean.startswith("LT"):
                container_lt_clean = "LT" + container_lt_clean
            # DEBUG : Afficher la comparaison
            Log.Message(f"Comparaison filtrage : container_lt_clean='{container_lt_clean}' (len={len(container_lt_clean)}) vs launch_number_clean='{launch_number_clean}' (len={len(launch_number_clean)})")
            if container_lt_clean.replace("LT", "").lstrip("0") == launch_number_clean.replace("LT", "").lstrip("0"):
                sns.append(container.SN)
        if not sns:
            Log.Warning(f"Aucun SN trouvé pour le numéro de lancement {launch_number}")
        else:
            Log.Message(f"SN trouvés pour {launch_number} : {', '.join(sns)}")
        return sns


def initialize_value_collector():
    """Initialise l'instance globale de ValueCollector si elle n'existe pas déjà."""
    global valueCollector
    if valueCollector is None:
        valueCollector = ValueCollector()
        Log.Message("Instance globale de ValueCollector initialisée")
    return valueCollector

# Initialiser l'instance globale après la définition de la classe
initialize_value_collector()



"""

Settings.LoadPaths()

thing = ValueCollector()

thing.LT = "LT2400182"
thing.numPlan = "23.199"

result = thing.FetchValuesInExcel()
print("result :", result)

"""