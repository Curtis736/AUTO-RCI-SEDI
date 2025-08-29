import docx
from docx.shared import Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from datetime import datetime
import os
import re

import Log
import Settings

# Vérifier si les importations supplémentaires nécessaires sont disponibles
try:
    import win32com.client
    import pythoncom
    has_win32com = True
except ImportError:
    has_win32com = False
    Log.Warning("win32com n'est pas disponible, certaines fonctionnalités avancées avec Word ne seront pas accessibles")

documentDefined = False
document : docx.Document
pathToDocument : str


# definition of constants

MARKER_REPLACE = "**"
MARKER_IMAGE = "§§"
MARKER_COLUMN = "$$"


# this dictionary contains replacement tags as key and their replacement as value
replacementMap = {}
# tags as key, and list of strings as value
columnContentMap = {}
# tags as key, and --- as value
imageMap = {}


def Open(path) :
    global document, pathToDocument, documentDefined
    
    # Vérifier si le chemin est valide
    if not os.path.exists(path):
        Log.Error(f"Le fichier modèle n'existe pas: {path}")
        return False
        
    if not path.endswith("docx") :
        Log.Error("error : file is not a docx")
        return False
    
    # Vérifier si Word est disponible en cas de problème
    word_check_done = False
    
    pathToDocument = path
    # cette ligne va causer un packagenotfounderror si le docx est vide (ou invalide)
    try :
        Log.Message(f"Ouverture du document Word: {path}")
        document = docx.Document(pathToDocument)
        documentDefined = True
        return True
    except Exception as e :
        Log.Error("couldn't open docx")
        Log.Error(str(e))
        
        # Si l'erreur semble être liée à un problème avec le document Office
        if "Package not found" in str(e) or "document.xml" in str(e):
            if not word_check_done:
                word_check_done = True
                if not CheckWordAvailability():
                    Log.Error("Veuillez vérifier que Microsoft Word est correctement installé et configuré")
        
        return False

def Save(filePath : str) :
    global document, pathToDocument, documentDefined
    if not documentDefined or document is None:
        Log.Error("ERREUR CRITIQUE: Document Word non défini ou est None")
        return False
    
    # Normaliser le chemin pour éviter les problèmes avec les slashes
    try:
        # Vérifier que le chemin n'est pas vide
        if not filePath:
            Log.Error("SAUVEGARDE - Chemin du fichier vide ou non défini")
            return False
            
        # Traiter le chemin selon qu'il est absolu ou relatif
        if os.path.isabs(filePath):
            # Chemin absolu
            filePath = os.path.normpath(filePath)
        else:
            # Chemin relatif - le convertir en absolu
            filePath = os.path.abspath(filePath)
            
        Log.Message(f"SAUVEGARDE - Tentative de sauvegarde du document à: {filePath}")
        
        # S'assurer que le répertoire de destination existe
        outDir = os.path.dirname(filePath)
        if not os.path.exists(outDir):
            try:
                os.makedirs(outDir, exist_ok=True)
                Log.Message(f"SAUVEGARDE - Création du répertoire: {outDir}")
            except Exception as e:
                Log.Error(f"SAUVEGARDE - Impossible de créer le répertoire {outDir}: {str(e)}")
                # Nous essaierons de sauvegarder quand même, l'erreur sera gérée plus bas
        else:
            Log.Message(f"SAUVEGARDE - Le répertoire existe déjà: {outDir}")
        
        # Tester l'accès en écriture au répertoire
        try:
            test_file = os.path.join(outDir, "test_write_access.tmp")
            with open(test_file, 'w') as f:
                f.write("test")
            os.remove(test_file)
            Log.Message(f"SAUVEGARDE - Le répertoire {outDir} est accessible en écriture")
        except Exception as e:
            Log.Warning(f"SAUVEGARDE - Le répertoire {outDir} pourrait ne pas être accessible en écriture: {str(e)}")
            Log.Warning("Nous allons quand même essayer de sauvegarder le document")
            # Nous ne sortons pas ici, nous continuons à essayer de sauvegarder
    
        # Essai de sauvegarde
        try:
            Log.Message(f"SAUVEGARDE - Tentative de sauvegarde avec docx.Document.save()")
            document.save(filePath)
            Log.Message(f"SAUVEGARDE - Document sauvegardé avec succès: {filePath}")
            
            # Vérifier immédiatement que le fichier a bien été créé
            if os.path.exists(filePath):
                Log.Message(f"SAUVEGARDE - Vérification réussie: le fichier {filePath} existe")
                # Fermer et libérer le document
                Close()
                return True
            else:
                Log.Error(f"SAUVEGARDE - ERREUR: Le fichier {filePath} n'a pas été créé malgré la réussite de document.save()")
                # Ne pas sortir, essayer des solutions alternatives
        except PermissionError as perm_e:
            Log.Error(f"SAUVEGARDE - Permission refusée: {str(perm_e)}")
            Log.Error("Le fichier est peut-être ouvert dans une autre application ou vous n'avez pas les droits d'écriture.")
            # Continuer avec les tentatives alternatives
        except Exception as e:
            Log.Error(f"SAUVEGARDE - Échec de la sauvegarde: {str(e)}")
            # Continuer avec les tentatives alternatives
        
        # Tentative alternative 1: Essayer de sauvegarder dans le répertoire courant
        try:
            fileName = os.path.basename(filePath)
            altPath = os.path.join(os.getcwd(), fileName)
            Log.Message(f"SAUVEGARDE - Tentative alternative 1: Sauvegarde dans le répertoire courant: {altPath}")
            document.save(altPath)
            
            if os.path.exists(altPath):
                Log.Message(f"SAUVEGARDE - Sauvegarde alternative réussie: {altPath}")
                # Fermer et libérer le document
                Close()
                return True
        except Exception as alt_e:
            Log.Error(f"SAUVEGARDE - Échec de la tentative alternative 1: {str(alt_e)}")
        
        # Tentative alternative 2: Essayer de sauvegarder dans un répertoire "output" créé spécifiquement
        try:
            fileName = os.path.basename(filePath)
            outputDir = os.path.join(os.getcwd(), "output")
            if not os.path.exists(outputDir):
                os.makedirs(outputDir, exist_ok=True)
            
            altPath = os.path.join(outputDir, fileName)
            Log.Message(f"SAUVEGARDE - Tentative alternative 2: Sauvegarde dans {altPath}")
            document.save(altPath)
            
            if os.path.exists(altPath):
                Log.Message(f"SAUVEGARDE - Sauvegarde alternative 2 réussie: {altPath}")
                # Fermer et libérer le document
                Close()
                return True
        except Exception as alt_e:
            Log.Error(f"SAUVEGARDE - Échec de la tentative alternative 2: {str(alt_e)}")
        
        # Tentative alternative 3: Essayer de sauvegarder avec un nom de fichier plus simple
        try:
            simplified_name = f"document_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
            altPath = os.path.join(os.getcwd(), simplified_name)
            Log.Message(f"SAUVEGARDE - Tentative alternative 3: Sauvegarde avec un nom simplifié: {altPath}")
            document.save(altPath)
            
            if os.path.exists(altPath):
                Log.Message(f"SAUVEGARDE - Sauvegarde alternative 3 réussie: {altPath}")
                # Fermer et libérer le document
                Close()
                return True
        except Exception as alt_e:
            Log.Error(f"SAUVEGARDE - Échec de la tentative alternative 3: {str(alt_e)}")
        
        # Si toutes les tentatives ont échoué
        Log.Error("SAUVEGARDE - ERREUR CRITIQUE: Impossible de sauvegarder le document après plusieurs tentatives")
        # Fermer et libérer le document même en cas d'échec
        Close()
        return False
        
    except Exception as e:
        Log.Error(f"SAUVEGARDE - Exception générale lors de la sauvegarde: {str(e)}")
        # Fermer et libérer le document même en cas d'échec
        Close()
        return False

def ReplaceTagsInString(string : str) :
    for tag, replacement in replacementMap.items() :
        string = string.replace(tag, replacement)
    return string

def SetReplacementMapping(key : str, value : str) :
    global replacementMap
    replacementMap[FormatTagForTextReplacement(key)] = value

def SetColumnFillMapping(key : str, value : str) :
    global columnContentMap
    columnContentMap[FormatTagForColumnFill(key)] = [value]


# methods to deal with the tags.
# they use the contents of the defined dictionaries at the top
def __ReplaceAllTags(pulseCallback = None) :
    tagList = []
    replacementList = []
    for key, value in replacementMap.items() :
        tagList.append(key)
        replacementList.append(value)
    
    print(tagList)
    
    __ReplaceAll(tagList, replacementList, pulseCallback)

# methods to deal with the tags.
# they use the contents of the defined dictionaries at the top
def ReplaceAllTags(pulseCallback = None) :
    # Nouvelle méthode améliorée:
    # 1. D'abord, trouver tous les tags non remplacés dans le document
    # 2. Essayer de trouver des valeurs pour ces tags
    # 3. Remplacer tous les tags trouvés
    
    # Trouver tous les tags non remplacés
    unreplaced_tags = GetUnreplacedTags()
    
    if unreplaced_tags:
        Log.Message(f"Trouvé {len(unreplaced_tags)} tags non remplacés dans le document")
        for tag in unreplaced_tags:
            # Extraire le nom du tag sans les délimiteurs
            if tag.startswith("**") and tag.endswith("**"):
                tag_name = tag[2:-2]  # Enlever les ** au début et à la fin
                Log.Message(f"Tag trouvé dans le document: {tag} (nom: {tag_name})")
                
                # Vérifier si une valeur existe pour ce tag
                if tag not in replacementMap:
                    # Le tag n'est pas encore dans le dictionnaire de remplacement
                    # Essayer de trouver une valeur dans les données chargées
                    found_value = None
                    
                    # 1. Chercher dans les variables prédéfinies de Settings
                    if hasattr(Settings, f"TAG_{tag_name.upper()}"):
                        tag_constant = getattr(Settings, f"TAG_{tag_name.upper()}")
                        Log.Message(f"Tag {tag_name} correspond à la constante Settings.{tag_constant}")
                    
                    # 2. Si le tag est déjà présent mais dans un format différent, l'utiliser
                    formatted_tag = FormatTagForTextReplacement(tag_name)
                    if formatted_tag in replacementMap:
                        found_value = replacementMap[formatted_tag]
                        Log.Message(f"Valeur trouvée pour {tag_name}: {found_value}")
                    
                    # 3. Si le tag n'a pas encore de valeur, mettre une valeur par défaut ou un message
                    if found_value is None:
                        # Utiliser une valeur vide pour éviter les erreurs
                        found_value = ""
                        Log.Warning(f"Aucune valeur trouvée pour le tag {tag}. Utilisation d'une chaîne vide.")
                    
                    # Ajouter au dictionnaire de remplacement
                    replacementMap[tag] = found_value
    
    # Mettre à jour le dictionnaire de remplacement pour s'assurer que tous les tags sont au format **TAG**
    enhanced_replacement_map = {}
    for key, value in replacementMap.items():
        # S'assurer que la clé est au format **TAG**
        if not (key.startswith("**") and key.endswith("**")):
            enhanced_key = f"**{key}**"
            enhanced_replacement_map[enhanced_key] = value
            Log.Message(f"Reformatage du tag {key} en {enhanced_key}")
        else:
            enhanced_replacement_map[key] = value
    
    # Fusionner avec le dictionnaire original
    replacementMap.update(enhanced_replacement_map)
    
    # Maintenant remplacer tous les tags
    tagList = []
    replacementList = []
    for key, value in replacementMap.items():
        tagList.append(key)
        replacementList.append(value)
    
    Log.Message(f"Remplacement de {len(tagList)} tags dans le document:")
    for i, tag in enumerate(tagList):
        Log.Message(f"  - {tag} -> {replacementList[i]}")
    
    # Appeler la fonction de remplacement
    __ReplaceAll(tagList, replacementList, pulseCallback)
    
    # Vérifier s'il reste des tags non remplacés
    remaining_tags = GetUnreplacedTags()
    if remaining_tags:
        Log.Warning(f"Il reste {len(remaining_tags)} tags non remplacés dans le document:")
        for tag in remaining_tags:
            Log.Warning(f"  - {tag}")
    else:
        Log.Message("Tous les tags ont été remplacés avec succès!")

def EnhancedTagReplacement(container):
    """
    Fonction améliorée pour le remplacement des tags qui s'assure que tous les tags
    dans le document sont correctement remplacés.
    
    Args:
        container: Conteneur contenant les valeurs à utiliser pour le remplacement
    """
    # Réinitialiser les dictionnaires de remplacement
    global replacementMap, columnContentMap
    replacementMap.clear()
    columnContentMap.clear()
    
    # 1. D'abord, analyser le document pour trouver tous les tags
    Log.Message("Analyse du document pour trouver tous les tags...")
    unreplaced_tags = GetUnreplacedTags()
    
    # Extraire les noms de tags sans les délimiteurs
    tag_names = []
    for tag in unreplaced_tags:
        if tag.startswith("**") and tag.endswith("**"):
            tag_name = tag[2:-2]  # Enlever les ** au début et à la fin
            tag_names.append(tag_name)
            Log.Message(f"Tag trouvé: {tag} (nom: {tag_name})")
    
    # 2. IMPORTANT: Ajout de TOUTES les valeurs disponibles dans le container aux dictionnaires de remplacement
    Log.Message("Ajout de toutes les valeurs du container aux dictionnaires de remplacement...")
    for key, value in container.tagAndValues.items():
        if value is not None and value != "":
            # Ajouter à la fois au format **TAG** et au format normal
            tag_key = f"**{key}**"
            replacementMap[tag_key] = value
            columnContentMap[f"$${key}$$"] = [value]
            Log.Message(f"Valeur ajoutée du container: {key} = {value}")
    
    # 3. Remplir le dictionnaire avec les valeurs explicitement demandées dans le document
    for tag_name in tag_names:
        tag_key = f"**{tag_name}**"
        
        # Si le tag est déjà dans le dictionnaire de remplacement, passer
        if tag_key in replacementMap:
            Log.Message(f"Tag {tag_name} déjà configuré avec valeur: {replacementMap[tag_key]}")
            continue
            
        # Chercher dans les valeurs du container
        if tag_name in container.tagAndValues:
            value = container.tagAndValues[tag_name]
            replacementMap[tag_key] = value
            columnContentMap[f"$${tag_name}$$"] = [value]
            Log.Message(f"Valeur trouvée pour {tag_name}: {value}")
        else:
            # Tag non trouvé dans le container, essayer de trouver une valeur alternative
            if tag_name == "DATE":
                # Utiliser la date actuelle
                from datetime import datetime
                current_date = datetime.now().strftime("%d/%m/%Y")
                replacementMap[tag_key] = current_date
                columnContentMap[f"$${tag_name}$$"] = [current_date]
                Log.Message(f"Utilisation de la date actuelle pour {tag_name}: {current_date}")
            elif tag_name == "LT" and hasattr(container, "LT"):
                # Utiliser le numéro de lancement du container
                replacementMap[tag_key] = container.LT
                columnContentMap[f"$${tag_name}$$"] = [container.LT]
                Log.Message(f"Utilisation du LT du container pour {tag_name}: {container.LT}")
            elif tag_name == "NUMPLAN" and hasattr(container, "numPlan"):
                # Utiliser le numéro de plan du container
                replacementMap[tag_key] = container.numPlan
                columnContentMap[f"$${tag_name}$$"] = [container.numPlan]
                Log.Message(f"Utilisation du NUMPLAN du container pour {tag_name}: {container.numPlan}")
            else:
                # Tag non reconnu, utiliser une chaîne vide
                replacementMap[tag_key] = ""
                columnContentMap[f"$${tag_name}$$"] = [""]
                Log.Warning(f"Aucune valeur trouvée pour le tag {tag_name}. Utilisation d'une chaîne vide.")
    
    # 4. Ajouter des alias pour les noms de tags courants qui pourraient être écrits différemment
    common_aliases = {
        "SN": ["NUMERO_SERIE", "NUMEROSERIE", "NUM_SERIE", "NUMSERIE"],
        "LT": ["LANCEMENT", "NUMERO_LANCEMENT", "NUM_LANCEMENT"],
        "NUMPLAN": ["PLAN", "NUMERO_PLAN", "NUM_PLAN"],
        "DATE": ["DATE_ESSAI", "DATEESSAI", "DATE_TEST", "DATETEST"],
        "IL_BEFORE": ["IL_AVANT", "ILAVANT", "PERTE_AVANT"],
        "IL_AFTER": ["IL_APRES", "ILAPRES", "PERTE_APRES"],
        "RL1": ["RETURN_LOSS1", "RETURNLOSS1", "RL_1"],
        "RL2": ["RETURN_LOSS2", "RETURNLOSS2", "RL_2"]
    }
    
    for base_tag, aliases in common_aliases.items():
        if base_tag in container.tagAndValues:
            base_value = container.tagAndValues[base_tag]
            if base_value is not None and base_value != "":
                # Ajouter tous les alias possibles
                for alias in aliases:
                    alias_key = f"**{alias}**"
                    if alias_key not in replacementMap:
                        replacementMap[alias_key] = base_value
                        columnContentMap[f"$${alias}$$"] = [base_value]
                        Log.Message(f"Ajout d'un alias pour {base_tag}: {alias} = {base_value}")
    
    # 5. Effectuer le remplacement
    Log.Message(f"Remplacement de {len(replacementMap)} tags dans le document...")
    ReplaceAllTags()
    
    # 6. Vérifier s'il reste des tags non remplacés
    remaining_tags = GetUnreplacedTags()
    if remaining_tags:
        Log.Warning(f"Il reste {len(remaining_tags)} tags non remplacés dans le document:")
        for tag in remaining_tags:
            Log.Warning(f"  - {tag}")
            
        # Nouvelle tentative de remplacement avec des formats différents
        second_attempt_replacements = {}
        for tag in remaining_tags:
            if tag.startswith("**") and tag.endswith("**"):
                tag_name = tag[2:-2]  # Enlever les ** au début et à la fin
                
                # Essayer d'autres formats courants
                if tag_name in container.tagAndValues:
                    value = container.tagAndValues[tag_name]
                    second_attempt_replacements[tag] = value
                    Log.Message(f"Deuxième tentative pour {tag}: {value}")
                # Essayer en minuscules
                elif tag_name.lower() in container.tagAndValues:
                    value = container.tagAndValues[tag_name.lower()]
                    second_attempt_replacements[tag] = value
                    Log.Message(f"Deuxième tentative (minuscules) pour {tag}: {value}")
                # Essayer en majuscules
                elif tag_name.upper() in container.tagAndValues:
                    value = container.tagAndValues[tag_name.upper()]
                    second_attempt_replacements[tag] = value
                    Log.Message(f"Deuxième tentative (majuscules) pour {tag}: {value}")
                # Essayer avec des underscores -> espaces
                elif tag_name.replace("_", " ") in container.tagAndValues:
                    value = container.tagAndValues[tag_name.replace("_", " ")]
                    second_attempt_replacements[tag] = value
                    Log.Message(f"Deuxième tentative (underscores) pour {tag}: {value}")
                # Essayer avec des espaces -> underscores
                elif tag_name.replace(" ", "_") in container.tagAndValues:
                    value = container.tagAndValues[tag_name.replace(" ", "_")]
                    second_attempt_replacements[tag] = value
                    Log.Message(f"Deuxième tentative (espaces) pour {tag}: {value}")
                else:
                    # Dernier recours: valeur vide
                    second_attempt_replacements[tag] = ""
                    Log.Warning(f"Aucune correspondance trouvée pour {tag}, utilisation d'une chaîne vide.")
        
        # Appliquer les remplacements de seconde tentative
        if second_attempt_replacements:
            replacementMap.update(second_attempt_replacements)
            tagList = []
            replacementList = []
            for key, value in second_attempt_replacements.items():
                tagList.append(key)
                replacementList.append(value)
            __ReplaceAll(tagList, replacementList)
            
            # Vérifier une dernière fois
            final_remaining_tags = GetUnreplacedTags()
            if not final_remaining_tags:
                Log.Message("Tous les tags ont été remplacés avec succès après la deuxième tentative!")
            else:
                Log.Warning(f"Il reste encore {len(final_remaining_tags)} tags non remplacés après la deuxième tentative.")
    else:
        Log.Message("Tous les tags ont été remplacés avec succès!")
    
    # 7. Remplir les colonnes après avoir remplacé les tags
    FillAllColumns()
    
    # Retourner True si tous les tags ont été remplacés
    final_check = GetUnreplacedTags()
    return len(final_check) == 0


"""
Find all the text-based tags in the documents, and attempts to replace them with a value from the container.
Return None if everything went fine, returns a list if some tags were not found in the container.
"""
def ReplaceAllTextTags(pulseCallback) :

    # verify the document is open
    if not documentDefined : return False
    
    __ReplaceAllTags(pulseCallback)

    missed = GetUnreplacedTags()

    if len(missed) == 0 : return True

    return missed



    

def FillAllColumns() :
    tagList = []
    replacementList = []
    for key, value in columnContentMap.items() :
        tagList.append(key)
        replacementList.append(value)
    
    print(tagList)
    
    __FillAllTables(tagList, replacementList)


def FindAllIndices(indiceTag) :
    allFounds = []
    i = 0
    for paragraph in document.paragraphs :
        allFounds = allFounds + FindTagOccurencesInPAragraph(paragraph, i, indiceTag)
        i += 1
    Log.Verbose("indices : " + str(allFounds))
    return allFounds

# return a list, with one tuple per occurence, whose only element is the index of the paragraph
def FindTagOccurencesInPAragraph(paragraph, paragraphIndex, tag) :
    found = []
    finding = 0
    index = 0
    for letter in paragraph.text :
        if finding > 0 :
            if letter == tag[finding] :
                finding += 1
                if finding == len(tag) :
                    found.append((paragraphIndex))
                    finding = 0
            else :
                finding = 0
        else :
            if letter == tag[finding] :
                finding += 1
        index += 1
    
    return found

def ReplaceAllInOrder(positionList, tag, replacements) :
    ReplaceInParagraphsInOrder(document.paragraphs, positionList, tag, replacements)

def ReplaceInParagraphsInOrder(paragraphs, positionList, tag, replacements) :
    i = 0
    for paragraphIndex in positionList :
        FindAndReplaceInParagraph(paragraphs[paragraphIndex], tag, replacements[i])
        i += 1


# put the string "content" in the cell, erasing anything that was in it previously
def FillCell(cell, content, alignment = WD_ALIGN_PARAGRAPH.LEFT) :
    # first we remove contents of all paragraphs
    for paragraph in cell.paragraphs :
        for run in paragraph.runs :
            run.text = ""
    
    # set the last run of the last paragraph to the correct value
    # a cell in a word document always have at least one paragraph
    paragraph = cell.paragraphs[len(cell.paragraphs) - 1]
    paragraph.alignment = alignment
    if len(paragraph.runs) == 0 :
        paragraph.add_run(content)
    else :
        run = paragraph.runs[len(paragraph.runs) - 1]
        run.text = content


def __ReplaceAll(tagsToReplace, tagReplacements, pulseCallback = None) :
    for i in range(len(tagsToReplace)) :
        FindAndReplaceInParagraphs(document.paragraphs, tagsToReplace[i], tagReplacements[i], pulseCallback)
        for table in document.tables :
            for row in table.rows :
                for cell in row.cells :
                    FindAndReplaceInParagraphs(cell.paragraphs, tagsToReplace[i], tagReplacements[i], pulseCallback)
        header = document.sections[0].header
        FindAndReplaceInParagraphs(header.paragraphs, tagsToReplace[i], tagReplacements[i], pulseCallback)
        FindAndReplaceInTables(header.tables, tagsToReplace[i], tagReplacements[i])


def FindAndReplaceInParagraphs(paragraphs, find, replace, pulseCallback = None) :
    for paragraph in paragraphs :
        FindAndReplaceInParagraph(paragraph, find, replace)
        if pulseCallback != None :
            pulseCallback()

def FindAndReplaceInTables(tables, find, replace) :
    for table in tables :
        for row in table.rows :
            for cell in row.cells :
                for paragraph in cell.paragraphs :
                    FindAndReplaceInParagraph(paragraph, find, replace)

# self explanatory
def FindAndReplaceInParagraph(paragraph, find, replace) :
    # first find in paragraph text :
    text = ""
    try :
        text = paragraph.text
    except Exception as e :
        Log.Error("couldn't obtain paragraph text")
        Log.Error(str(e))
        return -1

    if not find in text :
        return 0
    
    index = text.find(find)
    Log.Verbose("found " + find + " at " + str(index) + " to replace")

    if index == -1 :
        return 0

    endIndex = index + len(find)
    sizeDiff = len(replace) - len(find)
    currentIndex = 0
    for run in paragraph.runs :
        
        # it is important to read it now because we will change its length with the replacements.
        # it would lead to the program overdeleting if the replacement if shorter than the tag

        if currentIndex <= index and currentIndex + len(run.text) > index :
            if endIndex < currentIndex + len(run.text) :
                Log.Verbose(run.text[:(index - currentIndex)] + "-" + replace + "-" + run.text[(endIndex - currentIndex):])
                run.text = run.text[:(index - currentIndex)] + replace + run.text[(endIndex - currentIndex):]
                return 1
            else :
                #endIndex += (index - currentIndex) - len(run.text) + len(replace)
                prevLength = len(run.text)
                Log.Verbose(run.text + " -> " + run.text[:(index - currentIndex)] + replace)
                run.text = run.text[:(index - currentIndex)] + replace
                endIndex += len(run.text) - prevLength

        elif currentIndex >= index and currentIndex + len(run.text) <= endIndex :
            Log.Verbose("//" + run.text + "//")
            endIndex -= len(run.text)
            run.text = ""
        
        elif currentIndex < endIndex and currentIndex + len(run.text) >= endIndex :
            Log.Verbose("/" + run.text[:(endIndex - currentIndex)] + "/" + run.text[(endIndex - currentIndex + 1):])
            run.text = run.text[(endIndex - currentIndex):]
            return 1

        runLength = len(run.text)
        currentIndex += runLength
    
    return 1

# find something in a paragraph, replace it, and then erase all text after it
def FindAndReplaceInParagraphAndErase(paragraph, find, replace) :
    # first find in paragraph text :
    text = paragraph.text

    if not find in text :
        return
    
    index = text.find(find)
    Log.Verbose("found " + find + " at " + str(index) + " to replace and then erase the rest")

    currentIndex = 0
    replaced = False
    for run in paragraph.runs :
        
        if currentIndex <= index and currentIndex + len(run.text) > index :
            run.text = run.text[:(index - currentIndex)] + replace
            
            replaced = True

        elif replaced :
            run.text == ""

        currentIndex += len(run.text)

def AddRunInSameStyle(paragraph, string : str) :
    if len(paragraph.runs) == 0 :
        paragraph.add_run(string)
        return
    mostRecent = paragraph.runs[-1]
    lastFont = mostRecent.font
    run = paragraph.add_run(string)
    font = run.font
    font.size = lastFont.size
    font.bold = lastFont.bold
    font.italic = lastFont.italic
    font.color.rgb = lastFont.color.rgb
    return run


def __FillAllTables(columnTags, columnValues) :
    FillTables(document.tables, columnTags, columnValues)

def FillTables(tables, columnTags, columnValues) :
    x = 0
    tablesFilled = []
    while x < len(tables) :
        firstRow = tables[x].rows[0]
        y = 0
        while y < len(firstRow.cells) :
            for paragraph in firstRow.cells[y].paragraphs :
                for tag in columnTags :
                    if tag in paragraph.text :
                        FindAndReplaceInParagraph(paragraph, tag, "")
                        FillColumn(tables[x], y, columnValues[columnTags.index(tag)])
                        tablesFilled.append(x)
                """
                for run in paragraph.runs :
                    #print(run.text)
                    for tag in columnTags :
                        if tag in run.text :
                            run.text = run.text.replace("$$" + tag + "$$", "")
                            FillColumn(tables[x], y, columnValues[columnTags.index(tag)])
                            tablesFilled.append(x)
                """
            y +=1
        x += 1

# do not touch the topmost cell
def FillColumn(table, col, contents) :
    if len(table.rows) < len(contents) + 1 :
        for i in range(len(table.rows), len(contents) + 1) :
            table.add_row()

    for i in range(1, len(contents) + 1) :
        # first we remove contents of all paragraphs
        # vestigal code from when it was needed to replace the contents of the cells
        """
        for paragraph in table.rows[i].cells[col].paragraphs :
            for run in paragraph.runs :
                run.text = ""
        """

        # set the last run of the last paragraph to the correct value
        paragraph = table.rows[i].cells[col].paragraphs[len(table.rows[i].cells[col].paragraphs) - 1]
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        AddRunInSameStyle(paragraph, contents[i - 1])
        """
        if len(paragraph.runs) == 0 :
            paragraph.add_run(contents[i - 1])
        else :
            run = paragraph.runs[len(paragraph.runs) - 1]
            run.text = contents[i]
        """
    return True


def ReplaceImageContents(index, newImagePath):
    """Version simplifiée sans ImageReader"""
    try:
        image = document.inline_shapes[index] 
        blip = image._inline.graphic.graphicData.pic.blipFill.blip
        rId = blip.embed
        documentPart = document.part
        imagePart = documentPart.related_parts[rId]

        # Vérifier que le fichier existe
        if not os.path.exists(newImagePath):
            Log.Error(f"L'image {newImagePath} n'existe pas")
            return False

        # Lire directement le fichier image
        with open(newImagePath, "rb") as img:
            imagePart._blob = img.read()
        
        Log.Message(f"Image {newImagePath} remplacée avec succès")
        return True
        
    except Exception as e:
        Log.Error(f"Erreur lors du remplacement d'image: {str(e)}")
        return False

def ReplaceTagWithImageInText(tag, image_paths, titles=None):
    global documentDefined, document

    # make sure all the slashes are forward-facing
    image_paths = [path.replace("\\", "/") for path in image_paths]
    
    if not documentDefined or not document:
        Log.Error("Aucun document n'est ouvert ou initialisé")
        return False
    
    if titles is None:
        titles = [""] * len(image_paths)
    
    try:
        ReplaceTagWithImage(document.paragraphs, tag, image_paths, titles)
        return True
    except Exception as e:
        Log.Error(f"Erreur lors du remplacement du tag {tag} par des images: {str(e)}")
        return False

# !! can change the paragraph alignment, make sure to tell it in the doc !!
def ReplaceTagWithImage(paragraphs, tag, pics, titles) :
    x, y, z = 0, 0, 0

    Log.Verbose("Writer was requested to place " + str(pics) + " at " + tag)

    if len(pics) != len(titles) :
        Log.Error("image and title lists do not have the same length")
        return False

    while x < len(paragraphs) :
        y = 0
        scanProgress = 0
        buffer = ""
        while y < len(paragraphs[x].runs) :
            z = 0
            #print(paragraphs[x].runs[y].text)
            while z < len(paragraphs[x].runs[y].text) :
                letter = paragraphs[x].runs[y].text[z]
                if letter == "§" :
                    scanProgress += 1
                    if scanProgress >= 4 :
                        if buffer == tag :
                            FindAndReplaceInParagraphs(document.paragraphs, FormatTagForImagePlacement(tag), "")
                            #paragraphs[x].runs[y].text = paragraphs[x].runs[y].text.replace("§§" + buffer + "§§", "")
                            for i in range(len(pics)) :
                                # Modification: ne pas ajouter le titre avant l'image
                                run = AddRunInSameStyle(paragraphs[x], "")
                                Log.Verbose("pasting image " + pics[i] + " in the document")
                                run.add_picture(pics[i], width=Inches(6))

                                #if paragraphs[x].alignment == WD_ALIGN_PARAGRAPH.JUSTIFY :
                                # left alignment is forced
                                paragraphs[x].alignment = WD_ALIGN_PARAGRAPH.LEFT

                            
                elif scanProgress == 2 :
                    buffer += letter
                else :
                    scanProgress = 0
                    buffer = ""
                
                z += 1
            y += 1
        x += 1
    
    

def DocumentContainsInText(string : str) :
    global documentDefined, document
    
    if not documentDefined or not document:
        Log.Error("Aucun document n'est ouvert ou initialisé")
        return False
    
    try:
        header = document.sections[0].header
        for paragraph in header.paragraphs :
            if string in paragraph.text :
                return True
        
        for paragraph in document.paragraphs :
            if string in paragraph.text :
                return True
        
        return False
    except Exception as e:
        Log.Error(f"Erreur lors de la recherche de texte dans le document: {str(e)}")
        return False

def FormatTagForTextReplacement(tag : str) :
    return "**" + tag + "**"

def FormatTagForColumnFill(tag : str) :
    return "$$" + tag + "$$"

def FormatTagForImagePlacement(tag : str) :
    return "§§" + tag + "§§"

def GetAllParagraphs():
    """
    Récupère tous les paragraphes du document actuel.
    
    Returns:
        list: Liste de tous les paragraphes du document.
    """
    global documentDefined, document
    
    if not documentDefined or not document:
        Log.Error("Aucun document n'est ouvert ou initialisé")
        return []
    
    try:
        # Récupérer tous les paragraphes du document
        paragraphs = document.paragraphs
        return paragraphs
    except Exception as e:
        Log.Error(f"Erreur lors de la récupération des paragraphes: {str(e)}")
        return []

def GetUnreplacedTags():
    """
    Recherche dans le document les tags qui n'ont pas été remplacés.
    
    Returns:
        list: Liste des tags non remplacés trouvés dans le document
    """
    global documentDefined, document
    
    if not documentDefined or not document:
        Log.Error("Aucun document n'est ouvert ou initialisé")
        return []
    
    unreplaced_tags = []
    
    try:
        # Rechercher les tags de remplacement (**TAG**)
        tag_pattern1 = re.compile(r'\*\*(.*?)\*\*')
        
        # Rechercher les tags de colonnes ($$TAG$$)
        tag_pattern2 = re.compile(r'\$\$(.*?)\$\$')
        
        # Rechercher les tags d'images (§§TAG§§)
        tag_pattern3 = re.compile(r'§§(.*?)§§')
        
        # Parcourir tous les paragraphes du document
        for paragraph in document.paragraphs:
            if not paragraph or not hasattr(paragraph, 'text'):
                continue
                
            # Rechercher les tags dans le texte du paragraphe
            for match in tag_pattern1.finditer(paragraph.text):
                tag = match.group(0)  # Tag complet avec délimiteurs
                if tag not in unreplaced_tags:
                    unreplaced_tags.append(tag)
            
            for match in tag_pattern2.finditer(paragraph.text):
                tag = match.group(0)  # Tag complet avec délimiteurs
                if tag not in unreplaced_tags:
                    unreplaced_tags.append(tag)
            
            for match in tag_pattern3.finditer(paragraph.text):
                tag = match.group(0)  # Tag complet avec délimiteurs
                if tag not in unreplaced_tags:
                    unreplaced_tags.append(tag)
        
        # Parcourir également les tableaux
        for table in document.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        # Rechercher les tags dans le texte du paragraphe
                        for match in tag_pattern1.finditer(paragraph.text):
                            tag = match.group(0)  # Tag complet avec délimiteurs
                            if tag not in unreplaced_tags:
                                unreplaced_tags.append(tag)
                        
                        for match in tag_pattern2.finditer(paragraph.text):
                            tag = match.group(0)  # Tag complet avec délimiteurs
                            if tag not in unreplaced_tags:
                                unreplaced_tags.append(tag)
        
        # Parcourir les en-têtes
        try:
            for section in document.sections:
                header = section.header
                for paragraph in header.paragraphs:
                    # Rechercher les tags dans le texte du paragraphe
                    for match in tag_pattern1.finditer(paragraph.text):
                        tag = match.group(0)  # Tag complet avec délimiteurs
                        if tag not in unreplaced_tags:
                            unreplaced_tags.append(tag)
                    
                    for match in tag_pattern2.finditer(paragraph.text):
                        tag = match.group(0)  # Tag complet avec délimiteurs
                        if tag not in unreplaced_tags:
                            unreplaced_tags.append(tag)
        except Exception as e:
            Log.Warning(f"Erreur lors de la lecture des en-têtes: {str(e)}")
        
        return unreplaced_tags
    
    except Exception as e:
        Log.Error(f"Erreur lors de la recherche des tags non remplacés: {str(e)}")
        return []

def Close():
    """
    Ferme proprement le document Word actuel pour libérer les ressources.
    Cela peut être utile pour éviter des problèmes de verrous sur les fichiers.
    """
    global document, documentDefined
    
    if not documentDefined or document is None:
        Log.Message("Aucun document à fermer")
        return
    
    try:
        # Python-docx n'a pas de méthode close() explicite,
        # mais on peut simplement supprimer la référence au document
        document = None
        documentDefined = False
        Log.Message("Document fermé avec succès")
    except Exception as e:
        Log.Error(f"Erreur lors de la fermeture du document: {str(e)}")

def CheckWordAvailability():
    """
    Vérifie si Microsoft Word est correctement installé et accessible.
    
    Returns:
        bool: True si Word est disponible, False sinon
    """
    if not has_win32com:
        Log.Warning("Impossible de vérifier Word car win32com n'est pas disponible")
        return False
    
    try:
        # Vérifier si win32timezone est disponible
        try:
            import win32timezone
            Log.Message("Module win32timezone disponible")
        except ImportError as timezone_err:
            Log.Warning(f"Module win32timezone non disponible: {str(timezone_err)}")
            Log.Warning("Tentative d'utilisation alternative...")
            # Certaines installations fonctionnent sans win32timezone explicite
        
        # Initialiser COM pour ce thread
        try:
            pythoncom.CoInitialize()
            Log.Message("COM initialisé avec succès")
        except Exception as com_err:
            Log.Error(f"Erreur lors de l'initialisation COM: {str(com_err)}")
            return False
        
        # Tenter de créer une instance de Word
        try:
            word_app = win32com.client.Dispatch("Word.Application")
            Log.Message("Instance Word créée avec succès")
            
            # Fermer Word
            word_app.Quit()
            Log.Message("Instance Word fermée proprement")
            
            # Libérer COM
            pythoncom.CoUninitialize()
            
            Log.Message("Microsoft Word est correctement installé et accessible")
            return True
        except Exception as word_err:
            Log.Error(f"Erreur lors de la création de l'instance Word: {str(word_err)}")
            # Libérer COM en cas d'erreur
            try:
                pythoncom.CoUninitialize()
            except:
                pass
            return False
            
    except Exception as e:
        Log.Error(f"Erreur lors de la vérification de Word: {str(e)}")
        Log.Error("Microsoft Word n'est pas correctement installé ou accessible")
        
        # Essayer de donner des informations supplémentaires sur l'erreur
        if "Object library invalid" in str(e):
            Log.Error("La bibliothèque Word COM est invalide. Office pourrait être mal installé ou corrompu.")
        elif "win32timezone" in str(e):
            Log.Error("Problème avec le module win32timezone. Vérifiez votre installation de PyWin32.")
        
        # Libérer COM en cas d'erreur
        try:
            pythoncom.CoUninitialize()
        except:
            pass
            
        return False

def AddImagesFromFolder(folder_path, tag_content, sn, filter_pattern):
    """
    Ajoute des images depuis un dossier spécifique en remplaçant un tag personnalisé.
    
    Args:
        folder_path (str): Chemin vers le dossier contenant les images
        tag_content (str): Contenu du tag à remplacer (ex: "INTERFERO_E")
        sn (str): Numéro de série
        filter_pattern (str): Motif de filtre optionnel
    """
    global documentDefined, document
    
    if not documentDefined or not document:
        Log.Error("Aucun document n'est ouvert ou initialisé")
        return False
    
    try:
        # Importer la fonction de recherche d'images depuis DocumentGenerator
        from DocumentGenerator import FindImagesMatchingSN
        
        # Trouver les images correspondant au SN
        sn_images, all_images = FindImagesMatchingSN(folder_path, sn, filter_pattern)
        
        if not sn_images:
            Log.Warning(f"Aucune image trouvée pour SN{sn} dans le dossier {folder_path}")
            return False
        
        # Créer des titres pour les images
        titles = []
        for i, img_path in enumerate(sn_images):
            img_name = os.path.basename(img_path)
            titles.append(f"Image {i+1}: {img_name}")
        
        # Remplacer le tag par les images
        tag_to_replace = tag_content
        Log.Message(f"Remplacement du tag {tag_to_replace} par {len(sn_images)} image(s)")

        import ImageReader
        # Convert all the image files found into PNG images, temporarily
        for i in range(len(sn_images)) :
            if sn_images[i].endswith(".xlsx") or sn_images[i].endswith(".xlsm") :
                import ExcelController
                ExcelController.OpenExcel()
                path = ExcelController.FindGraphInExcelFile(sn_images[i], legendKeywords=[f"S/N{sn}", f"S/N {sn}", f"SN{sn}", f"SN {sn}"])
                if path != None : sn_images[i] = path
                else : Log.Error(f"Ne peut pas charger image {sn_images[i]}")
            else :
                path = ImageReader.GetImagePath(sn_images[i])
                if path != None : sn_images[i] = path
                else : Log.Error(f"Ne peut pas charger image {sn_images[i]}")
        
        # Utiliser la fonction existante pour remplacer le tag par les images
        Log.Verbose(f"Chemins des images : {sn_images}")
        ReplaceTagWithImageInText(tag_to_replace, sn_images, titles)
        
        ImageReader.Clear()
        return True
        
    except Exception as e:
        Log.Error(f"Erreur lors de l'ajout d'images depuis le dossier {folder_path}: {str(e)}")
        return False