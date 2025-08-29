"""
All data is now put under a fixed structure in the same folder. This means that it is now possible to collect more infos based on folder names and structures
"""

import os
import sys
import re

import Log
import Settings

# temporairement le chemin est une constante stockée ici

rootFolder = "X:/Tracabilite"  # ← MODIFIABLE
formFolder = "X:/Qualite/4_Public/A disposition/DOSSIER SMI/Formulaires/B3-PRODUCTION"  # ← MODIFIABLE


def GetPathLength(path : str) :

    path = path.strip("/").strip("\\")

    if "/" in path :
        return path.count("/") + 1
    elif "\\" in path :
        return path.count("\\") + 1
    elif path == "" :
        return 0
    else :
        return 1

# return None on error, otherwise a list of tuples (relativePath, re.Match)
def FindInFoldersRecursively(folderPath : str, expression : re.Pattern, maxDepth : int = -1) :

    if not os.path.isdir(folderPath) : return None

    absFolderPath = os.path.abspath(folderPath)

    results = [] # the list of matches, returns a tuple path, match
    waitList = [folderPath] # the queue for evaluation

    while len(waitList) > 0 :
        for element in os.scandir(waitList.pop()) :

            # only check dirs
            if element.is_dir() :
                # calculate the depth of the dir based on the root path
                absPath = os.path.abspath(element.path)
                relPath = os.path.relpath(element.path, absFolderPath)
                depth = GetPathLength(relPath)
                if depth < maxDepth or maxDepth == -1 :
                    waitList.append(absPath)
                
                result = expression.search(os.path.basename(relPath))
                if result != None :
                    results.append((absPath, result))

    return results

    

    

def GetPathToLT(lt : str) :

    results = FindInFoldersRecursively(rootFolder, re.compile(lt), 3)

    if len(results) == 0 :
        return False
    elif len(results) > 1 :
        Log.Error("there is more than one path to the same LT, shouldn't be possible")
        return False
    
    return results[0]


def GetMeasurementsFromLT(absPathToLt : str) :

    # the measurement file should be one directory up from the LT directory
    measurementDirectory = os.path.join(absPathToLt, os.path.pardir)
    measurementDirectory = os.path.normpath(measurementDirectory)

    # now try to find an excel file with the correct name
    expression = re.compile("(?i:mesures)")

    for thing in os.scandir(measurementDirectory) :
        if thing.is_file() :
            result = expression.search(thing.name)

            if result != None :
                return os.path.abspath(thing.path)
    
    return None

"""
Attempts to find a form with the specified number and indice.
Expects the number without the F, and the indice with just the letter.
Returns the path to the form if found, otherwise None.
"""
def GetForm(formNumber : str, indice : str) :

    regexNumber = re.compile(f"F\\s?{formNumber}")
    regexIndice = re.compile(f"(?i:ind)\\s?{indice}")

    queue = [formFolder]

    while len(queue) > 0 :

        currentFolder = queue.pop()
        for element in os.scandir(currentFolder) :

            if element.is_dir() :
                queue.append(element.path)
            
            elif element.is_file() :
                if regexNumber.search(element.name) :
                    if regexIndice.search(element.name) :
                        return os.path.join(currentFolder, element.path)
    
    return None

