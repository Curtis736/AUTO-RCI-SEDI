import os
import sys
import json

import Log


"""
Allow for easy retrieval from all kind of settings in the form of a JSON file.
Every file must be a JSON dictionary.
The idea is that the rest of the program only have to request reading and writing from the file and this program
will create and save them as needed.
"""


# constants
CONFIG_FILES_PATH = "config"


def _project_root() -> str:
    """Racine du dépôt (parent de src/), indépendamment du cwd."""
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


CWD = _project_root()

"""
A dictionary of all the loaded setting files. The key is the name of the file and the value is a dictionary from the contents of the file.
"""
__loadedFiles = {}


def EnsureDirectoryExists() :
    os.makedirs(CONFIG_FILES_PATH, exist_ok=True)


"""
Put the content of the requested config file in the loaded files, regardless of if it was already there or not.
Avoid using it directly outside of the file.
"""
def LoadConfig(name : str, reload = True) :

    if name in __loadedFiles :
        if reload :
            Log.Warning(f"Le fichier de config {name} était déjà chargé, possible que des valeurs soient écrasées")
        else :
            return

    filePath = os.path.join(CWD, CONFIG_FILES_PATH, name)
    if not filePath.endswith(".json") : filePath = filePath + ".json"

    EnsureDirectoryExists()
    print(filePath)

    # check if the file exists
    if not os.path.isfile(filePath) :
        # if the file does not exists, add its name to the loaded files anyways, with an empty dictionary
        # it will be created the next time settings are saved
        __loadedFiles[name] = {}
        Log.Verbose("created new config file " + name)
        return True
    
    else :
        # the file exists

        try :
            with open(filePath, "r") as file :
                contents = file.read()
            
            dictionary = json.loads(contents)

            __loadedFiles[name] = dictionary

            Log.Log(Log.Lvl.VERB, "loaded config file " + filePath)
            return True
        
        except Exception as e :
            
            Log.Error(f"Cannot open config file {name} :\n{str(e)}")
            __loadedFiles[name] = {}
            return False

"""
Save all the current settings to their respective files
"""
def SaveSettings() :

    for configName, values in __loadedFiles.items() :
        try :
            path = os.path.join(CONFIG_FILES_PATH, configName)
            path += ".json"
            with open(path, "w") as file :
                file.write(json.dumps(values, indent="\t"))
        except Exception as e :
            Log.Log(Log.Lvl.ERR, f"n'a pas pu enregistrer config {path} : {e}")


def __EnsureSettingExist(configName : str, key : str) :

    if not configName in __loadedFiles :

        if not LoadConfig(configName) :
            Log.Verbose("failed to load config " + configName)
            return False
    
    configSet : dict
    configSet = __loadedFiles[configName]
    
    if not key in configSet : configSet[key] = None

    return True

def CheckSettingExist(configName : str, key : str) :

    if not configName in __loadedFiles : return False
    if not key in __loadedFiles[configName] : return False
    return True


def GetConfigValue(configName : str, key : str) :

    if not __EnsureSettingExist(configName, key) : return None

    return __loadedFiles[configName][key]

"""
Format the config value as a string. Useful for text entries that store their contents in a config file
"""
def GetConfigValueString(configName : str, key : str) :
    val = GetConfigValue(configName, key)
    if val == None :
        return ""
    else :
        return str(val)

def SetConfigValue(configName : str, key : str, value) :

    if not __EnsureSettingExist(configName, key) : return False

    __loadedFiles[configName][key] = value



def VerifyCompleteness(configName : str, defaultDict : dict) :

    LoadConfig(configName)

    currentDic = __loadedFiles[configName]

    for key, value in defaultDict.items() :
        if not key in currentDic :
            currentDic[key] = value
        else :
            if currentDic[key] == None :
                currentDic[key] = value

def GetImageInfos():
    LoadConfig("images.json", reload=False)
    image_infos = __loadedFiles.get("images.json", {})
    # On suppose que le fichier images.json est un dictionnaire { "tag": "chemin", ... }
    return list(image_infos.items())
