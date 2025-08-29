import os
from enum import IntEnum


"""
Utility program to manage log levels, access, and file output.
As a guideline, use VERB for messages that might have debug values but are confusing / take too much space / are useless for the user
Use MSG for messages that provides info to the user
Use WARN and ERR as usual
"""


# list that contains every log line since the start of the program.
# sorted in order of addition, and contains tuples (level : int, message : str)
logBuffer = []

DEBUG_TO_PRINT = True

LOG_PATH = "Logs"

class Lvl(IntEnum) :
    SELF = -1
    VERB = 0
    MSG = 1
    WARN = 2
    ERR = 3

levelToString = ("[VERB]", "[MSG]", "[WARN]", "[ERROR]")

# list of callbacks for each log level
logCallbacks = [[] for i in range(4)]


def Init() :

    os.makedirs(LOG_PATH, exist_ok=True)


# main logging function
def Log(level : Lvl, message : str) :

    formattedString = " ".join((">>", levelToString[int(level)], message)) + "\n"
    logBuffer.append((int(level), formattedString))

    if (DEBUG_TO_PRINT) : print(formattedString)

    if level == Lvl.SELF : return

    relevantCallbacks = logCallbacks[int(level)]

    for callback in relevantCallbacks :
        try :
            callback(formattedString)
        except Exception as e :
            Log(Lvl.SELF, "cannot use callback : " + str(e))


def AddCallback(level : Lvl, callback) :
    logCallbacks[level].append(callback)


def CreateLogFile(name : str, levelMin : Lvl, levelMax : Lvl) :

    try :
        if not name.endswith(".log") : name = name + ".log"
        with open(os.path.join(LOG_PATH, name), "w", encoding="utf-8") as file :
            file.writelines([text for lvl, text in filter(lambda item : item[0] <= int(levelMax) and item[0] >= int(levelMin), logBuffer)])
    except Exception as e :
        print(f"error saving logs : {e}")



# /////////////////////////////////////////////// COMPATIBILITY FUNCTIONS //////////////////////////////////////////////////////

def Error(string : str) :
    Log(Lvl.ERR, string)

def Warning(string : str) :
    Log(Lvl.WARN, string)

def Verbose(string : str) :
    Log(Lvl.VERB, string)

def Message(string : str) :
    Log(Lvl.MSG, string)
