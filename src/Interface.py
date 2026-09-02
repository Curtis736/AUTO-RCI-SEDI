#!/usr/bin/env python
# -*- coding: utf-8 -*-

import init_paths  # noqa: F401 — cwd + sys.path avant les autres imports

import sys
import os
import traceback

import tkinter as tk
from tkinter import scrolledtext

import Log
import Settings
import DefaultSettings
import TkinterClasses
import ImageReader

import PathsInterface
import DocumentGenerator
import Selector
import ExcelController


import PdfGenerator
import FieldFillerInterface




quit = False





# Gérer explicitement l'importation de win32timezone
try:
    # Tenter d'importer les modules win32 nécessaires
    import win32timezone
    import pythoncom
    import win32com.client
    import pywintypes
    
    # Initialiser COM ici une fois pour toutes
    pythoncom.CoInitialize()
    print("Modules win32 importés avec succès et COM initialisé")
except ImportError as e:
    print(f"ATTENTION: Problème lors de l'importation des modules win32: {e}")
    print("L'application tentera de fonctionner sans ces modules, mais certaines fonctionnalités pourraient être limitées.")
    # Dans le cas particulier de win32timezone manquant
    if "win32timezone" in str(e):
        try:
            # Solution alternative pour win32timezone
            from win32com.client import pythoncom
            pythoncom.CoInitialize()
            print("Module pythoncom importé comme alternative et COM initialisé")
        except ImportError:
            print("Impossible d'importer pythoncom comme alternative")
        except Exception as e:
            print(f"Erreur avec l'alternative pour win32timezone: {str(e)}")







"""
Excuses d'avance pour le bazar incroyable qu'est ce machin
"""

# finish everything on application close
def OnClose() :
    global quit
    quit = True
    #DocumentGenerator.OnQuit()
    root.destroy()
    
    # Libérer COM avant de quitter
    try:
        pythoncom.CoUninitialize()
        print("COM uninitialize successful")
    except:
        pass
    
    Settings.SaveSettings()
    ExcelController.CloseExcel()
    Log.CreateLogFile("problems", Log.Lvl.WARN, Log.Lvl.ERR)
    Log.CreateLogFile("normal", Log.Lvl.MSG, Log.Lvl.ERR)
    Log.CreateLogFile("verbose", Log.Lvl.VERB, Log.Lvl.ERR)
    sys.exit()




try :
    # Message d'initialisation (sans callback pour le moment)
    print("\n=== DÉMARRAGE DE L'APPLICATION AUTO RCI ===")
    print("Initialisation de l'interface utilisateur...")

    # init les logs
    Log.Init()

    # génère valeurs par défaut des config si nécessaire
    Settings.LoadConfig("theme")
    Settings.LoadConfig("fields")

    DefaultSettings.VerifySettingsCompleteness()

    ImageReader.Init()

    # Créer la fenêtre principale
    root = tk.Tk()
    root.title("Auto RCI")
    root.protocol("WM_DELETE_WINDOW", OnClose)

    # Configurer la taille de la fenêtre : environ 70% de l'écran
    screenWidth = root.winfo_screenwidth()
    screenHeight = root.winfo_screenheight()
    window_width = int(screenWidth * 0.7)
    window_height = int(screenHeight * 0.7)  # Augmenter à 70% de la hauteur
    root.geometry(f"{window_width}x{window_height}+{int(screenWidth * 0.15)}+{int(screenHeight * 0.15)}")
    
    # Configurer les couleurs de l'interface
    root.configure(background=Settings.GetConfigValueString("theme", "background_color"))


    # créer l'espace où le contenu des onglets sera affiché
    tabDisplay = tk.Frame(root, background="red")
    tabDisplay.place(relx=0.025, y = 30, relwidth=0.66, relheight=0.9)
    #tabDisplay.grid(row=1, column=0, rowspan=9, columnspan=7, sticky="nsew")

    # Créer l'espace pour les onglets
    tabBar = TkinterClasses.TabBar(root, tabDisplay)
    tabBar.place(relx=0.025, y = 0, relwidth=0.95, height=30)

    # Affichage du répertoire
    directoryDisplay = scrolledtext.ScrolledText(root)
    directoryDisplay.bind("<Key>", lambda e: "break")
    directoryDisplay.config(wrap="none")
    directoryDisplay.configure(background=Settings.GetConfigValueString("theme", "entry_color"), foreground=Settings.GetConfigValueString("theme", "text_color"))
    directoryDisplay.place(relx=0.72, y=30, relwidth=0.26, relheight=0.9)
    directoryDisplay.tag_config("folder", foreground=Settings.GetConfigValueString("theme", "question_color"))
    directoryDisplay.tag_config("file", foreground="chartreuse3")
    directoryDisplay.tag_config("undefined", foreground=Settings.GetConfigValueString("theme", "error_color"))

    # Initialiser les références globales
    print("Initialisation des références globales...")
    Selector.root = root
    DocumentGenerator.root = root
    TkinterClasses.LoadingIcon.root = root
   


    # Initialiser l'interface avec les chemins
    PathsInterface.directoryDisplay = directoryDisplay
    
    # Initialiser les onglets
    print("Initialisation des onglets...")
    FieldFillerInterface.Init(tabBar)
    PathsInterface.Init(tabBar)
    DocumentGenerator.Init(tabBar)
    PdfGenerator.Init(tabBar)
    
    tabBar.tabs[0].Open()


    # Afficher un message sur les modifications récentes
    Log.Message("\n=== BIENVENUE DANS AUTO_RCI SIMPLIFIÉ ===")
    Log.Message("Version du 26/03/2024")
    Log.Message("\nModifications récentes:")
    Log.Message("1. Amélioration de la gestion des répertoires de sortie")
    Log.Message("2. Les documents sont maintenant générés même si le chemin de sortie n'existe pas")
    Log.Message("3. Correction du problème de création de documents dans les dossiers cibles")
    Log.Message("4. Les champs 'numéro de lancement (LT)' et 'numéro de plan' sont désormais préremplis avec des valeurs par défaut")
    Log.Message("5. Simplification de l'interface en supprimant les onglets non essentiels")
    Log.Message("6. Meilleure gestion des erreurs lors de la génération des documents")
    Log.Message("\nPour toute question ou signalement de bug, veuillez contacter le support technique.")
    Log.Message("=============================================\n")

    # S'assurer qu'au moins un onglet est ouvert
    Log.Message("Ouverture de l'onglet par défaut...")
    
    #windowList[0].Open()
    
    Log.Message("Initialisation terminée. Application prête à l'utilisation.")

    # Boucle principale
    while not quit:
        root.update_idletasks()
        root.update()

except Exception as e:
    error_message = traceback.format_exc()
    print(error_message)
    print(f"Erreur d'initialisation : {str(e)}")

    
    
    # Essayer d'afficher une boîte de dialogue d'erreur si possible
    try:
        import tkinter.messagebox as messagebox
        messagebox.showerror("Erreur critique", 
                            f"L'application a rencontré une erreur critique et doit être fermée.\n\n"
                            f"Erreur: {str(e)}\n\n"
                            f"Vérifiez les journaux pour plus de détails.")
    except Exception as dialog_err:
        print(f"Impossible d'afficher la boîte de dialogue d'erreur: {str(dialog_err)}")
    
    # Essayer de fermer proprement les ressources
    try:
        # Libérer COM avant de quitter
        try:
            pythoncom.CoUninitialize()
            print("COM uninitialize successful")
        except:
            pass
        
        ExcelController.CloseExcel()
        Log.Message("Ressources Excel libérées")
    except Exception as close_err:
        print(f"Erreur lors de la fermeture des ressources: {str(close_err)}")
    
    try :
        if "root" in globals() and root.winfo_exists():
            OnClose()
        else:
            Settings.SaveSettings()
            ExcelController.CloseExcel()
    except Exception as e :
        print(e)
        
    # Quitter sans utiliser input() qui peut causer 'lost sys.stdin'
    sys.exit(1)