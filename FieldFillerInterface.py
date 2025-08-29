import tkinter as tk
from tkinter import StringVar
from tkinter import ttk
import TkinterClasses

import Settings
import Log

import FileFinder

tab : TkinterClasses.Tab

textLT : StringVar
textForm : StringVar

def Init(tabBar : TkinterClasses.TabBar) :

	global tab
	global entryForm, entryLT, entryCategory

	tab = TkinterClasses.Tab(tabBar, "Find Data", OnOpen)


	
	description = tk.Label(tab, text="Remplis automatiquement les champs du programme avec les données ci-dessous",
						   foreground=Settings.GetConfigValueString("theme", "text_color"),
						   background=Settings.GetConfigValueString("theme", "background_color"))
	description.grid(column=1, row=1, columnspan=5, rowspan=1)

	entryLT = TkinterClasses.StorageTextField(tab, "fields", "LT", "LTxxx...xx")
	entryLT.grid(column=1, row=3, columnspan=2, rowspan=1)

	entryForm = TkinterClasses.StorageTextField(tab, "fields", "form", 'Formulaire + Indice\n(format "Fxxx A")')
	entryForm.grid(column=1, row=5, columnspan=2, rowspan=1)

	# éviter de hardcoder ça plus tard
	categoryMenuValues = ["RE", "PV", "RCI"]
	entryCategory = TkinterClasses.StorageDropdown(tab, "fields", "out_file_category", "catégorie fichier", categoryMenuValues)
	entryCategory.grid(column=3, row=3, columnspan=2, rowspan=1)

	fillButton = tk.Button(tab, text="fill", command=AutoFill)
	fillButton.grid(column=1, row=7, columnspan=2, rowspan=1)


	# output text window
	outputText = TkinterClasses.OutputTextWindow(tab)
	outputText.grid(column=3, row=4, rowspan=5, columnspan=10, sticky="nsew")


def OnOpen() :

	pass
	

def AutoFill() :
	err = False
	
	err = err or not AutoFillWorkdirAndMeasurePath()
	err = err or not AutoFillFormPath()

	# now change some field values
	Settings.SetConfigValue("paths", "generator_out_path", entryCategory.GetValue())
	Settings.SetConfigValue("fields", "generator_out_filename", f"**REF_SEDI**_SN**SN**_{entryCategory.GetValue()}")

	TkinterClasses.StorageField.LoadAll()

	if err :
		Log.Error("Le remplissage automatique à rencontré des erreurs")
	else :
		Log.Message("Remplissage automatique réussi")


def AutoFillWorkdirAndMeasurePath() :

	# the LT folder
	workFolder = FileFinder.GetPathToLT(entryLT.GetValue())
	print(workFolder)
	if workFolder == False :
		Log.Error("Le chemin vers le LT n'a pas été trouvé")
		return False
	else :
		# workfolder is a (str, re.Match) tuple
		Settings.SetConfigValue("paths", "root_work_dir", workFolder[0])

		measures = FileFinder.GetMeasurementsFromLT(workFolder[0])

		if measures == None :
			return False
		else :
			Settings.SetConfigValue("paths", "measurement_file", measures)
			return True
	

def AutoFillFormPath() :

	form = entryForm.GetValue()
	split = form.strip().split(" ")

	# normally the form has just one (or more) spaces between the num and the index, so both ends of the list should give us the number and the indice

	formNum = split[0]
	formIndice = split[-1]

	if len(split) < 2 :
		Log.Error(f"format du formulaire incorrect, {form}")
		return False

	# remove the initial F from the formNum
	if formNum.startswith("F") :
		formNum = formNum[1:]
	else :
		Log.Error(f"format du formulaire incorrect, {form}")
		return False
	
	formNum = formNum.strip()
	
	formPath = FileFinder.GetForm(formNum, formIndice)

	if formPath != None :
		Settings.SetConfigValue("paths", "template_file_path", formPath)
		return True
	else :
		Log.Error(f"Le formulaire {form} n'a pas été trouvé dans {FileFinder.formFolder}")
		return False