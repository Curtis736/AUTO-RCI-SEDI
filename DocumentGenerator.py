import tkinter as tk
from tkinter import StringVar
from tkinter import IntVar
from tkinter import filedialog as fd
from tkinter import scrolledtext
import os
import re

import TkinterClasses
import Settings
import Log
import ValueFetcher
import Writer
import ExcelController

# Fonction utilitaire pour s'assurer qu'un répertoire existe
def ensure_directory_exists(directory_path):
	"""S'assure qu'un répertoire existe et le crée si nécessaire.
	
	Args:
		directory_path: Chemin du répertoire à vérifier/créer
		
	Returns:
		bool: True si le répertoire existe ou a été créé avec succès, False sinon
	"""
	if not directory_path:
		Log.Error("Chemin de répertoire non spécifié")
		return False
	
	# Si le répertoire existe déjà, rien à faire
	if os.path.exists(directory_path) and os.path.isdir(directory_path):
		Log.Message(f"Le répertoire existe déjà: {directory_path}")
		return True
		
	# Créer le répertoire
	try:
		os.makedirs(directory_path, exist_ok=True)
		Log.Message(f"Répertoire créé avec succès: {directory_path}")
		
		# Vérifier que le répertoire a bien été créé
		if os.path.exists(directory_path) and os.path.isdir(directory_path):
			return True
		else:
			Log.Error(f"Le répertoire n'a pas pu être créé malgré l'absence d'erreur: {directory_path}")
			return False
	except Exception as make_err:
		Log.Error(f"Erreur lors de la création du répertoire: {str(make_err)}")
		
		# Essayer un chemin absolu simple comme solution de repli
		try:
			fallback_dir = os.path.join(os.getcwd(), "output")
			Log.Message(f"Utilisation d'un répertoire de secours: {fallback_dir}")
			os.makedirs(fallback_dir, exist_ok=True)
			return ensure_directory_exists(fallback_dir)
		except Exception as fallback_err:
			Log.Error(f"Impossible de créer le répertoire de secours: {str(fallback_err)}")
			return False




# ////////////////////////////////////// DEFINITION UI ////////////////////////////////////////////////////////////


# callbacks qui servent à intéragir avec la fenêtre de text principale

running = False

def ErrorCallback(string : str) :
	global textWidget
	textWidget.ErrorCallback(string)

def MessageCallback(string : str) :
	global textWidget
	textWidget.MessageCallback(string)

def AskQuestion(string) :
	global answer
	global textWidget
	answer = None
	textWidget.QuestionCallback(string)
	global blocked
	blocked = True
	while answer == None :
		root.update_idletasks()
		root.update()
	print(answer)
	if answer :
		Log.Message("-- OUI")
	else :
		Log.Message("-- NON")
	temp = answer
	print(temp)
	answer = None
	print(temp)
	blocked = False
	return temp

answer = None
root = None
textWidget = None  # Déclarer textWidget comme variable globale

indices = []

def Init(tabBar : TkinterClasses.TabBar) :

	global tab
	global window
	global textWidget
	global outputFileName, relOutDir

	tab = TkinterClasses.Tab(tabBar, "Générateur fiches RCI/RE", None)
	
	
	# Format du nom de fichier
	outputFileName = TkinterClasses.StorageTextField(tab, "fields", "generator_out_filename", "format du nom\nde fichier")
	outputFileName.grid(column=0, row=0, columnspan=2, sticky="we", padx=2)

	# Chemin de sortie
	"""Relative to the root path"""
	relOutDir = TkinterClasses.StorageTextField(tab, "paths", "generator_out_path", "chemin de sortie\n(relatif au dossier de lancement)")
	relOutDir.grid(column=0, row=1, columnspan=2, sticky="we", padx=2)
	

	# Zone de texte pour les messages
	"""
	global textWidget
	textWidget = scrolledtext.ScrolledText(tab, wrap=tk.WORD)
	textWidget.configure(background=Settings.GetConfigValueString("theme", "entry_color"), foreground=Settings.GetConfigValueString("theme", "text_color"))
	textWidget.bind("<Key>", lambda e: "break")
	textWidget.grid(column=3, row=0, columnspan=2, rowspan=7, sticky="nesw")

	textWidget.tag_config("question", foreground=Settings.GetConfigValueString("theme", "question_color"))
	textWidget.tag_config("error", foreground=Settings.GetConfigValueString("theme", "error_color"))
	"""

	
	textWidget = TkinterClasses.OutputTextWindow(tab)
	textWidget.grid(column=3, row=0, columnspan=15, rowspan=7, sticky="nesw")
	



	# Options pour les fichiers préexistants
	global duplicatesHandlingMenu
	duplicatesHandlingMenuValues = ["conserver l'original", "dupliquer", "remplacer", "demander"]
	duplicatesHandlingMenu = TkinterClasses.StorageDropdown(tab, "fields", "file_reaction_mode", "réaction aux fichiers préexistants", duplicatesHandlingMenuValues)
	duplicatesHandlingMenu.grid(column=1, row=2)



	# Icône de chargement
	global loadingIcon
	loadingIcon = TkinterClasses.LoadingIcon(tab)
	loadingIcon.SetText("---")
	loadingIcon.grid(column=1, row=4, sticky="we", columnspan=2, rowspan=1)

	# Bouton de génération
	GenOneDoc = tk.Button(tab, text="générer document", command=GenDoc, 
						  background=Settings.GetConfigValueString("theme", "unselected_color"), 
						  foreground=Settings.GetConfigValueString("theme", "text_color"))
	GenOneDoc.grid(column=1, row=5)

	# Boutons Oui/Non
	yes = tk.Button(tab, text="oui", command=Yes,
					background=Settings.GetConfigValueString("theme", "unselected_color"), 
					foreground=Settings.GetConfigValueString("theme", "text_color"))
	yes.grid(column=3, row=8, sticky="")

	no = tk.Button(tab, text="non", command=No,
					background=Settings.GetConfigValueString("theme", "unselected_color"), 
					foreground=Settings.GetConfigValueString("theme", "text_color"))
	no.grid(column=4, row=8, sticky="")

	# Initialisation du collecteur de valeurs
	global values
	values = ValueFetcher.ValueCollector()

	
	"""
	# Ajouter un label explicatif pour informer l'utilisateur des valeurs utilisées
	infoLabel = tk.Label(tab, text=f"Valeurs par défaut:\nLancement: {launchText.get()} | Plan: {numPlanText.get()}", 
						 background=Settings.backgroundColor, foreground=Settings.textColor)
	infoLabel.grid(column=0, row=8, columnspan=2, sticky="we", padx=2, pady=2)
	"""




# ///////////////////////////////////////////////////// END DEFINITION UI /////////////////////////////////////////////////////




def DisplayPaths():
	"""Affiche tous les chemins importants utilisés par l'application"""
	Log.Message("\n=== CHEMINS UTILISÉS ===")
	Log.Message(f"Dossier racine: {Settings.GetConfigValueString('paths', 'root_work_dir')}")
	Log.Message(f"Dossier de sortie RCI: {Settings.GetConfigValueString('paths', 'root_work_dir')}/{Settings.GetConfigValueString('paths', 'generator_out_path')}")
	Log.Message(f"Fichier Excel des mesures: {Settings.GetConfigValueString('paths', 'measurement_file')}")
	Log.Message(f"Template Word: {Settings.GetConfigValueString('paths', 'template_file_path')}")
	Log.Message("\n=== DOSSIERS D'IMAGES ===")
	# Note: imageInfos n'existe plus dans la nouvelle structure
	Log.Message("Configuration des images à vérifier")
	Log.Message("====================\n")





"""
Function called when the user requests the generation of a document
"""
def GenDoc() :
	"""
	Génère un document pour un SN spécifique ou tous les SN du LT actuel
	"""
	global running
	if running :
		return
	running = True
	tab.SetTabLock(True)
	loadingIcon.isloading = True
	loadingIcon.SetText("en cours")

	
	# Synchronisation automatique du LT depuis le chemin si vide
	if Settings.GetConfigValue("fields", "LT") == None :
		try:
			from PathsInterface import extract_lt_from_path
			lt_from_path = extract_lt_from_path(Settings.GetConfigValueString("paths", "root_work_dir"))
			if lt_from_path:
				Settings.SetConfigValue("fields", "LT", lt_from_path)
				ValueFetcher.valueCollector.LT = lt_from_path
				values.LT = lt_from_path
				Log.Message(f"[DEBUG] LT synchronisé automatiquement depuis le chemin : {lt_from_path}")
			else:
				Log.Warning("Impossible d'extraire le LT du chemin de travail.")
		except Exception as e:
			Log.Warning(f"Erreur lors de la synchronisation automatique du LT : {e}")

	# Les valeurs de lancement et plan sont déjà définies par défaut
	# Plus besoin de vérifier si les champs sont remplis
	
	# Mettre à jour les valeurs
	values.LT = Settings.GetConfigValueString("fields", "LT")
	values.numPlan = Settings.GetConfigValueString("fields", "num_plan")
	
	# Afficher le chemin de sortie avant la génération
	output_dir = os.path.abspath(os.path.join(Settings.GetConfigValueString("paths", "root_work_dir"), Settings.GetConfigValueString("paths", "generator_out_path")))
	Log.Message(f"\n=== DOSSIER DE SORTIE DES DOCUMENTS ===")
	Log.Message(f"Les documents seront générés dans le dossier :")
	Log.Message(f"{output_dir}")
	Log.Message(f"==========================================\n")
	
	
	# Demander à l'utilisateur de choisir une option
	question_avec_options = "Choisissez une option :\n\n"
	question_avec_options += "OPTION 1 = GÉNÉRER TOUS LES DOCUMENTS DU LT\n"
	question_avec_options += "→ Génère automatiquement un document pour chaque\n"
	question_avec_options += f"   numéro de série associé au LT {values.LT}\n\n"
	question_avec_options += "OPTION 2 = SÉLECTIONNER DES NUMÉROS DE SÉRIE SPÉCIFIQUES\n"
	question_avec_options += f"→ Vous permet de choisir un ou plusieurs numéros de série\n"
	question_avec_options += f"   parmi ceux disponibles pour le LT {values.LT}\n\n"
	question_avec_options += "Entrez le numéro de votre choix (1 ou 2) :"
	
	option_index = AskOptionIndex(question_avec_options, 2)
	
	if option_index == 1:
		# Option 1: Générer tous les SN du LT actuel
		Log.Message(f"Option choisie : Génération de tous les SN du LT {values.LT}")
		# Lire le fichier Excel et filtrer par LT
		if not values.FetchValuesInExcel():
			Log.Error("Échec de la lecture du fichier Excel")
			return False
		
		# Récupérer tous les SN du LT actuel
		lt_sns = values.FindSNFromLT(values.LT)
		if not lt_sns:
			Log.Error(f"Aucun SN trouvé pour le LT {values.LT}")
			return False
		
		Log.Message(f"Génération automatique des documents pour {len(lt_sns)} SN du LT {values.LT}")
		# Appeler GenerateSingleDocument avec la liste des SN du LT
		GenerateSingleDocument(lt_sns)
	elif option_index == 2:
		# Option 2: Sélectionner des SN spécifiques
		Log.Message("Option choisie : Sélection de SN spécifiques")
		selected_sns = SelectSingleSN()
		if selected_sns:
			Log.Message(f"Génération des documents pour les SN sélectionnés : {', '.join(selected_sns)}")
			GenerateSingleDocument(selected_sns)
		else:
			Log.Message("Aucun SN sélectionné, opération annulée")
	else:
		Log.Error("Option non valide")
		return False
	
	
	#output_dir = os.path.abspath(os.path.join(Settings.paths[Settings.PATH_ROOT], Settings.paths[Settings.PATH_RCI]))
	Log.Message(f"\n=== EMPLACEMENT DES DOCUMENTS GÉNÉRÉS ===")
	Log.Message(f"Les documents ont été générés dans le dossier :")
	Log.Message(f"{output_dir}")
	
	# Demander à l'utilisateur s'il souhaite ouvrir le dossier
	open_folder = AskQuestion("Voulez-vous ouvrir le dossier contenant les documents générés?")
	if open_folder:
		try:
			# Utiliser la commande appropriée selon le système d'exploitation
			import platform
			if platform.system() == "Windows":
				os.startfile(output_dir)
			elif platform.system() == "Darwin":  # macOS
				import subprocess
				subprocess.Popen(["open", output_dir])
			else:  # Linux et autres
				import subprocess
				subprocess.Popen(["xdg-open", output_dir])
			Log.Message("Ouverture du dossier de sortie...")
		except Exception as e:
			Log.Error(f"Impossible d'ouvrir le dossier : {str(e)}")
	
	Log.Message(f"===================================\n")
	

	ExcelController.CloseExcel()
	loadingIcon.isloading = False
	loadingIcon.SetText("")
	tab.SetTabLock(False)
	Log.CreateLogFile("verbose", Log.Lvl.VERB, Log.Lvl.ERR)
	running = False



def SelectSingleSN():
	"""
	Affiche une fenêtre permettant à l'utilisateur de sélectionner des SNs spécifiques
	Seuls les SNs associés au LT actuel sont affichés.
	"""
	# Lire le fichier Excel pour obtenir la liste des SN disponibles
	try:
		# Vérifier d'abord si les valeurs nécessaires sont définies
		if not values.LT or not values.numPlan:
			Log.Error("Le numéro de lancement ou le numéro de plan est manquant")
			Log.Message("Utilisation des valeurs par défaut pour le numéro de lancement et le numéro de plan")
			# Continuer avec les valeurs par défaut
			
		Log.Message("Tentative de lecture du fichier Excel...")
		if not values.FetchValuesInExcel():
			raise Exception("Échec de la lecture du fichier Excel")
		
		# Filtrer les SN pour ne garder que ceux associés au LT actuel
		lt_sns = values.FindSNFromLT(values.LT)
		if not lt_sns:
			Log.Error(f"Aucun SN trouvé pour le LT {values.LT}")
			Log.Message("Vérifiez que le LT est correct et que des données existent dans le fichier Excel")
			return None
		
		# Afficher une fenêtre de sélection avec seulement les SN du LT
		Log.Message(f"Veuillez sélectionner les SNs parmi les {len(lt_sns)} disponibles pour le LT {values.LT}")
		
		# Utiliser la fonction GetSelection existante pour afficher une liste de sélection
		# Sans limiter le nombre de sélections (max_selections=None)
		selected_indices = GetSelection(lt_sns, max_selections=None)
		
		if not selected_indices:
			Log.Message("Aucun SN sélectionné")
			return None
		
		# Récupérer tous les SNs sélectionnés
		selected_sns = [lt_sns[i] for i in selected_indices]
		Log.Message(f"SNs sélectionnés ({len(selected_sns)}) : {', '.join(selected_sns)}")
		return selected_sns
		
	except Exception as e:
		Log.Error(f"Erreur lors de la sélection des SNs : {str(e)}")
		
		Log.Message("Opération annulée")
		return None



def GenerateSingleDocument(selected_sns=None) :
	"""
	Génère un document pour chaque SN sélectionné
	selected_sns: peut être un seul SN (str) ou une liste de SNs
	"""

	workDir = Settings.GetConfigValueString("paths", "root_work_dir")

	# Convertir selected_sns en liste s'il ne l'est pas déjà
	if selected_sns and not isinstance(selected_sns, list):
		selected_sns = [selected_sns]
		
	# //////////////////////////// VERIFY USER INPUT ///////////////////////////////////////////

	# Afficher un message différent selon que des SNs spécifiques ont été sélectionnés ou non
	if selected_sns:
		Log.Message(f"Génération des documents pour les SNs ({len(selected_sns)}) : {', '.join(selected_sns)}")
	else:
		Log.Message(f"Génération des documents pour tous les SNs du LT {values.LT}")

	# do not start unless the way to handle multiple documents is defined
	if duplicatesHandlingMenu.GetValue() not in duplicatesHandlingMenu.GetList() :
		Log.Error("Veuillez choisir un mode de réaction aux documents préexistants")
		return False
	
	# verify the output filename template has at least the SN number in it
	if not "**SN**" in outputFileName.GetValue() :
		Log.Error("le nom final du fichier doit au moins contenir le numéro SN. Ajoutez **SN** au nom du fichier")
		return False
	
	# A VERIFIER
	outputFileName.Lock()
	#fileNameFormatFreezed = fileNameFormat.get()

	# ///////////////////////////////////////////////////// END VERIFY USER INPUT ////////////////////////////////////////////



	# A VERIFIER
	"""
	# this is necessary for compatibility with the rest of the program
	if not outputFolderPath.get().endswith("/") :
		outputFolderPath.set(outputFolderPath.get() + "/")

	Settings.paths[Settings.PATH_RCI] = outputFolderPath.get()
	"""
	
	# Vérifier et créer le répertoire de sortie principal dès le début
	output_directory = os.path.join(Settings.GetConfigValueString("paths", "root_work_dir"), Settings.GetConfigValueString("paths", "generator_out_path"))
	Log.Message(f"Vérification du répertoire de sortie : {output_directory}")
	

	if not ensure_directory_exists(output_directory):
		Log.Error("Impossible de créer ou d'accéder au répertoire de sortie principal")
		Log.Error(f"Chemin: {output_directory}")
		Log.Message("Vérifiez les permissions et que le chemin spécifié est valide")
		return False
	

	# //////////////////////////////////// START VALUES COLLECTION ///////////////////////////////////////////////
	
	# d'abord on essaye de ramasser toutes les infos sur le document
	values.LT = Settings.GetConfigValueString("fields", "LT")
	values.numPlan = Settings.GetConfigValueString("fields", "num_plan")


	if not values.GetFormulaireIndice(os.path.basename(Settings.GetConfigValueString("paths", "template_file"))) :
		Log.Error("n'a pas trouvé l'indice du formulaire")
	
	if not values.GetFormulaireNumber(os.path.basename(Settings.GetConfigValueString("paths", "template_file"))) :
		Log.Message("n'a pas trouvé le numéro du formulaire")
	

	try:
		# Tentative de lecture du fichier Excel
		Log.Message("\n=== LECTURE COMPLÈTE DU FICHIER EXCEL ===")
		Log.Message(f"Lecture de toutes les valeurs du fichier Excel: {Settings.GetConfigValueString('paths', 'measurement_file')}")
		if not values.FetchValuesInExcel():
			raise Exception("Échec de la lecture du fichier Excel")
		allContainers = values.containers
		if not allContainers:
			raise Exception("Aucune donnée n'a pu être récupérée du fichier Excel")
		Log.Message(f"Lecture du fichier Excel réussie. {len(allContainers)} SN trouvés.")
		#Log.Message(f"Toutes les valeurs du fichier Excel ont été lues et seront utilisées dans les documents générés.")
		Log.Message("=====================================\n")
	except Exception as e:
		Log.Error(f"Erreur lors de la lecture du fichier Excel: {str(e)}")
		
	
	# Filtrer les containers si un SN spécifique a été sélectionné
	if selected_sns:
		# Filtrer les containers pour les SN sélectionnés
		containers = [container for container in allContainers if container.SN in selected_sns]
		if not containers :
			Log.Error(f"tous les SNs sélectionnés n'ont pas pu être trouvés dans le fichier excel")
			return False
		Log.Message(f"Génération du document pour les SNs sélectionnés : {', '.join(selected_sns)}")
	else:
		# Filtrer les containers pour ne garder que ceux du LT actuel
		lt_sns = values.FindSNFromLT(values.LT)
		if not lt_sns:
			Log.Error(f"Aucun SN trouvé pour le LT {values.LT}")
			return False
		
		containers = [container for container in allContainers if container.SN in lt_sns]
		Log.Message(f"Génération automatique des documents pour {len(containers)} SN du LT {values.LT}")
		
		# Nous n'allons plus créer des versions E et SE séparées
		# Nous allons plutôt traiter les deux versions dans le même document
	Log.Message(f"Chaque document contiendra les versions E et SE du même SN")
	
	# CONTOURNEMENT TEMPORAIRE: Désactiver la recherche de plan
	"""
	if not values.GetPlan() :
		Log.Error("n'a pas trouvé de fichier pour le plan, abandon")
		return False
	
	if not ImageReader.SaveInTemp(values.pathToPlan) :
		Log.Error("n'a pas pu enregistrer l'image du plan dans temp, fin")
		return False
	
	pathToCachedPlan = ImageReader.GetPathToTempImage(values.pathToPlan)
	"""
	pathToCachedPlan = None
	
	values.UpdateDate()



	# /////////////////////////////////// END VALUES COLLECTION //////////////////////////////////////////////::



	# ////////////////////////////////// START IMAGE VERIFICATION ///////////////////////////////////////

	"""
	i = 0
	for tag, path in Settings.imageInfos :
		# if the path to the image folder exists, we verify it is correct
		if os.path.exists(workDir + "/" + path) :
			result = ImageReader.VerifyImageFilesIn(workDir + "/" + path, i, TkinterClasses.LoadingIcon.UpdateAll)
			if not result :
				Log.Error("[FATAL] quelque chose ne va pas dans les fichiers du dossier " + path)
				return False
		i += 1
	
	try :
		with open("foundImages.txt", "w") as f :
			f.write(str(ImageReader.imagesForSn))
	except :
		pass

	# Informer l'utilisateur qu'il va devoir sélectionner des images
	#Log.Message("Les images associées à chaque SN seront automatiquement sélectionnées et insérées dans le document.")
	#Log.Message("Les images sont filtrées par type (avec étiquette et sans étiquette).")
	"""

	# ////////////////////////////////// END IMAGE VERIFICATION //////////////////////////////////////


	errorCount = 0
	generated_files = []  # Liste pour stocker les chemins des fichiers générés


	# ///////////////////////////////// START DOCUMENT GENERATION ///////////////////////////////////////


	for container in containers :
		# each container object is a DocumentValues container
		# we create a new word for each of them
		Log.Message("------- DEBUT FICHE SN" + container.SN + " -------\n")

		errorHappened = False

		if not container.TryVerifyAllSpecs() :
			errorHappened = True
			Log.Error("une mesure ou plus n'est pas dans les specs pour " + container.SN)

		"""
		if len(container.undefinedValues) > 0 :
			errorHappened = True
			# Convertir la liste undefinedValues en chaîne de caractères
			undefined_values_str = ", ".join(container.undefinedValues) if isinstance(container.undefinedValues, list) else str(container.undefinedValues)
			Log.Error("il manque les mesures " + undefined_values_str + " au document " + container.SN)
		"""

		if not Writer.Open(Settings.GetConfigValueString("paths", "template_file_path")) :
			Log.Error("ne peut pas ouvrir le document de base")
			return

		# we get all the tags to replace and what they should be replaced with
		values.SetWriterTagLists(container)

		# Vérifier si le document est correctement initialisé
		if not Writer.documentDefined:
			Log.Error("Le document n'a pas été correctement initialisé")
			errorHappened = True
			continue

		# Traitement des tags d'images personnalisés (§§FOLDER_NAME§§)
		Log.Message("\n=== TRAITEMENT DES TAGS D'IMAGES PERSONNALISÉS ===")
		Log.Message(f"Recherche de tags d'images personnalisés pour SN{container.SN}")
		
		
		
		Log.Message("=== FIN DU TRAITEMENT DES TAGS PERSONNALISÉS ===\n")
		
		# ////////////////////////////////////// START IMAGES /////////////////////////////////////////////////////////////////

		imagesAdded = []
		# ESPACE TAGS PERSONNALISES

		# Utiliser la nouvelle fonction modulaire
		from CustomImageTagProcessor import process_custom_image_tags
		process_custom_image_tags(Writer.document, workDir, container)
		
		# ////////////////////////////////////// END IMAGES /////////////////////////////////////////////////////////////////
		
		# ///////////////////////////////////// START WRITING TEXT ///////////////////////////////////////////////////////////

		# then we ask the writer to replace them
		# Writer.ReplaceAllTags(pulseCallback=TkinterClasses.LoadingIcon.UpdateAll)
		# Utiliser la nouvelle fonction améliorée pour remplacer tous les tags
		Log.Message("Utilisation du remplaceur de tags amélioré pour garantir que tous les tags soient remplacés")
		result = Writer.ReplaceAllTextTags(pulseCallback=TkinterClasses.LoadingIcon.UpdateAll)
		if result == False :
			Log.Error("document is not open")
		elif result == True :
			pass # success
		else :
			# means the result is a list and some tags were not replaced
			errorHappened = True
			errorCount += 1
			Log.Error(f"Certains tags dans le documents n'ont pu être associés à aucune valeur : {result}")

		# now we fill the tables appropriately
		#Writer.FillAllColumns()

		# ////////////////////////////////////// END WRITING TEXT //////////////////////////////////////////////////////////////



		# Marquer le document comme incomplet si :
		# - Il y a des erreurs
		# - Il manque des valeurs
		# - Il manque des images requises
		# - Il y a des tags non remplacés
		isIncomplete = False
		incompleteReason = []
		
		if errorHappened:
			isIncomplete = True
			incompleteReason.append("Erreurs pendant la génération")


		# Créer le nom de fichier de base
		fileName = values.CreateNewFilenameFromString(outputFileName.GetValue())
		
		# Ajouter un marqueur d'incomplétude si nécessaire, mais limiter la longueur
		if isIncomplete:
			# Ajouter simplement "[INCOMPLET]" sans les détails pour éviter des noms trop longs
			fileName += " [INCOMPLET]"
			# Journaliser les raisons détaillées sans les inclure dans le nom du fichier
			Log.Error(f"Document marqué comme incomplet pour SN{container.SN}")
			for reason in incompleteReason:
				Log.Error(f"Raison : {reason}")

		filePath = os.path.join(workDir, Settings.GetConfigValueString("paths", "generator_out_path"), fileName + ".docx")
		
		# Vérifier si le répertoire de destination existe, le créer si nécessaire
		outDir = os.path.dirname(filePath)
		if not ensure_directory_exists(outDir):
			# Si la création du répertoire a échoué, on le signale et on passe au SN suivant
			Log.Error(f"Impossible de créer ou d'accéder au répertoire de destination pour SN{container.SN}")
			Log.Error(f"Chemin: {outDir}")
			errorHappened = True
			errorCount += 1
			continue  # Passer au SN suivant
			
		# if the file already exists we solve this differently depending on the option selected
		if os.path.exists(filePath) :
			if duplicatesHandlingMenu.GetValue() == duplicatesHandlingMenu.GetList()[0] :
				pass

			elif duplicatesHandlingMenu.GetValue() == duplicatesHandlingMenu.GetList()[1] :
				filePath = filePath.removesuffix(".docx")
				i = 2
				# try different numbers at the end until the name is free
				while i < 100 :
					if not os.path.exists(filePath + " " + str(i) + ".docx") :
						filePath = filePath + " " + str(i) + ".docx"
						result = Writer.Save(filePath)
						errorHappened = errorHappened or not result
						if result:
							generated_files.append(filePath)
						break
					i += 1

			elif duplicatesHandlingMenu.GetValue() == duplicatesHandlingMenu.GetList()[2] :
				result = Writer.Save(filePath)
				errorHappened = errorHappened or not result
				if result:
					generated_files.append(filePath)

			elif duplicatesHandlingMenu.GetValue() == duplicatesHandlingMenu.GetList()[3] :
				answer = AskQuestion("Le fichier " + fileName + " existe déjà\nVoulez-vous le remplacer ?\n")
				if answer :
					result = Writer.Save(filePath)
					errorHappened = errorHappened or not result
					if result:
						generated_files.append(filePath)
				else :
					pass
		else :
			result = Writer.Save(filePath)
			errorHappened = errorHappened or not result
			if result:
				generated_files.append(filePath)

		if errorHappened :
			Log.Error("fichier " + os.path.basename(filePath) + " contient des erreurs")
		else :
			Log.Message("fichier " + os.path.basename(filePath) + " traité avec succès")
			Log.Message(f"Fichier généré : {filePath}")
		string = "contient des images pour : "
		for item in imagesAdded :
			string += item + " "
		Log.Message(string)

		if errorHappened :
			errorCount += 1

		TkinterClasses.LoadingIcon.UpdateAll()

		Log.Message("------- FIN FICHE SN" + container.SN + " -------\n")


		root.update()

	# Afficher un résumé des fichiers générés
	if generated_files:
		Log.Message("\n=== RÉSUMÉ DES FICHIERS GÉNÉRÉS ===")
		for file_path in generated_files:
			Log.Message(f"- {os.path.basename(file_path)}")
		Log.Message(f"Nombre total de fichiers générés : {len(generated_files)}")
	else:
		Log.Message("\nAucun fichier n'a été généré.")

	Log.Message("Tous les documents terminés. Compte d'erreurs : " + str(errorCount))
	Log.CreateLogFile("problems", Log.Lvl.WARN, Log.Lvl.ERR)
	Log.CreateLogFile("normal", Log.Lvl.MSG, Log.Lvl.ERR)
	return True

def AddImageInDoc(workDir : str, pathId : int, tag : str, name : str, insertionList : list, SN : str, endType : list, titles : list) :
	"""
	Find the images for pathId, and use name for messages.
	Use this instead of directly using AddPicToDoc
	"""
	print("addimagetodoc ran")

	# Nous allons traiter les deux types (E et SE) dans le même document
	# Nous n'avons plus besoin d'extraire le type du SN car nous allons traiter les deux types
	
	# Message pour indiquer le traitement du document
	Log.Message(f"Traitement du document pour SN{SN} (types E et SE)")

	# Note: Settings.imageInfos n'existe plus dans la nouvelle structure
	# Pour l'instant, on utilise un chemin par défaut
	default_image_path = os.path.join(workDir, "images")
	
	# Vérifier si le chemin existe
	if not os.path.exists(default_image_path) :
		Log.Error(f"le chemin {default_image_path} n'existe pas, impossible d'ajouter l'image demandée")
		return False
	
	Log.Verbose(name + " images needed")
	
	# Utiliser AddPicToDoc pour les deux types (E et SE)
	path = default_image_path
	
	# Créer des titres descriptifs mais courts pour éviter les débordements
	e_titles = ["Image E 1", "Image E 2", "Image E 3", "Image E 4", "Image E 5"]
	se_titles = ["Image SE 1", "Image SE 2", "Image SE 3", "Image SE 4", "Image SE 5"]
	
	# Appeler AddPicToDoc pour le type E
	result_E = AddPicToDoc(path, pathId, SN, tag + "_E", SN, endType, e_titles, f"Sélection d'images pour {tag} - SN{SN}")
	
	# Appeler AddPicToDoc pour le type SE
	result_SE = AddPicToDoc(path, pathId, SN, tag + "_SE", SN, endType, se_titles, f"Sélection d'images pour {tag} - SN{SN}")
	
	# Si au moins un des deux types a été ajouté avec succès, considérer comme un succès
	if result_E or result_SE:
		Log.Verbose(name + " added successfully")
		insertionList.append(name)
		return True
	
	return True

def UpdateIndices() :
	if not Writer.Open(Settings.GetConfigValueString("paths", "template_file_path")) :
			Log.Error("ne peut pas ouvrir le document de base")
			return
	
	global indices
	# find all indices to fill in the document
	# Note: Settings.TAG_INDICE n'existe plus, utiliser une valeur par défaut
	indices = Writer.FindAllIndices(Writer.FormatTagForTextReplacement("**INDICE**"))
	IndiceCollectionString()

def IndiceCollectionString() :
	print("|" * (len(indices) - 1))
	# Note: indicesFieldText n'est pas défini, commenter cette ligne
	# indicesFieldText.set("|" * (len(indices) - 1))
	return "|" * (len(indices) - 1)

def IndicesFromString(string : str) :
	allIndices = string.split("|")
	if len(allIndices) != len(indices) :
		Log.Error("la longueur de la chaîne de caractère des indices est incorrect. Devrait être " + str(len(indices)))
		return
	
	return allIndices

def IsFileMatchingSN(file_name, SN):
	"""
	Fonction générique pour vérifier si un nom de fichier correspond à un numéro de série donné.
	Prend en compte différents formats de noms de fichiers.
	
	Parameters:
	- file_name: Nom du fichier à vérifier
	- SN: Numéro de série à rechercher
	
	Returns:
	- True si le fichier correspond au numéro de série, False sinon
	"""
	# Normaliser le numéro de série pour la recherche
	sn_normalized = SN.lstrip("0")  # Supprimer les zéros non significatifs
	sn_with_zeros = [SN, SN.zfill(2), SN.zfill(3)]  # Versions avec zéros
	
	# Format 1: Recherche directe des différentes variantes du numéro
	for sn_variant in sn_with_zeros:
		if f"SN{sn_variant}" in file_name or f"SN {sn_variant}" in file_name:
			return True
	
	# Format 2: Format spécifique "24.131 SN 03 SE" ou similaire
	# Recherche de motifs comme "SN 03" ou "SN 3" entourés d'espaces ou de caractères non alphanumériques
	pattern1 = re.compile(r'SN\s*' + re.escape(sn_normalized) + r'\b')
	pattern2 = re.compile(r'SN\s*' + re.escape(SN.zfill(2)) + r'\b')
	
	if pattern1.search(file_name) or pattern2.search(file_name):
		return True
	
	# Format 3: Recherche du numéro seul entouré d'espaces ou au début/fin du nom
	# Utile pour les formats comme "24.131 03 SE" où "SN" pourrait être omis
	pattern3 = re.compile(r'(?:^|\s)' + re.escape(sn_normalized) + r'(?:\s|$)')
	pattern4 = re.compile(r'(?:^|\s)' + re.escape(SN.zfill(2)) + r'(?:\s|$)')
	
	if pattern3.search(file_name) or pattern4.search(file_name):
		return True
	
	return False

def FindImagesMatchingSN(folder_path, SN, filter_pattern=None, image_extensions=None):
	"""
	Fonction générique pour trouver toutes les images correspondant à un numéro de série dans un dossier.
	
	Parameters:
	- folder_path: Chemin du dossier à explorer
	- SN: Numéro de série à rechercher
	- filter_pattern: Motif supplémentaire pour filtrer les images (optionnel)
	- image_extensions: Liste des extensions d'images à considérer (optionnel)
	
	Returns:
	- Tuple (sn_images, all_images) contenant les images correspondant au SN et toutes les images
	"""
	if image_extensions is None:
		# on peut prendre des images dans les fichiers excel
		image_extensions = ['.png', '.jpg', '.jpeg', '.bmp', '.tif', '.tiff', '.pdf', '.xlsx', '.xlsm']
	
	sn_images = []
	all_images = []
	
	# Vérifier si le dossier existe
	if not os.path.exists(folder_path):
		Log.Error(f"Le dossier {folder_path} n'existe pas")
		return sn_images, all_images
	
	# Journaliser les critères de recherche
	Log.Message(f"Recherche d'images pour SN{SN} dans {folder_path}" + 
				(f" avec filtre '{filter_pattern}'" if filter_pattern else ""))
	
	# Parcourir tous les fichiers du dossier et ses sous-dossiers
	for root, dirs, files in os.walk(folder_path):
		for file in files:
			file_path = os.path.join(root, file)
			file_name = os.path.basename(file_path)
			ext = os.path.splitext(file_name)[1].lower()
			
			if ext in image_extensions:

				if not file_name.startswith("~$") : # tthe symbol indicating temp files created by excel
					# Ajouter à la liste de toutes les images
					all_images.append(file_path)
					
					# Vérifier si le fichier correspond au numéro de série
					if IsFileMatchingSN(file_name, SN):
						# Si un filtre est spécifié, vérifier aussi le filtre
						if filter_pattern is None:
							sn_images.append(file_path)
							Log.Message(f"Image correspondant à SN{SN} trouvée: {file_name}")
						else:
							# Traitement spécial pour les filtres numériques (_01, _02, etc.)
							if filter_pattern.isdigit():
								# Si le filtre est un nombre, chercher ce nombre dans le nom du fichier
								# ou chercher un index correspondant si plusieurs fichiers sont trouvés
								sn_images.append(file_path)
								Log.Message(f"Image correspondant à SN{SN} trouvée (filtre numérique): {file_name}")
							elif filter_pattern in file_name:
								# Si le filtre est une chaîne, chercher cette chaîne dans le nom du fichier
								sn_images.append(file_path)
								Log.Message(f"Image correspondant à SN{SN} avec filtre '{filter_pattern}' trouvée: {file_name}")
	
	# Si le filtre est numérique et qu'on a trouvé plusieurs images, on peut essayer de sélectionner 
	# l'image correspondant à l'index indiqué par le filtre
	if filter_pattern and filter_pattern.isdigit() and len(sn_images) > 1:
		try:
			# Trier les images par nom de fichier pour assurer une sélection cohérente
			sn_images.sort(key=lambda x: os.path.basename(x))
			
			# Afficher les images triées pour le débogage
			Log.Message(f"Images triées pour le filtre '{filter_pattern}':")
			for i, img in enumerate(sn_images):
				Log.Message(f"  {i+1}: {os.path.basename(img)}")
			
			index = int(filter_pattern) - 1  # Convertir en index 0-based (01 -> index 0)
			if 0 <= index < len(sn_images):
				# Sélectionner uniquement l'image à l'index spécifié
				selected_image = sn_images[index]
				Log.Message(f"Sélection de l'image {index+1} sur {len(sn_images)} pour le filtre '{filter_pattern}': {os.path.basename(selected_image)}")
				return [selected_image], all_images
			else:
				Log.Warning(f"Index {index+1} hors limites pour le filtre '{filter_pattern}' (max: {len(sn_images)})")
		except ValueError:
			# En cas d'erreur, continuer avec toutes les images trouvées
			pass
	
	return sn_images, all_images

# number est le chiffre a trouver dans le nom du fichier pour le considérer valide.
# tag est le tag à rechercher pour qu'il se fasse remplacer
# endtype est une liste, dont il faut conserver la même instance pour différents appels de cette fonction pour un même fichier (DEPRECATED)

# A RETIRER
def AddPicToDoc(path : str, pathID : int, number : str, tag : str, SN : str, endType, titles : list, selection_title=None) :
	"""
	Avoid calling directly, use addimagetodoc instead
	"""
	# Vérifier si le chemin existe
	if not os.path.exists(path):
		Log.Error(f"Le chemin {path} n'existe pas")
		return False

	# Cas spécial pour SN07 - recherche élargie
	if SN == "07":
		Log.Message(f"Recherche spéciale pour SN07 - Élargissement des critères de recherche")
		return HandleSN07Images(path, pathID, number, tag, SN, endType, titles, selection_title)
	
	Log.Message(f"Recherche d'images pour SN{SN} dans {path}")
	
	# Utiliser la fonction générique pour trouver les images
	sn_images, all_images = FindImagesMatchingSN(path, SN)
	
	# Si aucune image n'est trouvée, proposer de sélectionner manuellement
	if not sn_images:
		Log.Message(f"Aucune image trouvée automatiquement pour SN{SN} dans {path}")
		
		# Demander à l'utilisateur s'il souhaite sélectionner manuellement une image
		select_manually = AskQuestion(f"Aucune image trouvée pour SN{SN}. Voulez-vous sélectionner manuellement une image?")
		
		if select_manually:
			# Utiliser le sélecteur d'images pour choisir manuellement
			Log.Message(f"Veuillez sélectionner manuellement une image pour SN{SN}")
			selected_images = ImageReader.SelectImageFromFolder(path, None, allow_multiple=True, 
															  window_title=f"Sélection d'images pour SN{SN} - {tag}")
			
			if selected_images and len(selected_images) > 0:
				sn_images = selected_images
				Log.Message(f"{len(sn_images)} image(s) sélectionnée(s) manuellement pour SN{SN}")
			else:
				Log.Message("Aucune image sélectionnée manuellement")
				return True
	elif len(sn_images) < 2:
		# Si une seule image est trouvée, proposer également de sélectionner manuellement
		Log.Message(f"Seulement {len(sn_images)} image(s) trouvée(s) pour SN{SN}")
		select_manually = AskQuestion(f"Peu d'images trouvées pour SN{SN}. Voulez-vous sélectionner manuellement?")
		
		if select_manually:
			selected_images = ImageReader.SelectImageFromFolder(path, None, allow_multiple=True, 
															  window_title=f"Sélection d'images pour SN{SN} - {tag}")
			
			if selected_images and len(selected_images) > 0:
				sn_images = selected_images
				Log.Message(f"{len(sn_images)} image(s) sélectionnée(s) manuellement pour SN{SN}")
	
	if not sn_images:
		Log.Message(f"Aucune image trouvée pour SN{SN} dans {path} et ses sous-dossiers")
		Log.Message(f"Vérifiez que les images pour SN{SN} existent et contiennent 'SN{SN}' dans leur nom de fichier")
		# Afficher les dossiers explorés pour aider au diagnostic
		Log.Message(f"Dossiers explorés: {path}")
		return True  # Continuer même si aucune image n'est trouvée
	
	# Filtrer les images selon le type (E ou SE) si nécessaire
	filtered_images = sn_images
	if "_E" in tag:
		filtered_images = [img for img in sn_images if any(indicator in img for indicator in ImageReader.etiquetteIndicators)]
		if not filtered_images:
			Log.Message(f"Aucune image avec étiquette trouvée pour SN{SN}")
			Log.Message(f"Vérifiez que les images contiennent un des indicateurs suivants: {ImageReader.etiquetteIndicators}")
	elif "_SE" in tag:
		filtered_images = [img for img in sn_images if not any(indicator in img for indicator in ImageReader.etiquetteIndicators)]
		if not filtered_images:
			Log.Message(f"Aucune image sans étiquette trouvée pour SN{SN}")
			Log.Message(f"Les images sans ces indicateurs sont considérées comme type SE: {ImageReader.etiquetteIndicators}")
	
	if not filtered_images:
		Log.Message(f"Aucune image du type approprié trouvée pour le tag {tag} - Document SN{SN}")
		return True  # Continuer même si aucune image n'est trouvée
	
	# Afficher les images trouvées pour aider au diagnostic
	Log.Message(f"Images trouvées pour SN{SN} ({len(filtered_images)}):")
	for img in filtered_images[:5]:  # Limiter à 5 pour éviter de surcharger le log
		Log.Message(f"  - {os.path.basename(img)}")
	if len(filtered_images) > 5:
		Log.Message(f"  - ... et {len(filtered_images) - 5} autres")
	
	# Sauvegarder les images dans le dossier temp
	temp_images = []
	for image_path in filtered_images:
		success = ImageReader.SaveInTemp(image_path)
		if success:
			image = ImageReader.GetPathToTempImage(image_path)
			if image:
				temp_images.append(image)
	
	if temp_images:
		# Créer des titres descriptifs pour chaque image
		image_titles = []
		for i in range(len(temp_images)):
			# Extraire le nom du fichier sans extension
			base_name = os.path.basename(filtered_images[i])
			file_name = os.path.splitext(base_name)[0]
			
			# Créer un titre descriptif incluant le numéro de série
			image_titles.append(f"Image SN{SN} - {i+1}")
		
		# Remplacer le tag par toutes les images sélectionnées
		Writer.ReplaceTagWithImageInText(tag, temp_images, image_titles)
		
		# Afficher un message pour chaque image insérée
		for image_path in filtered_images:
			Log.Message(f"Image {os.path.basename(image_path)} insérée dans le document pour SN{SN}")
		
		return True
	
	return True





def Yes() :
	global answer
	answer = True
	
	# Ouvrir le répertoire des documents générés dans l'explorateur Windows
	try:
		import os
		import subprocess
		output_dir = os.path.abspath(Settings.GetConfigValueString("paths", "generator_out_path"))
		if os.path.exists(output_dir):
			# Utiliser explorer.exe pour ouvrir le dossier
			subprocess.Popen(f'explorer "{output_dir}"')
			Log.Message(f"Ouverture du répertoire de sortie: {output_dir}")
		else:
			Log.Error(f"Le répertoire de sortie n'existe pas: {output_dir}")
	except Exception as e:
		Log.Error(f"Erreur lors de l'ouverture du répertoire: {str(e)}")
	
	# Forcer la mise à jour de l'interface
	root.update_idletasks()
	return True

def No() :
	global answer
	answer = False
	
	# Forcer la mise à jour de l'interface
	root.update_idletasks()
	return False


# create a popup window, with the SN numbers specified as strings in listSN.
# returns the indices of the selected SN
def GetSelection(listSN, max_selections=None) :
	"""
	Affiche une fenêtre pour sélectionner un ou plusieurs éléments d'une liste
	
	Parameters:
	- listSN: Liste des éléments à afficher
	- max_selections: Nombre maximum d'éléments sélectionnables (None pour pas de limite)
	
	Returns:
	- Liste des indices des éléments sélectionnés
	"""
	selection = []  # Déclaration locale pour éviter l'erreur nonlocal
	
	# Créer une fenêtre pour la sélection
	win = tk.Toplevel()
	win.title("Sélection des Numéros de Série (SN)")
	
	# Augmenter la taille de la fenêtre (2/3 de l'écran)
	screen_width = win.winfo_screenwidth()
	screen_height = win.winfo_screenheight()
	win_width = int(screen_width * 2/3)
	win_height = int(screen_height * 2/3)
	win.geometry(f"{win_width}x{win_height}+{int((screen_width - win_width)/2)}+{int((screen_height - win_height)/2)}")
	
	win.configure(background=Settings.GetConfigValueString("theme", "background_color"))
	
	# Ajouter un titre
	title_frame = tk.Frame(win, background="#3a7ebf")
	title_frame.pack(fill="x", pady=0)
	
	title_label = tk.Label(title_frame, text="SÉLECTION DES NUMÉROS DE SÉRIE", font=("Arial", 14, "bold"), 
						  background="#3a7ebf", foreground="white", pady=10)
	title_label.pack()
	
	# Ajouter un cadre d'instructions
	instructions_frame = tk.Frame(win, background=Settings.GetConfigValueString("theme", "background_color"))
	instructions_frame.pack(fill="x", padx=20, pady=(10, 5))
	
	instructions_label = tk.Label(instructions_frame, 
								 text="Sélectionnez les numéros de série pour lesquels générer des documents:",
								 background=Settings.GetConfigValueString("theme", "background_color"), foreground="#000000", 
								 font=("Arial", 11))
	instructions_label.pack(anchor="w")
	
	if max_selections:
		max_label = tk.Label(instructions_frame, 
							text=f"(Maximum {max_selections} sélections possibles)",
							background=Settings.GetConfigValueString("theme", "background_color"), foreground="#800000", 
							font=("Arial", 10, "italic"))
		max_label.pack(anchor="w")
	
	# Afficher le nombre total de SNs
	count_label = tk.Label(instructions_frame, 
						  text=f"Nombre total de SNs disponibles: {len(listSN)}",
						  background=Settings.GetConfigValueString("theme", "background_color"), foreground="#000000", 
						  font=("Arial", 10, "bold"))
	count_label.pack(anchor="w", pady=(5, 0))
	
	# Créer un cadre pour la liste
	list_frame = tk.Frame(win, background=Settings.GetConfigValueString("theme", "background_color"))
	list_frame.pack(fill="both", expand=True, padx=20, pady=10)
	
	# Créer une scrollbar
	scrollbar = tk.Scrollbar(list_frame)
	scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
	
	# Créer la liste
	listbox = tk.Listbox(list_frame, selectmode=tk.MULTIPLE, font=("Courier New", 12),
						 yscrollcommand=scrollbar.set, height=20, 
						 background="#ffffff", foreground="#000000",  # Fond blanc, texte noir
						 selectbackground="#3a7ebf", selectforeground="white")
	
	# Définir des couleurs alternées pour les lignes
	for i, item in enumerate(listSN):
		listbox.insert(tk.END, item)
		if i % 2 == 0:
			listbox.itemconfig(i, {'bg': '#f0f0f0'})  # Fond gris clair
		# Tous les textes sont en noir pour une meilleure visibilité
		listbox.itemconfig(i, {'fg': '#000000'})  # Texte noir pour tous les items
	
	listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
	scrollbar.config(command=listbox.yview)
	
	# Créer un cadre pour les boutons
	button_frame = tk.Frame(win, background=Settings.GetConfigValueString("theme", "background_color"))
	button_frame.pack(fill="x", padx=20, pady=15)
	
	# Ajouter un bouton pour sélectionner tous les éléments
	select_all_button = tk.Button(button_frame, text="Tout sélectionner", 
								 background="#4a8cce", foreground="white",
								 font=("Arial", 10, "bold"), padx=10, pady=5,
								 command=lambda: listbox.select_set(0, tk.END))
	select_all_button.pack(side=tk.LEFT, padx=5)
	
	# Ajouter un bouton pour effacer la sélection
	clear_button = tk.Button(button_frame, text="Effacer la sélection", 
							background="#d95555", foreground="white",
							font=("Arial", 10, "bold"), padx=10, pady=5,
							command=lambda: listbox.selection_clear(0, tk.END))
	clear_button.pack(side=tk.LEFT, padx=5)
	
	# Variable pour stocker la sélection
	final_selection = []
	
	# Fonction de validation
	def done():
		# Utilise final_selection comme une variable de la fonction externe
		nonlocal final_selection
		selected_items = listbox.curselection()
		
		# Vérifier si le nombre maximum de sélections est dépassé
		if max_selections and len(selected_items) > max_selections:
			Log.Warning(f"Trop d'éléments sélectionnés! Maximum: {max_selections}, Sélectionnés: {len(selected_items)}")
			if not AskQuestion(f"Vous avez sélectionné {len(selected_items)} éléments, mais le maximum est de {max_selections}.\nVoulez-vous continuer avec seulement les {max_selections} premiers éléments?"):
				return
			selected_items = selected_items[:max_selections]
		
		# Convertir en liste et stocker le résultat
		final_selection = list(selected_items)
		Log.Message(f"Indices sélectionnés: {final_selection}")
		win.destroy()
	
	# Ajouter un bouton de validation
	validate_button = tk.Button(button_frame, text="Valider la sélection", 
							  background="#55a155", foreground="white",
							  font=("Arial", 11, "bold"), padx=15, pady=7,
							  command=done)
	validate_button.pack(side=tk.RIGHT, padx=5)
	
	# Mettre à jour la fenêtre principale périodiquement pour éviter qu'elle ne se fige
	def update_window():
		if win.winfo_exists():
			win.after(100, update_window)
	
	update_window()
	
	# Attendre que la fenêtre soit fermée
	win.wait_window()
	
	# Retourner la sélection finale
	Log.Message(f"Sélection finale: {final_selection}")
	return final_selection

def DoneButton() :
	global doneSelecting
	doneSelecting = True


def UpdateRoot() :
	loadingIcon.Update()
	root.update()



def OnQuit() :
	# Enregistrer les chemins et valeurs qui sont encore utilisés
	#Settings.paths[Settings.PATH_RCI] = outputFolderPath.get()
	#Settings.fieldValues['FIELD_FILENAME'] = fileNameFormat.get()
	# Note: Ces méthodes n'existent plus dans la nouvelle structure
	# Settings.WriteFieldValues()
	# Settings.WritePaths()
	# Utiliser la nouvelle méthode pour sauvegarder
	Settings.SaveSettings()

def GenerateDocumentsForSNList(sn_list):
	"""
	Génère des documents pour une liste spécifique de numéros de série
	
	Parameters:
	- sn_list: Liste des numéros de série à traiter
	"""
	# Lire d'abord les données du fichier Excel
	try:
		Log.Message("Lecture du fichier Excel pour obtenir les données des SN...")
		if not values.FetchValuesInExcel():
			raise Exception("Échec de la lecture du fichier Excel")
		
		all_containers = values.containers
		if not all_containers:
			raise Exception("Aucune donnée n'a pu être récupérée du fichier Excel")
		
		# Créer un dictionnaire pour un accès rapide aux containers par SN
		container_dict = {container.SN: container for container in all_containers}
		
		# Vérifier que tous les SN appartiennent au LT actuel
		lt_sns = values.FindSNFromLT(values.LT)
		lt_sns_set = set(lt_sns)
		
		# Filtrer les SN qui existent dans le fichier Excel ET appartiennent au LT
		valid_sn_list = []
		for sn in sn_list:
			if sn in container_dict:
				if sn in lt_sns_set:
					valid_sn_list.append(sn)
					Log.Message(f"SN{sn} trouvé dans le fichier Excel et appartient au LT {values.LT}")
				else:
					Log.Warning(f"SN{sn} trouvé dans le fichier Excel mais n'appartient pas au LT {values.LT}")
			else:
				# Créer un container de test pour ce SN seulement s'il appartient au LT
				if sn in lt_sns_set:
					Log.Warning(f"SN{sn} non trouvé dans le fichier Excel mais appartient au LT {values.LT}, création d'un container de test")
					test_container = ValueFetcher.Container(sn)
					test_container.LT = values.LT
					test_container.numPlan = values.numPlan
					test_container.titrePlan = "Plan de test"
					test_container.date = "01/01/2023"
					test_container.undefinedValues = ""
					test_container.tagAndValues = {}
					
					# Note: Les anciens tags Settings.TAG_* n'existent plus
					# Utiliser les nouvelles clés de configuration
					test_container.tagAndValues["**SN**"] = test_container.SN
					test_container.tagAndValues["**LT**"] = test_container.LT
					test_container.tagAndValues["**NUMPLAN**"] = test_container.numPlan
					test_container.tagAndValues["**DATE**"] = test_container.date
					test_container.tagAndValues["**TITREPLAN**"] = test_container.titrePlan
					
					test_container.TryVerifyAllSpecs = lambda: True
					container_dict[sn] = test_container
					valid_sn_list.append(sn)
				else:
					Log.Warning(f"SN{sn} non trouvé dans le fichier Excel et n'appartient pas au LT {values.LT}")
		
		if not valid_sn_list:
			Log.Error(f"Aucun des SN du template n'a été trouvé dans le fichier Excel pour le LT {values.LT}")
			return False
		
		# Générer un document pour chaque SN valide
		for sn in valid_sn_list:
			Log.Message(f"\n=== GÉNÉRATION DU DOCUMENT POUR SN{sn} ===")
			GenerateSingleDocument(sn)
		
		return True
		
	except Exception as e:
		Log.Error(f"Erreur lors de la génération des documents pour les SN du template: {str(e)}")
		return False

def AskOptionIndex(question, max_options):
	"""
	Demande à l'utilisateur de choisir une option parmi plusieurs
	
	Parameters:
	- question: Question à poser à l'utilisateur
	- max_options: Nombre maximum d'options disponibles
	
	Returns:
	- Indice de l'option choisie (1-based)
	"""
	while True:
		option_str = AskInputQuestion(question)
		try:
			option = int(option_str)
			if 1 <= option <= max_options:
				return option
			else:
				Log.Error(f"Veuillez entrer un nombre entre 1 et {max_options}")
		except ValueError:
			Log.Error("Veuillez entrer un nombre valide")

def AskInputQuestion(question):
	"""
	Affiche une boîte de dialogue pour demander une entrée à l'utilisateur
	
	Parameters:
	- question: Question à poser à l'utilisateur
	
	Returns:
	- Réponse de l'utilisateur (chaîne de caractères)
	"""
	global input_answer
	input_answer = None
	
	# Créer une fenêtre de dialogue
	dialog = tk.Toplevel()
	dialog.title("Sélection d'option")
	dialog.geometry("550x350")
	dialog.configure(background=Settings.GetConfigValueString("theme", "background_color"))
	
	# Ajouter un titre
	title_frame = tk.Frame(dialog, background="#3a7ebf")
	title_frame.pack(fill="x", pady=0)
	
	title_label = tk.Label(title_frame, text="MENU DE SÉLECTION", font=("Arial", 14, "bold"), 
						  background="#3a7ebf", foreground="white", pady=10)
	title_label.pack()
	
	# Ajouter la question dans un cadre
	question_frame = tk.Frame(dialog, background=Settings.GetConfigValueString("theme", "background_color"), 
							 highlightbackground=Settings.GetConfigValueString("theme", "outline_color"), highlightthickness=1)
	question_frame.pack(fill="both", expand=True, padx=20, pady=15)
	
	# Traiter le texte pour mettre en évidence les options numérotées
	processed_text = question
	
	question_label = tk.Label(question_frame, text=processed_text, justify="left", wraplength=500,
							 background=Settings.GetConfigValueString("theme", "background_color"), foreground=Settings.GetConfigValueString("theme", "text_color"),
							 font=("Arial", 11))
	question_label.pack(pady=15, padx=15)
	
	# Ajouter un champ de saisie
	entry_frame = tk.Frame(dialog, background=Settings.GetConfigValueString("theme", "background_color"))
	entry_frame.pack(fill="x", padx=20, pady=10)
	
	entry_label = tk.Label(entry_frame, text="Votre choix :", font=("Arial", 11, "bold"),
						  background=Settings.GetConfigValueString("theme", "background_color"), foreground=Settings.GetConfigValueString("theme", "text_color"))
	entry_label.pack(side="left", padx=5)
	
	entry = tk.Entry(entry_frame, width=10, background=Settings.GetConfigValueString("theme", "entry_color"), foreground=Settings.GetConfigValueString("theme", "text_color"),
					font=("Arial", 12))
	entry.pack(side="left", padx=5)
	entry.focus_set()
	
	# Fonction pour valider la réponse
	def validate():
		global input_answer
		input_answer = entry.get()
		dialog.destroy()
	
	# Ajouter un bouton de validation
	button_frame = tk.Frame(dialog, background=Settings.GetConfigValueString("theme", "background_color"))
	button_frame.pack(fill="x", pady=15)
	
	button = tk.Button(button_frame, text="VALIDER", command=validate, width=15, height=2,
					  background="#4CAF50", foreground="white", font=("Arial", 11, "bold"),
					  activebackground="#45a049", activeforeground="white")
	button.pack()
	
	# Lier la touche Entrée à la validation
	entry.bind("<Return>", lambda event: validate())
	dialog.bind("<Return>", lambda event: validate())
	
	# Centrer la fenêtre
	dialog.update_idletasks()
	width = dialog.winfo_width()
	height = dialog.winfo_height()
	x = (dialog.winfo_screenwidth() // 2) - (width // 2)
	y = (dialog.winfo_screenheight() // 2) - (height // 2)
	dialog.geometry('{}x{}+{}+{}'.format(width, height, x, y))
	
	# Rendre la fenêtre modale
	dialog.transient(root)
	dialog.grab_set()
	
	# Attendre que l'utilisateur réponde
	while input_answer is None:
		try:
			root.update_idletasks()
			root.update()
		except:
			break
	
	return input_answer