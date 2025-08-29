import os

import tkinter as tk
from tkinter import ttk
from tkinter import StringVar
from tkinter import scrolledtext

import Settings
import Log


"""
A frame widget made to accept tabs objects, to display different widgets at the same spot depending on which tab was chosen.
Inherit from frame, and will contain the buttons to switch tabs.
Only one tab can be active at once.
"""
class TabBar(tk.Frame) :

	"""
	master : The widget in which the tab bar belongs
	display : a tk.Frame that will be used to display the contents of each tabs when they are active
	"""
	def __init__(self, master, display : tk.Frame, height=30) :
		super().__init__(master,
						highlightthickness=1,
						highlightbackground=Settings.GetConfigValueString("theme", "outline_color"),
						background=Settings.GetConfigValueString("theme", "background_color"))
		
		# if True, cannot change the active tab
		self.tablock = False

		# the list of tab objects
		self.tabs = []

		self.display = display
	

	def SetTabLock(self, val : bool) :

		# prevents tab switching
		self.tablock = val


				


"""
Class that represent a tabbed window that will add itself to a tab bar.
Put itself on the display widget of the tab bar and add a button in the tab bar.
"""
class Tab(tk.Frame) :

	tabLock = False

	"""
	master is the TabBar object
	display is the widget where the window content associated with
	"""
	def __init__(self, tabBar : TabBar, name : str, onOpenCallback = None) :

		self.__tabBar = tabBar

		super().__init__(tabBar.display,
						highlightthickness=1,
						highlightbackground=Settings.GetConfigValueString("theme", "outline_color"),
						background=Settings.GetConfigValueString("theme", "background_color"))
		
		self.__name = name


		# the button that appears in the tabbar and allow to choose this tab
		self.__button = tk.Button(tabBar, text=name,
								command=lambda: self.Open(), 
								foreground=Settings.GetConfigValueString("theme", "text_color"),
								background=Settings.GetConfigValueString("theme", "unselected_color"),
								activebackground=Settings.GetConfigValueString("theme", "selected_color"),
								activeforeground=Settings.GetConfigValueString("theme", "text_color"),
								relief=tk.RAISED,
								borderwidth=2,
								padx=10,
								pady=3
								)
		
		# add the button on the tab bar
		self.__button.pack(side=tk.LEFT, padx=2)
		

		# the widget that will be put over displayWidget at the same size, to display the tab contents
		self.contentWidget = tk.Frame(tabBar.display,
								background=Settings.GetConfigValueString("theme", "background_color")
								)
		
		self.isOpen = False

		# add itself to the list of tabs
		self.__tabBar.tabs.append(self)

		for i in range(20) :
			self.grid_rowconfigure(i, weight=1)
			self.grid_columnconfigure(i, weight=1)
				
		self.onOpenCallback = onOpenCallback

		self.Close()
		
	def Open(self) :
		
		# cannot open a new tab if tabs are locked
		if self.__tabBar.tablock : return

		# close every other tab
		for tab in self.__tabBar.tabs :

			if tab != self :
				tab.Close()
		
		# shows itself
		self.place(x=0, y=0, relwidth=1, relheight=1)
		self.isOpen = True

		if self.onOpenCallback != None : self.onOpenCallback()
	

	def Close(self) :

		# cannot close if tabs are locked
		if self.__tabBar.tablock : return

		# hides itself
		self.place_forget()
		self.isOpen = False
	
	# prevents tabs from switching
	def SetTabLock(self, val : bool) :
		self.__tabBar.SetTabLock(val)



"""
Base class for any user input fields that uses a StringVar.
Automatically updates a corresponding setting when modified.
Only one method has to be called to load all settings accross all fields, and same for saving.
This is useful to manage program closing and opening.
Also allow to make two different fields that will share a same value.
"""
class StorageField(tk.Frame) :

	# all instances of this class, children and everything
	instances = []

	def __init__(self, master : tk.Widget, config : str, key : str) :
		super().__init__(master, bg=Settings.GetConfigValueString("theme", "background_color"))
		
		StorageField.instances.append(self)

		self.stringVar = StringVar()
		# three arguments necessary
		self.stringVar.trace_add("write", lambda a,b,c : self.SaveValue())

		# intended to contain the input field using self.stringVar
		self.field = None

		self.config = config
		self.key = key

		self.blockSave = False

		self.LoadValue()

	
	def SaveValue(self) :
		if self.blockSave : return
		Settings.SetConfigValue(self.config, self.key, self.stringVar.get())
	
	def LoadValue(self) :
		self.blockSave = True
		self.stringVar.set(Settings.GetConfigValueString(self.config, self.key))
		self.blockSave = False
	

	def LoadAll() :
		for instance in StorageField.instances : instance.LoadValue()
	
	def SaveAll() :
		for instance in StorageField.instances : instance.SaveValue()

	
	def GetValue(self) :
		return self.stringVar.get()
	
	def SetValue(self, value : str) :
		self.stringVar.set(value)
	

	def Lock(self) :
		if hasattr(self.field, "state") :
			self.field.state = "disabled"
	
	def Unlock(self) :
		if hasattr(self.field, "state") :
			self.field.state = "enabled"



"""
Subclass for simple text input
"""
class StorageTextField(StorageField) :


	def __init__(self, master, config : str, key : str, label : str):
		super().__init__(master, config, key)

		#creates the caption on top of the field
		self.text = label
		self.caption = tk.Label(self, text=label, foreground=Settings.GetConfigValueString("theme", "text_color"), background=Settings.GetConfigValueString("theme", "background_color"))
		self.caption.pack(side=tk.TOP)

		# the field
		self.field = tk.Entry(self, textvariable=self.stringVar)
		self.field.pack(side=tk.BOTTOM)


class StorageDropdown(StorageField) :

	def __init__(self, master, config : str, key : str, label : str, values : list):
		super().__init__(master, config, key)

		#creates the caption on top of the field
		self.text = label
		self.caption = tk.Label(self, text=label, foreground=Settings.GetConfigValueString("theme", "text_color"), background=Settings.GetConfigValueString("theme", "background_color"))
		self.caption.pack(side=tk.TOP)

		style = ttk.Style()
		style.configure("SD.TCombobox", foreground=Settings.GetConfigValueString("theme", "text_color"), background=Settings.GetConfigValueString("theme", "background_color"))
		
		style.map("SD.TCombobox", foreground=[('active', Settings.GetConfigValueString("theme", "text_color")),
										('invalid', Settings.GetConfigValueString("theme", "text_color")),
										('alternate', Settings.GetConfigValueString("theme", "text_color"))])
		"""
										   ('readonly', Settings.GetConfigValueString("theme", "text_color")),
										   ('pressed', Settings.GetConfigValueString("theme", "text_color")),
										   ('focus', Settings.GetConfigValueString("theme", "text_color")),
										   ('hover', Settings.GetConfigValueString("theme", "text_color"))])
		"""

		# the field
		self.field = ttk.Combobox(self, textvariable=self.stringVar, values=values, style="SD.TCombobox")
		print(f"état {self.field.state()}")

		self.field.pack(side=tk.BOTTOM)

		self.__values = values
	
	def SaveValue(self) :
		super().SaveValue()
		print(f"état {self.field.state()}")

	def GetList(self) :
		return self.__values
	
	def SetList(self, values) :
		self.__values = values
		self.field.configure(values=self.__values)



class StorageLabel(StorageField) :

	def __init__(self, master, config : str, key : str, wrapCount = 0, wrapChar = " "):

		self.displayStringVar = StringVar()

		# the maximum desired amount of characters on a single line
		self.wrapCount = wrapCount
		# the character that allow newLines for text wrapping
		self.wrapChar = wrapChar

		super().__init__(master, config, key)

		self.field = tk.Label(self, textvariable=self.displayStringVar,
							background=Settings.GetConfigValueString("theme", "background_color"),
							foreground=Settings.GetConfigValueString("theme", "text_color"))
		self.field.pack(side=tk.TOP)
	
	def LoadValue(self) :
		super().LoadValue()
		self.Wrap()
	
	def SetValue(self, value):
		super().SetValue(value)
		self.Wrap()
		
	def Wrap(self) :
		# if wrapCount is 0 then wrapping is disabled
		if self.wrapCount == 0 :
			self.displayStringVar.set(self.stringVar.get())
			return
		
		# else wrapping is enabled

		# split on each wrapChar
		result = ""
		splitted = self.stringVar.get().split(self.wrapChar)
		charCount = 0
		for i in range(len(splitted)) :
			if charCount + len(splitted[i]) > self.wrapCount :
				# insert a newline to avoid going over the limit
				result += "\n" + splitted[i] + self.wrapChar
				charCount = 0
			else :
				result += splitted[i] + self.wrapChar
				charCount += len(splitted[i])
		
		result = result[:-len(self.wrapChar)]
		print(result)

		self.displayStringVar.set(result)


class LabelledCheckbox(tk.Frame) :

    def __init__(self, master, intVariable : tk.IntVar, label = "-") :
        super().__init__(master, highlightthickness=1, highlightbackground=Settings.GetConfigValue("theme", "outline_color"), background=Settings.GetConfigValue("theme", "background_color"))

        self.label = tk.Label(self, text=label, foreground=Settings.GetConfigValue("theme", "text_color"), background=Settings.GetConfigValue("theme", "background_color"))
        self.label.pack(side=tk.LEFT, fill="x")
        self.entry = tk.Checkbutton(self, onvalue=1, offvalue=0, variable=intVariable, background=Settings.GetConfigValue("theme", "background_color"), foreground="black")
        self.entry.pack(side=tk.RIGHT, padx=2)


class OutputTextWindow(tk.Frame) :

	def __init__(self, master) :
		super().__init__(master, background="red")

		# create the window that will display the text
		self.textWindow = scrolledtext.ScrolledText(self, wrap=tk.WORD)
		self.textWindow.configure(background=Settings.GetConfigValueString("theme", "entry_color"), foreground=Settings.GetConfigValueString("theme", "text_color"))
		self.textWindow.bind("<Key>", lambda e: "break")
		self.textWindow.place(x=0, y=0, relwidth=1, relheight=1)

		self.textWindow.tag_config("question", foreground=Settings.GetConfigValueString("theme", "question_color"))
		self.textWindow.tag_config("warning", foreground=Settings.GetConfigValueString("theme", "warning_color"))
		self.textWindow.tag_config("error", foreground=Settings.GetConfigValueString("theme", "error_color"))

		# connects to callbacks from Log
		Log.AddCallback(Log.Lvl.ERR, lambda x : self.ErrorCallback(x))
		Log.AddCallback(Log.Lvl.WARN, lambda x : self.WarningCallback(x))
		Log.AddCallback(Log.Lvl.MSG, lambda x : self.MessageCallback(x))
	
	def ErrorCallback(self, string : str) :
		self.textWindow.insert(tk.END, string, "error")
		self.textWindow.see(tk.END)
	
	def WarningCallback(self, string : str) :
		self.textWindow.insert(tk.END, string, "warning")
		self.textWindow.see(tk.END)

	def MessageCallback(self, string : str) :
		self.textWindow.insert(tk.END, string)
		self.textWindow.see(tk.END)

	def QuestionCallback(self, string : str) :
		self.textWindow.insert(tk.END, string, "question")
		self.textWindow.see(tk.END)



class ImageSelector(tk.Toplevel):
	"""
	Classe pour afficher une fenêtre de sélection d'images
	"""
	def __init__(self, master, images_list, callback, allow_multiple=False, window_title=None):
		"""
		Initialise la fenêtre de sélection d'images
		
		Parameters:
		- master: fenêtre parente
		- images_list: liste des chemins d'images à afficher
		- callback: fonction à appeler avec l'image sélectionnée
		- allow_multiple: permet la sélection de plusieurs images
		- window_title: titre personnalisé pour la fenêtre
		"""
		super().__init__(master, background=Settings.backgroundColor)
		
		# Définir le titre de la fenêtre
		if window_title:
			self.title(window_title)
		else:
			self.title("Sélection d'image")
		
		self.images_list = images_list
		self.callback = callback
		self.allow_multiple = allow_multiple
		self.selected_images = []
		
		# Configurer la taille de la fenêtre
		width = 500
		height = 400
		screen_width = self.winfo_screenwidth()
		screen_height = self.winfo_screenheight()
		x = (screen_width - width) // 2
		y = (screen_height - height) // 2
		self.geometry(f"{width}x{height}+{x}+{y}")
		
		# Créer le cadre principal
		main_frame = tk.Frame(self, background=Settings.backgroundColor)
		main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
		
		# Titre
		title_text = "Sélectionnez une ou plusieurs images" if allow_multiple else "Sélectionnez une image"
		title_label = tk.Label(main_frame, text=title_text, 
							  background=Settings.backgroundColor, 
							  foreground=Settings.textColor,
							  font=("Arial", 12, "bold"))
		title_label.pack(pady=(0, 10))
		
		# Liste des images
		list_frame = tk.Frame(main_frame, background=Settings.entryColor)
		list_frame.pack(fill=tk.BOTH, expand=True, pady=5)
		
		scrollbar = tk.Scrollbar(list_frame)
		scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
		
		# Configurer le mode de sélection en fonction de allow_multiple
		select_mode = tk.MULTIPLE if allow_multiple else tk.SINGLE
		
		self.listbox = tk.Listbox(list_frame, 
								 yscrollcommand=scrollbar.set,
								 background=Settings.entryColor,
								 foreground=Settings.textColor,
								 selectbackground=Settings.selectedColor,
								 selectforeground=Settings.textColor,
								 selectmode=select_mode,
								 font=("Arial", 10))
		self.listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
		scrollbar.config(command=self.listbox.yview)
		
		# Remplir la liste avec les noms des images
		for img in self.images_list:
			self.listbox.insert(tk.END, os.path.basename(img))
		
		# Boutons
		button_frame = tk.Frame(main_frame, background=Settings.backgroundColor)
		button_frame.pack(fill=tk.X, pady=10)
		
		select_button = tk.Button(button_frame, 
								 text="Sélectionner", 
								 command=self.select_image,
								 background=Settings.unselectedColor,
								 foreground=Settings.textColor)
		select_button.pack(side=tk.RIGHT, padx=5)
		
		cancel_button = tk.Button(button_frame, 
								 text="Annuler", 
								 command=self.cancel,
								 background=Settings.unselectedColor,
								 foreground=Settings.textColor)
		cancel_button.pack(side=tk.RIGHT, padx=5)
		
		# Configurer la fermeture de la fenêtre
		self.protocol("WM_DELETE_WINDOW", self.cancel)
		
		# Rendre la fenêtre modale
		self.transient(master)
		self.grab_set()
		master.wait_window(self)
	
	def select_image(self):
		"""Sélectionne l'image ou les images et appelle le callback"""
		selections = self.listbox.curselection()
		if selections:
			if self.allow_multiple:
				# Sélection multiple
				self.selected_images = [self.images_list[index] for index in selections]
				self.callback(self.selected_images)
			else:
				# Sélection unique
				index = selections[0]
				self.selected_images = [self.images_list[index]]
				self.callback(self.selected_images[0])
			self.destroy()
		else:
			# Aucune sélection
			if self.allow_multiple:
				self.callback([])
			else:
				self.callback(None)
			self.destroy()
	
	def cancel(self):
		"""Annule la sélection"""
		if self.allow_multiple:
			self.callback([])
		else:
			self.callback(None)
		self.destroy()
		

"""
weird thing
"""
class LoadingIcon(tk.Frame) :

	loadingIcons = []
	root = None

	def __init__(self, master) :
		self.displayText = ""
		self.step = 0
		self.text = StringVar()
		self.isloading = False
		super().__init__(master, background=Settings.GetConfigValueString("theme", "background_color"))
		self.label = tk.Label(self, 
							  textvariable=self.text, 
							  background=Settings.GetConfigValueString("theme", "background_color"), 
							  foreground=Settings.GetConfigValueString("theme", "text_color"))
		self.label.pack()
		self.text.set("---")
		LoadingIcon.loadingIcons.append(self)
	
	def SetText(self, string : str) :
		self.displayText = string
		self.Update()
	
	def Update(self) :
		if self.isloading :
			if self.step == 0 :
				self.text.set(self.displayText + "*..")
			elif self.step == 1 :
				self.text.set(self.displayText + ".*.")
			elif self.step == 2 :
				self.text.set(self.displayText + "..*")
			self.step += 1
			if self.step > 2 :
				self.step = 0
		else :
			self.step = 0
			self.text.set(self.displayText)
	
	@staticmethod
	def UpdateAll() :
		if LoadingIcon.root != None :
			
			for loadingIcon in LoadingIcon.loadingIcons :
				loadingIcon.Update()
			
			LoadingIcon.root.update()