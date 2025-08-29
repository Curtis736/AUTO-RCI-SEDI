# create a popup window, with the SN numbers specified as strings in listSN.
# returns the indices of the selected SN
import tkinter as tk
from functools import partial

import Log
import Settings

root = None

def GetSelection(listSN) :
    """
    returns the indices of the selected items
    """
    global doneSelecting
    doneSelecting = False

    selectionWindow = tk.Toplevel(background=Settings.backgroundColor)
    selectionWindow.title("selection SN")
    selectionWindow.protocol("WM_DELETE_WINDOW", OnExit)

    # we give it an appropriate size : about half the screen
    screenWidth = root.winfo_screenwidth()
    screenHeight = root.winfo_screenheight()
    # this method take for argument a string of format "widthxheight+offset_with+offset_height"
    selectionWindow.geometry((str)(screenWidth // 4) + "x" + (str)(screenHeight // 2) + "+" + (str)(screenWidth // 4) + "+" + (str)(screenHeight // 4))

    label = tk.Label(selectionWindow, text="choisissez les SN à traiter", background=Settings.backgroundColor, foreground=Settings.textColor)
    label.pack(side=tk.TOP, fill="x")

    options = listSN
    selectionBox = tk.Listbox(selectionWindow, selectmode="multiple", background=Settings.backgroundColor)
    selectionBox.pack(fill="both")

    btn = tk.Button(selectionWindow, text="done", command=DoneButton, background=Settings.unselectedColor, foreground=Settings.textColor)
    btn.pack(side="bottom", fill="none")

    for i in range(len(options)) :
        selectionBox.insert(tk.END, options[i])
        selectionBox.itemconfig(i, bg = "gray20" if i % 2 == 0 else "gray25")
        selectionBox.itemconfig(i, fg = Settings.textColor)
    
    global blocked
    blocked = True
    Log.Verbose("selector open")
    while not doneSelecting :
        root.update_idletasks()
        root.update()
    
    result = []
    for i in selectionBox.curselection() :
        result.append(i)
    
    selectionWindow.destroy()
    Log.Verbose("selector closed")

    Log.Verbose(str(result))
    blocked = False
    return result

class SelectionButton(tk.Button) :

    def __init__(self, master, text : str, index : int) :
        super().__init__(master, text=text, command= partial(OnClick, index))
        self.index = index
        self.selected = False
    
    def Select(self) :
        self.config(background="blue")
        self.selected = True
    
    def Deselect(self) :
        self.config(background="grey64")
        self.selected = False

def OnClick(index : int) :
    # resolves what gets selected / deselected
    global lastClick
    if shiftPressed :
        theRange = None
        if lastClick > index :
            theRange = range(index, lastClick + 1)
        else :
            theRange = range(lastClick, index + 1)
            
        for i in theRange :
            buttonList[i].Select()
    else :
        if not ctrlPressed :
            for button in buttonList :
                button.Deselect()
            buttonList[index].Select()
        else :
            if buttonList[index].selected :
                buttonList[index].Deselect()
            else :
                buttonList[index].Select()
    
    lastClick = index
            
    pass


lastClick = 0
shiftPressed = False
ctrlPressed = False
buttonList = []

def GetSelectionAdvanced(theList) :
    """
    List should be a list of str
    Returns the index of selected stuff
    """
    global doneSelecting
    doneSelecting = False

    selectionWindow = tk.Toplevel()
    selectionWindow.title("selection SN")
    selectionWindow.protocol("WM_DELETE_WINDOW", OnExit)
    selectionWindow.bind('<KeyPress>', OnKeyPress)
    selectionWindow.bind('<KeyRelease>', OnKeyRelease)

    # we give it an appropriate size : about half the screen
    screenWidth = root.winfo_screenwidth()
    screenHeight = root.winfo_screenheight()
    # this method take for argument a string of format "widthxheight+offset_with+offset_height"
    selectionWindow.geometry((str)(screenWidth // 4) + "x" + (str)(screenHeight // 2) + "+" + (str)(screenWidth // 4) + "+" + (str)(screenHeight // 4))

    label = tk.Label(selectionWindow, text="choisissez les SN à traiter\n(utilisez q au lieu de shift pour sélectionner\n plusieurs éléments d'un coup)\n(utilisez w au lieu de control)")
    label.pack(side=tk.TOP, fill="x")

    frame = tk.Frame(selectionWindow, highlightthickness=2, highlightbackground="black")
    frame.pack(fill="both")

    buttonFrame = tk.Text(frame, highlightthickness=2, highlightbackground="black")

    scroll = tk.Scrollbar(frame)
    scroll.pack(side=tk.RIGHT, fill="y")

    scroll.config(command=buttonFrame.xview)
    buttonFrame.config(yscrollcommand=scroll.set)

    i = 0
    buttonList.clear()
    global lastClick
    lastClick = 0
    for item in theList :
        button = SelectionButton(buttonFrame, item, i)
        buttonList.append(button)
        buttonFrame.window_create("end", window=button, stretch=True)
        buttonFrame.insert("end", "\n")
        i += 1
    buttonFrame.pack(fill="both")
    
    btn = tk.Button(selectionWindow, text="done", command=DoneButton)
    btn.pack(side="bottom", fill="none")
    
    global blocked
    blocked = True
    Log.Verbose("selector open")
    while not doneSelecting :
        root.update_idletasks()
        root.update()
    
    result = []
    for button in buttonList :
        if button.selected :
            result.append(button.index)
    
    selectionWindow.destroy()
    Log.Verbose("selector closed")

    Log.Verbose(str(result))
    blocked = False
    return result





def OnKeyPress(event) :
    global shiftPressed, ctrlPressed
    if event.char == "q" :
        shiftPressed = True
    if event.char == "w" :
        ctrlPressed = True

def OnKeyRelease(event) :
    global shiftPressed, ctrlPressed
    if event.char == "q" :
        shiftPressed = False
    if event.char == "w" :
        ctrlPressed = False


def OnExit() :
    DoneButton()

def DoneButton() :
    global doneSelecting
    doneSelecting = True