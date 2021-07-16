# Import library
import tkinter as tk
from tkinter import *
from tkinter import ttk
from tkinter import filedialog as fd
from tkinter.messagebox import showinfo

# Importing of function
from pandasPf import ConnectorType
from mufy import MufsNumbers
from PandasAeUn import AerialOrUnderground

# Adding new window
root = tk.Tk()
# Name of main window
root.title('Automatyzacja działań')
# Dimensions of main window
root.geometry("500x500")
# Icon
root.iconbitmap('C:/Users/marek.barul/Desktop/Python pliki/rozliczenia/FONBUD ICON.ico')

# Funkcja, will run when we click at button "Zaakceptuj"
def OnClick():
    ConnectorType(spliceCardsBox.get(), sourceFileBox.get())
    MufsNumbers(spliceCardsBox.get())
    AerialOrUnderground(spliceCardsBox.get())

# 1 Window for folder
spliceCardsFrame = LabelFrame(
    root, text="Podaj lokalizację folderu Splice Cards", padx=5, pady=5)
spliceCardsFrame.pack()
# Window to write a command
spliceCardsBox = Entry(spliceCardsFrame, width=70, borderwidth=5)
spliceCardsBox.insert(0, "")
# Location of window
spliceCardsBox.pack(pady=20)
# Open folder dialog box
def selectFolder():
    foldername = fd.askdirectory(
        title='Wybierz folder')
    spliceCardsBox.insert(0, f'{foldername}')
    showinfo(
        title='Wybrany folder',
        message=foldername
    )

openFolderButton = ttk.Button(
    spliceCardsFrame, text='...', command=selectFolder
)
openFolderButton.pack()

# 2 Frame for source file
sourceFileFrame = LabelFrame(
    root, text="Wybierz plik źródłowy", padx=5, pady=5)
sourceFileFrame.pack()
# 2.1 Window
sourceFileBox = Entry(sourceFileFrame, width=70, borderwidth=5)
sourceFileBox.insert(0, "")
sourceFileBox.pack(pady=20)

# Open files dialog box
def selectSourceFile():
    filetypes = (
        ('excel files', '*.xlsx'),
        ('all files', '*.*')
    )
    filename = fd.askopenfilename(
        title='Plik źródłowy',
        initialdir='/',
        filetypes=filetypes
    )
    sourceFileBox.insert(0, f'{filename}')
    showinfo(
        title='Wybrany plik',
        message=filename
    )

openFileButton = ttk.Button(
    sourceFileFrame, text='...', command=selectSourceFile
)
openFileButton.pack()

# Button of function OnClick
myButton = tk.Button(root, text="Zaakceptuj", width=10, command=OnClick)
myButton.pack()
# Button to quit from the app
buttonQuit = Button(root, text="Zakończ", width=10, command=root.quit)
buttonQuit.pack()

root.mainloop()
