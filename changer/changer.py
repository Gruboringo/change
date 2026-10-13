import  win32com.client #modyfikowanie skrótów 
import tkinter #do guziczków i kości GUI
import json #do odczytywania i robienia plikow typu JSON
from tkinter import filedialog #do okna wyboru guziczków 
from tkinter import messagebox

root = tkinter.Tk()
root.title("zmiana ikony") #nazwa okna guziczkowego 
from PIL import Image
from pathlib import Path #do ściezek plikow 
import sys
project_path = Path(sys.executable).parent if getattr(sys, "frozen", False) else Path(__file__).parent #pobiera sciezke w ktorej znajduje sie program 
backup_path=project_path/"backups"/"backup.json" #ustawia sciezke backupa tych JSOM 
import os #udostępnia funkcje systemowe
import sys
print(sys.executable)
print(os.getcwd())


file_path = "" #sciezka do sciezki lnk
icon_path = "" #sciezka do sciezki ico

def backup(file_path): #funckja do robienia backupa orginalnej sciezki ico 
    with open(backup_path, "r" ) as file: # open otwiera plik. pierwsza "rzecz" to lokalizacja pliku który trzeba otworzyć. drugie czyli "r" to tryb odczytu 
        backed = json.load(file) #json.load(file) zamienia to co jest w JSON na obiekt którym jest "backed" - dictonary
    if file_path in backed:
        print("juz jest")
        return
    else:
        shortcut = shell.CreateShortcut(file_path) #otwiera wybrany skrót Windows, czyli ten z file_path, 
        old_icon = shortcut.IconLocation #znajduje ikone skrótu 
        backed[file_path] = old_icon # zapisuje w słowniku: klucz = ścieżka skrótu, value = lokalizacja jego starej ikony
        save = json.dumps(backed) #zamienia dictonary na tekst w formacie JSON
        with open(backup_path, "w") as file: #"w" - tryb pisania 
            file.write(save) #wpisuje do pliku json to co jest w obiekcie "save"

def lnk(): #funckja do odczytu lokalizacji lnk
     global file_path #nawet po wyjsciu z funkcji filepath w globalu dostaję tą informacje 
     file_path = filedialog.askopenfilename(filetypes=[("Pliki lnk", "*.lnk")])  
     if not file_path:        
         messagebox.showinfo("Anulowano", "Nie wybrano pliku.")       
         return   
     backup(file_path)




button = tkinter.Button(root, text="wybierz skrót", command=lnk, width=7, height=8) #guziczek do wyboru pliku lnk
button.pack()



    
def convert_icon():
     global icon_path
     icon_path = filedialog.askopenfilename(filetypes=[("Obrazy", "*.png *.jpg *.jpeg *.ico")])
     if not icon_path:
          messagebox.showinfo("Anulowano", "Nie wybrano pliku.")
          return
     img = Image.open(icon_path)
     suffix = Path(icon_path).suffix.lower()
     if suffix!=".ico":
          source = Path(icon_path)
          new_icon=source.with_suffix(".ico")
          img.save(new_icon, format="ICO")
          icon_path = str(new_icon)
    

          
          
          
          





button = tkinter.Button(root, text="wybierz ikone", command=convert_icon, width=8, height=5) #guziczek do wyboru pliku ico i backupa

button.pack()

def change_Shorticon(file_path, icon_path):     #zmiana domyslnej ikonki na tą co sie chce 
     shortcut = shell.CreateShortcut(file_path)  ##otwiera wybrany skrót Windows, czyli ten z file_path
     old_icon = shortcut.IconLocation #znajduje lokalizacje ikonki
     print(old_icon) #chyba to usune
     shortcut.IconLocation = icon_path  #nadpisuje orginalną lokalizacje ikonki na nową      
     shortcut.Save()     #zapisuje
     messagebox.showinfo("Sukces!", "spoko udalo ci sie") #okej


    

def cos():
     if not file_path or not icon_path:
        messagebox.showwarning(
            "Brak danych",
            "Najpierw wybierz skrót i ikonę!"
        )
        return #posrednia ikona dla przucisku 
     change_Shorticon(file_path, icon_path)

button = tkinter.Button(root, text="Click Me", command=cos, width=20, height=5) #przycisk do wykonania funkcji change 
button.pack()
def returning_old_icon():
    file_path = filedialog.askopenfilename(filetypes=[("Pliki lnk", "*.lnk")])
    with open(backup_path, "r" ) as file: # open otwiera plik. pierwsza "rzecz" to lokalizacja pliku który trzeba otworzyć. drugie czyli "r" to tryb odczytu 
            backed = json.load(file) #json.load(file) zamienia to co jest w JSON na obiekt którym jest "backed" - dictonar
    if not file_path:
        return
    
    if file_path not in backed:
            messagebox.showinfo("?", "Plik nie ma Backupa!")
            return

    shortcut = shell.CreateShortcut(file_path)
    shortcut.IconLocation = backed[file_path]
    shortcut.Save()

def returning():
     returning_old_icon()

button = tkinter.Button(root, text="Return Original suko", command=returning, width=20, height=5) #przycisk do wykonania funkcji przywracającej 
button.pack()




shell = win32com.client.Dispatch("WScript.Shell")
root.mainloop() #zeby program sie nie wylaczal po sukcesywnej operacji idk 




