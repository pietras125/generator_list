#========================================IMPORT BIBLIOTEK===============================================================
import openpyxl, win32com.client, pathlib, shutil, os, tkinter as tk, sys, tkinter.font as font
from PyPDF2 import PdfFileMerger, PdfFileReader
from tkinter import messagebox
from tksheet import Sheet


#================================================ DEFINICJA FUNKCJI ====================================================
def kopiujZakres(poczatekKolumna, poczatekWiersz, koniecKolumna, koniecWiersz, sheet):
    zakresZaznaczenia = []
    for i in range(poczatekWiersz,koniecWiersz + 1,1):
        wybranyWiersz = []
        for j in range(poczatekKolumna,koniecKolumna+1,1):
            wybranyWiersz.append(sheet.cell(row = i, column = j).value)
        zakresZaznaczenia.append(wybranyWiersz)
    return zakresZaznaczenia

def wklejZakres(poczatekKolumna, poczatekWiersz, koniecKolumna, koniecWiersz, arkuszDoWklejenia, skopiowaneDane):
    ileWierszy = 0
    for i in range(poczatekWiersz,koniecWiersz+1,1):
        ileKolumn = 0
        for j in range(poczatekKolumna,koniecKolumna+1,1):
            arkuszDoWklejenia.cell(row = i, column = j).value = skopiowaneDane[ileWierszy][ileKolumn]
            ileKolumn += 1
        ileWierszy += 1

def wyczyscFolder():
    textbox.insert(tk.END,"Usuwanie poprzednich plików...")
    window.update()
    #jeśli istnieje plik PDF to usuń
    try:
        os.remove("GOTOWE\\WYDRUK.pdf")
    except OSError:
        pass
    #usuwanie wszystkiego w folderze
    folder = 'GOTOWE'
    for filename in os.listdir(folder):
        file_path = os.path.join(folder, filename)
        shutil.rmtree(file_path)
    textbox.insert(tk.END," GOTOWE\n")
    window.update()

def utworzFolderPDF():
    textbox.insert(tk.END,"Tworzenie folderu dla plików PDF...")
    window.update()
    os.mkdir('GOTOWE\\PDF')
    textbox.insert(tk.END," GOTOWE\n")
    window.update()

def pobierzListePlikowDoStruktury(folder):
    listaPlikow = []
    podzielonaLista = []
    listaPlikowWynikowa = []

    for filename in os.listdir(folder):
        listaPlikow.append(filename)

    for i in range(len(listaPlikow)):
        podzielonaLista = listaPlikow[i].split('_')
        numerSeryjny = podzielonaLista[0]
        nazwaMaszyny = podzielonaLista[1]
        numerListy = podzielonaLista[2]
        nazwaListyZRozszerzeniem = podzielonaLista[3]
        nazwaListy = nazwaListyZRozszerzeniem[:-5]
        listaPlikowWynikowa.append([numerSeryjny, nazwaMaszyny, numerListy, nazwaListy, listaPlikow[i]])
    return listaPlikowWynikowa

def wyswietlTabeleListMaterialowych(listaDoWyswietlenia, gdzieWyswietlic):
    oknoTabela = tk.Frame(gdzieWyswietlic, bg="green")
    rows = []
    for i in range(len(listaDoWyswietlenia)):
        cols = []
        for j in range(4):
            daneList = tk.Entry(oknoTabela, relief=tk.RIDGE)
            daneList.place(relwidth=0.25)
            daneList.grid(row=i, column=j)
            daneList.insert(tk.END, listaDoWyswietlenia[i][j])
            cols.append(daneList)
        daneList.place()
        rows.append(cols)
    oknoTabela.place(relx=0.5, y=5, anchor="n")
    oknoTabela.pack()

"""def stworzStruktureZDanychTabeli(listaDoUtworzeniaPlikow):
    rows = []
    for i in range(len(listaDoUtworzeniaPlikow)):
        cols = []
        for j in range(4):
            daneList.grid(row=i, column=j)
            listaDoUtworzeniaPlikow[i][j] = daneList.insert(tk.END, listaDoUtworzeniaPlikow[i][j])
            cols.append(daneList)
        daneList.place()
        rows.append(cols)
"""
#====================================GŁÓWNA FUNKCJA=====================================================================
def main():
    try:
        #DEZAKTYWACJA PRZYCISKU START
        button.configure(state = 'disabled')

        #USUNIĘCIE WSZYSTKICH POPRZEDNICH PLIKÓW WYNIKOWYCH
        wyczyscFolder()

        #UTWORZENIE FOLDERU DLA PLIKÓW PDF
        utworzFolderPDF()

        #SPRAWDZENIE POPRAWNOŚCI DANYCH

        #WYPISANIE LICZBY LIST
        ileList = sheet.get_total_rows()
        textbox.insert(tk.END,"Liczba list materiałowych do wygenerowania: " + str(ileList) + "\n")
        window.update()

        #WCZYTANIE DANYCH Z INPUTÓW
        konstruktor = InputKonstruktor.get()
        terminOd = InputTerminOd.get()
        terminDo = InputTerminDo.get()
        nrZlecenia = InputTerminOd.get()
        nrZlecenia = InputNumerZlecenia.get()

        #WCZYTANIE DANYCH Z TABELI PLIKÓW
        HEHEHE = Sheet.get_sheet_data(sheet)


    except:
        e = sys.exc_info()[1]
        textbox.insert(tk.END,"\n### Błąd działania programu: " + str(e) + " ###\n")
        textbox.see(tk.END)
        window.update()
        messagebox.showerror("BŁĄD!",  str(e))
        button.configure(state = 'active')


#================================================================= ZAWSZE URUCHAMIANE =======================================================================
#TWORZENIE OKNA GŁÓWNEGO
window = tk.Tk()
window.title( "Generator list materiałowych RADPAK - Piotr Jaroszewski" )
window.geometry("850x650")

#TWORZENIE RAMEK
ramkaGorna = tk.Frame(window)
ramkaGorna.pack()
ramkaDolna = tk.Frame(window)
ramkaDolna.pack()
ramka1 = tk.Frame(ramkaGorna, width=340, height=130, highlightbackground="black", highlightthickness=1)
ramka1.pack(side=tk.LEFT, expand=True)
ramka2 = tk.Frame(ramkaGorna, width=400, height=100)
ramka2.pack(side=tk.LEFT, expand=True)
ramka3 = tk.Frame(ramkaDolna, width=850, height=150)
ramka3.pack(side=tk.TOP, expand=True)

#TWORZENIE PÓL DANYCH
LabelKonstruktor = tk.Label(ramka1, text = "Konstruktor:")
LabelKonstruktor.place(x=10, y=10)
InputKonstruktor = tk.Entry(ramka1, bd = 2)
InputKonstruktor.place(x=105, y=10)

LabelTerminOd = tk.Label(ramka1, text = "Termin od:")
LabelTerminOd.place(x=10, y=40)
InputTerminOd = tk.Entry(ramka1, bd = 2)
InputTerminOd.place(x=105, y=40)

LabelTerminDo = tk.Label(ramka1, text = "Termin do:")
LabelTerminDo.place(x=10, y=70)
InputTerminDo = tk.Entry(ramka1, bd = 2)
InputTerminDo.place(x=105, y=70)

LabelNumerZlecenia = tk.Label(ramka1, text = "Numer zlecenia:")
LabelNumerZlecenia.place(x=10, y=100)
InputNumerZlecenia = tk.Entry(ramka1, bd = 2)
InputNumerZlecenia.place(x=105, y=100)

#TWORZENIE PRZYISKU START
czcionkaPrzycisk = tk.font.Font(family='Helvetica', size='13')
button = tk.Button(ramka1, text = "START", command = main, font = czcionkaPrzycisk)
button.pack()
button.place(anchor='center', rely=0.5, x=280)

#TWORZENIE OKNA LOGÓW
sb_textbox = tk.Scrollbar()
textbox = tk.Text(ramka3, width = 103, height = 8, yscrollcommand = sb_textbox.set)
sb_textbox.place(in_ = textbox, relx = 0.98, rely = 0, relheight = 1)
textbox.pack()
textbox.place(in_ = ramka3, relx=0.5, y=5, anchor="n")


#POBRANIE LISTY PLIKÓW Z FOLDERU "PLIKI_EPLAN"
listaMaszynWczytana = pobierzListePlikowDoStruktury('PLIKI_EPLAN')


#WYŚWIETLANIE TABELI DANYCH
tuplePliki = tuple([listaMaszynWczytana])
sheet = Sheet(
                ramka2,
                data = tuplePliki[0],
                #width = 500,
                headers = ['Numer seryjny','Nazwa maszyny','Nr listy','Nazwa listy','Plik'],
                set_all_heights_and_widths = True
            )
sheet.enable_bindings()
sheet.pack()

#===============================================================================================================================================================

tk.mainloop()
