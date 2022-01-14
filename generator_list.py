#========================================IMPORT BIBLIOTEK===============================================================
import openpyxl, win32com.client, pathlib, shutil, os, tkinter as tk, sys, tkinter.font as font
from PyPDF2 import PdfFileMerger, PdfFileReader


#===========================DEFINICJA FUNKCJI KOPIOWANIA I WKLEJANIA====================================================
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


#====================================GŁÓWNA FUNKCJA=====================================================================
def main():
    try:
        wyczyscFolder()


        #DEZAKTYWACJA PRZYCISKU START
        button.configure(state = 'disabled')


        #USUNIĘCIE WSZYSTKICH POPRZEDNICH PLIKÓW WYNIKOWYCH



        #UTWORZENIE FOLDERU DLA PLIKÓW PDF
        textbox.insert(tk.END,"Tworzenie folderu dla plików PDF...")
        window.update()
        os.mkdir('GOTOWE\\PDF')
        textbox.insert(tk.END," GOTOWE\n")
        window.update()


        #WCZYTANIE DANYCH Z GŁÓWNEGO PLIKU
        textbox.insert(tk.END,"Wczytywanie danych z głownego pliku...")
        plikGlowny = openpyxl.load_workbook('GŁÓWNY.xlsm', data_only=True)
        plikGlowny_Parametry = plikGlowny['PARAMETRY']
        konstruktor = plikGlowny_Parametry.cell(row=3,column=3).value
        terminOd = plikGlowny_Parametry.cell(row=4,column=3).value
        terminDo = plikGlowny_Parametry.cell(row=5,column=3).value
        nrZlecenia = plikGlowny_Parametry.cell(row=6,column=3).value
        dataDzisiaj = plikGlowny_Parametry.cell(row=7,column=3).value
        poprawnoscDanych = plikGlowny_Parametry.cell(row=2,column=4).value
        ilePustych = plikGlowny_Parametry.cell(row=2,column=5).value

        #sprawdzenie poprawności danych
        if konstruktor == None or terminOd == None or terminDo == None or nrZlecenia == None or dataDzisiaj == None:
            sys.exit("Niekompletne dane zlecenia w pliku GŁÓWNY.xlsx.")
        elif ilePustych == 100:
            sys.exit("Nie uzupełniono danych list w pliku GŁÓWNY.xlsx.")
        elif poprawnoscDanych != 0:
            sys.exit("Niekompletne dane list materiałowych w pliku GŁÓWNY.xlsx.")
        textbox.insert(tk.END," GOTOWE\n")
        window.update()


        #SPRAWDZENIE LICZBY LIST
        ileList = 0
        while plikGlowny_Parametry.cell(row=ileList+4,column=6).value != None:
            ileList = ileList + 1
        textbox.insert(tk.END,"Liczba list materiałowych do wygenerowania: " + str(ileList) + "\n")
        window.update()

        #TWORZENIE LIST MATEK
        textbox.insert(tk.END,"Tworzene list matek...\n")
        window.update()

        listaZespolow = []
        listaZespolowNumery = []
        for i in range(ileList):
            listaZespolow.append(plikGlowny_Parametry.cell(row=i+4,column=8).value)
            listaZespolowNumery.append(plikGlowny_Parametry.cell(row=i+4,column=9).value)

            #jeśli kolejny numer seryjny jest inny to twórz matkę
            if plikGlowny_Parametry.cell(row=i+5,column=6).value != plikGlowny_Parametry.cell(row=i+4,column=6).value or plikGlowny_Parametry.cell(row=i+5,column=7).value != plikGlowny_Parametry.cell(row=i+4,column=7).value :
                plikMatka = openpyxl.load_workbook('WZORY/MATKA.xlsx')
                plikMatka_ListaZespolow = plikMatka['Lista zespolow']
                #uzupełnienie danych
                plikMatka_ListaZespolow.cell(row=3,column=4).value = konstruktor
                plikMatka_ListaZespolow.cell(row=3,column=11).value = terminOd
                plikMatka_ListaZespolow.cell(row=3,column=13).value = terminDo
                plikMatka_ListaZespolow.cell(row=4,column=6).value = nrZlecenia
                plikMatka_ListaZespolow.cell(row=4,column=11).value = str(plikGlowny_Parametry.cell(row=i+4,column=7).value) + ' NR FABR. ' + str(plikGlowny_Parametry.cell(row=i+4,column=6).value)
                ileZespolow = len(listaZespolow)
                for j in range(ileZespolow):
                    plikMatka_ListaZespolow.cell(row=8+j,column=2).value = 'Z-' + str(plikGlowny_Parametry.cell(row=i+4,column=6).value) + '-' + str(listaZespolowNumery[j])
                    plikMatka_ListaZespolow.cell(row=8+j,column=6).value = str(plikGlowny_Parametry.cell(row=i+4,column=7).value) + ' ' + str(listaZespolow[j])
                    plikMatka_ListaZespolow.cell(row=8+j,column=7).value = str(1)

                #utworzenie folderu dla danej maszyny
                sciezkaFolderuMaszyny = 'GOTOWE\\' + plikGlowny_Parametry.cell(row=i+4,column=7).value + ' [' +plikGlowny_Parametry.cell(row=i+4,column=6).value + ']'
                os.mkdir(sciezkaFolderuMaszyny)

                #zapis pliku matka
                nazwaPlikuMatka = sciezkaFolderuMaszyny + '\\00 ' + plikGlowny_Parametry.cell(row=i+4,column=7).value + ' MATKA [' +plikGlowny_Parametry.cell(row=i+4,column=6).value + '].xlsx'
                plikMatka.save(filename = nazwaPlikuMatka)
                textbox.insert(tk.END,"Utworzono listę matkę EXCEL: " + nazwaPlikuMatka + "\n")
                textbox.see(tk.END)
                window.update()

                #wydruk do pdf pliku matka
                o = win32com.client.Dispatch("Excel.Application")
                o.Visible = False
                wb_path = str(pathlib.Path().absolute()) + '\\' + nazwaPlikuMatka
                wb = o.Workbooks.Open(wb_path)
                path_to_pdf = str(pathlib.Path().absolute()) + '\\GOTOWE\\PDF\\[' + plikGlowny_Parametry.cell(row=i+4,column=6).value + ']' + ' 00 ' +      plikGlowny_Parametry.cell(row=i+4,column=7).value + ' MATKA'
                wb.WorkSheets(1).ExportAsFixedFormat(0,path_to_pdf)
                wb.Close(True)
                textbox.insert(tk.END,"Utworzono listę matkę PDF:   " + "GOTOWE\\PDF\\[" + plikGlowny_Parametry.cell(row=i+4,column=6).value + "]" + " 00 " +      plikGlowny_Parametry.cell(row=i+4,column=7).value + " MATKA.pdf\n")
                textbox.see(tk.END)
                window.update()

                listaZespolow = []
                listaZespolowNumery = []


        #TWORZENIE LIST MATERIAŁOWYCH
        textbox.insert(tk.END,"Tworzenie list materiałowych...\n")
        window.update()
        numerZespolu = 0
        for k in range(ileList):
            #otwarcie pliku wzorcowego
            plikMaterialowa = openpyxl.load_workbook('WZORY/MATERIAŁOWA.xlsx')
            #otwarcie arkusza dane w pliku wzorcowym
            plikMaterialowa_Dane = plikMaterialowa['DANE']
            #uzupełnianie danych w arkuszu DANE
            plikMaterialowa_Dane.cell(row=2,column=5).value = str(plikGlowny_Parametry.cell(row=6,column=3).value) #numer zlecenia
            plikMaterialowa_Dane.cell(row=3,column=5).value = str(plikGlowny_Parametry.cell(row=k+4,column=6).value) #numer maszyny
            plikMaterialowa_Dane.cell(row=4,column=5).value = str(plikGlowny_Parametry.cell(row=k+4,column=7).value) #nazwa maszyny
            plikMaterialowa_Dane.cell(row=4,column=7).value = str(plikGlowny_Parametry.cell(row=k+4,column=8).value) #typ listy
            #plikMaterialowa_Dane.cell(row=4,column=7).value = str(plikGlowny_Parametry.cell(row=k+4,column=9).value) #numer listy
            plikMaterialowa_Dane.cell(row=11,column=2).value = str(plikGlowny_Parametry.cell(row=7,column=3).value) #dzisiejsza data
            plikMaterialowa_Dane.cell(row=12,column=2).value = str(plikGlowny_Parametry.cell(row=3,column=3).value) #konstruktor
            nrSeryjnyMaszyny = '['+plikGlowny_Parametry.cell(row=k+4,column=6).value+']'
            nazwaMaszyny = plikGlowny_Parametry.cell(row=k+4,column=7).value
            typListy = plikGlowny_Parametry.cell(row=k+4,column=8).value

            #nadawanie numeru zespołu - AUTOMATYCZNIE
            #numerZespolu = numerZespolu + 1
            #plikMaterialowa_Dane.cell(row=5,column=5).value = str(plikGlowny_Parametry.cell(row=k+4,column=6).value) + '-' + str(numerZespolu)

            #nadawanie numeru zespołu - NOWA WERSJA
            numerZespolu = str(plikGlowny_Parametry.cell(row=k+4,column=9).value)
            plikMaterialowa_Dane.cell(row=5,column=5).value = str(plikGlowny_Parametry.cell(row=k+4,column=6).value) + '-' + numerZespolu

            #otwarcie arkusza EPLAN w pliku wzorcowym
            plikMaterialowa_Eplan = plikMaterialowa['EPLAN']

            #otwarcie pliku wynikowego z eplan
            wynikowyEplan = str('PLIKI_EPLAN/'+plikGlowny_Parametry.cell(row=k+4,column=10).value)

            #sprawdzenie czy da się otworzyć plik listy wynikowej z EPLAN
            plikWynikowy = openpyxl.load_workbook(wynikowyEplan)

            #otwarcie arkusza EPLAN pliku wynikowego z eplan
            plikWynikowy_Eplan = plikWynikowy['Summarized parts list']

            #sprawdzenie ile jest materiałów na liscie
            ileMaterialow = 0
            while plikWynikowy_Eplan.cell(row=ileMaterialow+9,column=1).value != None:
                ileMaterialow = ileMaterialow + 1

            #skopiowanie listy materiałów z pliku wynikowego z eplan do gotowego pliku listy materiałowej
            zakresMaterialowZrodlo = kopiujZakres(1,9,8,ileMaterialow+9,plikWynikowy_Eplan)
            zakresMaterialowCel = wklejZakres(1,9,8,ileMaterialow+9,plikMaterialowa_Eplan,zakresMaterialowZrodlo)

            #ustawienie zakresu wydruku arkusza WW-A
            plikMaterialowa_WWA = plikMaterialowa['WW-A']
            if ileMaterialow <= 25:
                plikMaterialowa_WWA.print_area = 'A1:L37'
            elif (ileMaterialow > 25) and (ileMaterialow <= 50):
                plikMaterialowa_WWA.print_area = 'A1:L74'
            elif (ileMaterialow > 50) and (ileMaterialow <= 75):
                plikMaterialowa_WWA.print_area = 'A1:L111'
            elif (ileMaterialow > 75) and (ileMaterialow <= 100):
                plikMaterialowa_WWA.print_area = 'A1:A148'

            #zapis gotowego pliku
            sciezkaListyMaterialowej = 'GOTOWE\\' + nazwaMaszyny + ' ' + nrSeryjnyMaszyny + '\\'
            nazwaListyMaterialowej = '0' + str(numerZespolu) + ' ' + nazwaMaszyny + ' ' + typListy + ' ' + nrSeryjnyMaszyny
            plikMaterialowa.save(filename = sciezkaListyMaterialowej + nazwaListyMaterialowej + '.xlsx')
            textbox.insert(tk.END,"Utworzono listę materiałową EXCEL: " + sciezkaListyMaterialowej + nazwaListyMaterialowej + ".xlsx\n")
            textbox.see(tk.END)
            window.update()

            #wydruk do pdf arkusza WW-A
            o = win32com.client.Dispatch("Excel.Application")
            o.Visible = False
            wb_path = str(pathlib.Path().absolute()) + '\\' + sciezkaListyMaterialowej + nazwaListyMaterialowej
            wb = o.Workbooks.Open(wb_path)
            path_to_pdf = str(pathlib.Path().absolute()) + '\\GOTOWE\\PDF\\' + nrSeryjnyMaszyny + ' 0' + str(numerZespolu) + ' ' + nazwaMaszyny + ' ' + typListy
            wb.WorkSheets(6).PageSetup.PrintArea = plikMaterialowa_WWA.print_area
            wb.WorkSheets(6).ExportAsFixedFormat(0,path_to_pdf)
            wb.Close(True)
            textbox.insert(tk.END,"Utworzono listę materiałową PDF:   GOTOWE\\PDF\\" + nrSeryjnyMaszyny + " 0" + str(numerZespolu) + " " + nazwaMaszyny + " " + typListy + "\n")
            textbox.see(tk.END)
            window.update()

            if plikGlowny_Parametry.cell(row=k+4,column=6).value != plikGlowny_Parametry.cell(row=k+5,column=6).value:
                numerZespolu = 0


        #ŁĄCZENIE PDFÓW W JEDEN DO DRUKU
        textbox.insert(tk.END,"Scalanie plików PDF...")
        window.update()
        pdffiles = [f for f in os.listdir('GOTOWE\\PDF')]
        merger = PdfFileMerger()
        for pdf in pdffiles:
            merger.append('GOTOWE\\PDF\\'+pdf)
        merger.write('GOTOWE\\WYDRUK.pdf')
        merger.close()
        textbox.insert(tk.END," GOTOWE\n")
        textbox.see(tk.END)
        window.update()

        #USUNIĘCIE FOLDERU WYNIKOWEGO PDF
        textbox.insert(tk.END,"Usuwanie folderu wynikowego PDF...")
        window.update()
        shutil.rmtree('GOTOWE\\PDF')
        textbox.insert(tk.END," GOTOWE\n")
        window.update()
        button.configure(state = 'active')
        textbox.insert(tk.END,"WSZYSTKIE LISTY ZOSTAŁY UTWORZONE\n")
        textbox.see(tk.END)
        window.update()
        messagebox.showinfo("GOTOWE", "Utworzono wszystkie listy.")

    except:
        e = sys.exc_info()[1]
        textbox.insert(tk.END,"\n### Błąd działania programu: " + str(e) + " ###\n")
        textbox.see(tk.END)
        window.update()
        messagebox.showerror("BŁĄD!",  str(e))
        button.configure(state = 'active')


#==================================================================POWŁOKA GRAFICZNA==========================================================================
#TWORZENIE OKNA GŁÓWNEGO
window = tk.Tk()
window.title( "Generator list materiałowych RADPAK - Piotr Jaroszewski" )
window.geometry("850x500")

#TWORZENIE RAMKI DLA WSZYTKICH ELEMENTÓW

#TWORZENIE OKNA LOGÓW
sb_textbox = tk.Scrollbar(window)
textbox = tk.Text(window, width = 100, height = 5, yscrollcommand = sb_textbox.set)
textbox.pack()
sb_textbox.place(in_ = textbox, relx = 1., rely = 0, relheight = 1.)
textbox.place(x=10, y=50)

#TWORZENIE PRZYISKU START
czcionkaPrzycisk = tk.font.Font(family='Helvetica', size='13')
button = tk.Button(window, text = "START", command = main, font = czcionkaPrzycisk)
button.pack()
button.place(relx=0.5, rely=0.9, anchor='center')


#TWORZENIE TABELI DANYCH
#rows = []
#for i in range(5):
#    cols = []
#    for j in range(4):
#        daneList = tk.Entry(relief=tk.RIDGE)
#        daneList.grid(row=i, column=j)
#        daneList.insert(tk.END, '%d.%d' % (i, j))
#        cols.append(daneList)
#    rows.append(cols)

tk.mainloop()