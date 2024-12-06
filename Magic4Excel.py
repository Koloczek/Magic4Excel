from tkinter import Tk
from tkinter.filedialog import askopenfilename, asksaveasfilename
import pandas as pd
import time
import sys
import ctypes



print("Witaj Pastafarianinie!")
time.sleep(2)
print("Jeśli również masz w ciul powtarzających się rekordów w excelu...")
time.sleep(2)
print("i chcesz je przekopiować do innego arkusza...")
time.sleep(2)
print("To ten program jest dla ciebie!!!!!")
time.sleep(2)
print()
print("Wybierz plik gdzie znajdują się powtórzone rekordy: ")
time.sleep(2)



def choose_file():
    Tk().withdraw()
    file = askopenfilename(
        title="Wybierz plik Excel",
        filetype=[("Plik Excel", "*.xlsx *.xls")]
    )
    return file

def save_file():
    Tk().withdraw()
    file = asksaveasfilename(
        title="Zapisz wyniki jako",
        defaultextension=".xlsx",
        filetypes=[("Pliki Excel", "*.xlsx")]
    )
    return file

file = choose_file()
if not file:
    print("Jeśli nie chcesz to nie!")
    time.sleep(1)
    exit()

sheet = "Arkusz1"
df = pd.read_excel(file, sheet_name=sheet)

column= input("Podaj nazwę kolumny do sprawdzenia: ")
if column not in df.columns:
    print(f"Kolumna '{column}' nie istnieje w tym akruszu!")
    exit()

notduplicates = df[~df[column].duplicated(keep=False)]

print("Wybierz miejsce do zapisu i pomyśl jak nazwać plik!")
result = save_file()

notduplicates.to_excel(result, index=False)
time.sleep(2)
print(f"Zapisano do pliku '{result}'.")
time.sleep(2)
