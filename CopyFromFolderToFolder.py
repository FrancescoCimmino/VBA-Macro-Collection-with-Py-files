# CopyFromFolderToFolder
import os
import shutil
from tkinter import Tk
import pandas as pd
from tkinter.filedialog import askdirectory, askopenfilename
from tqdm import tqdm #opzionale, per la barra di avanzamento

# ------------------ SELEZIONE FILE EXCEL ------------------ #
def seleziona_file_excel():
    Tk().withdraw() # Nasconde la finestra principale di Tkinter
    return askopenfilename(title="Seleziona file Excel", filetypes=[("Excel files", "*.xlsx *.xls")])

# ------------------ SELEZIONE CARTELLA ------------------ #
def seleziona_cartella(titolo):
    Tk().withdraw()
    return askdirectory(title=titolo)
Cartella_sorgente = seleziona_cartella("Seleziona cartella sorgente")
Cartella_destinazione = seleziona_cartella("Seleziona cartella destinazione")

# ------------------ LETTURA FILE EXCEL ------------------ #
percorso_excel = seleziona_file_excel()
df = pd.read_excel(percorso_excel, sheet_name="Sheet1", usecols=[0])
df.dropna(inplace=True) # Rimuove eventuali righe vuote
nomi_file = df.iloc[:, 0].tolist() # Estrae i nomi dei file dalla prima colonna

# ------------------ LETTURA RIGHE-FILES ------------------ #
copiati = 0
for nome_file in tqdm(nomi_file):
    if os.path.exists(os.path.join(Cartella_sorgente, nome_file)):
        percorso_sorgente = os.path.join(Cartella_sorgente, nome_file)
        percorso_destinazione = os.path.join(Cartella_destinazione, nome_file)
        shutil.copy2(percorso_sorgente, percorso_destinazione)
        copiati += 1
    else: 
        print(f"File da copiare non trovato: {nome_file}")
print(f"Copia completata, file copiati: {copiati}, file non trovati: {len(nomi_file) - copiati} ")
