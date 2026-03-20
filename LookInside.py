import pandas as pd, os
from tkinter import Tk
from tkinter.filedialog import askdirectory, asksaveasfilename
from tqdm import tqdm 

# selezione cartella con tk solito
def Selected_Folder():
  Tk().withdraw()
  return askdirectory(title="Select Folder")
Selected_Folder = Selected_Folder()
all_rows = []

files_to_skip = ["thumbs.db", "desktop.ini", ".ds_store"]
for root, dirs, files in os.walk(Selected_Folder, topdown=True):
    
     # print(root.split("\\")[len(Selected_Folder.split("\\")):])
    for file in files:
        #oswalk returns a tuple with "dirpath", "dirnames", "filenames"
        
        Levels = root.split("\\")[len(Selected_Folder.split("\\")):]
        if file.lower() in files_to_skip:
            continue
        row = {"Level_1": Levels[0] if len(Levels) > 0 else None,
            "Level_2": Levels[1] if len(Levels) > 1 else None,
            "Level_3": Levels[2] if len(Levels) > 2 else None,
            "Level_4": Levels[3] if len(Levels) > 3 else None,
            "Level_5": Levels[4] if len(Levels) > 4 else None,
            "Level_6": Levels[5] if len(Levels) > 5 else None,
            "Level_7": Levels[6] if len(Levels) > 6 else None,
            "Level_8": Levels[7] if len(Levels) > 7 else None,
            "Level_9": Levels[8] if len(Levels) > 8 else None,
            "Level_10": Levels[9] if len(Levels) > 9 else None,
            "Files": file}
        all_rows.append(row)
df = pd.DataFrame.dropna(pd.DataFrame(all_rows))

def save_excel(df):
    Tk().withdraw()
    file_path = asksaveasfilename(
        
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx")],
        title="Salva file Excel",
        initialfile="LookInside_Results.xlsx"
        )
    if not file_path:
            print("Saving cancelled.")
            return
    df.to_excel(file_path, index=False)
save_excel(df)
