
import zipfile
import pandas as pd
import glob

files = sorted(glob.glob("BhavCopy*.zip"))
if files:
    with zipfile.ZipFile(files[-1], 'r') as z:
        csv_files = [f for f in z.namelist() if f.lower().endswith('.csv')]
        if csv_files:
            df = pd.read_csv(z.open(csv_files[0]), nrows=5)
            print("Columns:", df.columns.tolist())
else:
    print("No BhavCopy zip files found.")
