
import pandas as pd
import os

path = "Articulos.xlsx"
if not os.path.exists(path):
    print(f"File {path} not found")
else:
    try:
        df = pd.read_excel(path)
        print("Columns:", df.columns.tolist())
        print("First 3 rows:")
        print(df.head(3))
    except Exception as e:
        print(f"Error reading {path}: {e}")
