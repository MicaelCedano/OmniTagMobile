
import pandas as pd

try:
    df = pd.read_excel("plantilla_compra_iphone (7).xlsx")
    print("Columns:", df.columns.tolist())
    print("First few rows:\n", df.head())
except Exception as e:
    print(e)
