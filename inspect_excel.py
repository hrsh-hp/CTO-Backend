import pandas as pd

file_path = '/home/harsh2006/Projects/Python/cto-app/cto/DATALOGGER AND RTU SHEET 2025.xlsx'
try:
    df = pd.read_excel(file_path, header=None) # Read without header first to see layout
    print(df.head(20).to_markdown(index=False, numalign="left", stralign="left"))
    print("\nColumns:")
    print(df.columns.tolist())
except Exception as e:
    print(f"Error reading excel: {e}")
