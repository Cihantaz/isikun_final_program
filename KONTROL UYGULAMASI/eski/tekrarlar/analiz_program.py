import pandas as pd

program_file = r"C:\Users\cihan.tazeoz\Desktop\26 BAHAR FİNAL\UYGULAMA DOSYALARI\2026 Bahar Final Programı.xlsx"

xl = pd.ExcelFile(program_file)
print("Sayfalar:", xl.sheet_names)

for sheet in xl.sheet_names:
    print(f"\n{'='*60}")
    print(f"SAYFA: {sheet}")
    print('='*60)
    df = pd.read_excel(program_file, sheet_name=sheet)
    print(f"Satır: {len(df)}, Kolon: {list(df.columns)}")
    print(df.head(10))
