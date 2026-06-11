import pandas as pd
import os
import glob

# Dosya yollari
ana_program = r"C:\Users\cihan.tazeoz\Desktop\26 BAHAR FİNAL\UYGULAMA DOSYALARI\26 bahar final _ son.xlsx"
collision_file = r"C:\Users\cihan.tazeoz\Desktop\26 BAHAR FİNAL\UYGULAMA DOSYALARI\final_exam_collisions_14_05_2026_10_48_47.csv"
kisit_file = r"C:\Users\cihan.tazeoz\Desktop\26 BAHAR FİNAL\UYGULAMA DOSYALARI\ders_kisitleri.xlsx"
base_dir = r"C:\Users\cihan.tazeoz\Desktop\26 BAHAR FİNAL"

print("=" * 90)
print("ANA PROGRAM ANALIZI")
print("=" * 90)

# 1. Ana programi oku
df_ana = pd.read_excel(ana_program)
print(f"\nAna program dosyasi: {os.path.basename(ana_program)}")
print(f"Toplam satir: {len(df_ana)}")
print(f"Kolonlar: {list(df_ana.columns)}")
print("\nIlk 5 satir:")
print(df_ana.head())

# 2. Diger bolum programlarini listele
print("\n" + "=" * 90)
print("DIGER BOLUM PROGRAM DOSYALARI")
print("=" * 90)

excel_files = []
for f in glob.glob(os.path.join(base_dir, "*.xlsx")):
    fname = os.path.basename(f)
    if not fname.startswith("~$") and "25GÜZ" not in f:
        excel_files.append(f)

print(f"\nBulunan Excel dosyalari: {len(excel_files)}")
for f in excel_files:
    print(f"  - {os.path.basename(f)}")
