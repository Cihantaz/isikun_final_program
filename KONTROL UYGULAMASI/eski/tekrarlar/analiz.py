import pandas as pd
import sys

# Dosya yolları
collision_file = r"C:\Users\cihan.tazeoz\Desktop\26 BAHAR FİNAL\UYGULAMA DOSYALARI\final_exam_collisions_14_05_2026_10_48_47.csv"
schedule_file = r"C:\Users\cihan.tazeoz\Desktop\26 BAHAR FİNAL\UYGULAMA DOSYALARI\schedule (5).xlsx"
kisit_file = r"C:\Users\cihan.tazeoz\Desktop\26 BAHAR FİNAL\UYGULAMA DOSYALARI\ders_kisitleri.xlsx"

print("=" * 80)
print("1. ÇAKIŞMA DOSYASI ANALİZİ")
print("=" * 80)
collisions = pd.read_csv(collision_file)
print(f"Toplam satır: {len(collisions)}")
print(f"Kolonlar: {list(collisions.columns)}")
print("\nİlk 10 satır:")
print(collisions.head(10))

print("\n" + "=" * 80)
print("2. SCHEDULE DOSYASI ANALİZİ")
print("=" * 80)
schedule = pd.read_excel(schedule_file)
print(f"Toplam satır: {len(schedule)}")
print(f"Kolonlar: {list(schedule.columns)}")
print("\nİlk 10 satır:")
print(schedule.head(10))

print("\n" + "=" * 80)
print("3. KISIT DOSYASI ANALİZİ")
print("=" * 80)
kisit = pd.read_excel(kisit_file)
print(f"Toplam satır: {len(kisit)}")
print(f"Kolonlar: {list(kisit.columns)}")
print("\nİlk 10 satır:")
print(kisit.head(10))
