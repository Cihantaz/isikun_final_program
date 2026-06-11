import pandas as pd
from collections import defaultdict

schedule_file = r"C:\Users\cihan.tazeoz\Desktop\26 BAHAR FİNAL\UYGULAMA DOSYALARI\schedule (5).xlsx"
new_file = r"C:\Users\cihan.tazeoz\Desktop\26 BAHAR FİNAL\UYGULAMA DOSYALARI\YENILENMIS_PROGRAM_2026_BAHAR_FINAL.xlsx"
kisit_file = r"C:\Users\cihan.tazeoz\Desktop\26 BAHAR FİNAL\UYGULAMA DOSYALARI\ders_kisitleri.xlsx"

def parse_slot(slot_str):
    if pd.isna(slot_str):
        return None
    slot_str = str(slot_str).strip().upper()
    parts = slot_str.split('/')
    if len(parts) != 2:
        return None
    gun = int(parts[0].replace('G', ''))
    slot = int(parts[1].replace('S', ''))
    return (gun, slot)

# Yeni programi oku
df_new = pd.read_excel(new_file)
new_map = {}
for _, row in df_new.iterrows():
    code = str(row['Ders Kodu']).strip().upper()
    slot = parse_slot(row['Yeni Slot'])
    if slot:
        new_map[code] = slot

# Kisit dosyasini oku
kisit = pd.read_excel(kisit_file)
kisit_clean = kisit.dropna(how='all')

def parse_ders_list(val):
    if pd.isna(val):
        return []
    return [d.strip().upper() for d in str(val).split(',') if d.strip()]

# Farkli gun kisitlarini kontrol et
print("=" * 90)
print("KALAN DIFFDAY IHLALLERI (Yeni Program)")
print("=" * 90)

dd_count = 0
for _, row in kisit_clean.iterrows():
    val = row.get('C (Farkli Gun)', row.get('C (Farklı Gün)', ''))
    dersler = parse_ders_list(val)
    for i in range(len(dersler)):
        for j in range(i+1, len(dersler)):
            d1, d2 = dersler[i], dersler[j]
            s1 = new_map.get(d1)
            s2 = new_map.get(d2)
            if s1 and s2 and s1[0] == s2[0]:
                dd_count += 1
                print(f"{dd_count:2d}. {d1:12s} ({s1[0]}/S{s1[1]}) - {d2:12s} ({s2[0]}/S{s2[1]})")

print(f"\nToplam DiffDay ihlali: {dd_count}")

# Farkli slot kisitlarini kontrol et (sadece ayni gun + ayni slot)
print("\n" + "=" * 90)
print("KALAN DIFFSLOT IHLALLERI (Ayni gun + Ayni slot)")
print("=" * 90)

ds_count = 0
for _, row in kisit_clean.iterrows():
    val = row.get('D (Farkli Slot)', row.get('D (Farklı Slot)', ''))
    dersler = parse_ders_list(val)
    for i in range(len(dersler)):
        for j in range(i+1, len(dersler)):
            d1, d2 = dersler[i], dersler[j]
            s1 = new_map.get(d1)
            s2 = new_map.get(d2)
            if s1 and s2 and s1[0] == s2[0] and s1[1] == s2[1]:
                ds_count += 1
                print(f"{ds_count:2d}. {d1:12s} ({s1[0]}/S{s1[1]}) - {d2:12s} ({s2[0]}/S{s2[1]})")

print(f"\nToplam DiffSlot ihlali (ayni gun+slot): {ds_count}")

# Degisen dersleri listele
df_old = pd.read_excel(schedule_file)
old_map = {}
for _, row in df_old.iterrows():
    code = str(row['Ders Kodu']).strip().upper()
    slot = parse_slot(row['Slotlar'])
    if slot:
        old_map[code] = slot

print("\n" + "=" * 90)
print("DEGISEN DERSLER (Eski -> Yeni)")
print("=" * 90)
for _, row in df_new.iterrows():
    code = str(row['Ders Kodu']).strip().upper()
    if row['Değişti'] == 'EVET':
        old = old_map.get(code)
        new = new_map.get(code)
        print(f"  {code:12s}: G{old[0]}/S{old[1]} -> G{new[0]}/S{new[1]}")
