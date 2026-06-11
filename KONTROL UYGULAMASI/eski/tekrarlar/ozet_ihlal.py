import pandas as pd
from collections import defaultdict

collision_file = r"C:\Users\cihan.tazeoz\Desktop\26 BAHAR FİNAL\UYGULAMA DOSYALARI\final_exam_collisions_14_05_2026_10_48_47.csv"
schedule_file = r"C:\Users\cihan.tazeoz\Desktop\26 BAHAR FİNAL\UYGULAMA DOSYALARI\schedule (5).xlsx"
kisit_file = r"C:\Users\cihan.tazeoz\Desktop\26 BAHAR FİNAL\UYGULAMA DOSYALARI\ders_kisitleri.xlsx"

collisions = pd.read_csv(collision_file)
schedule = pd.read_excel(schedule_file)
kisit = pd.read_excel(kisit_file)

EXCLUDED_PREFIXES = ('ARCH', 'IMIM', 'GITA', 'INAR')

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

schedule_map = {}
for _, row in schedule.iterrows():
    ders_kodu = str(row['Ders Kodu']).strip().upper()
    slot_val = parse_slot(row['Slotlar'])
    if slot_val:
        schedule_map[ders_kodu] = slot_val

def parse_ders_list(val):
    if pd.isna(val):
        return []
    return [d.strip().upper() for d in str(val).split(',') if d.strip()]

kisit_clean = kisit.dropna(how='all')
farkli_gun_kisit = []
farkli_slot_kisit = []

for _, row in kisit_clean.iterrows():
    farkli_gun = parse_ders_list(row.get('C (Farkli Gun)', row.get('C (Farklı Gün)', '')))
    for i in range(len(farkli_gun)):
        for j in range(i+1, len(farkli_gun)):
            farkli_gun_kisit.append((farkli_gun[i], farkli_gun[j]))
    
    farkli_slot = parse_ders_list(row.get('D (Farkli Slot)', row.get('D (Farklı Slot)', '')))
    for i in range(len(farkli_slot)):
        for j in range(i+1, len(farkli_slot)):
            farkli_slot_kisit.append((farkli_slot[i], farkli_slot[j]))

farkli_gun_ihlal = []
for d1, d2 in farkli_gun_kisit:
    s1 = schedule_map.get(d1)
    s2 = schedule_map.get(d2)
    if s1 and s2 and s1[0] == s2[0]:
        farkli_gun_ihlal.append((d1, d2, s1, s2))

farkli_slot_ihlal = []
for d1, d2 in farkli_slot_kisit:
    s1 = schedule_map.get(d1)
    s2 = schedule_map.get(d2)
    if s1 and s2 and s1[1] == s2[1]:
        farkli_slot_ihlal.append((d1, d2, s1, s2))

# Tekrar edenleri kaldir
fg_unique = []
seen_fg = set()
for d1, d2, s1, s2 in farkli_gun_ihlal:
    key = tuple(sorted([d1, d2]))
    if key not in seen_fg:
        seen_fg.add(key)
        fg_unique.append((d1, d2, s1, s2))

fs_unique = []
seen_fs = set()
for d1, d2, s1, s2 in farkli_slot_ihlal:
    key = tuple(sorted([d1, d2]))
    if key not in seen_fs:
        seen_fs.add(key)
        fs_unique.append((d1, d2, s1, s2))

# Mimarlik haric
fg_haric = []
for d1, d2, s1, s2 in fg_unique:
    n1 = d1.replace('İ','I').replace('Ş','S').replace('Ğ','G').replace('Ü','U').replace('Ö','O').replace('Ç','C')
    n2 = d2.replace('İ','I').replace('Ş','S').replace('Ğ','G').replace('Ü','U').replace('Ö','O').replace('Ç','C')
    if not n1.startswith(EXCLUDED_PREFIXES) and not n2.startswith(EXCLUDED_PREFIXES):
        fg_haric.append((d1, d2, s1, s2))

fs_haric = []
for d1, d2, s1, s2 in fs_unique:
    n1 = d1.replace('İ','I').replace('Ş','S').replace('Ğ','G').replace('Ü','U').replace('Ö','O').replace('Ç','C')
    n2 = d2.replace('İ','I').replace('Ş','S').replace('Ğ','G').replace('Ü','U').replace('Ö','O').replace('Ç','C')
    if not n1.startswith(EXCLUDED_PREFIXES) and not n2.startswith(EXCLUDED_PREFIXES):
        fs_haric.append((d1, d2, s1, s2))

# Farkli slot ihlallerini kategorize et
fs_ayni_gun = [x for x in fs_haric if x[2][0] == x[3][0]]
fs_farkli_gun = [x for x in fs_haric if x[2][0] != x[3][0]]

print("=" * 80)
print("A. FARKLI GUN KISITI IHLALLERI (Tekil)")
print("=" * 80)
print(f"Toplam tekil ihlal: {len(fg_haric)}\n")
for i, (d1, d2, s1, s2) in enumerate(fg_haric, 1):
    print(f"{i:2d}. {d1:12s} (G{s1[0]}/S{s1[1]})  -  {d2:12s} (G{s2[0]}/S{s2[1]})  -> AYNI GUN: G{s1[0]}")

print("\n" + "=" * 80)
print("B. FARKLI SLOT KISITI IHLALLERI (Tekil)")
print("=" * 80)
print(f"Toplam tekil ihlal: {len(fs_haric)}")
print(f"  - Ayni gun / Ayni slot (TEHLIKELI): {len(fs_ayni_gun)}")
print(f"  - Farkli gun / Ayni slot (Riskli ama ogrenci cakismasi yok): {len(fs_farkli_gun)}\n")

print("--- B1. AYNI GUN + AYNI SLOT (Ogrenci Cakismasi Riski olanlar) ---")
for i, (d1, d2, s1, s2) in enumerate(fs_ayni_gun, 1):
    print(f"{i:2d}. {d1:12s} (G{s1[0]}/S{s1[1]})  -  {d2:12s} (G{s2[0]}/S{s2[1]})")

print("\n--- B2. FARKLI GUN + AYNI SLOT (Ogrenci Cakismasi YOK) ---")
for i, (d1, d2, s1, s2) in enumerate(fs_farkli_gun, 1):
    print(f"{i:2d}. {d1:12s} (G{s1[0]}/S{s1[1]})  -  {d2:12s} (G{s2[0]}/S{s2[1]})")

# Eksik dersler
missing_courses = set()
for _, row in collisions.iterrows():
    c1 = str(row['Course1']).strip().upper()
    c2 = str(row['Course2']).strip().upper()
    n1 = c1.replace('İ','I').replace('Ş','S').replace('Ğ','G').replace('Ü','U').replace('Ö','O').replace('Ç','C')
    n2 = c2.replace('İ','I').replace('Ş','S').replace('Ğ','G').replace('Ü','U').replace('Ö','O').replace('Ç','C')
    if not n1.startswith(EXCLUDED_PREFIXES) and not n2.startswith(EXCLUDED_PREFIXES):
        if c1 not in schedule_map:
            missing_courses.add(c1)
        if c2 not in schedule_map:
            missing_courses.add(c2)

core = sorted([c for c in missing_courses if c.startswith('CORE')])
proje = sorted([c for c in missing_courses if any(s in c for s in ['3910','3920','4901','4902','4910','4912','4920'])])
diger = sorted([c for c in missing_courses if c not in core and c not in proje])

print("\n" + "=" * 80)
print("C. SCHEDULE'DA OLMAYAN DERSLER")
print("=" * 80)
print(f"Toplam eksik ders: {len(missing_courses)}")
print(f"  [A] CORE dersleri          : {len(core)} adet")
print(f"  [B] Proje/Seminer dersleri : {len(proje)} adet")
print(f"  [C] Diger dersler          : {len(diger)} adet")
