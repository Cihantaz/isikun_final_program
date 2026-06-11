import pandas as pd
import os
import re
from datetime import datetime

# Dosya yollari
ana_program = r"C:\Users\cihan.tazeoz\Desktop\26 BAHAR FİNAL\UYGULAMA DOSYALARI\26 bahar final _ son.xlsx"
collision_file = r"C:\Users\cihan.tazeoz\Desktop\26 BAHAR FİNAL\UYGULAMA DOSYALARI\final_exam_collisions_14_05_2026_10_48_47.csv"
kisit_file = r"C:\Users\cihan.tazeoz\Desktop\26 BAHAR FİNAL\UYGULAMA DOSYALARI\ders_kisitleri.xlsx"

EXCLUDED_PREFIXES = ('ARCH', 'IMIM', 'GITA', 'INAR')

# 1. Ana programi oku
df_ana = pd.read_excel(ana_program)

# Slot bilgisini parse et (tarih ve saatten G/S cikar)
def parse_gs_from_program(row):
    """Tarih ve saat bilgisinden Gun/Slot cikar"""
    tarih = str(row.get('Sınav Tarihi', '')).strip()
    baslangic = str(row.get('Sınav Başlangıç Saati', '')).strip()
    
    # Tarih -> Gun mapping
    tarih_gun = {
        '08.06.2026': 1, '09.06.2026': 2, '10.06.2026': 3, '11.06.2026': 4,
        '12.06.2026': 5, '15.06.2026': 6, '16.06.2026': 7, '17.06.2026': 8,
        '18.06.2026': 9, '19.06.2026': 10,
    }
    
    # Saat -> Slot mapping
    saat_slot = {
        '08:30': 1, '11:30': 2, '14:30': 3,
    }
    
    gun = tarih_gun.get(tarih, 0)
    slot = saat_slot.get(baslangic, 0)
    return gun, slot

schedule_map = {}
fakulte_map = {}
for _, row in df_ana.iterrows():
    code = str(row['Ders Kodu']).strip().upper()
    fakulte = str(row.get('Fakülte Adı', '')).strip()
    gun, slot = parse_gs_from_program(row)
    if gun > 0 and slot > 0:
        schedule_map[code] = (gun, slot)
        fakulte_map[code] = fakulte

print(f"Ana programda parse edilen ders: {len(schedule_map)}")

# 2. Cakisma dosyasini oku
collisions = pd.read_csv(collision_file)
collision_pairs = {}
for _, row in collisions.iterrows():
    c1 = str(row['Course1']).strip().upper()
    c2 = str(row['Course2']).strip().upper()
    count = row['Common Student Count']
    n1 = c1.replace('İ','I').replace('Ş','S').replace('Ğ','G').replace('Ü','U').replace('Ö','O').replace('Ç','C')
    n2 = c2.replace('İ','I').replace('Ş','S').replace('Ğ','G').replace('Ü','U').replace('Ö','O').replace('Ç','C')
    if not n1.startswith(EXCLUDED_PREFIXES) and not n2.startswith(EXCLUDED_PREFIXES):
        key = tuple(sorted([c1, c2]))
        if key not in collision_pairs or collision_pairs[key] < count:
            collision_pairs[key] = count

# Cakisma analizi
cakisma_var = []
for (c1, c2), count in collision_pairs.items():
    s1 = schedule_map.get(c1)
    s2 = schedule_map.get(c2)
    if s1 and s2 and s1 == s2:
        cakisma_var.append((c1, c2, count, s1[0], s1[1]))

print(f"\nGERCEK CAKISMA (ayni gun+slot): {len(cakisma_var)}")

# 3. Kisit dosyasini oku
kisit = pd.read_excel(kisit_file)
kisit_clean = kisit.dropna(how='all')

def parse_ders_list(val):
    if pd.isna(val):
        return []
    return [d.strip().upper() for d in str(val).split(',') if d.strip()]

# Farkli gun ihlalleri
dd_ihlal = []
for _, row in kisit_clean.iterrows():
    val = row.get('C (Farkli Gun)', row.get('C (Farklı Gün)', ''))
    dersler = parse_ders_list(val)
    for i in range(len(dersler)):
        for j in range(i+1, len(dersler)):
            d1, d2 = dersler[i], dersler[j]
            s1 = schedule_map.get(d1)
            s2 = schedule_map.get(d2)
            if s1 and s2 and s1[0] == s2[0]:
                n1 = d1.replace('İ','I').replace('Ş','S').replace('Ğ','G').replace('Ü','U').replace('Ö','O').replace('Ç','C')
                n2 = d2.replace('İ','I').replace('Ş','S').replace('Ğ','G').replace('Ü','U').replace('Ö','O').replace('Ç','C')
                if not n1.startswith(EXCLUDED_PREFIXES) and not n2.startswith(EXCLUDED_PREFIXES):
                    dd_ihlal.append((d1, d2, s1, s2))

# Farkli slot ihlalleri (ayni gun + ayni slot)
ds_ihlal = []
for _, row in kisit_clean.iterrows():
    val = row.get('D (Farkli Slot)', row.get('D (Farklı Slot)', ''))
    dersler = parse_ders_list(val)
    for i in range(len(dersler)):
        for j in range(i+1, len(dersler)):
            d1, d2 = dersler[i], dersler[j]
            s1 = schedule_map.get(d1)
            s2 = schedule_map.get(d2)
            if s1 and s2 and s1[0] == s2[0] and s1[1] == s2[1]:
                n1 = d1.replace('İ','I').replace('Ş','S').replace('Ğ','G').replace('Ü','U').replace('Ö','O').replace('Ç','C')
                n2 = d2.replace('İ','I').replace('Ş','S').replace('Ğ','G').replace('Ü','U').replace('Ö','O').replace('Ç','C')
                if not n1.startswith(EXCLUDED_PREFIXES) and not n2.startswith(EXCLUDED_PREFIXES):
                    ds_ihlal.append((d1, d2, s1, s2))

print(f"Farkli gun ihlali: {len(dd_ihlal)}")
print(f"Farkli slot ihlali (ayni gun+slot): {len(ds_ihlal)}")

# 4. ITEC4431 analizi
print("\n" + "=" * 60)
print("ITEC4431 ANALIZI")
print("=" * 60)

itec_code = 'ITEC4431'
itec_slot = schedule_map.get(itec_code)
print(f"ITEC4431 mevcut slot: {itec_slot}")

# ITEC4431 ile cakisan dersleri bul
itec_cakisan = []
for (c1, c2), count in collision_pairs.items():
    if itec_code in (c1, c2):
        other = c2 if c1 == itec_code else c1
        other_slot = schedule_map.get(other)
        if other_slot:
            itec_cakisan.append((other, count, other_slot))

print(f"\nITEC4431 ile cakisan ders sayisi: {len(itec_cakisan)}")

# ITEC4431 icin uygun slotlar
n_days = 10
spd = 3
musait_slotlar = []
for g in range(1, n_days+1):
    for s in range(1, spd+1):
        uygun = True
        for other, count, other_slot in itec_cakisan:
            if other_slot == (g, s):
                uygun = False
                break
        if uygun:
            musait_slotlar.append((g, s))

print(f"\nITEC4431 icin uygun slotlar ({len(musait_slotlar)} adet):")
for g, s in musait_slotlar:
    print(f"  G{g}/S{s}")

# 5. Fakulte bazli analiz
print("\n" + "=" * 60)
print("FAKULTE BAZLI ANALIZ")
print("=" * 60)

fakulte_stats = {}
for code, fakulte in fakulte_map.items():
    if fakulte not in fakulte_stats:
        fakulte_stats[fakulte] = {'toplam': 0, 'cakisma': 0, 'dd': 0, 'ds': 0}
    fakulte_stats[fakulte]['toplam'] += 1

# Cakismalari fakulteye gore dagil
for c1, c2, count, g, s in cakisma_var:
    f1 = fakulte_map.get(c1, 'Bilinmiyor')
    f2 = fakulte_map.get(c2, 'Bilinmiyor')
    if f1 in fakulte_stats:
        fakulte_stats[f1]['cakisma'] += 1
    if f2 in fakulte_stats and f2 != f1:
        fakulte_stats[f2]['cakisma'] += 1

# Ihlalleri fakulteye gore dagil
for d1, d2, s1, s2 in dd_ihlal:
    f1 = fakulte_map.get(d1, 'Bilinmiyor')
    f2 = fakulte_map.get(d2, 'Bilinmiyor')
    if f1 in fakulte_stats:
        fakulte_stats[f1]['dd'] += 1
    if f2 in fakulte_stats and f2 != f1:
        fakulte_stats[f2]['dd'] += 1

for d1, d2, s1, s2 in ds_ihlal:
    f1 = fakulte_map.get(d1, 'Bilinmiyor')
    f2 = fakulte_map.get(d2, 'Bilinmiyor')
    if f1 in fakulte_stats:
        fakulte_stats[f1]['ds'] += 1
    if f2 in fakulte_stats and f2 != f1:
        fakulte_stats[f2]['ds'] += 1

for fakulte, stats in sorted(fakulte_stats.items()):
    print(f"\n{fakulte}:")
    print(f"  Toplam ders: {stats['toplam']}")
    print(f"  Cakisma ihlali: {stats['cakisma']}")
    print(f"  Farkli gun ihlali: {stats['dd']}")
    print(f"  Farkli slot ihlali: {stats['ds']}")

# Eksik dersler
print("\n" + "=" * 60)
print("EKSIK DERSLER (Ana programda olmayan ama cakismada olan)")
print("=" * 60)

missing = set()
for (c1, c2), count in collision_pairs.items():
    if c1 not in schedule_map:
        missing.add(c1)
    if c2 not in schedule_map:
        missing.add(c2)

print(f"Toplam eksik ders: {len(missing)}")
for c in sorted(missing)[:30]:
    print(f"  {c}")
