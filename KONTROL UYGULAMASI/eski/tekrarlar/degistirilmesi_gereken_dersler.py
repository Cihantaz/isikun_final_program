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

# Cakisma ciftlerini oku
collision_dict = {}
for _, row in collisions.iterrows():
    c1 = str(row['Course1']).strip().upper()
    c2 = str(row['Course2']).strip().upper()
    count = row['Common Student Count']
    n1 = c1.replace('İ','I').replace('Ş','S').replace('Ğ','G').replace('Ü','U').replace('Ö','O').replace('Ç','C')
    n2 = c2.replace('İ','I').replace('Ş','S').replace('Ğ','G').replace('Ü','U').replace('Ö','O').replace('Ç','C')
    if not n1.startswith(EXCLUDED_PREFIXES) and not n2.startswith(EXCLUDED_PREFIXES):
        key = tuple(sorted([c1, c2]))
        if key not in collision_dict or collision_dict[key] < count:
            collision_dict[key] = count

# Farkli gun kisitlarini oku (C kolonu)
farkli_gun_ciftler = []
for _, row in kisit_clean.iterrows():
    val = row.get('C (Farkli Gun)', row.get('C (Farklı Gün)', ''))
    dersler = parse_ders_list(val)
    for i in range(len(dersler)):
        for j in range(i+1, len(dersler)):
            farkli_gun_ciftler.append((dersler[i], dersler[j]))

# Farkli slot kisitlarini oku (D kolonu)
farkli_slot_ciftler = []
for _, row in kisit_clean.iterrows():
    val = row.get('D (Farkli Slot)', row.get('D (Farklı Slot)', ''))
    dersler = parse_ders_list(val)
    for i in range(len(dersler)):
        for j in range(i+1, len(dersler)):
            farkli_slot_ciftler.append((dersler[i], dersler[j]))

def unique_pairs(pairs):
    seen = set()
    result = []
    for d1, d2 in pairs:
        key = tuple(sorted([d1, d2]))
        if key not in seen:
            seen.add(key)
            result.append((d1, d2))
    return result

fg_unique = unique_pairs(farkli_gun_ciftler)
fs_unique = unique_pairs(farkli_slot_ciftler)

# Farkli gun ihlallerini bul (ayni gun)
fg_ihlal = []
for d1, d2 in fg_unique:
    s1 = schedule_map.get(d1)
    s2 = schedule_map.get(d2)
    if s1 and s2 and s1[0] == s2[0]:
        n1 = d1.replace('İ','I').replace('Ş','S').replace('Ğ','G').replace('Ü','U').replace('Ö','O').replace('Ç','C')
        n2 = d2.replace('İ','I').replace('Ş','S').replace('Ğ','G').replace('Ü','U').replace('Ö','O').replace('Ç','C')
        if not n1.startswith(EXCLUDED_PREFIXES) and not n2.startswith(EXCLUDED_PREFIXES):
            key = tuple(sorted([d1, d2]))
            ortak = collision_dict.get(key, 0)
            fg_ihlal.append((d1, d2, s1, s2, ortak))

# Farkli slot ihlallerini bul (ayni gun + ayni slot)
fs_ihlal = []
for d1, d2 in fs_unique:
    s1 = schedule_map.get(d1)
    s2 = schedule_map.get(d2)
    if s1 and s2 and s1[0] == s2[0] and s1[1] == s2[1]:
        n1 = d1.replace('İ','I').replace('Ş','S').replace('Ğ','G').replace('Ü','U').replace('Ö','O').replace('Ç','C')
        n2 = d2.replace('İ','I').replace('Ş','S').replace('Ğ','G').replace('Ü','U').replace('Ö','O').replace('Ç','C')
        if not n1.startswith(EXCLUDED_PREFIXES) and not n2.startswith(EXCLUDED_PREFIXES):
            key = tuple(sorted([d1, d2]))
            ortak = collision_dict.get(key, 0)
            fs_ihlal.append((d1, d2, s1, s2, ortak))

print("=" * 100)
print("DEGISTIRILMESI GEREKEN DERSLER - SON DURUM ANALIZI")
print("=" * 100)

print("\n" + "=" * 100)
print("A. FARKLI GUN KISITI IHLALLERI (Ayni gun, farkli slot -> Biri baska GUNE tasinsin)")
print("=" * 100)
print(f"{'#':<3} {'Ders 1':<12} {'Slot 1':<8} {'Ders 2':<12} {'Slot 2':<8} {'Ortak Ogr.':<10} {'Neden'}")
print("-" * 100)

for i, (d1, d2, s1, s2, ortak) in enumerate(fg_ihlal, 1):
    neden = "Kisit: Farkli gun olmali, ayni gun kalmis"
    print(f"{i:<3} {d1:<12} G{s1[0]}/S{s1[1]:<4} {d2:<12} G{s2[0]}/S{s2[1]:<4} {ortak:<10} {neden}")

print("\n" + "=" * 100)
print("B. FARKLI SLOT KISITI IHLALLERI (Ayni gun + Ayni slot -> Biri baska SLOTa/GUNE tasinsin)")
print("=" * 100)
print(f"{'#':<3} {'Ders 1':<12} {'Slot 1':<8} {'Ders 2':<12} {'Slot 2':<8} {'Ortak Ogr.':<10} {'Neden'}")
print("-" * 100)

for i, (d1, d2, s1, s2, ortak) in enumerate(fs_ihlal, 1):
    if d1 == d2:
        neden = "Kisit dosyasinda tekrar/hata (ayni ders kendisiyle)"
    else:
        neden = "Kisit: Farkli slot olmali, ayni gun+aynI slotta kalmis"
    print(f"{i:<3} {d1:<12} G{s1[0]}/S{s1[1]:<4} {d2:<12} G{s2[0]}/S{s2[1]:<4} {ortak:<10} {neden}")

# Her dersten kac kez gectigini say
problem_count = defaultdict(int)
for d1, d2, s1, s2, _ in fg_ihlal:
    if d1 != d2:
        problem_count[d1] += 1
        problem_count[d2] += 1
for d1, d2, s1, s2, _ in fs_ihlal:
    if d1 != d2:
        problem_count[d1] += 1
        problem_count[d2] += 1

print("\n" + "=" * 100)
print("C. EN COK PROBLEM CIKARAN DERSLER (Toplam ihlal sayisi)")
print("=" * 100)
for ders, count in sorted(problem_count.items(), key=lambda x: -x[1]):
    slot = schedule_map.get(ders)
    slot_str = f"G{slot[0]}/S{slot[1]}" if slot else "?"
    print(f"  {ders:<12} ({slot_str}) -> {count} kez ihlalde geciyor")

print("\n" + "=" * 100)
print("OZET")
print("=" * 100)
print(f"Farkli gun ihlali (degistirilmesi gereken cift) : {len(fg_ihlal)}")
print(f"Farkli slot ihlali (degistirilmesi gereken cift): {len(fs_ihlal)}")
print(f"Toplam problem cifti                            : {len(fg_ihlal) + len(fs_ihlal)}")
print(f"Etkilenen tekil ders sayisi                     : {len(problem_count)}")
print("\nNOT: Bu derslerin hepsi OGRENCI CAKISMASI acisindan sorunlu DEGIL.")
print("     Ancak kisit dosyasindaki kurallara uymuyorlar.")
print("     Degisiklik yapilirken OGRENCI CAKISMASI da goz onunde bulundurulmali.")
