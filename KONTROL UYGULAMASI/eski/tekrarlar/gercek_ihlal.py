import pandas as pd

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

# Farkli gun ihlali (ayni gun)
farkli_gun_ihlal = []
for d1, d2 in farkli_gun_kisit:
    s1 = schedule_map.get(d1)
    s2 = schedule_map.get(d2)
    if s1 and s2 and s1[0] == s2[0]:
        farkli_gun_ihlal.append((d1, d2, s1, s2))

# Farkli slot ihlali ama SADECE AYNI GUN olanlar (farkli gun sorun degil)
farkli_slot_ihlal_ayni_gun = []
for d1, d2 in farkli_slot_kisit:
    s1 = schedule_map.get(d1)
    s2 = schedule_map.get(d2)
    if s1 and s2 and s1[1] == s2[1] and s1[0] == s2[0]:
        farkli_slot_ihlal_ayni_gun.append((d1, d2, s1, s2))

# Tekrar kaldir
def unique_list(items):
    seen = set()
    result = []
    for d1, d2, s1, s2 in items:
        key = tuple(sorted([d1, d2]))
        if key not in seen:
            seen.add(key)
            result.append((d1, d2, s1, s2))
    return result

fg_unique = unique_list(farkli_gun_ihlal)
fs_ayni_gun_unique = unique_list(farkli_slot_ihlal_ayni_gun)

# Mimarlik haric
fg_haric = []
for d1, d2, s1, s2 in fg_unique:
    n1 = d1.replace('İ','I').replace('Ş','S').replace('Ğ','G').replace('Ü','U').replace('Ö','O').replace('Ç','C')
    n2 = d2.replace('İ','I').replace('Ş','S').replace('Ğ','G').replace('Ü','U').replace('Ö','O').replace('Ç','C')
    if not n1.startswith(EXCLUDED_PREFIXES) and not n2.startswith(EXCLUDED_PREFIXES):
        fg_haric.append((d1, d2, s1, s2))

fs_haric = []
for d1, d2, s1, s2 in fs_ayni_gun_unique:
    n1 = d1.replace('İ','I').replace('Ş','S').replace('Ğ','G').replace('Ü','U').replace('Ö','O').replace('Ç','C')
    n2 = d2.replace('İ','I').replace('Ş','S').replace('Ğ','G').replace('Ü','U').replace('Ö','O').replace('Ç','C')
    if not n1.startswith(EXCLUDED_PREFIXES) and not n2.startswith(EXCLUDED_PREFIXES):
        fs_haric.append((d1, d2, s1, s2))

# Cakisma dosyasinda bu ciftler arasinda ortak ogrenci var mi kontrol et
collision_pairs = set()
for _, row in collisions.iterrows():
    c1 = str(row['Course1']).strip().upper()
    c2 = str(row['Course2']).strip().upper()
    collision_pairs.add(tuple(sorted([c1, c2])))

print("=" * 80)
print("GERCEK IHLALLER (Sadece Ayni Gun)")
print("=" * 80)

print("\n[A] FARKLI GUN KISITI IHLALLERI (Ayni Gun = SORUN)")
print("-" * 80)
print(f"Toplam: {len(fg_haric)}\n")
for i, (d1, d2, s1, s2) in enumerate(fg_haric, 1):
    pair = tuple(sorted([d1, d2]))
    has_collision = "EVET" if pair in collision_pairs else "HAYIR"
    print(f"{i:2d}. {d1:12s} ({s1[0]}/S{s1[1]}) - {d2:12s} ({s2[0]}/S{s2[1]})  |  Ortak ogrenci: {has_collision}")

print("\n[B] FARKLI SLOT KISITI IHLALLERI (Sadece Ayni Gun + Ayni Slot)")
print("-" * 80)
print(f"Toplam: {len(fs_haric)}\n")
for i, (d1, d2, s1, s2) in enumerate(fs_haric, 1):
    pair = tuple(sorted([d1, d2]))
    has_collision = "EVET" if pair in collision_pairs else "HAYIR"
    print(f"{i:2d}. {d1:12s} (G{s1[0]}/S{s1[1]}) - {d2:12s} (G{s2[0]}/S{s2[1]})  |  Ortak ogrenci: {has_collision}")

print("\n" + "=" * 80)
print("OZET")
print("=" * 80)
print(f"Farkli gun ihlali (ayni gun)      : {len(fg_haric)}")
print(f"Farkli slot ihlali (ayni gun+slot): {len(fs_haric)}")
print(f"Toplam gercek ihlal               : {len(fg_haric) + len(fs_haric)}")
