import pandas as pd

# Dosya yollari
collision_file = r"C:\Users\cihan.tazeoz\Desktop\26 BAHAR FİNAL\UYGULAMA DOSYALARI\final_exam_collisions_14_05_2026_10_48_47.csv"
schedule_file = r"C:\Users\cihan.tazeoz\Desktop\26 BAHAR FİNAL\UYGULAMA DOSYALARI\schedule (5).xlsx"
kisit_file = r"C:\Users\cihan.tazeoz\Desktop\26 BAHAR FİNAL\UYGULAMA DOSYALARI\ders_kisitleri.xlsx"

# Okuma
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

# Mimarlik fakultesi haric tut
fg_ihlal_haric = []
for d1, d2, s1, s2 in farkli_gun_ihlal:
    n1 = d1.replace('İ','I').replace('Ş','S').replace('Ğ','G').replace('Ü','U').replace('Ö','O').replace('Ç','C')
    n2 = d2.replace('İ','I').replace('Ş','S').replace('Ğ','G').replace('Ü','U').replace('Ö','O').replace('Ç','C')
    if not n1.startswith(EXCLUDED_PREFIXES) and not n2.startswith(EXCLUDED_PREFIXES):
        fg_ihlal_haric.append((d1, d2, s1, s2))

fs_ihlal_haric = []
for d1, d2, s1, s2 in farkli_slot_ihlal:
    n1 = d1.replace('İ','I').replace('Ş','S').replace('Ğ','G').replace('Ü','U').replace('Ö','O').replace('Ç','C')
    n2 = d2.replace('İ','I').replace('Ş','S').replace('Ğ','G').replace('Ü','U').replace('Ö','O').replace('Ç','C')
    if not n1.startswith(EXCLUDED_PREFIXES) and not n2.startswith(EXCLUDED_PREFIXES):
        fs_ihlal_haric.append((d1, d2, s1, s2))

print("=" * 90)
print("FARKLI GUN KISITI IHLALLERI (SANAT TASARIM MIMARLIK FAKULTESI DISINDA)")
print("=" * 90)
print(f"Toplam ihlal: {len(fg_ihlal_haric)}\n")
for i, (d1, d2, s1, s2) in enumerate(fg_ihlal_haric, 1):
    print(f"{i:2d}. {d1:12s} (G{s1[0]:2d}/S{s1[1]})  -  {d2:12s} (G{s2[0]:2d}/S{s2[1]})  |  AYNI GUN: G{s1[0]}")

print("\n" + "=" * 90)
print("FARKLI SLOT KISITI IHLALLERI (SANAT TASARIM MIMARLIK FAKULTESI DISINDA)")
print("=" * 90)
print(f"Toplam ihlal: {len(fs_ihlal_haric)}\n")
for i, (d1, d2, s1, s2) in enumerate(fs_ihlal_haric, 1):
    gun_durum = "AYNI GUN" if s1[0] == s2[0] else "FARKLI GUN"
    print(f"{i:2d}. {d1:12s} (G{s1[0]:2d}/S{s1[1]})  -  {d2:12s} (G{s2[0]:2d}/S{s2[1]})  |  AYNI SLOT: S{s1[1]} ({gun_durum})")

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

print("\n" + "=" * 90)
print("SCHEDULE'DA OLMAYAN DERSLER (SANAT TASARIM MIMARLIK FAKULTESI DISINDA)")
print("=" * 90)
print(f"Toplam eksik ders: {len(missing_courses)}\n")

core = sorted([c for c in missing_courses if c.startswith('CORE')])
proje = sorted([c for c in missing_courses if any(s in c for s in ['3910','3920','4901','4902','4910','4912','4920'])])
diger = sorted([c for c in missing_courses if c not in core and c not in proje])

print(f"[A] CORE DERSLERI ({len(core)} adet):")
for i in range(0, len(core), 6):
    print("    " + ", ".join(core[i:i+6]))

print(f"\n[B] PROJE/SEMINER DERSLERI ({len(proje)} adet):")
for i in range(0, len(proje), 6):
    print("    " + ", ".join(proje[i:i+6]))

print(f"\n[C] DIGER DERSLER ({len(diger)} adet):")
for i in range(0, len(diger), 6):
    print("    " + ", ".join(diger[i:i+6]))

print("\n" + "=" * 90)
print("OZET")
print("=" * 90)
print(f"Farkli gun ihlali (mimarlik haric) : {len(fg_ihlal_haric)}")
print(f"Farkli slot ihlali (mimarlik haric): {len(fs_ihlal_haric)}")
print(f"Eksik ders (mimarlik haric)        : {len(missing_courses)}")
