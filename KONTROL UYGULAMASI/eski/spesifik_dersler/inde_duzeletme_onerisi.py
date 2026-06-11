import pandas as pd
from datetime import datetime, time as dt_time

# Load collision data
collisions = pd.read_csv('c:/Users/cihan.tazeoz/Desktop/26 BAHAR FİNAL/UYGULAMA DOSYALARI/final_exam_collisions_14_05_2026_10_48_47.csv')

# Collision dict
collision_dict = {}
for _, row in collisions.iterrows():
    c1 = str(row['Course1']).strip().upper()
    c2 = str(row['Course2']).strip().upper()
    cnt = int(row['Common Student Count'])
    collision_dict[(c1, c2)] = cnt
    collision_dict[(c2, c1)] = cnt

# Kullanıcının yeni programı
yeni_yerler = {
    'INDE3312': (1, 2), 'INDE3314': (1, 1), 'INDE4313': (1, 3),
    'INDE3145': (2, 1), 'INDE2156': (4, 1), 'INDE2002': (5, 3),
    'INDE3151': (8, 1), 'INDE4141': (8, 3), 'INDE2001': (5, 3),
    'INDE2211': (6, 3), 'INDE4003': (9, 3), 'INDE4181': (5, 1),
    'INDE4185': (9, 3), 'INDE2452': (6, 3), 'INDE4902': (10, 1),
    'INDE4912': (10, 1),
}

# DiffDay groups
gruplar = [
    ['INDE2001', 'INDE2001.L', 'MATH2201', 'MATH2105'],
    ['INDE2001', 'INDE2001-L', 'INDE2156', 'INDE2211', 'INDE2452', 'MATH2107', 'MATH2103'],
    ['INDE2002', 'MATH2201', 'MATH2107', 'MATH2105'],
    ['INDE2001', 'INDE2001-L', 'MATH2201', 'MATH2107', 'MATH2105'],
    ['INDE3151', 'INDE33314', 'INDE3312', 'BUSI4573', 'INDE3145', 'INDE4003', 'INDE4902'],
    ['INDE3312', 'INDE3314', 'BUSI4573', 'INDE3145', 'INDE2002', 'INDE4902'],
    ['INDE4185', 'INDE4141', 'INDE4181', 'INDE4403', 'INDE4902'],
    ['INDE2452', 'INDE4902'],
]

print("=" * 100)
print("DUZELTME ONERISI")
print("=" * 100)

oneri = {
    'INDE2452': (7, 3),    # G6S3 -> G7S3 (INDE2211 ile 18 ogrenci cakismasini coz)
    'INDE3314': (6, 3),    # G1S1 -> G6S3 (INDE3312 ile DiffDay ihlalini coz)
    'INDE2002': (5, 2),    # G5S3 -> G5S2 (INDE2001'den ayir)
}

print("\nONERILEN DEGISIKLIKLER:")
print("-" * 100)
for d, (g_yeni, s_yeni) in oneri.items():
    g_eski, s_eski = yeni_yerler[d]
    print(f"  {d}: G{g_eski}S{s_eski} -> G{g_yeni}S{s_yeni}")

# Apply changes to a copy
yeni_yerler_fix = dict(yeni_yerler)
for d, (g, s) in oneri.items():
    yeni_yerler_fix[d] = (g, s)

# Check same slot INDE conflicts
print("\n" + "=" * 100)
print("DUZELTMEDEN SONRA AYNI SLOT INDE DERSLERI")
print("=" * 100)

slot_dict = {}
for d, (g, s) in yeni_yerler_fix.items():
    key = f"G{g}S{s}"
    if key not in slot_dict:
        slot_dict[key] = []
    slot_dict[key].append(d)

for slot, dersler in sorted(slot_dict.items()):
    if len(dersler) > 1:
        print(f"\n  {slot}: {', '.join(dersler)}")
        for i in range(len(dersler)):
            for j in range(i+1, len(dersler)):
                cnt = collision_dict.get((dersler[i], dersler[j]), 0)
                print(f"    {dersler[i]} - {dersler[j]}: {cnt} ogrenci")

# DiffDay check after fix
print("\n" + "=" * 100)
print("DIFFDAY KONTROLU (DUZELTMEDEN SONRA)")
print("=" * 100)

# Load full program for MATH courses
df_full = pd.read_excel('c:/Users/cihan.tazeoz/Desktop/26 BAHAR FİNAL/UYGULAMA DOSYALARI/2026 Bahar Final Programı _03062026.xlsx')
df_full['Ders Kodu'] = df_full['Ders Kodu'].astype(str).str.strip().str.upper()

tarih_gun = {
    datetime(2026, 6, 8): 1, datetime(2026, 6, 9): 2, datetime(2026, 6, 10): 3,
    datetime(2026, 6, 11): 4, datetime(2026, 6, 12): 5, datetime(2026, 6, 15): 6,
    datetime(2026, 6, 16): 7, datetime(2026, 6, 17): 8, datetime(2026, 6, 18): 9,
    datetime(2026, 6, 19): 10,
}

def saat_to_slot(saat):
    if pd.isna(saat): return 0
    if isinstance(saat, dt_time):
        h = saat.hour
    else:
        s = str(saat).strip()
        if s == '-': return 0
        h = int(s.split(':')[0])
    if h == 8: return 1
    elif h == 11: return 2
    elif h == 14: return 3
    return 0

df_full['Gun'] = df_full['Sınav Tarihi'].apply(lambda x: tarih_gun.get(x, 0) if pd.notna(x) else 0)
df_full['Slot'] = df_full['Sınav Başlangıç Saati'].apply(saat_to_slot)

program_dict = {}
for _, row in df_full.iterrows():
    code = row['Ders Kodu']
    if code and str(code) != 'NAN':
        program_dict[code] = (int(row['Gun']), int(row['Slot']))

ihlal_sayisi = 0
for i, grup in enumerate(gruplar, 1):
    print(f"\n--- Grup {i}: {', '.join(grup)} ---")
    gunler = {}
    yoklar = []
    for d in grup:
        d_clean = d.replace('.L', '').replace('-L', '').upper()
        g = yeni_yerler_fix.get(d_clean)
        if g is None:
            g = program_dict.get(d_clean)
            if g is None:
                yoklar.append(d)
                continue
            g = g[0]
        else:
            g = g[0]
        
        if g not in gunler:
            gunler[g] = []
        gunler[g].append(d)
    
    if yoklar:
        print(f"  Programda bulunamayan: {', '.join(yoklar)}")
    
    ihlal = False
    for g, dersler in gunler.items():
        if len(dersler) > 1:
            print(f"  X IHLAL: Ayni gun (G{g}) -> {', '.join(dersler)}")
            ihlal = True
            ihlal_sayisi += 1
    
    if not ihlal and not yoklar:
        print(f"  OK - tum dersler farkli gunlerde")
    elif not ihlal and yoklar:
        print(f"  OK - mevcut dersler farkli gunlerde (bulunamayanlar haric)")

print(f"\nTOPLAM IHLAL: {ihlal_sayisi}")

# Load full program to check collisions after fix
print("\n" + "=" * 100)
print("OGRENCI CAKISMALARI (DUZELTMEDEN SONRA)")
print("=" * 100)

toplam_cakisma = 0
for d, (g, s) in sorted(yeni_yerler_fix.items()):
    cakisan = []
    for d2, (g2, s2) in program_dict.items():
        if d2 == d:
            continue
        if g2 == g and s2 == s:
            cnt = collision_dict.get((d, d2), 0)
            if cnt > 0:
                cakisan.append((d2, cnt))
    
    if cakisan:
        cakisan.sort(key=lambda x: -x[1])
        toplam = sum(c[1] for c in cakisan)
        toplam_cakisma += toplam
        print(f"\n{d} -> G{g}S{s} | {len(cakisan)} ders, {toplam} ogrenci")
        for d2, cnt in cakisan[:5]:
            print(f"    {d2:<15} {cnt:>3} ogrenci")
    else:
        print(f"{d} -> G{g}S{s} | CAKISMA YOK")

print(f"\nTOPLAM CAKISAN OGRENCI: {toplam_cakisma}")

