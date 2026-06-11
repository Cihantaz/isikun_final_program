import pandas as pd
from datetime import datetime, time

# Load new program
df = pd.read_excel('c:/Users/cihan.tazeoz/Desktop/26 BAHAR FİNAL/UYGULAMA DOSYALARI/2026 Bahar Final Programı _03062026.xlsx')
collisions = pd.read_csv('c:/Users/cihan.tazeoz/Desktop/26 BAHAR FİNAL/UYGULAMA DOSYALARI/final_exam_collisions_14_05_2026_10_48_47.csv')

# Clean
df['Ders Kodu'] = df['Ders Kodu'].astype(str).str.strip().str.upper()

tarih_gun = {
    datetime(2026, 6, 8): 1, datetime(2026, 6, 9): 2, datetime(2026, 6, 10): 3,
    datetime(2026, 6, 11): 4, datetime(2026, 6, 12): 5, datetime(2026, 6, 15): 6,
    datetime(2026, 6, 16): 7, datetime(2026, 6, 17): 8, datetime(2026, 6, 18): 9,
    datetime(2026, 6, 19): 10,
}

def saat_to_slot(saat):
    if pd.isna(saat): return 0
    if isinstance(saat, time):
        h = saat.hour
    else:
        s = str(saat).strip()
        if s == '-': return 0
        h = int(s.split(':')[0])
    if h == 8: return 1
    elif h == 11: return 2
    elif h == 14: return 3
    return 0

df['Gun'] = df['Sınav Tarihi'].apply(lambda x: tarih_gun.get(x, 0) if pd.notna(x) else 0)
df['Slot'] = df['Sınav Başlangıç Saati'].apply(saat_to_slot)

program_dict = {}
for _, row in df.iterrows():
    code = row['Ders Kodu']
    if code and str(code) != 'NAN':
        program_dict[code] = (int(row['Gun']), int(row['Slot']))

# Collision dict
collision_dict = {}
for _, row in collisions.iterrows():
    c1 = str(row['Course1']).strip().upper()
    c2 = str(row['Course2']).strip().upper()
    cnt = int(row['Common Student Count'])
    collision_dict[(c1, c2)] = cnt
    collision_dict[(c2, c1)] = cnt

# INDE courses to move (G7/G8/G9)
hedef_dersler = ['INDE2001', 'INDE2002', 'INDE2211', 'INDE3151', 'INDE4003', 'INDE4141', 'INDE4181', 'INDE4185']
# Also include G10 ones that need to be moved for DiffDay
hedef_dersler += ['INDE2452', 'INDE4902', 'INDE4912']

# Also G1 duplicates
hedef_dersler += ['INDE3312', 'INDE3314']

def cakisma_bul(d1, gun, slot):
    """Check collisions if d1 moves to (gun, slot)"""
    total = 0
    details = []
    for d2, (g2, s2) in program_dict.items():
        if d2 == d1:
            continue
        if g2 == gun and s2 == slot:
            cnt = collision_dict.get((d1, d2), 0)
            if cnt > 0:
                total += cnt
                details.append(f"{d2}({cnt})")
    return total, details

print("=" * 100)
print("G5/G6/G7 ALTERNATIFLERI (G7/G8/G9'daki INDE dersleri)")
print("=" * 100)

for d in hedef_dersler:
    if d not in program_dict:
        continue
    g_mevcut, s_mevcut = program_dict[d]
    print(f"\n{d} -> Mevcut: G{g_mevcut}S{s_mevcut}")
    
    for g in [5, 6, 7]:
        for s in [1, 2, 3]:
            total, details = cakisma_bul(d, g, s)
            durum = "UYGUN" if total == 0 else f"OLUMSUZ ({total} ogrenci)"
            detay = ", ".join(details[:3]) if details else "-"
            if len(details) > 3:
                detay += f" (+{len(details)-3})"
            print(f"  G{g}S{s}: {durum:<20} {detay}")

# DiffDay constraints between INDE courses
print()
print("=" * 100)
print("INDE DERSLERI ARASI DIFFDAY KONTROLU (G5/G6/G7 hedef)")
print("=" * 100)

# Groups with only INDE courses
inde_only_groups = [
    ['INDE3312', 'INDE3314'],  # G1 duplicate
    ['INDE4185', 'INDE4141', 'INDE4181', 'INDE4902'],  # Grup 7 subset
    ['INDE2452', 'INDE4902'],  # Grup 8
]

for grup in inde_only_groups:
    print(f"\nGrup: {', '.join(grup)}")
    mevcut = {d: program_dict.get(d, (None, None)) for d in grup}
    for d, (g, s) in mevcut.items():
        if g:
            print(f"  {d}: G{g}S{s}")
    
    # Check if any two are same day
    gunler = {}
    for d, (g, s) in mevcut.items():
        if g:
            if g not in gunler:
                gunler[g] = []
            gunler[g].append(d)
    for g, dersler in gunler.items():
        if len(dersler) > 1:
            print(f"  X IHLAL: G{g} -> {', '.join(dersler)}")

