import pandas as pd
from datetime import datetime, time as dt_time

# Load program
df = pd.read_excel('c:/Users/cihan.tazeoz/Desktop/26 BAHAR FİNAL/UYGULAMA DOSYALARI/2026 Bahar Final Programı _03062026.xlsx')
df['Ders Kodu'] = df['Ders Kodu'].astype(str).str.strip().str.upper()

collisions = pd.read_csv('c:/Users/cihan.tazeoz/Desktop/26 BAHAR FİNAL/UYGULAMA DOSYALARI/final_exam_collisions_14_05_2026_10_48_47.csv')
kisitlar = pd.read_excel('c:/Users/cihan.tazeoz/Desktop/26 BAHAR FİNAL/UYGULAMA DOSYALARI/ders_kisitleri.xlsx')

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

# Constraint types
header = list(kisitlar.columns)
use_cols = header[:4]
first_col = str(header[0]).strip().lower()
if len(header) >= 5 and (first_col.startswith('unnamed') or str(header[0]).strip().isdigit()):
    use_cols = header[1:5]

kisit_groups = []
for idx, row in kisitlar.iterrows():
    dersler = []
    for col in kisitlar.columns:
        if pd.notna(row[col]):
            val = str(row[col]).strip().upper()
            if val and val not in ['NAN']:
                dersler.append(val)
    tip = None
    for col in use_cols:
        if pd.notna(row[col]):
            val = str(row[col]).strip().upper()
            if val in ['SAMEDAY', 'SAMESLOT', 'DIFFDAY', 'DIFFSLOT']:
                tip = val
                break
    if tip and dersler:
        kisit_groups.append((tip, dersler))

# Changes to check
changes = {
    'TRAD2504': (1, 3),   # 08.06 14:30-17:30
    'BUSI3632': (4, 3),   # 11.06 14:30-17:30
    'BUSI4572': (4, 1),   # 11.06 08:30-11:30
    'BUSI3222': (2, 1),   # 09.06 08:30-11:30
}

# Apply changes to a copy
test_dict = dict(program_dict)
for d, (g, s) in changes.items():
    test_dict[d] = (g, s)

print("=" * 100)
print("DEGISIKLIK KONTROLU")
print("=" * 100)

for d, (g_yeni, s_yeni) in changes.items():
    g_eski, s_eski = program_dict.get(d, (None, None))
    print(f"\n{d}: G{g_eski}S{s_eski} -> G{g_yeni}S{s_yeni}")
    
    # Check collisions with all courses at new slot
    cakisan = []
    for d2, (g2, s2) in test_dict.items():
        if d2 == d:
            continue
        if g2 == g_yeni and s2 == s_yeni:
            cnt = collision_dict.get((d, d2), 0)
            if cnt > 0:
                cakisan.append((d2, cnt))
    
    if cakisan:
        cakisan.sort(key=lambda x: -x[1])
        toplam = sum(c[1] for c in cakisan)
        print(f"  CAKISMA: {len(cakisan)} ders, {toplam} ogrenci")
        for d2, cnt in cakisan:
            print(f"    {d2:<15} {cnt:>3} ogrenci")
    else:
        print(f"  CAKISMA YOK")

# Check constraints
print("\n" + "=" * 100)
print("KISIT KONTROLU")
print("=" * 100)

for tip, dersler in kisit_groups:
    ilgili = [d for d in dersler if d in changes]
    if not ilgili:
        continue
    
    print(f"\n{tip}: {', '.join(dersler)}")
    gunler = {}
    for d in dersler:
        g = test_dict.get(d, (None, None))
        if g:
            if g[0] not in gunler:
                gunler[g[0]] = []
            gunler[g[0]].append((d, g[1]))
    
    sorun = False
    if tip == 'DIFFDAY':
        for g, liste in gunler.items():
            if len(liste) > 1:
                print(f"  X IHLAL: Ayni gun G{g} -> {', '.join(d + 'S' + str(s) for d, s in liste)}")
                sorun = True
    elif tip == 'DIFFSLOT':
        for g, liste in gunler.items():
            if len(liste) > 1:
                print(f"  X IHLAL: Ayni gun/slot G{g} -> {', '.join(d + 'S' + str(s) for d, s in liste)}")
                sorun = True
    elif tip == 'SAMEDAY':
        gun_list = [g for g, _ in gunler.values()]
        if len(gun_list) != len(set(gun_list)):
            print(f"  X IHLAL: Farkli gunlerde olmali")
            sorun = True
    elif tip == 'SAMESLOT':
        slot_list = [s for _, s in gunler.values()]
        if len(slot_list) != len(set(slot_list)):
            print(f"  X IHLAL: Farkli slotlarda olmali")
            sorun = True
    
    if not sorun:
        print(f"  OK")

# Mutual check among the 4 courses
print("\n" + "=" * 100)
print("4 DERSIN BIRBIRIYLE KONTROLU")
print("=" * 100)

ders_list = list(changes.keys())
for i in range(len(ders_list)):
    for j in range(i+1, len(ders_list)):
        d1, d2 = ders_list[i], ders_list[j]
        g1, s1 = changes[d1]
        g2, s2 = changes[d2]
        cnt = collision_dict.get((d1, d2), 0)
        same_day = (g1 == g2)
        same_slot = (g1 == g2 and s1 == s2)
        durum = "AYNI GUN!" if same_day else "Farkli gun"
        if same_slot:
            durum = "AYNI SLOT!"
        print(f"  {d1} - {d2}: {cnt} ogrenci | {durum}")

