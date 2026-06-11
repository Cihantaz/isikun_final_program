import pandas as pd
from datetime import datetime, time

# Load new program
df = pd.read_excel('c:/Users/cihan.tazeoz/Desktop/26 BAHAR FİNAL/UYGULAMA DOSYALARI/2026 Bahar Final Programı _03062026.xlsx')
collisions = pd.read_csv('c:/Users/cihan.tazeoz/Desktop/26 BAHAR FİNAL/UYGULAMA DOSYALARI/final_exam_collisions_14_05_2026_10_48_47.csv')

# Clean course codes
df['Ders Kodu'] = df['Ders Kodu'].astype(str).str.strip().str.upper()

# Date to Gun mapping
tarih_gun = {
    datetime(2026, 6, 8): 1,
    datetime(2026, 6, 9): 2,
    datetime(2026, 6, 10): 3,
    datetime(2026, 6, 11): 4,
    datetime(2026, 6, 12): 5,
    datetime(2026, 6, 15): 6,
    datetime(2026, 6, 16): 7,
    datetime(2026, 6, 17): 8,
    datetime(2026, 6, 18): 9,
    datetime(2026, 6, 19): 10,
}

# Time to Slot mapping
def saat_to_slot(saat):
    if pd.isna(saat):
        return 0
    if isinstance(saat, time):
        h = saat.hour
    else:
        s = str(saat).strip()
        if s == '-':
            return 0
        h = int(s.split(':')[0])
    if h == 8:
        return 1
    elif h == 11:
        return 2
    elif h == 14:
        return 3
    return 0

# Build program dict
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

# INDE courses in program
inde_courses = [c for c in program_dict.keys() if c.startswith('INDE')]
print(f"Programdaki INDE ders sayisi: {len(inde_courses)}")
print()

print("=" * 90)
print("INDE DERSLERIN GUNCEL YERLERI")
print("=" * 90)
for c in sorted(inde_courses):
    g, s = program_dict[c]
    print(f"  {c:<15} -> G{g}S{s}")
print()

# Which INDE courses are in G7, G8, G9?
print("=" * 90)
print("G7/G8/G9'DAKI INDE DERSLERI")
print("=" * 90)
for c in sorted(inde_courses):
    g, s = program_dict[c]
    if g in [7, 8, 9]:
        print(f"  {c:<15} -> G{g}S{s}")
print()

# DiffDay violations in new program
print("=" * 90)
print("DIFFDAY KISIT KONTROLU (Guncel Program)")
print("=" * 90)

for i, grup in enumerate(gruplar, 1):
    print(f"\n--- Grup {i}: {', '.join(grup)} ---")
    gunler = {}
    yoklar = []
    for d in grup:
        d_clean = d.replace('.L', '').replace('-L', '').upper()
        # Try exact match first
        g = program_dict.get(d_clean)
        if g is None:
            yoklar.append(d)
        else:
            if g[0] not in gunler:
                gunler[g[0]] = []
            gunler[g[0]].append(d)
    
    if yoklar:
        print(f"  Programda bulunamayan: {', '.join(yoklar)}")
    
    ihlal = False
    for g, dersler in gunler.items():
        if len(dersler) > 1:
            print(f"  X IHLAL: Ayni gun (G{g}) -> {', '.join(dersler)}")
            ihlal = True
    
    if not ihlal and not yoklar:
        print(f"  OK - tum dersler farkli gunlerde")
    elif not ihlal and yoklar:
        print(f"  OK - mevcut dersler farkli gunlerde (bulunamayanlar haric)")

# Find empty slots in G5, G6, G7
print()
print("=" * 90)
print("G5/G6/G7'DEKI BOS SLOTLAR")
print("=" * 90)
for g in [5, 6, 7]:
    for s in [1, 2, 3]:
        occup = [c for c, (gg, ss) in program_dict.items() if gg == g and ss == s]
        print(f"  G{g}S{s}: {len(occup)} ders")
        if occup:
            print(f"    {', '.join(occup[:5])}{' ...' if len(occup) > 5 else ''}")

