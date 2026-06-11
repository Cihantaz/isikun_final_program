import pandas as pd

program = pd.read_excel('c:/Users/cihan.tazeoz/Desktop/26 BAHAR FİNAL/UYGULAMA DOSYALARI/YENILENMIS_PROGRAM_2026_BAHAR_FINAL_TARIH.xlsx')
program['Ders Kodu'] = program['Ders Kodu'].astype(str).str.strip().str.upper()
program['Gün'] = program['Gün'].fillna(0).astype(int)

program_dict = {}
for _, row in program.iterrows():
    code = row['Ders Kodu']
    if code and str(code) != 'nan':
        program_dict[code] = int(row['Gün'])

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

print("=" * 90)
print("DIFFDAY (FARKLI GUN) KISIT KONTROLU")
print("=" * 90)

for i, grup in enumerate(gruplar, 1):
    print(f"\n--- Grup {i}: {', '.join(grup)} ---")
    gunler = {}
    yoklar = []
    for d in grup:
        d_clean = d.replace('.L', '').replace('-L', '').upper()
        g = program_dict.get(d_clean)
        if g is None:
            yoklar.append(d)
        else:
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
    
    if not ihlal and not yoklar:
        print(f"  OK - tum dersler farkli gunlerde")
    elif not ihlal and yoklar:
        print(f"  OK - mevcut dersler farkli gunlerde (bulunamayanlar haric)")

