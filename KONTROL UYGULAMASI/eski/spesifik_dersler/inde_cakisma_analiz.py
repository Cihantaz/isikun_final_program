import pandas as pd

program = pd.read_excel('c:/Users/cihan.tazeoz/Desktop/26 BAHAR FİNAL/UYGULAMA DOSYALARI/YENILENMIS_PROGRAM_2026_BAHAR_FINAL_TARIH.xlsx')
collisions = pd.read_csv('c:/Users/cihan.tazeoz/Desktop/26 BAHAR FİNAL/UYGULAMA DOSYALARI/final_exam_collisions_14_05_2026_10_48_47.csv')

program['Ders Kodu'] = program['Ders Kodu'].astype(str).str.strip().str.upper()
program['Gün'] = program['Gün'].fillna(0).astype(int)
program['Slot'] = program['Slot'].fillna(0).astype(int)

program_dict = {}
for _, row in program.iterrows():
    code = row['Ders Kodu']
    if code and str(code) != 'nan':
        program_dict[code] = (int(row['Gün']), int(row['Slot']), str(row.get('Sınav Tarihi','')), str(row.get('Sınav Başlangıç Saati','')))

gun_map = {1:"Pzt",2:"Sal",3:"Crs",4:"Prs",5:"Cum",6:"Pzt",7:"Sal",8:"Crs",9:"Prs",10:"Cum"}

# Kullanıcının verdiği gruplar
gruplar = [
    ['INDE2001', 'INDE2001.L', 'MATH2201', 'MATH2105'],
    ['INDE2001', 'INDE2001-L', 'INDE2156', 'INDE2211', 'INDE2452', 'MATH2107', 'MATH2103'],
    ['INDE2002', 'MATH2201', 'MATH2107', 'MATH2105'],
    ['INDE2001', 'INDE2001-L', 'MATH2201', 'MATH2107', 'MATH2105'],
    ['INDE3151', 'INDE33314', 'INDE3312', 'BUSI4573', 'INDE3145', 'INDE4003', 'INDE4902'],
    ['INDE3312', 'INDE3314', 'BUSI4573', 'INDE3145', 'INDE2002', 'INDE4902'],
    ['INDE4185', 'INDE4141', 'INDE4181', 'INDE4403', 'INDE4902'],
]

ozel_dersler = ['INDE2452', 'INDE4902']

# Collision lookup
collision_dict = {}
for _, row in collisions.iterrows():
    c1 = str(row['Course1']).strip().upper()
    c2 = str(row['Course2']).strip().upper()
    cnt = int(row['Common Student Count'])
    collision_dict[(c1, c2)] = cnt
    collision_dict[(c2, c1)] = cnt

def cakisma_bul(d1, d2):
    return collision_dict.get((d1.upper(), d2.upper()), 0)

def slot_yaz(code):
    if code.upper() not in program_dict:
        return "YOK"
    g, s, t, sa = program_dict[code.upper()]
    return f"G{g}S{s} ({gun_map.get(g,'?')} {sa})"

print("=" * 100)
print("GRUP ICINDEKI CAKISMALAR")
print("=" * 100)

for i, grup in enumerate(gruplar, 1):
    print(f"\n--- Grup {i} ---")
    # Mevcut slotlar
    for d in grup:
        print(f"  {d:<15} -> {slot_yaz(d)}")
    
    # Çakışmalar
    cakislar = []
    for j in range(len(grup)):
        for k in range(j+1, len(grup)):
            cnt = cakisma_bul(grup[j], grup[k])
            if cnt > 0:
                cakislar.append((grup[j], grup[k], cnt))
    
    if cakislar:
        print("  CAKISMALAR:")
        for d1, d2, cnt in cakislar:
            print(f"    {d1} - {d2}: {cnt} ogrenci")
    else:
        print("  CAKISMA YOK")
    
    # Aynı gün/slot kontrolü
    slotlar = {}
    for d in grup:
        if d.upper() in program_dict:
            g, s, _, _ = program_dict[d.upper()]
            slot_key = f"G{g}S{s}"
            if slot_key not in slotlar:
                slotlar[slot_key] = []
            slotlar[slot_key].append(d)
    
    for sk, dersler in slotlar.items():
        if len(dersler) > 1:
            print(f"  AYNI SLOT ({sk}): {', '.join(dersler)}")

print()
print("=" * 100)
print(f"OZEL DERS ANALIZI: {', '.join(ozel_dersler)}")
print("=" * 100)

for d in ozel_dersler:
    print(f"\n{d} -> {slot_yaz(d)}")
    # Bu dersle çakışan tüm INDE dersleri
    cakislar = []
    for (c1, c2), cnt in collision_dict.items():
        if c1 == d.upper() and c2.startswith('INDE'):
            cakislar.append((c2, cnt, slot_yaz(c2)))
        elif c2 == d.upper() and c1.startswith('INDE'):
            cakislar.append((c1, cnt, slot_yaz(c1)))
    
    cakislar.sort(key=lambda x: -x[1])
    if cakislar:
        print("  INDE ile cakisanlar:")
        for dd, cnt, sl in cakislar:
            print(f"    {dd:<15} {cnt:>3} ogrenci  -> {sl}")
    else:
        print("  INDE ile cakisma yok")
    
    # Toplam çakışanlar
    tum = []
    for (c1, c2), cnt in collision_dict.items():
        if c1 == d.upper():
            tum.append((c2, cnt, slot_yaz(c2)))
        elif c2 == d.upper():
            tum.append((c1, cnt, slot_yaz(c1)))
    tum.sort(key=lambda x: -x[1])
    print(f"  Toplam {len(tum)} farkli dersle cakisiyor, toplam {sum(c[1] for c in tum)} ogrenci")
    if tum:
        print("  Tum cakisanlar:")
        for dd, cnt, sl in tum[:10]:
            print(f"    {dd:<15} {cnt:>3} ogrenci  -> {sl}")
        if len(tum) > 10:
            print(f"    ... ve {len(tum)-10} ders daha")

