import pandas as pd

# --- Load data ---
program = pd.read_excel('c:/Users/cihan.tazeoz/Desktop/26 BAHAR FİNAL/UYGULAMA DOSYALARI/YENILENMIS_PROGRAM_2026_BAHAR_FINAL_TARIH.xlsx')
collisions = pd.read_csv('c:/Users/cihan.tazeoz/Desktop/26 BAHAR FİNAL/UYGULAMA DOSYALARI/final_exam_collisions_14_05_2026_10_48_47.csv')
kisitlar = pd.read_excel('c:/Users/cihan.tazeoz/Desktop/26 BAHAR FİNAL/UYGULAMA DOSYALARI/ders_kisitleri.xlsx')

# Clean
program['Ders Kodu'] = program['Ders Kodu'].astype(str).str.strip().str.upper()
program['Gün'] = program['Gün'].fillna(0).astype(int)
program['Slot'] = program['Slot'].fillna(0).astype(int)

# BMED4112 collisions
c1 = collisions[collisions['Course1'].str.strip().str.upper() == 'BMED4112'][['Course2','Common Student Count']].rename(columns={'Course2':'Ders'})
c2 = collisions[collisions['Course2'].str.strip().str.upper() == 'BMED4112'][['Course1','Common Student Count']].rename(columns={'Course1':'Ders'})
cakisan = pd.concat([c1, c2])
cakisan['Ders'] = cakisan['Ders'].str.strip().str.upper()
cakisan_dict = dict(zip(cakisan['Ders'], cakisan['Common Student Count']))

# Program lookup: ders -> (gun, slot)
program_dict = {}
for _, row in program.iterrows():
    code = row['Ders Kodu']
    if code and str(code) != 'nan':
        program_dict[code] = (int(row['Gün']), int(row['Slot']))

print(f"BMED4112 mevcut: G{program_dict.get('BMED4112', 'BULUNAMADI')}")
print(f"Toplam cakisan ders: {len(cakisan_dict)}")
print()

# Date mapping
gun_map = {
    1: "08.06.2026 Pazartesi",
    2: "09.06.2026 Sali",
    3: "10.06.2026 Carsamba",
    4: "11.06.2026 Persembe",
    5: "12.06.2026 Cuma",
    6: "15.06.2026 Pazartesi",
    7: "16.06.2026 Sali",
    8: "17.06.2026 Carsamba",
    9: "18.06.2026 Persembe",
    10: "19.06.2026 Cuma",
}
slot_map = {1: "08:30-11:30", 2: "11:30-14:30", 3: "14:30-17:30"}

# Constraint types for BMED4112
kisit_df = pd.read_excel('c:/Users/cihan.tazeoz/Desktop/26 BAHAR FİNAL/UYGULAMA DOSYALARI/ders_kisitleri.xlsx')
header = list(kisit_df.columns)
use_cols = header[:4]
first_col = str(header[0]).strip().lower()
if len(header) >= 5 and (first_col.startswith('unnamed') or str(header[0]).strip().isdigit()):
    use_cols = header[1:5]

kisit_types = {}
for idx, row in kisit_df.iterrows():
    dersler = []
    for col in kisit_df.columns:
        if pd.notna(row[col]):
            val = str(row[col]).strip().upper()
            if val and val not in ['NAN']:
                dersler.append(val)
    if 'BMED4112' in dersler:
        tip = None
        for col in use_cols:
            if pd.notna(row[col]):
                val = str(row[col]).strip().upper()
                if val in ['SAMEDAY', 'SAMESLOT', 'DIFFDAY', 'DIFFSLOT']:
                    tip = val
                    break
        if tip:
            for d in dersler:
                if d != 'BMED4112':
                    kisit_types[d] = tip

print("BMED4112 kisitlari:")
for d, t in kisit_types.items():
    loc = program_dict.get(d, 'YOK')
    print(f"  {d}: {t} (mevcut: G{loc})")
print()

print("=" * 100)
print("BMED4112 ALTERNATIF ANALIZI (G1-G6)")
print("=" * 100)
print(f"{'Slot':<50} {'Durum':<12} {'Cakisan':<8} {'Ogrenci':<8} Detay")
print("-" * 100)

uygunlar = []
for g in range(1, 7):
    for s in range(1, 4):
        sorunlar = []
        toplam_cakisan = 0
        
        # Collision check
        for ders, (dg, ds) in program_dict.items():
            if ders == 'BMED4112':
                continue
            if dg == g and ds == s and ders in cakisan_dict:
                ogrenci = cakisan_dict[ders]
                sorunlar.append(f"CAKISMA: {ders} ({ogrenci} ogrenci)")
                toplam_cakisan += ogrenci
        
        # Constraint check
        for ders, tip in kisit_types.items():
            dg, ds = program_dict.get(ders, (None, None))
            if dg is None:
                continue
            if tip == 'SAMESLOT' and (dg != g or ds != s):
                sorunlar.append(f"SAMESLOT: {ders} mevcut G{dg}S{ds}")
            elif tip == 'SAMEDAY' and dg != g:
                sorunlar.append(f"SAMEDAY: {ders} mevcut G{dg}S{ds}")
            elif tip == 'DIFFSLOT' and dg == g and ds == s:
                sorunlar.append(f"DIFFSLOT: {ders} mevcut G{dg}S{ds}")
            elif tip == 'DIFFDAY' and dg == g:
                sorunlar.append(f"DIFFDAY: {ders} mevcut G{dg}S{ds}")
        
        slot_str = f"G{g}S{s} ({gun_map.get(g)} {slot_map.get(s)})"
        durum = "UYGUN" if not sorunlar else f"OLUMSUZ ({len(sorunlar)})"
        
        if sorunlar:
            detay = " | ".join(sorunlar[:2])
            if len(sorunlar) > 2:
                detay += f" (+{len(sorunlar)-2} daha)"
        else:
            detay = "-"
        
        if not sorunlar:
            uygunlar.append((g, s, slot_str))
        
        print(f"{slot_str:<50} {durum:<12} {len([s for s in sorunlar if 'CAKISMA' in s]):<8} {toplam_cakisan:<8} {detay}")

print()
print("=" * 100)
print("UYGUN ALTERNATIFLER (G1-G6)")
print("=" * 100)
if uygunlar:
    for g, s, st in uygunlar:
        print(f"  {st}")
else:
    print("  G1-G6 arasinda tamamen uygun slot bulunamadi.")
    print("  En az sorunlu alternatifler icin detayli analiz gerekli.")
