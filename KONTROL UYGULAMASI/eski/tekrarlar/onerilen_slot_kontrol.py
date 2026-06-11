import pandas as pd
from datetime import datetime

# Dosya yollari
ana_program = r"C:\Users\cihan.tazeoz\Desktop\26 BAHAR FİNAL\UYGULAMA DOSYALARI\2026 Bahar Final Programı _03062026.xlsx"
collision_file = r"C:\Users\cihan.tazeoz\Desktop\26 BAHAR FİNAL\UYGULAMA DOSYALARI\final_exam_collisions_14_05_2026_10_48_47.csv"
kisit_file = r"C:\Users\cihan.tazeoz\Desktop\26 BAHAR FİNAL\UYGULAMA DOSYALARI\ders_kisitleri.xlsx"

EXCLUDED_PREFIXES = ('ARCH', 'IMIM', 'GITA', 'INAR')

# 1. Ana programi oku
df_ana = pd.read_excel(ana_program)

# Slot parse fonksiyonu
def parse_gs_from_program(row):
    tarih = row['Sınav Tarihi']
    baslangic = row['Sınav Başlangıç Saati']
    
    tarih_gun = {
        '2026-06-08': 1, '2026-06-09': 2, '2026-06-10': 3, '2026-06-11': 4,
        '2026-06-12': 5, '2026-06-15': 6, '2026-06-16': 7, '2026-06-17': 8,
        '2026-06-18': 9, '2026-06-19': 10,
    }
    
    if isinstance(tarih, datetime):
        tarih_str = tarih.strftime('%Y-%m-%d')
        gun = tarih_gun.get(tarih_str, 0)
    else:
        gun = 0
    
    if pd.isna(baslangic):
        slot = 0
    elif isinstance(baslangic, str):
        if baslangic == '08:30': slot = 1
        elif baslangic == '11:30': slot = 2
        elif baslangic == '14:30': slot = 3
        else: slot = 0
    else:
        from datetime import time
        if isinstance(baslangic, time):
            s = baslangic.strftime('%H:%M')
            if s == '08:30': slot = 1
            elif s == '11:30': slot = 2
            elif s == '14:30': slot = 3
            else: slot = 0
        else:
            slot = 0
    
    return gun, slot

schedule_map = {}
for _, row in df_ana.iterrows():
    code = str(row['Ders Kodu']).strip().upper()
    gun, slot = parse_gs_from_program(row)
    if gun > 0 and slot > 0:
        schedule_map[code] = (gun, slot)

# 2. Cakisma dosyasini oku
collisions = pd.read_csv(collision_file)
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

# 3. Kisit dosyasini oku
kisit = pd.read_excel(kisit_file)
kisit_clean = kisit.dropna(how='all')

def parse_ders_list(val):
    if pd.isna(val):
        return []
    return [d.strip().upper() for d in str(val).split(',') if d.strip()]

diff_day_pairs = set()
diff_slot_pairs = set()
same_day_pairs = set()
same_slot_pairs = set()

for _, row in kisit_clean.iterrows():
    ayni_gun = parse_ders_list(row.get('A (Ayni Gun)', row.get('A (Aynı Gün)', '')))
    for i in range(len(ayni_gun)):
        for j in range(i+1, len(ayni_gun)):
            same_day_pairs.add(tuple(sorted([ayni_gun[i], ayni_gun[j]])))
    
    ayni_slot = parse_ders_list(row.get('B (Ayni Slot)', row.get('B (Aynı Slot)', '')))
    for i in range(len(ayni_slot)):
        for j in range(i+1, len(ayni_slot)):
            same_slot_pairs.add(tuple(sorted([ayni_slot[i], ayni_slot[j]])))
    
    farkli_gun = parse_ders_list(row.get('C (Farkli Gun)', row.get('C (Farklı Gün)', '')))
    for i in range(len(farkli_gun)):
        for j in range(i+1, len(farkli_gun)):
            diff_day_pairs.add(tuple(sorted([farkli_gun[i], farkli_gun[j]])))
    
    farkli_slot = parse_ders_list(row.get('D (Farkli Slot)', row.get('D (Farklı Slot)', '')))
    for i in range(len(farkli_slot)):
        for j in range(i+1, len(farkli_slot)):
            diff_slot_pairs.add(tuple(sorted([farkli_slot[i], farkli_slot[j]])))

# 4. ONERILEN SLOTLAR
oneriler = {
    'TRAD2504': (1, 2),   # G1/S2
    'BUSI3632': (4, 2),   # G4/S2
    'LOGI4558': (8, 1),   # G8/S1
    'BUSI4572': (4, 1),   # G4/S1
}

print("=" * 80)
print("ONERILEN SLOT KONTROLU")
print("=" * 80)

for code, (new_g, new_s) in oneriler.items():
    print(f"\n{'='*80}")
    print(f"DERS: {code} -> G{new_g}/S{new_s}")
    print(f"{'='*80}")
    
    # Mevcut slot
    old = schedule_map.get(code)
    if old:
        print(f"Mevcut slot: G{old[0]}/S{old[1]}")
    else:
        print(f"Mevcut slot: YOK (programda yok)")
    
    # A) Cakisma kontrolu
    print(f"\n[A] CAKISMA ANALIZI (G{new_g}/S{new_s}):")
    cakisan = []
    for (c1, c2), count in collision_dict.items():
        if code in (c1, c2):
            other = c2 if c1 == code else c1
            other_slot = schedule_map.get(other)
            if other_slot == (new_g, new_s):
                cakisan.append((other, count))
    
    if cakisan:
        print(f"  [-]  SORUN! {len(cakisan)} adet cakisma var:")
        for other, count in sorted(cakisan, key=lambda x: -x[1]):
            print(f"     {other} | Ortak: {count} ogrenci")
    else:
        print(f"  [+]  Cakisma YOK")
    
    # B) DiffDay kontrolu
    print(f"\n[B] FARKLI GUN KISITI:")
    dd_sorun = []
    for a, b in diff_day_pairs:
        if code in (a, b):
            other = b if a == code else a
            other_slot = schedule_map.get(other)
            if other_slot and other_slot[0] == new_g:
                dd_sorun.append((other, other_slot))
    
    if dd_sorun:
        print(f"  [-]  SORUN! {len(dd_sorun)} adet farkli gun kisiti ihlali:")
        for other, slot in dd_sorun:
            print(f"     {other} | Ayni gun: G{slot[0]} (farkli gun olmali)")
    else:
        print(f"  [+]  Uygun")
    
    # C) DiffSlot kontrolu
    print(f"\n[C] FARKLI SLOT KISITI:")
    ds_sorun = []
    for a, b in diff_slot_pairs:
        if code in (a, b):
            other = b if a == code else a
            other_slot = schedule_map.get(other)
            if other_slot and other_slot == (new_g, new_s):
                ds_sorun.append((other, other_slot))
    
    if ds_sorun:
        print(f"  [-]  SORUN! {len(ds_sorun)} adet farkli slot kisiti ihlali:")
        for other, slot in ds_sorun:
            print(f"     {other} | Ayni gun+slot: G{slot[0]}/S{slot[1]} (farkli slot olmali)")
    else:
        print(f"  [+]  Uygun")
    
    # D) SameDay kontrolu
    print(f"\n[D] AYNI GUN KISITI:")
    sd_sorun = []
    for a, b in same_day_pairs:
        if code in (a, b):
            other = b if a == code else a
            other_slot = schedule_map.get(other)
            if other_slot and other_slot[0] != new_g:
                sd_sorun.append((other, other_slot))
    
    if sd_sorun:
        print(f"  [-]  SORUN! {len(sd_sorun)} adet ayni gun kisiti ihlali:")
        for other, slot in sd_sorun:
            print(f"     {other} | Farkli gun: G{slot[0]} (ayni gun olmali)")
    else:
        print(f"  [+]  Uygun veya ayni gun kisiti yok")
    
    # E) SameSlot kontrolu
    print(f"\n[E] AYNI SLOT KISITI:")
    ss_sorun = []
    for a, b in same_slot_pairs:
        if code in (a, b):
            other = b if a == code else a
            other_slot = schedule_map.get(other)
            if other_slot and other_slot != (new_g, new_s):
                ss_sorun.append((other, other_slot))
    
    if ss_sorun:
        print(f"  [-]  SORUN! {len(ss_sorun)} adet ayni slot kisiti ihlali:")
        for other, slot in ss_sorun:
            print(f"     {other} | Farkli slot: G{slot[0]}/S{slot[1]} (ayni slot olmali)")
    else:
        print(f"  [+]  Uygun veya ayni slot kisiti yok")
    
    # GENEL SONUC
    print(f"\n[SONUC] {code} -> G{new_g}/S{new_s}:")
    toplam_sorun = len(cakisan) + len(dd_sorun) + len(ds_sorun) + len(sd_sorun) + len(ss_sorun)
    if toplam_sorun == 0:
        print(f"  [+] [+] [+]  TAMAMEN UYGUN! Hicbir sorun yok.")
    else:
        print(f"  [!]   {toplam_sorun} adet sorun bulundu. Yukaridaki detaylara bakin.")

# TUM DERSLERIN BIRBIRIYLE ETKILESIMI
print(f"\n{'='*80}")
print("TUM ONERILEN DERSLERIN BIRBIRIYLE KONTROLU")
print(f"{'='*80}")

tum_kodlar = list(oneriler.keys())
for i in range(len(tum_kodlar)):
    for j in range(i+1, len(tum_kodlar)):
        c1, c2 = tum_kodlar[i], tum_kodlar[j]
        s1, s2 = oneriler[c1], oneriler[c2]
        
        key = tuple(sorted([c1, c2]))
        
        # Cakisma var mi?
        cakisma = collision_dict.get(key, 0)
        
        # Ayni slot mu?
        ayni_slot = (s1 == s2)
        
        if cakisma > 0 and ayni_slot:
            print(f"  [-]  KRITIK: {c1} - {c2} | Ortak: {cakisma} ogrenci | AYNI SLOT: G{s1[0]}/S{s1[1]}")
        elif cakisma > 0:
            print(f"  [+]  {c1} - {c2} | Ortak: {cakisma} ogrenci | Farkli slot (sorun yok)")
        elif ayni_slot:
            print(f"  [!]   {c1} - {c2} | Ortak ogrenci yok | Ayni slot G{s1[0]}/S{s1[1]}")
