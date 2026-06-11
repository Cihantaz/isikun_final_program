import pandas as pd
from datetime import datetime

# Dosya yollari
ana_program = r"C:\Users\cihan.tazeoz\Desktop\26 BAHAR FİNAL\UYGULAMA DOSYALARI\2026 Bahar Final Programı _03062026.xlsx"
collision_file = r"C:\Users\cihan.tazeoz\Desktop\26 BAHAR FİNAL\UYGULAMA DOSYALARI\final_exam_collisions_14_05_2026_10_48_47.csv"
kisit_file = r"C:\Users\cihan.tazeoz\Desktop\26 BAHAR FİNAL\UYGULAMA DOSYALARI\ders_kisitleri.xlsx"

EXCLUDED_PREFIXES = ('ARCH', 'IMIM', 'GITA', 'INAR')

# Gun -> Tarih mapping
gun_tarih = {
    1: ('08.06.2026', 'Pazartesi'), 2: ('09.06.2026', 'Salı'),
    3: ('10.06.2026', 'Çarşamba'), 4: ('11.06.2026', 'Perşembe'),
    5: ('12.06.2026', 'Cuma'), 6: ('15.06.2026', 'Pazartesi'),
    7: ('16.06.2026', 'Salı'), 8: ('17.06.2026', 'Çarşamba'),
    9: ('18.06.2026', 'Perşembe'), 10: ('19.06.2026', 'Cuma'),
}
slot_saat = {1: '08:30 - 11:30', 2: '11:30 - 14:30', 3: '14:30 - 17:30'}

def gs_yaz(g, s):
    tarih, gun = gun_tarih.get(g, ('?', '?'))
    saat = slot_saat.get(s, '?')
    return f'{tarih} {gun}  {saat}'

# Ana programi oku
df_ana = pd.read_excel(ana_program)

def parse_gs_from_program(row):
    tarih = row['Sınav Tarihi']
    baslangic = row['Sınav Başlangıç Saati']
    tarih_gun = {'2026-06-08': 1, '2026-06-09': 2, '2026-06-10': 3, '2026-06-11': 4,
        '2026-06-12': 5, '2026-06-15': 6, '2026-06-16': 7, '2026-06-17': 8,
        '2026-06-18': 9, '2026-06-19': 10}
    if isinstance(tarih, datetime):
        gun = tarih_gun.get(tarih.strftime('%Y-%m-%d'), 0)
    else:
        gun = 0
    if pd.isna(baslangic):
        slot = 0
    elif isinstance(baslangic, str):
        slot = {'08:30': 1, '11:30': 2, '14:30': 3}.get(baslangic, 0)
    else:
        from datetime import time
        if isinstance(baslangic, time):
            slot = {'08:30': 1, '11:30': 2, '14:30': 3}.get(baslangic.strftime('%H:%M'), 0)
        else:
            slot = 0
    return gun, slot

schedule_map = {}
for _, row in df_ana.iterrows():
    code = str(row['Ders Kodu']).strip().upper()
    gun, slot = parse_gs_from_program(row)
    if gun > 0 and slot > 0:
        schedule_map[code] = (gun, slot)

# Cakisma dosyasi
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

# Kisit dosyasi
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
    for col, target in [(row.get('A (Ayni Gun)', row.get('A (Aynı Gün)', '')), same_day_pairs),
                        (row.get('B (Ayni Slot)', row.get('B (Aynı Slot)', '')), same_slot_pairs),
                        (row.get('C (Farkli Gun)', row.get('C (Farklı Gün)', '')), diff_day_pairs),
                        (row.get('D (Farkli Slot)', row.get('D (Farklı Slot)', '')), diff_slot_pairs)]:
        dersler = parse_ders_list(col)
        for i in range(len(dersler)):
            for j in range(i+1, len(dersler)):
                target.add(tuple(sorted([dersler[i], dersler[j]])))

# Kontrol fonksiyonu
def kontrol_et(code, new_g, new_s):
    sorunlar = []
    
    # Cakisma
    for (c1, c2), count in collision_dict.items():
        if code in (c1, c2):
            other = c2 if c1 == code else c1
            other_slot = schedule_map.get(other)
            if other_slot == (new_g, new_s):
                sorunlar.append(f'CAKISMA: {other} ({count} ogrenci)')
    
    # DiffDay
    for a, b in diff_day_pairs:
        if code in (a, b):
            other = b if a == code else a
            other_slot = schedule_map.get(other)
            if other_slot and other_slot[0] == new_g:
                sorunlar.append(f'DIFFDAY: {other} (ayni gun)')
    
    # DiffSlot
    for a, b in diff_slot_pairs:
        if code in (a, b):
            other = b if a == code else a
            other_slot = schedule_map.get(other)
            if other_slot and other_slot == (new_g, new_s):
                sorunlar.append(f'DIFFSLOT: {other} (ayni gun+slot)')
    
    # SameDay
    for a, b in same_day_pairs:
        if code in (a, b):
            other = b if a == code else a
            other_slot = schedule_map.get(other)
            if other_slot and other_slot[0] != new_g:
                sorunlar.append(f'SAMEDAY: {other} (farkli gun)')
    
    # SameSlot
    for a, b in same_slot_pairs:
        if code in (a, b):
            other = b if a == code else a
            other_slot = schedule_map.get(other)
            if other_slot and other_slot != (new_g, new_s):
                sorunlar.append(f'SAMESLOT: {other} (farkli slot)')
    
    return sorunlar

# TUM ONERILER
print("=" * 110)
print("TUM ONERILERIN DETAYLI ANALIZI (BUSI3222 G2/S1 DAHIL)")
print("=" * 110)

oneriler = [
    ('TRAD2504', (1, 1), (1, 2), (1, 3)),
    ('BUSI3632', (2, 2), (4, 2), (4, 3)),
    ('LOGI4558', (2, 1), (8, 1), None),
    ('BUSI4572', (3, 1), (4, 1), (5, 1)),
    ('BUSI3222', (4, 1), (2, 1), None),
]

print(f"\n{'Ders':<12} {'Mevcut (Tarih/Gun/Saat)':<45} {'Istenen':<45} {'Alternatif':<45} {'Durum'}")
print("-" * 110)

for code, mevcut, istenen, alternatif in oneriler:
    mg, ms = mevcut
    ig, is_ = istenen
    ag, as_ = alternatif if alternatif else (0, 0)
    
    mevcut_str = gs_yaz(mg, ms)
    istenen_str = gs_yaz(ig, is_)
    alt_str = gs_yaz(ag, as_) if alternatif else '-'
    
    # Istenen kontrol
    istenen_sorun = kontrol_et(code, ig, is_)
    # Alternatif kontrol
    alt_sorun = []
    if alternatif:
        alt_sorun = kontrol_et(code, ag, as_)
    
    if istenen_sorun:
        durum = f"OLUMSUZ ({len(istenen_sorun)} sorun)"
    else:
        durum = 'UYGUN'
    
    print(f"{code:<12} {mevcut_str:<45} {istenen_str:<45} {alt_str:<45} {durum}")
    
    if istenen_sorun:
        for s in istenen_sorun[:3]:
            print(f"  -> {s}")

# BUSI3222 ozel analiz
print("\n" + "=" * 110)
print("BUSI3222 -> G2/S1 OZEL ANALIZ")
print("=" * 110)
busi3222_sorun = kontrol_et('BUSI3222', 2, 1)
if busi3222_sorun:
    print(f"SORUNLAR ({len(busi3222_sorun)}):")
    for s in busi3222_sorun:
        print(f"  - {s}")
else:
    print("TAMAMEN UYGUN! Hicbir sorun yok.")

# 5 dersin birbiriyle cakismasi
print("\n" + "=" * 110)
print("5 DERSIN BIRBIRIYLE KONTROLU (YENI SLOTLAR)")
print("=" * 110)

yeni_slotlar = {
    'TRAD2504': (1, 3),
    'BUSI3632': (4, 3),
    'LOGI4558': (8, 1),
    'BUSI4572': (5, 1),
    'BUSI3222': (2, 1),
}

for i, (c1, s1) in enumerate(yeni_slotlar.items()):
    for j, (c2, s2) in enumerate(yeni_slotlar.items()):
        if i < j:
            key = tuple(sorted([c1, c2]))
            count = collision_dict.get(key, 0)
            if s1 == s2 and count > 0:
                print(f"[-] {c1} - {c2} | AYNI SLOT | Ortak: {count} ogrenci | {gs_yaz(*s1)}")
            elif s1 == s2:
                print(f"[!] {c1} - {c2} | AYNI SLOT | Ortak ogrenci yok | {gs_yaz(*s1)}")
            elif count > 0:
                print(f"[+] {c1} - {c2} | Farkli slot | Ortak: {count} ogrenci | {gs_yaz(*s1)} vs {gs_yaz(*s2)}")
