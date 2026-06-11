import pandas as pd
import os
from datetime import datetime, time

# Dosya yollari
ana_program = r"C:\Users\cihan.tazeoz\Desktop\26 BAHAR FİNAL\UYGULAMA DOSYALARI\26 bahar final _ son.xlsx"
collision_file = r"C:\Users\cihan.tazeoz\Desktop\26 BAHAR FİNAL\UYGULAMA DOSYALARI\final_exam_collisions_14_05_2026_10_48_47.csv"
kisit_file = r"C:\Users\cihan.tazeoz\Desktop\26 BAHAR FİNAL\UYGULAMA DOSYALARI\ders_kisitleri.xlsx"

EXCLUDED_PREFIXES = ('ARCH', 'IMIM', 'GITA', 'INAR')

# 1. Ana programi oku
df_ana = pd.read_excel(ana_program)

# Tarih -> Gun mapping (haftasonu dahil, G sayisi gun sirasina gore)
tarihler_raw = df_ana['Sınav Tarihi'].dropna().unique()
tarihler = []
for t in tarihler_raw:
    if isinstance(t, datetime):
        tarihler.append(t)
tarihler = sorted(tarihler)
# Sadece hafta ici gunleri al
tarih_gun = {}
gun_sayac = 1
for t in tarihler:
    if isinstance(t, datetime):
        tarih_str = t.strftime('%Y-%m-%d')
        # Haftasonu mu kontrol et (5=Cuma, 6=Cumartesi, 0=Pazar)
        wd = t.weekday()
        if wd < 5:  # Hafta ici
            tarih_gun[tarih_str] = gun_sayac
            gun_sayac += 1

print("Tarih -> Gun mapping:")
for t, g in tarih_gun.items():
    print(f"  {t} -> G{g}")

# Saat -> Slot mapping
def get_slot(baslangic):
    if pd.isna(baslangic):
        return 0
    if isinstance(baslangic, time):
        s = baslangic.strftime('%H:%M')
    else:
        s = str(baslangic).strip()
    if s == '08:30':
        return 1
    elif s == '11:30':
        return 2
    elif s == '14:30':
        return 3
    else:
        return 0  # Bilinmeyen slot

# Parse et
schedule_map = {}
fakulte_map = {}
bilinmeyen_slot = []
for _, row in df_ana.iterrows():
    code = str(row['Ders Kodu']).strip().upper()
    fakulte = str(row.get('Fakülte Adı', '')).strip()
    tarih = row['Sınav Tarihi']
    baslangic = row['Sınav Başlangıç Saati']
    
    if isinstance(tarih, datetime):
        tarih_str = tarih.strftime('%Y-%m-%d')
        gun = tarih_gun.get(tarih_str, 0)
    else:
        gun = 0
    
    slot = get_slot(baslangic)
    
    if gun > 0 and slot > 0:
        schedule_map[code] = (gun, slot)
        fakulte_map[code] = fakulte
    else:
        if gun > 0:
            bilinmeyen_slot.append((code, tarih_str, str(baslangic)))

print(f"\nAna programda parse edilen ders: {len(schedule_map)}")
print(f"Bilinmeyen slotlu ders: {len(bilinmeyen_slot)}")
for c, t, s in bilinmeyen_slot[:10]:
    print(f"  {c}: {t} {s}")

# 2. Cakisma dosyasini oku
collisions = pd.read_csv(collision_file)
collision_pairs = {}
for _, row in collisions.iterrows():
    c1 = str(row['Course1']).strip().upper()
    c2 = str(row['Course2']).strip().upper()
    count = row['Common Student Count']
    n1 = c1.replace('İ','I').replace('Ş','S').replace('Ğ','G').replace('Ü','U').replace('Ö','O').replace('Ç','C')
    n2 = c2.replace('İ','I').replace('Ş','S').replace('Ğ','G').replace('Ü','U').replace('Ö','O').replace('Ç','C')
    if not n1.startswith(EXCLUDED_PREFIXES) and not n2.startswith(EXCLUDED_PREFIXES):
        key = tuple(sorted([c1, c2]))
        if key not in collision_pairs or collision_pairs[key] < count:
            collision_pairs[key] = count

# Cakisma analizi
cakisma_var = []
for (c1, c2), count in collision_pairs.items():
    s1 = schedule_map.get(c1)
    s2 = schedule_map.get(c2)
    if s1 and s2 and s1 == s2:
        cakisma_var.append((c1, c2, count, s1[0], s1[1]))

print(f"\nGERCEK CAKISMA (ayni gun+slot): {len(cakisma_var)}")
if cakisma_var:
    for c1, c2, count, g, s in cakisma_var[:20]:
        print(f"  {c1} - {c2} | Ortak: {count} | G{g}/S{s}")

# 3. Kisit dosyasini oku
kisit = pd.read_excel(kisit_file)
kisit_clean = kisit.dropna(how='all')

def parse_ders_list(val):
    if pd.isna(val):
        return []
    return [d.strip().upper() for d in str(val).split(',') if d.strip()]

# Farkli gun ihlalleri
dd_ihlal = []
for _, row in kisit_clean.iterrows():
    val = row.get('C (Farkli Gun)', row.get('C (Farklı Gün)', ''))
    dersler = parse_ders_list(val)
    for i in range(len(dersler)):
        for j in range(i+1, len(dersler)):
            d1, d2 = dersler[i], dersler[j]
            s1 = schedule_map.get(d1)
            s2 = schedule_map.get(d2)
            if s1 and s2 and s1[0] == s2[0]:
                n1 = d1.replace('İ','I').replace('Ş','S').replace('Ğ','G').replace('Ü','U').replace('Ö','O').replace('Ç','C')
                n2 = d2.replace('İ','I').replace('Ş','S').replace('Ğ','G').replace('Ü','U').replace('Ö','O').replace('Ç','C')
                if not n1.startswith(EXCLUDED_PREFIXES) and not n2.startswith(EXCLUDED_PREFIXES):
                    dd_ihlal.append((d1, d2, s1, s2))

# Farkli slot ihlalleri (ayni gun + ayni slot)
ds_ihlal = []
for _, row in kisit_clean.iterrows():
    val = row.get('D (Farkli Slot)', row.get('D (Farklı Slot)', ''))
    dersler = parse_ders_list(val)
    for i in range(len(dersler)):
        for j in range(i+1, len(dersler)):
            d1, d2 = dersler[i], dersler[j]
            s1 = schedule_map.get(d1)
            s2 = schedule_map.get(d2)
            if s1 and s2 and s1[0] == s2[0] and s1[1] == s2[1]:
                n1 = d1.replace('İ','I').replace('Ş','S').replace('Ğ','G').replace('Ü','U').replace('Ö','O').replace('Ç','C')
                n2 = d2.replace('İ','I').replace('Ş','S').replace('Ğ','G').replace('Ü','U').replace('Ö','O').replace('Ç','C')
                if not n1.startswith(EXCLUDED_PREFIXES) and not n2.startswith(EXCLUDED_PREFIXES):
                    ds_ihlal.append((d1, d2, s1, s2))

print(f"\nFarkli gun ihlali: {len(dd_ihlal)}")
if dd_ihlal:
    for d1, d2, s1, s2 in dd_ihlal[:15]:
        print(f"  {d1} ({s1[0]}/S{s1[1]}) - {d2} ({s2[0]}/S{s2[1]})")

print(f"\nFarkli slot ihlali (ayni gun+slot): {len(ds_ihlal)}")
if ds_ihlal:
    for d1, d2, s1, s2 in ds_ihlal[:15]:
        print(f"  {d1} ({s1[0]}/S{s1[1]}) - {d2} ({s2[0]}/S{s2[1]})")

# 4. ITEC4431 analizi
print("\n" + "=" * 70)
print("ITEC4431 ANALIZI")
print("=" * 70)

itec_code = 'ITEC4431'
itec_slot = schedule_map.get(itec_code)
print(f"ITEC4431 mevcut slot: {itec_slot}")

# ITEC4431 ile cakisan dersleri bul
itec_cakisan = []
for (c1, c2), count in collision_pairs.items():
    if itec_code in (c1, c2):
        other = c2 if c1 == itec_code else c1
        other_slot = schedule_map.get(other)
        if other_slot:
            itec_cakisan.append((other, count, other_slot))

print(f"ITEC4431 ile cakisan ders sayisi: {len(itec_cakisan)}")
if itec_cakisan:
    print("Cakisan dersler:")
    for other, count, other_slot in sorted(itec_cakisan, key=lambda x: -x[1])[:15]:
        print(f"  {other} | Ortak: {count} | Slot: G{other_slot[0]}/S{other_slot[1]}")

# ITEC4431 icin uygun slotlar
n_days = 10
spd = 3
musait_slotlar = []
for g in range(1, n_days+1):
    for s in range(1, spd+1):
        uygun = True
        for other, count, other_slot in itec_cakisan:
            if other_slot == (g, s):
                uygun = False
                break
        if uygun:
            musait_slotlar.append((g, s))

print(f"\nITEC4431 icin UYGUN slotlar ({len(musait_slotlar)} adet):")
for g, s in musait_slotlar:
    print(f"  G{g}/S{s}")

# 5. Fakulte bazli analiz
print("\n" + "=" * 70)
print("FAKULTE BAZLI ANALIZ")
print("=" * 70)

fakulte_stats = {}
for code, fakulte in fakulte_map.items():
    if not fakulte:
        fakulte = 'Bilinmiyor'
    if fakulte not in fakulte_stats:
        fakulte_stats[fakulte] = {'toplam': 0, 'cakisma': 0, 'dd': 0, 'ds': 0}
    fakulte_stats[fakulte]['toplam'] += 1

# Cakismalari fakulteye gore dagil
for c1, c2, count, g, s in cakisma_var:
    f1 = fakulte_map.get(c1, 'Bilinmiyor')
    f2 = fakulte_map.get(c2, 'Bilinmiyor')
    if f1 in fakulte_stats:
        fakulte_stats[f1]['cakisma'] += 1
    if f2 in fakulte_stats and f2 != f1:
        fakulte_stats[f2]['cakisma'] += 1

# Ihlalleri fakulteye gore dagil
for d1, d2, s1, s2 in dd_ihlal:
    f1 = fakulte_map.get(d1, 'Bilinmiyor')
    f2 = fakulte_map.get(d2, 'Bilinmiyor')
    if f1 in fakulte_stats:
        fakulte_stats[f1]['dd'] += 1
    if f2 in fakulte_stats and f2 != f1:
        fakulte_stats[f2]['dd'] += 1

for d1, d2, s1, s2 in ds_ihlal:
    f1 = fakulte_map.get(d1, 'Bilinmiyor')
    f2 = fakulte_map.get(d2, 'Bilinmiyor')
    if f1 in fakulte_stats:
        fakulte_stats[f1]['ds'] += 1
    if f2 in fakulte_stats and f2 != f1:
        fakulte_stats[f2]['ds'] += 1

print(f"\n{'Fakulte':<35} {'Toplam':<8} {'Cakisma':<8} {'DD':<8} {'DS':<8}")
print("-" * 70)
for fakulte, stats in sorted(fakulte_stats.items()):
    print(f"{fakulte:<35} {stats['toplam']:<8} {stats['cakisma']:<8} {stats['dd']:<8} {stats['ds']:<8}")

# Eksik dersler
print("\n" + "=" * 70)
print("EKSIK DERSLER (Ana programda olmayan ama cakismada olan)")
print("=" * 70)

missing = set()
for (c1, c2), count in collision_pairs.items():
    if c1 not in schedule_map:
        missing.add(c1)
    if c2 not in schedule_map:
        missing.add(c2)

print(f"Toplam eksik ders: {len(missing)}")
for c in sorted(missing)[:30]:
    print(f"  {c}")
