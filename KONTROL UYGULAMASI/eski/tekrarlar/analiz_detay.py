import pandas as pd

# Dosya yolları
collision_file = r"C:\Users\cihan.tazeoz\Desktop\26 BAHAR FİNAL\UYGULAMA DOSYALARI\final_exam_collisions_14_05_2026_10_48_47.csv"
schedule_file = r"C:\Users\cihan.tazeoz\Desktop\26 BAHAR FİNAL\UYGULAMA DOSYALARI\schedule (5).xlsx"
kisit_file = r"C:\Users\cihan.tazeoz\Desktop\26 BAHAR FİNAL\UYGULAMA DOSYALARI\ders_kisitleri.xlsx"

# 1. Dosyaları oku
collisions = pd.read_csv(collision_file)
schedule = pd.read_excel(schedule_file)
kisit = pd.read_excel(kisit_file)

print("=" * 80)
print("VERİ YÜKLEME TAMAMLANDI")
print(f"Çakışma dosyası: {len(collisions)} satır")
print(f"Schedule dosyası: {len(schedule)} satır")
print(f"Kısıt dosyası: {len(kisit)} satır")
print("=" * 80)

# 2. Schedule dosyasından ders kodu -> (gün, slot) haritası oluştur
def parse_slot(slot_str):
    """G1/S3 -> (1, 3) veya G10/S2 -> (10, 2)"""
    if pd.isna(slot_str):
        return None
    slot_str = str(slot_str).strip().upper()
    # Format: GX/SY
    parts = slot_str.split('/')
    if len(parts) != 2:
        return None
    gun = int(parts[0].replace('G', ''))
    slot = int(parts[1].replace('S', ''))
    return (gun, slot)

schedule_map = {}
for _, row in schedule.iterrows():
    ders_kodu = str(row['Ders Kodu']).strip().upper()
    slot_val = parse_slot(row['Slotlar'])
    if slot_val:
        schedule_map[ders_kodu] = slot_val

print(f"\nSchedule'da slot atanmış ders sayısı: {len(schedule_map)}")

# 3. Çakışmaları analiz et - ARCH, İMİM, GİTA, INAR hariç
EXCLUDED_PREFIXES = ('ARCH', 'İMİM', 'GİTA', 'INAR')

cakisma_var = []      # Çakışma listesinde var ve schedule'da aynı G/S -> GERÇEK ÇAKIŞMA
cakisma_yok = []      # Çakışma listesinde var ama schedule'da farklı G veya farklı S -> ÇAKIŞMA YOK
schedule_bulunamadi = []  # Çakışma listesinde var ama schedule'da yok

for _, row in collisions.iterrows():
    c1 = str(row['Course1']).strip().upper()
    c2 = str(row['Course2']).strip().upper()
    count = row['Common Student Count']
    
    # Hariç tutulan prefix kontrolü
    if c1.startswith(EXCLUDED_PREFIXES) or c2.startswith(EXCLUDED_PREFIXES):
        continue
    
    s1 = schedule_map.get(c1)
    s2 = schedule_map.get(c2)
    
    if s1 is None or s2 is None:
        schedule_bulunamadi.append((c1, c2, count, s1, s2))
        continue
    
    if s1 == s2:
        cakisma_var.append((c1, c2, count, s1[0], s1[1], s2[0], s2[1]))
    else:
        cakisma_yok.append((c1, c2, count, s1[0], s1[1], s2[0], s2[1]))

print("\n" + "=" * 80)
print("ÇAKIŞMA ANALİZİ SONUÇLARI (ARCH, İMİM, GİTA, INAR hariç)")
print("=" * 80)
print(f"\nSchedule'da bulunamayan ders çifti: {len(schedule_bulunamadi)}")
print(f"GERÇEK ÇAKIŞMA (aynı G ve aynı S): {len(cakisma_var)}")
print(f"ÇAKIŞMA YOK (farklı G veya aynı G farklı S): {len(cakisma_yok)}")

# 4. Örnekler göster
print("\n" + "-" * 80)
print("GERÇEK ÇAKIŞMA OLAN ÖRNEKLER (ilk 15):")
print("-" * 80)
for i, (c1, c2, count, g1, s1, g2, s2) in enumerate(cakisma_var[:15], 1):
    print(f"{i}. {c1} - {c2} | Ortak öğrenci: {count} | G{g1}/S{s1} == G{g2}/S{s2}")

print("\n" + "-" * 80)
print("ÇAKIŞMA YOK (DOĞRU ATANMIŞ) ÖRNEKLER (ilk 15):")
print("-" * 80)
for i, (c1, c2, count, g1, s1, g2, s2) in enumerate(cakisma_yok[:15], 1):
    print(f"{i}. {c1} - {c2} | Ortak öğrenci: {count} | G{g1}/S{s1} vs G{g2}/S{s2}")

print("\n" + "-" * 80)
print("SCHEDULE'DA BULUNAMAYAN DERS ÇİFTLERİ (ilk 10):")
print("-" * 80)
for i, (c1, c2, count, s1, s2) in enumerate(schedule_bulunamadi[:10], 1):
    print(f"{i}. {c1} ({s1}) - {c2} ({s2}) | Ortak öğrenci: {count}")

# 5. Kısıt dosyasını da analiz et
print("\n" + "=" * 80)
print("KISIT DOSYASI ANALİZİ")
print("=" * 80)

# Kısıt dosyasını temizle
kisit_clean = kisit.dropna(how='all')
print(f"\nKısıt dosyası temiz satır sayısı: {len(kisit_clean)}")
print("\nKolonlar:", list(kisit.columns))

# Kısıtları parse et
def parse_ders_list(val):
    if pd.isna(val):
        return []
    return [d.strip().upper() for d in str(val).split(',') if d.strip()]

ayni_gun_kisit = []  # [(ders1, ders2), ...]
ayni_slot_kisit = []
farkli_gun_kisit = []
farkli_slot_kisit = []

for _, row in kisit_clean.iterrows():
    # Aynı gün kısıtları
    ayni_gun = parse_ders_list(row.get('A (Aynı Gün)', ''))
    for i in range(len(ayni_gun)):
        for j in range(i+1, len(ayni_gun)):
            ayni_gun_kisit.append((ayni_gun[i], ayni_gun[j]))
    
    # Aynı slot kısıtları
    ayni_slot = parse_ders_list(row.get('B (Aynı Slot)', ''))
    for i in range(len(ayni_slot)):
        for j in range(i+1, len(ayni_slot)):
            ayni_slot_kisit.append((ayni_slot[i], ayni_slot[j]))
    
    # Farklı gün kısıtları
    farkli_gun = parse_ders_list(row.get('C (Farklı Gün)', ''))
    for i in range(len(farkli_gun)):
        for j in range(i+1, len(farkli_gun)):
            farkli_gun_kisit.append((farkli_gun[i], farkli_gun[j]))
    
    # Farklı slot kısıtları
    farkli_slot = parse_ders_list(row.get('D (Farklı Slot)', ''))
    for i in range(len(farkli_slot)):
        for j in range(i+1, len(farkli_slot)):
            farkli_slot_kisit.append((farkli_slot[i], farkli_slot[j]))

print(f"\nAynı gün kısıt çifti: {len(ayni_gun_kisit)}")
print(f"Aynı slot kısıt çifti: {len(ayni_slot_kisit)}")
print(f"Farklı gün kısıt çifti: {len(farkli_gun_kisit)}")
print(f"Farklı slot kısıt çifti: {len(farkli_slot_kisit)}")

# Kısıt uyumluluk kontrolü
print("\n" + "-" * 80)
print("KISIT UYUMLULUK KONTROLÜ")
print("-" * 80)

ayni_gun_ihlal = []
ayni_slot_ihlal = []
farkli_gun_ihlal = []
farkli_slot_ihlal = []

for d1, d2 in ayni_gun_kisit:
    s1 = schedule_map.get(d1)
    s2 = schedule_map.get(d2)
    if s1 and s2 and s1[0] != s2[0]:
        ayni_gun_ihlal.append((d1, d2, s1, s2))

for d1, d2 in ayni_slot_kisit:
    s1 = schedule_map.get(d1)
    s2 = schedule_map.get(d2)
    if s1 and s2 and s1[1] != s2[1]:
        ayni_slot_ihlal.append((d1, d2, s1, s2))

for d1, d2 in farkli_gun_kisit:
    s1 = schedule_map.get(d1)
    s2 = schedule_map.get(d2)
    if s1 and s2 and s1[0] == s2[0]:
        farkli_gun_ihlal.append((d1, d2, s1, s2))

for d1, d2 in farkli_slot_kisit:
    s1 = schedule_map.get(d1)
    s2 = schedule_map.get(d2)
    if s1 and s2 and s1[1] == s2[1]:
        farkli_slot_ihlal.append((d1, d2, s1, s2))

print(f"Aynı gün kısıtı İHLAL: {len(ayni_gun_ihlal)}")
if ayni_gun_ihlal:
    for d1, d2, s1, s2 in ayni_gun_ihlal[:10]:
        print(f"  {d1} (G{s1[0]}/S{s1[1]}) - {d2} (G{s2[0]}/S{s2[1]})")

print(f"\nAynı slot kısıtı İHLAL: {len(ayni_slot_ihlal)}")
if ayni_slot_ihlal:
    for d1, d2, s1, s2 in ayni_slot_ihlal[:10]:
        print(f"  {d1} (G{s1[0]}/S{s1[1]}) - {d2} (G{s2[0]}/S{s2[1]})")

print(f"\nFarklı gün kısıtı İHLAL: {len(farkli_gun_ihlal)}")
if farkli_gun_ihlal:
    for d1, d2, s1, s2 in farkli_gun_ihlal[:10]:
        print(f"  {d1} (G{s1[0]}/S{s1[1]}) - {d2} (G{s2[0]}/S{s2[1]})")

print(f"\nFarklı slot kısıtı İHLAL: {len(farkli_slot_ihlal)}")
if farkli_slot_ihlal:
    for d1, d2, s1, s2 in farkli_slot_ihlal[:10]:
        print(f"  {d1} (G{s1[0]}/S{s1[1]}) - {d2} (G{s2[0]}/S{s2[1]})")

# 6. Özet
print("\n" + "=" * 80)
print("GENEL DEĞERLENDİRME")
print("=" * 80)
total_kisit_ihlal = len(ayni_gun_ihlal) + len(ayni_slot_ihlal) + len(farkli_gun_ihlal) + len(farkli_slot_ihlal)
print(f"""
1. ÇAKIŞMA ANALİZİ (Öğrenci bazlı):
   - Toplam çakışma çifti (hariçler hariç): {len(cakisma_var) + len(cakisma_yok)}
   - GERÇEK ÇAKIŞMA (aynı G+S): {len(cakisma_var)}
   - ÇAKIŞMA YOK (doğru ayrılmış): {len(cakisma_yok)}
   - Schedule'da bulunamayan: {len(schedule_bulunamadi)}

2. KISİYAT İHLALLERİ:
   - Toplam kısıt ihlali: {total_kisit_ihlal}
   - Aynı gün ihlali: {len(ayni_gun_ihlal)}
   - Aynı slot ihlali: {len(ayni_slot_ihlal)}
   - Farklı gün ihlali: {len(farkli_gun_ihlal)}
   - Farklı slot ihlali: {len(farkli_slot_ihlal)}
""")

if len(cakisma_var) == 0 and total_kisit_ihlal == 0:
    print("SONUÇ: Schedule dosyası MÜKEMMEL görünüyor! Çakışma yok, kısıt ihlali yok.")
elif len(cakisma_var) == 0:
    print("SONUÇ: Öğrenci çakışması YOK ancak bazı kısıtlar ihlal edilmiş.")
else:
    print(f"SONUÇ: {len(cakisma_var)} adet GERÇEK ÇAKIŞMA var! Schedule dosyası düzeltilmeli.")
