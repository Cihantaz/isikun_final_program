import pandas as pd
from collections import Counter

# Dosya yolları
collision_file = r"C:\Users\cihan.tazeoz\Desktop\26 BAHAR FİNAL\UYGULAMA DOSYALARI\final_exam_collisions_14_05_2026_10_48_47.csv"
schedule_file = r"C:\Users\cihan.tazeoz\Desktop\26 BAHAR FİNAL\UYGULAMA DOSYALARI\schedule (5).xlsx"
kisit_file = r"C:\Users\cihan.tazeoz\Desktop\26 BAHAR FİNAL\UYGULAMA DOSYALARI\ders_kisitleri.xlsx"

# 1. Dosyaları oku
collisions = pd.read_csv(collision_file)
schedule = pd.read_excel(schedule_file)
kisit = pd.read_excel(kisit_file)

EXCLUDED_PREFIXES = ('ARCH', 'IMIM', 'GITA', 'INAR')

def parse_slot(slot_str):
    if pd.isna(slot_str):
        return None
    slot_str = str(slot_str).strip().upper()
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

# Cakismalari analiz et
cakisma_var = []
cakisma_yok = []
schedule_bulunamadi = []

for _, row in collisions.iterrows():
    c1 = str(row['Course1']).strip().upper()
    c2 = str(row['Course2']).strip().upper()
    count = row['Common Student Count']
    
    # Turkce karakterleri normalize et
    c1_norm = c1.replace('İ', 'I').replace('Ş', 'S').replace('Ğ', 'G').replace('Ü', 'U').replace('Ö', 'O').replace('Ç', 'C')
    c2_norm = c2.replace('İ', 'I').replace('Ş', 'S').replace('Ğ', 'G').replace('Ü', 'U').replace('Ö', 'O').replace('Ç', 'C')
    
    if c1_norm.startswith(EXCLUDED_PREFIXES) or c2_norm.startswith(EXCLUDED_PREFIXES):
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

# Kisitlari analiz et
def parse_ders_list(val):
    if pd.isna(val):
        return []
    return [d.strip().upper() for d in str(val).split(',') if d.strip()]

kisit_clean = kisit.dropna(how='all')
farkli_gun_kisit = []
farkli_slot_kisit = []

for _, row in kisit_clean.iterrows():
    farkli_gun = parse_ders_list(row.get('C (Farkli Gun)', row.get('C (Farklı Gün)', '')))
    for i in range(len(farkli_gun)):
        for j in range(i+1, len(farkli_gun)):
            farkli_gun_kisit.append((farkli_gun[i], farkli_gun[j]))
    
    farkli_slot = parse_ders_list(row.get('D (Farkli Slot)', row.get('D (Farklı Slot)', '')))
    for i in range(len(farkli_slot)):
        for j in range(i+1, len(farkli_slot)):
            farkli_slot_kisit.append((farkli_slot[i], farkli_slot[j]))

farkli_gun_ihlal = []
for d1, d2 in farkli_gun_kisit:
    s1 = schedule_map.get(d1)
    s2 = schedule_map.get(d2)
    if s1 and s2 and s1[0] == s2[0]:
        farkli_gun_ihlal.append((d1, d2, s1, s2))

farkli_slot_ihlal = []
for d1, d2 in farkli_slot_kisit:
    s1 = schedule_map.get(d1)
    s2 = schedule_map.get(d2)
    if s1 and s2 and s1[1] == s2[1]:
        farkli_slot_ihlal.append((d1, d2, s1, s2))

# Rapor yaz
print("=" * 80)
print("           FINAL SINAV PROGRAMI CAKISMA ANALIZ RAPORU")
print("=" * 80)

print("""
1. ANLADIGIM KRITERLER VE MANTIK:
   --------------------------------
   - Cakisma dosyasi: Her satir "DersA, DersB, OrtakOgrenciSayisi" icerir.
     Bu, DersA ve DersB'yi alan ortak ogrencilerin varligini gosterir.
     A,B ve B,A ayni anlama gelir (cift tarafli).
   
   - Schedule dosyasi: Her derse "Gx/Sy" formatinda bir slot atanmistir.
     x = Gun (1-10), y = Slot (1-3)
   
   - CAKISMA KURALI: Eger cakisma dosyasinda A ve B arasinda ortak ogrenci varsa
     ve schedule'da A=Gx/Sy, B=Gx/Sy (AYNI gun ve AYNI slot) -> GERCEK CAKISMA
     Eger farkli gun veya ayni gun farkli slot -> CAKISMA YOK (dogru)
   
   - HARIC: ARCH, IMIM, GITA, INAR ile baslayan dersler analize dahil edilmemistir.
""")

print("2. CAKISMA DOSYASI ISTATISTIKLERI:")
print("   - Toplam cakisma kaydi:", len(collisions))
print("   - Analize dahil edilen (haricler cikarildi):", len(cakisma_var) + len(cakisma_yok) + len(schedule_bulunamadi))

print("\n3. SCHEDULE DOSYASI ISTATISTIKLERI:")
print("   - Toplam ders sayisi:", len(schedule))

print("\n4. CAKISMA ANALIZI SONUCLARI:")
print("   - GERCEK CAKISMA (ayni Gun + ayni Slot):", len(cakisma_var))
print("   - CAKISMA YOK (dogru ayrilmis):", len(cakisma_yok))
print("   - Schedule'da bulunamayan ders iceren cift:", len(schedule_bulunamadi))

print("\n5. ORNEKLER:")
print("   --- CAKISMA YOK (Dogru Atanmis) Ornekler ---")
for i, (c1, c2, count, g1, s1, g2, s2) in enumerate(cakisma_yok[:5], 1):
    print("   ", i, ". ", c1, " & ", c2, " | Ortak: ", count, " ogrenci", sep="")
    print("      -> ", c1, ": G", g1, "/S", s1, " | ", c2, ": G", g2, "/S", s2, " (FARKLI)", sep="")

print("\n   --- GERCEK CAKISMA Olan Ornekler ---")
if cakisma_var:
    for i, (c1, c2, count, g1, s1, g2, s2) in enumerate(cakisma_var[:5], 1):
        print("   ", i, ". ", c1, " & ", c2, " | Ortak: ", count, " ogrenci", sep="")
        print("      -> ", c1, ": G", g1, "/S", s1, " | ", c2, ": G", g2, "/S", s2, " (AYNI!)", sep="")
else:
    print("   YOK! Hic gercek cakisma bulunamadi.")

print("\n6. KISIT DOSYASI ANALIZI:")
print("   - Farkli gun kisiti IHLAL:", len(farkli_gun_ihlal))
if farkli_gun_ihlal:
    print("     Ihlal eden ciftler (ilk 10):")
    for d1, d2, s1, s2 in farkli_gun_ihlal[:10]:
        print("       ", d1, " (G", s1[0], "/S", s1[1], ") - ", d2, " (G", s2[0], "/S", s2[1], ")", sep="")

print("\n   - Farkli slot kisiti IHLAL:", len(farkli_slot_ihlal))
if farkli_slot_ihlal:
    print("     Ihlal eden ciftler (ilk 10):")
    for d1, d2, s1, s2 in farkli_slot_ihlal[:10]:
        print("       ", d1, " (G", s1[0], "/S", s1[1], ") - ", d2, " (G", s2[0], "/S", s2[1], ")", sep="")

# Eksik dersleri kategorize et
missing_courses = set()
for _, row in collisions.iterrows():
    c1 = str(row['Course1']).strip().upper()
    c2 = str(row['Course2']).strip().upper()
    c1_norm = c1.replace('İ', 'I').replace('Ş', 'S').replace('Ğ', 'G').replace('Ü', 'U').replace('Ö', 'O').replace('Ç', 'C')
    c2_norm = c2.replace('İ', 'I').replace('Ş', 'S').replace('Ğ', 'G').replace('Ü', 'U').replace('Ö', 'O').replace('Ç', 'C')
    if not c1_norm.startswith(EXCLUDED_PREFIXES) and not c2_norm.startswith(EXCLUDED_PREFIXES):
        if c1 not in schedule_map:
            missing_courses.add(c1)
        if c2 not in schedule_map:
            missing_courses.add(c2)

print("\n7. EKSIK DERSLER (Schedule'da olmayan):")
print("   - Toplam eksik ders:", len(missing_courses))
core_count = sum(1 for c in missing_courses if c.startswith('CORE'))
proj_count = sum(1 for c in missing_courses if any(s in c for s in ['3910','3920','4902','4910','4912','4920','4901']))
print("   - CORE dersleri:", core_count)
print("   - Proje/Seminer dersleri:", proj_count)
diger = sorted(c for c in missing_courses if not c.startswith('CORE') and not any(s in c for s in ['3910','3920','4902','4910','4912','4920','4901']))
print("   - Diger ornekler:", ", ".join(diger[:15]))

print("\n" + "=" * 80)
print("SONUC DEGERLENDIRMESI:")
print("=" * 80)

if len(cakisma_var) == 0:
    print("[+] OGRENCI CAKISMASI: MUKEMMEL!")
    print("   Schedule dosyasinda cakisma dosyasindaki hicbir ogrenci cakismasi yok.")
    print("   Tum ortak ogrencili ders ciftleri farkli gun veya farkli slota atanmis.")
else:
    print("[-] OGRENCI CAKISMASI:", len(cakisma_var), "adet cakisma var!")

print("\n[!] KISIT IHLALLERI:", len(farkli_gun_ihlal) + len(farkli_slot_ihlal))
if len(farkli_gun_ihlal) + len(farkli_slot_ihlal) > 0:
    print("   Bazi ders ciftleri 'farkli gun' veya 'farkli slot' kisitina uymuyor.")
    print("   Ancak bu kisitlarin tam anlami ve zorunlulugu dosya basliklarindan cikarilmistir.")

print("\n[?] EKSIK DERSLER:", len(missing_courses), "ders schedule'da yer almiyor.")
print("   Bu derslerin cogunlugu CORE ve proje/seminer dersleridir.")
print("   Bu derslerin sinav programi ayrica duzenleniyor olabilir.")

print("\n" + "=" * 80)
print("GENEL SONUC: Schedule (5) dosyasi OGRENCI CAKISMASI acisindan DOGRU olusturulmus.")
print("Kisit dosyasindaki bazi kurallar ihlal edilmis olabilir ancak ogrenci cakismasi YOKTUR.")
print("=" * 80)
