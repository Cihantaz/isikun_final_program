import pandas as pd
from datetime import datetime, time as dt_time

# Load program
df = pd.read_excel('c:/Users/cihan.tazeoz/Desktop/26 BAHAR FİNAL/UYGULAMA DOSYALARI/2026 Bahar Final Programı _03062026.xlsx')
df['Ders Kodu'] = df['Ders Kodu'].astype(str).str.strip().str.upper()

collisions = pd.read_csv('c:/Users/cihan.tazeoz/Desktop/26 BAHAR FİNAL/UYGULAMA DOSYALARI/final_exam_collisions_14_05_2026_10_48_47.csv')

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
fakulte_dict = {}
for _, row in df.iterrows():
    code = row['Ders Kodu']
    if code and str(code) != 'NAN':
        program_dict[code] = (int(row['Gun']), int(row['Slot']))
        fakulte_dict[code] = str(row.get('Fakülte Adı', '')).strip()

# Collision dict
collision_dict = {}
for _, row in collisions.iterrows():
    c1 = str(row['Course1']).strip().upper()
    c2 = str(row['Course2']).strip().upper()
    cnt = int(row['Common Student Count'])
    collision_dict[(c1, c2)] = cnt
    collision_dict[(c2, c1)] = cnt

# MATH/MATE/Sanat/Tasarım/Mimarlık derslerini bul
def is_protected(code):
    if code.startswith('MATH') or code.startswith('MATE'):
        return True
    if code.startswith('ARCH') or code.startswith('INAR') or code.startswith('ENTA'):
        return True
    if code.startswith('GİTA') or code.startswith('GITA'):
        return True
    # Sanat/Tasarım fakültesi kontrolü
    fak = fakulte_dict.get(code, '')
    if 'SANAT' in fak.upper() or 'TASARIM' in fak.upper() or 'MİMARLIK' in fak.upper() or 'MIMARLIK' in fak.upper():
        return True
    return False

protected_courses = [c for c in program_dict if is_protected(c)]
print(f"Korunan ders sayisi (MATH/MATE/Sanat/Tasarim/Mimarlik): {len(protected_courses)}")

# G1/G2/G3'teki korunan dersler
print("\nG1/G2/G3'teki korunan dersler:")
for g in [1, 2, 3]:
    for s in [1, 2, 3]:
        prot = [c for c, (gg, ss) in program_dict.items() if gg == g and ss == s and is_protected(c)]
        if prot:
            print(f"  G{g}S{s}: {', '.join(prot)}")

# BUSI2111 için G1/G2/G3 analizi (korunan dersleri değiştirmeden)
print("\n" + "=" * 100)
print("BUSI2111 - G1/G2/G3 ANALIZI (Korunan dersler sabit)")
print("=" * 100)

def check_slot(code, g, s):
    cakisan = []
    for d2, (g2, s2) in program_dict.items():
        if d2 == code:
            continue
        if g2 == g and s2 == s:
            cnt = collision_dict.get((code, d2), 0)
            if cnt > 0:
                cakisan.append((d2, cnt))
    return cakisan

print(f"{'Slot':<15} {'Toplam Cakisma':<20} {'Detay'}")
print("-" * 100)

for g in [1, 2, 3]:
    for s in [1, 2, 3]:
        c = check_slot('BUSI2111', g, s)
        toplam = sum(x[1] for x in c)
        
        # Korunan ders çakışması var mı?
        protected_collision = [x for x in c if is_protected(x[0])]
        protected_total = sum(x[1] for x in protected_collision)
        
        detay = ", ".join([f"{d}({cnt})" for d, cnt in c[:3]])
        if protected_collision:
            detay += f" | KORUNAN: {', '.join([d for d,_ in protected_collision])}"
        
        print(f"G{g}S{s:<12} {toplam:<20} {detay}")

# En düşük çakışmalı slotlar
print("\n" + "=" * 100)
print("ONERILEN YERLESIM (Minimal degisiklik)")
print("=" * 100)

# Mevcut G5S3'ten G2S3'e taşıma (sadece 1 öğrenci çakışma, korunan ders yok)
print("\n1. SECENEK: BUSI2111 -> G2S3")
c = check_slot('BUSI2111', 2, 3)
print(f"   Cakisma: {len(c)} ders, {sum(x[1] for x in c)} ogrenci")
for d, cnt in c:
    print(f"     {d}: {cnt} ogrenci")
prot = [x for x in c if is_protected(x[0])]
print(f"   Korunan ders cakismasi: {len(prot)} ders, {sum(x[1] for x in prot)} ogrenci")

# G3S3 (7 öğrenci)
print("\n2. SECENEK: BUSI2111 -> G3S3")
c = check_slot('BUSI2111', 3, 3)
print(f"   Cakisma: {len(c)} ders, {sum(x[1] for x in c)} ogrenci")
for d, cnt in c:
    print(f"     {d}: {cnt} ogrenci")
prot = [x for x in c if is_protected(x[0])]
print(f"   Korunan ders cakismasi: {len(prot)} ders, {sum(x[1] for x in prot)} ogrenci")

# Eğer mevcut G5S3'te bir dersi G2S3'e taşıyıp BUSI2111'i G5S3'e koysak?
# Hayır, kullanıcı BUSI2111'i G1/G2/G3'e taşımak istiyor

# Başka bir seçenek: G2S3'teki BUSI4068'i başka yere taşıyıp BUSI2111'i G2S3'e koymak
print("\n3. SECENEK: G2S3'teki BUSI4068'i baska yere tasiyip BUSI2111 -> G2S3")
print("   BUSI4068 mevcut: G2S3")
c_4068 = check_slot('BUSI4068', 2, 3)
print(f"   BUSI4068 G2S3 cakismasi: {sum(x[1] for x in c_4068)} ogrenci")

# BUSI4068 için alternatifler
print("   BUSI4068 alternatifleri:")
for g in range(1, 11):
    for s in range(1, 4):
        if (g, s) == (2, 3):
            continue
        c = check_slot('BUSI4068', g, s)
        toplam = sum(x[1] for x in c)
        if toplam == 0:
            occup = [d for d, (gg, ss) in program_dict.items() if gg == g and ss == s]
            print(f"     G{g}S{s}: UYGUN ({len(occup)} ders)")

