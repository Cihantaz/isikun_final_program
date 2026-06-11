import pandas as pd

collision_file = r"C:\Users\cihan.tazeoz\Desktop\26 BAHAR FİNAL\UYGULAMA DOSYALARI\final_exam_collisions_14_05_2026_10_48_47.csv"
schedule_file = r"C:\Users\cihan.tazeoz\Desktop\26 BAHAR FİNAL\UYGULAMA DOSYALARI\schedule (5).xlsx"
kisit_file = r"C:\Users\cihan.tazeoz\Desktop\26 BAHAR FİNAL\UYGULAMA DOSYALARI\ders_kisitleri.xlsx"

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

def parse_ders_list(val):
    if pd.isna(val):
        return []
    return [d.strip().upper() for d in str(val).split(',') if d.strip()]

kisit_clean = kisit.dropna(how='all')

# Farkli slot kisitlarini oku (D kolonu)
farkli_slot_ciftler = []
for _, row in kisit_clean.iterrows():
    val = row.get('D (Farkli Slot)', row.get('D (Farklı Slot)', ''))
    dersler = parse_ders_list(val)
    for i in range(len(dersler)):
        for j in range(i+1, len(dersler)):
            farkli_slot_ciftler.append((dersler[i], dersler[j]))

# Tekrar kaldir
seen = set()
fs_unique = []
for d1, d2 in farkli_slot_ciftler:
    key = tuple(sorted([d1, d2]))
    if key not in seen:
        seen.add(key)
        fs_unique.append((d1, d2))

print("=" * 90)
print("FARKLI SLOT KISITI ANALIZI (Kullanicinin kuralina gore)")
print("=" * 90)
print("\nKural: Farkli slot kisiti sadece iki ders AYNI GUNE atandiginda anlamli.")
print("       Farkli gunde olanlar zaten uyumlu sayilir.\n")

ayni_gun_ayni_slot = []
ayni_gun_farkli_slot = []
farkli_gun = []
bulunamadi = []

for d1, d2 in fs_unique:
    n1 = d1.replace('İ','I').replace('Ş','S').replace('Ğ','G').replace('Ü','U').replace('Ö','O').replace('Ç','C')
    n2 = d2.replace('İ','I').replace('Ş','S').replace('Ğ','G').replace('Ü','U').replace('Ö','O').replace('Ç','C')
    if n1.startswith(EXCLUDED_PREFIXES) or n2.startswith(EXCLUDED_PREFIXES):
        continue
    
    s1 = schedule_map.get(d1)
    s2 = schedule_map.get(d2)
    
    if s1 is None or s2 is None:
        bulunamadi.append((d1, d2, s1, s2))
        continue
    
    if s1[0] == s2[0]:  # Ayni gun
        if s1[1] == s2[1]:  # Ayni slot
            ayni_gun_ayni_slot.append((d1, d2, s1, s2))
        else:
            ayni_gun_farkli_slot.append((d1, d2, s1, s2))
    else:
        farkli_gun.append((d1, d2, s1, s2))

print(f"A. AYNI GUN + AYNI SLOT (GERCEK IHLAL): {len(ayni_gun_ayni_slot)}")
for i, (d1, d2, s1, s2) in enumerate(ayni_gun_ayni_slot, 1):
    print(f"   {i:2d}. {d1:12s} (G{s1[0]}/S{s1[1]}) - {d2:12s} (G{s2[0]}/S{s2[1]})")

print(f"\nB. AYNI GUN + FARKLI SLOT (Kurala uygun): {len(ayni_gun_farkli_slot)}")
for i, (d1, d2, s1, s2) in enumerate(ayni_gun_farkli_slot, 1):
    print(f"   {i:2d}. {d1:12s} (G{s1[0]}/S{s1[1]}) - {d2:12s} (G{s2[0]}/S{s2[1]})")

print(f"\nC. FARKLI GUN (Kurala uygun - kisi otomatik saglanir): {len(farkli_gun)}")
for i, (d1, d2, s1, s2) in enumerate(farkli_gun[:20], 1):
    print(f"   {i:2d}. {d1:12s} (G{s1[0]}/S{s1[1]}) - {d2:12s} (G{s2[0]}/S{s2[1]})")
if len(farkli_gun) > 20:
    print(f"   ... ve {len(farkli_gun) - 20} adet daha")

print(f"\nD. Schedule'da bulunamayan: {len(bulunamadi)}")
for d1, d2, s1, s2 in bulunamadi[:10]:
    print(f"   {d1} ({s1}) - {d2} ({s2})")

# Simdi power2 copy.py'deki kisit implementasyonunu analiz et
print("\n" + "=" * 90)
print("POWER2 COPY.PY KISIT IMPLEMENTASYON ANALIZI")
print("=" * 90)
print("""
1. load_group_constraints_excel fonksiyonu:
   - Excel'den ilk 4/5 kolonu okur.
   - Mapping: SameDay, SameSlot, DifferentDay, DifferentSlot
   - Duzeltilen hata: Unnamed kolonu atlayarak dogru kolonlari okuyor.

2. cpsat_solve fonksiyonunda grup kisitlari (satir 1275-1304):
   
   SameSlot:
     model.Add(t[a] == t[b]).OnlyEnforceIf(ba)
     -> Iki ders AYNI periyoda (gun+slot) atanir.
   
   DifferentSlot:
     model.Add(t[a] != t[b]).OnlyEnforceIf(ba)
     -> Iki ders FARKLI periyoda atanir.
     
     ONEMLI: t[a] != t[b] kisiti zaten otomatik olarak:
     - Farkli gun = farkli periyot = kural otomatik saglanir
     - Ayni gun ayni slot = ayni periyot = ihlal
     - Ayni gun farkli slot = farkli periyot = kural saglanir
     
     Yani power2 copy.py'deki implementasyon kullanicinin dedigi
     kurala UYGUNDUR. t[a] != t[b] sadece ayni gun ayni slot
     durumunu engeller; farkli gun ayni slot'a izin verir.
   
   SameDay:
     model.Add(da == db).OnlyEnforceIf(ba)
     -> Iki dersin gun degeri ayni olmali.
   
   DifferentDay:
     model.Add(da != db).OnlyEnforceIf(ba)
     -> Iki dersin gun degeri farkli olmali.

3. both_assigned fonksiyonu (satir 1269-1273):
   - bb = model.NewBoolVar(...)
   - model.AddBoolAnd([un[a].Not(), un[b].Not()]).OnlyEnforceIf(bb)
   - model.AddBoolOr([un[a], un[b]]).OnlyEnforceIf(bb.Not())
   
   Bu fonksiyon sadece IKI DERS DE atanmissa (unassigned degilse)
   grup kisitlarini uygular. Eger biri atanmamissa kisit
   uygulanmaz. Bu mantikli ve esnek bir yaklasimdir.

SONUC: power2 copy.py'deki DifferentSlot implementasyonu
kullanicinin "farkli slot sadece ayni gunde anlamli" kuralina
UYGUNDUR. t[a] != t[b] kisiti zaten bunu saglar.
""")

print("\n" + "=" * 90)
print("YENI OZET (Kullanicinin kuralina gore)")
print("=" * 90)
print(f"Farkli slot kisiti toplam cift     : {len(fs_unique)}")
print(f"  - Gercek ihlal (ayni gun+slot)   : {len(ayni_gun_ayni_slot)}")
print(f"  - Kurala uygun (ayni gun farkli) : {len(ayni_gun_farkli_slot)}")
print(f"  - Kurala uygun (farkli gun)      : {len(farkli_gun)}")
print(f"  - Bulunamadi                     : {len(bulunamadi)}")
