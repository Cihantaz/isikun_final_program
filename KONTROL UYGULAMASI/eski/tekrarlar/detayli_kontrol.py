import pandas as pd
from datetime import datetime, time as dt_time

# Load all data
df = pd.read_excel('c:/Users/cihan.tazeoz/Desktop/26 BAHAR FİNAL/UYGULAMA DOSYALARI/2026 Bahar Final Programı _03062026.xlsx')
df['Ders Kodu'] = df['Ders Kodu'].astype(str).str.strip().str.upper()

collisions = pd.read_csv('c:/Users/cihan.tazeoz/Desktop/26 BAHAR FİNAL/UYGULAMA DOSYALARI/final_exam_collisions_14_05_2026_10_48_47.csv')
kisitlar = pd.read_excel('c:/Users/cihan.tazeoz/Desktop/26 BAHAR FİNAL/UYGULAMA DOSYALARI/ders_kisitleri.xlsx')

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
for _, row in df.iterrows():
    code = row['Ders Kodu']
    if code and str(code) != 'NAN':
        program_dict[code] = (int(row['Gun']), int(row['Slot']))

# Collision dict
collision_dict = {}
for _, row in collisions.iterrows():
    c1 = str(row['Course1']).strip().upper()
    c2 = str(row['Course2']).strip().upper()
    cnt = int(row['Common Student Count'])
    collision_dict[(c1, c2)] = cnt
    collision_dict[(c2, c1)] = cnt

# Constraint dict
header = list(kisitlar.columns)
use_cols = header[:4]
first_col = str(header[0]).strip().lower()
if len(header) >= 5 and (first_col.startswith('unnamed') or str(header[0]).strip().isdigit()):
    use_cols = header[1:5]

def get_constraints(code):
    constraints = {}
    for idx, row in kisitlar.iterrows():
        dersler = []
        for col in kisitlar.columns:
            if pd.notna(row[col]):
                val = str(row[col]).strip().upper()
                if val and val not in ['NAN']:
                    dersler.append(val)
        if code in dersler:
            tip = None
            for col in use_cols:
                if pd.notna(row[col]):
                    val = str(row[col]).strip().upper()
                    if val in ['SAMEDAY', 'SAMESLOT', 'DIFFDAY', 'DIFFSLOT']:
                        tip = val
                        break
            if tip:
                for d in dersler:
                    if d != code:
                        constraints[d] = tip
    return constraints

def check_slot_detailed(code, g, s, program_dict):
    cakisan = []
    for d2, (g2, s2) in program_dict.items():
        if d2 == code:
            continue
        if g2 == g and s2 == s:
            cnt = collision_dict.get((code, d2), 0)
            if cnt > 0:
                cakisan.append((d2, cnt))
    
    kisit_sorun = []
    constraints = get_constraints(code)
    for d2, tip in constraints.items():
        g2, s2 = program_dict.get(d2, (None, None))
        if g2 is None:
            continue
        if tip == 'SAMESLOT' and (g2 != g or s2 != s):
            kisit_sorun.append(f"SAMESLOT: {d2} G{g2}S{s2}")
        elif tip == 'SAMEDAY' and g2 != g:
            kisit_sorun.append(f"SAMEDAY: {d2} G{g2}S{s2}")
        elif tip == 'DIFFSLOT' and g2 == g and s2 == s:
            kisit_sorun.append(f"DIFFSLOT: {d2} G{g2}S{s2}")
        elif tip == 'DIFFDAY' and g2 == g:
            kisit_sorun.append(f"DIFFDAY: {d2} G{g2}S{s2}")
    
    return cakisan, kisit_sorun

# ========== BUSI2111 ANALIZI ==========
print("=" * 100)
print("BUSI2111 ALTERNATIF ANALIZI")
print("=" * 100)

g_mevcut, s_mevcut = program_dict.get('BUSI2111', (None, None))
print(f"Mevcut: G{g_mevcut}S{s_mevcut}")
print()

constraints = get_constraints('BUSI2111')
if constraints:
    print("Kisitlari:")
    for d, t in constraints.items():
        loc = program_dict.get(d, 'YOK')
        print(f"  {d}: {t} (mevcut: G{loc})")
    print()

print(f"{'Slot':<15} {'Cakisan':<8} {'Ogr':<6} {'Kisit':<8} {'Durum':<20} {'Detay'}")
print("-" * 100)

uygunlar = []
for g in range(1, 11):
    for s in range(1, 4):
        cakisan, kisit_sorun = check_slot_detailed('BUSI2111', g, s, program_dict)
        toplam = sum(c[1] for c in cakisan)
        
        if not cakisan and not kisit_sorun:
            durum = "UYGUN"
            uygunlar.append((g, s))
        elif cakisan and not kisit_sorun:
            durum = f"CAKISMA ({toplam})"
        elif not cakisan and kisit_sorun:
            durum = f"KISIT ({len(kisit_sorun)})"
        else:
            durum = f"CAKISMA+KISIT"
        
        detay = ""
        if cakisan:
            detay += ", ".join([f"{d}({c})" for d, c in cakisan[:3]])
        if kisit_sorun:
            if detay: detay += " | "
            detay += kisit_sorun[0]
        
        print(f"G{g}S{s:<12} {len(cakisan):<8} {toplam:<6} {len(kisit_sorun):<8} {durum:<20} {detay}")

print()
print("UYGUN ALTERNATIFLER:")
for g, s in uygunlar:
    print(f"  G{g}S{s}")

# G2S2 ozel kontrol
print()
print("=" * 100)
print("BUSI2111 -> G2S2 OZEL KONTROL")
print("=" * 100)
c, k = check_slot_detailed('BUSI2111', 2, 2, program_dict)
print(f"Cakisma: {len(c)} ders, {sum(x[1] for x in c)} ogrenci")
for d2, cnt in c:
    print(f"  {d2}: {cnt} ogrenci")
print(f"Kisit: {len(k)}")
for ks in k:
    print(f"  {ks}")

# ========== TRAD2504 -> G3S1/G3S2 DETAYLI ==========
print()
print("=" * 100)
print("TRAD2504 -> G3S1 / G3S2 DETAYLI KONTROL")
print("=" * 100)

for g, s in [(3, 1), (3, 2)]:
    print(f"\n--- G{g}S{s} ---")
    c, k = check_slot_detailed('TRAD2504', g, s, program_dict)
    print(f"Cakisma: {len(c)} ders, {sum(x[1] for x in c)} ogrenci")
    for d2, cnt in c:
        print(f"  {d2}: {cnt} ogrenci")
    print(f"Kisit: {len(k)}")
    for ks in k:
        print(f"  {ks}")
    
    # Show all courses at this slot
    occup = [(d2, g2, s2) for d2, (g2, s2) in program_dict.items() if g2 == g and s2 == s]
    print(f"Bu slottaki tum dersler ({len(occup)}):")
    for d2, _, _ in sorted(occup):
        cnt = collision_dict.get(('TRAD2504', d2), 0)
        marker = "***" if cnt > 0 else ""
        print(f"  {d2:<15} {marker}")

# ========== BUSI3632 -> G7S2/G8S2 DETAYLI ==========
print()
print("=" * 100)
print("BUSI3632 -> G7S2 / G8S2 DETAYLI KONTROL")
print("=" * 100)

for g, s in [(7, 2), (8, 2)]:
    print(f"\n--- G{g}S{s} ---")
    c, k = check_slot_detailed('BUSI3632', g, s, program_dict)
    print(f"Cakisma: {len(c)} ders, {sum(x[1] for x in c)} ogrenci")
    for d2, cnt in c:
        print(f"  {d2}: {cnt} ogrenci")
    print(f"Kisit: {len(k)}")
    for ks in k:
        print(f"  {ks}")
    
    occup = [(d2, g2, s2) for d2, (g2, s2) in program_dict.items() if g2 == g and s2 == s]
    print(f"Bu slottaki tum dersler ({len(occup)}):")
    for d2, _, _ in sorted(occup):
        cnt = collision_dict.get(('BUSI3632', d2), 0)
        marker = "***" if cnt > 0 else ""
        print(f"  {d2:<15} {marker}")

