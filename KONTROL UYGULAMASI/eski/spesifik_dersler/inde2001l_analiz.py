import pandas as pd
from datetime import datetime, time as dt_time

# Load program
df_full = pd.read_excel('c:/Users/cihan.tazeoz/Desktop/26 BAHAR FİNAL/UYGULAMA DOSYALARI/2026 Bahar Final Programı _03062026.xlsx')
df_full['Ders Kodu'] = df_full['Ders Kodu'].astype(str).str.strip().str.upper()

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

df_full['Gun'] = df_full['Sınav Tarihi'].apply(lambda x: tarih_gun.get(x, 0) if pd.notna(x) else 0)
df_full['Slot'] = df_full['Sınav Başlangıç Saati'].apply(saat_to_slot)

program_dict = {}
for _, row in df_full.iterrows():
    code = row['Ders Kodu']
    if code and str(code) != 'NAN':
        program_dict[code] = (int(row['Gun']), int(row['Slot']))

# INDE2001-L hangi gunlerde olamaz (DiffDay kisitlari)
# Grup 1: INDE2001(G5), MATH2201(G4), MATH2105(G2)
# Grup 2: INDE2001(G5), INDE2156(G4), INDE2211(G6), INDE2452(G7), MATH2107(G6), MATH2103(G1)
# Grup 4: INDE2001(G5), MATH2201(G4), MATH2107(G6), MATH2105(G2)

yasak_gunler = {1, 2, 4, 5, 6, 7}  # MATH2103, MATH2105, MATH2201/INDE2156, INDE2001, MATH2107/INDE2211, INDE2452

print("=" * 80)
print("INDE2001-L ANALIZI")
print("=" * 80)
print(f"\nINDE2001 mevcut: G{program_dict.get('INDE2001', 'YOK')}")
print(f"Yasak gunler (DiffDay): {sorted(yasak_gunler)}")
print(f"Uygun gunler: {[g for g in range(1,11) if g not in yasak_gunler]}")

# Check each slot for occupancy and collisions (INDE2001-L has no collision data)
print("\n" + "=" * 80)
print("TUM SLOTLARDA DURUM (INDE2001-L icin)")
print("=" * 80)
print(f"{'Slot':<20} {'Doluluk':<10} {'INDE Dersleri':<30} {'Durum'}")
print("-" * 80)

for g in range(1, 11):
    for s in range(1, 4):
        occup = [c for c, (gg, ss) in program_dict.items() if gg == g and ss == s]
        inde_occup = [c for c in occup if c.startswith('INDE')]
        if g in yasak_gunler:
            durum = "YASAK (DiffDay)"
        else:
            durum = "UYGUN" if not inde_occup else f"UYGUN ({len(inde_occup)} INDE)"
        print(f"G{g}S{s:<15} {len(occup):<10} {', '.join(inde_occup[:3]):<30} {durum}")

# Find slots with no INDE courses in allowed days
print("\n" + "=" * 80)
print("ONERILEN SLOTLAR (INDE2001-L icin)")
print("=" * 80)

for g in range(1, 11):
    if g in yasak_gunler:
        continue
    for s in range(1, 4):
        occup = [c for c, (gg, ss) in program_dict.items() if gg == g and ss == s]
        inde_occup = [c for c in occup if c.startswith('INDE')]
        if not inde_occup:
            print(f"  G{g}S{s} - Bos slot (0 INDE, {len(occup)} diger ders)")
        else:
            print(f"  G{g}S{s} - {len(inde_occup)} INDE ders var: {', '.join(inde_occup)}")

