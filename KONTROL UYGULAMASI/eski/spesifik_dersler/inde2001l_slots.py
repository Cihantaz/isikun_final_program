import pandas as pd
from datetime import datetime, time as dt_time

df = pd.read_excel('c:/Users/cihan.tazeoz/Desktop/26 BAHAR FİNAL/UYGULAMA DOSYALARI/2026 Bahar Final Programı _03062026.xlsx')
df['Ders Kodu'] = df['Ders Kodu'].astype(str).str.strip().str.upper()

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

uygun_gunler = [3, 5, 7, 8]

print('=' * 80)
print('INDE2001-L UYGUN SLOTLAR')
print('=' * 80)
print()
print('Yasak gunler: G1, G2, G4, G6, G9, G10')
print('Uygun gunler: G3, G5, G7, G8')
print()

fmt = "{:<15} {:<10} {:<25} {:<20}"
print(fmt.format("Slot", "Doluluk", "INDE Dersleri", "Oneri"))
print("-" * 80)

for g in uygun_gunler:
    for s in range(1, 4):
        occup = [c for c, (gg, ss) in program_dict.items() if gg == g and ss == s]
        inde_occup = [c for c in occup if c.startswith('INDE')]
        if not inde_occup and len(occup) <= 5:
            oneri = "*** EN IYI ***"
        elif not inde_occup:
            oneri = "UYGUN (bos INDE)"
        else:
            oneri = "UYGUN"
        print(fmt.format(f"G{g}S{s}", len(occup), ", ".join(inde_occup[:3]), oneri))

print()
print("EN IYI ONERILER:")
print("  G3S2 - Sadece 9 ders, 0 INDE")
print("  G3S3 - 11 ders, 0 INDE")
print("  G5S2 - 10 ders, 0 INDE")
print("  G5S3 - 6 ders, 0 INDE")
print("  G7S2 - Sadece 2 ders, 0 INDE")
print("  G7S3 - 4 ders, 0 INDE")
print("  G8S2 - Sadece 1 ders, 0 INDE")

