import pandas as pd
from collections import defaultdict
from ortools.sat.python import cp_model

collision_file = r"C:\Users\cihan.tazeoz\Desktop\26 BAHAR FİNAL\UYGULAMA DOSYALARI\final_exam_collisions_14_05_2026_10_48_47.csv"
schedule_file = r"C:\Users\cihan.tazeoz\Desktop\26 BAHAR FİNAL\UYGULAMA DOSYALARI\schedule (5).xlsx"
kisit_file = r"C:\Users\cihan.tazeoz\Desktop\26 BAHAR FİNAL\UYGULAMA DOSYALARI\ders_kisitleri.xlsx"
out_file = r"C:\Users\cihan.tazeoz\Desktop\26 BAHAR FİNAL\UYGULAMA DOSYALARI\YENILENMIS_PROGRAM_2026_BAHAR_FINAL.xlsx"

EXCLUDED_PREFIXES = ('ARCH', 'IMIM', 'GITA', 'INAR')
CHANGEABLE_PREFIXES = ('COMP', 'SOFT', 'PSYC', 'PSKO')

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

# 1. Schedule'ı oku
schedule = pd.read_excel(schedule_file)
schedule_map = {}
schedule_names = {}
for _, row in schedule.iterrows():
    code = str(row['Ders Kodu']).strip().upper()
    name = str(row.get('Ders Adı', '')).strip()
    slot_val = parse_slot(row['Slotlar'])
    if slot_val:
        schedule_map[code] = slot_val
        schedule_names[code] = name

n_days = 10
spd = 3
P = n_days * spd

def period_index(day, slot):
    return (day - 1) * spd + slot

def period_to_day_slot(p):
    d = (p - 1) // spd + 1
    s = (p - 1) % spd + 1
    return d, s

# 2. Cakisma ciftlerini oku
collisions = pd.read_csv(collision_file)
collision_pairs = set()
for _, row in collisions.iterrows():
    c1 = str(row['Course1']).strip().upper()
    c2 = str(row['Course2']).strip().upper()
    n1 = c1.replace('İ','I').replace('Ş','S').replace('Ğ','G').replace('Ü','U').replace('Ö','O').replace('Ç','C')
    n2 = c2.replace('İ','I').replace('Ş','S').replace('Ğ','G').replace('Ü','U').replace('Ö','O').replace('Ç','C')
    if not n1.startswith(EXCLUDED_PREFIXES) and not n2.startswith(EXCLUDED_PREFIXES):
        if c1 in schedule_map and c2 in schedule_map:
            key = tuple(sorted([c1, c2]))
            collision_pairs.add(key)

# 3. Kisit dosyasini oku
kisit = pd.read_excel(kisit_file)
kisit_clean = kisit.dropna(how='all')

def parse_ders_list(val):
    if pd.isna(val):
        return []
    return [d.strip().upper() for d in str(val).split(',') if d.strip()]

same_day_pairs = set()
same_slot_pairs = set()
diff_day_pairs = set()
diff_slot_pairs = set()

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

# 4. Dersleri kategorize et
all_codes = sorted(schedule_map.keys())
changeable = [c for c in all_codes if c.startswith(CHANGEABLE_PREFIXES)]
fixed = [c for c in all_codes if c not in changeable]

print(f"Toplam ders: {len(all_codes)}")
print(f"Degistirilebilir (COMP/SOFT/PSYC/PSKO): {len(changeable)}")
print(f"Sabit (digerleri): {len(fixed)}")

# 5. CP-SAT modeli kur - SOFT CONSTRAINT yaklasimi
model = cp_model.CpModel()

t = {}
for c in all_codes:
    t[c] = model.NewIntVar(1, P, f"t_{c}")

def get_day_var(c):
    day = model.NewIntVar(1, n_days, f"day_{c}")
    tuples = [(p, period_to_day_slot(p)[0]) for p in range(1, P+1)]
    model.AddAllowedAssignments([t[c], day], tuples)
    return day

def get_slot_var(c):
    slot = model.NewIntVar(1, spd, f"slot_{c}")
    tuples = [(p, period_to_day_slot(p)[1]) for p in range(1, P+1)]
    model.AddAllowedAssignments([t[c], slot], tuples)
    return slot

day_var = {}
slot_var = {}
for c in all_codes:
    day_var[c] = get_day_var(c)
    slot_var[c] = get_slot_var(c)

# Sabit dersler mevcut slotlarinda kalir (hard)
for c in fixed:
    d, s = schedule_map[c]
    fp = period_index(d, s)
    model.Add(t[c] == fp)

# Degistirilebilir dersler icin domain tanimla
# Mevcut slotundan cok uzaklara gitmesin (opsiyonel)

# Cakisma kisitlari (SOFT - cezali)
terms = []
for a, b in collision_pairs:
    if a in all_codes and b in all_codes:
        viol = model.NewBoolVar(f"coll_{a}_{b}")
        model.Add(t[a] == t[b]).OnlyEnforceIf(viol)
        model.Add(t[a] != t[b]).OnlyEnforceIf(viol.Not())
        terms.append(10000 * viol)

# SameDay (SOFT)
for a, b in same_day_pairs:
    if a in all_codes and b in all_codes:
        viol = model.NewBoolVar(f"sd_{a}_{b}")
        model.Add(day_var[a] != day_var[b]).OnlyEnforceIf(viol)
        model.Add(day_var[a] == day_var[b]).OnlyEnforceIf(viol.Not())
        terms.append(500 * viol)

# SameSlot (SOFT) - ayni periyot
for a, b in same_slot_pairs:
    if a in all_codes and b in all_codes:
        viol = model.NewBoolVar(f"ss_{a}_{b}")
        model.Add(t[a] != t[b]).OnlyEnforceIf(viol)
        model.Add(t[a] == t[b]).OnlyEnforceIf(viol.Not())
        terms.append(500 * viol)

# DiffDay (SOFT)
for a, b in diff_day_pairs:
    if a in all_codes and b in all_codes:
        viol = model.NewBoolVar(f"dd_{a}_{b}")
        model.Add(day_var[a] == day_var[b]).OnlyEnforceIf(viol)
        model.Add(day_var[a] != day_var[b]).OnlyEnforceIf(viol.Not())
        terms.append(500 * viol)

# DiffSlot (SOFT) - periyot farkli
for a, b in diff_slot_pairs:
    if a in all_codes and b in all_codes:
        viol = model.NewBoolVar(f"ds_{a}_{b}")
        model.Add(t[a] == t[b]).OnlyEnforceIf(viol)
        model.Add(t[a] != t[b]).OnlyEnforceIf(viol.Not())
        terms.append(500 * viol)

# Mevcut slotlardan sapmayi minimize et (degistirilebilir dersler icin)
for c in changeable:
    curr_d, curr_s = schedule_map[c]
    curr_p = period_index(curr_d, curr_s)
    diff = model.NewIntVar(-P, P, f"diff_{c}")
    model.Add(diff == t[c] - curr_p)
    abs_diff = model.NewIntVar(0, P, f"abs_{c}")
    model.Add(abs_diff >= diff)
    model.Add(abs_diff >= -diff)
    terms.append(abs_diff)

model.Minimize(sum(terms))

solver = cp_model.CpSolver()
solver.parameters.max_time_in_seconds = 60.0
solver.parameters.num_search_workers = 8

print("\nModel cozuluyor (soft constraint yaklasimi)...")
status = solver.Solve(model)

print(f"Status: {solver.StatusName(status)}")

if status in (cp_model.OPTIMAL, cp_model.FEASIBLE):
    assigned = {}
    changes = []
    
    for c in all_codes:
        p = solver.Value(t[c])
        d, s = period_to_day_slot(p)
        assigned[c] = (d, s)
        if c in changeable:
            old_d, old_s = schedule_map[c]
            if (d, s) != (old_d, old_s):
                changes.append((c, old_d, old_s, d, s))
    
    print(f"\nDegistirilen ders sayisi: {len(changes)}")
    for c, od, os, nd, ns in changes:
        print(f"  {c}: G{od}/S{os} -> G{nd}/S{ns}")
    
    # Ihlalleri say
    coll_v = 0
    sd_v = 0
    ss_v = 0
    dd_v = 0
    ds_v = 0
    
    for a, b in collision_pairs:
        if assigned.get(a) == assigned.get(b):
            coll_v += 1
    for a, b in same_day_pairs:
        if assigned.get(a, (0,0))[0] != assigned.get(b, (0,0))[0]:
            sd_v += 1
    for a, b in same_slot_pairs:
        if assigned.get(a) != assigned.get(b):
            ss_v += 1
    for a, b in diff_day_pairs:
        if assigned.get(a, (0,0))[0] == assigned.get(b, (0,0))[0]:
            dd_v += 1
    for a, b in diff_slot_pairs:
        if assigned.get(a) == assigned.get(b):
            ds_v += 1
    
    print(f"\nKalan ihlaller:")
    print(f"  Cakisma: {coll_v}")
    print(f"  SameDay: {sd_v}")
    print(f"  SameSlot: {ss_v}")
    print(f"  DiffDay: {dd_v}")
    print(f"  DiffSlot: {ds_v}")
    
    # Excel olustur
    rows = []
    for c in sorted(all_codes):
        d, s = assigned[c]
        old_d, old_s = schedule_map[c]
        rows.append({
            'Ders Kodu': c,
            'Ders Adı': schedule_names.get(c, ''),
            'Yeni Slot': f"G{d}/S{s}",
            'Eski Slot': f"G{old_d}/S{old_s}",
            'Gün': d,
            'Slot': s,
            'Değişti': 'EVET' if (d,s) != (old_d, old_s) else 'HAYIR'
        })
    
    df = pd.DataFrame(rows)
    df.to_excel(out_file, index=False)
    print(f"\nDosya kaydedildi: {out_file}")
else:
    print("COZUM BULUNAMADI.")
