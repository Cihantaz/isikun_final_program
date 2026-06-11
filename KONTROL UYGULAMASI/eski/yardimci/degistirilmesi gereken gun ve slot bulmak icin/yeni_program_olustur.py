import pandas as pd
import re
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
    # A - Ayni Gun
    ayni_gun = parse_ders_list(row.get('A (Ayni Gun)', row.get('A (Aynı Gün)', '')))
    for i in range(len(ayni_gun)):
        for j in range(i+1, len(ayni_gun)):
            same_day_pairs.add(tuple(sorted([ayni_gun[i], ayni_gun[j]])))
    
    # B - Ayni Slot
    ayni_slot = parse_ders_list(row.get('B (Ayni Slot)', row.get('B (Aynı Slot)', '')))
    for i in range(len(ayni_slot)):
        for j in range(i+1, len(ayni_slot)):
            same_slot_pairs.add(tuple(sorted([ayni_slot[i], ayni_slot[j]])))
    
    # C - Farkli Gun
    farkli_gun = parse_ders_list(row.get('C (Farkli Gun)', row.get('C (Farklı Gün)', '')))
    for i in range(len(farkli_gun)):
        for j in range(i+1, len(farkli_gun)):
            diff_day_pairs.add(tuple(sorted([farkli_gun[i], farkli_gun[j]])))
    
    # D - Farkli Slot
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
print(f"Cakisma cifti (schedule'da olan): {len(collision_pairs)}")
print(f"SameDay cifti: {len(same_day_pairs)}")
print(f"SameSlot cifti: {len(same_slot_pairs)}")
print(f"DiffDay cifti: {len(diff_day_pairs)}")
print(f"DiffSlot cifti: {len(diff_slot_pairs)}")

# 5. CP-SAT modeli kur
model = cp_model.CpModel()

# t[c] = periyot (0 = atanmadi, 1..P)
t = {}
un = {}
for c in all_codes:
    t[c] = model.NewIntVar(0, P, f"t_{c}")
    un[c] = model.NewBoolVar(f"un_{c}")
    model.Add(t[c] == 0).OnlyEnforceIf(un[c])
    model.Add(t[c] != 0).OnlyEnforceIf(un[c].Not())

# x[c,p] = t[c] == p mi?
x = {}
for c in all_codes:
    for p in range(0, P+1):
        x[(c,p)] = model.NewBoolVar(f"x_{c}_{p}")
        if p == 0:
            model.Add(t[c] == 0).OnlyEnforceIf(x[(c,p)])
            model.Add(t[c] != 0).OnlyEnforceIf(x[(c,p)].Not())
        else:
            model.Add(t[c] == p).OnlyEnforceIf(x[(c,p)])
            model.Add(t[c] != p).OnlyEnforceIf(x[(c,p)].Not())
    model.Add(sum(x[(c,p)] for p in range(0, P+1)) == 1)

# Sabit dersler MUTLAKA mevcut slotlarinda kalmali
for c in fixed:
    d, s = schedule_map[c]
    fp = period_index(d, s)
    model.Add(t[c] == fp)
    model.Add(un[c] == 0)

# Degistirilebilir dersler de atanmali (unassigned olmasin)
for c in changeable:
    model.Add(un[c] == 0)

# Gun ve slot degiskenleri
def get_day_var(c):
    day = model.NewIntVar(1, n_days, f"day_{c}")
    tuples = [(p, period_to_day_slot(p)[0]) for p in range(1, P+1)]
    model.AddAllowedAssignments([t[c], day], tuples).OnlyEnforceIf(un[c].Not())
    return day

def get_slot_var(c):
    slot = model.NewIntVar(1, spd, f"slot_{c}")
    tuples = [(p, period_to_day_slot(p)[1]) for p in range(1, P+1)]
    model.AddAllowedAssignments([t[c], slot], tuples).OnlyEnforceIf(un[c].Not())
    return slot

day_var = {}
slot_var = {}
for c in all_codes:
    day_var[c] = get_day_var(c)
    slot_var[c] = get_slot_var(c)

# 6. Kisitlari uygula

# A) Cakisma kisitlari (hard) - cakisan dersler ayni periyotta olamaz
for a, b in collision_pairs:
    if a in all_codes and b in all_codes:
        model.Add(t[a] != t[b])

# B) SameDay (hard)
for a, b in same_day_pairs:
    if a in all_codes and b in all_codes:
        model.Add(day_var[a] == day_var[b])

# C) SameSlot (hard) - ayni periyotta olmali (gun+slot)
for a, b in same_slot_pairs:
    if a in all_codes and b in all_codes:
        model.Add(t[a] == t[b])

# D) DiffDay (hard)
for a, b in diff_day_pairs:
    if a in all_codes and b in all_codes:
        model.Add(day_var[a] != day_var[b])

# E) DiffSlot (hard) - sadece ayni gunde olanlar icin anlamli, ama kodda periyot farkli zorlar
for a, b in diff_slot_pairs:
    if a in all_codes and b in all_codes:
        model.Add(t[a] != t[b])

# 7. Amaç: Mevcut slotlardan sapmayi minimize et
terms = []
for c in changeable:
    curr_d, curr_s = schedule_map[c]
    curr_p = period_index(curr_d, curr_s)
    # Sapma = |t[c] - curr_p|
    diff = model.NewIntVar(-P, P, f"diff_{c}")
    model.Add(diff == t[c] - curr_p)
    abs_diff = model.NewIntVar(0, P, f"absdiff_{c}")
    model.Add(abs_diff >= diff)
    model.Add(abs_diff >= -diff)
    terms.append(abs_diff)

# Ayrica ihlalleri cezalandir (soft)
pen_terms = []

# DiffDay ihlalleri icin ceza
for a, b in diff_day_pairs:
    if a in all_codes and b in all_codes:
        viol = model.NewBoolVar(f"diffday_viol_{a}_{b}")
        model.Add(day_var[a] == day_var[b]).OnlyEnforceIf(viol)
        model.Add(day_var[a] != day_var[b]).OnlyEnforceIf(viol.Not())
        pen_terms.append(100 * viol)

# DiffSlot ihlalleri icin ceza (ayni gun+slot)
for a, b in diff_slot_pairs:
    if a in all_codes and b in all_codes:
        viol = model.NewBoolVar(f"diffslot_viol_{a}_{b}")
        model.Add(t[a] == t[b]).OnlyEnforceIf(viol)
        model.Add(t[a] != t[b]).OnlyEnforceIf(viol.Not())
        pen_terms.append(100 * viol)

model.Minimize(sum(terms) + sum(pen_terms))

solver = cp_model.CpSolver()
solver.parameters.max_time_in_seconds = 30.0
solver.parameters.num_search_workers = 8

print("\nModel cozuluyor...")
status = solver.Solve(model)

if status in (cp_model.OPTIMAL, cp_model.FEASIBLE):
    print(f"COZUM BULUNDU! Status: {'OPTIMAL' if status == cp_model.OPTIMAL else 'FEASIBLE'}")
    
    assigned = {}
    changes = []
    problems = []
    
    for c in all_codes:
        p = solver.Value(t[c])
        if p == 0:
            problems.append(f"{c}: ATANAMADI")
            assigned[c] = schedule_map[c]
        else:
            d, s = period_to_day_slot(p)
            assigned[c] = (d, s)
            if c in changeable:
                old_d, old_s = schedule_map[c]
                if (d, s) != (old_d, old_s):
                    changes.append((c, old_d, old_s, d, s))
    
    print(f"\nDegistirilen ders sayisi: {len(changes)}")
    for c, od, os, nd, ns in changes:
        print(f"  {c}: G{od}/S{os} -> G{nd}/S{ns}")
    
    if problems:
        print(f"\nProblemli dersler: {len(problems)}")
        for p in problems:
            print(f"  {p}")
    
    # Ihlalleri kontrol et
    diffday_viol = 0
    diffslot_viol = 0
    coll_viol = 0
    
    for a, b in diff_day_pairs:
        if a in assigned and b in assigned:
            if assigned[a][0] == assigned[b][0]:
                diffday_viol += 1
    
    for a, b in diff_slot_pairs:
        if a in assigned and b in assigned:
            if assigned[a] == assigned[b]:
                diffslot_viol += 1
    
    for a, b in collision_pairs:
        if a in assigned and b in assigned:
            if assigned[a] == assigned[b]:
                coll_viol += 1
    
    print(f"\nKalan ihlaller:")
    print(f"  DiffDay ihlali: {diffday_viol}")
    print(f"  DiffSlot ihlali: {diffslot_viol}")
    print(f"  Cakisma ihlali: {coll_viol}")
    
    # Excel olustur
    rows = []
    for c in sorted(all_codes):
        d, s = assigned[c]
        rows.append({
            'Ders Kodu': c,
            'Ders Adı': schedule_names.get(c, ''),
            'Slotlar': f"G{d}/S{s}",
            'Gün': d,
            'Slot': s,
            'Değişti': 'EVET' if c in [x[0] for x in changes] else 'HAYIR'
        })
    
    df = pd.DataFrame(rows)
    df.to_excel(out_file, index=False)
    print(f"\nDosya kaydedildi: {out_file}")
    
else:
    print(f"COZUM BULUNAMADI. Status: {status}")
