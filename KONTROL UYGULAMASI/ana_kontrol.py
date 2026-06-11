#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Final Sınav Programı Ana Kontrol ve Analiz Aracı

Tüm temel analizleri tek dosyada toplayan, modüler ana kontrol scripti.

Kullanım:
    python ana_kontrol.py                    -> Tam analiz raporu
    python ana_kontrol.py --ders KOD         -> Belirli ders için alternatif slot analizi
    python ana_kontrol.py --degisiklik FILE  -> JSON/CSV dosyasındaki değişiklikleri simüle et
"""

import pandas as pd
import sys
import os
import json
import glob
from datetime import datetime, time as dt_time
from collections import defaultdict, Counter

# =============================================================================
# 1. KONFIGURASYON
# =============================================================================

# --- Dosya Yolları ---
# Ana program (tarih/saat formatlı Excel)
ANA_PROGRAM = r"C:\Users\cihan.tazeoz\Desktop\26 BAHAR FİNAL\UYGULAMA DOSYALARI\2026 Bahar Final Programı _03062026.xlsx"
# Veya slot formatlı Excel: r"C:\Users\cihan.tazeoz\Desktop\26 BAHAR FİNAL\UYGULAMA DOSYALARI\schedule (5).xlsx"
SCHEDULE_FILE = r"C:\Users\cihan.tazeoz\Desktop\26 BAHAR FİNAL\UYGULAMA DOSYALARI\schedule (5).xlsx"

COLLISION_FILE = r"C:\Users\cihan.tazeoz\Desktop\26 BAHAR FİNAL\UYGULAMA DOSYALARI\final_exam_collisions_14_05_2026_10_48_47.csv"
KISIT_FILE = r"C:\Users\cihan.tazeoz\Desktop\26 BAHAR FİNAL\UYGULAMA DOSYALARI\ders_kisitleri.xlsx"

# --- Sabitler ---
EXCLUDED_PREFIXES = ('ARCH', 'IMIM', 'GITA', 'INAR')
N_DAYS = 10
N_SLOTS = 3

# Tarih -> Gün mapping (hafta içi)
TARIH_GUN = {
    datetime(2026, 6, 8): 1, datetime(2026, 6, 9): 2, datetime(2026, 6, 10): 3,
    datetime(2026, 6, 11): 4, datetime(2026, 6, 12): 5, datetime(2026, 6, 15): 6,
    datetime(2026, 6, 16): 7, datetime(2026, 6, 17): 8, datetime(2026, 6, 18): 9,
    datetime(2026, 6, 19): 10,
}

GUN_TARIH = {
    1: ('08.06.2026', 'Pazartesi'), 2: ('09.06.2026', 'Salı'),
    3: ('10.06.2026', 'Çarşamba'), 4: ('11.06.2026', 'Perşembe'),
    5: ('12.06.2026', 'Cuma'), 6: ('15.06.2026', 'Pazartesi'),
    7: ('16.06.2026', 'Salı'), 8: ('17.06.2026', 'Çarşamba'),
    9: ('18.06.2026', 'Perşembe'), 10: ('19.06.2026', 'Cuma'),
}
SLOT_SAAT = {1: '08:30 - 11:30', 2: '11:30 - 14:30', 3: '14:30 - 17:30'}


# =============================================================================
# 2. YARDIMCI FONKSIYONLAR
# =============================================================================

def normalize_turkce(text):
    """Türkçe karakterleri İngilizce karşılıklarına çevirir (büyük harf)."""
    return text.replace('İ', 'I').replace('Ş', 'S').replace('Ğ', 'G').replace('Ü', 'U').replace('Ö', 'O').replace('Ç', 'C')


def is_excluded(code):
    """Ders kodu hariç tutulan prefix'lerden biriyle mi başlıyor?"""
    norm = normalize_turkce(code)
    return norm.startswith(EXCLUDED_PREFIXES)


def saat_to_slot(saat):
    """Saat bilgisini slot numarasına çevirir."""
    if pd.isna(saat):
        return 0
    if isinstance(saat, dt_time):
        h = saat.hour
    else:
        s = str(saat).strip()
        if s == '-':
            return 0
        try:
            h = int(s.split(':')[0])
        except ValueError:
            return 0
    if h == 8:
        return 1
    elif h == 11:
        return 2
    elif h == 14:
        return 3
    return 0


def tarih_to_gun(tarih):
    """datetime veya string tarihi gün numarasına çevirir."""
    if pd.isna(tarih):
        return 0
    if isinstance(tarih, datetime):
        return TARIH_GUN.get(tarih, 0)
    if isinstance(tarih, str):
        tarih = tarih.strip()
        for dt, g in TARIH_GUN.items():
            if dt.strftime('%Y-%m-%d') == tarih or dt.strftime('%d.%m.%Y') == tarih:
                return g
    return 0


def parse_gs_string(slot_str):
    """'G1/S2' -> (1, 2)"""
    if pd.isna(slot_str):
        return None
    slot_str = str(slot_str).strip().upper()
    parts = slot_str.split('/')
    if len(parts) != 2:
        return None
    try:
        gun = int(parts[0].replace('G', ''))
        slot = int(parts[1].replace('S', ''))
        return (gun, slot)
    except ValueError:
        return None


def gs_yaz(g, s):
    """Gün/slot bilgisini okunabilir tarih/saat formatında döndürür."""
    tarih, gun = GUN_TARIH.get(g, ('?', '?'))
    saat = SLOT_SAAT.get(s, '?')
    return f'{tarih} {gun}  {saat}'


def parse_ders_list(val):
    """Virgülle ayrılmış ders kodu string'ini listeye çevirir."""
    if pd.isna(val):
        return []
    return [d.strip().upper() for d in str(val).split(',') if d.strip()]


# =============================================================================
# 3. VERI YUKLEME
# =============================================================================

class ProgramData:
    """Tüm veri kaynaklarını yükleyen ve işleyen ana sınıf."""

    def __init__(self, program_file=None, collision_file=None, kisit_file=None,
                 schedule_file=None, use_program_dates=True):
        """
        use_program_dates=True  -> program_file'dan tarih/saat oku
        use_program_dates=False -> schedule_file'dan G/S string oku
        """
        self.program_file = program_file or ANA_PROGRAM
        self.schedule_file = schedule_file or SCHEDULE_FILE
        self.collision_file = collision_file or COLLISION_FILE
        self.kisit_file = kisit_file or KISIT_FILE

        self.schedule_map = {}      # ders_kodu -> (gun, slot)
        self.fakulte_map = {}       # ders_kodu -> fakulte
        self.collision_dict = {}    # (c1, c2) -> count  (sorted key)
        self.kisit_groups = []      # [(tip, [dersler]), ...]
        self.kisit_pairs = {        # tip -> set of sorted tuples
            'SAMEDAY': set(), 'SAMESLOT': set(),
            'DIFFDAY': set(), 'DIFFSLOT': set()
        }

        self._load_program(use_program_dates)
        self._load_collisions()
        self._load_kisitlar()

    # --- Program ---
    def _load_program(self, use_dates):
        if use_dates and os.path.exists(self.program_file):
            df = pd.read_excel(self.program_file)
            df['Ders Kodu'] = df['Ders Kodu'].astype(str).str.strip().str.upper()
            for _, row in df.iterrows():
                code = row['Ders Kodu']
                if not code or code == 'NAN':
                    continue
                gun = tarih_to_gun(row.get('Sınav Tarihi'))
                slot = saat_to_slot(row.get('Sınav Başlangıç Saati'))
                if gun > 0 and slot > 0:
                    self.schedule_map[code] = (gun, slot)
                    self.fakulte_map[code] = str(row.get('Fakülte Adı', '')).strip()
        elif os.path.exists(self.schedule_file):
            df = pd.read_excel(self.schedule_file)
            for _, row in df.iterrows():
                code = str(row.get('Ders Kodu', '')).strip().upper()
                if not code:
                    continue
                slot_val = parse_gs_string(row.get('Slotlar'))
                if slot_val:
                    self.schedule_map[code] = slot_val
                    self.fakulte_map[code] = str(row.get('Fakülte Adı', '')).strip()
        else:
            raise FileNotFoundError("Ne program dosyası ne de schedule dosyası bulunamadı.")

    # --- Çakışmalar ---
    def _load_collisions(self):
        collisions = pd.read_csv(self.collision_file)
        for _, row in collisions.iterrows():
            c1 = str(row['Course1']).strip().upper()
            c2 = str(row['Course2']).strip().upper()
            count = int(row['Common Student Count'])
            if is_excluded(c1) or is_excluded(c2):
                continue
            key = tuple(sorted([c1, c2]))
            if key not in self.collision_dict or self.collision_dict[key] < count:
                self.collision_dict[key] = count

    # --- Kısıtlar ---
    def _load_kisitlar(self):
        kisit = pd.read_excel(self.kisit_file)
        kisit_clean = kisit.dropna(how='all')

        # Kolon adlarını tespit et
        header = list(kisit.columns)
        use_cols = header[:4]
        first_col = str(header[0]).strip().lower()
        if len(header) >= 5 and (first_col.startswith('unnamed') or str(header[0]).strip().isdigit()):
            use_cols = header[1:5]

        col_map = {
            'A (Aynı Gün)': 'SAMEDAY', 'A (Aynı Gun)': 'SAMEDAY',
            'B (Aynı Slot)': 'SAMESLOT', 'B (Aynı Slot)': 'SAMESLOT',
            'C (Farklı Gün)': 'DIFFDAY', 'C (Farkli Gun)': 'DIFFDAY',
            'D (Farklı Slot)': 'DIFFSLOT', 'D (Farkli Slot)': 'DIFFSLOT',
        }

        for _, row in kisit_clean.iterrows():
            for col_raw, tip in col_map.items():
                val = row.get(col_raw, '')
                dersler = parse_ders_list(val)
                if len(dersler) >= 2:
                    self.kisit_groups.append((tip, dersler))
                    for i in range(len(dersler)):
                        for j in range(i + 1, len(dersler)):
                            self.kisit_pairs[tip].add(tuple(sorted([dersler[i], dersler[j]])))

    # --- Yardımcı erişim ---
    def get_collisions_for(self, code):
        """Belirli dersin çakıştığı diğer dersleri döndürür: [(other, count), ...]"""
        result = []
        for (c1, c2), count in self.collision_dict.items():
            if code == c1:
                result.append((c2, count))
            elif code == c2:
                result.append((c1, count))
        return result

    def get_constraints_for(self, code):
        """Belirli dersin kısıtlarını döndürür: {other_code: tip, ...}"""
        result = {}
        for tip, dersler in self.kisit_groups:
            if code in dersler:
                for d in dersler:
                    if d != code:
                        result[d] = tip
        return result

    def get_courses_at_slot(self, g, s):
        """Belirli slottaki tüm dersleri döndürür."""
        return [c for c, (gg, ss) in self.schedule_map.items() if gg == g and ss == s]


# =============================================================================
# 4. ANALIZ FONKSIYONLARI
# =============================================================================

def analiz_cakisma(data: ProgramData):
    """Öğrenci çakışma analizi."""
    cakisma_var = []
    cakisma_yok = []
    bulunamadi = []

    for (c1, c2), count in data.collision_dict.items():
        s1 = data.schedule_map.get(c1)
        s2 = data.schedule_map.get(c2)
        if s1 is None or s2 is None:
            bulunamadi.append((c1, c2, count, s1, s2))
            continue
        if s1 == s2:
            cakisma_var.append((c1, c2, count, s1[0], s1[1]))
        else:
            cakisma_yok.append((c1, c2, count, s1[0], s1[1], s2[0], s2[1]))

    return {
        'var': cakisma_var,
        'yok': cakisma_yok,
        'bulunamadi': bulunamadi
    }


def analiz_kisit_ihlal(data: ProgramData):
    """Tüm kısıt ihlallerini bulur."""
    ihlaller = {
        'SAMEDAY': [], 'SAMESLOT': [],
        'DIFFDAY': [], 'DIFFSLOT': []
    }

    for tip, pairs in data.kisit_pairs.items():
        for d1, d2 in pairs:
            if is_excluded(d1) or is_excluded(d2):
                continue
            s1 = data.schedule_map.get(d1)
            s2 = data.schedule_map.get(d2)
            if not s1 or not s2:
                continue

            if tip == 'SAMEDAY' and s1[0] != s2[0]:
                ihlaller[tip].append((d1, d2, s1, s2))
            elif tip == 'SAMESLOT' and s1 != s2:
                ihlaller[tip].append((d1, d2, s1, s2))
            elif tip == 'DIFFDAY' and s1[0] == s2[0]:
                ihlaller[tip].append((d1, d2, s1, s2))
            elif tip == 'DIFFSLOT' and s1 == s2:
                ihlaller[tip].append((d1, d2, s1, s2))

    return ihlaller


def analiz_fakulte(data: ProgramData, cakisma_var):
    """Fakülte bazlı istatistikler."""
    stats = defaultdict(lambda: {'toplam': 0, 'cakisma': 0, 'dd': 0, 'ds': 0})
    for code, fakulte in data.fakulte_map.items():
        f = fakulte if fakulte else 'Bilinmiyor'
        stats[f]['toplam'] += 1

    for c1, c2, count, g, s in cakisma_var:
        f1 = data.fakulte_map.get(c1, 'Bilinmiyor')
        f2 = data.fakulte_map.get(c2, 'Bilinmiyor')
        stats[f1]['cakisma'] += 1
        if f2 != f1:
            stats[f2]['cakisma'] += 1

    ihlaller = analiz_kisit_ihlal(data)
    for d1, d2, s1, s2 in ihlaller['DIFFDAY']:
        f1 = data.fakulte_map.get(d1, 'Bilinmiyor')
        f2 = data.fakulte_map.get(d2, 'Bilinmiyor')
        stats[f1]['dd'] += 1
        if f2 != f1:
            stats[f2]['dd'] += 1

    for d1, d2, s1, s2 in ihlaller['DIFFSLOT']:
        f1 = data.fakulte_map.get(d1, 'Bilinmiyor')
        f2 = data.fakulte_map.get(d2, 'Bilinmiyor')
        stats[f1]['ds'] += 1
        if f2 != f1:
            stats[f2]['ds'] += 1

    return stats


def check_slot_detailed(data: ProgramData, code, g, s, custom_schedule=None):
    """
    Belirli bir ders için belirli slotun durumunu kontrol eder.
    custom_schedule verilirse onu kullanır (değişiklik simülasyonu için).
    """
    schedule = custom_schedule if custom_schedule else data.schedule_map

    # Çakışma kontrolü
    cakisan = []
    for d2, (g2, s2) in schedule.items():
        if d2 == code:
            continue
        if g2 == g and s2 == s:
            cnt = data.collision_dict.get(tuple(sorted([code, d2])), 0)
            if cnt > 0:
                cakisan.append((d2, cnt))

    # Kısıt kontrolü
    kisit_sorun = []
    constraints = data.get_constraints_for(code)
    for d2, tip in constraints.items():
        g2, s2 = schedule.get(d2, (None, None))
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


def find_alternatives(data: ProgramData, code, day_range=None, protected_checker=None):
    """
    Belirli ders için tüm uygun alternatif slotları bulur.
    protected_checker: (other_code) -> bool  fonksiyonu (örn: MATH/MATE koruma)
    """
    day_range = day_range or range(1, N_DAYS + 1)
    uygunlar = []
    for g in day_range:
        for s in range(1, N_SLOTS + 1):
            cakisan, kisit_sorun = check_slot_detailed(data, code, g, s)
            if protected_checker:
                cakisan = [(d, c) for d, c in cakisan if not protected_checker(d)]
            if not cakisan and not kisit_sorun:
                uygunlar.append((g, s))
    return uygunlar


def simulate_changes(data: ProgramData, changes):
    """
    Değişiklikleri simüle eder ve sonuçları raporlar.
    changes: {ders_kodu: (yeni_g, yeni_s), ...}
    """
    test_schedule = dict(data.schedule_map)
    for d, (g, s) in changes.items():
        test_schedule[d] = (g, s)

    print("\n" + "=" * 100)
    print("DEGISIKLIK SIMULASYONU")
    print("=" * 100)

    # Her değişiklik için ayrı kontrol
    for d, (g_yeni, s_yeni) in changes.items():
        g_eski, s_eski = data.schedule_map.get(d, (None, None))
        print(f"\n{d}: G{g_eski}S{s_eski} -> G{g_yeni}S{s_yeni}")
        cakisan, kisit = check_slot_detailed(data, d, g_yeni, s_yeni, test_schedule)
        if cakisan:
            cakisan.sort(key=lambda x: -x[1])
            toplam = sum(c[1] for c in cakisan)
            print(f"  CAKISMA: {len(cakisan)} ders, {toplam} ogrenci")
            for d2, cnt in cakisan:
                print(f"    {d2:<15} {cnt:>3} ogrenci")
        else:
            print(f"  CAKISMA YOK")
        if kisit:
            print(f"  KISIT IHLALI: {len(kisit)}")
            for k in kisit:
                print(f"    {k}")

    # Değişen derslerin birbiriyle kontrolü
    ders_list = list(changes.keys())
    if len(ders_list) > 1:
        print("\n" + "-" * 100)
        print("DEGISEN DERSLERIN BIRBIRIYLE KONTROLU:")
        for i in range(len(ders_list)):
            for j in range(i + 1, len(ders_list)):
                d1, d2 = ders_list[i], ders_list[j]
                g1, s1 = changes[d1]
                g2, s2 = changes[d2]
                cnt = data.collision_dict.get(tuple(sorted([d1, d2])), 0)
                same_slot = (g1 == g2 and s1 == s2)
                durum = "AYNI SLOT!" if same_slot else "Farkli slot"
                print(f"  {d1} - {d2}: {cnt} ogrenci | {durum}")

    # Kısıt kontrolü (değişen dersler dahil tüm kısıtlar)
    print("\n" + "-" * 100)
    print("TUM KISIT KONTROLU (DEGISIKLIKLER DAHIL):")
    ihlaller = analiz_kisit_ihlal(data)  # Not: bu orijinal schedule'a göre
    # Simüle edilmiş ihlalleri hesapla
    sim_ihlal = {'SAMEDAY': [], 'SAMESLOT': [], 'DIFFDAY': [], 'DIFFSLOT': []}
    for tip, pairs in data.kisit_pairs.items():
        for d1, d2 in pairs:
            if is_excluded(d1) or is_excluded(d2):
                continue
            s1 = test_schedule.get(d1)
            s2 = test_schedule.get(d2)
            if not s1 or not s2:
                continue
            if tip == 'SAMEDAY' and s1[0] != s2[0]:
                sim_ihlal[tip].append((d1, d2, s1, s2))
            elif tip == 'SAMESLOT' and s1 != s2:
                sim_ihlal[tip].append((d1, d2, s1, s2))
            elif tip == 'DIFFDAY' and s1[0] == s2[0]:
                sim_ihlal[tip].append((d1, d2, s1, s2))
            elif tip == 'DIFFSLOT' and s1 == s2:
                sim_ihlal[tip].append((d1, d2, s1, s2))

    total_ihlal = sum(len(v) for v in sim_ihlal.values())
    print(f"Toplam kisit ihlali (simulasyon sonrasi): {total_ihlal}")
    for tip, liste in sim_ihlal.items():
        if liste:
            print(f"  {tip}: {len(liste)}")


# =============================================================================
# 5. RAPORLAMA
# =============================================================================

def rapor_ozet(data: ProgramData):
    """Genel özet raporu üretir."""
    cakisma = analiz_cakisma(data)
    ihlaller = analiz_kisit_ihlal(data)
    fakulte_stats = analiz_fakulte(data, cakisma['var'])

    print("=" * 100)
    print("           FINAL SINAV PROGRAMI - ANA KONTROL RAPORU")
    print("=" * 100)

    print(f"\n[1] PROGRAM VERILERI")
    print(f"    - Toplam ders (schedule): {len(data.schedule_map)}")
    print(f"    - Toplam cakisma cifti: {len(data.collision_dict)}")
    print(f"    - Toplam kisit grubu: {len(data.kisit_groups)}")

    print(f"\n[2] OGRENCI CAKISMA ANALIZI")
    print(f"    - GERCEK CAKISMA (ayni gun+slot): {len(cakisma['var'])}")
    print(f"    - CAKISMA YOK (dogru ayrilmis)  : {len(cakisma['yok'])}")
    print(f"    - Schedule'da bulunamayan cift  : {len(cakisma['bulunamadi'])}")
    if cakisma['var']:
        print(f"\n    Ornekler (ilk 5):")
        for c1, c2, count, g, s in cakisma['var'][:5]:
            print(f"      {c1} - {c2} | {count} ogrenci | G{g}/S{s}")

    print(f"\n[3] KISIT IHLALLERI")
    total_ihlal = sum(len(v) for v in ihlaller.values())
    print(f"    - Toplam ihlal: {total_ihlal}")
    for tip, liste in ihlaller.items():
        print(f"      {tip:<12}: {len(liste)}")
        if liste:
            for d1, d2, s1, s2 in liste[:5]:
                print(f"         {d1} (G{s1[0]}/S{s1[1]}) - {d2} (G{s2[0]}/S{s2[1]})")

    print(f"\n[4] FAKULTE BAZLI ISTATISTIKLER")
    print(f"    {'Fakulte':<40} {'Toplam':<8} {'Cakisma':<8} {'DD':<8} {'DS':<8}")
    print("    " + "-" * 75)
    for fakulte, stats in sorted(fakulte_stats.items()):
        print(f"    {fakulte:<40} {stats['toplam']:<8} {stats['cakisma']:<8} {stats['dd']:<8} {stats['ds']:<8}")

    # Eksik dersler
    missing = set()
    for (c1, c2), count in data.collision_dict.items():
        if c1 not in data.schedule_map:
            missing.add(c1)
        if c2 not in data.schedule_map:
            missing.add(c2)

    core = [c for c in missing if c.startswith('CORE')]
    proje = [c for c in missing if any(s in c for s in ['3910', '3920', '4901', '4902', '4910', '4912', '4920'])]
    diger = [c for c in missing if c not in core and c not in proje]

    print(f"\n[5] EKSIK DERSLER (Schedule'da olmayan)")
    print(f"    - Toplam: {len(missing)}")
    print(f"    - CORE: {len(core)}")
    print(f"    - Proje/Seminer: {len(proje)}")
    print(f"    - Diger: {len(diger)}")
    if diger:
        print(f"      Ornekler: {', '.join(sorted(diger)[:15])}")

    print("\n" + "=" * 100)
    if len(cakisma['var']) == 0 and total_ihlal == 0:
        print("SONUC: Program MUKEMMEL! Hicbir cakisma ve kisit ihlali yok.")
    elif len(cakisma['var']) == 0:
        print("SONUC: Ogrenci cakismasi YOK ancak bazi kisitlar ihlal edilmis.")
    else:
        print(f"SONUC: {len(cakisma['var'])} adet GERCEK CAKISMA var! Duzenleme gerekli.")
    print("=" * 100)


def rapor_ders_analizi(data: ProgramData, code, day_range=None):
    """Belirli bir ders için detaylı alternatif analizi raporlar."""
    code = code.strip().upper()
    mevcut = data.schedule_map.get(code)

    print("=" * 100)
    print(f"{code} DETAYLI ANALIZI")
    print("=" * 100)
    if mevcut:
        print(f"Mevcut slot: G{mevcut[0]}S{mevcut[1]} ({gs_yaz(*mevcut)})")
    else:
        print("Mevcut slot: BULUNAMADI")

    # Çakışan dersler
    cakisan = data.get_collisions_for(code)
    if cakisan:
        print(f"\nCakisan dersler ({len(cakisan)} adet):")
        for other, count in sorted(cakisan, key=lambda x: -x[1])[:15]:
            other_slot = data.schedule_map.get(other)
            print(f"  {other:<15} {count:>4} ogrenci  (mevcut: G{other_slot[0]}S{other_slot[1]})")

    # Kısıtlar
    constraints = data.get_constraints_for(code)
    if constraints:
        print(f"\nKisitlari:")
        for d, t in constraints.items():
            loc = data.schedule_map.get(d, 'YOK')
            print(f"  {d}: {t} (mevcut: {loc})")

    # Alternatif slotlar
    day_range = day_range or range(1, N_DAYS + 1)
    print(f"\n{'Slot':<20} {'Cakisan':<8} {'Ogr':<6} {'Kisit':<8} {'Durum':<20} {'Detay'}")
    print("-" * 100)

    uygunlar = []
    for g in day_range:
        for s in range(1, N_SLOTS + 1):
            cakisan_list, kisit_sorun = check_slot_detailed(data, code, g, s)
            toplam = sum(c[1] for c in cakisan_list)

            if not cakisan_list and not kisit_sorun:
                durum = "UYGUN"
                uygunlar.append((g, s))
            elif cakisan_list and not kisit_sorun:
                durum = f"CAKISMA ({toplam})"
            elif not cakisan_list and kisit_sorun:
                durum = f"KISIT ({len(kisit_sorun)})"
            else:
                durum = f"CAKISMA+KISIT"

            detay = ""
            if cakisan_list:
                detay += ", ".join([f"{d}({c})" for d, c in cakisan_list[:3]])
            if kisit_sorun:
                if detay:
                    detay += " | "
                detay += kisit_sorun[0]

            print(f"G{g}S{s:<15} {len(cakisan_list):<8} {toplam:<6} {len(kisit_sorun):<8} {durum:<20} {detay}")

    print("\nUYGUN ALTERNATIFLER:")
    if uygunlar:
        for g, s in uygunlar:
            print(f"  G{g}S{s}  ({gs_yaz(g, s)})")
    else:
        print("  Belirtilen gun araliginda tamamen uygun slot bulunamadi.")


# =============================================================================
# 6. ANA CALISTIRMA
# =============================================================================

def main():
    import argparse
    parser = argparse.ArgumentParser(description='Final Sınav Programı Ana Kontrol')
    parser.add_argument('--ders', type=str, help='Tek ders analizi (örn: BUSI2111)')
    parser.add_argument('--gun-araligi', type=str, default='1-10', help='Ders analizi için gün aralığı (örn: 1-6)')
    parser.add_argument('--degisiklik', type=str, help='Değişiklik JSON dosyası (örn: {"BUSI2111": [2,3]})')
    parser.add_argument('--program', type=str, help='Ana program Excel dosyası')
    parser.add_argument('--schedule', type=str, help='Schedule Excel dosyası (G/S formatlı)')
    parser.add_argument('--collision', type=str, help='Çakışma CSV dosyası')
    parser.add_argument('--kisit', type=str, help='Kısıt Excel dosyası')
    parser.add_argument('--use-dates', action='store_true', help='Program dosyasından tarih/saat oku')
    args = parser.parse_args()

    # Dosya yollarını override et
    prog = args.program or ANA_PROGRAM
    sched = args.schedule or SCHEDULE_FILE
    coll = args.collision or COLLISION_FILE
    kis = args.kisit or KISIT_FILE
    use_dates = args.use_dates if args.program else False  # Default: schedule dosyasını kullan

    if not os.path.exists(prog) and not os.path.exists(sched):
        print("HATA: Ne program dosyası ne de schedule dosyası bulunamadi.")
        print(f"  Aranan program: {prog}")
        print(f"  Aranan schedule: {sched}")
        sys.exit(1)

    data = ProgramData(
        program_file=prog,
        schedule_file=sched,
        collision_file=coll,
        kisit_file=kis,
        use_program_dates=use_dates
    )

    if args.ders:
        # Gün aralığını parse et
        if '-' in args.gun_araligi:
            start, end = map(int, args.gun_araligi.split('-'))
            day_range = range(start, end + 1)
        else:
            day_range = range(1, N_DAYS + 1)
        rapor_ders_analizi(data, args.ders, day_range)

    elif args.degisiklik:
        with open(args.degisiklik, 'r', encoding='utf-8') as f:
            changes_raw = json.load(f)
        changes = {k: tuple(v) for k, v in changes_raw.items()}
        simulate_changes(data, changes)

    else:
        rapor_ozet(data)


if __name__ == "__main__":
    main()
