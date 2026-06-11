# Teknik Açıklama — Işık Üniversitesi Final Takvim Sistemi

Bu doküman, projedeki bileşenleri, veri akışını ve algoritmaları hızlıca anlamak için hazırlanmıştır. Hem AI agent’lar hem de insan geliştiriciler için referans niteliğindedir.

---

## 1. Proje Yapısı

```
final takvim new/
├── isikun_final_program.py          # Ana Flask uygulaması + CP-SAT solver
├── templates/
│   └── index.html                    # Web arayüzü (tek sayfa)
├── KONTROL UYGULAMASI/
│   ├── ana_kontrol.py                # Komut satırı analiz / kontrol aracı
│   ├── ANALIZ_RAPORU.txt             # Eski örnek rapor (salt okunur)
│   ├── NASIL_ANALIZ_YAPILDI.txt      # Eski adım adım notlar (salt okunur)
│   ├── ana_program_rapor.txt         # Eski rapor (salt okunur)
│   └── eski/                         # Arşivlenmiş eski script’ler
│       ├── spesifik_dersler/         # Ders bazlı eski analiz script’leri
│       ├── tekrarlar/                # Eski genel analiz script’leri
│       └── yardimci/                 # Eski yardımcı script’ler
└── .gitignore
```

**Kritik nokta:**
- `isikun_final_program.py` ve `KONTROL UYGULAMASI/ana_kontrol.py` birbirini çağırmaz; iki bağımsız araçtır.
- **Flask Uygulaması** → program *oluşturur* ve *optimize eder*.
- **Kontrol Aracı** → mevcut Excel programını *doğrular* ve *raporlar*.
- **Ancak** Flask uygulamasının kendi içinde de `run_kontrol_analizi`, `diagnose_all_issues` ve `analyze_assignment_issues` gibi kontrol/değerlendirme fonksiyonları vardır. Bu fonksiyonlar web arayüzündeki **Kontrol** ve **Değiştir** sekmelerini besler. `ana_kontrol.py` ise komut satırından çalıştırılan bağımsız bir analiz aracıdır.

---

## 2. `isikun_final_program.py` — Ana Uygulama

### 2.1 Teknoloji Yığını
- **Flask** — Web sunucusu & routing
- **openpyxl** — Excel okuma/yazma
- **pandas** — Veri çerçevesi işlemleri (isteğe bağlı ama önerilir)
- **OR-Tools (cp_model)** — CP-SAT optimizasyon motoru

### 2.2 State Yönetimi (`appstate`)
Tüm veriler global bir `appstate` instance'ında tutulur. `AppState` `@dataclass` olarak tanımlanmıştır ve tüm alanlar `field(default_factory=...)` ile başlatılır. Bu, class-level mutable defaults sorununu çözerek thread-safety'yi iyileştirir (session yok, tek kullanıcılı çalışma varsayılır):

| Alan | Tip | Açıklama |
|------|-----|----------|
| `calendar` | `dict` | Gün sayısı, slot sayısı, saat aralıkları, başlangıç/bitiş tarihleri |
| `courses` | `dict[str, dict]` | Ders kodu → ad, süre, öğrenci sayısı, tercih/engelli günler, sabit gün/slot, eğitmenler |
| `conflicts` | `list[tuple]` | `(dersA, dersB, weight)` öğrenci çakışma listesi |
| `group_constraints` | `list[dict]` | `(type, [dersler])` grup kısıtları |
| `instructors` | `dict` | Eğitmen ID → isim, e-posta |
| `instructor_unavailability` | `list[dict]` | Eğitmen ID → uygun olmadığı G/S |
| `slot_caps` | `dict` | `(gün, slot) → {min, max}` kapasite limitleri |
| `seeds` | `list[dict]` | Elle atanan dersler + lock bayrağı |
| `preview` / `preview_unassigned` | `dict` / `list` | Greedy sonucu atamalar |
| `final` / `final_unassigned` | `dict` / `list` | CP-SAT sonucu atamalar |
| `kontrol_results` | `dict` | Son kontrol çalıştırmasının özet sonuçları |
| `degistir_onerileri` | `list` | Otomatik değişiklik önerileri |

### 2.3 Veri Akışı

```
[Excel/CSV Yükleme]
        │
        ▼
[Feasibility Doctor]  ← 15+ çelişki / uyarı senaryosu tespiti
        │
        ▼
[Greedy Preview]      ← Hızlı, kural tabanlı heuristic atama
        │
        ▼
[CP-SAT Solver]       ← OR-Tools ile optimal atama (çakışma + kısıt + kapasite)
        │
        ▼
[Atama Analizi]       ← Yerleşim sonrası ihlal / kapasite raporu
        │
        ▼
[HTML Çıktı]          ← index.html üzerinden tablo / heatmap / ICS indirme
```

### 2.4 Solver Algoritmaları

#### A. Greedy Preview (`greedy_preview`)
- Dersleri öncelik sırasına koyar: **sabit atamalar önce**, sonra büyük öğrenci sayılı dersler.
- Her ders için aday slotları üretir (tercih günleri önce, engelli günler hariç).
- Çakışma + grup kısıtı + eğitmen uygunluğu + kapasite kontrolü yapar.
- İlk uyan slota yerleştirir.
- Yerleşemezse `preview_unassigned` listesine nedeniyle birlikte yazar.

#### B. CP-SAT Solver (`cpsat_solve`)
- Her ders için tüm `(gün, slot)` kombinasyonlarına bir `BoolVar` atar.
- **Hard constraints:**
  - Her ders en fazla bir slota atanır.
  - Sabit gün/slot zorunlu.
  - Engelli gün yasak.
  - Eğitmen uygun olmama kayıtları yasak.
  - Kapasite min/max sınırları.
- **Soft constraints (objective):**
  - Öğrenci çakışmalarını minimize et (ağırlıklı).
  - Grup kısıtlarını cezalandır (ihlal başına büyük penalty).
  - Seed atamalarına sadakat (lock’lu seed’ler zorunlu, diğerleri tercih).
- **Feasibility garantisi:** Atanamayan dersler `unassigned` olarak işaretlenir; model her zaman çözülebilir.

### 2.5 Kısıt Türleri

| Tür | Anlamı | CP-SAT Implementasyonu |
|-----|--------|------------------------|
| `SameDay` | Aynı gün | `da == db` |
| `SameSlot` | Aynı periyot (gün+slot) | `t[a] == t[b]` |
| `DifferentDay` | Farklı gün | `da != db` |
| `DifferentSlot` | Farklı periyot | `t[a] != t[b]` |

`DifferentSlot` önemli not: OR-Tools’ta `t[a] != t[b]` farklı periyot demektir; farklı gün otomatik olarak farklı periyottur, dolayısıyla farklı gündeki aynı saatli slotlar **ihlal sayılmaz**. Bu, kullanıcının istediği "farklı slot sadece aynı günde anlamlı" kuralına uygundur.

### 2.6 Feasibility Doctor (`diagnose_all_issues`)
Yerleşim öncesi / sonrası 7 kategoride 15+ senaryo tespit eder. Monolitik yapıdan çıkarılarak 7 alt fonksiyona bölünmüştür:

| Alt Fonksiyon | Sorumluluk |
|---------------|------------|
| `_check_calendar_issues` | Slot süreleri gün içine sığmıyor mu? |
| `_check_course_issues` | multi_slots taşma, sabit vs engelli gün çelişkisi, kayıtsız eğitmen ID |
| `_check_instructor_issues` | Kayıtsız eğitmen ID'leri |
| `_check_seed_issues` | Sabit çelişki, engelli gün, takvim dışı |
| `_check_capacity_issues` | Toplam talep > arz, min kapasite > öğrenci sayısı |
| `_check_conflict_data_issues` | Kayıp ders kodu, tekrar |
| `_check_group_conflicts` | SameDay + DifferentDay aynı anda, sabit slot vs grup kısıtı |

Ayrıca `get_conflict_pairs()` ve `get_group_pairs()` helper'ları ile çakışma/grup verileri her fonksiyonda yeniden oluşturulmaz.

---

## 3. `KONTROL UYGULAMASI/ana_kontrol.py` — Kontrol Aracı

### 3.1 Amaç
Halihazırda oluşturulmuş bir final programını (Excel) ve çakışma/kısıt dosyalarını alıp:
- Gerçek öğrenci çakışmalarını sayar,
- Kısıt ihlallerini listeler,
- Fakülte bazlı istatistik üretir,
- Tek bir ders için alternatif slotlar bulur,
- "Şu dersleri şu slotlara taşısak ne olur?" simülasyonu yapar.

### 3.2 Veri Modeli (`ProgramData`)
| Alan | Kaynak | Açıklama |
|------|--------|----------|
| `schedule_map` | Program Excel | `ders_kodu → (gün, slot)` |
| `fakulte_map` | Program Excel | `ders_kodu → fakülte adı` |
| `collision_dict` | Çakışma CSV | `sorted(c1,c2) → ortak öğrenci sayısı` |
| `kisit_groups` / `kisit_pairs` | Kısıt Excel | `SAMEDAY, SAMESLOT, DIFFDAY, DIFFSLOT` çiftleri |

### 3.3 Analiz Fonksiyonları
- `analiz_cakisma` → Aynı G+S’te olan çakışma çiftlerini bulur.
- `analiz_kisit_ihlal` → Tüm 4 kısıt türü için ihlalleri raporlar.
- `analiz_fakulte` → Fakülte bazlı toplam ders / çakışma / ihlal sayıları.
- `check_slot_detailed` → Tek bir `(ders, gün, slot)` kombinasyonu için çakışma + kısıt kontrolü.
- `simulate_changes` → Hipotetik taşıma sonuçlarını raporlar.

### 3.4 Bağımsızlık Notu
Bu araç **sadece okur**, programı değiştirmez. Solver ile entegrasyonu yoktur. Flask uygulamasından ayrı çalıştırılır.

---

## 4. Ortak Veri Formatları

### 4.1 Çakışma CSV (`final_exam_collisions_*.csv`)
```csv
Course1,Course2,Common Student Count
MATH2101,PHYS1001,42
```

### 4.2 Kısıt Excel (`ders_kisitleri.xlsx`)
İlk satır başlık, sonraki satırlar gruplar:
| A (Aynı Gün) | B (Aynı Slot) | C (Farklı Gün) | D (Farklı Slot) |
|--------------|---------------|----------------|-----------------|
| MATH2101,MATH2102 | | PHYS1001,PHYS1002 | |

Not: İlk kolon `Unnamed` veya rakam olabilir; parser otomatik atlar.

### 4.3 Program Excel — 2 format desteklenir

**Format A (Tarih + Saat)** — `ana_kontrol.py` default, `isikun_final_program.py` ders listesi için:
| Ders Kodu | Ders Adı | Fakülte Adı | Sınav Tarihi | Sınav Başlangıç Saati |
|-----------|----------|-------------|--------------|----------------------|
| MATH2101 | Calculus | Mühendislik | 2026-06-08 | 08:30 |

**Format B (Slot String)** — `ana_kontrol.py` alternatif:
| Ders Kodu | Slotlar |
|-----------|---------|
| MATH2101 | G1/S1 |

### 4.4 Ders Listesi Excel (Flask uygulaması için)
| ders kodu | ders adı | süre (dk) | öğrenci sayısı | tercih günleri | engelli günler | sabit gün | sabit slot | multi slots | 3 saatlik | eğitmen id’leri |

---

## 5. Geliştirici Notları

- **Türkçe karakter normalizasyonu:** `İ→I`, `Ş→S`, `Ğ→G`, `Ü→U`, `Ö→O`, `Ç→C` yapılır; `ARCH`, `IMIM`, `GITA`, `INAR` prefix’li dersler analiz dışı bırakılır.
- **Type safety:** `isikun_final_program.py`’de `to_int()`, `to_bool()`, `norm_code()`, `parse_days()` gibi yardımcılar sayesinde Excel hücre tipi (str/int/float/None) fark etmeksizin çalışılır.
- **Thread safety:** `AppState` artık `@dataclass` + `field(default_factory=...)` ile tanımlanır. Global `appstate = AppState()` instance'ı kullanılır. `reset_all_state()` yeni bir `AppState()` oluşturur. Çok kullanıcılı deployment için `flask.g` tabanlı refactor önerilir.
- **Magic number kaldırıldı:** `MAX_CAP = 10 ** 9` sabiti tanımlanmıştır; tüm kapasite sınırları bu sabit üzerinden yönetilir.
- **Code quality:** `base_code()` lowercase replace bug düzeltildi, `to_int()` sadeleştirildi, `COURSE_COLS` duplicate key (smart quote) temizlendi.
- **Dead code kaldırıldı:** `diagnose_all_issues` içindeki `a==b` ölü kod kontrolü kaldırıldı (`load_conflicts_csv` zaten eliyor). `analyze_assignment_issues` içindeki kullanılmayan `n_days`/`spd` değişkenleri kaldırıldı.
- **Performance:** `slot_caps_obj()` `_slot_caps_obj_cache` ile cache'lendi; `slot_caps` değiştiğinde otomatik invalidate edilir.

---

*Son güncelleme: 2026-06-11*
