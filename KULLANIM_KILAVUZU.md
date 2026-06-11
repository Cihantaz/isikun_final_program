# Kullanım Kılavuzu — Işık Üniversitesi Final Takvim Sistemi

Bu kılavuz, sistemin nasıl çalıştırılacağını, hangi dosyanın ne işe yaradığını ve beklenen veri formatlarını açıklar.

---

## 1. Sistem Bileşenleri

| Dosya / Klasör | Görevi | Kim kullanır? |
|----------------|--------|---------------|
| `isikun_final_program.py` | Web tabanlı program oluşturma & optimize etme | Final programını sıfırdan oluşturan kullanıcı |
| `KONTROL UYGULAMASI/ana_kontrol.py` | Mevcut programı analiz etme, çakışma/kısıt raporu çıkarma | Programı kontrol eden / düzeltme önerisi arayan kullanıcı |
| `templates/index.html` | Web arayüzü | Tarayıcı üzerinden görüntülenir |
| `KONTROL UYGULAMASI/eski/` | Arşivlenmiş eski script’ler | Gerektiğinde manuel bakılır |

---

## 2. `isikun_final_program.py` — Program Oluşturucu

### 2.1 Başlatma
```bash
pip install flask openpyxl pandas ortools
python isikun_final_program.py
```
Tarayıcıda: `http://127.0.0.1:5000`

### 2.2 Adım Adım Kullanım
1. **Takvim Ayarları** → Gün sayısı, slot sayısı, saatler, başlangıç/bitiş tarihlerini girin.
2. **Ders Listesi Yükle** → Excel dosyasını seçin (`*.xlsx`).
3. **Çakışmalar Yükle** → CSV dosyasını seçin (`*.csv`).
4. **Grup Kısıtları Yükle** → Excel dosyasını seçin (`*.xlsx`).
5. *(İsteğe bağlı)* Eğitmenler, eğitmen uygunlukları, slot kapasiteleri, elle atama (seed) ekleyin.
6. **Greedy Preview** → Hızlı önizleme alın.
7. **CP-SAT Solve** → Optimal final programını oluşturun.
8. **Kontrol Et** → Oluşan programdaki sorunları ve otomatik değişiklik önerilerini görün.
9. **Excel / ICS İndir** → Sonucu dışa aktarın.

### 2.3 Beklenen Yükleme Dosya Formatları

#### A. Ders Listesi Excel
**Sayfa adı önemli değil, ilk sayfa okunur.**

| ders kodu | ders adı | süre (dk) | öğrenci sayısı | tercih günleri | engelli günler | sabit gün | sabit slot | multi slots | 3 saatlik | eğitmen id’leri |
|-----------|----------|-----------|----------------|----------------|----------------|-----------|------------|-------------|-----------|-----------------|
| MATH2101 | Calculus I | 180 | 120 | 1,2 | 9,10 | | | 1 | TRUE | PROF01,ASST02 |

- **ders kodu**: Zorunlu, benzersiz, büyük harf önerilir.
- **süre (dk)**: Sınav süresi dakika cinsinden.
- **öğrenci sayısı**: Sınavı alacak toplam öğrenci.
- **tercih günleri / engelli günler**: Virgülle ayrılmış gün numaraları (örn: `1,3,5`).
- **sabit gün / sabit slot**: Dersin mutlaka yerleşmesi gereken G/S. Boş bırakılırsa esnek.
- **multi slots**: Süre birden fazla slot kaplıyorsa kaç slot tutacağı (örn: `2`).
- **3 saatlik**: `TRUE`/`1`/`YES` → süre otomatik 180 dk olur.
- **eğitmen id’leri**: Virgülle ayrılmış ID’ler. Sistemde tanımlı olmalı.

#### B. Çakışma CSV
```csv
Course1,Course2,Common Student Count
MATH2101,PHYS1001,42
MATH2101,CHEM1001,15
```
- İlk satır başlık olabilir; sistem otomatik algılar.
- Yalnızca sistemdeki ders kodlarıyla eşleşen satırlar işlenir.

#### C. Grup Kısıtları Excel
**4 kolonlu, ilk satır başlık:**

| A (Aynı Gün) | B (Aynı Slot) | C (Farklı Gün) | D (Farklı Slot) |
|--------------|---------------|----------------|-----------------|
| MATH2101,MATH2102 | PHYS1001,PHYS1002 | CHEM1001,CHEM1002 | BIOL1001,BIOL1002 |

- Her hücrede virgülle ayrılmış ders kodları.
- Aynı satırdaki dersler o kısıt grubunu oluşturur.
- Farklı satırlar bağımsız gruplardır.

---

## 3. `ana_kontrol.py` — Program Kontrol Aracı

### 3.1 Ne Zaman Kullanılır?
- Excel’den alınmış mevcut bir final programını doğrulamak istediğinizde.
- "Bu dersi başka güne taşısak çakışma olur mu?" sorusuna yanıt ararken.
- Fakülte bazlı çakışma özet raporu almak istediğinizde.

### 3.2 Çalıştırma

```bash
# Tam analiz raporu (özet)
python KONTROL_UYGULAMASI/ana_kontrol.py

# Belirli bir ders için alternatif slot ara
python KONTROL_UYGULAMASI/ana_kontrol.py --ders BUSI2111

# Sadece G1-G6 arası alternatif bak
python KONTROL_UYGULAMASI/ana_kontrol.py --ders BUSI2111 --gun-araligi 1-6

# Değişiklik simülasyonu (JSON dosyası ile)
python KONTROL_UYGULAMASI/ana_kontrol.py --degisiklik degisiklikler.json
```

### 3.3 Beklenen Girdi Dosyaları
Araç varsayılan olarak şu sabit yolları kullanır (değiştirilebilir):
- **Ana program:** `2026 Bahar Final Programı _03062026.xlsx`
  - Kolonlar: `Ders Kodu`, `Sınav Tarihi` (datetime), `Sınav Başlangıç Saati` (08:30/11:30/14:30), `Fakülte Adı`
- **Alternatif program:** `schedule (5).xlsx`
  - Kolonlar: `Ders Kodu`, `Slotlar` (örn: `G1/S2`), `Fakülte Adı`
- **Çakışma:** `final_exam_collisions_*.csv` (`Course1, Course2, Common Student Count`)
- **Kısıt:** `ders_kisitleri.xlsx` (A/B/C/D kolonları)

Parametrelerle kendi dosyalarınızı verebilirsiniz:
```bash
python ana_kontrol.py --program "benim_program.xlsx" --collision "cakisma.csv" --kisit "kisit.xlsx" --use-dates
```

### 3.4 Değişiklik Simülasyonu JSON Formatı
`degisiklikler.json`:
```json
{
  "TRAD2504": [1, 3],
  "BUSI3632": [4, 3],
  "BUSI4572": [4, 1]
}
```
Her anahtar ders kodu, değer `[gün, slot]` listesidir.

### 3.5 Çıktıların Yorumlanması
- **GERÇEK ÇAKIŞMA:** Aynı gün + aynı slot’ta ortak öğrencisi olan ders çifti. Mutlaka düzeltilmeli.
- **Farklı Gün İhlali:** `DifferentDay` kısıtlı dersler aynı günde kalmış.
- **Farklı Slot İhlali:** `DifferentSlot` kısıtlı dersler aynı periyotta kalmış.
- **Aynı Gün / Aynı Slot İhlali:** `SameDay` / `SameSlot` kısıtlı dersler farklı gün/slot’a dağılmış.
- **Eksik Ders:** Çakışma dosyasında var ama program Excel’inde yok. CORE / proje dersleri olabilir.

---

## 4. Sık Karşılaşılan Sorular

**S: Flask uygulaması ile kontrol aracı arasındaki fark nedir?**
> Flask uygulaması **program oluşturur** ve kendi içinde de **kontrol/değerlendirme** yapar (Kontrol ve Değiştir sekmeleri). `ana_kontrol.py` ise komut satırından çalıştırılan, mevcut bir Excel programını analiz eden bağımsız bir araçtır.

**S: `ARCH`, `IMIM`, `GITA`, `INAR` dersleri neden analize dahil edilmiyor?**
> Bu fakültelerin dersleri ayrı takvimde yönetilir; çakışma ve kısıt analizine dahil edilmez.

**S: Farklı slot ihlali gördüm ama dersler farklı günde. Sorun mu?**
> Hayır. `DifferentSlot` kısıtı sadece **aynı gün + aynı slot** durumunda ihlal sayılır. Farklı gün aynı saat normaldir.

**S: Değişiklik önerisi almak için Flask uygulamasındaki "Kontrol Et" butonunu mu kullanmalıyım?**
> Evet. Flask uygulaması kendi oluşturduğu program için otomatik değişiklik önerileri üretir. `ana_kontrol.py` ise dışarıdan gelen bir programı inceler.

---

## 5. Son Değişiklikler (2026-06-11)

| Değişiklik | Açıklama |
|------------|----------|
| **AppState Refactor** | Global state `@dataclass` + `field(default_factory=...)` ile yeniden yapılandırıldı. Thread-safety iyileştirildi. |
| **Feasibility Doctor** | `diagnose_all_issues()` 7 alt fonksiyona bölündü; okunabilirlik ve bakım kolaylığı arttı. |
| **Cache Helper'ları** | `get_conflict_pairs()` ve `get_group_pairs()` eklendi; çakışma verisi her fonksiyonda yeniden oluşturulmuyor. |
| **MAX_CAP Sabiti** | `10**9` magic number yerine `MAX_CAP = 10 ** 9` sabiti tanımlandı. |
| **Greedy Optimize** | `greedy_preview()` içindeki gereksiz ikinci `fits()` döngüsü kaldırıldı. |
| **Bugfix'ler** | `base_code()` lowercase replace bug düzeltildi, `to_int()` sadeleştirildi. |

---

*Son güncelleme: 2026-06-11*
