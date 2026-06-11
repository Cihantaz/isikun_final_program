import pandas as pd
import os
from datetime import datetime

# Dosya yollari
in_file = r"C:\Users\cihan.tazeoz\Desktop\26 BAHAR FİNAL\UYGULAMA DOSYALARI\YENILENMIS_PROGRAM_2026_BAHAR_FINAL.xlsx"
temp_file = r"C:\Users\cihan.tazeoz\Desktop\26 BAHAR FİNAL\UYGULAMA DOSYALARI\YENILENMIS_PROGRAM_2026_BAHAR_FINAL_TEMP.xlsx"
out_file = r"C:\Users\cihan.tazeoz\Desktop\26 BAHAR FİNAL\UYGULAMA DOSYALARI\YENILENMIS_PROGRAM_2026_BAHAR_FINAL.xlsx"

# Gun-Tarih eslestirmesi (haftasonu atlaniyor)
gun_tarihleri = {
    1: datetime(2026, 6, 8),   # Pazartesi
    2: datetime(2026, 6, 9),   # Sali
    3: datetime(2026, 6, 10),  # Carsamba
    4: datetime(2026, 6, 11),  # Persembe
    5: datetime(2026, 6, 12),  # Cuma
    6: datetime(2026, 6, 15),  # Pazartesi (haftasonu atlandi)
    7: datetime(2026, 6, 16),  # Sali
    8: datetime(2026, 6, 17),  # Carsamba
    9: datetime(2026, 6, 18),  # Persembe
    10: datetime(2026, 6, 19), # Cuma
}

gun_adlari = {
    1: "Pazartesi",
    2: "Salı",
    3: "Çarşamba",
    4: "Perşembe",
    5: "Cuma",
    6: "Pazartesi",
    7: "Salı",
    8: "Çarşamba",
    9: "Perşembe",
    10: "Cuma",
}

slot_saatleri = {
    1: ("08:30", "11:30"),
    2: ("11:30", "14:30"),
    3: ("14:30", "17:30"),
}

# Exceli oku
df = pd.read_excel(in_file)

# Yeni sutunlar ekle
sinav_tarihi = []
sinav_gunu = []
sinav_baslangic = []
sinav_bitis = []

for _, row in df.iterrows():
    gun = int(row['Gün'])
    slot = int(row['Slot'])
    
    tarih = gun_tarihleri[gun]
    gun_adi = gun_adlari[gun]
    baslangic, bitis = slot_saatleri[slot]
    
    sinav_tarihi.append(tarih.strftime("%d.%m.%Y"))
    sinav_gunu.append(gun_adi)
    sinav_baslangic.append(baslangic)
    sinav_bitis.append(bitis)

df['Sınav Tarihi'] = sinav_tarihi
df['Sınav Günü'] = sinav_gunu
df['Sınav Başlangıç Saati'] = sinav_baslangic
df['Sınav Bitiş Saati'] = sinav_bitis

# Sutun sirasini duzenle
sutun_sirasi = ['Ders Kodu', 'Ders Adı', 'Sınav Tarihi', 'Sınav Günü', 'Sınav Başlangıç Saati', 
                'Sınav Bitiş Saati', 'Gün', 'Slot', 'Yeni Slot', 'Eski Slot', 'Değişti']

# Mevcut sutunlardan sadece var olanlari al
mevcut_sutunlar = [s for s in sutun_sirasi if s in df.columns]
df = df[mevcut_sutunlar]

# Once gecici dosyaya kaydet
df.to_excel(temp_file, index=False)

# Eski dosyayi sil ve yenisini yeniden adlandir
try:
    if os.path.exists(out_file):
        os.remove(out_file)
    os.rename(temp_file, out_file)
    print(f"Dosya guncellendi: {out_file}")
except Exception as e:
    print(f"Dosya degistirilemedi: {e}")
    print(f"Gecici dosya kaydedildi: {temp_file}")

print(f"Toplam ders: {len(df)}")
print("\nOrnek satirlar:")
print(df.head(10).to_string(index=False))
