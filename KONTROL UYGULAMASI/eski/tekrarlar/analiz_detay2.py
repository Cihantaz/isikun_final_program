import pandas as pd

# Dosya yolları
collision_file = r"C:\Users\cihan.tazeoz\Desktop\26 BAHAR FİNAL\UYGULAMA DOSYALARI\final_exam_collisions_14_05_2026_10_48_47.csv"
schedule_file = r"C:\Users\cihan.tazeoz\Desktop\26 BAHAR FİNAL\UYGULAMA DOSYALARI\schedule (5).xlsx"

# 1. Dosyaları oku
collisions = pd.read_csv(collision_file)
schedule = pd.read_excel(schedule_file)

EXCLUDED_PREFIXES = ('ARCH', 'İMİM', 'GİTA', 'INAR')

# Schedule'daki ders kodları
schedule_courses = set(str(c).strip().upper() for c in schedule['Ders Kodu'])

# Çakışma dosyasındaki tüm ders kodları
collision_courses = set()
for _, row in collisions.iterrows():
    c1 = str(row['Course1']).strip().upper()
    c2 = str(row['Course2']).strip().upper()
    if not c1.startswith(EXCLUDED_PREFIXES) and not c2.startswith(EXCLUDED_PREFIXES):
        collision_courses.add(c1)
        collision_courses.add(c2)

# Schedule'da olmayan dersler
missing_in_schedule = collision_courses - schedule_courses
print(f"Çakışma dosyasındaki toplam ders sayısı (hariçler hariç): {len(collision_courses)}")
print(f"Schedule'daki ders sayısı: {len(schedule_courses)}")
print(f"Schedule'da olmayan ders sayısı: {len(missing_in_schedule)}")
print("\nSchedule'da olmayan dersler (ilk 50):")
for c in sorted(missing_in_schedule)[:50]:
    print(f"  {c}")

# Eksik derslerin çakışma dosyasındaki toplam çakışma sayısı
missing_collision_count = 0
for _, row in collisions.iterrows():
    c1 = str(row['Course1']).strip().upper()
    c2 = str(row['Course2']).strip().upper()
    if c1 in missing_in_schedule or c2 in missing_in_schedule:
        if not c1.startswith(EXCLUDED_PREFIXES) and not c2.startswith(EXCLUDED_PREFIXES):
            missing_collision_count += 1

print(f"\nEksik dersleri içeren çakışma çifti sayısı: {missing_collision_count}")
