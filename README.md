# MyLife — Oxford 5000™ Vocabulary & Smart Finance Tracker (Python to APK)

Aplikasi **MyLife** dibangun menggunakan **Python (Flet Mobile Framework)** dan **SQLite**, menggabungkan:
1. **Hafalan, Latihan & Tebakan 2.000 Kosakata Bahasa Inggris (American Oxford 5000™)** lengkap dengan **Tracker Bulanan & Tahunan**.
2. **Pencatat Keuangan Pintar (50/30/20 & 70/20/10 Budget Tracker)** sesuai `template_keuangan.xlsx` lengkap dengan **Set-Up Budget, Transaksi, Overview Bulanan, Overview Tahunan, Tracker Langganan, dan Target Tabungan Impian (Saving Goals)**.

---

## 🎨 Palet Warna Aplikasi
- `#0E0B1D` (`#OEOB1D`) — Background Utama (Midnight Deep Purple)
- `#2A1E4D` — Card & Kontainer Utama (Dark Royal Purple)
- `#5A3B8C` — Tombol & Aksen Utama (Vibrant Purple)
- `#8B7CC7` — Aksen Sekunder, Grafik & Badge (Soft Lavender Purple)
- `#C9C2E3` — Teks Utama (Pastel Lavender)
- `#3A3A52` — Border, Input Field & Elemen Pendukung (Muted Slate Purple)

---

## 📂 Daftar File Utama
- [`Oxford_5000_Kosakata.xlsx`](file:///d:/github/my-life/Oxford_5000_Kosakata.xlsx) — **(Tugas 1)** File Excel berisi seluruh **2.000 kosakata** dari `American_Oxford_5000 (1).pdf` dengan 3 kolom: **Kata Bahasa Inggris**, **Arti**, dan **Penjelasan**.
- [`main.py`](file:///d:/github/my-life/main.py) — **(Tugas 2)** Aplikasi utama Python Flet dengan navigasi 5 menu utama.
- [`database.py`](file:///d:/github/my-life/database.py) — Modul database SQLite (`mylife_app.db`) untuk menyimpan progres hafalan, log latihan/tebakan bulanan & tahunan, transaksi keuangan, langganan, dan target tabungan.
- [`excel_exporter.py`](file:///d:/github/my-life/excel_exporter.py) — Modul pembuat file Excel kosakata dan ekspor laporan aplikasi.
- [`pyproject.toml`](file:///d:/github/my-life/pyproject.toml) & [`build_apk.bat`](file:///d:/github/my-life/build_apk.bat) — Konfigurasi resmi kompilasi Python menjadi **Android `.apk`**.
- [`.github/workflows/build_apk.yml`](file:///d:/github/my-life/.github/workflows/build_apk.yml) — Workflow GitHub Actions untuk mem-build file `.apk` secara otomatis di cloud tanpa perlu menginstal Android Studio 5 GB di laptop.

---

## 🚀 Cara Menjalankan Aplikasi di Komputer / Laptop
Jalankan perintah berikut di terminal pada folder proyek:
```bash
python main.py
```

---

## 📱 Cara Mengubah menjadi File Android `.apk`

### Cara 1: Otomatis via GitHub Actions (Tanpa Install Android Studio)
Karena folder proyek ini berada di repositori Git (`d:\github\my-life`), Anda cukup melakukan `git push` ke GitHub:
1. Push seluruh file ke repositori GitHub Anda.
2. Buka tab **Actions** di repositori GitHub Anda, lalu pilih workflow **"Build Android APK (Flet Python)"**.
3. Setelah proses selesai, unduh file **`MyLife-Vocab-Finance-APK`** (`.apk`) yang siap diinstal di HP Android Anda.

### Cara 2: Build Lokal di Komputer Windows
Jika di komputer Anda sudah terinstal **Android Studio / Android SDK**:
1. Klik dua kali file [`build_apk.bat`](file:///d:/github/my-life/build_apk.bat) atau jalankan di terminal:
   ```bash
   flet build apk --verbose
   ```
2. File `.apk` hasil kompilasi akan tersimpan di folder `build/apk/`.

