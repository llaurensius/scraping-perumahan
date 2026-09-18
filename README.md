# Sikumbang - Scraping Perumahan Subsidi

Aplikasi web lokal berbasis Flask untuk mencari, memfilter, dan melakukan scraping/ekspor data perumahan subsidi langsung dari portal [sikumbang.tapera.go.id](https://sikumbang.tapera.go.id/).

---

## Prasyarat (Prerequisites)

- **Python 3.8+** terpasang di sistem Anda.
- Akses internet aktif (untuk mengambil data dari Sikumbang API).

---

## Panduan Instalasi & Menjalankan Aplikasi

Ikuti langkah-langkah berikut di terminal (PowerShell, Command Prompt, atau Bash) untuk menjalankan aplikasi:

### 1. Masuk ke Folder Proyek

Pastikan Anda berada di dalam folder proyek `scraping-perumahan`:

```bash
cd scraping-perumahan
```

### 2. (Opsional tapi Disarankan) Buat & Aktifkan Virtual Environment

Membuat virtual environment membantu mengisolasi dependensi aplikasi agar tidak bentrok dengan paket Python global.

- **Windows (PowerShell / Command Prompt):**
  ```powershell
  python -m venv venv
  .\venv\Scripts\activate
  ```
  *(Jika menggunakan PowerShell dan muncul error permission script, jalankan `Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass` terlebih dahulu)*

- **Linux / macOS:**
  ```bash
  python3 -m venv venv
  source venv/bin/activate
  ```

### 3. Pasang Dependensi

Pasang pustaka yang diperlukan (`Flask`, `requests`, `openpyxl`) menggunakan `requirements.txt`:

```bash
pip install -r requirements.txt
```

### 4. Jalankan Aplikasi

Jalankan server backend `app.py`:

```bash
python app.py
```

Setelah dijalankan, terminal akan menampilkan informasi berikut:
```text
==================================================
  SIKUMBANG - Scraping Perumahan Subsidi
  Buka browser: http://localhost:8081
==================================================
```

### 5. Buka di Browser

Buka peramban (browser) Anda dan akses alamat:
👉 **[http://localhost:8081](http://localhost:8081)** (atau `http://127.0.0.1:8081`)

---

## Fitur Aplikasi

- **Filter Wilayah & Asosiasi**: Filter data berdasarkan Provinsi, Kabupaten/Kota, Asosiasi Pengembang, dan Status (Subsidi / Komersil).
- **Pencarian Cepat**: Cari perumahan berdasarkan nama perumahan secara langsung.
- **Pilihan Interaktif**: Centang perumahan yang ingin dianalisis atau disimpan.
- **Penyimpanan Lokal**: Data perumahan yang dicentang tersimpan otomatis di peramban (`localStorage`) sehingga tidak hilang saat berpindah halaman.
- **Ekspor Data**: 
  - Format **CSV** (dengan dukungan encoding UTF-8 BOM untuk Microsoft Excel).
  - Format **Excel (.xlsx)** lengkap dengan styling header, format angka ribuan, dan penyesuaian lebar kolom otomatis.
  - File hasil ekspor otomatis disimpan di dalam folder `exports/` dan langsung diunduh lewat browser.

---

## Struktur Folder

```text
scraping-perumahan/
├── app.py              # Server backend Flask & proxy API Sikumbang
├── requirements.txt    # Daftar dependensi Python (Flask, requests, openpyxl)
├── templates/
│   └── index.html      # Tampilan antarmuka Web UI (HTML + CSS + JS)
├── exports/            # Folder penyimpanan file hasil ekspor (CSV/XLSX)
├── .gitignore          # Konfigurasi file yang diabaikan oleh Git
└── README.md           # Dokumentasi dan panduan penggunaan
```

---

## Dependensi Utama

- [Flask](https://flask.palletsprojects.com/) - Web framework server lokal
- [Requests](https://requests.readthedocs.io/) - Mengirim permintaan HTTP ke API Sikumbang
- [openpyxl](https://openpyxl.readthedocs.io/) - Pembuatan dan manipulasi file spreadsheet Excel (.xlsx)
