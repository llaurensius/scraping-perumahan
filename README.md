# Sikumbang - Scraping Perumahan Subsidi

Aplikasi web lokal untuk scraping data perumahan subsidi dari [sikumbang.tapera.go.id](https://sikumbang.tapera.go.id/).

## Cara Pakai

```bash
python app.py
```

Buka browser → **http://localhost:8081**

## Fitur

- Filter by Provinsi, Kabupaten/Kota, Asosiasi, Status (Subsidi/Komersil)
- Pencarian nama perumahan
- Centang lokasi yang diinginkan
- Export ke CSV atau XLSX
- Selection tersimpan otomatis di browser (localStorage)

## Dependencies

- Python 3.8+
- Flask
- requests
- openpyxl

Semua sudah tersedia di environment, tidak perlu install tambahan.

## File Structure

```
scraping-perumahan/
├── app.py              # Flask backend
├── templates/
│   └── index.html      # Frontend UI
├── exports/            # Generated export files
└── README.md
```
