# 🔍 Re-invent vs KRI

Aplikasi Streamlit **1 file (`app.py`)** untuk membandingkan data Re-invent dengan KRI.
Tujuan: semua yang ada di KRI harus ada di Re-invent.

## Jalankan lokal
```bash
pip install -r requirements.txt
streamlit run app.py
```
Sumber data (sidebar): **Upload file**, **file di folder `data/`** (nama berisi `reinvent` dan `kri`), atau **Data contoh (dummy)**.

## Deploy ke Streamlit Community Cloud
```bash
git init && git add . && git commit -m "Initial commit"
git branch -M main
git remote add origin https://github.com/<user>/reinvent-vs-kri.git
git push -u origin main
```
Buka <https://share.streamlit.io> → Create app → pilih repo → Main file path: `app.py` → Deploy.

## Keluaran
- `Reinvent_Updated.xlsx` — file utama, kolom persis Re-invent.
  Warna: 🟨 hostname beda IP sama · 🟧 Hostname/VM-Baremetal kosong · ⬜ dismantle · 🟩 baris baru dari KRI
- `Laporan_Perbandingan.xlsx` — ringkasan, selisih data, IP sama, dismantle, perlu dilengkapi, duplikat, log perubahan.

## ⚠️ Keamanan data
Jangan commit Excel asli ke repo publik (sudah dicegah `.gitignore`). Atur app Streamlit ke private jika datanya internal.
