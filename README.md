# SL & OTIF

Rekap **SL (Service Level) & OTIF (On Time In Full)** dari ekspor
`PURCHASE ORDER LINES VS ARRIVAL`. Berkas asli **tidak diubah**: header, kolom,
format, dan tanggal tetap apa adanya. Kolom bantuan ditambahkan, dan sheet
**SUMMARY** dibuat dengan **rumus hidup** (`SUMIFS`) sehingga tiap angka bisa
ditelusuri.

Tersedia tiga bentuk pemakaian:

1. **Web tanpa server** — satu halaman (`web/`), semua diproses di browser.
2. **Web + backend** — React (frontend) + Flask/openpyxl (backend).
3. **CLI** — batch dari terminal.

---

## Tech Stack

| Layer             | Teknologi                              | Versi            |
| ----------------- | -------------------------------------- | ---------------- |
| Frontend          | React + Vite + Tailwind CSS            | React 19, Vite 5 |
| Backend           | Python Flask                           | Flask 3.1        |
| Excel Processing  | openpyxl (pure Python)                 | openpyxl 3.1     |
| Excel (client-side) | ExcelJS                              | 4.4              |
| Runtime           | Python / Node                          | 3.12+, Node 20+  |

> openpyxl berjalan murni di Python. **Tidak butuh Microsoft Excel / WPS.**

---

## Cara Kerja

Untuk tiap baris PO dibandingkan **PLANNED RECEIPT DATE** dengan **ARRIVAL DATE**,
memakai toleransi (bawaan **7 hari**):

| Kolom                    | Rumus                                                                 |
| ------------------------ | --------------------------------------------------------------------- |
| `PERIODE`                | `=IF(O="","",TEXT(DATEVALUE(O),"MMM"))` — bulan dari planned receipt |
| `Planned + N`            | `=IF(O="","",DATEVALUE(O)+N)`                                         |
| `GAP`                    | `=IF(OR(O="",R=0),-1,<Planned+N>-DATEVALUE(R))`                       |
| `Remarks`                | `=IF(GAP>=0,"On Time","Late")`                                        |

Lalu di sheet **SUMMARY**, per bulan (`PERIODE`) dipisah `Late` / `On Time`:

| Kolom                | Rumus                                                            |
| -------------------- | ---------------------------------------------------------------- |
| `B` Late · PO QTY    | `=SUMIFS(PO QTY, PERIODE, bulan, Remarks, "Late")`               |
| `C` Late · ARR QTY   | `=SUMIFS(ARR QTY, PERIODE, bulan, Remarks, "Late")`              |
| `D` On Time · PO QTY | `=SUMIFS(PO QTY, PERIODE, bulan, Remarks, "On Time")`            |
| `E` On Time · ARR QTY| `=SUMIFS(ARR QTY, PERIODE, bulan, Remarks, "On Time")`           |
| `F` Total PO QTY     | `=B+D`                                                           |
| `G` Total ARR QTY    | `=C+E`                                                           |
| `H` **FULFIL**       | `=G/F`                                                           |
| `I` **OTIF**         | `=E/F`                                                           |

Karena `SUMIFS` menunjuk ke kolom bantuan, ubah tanggal atau qty di sheet
`Worksheet` maka SUMMARY ikut berubah. Setiap rumus juga menyertakan **nilai
cache** agar angka langsung tampil, dan berkas diset *recalculate on open*.

**Pertahankan format asli.** Generator bekerja *in-place* pada berkas sumber
(openpyxl memuat + menyimpan ulang), jadi warna, border, lebar kolom, tinggi
baris, dan `freeze panes` ikut terbawa. Kolom bantuan hanya **ditambahkan** di
kanan.

---

## Struktur File

```
sl-otif/
├─ sl_otif_core.py         # inti: parse, hitung, tulis rumus + cache (openpyxl)
├─ generate_sl_otif.py     # CLI: python generate_sl_otif.py
├─ backend/                # Flask API
│  ├─ app.py               # /api/health, /api/process
│  └─ requirements.txt
├─ frontend/               # React + Vite + Tailwind (UI utama)
│  ├─ src/App.jsx
│  ├─ src/index.css
│  ├─ tailwind.config.js
│  ├─ vite.config.js       # dev proxy /api -> 127.0.0.1:5000
│  ├─ vercel.json
│  └─ .env.example         # VITE_API_BASE
├─ web/                    # Web tanpa server (ExcelJS)
│  ├─ index.html
│  ├─ sl-otif-core.js
│  └─ exceljs.min.js
├─ render.yaml             # blueprint backend Flask (Render)
├─ .gitignore
└─ README.md
```

> Data mentah `*.xlsx` dikecualikan dari git (lihat `.gitignore`) karena berisi
> data nyata.

---

## Instalasi

Kebutuhan: **Python 3.12+** dan **Node 20+**.

### 1. CLI

```bash
pip install openpyxl

# proses semua .xlsx di folder skrip
python generate_sl_otif.py

# atau pilih berkas + folder output + toleransi
python generate_sl_otif.py "us.xlsx" -o hasil -n 7
```

Hasil tersimpan di `hasil/` dengan nama `2026_<NAMA>_SL & OTIF.xlsx`.

### 2. Web tanpa server (`web/`)

Tidak perlu install apa pun, tidak perlu server.

```bash
# cukup buka berkasnya
start web/index.html          # Windows
open  web/index.html          # macOS
```

Semua diproses di browser (ExcelJS). Bisa juga langsung di-deploy ke Vercel
sebagai situs statis.

### 3. Backend (Flask + openpyxl)

```bash
cd backend
pip install -r requirements.txt
python app.py                 # http://localhost:5000
```

Endpoint:

| Method | Path           | Keterangan                                            |
| ------ | -------------- | ----------------------------------------------------- |
| GET    | `/api/health`  | cek status                                            |
| POST   | `/api/process` | multipart: `files` (satu/banyak .xlsx), `tol` (opsional) |

Contoh:

```bash
curl -X POST http://localhost:5000/api/process \
  -F "files=@us.xlsx" -F "tol=7"
```

### 4. Frontend (React + Vite + Tailwind)

```bash
cd frontend
npm install
npm run dev                   # http://localhost:5173
```

Saat `npm run dev`, Vite mem-proxy `/api/*` ke `http://127.0.0.1:5000`
(jalankan backend lebih dulu). Untuk menunjuk backend lain, isi `.env`:

```
VITE_API_BASE=https://sl-otif-api.onrender.com
```

Build produksi:

```bash
npm run build                 # hasil di frontend/dist
npm run preview
```

---

## Deploy

### GitHub

```bash
git init
git add .
git commit -m "SL & OTIF: Flask + openpyxl, React + Vite, web standalone"
git branch -M main
git remote add origin https://github.com/<user>/sl-otif.git
git push -u origin main
```

### Vercel

Ada tiga opsi:

**A. Statis (tanpa backend)** — paling sederhana. Deploy folder `web/`
(Root Directory: `web`, framework: *Other*). Semua jalan di browser, tanpa batas
server.

**B. React di Vercel + Flask di Render** — memakai openpyxl.
- Vercel: Root Directory `frontend`, framework *Vite* (sudah ada `vercel.json`).
- Backend: Render → New → **Blueprint** → pilih repo (`render.yaml`). Start
  command memakai gunicorn.
- Set env di Vercel: `VITE_API_BASE=https://<backend>.onrender.com`.

**C. Semua di Vercel** — Flask dijalankan sebagai Vercel Function (Python
runtime mendukung WSGI `app`), digabung dengan frontend lewat *Services*.
Perhatikan batas serverless:

| Batas              | Nilai          | Catatan                                   |
| ------------------ | -------------- | ----------------------------------------- |
| Body request       | ± 4,5 MB       | cukup untuk berkas ratusan KB             |
| Durasi fungsi      | puluhan detik  | berkas sangat besar bisa timeout          |
| Cold start         | ada            | request pertama terasa lambat             |

---

## Catatan

- Tanggal di berkas sumber bertipe **teks** (ISO `YYYY-MM-DD`), karena itu
  dipakai `DATEVALUE` di rumus bantuan.
- Baris tanpa `ARRIVAL DATE` dihitung **Late** (GAP = -1).
- Berkas contoh menyimpan `Periode` sebagai teks sehingga rumus `TEXT(...)` di
  berkas itu tidak benar-benar menghitung; generator ini menulisnya sebagai
  rumus yang valid.

## Lisensi

MIT
