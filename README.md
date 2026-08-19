# SIMAMORA BPS — Sistem Monitoring & Evaluasi Kinerja

Aplikasi monitoring kinerja dengan 4 modul:
- **FRA** — capaian kinerja tiap tim + bukti dukung laporan
- **Analisis Kegiatan** — kendala, solusi, RTL (rencana tindak lanjut)
- **Tugas Tim & Target Tahunan** — target per tahun, realisasi & bukti dukung per triwulan (Q1–Q4)
- **IKU** — indikator kinerja utama + bukti laporan

Berbeda dari versi sebelumnya, versi ini punya **backend sungguhan**: semua data (teks) tersimpan
di `data/db.json` dan **semua berkas PDF/Word tersimpan fisik di server**, di folder `uploads/`.

## Struktur folder

```
simamora-server/
├── server.js          # backend Express (API + upload berkas)
├── package.json
├── data/
│   └── db.json         # dibuat otomatis saat pertama kali server jalan
├── uploads/             # tempat berkas PDF/Word bukti dukung disimpan
└── public/
    └── index.html       # frontend (dilayani langsung oleh server)
```

## Menjalankan di komputer lokal

Syarat: [Node.js](https://nodejs.org) versi 18 ke atas.

```bash
cd simamora-server
npm install
npm start
```

Lalu buka **http://localhost:3000** di browser. Data tim contoh akan otomatis terisi saat pertama kali dijalankan.

## Cara kerja penyimpanan berkas

- Saat pengguna mengunggah PDF/Word lewat form (FRA, IKU, atau triwulan Tugas Tim), berkas langsung
  dikirim ke endpoint `POST /api/upload`, disimpan di folder `uploads/` dengan nama unik (UUID),
  dan server mengembalikan info berkas (nama asli, ukuran, nama file di server).
- Info tersebut (bukan isi berkasnya) disimpan sebagai referensi di `data/db.json`.
- Untuk melihat/mengunduh, frontend membuka `GET /api/files/:filename`.
- Format yang diterima: `.pdf`, `.doc`, `.docx`. Ukuran maksimal saat ini **10MB** per berkas
  (bisa diubah lewat variabel `MAX_FILE_MB` di `server.js` dan `MAX_FILE_MB` di `public/index.html`).

## Deploy ke server sungguhan (production)

Aplikasi ini adalah aplikasi Node.js biasa, jadi bisa dijalankan di VPS/cloud (mis. server milik BPS,
DigitalOcean, AWS EC2, dsb). Langkah umum:

1. Salin seluruh folder `simamora-server` ke server.
2. Install Node.js 18+ di server tersebut.
3. `npm install --production`
4. Jalankan dengan process manager agar tetap hidup, contoh dengan [PM2](https://pm2.keymetrics.io/):
   ```bash
   npm install -g pm2
   pm2 start server.js --name simamora-bps
   pm2 save
   pm2 startup
   ```
5. Pasang reverse proxy **Nginx** di depan (agar bisa pakai domain & HTTPS):
   ```nginx
   server {
       listen 80;
       server_name simamora.bps-daerah.go.id;
       location / {
           proxy_pass http://localhost:3000;
           proxy_set_header Host $host;
           proxy_set_header X-Real-IP $remote_addr;
       }
       client_max_body_size 15M;  # penting agar upload berkas tidak ditolak Nginx
   }
   ```
6. Aktifkan HTTPS, misalnya dengan `certbot` (Let's Encrypt).
7. **Backup rutin** folder `data/` dan `uploads/` — keduanya adalah "database" aplikasi ini.

## Login & hak akses

Halaman login (split-screen, tema BPS) muncul otomatis untuk semua rute — tidak ada data yang bisa
diakses tanpa login. Akun bawaan (edit langsung di `USER_ACCOUNTS` pada `server.js` untuk mengganti):

| Username               | Password    | Peran    | Tim                                            |
|------------------------|-------------|----------|-------------------------------------------------|
| admin                  | admin123    | admin    | semua tim (akses penuh + Riwayat Aktivitas)      |
| operator_sosial        | opsosial1   | operator | Tim Statistik Sosial                             |
| operator_produksi      | opprod1     | operator | Tim Statistik Produksi                           |
| operator_distribusi    | opdist1     | operator | Tim Statistik Distribusi                         |
| operator_neraca        | opneraca1   | operator | Tim Neraca & Analisis Statistik                  |
| operator_ipds          | opipds1     | operator | Tim IPDS                                         |
| operator_tu            | optu1       | operator | Tim Tata Usaha                                   |

Operator hanya bisa melihat/mengubah data milik timnya sendiri; admin bisa mengelola semua tim, kelola
daftar tim, dan melihat halaman **Riwayat Aktivitas** (audit log). Session disimpan lewat `express-session`
(cookie `httpOnly`, berlaku 24 jam). Untuk produksi, set `SESSION_SECRET` di environment variable agar
tidak pakai nilai default.

## Validasi, sanitasi & backup

- Semua endpoint tulis (`POST`/`PUT`) divalidasi dengan **Joi** (tipe data, field wajib, rentang angka)
  sebelum masuk ke `db.json` — request lewat Postman/curl tanpa field yang benar akan ditolak (`400`).
  Input teks juga disanitasi (tag HTML dibuang) untuk mencegah injeksi skrip.
- Ukuran body JSON dibatasi (`256kb` umum, `64kb` untuk form url-encoded); upload berkas dibatasi
  `MAX_FILE_MB` (10MB) dan nama file disimpan ulang dengan UUID (tidak memakai nama asli) agar tidak
  bisa menimpa file sistem.
- `data/db.json` di-backup otomatis ke `data/backup/` setiap 24 jam (dan sekali saat server start).
  Semua perubahan (create/update/delete) tercatat di **Riwayat Aktivitas**.
- Export data tersedia dalam **JSON**, **Excel** (`exceljs`), dan **PDF** (`pdfkit`) — lihat tombol
  "Export" di tiap halaman modul, atau langsung lewat `GET /api/export/{json|excel|pdf}?collection=...`.

## Pengingat triwulan belum diisi

Dashboard menampilkan spanduk peringatan dan ikon lonceng di pojok kanan atas jika ada **Tugas Tim**
yang realisasinya belum diisi untuk triwulan yang sedang berjalan. Klik salah satu item untuk langsung
menuju halaman Tugas Tim dengan triwulan tersebut terbuka. Sumber data: `GET /api/reminders`.

**Opsional — pengingat via email (H-7 sebelum akhir triwulan):**
Fitur ini otomatis nonaktif jika environment variable SMTP tidak diisi (tidak mengganggu jalannya
server). Untuk mengaktifkan, set variabel berikut sebelum `npm start`:

```bash
SMTP_HOST=smtp.gmail.com
SMTP_PORT=587
SMTP_USER=akun@gmail.com
SMTP_PASS=app-password
SMTP_FROM="SIMAMORA BPS <akun@gmail.com>"
# opsional: alamat tujuan per tim, dipisah koma, format Nama Tim=email
REMINDER_EMAILS="Tim Statistik Sosial=sosial@bps.go.id,Tim Statistik Produksi=produksi@bps.go.id"
```

Server mengecek sekali sehari; email hanya dikirim jika sisa hari triwulan berjalan ≤ 7 hari **dan**
masih ada tugas yang belum diisi.

## Pencarian, filter & grafik

- Tiap halaman modul (FRA, Tugas, IKU, Kegiatan) punya kotak pencarian bebas, filter per Tim, dan
  filter Status (Tercapai / Perlu Perhatian / Tertinggal) — dihitung juga di sisi server lewat query
  string (`?search=&teamId=&status=`) sehingga export mengikuti filter yang sedang aktif.
- Dashboard menampilkan grafik tren realisasi per triwulan dan tabel progres per tim (canvas, tanpa
  dependensi eksternal di frontend).

## Yang masih bisa ditingkatkan lebih lanjut

- Migrasi dari `db.json` ke SQLite/PostgreSQL jika volume data & jumlah pengguna bersamaan makin besar.
- Notifikasi push/WhatsApp selain email untuk pengingat triwulan.
- Two-factor authentication untuk akun admin.

Jika perlu, semua ini bisa ditambahkan secara bertahap di atas struktur yang sudah ada.
