# ABelSekolah

Aplikasi bel sekolah otomatis. Tampilan (frontend) di Cloudflare Pages, data di Google Sheet lewat Google Apps Script.

```
index.html, config.js   -> frontend (GitHub -> Cloudflare Pages)
apps-script/Code.gs     -> backend API (Google Apps Script + Google Sheet)
```

## 1. Perbarui Apps Script
1. Buka Google Sheet Anda, lalu **Extensions -> Apps Script**.
2. Ganti isi `Code.gs` dengan `apps-script/Code.gs`. File `Index.html` di Apps Script boleh dihapus.
3. **Deploy -> Manage deployments -> Edit -> New version -> Deploy**. Pastikan *Execute as: Me* dan *Who has access: Anyone*. URL `/exec` tetap sama.
4. Buka URL `/exec` di browser. Jika muncul `{"ok":true,...}`, API sudah aktif.

## 2. Upload ke GitHub
Buat repository baru, lalu unggah `index.html`, `config.js`, `README.md`, dan folder `apps-script`.

## 3. Hubungkan ke Cloudflare Pages
1. Cloudflare Dashboard -> **Workers & Pages -> Create -> Pages -> Connect to Git**, pilih repository.
2. Framework preset: **None**. Build command: kosongkan. Build output directory: `/` (atau kosongkan).
3. Klik **Save and Deploy**. Aplikasi bisa dibuka di `https://nama-proyek.pages.dev`.

Jika URL deployment Apps Script berubah, cukup ubah `config.js`.

## PIN admin (opsional)
Isi sel di samping `pin` pada sheet **Pengaturan** (misalnya `1234`). Setelah itu PIN diminta saat mengubah jadwal, audio, logo, atau pengaturan. Kosongkan untuk mematikan. Karena repository dan URL API bisa dilihat publik, sebaiknya PIN diisi.
