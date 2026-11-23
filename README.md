# Jadwal Toko (OutboundPro)

Vercel (web + API) · Firebase Auth (login) · Google Apps Script + Google Sheets (database, backup, email alarm).

## Isi folder
- `public/index.html` — seluruh tampilan web
- `api/index.js` — API Vercel (cek login & hak akses, teruskan ke Apps Script)
- `Code.gs` — backend Google Apps Script
- `data/JPF_DETAIL_VS_REALISASI_2026-11.xlsx` — Excel DETAIL JPF vs Realisasi (+ Master Group, Rekap Branch, Toko Tutup) versi ringan (500 KB) untuk Impor Awal

## Cara update dari versi lama (data lama aman)
1. **Apps Script**: ganti isi `Code.gs` dengan file ini → jalankan `setup()` lalu `installTriggers()` (pasang backup 02:00 + email alarm 06:00) → *Deploy > Manage deployments > pensil > Version: New version > Deploy*. Pastikan zona waktu proyek = Asia/Jakarta.
2. **Vercel**: ganti `public/index.html` dan `api/index.js`, lalu deploy ulang.
3. **Web > Admin > Impor Awal dari File JPF**: pilih `data/JPF_DETAIL_VS_REALISASI_2026-11.xlsx` (menimpa master toko & group; backup otomatis).

Kolom baru hanya ditambahkan di kanan sheet yang sudah ada (TOKO: RIT · BLACKOUT: MODE, GESER_HARI, JML_TOKO, WAKTU, TGL_AKHIR · SL_yyyy-mm: RP_PB, QTY_PICK, RP_PICK, RP_SHIP, SL_QTY, TOT_ITEM). Sheet baru: TARGET, SETTING, BRANCH.

## Aturan yang dipakai
- **Target**: isi di Dashboard (total, atau per branch; yang kosong dibagi menurut jumlah toko). Status: TERCAPAI (aktual ≥ target) · ON TRACK / BERISIKO (proyeksi = aktual + sisa plan × % aktual berjalan) · TIDAK TERCAPAI (bulan selesai, aktual < target).
- **Hitung s/d tanggal**: otomatis mengikuti tanggal SL terakhir di-upload (bisa diubah manual).
- **Alarm** (dashboard, menu Tutup/GO, badge menu, email harian): GO hari ini / H-3, GO sudah jalan tapi belum ada SL, Tutup hari ini / H-3, Tutup tapi masih ada SL, toko > 4 hari belum kirim (dihitung dalam bulan berjalan), SL belum di-upload ≥ 2 hari. Email ke alamat di Admin > Pengaturan.
- **Libur / SO**: rentang tanggal, cakupan semua / branch / group, aksi *Kosongkan* atau *Geser +N hari* (seperti "jadwal geser" di file JPF). Tombol *Kembalikan* per entri dan *Kembalikan SEMUA* mengembalikan jadwal awal; riwayat mencatat jumlah toko.
- **Upload SL**: format `SL_DC_SS08_Per_Toko` (TANGGAL_PB, KD_STORE, NAMA_STORE, QTY_PB, RP_PB, QTY_PICK, RP_PICK, QTY_SHIP, RP_SHIP, SL_QTY, TOT_ITEM); format lama masih bisa. Baris tanpa QTY_SHIP/RP_SHIP tidak dihitung sebagai aktual.
- **Download**: Report JWC Toko (identik REP_JWC_TOKO: .xls HTML / .xlsx / .csv), plan atau aktual, per branch; juga matrix ringkas gaya JPF.

## Catatan data
- Kode **BJ74** muncul 2× di file JPF (baris 3049 & 3057, toko berbeda). Impor hanya memakai yang pertama — periksa kode toko yang kedua.
- Kolom **RIT** di Excel `data/` diambil dari file REP_JWC_TOKO Agustus 2026 (3.258 toko cocok, sisanya default 1). Ubah per toko lewat detail toko.
