# Simulasi Register Surat Desa (Offline / PWA)

Ini adalah aplikasi simulasi berbasis website yang fokus pada:
- Nomor register **urut per klasifikasi** (Surat Keluar & Surat Keterangan)
- Counter **reset per tahun** (tahun diambil dari tanggal surat)
- Data tersimpan lokal (localStorage), bisa dipakai offline

## Cara menjalankan di laptop/PC
1. Ekstrak file ZIP ini.
2. Jalankan server lokal (pilih salah satu):
   - Python: `python -m http.server 8000`
   - Node: `npx serve .`
3. Buka browser ke: `http://localhost:8000`

> Catatan: PWA/service-worker butuh dibuka lewat http(s), bukan file://

## Cara install di Android (paling mudah)
1. Pindahkan folder hasil ekstrak ke laptop/PC dan jalankan server lokal seperti di atas, lalu akses dari HP via jaringan yang sama, **atau** upload folder ini ke hosting (HTTPS).
2. Buka situsnya via Chrome Android.
3. Menu ⋮ → **Tambahkan ke layar utama** (atau Install app).

## Export / Import
- Export JSON untuk backup/manual pindah perangkat.
- Import JSON untuk restore.

## Catatan penting
Ini simulasi single-user (lokal). Untuk kantor desa multi-user, generator nomor urut harus pakai database + transaksi supaya tidak dobel.
