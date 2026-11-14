# Jadwalin: akun & data di Google Sheets (tanpa Firebase)

File: `index.html` (web), `apps-script.gs` (server), `PANDUAN.md` (ini).

## 1. Pasang server di Google Sheets
1. Buat Google Sheet baru (nama bebas, misalnya "Database Jadwalin"), atau pakai sheet lama kamu.
2. **Ekstensi > Apps Script**, hapus isi bawaan, tempel seluruh isi `apps-script.gs`, lalu simpan.
3. **Izinkan pengiriman email (wajib, agar Lupa kata sandi jalan):**
   - Di editor Apps Script, pilih fungsi **`izinkanEmail`** di menu dropdown atas, klik **Run**.
   - Klik **Review permissions**, pilih akunmu, **Advanced > Go to ... (unsafe) > Allow**.
   - Email tes akan masuk ke inbox kamu. Itu tanda izin sudah aktif.
4. **Deploy > New deployment > Web app**. Execute as: **Me**. Who has access: **Anyone**. Klik Deploy dan salin **URL Web app**.
   - Kalau kamu sudah punya deployment lama: **Deploy > Manage deployments > ikon pensil > Version: New version > Deploy**. URL-nya tetap sama.
5. Buka `index.html`, pastikan `API_URL` di bagian `PENGATURAN` berisi URL tadi.

Setiap kali `apps-script.gs` diubah, buat versi deployment baru (langkah 4), kalau tidak perubahannya tidak aktif.

Sheet lama tidak perlu diubah manual: kolom baru (foto, sosmed, kolab, fee, mulai, kirim, brand, pic, kontak, bayar, dibayar, keranjang, boost) ditambahkan otomatis saat pertama dipakai.

## 2. Hosting di Vercel
1. Taruh `index.html` sendirian di satu folder (misalnya `jadwalin/`).
2. Cara paling mudah: buat repo GitHub berisi file itu, lalu di vercel.com pilih **Add New > Project**, impor repo, **Framework Preset: Other**, kosongkan Build Command, lalu **Deploy**.
3. Alternatif tanpa GitHub: install Vercel CLI (`npm i -g vercel`), masuk ke folder tadi, jalankan `vercel --prod`.
4. Setiap kali `index.html` berubah, push ke GitHub (atau jalankan `vercel --prod` lagi) dan Vercel akan memperbarui situsnya.

Tidak perlu pengaturan CORS atau environment variable: web memanggil Apps Script langsung dari browser.

## Fitur
- **Sidebar**: Jadwal konten, Profil & akun, Ekspor ke spreadsheet (CSV), Impor dari spreadsheet, Ganti tema, Keluar akun. Di HP sidebar muncul lewat tombol menu.
- **Edit konten**: tombol **Edit** di tiap kartu membuka jendela untuk mengubah semua isian (judul, platform, tahap, deadline, kerja sama, fee, SOW, catatan).
- **Jenis kerja sama**: pilih *Tanpa kerja sama*, *Paid collab*, atau *Barter*. Untuk Paid collab muncul isian **Fee**, dan kartu menampilkan tulisan seperti "Paid collab · Fee Rp 2.500.000".
- **Status barang (barter)**: pilih *Belum dikirim*, *Masih dalam perjalanan*, atau *Sudah sampai*. Statusnya tampil sebagai label berwarna di kartu.
- **Rentang tanggal**: isi *Mulai* (opsional) dan *Deadline / sampai tanggal*. Kartu menampilkan "Periode 10 Nov – 20 Nov" bila keduanya diisi.
- **Status pembayaran (paid collab)**: pilih *Belum dibayar*, *DP masuk* (isi jumlah yang sudah dibayar), atau *Lunas*. Statusnya tampil sebagai label di kartu.
- **Info brand dan PIC**: isi nama brand, nama PIC, dan kontak PIC untuk konten kerja sama (paid atau barter). Kontak berupa nomor WA atau email otomatis menjadi tautan.
- **TikTok**: kalau platform TikTok dipilih, muncul isian **Keranjang kuning** (tanpa, belum dipasang, sudah dipasang) dan **Boost code**. Boost code di kartu punya tombol **Salin**.
- **Pendapatan**: menu baru di sidebar. Menampilkan total fee, yang sudah masuk, dan yang belum masuk per bulan (menurut bulan deadline), plus total yang belum masuk dari semua bulan. Status pembayaran bisa diubah langsung dari daftar.
- **Tampilan HP**: kolom papan tersusun ke bawah dan bisa dilipat dengan mengetuk judul kolom; form tersusun satu kolom.
- **Profil**: foto profil (otomatis dikecilkan), nama, dan tautan TikTok, Instagram, Threads, YouTube, X, serta website. Boleh isi `@username` atau link lengkap.
- **Ganti kata sandi**: di halaman Profil, wajib mengisi kata sandi lama. Perangkat lain otomatis keluar.
- CSV ekspor/impor sekarang punya kolom **Mulai**, **Kerja sama**, **Fee**, dan **Status kirim**. CSV lama tanpa kolom itu tetap bisa diimpor.

## Lupa kata sandi
- Di layar masuk ada tombol **Lupa kata sandi?**. Pengguna memasukkan email, menerima **kode 6 angka** (berlaku 15 menit), lalu membuat kata sandi baru.
- Email dikirim lewat `MailApp` dari akun Google pemilik script. Batas kirim sekitar 100 email per hari untuk Gmail biasa.
- **Kalau muncul "Email gagal dikirim"**, pesan di layar sekarang menyertakan penyebabnya. Penyebab paling umum: izin kirim email belum diberikan. Jalankan `izinkanEmail` (langkah 1.3), lalu buat versi deployment baru.
- Kode salah 5 kali membuat akun dikunci 10 menit dan kodenya hangus.
- **Cadangan manual**: hapus baris orang itu di tab **Users**. Orang itu mendaftar ulang dengan email yang sama dan data lamanya muncul lagi (terikat ke email). Lakukan hanya setelah yakin orangnya benar.

## Batasan yang perlu kamu tahu
- **Jaga sheet ini tetap privat.** Pemilik sheet bisa melihat email, foto, tautan sosmed, dan semua konten semua akun.
- Keamanannya lebih sederhana daripada Firebase: tidak ada verifikasi email saat daftar, dan hash memakai SHA-256 berulang (bukan bcrypt). Cocok untuk tim kecil.
- Apps Script agak lambat (1 sampai 3 detik per aksi) dan punya kuota harian. Maksimal 500 konten per akun.
- Foto profil disimpan langsung di sheet dalam ukuran kecil (sekitar 10 sampai 30 KB).
