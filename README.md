# Materi Generator

Mengubah naskah materi (JSON) menjadi deck PPTX dan/atau PDF sesuai `Template Desain Materi.pptx`.

## Cara pakai

```bash
python generator/buat_materi.py konten/tiu_penalaran_logis.json --format pptx
python generator/buat_materi.py konten/tiu_penalaran_logis.json --format pdf
python generator/buat_materi.py konten/tiu_penalaran_logis.json --format keduanya
```

Hasil masuk ke `output/`. Cek teks meluap dan render PNG per slide:

```bash
python generator/periksa.py "output/TIU - Penalaran Logis.pptx" --png output/pratinjau
```

Syarat: Windows dengan Microsoft PowerPoint (untuk PDF dan pemeriksaan) dan Python 3.10+. Pasang paket dengan:

```bash
pip install -r requirements-lokal.txt
```

`requirements.txt` hanya berisi paket untuk membuat PPTX dan aplikasi web (dipakai Vercel). `requirements-lokal.txt` menambahkan PyMuPDF (ekstraksi soal) dan pywin32 (ekspor PDF dan pemeriksaan lewat PowerPoint).

## Aplikasi web (Vercel)

`app.py` adalah aplikasi Flask: unggah naskah materi JSON, lalu unduh deck PPTX. PDF dan pemeriksaan tetap dibuat di Windows lokal karena butuh PowerPoint. Deck latihan soal (`buat_latihan.py`) juga tetap lokal, karena butuh folder gambar figur.

Jalankan di komputer sendiri:

```bash
python app.py --template "path/ke/Template Desain Materi.pptx"
```

Lalu buka http://127.0.0.1:5000. Tanpa `--template`, aplikasi memakai template di akar proyek.

### Deploy

1. Simpan `Template Desain Materi.pptx` (dan bila perlu `latar_cover.png`) di penyimpanan yang bisa diunduh lewat URL, misalnya Vercel Blob. Pakai penyimpanan privat dengan token bila tersedia. Kalau tidak ada, pakai URL acak yang tidak mudah ditebak. Jangan taruh berkas ini di repo.
2. Impor repo ini di Vercel. Flask terdeteksi otomatis dari `app.py`. Python mengikuti `.python-version` (3.12).
3. Isi **Environment Variables** di pengaturan proyek, lalu deploy ulang:

| Variabel | Wajib | Isi |
|---|---|---|
| `MATERI_TEMPLATE_URL` | ya | URL unduhan `Template Desain Materi.pptx` |
| `MATERI_SANDI` | ya* | Sandi akses. Browser meminta nama pengguna (bebas) dan sandi ini |
| `MATERI_COVER_URL` | tidak | URL foto latar sampul; tanpa ini sampul dibuat tanpa foto |
| `MATERI_ASET_TOKEN` | tidak | Dikirim sebagai `Authorization: Bearer <token>` saat mengunduh dua URL di atas |
| `MATERI_TANPA_SANDI` | tidak | Isi `1` hanya jika akses sudah dilindungi **Vercel Authentication** (Settings → Deployment Protection → All Deployments) |

\* Tanpa `MATERI_SANDI` dan tanpa `MATERI_TANPA_SANDI=1`, aplikasi di Vercel menolak semua akses. Ini supaya konten tidak terbuka untuk umum karena lupa disetel.

4. Buka `/api/status` untuk mengecek template dan foto sampul sudah terbaca.

Catatan:
- Vercel membatasi request dan response 4,5 MB. Naskah JSON jauh di bawah batas itu. Deck tanpa foto sampul sekitar 0,4 MB, jadi foto sampul harus cukup kecil (di bawah ±3,5 MB) agar deck tetap bisa diunduh.
- Aset hasil unduhan disimpan sementara di `/tmp` per instance. Setelah mengganti template, pakai URL baru (Vercel Blob membuat URL baru tiap unggah) supaya instance yang sedang jalan mengunduh ulang.
- `.vercelignore` mencegah template, PDF, dan konten soal ikut terunggah jika deploy lewat `vercel` CLI dari komputer ini.

## Font rumus di server

Pengukur teks butuh lebar glyph Cambria Math untuk menebak pindah baris. Font ini milik Microsoft dan tidak ada di server Linux, jadi repo hanya menyimpan tabel lebarnya di `generator/assets/metrik_cambria_math.json`, bukan file fontnya. Hasil ukurnya sama dengan pengukuran langsung dari font. Tabel dibuat ulang (di Windows, butuh `fonttools`) dengan:

```bash
python generator/buat_metrik_rumus.py
```

## Berkas yang tidak ikut di repo

Berkas berikut milik Cerebrum dan sengaja dikecualikan lewat `.gitignore`. Salin manual dari penyimpanan internal bila perlu:

| Berkas | Fungsi | Jika tidak ada |
|---|---|---|
| `Template Desain Materi.pptx` (akar proyek) | Dasar master, tema, dan font tertanam. Ukuran slide wajib 20 × 11,25 inci (16:9); tata letak generator mengikuti ukuran ini | Generator berhenti dengan pesan; wajib ada |
| `generator/assets/latar_cover.png` | Foto latar sampul | Sampul dibuat tanpa foto |
| `konten/latihan_*.json`, `konten/gambar/` | Soal dan pembahasan tryout hasil ekstraksi | Buat ulang dengan `ekstrak_soal.py` dari PDF sumber |
| `output/` | Hasil generate | Dibuat otomatis |

Font Poppins di `generator/assets/fonts/` berlisensi SIL Open Font License 1.1 (lihat `OFL.txt` di folder yang sama).

## Deck latihan soal dari PDF

Mengubah PDF "Latihan SKD CPNS" (soal + pembahasan) menjadi deck: satu slide soal, lalu slide jawaban dan pembahasan.

```bash
python generator/ekstrak_soal.py "path/ke/TIU Paket 1.pdf" --keluar konten/latihan_tiu_paket1.json --paket "Paket 1"
python generator/buat_latihan.py konten/latihan_tiu_paket1.json --format keduanya
```

- Rumus di PDF sumber berupa gambar. Ekstraktor menerjemahkannya lewat `generator/label_rumus.json` (452 rumus dari Paket 1 dan 2). Jika PDF baru memuat rumus yang belum dikenal, ekstraktor memberi peringatan; jalankan `python generator/ekstrak_soal.py a.pdf --lembar folder` untuk membuat lembar gambar bernomor, lalu tambahkan labelnya.
- Gambar figur disimpan di `konten/gambar/<nama naskah>/` pada 300 dpi.
- Naskah JSON hasil ekstraksi boleh disunting (salah ketik, kunci, pembahasan) sebelum deck dibuat ulang.

## Struktur naskah

| Bagian | Isi |
|---|---|
| `meta` | `judul`, `subtes`, `nama_file` |
| `peta_subtes` | `judul`, `subjudul`, `aktif` (indeks kartu yang disorot), `kartu[]` (`judul`, `deskripsi`, `ikon`) |
| `pendahuluan` | `paragraf[]`, `pemantik`, `ikon` |
| `tujuan` | `pengantar`, `butir[]` (`level`, `kko`, `teks`) |
| `materi[]` | `subbab`, `ikon`, `slide[]` |
| `penutup` | `judul`, `ringkasan[]` (`judul`, `teks`), `pesan` |

Tipe slide materi (`tipe`): `tabel`, `tabel_kebenaran`, `banding`, `notasi`, `kartu`, `alur`, `langkah`, `teks_ikon`.
Setiap slide boleh punya `catatan` dengan `jenis` `peringatan` atau `tips`. Contoh lengkap ada di `konten/tiu_penalaran_logis.json`.

Markup teks: `**tebal**` dan `$p → q$` untuk simbol/rumus (ditulis dengan Cambria Math karena Poppins tidak punya glyph logika).

Ikon: `globe`, `strategi`, `grafik`, `checklist`, `roket`, `peringatan`, `presenter`, `bingung`.

## Perbaikan terhadap template asli

- Font tebal dirujuk sebagai `Poppins` + bold, sehingga font tertanam benar-benar dipakai (sebelumnya jatuh ke Arial).
- Satu keluarga font (Poppins); Inter dan Arial Nova dihapus dari file.
- Palet brand dan Poppins dimasukkan ke tema PowerPoint.
- Teks rata kiri, posisi judul dan chip seragam di semua slide, kontras label bagian dinaikkan.
- Penomoran langkah otomatis dan benar, nomor halaman, navigasi bawah menandai bagian aktif dan bisa diklik (juga di PDF).
- Komponen baru untuk TIU: kotak notasi, tabel kebenaran, alur premis, catatan peringatan/tips, kartu banding.
