# Materi Generator

Mengubah naskah materi (JSON) menjadi deck PPTX dan/atau PDF dengan dua pilihan tampilan (16 × 11,25 inci):

- **Opsi B "Ceria & Membulat"** (bawaan): navy `13215A`, kuning `FBBB2B`, kartu putih membulat.
- **General**: mengikuti `Template Materi General.pptx` (foto sampul, halaman berbingkai navy/kuning, gradien kuning dan navy-abu).

Dokumen materi di Google Docs bisa dikonversi otomatis menjadi naskah, lalu ditinjau dan disunting tim akademik lewat dasbor web.

## Alur kerja

1. **Dokumen → naskah.** Dokumen materi (templat booklet SKD) diunduh sebagai `.docx` dan dikonversi oleh `generator/impor_katalog.py` menjadi `konten/*.json` beserta gambarnya di `konten/media/`.
2. **Tinjau dan sunting.** Di dasbor web, tim akademik melihat 15 materi TIU Verbal, mengambil tugas tinjauan, memberi status dan catatan, lalu membuka editor untuk mengubah isi slide dengan pratinjau langsung.
3. **Ekspor.** Unduh PPTX dari dasbor atau editor. PDF dibuat di Windows lokal (butuh PowerPoint).

## Cara pakai (baris perintah)

```bash
python generator/buat_materi.py konten/tiu_penalaran_logis.json --format pptx
python generator/buat_materi.py konten/tiu_penalaran_logis.json --format pdf
python generator/buat_materi.py konten/tiu_penalaran_logis.json --format keduanya
```

Tambahkan `--tema general` untuk tampilan General (berkas hasil diberi akhiran " (General)"):

```bash
python generator/buat_materi.py konten/tiu_penalaran_logis.json --tema general
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

## Situs unduhan statis (tanpa generator)

Folder `situs/` adalah halaman unduhan siap pakai: semua materi (15 materi TIU Verbal + Penalaran Logis) dalam dua tampilan, Opsi B dan General, masing-masing PPTX dan PDF, dengan gambar sampul. Pengunjung memasukkan sandi, memilih tampilan, lalu mengunduh; tidak ada server atau generator yang berjalan.

Bangun ulang setelah naskah atau template berubah (Windows + PowerPoint, karena PDF dan sampul diekspor PowerPoint):

```bash
python generator/buat_situs.py --template-opsi-b "path/Template Referensi Opsi B - Ceria & Membulat.pptx" --template-general "path/Template Materi General.pptx"
```

Pasang `requirements-lokal.txt` dulu: PyMuPDF menggabungkan gambar latar yang disimpan berulang oleh PowerPoint, sehingga PDF General turun dari sekitar 15 MB menjadi sekitar 2 MB tanpa mengubah tampilan.

Isi `situs/`: `index.html` (halaman), `data.js` (daftar materi dan ukuran berkas, dibuat skrip), `unduh/<tema>/` (PPTX dan PDF), `sampul/<tema>/` (gambar slide pertama), `font/` (Poppins), `middleware.ts` (kunci sandi), `vercel.json`.

Deploy di Vercel: buat proyek dari repo yang sama dengan Root Directory = `situs`, Framework Preset = Other, tanpa perintah build. Isi satu env variable, `MATERI_SANDI`, dengan sandi akses. `middleware.ts` meminta sandi itu (kotak login bawaan browser, nama pengguna bebas) sebelum halaman atau berkas apa pun dibuka; tanpa `MATERI_SANDI` situs menjawab 503. Vercel Authentication di Settings › Deployment Protection boleh dimatikan karena aksesnya sudah dijaga sandi. Server lokal (`python -m http.server`) tidak memakai sandi.

## Impor dokumen Google Docs

`konten/katalog.json` mendaftar 15 dokumen di kolom D sheet "Kerangka Course" (Mapping SKD 2027): kelompok, bab, baris sheet, dan ID dokumen. Untuk membuat ulang naskah setelah dokumen diperbarui:

1. Unduh dokumen sebagai Word (File › Download › Microsoft Word) ke satu folder, beri nama `<doc_id>.docx`.
2. Jalankan:

```bash
python generator/impor_katalog.py --docx folder_docx
```

Hanya entri dengan `"otomatis": true` yang ditimpa; `va-01` (Verbal Analogi) disusun manual. Satu dokumen bisa dikonversi langsung dengan `generator/ekstrak_gdoc.py`.

Yang dikonversi apa adanya: paragraf, subjudul, daftar, tabel (termasuk sel gabungan), gambar beserta keterangannya, persamaan Word (menjadi `$…$`), serta contoh soal (soal, opsi A–E, kunci, pembahasan). Yang tidak ada di dokumen dibiarkan kosong untuk dilengkapi di editor: pertanyaan pemantik, level Bloom (lingkaran menampilkan nomor bila kosong), dan teks ringkasan penutup. Dasbor menandai materi yang masih perlu dilengkapi.

Catatan: tautan "05 Logika Pernyataan" di sheet (baris 39) kurang satu huruf sehingga tidak bisa dibuka; katalog memakai ID yang benar.

## Aplikasi web (Vercel)

`app.py` adalah aplikasi Flask:

| Halaman | Isi |
|---|---|
| `/` | Dasbor: daftar materi, tautan dokumen sumber, status tinjauan (Belum ditinjau / Sedang ditinjau / Perlu revisi / Disetujui), peninjau, catatan, unduh PPTX |
| `/materi/<id>` | Editor: pratinjau slide dari deck yang sama dengan hasil unduhan, form isi per slide, simpan draf, buang draf, unduh PPTX |
| `/unggah` | Buat PPTX dari berkas naskah JSON sendiri |

Draf suntingan tidak mengubah berkas di repo. Unduhan memakai draf bila ada. Bila dua orang menyunting materi yang sama, penyimpanan kedua ditolak dengan pesan agar suntingan pertama tidak tertimpa.

Pilihan **Tampilan PPT** (Opsi B atau General) ada di dasbor, editor, dan halaman unggah; pilihan disimpan di browser dan dipakai untuk pratinjau serta unduhan.

Jalankan di komputer sendiri (draf dan status disimpan di `data/`, diabaikan git):

```bash
python app.py --template "path/ke/Template Referensi Opsi B - Ceria & Membulat.pptx" --template-general "path/ke/Template Materi General.pptx"
```

Lalu buka http://127.0.0.1:5000. Tanpa `--template`, aplikasi memakai template di akar proyek.

### Deploy

1. Simpan template Opsi B di penyimpanan yang bisa diunduh lewat URL, misalnya Vercel Blob. Pakai penyimpanan privat dengan token bila tersedia; kalau tidak ada, pakai URL acak yang tidak mudah ditebak. Jangan taruh berkas ini di repo.
2. Impor repo ini di Vercel. Flask terdeteksi otomatis dari `app.py`. Python mengikuti `.python-version` (3.12).
3. Di proyek Vercel, buka **Storage › Upstash for Redis**, buat database gratis, lalu sambungkan ke proyek. Variabel `KV_REST_API_URL` dan `KV_REST_API_TOKEN` terisi otomatis.
4. Isi **Environment Variables**, lalu deploy ulang:

| Variabel | Wajib | Isi |
|---|---|---|
| `MATERI_TEMPLATE_URL` | ya | URL unduhan `Template Referensi Opsi B - Ceria & Membulat.pptx` |
| `MATERI_TEMPLATE_GENERAL_URL` | tidak | URL unduhan `Template Materi General.pptx`; tanpa ini pilihan General menampilkan pesan |
| `MATERI_SANDI` | ya* | Sandi akses. Browser meminta nama pengguna (bebas) dan sandi ini |
| `MATERI_ASET_TOKEN` | tidak | Dikirim sebagai `Authorization: Bearer <token>` saat mengunduh template |
| `MATERI_TANPA_SANDI` | tidak | Isi `1` hanya jika akses sudah dilindungi **Vercel Authentication** (Settings › Deployment Protection › All Deployments) |
| `KV_REST_API_URL`, `KV_REST_API_TOKEN` | ya** | Diisi otomatis oleh integrasi Upstash (nama `UPSTASH_REDIS_REST_URL`/`_TOKEN` juga diterima) |

\* Tanpa `MATERI_SANDI` dan tanpa `MATERI_TANPA_SANDI=1`, aplikasi di Vercel menolak semua akses agar konten tidak terbuka untuk umum karena lupa disetel.
\** Tanpa Upstash, aplikasi di Vercel berjalan dalam mode baca saja: dasbor, editor, dan unduhan jalan, tetapi status, catatan, dan draf tidak bisa disimpan.

5. Buka `/api/status` untuk mengecek template dan penyimpanan sudah terbaca.

Catatan:
- Vercel membatasi request dan response 4,5 MB. Deck terbesar saat ini (Logika Kuantor, banyak gambar) sekitar 1,8 MB.
- Template hasil unduhan disimpan sementara di `/tmp` per instance. Setelah mengganti template, pakai URL baru supaya instance yang sedang jalan mengunduh ulang.
- `.vercelignore` mencegah template, PDF, dan konten soal tryout ikut terunggah jika deploy lewat `vercel` CLI dari komputer ini.

## Struktur naskah

| Bagian | Isi |
|---|---|
| `meta` | `judul`, `subjudul` (opsional), `subtes`, `nama_file`, `sumber` |
| `peta_subtes` | `judul`, `subjudul`, `aktif` (indeks kartu yang disorot), `kartu[]` (`judul`, `deskripsi`, `ikon`) |
| `pendahuluan` | `paragraf[]`, `pemantik` (opsional), `ikon` atau `gambar` (path di `konten/`) |
| `tujuan` | `pengantar`, `butir[]` (`level` opsional, `kko`, `teks`) |
| `materi[]` | `subbab`, `ikon`, `slide[]` |
| `penutup` | `judul`, `ringkasan[]` (`judul`, `teks` opsional), `pesan` |

Tipe slide materi (`tipe`):

- Tata letak khusus: `tabel`, `tabel_kebenaran`, `banding`, `notasi`, `kartu`, `alur`, `langkah`, `teks_ikon`. Setiap slide boleh punya `catatan` dengan `jenis` `peringatan` atau `tips`. Contoh lengkap ada di `konten/tiu_penalaran_logis.json` dan `konten/tiu_verbal_analogi.json`.
- `blok`: isi umum hasil konversi dokumen, berupa daftar `blok[]` dengan `jenis` `paragraf`, `subjudul`, `sorot`, `poin`, `gambar`, `tabel`, atau `catatan`. Bila slide penuh, generator membuat slide lanjutan; slide yang dimulai tepat di subjudul memakai subjudul itu sebagai judul.
- `contoh_soal`: `judul`, `label` (opsional), `soal[]` (blok), `opsi` (`A`–`E`), `kunci`, `pembahasan[]` (blok). Menghasilkan slide soal dan slide pembahasan yang saling bertaut. Opsi yang semuanya pendek disusun mendatar.

Markup teks: `**tebal**` dan `$p → q$` untuk simbol/rumus (ditulis dengan Cambria Math karena Poppins tidak punya glyph logika).

Ikon: `globe`, `strategi`, `grafik`, `checklist`, `roket`, `peringatan`, `presenter`, `bingung`. Varian navy dan putih untuk desain Opsi B ada di `generator/assets/ikon/` (dibuat ulang dengan `python generator/buat_ikon.py`).

## Desain Opsi B

- Latar slide isi `F4F6FB`, kartu putih membulat, kartu sorotan navy, catatan kuning muda dengan lingkaran simbol navy (`!` peringatan, `i` tips, `?` pemantik).
- Chip bagian berupa pil navy dengan ikon dalam lingkaran kuning; navigasi bawah pil navy dengan sorotan kuning dan bisa diklik (juga di PDF); nomor halaman dalam lingkaran kuning.
- Sampul dan penutup memakai bentuk bulat bertumpuk navy-kuning; peta subtes bergaya daftar isi bernomor; tujuan berupa kartu dengan level Bloom.
- Palet Opsi B dan Poppins dimasukkan ke tema PowerPoint; Poppins regular dan bold tetap tertanam.
- Generator menolak template dengan ukuran slide selain 16 × 11,25 inci agar tata letak tidak bergeser diam-diam.

## Tampilan General

`generator/tema_general.py` mengambil grafis langsung dari slide contoh `Template Materi General.pptx` saat deck dibuat, sehingga foto dan ilustrasi perusahaan tidak perlu disalin ke repo:

- Sampul, Pendahuluan (bingkai navy), Tujuan, Penutup (bingkai kuning), dan Terima kasih disalin dari slide contohnya; teks contoh diganti isi naskah. Slide contoh dikenali dari teks besarnya (MATERI BESAR, PENDAHULUAN, TUJUAN PEMBELAJARAN, PENUTUP, TERIMA KASIH) dan slide isi bertabel, bukan dari nomor urutnya.
- Slide isi memakai chip penanda dan hiasan samar (lingkaran kuning, gelombang garis) dari template; komponennya sama dengan Opsi B, dengan token warna General: kartu gradien kuning muda ke putih, kartu gelap gradien navy ke abu-abu, catatan kuning.
- Navigasi digambar ulang mengikuti gaya template (pil bergaris navy, bagian aktif navy) agar bagian aktif benar dan tetap bisa diklik; di slide berbingkai warnanya putih.
- Sampul memakai `meta.materi_besar` (bila ada) atau subtes sebagai teks besar, dan `meta.tutor` (opsional) untuk nama tutor.
- Gambar template yang besar diperkecil saat deck dibuat (hiasan gelombang 3,3 MB menjadi ±0,2 MB) sehingga deck General sekitar 1,3–2,8 MB.
- Deck latihan soal (`buat_latihan.py`) tetap memakai tampilan Opsi B.

## Font rumus di server

Pengukur teks butuh lebar glyph Cambria Math untuk menebak pindah baris. Font ini milik Microsoft dan tidak ada di server Linux, jadi repo hanya menyimpan tabel lebarnya di `generator/assets/metrik_cambria_math.json`, bukan file fontnya. Hasil ukurnya sama dengan pengukuran langsung dari font. Tabel dibuat ulang (di Windows, butuh `fonttools`) dengan:

```bash
python generator/buat_metrik_rumus.py
```

## Berkas yang tidak ikut di repo

Berkas berikut milik Cerebrum dan sengaja dikecualikan lewat `.gitignore`. Salin manual dari penyimpanan internal bila perlu:

| Berkas | Fungsi | Jika tidak ada |
|---|---|---|
| `Template Referensi Opsi B - Ceria & Membulat.pptx` (akar proyek) | Dasar master, tema, dan font tertanam. Ukuran slide wajib 16 × 11,25 inci | Generator berhenti dengan pesan; wajib ada |
| `Template Materi General.pptx` (akar proyek) | Sumber grafis tampilan General | Hanya pilihan General yang tidak bisa dipakai |
| `konten/latihan_*.json`, `konten/gambar/` | Soal dan pembahasan tryout hasil ekstraksi | Buat ulang dengan `ekstrak_soal.py` dari PDF sumber |
| `output/`, `data/` | Hasil generate; draf dan status dasbor lokal | Dibuat otomatis |

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
