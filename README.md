# Materi Generator

Mengubah naskah materi (JSON) menjadi deck PPTX dan/atau PDF dengan dua pilihan tampilan (16 × 11,25 inci):

- **Opsi B "Ceria & Membulat"** (bawaan): navy `13215A`, kuning `FBBB2B`, kartu putih membulat.
- **General**: mengikuti `Template Materi General.pptx` (foto sampul, halaman berbingkai navy/kuning, gradien kuning dan navy-abu).

Dokumen materi di Google Docs bisa dikonversi otomatis menjadi naskah, lalu ditinjau dan disunting tim akademik lewat dasbor web.

## Alur kerja

1. **Dokumen → naskah.** Dokumen materi (templat booklet SKD) diunduh sebagai `.docx` dan dikonversi oleh `generator/impor_katalog.py` menjadi `konten/*.json` beserta gambarnya di `konten/media/`.
2. **Tinjau dan sunting.** Di dasbor web, tim akademik melihat 76 deck TIU (Verbal, Numerik, dan Figural; materi dan contoh soal), mengambil tugas tinjauan, memberi status dan catatan, lalu membuka editor untuk mengubah isi slide dengan pratinjau langsung.
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

Folder `situs/` adalah satu halaman unduhan untuk semua aplikasi. Tab di bagian atas memilih aplikasinya: **JadiASN · TIU** (77 deck: 28 Verbal, 42 Numerik termasuk Data & Deret, dan 7 Figural), **JadiOJK · PCAM OJK** (57 deck TKU PCAM OJK: 29 materi dan 28 contoh soal), dan **Cerebrum · UTBK** (243 deck UTBK-SNBT per Bab: 122 materi dan 121 contoh soal). Warna slide dan halaman mengikuti palet aplikasi itu. Setiap materi berbentuk General dengan sampul foto, dalam PPTX dan PDF. Pengunjung memasukkan sandi, memilih aplikasi, lalu mengunduh; tidak ada server atau generator yang berjalan. Tab terakhir yang dipilih diingat, dan tautan bisa langsung menuju satu aplikasi (`…/#ojk`). Slicer **Semua / Materi / Contoh soal** menyaring jenis deck bersama saringan kelompok dan pencarian; pilihannya ikut di alamat (`…/#tiu/soal`, `…/#ojk/materi`). Jenis deck ditentukan `jenis_deck` di `generator/buat_materi.py` dan disimpan sebagai `"jenis"` di `data.json`: deck dengan `bagian: "soal"`, berjudul "Contoh Soal …", atau yang seluruh isinya soal (Figural Serial dan Matriks) dihitung contoh soal.

Palet platform ada di `generator/palet.py` (warna dari proyek ICON Aplikasi generator, `catalog/brands.json`). Palet mengganti navy dan kuning template di semua bentuk, SVG, dan PNG berwarna brand; foto tidak diubah. Palet `cerebrum` memakai merah bata dan kuning untuk bidang dan judul, dan biru tua (`teks_biasa`) untuk teks isi yang tidak tebal. Generator biasa juga bisa memakainya: `python generator/buat_materi.py naskah.json --tema general --palet jadiasn`.

Bangun ulang setelah naskah atau template berubah (Windows + PowerPoint, karena PDF dan sampul diekspor PowerPoint):

```bash
python generator/buat_situs.py --template-general "path/Template Materi General.pptx"
python generator/buat_situs.py --katalog konten/ojk/katalog.json --template-general "path/Template Materi General.pptx"
python generator/buat_situs.py --katalog konten/utbk/katalog.json --template-general "path/Template Materi General.pptx"
```

Setiap perintah membangun satu aplikasi ke `situs/<id>/` (id dari `"aplikasi"` di katalog) dan menambahkan tabnya ke `situs/aplikasi.json`; aplikasi lain tidak tersentuh. Palet diambil dari `"palet"` di katalog. Bawaannya `--tema general` (satu tampilan, jadi halaman tidak menampilkan pilihan). Pakai `--palet tanpa` untuk warna asli template, atau tambahkan `opsi-b` ke `--tema` (dengan `--template-opsi-b "path/Template Referensi Opsi B - Ceria & Membulat.pptx"`) untuk menyediakan Opsi B juga.

Ukuran: TIU sekitar 198 MB, OJK sekitar 142 MB, dan UTBK sekitar 561 MB (243 deck per Bab), jadi seluruh situs sekitar 900 MB. Menurut dokumentasi Vercel (dicek November 2026), batas 100 MB di paket Hobby hanya untuk unggahan lewat CLI; untuk deploy dari GitHub tidak ada batas ukuran tertulis. Yang perlu dijaga adalah kuota 10 GB penyimpanan deployment per tim Hobby: bila terlampaui, deploy baru bisa ditolak sampai deployment lama dihapus.

Untuk memperbarui daftar saja (kelompok, tautan dokumen, jenis deck) tanpa membangun deck, jalankan perintah yang sama dengan `--hanya-data`; PowerPoint dan template tidak dibutuhkan.

Pasang `requirements-lokal.txt` dulu: PyMuPDF menggabungkan gambar latar yang disimpan berulang oleh PowerPoint dan menyimpan hiasan abu-abu sebagai gambar satu kanal, sehingga PDF General turun dari sekitar 15 MB menjadi sekitar 1,5–2 MB tanpa mengubah tampilan.

Isi `situs/`: `index.html` (halaman), `aplikasi.json` (daftar tab), `<id>/data.json` (daftar materi, ukuran berkas, dan warna halaman aplikasi itu; dibuat skrip), `<id>/unduh/general/` (PPTX dan PDF), `<id>/sampul/general/` (gambar slide pertama), `font/` (Poppins), `middleware.ts` (kunci sandi), `vercel.json`.

Deploy di Vercel: buat proyek dari repo yang sama dengan Root Directory = `situs`, Framework Preset = Other, tanpa perintah build. Isi satu env variable, `MATERI_SANDI`, dengan sandi akses. `middleware.ts` meminta sandi itu (kotak login bawaan browser, nama pengguna bebas) sebelum halaman atau berkas apa pun dibuka; tanpa `MATERI_SANDI` situs menjawab 503. Vercel Authentication di Settings › Deployment Protection boleh dimatikan karena aksesnya sudah dijaga sandi. Server lokal (`python -m http.server`) tidak memakai sandi.

## Impor dokumen Google Docs

`konten/katalog.json` mendaftar 46 dokumen TIU di kolom D sheet "Kerangka Course" (Mapping SKD 2027) sebagai 75 deck: kelompok, bab, baris sheet, dan ID dokumen. Untuk membuat ulang naskah setelah dokumen diperbarui:

1. Unduh dokumen sebagai Word (File › Download › Microsoft Word) ke satu folder, beri nama `<doc_id>.docx`.
2. Jalankan:

```bash
python generator/impor_katalog.py --docx folder_docx
```

Soal yang ditulis di dokumen materi selalu dipisah menjadi deck sendiri, baik yang punya baris "Contoh Soal" di sheet (mis. Silogisme 02–06) maupun yang tidak (mis. Rasio; `baris_sheet`-nya memakai baris materi). Contoh soal yang ditulis di tengah Materi (heading "Contoh Soal — Metode Substitusi" di Persamaan Linear) ikut dipisah (`_soal_sisipan`). Entri `"bagian": "soal"` (mis. `vs-02-soal`) memakai `doc_id` yang sama, dan importer memisahkan satu dokumen menjadi dua naskah: materi (`tiu_vs02_logika_kondisional.json`) dan contoh soal (`tiu_vs02_contoh_soal.json`). Deck contoh soal hanya memuat sampul, peta materi, soal, dan pembahasan; importer menandai soal yang tidak punya kunci ("TANPA KUNCI").

Hanya entri dengan `"otomatis": true` yang ditimpa. Satu dokumen bisa dikonversi langsung dengan `generator/ekstrak_gdoc.py`.

### Yang masuk PPT: wajib, pelengkap, dan tidak masuk

Highlight berarti *wajib* masuk, bukan satu-satunya yang boleh masuk. Ada tiga tingkat:

1. **Wajib**: bagian yang diberi sorotan (highlight) kuning atau berlabel `[WAJIB PPT]` di dokumen sumber (rinciannya di bawah).
2. **Pelengkap**: tambahan di luar sorotan supaya isi wajib tidak pincang (`pelengkap()` di `generator/ekstrak_gdoc.py`): gambar dan tabel di bawah judul yang sudah punya isi masuk, langkah atau butir daftar yang bolong, dan pilihan per dokumen di `konten/pelengkap.json` / `konten/ojk/pelengkap.json` / `konten/utbk/pelengkap.json` (inti subbab tanpa sorotan, definisi, tujuan dan pesan penutup yang kosong). Hapus satu butir di berkas itu untuk mencoret tambahannya. Katalog `"pelengkap": "minimal"` (Pengenalan OJK, Perencanaan Keuangan, Dana Pensiun) membatasi tabel ke yang diapit isi wajib.
3. **Tidak masuk**: `[PENGAYAAN]` (kecuali di bawah label "semua contoh soal wajib"), isi di luar label "hanya foto dan profil", soal tanpa sorotan, penjelasan lanjutan di subbab yang sudah punya sorotan, rujukan, dan catatan penyusun (`CATATAN_PRODUKSI`, mis. "PPT Materi OLD mencantumkan D").

Setiap tambahan ditulis di **catatan pembicara** slide-nya ("TAMBAHAN DI LUAR HIGHLIGHT") dan disimpan di naskah (`tambahan` pada blok) supaya tim bisa meninjau dan mencoretnya.

Aturan bagian wajib: Konverter membacanya dari berkas Word (teks, gambar, rumus, dan sel tabel punya tanda sorotan sendiri); lihat `tandai()` dan `saring()` di `generator/ekstrak_gdoc.py`.

- **Paragraf** yang disorot sebagian: dipakai kalimat yang minimal separuh hurufnya disorot. **Tabel**: header dan baris yang disorot.
- **Label** di judul: `[WAJIB PPT]` berarti seluruh isi bagian itu masuk; `[... yang dihighlight wajib ...]` dan `[PENGAYAAN]` berarti sorotan saja; `[semua contoh soal wajib ada di PPT]` memasukkan semua soal, termasuk yang berlabel `[PENGAYAAN]`; `[Sertakan gambar pada slide]` memasukkan gambarnya; `[... hanya foto dan profil ...]` memasukkan tabel foto sebagai kartu profil.
- **Judul** subbab dan subjudul tetap ada bila isinya masuk.
- **Gambar** ikut teksnya, karena sorotan di Google Docs hanya tampak pada teks sehingga gambar jarang ikut disorot (lihat `_ikutkan_gambar()`): gambar masuk bila disorot, keterangannya disorot, diapit paragraf yang masuk, atau, di dokumen yang penulisnya tidak pernah menyorot gambar di bagian materi (mis. Fill Shape dan Drag the Line), bila ada isi yang masuk di bawah judul yang sama. Di dokumen yang sebagian gambar materinya disorot (Verbal, Numerik), gambar lain tidak dianggap wajib. Gambar di dokumen semacam itu tetap masuk sebagai pelengkap bila subbabnya masuk.
- **Soal** masuk bila minimal separuh pertanyaan dan opsinya disorot. Pertanyaan, opsi, dan kuncinya utuh; pembahasannya hanya bagian yang disorot (jawaban uraian tetap ikut). Bila yang ditandai hanya judul soalnya (Badges), soal dan pembahasannya masuk utuh.
- **Pendahuluan, tujuan, dan pesan penutup** yang tidak ditandai dilewati (`meta.lewati_kosong`); ringkasan penutup tetap berisi judul subbab.
- **Teks merah**: Badges tidak memakai highlight; penandanya teks merah (`"penanda": "merah"` di katalog), dibaca sama seperti highlight kuning. Di dokumen lain warna teks bukan penanda; Perkalian dan Pembagian Setara memakai `"coret": true` supaya coretan merah (`<s>…</s>`) ikut. Sorotan kuning yang dipasang sebagai latar paragraf juga dibaca.
- Dokumen yang tidak punya tanda sama sekali dikonversi utuh (saat ini tidak ada). Figural (PDF) dan Penalaran Logis tidak memakai aturan ini.

Yang ditandai dikonversi apa adanya: paragraf, subjudul, daftar, tabel (termasuk sel gabungan), gambar beserta keterangannya, persamaan Word (menjadi `$…$`, lewat `generator/omml.py`: pecahan `a/b`, pangkat `x^{2}`, indeks `U_{n}`, akar `√x`), serta contoh soal (soal, opsi A–E, kunci, pembahasan). Yang tidak ada di dokumen dibiarkan kosong untuk dilengkapi di editor: pertanyaan pemantik, level Bloom (lingkaran menampilkan nomor bila kosong), dan teks ringkasan penutup. Dasbor menandai materi yang masih perlu dilengkapi.

Catatan: tautan "05 Logika Pernyataan" di sheet (baris 39) kurang satu huruf sehingga tidak bisa dibuka; katalog memakai ID yang benar.

### Koreksi QC (`koreksi.json`)

Koreksi tim QC yang belum ada di dokumen sumber ditulis di `konten/koreksi.json` / `konten/ojk/koreksi.json` / `konten/utbk/koreksi.json` per ID materi, dan diterapkan setiap kali impor (`koreksi()` di `generator/ekstrak_gdoc.py`): `{"cari": "√(3/2)", "ganti": "√3/2", "asal": "QC …"}` untuk teks dan sel tabel, atau `{"hapus_gambar_di": "judul subbab"}`. Bila dokumennya sudah diperbaiki, importer menampilkan "PERINGATAN koreksi tidak ditemukan"; hapus butirnya.

### Figural (PDF)

Dokumen Figural di sheet (baris 76–82) berupa PDF hasil ekspor Canva, bukan Google Docs. Entrinya di katalog punya `"url"` (tautan Drive) dan `"otomatis": false`, dan naskahnya dibuat skrip tersendiri (butuh PyMuPDF dan Pillow):

```bash
python generator/ekstrak_figural.py --pdf folder_pdf
```

`folder_pdf` berisi `<doc_id>.pdf`. Teks dan kunci diambil dari PDF; gambar soal dan pembahasan dipotong dari halaman, jadi opsi A–E ada di dalam gambar dan `opsi` di naskah kosong. Letak potongan mengikuti susunan halaman PDF saat ini; bila PDF diganti, periksa lagi hasil potongannya. PDF tidak punya pendahuluan, tujuan, dan penutup, sehingga naskahnya diberi `"lewati_kosong": true` dan slide itu tidak dibuat sampai isinya dilengkapi.

### Materi PCAM OJK (JadiOJK)

`konten/ojk/katalog.json` mendaftar 28 dokumen dan 2 deck Canva di kolom File sheet "Kerangka Course" (Mapping Tahap 2 PCAM OJK 2026) sebagai 57 deck (contoh soal setiap bab dipisah menjadi deck sendiri): subtes TKU (Penalaran Induktif, Penalaran Deduktif, Bahasa Inggris, Pengetahuan Sektor Jasa Keuangan), materi, bab, dan ID dokumen. Katalog ini juga menyimpan pengaturan deck: subtes di sampul ("Tes Kemampuan Umum"), peta subtes TKU, awalan nama berkas ("OJK TKU"), merek situs, dan palet `jadiojk`. Naskah dan gambar ada di `konten/ojk/`.

```bash
python generator/impor_katalog.py --katalog konten/ojk/katalog.json --docx folder_docx_ojk
python generator/buat_situs.py --katalog konten/ojk/katalog.json --template-general "path/Template Materi General.pptx"
```

Perintah kedua membangun materi OJK ke `situs/ojk/` dan menambahkan tab "JadiOJK · PCAM OJK" di situs unduhan yang sama dengan TIU (warna halaman JadiOJK).

Materi Bahasa Inggris (baris 16–20) memakai `"miring": true` di katalog supaya cetak miring dokumen ikut (`<i>…</i>` di naskah), karena soal Sentence Improvement menanyakan "the italicised part". Reading Comprehension (baris 21–22) berupa deck Canva, jadi naskahnya disusun manual (`"otomatis": false`, `"url"` tautan Canva); latihan di deck materinya dipindah ke deck contoh soal. Soal bacaan diberi `"bacaan": true` supaya teks, pertanyaan, dan opsi diusahakan muat di satu slide (huruf diperkecil sampai 16 pt).

Konverter mengenali format dokumen OJK: judul bagian bernomor ("1. PENDAHULUAN", "C. Materi Pembelajaran"), catatan penulis dalam kurung siku ("[WAJIB PPT]", "[PENGAYAAN]") yang dibuang dari teks (dan menentukan isi yang masuk, lihat di atas), bagian yang dilewati (Kata Pengantar, Capaian Pembelajaran, Daftar Rujukan), soal per Heading 2 ("4.1 …", "Contoh Soal 1 — …"), dan kunci "Jawaban C. …". Tabel yang punya kolom foto (mis. Profil Anggota Dewan Komisioner OJK) menjadi kartu profil berisi foto, nama, dan jabatan, sesuai catatan "[yang harus ada di PPT hanya foto dan profil]"; kolom uraian panjang tidak ikut. Tabel, daftar butir, dan kartu profil yang lebih tinggi dari satu halaman dipecah ke halaman berikutnya.

"Pengenalan OJK dan Waspada Investasi" (13 MB) terlalu besar untuk diekspor lewat Google Drive API; unduh manual sebagai Word, simpan sebagai `<doc_id>.docx`, lalu jalankan impor dengan `--id sjk-01`. Untuk membangun ulang satu materi saja di situs (tanpa membangun ulang yang lain), tambahkan `--id sjk-01` ke perintah `buat_situs.py`.

### Materi UTBK (Cerebrum)

`konten/utbk/katalog.json` mendaftar 46 dokumen di kolom File sheet "Kerangka Course" (Mapping UTBK 2027) sebagai 243 deck, **satu deck materi dan satu deck contoh soal per baris Bab** di sheet (mis. Penalaran Verbal menjadi Generalisasi, Analogi, Hubungan Kausalitas, Kesesuaian Pernyataan, dan Evaluasi Argumen). Kelompoknya mengikuti sheet (PU, PPU-PBM, PK-PM, LBI, LBE); nama subtes di sampul diambil dari `"nama_subtes"`, dan nama Bab menjadi judul deck. Baris 10–15 (Silogisme dan Logika) belum punya dokumen.

21 dokumen mencakup beberapa Bab. Dokumen itu ada di `"sumber"` katalog dan dibaca sekali, lalu `generator/pisah_bab.py` membaginya ke entri `materi` yang `"sumber"`-nya sama (`pu-01-1` … `pu-01-5`, contoh soalnya `pu-01-1-soal` …):

- Subbab yang namanya sama dengan Bab masuk deck Bab itu (`"bab"` memakai nama subbab di dokumen; `"bab_sheet"` mencatat nama di sheet bila berbeda, mis. "Analogi Kata/Frasa"). Subbab pengantar yang bukan Bab ("Penalaran Induktif", "Morfologi") ikut di awal setiap deck Bab sesudahnya.
- Rangkuman: butir, paragraf, atau baris tabel yang diawali nama Bab masuk deck Bab itu; sisanya masuk semua deck.
- Tujuan pembelajaran: butir berawalan nama Bab ("Generalisasi: membedakan …; menyimpulkan …") dipecah per kata kerja; tanpa awalan, butir dibagi rata menurut urutan Bab.
- Pendahuluan dan pesan penutup tetap utuh karena teksnya membahas seluruh materi. Ringkasan penutup berisi subbab deck, atau judul bagian di dalam Bab.
- Contoh soal mengikuti label subbabnya. Soal tanpa label dicocokkan dari nama Bab di judulnya, atau dari `"nomor_soal"` di entri katalog (Aritmatika Sosial: soal 1–2 Aritmatika Sosial, 3–4 Harga, dan seterusnya). Nomor soal diurutkan ulang per deck. Bab tanpa soal (Statistika) tidak punya deck contoh soal.

`pelengkap.json` dan `koreksi.json` memakai ID dokumen (`pk-05`), bukan ID Bab.

```bash
python generator/impor_katalog.py --katalog konten/utbk/katalog.json --docx folder_docx_utbk
python generator/buat_situs.py --katalog konten/utbk/katalog.json --template-general "path/Template Materi General.pptx"
```

Dokumen UTBK ditulis beberapa tim dengan kerangka berbeda, jadi katalognya memakai `"kerangka": "utbk"` dan setiap dokumen disusun ulang dulu oleh `generator/rapikan_utbk.py`: subbab bernomor di Heading 1/2/3, contoh soal yang ditulis di tengah subbab (dikumpulkan ke deck contoh soal dengan nama subbabnya di samping judul), soal berupa heading, paragraf tebal, atau daftar bernomor, opsi dalam satu paragraf, bacaan untuk beberapa soal ("Teks untuk Soal 1–3" menjadi slide sendiri), tabel tujuan (IPK), ganti baris di dalam paragraf, dan nomor daftar otomatis.

Penanda wajib tiap kelompok berbeda:

- **PU dan Bilangan dan Aritmetika**: teks merah (`"penanda": "merah"`).
- **PPU, PBM, dan PK-PM lainnya**: highlight kuning dan label `[WAJIB PPT]`; label `[semua contoh soal wajib ada di PPT]` memasukkan semua soal.
- **LBI dan LBE**: latar hijau muda (`"penanda": "hijau"`); tidak ada keterangan arti warnanya di dokumen. LBE memberi label "(SEMUA SOAL HARUS ADA DI PPT)". Catatan `[WAJIB PPT]` di sela contoh soal ikut deck contoh soal di letaknya. Soal LBI tidak ditandai sama sekali; soalnya tetap dimasukkan dan dicatat sebagai tambahan ("contoh soal tanpa tanda").

Tujuan pembelajaran atau penutup yang sama sekali tidak ditandai diisi dari dokumennya dan dicatat sebagai tambahan (`isi_bagian_kosong`). Rumus PK-PM yang di dokumen berupa gambar satu baris (16 px) diketik ulang di `"rumus_gambar"` katalog (nama berkas gambar → rumus), karena gambarnya buram di slide; bila gambar di dokumen diganti, sesuaikan pemetaannya. Gambar lain diperbesar 1,5 kali dari ukuran dokumennya.

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
| `meta` | `judul`, `subjudul` (opsional), `subtes`, `nama_file`, `sumber`, `bagian` (`"soal"` untuk deck contoh soal), `lewati_kosong` (opsional; lewati slide pendahuluan, tujuan, dan penutup yang masih kosong) |
| `peta_subtes` | `judul`, `subjudul`, `aktif` (indeks kartu yang disorot), `kartu[]` (`judul`, `deskripsi`, `ikon`) |
| `pendahuluan` | `paragraf[]`, `pemantik` (opsional), `ikon` atau `gambar` (path di `konten/`) |
| `tujuan` | `pengantar`, `butir[]` (`level` opsional, `kko`, `teks`) |
| `materi[]` | `subbab`, `ikon`, `slide[]` |
| `penutup` | `judul`, `ringkasan[]` (`judul`, `teks` opsional), `pesan` |

Tipe slide materi (`tipe`):

- Tata letak khusus: `tabel`, `tabel_kebenaran`, `banding`, `notasi`, `kartu`, `alur`, `langkah`, `teks_ikon`. Setiap slide boleh punya `catatan` dengan `jenis` `peringatan` atau `tips`. Contoh lengkap ada di `konten/tiu_penalaran_logis.json` dan `konten/tiu_verbal_analogi.json`.
- `blok`: isi umum hasil konversi dokumen, berupa daftar `blok[]` dengan `jenis` `paragraf`, `subjudul`, `sorot`, `poin`, `gambar`, `tabel`, atau `catatan`. `subjudul` boleh punya `tingkat` (1 = Heading di dokumen, 2 = paragraf tebal). Slide `blok` yang berurutan dalam satu subbab mengalir bersama; lihat "Aturan tata letak slide blok".
- `contoh_soal`: `judul`, `label` (opsional), `soal[]` (blok), `opsi` (`A`–`E`), `kunci`, `pembahasan[]` (blok). Menghasilkan slide soal dan slide pembahasan yang saling bertaut. Opsi yang semuanya pendek disusun mendatar.

### Aturan tata letak slide blok

Dipakai di semua deck (`_slide_blok`, `_rencana`, dan `_halaman` di `generator/buat_materi.py`; `GAYA_JUDUL` dan `alir_teks` di `generator/desain.py`):

- **Ukuran huruf mengikuti tingkat judul di dokumen, bukan letaknya di slide.** Nama subbab di pita atas; judul bagian (tingkat 0) 26 pt; subjudul tingkat 1 22 pt; subjudul tingkat 2 dan isi 20 pt. Subjudul yang setingkat di dokumen selalu sama besar, walau jatuh di awal slide. Deck yang tidak punya tingkat 1 menaikkan semua tingkat 2 menjadi tingkat 1.
- **Judul tidak ditulis dua kali.** Judul slide yang sama dengan nama subbab hanya tampil di pita. Di halaman sambungan, judul yang isinya masih berlanjut diulang dengan ukuran yang sama, tanpa kata "(lanjutan)".
- **Yang menjadi judul** hanya judul. Kalimat langkah ("Langkah 2. Uji frasa …."), "Jawaban: 36", "Konsep: …", dan baris bagan ("↓ …", "… | …") adalah isi.
- **Rata kanan-kiri dan jarak seragam.** Paragraf, subjudul, dan butir yang berurutan berada di satu kotak teks, jadi jarak antarparagrafnya sama dan mudah disunting di PowerPoint.
- **Halaman tidak bolong.** Isi dibagi rata antarhalaman; halaman terakhir yang nyaris kosong dihapus dengan merapatkan jarak atau mengecilkan gambar (paling kecil 65%); halaman tunggal yang isinya sedikit dilonggarkan. Judul dan kalimat pengantar ("Contohnya:") selalu sehalaman dengan isinya, dan tabel atau daftar tidak menyisakan satu baris sendirian.
- **Tujuan pembelajaran** satu ukuran huruf untuk semua butir; lebih dari 9 butir dibagi ke dua slide. Penutup memuat paling banyak 10 judul subbab.

Markup teks: `**tebal**` dan `$p → q$` untuk simbol/rumus (ditulis dengan Cambria Math karena Poppins tidak punya glyph logika). Di dalam `$…$`, `^{…}` menjadi superskrip dan `_{…}` subskrip, mis. `$x^{2}$` dan `$U_{n}$`; pecahan ditulis mendatar (`3/4`).

Ikon: `globe`, `strategi`, `grafik`, `checklist`, `roket`, `peringatan`, `presenter`, `bingung`. Varian navy dan putih untuk desain Opsi B ada di `generator/assets/ikon/` (dibuat ulang dengan `python generator/buat_ikon.py`).

## Desain Opsi B

- Latar slide isi `F4F6FB`, kartu putih membulat, kartu sorotan navy, catatan kuning muda dengan lingkaran simbol navy (`!` peringatan, `i` tips, `?` pemantik).
- Chip bagian berupa pil navy dengan ikon dalam lingkaran kuning; navigasi bawah pil navy dengan sorotan kuning dan bisa diklik (juga di PDF); nomor halaman dalam lingkaran kuning.
- Sampul dan penutup memakai bentuk bulat bertumpuk navy-kuning; peta subtes bergaya daftar isi bernomor; tujuan berupa kartu dengan level Bloom.
- Palet Opsi B dan Poppins dimasukkan ke tema PowerPoint; Poppins regular dan bold tetap tertanam.
- Generator menolak template dengan ukuran slide selain 16 × 11,25 inci agar tata letak tidak bergeser diam-diam.

## Tampilan General

`generator/tema_general.py` mengambil grafis langsung dari slide contoh `Template Materi General.pptx` saat deck dibuat, sehingga foto dan ilustrasi perusahaan tidak perlu disalin ke repo:

- Sampul, Pendahuluan (bingkai navy), Tujuan, Penutup (bingkai kuning), dan Terima kasih disalin dari slide contohnya; teks contoh diganti isi naskah. Slide contoh dikenali dari teks besarnya (MATERI BESAR, PENDAHULUAN, TUJUAN PEMBELAJARAN, PENUTUP, TERIMA KASIH) dan slide isi bertabel, bukan dari nomor urutnya. Di Penutup, butir ringkasan disusun dua kolom bila perlu, dan pesan yang terlalu panjang dipindah ke slide penutup kedua.
- Slide isi memakai chip penanda dan hiasan samar (lingkaran kuning, gelombang garis) dari template; komponennya sama dengan Opsi B, dengan token warna General: kartu gradien kuning muda ke putih, kartu gelap gradien navy ke abu-abu, catatan kuning.
- Navigasi digambar ulang mengikuti gaya template (pil bergaris navy, bagian aktif navy) agar bagian aktif benar dan tetap bisa diklik; di slide berbingkai warnanya putih.
- Sampul memakai slide contoh berfoto ("Judul Materi (boleh panjang)"), dengan `meta.materi_besar` (bila ada) atau subtes sebagai teks besar, dan `meta.tutor` (opsional) untuk nama tutor.
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
