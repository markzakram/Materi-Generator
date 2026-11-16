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
pip install -r requirements.txt
```

## Berkas yang tidak ikut di repo

Berkas berikut milik Cerebrum dan sengaja dikecualikan lewat `.gitignore`. Salin manual dari penyimpanan internal bila perlu:

| Berkas | Fungsi | Jika tidak ada |
|---|---|---|
| `Template Desain Materi.pptx` (akar proyek) | Dasar master, tema, dan font tertanam | Generator berhenti dengan pesan; wajib ada |
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
