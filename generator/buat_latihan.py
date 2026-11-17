"""Generator deck latihan soal: naskah JSON (hasil ekstrak_soal.py) -> PPTX/PDF.

Alur deck: sampul, peta soal, lalu untuk tiap soal satu slide soal diikuti slide
jawaban + pembahasan (bersambung otomatis bila panjang), kunci jawaban, penutup.

    python generator/buat_latihan.py konten/latihan_tiu_paket1.json --format pptx
    python generator/buat_latihan.py konten/latihan_tiu_paket1.json --format keduanya
"""
from __future__ import annotations

import argparse
import json
import os
import re
import shutil
import sys
import tempfile

from pptx.enum.shapes import MSO_SHAPE

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import desain as D  # noqa: E402
from buat_materi import (AKAR, TEMPLATE, TemplateTidakAda, TemplateTidakCocok,  # noqa: E402
                         ekspor_pdf, siapkan_presentasi)
from desain import P  # noqa: E402

NAV_LATIHAN = ["Verbal", "Numerik", "Figural", "Kunci Jawaban"]
IKON_NAV = {"Verbal": "checklist", "Numerik": "grafik", "Figural": "strategi"}

# urutan pemeriksaan penting: kata kunci yang lebih spesifik lebih dulu
KELOMPOK = [
    ("Figural", ("Figural",), "strategi", "Figural"),
    ("Penalaran Logis", ("Premis",), "checklist", "Verbal"),
    ("Deret & Pola", ("Deret", "Pola Bilangan"), "roket", "Numerik"),
    ("Soal Cerita & Tabel", ("Soal Cerita", "Tabel Pernyataan"), "globe", "Numerik"),
    ("Verbal & Analitis", ("Kata", "Argumen", "Susunan"), "strategi", "Verbal"),
    ("Berhitung & Aljabar", ("Operasi", "Setara", "Aljabar", "Bilangan"), "grafik", "Numerik"),
]

BADAN, RUMUS, JUDUL_KECIL = 19, 22, 19
Y_ISI = 2.1
JARAK_TEKS, JARAK_BLOK = 0.12, 0.28


def kelompok(subtipe):
    for nama, kunci, ikon, nav in KELOMPOK:
        if any(k in subtipe for k in kunci):
            return nama, ikon, nav
    return "Verbal & Analitis", "strategi", "Verbal"


def rentang(nomor):
    nomor = sorted(nomor)
    bagian, awal = [], nomor[0]
    for a, b in zip(nomor, nomor[1:] + [None]):
        if b != a + 1:
            bagian.append(f"{awal}" if awal == a else f"{awal}–{a}")
            awal = b
    return ", ".join(bagian)


def polos(teks):
    return re.sub(r"\*\*|\$", "", teks)


# ---------------------------------------------------------------- unit ---
class Unit:
    """Potongan isi pembahasan yang bisa diukur lalu digambar pada posisi y."""

    def __init__(self, tinggi, gambar, jarak=JARAK_TEKS, judul=False):
        self.tinggi, self.gambar, self.jarak, self.judul = tinggi, gambar, jarak, judul
        self.h_min = None  # diisi untuk gambar yang boleh diperkecil


class UnitGambar(Unit):
    def __init__(self, path, w, h, skala_min=0.65):
        super().__init__(h, self._gambar, JARAK_BLOK)
        self.path, self.w = path, w
        self.h_min = h * skala_min

    def perkecil(self, kurang):
        baru = self.tinggi - kurang
        self.w *= baru / self.tinggi
        self.tinggi = baru

    def _gambar(self, s, y):
        D.gambar(s, self.path, (D.LEBAR_SLIDE - self.w) / 2, y, w=self.w, h=self.tinggi, nama="Gambar pembahasan")


class PembangunLatihan:
    def __init__(self, prs, naskah, folder_naskah):
        self.prs, self.n = prs, naskah
        self.folder_gambar = os.path.join(folder_naskah, naskah["meta"]["folder_gambar"])
        self.layout = next((l for l in prs.slide_layouts if l.name == "Blank"), prs.slide_layouts[-1])
        self.slides, self.target, self.tautan = [], {}, []

    # kerangka -----------------------------------------------------------
    def baru(self, aktif=None, nav=True, kunci=None):
        s = self.prs.slides.add_slide(self.layout)
        for ph in list(s.placeholders):
            ph._element.getparent().remove(ph._element)
        D.latar(s, "latar" if nav else "putih")
        tombol = D.navbar(s, aktif, NAV_LATIHAN) if nav else None
        self.slides.append((s, tombol))
        if kunci and kunci not in self.target:
            self.target[kunci] = s
        return s

    def taut(self, shp, kunci):
        self.tautan.append((shp, kunci))

    def selesaikan(self):
        total = len(self.slides)
        for i, (s, tombol) in enumerate(self.slides, 1):
            if tombol is None:
                continue
            D.nomor_halaman(s, i, total)
            for label, shp in tombol.items():
                tujuan = self.target.get(("nav", label))
                if tujuan is not None and tujuan is not s:
                    shp.click_action.target_slide = tujuan
        for shp, kunci in self.tautan:
            if kunci in self.target:
                shp.click_action.target_slide = self.target[kunci]

    def path_gambar(self, nama):
        return os.path.join(self.folder_gambar, nama)

    # sampul, peta, kunci, penutup -----------------------------------------
    def sampul(self):
        m = self.n["meta"]
        s = self.baru(nav=False)
        tag = " · ".join(x for x in (m.get("paket"), m.get("subtes")) if x)
        D.sampul(s, m["judul"], tag.upper() or None, f"{len(self.n['soal'])} soal dengan pembahasan",
                 ikon="checklist")

    def peta(self):
        """Peta kelompok soal bergaya daftar isi Opsi B; kartu pertama navy, kartu bisa diklik."""
        m, soal = self.n["meta"], self.n["soal"]
        s = self.baru()
        sub = f"{m.get('paket', '')} · {len(soal)} soal".strip(" ·")
        y = D.kerangka_isi(s, "Peta Soal", "checklist", m["judul"], sub)
        grup = {}
        for so in soal:
            nama, ikon, _ = kelompok(so["subtipe"])
            grup.setdefault(nama, {"ikon": ikon, "nomor": []})["nomor"].append(so["nomor"])
        daftar = sorted(grup.items(), key=lambda kv: min(kv[1]["nomor"]))
        n = len(daftar)
        celah = 0.25
        w = (D.LEBAR_ISI - celah * (n - 1)) / n
        inset = (0.3, 0.3, 0.3, 0.3)
        isi = []
        for i, (nama, g) in enumerate(daftar):
            gelap = i == 0
            isi.append([P(nama, 17, "putih" if gelap else "navy", tebal=True, spasi=1.2, setelah=6),
                        P(f"Soal {rentang(g['nomor'])}", 14, "kuning" if gelap else "navy", tebal=True, spasi=1.3,
                          setelah=2),
                        P(f"{len(g['nomor'])} soal", 13, "redup" if gelap else "teks", spasi=1.3)])
        d_ikon = 0.64
        h = max(3.0, max(D.tinggi_paragraf(p, w, inset) for p in isi) + d_ikon + 0.3)
        for i, ((nama, g), par) in enumerate(zip(daftar, isi)):
            x = D.KIRI + i * (w + celah)
            kartu = D.kotak(s, x, y, w, h, warna="navy" if i == 0 else "putih", radius=D.RADIUS,
                            nama=f"Kelompok {nama}")
            D.teks(s, x, y, w, h - d_ikon, par, inset=inset)
            D.ubin_ikon(s, x + inset[0], y + h - inset[3] - d_ikon, d_ikon, g["ikon"])
            self.taut(kartu, ("soal", min(g["nomor"])))
        D.catatan(s, y + h + 0.4, "tips",
                  "Kerjakan setiap soal lebih dulu, lalu buka slide berikutnya untuk jawaban dan pembahasan. "
                  "Klik kartu atau menu bawah untuk berpindah bagian.", judul_catatan="Petunjuk", ukuran=18)

    def kunci_jawaban(self):
        soal = self.n["soal"]
        s = self.baru("Kunci Jawaban", kunci=("nav", "Kunci Jawaban"))
        D.chip(s, "Kunci Jawaban", "checklist")
        kol = 7
        celah = 0.2
        w = (D.LEBAR_ISI - celah * (kol - 1)) / kol
        h = 1.12
        y0 = 2.15
        for i, so in enumerate(soal):
            x = D.KIRI + (i % kol) * (w + celah)
            y = y0 + (i // kol) * (h + celah)
            sel = D.kotak(s, x, y, w, h, warna="tint", radius=0.12, nama=f"Kunci {so['nomor']}")
            D.isi_teks(sel.text_frame, [
                P(f"Soal {so['nomor']}", 14, "teks", rata="c", spasi=1.1, setelah=2),
                P(so["kunci"] or "–", 30, "teal", tebal=True, rata="c", spasi=1.0)], jangkar="m")
            self.taut(sel, ("bahas", so["nomor"]))
        baris = (len(soal) + kol - 1) // kol
        D.teks(s, D.KIRI, y0 + baris * (h + celah) + 0.05, D.LEBAR_ISI, 0.45,
               [P("Klik nomor soal untuk membuka pembahasannya.", 16, "teks", spasi=1.2)], nama="Keterangan")

    def penutup(self):
        s = self.baru(nav=False)
        D.terima_kasih(s, "Selamat berlatih dan sampai jumpa di paket berikutnya.")

    # slide soal -------------------------------------------------------------
    def _tag(self, s, w_chip, so):
        x = D.KIRI + w_chip + 0.3
        y = D.Y_CHIP + (D.H_CHIP - 0.46) / 2
        for t in (so["subtipe"], so.get("level"), so.get("bloom")):
            if t:
                _, w = D.pil(s, x, y, t)
                x += w + 0.15

    def _opsi_mendatar(self, opsi):
        return len(opsi) == 5 and all(len(polos(v)) <= 24 for v in opsi.values())

    def _tata_soal(self, s, so, uk_soal, uk_opsi, gambar, rapat=False):
        """Ukur (gambar=False) atau gambar tata letak soal. Mengembalikan y akhir."""
        y = Y_ISI
        opsi = so["opsi"]
        jarak_teks = 0.05 if rapat else JARAK_TEKS
        t_min = 0.42 if rapat else 0.5
        # tinggi blok opsi lebih dulu agar gambar soal bisa menyesuaikan sisa ruang
        mendatar = self._opsi_mendatar(opsi)
        lw = (D.LEBAR_ISI - 0.25 * 4) / 5
        if not opsi:
            h_opsi = 0
        elif mendatar:
            h_opsi = max(max(0.85, D.tinggi_paragraf([P(v, uk_opsi + 2, spasi=1.25)], lw - 0.85) + 0.3)
                         for v in opsi.values())
        else:
            h_opsi = sum(self._h_opsi(v, uk_opsi, rapat) for v in opsi.values()) + \
                self._celah_opsi(rapat) * (len(opsi) - 1)

        blok = so["soal"]
        h_tetap = 0
        for b in blok:
            if b["jenis"] == "teks":
                h_tetap += self._h_teks_soal(b, uk_soal) + jarak_teks
            elif b["jenis"] == "tabel":
                kol, baris = self._kepala(b["baris"])
                _, w, _ = self._letak_tabel(kol, baris)
                h_tetap += D.tinggi_tabel(kol, baris, ukuran=uk_opsi, w=w, tinggi_min=t_min) + JARAK_BLOK
        n_gbr = sum(1 for b in blok if b["jenis"] == "gambar")
        sisa = D.BATAS_BAWAH - Y_ISI - h_tetap - (h_opsi + 0.3 if opsi else 0) - JARAK_BLOK * n_gbr
        h_gbr_maks = max(1.5, sisa / max(1, n_gbr))

        for b in blok:
            if b["jenis"] == "teks":
                h = self._h_teks_soal(b, uk_soal)
                if gambar:
                    D.teks(s, D.KIRI, y, D.LEBAR_ISI, h, self._p_teks_soal(b, uk_soal), nama="Teks soal")
                y += h + jarak_teks
            elif b["jenis"] == "tabel":
                kol, baris = self._kepala(b["baris"])
                x, w, tengah = self._letak_tabel(kol, baris)
                if gambar:
                    y = D.tabel(s, y + 0.08, kol, baris, ukuran=uk_opsi, x=x, w=w, tengah=tengah,
                                tinggi_min=t_min) + JARAK_BLOK
                else:
                    y += D.tinggi_tabel(kol, baris, ukuran=uk_opsi, w=w, tinggi_min=t_min) + 0.08 + JARAK_BLOK
            elif b["jenis"] == "gambar":
                w, h = self._ukuran_gambar(b, h_gbr_maks)
                if gambar:
                    D.gambar(s, self.path_gambar(b["file"]), (D.LEBAR_SLIDE - w) / 2, y + 0.05, w=w, h=h,
                             nama="Gambar soal")
                y += h + JARAK_BLOK
        if opsi:
            y += 0.15
            if gambar:
                self._gambar_opsi(s, y, opsi, uk_opsi, mendatar, lw, h_opsi, rapat)
            y += h_opsi
        return y

    @staticmethod
    def _h_opsi(isi, ukuran, rapat):
        return max(0.5 if rapat else 0.62,
                   D.tinggi_paragraf([P(isi, ukuran, spasi=1.3)], D.LEBAR_ISI - 1.2) + (0.14 if rapat else 0.24))

    @staticmethod
    def _celah_opsi(rapat):
        return 0.07 if rapat else 0.13

    @staticmethod
    def _letak_tabel(kol, baris):
        """Tabel berisi angka/isi pendek dibuat ringkas dan rata tengah."""
        semua = ([kol] if kol else []) + baris
        n = max(len(b) for b in semua)
        isi_maks = max(len(polos(c)) for b in baris for c in b) if baris else 0
        kepala_maks = max((len(polos(c)) for c in kol), default=0) if kol else 0
        if isi_maks <= 8 and kepala_maks <= 20:
            w = min(D.LEBAR_ISI, n * (2.2 if not kol else max(2.6, kepala_maks * 0.16)))
            return (D.LEBAR_SLIDE - w) / 2, w, True
        return D.KIRI, D.LEBAR_ISI, False

    def _p_teks_soal(self, b, ukuran):
        if b.get("rumus"):
            isi = b["teks"] if "$" in b["teks"] else f"${b['teks']}$"
            return [P(isi, ukuran + 2, "deep", spasi=1.3)]
        return [P(b["teks"], ukuran, "teks", spasi=1.35)]

    def _h_teks_soal(self, b, ukuran):
        return D.tinggi_paragraf(self._p_teks_soal(b, ukuran), D.LEBAR_ISI)

    def _gambar_opsi(self, s, y, opsi, ukuran, mendatar, lw, h_total, rapat=False):
        d = 0.5
        if mendatar:
            for i, (huruf, isi) in enumerate(opsi.items()):
                x = D.KIRI + i * (lw + 0.25)
                D.kotak(s, x, y, lw, h_total, warna="tint", radius=0.12, nama=f"Opsi {huruf}")
                D.lingkaran(s, x + 0.2, y + (h_total - d) / 2, d, "kuning", huruf, 18, "navy", nama=f"Huruf {huruf}")
                D.teks(s, x + 0.85, y, lw - 0.95, h_total, [P(isi, ukuran + 2, "teks", spasi=1.25)], jangkar="m")
            return
        if rapat:
            d = 0.42
        for huruf, isi in opsi.items():
            p = [P(isi, ukuran, "teks", spasi=1.3)]
            h = self._h_opsi(isi, ukuran, rapat)
            D.kotak(s, D.KIRI, y, D.LEBAR_ISI, h, warna="tint", radius=0.1, nama=f"Opsi {huruf}")
            D.lingkaran(s, D.KIRI + 0.15, y + (h - d) / 2, d, "kuning", huruf, 16 if rapat else 18, "navy",
                        nama=f"Huruf {huruf}")
            D.teks(s, D.KIRI + 0.9, y, D.LEBAR_ISI - 1.2, h, p, jangkar="m")
            y += h + self._celah_opsi(rapat)

    def _ukuran_gambar(self, b, h_maks, skala_maks=2.0):
        w_in, h_in = b["lebar_pt"] / 72, b["tinggi_pt"] / 72
        skala = min(skala_maks, D.LEBAR_ISI / w_in, h_maks / h_in)
        return w_in * skala, h_in * skala

    @staticmethod
    def _kepala(baris):
        if baris and all(c.startswith("**") and c.endswith("**") for c in baris[0] if c):
            return baris[0], baris[1:]
        return None, baris

    def slide_soal(self, so):
        _, _, nav = kelompok(so["subtipe"])
        s = self.baru(nav, kunci=("soal", so["nomor"]))
        if ("nav", nav) not in self.target:
            self.target[("nav", nav)] = s
        w_chip = D.chip(s, f"Soal {so['nomor']}", IKON_NAV[nav])
        self._tag(s, w_chip, so)
        tingkat = ((22, 20, False), (21, 19, False), (20, 18, False), (20, 18, True), (19, 17, True),
                   (18, 17, True), (17, 16, True), (16, 15, True))
        for uk_soal, uk_opsi, rapat in tingkat:
            if self._tata_soal(None, so, uk_soal, uk_opsi, gambar=False, rapat=rapat) <= D.BATAS_BAWAH:
                break
        else:
            print(f"PERINGATAN: soal {so['nomor']} tetap melebihi batas slide pada ukuran terkecil.")
        self._tata_soal(s, so, uk_soal, uk_opsi, gambar=True, rapat=rapat)
        self.taut(D.tombol(s, D.KIRI, 10.36, "Lihat pembahasan ›"), ("bahas", so["nomor"]))

    # slide pembahasan --------------------------------------------------------
    def _unit_teks(self, par, w=D.LEBAR_ISI, x=D.KIRI, inset=(0, 0, 0, 0), jarak=JARAK_TEKS, judul=False,
                   latar=None):
        h = D.tinggi_paragraf(par, w, inset)

        def gambar(s, y):
            if latar:
                D.kotak(s, x, y, w, h, warna=latar, radius=0.1)
            D.teks(s, x, y, w, h, par, inset=inset)
        return Unit(h, gambar, jarak, judul)

    def _unit_pembahasan(self, so):
        units, buf_rumus = [], []

        def tutup_rumus():
            if buf_rumus:
                par = [P(t if "$" in t else f"${t}$", RUMUS, "deep", spasi=1.35) for t in buf_rumus]
                units.append(self._unit_teks(par, inset=(0.4, 0.14, 0.4, 0.14), latar="tint", jarak=JARAK_BLOK))
                buf_rumus.clear()

        for b in so["pembahasan"]:
            if b["jenis"] == "teks" and b.get("rumus"):
                buf_rumus.append(b["teks"])
                continue
            tutup_rumus()
            if b["jenis"] == "teks":
                t = b["teks"]
                if t.startswith("**Cara cepat:**"):
                    isi = t.replace("**Cara cepat:**", "").strip()
                    h = D.tinggi_catatan(isi, judul_catatan="Cara cepat", ukuran=18)
                    units.append(Unit(h, lambda s, y, isi=isi: D.catatan(s, y, "tips", isi, judul_catatan="Cara cepat",
                                                                          ukuran=18), JARAK_BLOK))
                elif re.fullmatch(r"\*\*[^*]+\*\*", t) or re.fullmatch(r"\*\*[^*]+:\*\*\s*", t):
                    units.append(self._unit_teks([P(t, JUDUL_KECIL, "teal", tebal=True, spasi=1.25)], judul=True))
                elif t.startswith("Jadi, jawaban yang benar"):
                    units.append(self._unit_teks([P(f"**{t}**", BADAN, "teal", spasi=1.3)], jarak=JARAK_BLOK))
                else:
                    # kalimat pengantar pendek berakhiran ':' ikut pindah bersama isi berikutnya
                    pengantar = t.rstrip().endswith(":") and len(polos(t)) < 90
                    units.append(self._unit_teks([P(t, BADAN, "teks", spasi=1.32)], judul=pengantar))
            elif b["jenis"] == "tabel":
                kol, baris = self._kepala(b["baris"])
                x, w, tengah = self._letak_tabel(kol, baris)
                h = D.tinggi_tabel(kol, baris, ukuran=17, w=w, tinggi_min=0.46)
                units.append(Unit(h, lambda s, y, k=kol, r=baris, x=x, w=w, c=tengah: D.tabel(
                    s, y, k, r, ukuran=17, x=x, w=w, tengah=c, tinggi_min=0.46), JARAK_BLOK))
            elif b["jenis"] == "gambar":
                w, h = self._ukuran_gambar(b, 6.6)
                units.append(UnitGambar(self.path_gambar(b["file"]), w, h))
            elif b["jenis"] == "analisis":
                units.append(self._unit_teks([P("**Analisis setiap opsi**", JUDUL_KECIL, "teal", spasi=1.25)],
                                             judul=True))
                for c in b["catatan"]:
                    units.append(self._unit_teks([P(c, 17, "teks", spasi=1.3)]))
                for it in b["item"]:
                    units.append(self._unit_analisis(it))
        tutup_rumus()
        return units

    def _unit_analisis(self, it):
        d = 0.5
        rumus = f"${it['rumus']}$  " if it["rumus"] else ""
        status = it["status"][0].upper() + it["status"][1:]
        par = [P(it["teks"], 17, "teks", tebal=True, spasi=1.25, setelah=4),
               P(f"{rumus}{it['alasan']}", 17, "teks", spasi=1.3, setelah=4),
               P(f"**{status}**", 15, "teal" if it["benar"] else "slate", spasi=1.2)]
        inset = (0.95, 0.16, 0.3, 0.16)
        h = D.tinggi_paragraf(par, D.LEBAR_ISI, inset)

        def gambar(s, y):
            D.kotak(s, D.KIRI, y, D.LEBAR_ISI, h, warna="sage" if it["benar"] else "tint", radius=0.1,
                    nama=f"Analisis {it['opsi']}")
            b = D.kotak(s, D.KIRI + 0.22, y + 0.18, d, d, gradien=D.GRADIEN_BRAND if it["benar"] else None,
                        warna=None if it["benar"] else "slate", bentuk=MSO_SHAPE.OVAL, radius=0)
            D.isi_teks(b.text_frame, [P(it["opsi"], 17, "putih", tebal=True, rata="c", spasi=1.0)], jangkar="m")
            D.teks(s, D.KIRI, y, D.LEBAR_ISI, h, par, inset=inset)
        return Unit(h, gambar, 0.12)

    def _kepala_bahas(self, so, nav, lanjutan):
        s = self.baru(nav, kunci=None if lanjutan else ("bahas", so["nomor"]))
        label = f"Pembahasan Soal {so['nomor']}" + (" (lanjutan)" if lanjutan else "")
        w_chip = D.chip(s, label, IKON_NAV[nav])
        if so["kunci"]:
            D.pil(s, D.KIRI + w_chip + 0.3, D.Y_CHIP + (D.H_CHIP - 0.56) / 2, f"Jawaban {so['kunci']}",
                  ukuran=18, warna="kuning", warna_teks="navy", h=0.56, nama="Jawaban")
        self.taut(D.tombol(s, D.KIRI, 10.36, "‹ Kembali ke soal"), ("soal", so["nomor"]))
        return s

    def slide_pembahasan(self, so):
        _, _, nav = kelompok(so["subtipe"])
        units = self._unit_pembahasan(so)
        s = self._kepala_bahas(so, nav, False)
        y = Y_ISI
        if so["kunci"] and so["opsi"].get(so["kunci"]):
            par = [P(f"**Kunci {so['kunci']}.** {so['opsi'][so['kunci']]}", BADAN, "deep", spasi=1.3)]
            u = self._unit_teks(par, inset=(0.35, 0.16, 0.35, 0.16), latar="sage", jarak=JARAK_BLOK)
            u.gambar(s, y)
            y += u.tinggi + JARAK_BLOK
        # rencanakan halaman dulu (gambar bisa diperkecil agar sisa kecil tidak pindah slide)
        halaman, y0 = [[]], y
        i = 0
        while i < len(units):
            u = units[i]
            perlu = u.tinggi
            if u.judul and i + 1 < len(units):
                perlu += u.jarak + units[i + 1].tinggi  # judul ikut dengan isi berikutnya
            if y + perlu > D.BATAS_BAWAH and halaman[-1]:
                sisa = sum(v.tinggi + v.jarak for v in units[i:]) - units[-1].jarak
                kurang = y + sisa - D.BATAS_BAWAH
                gbr = [v for v in halaman[-1] if v.h_min is not None]
                cadangan = sum(v.tinggi - v.h_min for v in gbr)
                if gbr and 0 < kurang <= cadangan:
                    for v in gbr:
                        v.perkecil(kurang * (v.tinggi - v.h_min) / cadangan)
                    y -= kurang
                else:
                    halaman.append([])
                    y = Y_ISI
            halaman[-1].append(u)
            y += u.tinggi + u.jarak
            i += 1
        for n, isi in enumerate(halaman):
            if n > 0:
                s = self._kepala_bahas(so, nav, True)
            y = y0 if n == 0 else Y_ISI
            for u in isi:
                u.gambar(s, y)
                y += u.tinggi + u.jarak

    def bangun(self):
        self.sampul()
        self.peta()
        for so in self.n["soal"]:
            self.slide_soal(so)
            self.slide_pembahasan(so)
        self.kunci_jawaban()
        self.penutup()
        self.selesaikan()


def main(argv=None):
    ap = argparse.ArgumentParser(description="Buat deck latihan soal PPTX/PDF dari naskah JSON.")
    ap.add_argument("naskah")
    ap.add_argument("--format", choices=["pptx", "pdf", "keduanya"], default="pptx")
    ap.add_argument("--keluar", default=os.path.join(AKAR, "output"))
    ap.add_argument("--template", default=TEMPLATE)
    a = ap.parse_args(argv)
    with open(a.naskah, encoding="utf-8") as f:
        naskah = json.load(f)
    os.makedirs(a.keluar, exist_ok=True)
    try:
        prs = siapkan_presentasi(a.template)
    except (TemplateTidakAda, TemplateTidakCocok) as e:
        sys.exit(str(e))
    b = PembangunLatihan(prs, naskah, os.path.dirname(os.path.abspath(a.naskah)))
    b.bangun()
    nama = naskah["meta"]["nama_file"]
    hasil = []
    if a.format == "pdf":
        tmp = tempfile.mkdtemp()
        sementara = os.path.join(tmp, "latihan.pptx")
        prs.save(sementara)
        ekspor_pdf(sementara, os.path.join(a.keluar, nama + ".pdf"))
        shutil.rmtree(tmp, ignore_errors=True)
        hasil.append(os.path.join(a.keluar, nama + ".pdf"))
    else:
        path = os.path.join(a.keluar, nama + ".pptx")
        prs.save(path)
        hasil.append(path)
        if a.format == "keduanya":
            ekspor_pdf(path, os.path.join(a.keluar, nama + ".pdf"))
            hasil.append(os.path.join(a.keluar, nama + ".pdf"))
    print(f"{len(b.slides)} slide dibuat:")
    for h in hasil:
        print("  ", h)


if __name__ == "__main__":
    main()
