"""Generator materi Cerebrum: naskah JSON -> PPTX dan/atau PDF.

Contoh:
    python generator/buat_materi.py konten/tiu_penalaran_logis.json --format pptx
    python generator/buat_materi.py konten/tiu_penalaran_logis.json --format pdf
    python generator/buat_materi.py konten/tiu_penalaran_logis.json --format keduanya

PDF diekspor oleh Microsoft PowerPoint (lewat COM), sehingga tampilannya sama
persis dengan PPTX. Template dipakai sebagai dasar agar tema, master, dan font
Poppins yang tertanam ikut terbawa.
"""
from __future__ import annotations

import argparse
import json
import os
import shutil
import sys
import tempfile

from lxml import etree
from pptx import Presentation
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml.ns import qn

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import desain as D  # noqa: E402
import tema_general  # noqa: E402
from desain import P  # noqa: E402

AKAR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TEMA = ("opsi-b", "general")
NAMA_TEMPLATE_TEMA = {
    "opsi-b": "Template Referensi Opsi B - Ceria & Membulat.pptx",
    "general": "Template Materi General.pptx",
}
NAMA_TEMPLATE = NAMA_TEMPLATE_TEMA["opsi-b"]
TEMPLATE_TEMA = {t: os.path.join(AKAR, n) for t, n in NAMA_TEMPLATE_TEMA.items()}
TEMPLATE = TEMPLATE_TEMA["opsi-b"]

# Palet Opsi B untuk tema PowerPoint (template bawaannya masih memakai palet lama).
TEMA_WARNA = {
    "dk1": "000000", "lt1": "FFFFFF", "dk2": "13215A", "lt2": "F4F6FB",
    "accent1": "FBBB2B", "accent2": "13215A", "accent3": "4A5578", "accent4": "FFE6AE",
    "accent5": "8A93AD", "accent6": "C8D1E8", "hlink": "13215A", "folHlink": "4A5578",
}


# ------------------------------------------------------ persiapan file ---
class TemplateTidakAda(FileNotFoundError):
    pass


class TemplateTidakCocok(ValueError):
    pass


def siapkan_presentasi(path_template, tema="opsi-b"):
    """Buka template, aktifkan tema, lalu buang slide contohnya.

    Tema "general" memakai slide contoh template sebagai sumber salinan (sampul, bingkai,
    hiasan), jadi slide itu dikenali dulu sebelum dihapus dari deck.
    """
    if tema not in TEMA:
        raise ValueError(f"Tema tidak dikenal: {tema} (pilihan: {', '.join(TEMA)})")
    nama = NAMA_TEMPLATE_TEMA[tema]
    if not path_template or not os.path.exists(path_template):
        raise TemplateTidakAda(
            f"Template tidak ditemukan: {path_template}\n"
            f"Salin '{nama}' ke folder akar proyek atau pakai --template.")
    prs = Presentation(path_template)
    lebar, tinggi = prs.slide_width / 914400, prs.slide_height / 914400
    if abs(lebar - D.LEBAR_SLIDE) > 0.01 or abs(tinggi - D.TINGGI_SLIDE) > 0.01:
        raise TemplateTidakCocok(
            f"Ukuran slide template {lebar:g} x {tinggi:g} inci, sedangkan generator ditata untuk "
            f"{D.LEBAR_SLIDE:g} x {D.TINGGI_SLIDE:g} inci. Pakai '{nama}'.")
    if tema == "general":
        try:
            D.pakai_tema("general", tema_general.TemaGeneral(list(prs.slides)))
        except tema_general.TemplateGeneralTidakCocok as e:
            raise TemplateTidakCocok(str(e)) from e
    else:
        D.pakai_tema("opsi-b")
    daftar = prs.slides._sldIdLst
    for sid in list(daftar):
        rid = sid.get(qn("r:id"))
        daftar.remove(sid)
        prs.part.drop_rel(rid)
    _rapikan_font(prs)
    _rapikan_tema(prs)
    return prs


def _rapikan_font(prs):
    """Sisakan satu keluarga 'Poppins' (regular + bold) di daftar font tertanam.

    Ada template yang menamai font tebal sebagai keluarga 'Poppins Bold', padahal nama
    internalnya 'Poppins' gaya Bold, sehingga PowerPoint memakai font pengganti. Template
    Opsi B sudah benar (satu entri 'Poppins' berisi regular dan bold); keduanya didukung.
    """
    lst = prs.part._element.find(qn("p:embeddedFontLst"))
    if lst is None:
        return
    peta = {}  # (keluarga, gaya) -> rId
    for ef in lst.findall(qn("p:embeddedFont")):
        nama = ef.find(qn("p:font")).get("typeface")
        for gaya in ("regular", "bold", "italic", "boldItalic"):
            el = ef.find(qn(f"p:{gaya}"))
            if el is not None:
                peta[(nama, gaya)] = el.get(qn("r:id"))
    reg = peta.get(("Poppins", "regular"))
    tebal = peta.get(("Poppins", "bold")) or peta.get(("Poppins Bold", "regular"))
    for ef in list(lst):
        lst.remove(ef)
    if reg:
        ef = etree.SubElement(lst, qn("p:embeddedFont"))
        etree.SubElement(ef, qn("p:font"), typeface="Poppins", panose="00000500000000000000", charset="1")
        etree.SubElement(ef, qn("p:regular")).set(qn("r:id"), reg)
        if tebal:
            etree.SubElement(ef, qn("p:bold")).set(qn("r:id"), tebal)
    for rid in set(peta.values()):
        if rid and rid not in (reg, tebal):
            prs.part.drop_rel(rid)


def _rapikan_tema(prs):
    """Masukkan palet brand dan font Poppins ke tema (bawaan template masih Office)."""
    part = prs.slide_master.part.part_related_by(RT.THEME)
    root = part._element if hasattr(part, "_element") else etree.fromstring(part.blob)
    a = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
    skema = root.find(f".//{a}clrScheme")
    skema.set("name", "Cerebrum Materi")
    for tag, nilai in TEMA_WARNA.items():
        el = skema.find(f"{a}{tag}")
        for anak in list(el):
            el.remove(anak)
        etree.SubElement(el, f"{a}srgbClr", val=nilai)
    fonts = root.find(f".//{a}fontScheme")
    fonts.set("name", "Cerebrum Materi")
    for grup in ("majorFont", "minorFont"):
        fonts.find(f"{a}{grup}/{a}latin").set("typeface", "Poppins")
    if not hasattr(part, "_element"):
        part._blob = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


# ------------------------------------------------------------ pembangun ---
# Blok yang boleh dipecah ke halaman berikut: jenis -> (kunci daftar isi, langkah, minimal di halaman berisi)
BISA_DIPECAH = {"tabel": ("baris", 1, 2), "poin": ("butir", 1, 2), "profil": ("item", D.KOLOM_PROFIL, D.KOLOM_PROFIL)}


class Pembangun:
    varian_sampul = "putih"  # sampul General: 'putih' atau 'foto' (diatur buat_deck)

    def __init__(self, prs, konten, folder=None):
        self.prs, self.k = prs, konten
        self.folder = folder or os.path.join(AKAR, "konten")  # dasar path gambar di naskah
        self.layout = next((l for l in prs.slide_layouts if l.name == "Blank"), prs.slide_layouts[-1])
        self.slides = []
        self.awal_bagian = {}
        self.peringatan = []
        self.asal = []  # bagian naskah asal tiap slide (mis. "materi.2.slide.0"), untuk editor
        self.sumber = None
        self.bingkai = set()  # slide berbingkai warna (nomor halaman versi terang)
        self.general = D.TEMA == "general"

    # kerangka -----------------------------------------------------------
    def baru(self, bagian=None, nav=True, hias=True, bingkai=False):
        """Slide kosong berlatar tema. bingkai: slide berlatar warna (navigasi versi terang)."""
        s = self.prs.slides.add_slide(self.layout)
        for ph in list(s.placeholders):
            ph._element.getparent().remove(ph._element)
        D.latar(s, "latar" if nav else "putih")
        if nav and hias and D._kait("hias"):
            D._KAIT.hias(s)
        if bingkai:
            self.bingkai.add(id(s))
        tombol = D.navbar(s, bagian, gelap=bingkai) if nav else None
        self.slides.append((s, tombol))
        self.asal.append(self.sumber)
        if bagian and bagian not in self.awal_bagian:
            self.awal_bagian[bagian] = s
        return s

    def cek_batas(self, y, nama):
        if y > D.BATAS_BAWAH + 0.01:
            self.peringatan.append(f"Slide '{nama}': konten sampai {y:.2f} in (batas {D.BATAS_BAWAH}).")

    def selesaikan(self):
        total = len(self.slides)
        for i, (s, tombol) in enumerate(self.slides, 1):
            if tombol is None:
                continue
            D.nomor_halaman(s, i, total, gelap=id(s) in self.bingkai)
            for label, shp in tombol.items():
                tujuan = self.awal_bagian.get(label)
                if tujuan is not None and tujuan is not s:
                    shp.click_action.target_slide = tujuan

    # slide khusus ---------------------------------------------------------
    def sampul(self):
        m = self.k["meta"]
        s = self.baru(nav=False)
        if self.general:
            D._KAIT.sampul(s, (m.get("materi_besar") or m["subtes"]).upper(), m["judul"], m.get("tutor"),
                           varian=self.varian_sampul)
            return
        D.sampul(s, m["judul"], m["subtes"].upper(), m.get("subjudul"), ikon=m.get("ikon", "strategi"))

    def penutup_akhir(self):
        s = self.baru(nav=False)
        if self.general:
            D._KAIT.terima_kasih(s)
            return
        D.terima_kasih(s)

    def peta_subtes(self):
        """Peta materi bergaya daftar isi Opsi B: kartu bernomor, kartu aktif berwarna navy."""
        p = self.k["peta_subtes"]
        s = self.baru()
        y = D.kerangka_isi(s, p.get("subjudul", "Peta Materi"), "checklist", p["judul"])
        kartu = p["kartu"]
        n = len(kartu)
        celah = 0.25
        w = (D.LEBAR_ISI - celah * (n - 1)) / n
        inset = (0.3, 0.35, 0.3, 0.35)
        aktif = p.get("aktif", 0)
        # General: isi kartu sedikit lebih besar dan kartu di tengah ruang judul-navigasi,
        # supaya bagian bawah slide tidak kosong.
        if self.general:
            u_nomor, u_judul, u_desk, d_ikon = (50 if n > 5 else 58), 20, 16, 0.85
        else:
            u_nomor, u_judul, u_desk, d_ikon = (44 if n > 5 else 54), 17, 14, 0.7
        isi = []
        for i, k in enumerate(kartu):
            gelap = i == aktif
            isi.append([P(f"{i + 1:02d}", u_nomor, "kuning_di_gelap" if gelap else "navy", tebal=True,
                          spasi=1.0, setelah=14),
                        P(k["judul"], u_judul, "putih" if gelap else "navy", tebal=True, spasi=1.2, setelah=6),
                        P(k["deskripsi"], u_desk, "redup" if gelap else "teks", spasi=1.35)])
        h = max(5.6, max(D.tinggi_paragraf(par, w, inset) for par in isi) + d_ikon + 0.35)
        if self.general:
            ruang = D.BATAS_BAWAH - 0.15 - y
            h = min(max(h, 6.0), ruang)
            y += (ruang - h) / 2
        for i, (k, par) in enumerate(zip(kartu, isi)):
            x = D.KIRI + i * (w + celah)
            D.kotak(s, x, y, w, h, warna="kartu_gelap" if i == aktif else "kartu", radius=D.RADIUS,
                    nama=f"Kartu materi {i + 1}")
            D.teks(s, x, y, w, h - d_ikon, par, inset=inset)
            D.ubin_ikon(s, x + inset[0], y + h - inset[3] - d_ikon, d_ikon, k.get("ikon", "globe"))
        self.cek_batas(y + h, "Peta subtes")

    def pendahuluan(self):
        p, m = self.k["pendahuluan"], self.k["meta"]
        if self.general:
            s = self.baru("Pendahuluan", hias=False, bingkai=True)
            self.cek_batas(D._KAIT.pendahuluan(s, p.get("paragraf", []), p.get("pemantik")), "Pendahuluan")
            return
        s = self.baru("Pendahuluan")
        y = D.kerangka_isi(s, "Pendahuluan", p.get("ikon_chip", "globe"), p.get("judul") or f"Mengenal {m['judul']}")
        w_panel = 5.6
        x_panel = D.KANAN - w_panel
        w = x_panel - 0.5 - D.KIRI
        par = [P(t, 20, "teks", spasi=1.45, setelah=14) for t in p["paragraf"]]
        h = D.tinggi_paragraf(par, w)
        D.teks(s, D.KIRI, y, w, h, par, nama="Paragraf pendahuluan")
        h_panel = max(4.0, min(h, 4.8))
        path_gambar = os.path.join(self.folder, p["gambar"]) if p.get("gambar") else None
        if path_gambar and os.path.exists(path_gambar):
            # gambar dari dokumen sumber, dipaskan ke dalam kartu putih
            D.kotak(s, x_panel, y, w_panel, h_panel, warna="putih", radius=D.RADIUS, nama="Panel gambar")
            pic = D.gambar(s, path_gambar, 0, 0, w=w_panel - 0.5, nama="Gambar pendahuluan")
            skala = min(1.0, (h_panel - 0.5) / (pic.height / 914400))
            pic.width, pic.height = int(pic.width * skala), int(pic.height * skala)
            pic.left = int((x_panel + (w_panel - pic.width / 914400) / 2) * 914400)
            pic.top = int((y + (h_panel - pic.height / 914400) / 2) * 914400)
        else:
            D.kotak(s, x_panel, y, w_panel, h_panel, warna="kartu_gelap", radius=D.RADIUS, nama="Panel ilustrasi")
            ukuran = min(3.0, h_panel - 0.8)
            ilustrasi = D.ikon_path(p.get("ikon", "presenter"), "putih")
            pic = D.gambar(s, ilustrasi, 0, 0, h=ukuran, nama="Ilustrasi")
            pic.left = int((x_panel + (w_panel - pic.width / 914400) / 2) * 914400)
            pic.top = int((y + (h_panel - ukuran) / 2) * 914400)
        y = max(y + h, y + h_panel)
        if p.get("pemantik"):
            y = D.catatan(s, y + 0.35, "pemantik", p["pemantik"], judul_catatan="Pertanyaan pemantik", ukuran=20)
        self.cek_batas(y, "Pendahuluan")

    def tujuan(self):
        """Kartu tujuan: lingkaran kuning berisi level Bloom, kartu pertama navy."""
        t = self.k["tujuan"]
        if self.general:
            s = self.baru("Tujuan Pembelajaran", hias=False)
            self.cek_batas(D._KAIT.tujuan(s, t.get("pengantar"), t["butir"]), "Tujuan")
            return
        s = self.baru("Tujuan Pembelajaran")
        y = D.kerangka_isi(s, "Tujuan Pembelajaran", t.get("ikon_chip", "strategi"), t.get("judul") or t["pengantar"])
        butir = t["butir"]
        n = len(butir)
        kolom = 1 if n == 1 else (3 if n in (3, 5, 6) else 2)
        celah = 0.3
        lw = (D.LEBAR_ISI - celah * (kolom - 1)) / kolom
        d = 1.0
        inset = (0.45, 0.4 + d + 0.25, 0.45, 0.4)
        ukuran = 21 if kolom == 3 else 22
        for awal in range(0, n, kolom):
            baris = butir[awal:awal + kolom]
            par = []
            for j, b in enumerate(baris):
                gelap = awal + j == 0
                segs = [(b["kko"], True, False, "putih" if gelap else "navy")] + D.segmen(" " + b["teks"])
                par.append([P(None, ukuran, "redup" if gelap else "teks", spasi=1.35, segs=segs)])
            h = max(2.6, max(D.tinggi_paragraf(p, lw, inset) for p in par))
            for j, (b, p) in enumerate(zip(baris, par)):
                x = D.KIRI + j * (lw + celah)
                D.kotak(s, x, y, lw, h, warna="kartu_gelap" if awal + j == 0 else "kartu", radius=D.RADIUS,
                        nama="Butir tujuan")
                level = b.get("level") or str(awal + j + 1)
                D.lingkaran(s, x + inset[0], y + 0.4, d, "kuning", level, 24, "di_kuning", nama=f"Level {level}")
                D.teks(s, x, y, lw, h, p, inset=inset)
            y += h + celah
        self.cek_batas(y - celah, "Tujuan")

    def materi(self):
        for i, sub in enumerate(self.k["materi"]):
            for j, sl in enumerate(sub["slide"]):
                self.sumber = f"materi.{i}.slide.{j}"
                if sl["tipe"] == "blok":
                    self._slide_blok(sub, sl)
                elif sl["tipe"] == "contoh_soal":
                    self._contoh_soal(sub, sl)
                else:
                    self._slide_materi(sub, sl)

    def _tata_blok(self, buat_slide, blok, gaya=None, nama="Blok", judul_dari_subjudul=False):
        """Tata blok berurutan; buat_slide(lanjutan, judul) -> (slide, y). Slide baru dibuat bila penuh.

        Jika judul_dari_subjudul, halaman baru yang dimulai tepat di subjudul memakai subjudul itu
        sebagai judul slide; halaman lanjutan lain diberi judul subjudul terakhir + "(lanjutan)".
        """
        s, y = buat_slide(False, None)
        kosong = True
        w = D.LEBAR_ISI
        jarak = 0.1 if gaya == "soal" else D.JARAK_BLOK
        subjudul_akhir = None
        blok = list(blok)
        i = 0
        while i < len(blok):
            b = blok[i]
            sisa = D.BATAS_BAWAH - y
            h = D.tinggi_blok(b, w, self.folder, gaya, sisa)
            pindah_dulu = False
            pecah = BISA_DIPECAH.get(b["jenis"])
            if pecah and h > sisa and len(b[pecah[0]]) > pecah[1]:
                # Blok lebih tinggi dari sisa halaman: pecah (tabel per baris dengan header diulang,
                # daftar per butir, kartu profil per baris kartu); sisanya ke halaman berikut.
                kunci, langkah, minimal = pecah
                k = self._bagian_muat(b, kunci, langkah, w, sisa)
                if k >= (langkah if kosong else minimal):
                    blok[i:i + 1] = [dict(b, **{kunci: b[kunci][:k]}), dict(b, **{kunci: b[kunci][k:]})]
                    b = blok[i]
                    h = D.tinggi_blok(b, w, self.folder, gaya, sisa)
                elif not kosong:
                    pindah_dulu = True  # pecah lagi di halaman baru yang lebih lapang
            if b["jenis"] == "gambar":
                muat = D.skala_min_gambar(b, w) <= sisa or kosong
            else:
                muat = (h <= sisa or kosong) and not pindah_dulu
                if muat and b["jenis"] == "subjudul" and i + 1 < len(blok) and not kosong:
                    berikut = blok[i + 1]
                    h2 = D.skala_min_gambar(berikut, w) if berikut["jenis"] == "gambar" else                         D.tinggi_blok(berikut, w, self.folder, gaya)
                    muat = h + jarak + min(h2, 1.2) <= sisa  # subjudul ikut ke isi berikutnya
            if not muat:
                self.cek_batas(y - jarak, nama)
                if judul_dari_subjudul and b["jenis"] == "subjudul":
                    s, y = buat_slide(True, b["teks"])
                    subjudul_akhir = b["teks"]
                    kosong = True
                    i += 1
                    continue
                s, y = buat_slide(True, f"{subjudul_akhir} (lanjutan)" if subjudul_akhir else None)
                kosong = True
                if pindah_dulu:
                    continue  # ukur ulang blok ini di halaman baru
                sisa = D.BATAS_BAWAH - y
            if b["jenis"] == "subjudul":
                subjudul_akhir = b["teks"]
                if not kosong:
                    y += 0.1
            y = D.gambar_blok(s, b, D.KIRI, y, w, self.folder, gaya, sisa) + jarak
            kosong = False
            i += 1
        y -= jarak
        self.cek_batas(y, nama)
        return s, y

    def _bagian_muat(self, b, kunci, langkah, w, sisa):
        """Jumlah isi terbanyak (kelipatan langkah) dari b[kunci] yang muat dalam tinggi sisa."""
        n = len(b[kunci])
        for k in range((n - 1) // langkah * langkah, 0, -langkah):
            if D.tinggi_blok(dict(b, **{kunci: b[kunci][:k]}), w, self.folder) <= sisa:
                return k
        return 0

    def _slide_blok(self, sub, sl):
        def buat(lanjutan, judul=None):
            s = self.baru("Pendalaman Materi")
            judul = judul or (sl["judul"] + (" (lanjutan)" if lanjutan else ""))
            return s, D.kerangka_isi(s, sub["subbab"], sub.get("ikon", "checklist"), judul,
                                     None if lanjutan else sl.get("pengantar"))

        self._tata_blok(buat, sl["blok"], nama=sl["judul"], judul_dari_subjudul=True)

    def _contoh_soal(self, sub, sl):
        """Slide soal (stem + opsi) lalu slide pembahasan; tombol menautkan keduanya."""
        judul = sl.get("judul") or "Contoh Soal"
        tag = [t for t in (sl.get("label"),) if t]
        soal_slide, bahas_slide = [], []

        def kepala_soal(lanjutan, _judul=None):
            s = self.baru("Pendalaman Materi")
            w_chip = D.chip(s, judul + (" (lanjutan)" if lanjutan else ""), sub.get("ikon", "checklist"))
            x = D.X_CHIP + w_chip + 0.25
            for t in tag:
                _, w = D.pil(s, x, D.Y_CHIP + (D.H_CHIP - 0.46) / 2, t, ukuran=13)
                x += w + 0.15
            soal_slide.append(s)
            return s, D.Y_CHIP + D.H_CHIP + 0.45

        s, y = self._tata_blok(kepala_soal, sl.get("soal", []), gaya="soal", nama=judul)
        opsi = sl.get("opsi") or {}
        if opsi:
            h = D.tinggi_opsi(opsi)
            if y + 0.35 + h > D.BATAS_BAWAH and sl.get("soal"):
                s, y = kepala_soal(True)
                y -= 0.35
            y = D.gambar_opsi(s, y + 0.35, opsi)
            self.cek_batas(y, judul)
        tombol_bahas = D.tombol(s, D.KIRI, 10.36, "Lihat pembahasan ›")

        kunci = sl.get("kunci", "")
        blok = list(sl.get("pembahasan", []))
        if kunci and opsi.get(kunci):
            blok.insert(0, {"jenis": "kunci", "teks": f"**Kunci {kunci}.** {opsi[kunci]}"})

        def kepala_bahas(lanjutan, _judul=None):
            s = self.baru("Pendalaman Materi")
            w_chip = D.chip(s, f"Pembahasan {judul}" + (" (lanjutan)" if lanjutan else ""), "strategi")
            if kunci:
                D.pil(s, D.X_CHIP + w_chip + 0.25, D.Y_CHIP + (D.H_CHIP - 0.56) / 2, f"Jawaban {kunci}",
                      ukuran=18, warna="kuning", warna_teks="di_kuning", h=0.56, nama="Jawaban")
            kembali = D.tombol(s, D.KIRI, 10.36, "‹ Kembali ke soal")
            kembali.click_action.target_slide = soal_slide[0]
            bahas_slide.append(s)
            return s, D.Y_CHIP + D.H_CHIP + 0.45

        if blok:
            self._tata_blok(kepala_bahas, blok, nama=f"Pembahasan {judul}")
            tombol_bahas.click_action.target_slide = bahas_slide[0]

    def _slide_materi(self, sub, sl):
        s = self.baru("Pendalaman Materi")
        y = D.kerangka_isi(s, sub["subbab"], sub.get("ikon", "checklist"), sl["judul"], sl.get("pengantar"))
        tipe = sl["tipe"]
        if tipe == "tabel":
            y = D.tabel(s, y, sl["kolom"], sl["baris"], sl.get("lebar"), sl.get("ukuran", 20))
        elif tipe == "tabel_kebenaran":
            w = 10.5
            x = (D.LEBAR_SLIDE - w) / 2
            y = D.tabel(s, y, sl["kolom"], sl["baris"], x=x, w=w, tengah=True,
                        sorot=sl.get("sorot"), kebenaran=True, ukuran=sl.get("ukuran", 20))
            if sl.get("keterangan"):
                D.teks(s, x, y + 0.1, w, 0.4, [P(sl["keterangan"], 15, "teks", rata="r", spasi=1.2)],
                       nama="Keterangan tabel")
                y += 0.5
        elif tipe == "banding":
            y = D.kartu_banding(s, y, sl["kartu"])
        elif tipe == "notasi":
            y = D.baris_notasi(s, y, sl["baris"])
            if sl.get("ilustrasi"):
                y = D.ilustrasi_setara(s, y + D.JARAK, sl["ilustrasi"]["kiri"], sl["ilustrasi"]["kanan"])
        elif tipe == "kartu":
            y = D.grid_kartu(s, y, sl["kartu"], sl.get("kolom", 2))
        elif tipe == "alur":
            y = D.alur_premis(s, y, sl["premis"], sl["langkah"])
        elif tipe == "langkah":
            y = D.daftar_langkah(s, y, sl["langkah"])
        elif tipe == "teks_ikon":
            y0 = y
            w_panel = 4.6
            x_panel = D.KANAN - w_panel
            y = D.poin_berbutir(s, y, sl["poin"], w=x_panel - 0.5 - D.KIRI)
            h_panel = max(y - y0, 3.4)
            D.kotak(s, x_panel, y0, w_panel, h_panel, warna="kartu_gelap", radius=D.RADIUS, nama="Panel ilustrasi")
            ukuran = min(2.8, h_panel - 0.6)
            pic = D.gambar(s, D.ikon_path(sl.get("ikon", "bingung"), "putih"), 0, 0, h=ukuran, nama="Ilustrasi")
            pic.left = int((x_panel + (w_panel - pic.width / 914400) / 2) * 914400)
            pic.top = int((y0 + (h_panel - ukuran) / 2) * 914400)
            y = y0 + h_panel
        else:
            raise ValueError(f"Tipe slide tidak dikenal: {tipe}")
        if sl.get("catatan"):
            c = sl["catatan"]
            y = D.catatan(s, y + D.JARAK, c["jenis"], c["teks"])
        self.cek_batas(y, sl["judul"])

    def penutup(self):
        """Ringkasan bernomor di kiri, kartu pesan penutup navy di kanan."""
        p = self.k["penutup"]
        if self.general:
            s = self.baru("Penutup", hias=False, bingkai=True)
            self.cek_batas(D._KAIT.penutup(s, p.get("ringkasan", []), p.get("pesan")), "Penutup")
            return
        s = self.baru("Penutup")
        y = D.kerangka_isi(s, "Penutup", p.get("ikon_chip", "roket"), p.get("judul", "Yang Perlu Diingat"))
        ada_pesan = bool(p.get("pesan"))
        ada_ringkasan = bool(p.get("ringkasan"))
        w_pesan = (5.3 if ada_ringkasan else D.LEBAR_ISI) if ada_pesan else 0
        w_baris = D.LEBAR_ISI - w_pesan - 0.3 if ada_pesan and ada_ringkasan else D.LEBAR_ISI
        d = 0.72
        inset = (0.3 + d + 0.3, 0.2, 0.35, 0.2)
        y0 = y
        for i, r in enumerate(p["ringkasan"], 1):
            isi = f"**{r['judul']}:** {r['teks']}" if r.get("teks") else f"**{r['judul']}**"
            par = [P(isi, 18, "navy", spasi=1.3)]
            h = max(1.0, D.tinggi_paragraf(par, w_baris, inset))
            D.kotak(s, D.KIRI, y, w_baris, h, warna="kartu", radius=D.RADIUS, nama=f"Ringkasan {i}")
            D.lingkaran(s, D.KIRI + 0.3, y + (h - d) / 2, d, "kuning", i, 19, "di_kuning", nama=f"Nomor {i}")
            D.teks(s, D.KIRI, y, w_baris, h, par, inset=inset, jangkar="m")
            y += h + 0.2
        y_akhir = y - 0.2 if ada_ringkasan else y0
        if not ada_pesan:
            self.cek_batas(y_akhir, "Penutup")
            return
        sisi = (0.45, 0.45, 0.45, 0.45)
        tinggi_maks = max(y_akhir, D.BATAS_BAWAH) - y0
        for ukuran in (24, 22, 20, 18, 17, 16):  # pesan panjang dari dokumen dikecilkan agar muat
            pp = [P("Pesan penutup", 16, "kuning_di_gelap", tebal=True, spasi=1.2, setelah=10),
                  P(p["pesan"], ukuran, "putih", tebal=ukuran >= 20, spasi=1.3)]
            if D.tinggi_paragraf(pp, w_pesan, sisi) <= tinggi_maks:
                break
        h_pesan = max(y_akhir - y0, D.tinggi_paragraf(pp, w_pesan, sisi))
        x_pesan = D.KANAN - w_pesan
        D.kotak(s, x_pesan, y0, w_pesan, h_pesan, warna="kartu_gelap", radius=D.RADIUS, nama="Pesan penutup")
        D.teks(s, x_pesan, y0, w_pesan, h_pesan, pp, inset=sisi, jangkar="m")
        self.cek_batas(max(y_akhir, y0 + h_pesan), "Penutup")

    def bangun(self):
        for sumber, langkah in (("meta", self.sampul), ("peta_subtes", self.peta_subtes),
                                ("pendahuluan", self.pendahuluan), ("tujuan", self.tujuan),
                                ("materi", self.materi), ("penutup", self.penutup), (None, self.penutup_akhir)):
            self.sumber = sumber
            langkah()
        self.selesaikan()


# ---------------------------------------------------------------- ekspor ---
def ekspor_pdf(path_pptx, path_pdf):
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    app = win32com.client.Dispatch("PowerPoint.Application")
    try:
        pres = app.Presentations.Open(os.path.abspath(path_pptx), True, False, False)
        try:
            pres.SaveAs(os.path.abspath(path_pdf), 32)  # 32 = ppSaveAsPDF
        finally:
            pres.Close()
    finally:
        if app.Presentations.Count == 0:
            app.Quit()


def nama_file(konten):
    return konten["meta"].get("nama_file") or konten["meta"]["judul"]


def buat_deck(konten, path_template=None, folder=None, tema="opsi-b", sampul="putih", palet=None):
    """Bangun deck dari naskah (dict); folder = dasar path gambar di naskah. Hasil: (Presentation, Pembangun).

    sampul: varian sampul General, 'putih' atau 'foto'. palet: nama palet platform (mis. 'jadiasn',
    lihat palet.py) untuk mewarnai ulang navy/kuning template; None = warna template.

    Tema disimpan sebagai keadaan modul desain, jadi pemanggil yang berjalan paralel (server
    web) harus membangun deck satu per satu.
    """
    prs = siapkan_presentasi(path_template or TEMPLATE_TEMA.get(tema, TEMPLATE), tema)
    if palet:
        import palet as PL
        PL.pakai_token(palet)  # token kontras palet berlaku selama deck dibangun
    b = Pembangun(prs, konten, folder)
    b.varian_sampul = sampul
    b.bangun()
    if palet:
        PL.terapkan(prs, palet)
    return prs, b


def main(argv=None):
    ap = argparse.ArgumentParser(description="Buat materi PPTX/PDF dari naskah JSON.")
    ap.add_argument("naskah", help="Path file naskah .json")
    ap.add_argument("--format", choices=["pptx", "pdf", "keduanya"], default="pptx")
    ap.add_argument("--keluar", default=os.path.join(AKAR, "output"), help="Folder hasil")
    ap.add_argument("--tema", choices=TEMA, default="opsi-b",
                    help="Tampilan: opsi-b (Ceria & Membulat) atau general (Template Materi General)")
    ap.add_argument("--template", default=None, help="Path template (bawaan: template tema di akar proyek)")
    ap.add_argument("--sampul", choices=["putih", "foto"], default="putih",
                    help="Varian sampul General: putih (lengkung navy) atau foto (panel navy di atas foto)")
    ap.add_argument("--palet", default=None, help="Palet platform untuk warna slide, mis. jadiasn")
    arg = ap.parse_args(argv)

    with open(arg.naskah, encoding="utf-8") as f:
        konten = json.load(f)
    nama = nama_file(konten) + (" (General)" if arg.tema == "general" else "")
    os.makedirs(arg.keluar, exist_ok=True)

    try:
        prs, b = buat_deck(konten, arg.template, os.path.dirname(os.path.abspath(arg.naskah)), arg.tema,
                           sampul=arg.sampul, palet=arg.palet)
    except (TemplateTidakAda, TemplateTidakCocok, ValueError) as e:
        sys.exit(str(e))
    for w in b.peringatan:
        print("PERINGATAN:", w)

    hasil = []
    path_pptx = os.path.join(arg.keluar, nama + ".pptx")
    if arg.format == "pdf":
        tmp = tempfile.mkdtemp()
        sementara = os.path.join(tmp, "materi.pptx")
        prs.save(sementara)
        path_pdf = os.path.join(arg.keluar, nama + ".pdf")
        ekspor_pdf(sementara, path_pdf)
        shutil.rmtree(tmp, ignore_errors=True)
        hasil.append(path_pdf)
    else:
        prs.save(path_pptx)
        hasil.append(path_pptx)
        if arg.format == "keduanya":
            path_pdf = os.path.join(arg.keluar, nama + ".pdf")
            ekspor_pdf(path_pptx, path_pdf)
            hasil.append(path_pdf)
    print(f"{len(b.slides)} slide dibuat:")
    for h in hasil:
        print("  ", h)


if __name__ == "__main__":
    main()
