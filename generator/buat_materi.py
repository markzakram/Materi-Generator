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
import re
import shutil
import sys
import tempfile

from lxml import etree
from PIL import Image
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
            D._KAIT.sampul(s, (m.get("materi_besar") or m["subtes"]).upper(), m["judul"], m.get("tutor"))
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
            gambar = None
            if p.get("gambar") and os.path.exists(os.path.join(self.folder, p["gambar"])):
                with Image.open(os.path.join(self.folder, p["gambar"])) as im:
                    gambar = {"jenis": "gambar", "file": p["gambar"], "lebar": 6.0, "tinggi": 6.0 * im.height / im.width}
                if p.get("keterangan"):
                    gambar["keterangan"] = p["keterangan"]
            par, pemantik = p.get("paragraf", []), p.get("pemantik")
            muat = gambar and D._KAIT.tinggi_gambar_pendahuluan(par, pemantik, gambar, self.folder)
            for isi in ([(par, pemantik, gambar)] if muat or not gambar else [(par, pemantik, None), ([], None, gambar)]):
                s = self.baru("Pendahuluan", hias=False, bingkai=True)
                self.cek_batas(D._KAIT.pendahuluan(s, *isi, folder=self.folder), "Pendahuluan")
            return
        s = self.baru("Pendahuluan")
        y = D.kerangka_isi(s, "Pendahuluan", p.get("ikon_chip", "globe"), p.get("judul") or f"Mengenal {m['judul']}")
        w_panel = 5.6
        x_panel = D.KANAN - w_panel
        w = x_panel - 0.5 - D.KIRI
        par = [P(t, 20, "teks", spasi=1.45, setelah=14) for t in p["paragraf"] + ([p["keterangan"]] if p.get("keterangan") else [])]
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
            n = len(t["butir"])
            per = n if n <= 9 else -(-n // -(-n // 8))  # lebih dari 9 butir: bagi rata ke beberapa slide
            for i in range(0, max(n, 1), max(per, 1)):
                s = self.baru("Tujuan Pembelajaran", hias=False)
                self.cek_batas(D._KAIT.tujuan(s, t.get("pengantar"), t["butir"][i:i + per], i), "Tujuan")
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
        self._siapkan_tingkat()
        for i, sub in enumerate(self.k["materi"]):
            for j, sl in enumerate(sub["slide"]):
                self.sumber = f"materi.{i}.slide.{j}"
                if sl["tipe"] == "blok":
                    if j and sub["slide"][j - 1]["tipe"] == "blok":
                        continue  # sudah ikut slide 'blok' pertama di kelompoknya
                    kumpulan = []
                    for k in range(j, len(sub["slide"])):
                        if sub["slide"][k]["tipe"] != "blok":
                            break
                        kumpulan.append((f"materi.{i}.slide.{k}", sub["slide"][k]))
                    self._slide_blok(sub, kumpulan)
                elif sl["tipe"] == "contoh_soal":
                    self._contoh_soal(sub, sl)
                else:
                    self._slide_materi(sub, sl)

    # ---- tata letak blok -------------------------------------------------
    Y_ISI = D.Y_CHIP + D.H_CHIP + 0.42

    def _siapkan_tingkat(self):
        """Paragraf tebal (tingkat 2) dinaikkan menjadi subjudul biasa bila deck tidak punya subjudul lain,
        supaya ukuran subjudul seragam di seluruh deck."""
        ada = any(b["jenis"] == "subjudul" and b.get("tingkat", 1) == 1 for sub in self.k["materi"] for sl in sub["slide"]
                  if sl["tipe"] == "blok" for b in sl["blok"])
        self.tingkat_tebal = 2 if ada else 1

    def _tingkat(self, b):
        t = b.get("tingkat", 1)
        return self.tingkat_tebal if t == 2 else t

    def _celah(self, a, b, jarak):
        """Jarak tegak antara dua blok berurutan di satu halaman."""
        if a is None:
            return 0.0
        if b["jenis"] == "subjudul":
            return jarak + (0.04 if a["jenis"] == "subjudul" else {0: 0.2, 1: 0.14, 2: 0.06}[b.get("tingkat", 1)])
        if a["jenis"] == "subjudul":
            return 0.12 if b["jenis"] in D.TEKS_ALIR else 0.16
        return jarak

    def _rencana(self, blok, y0, batas, gaya, jarak, ulang=True):
        """Bagi blok ke halaman -> [[blok, ...], ...]; tiap blok diberi '_h' (tinggi). Di halaman sambungan, judul yang
        isinya masih berlanjut diulang dengan ukuran yang sama ('ulang')."""
        w = D.LEBAR_ISI
        antre, halaman, kini, y, aktif, i = list(blok), [], [], y0, {}, 0
        while i < len(antre):
            b = antre[i]
            celah = self._celah(kini[-1] if kini else None, b, jarak)
            sisa = batas - y - celah
            kosong = all(x.get("ulang") for x in kini)
            h = D.tinggi_blok(b, w, self.folder, gaya, sisa)
            pindah, pecah = False, BISA_DIPECAH.get(b["jenis"])
            if pecah and h > sisa and len(b[pecah[0]]) > pecah[1]:
                # Blok lebih tinggi dari sisa halaman: pecah (tabel per baris dengan header diulang, daftar per
                # butir, kartu profil per baris kartu); sisanya ke halaman berikut.
                kunci, langkah, minimal = pecah
                k = self._bagian_muat(b, kunci, langkah, w, sisa, gaya)
                if len(b[kunci]) - k == 1 and b["jenis"] != "profil":  # satu baris/butir tidak ditinggal sendirian
                    k = max(k - (2 if b["jenis"] == "tabel" else 1), min(k, 3))
                if k >= (langkah if kosong else minimal):
                    antre[i:i + 1] = [dict(b, **{kunci: b[kunci][:k]}), dict(b, **{kunci: b[kunci][k:]})]
                    b = antre[i]
                    h = D.tinggi_blok(b, w, self.folder, gaya, sisa)
                elif not kosong:
                    pindah = True  # pecah lagi di halaman baru yang lebih lapang
            if b["jenis"] == "gambar":
                muat = D.skala_min_gambar(b, w) <= sisa or kosong
            else:
                muat = (h <= sisa or kosong) and not pindah
            if muat and self._pembuka(b) and i + 1 < len(antre) and not kosong:
                muat = self._tinggi_perlu(antre, i, w, gaya, jarak) <= sisa  # judul tidak tertinggal di dasar halaman
            if not muat:
                halaman.append(kini)
                kini, y = [], y0
                for u in self._judul_ulang(aktif, b) if ulang else ():
                    c = self._celah(kini[-1] if kini else None, u, jarak)
                    u = dict(u, ulang=True, _h=D.tinggi_blok(u, w, self.folder, gaya))
                    u.pop("tambahan", None)
                    kini.append(u)
                    y += c + u["_h"]
                continue  # ukur ulang blok ini di halaman baru
            kini.append(dict(b, _h=h))
            y += celah + h
            if b["jenis"] == "subjudul":
                t = b.get("tingkat", 1)
                aktif = {k: v for k, v in aktif.items() if k < t}
                aktif[t] = b
            i += 1
        halaman.append(kini)
        return halaman

    @staticmethod
    def _nomor(b):
        """Nomor butir paragraf ('4. …' -> 4, 'B. …' -> 2), atau None."""
        m = re.match(r"^\(?(\d{1,2}|[A-Ea-e])[.)]\s", D.polos(b["teks"]).strip()) if b["jenis"] == "paragraf" else None
        return None if not m else int(m.group(1)) if m.group(1).isdigit() else ord(m.group(1).upper()) - 64

    def _yatim(self, halaman):
        """Jumlah batas halaman yang meninggalkan satu butir daftar bernomor sendirian."""
        n = 0
        for a, b in zip(halaman, halaman[1:]):
            b = [x for x in b if not x.get("ulang")]
            akhir, awal = (self._nomor(a[-1]), self._nomor(b[0])) if a and b else (None, None)
            if akhir and awal == akhir + 1:
                sisa1 = len(b) < 2 or self._nomor(b[1]) != awal + 1
                awal1 = len(a) < 2 or self._nomor(a[-2]) != akhir - 1
                n += sisa1 + awal1
        return n

    @staticmethod
    def _pembuka(b):
        """Judul, atau kalimat pengantar ('Contohnya:', 'Perhatikan kalimat berikut.'): harus sehalaman dengan isinya."""
        if b["jenis"] != "paragraf":
            return b["jenis"] == "subjudul"
        isi = D.polos(b["teks"]).strip()
        return len(isi) <= 110 and (isi.endswith(":") or bool(re.search(r"\b(berikut|berikut ini|di bawah ini)\.$", isi)))

    @staticmethod
    def _judul_ulang(aktif, b):
        """Judul yang diulang di puncak halaman sambungan: judul bagian, lalu subjudul terdekat bila isinya masih berlanjut."""
        t = b.get("tingkat", 1) if b["jenis"] == "subjudul" else 9
        hasil = [aktif[0]] if 0 in aktif and t > 0 else []
        dalam = [k for k in aktif if k > 0]
        if t == 9 and dalam:
            hasil.append(aktif[max(dalam)])
        return hasil

    def _tinggi_isi(self, hal, jarak):
        return sum(b["_h"] + self._celah(a, b, jarak) for a, b in zip([None] + hal[:-1], hal))

    def _halaman(self, blok, y0, gaya, jarak, ulang, rapikan):
        """(halaman, gaya, jarak): pembagian halaman yang rapat, seimbang, dan tidak menyisakan halaman nyaris kosong."""
        B = D.BATAS_BAWAH
        ruang = B - y0
        hal = self._rencana(blok, y0, B, gaya, jarak, ulang)
        if not rapikan:
            return hal, gaya, jarak
        potong = lambda r: sum(not b.get("ulang") for h in r for b in h)  # noqa: E731
        gbr = lambda r: [b["_h"] for h in r for b in h if b["jenis"] == "gambar"]  # noqa: E731
        if len(hal) > 1:  # halaman terakhir nyaris kosong: rapatkan jarak atau kecilkan gambar supaya hemat satu halaman
            sisa = self._tinggi_isi(hal[-1], jarak) / ruang
            coba2 = [(1, r) for r in (0.8, 0.6, 0.45)] if sisa < 0.25 else []
            if gbr(hal) and sisa < 0.36:
                coba2 += [(s, r) for r in (0.8, 0.6) for s in (0.85, 0.75, 0.65)]
            for skala, r in coba2:
                isi = [dict(b, lebar=b.get("lebar", 4) * skala, tinggi=b.get("tinggi", 3) * skala)
                       if b["jenis"] == "gambar" and b.get("lebar", 4) * skala >= 2.1 else b for b in blok]
                coba = self._rencana(isi, y0, B, gaya, jarak * r, ulang)
                if len(coba) < len(hal):
                    blok, hal, jarak = isi, coba, jarak * r
                    break
        if len(hal) > 1:
            # bagi rata: batas halaman terendah yang jumlah halamannya tetap. Daftar bernomor yang menyisakan satu
            # butir di halaman lain dihindari, bila perlu dengan merapatkan jarak sedikit.
            terbaik = (self._yatim(hal), hal, jarak)
            for r in (1, 0.85, 0.7, 0.55):
                rapat = []
                for k in range(10, 21):
                    coba = self._rencana(blok, y0, y0 + ruang * k / 20, gaya, jarak * r, ulang)
                    if len(coba) == len(hal) and potong(coba) <= potong(hal) and all(
                            a >= 0.97 * b for a, b in zip(gbr(coba), gbr(hal))) and all(
                            self._tinggi_isi(h, jarak * r) <= ruang + 0.01 for h in coba):
                        rapat.append((self._yatim(coba), coba, jarak * r))
                if rapat:
                    pilih = min(rapat, key=lambda x: x[0])
                    if pilih[0] < terbaik[0] or r == 1 and pilih[0] == terbaik[0]:
                        terbaik = pilih
                if not terbaik[0]:
                    break
            _, hal, jarak = terbaik
        elif gaya is None and self._tinggi_isi(hal[0], jarak) < 0.5 * ruang:
            # isi sedikit: gambar diperbesar dan spasi dilonggarkan (ukuran huruf tetap)
            for skala in (1.4, 1.25, 1.0):
                besar = [dict(b, lebar=b.get("lebar", 4) * skala, tinggi=b.get("tinggi", 3) * skala) if b["jenis"] == "gambar" else b
                         for b in blok]
                coba = self._rencana(besar, y0, B, "lega", jarak * 1.5, ulang)
                if len(coba) == 1 and self._tinggi_isi(coba[0], jarak * 1.5) <= 0.85 * ruang:
                    return coba, "lega", jarak * 1.5
        return hal, gaya, jarak

    def _tata_blok(self, buat_slide, blok, gaya=None, nama="Blok", ulang=False, rapikan=True, y0=None):
        """Tata blok berurutan; buat_slide(lanjutan, sumber) -> (slide, y). Slide baru dibuat bila penuh.

        Paragraf, subjudul, dan butir yang berurutan menjadi satu kotak teks sehingga jarak antarparagrafnya seragam
        dan mudah disunting. ulang: judul yang isinya berlanjut diulang di halaman sambungan.
        """
        w = D.LEBAR_ISI
        y0 = self.Y_ISI if y0 is None else y0
        jarak = 0.12 if (gaya or "").startswith("soal") else D.JARAK_BLOK
        blok = [dict(b, tingkat=self._tingkat(b)) if b["jenis"] == "subjudul" else dict(b) for b in blok]
        halaman, gaya, jarak = self._halaman(blok, y0, gaya, jarak, ulang, rapikan)
        for n, hal in enumerate(halaman):
            s, y = buat_slide(n > 0, next((b["_sumber"] for b in hal if b.get("_sumber")), None))
            i = 0
            while i < len(hal):
                b = hal[i]
                y += self._celah(hal[i - 1] if i else None, b, jarak)
                if b["jenis"] in D.TEKS_ALIR:
                    isi = []
                    while i < len(hal) and hal[i]["jenis"] in D.TEKS_ALIR:
                        lanjut = i + 1 < len(hal) and hal[i + 1]["jenis"] in D.TEKS_ALIR
                        isi.append((hal[i], self._celah(hal[i], hal[i + 1], jarak) if lanjut else 0))
                        self.catat_tambahan(s, hal[i].get("tambahan"))
                        i += 1
                    y = D.alir_teks(s, isi, D.KIRI, y, w, gaya)
                    continue
                y = D.gambar_blok(s, b, D.KIRI, y, w, self.folder, gaya, D.BATAS_BAWAH - y)
                self.catat_tambahan(s, b.get("tambahan"))
                i += 1
            self.cek_batas(y, nama)
        return s, y

    def _tinggi_perlu(self, blok, i, w, gaya, jarak):
        """Tinggi terkecil agar blok[i] layak berada di suatu halaman: potongan pertama untuk blok yang bisa
        dipecah, ukuran terkecil untuk gambar, dan subjudul bersama isi di bawahnya."""
        b = blok[i]
        if b["jenis"] == "gambar":
            return D.skala_min_gambar(b, w)
        pecah = BISA_DIPECAH.get(b["jenis"])
        if pecah and len(b[pecah[0]]) > pecah[2]:
            b = dict(b, **{pecah[0]: b[pecah[0]][:pecah[2]]})
        h = D.tinggi_blok(b, w, self.folder, gaya)
        if i + 1 < len(blok) and (b["jenis"] == "subjudul" or self._pembuka(b) and blok[i + 1]["jenis"] != "subjudul"):
            perlu = self._tinggi_perlu(blok, i + 1, w, gaya, jarak)
            h += self._celah(b, blok[i + 1], jarak) + perlu
            if perlu < 0.5 and blok[i + 1]["jenis"] == "paragraf" and i + 2 < len(blok) and blok[i + 2]["jenis"] == "paragraf":
                h += jarak + self._tinggi_perlu(blok, i + 2, w, gaya, jarak)  # tidak ditinggal bersama satu butir pendek
        return h

    def _bagian_muat(self, b, kunci, langkah, w, sisa, gaya=None):
        """Jumlah isi terbanyak (kelipatan langkah) dari b[kunci] yang muat dalam tinggi sisa."""
        n = len(b[kunci])
        for k in range((n - 1) // langkah * langkah, 0, -langkah):
            if D.tinggi_blok(dict(b, **{kunci: b[kunci][:k]}), w, self.folder, gaya) <= sisa:
                return k
        return 0

    def _slide_blok(self, sub, kumpulan):
        """Slide 'blok' berurutan dalam satu subbab mengalir bersama. Nama subbab ada di pita atas; judul slide yang
        berbeda dari nama subbabnya menjadi judul bagian (tingkat 0) di dalam isi, jadi judul tidak tertulis dua kali."""
        sama = lambda a, b: re.sub(r"\W+", "", D.polos(a)).lower() == re.sub(r"\W+", "", D.polos(b)).lower()  # noqa: E731
        isi = []
        for sumber, sl in kumpulan:
            blok = [dict(b) for b in sl["blok"]]
            if blok and blok[0]["jenis"] == "subjudul" and sama(blok[0]["teks"], sl["judul"]):
                blok.pop(0)
            if sl.get("pengantar"):
                blok.insert(0, {"jenis": "paragraf", "teks": sl["pengantar"]})
            if not sama(sl["judul"], sub["subbab"]):
                blok.insert(0, {"jenis": "subjudul", "tingkat": 0, "teks": sl["judul"]})
            isi += [dict(b, _sumber=sumber) for b in blok]

        def buat(lanjutan, sumber=None):
            self.sumber = sumber or self.sumber
            s = self.baru("Pendalaman Materi")
            D.chip(s, sub["subbab"], sub.get("ikon", "checklist"))
            return s, self.Y_ISI

        self._tata_blok(buat, isi, nama=sub["subbab"], ulang=True)

    def _contoh_soal(self, sub, sl):
        """Slide soal (stem + opsi) lalu slide pembahasan; tombol menautkan keduanya."""
        judul = sl.get("judul") or "Contoh Soal"
        tag = [t for t in (sl.get("label"),) if t]
        soal_slide, bahas_slide = [], []

        def kepala_soal(lanjutan, _judul=None):
            s = self.baru("Pendalaman Materi")
            w_chip = D.chip(s, judul, sub.get("ikon", "checklist"))
            x = D.X_CHIP + w_chip + 0.25
            for t in tag:  # label panjang dikecilkan supaya tidak keluar tepi slide
                ruang = D.LEBAR_SLIDE - 0.3 - x
                ukuran = next((u for u in (13, 12, 11, 10, 9) if D.lebar_pt(t, u, True) / 72 + 0.5 <= ruang), None)
                if ukuran:
                    _, w = D.pil(s, x, D.Y_CHIP + (D.H_CHIP - 0.46) / 2, t, ukuran=ukuran)
                    x += w + 0.15
            soal_slide.append(s)
            return s, self.Y_ISI

        opsi = sl.get("opsi") or {}
        gaya, u_opsi = "soal", 18
        if sl.get("bacaan"):  # soal bacaan: teks, pertanyaan, dan opsi sebisa mungkin di satu slide
            for u in (20, 19, 18, 17, 16):
                h = D.Y_CHIP + D.H_CHIP + 0.8 + D.tinggi_opsi(opsi, ukuran=min(18, u)) + sum(
                    D.tinggi_blok(b, D.LEBAR_ISI, self.folder, f"soal:{u}") + 0.1 for b in sl["soal"])
                if h <= D.BATAS_BAWAH:
                    gaya, u_opsi = f"soal:{u}", min(18, u)
                    break
        s, y = self._tata_blok(kepala_soal, sl.get("soal", []), gaya=gaya, nama=judul, rapikan=False)
        if opsi:
            h = D.tinggi_opsi(opsi, ukuran=u_opsi)
            if y + 0.35 + h > D.BATAS_BAWAH and sl.get("soal"):
                s, y = kepala_soal(True)
                y -= 0.35
            y = D.gambar_opsi(s, y + 0.35, opsi, ukuran=u_opsi)
            self.cek_batas(y, judul)
        kunci = sl.get("kunci", "")
        blok = [b for b in sl.get("pembahasan", [])  # 'Kunci jawaban: C' sudah tertulis di pil 'Jawaban C'
                if not (kunci and b["jenis"] == "kunci" and re.fullmatch(rf"kunci(?: jawaban)?\s*:?\s*{kunci}\.?", D.polos(b["teks"]), re.I))]
        if kunci and opsi.get(kunci):
            blok.insert(0, {"jenis": "kunci", "teks": f"**Kunci {kunci}.** {opsi[kunci]}"})
        if blok:
            tombol_bahas = D.tombol(s, D.KIRI, 10.36, "Lihat pembahasan ›")

        def kepala_bahas(lanjutan, _judul=None):
            s = self.baru("Pendalaman Materi")
            w_chip = D.chip(s, f"Pembahasan {judul}", "strategi")
            if kunci:
                D.pil(s, D.X_CHIP + w_chip + 0.25, D.Y_CHIP + (D.H_CHIP - 0.56) / 2, f"Jawaban {kunci}",
                      ukuran=18, warna="kuning", warna_teks="di_kuning", h=0.56, nama="Jawaban")
            kembali = D.tombol(s, D.KIRI, 10.36, "‹ Kembali ke soal")
            kembali.click_action.target_slide = soal_slide[0]
            for t in (x for x in (sl.get("label_bahas"),) if x):
                D.pil(s, D.X_CHIP + w_chip + (2.6 if kunci else 0.25), D.Y_CHIP + (D.H_CHIP - 0.46) / 2, t, ukuran=13)
            bahas_slide.append(s)
            return s, self.Y_ISI

        if blok:
            self._tata_blok(kepala_bahas, blok, nama=f"Pembahasan {judul}")
            tombol_bahas.click_action.target_slide = bahas_slide[0]

    def _slide_materi(self, sub, sl):
        s = self.baru("Pendalaman Materi")
        judul = None if D.polos(sl["judul"]).strip().lower() == D.polos(sub["subbab"]).strip().lower() else sl["judul"]
        y = D.kerangka_isi(s, sub["subbab"], sub.get("ikon", "checklist"), judul, sl.get("pengantar"))
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
            ringkasan, pesan = p.get("ringkasan", []), p.get("pesan")
            # pesan yang panjang dipindah ke slide penutup kedua supaya semua butir ringkasan tetap tampil
            muat = not (ringkasan and pesan) or D._KAIT.susun_penutup(ringkasan, pesan)
            for r, ps in [(ringkasan, pesan)] if muat else [(ringkasan, None), ([], pesan)]:
                s = self.baru("Penutup", hias=False, bingkai=True)
                self.cek_batas(D._KAIT.penutup(s, r, ps), "Penutup")
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
            if self._bagian_dilewati(sumber):
                continue
            self.sumber = sumber
            langkah()
            if sumber in ("pendahuluan", "tujuan", "penutup"):
                self.catat_tambahan(self.prs.slides[len(self.prs.slides) - 1], (self.k.get(sumber) or {}).get("tambahan"))
        self.selesaikan()

    def catat_tambahan(self, slide, isi):
        """Tandai isi di luar highlight di catatan pembicara, supaya tim bisa meninjau dan mencoretnya."""
        if not isi:
            return
        tf = slide.notes_slide.notes_text_frame
        baris = [p.text for p in tf.paragraphs if p.text] or ["TAMBAHAN DI LUAR HIGHLIGHT (boleh dicoret tim akademik):"]
        for t in isi:
            teks = D.polos(t["isi"])
            butir = f"- {teks if len(teks) <= 160 else teks[:157] + '...'} [{t['alasan']}]"
            if butir not in baris or t["isi"] == "Gambar":
                baris.append(butir)
        tf.text = "\n".join(baris)

    def _bagian_dilewati(self, sumber):
        """Deck contoh soal (meta.bagian 'soal') dan naskah dari PDF (meta.lewati_kosong) tidak punya
        pendahuluan, tujuan, dan penutup sendiri; slidenya dilewati selama masih kosong."""
        meta = self.k["meta"]
        if not (meta.get("bagian") == "soal" or meta.get("lewati_kosong")) \
                or sumber not in ("pendahuluan", "tujuan", "penutup"):
            return False
        isi = self.k.get(sumber) or {}
        return not any(isi.get(x) for x in ("paragraf", "pemantik", "gambar", "butir", "ringkasan", "pesan"))


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


def jenis_deck(naskah):
    """'soal' untuk deck contoh soal, selain itu 'materi' (dipakai saringan di situs dan dasbor)."""
    meta, slide = naskah.get("meta") or {}, [s for sub in naskah.get("materi") or [] for s in sub.get("slide", [])]
    semua_soal = bool(slide) and all(s.get("tipe") == "contoh_soal" for s in slide)
    return "soal" if meta.get("bagian") == "soal" or meta.get("judul", "").startswith("Contoh Soal") or semua_soal else "materi"


def buat_deck(konten, path_template=None, folder=None, tema="opsi-b", palet=None):
    """Bangun deck dari naskah (dict); folder = dasar path gambar di naskah. Hasil: (Presentation, Pembangun).

    palet: nama palet platform (mis. 'jadiasn',
    lihat palet.py) untuk mewarnai ulang navy/kuning template; None = warna template.

    Tema disimpan sebagai keadaan modul desain, jadi pemanggil yang berjalan paralel (server
    web) harus membangun deck satu per satu.
    """
    prs = siapkan_presentasi(path_template or TEMPLATE_TEMA.get(tema, TEMPLATE), tema)
    if palet:
        import palet as PL
        PL.pakai_token(palet)  # token kontras palet berlaku selama deck dibangun
    b = Pembangun(prs, konten, folder)
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
    ap.add_argument("--palet", default=None, help="Palet platform untuk warna slide, mis. jadiasn")
    arg = ap.parse_args(argv)

    with open(arg.naskah, encoding="utf-8") as f:
        konten = json.load(f)
    nama = nama_file(konten) + (" (General)" if arg.tema == "general" else "")
    os.makedirs(arg.keluar, exist_ok=True)

    try:
        prs, b = buat_deck(konten, arg.template, os.path.dirname(os.path.abspath(arg.naskah)), arg.tema,
                           palet=arg.palet)
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
