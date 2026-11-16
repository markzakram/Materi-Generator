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
from desain import P  # noqa: E402

AKAR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TEMPLATE = os.path.join(AKAR, "Template Desain Materi.pptx")

TEMA_WARNA = {
    "dk1": "000000", "lt1": "FFFFFF", "dk2": "004B66", "lt2": "EAF6FB",
    "accent1": "00AEEF", "accent2": "0A6486", "accent3": "364D61", "accent4": "CFE2BE",
    "accent5": "004B66", "accent6": "F7A93B", "hlink": "0A6486", "folHlink": "364D61",
}


# ------------------------------------------------------ persiapan file ---
class TemplateTidakAda(FileNotFoundError):
    pass


class TemplateTidakCocok(ValueError):
    pass


def siapkan_presentasi(path_template):
    if not path_template or not os.path.exists(path_template):
        raise TemplateTidakAda(
            f"Template tidak ditemukan: {path_template}\n"
            "Salin 'Template Desain Materi.pptx' ke folder akar proyek atau pakai --template.")
    prs = Presentation(path_template)
    lebar, tinggi = prs.slide_width / 914400, prs.slide_height / 914400
    if abs(lebar - D.LEBAR_SLIDE) > 0.01 or abs(tinggi - D.TINGGI_SLIDE) > 0.01:
        raise TemplateTidakCocok(
            f"Ukuran slide template {lebar:g} x {tinggi:g} inci, sedangkan generator ditata untuk "
            f"{D.LEBAR_SLIDE:g} x {D.TINGGI_SLIDE:g} inci. Pakai 'Template Desain Materi.pptx' yang benar.")
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

    Template menamai font tebal sebagai keluarga 'Poppins Bold', padahal nama
    internalnya 'Poppins' gaya Bold, sehingga PowerPoint memakai font pengganti.
    """
    lst = prs.part._element.find(qn("p:embeddedFontLst"))
    if lst is None:
        return
    peta = {}
    for ef in lst.findall(qn("p:embeddedFont")):
        nama = ef.find(qn("p:font")).get("typeface")
        reg = ef.find(qn("p:regular"))
        peta[nama] = reg.get(qn("r:id")) if reg is not None else None
    reg, tebal = peta.get("Poppins"), peta.get("Poppins Bold")
    for ef in list(lst):
        lst.remove(ef)
    if reg:
        ef = etree.SubElement(lst, qn("p:embeddedFont"))
        etree.SubElement(ef, qn("p:font"), typeface="Poppins", panose="00000500000000000000", charset="1")
        etree.SubElement(ef, qn("p:regular")).set(qn("r:id"), reg)
        if tebal:
            etree.SubElement(ef, qn("p:bold")).set(qn("r:id"), tebal)
    for rid in peta.values():
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
class Pembangun:
    def __init__(self, prs, konten):
        self.prs, self.k = prs, konten
        self.layout = next((l for l in prs.slide_layouts if l.name == "Blank"), prs.slide_layouts[-1])
        self.slides = []
        self.awal_bagian = {}
        self.peringatan = []

    # kerangka -----------------------------------------------------------
    def baru(self, bagian=None, nav=True):
        s = self.prs.slides.add_slide(self.layout)
        for ph in list(s.placeholders):
            ph._element.getparent().remove(ph._element)
        tombol = D.navbar(s, bagian) if nav else None
        self.slides.append((s, tombol))
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
            D.nomor_halaman(s, i, total)
            for label, shp in tombol.items():
                tujuan = self.awal_bagian.get(label)
                if tujuan is not None and tujuan is not s:
                    shp.click_action.target_slide = tujuan

    # slide khusus ---------------------------------------------------------
    def _judul_besar(self, s, isi, y, h, maks, lebar=D.LEBAR_SLIDE - 2.5):
        ukuran = maks
        while D.lebar_segs(D.segmen(isi, True), ukuran) > lebar * 0.95 and ukuran > 36:
            ukuran -= 2
        D.teks(s, (D.LEBAR_SLIDE - lebar) / 2, y, lebar, h,
               [P(isi, ukuran, tebal=True, rata="c", spasi=1.05, gradien=D.GRADIEN_BRAND)],
               jangkar="m", nama="Judul besar")

    def _latar(self, s):
        foto = D.LATAR_COVER
        if not foto or not os.path.exists(foto):  # foto stok tidak ikut di repo; sampul tetap jadi tanpa foto
            return
        pic = D.gambar(s, foto, 0, 0, w=D.LEBAR_SLIDE, h=D.TINGGI_SLIDE, nama="Latar")
        D.opasitas_gambar(pic, 31)

    def sampul(self):
        m = self.k["meta"]
        s = self.baru(nav=False)
        self._latar(s)
        self._judul_besar(s, m["judul"].upper(), 4.15, 1.5, 80)
        ukuran = 24
        w = D.lebar_pt(m["subtes"], ukuran, True) / 72 + 1.0
        pil = D.kotak(s, (D.LEBAR_SLIDE - w) / 2, 5.85, w, 0.66, warna="slate", radius=0.12, nama="Nama subtes")
        D.isi_teks(pil.text_frame, [P(m["subtes"], ukuran, "putih", tebal=True, rata="c", spasi=1.1)], jangkar="m")

    def penutup_akhir(self):
        s = self.baru(nav=False)
        self._latar(s)
        self._judul_besar(s, "TERIMA KASIH", 4.85, 1.5, 80)

    def peta_subtes(self):
        p = self.k["peta_subtes"]
        s = self.baru()
        self._judul_besar(s, p["judul"], 1.75, 1.4, 76)
        D.teks(s, 2.0, 3.2, D.LEBAR_SLIDE - 4.0, 0.6,
               [P(p["subjudul"], 28, tebal=True, rata="c", spasi=1.1, gradien=D.GRADIEN_BRAND)],
               jangkar="m", nama="Subjudul")
        kartu = p["kartu"]
        n = len(kartu)
        total_w, celah = D.LEBAR_SLIDE - 1.2, 0.2
        w = (total_w - celah * (n - 1)) / n
        x0, y, h = (D.LEBAR_SLIDE - total_w) / 2, 4.75, 3.6
        sisi = (0.2, 0, 0.2, 0)
        judul_k = [[P(k["judul"], 17, "putih", tebal=True, rata="c", spasi=1.2)] for k in kartu]
        h_judul = max(D.tinggi_paragraf(j, w, sisi) for j in judul_k)
        for i, k in enumerate(kartu):
            x = x0 + i * (w + celah)
            aktif = i == p.get("aktif", 0)
            if aktif:
                D.kotak(s, x, y, w, h, gradien=D.GRADIEN_BRAND, radius=0.12, nama=f"Kartu materi {i + 1}")
            else:
                D.kotak(s, x, y, w, h, warna="slate", radius=0.12, nama=f"Kartu materi {i + 1}")
            D.teks(s, x, y + 0.35, w, h_judul, judul_k[i], inset=sisi, jangkar="m")
            D.teks(s, x, y + 0.35 + h_judul + 0.15, w, 1.3,
                   [P(k["deskripsi"], 13, "putih" if aktif else "redup", rata="c", spasi=1.35)], inset=sisi)
            ukuran = 0.64
            D.ubin_ikon(s, x + (w - ukuran) / 2, y + h - 1.1, ukuran, k.get("ikon", "globe"))

    def pendahuluan(self):
        p = self.k["pendahuluan"]
        s = self.baru("Pendahuluan")
        D.chip(s, "Pendahuluan", p.get("ikon_chip", "globe"))
        y = D.Y_CHIP + D.H_CHIP + 0.5
        par = [P(t, 22, "teks", spasi=1.4, setelah=16) for t in p["paragraf"]]
        w_ikon = 3.6
        x_ikon = D.KANAN - w_ikon - 0.3
        w = x_ikon - 1.0 - D.KIRI
        h = D.tinggi_paragraf(par, w)
        D.teks(s, D.KIRI, y, w, h, par, nama="Paragraf pendahuluan")
        ikon = D.gambar(s, D.ikon_path(p.get("ikon", "presenter")), x_ikon, y + 0.2, w=w_ikon, nama="Ilustrasi")
        bawah = max(y + h, (ikon.top + ikon.height) / 914400)
        y = D.catatan(s, bawah + 0.45, "pemantik", p["pemantik"], judul_catatan="Pertanyaan pemantik")
        self.cek_batas(y, "Pendahuluan")

    def tujuan(self):
        t = self.k["tujuan"]
        s = self.baru("Tujuan Pembelajaran")
        D.chip(s, "Tujuan Pembelajaran", t.get("ikon_chip", "strategi"))
        y = D.pengantar(s, t["pengantar"], D.Y_CHIP + D.H_CHIP + 0.5) + 0.3
        for b in t["butir"]:
            segs = [(b["kko"], True, False, "teal")] + D.segmen(" " + b["teks"])
            par = [P(None, 22, "teks", spasi=1.3, segs=segs)]
            inset = (1.85, 0.24, 0.4, 0.24)
            h = max(1.0, D.tinggi_paragraf(par, D.LEBAR_ISI, inset))
            D.kotak(s, D.KIRI, y, D.LEBAR_ISI, h, warna="tint", radius=0.1, nama="Butir tujuan")
            lencana = D.kotak(s, D.KIRI + 0.35, y + (h - 0.58) / 2, 1.15, 0.58, gradien=D.GRADIEN_BRAND,
                              radius=0.29, nama=f"Level {b['level']}")
            D.isi_teks(lencana.text_frame, [P(b["level"], 20, "putih", tebal=True, rata="c", spasi=1.0)],
                       jangkar="m")
            D.teks(s, D.KIRI, y, D.LEBAR_ISI, h, par, inset=inset, jangkar="m")
            y += h + 0.22
        self.cek_batas(y - 0.22, "Tujuan")

    def materi(self):
        for sub in self.k["materi"]:
            for sl in sub["slide"]:
                self._slide_materi(sub, sl)

    def _slide_materi(self, sub, sl):
        s = self.baru("Materi")
        y = D.kerangka_isi(s, sub["subbab"], sub.get("ikon", "checklist"), sl["judul"], sl.get("pengantar"))
        tipe = sl["tipe"]
        if tipe == "tabel":
            y = D.tabel(s, y, sl["kolom"], sl["baris"], sl.get("lebar"), sl.get("ukuran", 20))
        elif tipe == "tabel_kebenaran":
            w = 12.5
            x = (D.LEBAR_SLIDE - w) / 2
            y = D.tabel(s, y, sl["kolom"], sl["baris"], x=x, w=w, tengah=True,
                        sorot=sl.get("sorot"), kebenaran=True, ukuran=sl.get("ukuran", 20))
            if sl.get("keterangan"):
                D.teks(s, x, y + 0.1, w, 0.4, [P(sl["keterangan"], 16, "teks", rata="r", spasi=1.2)],
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
            w_ikon = 3.2
            x_ikon = D.KANAN - w_ikon - 0.3
            y = D.poin_berbutir(s, y, sl["poin"], w=x_ikon - 1.0 - D.KIRI)
            ikon = D.gambar(s, D.ikon_path(sl.get("ikon", "bingung")), x_ikon, y0 - 0.1, w=w_ikon,
                            nama="Ilustrasi")
            y = max(y, (ikon.top + ikon.height) / 914400)
        else:
            raise ValueError(f"Tipe slide tidak dikenal: {tipe}")
        if sl.get("catatan"):
            c = sl["catatan"]
            y = D.catatan(s, y + D.JARAK, c["jenis"], c["teks"])
        self.cek_batas(y, sl["judul"])

    def penutup(self):
        p = self.k["penutup"]
        s = self.baru("Penutup")
        y = D.kerangka_isi(s, "Penutup", p.get("ikon_chip", "roket"), p.get("judul", "Ringkasan"))
        item = [dict(r, nomor=i) for i, r in enumerate(p["ringkasan"], 1)]
        item.append({"pesan": p["pesan"], "warna": "gradien"})

        def pembuat(k, lw):
            if "pesan" in k:
                return [P("Pesan penutup", 16, "sage", tebal=True, spasi=1.2, setelah=6),
                        P(k["pesan"], 20, "putih", tebal=True, spasi=1.35)]
            segs = [(f"{k['nomor']}  ", True, False, "sage"), (k["judul"], True, False, "putih")]
            return [P(None, 22, segs=segs, spasi=1.2, setelah=8), P(k["teks"], 19, "putih", spasi=1.35)]

        y = D.grid_kartu(s, y, item, kolom=3, pembuat=pembuat, inset=(0.32, 0.28, 0.32, 0.28),
                         warna_list=["slate", "teal", "slate", "teal", "slate"])
        self.cek_batas(y, "Penutup")

    def bangun(self):
        self.sampul()
        self.peta_subtes()
        self.pendahuluan()
        self.tujuan()
        self.materi()
        self.penutup()
        self.penutup_akhir()
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


def buat_deck(konten, path_template=TEMPLATE):
    """Bangun deck dari naskah (dict). Hasil: (Presentation, Pembangun)."""
    prs = siapkan_presentasi(path_template)
    b = Pembangun(prs, konten)
    b.bangun()
    return prs, b


def main(argv=None):
    ap = argparse.ArgumentParser(description="Buat materi PPTX/PDF dari naskah JSON.")
    ap.add_argument("naskah", help="Path file naskah .json")
    ap.add_argument("--format", choices=["pptx", "pdf", "keduanya"], default="pptx")
    ap.add_argument("--keluar", default=os.path.join(AKAR, "output"), help="Folder hasil")
    ap.add_argument("--template", default=TEMPLATE)
    arg = ap.parse_args(argv)

    with open(arg.naskah, encoding="utf-8") as f:
        konten = json.load(f)
    nama = nama_file(konten)
    os.makedirs(arg.keluar, exist_ok=True)

    try:
        prs, b = buat_deck(konten, arg.template)
    except (TemplateTidakAda, TemplateTidakCocok) as e:
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
