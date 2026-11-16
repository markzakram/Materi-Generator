"""Sistem desain materi Cerebrum: token, pengukur teks, dan komponen slide.

Semua ukuran dalam inci pada kanvas 16 x 11,25 inci (ukuran template asli).
Teks memakai Poppins (tertanam di file), simbol logika/rumus memakai Cambria Math
karena Poppins tidak punya glyph seperti → ∨ ∧ ≡.

Markup teks kaya yang didukung di naskah:
  **tebal**     -> Poppins Bold
  $p → q$       -> Cambria Math (tidak dipotong saat pindah baris)
"""
from __future__ import annotations

import json
import os
import re
from functools import lru_cache

from lxml import etree
from PIL import Image, ImageDraw, ImageFont
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

ASET = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets")
# Foto latar sampul (tidak ikut di repo). Aplikasi web boleh menimpanya dengan
# path hasil unduhan; jika berkasnya tidak ada, sampul dibuat tanpa foto.
LATAR_COVER = os.path.join(ASET, "latar_cover.png")

# ---------------------------------------------------------------- token ---
LEBAR_SLIDE, TINGGI_SLIDE = 16.0, 11.25

WARNA = {
    "cyan": "00AEEF",
    "deep": "004B66",
    "teal": "0A6486",
    "slate": "364D61",
    "sage": "CFE2BE",
    "putih": "FFFFFF",
    "teks": "364D61",
    "tint": "EAF6FB",
    "tint_amber": "FFF4E0",
    "garis": "D3E0E8",
    "sorot": "FFE3A8",
    "redup": "DCE8EF",
}

FONT = "Poppins"
FONT_RUMUS = "Cambria Math"

KIRI = 0.62
LEBAR_ISI = LEBAR_SLIDE - 2 * KIRI  # 14.76
BATAS_BAWAH = 9.9
JARAK = 0.35

Y_CHIP, H_CHIP = 0.85, 0.86
Y_JUDUL = 1.98

IKON = {
    "globe": "ikon_globe.png",
    "strategi": "ikon_strategi.png",
    "grafik": "ikon_grafik.png",
    "checklist": "ikon_checklist.png",
    "roket": "ikon_roket.png",
    "peringatan": "ikon_peringatan.png",
    "presenter": "ikon_presenter.png",
    "bingung": "ikon_bingung.png",
}

GRADIEN_BRAND = [(0, "00AEEF", 100), (100, "004B66", 100)]
GRADIEN_LEMBUT = [(0, "00AEEF", 67), (100, "CFE2BE", 67)]
GRADIEN_TEKS = [(0, "0A6486", 100), (100, "004B66", 100)]


def ikon_path(nama: str) -> str:
    return os.path.join(ASET, IKON.get(nama, IKON["globe"]))


# ------------------------------------------------------- pengukur teks ---
# Layout dasar dipaksa agar hasil ukur sama di Windows dan Linux (Pillow di
# Linux bisa memakai raqm yang menerapkan kerning GPOS).
_LAYOUT = ImageFont.Layout.BASIC
_FONT_FILE = {
    False: os.path.join(ASET, "fonts", "Poppins-Regular.ttf"),
    True: os.path.join(ASET, "fonts", "Poppins-Bold.ttf"),
}
# Cambria Math tidak boleh ikut di repo dan tidak ada di server Linux, jadi yang
# dipakai tabel lebarnya (lihat buat_metrik_rumus.py).
_METRIK_RUMUS = os.path.join(ASET, "metrik_cambria_math.json")


@lru_cache(maxsize=None)
def _ft(bold: bool):
    return ImageFont.truetype(_FONT_FILE[bold], 400, layout_engine=_LAYOUT)


@lru_cache(maxsize=None)
def _metrik_rumus():
    with open(_METRIK_RUMUS, encoding="utf-8") as f:
        m = json.load(f)
    return m["lebar"], m["kern"], m["notdef"], m["ukuran"]


def _lebar_rumus(teks: str) -> float:
    """Lebar teks Cambria Math dalam px pada ukuran 400 (sama dengan Pillow)."""
    lebar, kern, notdef, ukuran = _metrik_rumus()
    w = sum(lebar.get(ch, notdef) for ch in teks)
    w += sum(kern.get(a + b, 0.0) for a, b in zip(teks, teks[1:]))
    return w * 400.0 / ukuran


@lru_cache(maxsize=None)
def ada_glyph_poppins(ch: str) -> bool:
    """True jika Poppins punya glyph untuk karakter ini."""
    if ch.isspace() or ord(ch) < 0x250:
        return True
    ft = ImageFont.truetype(_FONT_FILE[False], 40, layout_engine=_LAYOUT)

    def gambar(c):
        im = Image.new("L", (60, 60), 0)
        ImageDraw.Draw(im).text((5, 0), c, font=ft, fill=255)
        return im.tobytes()

    return gambar(ch) != gambar("\ue000")


def lebar_pt(teks: str, ukuran: float, tebal=False, rumus=False) -> float:
    if not teks:
        return 0.0
    w = _lebar_rumus(teks) if rumus else _ft(bool(tebal)).getlength(teks)
    return w * ukuran / 400.0


def kutip_lengkung(teks: str) -> str:
    hasil, buka = [], True
    for ch in teks:
        if ch == '"':
            hasil.append("\u201c" if buka else "\u201d")
            buka = not buka
        else:
            hasil.append(ch)
    return "".join(hasil)


def segmen(teks: str, tebal=False, warna=None):
    """Ubah markup menjadi daftar segmen (teks, tebal, rumus, warna)."""
    keluar = []
    for bagian in re.split(r"(\$[^$]+\$)", teks):
        if not bagian:
            continue
        if bagian.startswith("$") and bagian.endswith("$"):
            isi = bagian[1:-1].replace(" ", "\u00a0")
            keluar.append((isi, False, True, warna))
            continue
        bagian = kutip_lengkung(bagian)
        for i, potong in enumerate(re.split(r"\*\*", bagian)):
            if not potong:
                continue
            b = tebal or (i % 2 == 1)
            # karakter yang tak ada di Poppins dialihkan ke Cambria Math
            buf, buf_rumus = "", False
            for ch in potong:
                r = not ada_glyph_poppins(ch)
                if buf and r != buf_rumus:
                    keluar.append((buf, b and not buf_rumus, buf_rumus, warna))
                    buf = ""
                buf += ch
                buf_rumus = r
            if buf:
                keluar.append((buf, b and not buf_rumus, buf_rumus, warna))
    return keluar


SKALA_RUMUS = 1.15  # Cambria Math tampak lebih kecil dari Poppins pada ukuran pt yang sama


def campuran(segs) -> bool:
    return any(not r for _, _, r, _ in segs)


def ukuran_run(ukuran, rumus, campur):
    return round(ukuran * SKALA_RUMUS, 1) if (rumus and campur) else ukuran


def jumlah_baris(segs, ukuran: float, lebar_in: float) -> int:
    """Simulasi pindah baris ala PowerPoint (greedy, putus di spasi)."""
    batas = lebar_in * 72 * 0.965
    campur = campuran(segs)
    kata, kini = [], []
    for teks, b, r, _ in segs:
        if r:
            kini.append((teks, b, r))
            continue
        for bag in re.split(r"( +)", teks):
            if not bag:
                continue
            if bag.strip(" ") == "":
                if kini:
                    kata.append(kini)
                kini = []
                kata.append(None)  # penanda spasi
            else:
                kini.append((bag, b, r))
    if kini:
        kata.append(kini)

    spasi = lebar_pt(" ", ukuran)
    baris, isi, ada = 1, 0.0, False
    tunda_spasi = False
    for k in kata:
        if k is None:
            tunda_spasi = ada
            continue
        w = sum(lebar_pt(t, ukuran_run(ukuran, r, campur), b, r) for t, b, r in k)
        tambah = (spasi if tunda_spasi else 0) + w
        if ada and isi + tambah > batas:
            baris += 1
            isi, ada = w, True
        else:
            isi += tambah
            ada = True
        tunda_spasi = False
    return baris


def lebar_segs(segs, ukuran):
    campur = campuran(segs)
    return sum(lebar_pt(t, ukuran_run(ukuran, r, campur), b, r) for t, b, r, _ in segs) / 72.0


# --------------------------------------------------------- paragraf ---
class P:
    """Spesifikasi satu paragraf teks kaya."""

    def __init__(self, teks, ukuran=20, warna="teks", tebal=False, rumus=False,
                 rata="l", spasi=1.3, setelah=0, segs=None, gradien=None):
        self.ukuran, self.warna, self.rata = ukuran, warna, rata
        self.spasi, self.setelah, self.gradien = spasi, setelah, gradien
        if segs is not None:
            self.segs = segs
        elif rumus:
            self.segs = [(teks.replace(" ", "\u00a0"), False, True, None)]
        else:
            self.segs = segmen(teks, tebal=tebal)

    def tinggi_pt(self, lebar_in):
        return jumlah_baris(self.segs, self.ukuran, lebar_in) * self.ukuran * self.spasi


def tinggi_paragraf(paragraf, lebar_in, inset=(0, 0, 0, 0)):
    """Tinggi (inci) blok paragraf; inset = (kiri, atas, kanan, bawah)."""
    lebar = lebar_in - inset[0] - inset[2]
    total = 0.0
    for i, p in enumerate(paragraf):
        total += p.tinggi_pt(lebar)
        if i < len(paragraf) - 1:
            total += p.setelah
    return total / 72.0 + inset[1] + inset[3]


def _rgb(warna):
    return RGBColor.from_string(WARNA.get(warna, warna))


def _gradfill_xml(stops, sudut=45):
    gs = "".join(
        f'<a:gs pos="{pos * 1000}"><a:srgbClr val="{WARNA.get(c, c)}">'
        f'<a:alpha val="{a * 1000}"/></a:srgbClr></a:gs>'
        for pos, c, a in stops
    )
    return etree.fromstring(
        '<a:gradFill xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        f'rotWithShape="1"><a:gsLst>{gs}</a:gsLst><a:lin ang="{int(sudut * 60000)}" scaled="0"/></a:gradFill>'
    )


def isi_teks(tf, paragraf, inset=(0, 0, 0, 0), jangkar="t"):
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.margin_left, tf.margin_top = Inches(inset[0]), Inches(inset[1])
    tf.margin_right, tf.margin_bottom = Inches(inset[2]), Inches(inset[3])
    tf.vertical_anchor = {"t": MSO_ANCHOR.TOP, "m": MSO_ANCHOR.MIDDLE, "b": MSO_ANCHOR.BOTTOM}[jangkar]
    rata = {"l": PP_ALIGN.LEFT, "c": PP_ALIGN.CENTER, "r": PP_ALIGN.RIGHT}
    for i, spec in enumerate(paragraf):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = rata[spec.rata]
        p.line_spacing = Pt(round(spec.ukuran * spec.spasi, 1))
        p.space_before = Pt(0)
        p.space_after = Pt(spec.setelah if i < len(paragraf) - 1 else 0)
        campur = campuran(spec.segs)
        for teks, tebal, rumus, warna in spec.segs:
            r = p.add_run()
            r.text = teks
            f = r.font
            f.size = Pt(ukuran_run(spec.ukuran, rumus, campur))
            f.bold = bool(tebal)
            f.name = FONT_RUMUS if rumus else FONT
            if spec.gradien:
                rpr = r._r.get_or_add_rPr()
                rpr.insert(0, _gradfill_xml(spec.gradien))
            else:
                f.color.rgb = _rgb(warna or spec.warna)


def teks(slide, x, y, w, h, paragraf, inset=(0, 0, 0, 0), jangkar="t", nama=None):
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    if nama:
        tb.name = nama
    isi_teks(tb.text_frame, paragraf, inset, jangkar)
    return tb


# ------------------------------------------------------------ bentuk ---
def _hapus_gaya(shp):
    st = shp._element.find(qn("p:style"))
    if st is not None:
        shp._element.remove(st)


def _spPr(shp):
    return shp._element.spPr


def _ganti_fill(shp, el):
    sp = _spPr(shp)
    for tag in ("a:noFill", "a:solidFill", "a:gradFill", "a:blipFill", "a:pattFill"):
        for old in sp.findall(qn(tag)):
            sp.remove(old)
    geom = sp.find(qn("a:prstGeom"))
    if geom is None:
        geom = sp.find(qn("a:custGeom"))
    geom.addnext(el)


def bayangan(shp, alpha=22, blur=0.09, jarak=0.03):
    sp = _spPr(shp)
    for old in sp.findall(qn("a:effectLst")):
        sp.remove(old)
    eff = etree.fromstring(
        '<a:effectLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        f'<a:outerShdw blurRad="{int(blur * 914400)}" dist="{int(jarak * 914400)}" dir="5400000" '
        'algn="t" rotWithShape="0"><a:srgbClr val="004B66">'
        f'<a:alpha val="{alpha * 1000}"/></a:srgbClr></a:outerShdw></a:effectLst>'
    )
    ln = sp.find(qn("a:ln"))
    (ln if ln is not None else sp[-1]).addnext(eff)


def kotak(slide, x, y, w, h, warna=None, gradien=None, sudut=45, radius=0.08,
          garis=None, tebal_garis=1.0, bentuk=None, nama=None):
    jenis = bentuk or (MSO_SHAPE.ROUNDED_RECTANGLE if radius else MSO_SHAPE.RECTANGLE)
    shp = slide.shapes.add_shape(jenis, Inches(x), Inches(y), Inches(w), Inches(h))
    _hapus_gaya(shp)
    if nama:
        shp.name = nama
    if jenis == MSO_SHAPE.ROUNDED_RECTANGLE:
        shp.adjustments[0] = min(0.5, radius / min(w, h))
    if gradien:
        _ganti_fill(shp, _gradfill_xml(gradien, sudut))
    elif warna:
        shp.fill.solid()
        shp.fill.fore_color.rgb = _rgb(warna)
    else:
        shp.fill.background()
    if garis:
        shp.line.color.rgb = _rgb(garis)
        shp.line.width = Pt(tebal_garis)
    else:
        shp.line.fill.background()
    return shp


def gambar(slide, path, x, y, w=None, h=None, nama=None):
    with Image.open(path) as im:
        rasio = im.height / im.width
    if w is not None and h is None:
        h = w * rasio
    elif h is not None and w is None:
        w = h / rasio
    pic = slide.shapes.add_picture(path, Inches(x), Inches(y), Inches(w), Inches(h))
    if nama:
        pic.name = nama
    return pic


def opasitas_gambar(pic, persen):
    """Atur opasitas gambar (0-100), seperti latar sampul template (31%)."""
    blip = pic._element.find(".//" + qn("a:blip"))
    etree.SubElement(blip, qn("a:alphaModFix"), amt=str(int(persen * 1000)))


def ubin_ikon(slide, x, y, ukuran, ikon, besar_ikon=0.55):
    """Ubin putih membulat berisi ikon (motif kartu & chip template)."""
    t = kotak(slide, x, y, ukuran, ukuran * 1.08, warna="putih", radius=ukuran * 0.24)
    bayangan(t, alpha=25)
    s = ukuran * besar_ikon
    gambar(slide, ikon_path(ikon), x + (ukuran - s) / 2, y + (ukuran * 1.08 - s) / 2, w=s, h=s)
    return t


# --------------------------------------------------------- kerangka ---
NAV = ["Pendahuluan", "Tujuan Pembelajaran", "Materi", "Penutup"]


def navbar(slide, aktif=None, label=None):
    """Bilah navigasi bawah; label aktif diberi pil putih. Mengembalikan {label: shape}."""
    label = label or NAV
    x0, y0, w0, h0 = 4.20, 10.26, 7.63, 0.71
    kotak(slide, x0, y0, w0, h0, gradien=GRADIEN_LEMBUT, radius=0.06, nama="Navigasi")
    ukuran = 14
    lebar = [lebar_pt(n, ukuran, True) / 72 for n in label]
    pad = 0.22
    sisa = w0 - 2 * pad - sum(lebar) - 2 * 0.2 * len(label)
    celah = sisa / (len(label) - 1)
    x = x0 + pad
    hasil = {}
    for nama, lw in zip(label, lebar):
        kw = lw + 0.4
        if nama == aktif:
            pil = kotak(slide, x, y0 + 0.11, kw, h0 - 0.22, warna="putih", radius=0.25,
                        nama=f"Nav aktif {nama}")
            bayangan(pil, alpha=18, blur=0.06, jarak=0.02)
        tb = teks(slide, x, y0, kw, h0,
                  [P(nama, ukuran, "deep" if nama == aktif else "teal", tebal=True, rata="c", spasi=1.1)],
                  jangkar="m", nama=f"Nav {nama}")
        hasil[nama] = tb
        x += kw + celah
    return hasil


def nomor_halaman(slide, nomor, total):
    teks(slide, 13.9, 10.33, 1.48, 0.57,
         [P(f"{nomor} / {total}", 13, "teal", tebal=True, rata="r", spasi=1.1)],
         jangkar="m", nama="Nomor halaman")


def chip(slide, label, ikon):
    ukuran = 23
    w = 0.95 + lebar_pt(label, ukuran, True) / 72 + 0.45
    kotak(slide, KIRI, Y_CHIP, w, H_CHIP, gradien=GRADIEN_LEMBUT, radius=0.08, nama="Chip bagian")
    ubin_ikon(slide, KIRI + 0.15, Y_CHIP + 0.12, 0.58, ikon)
    teks(slide, KIRI + 0.89, Y_CHIP, w - 0.89, H_CHIP,
         [P(label, ukuran, tebal=True, spasi=1.1, gradien=GRADIEN_TEKS)], jangkar="m", nama="Label bagian")
    return w


def pil(slide, x, y, isi, ukuran=14, warna="tint", warna_teks="teal", h=0.46, tebal=True, gradien=None, nama=None):
    """Pil kecil (tag/label). Mengembalikan (shape, lebar)."""
    w = lebar_pt(isi, ukuran, tebal) / 72 + 0.5
    shp = kotak(slide, x, y, w, h, warna=None if gradien else warna, gradien=gradien, radius=h / 2,
                nama=nama or f"Tag {isi}")
    isi_teks(shp.text_frame, [P(isi, ukuran, warna_teks, tebal=tebal, rata="c", spasi=1.0)], jangkar="m")
    return shp, w


def tombol(slide, x, y, isi, rata="l", w=None, h=0.5):
    """Tombol tautan kecil di area bawah (mis. 'Lihat pembahasan ›')."""
    w = w or lebar_pt(isi, 13, True) / 72 + 0.55
    if rata == "r":
        x = x - w
    shp = kotak(slide, x, y, w, h, warna="tint", radius=h / 2, nama=f"Tombol {isi}")
    isi_teks(shp.text_frame, [P(isi, 13, "teal", tebal=True, rata="c", spasi=1.0)], jangkar="m")
    return shp


def judul(slide, isi, y=Y_JUDUL, ukuran=44):
    segs = segmen(isi, tebal=True)
    while lebar_segs(segs, ukuran) > LEBAR_ISI * 0.97 and ukuran > 30:
        ukuran -= 2
    teks(slide, KIRI, y, LEBAR_ISI, 0.85, [P(isi, ukuran, "teal", tebal=True, spasi=1.1)], nama="Judul")
    return y + 0.85


def pengantar(slide, isi, y, lebar=13.6, ukuran=22):
    p = [P(isi, ukuran, "teks", spasi=1.3)]
    h = tinggi_paragraf(p, lebar)
    teks(slide, KIRI, y, lebar, h, p, nama="Pengantar")
    return y + h


def kerangka_isi(slide, label_chip, ikon, judul_slide=None, isi_pengantar=None):
    """Chip + judul + pengantar. Mengembalikan y awal area konten."""
    chip(slide, label_chip, ikon)
    y = Y_JUDUL
    if judul_slide:
        y = judul(slide, judul_slide)
    else:
        y = Y_CHIP + H_CHIP + 0.35
    if isi_pengantar:
        y = pengantar(slide, isi_pengantar, y + 0.05)
    return y + JARAK


# ----------------------------------------------------------- komponen ---
def _par_catatan(isi, judul_catatan, ukuran):
    par = []
    if judul_catatan:
        par.append(P(judul_catatan, 18, "teal", tebal=True, spasi=1.2, setelah=4))
    par.append(P(isi, ukuran, "teks", spasi=1.3))
    return par, (1.25, 0.24, 0.35, 0.24)


def tinggi_catatan(isi, w=LEBAR_ISI, judul_catatan=None, ukuran=20):
    par, inset = _par_catatan(isi, judul_catatan, ukuran)
    return max(1.0, tinggi_paragraf(par, w, inset))


def catatan(slide, y, jenis, isi, x=KIRI, w=LEBAR_ISI, judul_catatan=None, ukuran=20):
    """Kotak catatan: peringatan (amber) / tips (cyan) / pemantik."""
    latar = "tint_amber" if jenis == "peringatan" else "tint"
    ikon = {"peringatan": "peringatan", "tips": "strategi", "pemantik": "globe"}.get(jenis, "strategi")
    par, inset = _par_catatan(isi, judul_catatan, ukuran)
    h = max(1.0, tinggi_paragraf(par, w, inset))
    kotak(slide, x, y, w, h, warna=latar, radius=0.1, nama=f"Catatan {jenis}")
    if jenis == "peringatan":
        gambar(slide, ikon_path(ikon), x + 0.32, y + (h - 0.5) / 2, w=0.58)
    else:
        ubin_ikon(slide, x + 0.3, y + (h - 0.65) / 2, 0.6, ikon)
    teks(slide, x, y, w, h, par, inset=inset, jangkar="m")
    return y + h


def _siapkan_tabel(kolom, baris, lebar, ukuran, w, tengah, sorot, kebenaran, tinggi_min):
    n = len(kolom) if kolom else max(len(b) for b in baris)
    if not lebar:  # lebar kolom mengikuti panjang isi terpanjang
        semua = ([kolom] if kolom else []) + [list(b) + [""] * (n - len(b)) for b in baris]
        panjang = [max(len(re.sub(r"\*\*|\$", "", r[j])) for r in semua) for j in range(n)]
        lebar = [min(42, max(6, p)) for p in panjang]
    rasio = lebar
    lebar_kol = [w * r / sum(rasio) for r in rasio]
    inset = (0.22, 0.12, 0.22, 0.12)
    rata = "c" if tengah else "l"

    def spes(isi, header=False, r=None, c=None):
        rumus_saja = bool(re.fullmatch(r"\$[^$]+\$", isi.strip()))
        if header:
            return P(isi.replace("**", ""), ukuran, "putih", tebal=True, rata=rata, spasi=1.25)
        if kebenaran:
            terang = sorot and [r, c] == list(sorot)
            return P(isi, ukuran + 4, "deep" if isi == "B" else "slate", tebal=(isi == "B" or terang),
                     rata="c", spasi=1.2)
        if rumus_saja:
            return P(isi, ukuran + 4, "deep", rata=rata, spasi=1.2)
        return P(isi, ukuran, "teks", rata=rata, spasi=1.25)

    sel = ([[spes(k, header=True) for k in kolom]] if kolom else []) + [
        [spes(v, r=i, c=j) for j, v in enumerate(list(b) + [""] * (n - len(b)))] for i, b in enumerate(baris)
    ]
    tinggi = [max(max(tinggi_paragraf([p], lw, inset) for p, lw in zip(row, lebar_kol)), tinggi_min)
              for row in sel]
    return sel, lebar_kol, inset, tinggi


def tinggi_tabel(kolom, baris, lebar=None, ukuran=20, w=LEBAR_ISI, tinggi_min=0.62):
    return sum(_siapkan_tabel(kolom, baris, lebar, ukuran, w, False, None, False, tinggi_min)[3])


def tabel(slide, y, kolom, baris, lebar=None, ukuran=20, x=KIRI, w=LEBAR_ISI,
          tengah=False, sorot=None, kebenaran=False, tinggi_min=0.62):
    """Tabel dengan header slate (kolom=None: tanpa header); sel berisi markup teks kaya."""
    sel, lebar_kol, inset, tinggi = _siapkan_tabel(kolom, baris, lebar, ukuran, w, tengah, sorot,
                                                   kebenaran, tinggi_min)
    ada_header = bool(kolom)
    n = len(lebar_kol)
    total = sum(tinggi)
    gf = slide.shapes.add_table(len(sel), n, Inches(x), Inches(y), Inches(w), Inches(total))
    gf.name = "Tabel"
    tbl = gf.table
    tblPr = tbl._tbl.tblPr
    for attr in ("firstRow", "bandRow"):
        tblPr.set(attr, "0")
    sid = tblPr.find(qn("a:tableStyleId"))
    if sid is not None:
        tblPr.remove(sid)
    for j, lw in enumerate(lebar_kol):
        tbl.columns[j].width = Inches(lw)
    for i, row in enumerate(sel):
        tbl.rows[i].height = Inches(tinggi[i])
        for j, p in enumerate(row):
            c = tbl.cell(i, j)
            c.text_frame.text = ""
            isi_teks(c.text_frame, [p], inset, "m")
            c.margin_left, c.margin_top = Inches(inset[0]), Inches(inset[1])
            c.margin_right, c.margin_bottom = Inches(inset[2]), Inches(inset[3])
            c.vertical_anchor = MSO_ANCHOR.MIDDLE
            r = i - 1 if ada_header else i  # indeks baris isi
            if ada_header and i == 0:
                isi = "slate"
            elif kebenaran and sorot and [r, j] == list(sorot):
                isi = "sorot"
            elif kebenaran and sorot and r == sorot[0]:
                isi = "tint_amber"
            else:
                isi = "putih" if r % 2 == 0 else "tint"
            _sel_gaya(c, isi)
    return y + total


def _sel_gaya(cell, isi, garis="garis"):
    tcPr = cell._tc.get_or_add_tcPr()
    for el in list(tcPr):
        tcPr.remove(el)
    w = int(Pt(1))
    for sisi in ("a:lnL", "a:lnR", "a:lnT", "a:lnB"):
        ln = etree.SubElement(tcPr, qn(sisi), w=str(w), cap="flat", cmpd="sng", algn="ctr")
        sf = etree.SubElement(ln, qn("a:solidFill"))
        etree.SubElement(sf, qn("a:srgbClr"), val=WARNA[garis])
        etree.SubElement(ln, qn("a:prstDash"), val="solid")
    sf = etree.SubElement(tcPr, qn("a:solidFill"))
    etree.SubElement(sf, qn("a:srgbClr"), val=WARNA[isi])
    tcPr.set("anchor", "ctr")


def kartu_isi(item, lebar):
    par = [P(item["judul"], 26, "putih", tebal=True, spasi=1.15, setelah=8)]
    if item.get("rumus"):
        par.append(P(item["rumus"], 28, "sage", spasi=1.2, setelah=8))
    if item.get("isi"):
        par.append(P(item["isi"], 20, "putih", spasi=1.3, setelah=10))
    if item.get("contoh"):
        par.append(P("**Contoh:** " + item["contoh"], 18, "redup", spasi=1.3))
    return par


def grid_kartu(slide, y, kartu, kolom=2, x=KIRI, w=LEBAR_ISI, celah=0.3, pembuat=kartu_isi,
               inset=(0.35, 0.3, 0.35, 0.3), warna_list=None):
    lw = (w - celah * (kolom - 1)) / kolom
    baris = [kartu[i:i + kolom] for i in range(0, len(kartu), kolom)]
    pola = warna_list or (["slate", "teal", "teal", "slate"] if kolom == 2 else ["slate", "teal"])
    idx = 0
    for b in baris:
        par = [pembuat(k, lw) for k in b]
        h = max(tinggi_paragraf(p, lw, inset) for p in par)
        for j, (k, p) in enumerate(zip(b, par)):
            warna = k.get("warna") or pola[idx % len(pola)]
            idx += 1
            kx = x + j * (lw + celah)
            if warna == "gradien":
                kotak(slide, kx, y, lw, h, gradien=GRADIEN_BRAND, radius=0.1)
            else:
                kotak(slide, kx, y, lw, h, warna=warna, radius=0.1)
            teks(slide, kx, y, lw, h, p, inset=inset)
        y += h + celah
    return y - celah


def kartu_banding(slide, y, kartu, x=KIRI, w=LEBAR_ISI, celah=0.3):
    lw = (w - celah) / 2
    inset = (0.4, 0.32, 0.4, 0.32)
    kol_rumus = 2.9
    blok = []
    for k in kartu:
        kepala = P(k["judul"], 26, "putih", tebal=True, spasi=1.15)
        hk = tinggi_paragraf([kepala], lw, (inset[0], 0, inset[2], 0))
        baris = []
        for rumus, ket in k["baris"]:
            pr = P(rumus, 30, "sage", spasi=1.2)
            pk = P(ket, 20, "putih", spasi=1.3)
            hb = max(tinggi_paragraf([pr], kol_rumus), tinggi_paragraf([pk], lw - inset[0] - inset[2] - kol_rumus))
            baris.append((pr, pk, max(hb, 0.62)))
        blok.append((k, kepala, hk, baris))
    h = max(inset[1] + hk + 0.25 + sum(b[2] for b in baris) + 0.18 * (len(baris) - 1) + inset[3]
            for _, _, hk, baris in blok)
    for j, (k, kepala, hk, baris) in enumerate(blok):
        kx = x + j * (lw + celah)
        kotak(slide, kx, y, lw, h, warna=k.get("warna", "slate"), radius=0.1)
        teks(slide, kx + inset[0], y + inset[1], lw - inset[0] - inset[2], hk, [kepala])
        by = y + inset[1] + hk + 0.25
        for pr, pk, hb in baris:
            latar = kotak(slide, kx + 0.22, by - 0.06, lw - 0.44, hb + 0.12, warna="putih", radius=0.08,
                          nama="Baris banding")
            _transparan(latar, 88)
            teks(slide, kx + inset[0], by, kol_rumus, hb, [pr], jangkar="m")
            teks(slide, kx + inset[0] + kol_rumus, by, lw - inset[0] - inset[2] - kol_rumus, hb, [pk], jangkar="m")
            by += hb + 0.18
    return y + h


def _transparan(shp, persen):
    """Atur transparansi isian solid (0-100)."""
    sf = _spPr(shp).find(qn("a:solidFill"))
    clr = sf[0]
    for a in clr.findall(qn("a:alpha")):
        clr.remove(a)
    etree.SubElement(clr, qn("a:alpha"), val=str(int((100 - persen) * 1000)))


def baris_notasi(slide, y, baris, x=KIRI, w=LEBAR_ISI, kol_rumus=5.6):
    for b in baris:
        pr = P(b["rumus"], 32, "deep", spasi=1.2)
        pk = P(b["ket"], 20, "teks", spasi=1.3)
        h = max(tinggi_paragraf([pr], kol_rumus), tinggi_paragraf([pk], w - kol_rumus - 0.9)) + 0.44
        h = max(h, 1.05)
        kotak(slide, x, y, w, h, warna="tint", radius=0.1, nama="Baris notasi")
        teks(slide, x + 0.4, y, kol_rumus, h, [pr], jangkar="m")
        teks(slide, x + 0.4 + kol_rumus, y, w - kol_rumus - 0.8, h, [pk], jangkar="m")
        y += h + 0.2
    return y - 0.2


def panah(slide, x, y, w=0.5, h=0.5, warna="cyan"):
    shp = slide.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Inches(x), Inches(y), Inches(w), Inches(h))
    _hapus_gaya(shp)
    shp.fill.solid()
    shp.fill.fore_color.rgb = _rgb(warna)
    shp.line.fill.background()
    shp.name = "Panah"
    return shp


def ilustrasi_setara(slide, y, kiri, kanan, x=KIRI, w=LEBAR_ISI):
    lebar_panah = 0.9
    lw = (w - lebar_panah) / 2
    inset = (0.35, 0.28, 0.35, 0.28)
    pk = [P("Kalimat", 16, "sage", tebal=True, spasi=1.2, setelah=4), P(kiri, 20, "putih", spasi=1.3)]
    pn = [P("Bentuk setara", 16, "sage", tebal=True, spasi=1.2, setelah=4), P(kanan, 20, "putih", spasi=1.3)]
    h = max(tinggi_paragraf(pk, lw, inset), tinggi_paragraf(pn, lw, inset))
    kotak(slide, x, y, lw, h, warna="slate", radius=0.1)
    teks(slide, x, y, lw, h, pk, inset=inset, jangkar="m")
    panah(slide, x + lw + 0.2, y + h / 2 - 0.25)
    kotak(slide, x + lw + lebar_panah, y, lw, h, warna="teal", radius=0.1)
    teks(slide, x + lw + lebar_panah, y, lw, h, pn, inset=inset, jangkar="m")
    return y + h


def daftar_langkah(slide, y, langkah, x=KIRI, w=LEBAR_ISI, celah=0.17):
    d = 0.66
    for i, l in enumerate(langkah, 1):
        isi = f"**{l['judul']}:** {l['isi']}" if l.get("judul") else l["isi"]
        p = [P(isi, 20, "putih", spasi=1.3)]
        inset = (0.35, 0.16, 0.35, 0.16)
        lw = w - d - 0.25
        h = max(0.8, tinggi_paragraf(p, lw, inset))
        lingkar = kotak(slide, x, y + (h - d) / 2, d, d, gradien=GRADIEN_BRAND, bentuk=MSO_SHAPE.OVAL,
                        radius=0, nama=f"Nomor {i}")
        isi_teks(lingkar.text_frame, [P(str(i), 24, "putih", tebal=True, rata="c", spasi=1.0)], jangkar="m")
        kotak(slide, x + d + 0.25, y, lw, h, warna="teal", radius=0.08, nama=f"Langkah {i}")
        teks(slide, x + d + 0.25, y, lw, h, p, inset=inset, jangkar="m")
        y += h + celah
    return y - celah


def alur_premis(slide, y, premis, langkah, x=KIRI, w=LEBAR_ISI):
    """Baris kartu premis, lalu alur langkah berpanah menuju kesimpulan."""
    inset = (0.3, 0.24, 0.3, 0.24)
    sisi = (inset[0], 0, inset[2], 0)

    def baris_kartu(y, item, lebar_kartu, celah, warna_label, warna_teks, ukuran_rumus, gaya):
        atas = [[P(it["label"], 16, warna_label, tebal=True, spasi=1.2, setelah=4),
                 P(it["teks"], 19, warna_teks, spasi=1.3)] for it in item]
        h_atas = max(tinggi_paragraf(a, lebar_kartu, sisi) for a in atas)
        ada_rumus = any(it.get("rumus") for it in item)
        h_rumus = ukuran_rumus * 1.2 / 72 if ada_rumus else 0
        h = inset[1] + h_atas + (0.14 + h_rumus if ada_rumus else 0) + inset[3]
        for i, (it, a) in enumerate(zip(item, atas)):
            kx = x + i * (lebar_kartu + celah)
            gaya(i, kx, h)
            teks(slide, kx, y + inset[1], lebar_kartu, h_atas, a, inset=sisi)
            if it.get("rumus"):
                teks(slide, kx, y + h - inset[3] - h_rumus, lebar_kartu, h_rumus,
                     [P(it["rumus"], ukuran_rumus, warna_teks if warna_teks == "putih" else "deep", spasi=1.2)],
                     inset=sisi, nama="Rumus")
        return h

    celah = 0.3
    lw = (w - celah * (len(premis) - 1)) / len(premis)
    h = baris_kartu(y, premis, lw, celah, "teal", "teks", 28,
                    lambda i, kx, h: kotak(slide, kx, y, lw, h, warna="tint", radius=0.1, nama=f"Premis {i + 1}"))
    y += h + 0.45

    tengah, akhir = langkah[:-1], langkah[-1]
    lp = 0.8
    lw2 = (w - lp * len(tengah)) / (len(tengah) + 1)

    def gaya_langkah(i, kx, h):
        kotak(slide, kx, y, lw2, h, warna="teal", radius=0.1, nama=f"Langkah {i + 1}")
        panah(slide, kx + lw2 + 0.15, y + h / 2 - 0.25)

    h2 = baris_kartu(y, tengah, lw2, lp, "sage", "putih", 30, gaya_langkah)
    kx = x + len(tengah) * (lw2 + lp)
    kotak(slide, kx, y, lw2, h2, gradien=GRADIEN_BRAND, radius=0.1, nama="Kesimpulan")
    teks(slide, kx, y, lw2, h2,
         [P(akhir["label"], 16, "sage", tebal=True, spasi=1.2, setelah=6),
          P(akhir["teks"], 22, "putih", tebal=True, spasi=1.3)], inset=inset, jangkar="m")
    return y + h2


def poin_berbutir(slide, y, poin, x=KIRI, w=9.8, ukuran=22):
    for isi in poin:
        p = [P(isi, ukuran, "teks", spasi=1.3)]
        h = tinggi_paragraf(p, w - 0.5)
        kotak(slide, x + 0.02, y + ukuran * 1.3 / 72 / 2 - 0.08, 0.16, 0.16, warna="cyan",
              bentuk=MSO_SHAPE.OVAL, radius=0, nama="Butir")
        teks(slide, x + 0.45, y, w - 0.45, h, p)
        y += h + 0.18
    return y - 0.18
