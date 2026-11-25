"""Sistem desain materi Cerebrum, gaya Opsi B "Ceria & Membulat": token, pengukur teks, komponen.

Semua ukuran dalam inci pada kanvas 16 x 11,25 inci, ukuran 'Template Referensi Opsi B -
Ceria & Membulat.pptx'. Palet: navy 13215A, kuning FBBB2B, kuning muda FFE6AE, teks sekunder
4A5578, latar slide isi F4F6FB. Bentuk membulat, warna rata (tanpa gradien dan bayangan).
Teks memakai Poppins (tertanam di file), simbol logika/rumus memakai Cambria Math karena
Poppins tidak punya glyph seperti → ∨ ∧ ≡.

Markup teks kaya yang didukung di naskah:
  **tebal**     -> Poppins Bold
  <i>miring</i> -> Poppins miring; <s>coret</s> -> coretan merah; <r>merah</r> -> teks merah
  $p → q$       -> Cambria Math (tidak dipotong saat pindah baris)
"""
from __future__ import annotations

import datetime
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
from pptx.util import Inches, Pt

ASET = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets")

# ---------------------------------------------------------------- token ---
LEBAR_SLIDE, TINGGI_SLIDE = 16.0, 11.25

WARNA = {
    "navy": "13215A",
    "navy_muda": "24336F",
    "kuning": "FBBB2B",
    "kuning_muda": "FFE6AE",
    "teks": "4A5578",
    "latar": "F4F6FB",
    "putih": "FFFFFF",
    "redup": "C8D1E8",
    "abu": "8A93AD",
    "baris": "F7F8FB",
    "garis": "E3E7F0",
    # nama dari desain lama, dipetakan ke palet Opsi B (masih dipakai buat_latihan.py)
    "deep": "13215A",
    "teal": "13215A",
    "slate": "13215A",
    "cyan": "FBBB2B",
    "sage": "FFE6AE",
    "tint": "FFFFFF",
    "tint_amber": "FFE6AE",
    "sorot": "FBBB2B",
    # token isian bentuk (berbeda per tema; nilainya bisa hex, gradien, atau warna transparan)
    "kartu": "FFFFFF",
    "kartu_gelap": "13215A",
    "baris_terang": "F4F6FB",
    "baris_gelap": "24336F",
    "isi_catatan": "FFE6AE",
    # pasangan kontras: teks/ikon di atas kuning, dan aksen kuning di atas latar gelap. Palet platform
    # yang aksennya gelap (mis. JadiOJK, merah) menimpanya agar tetap terbaca (lihat palet.py).
    "di_kuning": "13215A",
    "kuning_di_gelap": "FBBB2B",
}

# Tema: "opsi-b" (bawaan) dan "general" (hiasan dari 'Template Materi General.pptx', lihat
# tema_general.py). Token tema menimpa WARNA; kait menggantikan chip, navigasi, nomor, judul.
TEMA = "opsi-b"
_KAIT = None
_WARNA_DASAR = dict(WARNA)
TOKEN_TEMA = {
    "opsi-b": {},
    "general": {
        "latar": "FFFFFF",
        "kartu": {"stops": [(0, "FFE6AE", 100), (100, "FFFFFF", 100)], "sudut": 90},
        "kartu_gelap": {"stops": [(0, "13215A", 100), (100, "595858", 100)], "sudut": 0},
        "baris_terang": "FFFFFF",
        "baris_gelap": {"warna": "FFFFFF", "alpha": 16},
        "isi_catatan": "FBBB2B",
        "baris": "FFFFFF",
        "tint": {"stops": [(0, "FFE6AE", 100), (100, "FFFFFF", 100)], "sudut": 90},
    },
}
RADIUS_TEMA = {"opsi-b": 0.4, "general": 0.12}

FONT = "Poppins"
FONT_RUMUS = "Cambria Math"

X_CHIP = 0.60  # chip dan judul sedikit menjorok keluar dari area isi, seperti di template
KIRI = 0.70
LEBAR_ISI = 14.65
KANAN = KIRI + LEBAR_ISI  # 15.35
LEBAR_PENGANTAR = 14.0
BATAS_BAWAH = 9.85
JARAK = 0.3

Y_CHIP, H_CHIP = 0.55, 0.75
Y_JUDUL, H_JUDUL = 1.45, 0.85
Y_NAV, H_NAV = 10.30, 0.60
RADIUS = 0.4  # sudut kartu (diatur ulang oleh pakai_tema)

HAK_CIPTA = f"© {datetime.date.today().year} PT Cerebrum Edukanesia Nusantara. All Rights Reserved."

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

# Opsi B memakai warna rata; konstanta gradien dipertahankan agar kode lama tetap jalan.
GRADIEN_BRAND = [(0, "navy", 100), (100, "navy", 100)]
GRADIEN_LEMBUT = [(0, "kuning_muda", 100), (100, "kuning_muda", 100)]
GRADIEN_TEKS = [(0, "navy", 100), (100, "navy", 100)]


def pakai_tema(nama, kait=None):
    """Aktifkan tema untuk deck berikutnya. kait: objek dengan metode chip/navbar/nomor_halaman/judul."""
    global TEMA, _KAIT, RADIUS
    if nama not in TOKEN_TEMA:
        raise ValueError(f"Tema tidak dikenal: {nama} (pilihan: {', '.join(TOKEN_TEMA)})")
    WARNA.clear()
    WARNA.update(_WARNA_DASAR)
    WARNA.update(TOKEN_TEMA[nama])
    TEMA, _KAIT, RADIUS = nama, kait, RADIUS_TEMA[nama]


def _kait(nama):
    return getattr(_KAIT, nama, None) if _KAIT is not None else None


def ikon_path(nama: str, warna: str | None = None) -> str:
    """Path ikon; warna 'navy' atau 'putih' memakai varian dari generator/buat_ikon.py."""
    berkas = IKON.get(nama, IKON["globe"])
    if warna:
        varian = os.path.join(ASET, "ikon", warna, berkas)
        if os.path.exists(varian):
            return varian
    return os.path.join(ASET, berkas)


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
    if rumus in ("sup", "sub"):  # pangkat/indeks digambar PowerPoint sekitar 2/3 ukuran run
        w *= 0.67
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


def _segmen_rumus(isi, warna=None):
    """Isi rumus -> segmen; pangkat ^{..} dan indeks _{..} menjadi run superskrip/subskrip."""
    keluar = []
    for bagian in re.split(r"(\^\{[^{}]*\}|_\{[^{}]*\})", isi.replace(" ", " ")):
        if bagian.startswith(("^{", "_{")) and bagian.endswith("}"):
            if bagian[2:-1]:
                keluar.append((bagian[2:-1], False, "sup" if bagian[0] == "^" else "sub", warna))
        elif bagian:
            keluar.append((bagian, False, True, warna))
    return keluar


class Tebal(int):
    """Nilai 'tebal' sebuah segmen (0/1) yang sekaligus membawa cetak miring dan coretan."""

    miring = coret = False


def _gaya(tebal, miring, coret=False):
    if not miring and not coret:
        return tebal
    g = Tebal(bool(tebal))
    g.miring, g.coret = miring, coret
    return g


def polos(teks: str) -> str:
    """Teks tanpa penanda markup (tebal, miring, rumus)."""
    return re.sub(r"\*\*|</?[isr]>|\$", "", teks)


def segmen(teks: str, tebal=False, warna=None):
    """Ubah markup menjadi daftar segmen (teks, tebal, rumus, warna).

    rumus: False, True (Cambria Math), atau "sup"/"sub" untuk pangkat dan indeks di dalam $..$.
    tebal: bool, atau Tebal dengan .miring (cetak miring) dan .coret (coretan).
    """
    keluar = []
    for bagian in re.split(r"(\$[^$]+\$)", teks):
        if not bagian:
            continue
        if bagian.startswith("$") and bagian.endswith("$"):
            keluar.extend(_segmen_rumus(bagian[1:-1], warna))
            continue
        bagian = kutip_lengkung(bagian)
        b_kini, nyala = False, {"i": False, "s": False, "r": False}
        for potong in re.split(r"(\*\*|</?[isr]>)", bagian):
            if potong == "**":
                b_kini = not b_kini
                continue
            if re.fullmatch(r"</?[isr]>", potong):
                nyala[potong[-2]] = potong[1] != "/"
                continue
            if not potong:
                continue
            b, miring, coret = tebal or b_kini, nyala["i"], nyala["s"]
            w = "FF0000" if coret or nyala["r"] else warna
            # karakter yang tak ada di Poppins dialihkan ke Cambria Math
            buf, buf_rumus = "", False
            for ch in potong:
                r = not ada_glyph_poppins(ch)
                if buf and r != buf_rumus:
                    keluar.append((buf, _gaya(b and not buf_rumus, miring, coret), buf_rumus, w))
                    buf = ""
                buf += ch
                buf_rumus = r
            if buf:
                keluar.append((buf, _gaya(b and not buf_rumus, miring, coret), buf_rumus, w))
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
            self.segs = _segmen_rumus(teks)
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


def _hex(warna):
    """Hex dari nama warna/token (gradien -> warna titik pertama, transparan -> warnanya)."""
    nilai = WARNA.get(warna, warna)
    if isinstance(nilai, dict):
        return nilai.get("warna") or nilai["stops"][0][1]
    return nilai


def _rgb(warna):
    return RGBColor.from_string(_hex(warna))


def _warna_rata(stops):
    """Warna tunggal jika semua titik gradien sama dan tidak transparan, selain itu None."""
    warna = {WARNA.get(c, c) for _, c, _ in stops}
    return warna.pop() if len(warna) == 1 and all(a == 100 for _, _, a in stops) else None


def _gradfill_xml(stops, sudut=45):
    gs = "".join(
        f'<a:gs pos="{pos * 1000}"><a:srgbClr val="{_hex(c)}">'
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
    rata = {"l": PP_ALIGN.LEFT, "c": PP_ALIGN.CENTER, "r": PP_ALIGN.RIGHT, "j": PP_ALIGN.JUSTIFY}
    for i, spec in enumerate(paragraf):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = rata[spec.rata]
        p.line_spacing = Pt(round(spec.ukuran * spec.spasi, 1))
        p.space_before = Pt(0)
        p.space_after = Pt(spec.setelah if i < len(paragraf) - 1 else 0)
        campur = campuran(spec.segs)
        rata_gradien = _warna_rata(spec.gradien) if spec.gradien else None
        for teks, tebal, rumus, warna in spec.segs:
            r = p.add_run()
            r.text = teks
            f = r.font
            f.size = Pt(ukuran_run(spec.ukuran, rumus, campur))
            f.bold = bool(tebal)
            if getattr(tebal, "miring", False):
                f.italic = True
            if getattr(tebal, "coret", False):
                r._r.get_or_add_rPr().set("strike", "sngStrike")
            f.name = FONT_RUMUS if rumus else FONT
            if rumus in ("sup", "sub"):
                r._r.get_or_add_rPr().set("baseline", "30000" if rumus == "sup" else "-25000")
            if spec.gradien and not rata_gradien:
                rpr = r._r.get_or_add_rPr()
                rpr.insert(0, _gradfill_xml(spec.gradien))
            else:
                f.color.rgb = _rgb(rata_gradien or warna or spec.warna)


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


def latar(slide, warna="latar"):
    """Warna latar slide (bukan bentuk), seperti template Opsi B."""
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = _rgb(warna)


def kotak(slide, x, y, w, h, warna=None, gradien=None, sudut=45, radius=0.08,
          garis=None, tebal_garis=1.0, bentuk=None, nama=None):
    jenis = bentuk or (MSO_SHAPE.ROUNDED_RECTANGLE if radius else MSO_SHAPE.RECTANGLE)
    shp = slide.shapes.add_shape(jenis, Inches(x), Inches(y), Inches(w), Inches(h))
    _hapus_gaya(shp)
    if nama:
        shp.name = nama
    if jenis == MSO_SHAPE.ROUNDED_RECTANGLE:
        shp.adjustments[0] = min(0.5, radius / min(w, h))
    nilai = WARNA.get(warna, warna) if isinstance(warna, str) else warna
    if gradien and not _warna_rata(gradien):
        _ganti_fill(shp, _gradfill_xml(gradien, sudut))
    elif isinstance(nilai, dict) and "stops" in nilai:
        _ganti_fill(shp, _gradfill_xml(nilai["stops"], nilai.get("sudut", 90)))
    elif gradien or nilai:
        shp.fill.solid()
        shp.fill.fore_color.rgb = _rgb(_warna_rata(gradien) if gradien else warna)
        if isinstance(nilai, dict) and nilai.get("alpha") is not None:
            clr = _spPr(shp).find(qn("a:solidFill"))[0]
            etree.SubElement(clr, qn("a:alpha"), val=str(int(nilai["alpha"] * 1000)))
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
    """Atur opasitas gambar (0-100)."""
    blip = pic._element.find(".//" + qn("a:blip"))
    etree.SubElement(blip, qn("a:alphaModFix"), amt=str(int(persen * 1000)))


def lingkaran(slide, x, y, d, warna, isi=None, ukuran=None, warna_teks="navy", nama=None):
    """Lingkaran berwarna, opsional berisi teks tebal di tengah (nomor, level, simbol)."""
    o = kotak(slide, x, y, d, d, warna=warna, bentuk=MSO_SHAPE.OVAL, radius=0, nama=nama)
    if isi is not None:
        ukuran = ukuran or round(d * 72 * 0.4)
        isi_teks(o.text_frame, [P(str(isi), ukuran, warna_teks, tebal=True, rata="c", spasi=1.0)], jangkar="m")
    return o


def ikon_di_kuning():
    """Varian ikon untuk di atas lingkaran kuning: putih bila palet memakai teks terang di atas aksen."""
    return "putih" if _hex(WARNA["di_kuning"]).upper() in ("FFFFFF", "FFF") else "navy"


def ubin_ikon(slide, x, y, ukuran, ikon, besar_ikon=0.6, warna="kuning", warna_ikon="navy"):
    """Lingkaran kuning berisi ikon navy (motif chip dan kartu Opsi B)."""
    if warna == "kuning" and warna_ikon == "navy":
        warna_ikon = ikon_di_kuning()
    o = lingkaran(slide, x, y, ukuran, warna, nama="Ikon")
    s = ukuran * besar_ikon
    gambar(slide, ikon_path(ikon, warna_ikon), x + (ukuran - s) / 2, y + (ukuran - s) / 2, w=s, h=s)
    return o


def panah(slide, x, y, w=0.5, h=0.5, warna="kuning"):
    shp = slide.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Inches(x), Inches(y), Inches(w), Inches(h))
    _hapus_gaya(shp)
    shp.fill.solid()
    shp.fill.fore_color.rgb = _rgb(warna)
    shp.line.fill.background()
    shp.name = "Panah"
    return shp


def isian_kartu(warna):
    """Warna kartu (navy/putih) -> token isian tema (kartu_gelap/kartu)."""
    return {"navy": "kartu_gelap", "putih": "kartu"}.get(warna, warna)


def _warna_kartu(warna, bawaan="putih"):
    """Nama warna kartu dari naskah (termasuk nama desain lama) ke palet Opsi B."""
    peta = {"teal": "navy", "slate": "navy", "deep": "navy", "gradien": "navy", "tint": "putih",
            "sage": "kuning_muda", "navy": "navy", "putih": "putih", "kuning": "kuning",
            "kuning_muda": "kuning_muda"}
    return peta.get(warna, bawaan) if warna else bawaan


# --------------------------------------------------------- kerangka ---
NAV = ["Pendahuluan", "Tujuan Pembelajaran", "Pendalaman Materi", "Penutup"]


def navbar(slide, aktif=None, label=None, gelap=False):
    """Pil navigasi navy di bawah; bagian aktif diberi pil kuning. Mengembalikan {label: shape}.
    gelap: slide berlatar gelap/berwarna (dipakai tema yang punya slide berbingkai)."""
    label = label or NAV
    if _kait("navbar"):
        return _KAIT.navbar(slide, aktif, label, gelap)
    w0 = 8.0
    x0 = (LEBAR_SLIDE - w0) / 2
    kotak(slide, x0, Y_NAV, w0, H_NAV, warna="navy", radius=H_NAV / 2, nama="Navigasi")
    huruf = [n.upper() for n in label]
    pad, tambah = 0.12, 0.3
    for ukuran in (11, 10, 9):
        lebar = [lebar_pt(t, ukuran, True) / 72 + tambah for t in huruf]
        celah = (w0 - 2 * pad - sum(lebar)) / max(1, len(label) - 1)
        if celah >= 0.05:
            break
    x = x0 + pad
    hasil = {}
    for nama, t, lw in zip(label, huruf, lebar):
        if nama == aktif:
            kotak(slide, x, Y_NAV + 0.09, lw, H_NAV - 0.18, warna="kuning", radius=(H_NAV - 0.18) / 2,
                  nama=f"Nav aktif {nama}")
        hasil[nama] = teks(slide, x, Y_NAV, lw, H_NAV,
                           [P(t, ukuran, "navy" if nama == aktif else "putih", tebal=True, rata="c", spasi=1.0)],
                           jangkar="m", nama=f"Nav {nama}")
        x += lw + celah
    return hasil


def nomor_halaman(slide, nomor, total, gelap=False):
    if _kait("nomor_halaman"):
        return _KAIT.nomor_halaman(slide, nomor, total, gelap)
    lingkaran(slide, KANAN - 0.75, Y_NAV, 0.6, "kuning", nomor, 13 if nomor < 100 else 11, "navy",
              nama="Nomor halaman")


def chip(slide, label, ikon):
    """Pil navy berisi lingkaran kuning + ikon dan label bagian. Mengembalikan lebarnya."""
    if _kait("chip"):
        return _KAIT.chip(slide, label, ikon)
    ukuran = 18
    tw = lebar_pt(label, ukuran, True) / 72
    w = 0.82 + tw + 0.35
    kotak(slide, X_CHIP, Y_CHIP, w, H_CHIP, warna="navy", radius=H_CHIP / 2, nama="Chip bagian")
    ubin_ikon(slide, X_CHIP + 0.10, Y_CHIP + 0.08, 0.59, ikon, besar_ikon=0.66)
    teks(slide, X_CHIP + 0.82, Y_CHIP, tw + 0.3, H_CHIP, [P(label, ukuran, "putih", tebal=True, spasi=1.1)],
         jangkar="m", nama="Label bagian")
    return w


def pil(slide, x, y, isi, ukuran=14, warna="putih", warna_teks="navy", h=0.46, tebal=True, gradien=None,
        nama=None):
    """Pil kecil (tag/label). Mengembalikan (shape, lebar)."""
    w = lebar_pt(isi, ukuran, tebal) / 72 + 0.5
    shp = kotak(slide, x, y, w, h, warna=None if gradien else warna, gradien=gradien, radius=h / 2,
                nama=nama or f"Tag {isi}")
    isi_teks(shp.text_frame, [P(isi, ukuran, warna_teks, tebal=tebal, rata="c", spasi=1.0)], jangkar="m")
    return shp, w


def tombol(slide, x, y, isi, rata="l", w=None, h=0.5):
    """Tombol tautan kecil di area bawah (mis. 'Lihat pembahasan ›')."""
    ukuran = 13
    if TEMA == "general" and y > 10:  # navigasi General dimulai di x 2,22: tombol kecil di paling kiri
        ukuran, x, h, y = 10, 0.12, 0.44, 10.15
    w = w or lebar_pt(isi, ukuran, True) / 72 + 0.4
    if rata == "r":
        x = x - w
    shp = kotak(slide, x, y, w, h, warna="kartu", radius=h / 2, nama=f"Tombol {isi}")
    isi_teks(shp.text_frame, [P(isi, ukuran, "navy", tebal=True, rata="c", spasi=1.0)], jangkar="m")
    return shp


def judul(slide, isi, y=Y_JUDUL, ukuran=38):
    if _kait("judul"):
        return _KAIT.judul(slide, isi, y)
    w = LEBAR_SLIDE - 2 * X_CHIP
    segs = segmen(isi, tebal=True)
    while lebar_segs(segs, ukuran) > w * 0.97 and ukuran > 28:
        ukuran -= 2
    teks(slide, X_CHIP, y, w, H_JUDUL, [P(isi, ukuran, "navy", tebal=True, spasi=1.1)], nama="Judul")
    return y + H_JUDUL


def pengantar(slide, isi, y, lebar=LEBAR_PENGANTAR, ukuran=19):
    p = [P(isi, ukuran, "teks", spasi=1.35)]
    h = tinggi_paragraf(p, lebar)
    teks(slide, KIRI, y, lebar, h, p, nama="Pengantar")
    return y + h


def kerangka_isi(slide, label_chip, ikon, judul_slide=None, isi_pengantar=None):
    """Chip + judul + pengantar. Mengembalikan y awal area konten."""
    chip(slide, label_chip, ikon)
    y = judul(slide, judul_slide) if judul_slide else Y_CHIP + H_CHIP + 0.35
    if isi_pengantar:
        return pengantar(slide, isi_pengantar, y) + 0.35
    return y + 0.15


def _hiasan_sampul(slide, ikon):
    """Tiga bentuk membulat bertumpuk (navy, kuning, putih) dengan ikon, terpotong tepi kanan."""
    kotak(slide, 8.2, 1.2, 10.0, 8.9, warna="navy", radius=4.45, nama="Hiasan navy")
    kotak(slide, 9.6, 2.35, 8.6, 6.6, warna="kuning", radius=3.3, nama="Hiasan kuning")
    kotak(slide, 11.1, 3.6, 7.0, 4.1, warna="putih", radius=2.05, nama="Hiasan putih")
    gambar(slide, ikon_path(ikon, "navy"), 12.0, 4.25, w=2.8, h=2.8, nama="Ikon sampul")


def _judul_sampul(slide, isi, y, lebar=7.3, maks=64):
    segs = segmen(isi, tebal=True)
    kata_terpanjang = max(isi.split(), key=lambda k: lebar_pt(k, 10, True))
    ukuran = maks
    while ukuran > 36 and (jumlah_baris(segs, ukuran, lebar) > 2 or
                           lebar_pt(kata_terpanjang, ukuran, True) / 72 > lebar * 0.95):
        ukuran -= 4
    h = jumlah_baris(segs, ukuran, lebar) * ukuran * 1.05 / 72
    teks(slide, 0.8, y, lebar + 0.2, h + 0.1, [P(isi, ukuran, "navy", tebal=True, spasi=1.05)], nama="Judul sampul")
    return y + h


def sampul(slide, judul_sampul, tag=None, subjudul=None, ikon="strategi"):
    latar(slide, "putih")
    _hiasan_sampul(slide, ikon)
    if tag:
        pil(slide, 0.8, 3.0, tag, ukuran=15, warna="kuning", warna_teks="navy", h=0.55, nama="Tag subtes")
    y = _judul_sampul(slide, judul_sampul, 3.75)
    if subjudul:
        p = [P(subjudul, 22, "teks", spasi=1.3)]
        teks(slide, 0.8, y + 0.3, 7.0, tinggi_paragraf(p, 7.0), p, nama="Subjudul sampul")
    teks(slide, 0.8, 10.25, 9.0, 0.4, [P(HAK_CIPTA, 12, "abu", spasi=1.2)], nama="Hak cipta")


def terima_kasih(slide, pesan="Sampai jumpa di materi berikutnya.", ikon="roket"):
    latar(slide, "putih")
    _hiasan_sampul(slide, ikon)
    y = _judul_sampul(slide, "Terima kasih", 4.2)
    teks(slide, 0.8, y + 0.3, 7.0, 0.6, [P(pesan, 22, "teks", spasi=1.3)], nama="Pesan akhir")
    teks(slide, 0.8, 10.25, 9.0, 0.4, [P(HAK_CIPTA, 12, "abu", spasi=1.2)], nama="Hak cipta")


# ----------------------------------------------------------- komponen ---
SIMBOL_CATATAN = {"peringatan": "!", "tips": "i", "pemantik": "?", "info": "i"}


def _par_catatan(isi, judul_catatan, ukuran):
    par = []
    if judul_catatan:
        par.append(P(judul_catatan, 16, "navy", tebal=True, spasi=1.2, setelah=4))
    par.append(P(isi, ukuran, "navy", spasi=1.3))
    return par, (1.65, 0.28, 0.45, 0.28)


def tinggi_catatan(isi, w=LEBAR_ISI, judul_catatan=None, ukuran=19):
    par, inset = _par_catatan(isi, judul_catatan, ukuran)
    return max(1.15, tinggi_paragraf(par, w, inset))


def catatan(slide, y, jenis, isi, x=KIRI, w=LEBAR_ISI, judul_catatan=None, ukuran=19):
    """Kotak kuning muda dengan lingkaran navy bersimbol: peringatan (!), tips (i), pemantik (?)."""
    par, inset = _par_catatan(isi, judul_catatan, ukuran)
    h = max(1.15, tinggi_paragraf(par, w, inset))
    kotak(slide, x, y, w, h, warna="isi_catatan", radius=RADIUS, nama=f"Catatan {jenis}")
    d = 0.85
    lingkaran(slide, x + 0.45, y + (h - d) / 2, d, "navy", SIMBOL_CATATAN.get(jenis, "i"), 24, "putih",
              nama="Simbol catatan")
    teks(slide, x, y, w, h, par, inset=inset, jangkar="m")
    return y + h


def _siapkan_tabel(kolom, baris, lebar, ukuran, w, tengah, sorot, kebenaran, tinggi_min):
    n = len(kolom) if kolom else max(len(b) for b in baris)
    if not lebar:  # lebar kolom mengikuti panjang isi terpanjang
        semua = ([kolom] if kolom else []) + [list(b) + [""] * (n - len(b)) for b in baris]
        panjang = [max(len(polos(r[j])) for r in semua) for j in range(n)]
        lebar = [min(42, max(6, p)) for p in panjang]
    rasio = lebar
    lebar_kol = [w * r / sum(rasio) for r in rasio]
    inset = (0.22, 0.12, 0.22, 0.12)
    rata = "c" if tengah else "l"
    if not kebenaran:  # kolom tidak lebih sempit dari kata terpanjangnya (kata tidak terpotong di tengah)
        def kata_terlebar(j):
            kata = [(x, max(14, ukuran - 2), True) for x in polos(kolom[j]).split()] if kolom else []
            kata += [(x, ukuran, "**" in b[j]) for b in baris if j < len(b) and "$" not in b[j] for x in polos(b[j]).split()]
            return max((lebar_pt(x, u, t) for x, u, t in kata), default=0) / 72 + inset[0] + inset[2] + 0.05

        minimal = [min(kata_terlebar(j), w * 0.45) for j in range(n)]
        if sum(minimal) < w:
            for _ in range(n):
                kurang = [j for j in range(n) if lebar_kol[j] < minimal[j] - 1e-6]
                if not kurang:
                    break
                longgar = [j for j in range(n) if lebar_kol[j] > minimal[j]]
                butuh = sum(minimal[j] - lebar_kol[j] for j in kurang)
                sisa = sum(lebar_kol[j] - minimal[j] for j in longgar)
                for j in kurang:
                    lebar_kol[j] = minimal[j]
                for j in longgar:
                    lebar_kol[j] -= butuh * (lebar_kol[j] - minimal[j]) / sisa

    def spes(isi, header=False, r=None, c=None):
        rumus_saja = bool(re.fullmatch(r"\$[^$]+\$", isi.strip()))
        if header:
            return P(re.sub(r"\*\*|</?[isr]>", "", isi), max(14, ukuran - 2), "putih", tebal=True, rata=rata, spasi=1.25)
        if kebenaran:
            terang = sorot and [r, c] == list(sorot)
            return P(isi, ukuran + 4, "navy" if isi == "B" else "teks", tebal=(isi == "B" or terang),
                     rata="c", spasi=1.2)
        if rumus_saja:
            return P(isi, ukuran + 4, "navy", rata=rata, spasi=1.2)
        return P(isi, ukuran, "navy", rata=rata, spasi=1.25)

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
    """Tabel dengan header navy (kolom=None: tanpa header); sel berisi markup teks kaya."""
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
                isi = "navy"
            elif kebenaran and sorot and [r, j] == list(sorot):
                isi = "kuning"
            elif kebenaran and sorot and r == sorot[0]:
                isi = "kuning_muda"
            else:
                isi = "putih" if r % 2 == 0 else "baris"
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


def kartu_isi(item, lebar, gelap=False):
    wj, wr, wi = ("putih", "kuning_di_gelap", "redup") if gelap else ("navy", "navy", "teks")
    par = [P(item["judul"], 22, wj, tebal=True, spasi=1.15, setelah=8)]
    if item.get("rumus"):
        par.append(P(item["rumus"], 26, wr, spasi=1.2, setelah=8))
    if item.get("isi"):
        par.append(P(item["isi"], 18, wi, spasi=1.35, setelah=10))
    for butir in item.get("poin", []):
        par.append(P("•  " + butir, 18, wi, spasi=1.35))
    if item.get("contoh"):
        par.append(P("**Contoh:** " + item["contoh"], 16, wi, spasi=1.35))
    return par


def grid_kartu(slide, y, kartu, kolom=2, x=KIRI, w=LEBAR_ISI, celah=0.3, pembuat=kartu_isi,
               inset=(0.4, 0.35, 0.4, 0.35), warna_list=None):
    """Kartu putih membulat (atau navy jika 'warna' kartu navy) dalam grid."""
    lw = (w - celah * (kolom - 1)) / kolom
    baris = [kartu[i:i + kolom] for i in range(0, len(kartu), kolom)]
    pola = warna_list or ["putih"]
    idx = 0
    for b in baris:
        warna = []
        for k in b:
            warna.append(_warna_kartu(k.get("warna"), pola[idx % len(pola)]))
            idx += 1
        par = [pembuat(k, lw, wn == "navy") for k, wn in zip(b, warna)]
        h = max(tinggi_paragraf(p, lw, inset) for p in par)
        for j, (p, wn) in enumerate(zip(par, warna)):
            kx = x + j * (lw + celah)
            kotak(slide, kx, y, lw, h, warna=isian_kartu(wn), radius=RADIUS, nama="Kartu")
            teks(slide, kx, y, lw, h, p, inset=inset)
        y += h + celah
    return y - celah


def kartu_banding(slide, y, kartu, x=KIRI, w=LEBAR_ISI, celah=0.3):
    """Dua kartu berdampingan: kiri navy, kanan putih; tiap baris berisi rumus dan keterangan."""
    lw = (w - celah) / 2
    inset = (0.45, 0.4, 0.45, 0.4)
    kol_rumus = 3.0
    blok = []
    for j, k in enumerate(kartu):
        gelap = _warna_kartu(k.get("warna")) == "navy" if k.get("warna") in ("navy", "putih") else j == 0
        kepala = P(k["judul"], 22, "putih" if gelap else "navy", tebal=True, spasi=1.15)
        hk = tinggi_paragraf([kepala], lw, (inset[0], 0, inset[2], 0))
        baris = []
        for rumus, ket in k["baris"]:
            pr = P(rumus, 26, "kuning_di_gelap" if gelap else "navy", spasi=1.2)
            pk = P(ket, 18, "putih" if gelap else "teks", spasi=1.3)
            hb = max(tinggi_paragraf([pr], kol_rumus), tinggi_paragraf([pk], lw - inset[0] - inset[2] - kol_rumus))
            baris.append((pr, pk, max(hb, 0.62)))
        blok.append((gelap, kepala, hk, baris))
    h = max(inset[1] + hk + 0.25 + sum(b[2] for b in baris) + 0.18 * (len(baris) - 1) + inset[3]
            for _, _, hk, baris in blok)
    for j, (gelap, kepala, hk, baris) in enumerate(blok):
        kx = x + j * (lw + celah)
        kotak(slide, kx, y, lw, h, warna="kartu_gelap" if gelap else "kartu", radius=RADIUS, nama="Kartu banding")
        teks(slide, kx + inset[0], y + inset[1], lw - inset[0] - inset[2], hk, [kepala])
        by = y + inset[1] + hk + 0.25
        for pr, pk, hb in baris:
            kotak(slide, kx + 0.22, by - 0.06, lw - 0.44, hb + 0.12, warna="baris_gelap" if gelap else "baris_terang",
                  radius=0.2, nama="Baris banding")
            teks(slide, kx + inset[0], by, kol_rumus, hb, [pr], jangkar="m")
            teks(slide, kx + inset[0] + kol_rumus, by, lw - inset[0] - inset[2] - kol_rumus, hb, [pk], jangkar="m")
            by += hb + 0.18
    return y + h


def baris_notasi(slide, y, baris, x=KIRI, w=LEBAR_ISI, kol_rumus=5.6):
    """Baris putih membulat: rumus besar di kiri, keterangan di kanan."""
    lebar_ket = w - 0.45 - kol_rumus - 0.25 - 0.4
    for b in baris:
        pr = P(b["rumus"], 30, "navy", spasi=1.2)
        pk = P(b["ket"], 18, "teks", spasi=1.35)
        h = max(tinggi_paragraf([pr], kol_rumus), tinggi_paragraf([pk], lebar_ket)) + 0.44
        h = max(h, 1.05)
        kotak(slide, x, y, w, h, warna="kartu", radius=RADIUS, nama="Baris notasi")
        teks(slide, x + 0.45, y, kol_rumus, h, [pr], jangkar="m")
        teks(slide, x + 0.45 + kol_rumus + 0.25, y, lebar_ket, h, [pk], jangkar="m")
        y += h + 0.2
    return y - 0.2


def ilustrasi_setara(slide, y, kiri, kanan, x=KIRI, w=LEBAR_ISI):
    """Kotak contoh navy: kalimat dan bentuk setaranya dihubungkan lambang ≡."""
    par = [P("Contoh", 15, "kuning_di_gelap", tebal=True, spasi=1.2, setelah=6),
           P(f"“{kiri}”  $≡$  “{kanan}”", 19, "putih", spasi=1.35)]
    inset = (0.45, 0.35, 0.45, 0.35)
    h = tinggi_paragraf(par, w, inset)
    kotak(slide, x, y, w, h, warna="kartu_gelap", radius=RADIUS, nama="Contoh setara")
    teks(slide, x, y, w, h, par, inset=inset, jangkar="m")
    return y + h


def _kapital(s):
    return s[:1].upper() + s[1:]


def daftar_langkah(slide, y, langkah, x=KIRI, w=LEBAR_ISI, celah=0.25):
    """Alur bernomor mendatar (2-5 langkah, kartu terakhir navy); jika terlalu panjang, daftar menurun."""
    n = len(langkah)
    if 2 <= n <= 5:
        d = 0.9
        lw = (w - celah * (n - 1)) / n
        inset = (0.3, 0.3, 0.3, 0.3)
        pars = []
        for i, l in enumerate(langkah):
            gelap = i == n - 1
            p = []
            if l.get("judul"):
                p.append(P(l["judul"], 20, "putih" if gelap else "navy", tebal=True, spasi=1.2, setelah=6))
            p.append(P(_kapital(l["isi"]), 17, "redup" if gelap else "teks", spasi=1.35))
            pars.append(p)
        h = max(2.0, max(tinggi_paragraf(p, lw, inset) for p in pars))
        if h <= 4.4:
            x_awal, x_akhir = x + lw / 2, x + (n - 1) * (lw + celah) + lw / 2
            kotak(slide, x_awal, y + d / 2 - 0.04, x_akhir - x_awal, 0.08, warna="kuning", radius=0,
                  nama="Garis alur")
            yk = y + d + 0.35
            for i, p in enumerate(pars):
                kx = x + i * (lw + celah)
                navy = i % 2 == 0
                lingkaran(slide, kx + (lw - d) / 2, y, d, "navy" if navy else "kuning", i + 1, 24,
                          "putih" if navy else "di_kuning", nama=f"Nomor {i + 1}")
                kotak(slide, kx, yk, lw, h, warna="kartu_gelap" if i == n - 1 else "kartu", radius=RADIUS,
                      nama=f"Langkah {i + 1}")
                teks(slide, kx, yk, lw, h, p, inset=inset)
            return yk + h
    d = 0.7
    for i, l in enumerate(langkah, 1):
        isi = f"**{l['judul']}:** {l['isi']}" if l.get("judul") else _kapital(l["isi"])
        p = [P(isi, 19, "navy", spasi=1.3)]
        inset = (1.3, 0.22, 0.4, 0.22)
        h = max(0.95, tinggi_paragraf(p, w, inset))
        kotak(slide, x, y, w, h, warna="kartu", radius=RADIUS, nama=f"Langkah {i}")
        lingkaran(slide, x + 0.3, y + (h - d) / 2, d, "kuning", i, 20, "di_kuning", nama=f"Nomor {i}")
        teks(slide, x, y, w, h, p, inset=inset, jangkar="m")
        y += h + 0.2
    return y - 0.2


def alur_premis(slide, y, premis, langkah, x=KIRI, w=LEBAR_ISI):
    """Baris kartu premis (putih), lalu alur langkah navy berpanah menuju kesimpulan kuning."""
    inset = (0.3, 0.26, 0.3, 0.26)
    sisi = (inset[0], 0, inset[2], 0)

    def baris_kartu(y, item, lebar_kartu, celah, warna_label, warna_teks, ukuran_rumus, gaya):
        atas = [[P(it["label"], 15, warna_label, tebal=True, spasi=1.2, setelah=4),
                 P(it["teks"], 18, warna_teks, spasi=1.3)] for it in item]
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
                     [P(it["rumus"], ukuran_rumus, "putih" if warna_teks == "putih" else "navy", spasi=1.2)],
                     inset=sisi, nama="Rumus")
        return h

    celah = 0.3
    lw = (w - celah * (len(premis) - 1)) / len(premis)
    h = baris_kartu(y, premis, lw, celah, "navy", "teks", 26,
                    lambda i, kx, h: kotak(slide, kx, y, lw, h, warna="kartu", radius=RADIUS,
                                           nama=f"Premis {i + 1}"))
    y += h + 0.45

    tengah, akhir = langkah[:-1], langkah[-1]
    lp = 0.8
    lw2 = (w - lp * len(tengah)) / (len(tengah) + 1)

    def gaya_langkah(i, kx, h):
        kotak(slide, kx, y, lw2, h, warna="kartu_gelap", radius=RADIUS, nama=f"Langkah {i + 1}")
        panah(slide, kx + lw2 + 0.15, y + h / 2 - 0.25)

    h2 = baris_kartu(y, tengah, lw2, lp, "kuning", "putih", 28, gaya_langkah)
    kx = x + len(tengah) * (lw2 + lp)
    kotak(slide, kx, y, lw2, h2, warna="kuning", radius=RADIUS, nama="Kesimpulan")
    teks(slide, kx, y, lw2, h2,
         [P(akhir["label"], 15, "navy", tebal=True, spasi=1.2, setelah=6),
          P(akhir["teks"], 21, "navy", tebal=True, spasi=1.3)], inset=inset, jangkar="m")
    return y + h2


def poin_berbutir(slide, y, poin, x=KIRI, w=13.0, ukuran=20):
    for isi in poin:
        p = [P(isi, ukuran, "navy", spasi=1.3)]
        h = tinggi_paragraf(p, w - 0.5)
        lingkaran(slide, x + 0.04, y + ukuran * 1.3 / 72 / 2 - 0.12, 0.24, "kuning", nama="Butir")
        teks(slide, x + 0.5, y, w - 0.5, h, p)
        y += h + 0.22
    return y - 0.22


# -------------------------------------------------------------- blok ---
# Blok adalah isi umum hasil ekstraksi dokumen (ekstrak_gdoc.py): paragraf, subjudul,
# sorot (baris rumus/penegasan), poin, gambar, tabel, catatan, dan kunci jawaban.
JARAK_BLOK = 0.2


def _gaya_paragraf(gaya):
    if (gaya or "").startswith("soal"):  # 'soal' atau 'soal:17' (ukuran diperkecil untuk bacaan panjang)
        return int(gaya.partition(":")[2] or 20), "navy"
    return 19, "teks"


def _par_blok(b, gaya=None):
    j = b["jenis"]
    if j == "paragraf":
        ukuran, warna = _gaya_paragraf(gaya)
        return [P(b["teks"], ukuran, warna, spasi=1.4)], (0, 0, 0, 0)
    if j == "subjudul":
        return [P(b["teks"], 21, "navy", tebal=True, spasi=1.2)], (0, 0.08, 0, 0)
    if j == "sorot":
        return [P(b["teks"], 22, "navy", spasi=1.3)], (0.45, 0.2, 0.45, 0.2)
    if j == "kunci":
        return [P(b["teks"], 19, "navy", spasi=1.35)], (0.4, 0.22, 0.4, 0.22)
    return None, None


def _ukur_gambar(b, w, h_maks, folder):
    """(lebar, tinggi gambar, tinggi keterangan) dalam inci, dipaskan ke ruang yang ada."""
    w_nat, h_nat = b.get("lebar", 4) * 1.5, b.get("tinggi", 3) * 1.5
    pk = [P(b["keterangan"], 14, "teks", rata="c", spasi=1.3)] if b.get("keterangan") else None
    h_ket = tinggi_paragraf(pk, w) + 0.12 if pk else 0
    skala = min(1.0, (min(w, 12.5) - 0.4) / w_nat, max(0.1, (h_maks - h_ket - 0.4)) / h_nat)
    return w_nat * skala, h_nat * skala, h_ket, skala


def tinggi_blok(b, w=LEBAR_ISI, folder=".", gaya=None, h_maks=99.0):
    j = b["jenis"]
    par, inset = _par_blok(b, gaya)
    if par:
        return tinggi_paragraf(par, w, inset)
    if j == "poin":
        total = 0.0
        for isi in b["butir"]:
            total += tinggi_paragraf([P(isi, 19, "navy", spasi=1.3)], w - 0.5) + 0.22  # = jarak di poin_berbutir
        return total - 0.22
    if j == "gambar":
        gw, gh, h_ket, _ = _ukur_gambar(b, w, h_maks, folder)
        return gh + 0.4 + h_ket
    if j == "tabel":
        return tinggi_tabel(b["kolom"], b["baris"], b.get("lebar"), b.get("ukuran", 16), w)
    if j == "catatan":
        return tinggi_catatan(b["teks"], w, b.get("judul"), 18)
    if j == "profil":
        _, _, tinggi = _ukur_profil(b, w)
        return sum(tinggi) + (len(tinggi) - 1) * CELAH_PROFIL
    return 0.0


# Kartu profil (foto, nama, jabatan) dalam grid; dipakai untuk tabel profil di dokumen sumber.
KOLOM_PROFIL, CELAH_PROFIL = 4, 0.25
INSET_PROFIL = (0.2, 0.1, 0.2, 0.18)


def _par_profil(it):
    par = [P(it.get("nama") or "", 15, "navy", tebal=True, rata="c", spasi=1.15, setelah=4)]
    if it.get("peran"):
        par.append(P(it["peran"], 11, "teks", rata="c", spasi=1.25))
    return par


def _ukur_profil(b, w):
    """(lebar kartu, sisi foto, [tinggi kartu per baris grid])."""
    lw = (w - CELAH_PROFIL * (KOLOM_PROFIL - 1)) / KOLOM_PROFIL
    foto = min(1.15, lw * 0.4)
    item = b["item"]
    tinggi = []
    for i in range(0, max(1, len(item)), KOLOM_PROFIL):
        h_teks = max((tinggi_paragraf(_par_profil(it), lw, INSET_PROFIL) for it in item[i:i + KOLOM_PROFIL]),
                     default=0.5)
        tinggi.append(0.25 + foto + 0.1 + h_teks)
    return lw, foto, tinggi


def kartu_profil(slide, y, b, x=KIRI, w=LEBAR_ISI, folder="."):
    lw, foto, tinggi = _ukur_profil(b, w)
    for i, it in enumerate(b["item"]):
        baris = i // KOLOM_PROFIL
        h_kartu = tinggi[baris]
        kx = x + (i % KOLOM_PROFIL) * (lw + CELAH_PROFIL)
        ky = y + sum(tinggi[:baris]) + baris * CELAH_PROFIL
        kotak(slide, kx, ky, lw, h_kartu, warna="kartu", radius=RADIUS, nama=f"Profil {i + 1}")
        path = os.path.join(folder, it["file"]) if it.get("file") else None
        if path and os.path.exists(path):
            gambar(slide, path, kx + (lw - foto) / 2, ky + 0.25, w=foto, h=foto, nama=f"Foto {i + 1}")
        teks(slide, kx, ky + 0.25 + foto + 0.1, lw, h_kartu - foto - 0.35, _par_profil(it), inset=INSET_PROFIL,
             nama=f"Nama {i + 1}")
    return y + sum(tinggi) + (len(tinggi) - 1) * CELAH_PROFIL


def skala_min_gambar(b, w=LEBAR_ISI):
    """Tinggi terkecil yang masih layak untuk gambar (60% ukuran alaminya)."""
    _, gh, h_ket, _ = _ukur_gambar(b, w, 99.0, ".")
    return gh * 0.6 + 0.4 + h_ket


def gambar_blok(slide, b, x, y, w=LEBAR_ISI, folder=".", gaya=None, h_maks=99.0):
    """Gambar satu blok pada posisi y. Mengembalikan y bawahnya."""
    j = b["jenis"]
    par, inset = _par_blok(b, gaya)
    if par:
        h = tinggi_paragraf(par, w, inset)
        if j == "sorot":
            kotak(slide, x, y, w, h, warna="kartu", radius=RADIUS, nama="Sorotan")
        elif j == "kunci":
            kotak(slide, x, y, w, h, warna="kuning_muda", radius=0.3, nama="Kunci jawaban")
        teks(slide, x, y, w, h, par, inset=inset, jangkar="m" if j in ("sorot", "kunci") else "t")
        return y + h
    if j == "poin":
        return poin_berbutir(slide, y, b["butir"], x, w, ukuran=19)
    if j == "gambar":
        gw, gh, h_ket, _ = _ukur_gambar(b, w, h_maks, folder)
        kw = gw + 0.4
        kx = x + (w - kw) / 2
        kotak(slide, kx, y, kw, gh + 0.4, warna="putih", radius=0.25, nama="Bingkai gambar")
        path = os.path.join(folder, b["file"])
        if os.path.exists(path):
            gambar(slide, path, kx + 0.2, y + 0.2, w=gw, h=gh, nama="Gambar materi")
        y += gh + 0.4
        if b.get("keterangan"):
            p = [P(b["keterangan"], 14, "teks", rata="c", spasi=1.3)]
            teks(slide, x, y + 0.08, w, h_ket, p, nama="Keterangan gambar")
            y += h_ket
        return y
    if j == "tabel":
        return tabel(slide, y, b["kolom"], b["baris"], b.get("lebar"), b.get("ukuran", 16), x=x, w=w)
    if j == "catatan":
        return catatan(slide, y, b.get("catatan", "tips"), b["teks"], x=x, w=w, judul_catatan=b.get("judul"),
                       ukuran=18)
    if j == "profil":
        return kartu_profil(slide, y, b, x, w, folder)
    return y


def _opsi_mendatar(opsi, w, ukuran):
    """Semua opsi pendek, dan kata terpanjangnya muat di kartu (kalau tidak, kata terpotong di tengah)."""
    isi = [polos(v) for v in opsi.values()]
    if not 2 <= len(isi) <= 5 or any(len(v) > 26 for v in isi):
        return False
    muat = ((w - 0.2 * (len(isi) - 1)) / len(isi) - 0.95) * 72
    if any("$" in v and lebar_segs(segmen(v), ukuran) > muat * 0.95 for v in opsi.values()):
        return False  # rumus tidak dipotong di spasi, jadi harus muat satu baris
    return all(lebar_pt(k, ukuran) <= muat * 0.98 for v in isi for k in v.split())


def _ukur_opsi(opsi, w, ukuran):
    """(mendatar, lebar kartu, daftar tinggi baris)."""
    if _opsi_mendatar(opsi, w, ukuran):
        lw = (w - 0.2 * (len(opsi) - 1)) / len(opsi)
        h = max(0.85, max(tinggi_paragraf([P(v, ukuran, "navy", spasi=1.25)], lw - 0.95) + 0.3 for v in opsi.values()))
        return True, lw, [h]
    return False, w, [max(0.56, tinggi_paragraf([P(v, ukuran, "navy", spasi=1.3)], w - 1.25) + 0.2)
                      for v in opsi.values()]


def tinggi_opsi(opsi, w=LEBAR_ISI, ukuran=18):
    if not opsi:
        return 0.0
    mendatar, _, tinggi = _ukur_opsi(opsi, w, ukuran)
    return tinggi[0] if mendatar else sum(tinggi) + 0.12 * (len(tinggi) - 1)


def gambar_opsi(slide, y, opsi, x=KIRI, w=LEBAR_ISI, ukuran=18):
    """Pilihan jawaban dengan lingkaran huruf kuning: mendatar bila semua opsi pendek, selain itu menurun."""
    d = 0.48
    mendatar, lw, tinggi = _ukur_opsi(opsi, w, ukuran)
    for i, (huruf, isi) in enumerate(opsi.items()):
        h = tinggi[0] if mendatar else tinggi[i]
        kx = x + i * (lw + 0.2) if mendatar else x
        p = [P(isi, ukuran, "navy", spasi=1.25 if mendatar else 1.3)]
        kotak(slide, kx, y, lw, h, warna="kartu", radius=min(RADIUS, h / 2), nama=f"Opsi {huruf}")
        lingkaran(slide, kx + 0.18, y + (h - d) / 2, d, "kuning", huruf, 17, "di_kuning", nama=f"Huruf {huruf}")
        teks(slide, kx + 0.82, y, lw - 0.95, h, p, jangkar="m")
        if not mendatar:
            y += h + 0.12
    return y + tinggi[0] if mendatar else y - 0.12
