"""Palet platform: warnai ulang deck jadi (XML, SVG, PNG) dari warna template ke warna platform.

Template General memakai navy 13215A dan kuning FBBB2B. Palet platform mengganti warna itu di
semua bentuk (termasuk gradien), di gambar SVG, dan di gambar PNG berwarna brand (ikon navy,
lengkungan kuning sampul foto) dengan menggeser rona piksel yang dekat warna asal. Foto JPEG
tidak disentuh.

Warna platform diambil dari proyek ICON Aplikasi generator, catalog/brands.json.
"""
from __future__ import annotations

import io
import re

from pptx.oxml.ns import qn

PALET = {
    "jadiasn": {
        "nama": "JadiASN",
        # peran dari brands.json: primary, primary_shade, primary_alt, accent, accent_shade, light
        "peran": {"primer": "3D6E99", "navy": "052C5F", "primer_muda": "4C8FC1",
                  "aksen": "E8B41E", "aksen_tua": "876000", "terang": "F9F3F1"},
        # warna template/generator -> warna platform
        "ganti": {
            "13215A": "052C5F",  # navy
            "24336F": "214D7C",  # navy muda: antara navy dan biru primer
            "FBBB2B": "E8B41E",  # kuning
            "FFE6AE": "F8E8BB",  # kuning muda: aksen 30% di atas putih
            "595858": "3D6E99",  # ujung gradien kartu aktif: abu -> biru primer
            "C68700": "3D6E99",  # judul materi di sampul putih (lihat tema_general.WARNA_JUDUL)
        },
        # pasangan rona untuk gambar PNG
        "rona": [("13215A", "052C5F"), ("FBBB2B", "E8B41E")],
    },
}

TOLERANSI_RONA = 8  # dari 256 (sekitar 11 derajat)
SATURASI_MIN = 90  # dari 255; abu-abu, putih, dan kulit tidak ikut berubah


def _hsv(hexa):
    from PIL import Image
    return Image.new("RGB", (1, 1), "#" + hexa).convert("HSV").getpixel((0, 0))


def _warnai_png(blob, rona):
    """Geser rona piksel yang dekat warna asal ke warna tujuan; alfa dan piksel lain tetap."""
    from PIL import Image, ImageChops

    with Image.open(io.BytesIO(blob)) as asli:
        asli.load()
        mode = asli.mode
        rgba = asli.convert("RGBA")
    alfa = rgba.getchannel("A")
    rgb = rgba.convert("RGB")
    h, s, v = rgb.convert("HSV").split()
    hasil = rgb
    berubah = False
    for asal, tujuan in rona:
        ha, sa, va = _hsv(asal)
        ht, st, vt = _hsv(tujuan)
        dekat = h.point(lambda x: 255 if min(abs(x - ha), 256 - abs(x - ha)) <= TOLERANSI_RONA else 0)
        jenuh = s.point(lambda x: 255 if x >= SATURASI_MIN else 0)
        topeng = ImageChops.multiply(dekat, jenuh)
        if not topeng.getbbox():
            continue
        berubah = True
        h2 = h.point(lambda x: (x + ht - ha) % 256)
        s2 = s.point(lambda x: max(0, min(255, round(x * st / max(sa, 1)))))
        v2 = v.point(lambda x: max(0, min(255, round(x * vt / max(va, 1)))))
        baru = Image.merge("HSV", (h2, s2, v2)).convert("RGB")
        hasil = Image.composite(baru, hasil, topeng)
    if not berubah:
        return None
    hasil.putalpha(alfa)
    keluar = io.BytesIO()
    if mode == "P":  # pertahankan ukuran kecil gambar berpalet
        hasil.quantize(colors=256, method=Image.Quantize.FASTOCTREE).save(keluar, "PNG", optimize=True)
    else:
        hasil.save(keluar, "PNG", optimize=True)
    return keluar.getvalue()


def terapkan(prs, nama):
    """Warnai ulang seluruh presentasi dengan palet nama (mis. 'jadiasn')."""
    if not nama:
        return
    if nama not in PALET:
        raise ValueError(f"Palet tidak dikenal: {nama} (pilihan: {', '.join(PALET)})")
    p = PALET[nama]
    ganti = {k.upper(): v for k, v in p["ganti"].items()}
    pola_svg = re.compile("#(" + "|".join(ganti) + ")", re.IGNORECASE)
    for part in prs.part.package.iter_parts():
        el = getattr(part, "_element", None)
        if el is not None:
            for c in el.iter(qn("a:srgbClr")):
                baru = ganti.get((c.get("val") or "").upper())
                if baru:
                    c.set("val", baru)
            continue
        jenis = getattr(part, "content_type", "")
        if jenis == "image/svg+xml":
            teks = part.blob.decode("utf-8", "ignore")
            hasil = pola_svg.sub(lambda m: "#" + ganti[m.group(1).upper()], teks)
            if hasil != teks:
                part._blob = hasil.encode("utf-8")
        elif jenis == "image/png":
            hasil = _warnai_png(part.blob, p["rona"])
            if hasil is not None:
                part._blob = hasil
