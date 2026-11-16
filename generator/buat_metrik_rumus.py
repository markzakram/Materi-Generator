"""Buat tabel lebar glyph Cambria Math untuk pengukur teks rumus.

Generator mengukur lebar teks agar bisa menebak pindah baris. Cambria Math
milik Microsoft dan tidak ada di server Linux (Vercel), jadi yang disimpan di
repo hanya angka lebarnya: lebar tiap karakter dan pasangan kerning, persis
seperti hasil Pillow (layout dasar) pada ukuran 400 px. Fontnya sendiri tidak
ikut didistribusikan.

Jalankan ulang di Windows hanya bila ingin memperbarui tabel:

    python generator/buat_metrik_rumus.py
"""
from __future__ import annotations

import argparse
import json
import os

from fontTools.ttLib import TTCollection
from PIL import ImageFont

UKURAN = 400
KELUAR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", "metrik_cambria_math.json")


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("--font", default="C:/Windows/Fonts/cambria.ttc")
    ap.add_argument("--indeks", type=int, default=1, help="Indeks Cambria Math di dalam .ttc")
    ap.add_argument("--keluar", default=KELUAR)
    a = ap.parse_args(argv)

    tt = TTCollection(a.font).fonts[a.indeks]
    nama = tt["name"].getDebugName(4)
    cmap = tt.getBestCmap()
    ft = ImageFont.truetype(a.font, UKURAN, index=a.indeks, layout_engine=ImageFont.Layout.BASIC)

    def bulat(x):
        return round(float(x), 3)

    lebar = {chr(cp): bulat(ft.getlength(chr(cp))) for cp in sorted(cmap) if cp >= 0x20}
    kosong = next(cp for cp in range(0xE000, 0xF900) if cp not in cmap)
    notdef = bulat(ft.getlength(chr(kosong)))

    # pasangan kerning dari tabel 'kern' (yang dipakai Pillow pada layout dasar)
    glyph_ke_char = {}
    for cp, g in cmap.items():
        if cp >= 0x20:
            glyph_ke_char.setdefault(g, []).append(chr(cp))
    kern = {}
    if "kern" in tt:
        for sub in tt["kern"].kernTables:
            for (gk, gn) in getattr(sub, "kernTable", {}):
                for ck in glyph_ke_char.get(gk, []):
                    for cn in glyph_ke_char.get(gn, []):
                        k = bulat(ft.getlength(ck + cn) - ft.getlength(ck) - ft.getlength(cn))
                        if k:
                            kern[ck + cn] = k

    data = {
        "keterangan": f"Lebar glyph {nama} dalam px pada ukuran {UKURAN} (Pillow, layout dasar). "
                      "Dibuat oleh generator/buat_metrik_rumus.py; jangan disunting manual.",
        "ukuran": UKURAN,
        "notdef": notdef,
        "lebar": lebar,
        "kern": kern,
    }
    with open(a.keluar, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, separators=(",", ":"), sort_keys=True)
    print(f"{nama}: {len(lebar)} karakter, {len(kern)} pasangan kerning -> {a.keluar}")


if __name__ == "__main__":
    main()
