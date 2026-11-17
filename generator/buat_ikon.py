"""Buat varian warna ikon untuk desain Opsi B (navy dan putih).

Ikon garis di generator/assets berwarna hitam atau teal. Template Opsi B memakai
ikon yang sama dalam warna navy (di atas lingkaran kuning atau latar putih) dan
putih (di atas kartu navy). Skrip ini hanya mengganti warna dan mempertahankan
transparansi. Jalankan ulang bila ikon sumber berubah:

    python generator/buat_ikon.py
"""
from __future__ import annotations

import os

from PIL import Image

ASET = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets")
IKON_GARIS = ["checklist", "globe", "grafik", "roket", "strategi", "bingung", "presenter"]
WARNA = {"navy": (0x13, 0x21, 0x5A), "putih": (0xFF, 0xFF, 0xFF)}


def main():
    for nama_warna, rgb in WARNA.items():
        folder = os.path.join(ASET, "ikon", nama_warna)
        os.makedirs(folder, exist_ok=True)
        for nama in IKON_GARIS:
            with Image.open(os.path.join(ASET, f"ikon_{nama}.png")) as im:
                alfa = im.convert("RGBA").getchannel("A")
            hasil = Image.new("RGBA", alfa.size, rgb + (0,))
            hasil.putalpha(alfa)
            hasil.save(os.path.join(folder, f"ikon_{nama}.png"), optimize=True)
        print(f"{nama_warna}: {len(IKON_GARIS)} ikon -> {folder}")


if __name__ == "__main__":
    main()
