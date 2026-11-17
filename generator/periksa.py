"""QA hasil generator: deteksi teks meluap lewat PowerPoint dan render slide ke PNG.

    python generator/periksa.py "output/TIU - Penalaran Logis.pptx" --png folder_gambar
"""
import argparse
import os

import pythoncom
import win32com.client

BATAS_NAV = 10.26 * 72  # tepi atas bilah navigasi (pt)


def _bentuk(shapes):
    for sh in shapes:
        if sh.Type == 6:  # grup
            yield from _bentuk(sh.GroupItems)
        else:
            yield sh


def periksa(path, folder_png=None, lebar_px=1600):
    pythoncom.CoInitialize()
    app = win32com.client.Dispatch("PowerPoint.Application")
    pres = app.Presentations.Open(os.path.abspath(path), True, False, False)
    masalah = []
    try:
        W, H = pres.PageSetup.SlideWidth, pres.PageSetup.SlideHeight
        for s in pres.Slides:
            for sh in _bentuk(s.Shapes):
                frames = []
                if sh.HasTable:
                    t = sh.Table
                    for r in range(1, t.Rows.Count + 1):
                        for c in range(1, t.Columns.Count + 1):
                            frames.append((f"{sh.Name} [{r},{c}]", t.Cell(r, c).Shape, None))
                elif sh.HasTextFrame:
                    frames.append((sh.Name, sh, sh))
                for nama, obj, kotak in frames:
                    tf = obj.TextFrame2
                    if not tf.HasText:
                        continue
                    if kotak is not None:
                        tersedia = kotak.Height - tf.MarginTop - tf.MarginBottom
                        tinggi = tf.TextRange.BoundHeight
                        if tinggi > tersedia + 3:
                            masalah.append(f"slide {s.SlideIndex}: '{nama}' meluap {tinggi - tersedia:.0f}pt "
                                           f"({tf.TextRange.Text[:45]!r})")
                        bawah = kotak.Top + max(kotak.Height, tinggi)
                        # tombol tautan sengaja berada sejajar navigasi (di sebelahnya) pada tema General
                        if kotak.Top < BATAS_NAV < bawah and not nama.startswith(("Nav", "Nomor", "Hak cipta", "Tombol")):
                            masalah.append(f"slide {s.SlideIndex}: '{nama}' menabrak navigasi")
                        if kotak.Left + kotak.Width > W + 1 and kotak.Left > 0:
                            masalah.append(f"slide {s.SlideIndex}: '{nama}' keluar tepi kanan")
            if folder_png:
                os.makedirs(folder_png, exist_ok=True)
                s.Export(os.path.abspath(os.path.join(folder_png, f"slide-{s.SlideIndex:02d}.png")), "PNG",
                         lebar_px, int(lebar_px * H / W))
    finally:
        pres.Close()
        if app.Presentations.Count == 0:
            app.Quit()
    return masalah


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("pptx")
    ap.add_argument("--png")
    a = ap.parse_args()
    hasil = periksa(a.pptx, a.png)
    print("\n".join(hasil) if hasil else "Tidak ada teks yang meluap.")
