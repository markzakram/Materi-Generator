"""Bangun situs unduhan statis: semua materi General dengan pilihan sampul putih atau foto, PPTX + PDF.

Situs di folder situs/ tidak butuh server atau generator: cukup index.html, data.js, dan
berkas unduhan. Butuh Windows + Microsoft PowerPoint untuk PDF dan gambar sampul. Slide diwarnai
dengan palet platform JadiASN (--palet tanpa untuk warna template).

    python generator/buat_situs.py --template-general "path/Template Materi General.pptx"

Tambahkan --tema general general-foto opsi-b --template-opsi-b "path/Template Referensi Opsi B -
Ceria & Membulat.pptx" untuk menyediakan Opsi B juga.
"""
from __future__ import annotations

import argparse
import datetime
import io
import json
import os
import re
import shutil
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import buat_materi as BM  # noqa: E402
import desain as D  # noqa: E402

AKAR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
KONTEN = os.path.join(AKAR, "konten")
# Pilihan di situs: id -> (tema generator, varian sampul, nama di halaman, akhiran nama berkas unduhan)
VARIAN = {
    "general": ("general", "putih", "Sampul putih", " (General)"),
    "general-foto": ("general", "foto", "Sampul foto", " (General, sampul foto)"),
    "opsi-b": ("opsi-b", "putih", "Opsi B · Ceria & Membulat", ""),
}
TAMBAHAN = [  # naskah di luar katalog sheet yang ikut ditampilkan
    {"id": "penalaran-logis", "kelompok": "Lainnya", "bab": "Penalaran Logis", "naskah": "tiu_penalaran_logis.json"},
]


def _slug(nama):
    return re.sub(r"[^a-z0-9]+", "-", os.path.splitext(nama)[0].lower()).strip("-")


def _nama_aman(nama):
    return re.sub(r'[\\/:*?"<>|]+', "-", nama).strip(" .-") or "materi"


def _rampingkan_pdf(path):
    """Gabungkan gambar kembar di PDF hasil PowerPoint (lossless).

    PowerPoint menyimpan ulang gambar latar yang sama di setiap halaman, sehingga PDF General
    bisa 15 MB; setelah objek kembar digabung ukurannya sekitar 2 MB dan tampilannya identik.
    """
    try:
        import pymupdf
    except ImportError:
        print("   (PyMuPDF tidak terpasang; PDF tidak dirampingkan)")
        return
    from PIL import Image, ImageChops

    sementara = path + ".tmp"
    with pymupdf.open(path) as doc:
        # PowerPoint menyimpan semua gambar sebagai RGB; hiasan abu-abu (garis gelombang, sekitar
        # 0,7 MB per PDF) cukup satu kanal. Alfa ada di SMask terpisah dan tidak berubah.
        for xref in range(1, doc.xref_length()):
            if (doc.xref_get_key(xref, "Subtype")[1] != "/Image"
                    or doc.xref_get_key(xref, "ColorSpace")[1] != "/DeviceRGB"
                    or doc.xref_get_key(xref, "Filter")[1] != "/FlateDecode"
                    or doc.xref_get_key(xref, "DecodeParms")[1] != "null"
                    or len(doc.xref_stream_raw(xref) or b"") < 100_000):
                continue
            w, h = int(doc.xref_get_key(xref, "Width")[1]), int(doc.xref_get_key(xref, "Height")[1])
            mentah = doc.xref_stream(xref)  # piksel RGB gambar dasar, tanpa SMask
            if len(mentah) != w * h * 3:
                continue
            im = Image.frombytes("RGB", (w, h), mentah)
            r, g, b = im.split()
            beda = ImageChops.lighter(ImageChops.difference(r, g), ImageChops.difference(g, b))
            smask = doc.xref_get_key(xref, "SMask")
            if smask[0] == "xref":  # piksel yang tak terlihat (alfa 0) boleh berwarna apa saja
                alfa = doc.xref_stream(int(smask[1].split()[0]))
                if len(alfa) == w * h:
                    terlihat = Image.frombytes("L", (w, h), alfa).point(lambda a: 255 if a else 0)
                    beda = ImageChops.multiply(beda, terlihat)
            if beda.getextrema()[1] > 8:
                continue
            doc.update_stream(xref, im.convert("L").tobytes())
            doc.xref_set_key(xref, "ColorSpace", "/DeviceGray")
            doc.xref_set_key(xref, "BitsPerComponent", "8")
            doc.xref_set_key(xref, "DecodeParms", "null")
        doc.save(sementara, garbage=4, deflate=True, clean=True)
    os.replace(sementara, path)


def _pdf_dan_sampul(daftar):
    """Ekspor PDF dan gambar slide pertama lewat satu instance PowerPoint."""
    import pythoncom
    import win32com.client
    from PIL import Image

    pythoncom.CoInitialize()
    app = win32com.client.Dispatch("PowerPoint.Application")
    try:
        for pptx, pdf, sampul in daftar:
            pres = app.Presentations.Open(os.path.abspath(pptx), True, False, False)
            try:
                pres.SaveAs(os.path.abspath(pdf), 32)  # 32 = ppSaveAsPDF
                png = sampul[:-4] + ".png"
                pres.Slides(1).Export(os.path.abspath(png), "PNG", 960, 675)
            finally:
                pres.Close()
            _rampingkan_pdf(pdf)
            with Image.open(png) as im:
                im.convert("RGB").save(sampul, "JPEG", quality=82, optimize=True, progressive=True)
            os.remove(png)
            print(f"   PDF ({os.path.getsize(pdf) / 1e6:.1f} MB) + sampul:", os.path.basename(pdf))
    finally:
        if app.Presentations.Count == 0:
            app.Quit()


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("--template-opsi-b", default=BM.TEMPLATE_TEMA["opsi-b"])
    ap.add_argument("--template-general", default=BM.TEMPLATE_TEMA["general"])
    ap.add_argument("--keluar", default=os.path.join(AKAR, "situs"))
    ap.add_argument("--tanpa-pdf", action="store_true", help="Lewati PDF dan sampul (tanpa PowerPoint)")
    ap.add_argument("--tema", nargs="+", choices=list(VARIAN), default=["general", "general-foto"],
                    help="Pilihan yang disediakan di situs (bawaan: General sampul putih dan sampul foto)")
    ap.add_argument("--palet", default="jadiasn",
                    help="Palet platform untuk warna slide (lihat generator/palet.py); 'tanpa' = warna template")
    a = ap.parse_args(argv)
    palet = None if a.palet == "tanpa" else a.palet
    semua = {"opsi-b": a.template_opsi_b, "general": a.template_general}
    template = {v: semua[VARIAN[v][0]] for v in VARIAN if v in a.tema}

    with open(os.path.join(KONTEN, "katalog.json"), encoding="utf-8") as f:
        katalog = json.load(f)
    entri = [dict(m, url=f"https://docs.google.com/document/d/{m['doc_id']}/edit") for m in katalog["materi"]]
    entri += [dict(m, url=None) for m in TAMBAHAN]

    for sub in ("unduh", "sampul"):
        shutil.rmtree(os.path.join(a.keluar, sub), ignore_errors=True)
    for tema in template:
        os.makedirs(os.path.join(a.keluar, "unduh", tema), exist_ok=True)
        os.makedirs(os.path.join(a.keluar, "sampul", tema), exist_ok=True)
    os.makedirs(os.path.join(a.keluar, "font"), exist_ok=True)
    for f in ("Poppins-Regular.ttf", "Poppins-Bold.ttf"):
        shutil.copy(os.path.join(D.ASET, "fonts", f), os.path.join(a.keluar, "font", f))

    data, ekspor = [], []
    for m in entri:
        with open(os.path.join(KONTEN, m["naskah"]), encoding="utf-8") as f:
            naskah = json.load(f)
        slug = _slug(m["naskah"])
        item = {"id": m["id"], "kelompok": m["kelompok"], "bab": m["bab"], "judul": naskah["meta"]["judul"],
                "dokumen": m["url"], "tema": {}}
        for tema, path in template.items():
            tema_gen, sampul_gen, _, akhiran = VARIAN[tema]
            prs, b = BM.buat_deck(naskah, path, KONTEN, tema_gen, sampul=sampul_gen, palet=palet)
            nama = _nama_aman(BM.nama_file(naskah)) + akhiran
            pptx = os.path.join(a.keluar, "unduh", tema, slug + ".pptx")
            pdf = os.path.join(a.keluar, "unduh", tema, slug + ".pdf")
            sampul = os.path.join(a.keluar, "sampul", tema, slug + ".jpg")
            prs.save(pptx)
            ekspor.append((pptx, pdf, sampul))
            item["tema"][tema] = {"slide": len(b.slides), "nama": nama,
                                  "pptx": f"unduh/{tema}/{slug}.pptx", "pdf": f"unduh/{tema}/{slug}.pdf",
                                  "sampul": f"sampul/{tema}/{slug}.jpg"}
            print(f"{m['id']:16s} {tema:12s} {len(b.slides):3d} slide" +
                  (f"  PERINGATAN: {b.peringatan}" if b.peringatan else ""))
        data.append(item)

    if not a.tanpa_pdf:
        _pdf_dan_sampul(ekspor)
    for item in data:  # ukuran berkas untuk ditampilkan di halaman
        for t in item["tema"].values():
            for jenis in ("pptx", "pdf"):
                path = os.path.join(a.keluar, t[jenis])
                t[f"ukuran_{jenis}"] = os.path.getsize(path) if os.path.exists(path) else None

    manifest = {"judul": katalog.get("judul", "Materi"), "sumber_sheet": katalog.get("sumber_sheet"),
                "dibuat": datetime.date.today().isoformat(), "palet": palet,
                "tema": [{"id": t, "nama": VARIAN[t][2]} for t in template], "materi": data}
    with open(os.path.join(a.keluar, "data.js"), "w", encoding="utf-8") as f:
        f.write("// Dibuat oleh generator/buat_situs.py; jangan disunting manual.\n")
        f.write("window.MATERI = " + json.dumps(manifest, ensure_ascii=False, indent=1) + ";\n")
    total = sum(os.path.getsize(os.path.join(r, n)) for r, _, fs in os.walk(a.keluar) for n in fs)
    print(f"{len(data)} materi x {len(template)} tampilan -> {a.keluar} ({total / 1e6:.1f} MB)")


if __name__ == "__main__":
    main()
