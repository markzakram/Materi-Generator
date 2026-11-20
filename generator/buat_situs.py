"""Bangun situs unduhan statis: semua materi General dengan pilihan sampul putih atau foto, PPTX + PDF.

Satu situs (folder situs/) untuk semua aplikasi: setiap katalog menjadi satu tab (aplikasi.json) dengan
data dan berkasnya di situs/<aplikasi>/. Tidak butuh server atau generator. Butuh Windows + Microsoft
PowerPoint untuk PDF dan gambar sampul. Slide dan halaman diwarnai dengan palet di katalog.

    python generator/buat_situs.py --template-general "path/Template Materi General.pptx"
    python generator/buat_situs.py --katalog konten/ojk/katalog.json --template-general "path/..."

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
    ap.add_argument("--katalog", default=os.path.join(KONTEN, "katalog.json"),
                    help="Katalog materi (bawaan TIU; OJK: konten/ojk/katalog.json). Naskah dibaca dari foldernya.")
    ap.add_argument("--keluar", default=os.path.join(AKAR, "situs"),
                    help="Folder situs; materi aplikasi ditulis ke <keluar>/<aplikasi.id katalog>/")
    ap.add_argument("--tanpa-pdf", action="store_true", help="Lewati PDF dan sampul (tanpa PowerPoint)")
    ap.add_argument("--tema", nargs="+", choices=list(VARIAN), default=["general", "general-foto"],
                    help="Pilihan yang disediakan di situs (bawaan: General sampul putih dan sampul foto)")
    ap.add_argument("--palet", default=None,
                    help="Palet platform untuk warna slide (lihat generator/palet.py; bawaan: 'palet' di katalog, "
                         "lalu jadiasn); 'tanpa' = warna template")
    ap.add_argument("--id", nargs="+", default=None,
                    help="Hanya bangun ulang materi ini (mis. sjk-01); materi lain di data.json tetap dipakai")
    a = ap.parse_args(argv)
    semua = {"opsi-b": a.template_opsi_b, "general": a.template_general}
    template = {v: semua[VARIAN[v][0]] for v in VARIAN if v in a.tema}

    konten = os.path.dirname(os.path.abspath(a.katalog))
    with open(a.katalog, encoding="utf-8") as f:
        katalog = json.load(f)
    utama = os.path.samefile(konten, KONTEN)
    palet = a.palet or katalog.get("palet", "jadiasn")
    palet = None if palet == "tanpa" else palet
    app = katalog.get("aplikasi") or {"id": "tiu" if utama else os.path.basename(konten),
                                      "nama": katalog.get("judul", "Materi"), "urutan": 99}
    entri = [dict(m, url=f"https://docs.google.com/document/d/{m['doc_id']}/edit") for m in katalog["materi"]]
    entri += [dict(m, url=None) for m in (TAMBAHAN if utama else [])]
    entri = [m for m in entri if os.path.exists(os.path.join(konten, m["naskah"]))
             or print(f"Lewati {m['id']}: naskah {m['naskah']} belum ada")]

    # Satu situs untuk semua aplikasi: index.html memuat aplikasi.json (daftar tab) dan
    # <aplikasi>/data.json; berkas unduhan di <keluar>/<aplikasi>/unduh dan /sampul.
    folder_app = os.path.join(a.keluar, app["id"])
    lama = {}
    if a.id:  # bangun sebagian: materi lain diambil dari data.json yang ada
        path_data = os.path.join(folder_app, "data.json")
        if not os.path.exists(path_data):
            sys.exit(f"--id butuh {path_data} dari build sebelumnya; jalankan dulu tanpa --id.")
        with open(path_data, encoding="utf-8") as f:
            lama = {x["id"]: x for x in json.load(f)["materi"]}
        tidak_ada = set(a.id) - {m["id"] for m in entri}
        if tidak_ada:
            sys.exit(f"Tidak ada di katalog (atau naskahnya belum ada): {', '.join(sorted(tidak_ada))}")
    else:
        for sub in ("unduh", "sampul"):
            shutil.rmtree(os.path.join(folder_app, sub), ignore_errors=True)
    for tema in template:
        os.makedirs(os.path.join(folder_app, "unduh", tema), exist_ok=True)
        os.makedirs(os.path.join(folder_app, "sampul", tema), exist_ok=True)
    os.makedirs(os.path.join(a.keluar, "font"), exist_ok=True)
    for f in ("Poppins-Regular.ttf", "Poppins-Bold.ttf"):
        shutil.copy(os.path.join(D.ASET, "fonts", f), os.path.join(a.keluar, "font", f))

    data, ekspor = [], []
    for m in entri:
        if a.id and m["id"] not in a.id:
            if m["id"] in lama:
                data.append(lama[m["id"]])
            else:
                print(f"Lewati {m['id']}: belum pernah dibangun (jalankan tanpa --id)")
            continue
        with open(os.path.join(konten, m["naskah"]), encoding="utf-8") as f:
            naskah = json.load(f)
        slug = _slug(m["naskah"])
        item = {"id": m["id"], "kelompok": m["kelompok"], "bab": m["bab"], "judul": naskah["meta"]["judul"],
                "dokumen": m["url"], "tema": {}}
        for tema, path in template.items():
            tema_gen, sampul_gen, _, akhiran = VARIAN[tema]
            prs, b = BM.buat_deck(naskah, path, konten, tema_gen, sampul=sampul_gen, palet=palet)
            nama = _nama_aman(BM.nama_file(naskah)) + akhiran
            pptx = os.path.join(folder_app, "unduh", tema, slug + ".pptx")
            pdf = os.path.join(folder_app, "unduh", tema, slug + ".pdf")
            sampul = os.path.join(folder_app, "sampul", tema, slug + ".jpg")
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
                path = os.path.join(folder_app, t[jenis])
                t[f"ukuran_{jenis}"] = os.path.getsize(path) if os.path.exists(path) else None

    manifest = {"judul": katalog.get("judul", "Materi"), "merek": katalog.get("merek", "Cerebrum · Materi"),
                "sumber_sheet": katalog.get("sumber_sheet"),
                "dibuat": datetime.date.today().isoformat(), "palet": palet,
                "tema": [{"id": t, "nama": VARIAN[t][2]} for t in template], "materi": data}
    simpan_aplikasi(a.keluar, app, manifest)
    ukuran = lambda d: sum(os.path.getsize(os.path.join(r, n)) for r, _, fs in os.walk(d) for n in fs)  # noqa: E731
    dibangun = f"{len(a.id)} dibangun ulang, " if a.id else ""
    print(f"{len(data)} materi ({dibangun}{len(template)} tampilan) -> {folder_app} ({ukuran(folder_app) / 1e6:.1f} MB; "
          f"seluruh situs {ukuran(a.keluar) / 1e6:.1f} MB)")


def simpan_aplikasi(situs, app, manifest):
    """Tulis <situs>/<id>/data.json (dengan warna halaman dari palet) dan perbarui daftar tab aplikasi.json."""
    import palet as PL

    manifest = dict(manifest, css=PL.css_situs(manifest.get("palet") or "jadiasn"))
    folder = os.path.join(situs, app["id"])
    os.makedirs(folder, exist_ok=True)
    with open(os.path.join(folder, "data.json"), "w", encoding="utf-8") as f:
        json.dump(manifest, f, ensure_ascii=False, indent=1)
        f.write("\n")
    path = os.path.join(situs, "aplikasi.json")
    daftar = []
    if os.path.exists(path):
        with open(path, encoding="utf-8") as f:
            daftar = json.load(f)
    daftar = [x for x in daftar if x["id"] != app["id"]]
    daftar.append({"id": app["id"], "nama": app["nama"], "urutan": app.get("urutan", 99)})
    daftar.sort(key=lambda x: (x["urutan"], x["id"]))
    with open(path, "w", encoding="utf-8") as f:
        json.dump(daftar, f, ensure_ascii=False, indent=1)
        f.write("\n")


if __name__ == "__main__":
    main()
