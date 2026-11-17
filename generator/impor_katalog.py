"""Buat ulang naskah materi dari dokumen Google Docs yang terdaftar di konten/katalog.json.

Unduh setiap dokumen sebagai .docx (Google Docs: File > Download > Microsoft Word) ke satu
folder dengan nama <doc_id>.docx, lalu jalankan:

    python generator/impor_katalog.py --docx folder_docx [--id vs-01 vn-03]

Hanya entri dengan "otomatis": true yang ditimpa. Gambar disimpan di konten/media/<naskah>/.
"""
from __future__ import annotations

import argparse
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import ekstrak_gdoc as E  # noqa: E402

AKAR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
KONTEN = os.path.join(AKAR, "konten")
KELOMPOK = {"Verbal Analogi": 0, "Verbal Silogisme": 1, "Verbal Analitis": 2}


def url_dokumen(doc_id):
    return f"https://docs.google.com/document/d/{doc_id}/edit"


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("--docx", required=True, help="Folder berisi <doc_id>.docx")
    ap.add_argument("--katalog", default=os.path.join(KONTEN, "katalog.json"))
    ap.add_argument("--id", nargs="*", help="Hanya entri ini (mis. vs-01)")
    a = ap.parse_args(argv)
    with open(a.katalog, encoding="utf-8") as f:
        katalog = json.load(f)
    hilang = []
    for m in katalog["materi"]:
        if not m.get("otomatis") or (a.id and m["id"] not in a.id):
            continue
        path = os.path.join(a.docx, m["doc_id"] + ".docx")
        if not os.path.exists(path):
            hilang.append(m["dokumen"])
            continue
        keluar = os.path.join(KONTEN, m["naskah"])
        slug = os.path.splitext(m["naskah"])[0]
        naskah = E.susun(E.baca_docx(path), KONTEN, slug, url_dokumen(m["doc_id"]),
                         KELOMPOK.get(m["kelompok"], 0), f"TIU - {m['bab']}".replace("/", "-"))
        naskah["meta"]["id_katalog"] = m["id"]
        with open(keluar, "w", encoding="utf-8") as f:
            json.dump(naskah, f, ensure_ascii=False, indent=2)
            f.write("\n")
        n_slide = sum(len(s["slide"]) for s in naskah["materi"])
        n_soal = sum(1 for s in naskah["materi"] for x in s["slide"] if x["tipe"] == "contoh_soal")
        print(f"{m['id']}: {naskah['meta']['judul']} — {n_slide} slide isi, {n_soal} contoh soal")
    if hilang:
        print("Belum diunduh:", "; ".join(hilang))


if __name__ == "__main__":
    main()
