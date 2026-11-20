"""Buat ulang naskah materi dari dokumen Google Docs yang terdaftar di konten/katalog.json.

Unduh setiap dokumen sebagai .docx (Google Docs: File > Download > Microsoft Word) ke satu
folder dengan nama <doc_id>.docx, lalu jalankan:

    python generator/impor_katalog.py --docx folder_docx [--id vs-01 vn-03]
    python generator/impor_katalog.py --katalog konten/ojk/katalog.json --docx folder_docx_ojk

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
    # Pengaturan per katalog (bawaan: TIU). Naskah disimpan di folder katalog, gambar di media/.
    folder = os.path.dirname(os.path.abspath(a.katalog))
    subtes = katalog.get("subtes", E.SUBTES_TIU)
    peta = katalog.get("peta")
    indeks_peta = katalog.get("indeks_peta", KELOMPOK)
    awalan = katalog.get("awalan_nama", "TIU")
    hilang = []
    for m in katalog["materi"]:
        if not m.get("otomatis") or (a.id and m["id"] not in a.id):
            continue
        path = os.path.join(a.docx, m["doc_id"] + ".docx")
        if not os.path.exists(path):
            hilang.append(m["dokumen"])
            continue
        keluar = os.path.join(folder, m["naskah"])
        slug = os.path.splitext(m["naskah"])[0]
        aktif = indeks_peta.get(m.get("subtes"), indeks_peta.get(m["kelompok"], 0))
        naskah = E.susun(E.baca_docx(path), folder, slug, url_dokumen(m["doc_id"]), aktif,
                         f"{awalan} - {m['bab']}".replace("/", "-"), subtes=subtes, peta=peta,
                         judul=m["bab"] if "peta" in katalog else None)
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
