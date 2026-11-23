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
KELOMPOK = {"Verbal Analogi": 0, "Verbal Silogisme": 1, "Verbal Analitis": 2, "Numerik": 3, "Data & Deret": 4, "Figural": 5}


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
    # Entri "bagian": "soal" = baris 'Contoh Soal' di sheet yang soalnya ditulis di dokumen materi (doc_id sama).
    pasangan = {m["doc_id"]: m for m in katalog["materi"] if m.get("bagian") == "soal"}
    for m in katalog["materi"]:
        soal = pasangan.get(m["doc_id"])
        if m is soal or not m.get("otomatis"):
            continue
        if a.id and m["id"] not in a.id and not (soal and soal["id"] in a.id):
            continue
        path = os.path.join(a.docx, m["doc_id"] + ".docx")
        if not os.path.exists(path):
            hilang.append(m["dokumen"])
            continue
        slug = os.path.splitext(m["naskah"])[0]
        aktif = indeks_peta.get(m.get("subtes"), indeks_peta.get(m["kelompok"], 0))
        naskah = E.susun(E.baca_docx(path), folder, slug, url_dokumen(m["doc_id"]), aktif,
                         f"{awalan} - {m['bab']}".replace("/", "-"), subtes=subtes, peta=peta,
                         judul=m["bab"] if "peta" in katalog else None)
        keluaran = [(m, naskah)]
        if soal:
            materi, contoh = E.pisah_soal(naskah, f"{awalan} - {soal['bab']}".replace("/", "-"))
            keluaran = [(m, materi), (soal, contoh)]
        for entri, isi in keluaran:
            isi["meta"]["id_katalog"] = entri["id"]
            with open(os.path.join(folder, entri["naskah"]), "w", encoding="utf-8") as f:
                json.dump(isi, f, ensure_ascii=False, indent=2)
                f.write("\n")
            slide = [x for s in isi["materi"] for x in s["slide"]]
            tanpa_kunci = [x["judul"] for x in slide if x["tipe"] == "contoh_soal" and x["opsi"] and not x["kunci"]]
            print(f"{entri['id']}: {isi['meta']['judul']} — {len(slide)} slide isi, "
                  f"{sum(x['tipe'] == 'contoh_soal' for x in slide)} contoh soal"
                  + (f"  TANPA KUNCI: {', '.join(tanpa_kunci)}" if tanpa_kunci else ""))
    if hilang:
        print("Belum diunduh:", "; ".join(hilang))


if __name__ == "__main__":
    main()
