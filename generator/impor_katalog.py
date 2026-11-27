"""Buat ulang naskah materi dari dokumen Google Docs yang terdaftar di konten/katalog.json.

Unduh setiap dokumen sebagai .docx (Google Docs: File > Download > Microsoft Word) ke satu
folder dengan nama <doc_id>.docx, lalu jalankan:

    python generator/impor_katalog.py --docx folder_docx [--id vs-01 vn-03]
    python generator/impor_katalog.py --katalog konten/ojk/katalog.json --docx folder_docx_ojk

Hanya entri dengan "otomatis": true yang ditimpa. Gambar disimpan di konten/media/<naskah>/.
Entri dengan "miring": true mempertahankan cetak miring dokumen (materi Bahasa Inggris); entri dengan
"penanda": "merah" memakai teks merah sebagai penanda wajib PPT, seperti highlight kuning (Badges), dan "hijau"
memakai latar hijau muda (UTBK Literasi); "coret": true mempertahankan coretan merah; "pelengkap": "minimal" membatasi
tambahan di luar sorotan. Katalog dengan "kerangka": "utbk" dirapikan dulu oleh rapikan_utbk.py.
Dokumen di katalog["sumber"] mencakup beberapa Bab: dibaca sekali, lalu dibagi pisah_bab.py ke entri materi yang
"sumber"-nya sama (satu deck materi dan satu deck contoh soal per Bab). --id menerima ID dokumen atau ID Bab.
"""
from __future__ import annotations

import argparse
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import ekstrak_gdoc as E  # noqa: E402
import pisah_bab as PB  # noqa: E402
import rapikan_utbk as R  # noqa: E402

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
    # pelengkap.json: tambahan di luar sorotan yang dipilih per dokumen (lihat pelengkap() di ekstrak_gdoc.py)
    path_pilihan = os.path.join(folder, "pelengkap.json")
    pilihan = {}
    if os.path.exists(path_pilihan):
        with open(path_pilihan, encoding="utf-8") as f:
            pilihan = json.load(f)["dokumen"]
    # koreksi.json: koreksi tim QC yang belum ada di dokumen sumber (lihat koreksi() di ekstrak_gdoc.py)
    path_koreksi = os.path.join(folder, "koreksi.json")
    perbaikan = {}
    if os.path.exists(path_koreksi):
        with open(path_koreksi, encoding="utf-8") as f:
            perbaikan = json.load(f)["dokumen"]
    # Entri "bagian": "soal" = baris 'Contoh Soal' di sheet yang soalnya ditulis di dokumen materi (doc_id sama).
    # Entri dengan "sumber" adalah satu Bab dari dokumen di katalog["sumber"] yang mencakup beberapa Bab (pisah_bab.py).
    pasangan = {m["doc_id"]: m for m in katalog["materi"] if m.get("bagian") == "soal" and not m.get("sumber")}
    per_bab = {}
    for m in katalog["materi"]:
        if m.get("sumber"):
            per_bab.setdefault(m["sumber"], []).append(m)
    tugas = [(m, pasangan.get(m["doc_id"]), None) for m in katalog["materi"] if not m.get("sumber")]
    tugas += [(s, None, per_bab.get(s["id"], [])) for s in katalog.get("sumber", [])]
    for m, soal, bab in tugas:
        if m is soal or not m.get("otomatis") or m.get("bagian") == "soal":
            continue
        if a.id and m["id"] not in a.id and not (soal and soal["id"] in a.id) and not any(x["id"] in a.id for x in bab or []):
            continue
        path = os.path.join(a.docx, m["doc_id"] + ".docx")
        if not os.path.exists(path):
            hilang.append(m["dokumen"])
            continue
        slug = m.get("media") or os.path.splitext(m["naskah"])[0]
        aktif = indeks_peta.get(m.get("subtes"), indeks_peta.get(m["kelompok"], 0))
        elemen = E.baca_docx(path, miring=bool(m.get("miring")), penanda=m.get("penanda", ""), coret=bool(m.get("coret")),
                             rumus_gambar=m.get("rumus_gambar"), ganti_baris=katalog.get("kerangka") == "utbk")
        elemen = E.koreksi(elemen, perbaikan.get(m["id"], []))
        if katalog.get("kerangka") == "utbk":
            elemen = R.rapikan(elemen)
        naskah = E.susun(elemen, folder, slug, url_dokumen(m["doc_id"]), aktif,
                         f"{m.get('awalan', awalan)} - {m['bab']}".replace("/", "-"),
                         subtes=katalog.get("nama_subtes", {}).get(m.get("subtes"), subtes), peta=peta,
                         judul=m["bab"] if "peta" in katalog else None,
                         pilihan=pilihan.get(m["id"], []), minimal=m.get("pelengkap") == "minimal",
                         isi_kosong=katalog.get("kerangka") == "utbk")
        keluaran = [(m, naskah)]
        if bab is not None:
            print(f"{m['id']} ({m['bab']}) dibagi per Bab:")
            materi, contoh = E.pisah_soal(naskah, "")
            keluaran = PB.bagi(materi, contoh, bab, m.get("awalan", awalan))
        elif soal:
            materi, contoh = E.pisah_soal(naskah, f"{m.get('awalan', awalan)} - {soal['bab']}".replace("/", "-"))
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
