"""Ubah dokumen materi Google Docs (ekspor .docx) menjadi naskah JSON generator.

Dokumen mengikuti templat booklet materi SKD: Judul, Pendahuluan, Tujuan Pembelajaran,
Materi (Heading 2 = subbab, Heading 3 = slide), Contoh Soal dan Pembahasan, Penutup.
Sebagian dokumen memakai Heading 1 langsung sebagai subbab; keduanya didukung.

Isi tidak dikarang ulang: paragraf, tabel, gambar, dan persamaan diambil apa adanya.
Bagian yang tidak ada di dokumen (pertanyaan pemantik, level Bloom) dibiarkan kosong
agar diisi tim akademik lewat editor.

    python generator/ekstrak_gdoc.py "01 Susunan & Posisi.docx" --keluar konten/tiu_vn01.json \
        --sumber https://docs.google.com/document/d/... --kelompok 2
"""
from __future__ import annotations

import argparse
import json
import os
import re
import sys
import zipfile

from lxml import etree

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
}
W, M = "{%s}" % NS["w"], "{%s}" % NS["m"]
EMU = 914400

SUBTES_TIU = "Tes Intelegensia Umum"
PETA_TIU = {
    "judul": SUBTES_TIU,
    "subjudul": "Peta Materi Subtes TIU",
    "kartu": [
        {"judul": "Verbal Analogi", "deskripsi": "Jenis hubungan antarkata dan analogi dua kata", "ikon": "checklist"},
        {"judul": "Verbal Silogisme", "deskripsi": "Operator logika, kondisional, ekuivalensi, kuantor", "ikon": "strategi"},
        {"judul": "Verbal Analitis", "deskripsi": "Posisi, urutan, jadwal, rute, dan argumen", "ikon": "globe"},
        {"judul": "Numerik", "deskripsi": "Berhitung, aljabar, geometri, dan perbandingan", "ikon": "grafik"},
        {"judul": "Data & Deret", "deskripsi": "Kecukupan data, tabel pernyataan, deret angka", "ikon": "roket"},
        {"judul": "Figural", "deskripsi": "Transformasi, analogi, serial, dan matriks gambar", "ikon": "strategi"},
    ],
}
IKON_SUBBAB = ["checklist", "globe", "strategi", "grafik", "roket"]
SIMBOL_SOROT = set("→←↔⟷=≡¬∧∨⊬∴⇒⇔①②③[]•×÷≠≤≥")


# ------------------------------------------------------------ baca docx ---
def _tebal(r):
    b = r.find(f"{W}rPr/{W}b")
    return b is not None and b.get(f"{W}val") not in ("0", "false")


def _miring(r):
    i = r.find(f"{W}rPr/{W}i")
    return i is not None and i.get(f"{W}val") not in ("0", "false")


def _gabung(potongan):
    """[(teks, jenis)] -> markup: jenis 'b' = tebal, 'm' = rumus, '' = biasa."""
    keluar, kini, isi = [], None, ""
    for t, j in potongan + [("", "#")]:
        if j != kini:
            if isi:
                if kini == "b" and isi.strip():
                    awal, akhir = len(isi) - len(isi.lstrip()), len(isi.rstrip())
                    keluar.append(isi[:awal] + "**" + isi[awal:akhir] + "**" + isi[akhir:])
                elif kini == "m" and isi.strip():
                    keluar.append("$" + isi.strip() + "$")
                else:
                    keluar.append(isi)
            kini, isi = j, ""
        isi += t
    return re.sub(r"[ \t]+", " ", "".join(keluar)).strip()


def _teks_paragraf(p):
    """Teks paragraf dengan markup, daftar gambar (rId, lebar, tinggi), dan status tebal/miring penuh."""
    potongan, gambar = [], []
    semua_tebal, semua_miring, ada_teks = True, True, False

    def jalan(el):
        nonlocal semua_tebal, semua_miring, ada_teks
        for anak in el:
            tag = anak.tag
            if tag == f"{W}r":
                for d in anak.iter(f"{W}drawing"):
                    blip = d.find(".//{%s}blip" % NS["a"])
                    ext = d.find(".//{%s}extent" % NS["wp"])
                    if blip is not None:
                        gambar.append((blip.get("{%s}embed" % NS["r"]),
                                       int(ext.get("cx")) / EMU if ext is not None else None,
                                       int(ext.get("cy")) / EMU if ext is not None else None))
                t = ""
                for x in anak:
                    if x.tag == f"{W}t":
                        t += x.text or ""
                    elif x.tag in (f"{W}tab", f"{W}br"):
                        t += " "
                if t:
                    if t.strip():
                        ada_teks = True
                        semua_tebal &= _tebal(anak)
                        semua_miring &= _miring(anak)
                    potongan.append((t, "b" if _tebal(anak) else ""))
            elif tag in (f"{M}oMath", f"{M}oMathPara"):
                t = "".join(x.text or "" for x in anak.iter(f"{M}t"))
                if t.strip():
                    ada_teks = True
                    semua_tebal = False
                    potongan.append((" " + t + " ", "m"))
            elif tag in (f"{W}hyperlink", f"{W}smartTag", f"{W}ins", f"{W}sdt", f"{W}sdtContent"):
                jalan(anak)

    jalan(p)
    return _gabung(potongan), gambar, (semua_tebal and ada_teks), (semua_miring and ada_teks)


def baca_docx(path):
    """Daftar elemen berurutan: judul/subjudul/h1/h2/h3/paragraf/tabel/gambar."""
    z = zipfile.ZipFile(path)
    doc = etree.fromstring(z.read("word/document.xml"))
    gaya = etree.fromstring(z.read("word/styles.xml"))
    nama_gaya = {}
    for s in gaya.iter(f"{W}style"):
        n = s.find(f"{W}name")
        nama_gaya[s.get(f"{W}styleId")] = (n.get(f"{W}val") if n is not None else s.get(f"{W}styleId")).lower()
    rels = etree.fromstring(z.read("word/_rels/document.xml.rels"))
    target = {r.get("Id"): r.get("Target") for r in rels}

    elemen = []
    for el in doc.find(f"{W}body"):
        if el.tag == f"{W}p":
            ps = el.find(f"{W}pPr/{W}pStyle")
            st = nama_gaya.get(ps.get(f"{W}val"), "") if ps is not None else ""
            daftar = el.find(f"{W}pPr/{W}numPr") is not None
            teks, gambar, tebal, miring = _teks_paragraf(el)
            jenis = {"title": "judul", "subtitle": "subjudul", "heading 1": "h1", "heading 2": "h2",
                     "heading 3": "h3", "heading 4": "h3"}.get(st, "paragraf")
            for rid, w, h in gambar:
                berkas = target.get(rid, "")
                elemen.append({"jenis": "gambar", "data": z.read("word/" + berkas) if berkas else b"",
                               "ext": os.path.splitext(berkas)[1].lower(), "w": w, "h": h})
            polos = teks.replace("**", "").strip()
            if polos or jenis != "paragraf":
                elemen.append({"jenis": jenis, "teks": teks, "polos": polos, "tebal": tebal,
                               "miring": miring, "daftar": daftar})
        elif el.tag == f"{W}tbl":
            baris = []
            for tr in el.findall(f"{W}tr"):
                sel = []
                for tc in tr.findall(f"{W}tc"):
                    tcPr = tc.find(f"{W}tcPr")
                    rentang = tcPr.find(f"{W}gridSpan") if tcPr is not None else None
                    lanjut = tcPr.find(f"{W}vMerge") if tcPr is not None else None
                    isi = [_teks_paragraf(p)[0] for p in tc.findall(f".//{W}p")]
                    teks = "" if lanjut is not None and lanjut.get(f"{W}val") != "restart" else \
                        " ".join(x for x in isi if x)
                    sel.append(teks)
                    # sel gabungan mendatar diperluas agar kolom tetap sejajar
                    sel.extend([""] * (int(rentang.get(f"{W}val")) - 1 if rentang is not None else 0))
                baris.append(sel)
            if baris:
                elemen.append({"jenis": "tabel", "baris": baris})
    return elemen


# ------------------------------------------------------------- susun ---
def _bersih(judul):
    t = judul.replace("**", "").strip()
    return re.sub(r"^\d+[.)]?\s+", "", t).strip()


def _kalimat_pertama(teks):
    t = teks.replace("**", "")
    m = re.match(r"(.+?[.!?])(\s|$)", t)
    return (m.group(1) if m else t).strip()


class Penyusun:
    def __init__(self, folder_naskah, slug):
        self.folder_naskah, self.slug = folder_naskah, slug
        self.n_gambar = 0

    def simpan_gambar(self, el):
        if el["ext"] not in (".png", ".jpg", ".jpeg", ".gif", ".bmp"):
            return None
        self.n_gambar += 1
        rel = f"media/{self.slug}/gambar-{self.n_gambar:02d}{el['ext']}"
        path = os.path.join(self.folder_naskah, rel)
        os.makedirs(os.path.dirname(path), exist_ok=True)
        with open(path, "wb") as f:
            f.write(el["data"])
        return rel

    def blok_dari(self, elemen):
        """Elemen isi -> blok naskah (paragraf, subjudul, sorot, poin, gambar, tabel, catatan)."""
        blok, i = [], 0
        while i < len(elemen):
            el = elemen[i]
            j = el["jenis"]
            if j == "gambar":
                rel = self.simpan_gambar(el)
                if rel:
                    b = {"jenis": "gambar", "file": rel, "lebar": round(el["w"] or 4, 2), "tinggi": round(el["h"] or 3, 2)}
                    if i + 1 < len(elemen) and elemen[i + 1]["jenis"] == "paragraf" and \
                            re.match(r"^(Gambar|Diagram|Ilustrasi)\s*\d*[.:]", elemen[i + 1]["polos"]):
                        b["keterangan"] = elemen[i + 1]["teks"]
                        i += 1
                    blok.append(b)
            elif j == "tabel":
                baris = [r for r in el["baris"] if any(c.strip() for c in r)]
                if len(baris) >= 2 and len(baris[0]) >= 1:
                    n = max(len(r) for r in baris)
                    baris = [r + [""] * (n - len(r)) for r in baris]
                    blok.append({"jenis": "tabel", "kolom": [c.replace("**", "") for c in baris[0]],
                                 "baris": baris[1:]})
                elif baris:
                    for r in baris:
                        blok.append({"jenis": "paragraf", "teks": " · ".join(c for c in r if c)})
            elif j in ("h3", "h2", "h1"):
                blok.append({"jenis": "subjudul", "teks": _bersih(el["teks"])})
            else:
                t, polos = el["teks"], el["polos"]
                if re.fullmatch(r"cara cepat:?", polos.lower()) and i + 1 < len(elemen) and elemen[i + 1]["jenis"] == "paragraf":
                    blok.append({"jenis": "catatan", "catatan": "tips", "judul": "Cara cepat",
                                 "teks": elemen[i + 1]["teks"]})
                    i += 1
                elif el["daftar"] or re.match(r"^[•▪●◦\-–]\s+", polos):
                    butir = re.sub(r"^(\*\*)?[•▪●◦\-–]\s+", r"\1", t)
                    if blok and blok[-1]["jenis"] == "poin":
                        blok[-1]["butir"].append(butir)
                    else:
                        blok.append({"jenis": "poin", "butir": [butir]})
                elif re.fullmatch(r"\$[^$]+\$", t.strip()):
                    blok.append({"jenis": "sorot", "teks": t.strip()})
                elif el["tebal"] and (any(c in SIMBOL_SOROT for c in polos.rstrip(":")) or
                                      re.search(r"\w\s+[:=]\s+\w", polos)):
                    blok.append({"jenis": "sorot", "teks": polos})
                elif el["tebal"] and len(polos) <= 90:
                    blok.append({"jenis": "subjudul", "teks": polos.rstrip(":")})
                elif re.match(r"^(Langkah\s+\d+\b|Kesimpulan$|Pemisalan$)", polos) and len(polos) <= 90:
                    blok.append({"jenis": "subjudul", "teks": polos})
                else:
                    blok.append({"jenis": "paragraf", "teks": t})
            i += 1
        return blok


def _bagi_bagian(elemen):
    """[(judul_h1, [elemen...])] beserta judul dan subjudul dokumen."""
    judul, subjudul, bagian, kini = "", "", [], None
    for el in elemen:
        if el["jenis"] == "judul":
            judul = el["polos"]
        elif el["jenis"] == "subjudul":
            subjudul = el["polos"]
        elif el["jenis"] == "h1":
            if not el["polos"]:
                continue  # heading kosong di dokumen: abaikan
            kini = (el["polos"], [])
            bagian.append(kini)
        elif kini is not None:
            kini[1].append(el)
    return judul, subjudul, bagian


def _soal(penyusun, elemen):
    """Bagian 'Contoh Soal dan Pembahasan' -> daftar slide contoh_soal."""
    hasil, label, kini, mode = [], None, None, None
    for el in elemen:
        j, polos = el["jenis"], el.get("polos", "")
        if j in ("h2", "h3") and re.match(r"^soal\s*\d+", polos, re.I):
            kini = {"tipe": "contoh_soal", "judul": _bersih(el["teks"]), "soal": [], "opsi": {}, "kunci": "",
                    "pembahasan": []}
            if label:
                kini["label"] = label
            hasil.append(kini)
            mode = "soal"
            continue
        if j in ("h2", "h3"):
            label = _bersih(el["teks"])
            continue
        if kini is None:
            continue
        m_jawab = re.match(r"^jawaban\b[^:]{0,20}:?\s*\(?([A-E])\)?\.?$", polos, re.I)
        if mode in ("soal", "opsi") and m_jawab:
            kini["kunci"] = m_jawab.group(1).upper()
            mode = "kunci"
            continue
        if re.fullmatch(r"pembahasan\s*:?", polos, re.I):
            mode = "bahas"
            continue
        m_opsi = re.match(r"^\(?([A-E])[.)]\s+(.*)$", polos) if j == "paragraf" else None
        if mode in ("soal", "opsi") and m_opsi and (m_opsi.group(1) == "A" or mode == "opsi"):
            teks = re.sub(r"^(\*\*)?\(?[A-E][.)]\s+", r"\1", el["teks"])
            kini["opsi"][m_opsi.group(1)] = teks
            mode = "opsi"
            continue
        kini["pembahasan" if mode in ("bahas", "kunci") else "soal"].append(el)
    for s in hasil:
        s["soal"] = penyusun.blok_dari(s["soal"])
        s["pembahasan"] = penyusun.blok_dari(s["pembahasan"])
    return hasil


def susun(elemen, folder_naskah, slug, sumber="", kelompok=0, nama_file=None, subtes=SUBTES_TIU):
    penyusun = Penyusun(folder_naskah, slug)
    judul, subjudul, bagian = _bagi_bagian(elemen)
    judul_bersih = re.sub(r"^\d{2}\s+", "", judul).strip() or slug  # buang nomor urut "01 "
    naskah = {
        "meta": {"judul": judul_bersih, "subtes": subtes,
                 "nama_file": nama_file or f"TIU - {judul_bersih}", "sumber": sumber},
        "peta_subtes": dict(PETA_TIU, aktif=kelompok),
        "pendahuluan": {"paragraf": [], "pemantik": "", "ikon": "presenter"},
        "tujuan": {"pengantar": "Setelah mempelajari materi ini, peserta mampu:", "butir": []},
        "materi": [],
        "penutup": {"judul": "Yang Perlu Diingat", "ringkasan": [], "pesan": ""},
    }
    if subjudul and "•" not in subjudul:
        naskah["meta"]["subjudul"] = subjudul

    def tambah_subbab(nama, isi, h2_jadi_slide):
        """Elemen satu subbab -> slide 'blok'. Heading 2 memulai slide baru hanya bila subbabnya
        Heading 1; Heading 3 menjadi subjudul dan generator memecah halaman bila penuh."""
        nama = _bersih(nama)
        slides, kini = [], {"tipe": "blok", "judul": nama, "blok": []}
        pemecah = ("h2",) if h2_jadi_slide else ()
        buf = []

        def tutup():
            if buf or not slides:
                kini["blok"] = penyusun.blok_dari(buf)
                if kini["blok"]:
                    slides.append(dict(kini))

        for el in isi:
            if el["jenis"] in pemecah and el["polos"]:
                tutup()
                buf = []
                kini = {"tipe": "blok", "judul": _bersih(el["teks"]), "blok": []}
            else:
                buf.append(el)
        tutup()
        if slides:
            ikon = IKON_SUBBAB[len(naskah["materi"]) % len(IKON_SUBBAB)]
            naskah["materi"].append({"subbab": nama, "ikon": ikon, "slide": slides})

    for nama, isi in bagian:
        kunci = nama.lower().strip(" :")
        if kunci == "pendahuluan":
            blok = penyusun.blok_dari(isi)
            naskah["pendahuluan"]["paragraf"] = [b["teks"] for b in blok if b["jenis"] == "paragraf"]
            gambar = [b for b in blok if b["jenis"] == "gambar"]
            if gambar:
                naskah["pendahuluan"]["gambar"] = gambar[0]["file"]
        elif kunci == "tujuan pembelajaran":
            for el in isi:
                if el["jenis"] != "paragraf" or not el["polos"]:
                    continue
                t = re.sub(r"^(\*\*)?[•▪●◦\-–]\s+", r"\1", el["teks"]).replace("**", "")
                kko, _, sisa = t.partition(" ")
                naskah["tujuan"]["butir"].append({"level": "", "kko": kko, "teks": sisa})
        elif kunci == "materi":
            sub, isi_sub = None, []
            for el in isi:
                if el["jenis"] == "h2" and el["polos"]:
                    if sub is not None or isi_sub:
                        tambah_subbab(sub or "Materi", isi_sub, False)
                    sub, isi_sub = el["teks"], []
                else:
                    isi_sub.append(el)
            if sub is not None or isi_sub:
                tambah_subbab(sub or "Materi", isi_sub, False)
        elif kunci.startswith("contoh soal") or kunci.startswith("latihan soal"):
            soal = _soal(penyusun, isi)
            if soal:
                naskah["materi"].append({"subbab": "Contoh Soal", "ikon": "checklist", "slide": soal})
        elif kunci == "penutup":
            naskah["penutup"]["pesan"] = " ".join(
                el["teks"] for el in isi if el["jenis"] == "paragraf" and el["polos"])
        elif kunci in ("sumber", "referensi", "daftar pustaka"):
            continue
        else:
            tambah_subbab(nama, isi, True)

    naskah["penutup"]["ringkasan"] = [
        {"judul": sub["subbab"], "teks": ""} for sub in naskah["materi"] if sub["subbab"] != "Contoh Soal"][:6]
    return naskah


def main(argv=None):
    ap = argparse.ArgumentParser(description="Ubah dokumen materi (.docx dari Google Docs) menjadi naskah JSON.")
    ap.add_argument("docx")
    ap.add_argument("--keluar", required=True, help="Path naskah .json (gambar disimpan di media/ sebelahnya)")
    ap.add_argument("--sumber", default="", help="URL dokumen sumber")
    ap.add_argument("--kelompok", type=int, default=0, help="Indeks kartu aktif di peta subtes TIU")
    ap.add_argument("--nama-file", default=None)
    a = ap.parse_args(argv)
    folder = os.path.dirname(os.path.abspath(a.keluar))
    slug = os.path.splitext(os.path.basename(a.keluar))[0]
    naskah = susun(baca_docx(a.docx), folder, slug, a.sumber, a.kelompok, a.nama_file)
    with open(a.keluar, "w", encoding="utf-8") as f:
        json.dump(naskah, f, ensure_ascii=False, indent=2)
        f.write("\n")
    n_slide = sum(len(s["slide"]) for s in naskah["materi"])
    n_soal = sum(1 for s in naskah["materi"] for x in s["slide"] if x["tipe"] == "contoh_soal")
    print(f"{naskah['meta']['judul']}: {len(naskah['materi'])} subbab, {n_slide} slide isi "
          f"({n_soal} contoh soal) -> {a.keluar}")


if __name__ == "__main__":
    main()
