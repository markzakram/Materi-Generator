"""Ekstraksi PDF soal (format 'Latihan SKD CPNS', dibuat jsPDF) menjadi naskah JSON.

PDF sumber menyimpan rumus sebagai gambar kecil tanpa teks alternatif. Setiap gambar
rumus dikenali dari hash pikselnya lalu diterjemahkan lewat kamus
`generator/label_rumus.json` ({hash: "p → ¬q"}). Pecahan ditulis `[a/b]` di kamus.
Gambar rumus yang belum berlabel bisa dikumpulkan ke lembar kontak PNG.

    python generator/ekstrak_soal.py "TIU Paket 1.pdf" --keluar konten/latihan_tiu_paket1.json
    python generator/ekstrak_soal.py a.pdf b.pdf --lembar folder_lembar   # lembar label rumus baru
"""
from __future__ import annotations

import argparse
import io
import json
import os
import re

import pymupdf

AKAR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LABEL = os.path.join(os.path.dirname(os.path.abspath(__file__)), "label_rumus.json")

TINGGI_RUMUS_MAKS = 22  # pt; gambar lebih tinggi dianggap gambar/figur
TEPI_KIRI, BARIS_PENUH = 50, 500  # pt; untuk menyambung baris menjadi paragraf
POLA_SOAL = re.compile(r"^\*\*Soal (\d+) \| (\w+) \| (.+?)(?: \| (Mudah|Sedang|Sulit))?(?: \| (C\d))?\*\*$")
POLA_BARU = re.compile(r"^(\*\*)?(Soal \d+ \||[A-E]\. |Premis \d|Langkah \d|Jawaban:|Pembahasan:|Kesimpulan|"
                       r"Analisis setiap|Cara cepat|Pernyataan \(|\(\d\) |\d\. |\$→\$|Jadi,|\$[A-Za-z]\$:)")
ANGKA = re.compile(r"[0-9.,()+×÷−=≠≤≥<>%° \-]+")


def muat_label():
    if os.path.exists(LABEL):
        with open(LABEL, encoding="utf-8") as f:
            return json.load(f)
    return {}


def linear(label: str) -> str:
    """Ubah [a/b] menjadi a/b, dengan kurung bila perlu."""
    def bungkus(s):
        return f"({s})" if re.search(r"[ +−×÷]", s) else s

    keluar, pos = [], 0
    for m in re.finditer(r"\[([^\[\]/]+)/([^\[\]]+)\]", label):
        sebelum = label[: m.start()].rstrip()
        sesudah = label[m.end():].lstrip()
        frac = f"{bungkus(m.group(1).strip())}/{bungkus(m.group(2).strip())}"
        if (sebelum.endswith(("×", "÷")) or sesudah.startswith(("×", "÷"))):
            frac = f"({frac})"
        keluar.append(label[pos:m.start()] + frac)
        pos = m.end()
    keluar.append(label[pos:])
    return "".join(keluar)


# ------------------------------------------------------------ item halaman ---
class Item:
    __slots__ = ("x0", "y0", "x1", "y1", "jenis", "isi", "tebal", "xref")

    def __init__(self, bbox, jenis, isi, tebal=False, xref=None):
        self.x0, self.y0, self.x1, self.y1 = bbox
        self.jenis, self.isi, self.tebal, self.xref = jenis, isi, tebal, xref

    @property
    def yc(self):
        return (self.y0 + self.y1) / 2

    @property
    def xc(self):
        return (self.x0 + self.x1) / 2

    @property
    def h(self):
        return self.y1 - self.y0


def item_halaman(page):
    items = []
    for b in page.get_text("dict")["blocks"]:
        for l in b.get("lines", []):
            for s in l["spans"]:
                if s["text"]:
                    items.append(Item(s["bbox"], "teks", s["text"], tebal="Bold" in s["font"]))
    for info in page.get_image_info(hashes=True, xrefs=True):
        bb = info["bbox"]
        jenis = "rumus" if (bb[3] - bb[1]) <= TINGGI_RUMUS_MAKS else "gambar"
        items.append(Item(bb, jenis, info["digest"].hex()[:16], xref=info["xref"]))
    return items


def susun_baris(items):
    """Kelompokkan item menjadi baris berdasarkan tumpang-tindih vertikal."""
    baris = []  # [y0, y1, [item]]
    for it in sorted(items, key=lambda i: (i.yc, i.x0)):
        for b in baris:
            if it.jenis == "gambar" or b[2][0].jenis == "gambar":
                continue
            tumpang = min(b[1], it.y1) - max(b[0], it.y0)
            if tumpang >= 0.5 * min(b[1] - b[0], it.h):
                b[2].append(it)
                if it.jenis == "teks":
                    b[0], b[1] = min(b[0], it.y0), max(b[1], it.y1)
                break
        else:
            baris.append([it.y0, it.y1, [it]])
    hasil = [sorted(b[2], key=lambda i: i.x0) for b in baris]
    hasil.sort(key=lambda b: min(i.y0 for i in b))
    return hasil


# ---------------------------------------------------------------- token ---
def teks_baris(items, label, belum):
    """Gabungkan item satu baris menjadi markup (**tebal**, $rumus$)."""
    keping, prev = [], None  # (jenis, teks, tebal)
    for it in items:
        if prev is not None:
            celah = it.x0 - prev.x1
            if celah > 1.2 and not (keping and keping[-1][1].endswith(" ")) and not (
                    it.jenis == "teks" and it.isi.startswith(" ")):
                keping.append(("spasi", " ", False))
        if it.jenis == "teks":
            for bag in re.split(r"( +)", it.isi):
                if not bag:
                    continue
                if bag.strip() == "":
                    keping.append(("spasi", " ", False))
                else:  # angka_t = angka/tanda yang berasal dari lapisan teks
                    keping.append(("angka_t" if ANGKA.fullmatch(bag) else "teks", bag, it.tebal))
        else:
            lab = label.get(it.isi)
            if lab is None:
                belum.add(it.isi)
                lab = f"⟦{it.isi}⟧"
            lab = linear(lab)
            jenis = "angka" if ANGKA.fullmatch(lab) else "rumus"
            keping.append((jenis, lab, False))
        prev = it

    # gabungkan keping rumus/angka yang bersambung tanpa spasi menjadi satu rumus
    grup, i = [], 0
    while i < len(keping):
        j = i
        while j < len(keping) and keping[j][0] in ("rumus", "angka", "angka_t"):
            j += 1
        if j > i:
            blok = keping[i:j]
            if any(k[0] == "rumus" for k in blok):
                isi = "".join(k[1] for k in blok)
                depan, belakang = "", ""
                m = re.search(r"[.,;:]+$", isi)  # tanda baca akhir kalimat
                if m and blok[-1][0] == "angka_t":
                    isi, belakang = isi[: m.start()], m.group(0)
                while isi.startswith("(") and isi.count("(") > isi.count(")"):
                    isi, depan = isi[1:], depan + "("
                while isi.endswith(")") and isi.count(")") > isi.count("("):
                    isi, belakang = isi[:-1], ")" + belakang
                tebal = any(k[2] for k in blok)
                if depan:
                    grup.append(("angka_t", depan, tebal))
                grup.append(("rumus", isi, False))
                if belakang:
                    grup.append(("angka_t", belakang, tebal))
            else:
                grup.extend(blok)
            i = j
        else:
            grup.append(keping[i])
            i += 1
    # angka di tengah teks tebal ikut tebal (mis. "Langkah 1 (Premis 1 + 2):")
    tebal_terakhir = False
    for n, (jenis, t, tebal) in enumerate(grup):
        if jenis == "teks":
            tebal_terakhir = tebal
        elif jenis in ("angka", "angka_t") and tebal_terakhir:
            berikut = next((k for k in grup[n + 1:] if k[0] == "teks"), None)
            if berikut is not None and berikut[2]:  # hanya bila diapit teks tebal
                grup[n] = (jenis, t, True)

    # baris rumus tampilan: seluruh isinya berasal dari gambar rumus
    hanya_rumus = all(k[0] in ("rumus", "angka", "spasi") for k in grup) and any(
        it.jenis == "rumus" for it in items)
    if hanya_rumus:
        isi = "".join(k[1] for k in grup).strip()
        return (f"${isi}$" if any(k[0] == "rumus" for k in grup) else isi), True

    out = []
    for jenis, t, tebal in grup:
        if jenis == "rumus":
            out.append(f"${t}$")
        elif tebal:
            out.append(f"**{t}**")
        else:
            out.append(t)
    s = "".join(out)
    s = re.sub(r"\*\*(\s*)\*\*", r"\1", s)
    return re.sub(r"[ \t]+", " ", s).strip(), False


# ---------------------------------------------------------------- tabel ---
def _rapatkan(nilai, toleransi=2.0):
    keluar = []
    for v in sorted(nilai):
        if not keluar or v - keluar[-1] > toleransi:
            keluar.append(v)
    return keluar


def grid_kotak(page):
    """Tabel yang digambar sebagai kotak garis per sel -> [(bbox, [y batas], [x batas])]."""
    kotak = [pymupdf.Rect(it[1]) for g in page.get_drawings() if g.get("type") in ("s", "fs")
             for it in g["items"] if it[0] == "re"]
    kotak = [r for r in kotak if r.width > 8 and r.height > 8]
    induk = list(range(len(kotak)))

    def cari(i):
        while induk[i] != i:
            induk[i] = induk[induk[i]]
            i = induk[i]
        return i

    for i, a in enumerate(kotak):
        aa = pymupdf.Rect(a.x0 - 1, a.y0 - 1, a.x1 + 1, a.y1 + 1)
        for j in range(i + 1, len(kotak)):
            if aa.intersects(kotak[j]):
                induk[cari(i)] = cari(j)
    grup = {}
    for i, r in enumerate(kotak):
        grup.setdefault(cari(i), []).append(r)
    hasil = []
    for rs in grup.values():
        if len(rs) < 2:
            continue
        bb = pymupdf.Rect(rs[0])
        for r in rs[1:]:
            bb |= r
        ys = _rapatkan([v for r in rs for v in (r.y0, r.y1)])
        xs = _rapatkan([v for r in rs for v in (r.x0, r.x1)])
        if len(ys) >= 2 and len(xs) >= 3:
            hasil.append((bb, ys, xs))
    return hasil


def _isi_grid(bb, ys, xs, items, label, belum, terpakai):
    data = []
    for y0, y1 in zip(ys, ys[1:]):
        kolom = [[] for _ in xs[1:]]
        for it in items:
            if y0 <= it.yc <= y1 and bb.x0 - 40 <= it.xc <= bb.x1 + 40:
                k = min(len(kolom) - 1, max(0, sum(1 for x in xs[1:-1] if it.xc >= x)))
                kolom[k].append(it)
                terpakai.add(id(it))
        data.append([re.sub(r"\*\*(\s*)\*\*", r"\1", " ".join(teks_baris(b, label, belum)[0]
                                                              for b in susun_baris(kol))) for kol in kolom])
    return data


def ambil_tabel(page, items, label, belum):
    """Deteksi tabel bergaris: grid kotak per sel, lalu find_tables untuk sisanya."""
    hasil, terpakai = [], set()
    area = []
    for bb, ys, xs in grid_kotak(page):
        data = _isi_grid(bb, ys, xs, items, label, belum, terpakai)
        if any(any(c for c in b) for b in data):
            hasil.append((bb.y0, {"jenis": "tabel", "baris": data}))
            area.append(bb)
    try:
        tabs = page.find_tables().tables
    except Exception:
        tabs = []
    for t in tabs:
        bb = pymupdf.Rect(t.bbox)
        if bb.height < 15 or t.col_count < 2 or any(bb.intersects(a) for a in area):
            continue
        baris = [[pymupdf.Rect(c) for c in row.cells if c] for row in t.rows]
        baris = [b for b in baris if b]
        ys = _rapatkan([b[0].y0 for b in baris] + [baris[-1][0].y1])
        xs = _rapatkan([c.x0 for c in baris[0]] + [baris[0][-1].x1])
        data = _isi_grid(bb, ys, xs, items, label, belum, terpakai)
        if data:
            hasil.append((bb.y0, {"jenis": "tabel", "baris": data}))
    return hasil, terpakai


# ------------------------------------------------------------- dokumen ---
def alir_dokumen(path, label, folder_gambar=None, awalan="gbr"):
    """PDF -> daftar entri berurutan: teks (per baris), tabel, gambar."""
    doc = pymupdf.open(path)
    belum = set()
    alir = []
    no_gbr = 0
    for page in doc:
        items = item_halaman(page)
        tabel, terpakai = ambil_tabel(page, items, label, belum)
        entri = list(tabel)
        for b in susun_baris([i for i in items if id(i) not in terpakai]):
            y0 = min(i.y0 for i in b)
            if len(b) == 1 and b[0].jenis == "gambar":
                g = b[0]
                no_gbr += 1
                nama = f"{awalan}_{page.number + 1:02d}_{no_gbr:02d}.png"
                if folder_gambar:
                    os.makedirs(folder_gambar, exist_ok=True)
                    pix = page.get_pixmap(clip=pymupdf.Rect(g.x0, g.y0, g.x1, g.y1), dpi=300)
                    pix.save(os.path.join(folder_gambar, nama))
                entri.append((y0, {"jenis": "gambar", "file": nama, "lebar_pt": round(g.x1 - g.x0, 1),
                                   "tinggi_pt": round(g.y1 - g.y0, 1)}))
                continue
            teks, rumus_saja = teks_baris(b, label, belum)
            if re.fullmatch(r"\d+ / \d+", teks.replace("*", "")):
                continue  # nomor halaman
            entri.append((y0, {"jenis": "teks", "teks": teks, "rumus": rumus_saja,
                               "x0": min(i.x0 for i in b), "x1": max(i.x1 for i in b)}))
        entri.sort(key=lambda e: e[0])
        alir.extend(e for _, e in entri)
    return alir, belum


def jadi_paragraf(alir):
    """Sambung baris yang terpotong lebar halaman menjadi satu paragraf."""
    hasil = []
    for e in alir:
        prev = hasil[-1] if hasil else None
        # tabel yang terpotong pergantian halaman (header diulang) disatukan kembali
        if e["jenis"] == "tabel" and prev and prev["jenis"] == "tabel" and \
                e["baris"] and prev["baris"] and e["baris"][0] == prev["baris"][0]:
            prev["baris"].extend(e["baris"][1:])
            continue
        baru = POLA_BARU.match(e["teks"]) if e["jenis"] == "teks" else None
        if baru and re.match(r"^\d\. ", e["teks"]) and prev and prev["jenis"] == "teks" and \
                not prev["teks"].rstrip("*$ ").endswith((".", ":", "?", "!", ")")):
            baru = None  # "… adalah / 2. Nilai …" = lanjutan kalimat, bukan daftar bernomor
        if (e["jenis"] == "teks" and prev and prev["jenis"] == "teks" and not e["rumus"]
                and not prev["rumus"] and (prev["x1"] > BARIS_PENUH or e["teks"][:1].islower())
                and e["x0"] < TEPI_KIRI and not baru):
            prev = hasil[-1]
            prev["teks"] = re.sub(r"\*\*(\s*)\*\*", r"\1", prev["teks"] + " " + e["teks"])
            prev["x1"] = e["x1"]
            continue
        hasil.append(dict(e))
    for e in hasil:
        e.pop("x0", None)
        e.pop("x1", None)
        if e["jenis"] == "teks" and not e["rumus"]:
            e.pop("rumus")
    return hasil


# ---------------------------------------------------------------- soal ---
POLA_ANALISIS = re.compile(r"^\$→\$\s*(\$[^$]+\$)?\s*—\s*(.*?)\s*\((?:\*\*)?([^()]+?)(?:\*\*)?\)(?:\*\*)?\s*$")


def susun_analisis(blok):
    """Ubah rangkaian 'X. opsi' + '→ rumus — alasan (status)' menjadi blok analisis."""
    keluar, i = [], 0
    while i < len(blok):
        b = blok[i]
        if b["jenis"] == "teks" and b["teks"].strip("*").startswith("Analisis setiap opsi"):
            catatan, item, j = [], [], i + 1
            while j < len(blok) and blok[j]["jenis"] == "teks":
                t = blok[j]["teks"]
                m_opsi = re.match(r"^([A-E])\. (.*)$", t)
                if m_opsi and j + 1 < len(blok) and blok[j + 1]["jenis"] == "teks":
                    m = POLA_ANALISIS.match(blok[j + 1]["teks"])
                    if m:
                        status = m.group(3).strip()
                        item.append({"opsi": m_opsi.group(1), "teks": m_opsi.group(2),
                                     "rumus": (m.group(1) or "").strip("$"), "alasan": m.group(2),
                                     "status": status, "benar": not status.startswith("bukan")})
                        j += 2
                        continue
                if item:
                    break
                catatan.append(t)
                j += 1
            if item:
                keluar.append({"jenis": "analisis", "catatan": catatan, "item": item})
                i = j
                continue
        keluar.append(b)
        i += 1
    return keluar


def urai_soal(blok_semua):
    meta, soal, kini, tahap = [], [], None, None
    for b in blok_semua:
        t = b.get("teks", "")
        m = POLA_SOAL.match(t) if b["jenis"] == "teks" else None
        if m:
            kini = {"nomor": int(m.group(1)), "subtes": m.group(2), "subtipe": m.group(3),
                    "level": m.group(4) or "", "bloom": m.group(5) or "",
                    "soal": [], "opsi": {}, "kunci": "", "pembahasan": []}
            soal.append(kini)
            tahap = "soal"
            continue
        if kini is None:
            meta.append(t.replace("**", ""))
            continue
        mj = re.fullmatch(r"\*\*Jawaban: ([A-E])\*\*", t) if b["jenis"] == "teks" else None
        if tahap in ("soal", "opsi") and mj:
            kini["kunci"] = mj.group(1)
            tahap = "jawab"
            continue
        if tahap == "jawab" and t == "**Pembahasan:**":
            tahap = "pembahasan"
            continue
        mo = re.match(r"^([A-E])\. (.*)$", t) if b["jenis"] == "teks" else None
        if tahap in ("soal", "opsi") and mo:
            kini["opsi"][mo.group(1)] = mo.group(2)
            tahap = "opsi"
            continue
        if tahap in ("soal", "opsi"):
            kini["soal"].append(b)
        elif tahap == "pembahasan":
            kini["pembahasan"].append(b)
    for s in soal:
        s["pembahasan"] = susun_analisis(s["pembahasan"])
    return meta, soal


def ekstrak(path, label, folder_gambar, awalan):
    alir, belum = alir_dokumen(path, label, folder_gambar, awalan)
    meta, soal = urai_soal(jadi_paragraf(alir))
    return meta, soal, belum


def lembar_kontak(paths, label, folder, per_lembar=40):
    """Buat lembar PNG berisi gambar rumus yang belum berlabel, bernomor urut."""
    from PIL import Image, ImageDraw, ImageFont

    unik = {}
    for path in paths:
        doc = pymupdf.open(path)
        for page in doc:
            for info in page.get_image_info(hashes=True, xrefs=True):
                bb = info["bbox"]
                kunci = info["digest"].hex()[:16]
                if bb[3] - bb[1] > TINGGI_RUMUS_MAKS or kunci in label or kunci in unik:
                    continue
                unik[kunci] = doc.extract_image(info["xref"])["image"]
    os.makedirs(folder, exist_ok=True)
    urut, indeks = list(unik), {}
    font = ImageFont.truetype("arial.ttf", 22)
    for n in range(0, len(urut), per_lembar):
        bagian = urut[n:n + per_lembar]
        lembar = Image.new("RGB", (1500, 70 * ((len(bagian) + 1) // 2) + 10), "white")
        dr = ImageDraw.Draw(lembar)
        for j, k in enumerate(bagian):
            indeks[n + j + 1] = k
            im = Image.open(io.BytesIO(unik[k])).convert("RGB")
            skala = min(1.6, 600 / im.width, 58 / im.height)
            im = im.resize((max(1, int(im.width * skala)), max(1, int(im.height * skala))))
            x, y = 10 + (j % 2) * 750, 5 + (j // 2) * 70
            dr.text((x, y + 18), f"{n + j + 1:3d}", fill="red", font=font)
            lembar.paste(im, (x + 60, y + 5))
        lembar.save(os.path.join(folder, f"lembar_{n // per_lembar + 1:02d}.png"))
    with open(os.path.join(folder, "indeks.json"), "w", encoding="utf-8") as f:
        json.dump(indeks, f, indent=1)
    return len(urut)


def main(argv=None):
    ap = argparse.ArgumentParser(description="Ekstrak PDF soal menjadi naskah JSON latihan.")
    ap.add_argument("pdf", nargs="+")
    ap.add_argument("--keluar", help="file JSON naskah (satu PDF)")
    ap.add_argument("--judul", default="Latihan Soal TIU")
    ap.add_argument("--paket", default="")
    ap.add_argument("--lembar", help="folder lembar kontak rumus yang belum berlabel")
    a = ap.parse_args(argv)
    label = muat_label()
    if a.lembar:
        print(f"{lembar_kontak(a.pdf, label, a.lembar)} rumus belum berlabel -> {a.lembar}")
        return
    path = a.pdf[0]
    keluar = a.keluar or os.path.join(AKAR, "konten", os.path.splitext(os.path.basename(path))[0] + ".json")
    nama = os.path.splitext(os.path.basename(keluar))[0]
    folder_gambar = os.path.join(os.path.dirname(keluar), "gambar", nama)
    meta, soal, belum = ekstrak(path, label, folder_gambar, "g")
    if belum:
        print(f"PERINGATAN: {len(belum)} rumus belum berlabel; jalankan dengan --lembar.")
    info = next((m for m in meta if "soal" in m), "")
    naskah = {
        "meta": {"judul": a.judul, "paket": a.paket, "subtes": "Tes Intelegensia Umum", "info": info,
                 "nama_file": f"{a.judul} - {a.paket}".strip(" -"), "folder_gambar": f"gambar/{nama}",
                 "sumber": os.path.basename(path)},
        "soal": soal,
    }
    with open(keluar, "w", encoding="utf-8") as f:
        json.dump(naskah, f, ensure_ascii=False, indent=1)
    kosong = [s["nomor"] for s in soal if not s["kunci"]]
    print(f"{len(soal)} soal -> {keluar}" + (f" | tanpa kunci: {kosong}" if kosong else ""))


if __name__ == "__main__":
    main()
