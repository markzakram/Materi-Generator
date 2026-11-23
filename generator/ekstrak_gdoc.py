"""Ubah dokumen materi Google Docs (ekspor .docx) menjadi naskah JSON generator.

Dokumen mengikuti templat booklet materi SKD: Judul, Pendahuluan, Tujuan Pembelajaran,
Materi (Heading 2 = subbab, Heading 3 = slide), Contoh Soal dan Pembahasan, Penutup.
Sebagian dokumen memakai Heading 1 langsung sebagai subbab; keduanya didukung.

Isi tidak dikarang ulang: paragraf, tabel, gambar, dan persamaan diambil apa adanya. Yang masuk naskah hanya
bagian yang ditandai penulis: sorotan kuning (highlight), atau bagian berlabel [WAJIB PPT] (lihat tandai()).
Bagian yang tidak ada di dokumen (pertanyaan pemantik, level Bloom) dibiarkan kosong
agar diisi tim akademik lewat editor.

    python generator/ekstrak_gdoc.py "01 Susunan & Posisi.docx" --keluar konten/tiu_vn01.json \
        --sumber https://docs.google.com/document/d/... --kelompok 2
"""
from __future__ import annotations

import argparse
import copy
import io
import json
import os
import re
import sys
import zipfile

from lxml import etree

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import omml  # noqa: E402

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


def _disorot(rPr):
    """Run (teks, rumus, gambar) atau tanda paragraf yang diberi sorotan warna."""
    h = rPr.find(f"{W}highlight") if rPr is not None else None
    return h is not None and h.get(f"{W}val") not in (None, "none", "white")


def _kalimat_disorot(potongan):
    """Potongan [(teks, jenis, sorot)] dari kalimat yang sebagian besar hurufnya disorot. Kalimat utuh dipakai
    supaya sorotan yang mulai di tengah kalimat tetap terbaca."""
    polos = "".join(x[0] for x in potongan)
    batas, awal = [], 0
    for m in re.finditer(r"[.!?][”\"')]?\s+(?=\S)", polos):
        batas.append((awal, m.end()))
        awal = m.end()
    batas.append((awal, len(polos)))
    pakai = [False] * len(polos)
    for a, b in batas:
        huruf = [i for i in range(a, b) if not polos[i].isspace()]
        i = 0
        sorot = 0
        pos = 0
        for teks, _, s in potongan:
            if s:
                sorot += sum(1 for k in range(max(a, pos), min(b, pos + len(teks))) if not polos[k].isspace())
            pos += len(teks)
        if huruf and sorot >= 0.5 * len(huruf):
            for i in range(a, b):
                pakai[i] = True
    hasil, pos = [], 0
    for teks, jenis, _ in potongan:
        if jenis == "m":  # rumus tidak dipotong
            isi = teks if any(pakai[pos:pos + len(teks)]) else ""
        else:
            isi = "".join(c for k, c in enumerate(teks, pos) if pakai[k])
        if isi:
            hasil.append((isi, jenis))
        pos += len(teks)
    return hasil


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
    """Teks paragraf dengan markup, daftar gambar (rId, lebar, tinggi, disorot), status tebal/miring penuh, dan
    teks yang hanya memuat kalimat disorot."""
    potongan, gambar = [], []
    semua_tebal, semua_miring, ada_teks = True, True, False
    par_disorot = _disorot(p.find(f"{W}pPr/{W}rPr"))

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
                                       int(ext.get("cy")) / EMU if ext is not None else None,
                                       _disorot(anak.find(f"{W}rPr")) or par_disorot))
                t = ""
                for x in anak:
                    if x.tag == f"{W}t":
                        t += x.text or ""
                    elif x.tag in (f"{W}tab", f"{W}br"):
                        t += " "
                posisi = anak.find(f"{W}rPr/{W}vertAlign")
                naik = {"superscript": "^", "subscript": "_"}.get(posisi.get(f"{W}val") if posisi is not None else None)
                s = _disorot(anak.find(f"{W}rPr"))
                if t.strip() and naik:  # superskrip/subskrip teks biasa: cm^{2}, H_{2}O
                    ada_teks = True
                    potongan.append((f"{naik}{{{t.strip()}}}", "m", s))
                elif t:
                    if t.strip():
                        ada_teks = True
                        semua_tebal &= _tebal(anak)
                        semua_miring &= _miring(anak)
                    potongan.append((t, "b" if _tebal(anak) else "", s))
            elif tag in (f"{M}oMath", f"{M}oMathPara"):
                t = omml.teks(anak)  # pecahan, pangkat, akar dibaca strukturnya (bukan digabung teksnya)
                if t.strip():
                    ada_teks = True
                    semua_tebal = False
                    potongan.append((" " + t + " ", "m", any(_disorot(r.find(f"{W}rPr")) for r in anak.iter(f"{M}r"))))
            elif tag in (f"{W}hyperlink", f"{W}smartTag", f"{W}ins", f"{W}sdt", f"{W}sdtContent"):
                jalan(anak)

    jalan(p)
    return (_gabung([x[:2] for x in potongan]), gambar, (semua_tebal and ada_teks), (semua_miring and ada_teks),
            _gabung(_kalimat_disorot(potongan)))


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
            teks, gambar, tebal, miring, sorot = _teks_paragraf(el)
            jenis = {"title": "judul", "subtitle": "subjudul", "heading 1": "h1", "heading 2": "h2",
                     "heading 3": "h3", "heading 4": "h3"}.get(st, "paragraf")
            for rid, w, h, s in gambar:
                berkas = target.get(rid, "")
                elemen.append({"jenis": "gambar", "data": z.read("word/" + berkas) if berkas else b"",
                               "ext": os.path.splitext(berkas)[1].lower(), "w": w, "h": h, "sorot": s})
            polos = teks.replace("**", "").strip()
            if polos or jenis != "paragraf":
                elemen.append({"jenis": jenis, "teks": teks, "polos": polos, "tebal": tebal,
                               "miring": miring, "daftar": daftar, "sorot": sorot})
        elif el.tag == f"{W}tbl":
            baris, par_baris, gambar_baris, sorot_baris = [], [], [], []
            for tr in el.findall(f"{W}tr"):
                sel, par_sel, gambar_sel, sorot_sel = [], [], [], []
                for tc in tr.findall(f"{W}tc"):
                    tcPr = tc.find(f"{W}tcPr")
                    rentang = tcPr.find(f"{W}gridSpan") if tcPr is not None else None
                    lanjut = tcPr.find(f"{W}vMerge") if tcPr is not None else None
                    hasil = [_teks_paragraf(p) for p in tc.findall(f".//{W}p")]
                    isi = [h[0] for h in hasil]
                    sambungan = lanjut is not None and lanjut.get(f"{W}val") != "restart"
                    teks = "" if sambungan else " ".join(x for x in isi if x)
                    sel.append(teks)
                    if teks:
                        sorot_sel.append(bool(" ".join(h[4] for h in hasil).strip()))
                    par_sel.append([] if sambungan else [x for x in isi if x])
                    # gambar di dalam sel (mis. foto di tabel profil)
                    gambar_sel.append([] if sambungan else [
                        {"jenis": "gambar", "data": z.read("word/" + target[rid]), "ext": os.path.splitext(target[rid])[1].lower(),
                         "w": w, "h": h} for x in hasil for rid, w, h, _ in x[1] if target.get(rid)])
                    # sel gabungan mendatar diperluas agar kolom tetap sejajar
                    tambah = int(rentang.get(f"{W}val")) - 1 if rentang is not None else 0
                    sel.extend([""] * tambah)
                    par_sel.extend([[]] * tambah)
                    gambar_sel.extend([[]] * tambah)
                baris.append(sel)
                par_baris.append(par_sel)
                gambar_baris.append(gambar_sel)
                sorot_baris.append(sum(sorot_sel) >= 0.5 * len(sorot_sel) if sorot_sel else False)
            if baris:
                elemen.append({"jenis": "tabel", "baris": baris, "par": par_baris, "gambar": gambar_baris,
                               "sorot_baris": sorot_baris})
    return elemen


# ------------------------------------------------------------- susun ---
# catatan penulis untuk pembuat PPT, mis. [WAJIB PPT], [PENGAYAAN], [YANG HARUS ADA DI PPT HANYA FOTO ...]
CATATAN_PENULIS = re.compile(r"\s*\[[^\]]*\b(?:WAJIB|PENGAYAAN|PPT|SEMUA|Sertakan|slide)\b[^\]]*\]", re.I)
_KATA_KECIL = {"dan", "atau", "di", "ke", "dari", "yang", "untuk", "pada", "dengan", "dalam", "vs"}


def _rapikan_kapital(t):
    """Judul huruf besar semua ('3.1 PERUSAHAAN MODAL VENTURA') -> huruf judul; singkatan pendek tetap."""
    huruf = re.sub(r"[^A-Za-z]", "", t)
    if len(huruf) < 6 or not huruf.isupper():
        return t
    return " ".join(w.lower() if w.lower() in _KATA_KECIL else
                    (w.capitalize() if len(re.sub(r"[^A-Za-z]", "", w)) > 3 else w) for w in t.split())


def _bersih(judul):
    """Judul tanpa markup, catatan penulis, dan nomor ('3.1', '2.', 'B.')."""
    t = CATATAN_PENULIS.sub("", judul.replace("**", "")).strip()
    t = re.sub(r"^(?:\d+(?:\.\d+)+\.?|\d+[.)]?|[A-H][.)])\s+", "", t).strip()
    return _rapikan_kapital(t)


def _kunci_bagian(nama):
    """Nama bagian Heading 1 untuk dicocokkan: '1. PENDAHULUAN', 'A. Pendahuluan [WAJIB PPT]' -> 'pendahuluan'."""
    return _bersih(nama).lower().strip(" :")


POLA_BAGIAN = re.compile(r"^[A-H]\.\s+(pendahuluan|tujuan pembelajaran|materi pembelajaran|strategi|contoh soal|"
                         r"pola soal|kesalahan umum|penutup)\b", re.I)


def _rapikan_elemen(elemen):
    """Perbaiki gaya yang keliru di dokumen sebelum disusun.

    - judul bagian yang ditulis sebagai paragraf biasa ('E. Contoh Soal [...]') -> Heading 1
    - Heading 2/3 yang sebenarnya kalimat panjang -> paragraf
    - catatan penulis ([WAJIB PPT], [PENGAYAAN], [Sertakan gambar pada slide]) dibuang
    """
    hasil = []
    for el in elemen:
        if "polos" not in el:
            hasil.append(el)
            continue
        el = dict(el, **{k: CATATAN_PENULIS.sub("", el[k]).strip() for k in ("teks", "polos", "sorot")})
        j, polos = el["jenis"], el["polos"]
        if j == "paragraf" and not polos:
            continue
        if j == "paragraf" and len(polos) <= 70 and POLA_BAGIAN.match(polos):
            el["jenis"] = "h1"
        elif j in ("h2", "h3") and (len(polos) > 100 or (polos.endswith(".") and len(polos.split()) > 12)):
            el["jenis"] = "paragraf"
        hasil.append(el)
    return hasil


# ------------------------------------------------------- bagian yang masuk PPT ---
LABEL_WAJIB = re.compile(r"\[[^\]]*\b(?:wajib|harus ada)\b[^\]]*\]", re.I)  # [WAJIB PPT], [semua contoh soal wajib ada di PPT]
LABEL_SOROT_SAJA = re.compile(r"hig\w*light|hilight|disorot|sorotan", re.I)  # [contoh soal yang dihighlight wajib ada ...]
LABEL_PROFIL = re.compile(r"\[[^\]]*foto dan profil[^\]]*\]", re.I)
LABEL_GAMBAR = re.compile(r"\[[^\]]*sertakan gambar[^\]]*\]", re.I)
LABEL_PENGAYAAN = re.compile(r"\[[^\]]*pengayaan[^\]]*\]", re.I)
TINGKAT = {"h1": 1, "h2": 2, "h3": 3}


def ada_penanda(elemen):
    """Dokumen memakai sorotan atau label [WAJIB PPT]? Bila tidak, seluruh isinya dipakai."""
    return any(el.get("sorot") or any(el.get("sorot_baris", [])) or LABEL_WAJIB.search(el.get("teks", "")) for el in elemen)


def tandai(elemen):
    """Beri el['masuk'] (boleh masuk PPT) dan el['utuh'] (pakai teks lengkap, bukan kalimat disorot saja).

    Aturan tim akademik: yang masuk hanya bagian bersorotan. Judul berlabel [WAJIB PPT] berarti seluruh isi
    bagian itu masuk; [PENGAYAAN] mengembalikannya ke sorotan saja. Label di paragraf ('Gambar 2 [WAJIB PPT]. ...')
    berlaku untuk paragraf itu dan gambar di dekatnya.
    """
    if not ada_penanda(elemen):
        return [dict(el, masuk=True, utuh=True) for el in elemen]
    hasil, aturan, gambar_wajib = [], [], False
    for el in elemen:
        j, teks = el["jenis"], el.get("teks", "")
        tingkat = TINGKAT.get(j) or (1 if j == "paragraf" and len(teks) <= 110 and POLA_BAGIAN.match(el["polos"]) else None)
        wajib_label = bool(LABEL_WAJIB.search(teks)) and not LABEL_SOROT_SAJA.search(teks)
        if tingkat and el.get("polos"):
            while aturan and aturan[-1][0] >= tingkat:
                aturan.pop()
            if LABEL_PROFIL.search(teks):
                aturan.append((tingkat, "profil"))
            elif wajib_label:
                aturan.append((tingkat, "wajib"))
            elif LABEL_PENGAYAAN.search(teks):
                aturan.append((tingkat, "sorot"))
        kini = aturan[-1][1] if aturan else "sorot"
        if j == "gambar":
            el = dict(el, masuk=bool(el["sorot"]) or kini == "wajib" or gambar_wajib, utuh=True)
        elif j == "tabel":
            foto = any(g for r in el["gambar"] for g in r)
            utuh = kini == "wajib" or (kini == "profil" and foto)
            el = dict(el, masuk=utuh or any(el["sorot_baris"]), utuh=utuh)
        else:
            if wajib_label and not tingkat and not CATATAN_PENULIS.sub("", el["polos"]).strip() and                     hasil and hasil[-1]["jenis"] in TINGKAT:  # label di baris sendiri, tepat di bawah judul
                aturan.append((TINGKAT[hasil[-1]["jenis"]], "wajib"))
            utuh = kini == "wajib" or (wajib_label and not tingkat)
            el = dict(el, masuk=utuh or bool(el["sorot"].strip()), utuh=utuh)
            if el["polos"]:
                gambar_wajib = bool(LABEL_GAMBAR.search(teks)) or (wajib_label and not tingkat)
            # keterangan gambar yang masuk (berlabel wajib atau disorot): gambar tepat di atasnya ikut
            if el["masuk"] and not tingkat and hasil and hasil[-1]["jenis"] == "gambar" and                     (wajib_label or re.match(r"^(Gambar|Diagram|Ilustrasi)\s*\d*[.:]", el["polos"])):
                hasil[-1]["masuk"] = True
        hasil.append(el)
    for i, el in enumerate(hasil):  # gambar di antara dua paragraf yang masuk (di subbab yang sama) ikut
        if el["jenis"] == "gambar" and not el["masuk"]:
            el["masuk"] = _tetangga_masuk(hasil[i - 1::-1] if i else []) and _tetangga_masuk(hasil[i + 1:])
    return hasil


def _tetangga_masuk(elemen):
    """Paragraf pertama di elemen masuk PPT? Berhenti di judul subbab."""
    for el in elemen:
        if el["jenis"] in ("h1", "h2"):
            return False
        if el.get("polos"):
            return el["masuk"]
    return False


def _seperti_judul(el):
    """Tingkat judul elemen: Heading, atau paragraf tebal pendek yang menjadi subjudul di slide (4)."""
    if el["jenis"] in TINGKAT:
        return TINGKAT[el["jenis"]]
    if el["jenis"] != "paragraf" or len(el["polos"]) > 90 or el["daftar"]:
        return None
    polos = el["polos"]
    if re.fullmatch(r"cara cepat:?", polos.lower()) or re.match(r"^(Langkah\s+\d+\b|Kesimpulan$|Pemisalan$)", polos):
        return 4
    simbol = any(c in SIMBOL_SOROT for c in polos.rstrip(":")) or re.search(r"\w\s+[:=]\s+\w", polos)
    return 4 if el["tebal"] and not simbol else None


def saring(elemen):
    """Elemen satu bagian (bukan soal) yang masuk PPT. Judul ikut bila ada isi di bawahnya yang masuk."""
    ikut = [False] * len(elemen)
    for i, el in enumerate(elemen):
        tk = _seperti_judul(el)
        if tk is None:
            ikut[i] = el["masuk"]
            continue
        for lain in elemen[i + 1:]:
            tl = _seperti_judul(lain)
            if tl is not None and tl <= tk:
                break
            if lain["masuk"] and (tl is None or lain.get("sorot") and tl == 4):
                ikut[i] = True
                break
        if tk >= 3 and el["masuk"]:
            ikut[i] = True  # subjudul yang disorot tetap masuk walau isinya tidak
    hasil = []
    for el, ya in zip(elemen, ikut):
        if not ya:
            continue
        if el["jenis"] == "tabel" and not el["utuh"]:
            pakai = [0] + [r for r in range(1, len(el["baris"])) if el["sorot_baris"][r]]  # header selalu ikut
            el = dict(el, **{k: [el[k][r] for r in pakai] for k in ("baris", "par", "gambar")})
        elif el["jenis"] == "paragraf" and not el["utuh"] and el["sorot"].strip():
            el = dict(el, teks=el["sorot"], polos=el["sorot"].replace("**", "").strip())
        hasil.append(el)
    return hasil


def _kalimat_pertama(teks):
    t = teks.replace("**", "")
    m = re.match(r"(.+?[.!?])(\s|$)", t)
    return (m.group(1) if m else t).strip()


class Penyusun:
    def __init__(self, folder_naskah, slug):
        self.folder_naskah, self.slug = folder_naskah, slug
        self.n_gambar = 0

    def simpan_gambar(self, el, foto_persegi=False):
        """Simpan gambar ke media/<slug>/. Gambar lebih lebar dari 1600 px diperkecil; foto profil
        dipotong persegi (bagian atas, tempat wajah) dan disimpan sebagai JPEG 480 px."""
        if el["ext"] not in (".png", ".jpg", ".jpeg", ".gif", ".bmp"):
            return None
        data, ext = el["data"], el["ext"]
        try:
            from PIL import Image
            with Image.open(io.BytesIO(data)) as im:
                im.load()
                if foto_persegi:
                    s = min(im.size)
                    im = im.crop(((im.width - s) // 2, 0, (im.width - s) // 2 + s, s)).resize((480, 480), Image.LANCZOS)
                    latar = Image.new("RGB", im.size, "white")
                    latar.paste(im.convert("RGBA"), mask=im.convert("RGBA").getchannel("A"))
                    keluar = io.BytesIO()
                    latar.save(keluar, "JPEG", quality=85, optimize=True)
                    data, ext = keluar.getvalue(), ".jpg"
                elif im.width > 1600:
                    im = im.resize((1600, round(im.height * 1600 / im.width)), Image.LANCZOS)
                    keluar = io.BytesIO()
                    im.save(keluar, "PNG" if ext == ".png" else "JPEG", optimize=True, **({} if ext == ".png" else {"quality": 88}))
                    if len(keluar.getvalue()) < len(data):  # pakai yang lebih kecil
                        data = keluar.getvalue()
        except Exception:  # gambar yang tidak terbaca Pillow disimpan apa adanya
            pass
        self.n_gambar += 1
        rel = f"media/{self.slug}/gambar-{self.n_gambar:02d}{ext}"
        path = os.path.join(self.folder_naskah, rel)
        os.makedirs(os.path.dirname(path), exist_ok=True)
        with open(path, "wb") as f:
            f.write(data)
        return rel

    def blok_profil(self, el):
        """Tabel dengan kolom foto -> blok 'profil' (foto, nama, jabatan), atau None.

        Seperti instruksi di dokumen OJK ('yang harus ada di PPT hanya foto dan profil'), kartu memuat
        foto dan kolom teks terpendek (nama + jabatan); kolom uraian panjang tidak ikut.
        """
        gambar = el.get("gambar") or []
        baris = el["baris"]
        if len(baris) < 3 or not any(g for r in gambar for g in r):
            return None
        data = list(range(1, len(baris)))  # baris 0 = header
        n_kol = max(len(r) for r in baris)
        kol_foto = next((c for c in range(n_kol)
                         if sum(1 for r in data if c < len(gambar[r]) and gambar[r][c]) >= max(2, 0.6 * len(data))), None)
        if kol_foto is None:
            return None
        nomor = re.compile(r"^\s*\d+[.)]?\s*$")
        kandidat = [c for c in range(n_kol) if c != kol_foto and any(
            c < len(baris[r]) and baris[r][c].strip() and not nomor.match(baris[r][c]) for r in data)]
        if not kandidat:
            return None
        rata = lambda c: sum(len(baris[r][c]) for r in data if c < len(baris[r])) / len(data)  # noqa: E731
        kol_teks = min(kandidat, key=rata)
        item = []
        for r in data:
            foto = gambar[r][kol_foto][0] if gambar[r][kol_foto] else None
            par = el["par"][r][kol_teks] if kol_teks < len(el["par"][r]) else []
            if not par and not foto:
                continue
            nama = (par[0] if par else "").replace("**", "")
            peran = " ".join(par[1:]).replace("**", "")
            if not peran or len(nama) > 40:  # nama dan jabatan dalam satu paragraf: pisah di awal jabatan
                m = re.search(r"\s(?=(?:Ketua|Wakil|Kepala|Anggota|Deputi|Direktur|Komisaris|Sekretaris)\b)", nama)
                if m and m.start() >= 4:
                    nama, peran = nama[:m.start()].strip(), (nama[m.end():].strip() + " " + peran).strip()
            item.append({"file": self.simpan_gambar(foto, foto_persegi=True) if foto else None,
                         "nama": nama, "peran": peran})
        return {"jenis": "profil", "item": item} if item else None

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
                profil = self.blok_profil(el)
                baris = [r for r in el["baris"] if any(c.strip() for c in r)]
                if profil:
                    blok.append(profil)
                elif len(baris) >= 2 and len(baris[0]) >= 1:
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
                m_cepat = re.match(r"^cara cepat\s*[:.]\s*(.+)$", polos, re.I)
                if re.fullmatch(r"cara cepat:?", polos.lower()) and i + 1 < len(elemen) and elemen[i + 1]["jenis"] == "paragraf":
                    blok.append({"jenis": "catatan", "catatan": "tips", "judul": "Cara cepat",
                                 "teks": elemen[i + 1]["teks"]})
                    i += 1
                elif m_cepat:  # 'Cara cepat. Ikuti siapa ...' dalam satu paragraf
                    blok.append({"jenis": "catatan", "catatan": "tips", "judul": "Cara cepat", "teks": m_cepat.group(1)})
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


POLA_JUDUL_SOAL = re.compile(r"^(contoh\s+)?soal\s*\d+", re.I)
# 'Jawaban: B', 'Jawaban C. Penyelenggara', 'Jawaban pada PPT: B'; bukan 'Jawaban: E5.' atau 'Jawaban: Selasa'
POLA_KUNCI = re.compile(r"^(?i:jawaban)\b(?:[^:]{0,20}:)?\s*\(?([A-E])\)?(?=$|[.),;\s])")


POLA_KUNCI_AKHIR = re.compile(r"\b[Jj]awaban(?:nya)?(?:\s+(?:yang\s+)?(?:benar|tepat))?(?:\s+adalah)?\s*:?\s*\(?([A-E])\)?\s*\.?$")


def _soal(penyusun, elemen):
    """Bagian 'Contoh Soal dan Pembahasan' -> daftar slide contoh_soal.

    Soal dimulai di heading 'Soal 1'/'Contoh Soal 1'. Bila tidak ada heading seperti itu, setiap
    Heading 2 memulai soal baru ('4.1 Mengenali platform', 'Contoh 2 ...').
    """
    hasil, label, kini, mode = [], None, None, None
    pengantar = []  # kalimat sebelum soal pertama
    per_judul = any(e["jenis"] in ("h2", "h3") and POLA_JUDUL_SOAL.match(_bersih(e.get("polos", ""))) for e in elemen)
    for el in elemen:
        j, polos = el["jenis"], el.get("polos", "")
        mulai = (POLA_JUDUL_SOAL.match(_bersih(polos)) if per_judul else j == "h2") and j in ("h2", "h3")
        if mulai and polos:
            kini = {"tipe": "contoh_soal", "judul": _bersih(el["teks"]), "soal": [], "opsi": {}, "kunci": "",
                    "pembahasan": [], "_tanda": []}
            if label:
                kini["label"] = label
            hasil.append(kini)
            mode = "soal"
            continue
        if j in ("h2", "h3") and not polos:
            continue
        if j in ("h2", "h3") and kini is not None and re.fullmatch(r"pembahasan\s*:?", polos, re.I):
            mode = "bahas"
            continue
        if j in ("h2", "h3") and (kini is None or per_judul and j == "h2"):
            label = _bersih(el["teks"])
            continue
        if kini is None:
            pengantar.append(el)
            continue
        m_jawab = POLA_KUNCI.match(polos)
        if m_jawab and (mode in ("soal", "opsi") or not kini["kunci"]):
            kini["kunci"] = m_jawab.group(1).upper()
            mode = "kunci"
            # 'Jawaban B. Polis B memenuhi ...': alasan di baris kunci (bukan ulangan teks opsi) ikut ke pembahasan
            alasan = polos[m_jawab.end():].lstrip(" .),:;–-").strip()
            rata = lambda x: re.sub(r"[\W_]+", "", x.lower())  # noqa: E731
            if len(alasan.split()) >= 4:
                a, o = rata(alasan), rata(kini["opsi"].get(kini["kunci"], ""))
                if not (a.startswith(o[:40]) or o.startswith(a[:40])):
                    kini["pembahasan"].append(dict(el, teks=alasan, polos=alasan, sorot=alasan if el["masuk"] else "", tebal=False))
            continue
        if mode in ("soal", "opsi") and re.match(r"^jawaban\b", polos, re.I):
            mode = "bahas"  # jawaban berupa uraian (bukan huruf opsi): bagian dari pembahasan
        m_bahas = re.match(r"^pembahasan\s*[:.]?\s*(.*)$", polos, re.I)
        if m_bahas and j == "paragraf":
            mode = "bahas"
            if m_bahas.group(1):  # 'Pembahasan: Penyelenggara mengelola ...'
                kini["pembahasan"].append(dict(el, teks=m_bahas.group(1), polos=m_bahas.group(1), tebal=False, sorot=re.sub(
                    r"^(\*\*)?pembahasan\s*[:.]?\s*(\*\*)?\s*", "", el["sorot"], flags=re.I)))
            continue
        m_opsi = re.match(r"^\(?([A-E])[.)]\s+(.*)$", polos) if j == "paragraf" else None
        if mode in ("soal", "opsi") and m_opsi and (m_opsi.group(1) == "A" or mode == "opsi"):
            teks = re.sub(r"^(\*\*)?\(?[A-E][.)]\s+", r"\1", el["teks"])
            kini["opsi"][m_opsi.group(1)] = teks
            kini["_tanda"].append((len(polos), el["masuk"]))
            mode = "opsi"
            continue
        if mode not in ("bahas", "kunci"):
            kini["_tanda"].append((len(polos) or 20, el["masuk"]))
        kini["pembahasan" if mode in ("bahas", "kunci") else "soal"].append(el)
    # soal masuk (pertanyaan, opsi, dan kunci utuh) bila sebagian besar pertanyaan dan opsinya disorot;
    # pembahasannya hanya bagian yang disorot, kecuali jawaban uraian
    hasil = [s for s in hasil if sum(n for n, ya in s["_tanda"] if ya) >= 0.5 * sum(n for n, _ in s["_tanda"])]
    for s in hasil:
        del s["_tanda"]
        if s["opsi"] and not s["kunci"]:  # kunci hanya ditulis di penutup pembahasan: 'Jadi, ... Jawaban A.'
            akhir = [e.get("polos", "").strip() for e in s["pembahasan"] if e["jenis"] == "paragraf"][-2:]
            kunci = [m.group(1) for m in map(POLA_KUNCI_AKHIR.search, akhir) if m and m.group(1) in s["opsi"]]
            s["kunci"] = kunci[-1] if kunci else ""
        s["soal"] = penyusun.blok_dari(s["soal"])
        s["pembahasan"] = penyusun.blok_dari(saring([
            dict(e, masuk=True, utuh=True) if not s["opsi"] and re.match(r"^jawaban\b", e.get("polos", ""), re.I) else e
            for e in s["pembahasan"]]))
    awal = penyusun.blok_dari(saring(pengantar)) if hasil else []
    if awal:
        hasil.insert(0, {"tipe": "blok", "judul": "Contoh Soal", "blok": awal})
    return hasil


LEWATI_BAGIAN = ("sumber", "referensi", "daftar pustaka", "daftar rujukan", "kata pengantar", "capaian pembelajaran",
                 "pemetaan referensi")


def susun(elemen, folder_naskah, slug, sumber="", kelompok=0, nama_file=None, subtes=SUBTES_TIU, peta=None,
          judul=None):
    """peta: peta subtes (bawaan PETA_TIU); judul: judul materi bila dokumen tidak memakai gaya Title."""
    penyusun = Penyusun(folder_naskah, slug)
    judul_dok, subjudul, bagian = _bagi_bagian(_rapikan_elemen(tandai(elemen)))
    judul_bersih = judul or re.sub(r"^\d{2}\s+", "", judul_dok).strip() or slug  # buang nomor urut "01 "
    naskah = {
        "meta": {"judul": judul_bersih, "subtes": subtes,
                 "nama_file": nama_file or f"TIU - {judul_bersih}", "sumber": sumber},
        "peta_subtes": dict(peta or PETA_TIU, aktif=kelompok),
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
        kunci = _kunci_bagian(nama)
        if kunci in ("tujuan", "tujuan pembelajaran"):
            kunci = "tujuan pembelajaran"
        elif kunci == "materi pembelajaran":
            kunci = "materi"
        elif kunci.startswith("penutup"):
            kunci = "penutup"
        if kunci.startswith(LEWATI_BAGIAN):
            continue
        if not kunci.startswith(("contoh soal", "latihan soal")):
            isi = saring(isi)  # soal disaring per soal di _soal
        if kunci == "pendahuluan":
            blok = penyusun.blok_dari(isi)
            naskah["pendahuluan"]["paragraf"] = [x for b in blok for x in (
                [b["teks"]] if b["jenis"] == "paragraf" else [f"**{b['teks']}**"] if b["jenis"] == "subjudul"
                else [f"• {x}" for x in b["butir"]] if b["jenis"] == "poin" else [])]
            gambar = [b for b in blok if b["jenis"] == "gambar"]
            if gambar:
                naskah["pendahuluan"]["gambar"] = gambar[0]["file"]
                naskah["pendahuluan"]["paragraf"] += [b["keterangan"] for b in gambar[:1] if b.get("keterangan")]
        elif kunci == "tujuan pembelajaran":
            for el in isi:
                if el["jenis"] != "paragraf" or not el["polos"] or el["polos"].endswith(":"):
                    continue  # kalimat pengantar 'Setelah mempelajari materi ini, peserta mampu:'
                t = re.sub(r"^(\*\*)?[•▪●◦\-–]\s+", r"\1", el["teks"]).replace("**", "")
                t = re.sub(r"^\d+[.)]\s+", "", t)
                kko, _, sisa = t.partition(" ")
                naskah["tujuan"]["butir"].append({"level": "", "kko": kko, "teks": sisa})
        elif kunci == "materi":
            sub, isi_sub = None, []
            for el in isi:
                if el["jenis"] == "h2" and el["polos"]:
                    if sub is not None or isi_sub:
                        tambah_subbab(sub or judul_bersih, isi_sub, False)
                    sub, isi_sub = el["teks"], []
                else:
                    isi_sub.append(el)
            if sub is not None or isi_sub:
                tambah_subbab(sub or judul_bersih, isi_sub, False)
        elif kunci.startswith("contoh soal") or kunci.startswith("latihan soal"):
            soal = _soal(penyusun, isi)
            if soal:
                naskah["materi"].append({"subbab": "Contoh Soal", "ikon": "checklist", "slide": soal})
        elif kunci == "penutup":
            pesan, lain = [], []
            for el in isi:
                if el["jenis"] == "h2" and _kunci_bagian(el["polos"]).startswith(LEWATI_BAGIAN + ("materi inti yang wajib dimuat",)):
                    break  # catatan penyusun di ujung penutup, mis. 'PEMETAAN REFERENSI TAMBAHAN DAN CATATAN QC'
                if el["jenis"] == "paragraf" and el["polos"]:
                    pesan.append(el["teks"])
                elif el["jenis"] in ("tabel", "gambar"):
                    lain.append(el)
            naskah["penutup"]["pesan"] = " ".join(pesan)
            if lain:  # tabel rangkuman tidak muat di kartu pesan: jadi slide sendiri
                tambah_subbab("Rangkuman", lain, False)
        elif kunci in ("sumber", "referensi", "daftar pustaka"):
            continue
        else:
            tambah_subbab(nama, isi, True)

    naskah["penutup"]["ringkasan"] = [
        {"judul": sub["subbab"], "teks": ""} for sub in naskah["materi"] if sub["subbab"] not in ("Contoh Soal", "Rangkuman")][:6]
    if not (naskah["pendahuluan"]["paragraf"] and naskah["tujuan"]["butir"]):
        naskah["meta"]["lewati_kosong"] = True  # bagian yang tidak ditandai penulis: slidenya dilewati
    return naskah


def pisah_soal(naskah, nama_file_soal):
    """(naskah materi, naskah contoh soal) dari dokumen yang memuat keduanya.

    Naskah contoh soal hanya berisi soal dan pembahasannya (meta.bagian = "soal"); pendahuluan, tujuan,
    dan penutup milik materi, jadi dikosongkan dan slidenya dilewati buat_materi.
    """
    def berisi_soal(sub):
        return any(s["tipe"] == "contoh_soal" for s in sub["slide"])

    materi, soal = copy.deepcopy(naskah), copy.deepcopy(naskah)
    materi["materi"] = [s for s in materi["materi"] if not berisi_soal(s)]
    soal["materi"] = [s for s in soal["materi"] if berisi_soal(s)]
    judul = naskah["meta"]["judul"].removeprefix("Soal ")  # 'Soal Cerita Aljabar' -> 'Contoh Soal Cerita Aljabar'
    soal["meta"].update(judul=f"Contoh Soal {judul}", nama_file=nama_file_soal, bagian="soal")
    soal["meta"].pop("subjudul", None)
    soal["pendahuluan"] = {"paragraf": [], "pemantik": "", "ikon": "presenter"}
    soal["tujuan"] = dict(naskah["tujuan"], butir=[])
    soal["penutup"] = {"judul": "Yang Perlu Diingat", "ringkasan": [], "pesan": ""}
    return materi, soal


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
