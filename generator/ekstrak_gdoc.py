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


GANTI_BARIS = " "
LATAR_HIJAU = ("d9ead3", "b6d7a8")  # hijau muda Google Docs


def _disorot(rPr, penanda=""):
    """Run (teks, rumus, gambar) atau tanda paragraf yang diberi sorotan warna. penanda: 'merah' = teks merah juga
    penanda; 'hijau' = teks berlatar hijau muda juga penanda."""
    if rPr is None:
        return False
    h = rPr.find(f"{W}highlight")
    if h is not None and h.get(f"{W}val") not in (None, "none", "white"):
        return True
    if penanda == "hijau":
        s = rPr.find(f"{W}shd")
        return s is not None and (s.get(f"{W}fill") or "").lower() in LATAR_HIJAU
    return penanda == "merah" and _warna_merah(rPr)


def _warna_merah(rPr):
    c = rPr.find(f"{W}color")
    return c is not None and bool(re.fullmatch(r"[89a-f][0-9a-f][0-3][0-9a-f][0-3][0-9a-f]", (c.get(f"{W}val") or "").lower()))


def _merah(r):
    rPr = r.find(f"{W}rPr")
    return rPr is not None and _warna_merah(rPr)


def _latar_kuning(p):
    """Paragraf yang sorotan kuningnya dipasang sebagai latar paragraf, bukan di teksnya."""
    s = p.find(f"{W}pPr/{W}shd")
    return s is not None and (s.get(f"{W}fill") or "").lower() == "ffff00"


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


def _polos(teks):
    """Teks tanpa penanda tebal (**), miring (<i>), coret (<s>), dan merah (<r>)."""
    return re.sub(r"\*\*|</?[isr]>", "", teks)


def _gabung(potongan):
    """[(teks, jenis)] -> markup. jenis 'm' = rumus; selain itu gabungan huruf 'b' tebal, 'i' miring, 's' coret
    merah, 'r' teks merah; '' = biasa."""
    keluar, kini, isi = [], None, ""
    for t, j in potongan + [("", "#")]:
        if j != kini:
            if isi:
                if kini and kini != "m" and isi.strip():
                    awal, akhir = len(isi) - len(isi.lstrip()), len(isi.rstrip())
                    tanda = [x for x in "isr" if x in kini]
                    buka = ("**" if "b" in kini else "") + "".join(f"<{x}>" for x in tanda)
                    tutup = "".join(f"</{x}>" for x in reversed(tanda)) + ("**" if "b" in kini else "")
                    keluar.append(isi[:awal] + buka + isi[awal:akhir] + tutup + isi[akhir:])
                elif kini == "m" and isi.strip():
                    keluar.append("$" + isi.strip() + "$")
                else:
                    keluar.append(isi)
            kini, isi = j, ""
        isi += t
    for k in range(1, len(keluar)):  # rumus Word yang menempel ke kata: '75%' + 'lebih sedikit'
        a, b = keluar[k - 1], keluar[k]
        if a[-1:] == "$" and len(a) > 2 and b[:1].isalpha() or b[:1] == "$" and len(b) > 2 and a[-1:].isalpha():
            keluar[k] = " " + b
    return re.sub(r"[ \t]+", " ", "".join(keluar)).strip()


def _coret(r):
    s = r.find(f"{W}rPr/{W}strike")
    return s is not None and s.get(f"{W}val") not in ("0", "false")


def _teks_paragraf(p, miring=False, penanda="", coret=False, rumus=None, ganti_baris=False):
    """Teks paragraf dengan markup, daftar gambar (rId, lebar, tinggi, disorot), status tebal/miring penuh, dan
    teks yang hanya memuat kalimat disorot. miring: cetak miring ikut dicatat sebagai <i>..</i>; coret: coretan
    (<s>..</s>) dan angka merah (<r>..</r>) ikut dicatat."""
    potongan, gambar = [], []
    semua_tebal, semua_miring, ada_teks = True, True, False
    latar = _latar_kuning(p)
    par_disorot = latar or _disorot(p.find(f"{W}pPr/{W}rPr"), penanda)

    def jalan(el):
        nonlocal semua_tebal, semua_miring, ada_teks
        for anak in el:
            tag = anak.tag
            if tag == f"{W}r":
                for d in anak.iter(f"{W}drawing"):
                    blip = d.find(".//{%s}blip" % NS["a"])
                    if rumus and blip is not None and blip.get("{%s}embed" % NS["r"]) in rumus:  # rumus berupa gambar di tengah kalimat
                        ada_teks, semua_tebal = True, False
                        potongan.append((" " + rumus[blip.get("{%s}embed" % NS["r"])] + " ", "m",
                                         _disorot(anak.find(f"{W}rPr"), penanda) or par_disorot))
                        continue
                    ext = d.find(".//{%s}extent" % NS["wp"])
                    if blip is not None:
                        gambar.append((blip.get("{%s}embed" % NS["r"]),
                                       int(ext.get("cx")) / EMU if ext is not None else None,
                                       int(ext.get("cy")) / EMU if ext is not None else None,
                                       _disorot(anak.find(f"{W}rPr"), penanda) or par_disorot))
                t = ""
                for x in anak:
                    if x.tag == f"{W}t":
                        t += x.text or ""
                    elif ganti_baris and x.tag == f"{W}br" and x.get(f"{W}type") != "page":
                        t += GANTI_BARIS
                    elif x.tag in (f"{W}tab", f"{W}br"):
                        t += " "
                posisi = anak.find(f"{W}rPr/{W}vertAlign")
                naik = {"superscript": "^", "subscript": "_"}.get(posisi.get(f"{W}val") if posisi is not None else None)
                s = latar or _disorot(anak.find(f"{W}rPr"), penanda)
                for ke, t in enumerate(t.split(GANTI_BARIS)):
                    if ke:
                        potongan.append((GANTI_BARIS, "br", s))
                    if t.strip() and naik:  # superskrip/subskrip teks biasa: cm^{2}, H_{2}O
                        ada_teks = True
                        potongan.append((f"{naik}{{{t.strip()}}}", "m", s))
                    elif t:
                        if t.strip():
                            ada_teks = True
                            semua_tebal &= _tebal(anak)
                            semua_miring &= _miring(anak)
                        gaya = ("b" if _tebal(anak) else "") + ("i" if miring and _miring(anak) else "")
                        if coret and t.strip():
                            gaya += "s" if _coret(anak) else "r" if _merah(anak) else ""
                        potongan.append((t, gaya, s))
            elif tag in (f"{M}oMath", f"{M}oMathPara"):
                t = omml.teks(anak)  # pecahan, pangkat, akar dibaca strukturnya (bukan digabung teksnya)
                if t.strip():
                    ada_teks = True
                    semua_tebal = False
                    potongan.append((" " + t + " ", "m", latar or any(
                        _disorot(r.find(f"{W}rPr"), penanda) for r in anak.iter(f"{M}r"))))
            elif tag in (f"{W}hyperlink", f"{W}smartTag", f"{W}ins", f"{W}sdt", f"{W}sdtContent"):
                jalan(anak)

    jalan(p)
    return (_gabung([x[:2] for x in potongan]), gambar, (semua_tebal and ada_teks), (semua_miring and ada_teks),
            _gabung(_kalimat_disorot(potongan)))


def _format_nomor(z):
    """{(numId, ilvl): 'bullet' | 'decimal' | 'lowerLetter' | ...} dari numbering.xml."""
    if "word/numbering.xml" not in z.namelist():
        return {}
    akar = etree.fromstring(z.read("word/numbering.xml"))
    fmt = {}
    for a in akar.findall(f"{W}abstractNum"):
        for lv in a.findall(f"{W}lvl"):
            n = lv.find(f"{W}numFmt")
            fmt[a.get(f"{W}abstractNumId"), lv.get(f"{W}ilvl")] = n.get(f"{W}val") if n is not None else ""
    hasil = {}
    for n in akar.findall(f"{W}num"):
        abst = n.find(f"{W}abstractNumId")
        for (a, lv), f in fmt.items():
            if abst is not None and abst.get(f"{W}val") == a:
                hasil[n.get(f"{W}numId"), lv] = f
    return hasil


def baca_docx(path, miring=False, penanda="", coret=False, rumus_gambar=None, ganti_baris=False):
    """Daftar elemen berurutan: judul/subjudul/h1/h2/h3/paragraf/tabel/gambar.

    miring: pertahankan cetak miring (soal Bahasa Inggris menanyakan 'the italicised part').
    penanda: 'merah' = teks merah adalah penanda wajib PPT, seperti highlight kuning (Badges, UTBK PU);
    'hijau' = teks berlatar hijau muda adalah penandanya (UTBK Literasi).
    coret: pertahankan coretan dan angka merah (pencoretan nol di Perkalian dan Pembagian Setara).
    rumus_gambar: {nama berkas gambar: rumus} untuk rumus yang disisipkan sebagai gambar di tengah kalimat;
    ganti_baris: ganti baris di dalam paragraf dipertahankan sebagai GANTI_BARIS."""
    z = zipfile.ZipFile(path)
    doc = etree.fromstring(z.read("word/document.xml"))
    gaya = etree.fromstring(z.read("word/styles.xml"))
    nama_gaya = {}
    for s in gaya.iter(f"{W}style"):
        n = s.find(f"{W}name")
        nama_gaya[s.get(f"{W}styleId")] = (n.get(f"{W}val") if n is not None else s.get(f"{W}styleId")).lower()
    rels = etree.fromstring(z.read("word/_rels/document.xml.rels"))
    target = {r.get("Id"): r.get("Target") for r in rels}
    format_nomor = _format_nomor(z)
    nama_gambar = {k: os.path.basename(v or "") for k, v in target.items()}
    rumus = {k: rumus_gambar[v] for k, v in nama_gambar.items() if v in rumus_gambar} if rumus_gambar else None

    elemen = []
    for el in doc.find(f"{W}body"):
        if el.tag == f"{W}p":
            ps = el.find(f"{W}pPr/{W}pStyle")
            st = nama_gaya.get(ps.get(f"{W}val"), "") if ps is not None else ""
            numPr = el.find(f"{W}pPr/{W}numPr")
            daftar = numPr is not None
            num = None if numPr is None else tuple(
                (numPr.find(f"{W}{k}").get(f"{W}val") if numPr.find(f"{W}{k}") is not None else "0") for k in ("numId", "ilvl"))
            fmt = format_nomor.get(num, "")
            teks, gambar, tebal, semua_miring, sorot = _teks_paragraf(el, miring, penanda, coret, rumus, ganti_baris)
            jenis = {"title": "judul", "subtitle": "subjudul", "heading 1": "h1", "heading 2": "h2",
                     "heading 3": "h3", "heading 4": "h3"}.get(st, "paragraf")
            for rid, w, h, s in gambar:
                berkas = target.get(rid, "")
                elemen.append({"jenis": "gambar", "data": z.read("word/" + berkas) if berkas else b"",
                               "ext": os.path.splitext(berkas)[1].lower(), "w": w, "h": h, "sorot": s})
            polos = _polos(teks).strip()
            if polos or jenis != "paragraf":
                elemen.append({"jenis": jenis, "teks": teks, "polos": polos, "tebal": tebal,
                               "miring": semua_miring, "daftar": daftar, "sorot": sorot, "gaya": st, "fmt": fmt, "num": num})
        elif el.tag == f"{W}tbl":
            baris, par_baris, gambar_baris, sorot_baris = [], [], [], []
            for tr in el.findall(f"{W}tr"):
                sel, par_sel, gambar_sel, sorot_sel = [], [], [], []
                for tc in tr.findall(f"{W}tc"):
                    tcPr = tc.find(f"{W}tcPr")
                    rentang = tcPr.find(f"{W}gridSpan") if tcPr is not None else None
                    lanjut = tcPr.find(f"{W}vMerge") if tcPr is not None else None
                    hasil = [_teks_paragraf(p, miring, penanda, coret) for p in tc.findall(f".//{W}p")]
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
    t = CATATAN_PENULIS.sub("", _polos(judul)).strip()
    t = re.sub(r"^(?:\d+(?:\.\d+)+\.?|\d+[.)]?|[A-H][.)])\s+", "", t).strip()
    return _rapikan_kapital(t)


def _kunci_bagian(nama):
    """Nama bagian Heading 1 untuk dicocokkan: '1. PENDAHULUAN', 'A. Pendahuluan [WAJIB PPT]' -> 'pendahuluan'."""
    return _bersih(nama).lower().strip(" :")


POLA_BAGIAN = re.compile(r"^[A-H]\.\s+(pendahuluan|tujuan pembelajaran|materi pembelajaran|strategi|contoh soal|"
                         r"pola soal|kesalahan umum|penutup)\b", re.I)


# Kalimat untuk penyusun atau tim produksi, bukan untuk peserta: tidak masuk PPT walau berada di bagian wajib.
CATATAN_PRODUKSI = [(re.compile(p), g) for p, g in (
    (r"PPT (?:Materi OLD )?(?:mencantumkan|memilih)\b[^.]*\.\s*", ""),
    (r"(Untuk membaca susunan) pada PPT,", r"\1,"),
    (r"Dalam materi PPT, contoh gambar sumber dipertahankan[^.]*\.\s*", ""),
    (r"\s*Catatan QC: dokumen sumber.*$", ""),
    (r"Dalam dokumen Word ini, [^.]*\.\s*", ""),
    (r"^Catatan ketelitian durasi: sumber rujukan.*$", ""),
    (r"^Contoh berikut adalah latihan pembelajaran yang disusun dari blueprint.*$", ""),
    (r"^Contoh 1 mempertahankan nama, durasi, dan batas waktu pada riset.*$", ""),
    (r"^Materi ini menyertakan contoh pemeriksaan kasus tidak valid sebagai ilustrasi QC.*$", ""),
    (r"^(?:•\s*)?Mengenali soal yang tidak memiliki solusi .*dikoreksi saat produksi\.?$", ""),
    (r"^Gambar berikut (?:diambil dari halaman contoh soal pada referensi|menggunakan halaman contoh soal asli referensi).*$", ""),
    # UTBK
    (r"Strategi ini sangat relevan untuk IPK [A-Z.0-9*]+\s*", ""),
    (r"; kunci sumber tidak valid(?=\.)", ""),
    (r"\s*\(contoh penguatan IPK\)", ""),
    (r"\bKoreksi kunci pada rasio tiga kelompok", "Rasio tiga kelompok"),
    (r"\s*\(Bukan Subbab Mapping\)", ""),
    (r" pada bank soal internal(?= digunakan)", ""),
    (r"\be-book ini memadukan", "materi ini memadukan"),
)]


def _buang_catatan(t):
    for pola, ganti in CATATAN_PRODUKSI:
        t = pola.sub(ganti, t)
    t = re.sub(r"\s*\*\*\s*\*\*\s*|\s*<([isr])>\s*</\1>\s*", " ", t)  # sisa label yang dicetak tebal: '**  **'
    return re.sub(r"^((?:\*\*|<[isr]>)+)\s+", r"\1", t.strip())  # sisa spasi label di dalam teks tebal


def koreksi(elemen, daftar):
    """Koreksi tim QC yang belum ada di dokumen (koreksi.json): {"cari", "ganti"} di semua teks dan sel tabel, atau
    {"hapus_gambar_di": judul}. Bila dokumennya sudah diperbaiki, butirnya tidak ditemukan lagi dan diberi peringatan."""
    for butir in daftar:
        kena = 0
        if butir.get("hapus_gambar_di"):
            judul, sisa = "", []
            for el in elemen:
                if el["jenis"] in TINGKAT and el.get("polos"):
                    judul = _bersih(el["teks"]).lower()
                if el["jenis"] == "gambar" and judul == butir["hapus_gambar_di"]:
                    kena += 1
                else:
                    sisa.append(el)
            elemen = sisa
        else:
            a, b = butir["cari"], butir["ganti"]
            for el in elemen:
                for k in ("teks", "polos", "sorot"):
                    if isinstance(el.get(k), str) and a in el[k]:
                        el[k] = el[k].replace(a, b)
                        kena += 1
                if el["jenis"] == "tabel":
                    for r, par in zip(el["baris"], el["par"]):
                        for c, isi in enumerate(r):
                            if a in isi:
                                r[c] = isi.replace(a, b)
                                kena += 1
                        for isi in par:
                            isi[:] = [x.replace(a, b) for x in isi]
        if not kena:
            print(f"  PERINGATAN koreksi tidak ditemukan (dokumen sudah diperbaiki?): {butir.get('cari') or butir}")
    return elemen


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
        el = dict(el, **{k: _buang_catatan(CATATAN_PENULIS.sub("", el[k]).strip()) for k in ("teks", "polos", "sorot")})
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
LABEL_SEMUA = re.compile(r"\[[^\]]*\bsemua\b[^\]]*\]", re.I)  # [semua contoh soal wajib ada ...]: soal [PENGAYAAN] pun masuk
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
                aturan.append((tingkat, "semua" if LABEL_SEMUA.search(teks) else "wajib"))
            elif LABEL_PENGAYAAN.search(teks) and not (aturan and aturan[-1][1] == "semua"):
                aturan.append((tingkat, "pengayaan"))
        kini = lingkup = aturan[-1][1] if aturan else "sorot"
        kini = {"semua": "wajib", "pengayaan": "sorot"}.get(kini, kini)
        el = dict(el, lingkup=lingkup)
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
    _ikutkan_gambar(hasil)
    for el in hasil:
        if el.get("paksa"):  # soal tanpa tanda yang tetap dimasukkan (rapikan_utbk)
            el.update(masuk=True, utuh=True)
    return hasil


def _ikutkan_gambar(elemen):
    """Gambar yang tidak disorot ikut teksnya.

    Sorotan di Google Docs hanya tampak pada teks, jadi gambar jarang ikut disorot walau dipakai oleh kalimat
    yang disorot ('Hijau menempati bagian 1, kemudian 2 ...').
    - Gambar di antara paragraf yang masuk (sampai batas subbab) selalu ikut.
    - Bila penulis tidak pernah menyorot gambar di bagian materi, gambar ikut subbabnya: masuk bila ada isi yang
      masuk di bawah judul yang sama. Gambar pembuka Heading 2 juga melihat Heading 3 di bawahnya.
    - Bila penulis menyorot sebagian gambar materi, gambar lain dianggap sengaja tidak ditandai.
    """
    soal, bagian_soal = [], False
    for el in elemen:
        if el["jenis"] == "h1":
            bagian_soal = _kunci_bagian(el["polos"]).startswith(("contoh soal", "latihan soal"))
        soal.append(bagian_soal)
    disorot = any(el["jenis"] == "gambar" and el["sorot"] and not s for el, s in zip(elemen, soal))
    tambah = []
    for i, el in enumerate(elemen):
        if el["jenis"] != "gambar" or el["masuk"] or el["lingkup"] == "profil":
            continue  # bagian 'hanya foto dan profil' tidak memuat gambar lain
        atas, bawah = _tetangga_masuk(elemen[i - 1::-1] if i else []), _tetangga_masuk(elemen[i + 1:])
        if False not in (atas, bawah) and True in (atas, bawah):
            tambah.append(i)
        elif not disorot and not soal[i]:
            judul = next((k for k in range(i, -1, -1) if elemen[k]["jenis"] in TINGKAT), None)
            if judul is None:
                continue  # sampul dokumen
            batas = TINGKAT[elemen[judul]["jenis"]]
            for lain in elemen[judul + 1:]:
                if TINGKAT.get(lain["jenis"], 9) <= batas:
                    break
                if lain["masuk"] and lain["jenis"] not in TINGKAT and lain["jenis"] != "gambar":
                    tambah.append(i)
                    break
    for i in tambah:
        elemen[i]["masuk"] = True
        if not soal[i]:
            elemen[i]["tambahan"] = "gambar di subbab yang masuk"


NOMOR_LANGKAH = re.compile(r"^(\d+)[.)]\s")


def _butir(el):
    return el["jenis"] == "paragraf" and (el["daftar"] or bool(re.match(r"^[•▪●◦]\s", el["polos"])))


def pelengkap(elemen, pilihan=(), minimal=False):
    """Tambahan di luar sorotan, supaya isi wajib tidak pincang. Setiap tambahan diberi el['tambahan'] (alasan).

    Highlight berarti wajib masuk, bukan satu-satunya yang boleh masuk. Yang ditambahkan:
    - gambar dan tabel di bawah judul yang sudah punya isi masuk (beserta keterangan gambarnya);
    - langkah bernomor atau butir daftar yang hanya sebagian ditandai;
    - pilihan per dokumen di pelengkap.json (subbab tanpa tanda, definisi, tujuan dan pesan penutup kosong).
    Bagian [PENGAYAAN] dan 'hanya foto dan profil' tidak ditambah. minimal: tabel hanya bila diapit isi yang masuk
    (dokumen yang menyebut hanya bagian bertanda yang perlu ditampilkan).
    """
    if not any("lingkup" in el for el in elemen):
        return elemen  # dokumen tanpa tanda: seluruhnya sudah masuk

    def tambah(el, alasan):
        if not el["masuk"]:
            el.update(masuk=True, utuh=True, tambahan=alasan)

    bagian, soal, judul, konteks = "", False, {}, []
    for i, el in enumerate(elemen):
        t = TINGKAT.get(el["jenis"])
        if t and el["polos"]:
            judul = {k: v for k, v in judul.items() if k < t}
            judul[t] = (i, _bersih(el["teks"]).lower())
            if t == 1:
                bagian = _kunci_bagian(el["polos"])
                soal = bagian.startswith(("contoh soal", "latihan soal"))
        konteks.append((bagian, soal, dict(judul)))
    for butir in pilihan:
        kena, dipakai = 0, set()
        for el, (_, _, jd) in zip(elemen, konteks):
            if butir["di"].lower() not in [v for _, v in jd.values()] or el["jenis"] in TINGKAT:
                continue
            if el["jenis"] in ("gambar", "tabel"):
                ya = butir.get(el["jenis"], False)
            else:
                isi = re.sub(r"\s", "", el["polos"])
                cocok = [a for a in butir.get("awal", []) if isi.startswith(re.sub(r"\s", "", a))]
                dipakai.update(cocok)
                ya = bool(cocok)
            if ya:
                tambah(el, butir["alasan"])
                kena += 1
        sisa = [a for a in butir.get("awal", []) if a not in dipakai]
        if sisa or not kena:
            print(f"  PERINGATAN pelengkap '{butir['di']}': tidak ditemukan {sisa or butir}")
    for i, (el, (bag, di_soal, jd)) in enumerate(zip(elemen, konteks)):
        if not bag or bag.startswith(LEWATI_BAGIAN) or el["lingkup"] in ("profil", "pengayaan"):
            continue
        if el["jenis"] in ("gambar", "tabel") and not el["masuk"] and not di_soal and jd:
            awal = max(k for k, _ in jd.values())
            batas = TINGKAT[elemen[awal]["jenis"]]
            ada = False
            for lain in elemen[awal + 1:]:
                if TINGKAT.get(lain["jenis"], 9) <= batas:
                    break
                ada |= lain["masuk"] and lain["jenis"] not in TINGKAT and "tambahan" not in lain
            diapit = i and elemen[i - 1]["masuk"] and i + 1 < len(elemen) and elemen[i + 1]["masuk"]
            if ada and (not minimal or diapit):
                tambah(el, f"{el['jenis']} di subbab yang masuk")
                ket = elemen[i + 1] if i + 1 < len(elemen) else None
                if el["jenis"] == "gambar" and ket and "internal" not in ket.get("polos", "") and \
                        re.match(r"^(Gambar|Diagram|Ilustrasi)\s*\d*[.:]", ket.get("polos", "")):
                    tambah(ket, "keterangan gambar")
        if el["jenis"] != "paragraf":
            continue
        nomor = NOMOR_LANGKAH.match(el["polos"])
        if nomor and nomor.group(1) == "1":
            langkah = []
            for lain in elemen[i:]:
                m = NOMOR_LANGKAH.match(lain["polos"]) if lain["jenis"] == "paragraf" else None
                if not m or int(m.group(1)) != len(langkah) + 1:
                    break
                langkah.append(lain)
            if len(langkah) > 1 and any(x["masuk"] for x in langkah):
                for x in langkah:
                    tambah(x, "langkah yang bolong")
        if _butir(el) and not (i and elemen[i - 1]["jenis"] == "paragraf" and _butir(elemen[i - 1])):
            daftar = []
            for lain in elemen[i:]:
                if lain["jenis"] != "paragraf" or not _butir(lain):
                    break
                daftar.append(lain)
            if len(daftar) > 1 and any(x["masuk"] and "tambahan" not in x for x in daftar):
                for x in daftar:
                    tambah(x, "butir daftar yang bolong")
    return elemen


def _tetangga_masuk(elemen):
    """Paragraf pertama di elemen masuk PPT? None bila bertemu judul subbab lebih dulu."""
    for el in elemen:
        if el["jenis"] in ("h1", "h2"):
            return None
        if el.get("polos"):
            return el["masuk"]
    return None


# 'Langkah 2. Kalimat lengkap.' adalah isi, bukan judul (judul langkah tidak diakhiri titik). Tanpa ini baris yang
# pendek menjadi judul dan baris yang panjang menjadi paragraf.
LANGKAH_KALIMAT = re.compile(r"^Langkah\s+\d+[.:]?\s+[^\s(—–-].*\.$", re.S)


def _seperti_judul(el):
    """Tingkat judul elemen: Heading, atau paragraf tebal pendek yang menjadi subjudul di slide (4)."""
    if el["jenis"] in TINGKAT:
        return TINGKAT[el["jenis"]]
    if el["jenis"] != "paragraf" or len(el["polos"]) > 90 or el["daftar"]:
        return None
    polos = el["polos"]
    if re.fullmatch(r"cara cepat:?", polos.lower()) or (
            re.match(r"^(Langkah\s+\d+\b|Kesimpulan$|Pemisalan$)", polos) and not LANGKAH_KALIMAT.match(polos)):
        return 4
    simbol = any(c in SIMBOL_SOROT for c in polos.rstrip(":")) or re.search(r"\w\s+[:=]\s+\w", polos)
    return 4 if el["tebal"] and not simbol and not re.match(r"^[^:]+: .+\.$", polos) else None


def saring(elemen):
    """Elemen satu bagian (bukan soal) yang masuk PPT. Judul ikut bila ada isi di bawahnya yang masuk."""
    ikut = [False] * len(elemen)
    for i, el in enumerate(elemen):
        tk = _seperti_judul(el)
        if tk is None:
            ikut[i] = el["masuk"]
            continue
        label = tk == 4 and not el["masuk"] and (el.get("ketat") or re.fullmatch(r"(contoh|misalnya|ilustrasi)\s*:?", el["polos"], re.I))
        for k, lain in enumerate(elemen[i + 1:]):
            tl = _seperti_judul(lain)
            if tl is not None and tl <= tk:
                break
            if label and k:
                break  # label 'Contoh:' tanpa tanda hanya ikut bila isi tepat di bawahnya masuk
            if tk == 4 and lain["jenis"] == "gambar" and not lain.get("sorot"):
                continue  # gambar yang hanya ikut subbabnya tidak menarik teks tebal di atasnya
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
            el = dict(el, teks=el["sorot"], polos=_polos(el["sorot"]).strip())
        hasil.append(el)
    return hasil


def _kalimat_pertama(teks):
    t = _polos(teks)
    m = re.match(r"(.+?[.!?])(\s|$)", t)
    return (m.group(1) if m else t).strip()


def _catatan_tambahan(el):
    isi = el.get("polos") or ("Tabel: " + " | ".join(_polos(c) for c in el["baris"][0] if c) if el["jenis"] == "tabel" else "Gambar")
    return {"alasan": el["tambahan"], "isi": isi}


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
            nama = _polos(par[0] if par else "")
            peran = _polos(" ".join(par[1:]))
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
            j, mulai = el["jenis"], i
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
                # sel berisi beberapa paragraf ('Premis I:' / 'Premis II:' / 'Kesimpulan:') tetap beberapa baris
                baris = [["\n".join(p) if len(p) > 1 else c for c, p in zip(r, par)] if len(par) == len(r) else r
                         for r, par in zip(el["baris"], el["par"]) if any(c.strip() for c in r)]
                if profil:
                    blok.append(profil)
                elif len(baris) >= 2 and len(baris[0]) >= 1:
                    n = max(len(r) for r in baris)
                    baris = [r + [""] * (n - len(r)) for r in baris]
                    blok.append({"jenis": "tabel", "kolom": [_polos(c).replace("\n", " ") for c in baris[0]],
                                 "baris": baris[1:]})
                elif baris:
                    for r in baris:
                        blok.append({"jenis": "paragraf", "teks": " · ".join(c for c in r if c)})
            elif j in ("h3", "h2", "h1"):
                if _bersih(el["teks"]):  # heading kosong di dokumen: abaikan
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
                    butir = re.sub(r"^((?:\*\*|<[isr]>)*)[•▪●◦\-–]\s+", r"\1", t)
                    if blok and blok[-1]["jenis"] == "poin":
                        blok[-1]["butir"].append(butir)
                    else:
                        blok.append({"jenis": "poin", "butir": [butir]})
                elif re.fullmatch(r"\$[^$]+\$", t.strip()):
                    blok.append({"jenis": "sorot", "teks": t.strip()})
                elif el["tebal"] and re.fullmatch(r"[A-Z][a-z]+ [a-z]+ • [^→←↔=≡]{12,}", polos):
                    # judul bagan: 'Garis waktu • Latar dan kejadian'; tingkat 2 = paragraf tebal, di bawah Heading
                    blok.append({"jenis": "subjudul", "teks": polos.rstrip(":"), "tingkat": 2})
                elif re.match(r"^Jawaban\s*:\s*\S", polos) and len(polos) <= 200:
                    # 'Jawaban: 36' adalah isi, bukan judul bagian; tampilannya sama di semua soal
                    isi = re.sub(r"^Jawaban\s*:\s*", "", t.replace("**", "")).strip()
                    blok.append({"jenis": "paragraf", "teks": f"**Jawaban:** {isi}" if re.search(r"\$|</?[isr]>", isi) else f"**Jawaban: {isi}**"})
                elif el["tebal"] and (re.match(r"^Konsep\s*:\s*\S", polos) or re.match(r"^Kesimpulan\s*:\s*[A-Z][^=→≡]+\.$", polos)):
                    blok.append({"jenis": "paragraf", "teks": t})
                elif el["tebal"] and (any(c in SIMBOL_SOROT for c in polos.rstrip(":")) or
                                      re.search(r"\w\s+[:=]\s+\w", polos) or
                                      re.search(r"^[↑↓]| \| ", polos) or  # baris bagan
                                      re.match(r"^1\. \D+ 2\. \D+ 3\. \D", polos)):  # '1. Posisi 2. Petunjuk 3. Pasangan'
                    if polos.startswith("↓ ") and blok and blok[-1].get("tingkat") == 2 and blok[-1]["jenis"] == "subjudul":
                        blok[-1] = {k: v for k, v in dict(blok[-1], jenis="sorot").items() if k != "tingkat"}
                    blok.append({"jenis": "sorot", "teks": polos})
                elif el["tebal"] and len(polos) <= 90:
                    blok.append({"jenis": "subjudul", "teks": polos.rstrip(":"), "tingkat": 2})
                elif LANGKAH_KALIMAT.match(polos):  # labelnya saja yang tebal
                    blok.append({"jenis": "paragraf", "teks": re.sub(
                        r"^(?:\*\*)?(Langkah\s+\d+[.:]?)(?:\s*\*\*)?\s*", r"**\1** ", t, count=1)})
                elif re.match(r"^(Langkah\s+\d+\b|Kesimpulan$|Pemisalan$)", polos) and len(polos) <= 90:
                    blok.append({"jenis": "subjudul", "teks": polos, "tingkat": 2})
                else:
                    blok.append({"jenis": "paragraf", "teks": t})
            for x in elemen[mulai:i + 1]:
                if x.get("tambahan") and blok:  # dicatat di catatan pembicara slide dan daftar tinjauan
                    blok[-1].setdefault("tambahan", []).append(_catatan_tambahan(x))
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
    ditandai = any(e.get("awal_soal") for e in elemen)  # awal soal dan bacaan sudah ditandai (rapikan_utbk)
    bacaan = None
    for el in elemen:
        j, polos = el["jenis"], el.get("polos", "")
        if el.get("zona"):  # contoh soal subbab berikutnya
            label, kini, mode, bacaan = None, None, None, None
            continue
        if el.get("awal_bacaan"):
            bacaan = {"tipe": "blok", "judul": el["teks"], "blok": []}
            hasil.append(bacaan)
            kini, mode = None, None
            continue
        if ditandai:
            mulai = el.get("awal_soal")
            per_judul = True
        else:
            mulai = (POLA_JUDUL_SOAL.match(_bersih(polos)) if per_judul else j == "h2") and j in ("h2", "h3")
        if mulai:
            bacaan = None
        if bacaan is not None:
            bacaan["blok"].append(el)
            continue
        if mulai and polos:
            kini = {"tipe": "contoh_soal", "judul": _bersih(el["teks"]), "soal": [], "opsi": {}, "kunci": "",
                    "pembahasan": [], "_tanda": [], "_judul": bool(el.get("sorot", "").strip())}
            if el.get("label") or label:
                kini["label"] = el.get("label") or label
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
                    r"^(\*\*|<i>)*pembahasan\s*[:.]?\s*(\*\*|</i>)*\s*", "", el["sorot"], flags=re.I)))
            continue
        m_opsi = re.match(r"^\(?([A-E])[.)]\s+(.*)$", polos) if j == "paragraf" else None
        if mode in ("soal", "opsi") and m_opsi and (m_opsi.group(1) == "A" or mode == "opsi"):
            teks = re.sub(r"^((?:\*\*|<[isr]>)*)\(?[A-E][.)]\s+", r"\1", el["teks"])
            kini["opsi"][m_opsi.group(1)] = teks
            kini["_tanda"].append((len(polos), el["masuk"]))
            mode = "opsi"
            continue
        if mode not in ("bahas", "kunci"):
            kini["_tanda"].append((len(polos) or 20, el["masuk"]))
        kini["pembahasan" if mode in ("bahas", "kunci") else "soal"].append(el)
    # soal masuk (pertanyaan, opsi, dan kunci utuh) bila sebagian besar pertanyaan dan opsinya disorot;
    # pembahasannya hanya bagian yang disorot, kecuali jawaban uraian. Bila yang ditandai hanya judul soalnya
    # (isinya tidak ada yang ditandai), soal dan pembahasannya masuk utuh.
    teks_bacaan = [s for s in hasil if s["tipe"] == "blok"]
    for s in hasil:
        if s["tipe"] == "blok":
            continue
        s["_judul"] = s["_judul"] and not any(ya for _, ya in s["_tanda"]) and not any(e["masuk"] for e in s["pembahasan"])
        if s["_judul"]:
            s["pembahasan"] = [dict(e, masuk=True, utuh=True) for e in s["pembahasan"]]
    hasil = [s for s in hasil if s["tipe"] == "blok" or s["_judul"] or
             sum(n for n, ya in s["_tanda"] if ya) >= 0.5 * sum(n for n, _ in s["_tanda"])]
    # bacaan ikut bila ada soal di bawahnya yang masuk
    hasil = [s for k, s in enumerate(hasil) if s["tipe"] != "blok" or k + 1 < len(hasil) and hasil[k + 1]["tipe"] != "blok"]
    for s in teks_bacaan:
        s["blok"] = penyusun.blok_dari([dict(e, masuk=True, utuh=True) for e in s["blok"]]) if s in hasil else []
    for s in hasil:
        if s["tipe"] == "blok":
            continue
        del s["_tanda"], s["_judul"]
        if s["opsi"] and not s["kunci"]:  # kunci hanya ditulis di penutup pembahasan: 'Jadi, ... Jawaban A.'
            akhir = [e.get("polos", "").strip() for e in s["pembahasan"] if e["jenis"] == "paragraf"][-2:]
            kunci = [m.group(1) for m in map(POLA_KUNCI_AKHIR.search, akhir) if m and m.group(1) in s["opsi"]]
            s["kunci"] = kunci[-1] if kunci else ""
        s["soal"] = penyusun.blok_dari(s["soal"])
        if s["opsi"] and s["soal"] and s["soal"][-1]["jenis"] == "subjudul" and                 re.fullmatch(r"pilihan jawaban", s["soal"][-1]["teks"], re.I):
            s["soal"].pop()  # label di atas opsi; opsinya sudah punya kartu sendiri
        s["pembahasan"] = penyusun.blok_dari(saring([
            dict(e, masuk=True, utuh=True) if not s["opsi"] and re.match(r"^jawaban\b", e.get("polos", ""), re.I) else e
            for e in s["pembahasan"]]))
    awal = penyusun.blok_dari(saring([e for e in pengantar if "masuk" in e])) if hasil else []
    if awal:
        hasil.insert(0, {"tipe": "blok", "judul": "Contoh Soal", "blok": awal})
    return hasil


PEMBUKA_TUJUAN = (r"(?i:(?:setelah mempelajari [^,.]*,\s*)?(?:kamu|peserta|pembaca|anda|siswa)\s+(?:juga\s+)?"
                  r"(?:diharapkan\s+)?(?:dapat|mampu|bisa)\s+)")


def _kapital_awal(t):
    return t[:1].upper() + t[1:]


def _soal_sisipan(isi):
    """(isi materi, isi soal): 'Contoh Soal — Metode Substitusi' yang ditulis di tengah Materi dipisah, karena contoh
    soal selalu menjadi deck sendiri."""
    materi, soal, tingkat = [], [], None
    for el in isi:
        t = TINGKAT.get(el["jenis"])
        if t and tingkat and t <= tingkat:
            tingkat = None
        if t and re.match(r"^contoh soal\b", _bersih(el["teks"]), re.I):
            tingkat, el = t, dict(el, jenis="h2")
        (soal if tingkat else materi).append(el)
    return materi, soal


LEWATI_BAGIAN = ("sumber", "referensi", "daftar pustaka", "daftar rujukan", "kata pengantar", "capaian pembelajaran",
                 "pemetaan referensi")


def isi_bagian_kosong(elemen):
    """Tujuan pembelajaran dan penutup yang sama sekali tidak ditandai tetap diisi, supaya slidenya tidak kosong
    (tercatat sebagai tambahan): seluruh butir tujuan dan paragraf pertama penutup."""
    if not any("lingkup" in el for el in elemen):
        return elemen
    for nama, alasan, banyak in (("tujuan pembelajaran", "tujuan pembelajaran kosong", None), ("penutup", "pesan penutup kosong", 1)):
        awal = next((i for i, el in enumerate(elemen) if el["jenis"] == "h1" and _kunci_bagian(el["polos"]) == nama), None)
        if awal is None:
            continue
        akhir = next((i for i in range(awal + 1, len(elemen)) if elemen[i]["jenis"] == "h1"), len(elemen))
        isi = [el for el in elemen[awal + 1:akhir] if el["jenis"] == "paragraf" and el["polos"]]
        if isi and not any(el["masuk"] for el in isi):
            for el in isi[:banyak]:
                el.update(masuk=True, utuh=True, tambahan=alasan)
    return elemen


def susun(elemen, folder_naskah, slug, sumber="", kelompok=0, nama_file=None, subtes=SUBTES_TIU, peta=None,
          judul=None, pilihan=(), minimal=False, isi_kosong=False):
    """peta: peta subtes (bawaan PETA_TIU); judul: judul materi bila dokumen tidak memakai gaya Title;
    pilihan, minimal: lihat pelengkap(); isi_kosong: lihat isi_bagian_kosong()."""
    penyusun = Penyusun(folder_naskah, slug)
    elemen = pelengkap(_rapikan_elemen(tandai(elemen)), pilihan, minimal)
    judul_dok, subjudul, bagian = _bagi_bagian(isi_bagian_kosong(elemen) if isi_kosong else elemen)
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
            if (el["jenis"] in pemecah or el.get("slide_baru")) and el["polos"]:
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
                if gambar[0].get("tambahan"):
                    naskah["pendahuluan"]["tambahan"] = gambar[0]["tambahan"]
                if gambar[0].get("keterangan"):
                    naskah["pendahuluan"]["keterangan"] = gambar[0]["keterangan"]
        elif kunci == "tujuan pembelajaran":
            for el in isi:
                if el["jenis"] != "paragraf" or not el["polos"] or el["polos"].endswith(":"):
                    continue  # kalimat pengantar 'Setelah mempelajari materi ini, peserta mampu:'
                t = re.sub(r"^[•▪●◦\-–]\s+", "", _polos(el["teks"]))
                t = re.sub(r"^\d+[.)]\s+", "", t)
                # tujuan berupa kalimat: 'Setelah mempelajari bab ini, kamu dapat menemukan ... Kamu juga dapat ...'
                if el.get("kko"):  # 'Pola Barisan — mengenali keteraturan ...'
                    naskah["tujuan"]["butir"].append({"level": "", "kko": el["kko"], "teks": t})
                    t = ""
                for butir in re.split(rf"(?<=\.)\s+(?={PEMBUKA_TUJUAN})", t) if t else ():
                    kko, _, sisa = _kapital_awal(re.sub(rf"^{PEMBUKA_TUJUAN}", "", butir)).partition(" ")
                    naskah["tujuan"]["butir"].append({"level": "", "kko": kko, "teks": sisa})
                if el.get("tambahan"):
                    naskah["tujuan"].setdefault("tambahan", []).append(_catatan_tambahan(el))
        elif kunci == "materi":
            isi, sisipan = _soal_sisipan(isi)
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
            soal = _soal(penyusun, sisipan)
            if soal:
                naskah["materi"].append({"subbab": "Contoh Soal", "ikon": "checklist", "slide": soal})
        elif kunci.startswith("contoh soal") or kunci.startswith("latihan soal"):
            soal = _soal(penyusun, isi)
            if soal:
                naskah["materi"].append({"subbab": "Contoh Soal", "ikon": "checklist", "slide": soal})
        elif kunci == "penutup":
            pesan, lain = [], []
            for el in isi:
                if el["jenis"] == "h2" and _kunci_bagian(el["polos"]).startswith(
                        LEWATI_BAGIAN + ("materi inti yang wajib dimuat", "ringkasan bagian yang wajib")):
                    break  # catatan penyusun di ujung penutup, mis. 'PEMETAAN REFERENSI TAMBAHAN DAN CATATAN QC'
                if el["jenis"] == "paragraf" and el["polos"]:
                    pesan.append(el["teks"])
                    if el.get("tambahan"):
                        naskah["penutup"].setdefault("tambahan", []).append(_catatan_tambahan(el))
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
        {"judul": sub["subbab"], "teks": ""} for sub in naskah["materi"] if sub["subbab"] not in ("Contoh Soal", "Rangkuman")][:10]
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
    soal["tujuan"] = {"pengantar": naskah["tujuan"]["pengantar"], "butir": []}
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
