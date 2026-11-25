"""Rapikan dokumen materi UTBK (Cerebrum) menjadi kerangka yang dikenal ekstrak_gdoc.susun().

Dokumen UTBK ditulis beberapa tim dengan kerangka berbeda: subbab bernomor di Heading 1/2/3, contoh soal di tengah
subbab atau di akhir dokumen, soal berupa heading, paragraf tebal, atau daftar bernomor, dan bacaan bersama untuk
beberapa soal. Semuanya disusun ulang menjadi bagian Heading 1 Pendahuluan, Tujuan Pembelajaran, Materi (Heading 2 =
subbab), Contoh Soal (setiap soal ditandai 'awal_soal', bacaan ditandai 'awal_bacaan'), dan Penutup. Teksnya tidak diubah.
"""
from __future__ import annotations

import re

import ekstrak_gdoc as E

LEWATI = ("kata pengantar", "capaian pembelajaran", "peta materi", "keterangan penanda", "rujukan", "daftar pustaka",
          "daftar rujukan", "referensi", "sumber")
SKALA_GAMBAR = 1.5
RANGKUMAN = ("rangkuman", "rangkuman cepat", "ringkasan")
JUDUL = ("h1", "h2", "h3")
LABEL_KURUNG = re.compile(r"\((?=[^)]*\bPPT\b)([^)]*)\)|[{\[](?=[^}\]]*\bPPT\b)([^}\]]*)[}\]]", re.I)
LABEL_SEMUA_SOAL = re.compile(r"semua (?:contoh(?: sol| soal)?|soal)\b", re.I)
POLA_SOAL = re.compile(r"^(contoh soal|contoh|soal)\s*(?:nomor\s*|no\.?\s*)?(\d+)\b\s*(.*)$", re.I)
NOMOR = re.compile(r"^(\d+(?:\.\d+)*)\.?\s+(?=\S)")
NOMOR_BAB_SOAL = re.compile(r"^\d+\.(\d+)\.?\s+(?=\S)")
FORMAT_SOAL = re.compile(r"\s*(?:[—–-]\s*)?\(?(?:short answer|pilihan ganda|PGK)\b[^)]*\)?\s*$|\s*\((?:table|tabel)\)\s*$", re.I)
# label isi pembahasan yang berdiri sendiri di satu baris
LABEL_ISI = re.compile(r"^(jawaban|konsep|konsep yang dipakai|petunjuk dalam bacaan|cara membaca soal|cara cepat|pembahasan|"
                       r"pembahasan benar|analisis (?:pilihan |opsi )?(?:jawaban|opsi|lain)|analisis opsi(?: lain| jawaban)?|"
                       r"analisis pilihan(?: jawaban)?|penjelasan opsi salah|penjelasan setiap pernyataan|"
                       r"mengapa pilihan lain kurang tepat|strategi menjawab)\s*:?$", re.I)


def _label(teks):
    """Label penulis dalam kurung apa pun -> kurung siku, dengan kata yang dikenal tandai()."""
    def ganti(m):
        isi = m.group(1) or m.group(2)
        if "=" in isi:  # keterangan warna: [MERAH = WAJIB PPT]
            return ""
        return "[semua contoh soal wajib ada di PPT]" if LABEL_SEMUA_SOAL.search(isi) else f"[{isi}]"
    return re.sub(r"\s{2,}", " ", LABEL_KURUNG.sub(ganti, teks)).strip()


def _tanpa_label(teks):
    return re.sub(r"\[[^\]]*\bPPT\b[^\]]*\]", "", teks).strip()


SINGKATAN = {"FPB", "KPK", "PPU", "PBM", "PU", "PK", "PM", "UTBK", "SNBT", "SPLDV", "EYD", "KBBI"}


def _kapital(t):
    """'IDE POKOK' -> 'Ide Pokok', 'MAJAS/PERUMPAMAAN' -> 'Majas/Perumpamaan'; singkatan tetap."""
    huruf = re.sub(r"[^A-Za-z]", "", t)
    if len(huruf) < 6 or not huruf.isupper():
        return t

    def kata(m):
        w = m.group()
        return w if w in SINGKATAN else w.lower() if w.lower() in E._KATA_KECIL else w.capitalize()
    t = re.sub(r"[A-Za-z]+", kata, t)
    return t[:1].upper() + t[1:]


def _judul_bersih(el):
    """Judul tanpa nomor ('3.1.2', 'A.', 'b.'), label, dan titik dua."""
    t = re.sub(r"^[A-Za-z][.)]\s+", "", _tanpa_label(el["polos"]))
    return E._bersih(_kapital(re.sub(r"^\d+(?:\.\d+)*[.)]?\s+", "", t))).strip(" :")


def _nama(el):
    """Nama judul huruf kecil."""
    return _judul_bersih(el).strip(" .").lower()


def _bukan_judul(polos):
    """Heading yang sebenarnya kalimat: 'nyatakan $x$ sebagai fungsi $y$;'."""
    if re.match(r"^[a-h][.)]\s", polos):
        return False
    return bool(polos[:1].islower() or polos.endswith((";", ",")) or "$" in polos and len(polos.split()) > 8 or
                len(polos) > 100 or polos.endswith(".") and len(polos.split()) > 12)


def _el(jenis, teks, **k):
    return {"jenis": jenis, "teks": teks, "polos": E._polos(teks), "tebal": True, "miring": False, "daftar": False,
            "sorot": "", "gaya": "", "fmt": "", "num": None, **k}


def _per_baris(el):
    """Paragraf dengan ganti baris di dalamnya -> satu paragraf per baris ('Konsep:' + isinya, opsi A-E)."""
    if el.get("jenis") != "paragraf" or E.GANTI_BARIS not in el["teks"]:
        return [dict(el, **{k: el[k].replace(E.GANTI_BARIS, " ") for k in ("teks", "polos", "sorot") if k in el and isinstance(el[k], str)})]
    sorot = E._polos(el["sorot"])
    hasil = []
    for t in el["teks"].split(E.GANTI_BARIS):
        t = t.strip()
        polos = E._polos(t).strip()
        if polos:
            tebal = el["tebal"] or bool(re.fullmatch(r"\*\*[^*]+\*\*", t))
            hasil.append(dict(el, teks=t, polos=polos, tebal=tebal, sorot=t if polos in sorot else "",
                              daftar=el["daftar"] and not hasil, fmt=el["fmt"] if not hasil else ""))
    return hasil


def _tabel_bergambar(el):
    """Tabel yang selnya berisi gambar rumus atau bagan -> subjudul (kolom pertama), teks, dan gambarnya per baris,
    karena sel tabel di slide hanya memuat teks."""
    if el.get("jenis") != "tabel" or not any(g for r in el["gambar"] for g in r):
        return [el]
    hasil = []
    for r in range(1, len(el["baris"])):
        tanda = el["sorot_baris"][r]
        for c, sel in enumerate(el["baris"][r]):
            if c == 0 and sel.strip():
                hasil.append(_el("h3", sel, sorot=sel if tanda else ""))
            elif sel.strip():
                hasil += [_el("paragraf", t, tebal=False, sorot=t if tanda else "") for t in el["par"][r][c]]
            hasil += [dict(g, sorot=tanda) for g in el["gambar"][r][c]]
    return hasil


def _siapkan(elemen):
    hasil, urutan = [], {}
    for el in (x for e in elemen for t in _tabel_bergambar(e) for x in _per_baris(t)):
        el = dict(el)
        if el["jenis"] == "gambar" and el.get("w") and el.get("h"):
            skala = min(SKALA_GAMBAR, 10.5 / el["w"], 4.8 / el["h"])  # gambar dokumen A4 terlalu kecil di slide 16 inci
            if skala > 1:
                el.update(w=el["w"] * skala, h=el["h"] * skala)
        if "polos" not in el:
            hasil.append(el)
            continue
        for k in ("teks", "polos", "sorot"):
            el[k] = _label(el[k])
        if re.fullmatch(r"\[[^\]]*wajib[^\]]*\]", el["polos"], re.I) and hasil and re.fullmatch(r"\$[^$]+\$", hasil[-1].get("teks", "")):
            hasil[-1]["sorot"] = hasil[-1]["teks"]  # label di bawah rumus (aslinya gambar rumus)
            continue
        el["ketat"] = True  # paragraf tebal tanpa tanda tidak dianggap judul isi bertanda yang jauh di bawahnya
        if el["jenis"] == "paragraf" and el.get("gaya") in ("heading 5", "heading 6"):
            el.update(jenis="h3", daftar=False)
        elif el["jenis"] == "paragraf" and el["daftar"] and el["tebal"] and len(_tanpa_label(el["polos"])) <= 80 and \
                not el["polos"].rstrip().endswith((".", ",", ";")):
            el["daftar"] = False  # butir yang seluruhnya tebal adalah subjudul: 'Strategi Mengerjakan Soal Peluang'
        if el["jenis"] in JUDUL:
            if not el["polos"]:
                continue
            if _bukan_judul(el["polos"]):
                el["jenis"] = "paragraf"
        if el["jenis"] == "paragraf" and el.get("fmt") == "decimal" and el["polos"]:
            # nomor daftar otomatis Word ditulis di teksnya, karena jawaban soal merujuk nomor itu ('1. Benar')
            urutan[el["num"]] = n = urutan.get(el["num"], 0) + 1
            for k in ("teks", "polos", "sorot"):
                el[k] = f"{n}. {el[k]}" if el[k].strip() else el[k]
            el["daftar"] = False
        if el["polos"] or el["jenis"] not in ("paragraf",):
            hasil.append(el)
    return hasil


def _kepala(el):
    """Elemen yang boleh menjadi nama bagian: heading atau paragraf tebal pendek."""
    return el["jenis"] in JUDUL or el["jenis"] == "paragraf" and el["tebal"] and len(el["polos"]) <= 40


def _awal_soal(el):
    """(nomor, sisa judul) bila elemen adalah judul soal 'Contoh Soal 1 — ...' / 'Soal 2', selain itu None."""
    if "polos" not in el or el["jenis"] == "tabel":
        return None
    polos = _tanpa_label(el["polos"])
    m = POLA_SOAL.match(polos)
    if not m or len(polos) > 130:
        return None
    sisa = m.group(3)
    if el["jenis"] not in JUDUL and not el["tebal"] and not re.match(r"^([—–:-].*)?$", sisa):
        return None
    return int(m.group(2)), FORMAT_SOAL.sub("", re.sub(r"^[—–:-]\s*", "", sisa)).strip(), m.group(1).lower()


def rapikan(elemen):
    elemen = _siapkan(elemen)
    bagian = {"pendahuluan": [], "tujuan": [], "materi": [], "penutup": []}
    zona = []  # [{label, semua, bab, awal: [el], soal: [(judul, label, sorot, [el])]}]
    subbab_kini = None
    keadaan, z = "lewati", None
    soal_akhir = _awal_soal_akhir(elemen)
    nomor = [m.group(1) for el in elemen if el["jenis"] in JUDUL and (m := NOMOR.match(el["polos"])) and "." in m.group(1)]
    n2, n3 = sum(x.count(".") == 1 for x in nomor), sum(x.count(".") == 2 for x in nomor)
    tingkat_subbab = 2 if n3 >= 2 and n2 - _jumlah_soal_bab(elemen) <= 2 else 1  # titik pada nomor subbab

    def mulai_zona(label, semua, bab=False, akhir=False):
        nonlocal z, keadaan
        z = {"label": label, "semua": semua, "bab": bab, "akhir": akhir, "awal": [], "soal": []}
        zona.append(z)
        keadaan = "soal"

    def tambah_materi(el):
        bagian["materi"].append(el)

    def subbab(nama, el):
        nonlocal subbab_kini, keadaan, z
        subbab_kini, keadaan, z = nama, "materi", None
        tambah_materi(dict(el, jenis="h2", teks=nama, polos=nama, sorot=""))

    i = -1
    while (i := i + 1) < len(elemen):
        el = elemen[i]
        if i == soal_akhir:
            mulai_zona(None, "tanpa tanda", akhir=True)
        if "polos" not in el:
            if keadaan == "soal":
                (z["soal"][-1][3] if z["soal"] else z["awal"]).append(el)
            elif keadaan in ("rangkuman penutup",):
                tambah_materi(el)
            elif keadaan in bagian:
                bagian[keadaan].append(el)
            continue
        j, polos = el["jenis"], el["polos"]
        nm = _nama(el)
        di_akhir = keadaan == "soal" and z["akhir"]
        # ---- nama bagian ----
        if _kepala(el) and not di_akhir:
            baru = ("lewati" if nm in LEWATI else "pendahuluan" if nm == "pendahuluan" else
                    "tujuan" if nm in ("tujuan", "tujuan pembelajaran") else
                    "materi" if nm in ("materi", "materi pembelajaran") else "penutup" if nm == "penutup" else None)
            if baru:
                keadaan, z = baru, None
                if baru == "materi":
                    subbab_kini = None
                continue
            if nm in RANGKUMAN and (keadaan != "soal" or j in JUDUL):
                if keadaan == "penutup":
                    keadaan = "rangkuman penutup"
                    tambah_materi(_el("h2", "Rangkuman"))
                else:
                    subbab("Rangkuman", el)
                continue
        # ---- awal zona contoh soal ----
        label_semua = bool(E.LABEL_SEMUA.search(polos))
        if label_semua and not _tanpa_label(polos) or label_semua and j == "paragraf" and LABEL_SEMUA_AWAL.match(polos) and nm not in ("contoh", "contoh soal"):
            mulai_zona(None, True, akhir=True)
            sisa = _tanpa_label(polos)
            if not sisa:
                continue
            el = dict(el, teks=_tanpa_label(el["teks"]), polos=sisa, sorot="")
            polos = sisa
        elif re.fullmatch(r"contoh(?: soal)?(?: dan pembahasan)?", nm) and (j in JUDUL or label_semua) and (nm != "contoh" or label_semua):
            bab = j == "h1"
            m = NOMOR.match(polos)
            sejajar = m and m.group(1).count(".") == tingkat_subbab  # zona untuk seluruh bab, bukan subbab terakhir
            mulai_zona(None if bab or sejajar else subbab_kini, label_semua, bab=bab)
            continue
        awal = _awal_soal(el)
        if awal and keadaan != "soal" and j in JUDUL and keadaan != "lewati":  # 'Contoh Soal 3 — ...' langsung di subbab
            mulai_zona(subbab_kini, False)
        # ---- isi zona contoh soal ----
        if keadaan == "soal":
            judul = None
            if awal:
                judul = (f"{'Soal' if awal[2] == 'soal' else 'Contoh Soal'} {awal[0]}" + (f" — {awal[1]}" if awal[1] else ""), z["label"])
            elif z["bab"] and (m := NOMOR_BAB_SOAL.match(polos)) and (j in JUDUL or el["tebal"]):
                bab_soal, _, nama = _tanpa_label(polos[m.end():]).rpartition(" — ")
                judul = (f"Contoh Soal {m.group(1)} — {nama}", bab_soal or None)
            elif j == "paragraf" and el.get("fmt") == "decimal" and _soal_bernomor(elemen, i, z):
                judul = (f"Contoh Soal {len(z['soal']) + 1} — {_judul_bersih(el)}", z["label"])
            if judul:
                z["soal"].append((*judul, el["sorot"], []))
                continue
            if j in ("h1", "h2") and not z["akhir"] or j == "h3" and NOMOR.match(polos) and "." in NOMOR.match(polos).group(1) \
                    and not z["akhir"] and not z["bab"]:
                keadaan, z = "materi", None  # judul subbab berikutnya: teruskan ke bawah
            else:
                m = re.match(r"^\[[^\]]*wajib[^\]]*\]\s*(?=\S)", polos, re.I)
                if m and not z["akhir"] and not z["bab"] and j == "paragraf":  # catatan inti di sela contoh soal: milik materi
                    isi = bagian["materi"]
                    tempat = len(isi)  # sesudah isi bertanda terakhir, bukan di bawah judul yang isinya tidak ditandai
                    while tempat > 1 and isi[tempat - 1]["jenis"] != "h2" and not (
                            isi[tempat - 1]["jenis"] == "paragraf" and isi[tempat - 1]["sorot"].strip() or any(isi[tempat - 1].get("sorot_baris", []))):
                        tempat -= 1
                    isi.insert(tempat, el)
                    continue
                if j in JUDUL:
                    el = dict(el, jenis="paragraf", tebal=not z["akhir"] or el["tebal"])
                (z["soal"][-1][3] if z["soal"] else z["awal"]).append(el)
                continue
        # ---- isi bagian lain ----
        if j in JUDUL:
            m = NOMOR.match(polos)
            dalam = m.group(1).count(".") if m else None
            nama = _judul_bersih(el)
            if j == "h1" and dalam == 0:  # '1. MAKNA KATA DAN FRASA' judul bab; '2. STRATEGI TERPADU ...' subbab
                keadaan, z = "materi", None
                if m.group(1) == "1":
                    subbab_kini = None
                else:
                    subbab(nama, el)
            elif j == "h2" and dalam is None or j == "h2" and dalam == 0:
                subbab(nama, el)
            elif dalam is not None and 1 <= dalam <= tingkat_subbab:
                subbab(nama, el)
            elif dalam == 2 and tingkat_subbab == 1:
                keadaan = "materi"
                tambah_materi(dict(el, jenis="h3", teks=nama, polos=nama, slide_baru=True))
            elif keadaan == "penutup":
                bagian["penutup"].append(dict(el, jenis="paragraf"))
            elif keadaan in ("pendahuluan", "tujuan", "lewati"):
                if keadaan != "lewati":
                    bagian[keadaan].append(dict(el, jenis="h3"))
            else:
                keadaan = "materi"
                tambah_materi(dict(el, jenis="h3", teks=nama, polos=nama))
            continue
        if keadaan == "rangkuman penutup":
            bagian_rangkuman = bagian["materi"]
            sisa = [x for x in elemen[i:_batas(elemen, i)]]
            ada_butir = any(x["jenis"] == "paragraf" and E._butir(x) for x in sisa)
            (bagian_rangkuman if ada_butir or E._butir(el) else bagian["penutup"]).append(el)
        elif keadaan in bagian:
            bagian[keadaan].append(el)
    return _susun_ulang(bagian, zona)


LABEL_SEMUA_AWAL = re.compile(r"^\[[^\]]*\bsemua\b[^\]]*\]", re.I)


def _batas(elemen, i):
    """Indeks heading berikutnya (akhir bagian)."""
    return next((k for k in range(i, len(elemen)) if elemen[k]["jenis"] in JUDUL), len(elemen))


def _jumlah_soal_bab(elemen):
    """Banyak heading '4.1 ...' di bawah Heading 1 'CONTOH SOAL' (bukan subbab)."""
    n, di_soal = 0, False
    for el in elemen:
        if el["jenis"] == "h1":
            di_soal = _nama(el).startswith("contoh soal")
        elif di_soal and el["jenis"] in JUDUL and NOMOR_BAB_SOAL.match(el["polos"]):
            n += 1
    return n


def _awal_soal_akhir(elemen):
    """Dokumen tanpa judul 'Contoh Soal' yang menaruh 'Soal 1' di ujung, sesudah Penutup (LBI): indeks awal bacaannya."""
    if any(el.get("polos") and (re.match(r"^contoh soal", _nama(el)) and el["jenis"] in JUDUL or
                                E.LABEL_SEMUA.search(el["polos"])) for el in elemen):
        return None
    penutup = [i for i, el in enumerate(elemen) if el.get("polos") and _kepala(el) and _nama(el) == "penutup"]
    if not penutup:
        return None
    soal = next((i for i in range(penutup[-1], len(elemen)) if _awal_soal(elemen[i])), None)
    if soal is None:
        return None
    isi = [i for i in range(penutup[-1] + 1, soal) if elemen[i].get("polos")]
    tanpa_tanda = [i for i in isi if not elemen[i]["sorot"].strip()]
    return tanpa_tanda[0] if tanpa_tanda else soal


def _soal_bernomor(elemen, i, z):
    """Butir bernomor yang memulai soal: pendek, diikuti isi bukan daftar, dan soal sebelumnya sudah dijawab."""
    el = elemen[i]
    if len(el["polos"]) > 110 or el["polos"].endswith((".", ";", ",")):
        return False
    if z["soal"] and not any(re.match(r"^jawaban\b", x.get("polos", ""), re.I) for x in z["soal"][-1][3]):
        return False
    for lain in elemen[i + 1:i + 40]:
        if lain["jenis"] in JUDUL or lain.get("fmt") == "decimal":
            return False
        if re.match(r"^jawaban\b", lain.get("polos", ""), re.I):
            return True
    return False


# ------------------------------------------------------------- isi soal ---
def _pecah_opsi(el):
    """'Pertanyaan .... A. 1 : 2 B. 2 : 3 C. ...' dalam satu paragraf -> pertanyaan dan opsi terpisah."""
    t = el["teks"]
    tanpa_titik = re.match(r"^A\s+(.+?)\s+B\s+(.+?)\s+C\s+(.+?)\s+D\s+(.+?)(?:\s+E\s+(.+))?$", t)
    if tanpa_titik:  # 'A vitamin : sayur B kalsium : tulang C ...'
        t = " ".join(f"{h}. {x}" for h, x in zip("ABCDE", tanpa_titik.groups()) if x)
    m = [x for x in re.finditer(r"(?:(?<=\s)|^)([A-E])[.)]\s+", t)]
    huruf = "".join(x.group(1) for x in m)
    k = huruf.find("ABCD")
    if k < 0 or len(m) == 1:
        return [el]
    m = m[k:k + (5 if huruf[k:k + 5] == "ABCDE" else 4)]
    batas = [x.start() for x in m] + [len(t)]
    potong = [t[:batas[0]]] + [t[a:b] for a, b in zip(batas, batas[1:])]
    hasil = []
    for p in potong:
        p = p.strip()
        if p.replace("*", ""):
            hasil.append(dict(el, teks=p, polos=E._polos(p), sorot=p if el["sorot"].strip() else "", daftar=False))
    return hasil


def _pecah_analisis(el):
    """'A. Salah, ... B. Salah, ... C. Benar, ...' dalam satu paragraf pembahasan -> satu paragraf per opsi."""
    bagian = [x.strip() for x in re.split(r"(?<=[.;\s])(?=[B-E]\.\s+\S)", el["teks"])]
    if not re.match(r"^A\.\s", el["polos"]) or [x[:1] for x in bagian] not in (list("ABCD"), list("ABCDE")):
        return [el]
    return [dict(el, teks=x, polos=E._polos(x), sorot=x if el["sorot"].strip() else "", daftar=False) for x in bagian]


def _isi_soal(isi):
    hasil = []
    for el in isi:
        if el["jenis"] != "paragraf":
            hasil.append(el)
            continue
        polos = el["polos"]
        if re.match(r"^kunci(?: jawaban)?\s*:", polos, re.I):  # 'Kunci: Benar – Benar – Salah'
            el = dict(el, **{k: re.sub(r"^((?:\*\*|<[isr]>)*)kunci(?: jawaban)?", r"\1Jawaban", el[k], flags=re.I)
                             for k in ("teks", "polos", "sorot")})
            polos = el["polos"]
        elif re.match(r"^jawaban model\b", polos, re.I):
            el = dict(el, **{k: re.sub(r"jawaban model", "Jawaban", el[k], flags=re.I) for k in ("teks", "polos", "sorot")})
        if re.match(r"^bentuk soal\s*:", polos, re.I):  # keterangan format untuk penyusun
            continue
        m = re.match(r"^pembahasan (opsi lain|benar)\s*:?$", polos, re.I)
        if m:
            t = "Opsi lain" if m.group(1).lower() == "opsi lain" else "Alasan jawaban benar"
            el = dict(el, teks=t, polos=t, sorot=t if el["sorot"].strip() else "", tebal=True)
        elif re.match(r"^([A-E])\.\s+\1\s+(?:salah|benar|tepat)\b", polos):  # 'A. A salah karena ...'
            el = dict(el, **{k: re.sub(r"^((?:\*\*|<[isr]>)*)([A-E])\.\s*(?:\*\*)?\s*(?=\2\s)", r"\1", el[k]) for k in ("teks", "polos", "sorot")})
        elif re.match(r"^jawaban\s*:\s*\S", polos, re.I) and len(polos) < 60:
            el = dict(el, tebal=True, daftar=False)
        elif LABEL_ISI.match(polos) and len(polos) < 60:
            el = dict(el, tebal=True, daftar=False)
        hasil.append(el)
    jawab = next((i for i, el in enumerate(hasil) if re.match(r"^jawaban\b", el.get("polos", ""), re.I)), None)
    batas = jawab if jawab is not None else len(hasil)
    # opsi 'A sepatu bola', 'B lapangan ...' tanpa titik
    i = 0
    while i + 3 < batas:
        grup = hasil[i:min(i + 5, batas)]
        huruf = "".join(g["polos"][:1] if g["jenis"] == "paragraf" and re.match(r"^[A-E]\s+\S", g["polos"]) else "-" for g in grup)
        n = 5 if huruf == "ABCDE" else 4 if huruf.startswith("ABCD") else 0
        if n:
            for k in range(i, i + n):
                hasil[k] = dict(hasil[k], **{c: re.sub(r"^((?:\*\*|<[isr]>)*)([A-E])(\*\*)?\s+", r"\1\2.\3 ", hasil[k][c]) for c in ("teks", "polos", "sorot")})
            break
        i += 1
    soal = [x for el in hasil[:batas] for x in (_pecah_opsi(el) if el["jenis"] == "paragraf" else [el])]
    sisa = [x for el in hasil[batas:] for x in (_pecah_analisis(el) if el["jenis"] == "paragraf" else [el])]
    # kalimat tanya yang dicetak tebal tetap bagian soal, bukan subjudul
    soal = [dict(el, tebal=False) if el["jenis"] == "paragraf" and el["tebal"] and re.search(r"([?…]|\.\.\.\.?|:)$", el["polos"])
            else el for el in soal]
    # 'Konsep: ...' yang ditulis sebelum 'Jawaban: ...'
    k = next((i for i, el in enumerate(soal) if re.match(r"^konsep\b", el.get("polos", ""), re.I)), None)
    if k is not None and sisa:
        sisa = [sisa[0]] + soal[k:] + sisa[1:]
        soal = soal[:k]
    return soal + sisa


def _teks_untuk(awal):
    """Judul slide bacaan dan isinya: 'Bacalah teks berikut untuk menjawab soal nomor 1–3!' -> 'Teks untuk Soal 1–3'."""
    judul, isi = "Teks Bacaan", list(awal)
    for i, el in enumerate(awal[:2]):
        m = re.match(r"^baca(?:lah)?\b.*?(?:soal nomor|nomor|soal)\s*(\d+)\s*(?:[–—-]|sampai|dan)\s*(\d+)", el.get("polos", ""), re.I)
        if m:
            judul = f"Teks untuk Soal {m.group(1)}–{m.group(2)}"
            isi = awal[:i] + awal[i + 1:]
            break
    return judul, isi


def _tujuan(isi):
    """Butir 'Pola Barisan — mengenali ...' dan tabel IPK -> paragraf dengan nama subbab sebagai pembuka (kko)."""
    hasil = []
    for el in isi:
        if el["jenis"] == "tabel":
            for baris, tanda in list(zip(el["baris"], el["sorot_baris"]))[1:]:
                sel = [E._polos(c).strip() for c in baris]
                for n in range(1, len(sel)):  # 'IPK 1: Memahami' | 'Populasi, sampel, ...' -> 'memahami populasi, sampel, ...'
                    kerja = re.search(r"[:—-] *(Me[a-z]+) *$", E._polos(el["baris"][0][n]))
                    if kerja and sel[n] and not re.match(r"Me[a-z]+ ", sel[n]):
                        sel[n] = kerja.group(1) + " " + sel[n][:1].lower() + sel[n][1:]
                    sel[n] = sel[n][:1].lower() + sel[n][1:] if re.match(r"Me[a-z]+ ", sel[n]) else sel[n]
                teks = " ".join(c.rstrip(".; ") + (";" if n < len(sel) - 2 else ".") for n, c in enumerate(sel[1:]) if c)
                if sel and sel[0] and teks:
                    hasil.append(_el("paragraf", teks[:1].lower() + teks[1:], kko=f"{sel[0]}:", tebal=False, sorot=teks if tanda else ""))
            continue
        if el["jenis"] == "paragraf":
            m = re.match(r"^(?:\d+[.)]\s+)?([^—]{4,60}?)\s+—\s+(\S.*)$", el["polos"])
            if not m and not el["polos"].endswith(":"):
                m = re.match(r"^(?:[•▪●]\s*)?([A-Z][^:.;—]{3,44}):\s+(\S.*)$", el["polos"])  # 'Persamaan Linear: menguasai ...'
            if m:
                sorot = E._polos(el["sorot"]).strip()
                sisa = re.match(r"^(?:\d+[.)]\s+)?(?:[•▪●]\s*)?" + re.escape(m.group(1)) + r"\s*[—:]\s*(\S.*)$", sorot)
                sorot = sisa.group(1) if sisa else sorot
                el = dict(el, teks=m.group(2)[:1].lower() + m.group(2)[1:], kko=f"{m.group(1)}:", sorot=sorot[:1].lower() + sorot[1:])
        hasil.append(el)
    return hasil


def _susun_ulang(bagian, zona):
    out = []
    for nama, judul in (("pendahuluan", "Pendahuluan"), ("tujuan", "Tujuan Pembelajaran"), ("materi", "Materi")):
        isi = _tujuan(bagian[nama]) if nama == "tujuan" else bagian[nama]
        if isi:
            out += [_el("h1", judul)] + isi
    ada = [z for z in zona if z["soal"]]
    if ada:
        out.append(_el("h1", "Contoh Soal"))
        for z in ada:
            semua = " [semua contoh soal wajib ada di PPT]" if z["semua"] is True else ""
            # Heading 2 membatasi label 'semua contoh soal' pada zonanya
            out.append(_el("h2", "Contoh Soal" + semua, zona=True))
            if z["akhir"] and z["awal"]:
                judul, isi = _teks_untuk(z["awal"])
                out += [_el("h3", judul, awal_bacaan=True)] + isi
            else:
                out += z["awal"]
            for judul, label, sorot, isi in z["soal"]:
                isi = _isi_soal(isi)
                if z["semua"] == "tanpa tanda" and isi:
                    # soal yang sama sekali tidak ditandai penulis: tetap masuk, tercatat di catatan pembicara
                    isi = [dict(el, paksa=True) for el in isi]
                    isi[0]["tambahan"] = "contoh soal tanpa tanda"
                out += [_el("h3", judul, awal_soal=True, label=label, sorot=sorot)] + isi
    if bagian["penutup"]:
        out += [_el("h1", "Penutup")] + bagian["penutup"]
    return out
