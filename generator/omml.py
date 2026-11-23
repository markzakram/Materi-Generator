"""Rumus Word (OMML) -> teks satu baris dengan markup pangkat ^{..} dan indeks _{..}.

Meratakan rumus dengan menggabung teksnya saja menghilangkan arti: pecahan 280/2 menjadi "2802",
(1+i)^n menjadi "(1+i)n", akar dan kurung hilang. Modul ini membaca strukturnya:

    pecahan          280/2, (a+b)/(c+d), bilangan campuran 2 3/4
    pangkat, indeks  (1+i)^{n}, x_{1}
    akar             √16, √(x+1), ∛8
    kurung, garis atas, sigma, matriks, himpunan persamaan
"""
from __future__ import annotations

import re

from lxml import etree

M = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"

_NAIK = re.compile(r"[\^_]\{[^{}]*\}")
_PASANGAN = {"(": ")", "[": "]", "{": "}"}


def _nama(el):
    return etree.QName(el).localname if isinstance(el.tag, str) and el.tag.startswith(M) else ""


def _atribut(el, *jalur):
    """Nilai m:val dari el/jalur..., None bila tidak ada."""
    for n in jalur:
        el = el.find(M + n) if el is not None else None
    return el.get(M + "val") if el is not None else None


def _terkurung(t):
    """True bila seluruh t berada di dalam satu pasang kurung."""
    t = t.strip()
    if len(t) < 2:
        return False
    if t[0] == "|":
        return t.endswith("|") and t.count("|") == 2
    tutup, dalam = _PASANGAN.get(t[0]), 0
    if not tutup:
        return False
    for i, ch in enumerate(t):
        dalam += (ch == t[0]) - (ch == tutup)
        if dalam == 0:
            return i == len(t) - 1
    return False


def _kurung(t):
    t = t.strip()
    return t if _terkurung(t) else f"({t})"


def _akar_utuh(t):
    return t[:1] in "√∛" and (_terkurung(t[1:]) or re.fullmatch(r"\d+(?:[.,]\d+)?|[^\W\d_]", t[1:]) is not None)


def _pembilang(t):
    """Satu suku (2x, 15, realisasi) ditulis apa adanya; selain itu diberi kurung."""
    t = t.strip()
    inti = _NAIK.sub("", t)
    return t if re.fullmatch(r"[−–-]?[\w.,%°̅]+", inti) or _akar_utuh(inti) else _kurung(t)


def _penyebut(t):
    """Hanya bilangan, satu huruf (boleh berindeks), atau satu kata yang tanpa kurung: 1/2x dan 1/ab rancu."""
    t = t.strip()
    inti = _NAIK.sub("", t)
    aman = re.fullmatch(r"\d[\d.,]*%?|[^\W\d_][₀-₉]*|[^\W\d_]{4,}", inti) or _akar_utuh(inti)
    return t if aman else _kurung(t)


def _datar(t):
    """Pangkat/indeks bersarang tidak didukung markup: x^{2^{n}} -> x^{2^n}."""
    return _NAIK.sub(lambda m: m.group(0)[0] + m.group(0)[2:-1], t)


def _gabung_diakritik(t, tanda):
    return "".join(ch + tanda if not ch.isspace() else ch for ch in t)


def _urut(bagian):
    """Gabungkan [(teks, jenis)] setingkat; pecahan dan akar yang menempel dibuat tidak rancu."""
    out = ""
    bagian = [b for b in bagian if b[0]]
    for i, (t, jenis) in enumerate(bagian):
        kanan = bagian[i + 1][0] if i + 1 < len(bagian) else ""
        rapat_kanan = re.match(r"[\w(\[√∛|]", kanan) is not None
        if jenis == "f":
            angka = re.fullmatch(r"\d[\d.,]*/\d[\d.,]*", t) is not None
            if angka and re.search(r"(?:^|[^\w.,])\d[\d.,]*\s*$", out):  # bilangan campuran: 2 3/4
                out = out.rstrip() + " " + t
            elif rapat_kanan or re.search(r"[\w)\]}|°%]$", out):
                out += _kurung(t)
            else:
                out += t
        elif jenis == "rad" and rapat_kanan and not t.endswith(")"):
            out += re.sub(r"([√∛])(.+)$", r"\1(\2)", t)
        else:
            out += t
    return out


def _isi(el, nama=None):
    """Teks isi el (atau anak bernama m:nama)."""
    if nama is not None:
        el = el.find(M + nama)
    if el is None:
        return ""
    return _urut([_satu(x) for x in el if _nama(x) and not _nama(x).endswith("Pr")])


def _satu(el):
    """(teks, jenis) satu unsur rumus."""
    n = _nama(el)
    if n == "r":
        return "".join(x.text or "" for x in el.iter(M + "t")), "r"
    if n == "t":
        return el.text or "", "r"
    if n == "f":
        atas, bawah = _isi(el, "num"), _isi(el, "den")
        if _atribut(el, "fPr", "type") == "noBar":  # susunan tanpa garis pecahan, mis. koefisien binomial
            return f"({atas.strip()} | {bawah.strip()})", "r"
        return (f"{_pembilang(atas)}/{_penyebut(bawah)}" if bawah.strip() else atas), "f"
    if n in ("sSup", "sSub", "sSubSup", "sPre"):
        dasar = _isi(el, "e")
        if "/" in _NAIK.sub("", dasar) and not _terkurung(dasar):
            dasar = _kurung(dasar)
        bawah, atas = _datar(_isi(el, "sub")).strip(), _datar(_isi(el, "sup")).strip()
        naik = (f"_{{{bawah}}}" if bawah else "") + (f"^{{{atas}}}" if atas else "")
        return (naik + dasar if n == "sPre" else dasar.rstrip() + naik), "r"
    if n == "rad":
        isi, pangkat = _isi(el, "e").strip(), _isi(el, "deg").strip()
        polos = re.fullmatch(r"\d+(?:[.,]\d+)?|[^\W\d_]", isi) is not None
        isi = isi if polos else _kurung(isi)
        if pangkat in ("", "2"):
            return "√" + isi, "rad"
        return ("∛" + isi if pangkat == "3" else f"^{{{_datar(pangkat)}}}√{isi}"), "rad"
    if n == "d":
        pr = el.find(M + "dPr")
        ambil = lambda k, bawaan: bawaan if pr is None or pr.find(M + k) is None else (pr.find(M + k).get(M + "val") or "")  # noqa: E731
        isi = ambil("sepChr", "|").join(_isi(x) for x in el.findall(M + "e"))
        return ambil("begChr", "(") + isi + ambil("endChr", ")"), "r"
    if n == "nary":
        ch = _atribut(el, "naryPr", "chr") or "∫"
        bawah, atas = _datar(_isi(el, "sub")).strip(), _datar(_isi(el, "sup")).strip()
        return ch + (f"_{{{bawah}}}" if bawah else "") + (f"^{{{atas}}}" if atas else "") + " " + _isi(el, "e"), "r"
    if n == "func":
        return _isi(el, "fName").rstrip() + " " + _isi(el, "e"), "r"
    if n in ("limLow", "limUpp"):
        batas = _datar(_isi(el, "lim")).strip()
        return _isi(el, "e").rstrip() + (f"{'_' if n == 'limLow' else '^'}{{{batas}}}" if batas else ""), "r"
    if n == "bar":
        return _gabung_diakritik(_isi(el, "e"), "̅" if _atribut(el, "barPr", "pos") == "top" else "̲"), "r"
    if n == "acc":
        return _gabung_diakritik(_isi(el, "e"), _atribut(el, "accPr", "chr") or "̂"), "r"
    if n == "eqArr":
        return "; ".join(x for x in (_isi(e).strip() for e in el.findall(M + "e")) if x), "r"
    if n == "m":
        return "[" + "; ".join("  ".join(_isi(e).strip() for e in mr.findall(M + "e")) for mr in el.findall(M + "mr")) + "]", "r"
    if n.endswith("Pr"):
        return "", "r"
    return _isi(el), "r"


def teks(el):
    """Teks satu baris untuk m:oMath atau m:oMathPara."""
    if _nama(el) == "oMathPara":
        return "; ".join(x for x in (teks(o).strip() for o in el.findall(M + "oMath")) if x)
    return _isi(el)
