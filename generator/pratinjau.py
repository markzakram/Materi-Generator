"""Ubah deck hasil generator menjadi data bentuk (JSON) untuk pratinjau slide di browser.

Pratinjau tidak merender ulang tata letak: posisi, warna, dan teks diambil dari PPTX yang
sama yang akan diunduh, sehingga yang dilihat di editor sesuai dengan hasil akhirnya.
Didukung: persegi (membulat), oval, panah, bentuk bebas (custGeom -> path SVG), grup,
gambar, tabel, isian warna/gradien/gambar, garis tepi, rotasi, dan latar slide.
"""
from __future__ import annotations

import base64
import hashlib

from pptx.oxml.ns import qn

EMU = 914400
R_EMBED = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed"


def _inci(v):
    return round((v or 0) / EMU, 4)


def _warna(el):
    """(hex, alpha 0-1) dari elemen warna (srgbClr); None bila bukan warna RGB."""
    if el is None or el.tag != qn("a:srgbClr"):
        return None
    a = el.find(qn("a:alpha"))
    return el.get("val"), round(int(a.get("val")) / 100000, 3) if a is not None else 1.0


def _warna_isian(el):
    """Warna solid (hex) dari spPr/tcPr/bgPr/rPr, None bila tanpa isian solid."""
    if el is None:
        return None
    sf = el.find(qn("a:solidFill"))
    if sf is not None and len(sf):
        return sf[0].get("val")
    return None


class _Gambar:
    def __init__(self):
        self.peta = {}

    def kunci(self, blob, jenis):
        k = hashlib.sha1(blob).hexdigest()[:16]
        if k not in self.peta:
            self.peta[k] = f"data:{jenis};base64,{base64.b64encode(blob).decode()}"
        return k


def _isian(pr, part, gambar):
    """spPr/bgPr -> {"t": "solid"|"grad"|"img", ...} atau None."""
    if pr is None:
        return None
    sf = pr.find(qn("a:solidFill"))
    if sf is not None and len(sf):
        w = _warna(sf[0])
        return {"t": "solid", "c": w[0], "a": w[1]} if w else None
    gf = pr.find(qn("a:gradFill"))
    if gf is not None:
        stops = []
        for gs in gf.iter(qn("a:gs")):
            w = _warna(gs[0]) if len(gs) else None
            if w:
                stops.append([int(gs.get("pos", 0)) / 1000, w[0], w[1]])
        lin = gf.find(qn("a:lin"))
        return {"t": "grad", "s": stops, "ang": int(lin.get("ang", 0)) / 60000 if lin is not None else 90}
    bf = pr.find(qn("a:blipFill"))
    if bf is not None:
        blip = bf.find(qn("a:blip"))
        rid = blip.get(R_EMBED) if blip is not None else None
        if rid:
            try:
                p = part.related_part(rid)
                return {"t": "img", "src": gambar.kunci(p.blob, p.content_type)}
            except KeyError:
                return None
    return None


def _garis(spPr):
    ln = spPr.find(qn("a:ln")) if spPr is not None else None
    if ln is None or ln.find(qn("a:noFill")) is not None:
        return None
    sf = ln.find(qn("a:solidFill"))
    w = _warna(sf[0]) if sf is not None and len(sf) else None
    if not w:
        return None
    return {"c": w[0], "w": round(int(ln.get("w", 12700)) / 12700, 2)}


def _path_svg(custGeom):
    """custGeom -> (d, lebar, tinggi) dalam satuan path pertama."""
    paths = custGeom.findall(f"{qn('a:pathLst')}/{qn('a:path')}")
    if not paths:
        return None
    w0, h0 = int(paths[0].get("w", 1)) or 1, int(paths[0].get("h", 1)) or 1
    d = []
    for p in paths:
        sx = w0 / (int(p.get("w", w0)) or w0)
        sy = h0 / (int(p.get("h", h0)) or h0)
        for cmd in p:
            pts = [(int(pt.get("x")) * sx, int(pt.get("y")) * sy) for pt in cmd.findall(qn("a:pt"))]
            tag = cmd.tag.split("}")[1]
            koord = " ".join(f"{x:.0f} {y:.0f}" for x, y in pts)
            if tag == "moveTo":
                d.append("M" + koord)
            elif tag == "lnTo":
                d.append("L" + koord)
            elif tag == "cubicBezTo":
                d.append("C" + koord)
            elif tag == "quadBezTo":
                d.append("Q" + koord)
            elif tag == "close":
                d.append("Z")
    return {"d": " ".join(d), "w": w0, "h": h0}


def _teks(tf_el):
    """txBody -> dict inset, jangkar, paragraf (perataan, spasi, run)."""
    if tf_el is None:
        return None
    body = tf_el.find(qn("a:bodyPr"))
    inset = [int(body.get(k, d)) / EMU for k, d in (("lIns", 91440), ("tIns", 45720), ("rIns", 91440),
                                                      ("bIns", 45720))] if body is not None else [0.1, 0.05, 0.1, 0.05]
    jangkar = body.get("anchor", "t") if body is not None else "t"
    paragraf = []
    for p in tf_el.findall(qn("a:p")):
        ppr = p.find(qn("a:pPr"))
        lnspc = ppr.find(f"{qn('a:lnSpc')}/{qn('a:spcPts')}") if ppr is not None else None
        lnpct = ppr.find(f"{qn('a:lnSpc')}/{qn('a:spcPct')}") if ppr is not None else None
        aft = ppr.find(f"{qn('a:spcAft')}/{qn('a:spcPts')}") if ppr is not None else None
        run = []
        for r in p.findall(qn("a:r")):
            rpr = r.find(qn("a:rPr"))
            t = r.find(qn("a:t"))
            latin = rpr.find(qn("a:latin")) if rpr is not None else None
            run.append({
                "t": t.text if t is not None and t.text else "",
                "pt": int(rpr.get("sz", 1800)) / 100 if rpr is not None else 18,
                "b": rpr is not None and rpr.get("b") in ("1", "true"),
                "i": rpr is not None and rpr.get("i") in ("1", "true"),
                "s": rpr is not None and rpr.get("strike") == "sngStrike",
                "c": _warna_isian(rpr) or "000000",
                "f": latin.get("typeface") if latin is not None else "Poppins",
            })
            naik = int(rpr.get("baseline", 0)) if rpr is not None else 0
            if naik:  # pangkat/indeks rumus
                run[-1]["v"] = "super" if naik > 0 else "sub"
        paragraf.append({
            "al": {"ctr": "c", "r": "r", "just": "j"}.get(ppr.get("algn") if ppr is not None else None, "l"),
            "ls": int(lnspc.get("val")) / 100 if lnspc is not None else None,
            "lp": int(lnpct.get("val")) / 100000 if lnpct is not None else None,
            "sa": int(aft.get("val")) / 100 if aft is not None else 0,
            "r": run,
        })
    return {"in": [round(v, 4) for v in inset], "an": jangkar, "p": paragraf}


def _xfrm(el):
    for jalur in (("p:spPr", "a:xfrm"), ("p:grpSpPr", "a:xfrm"), ("p:xfrm",)):
        x = el.find("/".join(qn(b) for b in jalur))
        if x is not None:
            return x
    return None


class _Transform:
    """Peta koordinat anak grup ke koordinat slide (EMU)."""

    def __init__(self, ox=0, oy=0, sx=1.0, sy=1.0):
        self.ox, self.oy, self.sx, self.sy = ox, oy, sx, sy

    def titik(self, x, y):
        return self.ox + x * self.sx, self.oy + y * self.sy

    def anak(self, xfrm):
        off, ext = xfrm.find(qn("a:off")), xfrm.find(qn("a:ext"))
        choff, chext = xfrm.find(qn("a:chOff")), xfrm.find(qn("a:chExt"))
        gx, gy = self.titik(int(off.get("x")), int(off.get("y")))
        gw, gh = int(ext.get("cx")) * self.sx, int(ext.get("cy")) * self.sy
        cx, cy = (int(choff.get("x")), int(choff.get("y"))) if choff is not None else (0, 0)
        cw = int(chext.get("cx")) if chext is not None and int(chext.get("cx")) else int(ext.get("cx")) or 1
        ch = int(chext.get("cy")) if chext is not None and int(chext.get("cy")) else int(ext.get("cy")) or 1
        sx, sy = gw / cw, gh / ch
        return _Transform(gx - cx * sx, gy - cy * sy, sx, sy)


def _kotak(el, tr):
    x = _xfrm(el)
    if x is None:
        return None
    off, ext = x.find(qn("a:off")), x.find(qn("a:ext"))
    if off is None or ext is None:
        return None
    px, py = tr.titik(int(off.get("x")), int(off.get("y")))
    d = {"x": _inci(px), "y": _inci(py), "w": _inci(int(ext.get("cx")) * tr.sx), "h": _inci(int(ext.get("cy")) * tr.sy)}
    rot = int(x.get("rot", 0)) / 60000
    if rot:
        d["rot"] = round(rot, 2)
    if x.get("flipH") in ("1", "true"):
        d["fh"] = True
    if x.get("flipV") in ("1", "true"):
        d["fv"] = True
    return d


def _bentuk(el, part, gambar, tr, hasil):
    tag = el.tag
    if tag == qn("p:grpSp"):
        x = _xfrm(el)
        anak = tr.anak(x) if x is not None else tr
        for c in el:
            _bentuk(c, part, gambar, anak, hasil)
        return
    if tag not in (qn("p:sp"), qn("p:pic"), qn("p:graphicFrame"), qn("p:cxnSp")):
        return
    d = _kotak(el, tr)
    if d is None:
        return
    if tag == qn("p:pic"):
        blip = el.find(f"{qn('p:blipFill')}/{qn('a:blip')}")
        if blip is not None and blip.get(R_EMBED):
            p = part.related_part(blip.get(R_EMBED))
            d.update(k="img", src=gambar.kunci(p.blob, p.content_type))
            hasil.append(d)
        return
    if tag == qn("p:graphicFrame"):
        tbl = el.find(f".//{qn('a:tbl')}")
        if tbl is None:
            return
        d.update(k="tbl", cw=[_inci(int(c.get("w"))) for c in tbl.find(qn("a:tblGrid"))],
                 rh=[_inci(int(r.get("h"))) for r in tbl.findall(qn("a:tr"))],
                 c=[[{"bg": _warna_isian(tc.find(qn("a:tcPr"))), "tx": _teks(tc.find(qn("a:txBody")))}
                     for tc in tr_.findall(qn("a:tc"))] for tr_ in tbl.findall(qn("a:tr"))])
        hasil.append(d)
        return
    spPr = el.find(qn("p:spPr"))
    d["k"] = "box"
    isian = _isian(spPr, part, gambar)
    if isian:
        d["fill"] = isian
        if isian["t"] == "solid" and isian["a"] == 1:
            d["bg"] = isian["c"]
    garis = _garis(spPr)
    if garis:
        d["ln"] = garis
    prst = spPr.find(qn("a:prstGeom")) if spPr is not None else None
    cust = spPr.find(qn("a:custGeom")) if spPr is not None else None
    if prst is not None:
        jenis = prst.get("prst")
        if jenis == "ellipse":
            d["r"] = "oval"
        elif jenis == "rightArrow":
            d["r"] = "panah"
        elif jenis == "roundRect":
            gd = prst.find(f"{qn('a:avLst')}/{qn('a:gd')}")
            adj = int(gd.get("fmla", "val 16667").split()[-1]) / 100000 if gd is not None else 0.16667
            d["r"] = round(adj * min(d["w"], d["h"]), 4)
    elif cust is not None:
        geo = _path_svg(cust)
        if geo and geo["d"]:
            d["geo"] = geo
    tx = el.find(qn("p:txBody"))
    if tx is not None and "".join(t.text or "" for t in tx.iter(qn("a:t"))).strip():
        d["tx"] = _teks(tx)
    if "fill" in d or "ln" in d or "tx" in d:
        hasil.append(d)


def dump_deck(prs, asal=None):
    """Seluruh deck -> {"lebar", "tinggi", "slide": [...], "gambar": {kunci: dataURI}}."""
    gambar, slide = _Gambar(), []
    for i, s in enumerate(prs.slides):
        bgpr = s._element.find(f"{qn('p:cSld')}/{qn('p:bg')}/{qn('p:bgPr')}")
        isian_bg = _isian(bgpr, s.part, gambar)
        bentuk = []
        for el in s.shapes._spTree.iterchildren():
            _bentuk(el, s.part, gambar, _Transform(), bentuk)
        slide.append({
            "bg": isian_bg["c"] if isian_bg and isian_bg["t"] == "solid" else "FFFFFF",
            "bgimg": isian_bg["src"] if isian_bg and isian_bg["t"] == "img" else None,
            "asal": asal[i] if asal and i < len(asal) else None,
            "b": bentuk,
        })
    return {"lebar": _inci(prs.slide_width), "tinggi": _inci(prs.slide_height), "slide": slide, "gambar": gambar.peta}
