"""Ubah deck hasil generator menjadi data bentuk (JSON) untuk pratinjau slide di browser.

Pratinjau tidak merender ulang tata letak: posisi, warna, dan teks diambil dari PPTX yang
sama yang akan diunduh, sehingga yang dilihat di editor sesuai dengan hasil akhirnya.
Hanya bentuk yang dipakai generator yang didukung: persegi (membulat), oval, panah,
kotak teks, gambar, dan tabel.
"""
from __future__ import annotations

import base64
import hashlib

from pptx.enum.shapes import MSO_SHAPE, MSO_SHAPE_TYPE
from pptx.oxml.ns import qn

EMU = 914400


def _inci(v):
    return round((v or 0) / EMU, 4)


def _warna_isian(el):
    """Warna solid (hex) dari spPr/tcPr/bgPr, None bila tanpa isian."""
    if el is None:
        return None
    sf = el.find(qn("a:solidFill"))
    if sf is not None and len(sf):
        return sf[0].get("val")
    return None


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
        aft = ppr.find(f"{qn('a:spcAft')}/{qn('a:spcPts')}") if ppr is not None else None
        run = []
        for r in p.findall(qn("a:r")):
            rpr = r.find(qn("a:rPr"))
            t = r.find(qn("a:t"))
            latin = rpr.find(qn("a:latin")) if rpr is not None else None
            run.append({
                "t": t.text if t is not None and t.text else "",
                "pt": int(rpr.get("sz", 1800)) / 100 if rpr is not None else 18,
                "b": rpr is not None and rpr.get("b") == "1",
                "c": _warna_isian(rpr) or "000000",
                "f": latin.get("typeface") if latin is not None else "Poppins",
            })
        paragraf.append({
            "al": {"ctr": "c", "r": "r"}.get(ppr.get("algn") if ppr is not None else None, "l"),
            "ls": int(lnspc.get("val")) / 100 if lnspc is not None else None,
            "sa": int(aft.get("val")) / 100 if aft is not None else 0,
            "r": run,
        })
    return {"in": [round(v, 4) for v in inset], "an": jangkar, "p": paragraf}


def _gambar(shp, gambar):
    img = shp.image
    kunci = hashlib.sha1(img.blob).hexdigest()[:16]
    if kunci not in gambar:
        gambar[kunci] = f"data:{img.content_type};base64,{base64.b64encode(img.blob).decode()}"
    return kunci


def _bentuk(shp, gambar):
    d = {"x": _inci(shp.left), "y": _inci(shp.top), "w": _inci(shp.width), "h": _inci(shp.height)}
    if shp.shape_type == MSO_SHAPE_TYPE.PICTURE:
        d.update(k="img", src=_gambar(shp, gambar))
        return d
    if shp.has_table:
        tbl = shp.table
        d.update(k="tbl", cw=[_inci(c.width) for c in tbl.columns], rh=[_inci(r.height) for r in tbl.rows],
                 c=[[{"bg": _warna_isian(cell._tc.find(qn("a:tcPr"))),
                      "tx": _teks(cell._tc.find(qn("a:txBody")))} for cell in row.cells] for row in tbl.rows])
        return d
    sp = shp._element
    spPr = sp.find(qn("p:spPr"))
    d["k"] = "box"
    d["bg"] = _warna_isian(spPr)
    if shp.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE:
        jenis = shp.auto_shape_type
        if jenis == MSO_SHAPE.OVAL:
            d["r"] = "oval"
        elif jenis == MSO_SHAPE.RIGHT_ARROW:
            d["r"] = "panah"
        elif jenis == MSO_SHAPE.ROUNDED_RECTANGLE:
            d["r"] = round(shp.adjustments[0] * min(d["w"], d["h"]), 4)
    if shp.has_text_frame and shp.text_frame.text.strip():
        d["tx"] = _teks(sp.find(qn("p:txBody")))
    return d


def dump_deck(prs, asal=None):
    """Seluruh deck -> {"lebar", "tinggi", "slide": [...], "gambar": {kunci: dataURI}}."""
    gambar, slide = {}, []
    for i, s in enumerate(prs.slides):
        bg = s._element.find(f"{qn('p:cSld')}/{qn('p:bg')}/{qn('p:bgPr')}")
        slide.append({
            "bg": _warna_isian(bg) or "FFFFFF",
            "asal": asal[i] if asal and i < len(asal) else None,
            "b": [_bentuk(shp, gambar) for shp in s.shapes],
        })
    return {"lebar": _inci(prs.slide_width), "tinggi": _inci(prs.slide_height), "slide": slide, "gambar": gambar}
