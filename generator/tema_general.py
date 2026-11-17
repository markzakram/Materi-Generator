"""Tema "General": tampilan mengikuti 'Template Materi General.pptx'.

Slide sampul, pendahuluan, tujuan, penutup, dan terima kasih disalin dari slide contoh di
template (foto, ilustrasi, bingkai, dan hiasan tetap persis), lalu teks contohnya diganti
isi naskah. Slide isi memakai chip dan hiasan samar dari template; komponen isinya (tabel,
kartu, catatan, soal) memakai tata letak generator dengan token warna tema General.
Navigasi digambar ulang agar bagian aktifnya benar dan tetap bisa diklik.

Slide contoh dikenali dari teks besarnya (MATERI BESAR, PENDAHULUAN, TUJUAN PEMBELAJARAN,
PENUTUP, TERIMA KASIH) dan dari slide isi yang berisi tabel, bukan dari nomor urutnya.
"""
from __future__ import annotations

import copy
import io

from PIL import Image
from pptx.opc.packuri import PackURI
from pptx.oxml.ns import qn

import desain as D
from desain import P

R_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
EMU = 914400
NAVBAR = (2.22, 10.04, 11.56, 0.65)  # x, y, w, h seperti di template
BATAS_HALAMAN = 9.3  # bawah halaman putih pada slide berbingkai


class TemplateGeneralTidakCocok(ValueError):
    pass


# ------------------------------------------------------------ bantuan ---
def _semua_bentuk(shapes):
    for sh in shapes:
        yield sh
        if sh.shape_type == 6:  # grup
            yield from _semua_bentuk(sh.shapes)


def _teks_besar(slide, minimal=40):
    """Teks (huruf besar) dalam bentuk berukuran >= minimal pt."""
    hasil = []
    for sh in _semua_bentuk(slide.shapes):
        if not (sh.has_text_frame and sh.text_frame.text.strip()):
            continue
        ukuran = [r.font.size.pt for p in sh.text_frame.paragraphs for r in p.runs if r.font.size]
        if ukuran and max(ukuran) >= minimal:
            hasil.append(" ".join(sh.text_frame.text.split()).upper())
    return hasil


def _teks_dalam(el):
    return " ".join(t.text or "" for t in el.iter(qn("a:t")))


def _navbar_template(el):
    """Navigasi bawaan template: grupnya (berisi label PENDALAMAN MATERI) atau elemen lepas
    di pita navigasi bawah (slide Penutup menaruh label dan sorotan tanpa grup)."""
    if el.tag == qn("p:grpSp") and "PENDALAMAN MATERI" in _teks_dalam(el).upper():
        return True
    return _posisi(el)[1] >= 9.95


def _isian_gambar(el):
    sp = el.find(qn("p:spPr"))
    return sp is not None and sp.find(qn("a:blipFill")) is not None


def _xfrm(el):
    for jalur in ("p:spPr/a:xfrm", "p:grpSpPr/a:xfrm", "p:xfrm"):
        x = el.find("/".join(qn(b) for b in jalur.split("/")))
        if x is not None:
            return x
    return None


def _posisi(el):
    x = _xfrm(el)
    if x is None:
        return 0, 0, 0, 0
    off, ext = x.find(qn("a:off")), x.find(qn("a:ext"))
    return (int(off.get("x")) / EMU, int(off.get("y")) / EMU, int(ext.get("cx")) / EMU, int(ext.get("cy")) / EMU)


def _perkecil(blob, jenis):
    """Gambar besar -> lebar maks 1600 px (PNG dikuantisasi 96 warna, JPEG kualitas 82)."""
    with Image.open(io.BytesIO(blob)) as im:
        im.load()
        if im.width > 1600:
            im = im.resize((1600, round(im.height * 1600 / im.width)), Image.LANCZOS)
        hasil = io.BytesIO()
        if jenis == "image/png":
            if im.mode not in ("RGBA", "RGB"):
                im = im.convert("RGBA")
            im.quantize(colors=96, method=Image.Quantize.FASTOCTREE).save(hasil, "PNG", optimize=True)
        else:
            im.convert("RGB").save(hasil, "JPEG", quality=82, optimize=True, progressive=True)
    return hasil.getvalue() if len(hasil.getvalue()) < len(blob) else blob


class TemaGeneral:
    nama = "general"

    def __init__(self, slides):
        self.sumber = self._kenali(slides)
        self._siapkan_media(slides)

    @staticmethod
    def _siapkan_media(slides):
        """Beri nama unik pada gambar template dan perkecil yang besar.

        Setelah slide contoh dilepas, python-pptx tidak lagi melihat gambar template dan bisa
        memakai nama yang sama (mis. image10.png) untuk gambar baru, sehingga PPTX berisi nama
        ganda. Foto dan hiasan besar (hingga 3,3 MB) juga diperkecil agar deck tetap di bawah
        batas unduhan Vercel 4,5 MB.
        """
        terlihat = set()
        for s in slides:
            for rel in s.part.rels.values():
                if rel.is_external:
                    continue
                part = rel.target_part
                if not str(part.partname).startswith("/ppt/media/") or id(part) in terlihat:
                    continue
                terlihat.add(id(part))
                part.partname = PackURI(f"/ppt/media/general{len(terlihat):02d}.{part.partname.ext}")
                if len(part.blob) > 300_000 and part.content_type in ("image/png", "image/jpeg"):
                    part._blob = _perkecil(part.blob, part.content_type)

    # ------------------------------------------------------- slide contoh
    @staticmethod
    def _kenali(slides):
        hasil = {}
        for s in slides:
            besar = _teks_besar(s)
            gabung = " ".join(besar)
            if "MATERI BESAR" in gabung and "BOLEH PANJANG" not in " ".join(
                    _teks_besar(s, 0)).upper().replace("(", " ").replace(")", " "):
                hasil.setdefault("sampul", s)
            elif "TERIMA KASIH" in gabung:
                hasil["terima_kasih"] = s  # varian terakhir (lingkaran navy-kuning)
            elif "PENDAHULUAN" in besar:
                hasil.setdefault("pendahuluan", s)
            elif "PENUTUP" in besar:
                hasil.setdefault("penutup", s)
            elif "TUJUAN PEMBELAJARAN" in besar:
                hasil.setdefault("tujuan", s)
            elif any(sh.has_table for sh in s.shapes) and "isi" not in hasil:
                hasil["isi"] = s
        kurang = [k for k in ("sampul", "pendahuluan", "tujuan", "isi", "penutup", "terima_kasih") if k not in hasil]
        if kurang:
            raise TemplateGeneralTidakCocok(
                "Template General tidak memuat slide contoh: " + ", ".join(kurang)
                + ". Pakai 'Template Materi General.pptx' yang asli.")
        return hasil

    # ------------------------------------------------------------ salin
    def salin(self, sumber, slide, el, belakang=True, xfrm=None):
        """Salin satu elemen (bentuk/grup/gambar) dari slide contoh ke slide; relasi gambar dipetakan."""
        baru = copy.deepcopy(el)
        for node in list(baru.iter()):
            for attr, rid in list(node.attrib.items()):
                if not attr.startswith(R_NS):
                    continue
                rel = sumber.part.rels[rid]
                if rel.reltype.endswith("/slide") or rel.reltype.endswith("/notesSlide"):
                    induk = node.getparent()  # tautan ke slide contoh dibuang
                    if induk is not None:
                        induk.remove(node)
                    break
                if rel.is_external:
                    node.set(attr, slide.part.relate_to(rel.target_ref, rel.reltype, is_external=True))
                else:
                    node.set(attr, slide.part.relate_to(rel.target_part, rel.reltype))
        for f in baru.iter(qn("a:latin"), qn("a:ea"), qn("a:cs"), qn("a:sym")):
            if f.get("typeface") == "Poppins Bold":  # keluarga 'Poppins Bold' -> Poppins tebal
                f.set("typeface", "Poppins")
                rpr = f.getparent()
                if rpr is not None and rpr.tag in (qn("a:rPr"), qn("a:defRPr"), qn("a:endParaRPr")):
                    rpr.set("b", "1")
        if xfrm is not None:
            x = _xfrm(baru)
            x.find(qn("a:off")).set("x", str(int(xfrm[0] * EMU)))
            x.find(qn("a:off")).set("y", str(int(xfrm[1] * EMU)))
            x.find(qn("a:ext")).set("cx", str(int(xfrm[2] * EMU)))
            x.find(qn("a:ext")).set("cy", str(int(xfrm[3] * EMU)))
        pohon = slide.shapes._spTree
        if belakang:
            pohon.insert(2, baru)  # di belakang bentuk lain (setelah nvGrpSpPr dan grpSpPr)
        else:
            pohon.append(baru)
        return baru

    def salin_slide(self, jenis, slide, kecuali):
        """Salin latar dan semua bentuk slide contoh, kecuali yang cocok dengan kecuali(el)."""
        sumber = self.sumber[jenis]
        bg = sumber._element.find(f"{qn('p:cSld')}/{qn('p:bg')}")
        if bg is not None:
            cSld = slide._element.find(qn("p:cSld"))
            lama = cSld.find(qn("p:bg"))
            if lama is not None:
                cSld.remove(lama)
            cSld.insert(0, copy.deepcopy(bg))
        elemen = [el for el in sumber.shapes._spTree.iterchildren()
                  if el.tag in (qn("p:sp"), qn("p:grpSp"), qn("p:pic"), qn("p:graphicFrame"), qn("p:cxnSp"))]
        for el in reversed(elemen):  # disisipkan ke belakang satu per satu: urutan tetap sama
            if not kecuali(el):
                baru = self.salin(sumber, slide, el)
                self._potong_kanan(baru)

    @staticmethod
    def _potong_kanan(el):
        """Kotak teks rata kiri yang melewati tepi kanan slide dipotong lebarnya (teks tidak bergeser)."""
        if el.tag != qn("p:sp") or not _teks_dalam(el).strip():
            return
        x, _, w, _ = _posisi(el)
        rata = {p.get("algn") for p in el.iter(qn("a:pPr"))}
        if x > 0 and x + w > D.LEBAR_SLIDE and not rata & {"r", "ctr", "just", "dist"}:
            _xfrm(el).find(qn("a:ext")).set("cx", str(int((D.LEBAR_SLIDE - x) * EMU)))

    @staticmethod
    def _teks_contoh(el, *kata):
        t = _teks_dalam(el).upper()
        return el.tag == qn("p:sp") and any(k in t for k in kata)

    # ------------------------------------------------------------ kait
    def hias(self, slide):
        """Hiasan samar slide isi: gambar bebas di slide contoh isi (selain chip, tabel, navigasi)."""
        sumber = self.sumber["isi"]
        for el in reversed(list(sumber.shapes._spTree.iterchildren(qn("p:sp")))):
            if _isian_gambar(el):
                self.salin(sumber, slide, el)

    def chip(self, slide, label, ikon):
        sumber = self.sumber["isi"]
        grup = next(el for el in sumber.shapes._spTree.iterchildren(qn("p:grpSp"))
                    if _posisi(el)[1] < 0.6 and _posisi(el)[0] < 1.0 and not _navbar_template(el))
        self.salin(sumber, slide, grup, belakang=False)
        x, y, w, h = _posisi(grup)
        lebar = w - 1.4
        ukuran = 26
        while D.lebar_pt(label, ukuran, True) / 72 > lebar * 0.97 and ukuran > 15:
            ukuran -= 1
        D.teks(slide, x + 1.17, y, lebar, h, [P(label, ukuran, "putih", tebal=True, spasi=1.05)], jangkar="m",
               nama="Label bagian")
        return w + 0.05

    def judul(self, slide, isi, y=D.Y_JUDUL):
        ukuran = 30
        segs = D.segmen(isi, tebal=True)
        while D.lebar_segs(segs, ukuran) > D.LEBAR_ISI * 0.97 and ukuran > 22:
            ukuran -= 2
        D.teks(slide, D.KIRI, y - 0.05, D.LEBAR_ISI, 0.75, [P(isi, ukuran, "navy", tebal=True, spasi=1.1)],
               nama="Judul")
        return y + 0.75

    def navbar(self, slide, aktif, label, gelap=False):
        """Pil putih bergaris navy (atau putih di slide berbingkai), bagian aktif pil navy/putih."""
        x0, y0, w0, h0 = NAVBAR
        garis = "putih" if gelap else "navy"
        pil = D.kotak(slide, x0, y0, w0, h0, warna=None if gelap else "putih", radius=h0 / 2, garis=garis,
                      tebal_garis=1.5, nama="Navigasi")
        if gelap:
            pil.fill.background()
        # ikon kaca pembesar dari template, di ujung kanan
        sumber = self.sumber["isi"]
        nav = next((el for el in sumber.shapes._spTree.iterchildren(qn("p:grpSp")) if _navbar_template(el)), None)
        ikon = next((el for el in nav.iter(qn("p:sp")) if _isian_gambar(el)), None) if nav is not None else None
        d = 0.46
        if ikon is not None:
            self.salin(sumber, slide, ikon, belakang=False, xfrm=(x0 + w0 - d - 0.12, y0 + (h0 - d) / 2, d, d))
        huruf = [n.upper() for n in label]
        ukuran = 10
        lebar = [D.lebar_pt(t, ukuran, True) / 72 + 0.36 for t in huruf]
        ruang = w0 - 0.3 - d - 0.25
        celah = (ruang - sum(lebar)) / max(1, len(label) - 1)
        x = x0 + 0.15
        hasil = {}
        for nama, t, lw in zip(label, huruf, lebar):
            if nama == aktif:
                D.kotak(slide, x, y0 + 0.11, lw, h0 - 0.22, warna="putih" if gelap else "navy",
                        radius=(h0 - 0.22) / 2, nama=f"Nav aktif {nama}")
                warna = "navy" if gelap else "putih"
            else:
                warna = "putih" if gelap else "navy"
            hasil[nama] = D.teks(slide, x, y0, lw, h0, [P(t, ukuran, warna, tebal=True, rata="c", spasi=1.0)],
                                 jangkar="m", nama=f"Nav {nama}")
            x += lw + celah
        return hasil

    def nomor_halaman(self, slide, nomor, total, gelap=False):
        D.lingkaran(slide, 14.55, NAVBAR[1] + 0.07, 0.52, "putih" if gelap else "navy", nomor,
                    12 if nomor < 100 else 10, "navy" if gelap else "putih", nama="Nomor halaman")

    # ----------------------------------------------------- slide khusus
    def sampul(self, slide, besar, judul, tutor=None):
        self.salin_slide("sampul", slide, lambda el: el.tag == qn("p:sp") and _teks_dalam(el).strip() != "")
        lebar = 7.2
        ukuran = 70
        while D.lebar_pt(besar, ukuran, True) / 72 > lebar and ukuran > 32:
            ukuran -= 2
        D.teks(slide, 2.58, 4.17, lebar + 0.3, 1.3, [P(besar, ukuran, "000000", tebal=True, spasi=1.0)],
               jangkar="b", nama="Materi besar")
        segs = D.segmen(judul, tebal=True)
        ukuran = 46
        while D.jumlah_baris(segs, ukuran, lebar) > 3 and ukuran > 26:
            ukuran -= 2
        h = D.jumlah_baris(segs, ukuran, lebar) * ukuran * 1.1 / 72 + 0.1
        D.teks(slide, 2.58, 5.6, lebar + 0.3, h, [P(judul, ukuran, "000000", tebal=True, spasi=1.1)], nama="Judul sampul")
        if tutor:
            D.teks(slide, 6.07, 9.34, 4.25, 0.5, [P(tutor, 22, "putih", tebal=True, rata="c", spasi=1.1)],
                   jangkar="m", nama="Nama tutor")
        D.teks(slide, 5.96, 10.05, 8.92, 0.4, [P(D.HAK_CIPTA, 16, "000000", rata="r", spasi=1.1)], nama="Hak cipta")

    def _isi_halaman(self, slide, par, y, w, h_maks, x=1.12):
        """Paragraf di halaman berbingkai; ukuran diperkecil sampai muat h_maks."""
        for ukuran in (24, 22, 20, 19, 18, 17, 16, 15, 14):
            p = [P(t, ukuran, "000000", rata="j", spasi=1.4, setelah=ukuran * 0.6) for t in par]
            h = D.tinggi_paragraf(p, w)
            if h <= h_maks:
                break
        D.teks(slide, x, y, w, h, p, nama="Isi halaman")
        return y + h

    def pendahuluan(self, slide, paragraf, pemantik=None, gambar=None):
        self.salin_slide("pendahuluan", slide,
                         lambda el: _navbar_template(el) or self._teks_contoh(el, "LOREM"))
        w = 12.37
        h_pemantik = D.tinggi_catatan(pemantik, w, "Pertanyaan pemantik", 19) + 0.3 if pemantik else 0
        y = self._isi_halaman(slide, paragraf, 2.02, w, BATAS_HALAMAN - 2.02 - h_pemantik)
        if pemantik:
            y = D.catatan(slide, y + 0.3, "pemantik", pemantik, x=1.12, w=w, judul_catatan="Pertanyaan pemantik",
                          ukuran=19)
        return y

    def tujuan(self, slide, pengantar, butir):
        self.salin_slide("tujuan", slide, lambda el: _navbar_template(el) or (
            el.tag == qn("p:grpSp") and el.find(f".//{qn('a:gradFill')}") is not None))
        y = 2.1
        if pengantar:
            D.teks(slide, 0.78, y, 14.43, 0.4, [P(pengantar, 18, "navy", spasi=1.2)], nama="Pengantar tujuan")
            y += 0.55
        n = max(1, len(butir))
        celah = 0.15
        h_maks = min(1.12, (9.75 - y - celah * (n - 1)) / n)
        d = 0.66
        for i, b in enumerate(butir):
            segs = [(b["kko"], True, False, "putih")] + D.segmen(" " + b["teks"])
            for ukuran in (20, 19, 18, 17, 16, 15):
                p = [P(None, ukuran, "putih", spasi=1.3, segs=segs)]
                inset = (0.35 + d + 0.3, 0.12, 0.4, 0.12)
                h = max(h_maks, D.tinggi_paragraf(p, 14.43, inset))
                if h <= max(h_maks, 1.12):
                    break
            D.kotak(slide, 0.78, y, 14.43, h, warna="kartu_gelap", radius=0, nama="Butir tujuan")
            level = b.get("level") or str(i + 1)
            D.lingkaran(slide, 0.78 + 0.35, y + (h - d) / 2, d, "kuning", level, 17, "navy", nama=f"Level {level}")
            D.teks(slide, 0.78, y, 14.43, h, p, inset=inset, jangkar="m")
            y += h + celah
        return y - celah

    def penutup(self, slide, ringkasan, pesan):
        self.salin_slide("penutup", slide, lambda el: _navbar_template(el) or self._teks_contoh(el, "LOREM"))
        x, w = 1.12, 12.6
        y = 2.0
        d = 0.6
        for i, r in enumerate(ringkasan, 1):
            isi = f"**{r['judul']}:** {r['teks']}" if r.get("teks") else f"**{r['judul']}**"
            p = [P(isi, 18, "navy", spasi=1.3)]
            inset = (0.25 + d + 0.25, 0.14, 0.3, 0.14)
            h = max(0.8, D.tinggi_paragraf(p, w, inset))
            D.kotak(slide, x, y, w, h, warna="kartu", radius=D.RADIUS, nama=f"Ringkasan {i}")
            D.lingkaran(slide, x + 0.25, y + (h - d) / 2, d, "navy", i, 16, "putih", nama=f"Nomor {i}")
            D.teks(slide, x, y, w, h, p, inset=inset, jangkar="m")
            y += h + 0.15
        if pesan:
            sisi = (0.4, 0.3, 0.4, 0.3)
            for ukuran in (22, 20, 19, 18, 17, 16):
                pp = [P("Pesan penutup", 14, "kuning", tebal=True, spasi=1.2, setelah=6),
                      P(pesan, ukuran, "putih", tebal=ukuran >= 19, spasi=1.3)]
                h = D.tinggi_paragraf(pp, w, sisi)
                if y + 0.15 + h <= BATAS_HALAMAN:
                    break
            D.kotak(slide, x, y + 0.15, w, h, warna="kartu_gelap", radius=D.RADIUS, nama="Pesan penutup")
            D.teks(slide, x, y + 0.15, w, h, pp, inset=sisi, jangkar="m")
            y += 0.15 + h
        return y

    def terima_kasih(self, slide):
        self.salin_slide("terima_kasih", slide, lambda el: False)
