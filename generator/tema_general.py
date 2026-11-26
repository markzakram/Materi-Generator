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

from PIL import Image, ImageChops
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


def _abu_abu(im):
    """True bila piksel yang terlihat praktis tanpa warna (R, G, B hampir sama).

    Piksel hampir transparan (alfa < 24) diabaikan, dan sisa berwarna maks 0,05% piksel ditoleransi:
    garis gelombang template punya sedikit tepi berwarna yang tidak tampak.
    """
    r, g, b = im.convert("RGB").split()
    beda = ImageChops.lighter(ImageChops.difference(r, g), ImageChops.difference(g, b))
    if "A" in im.getbands():
        beda = ImageChops.multiply(beda, im.getchannel("A").point(lambda a: 255 if a >= 24 else 0))
    return sum(beda.histogram()[11:]) <= im.width * im.height * 0.0005


LEBAR_MAKS, LEBAR_MAKS_ABU = 1600, 1200


def _perkecil(blob, jenis):
    """Gambar besar -> lebar maks 1600 px (PNG dikuantisasi 96 warna, JPEG kualitas 82).

    Hiasan abu-abu (garis gelombang, titik raster) cukup selebar 1200 px: hiasan samar ini tidak butuh
    resolusi tinggi. Di PDF, buat_situs menyimpannya sebagai gambar satu kanal.
    """
    with Image.open(io.BytesIO(blob)) as im:
        im.load()
        abu = jenis == "image/png" and _abu_abu(im)
        maks = LEBAR_MAKS_ABU if abu else LEBAR_MAKS
        if im.width > maks:
            im = im.resize((maks, round(im.height * maks / im.width)), Image.LANCZOS)
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
                if len(part.blob) > 150_000 and part.content_type in ("image/png", "image/jpeg"):
                    part._blob = _perkecil(part.blob, part.content_type)

    # ------------------------------------------------------- slide contoh
    @staticmethod
    def _kenali(slides):
        hasil = {}
        for s in slides:
            besar = _teks_besar(s)
            gabung = " ".join(besar)
            if "MATERI BESAR" in gabung:
                # template punya dua sampul; yang dipakai sampul foto (panel navy, 'Judul Materi (boleh panjang)')
                if "BOLEH PANJANG" in " ".join(_teks_besar(s, 0)).upper().replace("(", " ").replace(")", " "):
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
    VARIAN_SAMPUL = ("putih", "foto")

    def sampul(self, slide, besar, judul, tutor=None):
        """Sampul foto: panel navy di atas foto, bab putih dan judul kuning di dalam panel (x 0,36-10,19)."""
        # kotak abu 'Nama Tutor' di kanan bawah hanya dipakai bila ada nama tutor
        self.salin_slide("sampul", slide, lambda el: (
            (el.tag == qn("p:sp") and _teks_dalam(el).strip() != "")
            or (not tutor and el.tag == qn("p:grpSp") and _posisi(el)[0] > 10 and _posisi(el)[1] > 8.5)))
        self._judul_sampul(slide, 1.0, 8.2, 5.7, besar, judul, "putih", "kuning_di_gelap")
        if tutor:
            D.teks(slide, 11.36, 10.05, 4.25, 0.5, [P(tutor, 22, "putih", tebal=True, rata="c", spasi=1.1)],
                   jangkar="m", nama="Nama tutor")
        D.teks(slide, 0.36, 9.18, 8.92, 0.4, [P(D.HAK_CIPTA, 16, "000000", rata="r", spasi=1.1)], nama="Hak cipta")

    @staticmethod
    def _judul_sampul(slide, x, lebar, y_dasar, besar, judul, warna_bab, warna_judul):
        """Seperti template: materi besar (bab) jauh lebih besar dari judul materi (subbab).

        Bab boleh dua baris agar tetap besar; garis bawahnya di y_dasar, judul tepat di bawahnya.
        """
        def muat(teks, ukuran, maks_baris):
            kata = max(D.lebar_pt(k, ukuran, True) for k in teks.split()) / 72
            return kata <= lebar * 0.95 and D.jumlah_baris(D.segmen(teks, tebal=True), ukuran, lebar) <= maks_baris

        ukuran = 64
        while not muat(besar, ukuran, 2) and ukuran > 36:
            ukuran -= 2
        h = D.jumlah_baris(D.segmen(besar, tebal=True), ukuran, lebar) * ukuran / 72 + 0.12
        D.teks(slide, x, y_dasar - h, lebar, h,
               [P(besar, ukuran, warna_bab, tebal=True, spasi=1.0)], jangkar="b", nama="Materi besar")
        u_judul = min(40, int(ukuran * 0.68))
        while not muat(judul, u_judul, 2) and u_judul > 24:
            u_judul -= 2
        h = D.jumlah_baris(D.segmen(judul, tebal=True), u_judul, lebar) * u_judul * 1.15 / 72 + 0.12
        D.teks(slide, x, y_dasar + 0.12, lebar, h,
               [P(judul, u_judul, warna_judul, tebal=True, spasi=1.1)], nama="Judul sampul")

    def _isi_halaman(self, slide, par, y, w, h_maks, x=1.12):
        """Paragraf di halaman berbingkai; ukuran diperkecil sampai muat h_maks."""
        for ukuran in (24, 22, 20, 19, 18, 17, 16, 15, 14):
            p = [P(t, ukuran, "000000", rata="j", spasi=1.4, setelah=ukuran * 0.6) for t in par]
            # cadangan 4% + 0,05 in: rata kanan-kiri di PowerPoint sedikit lebih tinggi dari perkiraan
            h = D.tinggi_paragraf(p, w) * 1.04 + 0.05
            if h <= h_maks:
                break
        D.teks(slide, x, y, w, h, p, nama="Isi halaman")
        return y + h

    W_PENDAHULUAN, Y_PENDAHULUAN = 12.37, 2.02

    def _tinggi_pemantik(self, pemantik):
        return D.tinggi_catatan(pemantik, self.W_PENDAHULUAN, "Pertanyaan pemantik", 19) + 0.3 if pemantik else 0

    def tinggi_gambar_pendahuluan(self, paragraf, pemantik, gambar, folder):
        """Ruang (inci) untuk gambar di bawah paragraf 20 pt; 0 bila gambar harus ke slide sendiri."""
        p = [P(t, 20, "000000", rata="j", spasi=1.4, setelah=12) for t in paragraf]
        h_teks = D.tinggi_paragraf(p, self.W_PENDAHULUAN) * 1.04 + 0.3 if p else 0
        sisa = BATAS_HALAMAN - self.Y_PENDAHULUAN - self._tinggi_pemantik(pemantik) - h_teks
        h_penuh = D.tinggi_blok(gambar, self.W_PENDAHULUAN, folder)
        return min(sisa, h_penuh) if sisa >= 0.75 * h_penuh or not paragraf else 0

    def pendahuluan(self, slide, paragraf, pemantik=None, gambar=None, folder="."):
        """gambar: blok gambar di bawah paragraf (pastikan muat lewat tinggi_gambar_pendahuluan)."""
        self.salin_slide("pendahuluan", slide,
                         lambda el: _navbar_template(el) or self._teks_contoh(el, "LOREM"))
        w, y = self.W_PENDAHULUAN, self.Y_PENDAHULUAN
        h_gambar = self.tinggi_gambar_pendahuluan(paragraf, pemantik, gambar, folder) if gambar else 0
        if paragraf:
            y = self._isi_halaman(slide, paragraf, y, w, BATAS_HALAMAN - y - self._tinggi_pemantik(pemantik)
                                  - (h_gambar + 0.2 if h_gambar else 0))
        if h_gambar:
            y = D.gambar_blok(slide, gambar, 1.12, y + 0.15 if paragraf else y, w, folder, h_maks=h_gambar)
        if pemantik:
            y = D.catatan(slide, y + 0.3, "pemantik", pemantik, x=1.12, w=w, judul_catatan="Pertanyaan pemantik",
                          ukuran=19)
        return y

    def tujuan(self, slide, pengantar, butir, awal=0):
        """awal: nomor butir pertama dikurangi satu (butir yang dibagi ke beberapa slide)."""
        self.salin_slide("tujuan", slide, lambda el: _navbar_template(el) or (
            el.tag == qn("p:grpSp") and el.find(f".//{qn('a:gradFill')}") is not None))
        y = 2.1
        if pengantar:
            D.teks(slide, 0.78, y, 14.43, 0.4, [P(pengantar, 18, "navy", spasi=1.2)], nama="Pengantar tujuan")
            y += 0.55
        n = max(1, len(butir))
        celah = 0.15 if n <= 5 else 0.1
        h_maks = min(1.12, (9.75 - y - celah * (n - 1)) / n)  # jatah per butir agar semua muat
        d = min(0.66, h_maks - 0.14)
        inset = (0.35 + d + 0.3, 0.1, 0.4, 0.1)
        isi = [[(b["kko"], True, False, "putih")] + D.segmen(" " + b["teks"]) for b in butir]
        tinggi = lambda segs, u: D.tinggi_paragraf([P(None, u, "putih", spasi=1.3, segs=segs)], 14.43, inset)  # noqa: E731
        # satu ukuran huruf untuk semua butir: terbesar yang membuat butir terpanjang pun muat
        ukuran = next((u for u in (20, 19, 18, 17, 16, 15, 14, 13) if all(tinggi(s, u) <= h_maks for s in isi)), 13)
        for i, b in enumerate(butir):
            p = [P(None, ukuran, "putih", spasi=1.3, segs=isi[i])]
            h = max(h_maks, tinggi(isi[i], ukuran))
            D.kotak(slide, 0.78, y, 14.43, h, warna="kartu_gelap", radius=0, nama="Butir tujuan")
            level = b.get("level") or str(awal + i + 1)
            D.lingkaran(slide, 0.78 + 0.35, y + (h - d) / 2, d, "kuning", level, min(17, round(d * 26)), "di_kuning",
                        nama=f"Level {level}")
            D.teks(slide, 0.78, y, 14.43, h, p, inset=inset, jangkar="m")
            y += h + celah
        return y - celah

    # ukuran slide penutup (inci): isi x, lebar, y awal, diameter nomor, jarak antarbutir
    PENUTUP = {"x": 1.12, "w": 12.6, "y": 2.0, "d": 0.6, "celah": 0.15, "sisi": (0.4, 0.3, 0.4, 0.3)}

    def susun_penutup(self, ringkasan, pesan):
        """(kolom, baris butir, paragraf pesan, tinggi pesan) bila semuanya muat di satu slide, atau None.

        Pesan dikecilkan dulu (22 sampai 14 pt), lalu butir ringkasan dibuat dua kolom. Pesan tanpa ringkasan
        boleh sampai 12 pt karena tidak bisa dibagi lagi.
        """
        u = self.PENUTUP
        inset = (0.25 + u["d"] + 0.25, 0.14, 0.3, 0.14)
        par = [[P(f"**{r['judul']}:** {r['teks']}" if r.get("teks") else f"**{r['judul']}**", 18, "navy", spasi=1.3)]
               for r in ringkasan]
        pilihan = [(1, range(22, 15, -1)), (2, range(22, 13, -1)), (1, range(15, 13, -1))] if par else [(1, range(22, 11, -1))]
        for kolom, deret in pilihan:
            lw = (u["w"] - u["celah"] * (kolom - 1)) / kolom
            baris = [(max(max(0.8, D.tinggi_paragraf(p, lw, inset)) for p in par[i:i + kolom]), par[i:i + kolom])
                     for i in range(0, len(par), kolom)]
            for ukuran in (deret if pesan else (0,)):
                pp, hp = None, 0
                if pesan:
                    pp = [P("Pesan penutup", 14, "kuning_di_gelap", tebal=True, spasi=1.2, setelah=6),
                          P(pesan, ukuran, "putih", tebal=ukuran >= 19, spasi=1.3, rata="j")]
                    hp = D.tinggi_paragraf(pp, u["w"], u["sisi"]) + 0.23
                if u["y"] + sum(h + u["celah"] for h, _ in baris) + hp <= BATAS_HALAMAN:
                    return kolom, baris, pp, hp
        return None

    def penutup(self, slide, ringkasan, pesan):
        self.salin_slide("penutup", slide, lambda el: _navbar_template(el) or self._teks_contoh(el, "LOREM"))
        u = self.PENUTUP
        x, w, y, d, celah = u["x"], u["w"], u["y"], u["d"], u["celah"]
        inset = (0.25 + d + 0.25, 0.14, 0.3, 0.14)
        n = len(ringkasan)
        while (susunan := self.susun_penutup(ringkasan[:n], pesan)) is None and n:
            n -= 1  # hanya bila butirnya sendiri lebih tinggi dari satu halaman
        kolom, baris, pp, hp = susunan or (1, [], None, 0)
        lw = (w - celah * (kolom - 1)) / kolom
        i = 0
        for h, isi in baris:
            for j, p in enumerate(isi):
                i += 1
                kx = x + j * (lw + celah)
                D.kotak(slide, kx, y, lw, h, warna="kartu", radius=D.RADIUS, nama=f"Ringkasan {i}")
                D.lingkaran(slide, kx + 0.25, y + (h - d) / 2, d, "navy", i, 16, "putih", nama=f"Nomor {i}")
                D.teks(slide, kx, y, lw, h, p, inset=inset, jangkar="m")
            y += h + celah
        if pp:
            h = hp - 0.15
            D.kotak(slide, x, y + 0.15, w, h, warna="kartu_gelap", radius=D.RADIUS, nama="Pesan penutup")
            D.teks(slide, x, y + 0.15, w, h, pp, inset=u["sisi"], jangkar="m")
            y += 0.15 + h
        return y

    def terima_kasih(self, slide):
        self.salin_slide("terima_kasih", slide, lambda el: False)
