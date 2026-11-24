"""PDF Figural (ekspor deck Canva, 1440 x 810 pt) -> naskah JSON generator.

Dokumen Figural di sheet 'Kerangka Course' berupa PDF, bukan Google Docs. Teks dan kunci diambil dari
PDF; gambar soal dipotong dari halaman. Bagian yang tidak ada di PDF (pendahuluan, tujuan, penutup)
dibiarkan kosong dan slidenya dilewati (meta.lewati_kosong).

    python generator/ekstrak_figural.py --pdf folder_pdf [--keluar konten] [--id fi-01]

folder_pdf berisi <doc_id>.pdf sesuai konten/katalog.json (entri kelompok "Figural"). Butuh PyMuPDF.
"""
from __future__ import annotations

import argparse
import json
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import ekstrak_gdoc as E  # noqa: E402

AKAR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
ZOOM = 2.0  # piksel per pt saat memotong gambar
KE_INCI = 0.8 / 72  # lebar halaman PDF 20 in -> slide 16 in
BIRU = (42, 70, 139)  # panel biru halaman transformasi
# (id entri, halaman) -> potongan gambar yang tidak bisa dihitung otomatis: daftar kotak pt
POTONG = {
    ("fi-02", 2): [(70, 268, 1400, 636)],
    ("fi-05", 7): [(88, 240, 790, 806)],
}
TEKS_DI_GAMBAR = {("fi-05", 7)}  # teksnya sudah ikut di dalam potongan gambar (tidak ditulis ulang)
KOREKSI = {"Pencermninan": "Pencerminan"}


def baris_teks(hal):
    """[(x0, y0, x1, y1, ukuran, teks)] urut dari atas."""
    hasil = []
    for b in hal.get_text("dict")["blocks"]:
        for ln in b.get("lines", []):
            t = "".join(s["text"] for s in ln["spans"]).replace("\u200b", "").strip()
            for salah, benar in KOREKSI.items():
                t = t.replace(salah, benar)
            if t:
                hasil.append((*ln["bbox"], ln["spans"][0]["size"], t))
    return sorted(hasil, key=lambda b: (round(b[1] / 6), b[0]))


def blok_teks(baris):
    """Baris-baris teks -> blok paragraf dan poin. Paragraf baru dimulai oleh jeda lebar, baris berbutir
    ('- ' atau menjorok), atau baris sebelumnya yang berhenti jauh sebelum tepi kanan."""
    hasil = []
    if not baris:
        return hasil
    kiri, kanan = min(b[0] for b in baris), max(b[2] for b in baris)
    akhir = None
    for x0, y0, x1, y1, _, t in baris:
        butir = t.startswith("- ") or x0 - kiri > 20
        isi = t[2:].strip() if t.startswith("- ") else t
        lanjut = akhir and not t.startswith("- ") and y0 - akhir["y1"] < 0.55 * (y1 - y0) and butir == akhir["butir"] \
            and not re.match(r"^[A-E](, [A-E])* ?:", t) \
            and not (re.match(r"^[A-Z(]", t) and (akhir["x1"] < kanan - 130 or akhir["teks"].endswith((".", ":"))))
        if lanjut:
            akhir["teks"] += ("" if akhir["teks"].endswith("-") else " ") + isi
        else:
            akhir = {"teks": isi, "butir": butir}
            hasil.append(akhir)
        akhir.update(y1=y1, x1=x1)
    blok = []
    for p in hasil:
        teks = re.sub(r"(\w)- (\w)", r"\1-\2", p["teks"])
        if not p["butir"]:
            blok.append({"jenis": "paragraf", "teks": teks})
        elif blok and blok[-1]["jenis"] == "poin":
            blok[-1]["butir"].append(teks)
        else:
            blok.append({"jenis": "poin", "butir": [teks]})
    return blok


class Pengambil:
    def __init__(self, dok, entri, folder):
        import pymupdf
        self.pymupdf, self.dok, self.entri, self.folder = pymupdf, dok, entri, folder
        self.kode = entri["id"]
        self.slug = os.path.splitext(entri["naskah"])[0]
        self.n = 0

    def potong(self, hal, kotak, tutup=()):
        """Render bagian halaman; baris teks di tutup (kotak pt) dihapus dari gambar."""
        from PIL import Image, ImageDraw
        p = hal.get_pixmap(matrix=self.pymupdf.Matrix(ZOOM, ZOOM), clip=self.pymupdf.Rect(*kotak), alpha=False)
        im = Image.frombytes("RGB", (p.width, p.height), p.samples)
        d = ImageDraw.Draw(im)
        for x0, y0, x1, y1 in tutup:
            d.rectangle([(x0 - kotak[0] - 2) * ZOOM, (y0 - kotak[1]) * ZOOM, (x1 - kotak[0] + 2) * ZOOM, (y1 - kotak[1]) * ZOOM],
                        fill="white")
        return im

    def simpan(self, im, pangkas=True):
        """Simpan gambar (tepi putih dipangkas) dan kembalikan blok gambar."""
        from PIL import Image, ImageChops
        if pangkas:
            bb = ImageChops.difference(im, Image.new("RGB", im.size, "white")).point(lambda v: 255 if v > 24 else 0).getbbox()
            if bb:
                m = 8
                im = im.crop((max(0, bb[0] - m), max(0, bb[1] - m), min(im.width, bb[2] + m), min(im.height, bb[3] + m)))
        self.n += 1
        rel = f"media/{self.slug}/gambar-{self.n:02d}.png"
        path = os.path.join(self.folder, rel)
        os.makedirs(os.path.dirname(path), exist_ok=True)
        im.quantize(colors=128, method=Image.Quantize.MEDIANCUT).save(path, "PNG", optimize=True)
        return {"jenis": "gambar", "file": rel, "lebar": round(im.width / ZOOM * KE_INCI, 2),
                "tinggi": round(im.height / ZOOM * KE_INCI, 2)}

    def gambar_halaman(self, no):
        """Blok gambar halaman no: potongan khusus, atau gabungan kotak gambar raster di halaman."""
        hal = self.dok[no - 1]
        if (self.kode, no) in POTONG:
            return [self.simpan(self.potong(hal, k)) for k in POTONG[(self.kode, no)]]
        kotak = [g["bbox"] for g in hal.get_image_info()]
        if not kotak:
            return []
        baris = baris_teks(hal)
        tutup = [b[:4] for b in baris if b[4] < 60]
        x0, y0 = min(k[0] for k in kotak), min(k[1] for k in kotak)
        x1, y1 = max(k[2] for k in kotak), max(k[3] for k in kotak)
        # kotak "?" di ujung deret digambar sebagai teks besar, di luar gambar raster; hiasan di bawahnya ditutup
        for b in baris:
            if b[4] > 60:
                kanan = (b[0] + b[2]) / 2 + 95
                tutup.append((x1 + 5, b[3] + 40, kanan + 5, y1 + 5))
                x1 = kanan
        return [self.simpan(self.potong(hal, (x0, y0, x1, y1), tutup))]

    # ------------------------------------------------------------- soal ---
    def soal(self):
        """Pasangan halaman soal -> kunci menjadi slide contoh_soal; [(kelompok, slide)]."""
        hasil, jenis = [], ""
        for no in range(2, len(self.dok)):
            baris = [b for b in baris_teks(self.dok[no - 1]) if b[4] < 40]  # judul halaman (49 pt) dilewati
            kunci = next((re.match(r"^Kunci jawaban\s*:\s*([A-E])\b", b[5], re.I) for b in baris
                          if re.match(r"^Kunci jawaban", b[5], re.I)), None)
            if kunci:
                isi = [b for b in baris if not re.match(r"^Kunci jawaban", b[5], re.I)]
                blok = [{"jenis": "kunci", "teks": f"**Kunci jawaban: {kunci.group(1).upper()}**"}]
                if (self.kode, no) == ("fi-05", 5):
                    blok += self.matriks_berwarna(self.dok[no - 1], isi)
                else:
                    blok += ([] if (self.kode, no) in TEKS_DI_GAMBAR else blok_teks(isi)) + self.gambar_halaman(no)
                hasil[-1][1].update(kunci=kunci.group(1).upper(), pembahasan=blok)
                continue
            nomor = next((m.group(1) for b in baris for m in [re.match(r"^Contoh\s*(\d+)\s*:?$", b[5])] if m), None)
            for b in baris:
                if re.match(r"^\d+ Kotak$", b[5]):
                    jenis = b[5]
            teks = blok_teks([b for b in baris if not re.match(r"^(Contoh\s*\d+\s*:?|\d+ Kotak|[•\-]|[A-E]|\?)$", b[5])])
            judul = f"{jenis} – Contoh {nomor}" if jenis else f"Contoh {nomor}"
            hasil.append((jenis or "Contoh Soal", {
                "tipe": "contoh_soal", "judul": judul, "opsi": {}, "kunci": "", "pembahasan": [],
                "soal": teks + self.gambar_halaman(no)}))
        return hasil

    def matriks_berwarna(self, hal, baris):
        """Pembahasan matriks di panel biru: dua kotak 3x3 bertanda merah dan dua deret bentuk."""
        hasil = []
        for (y_dari, y_sampai, y_akhir), matriks, deret in (((170, 390, 425), (80, 178, 346, 433), (377, 290, 800, 397)),
                                                           ((430, 655, 720), (80, 435, 346, 701), (377, 554, 800, 660))):
            hasil += blok_teks([b for b in baris if y_dari <= b[1] < y_sampai])
            hasil.append(self.simpan(self._tanpa_biru(self.potong(hal, matriks))))
            hasil.append(self.simpan(self._deret_putih(self.potong(hal, deret))))
            hasil += blok_teks([b for b in baris if y_sampai <= b[1] < y_akhir])
        return hasil

    @staticmethod
    def _tanpa_biru(im):
        """Latar panel biru -> putih."""
        px = im.load()
        for x in range(im.width):
            for y in range(im.height):
                if all(abs(a - b) < 45 for a, b in zip(px[x, y], BIRU)):
                    px[x, y] = (255, 255, 255)
        return im

    @staticmethod
    def _deret_putih(im):
        """Deret kotak putih berpanah di latar biru -> latar putih, panah gelap, kotak berbingkai."""
        from PIL import ImageDraw
        lebar, tinggi = im.size
        px = im.load()
        di_kotak = [sum(px[x, y][0] > 200 or max(px[x, y]) < 60 for y in range(tinggi)) > tinggi * 0.5 for x in range(lebar)]
        for x in range(lebar):
            for y in range(tinggi):
                t = min(1, max(0, (px[x, y][0] - BIRU[0]) / (255 - BIRU[0])))
                biru = abs(px[x, y][2] - px[x, y][0]) > 40  # piksel kebiruan = latar atau tepinya
                if not di_kotak[x]:
                    v = round(255 - t * 195)
                    px[x, y] = (v, v, v)
                elif biru:
                    px[x, y] = (255, 255, 255)
        d = ImageDraw.Draw(im)
        awal = None
        for x in range(lebar + 1):
            if x < lebar and di_kotak[x]:
                awal = x if awal is None else awal
            elif awal is not None:
                d.rectangle([awal, 2, x - 1, tinggi - 3], outline=(90, 90, 90), width=2)
                awal = None
        return im

    # ---------------------------------------------------------- transformasi ---
    def kotak_transformasi(self, hal, y_dari, y_sampai):
        """Kotak-kotak putih di satu baris halaman biru, disusun ulang di latar putih."""
        from PIL import Image, ImageChops, ImageDraw
        kolom = ((176, 363), (476, 664), (776, 964), (1077, 1264))
        petak = []
        for x0, x1 in kolom:
            im = self.potong(hal, (x0 - 6, y_dari, x1 + 6, y_sampai))
            bb = ImageChops.difference(im, Image.new("RGB", im.size, BIRU)).convert("L").point(lambda v: 255 if v > 40 else 0).getbbox()
            if bb and bb[2] - bb[0] > 200 and bb[3] - bb[1] > 200:
                petak.append(im.crop(bb))
        if not petak:
            return None
        s = 380
        petak = [p.resize((s, s), Image.LANCZOS) for p in petak]
        jarak = 90
        lembar = Image.new("RGB", (len(petak) * s + (len(petak) - 1) * jarak + 8, s + 8), "white")
        d = ImageDraw.Draw(lembar)
        for i, p in enumerate(petak):
            x = 4 + i * (s + jarak)
            lembar.paste(p, (x, 4))
            d.rectangle([x, 4, x + s, 4 + s], outline=(90, 90, 90), width=3)
        return lembar

    def transformasi(self):
        materi = []
        # halaman 2: daftar isi dua kolom
        kartu = []
        baris = [b for b in baris_teks(self.dok[1]) if b[4] < 40 and b[5] not in "•"]
        for kolom in ((0, 700), (700, 1440)):
            isi = [b for b in baris if kolom[0] <= b[0] < kolom[1]]
            kiri = min(b[0] for b in isi)
            for b in sorted(isi, key=lambda b: b[1]):
                if b[0] - kiri < 8:
                    kartu.append({"judul": b[5], "poin": []})
                else:
                    kartu[-1]["poin"].append(b[5])
        materi.append({"subbab": "Macam-Macam Transformasi", "ikon": "checklist",
                       "slide": [{"tipe": "kartu", "judul": "Macam-Macam Transformasi Figural", "kolom": 3, "kartu": kartu}]})
        for no in range(3, len(self.dok)):
            hal = self.dok[no - 1]
            baris = baris_teks(hal)
            subbab = next(b[5] for b in baris if 28 <= b[4] < 40)
            label = [b for b in baris if b[5].startswith("Contoh")]
            jenis = next(b[5] for b in baris if b[4] < 28 and not b[5].startswith("Contoh") and b[5] != "•")
            blok = []
            for i, lb in enumerate(label):
                bawah = label[i + 1][1] - 4 if i + 1 < len(label) else 752
                im = self.kotak_transformasi(hal, lb[3] + 2, bawah)
                blok.append({"jenis": "subjudul", "teks": lb[5]})
                if im is not None:
                    g = self.simpan(im, pangkas=False)
                    g.update(lebar=round(g["lebar"] / 1.5, 2), tinggi=round(g["tinggi"] / 1.5, 2))  # dua baris per slide
                    blok.append(g)
            if not materi or materi[-1]["subbab"] != subbab:
                materi.append({"subbab": subbab, "ikon": E.IKON_SUBBAB[len(materi) % len(E.IKON_SUBBAB)], "slide": []})
            materi[-1]["slide"].append({"tipe": "blok", "judul": jenis, "blok": blok})
        return materi

    # --------------------------------------------------------------- metode ---
    def metode(self):
        slide = []
        for no in range(2, len(self.dok)):
            hal = self.dok[no - 1]
            baris = baris_teks(hal)
            judul = next((b[5] for b in baris if 40 <= b[4] < 60), "Metode Pengerjaan")
            isi = [b for b in baris if b[4] < 40 and b[5] != "•"]
            label = [b for b in isi if b[5].startswith("Metode ")]
            blok = blok_teks([b for b in isi if b not in label])
            if label:  # tiap metode: subjudul lalu gambar contohnya
                kotak = [g["bbox"] for g in hal.get_image_info()]
                x0, x1, dasar = min(k[0] for k in kotak), max(k[2] for k in kotak), max(k[3] for k in kotak)
                for i, lb in enumerate(label):
                    bawah = label[i + 1][1] if i + 1 < len(label) else dasar
                    blok.append({"jenis": "subjudul", "teks": lb[5]})
                    blok.append(self.simpan(self.potong(hal, (x0, lb[1], x1, bawah), [lb[:4]])))
            else:
                blok += self.gambar_halaman(no)
            slide.append({"tipe": "blok", "judul": judul, "blok": blok})
        return [{"subbab": slide[0]["judul"], "ikon": "strategi", "slide": slide}]


def susun(entri, dok, folder):
    a = Pengambil(dok, entri, folder)
    soal = entri["id"].endswith("-soal")
    if entri["id"] == "fi-01":
        materi = a.transformasi()
    elif soal or entri["id"] in ("fi-04", "fi-05"):
        materi = []
        for kelompok, s in a.soal():
            if not materi or materi[-1]["subbab"] != kelompok:
                materi.append({"subbab": kelompok, "ikon": "checklist", "slide": []})
            materi[-1]["slide"].append(s)
    else:
        materi = a.metode()
    judul = entri["bab"].replace("Contoh Soal Figural - ", "Contoh Soal ").replace("Figural - ", "")
    meta = {"judul": judul, "subtes": E.SUBTES_TIU, "nama_file": f"TIU - {entri['bab']}", "sumber": entri["url"],
            "id_katalog": entri["id"], "lewati_kosong": True}
    if soal:
        meta["bagian"] = "soal"
    return {"meta": meta, "peta_subtes": dict(E.PETA_TIU, aktif=5),
            "pendahuluan": {"paragraf": [], "pemantik": "", "ikon": "presenter"},
            "tujuan": {"pengantar": "Setelah mempelajari materi ini, peserta mampu:", "butir": []},
            "materi": materi, "penutup": {"judul": "Yang Perlu Diingat", "ringkasan": [], "pesan": ""}}


def main(argv=None):
    import pymupdf
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("--pdf", required=True, help="Folder berisi <doc_id>.pdf")
    ap.add_argument("--katalog", default=os.path.join(AKAR, "konten", "katalog.json"))
    ap.add_argument("--keluar", default=None, help="Folder naskah (bawaan: folder katalog)")
    ap.add_argument("--id", nargs="*")
    a = ap.parse_args(argv)
    folder = a.keluar or os.path.dirname(os.path.abspath(a.katalog))
    with open(a.katalog, encoding="utf-8") as f:
        katalog = json.load(f)
    for m in katalog["materi"]:
        if m["kelompok"] != "Figural" or (a.id and m["id"] not in a.id):
            continue
        naskah = susun(m, pymupdf.open(os.path.join(a.pdf, m["doc_id"] + ".pdf")), folder)
        with open(os.path.join(folder, m["naskah"]), "w", encoding="utf-8") as f:
            json.dump(naskah, f, ensure_ascii=False, indent=2)
            f.write("\n")
        slide = [s for sb in naskah["materi"] for s in sb["slide"]]
        print(f"{m['id']}: {naskah['meta']['judul']} — {len(slide)} slide isi, "
              f"{sum(s['tipe'] == 'contoh_soal' for s in slide)} contoh soal, "
              f"kunci {[s['kunci'] for s in slide if s['tipe'] == 'contoh_soal']}")


if __name__ == "__main__":
    main()
