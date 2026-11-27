"""Pecah naskah satu dokumen menjadi naskah per Bab (baris Bab di sheet), untuk dokumen yang mencakup beberapa Bab.

Dipakai impor_katalog.py untuk entri katalog yang punya "sumber" (lihat katalog["sumber"]). Aturannya:

- Subbab yang namanya sama dengan Bab masuk deck Bab itu. Subbab pengantar yang tidak bernama Bab (mis. "Penalaran
  Induktif") ikut di awal setiap deck Bab sesudahnya, sampai ada subbab pengantar berikutnya.
- Subbab "Rangkuman": butir, paragraf, atau baris tabel yang diawali nama Bab masuk deck Bab itu; sisanya ke semua deck.
- Tujuan pembelajaran: butir berawalan nama Bab ("Generalisasi: membedakan …; menyimpulkan …") masuk deck Bab itu dan
  dipecah per kata kerja; tanpa awalan, butir dibagi rata menurut urutan Bab.
- Pendahuluan dan pesan penutup tetap utuh (teksnya membahas seluruh materi); ringkasan penutup berisi subbab deck,
  atau judul bagian di dalam Bab bila deck hanya punya satu subbab.
- Contoh soal: label soal (nama subbab di samping judul) menentukan Bab; bila label tidak ada, nama Bab di judul soal,
  atau "nomor_soal" di entri katalog (urutan soal di dokumen, mulai 1).
"""
from __future__ import annotations

import copy
import re


def kunci(t):
    """Bentuk pembanding nama: huruf kecil, tanpa markup, '/' dan 'dan' diabaikan."""
    t = re.sub(r"\*\*|</?[a-z]+>|\$", "", t or "").lower().replace("/", " ")
    return " ".join(w for w in re.findall(r"[a-z0-9]+", t) if w != "dan")


def _awal(teks, daftar):
    """Indeks Bab yang namanya mengawali teks (setelah nomor), atau None. Nama terpanjang menang."""
    k = kunci(re.sub(r"^\s*(?:\*\*)?\(?\d{1,2}[.)]\s*", "", teks or ""))
    cocok = [(len(b), i) for i, b in enumerate(daftar) if b and (k == b or k.startswith(b + " "))]
    return max(cocok)[1] if cocok else None


def _rangkuman(sub, daftar):
    """{indeks Bab atau None (semua): subbab rangkuman berisi bagian untuk Bab itu}."""
    hasil = {}
    for i in list(range(len(daftar))) + [None]:
        s = copy.deepcopy(sub)
        for sl in s["slide"]:
            blok = []
            for b in sl.get("blok", []):
                if b["jenis"] in ("paragraf", "subjudul", "sorot"):
                    j = _awal(b["teks"], daftar)
                    if j is None:
                        blok.append(b)
                    elif j == i:  # '3. Penentuan …' tinggal satu butir: nomornya dibuang
                        blok.append(dict(b, teks=re.sub(r"^(\s*(?:\*\*)?)\(?\d{1,2}[.)]\s*", r"\1", b["teks"])))
                elif b["jenis"] == "poin":
                    butir = [x for x in b["butir"] if _awal(x, daftar) in (None, i)]
                    if butir:
                        blok.append(dict(b, butir=butir))
                elif b["jenis"] == "tabel":
                    baris = [r for r in b["baris"] if _awal(r[0] if r else "", daftar) in (None, i)]
                    if baris:
                        blok.append(dict(b, baris=baris))
                else:
                    blok.append(b)
            sl["blok"] = blok
        s["slide"] = [sl for sl in s["slide"] if sl.get("blok") or sl["tipe"] != "blok"]
        hasil[i] = s
    return hasil


KATA_KERJA = re.compile(r"^(me|ber|di|ter)[a-z]+$")


def _tujuan_bab(butir, daftar, i, n_bab):
    """Butir tujuan untuk Bab ke-i."""
    berawalan = [b for b in butir if _awal(b.get("kko", "") + " " + b["teks"], daftar) is not None]
    if berawalan:
        hasil = []
        for b in butir:
            if _awal(b.get("kko", "") + " " + b["teks"], daftar) != i:
                continue
            # "Generalisasi:" + "membedakan …; menyimpulkan …; serta mengevaluasi …." -> satu butir per kata kerja
            teks = b["teks"] if b.get("kko", "").rstrip().endswith(":") else re.sub(r"^[^:]+:\s*", "", b["teks"])
            bagian = [re.sub(r"^(serta|dan)\s+", "", x.strip()).rstrip(" .;") for x in re.split(r";\s*", teks) if x.strip()]
            if bagian and all(KATA_KERJA.match(x.split()[0].lower()) for x in bagian):
                hasil += [dict(b, kko=x.split()[0].capitalize(), teks=x.partition(" ")[2] + ".") for x in bagian]
            else:
                hasil.append(dict(b, kko=teks.split()[0].capitalize(), teks=teks.partition(" ")[2]))
        return hasil
    if butir and len(butir) % n_bab == 0:
        k = len(butir) // n_bab
        return butir[i * k:(i + 1) * k]
    print(f"  PERINGATAN tujuan tidak bisa dibagi per Bab ({len(butir)} butir, {n_bab} Bab); semua butir dipakai")
    return list(butir)


def _tujuan_asli(butir, daftar, i, n_bab):
    """Kunci teks butir tujuan dokumen (sebelum dipecah) yang dipakai Bab ke-i."""
    teks = [kunci(b.get("kko", "") + " " + b["teks"]) for b in butir]
    if any(_awal(b.get("kko", "") + " " + b["teks"], daftar) is not None for b in butir):
        return {t for b, t in zip(butir, teks) if _awal(b.get("kko", "") + " " + b["teks"], daftar) == i}
    if butir and len(butir) % n_bab == 0:
        k = len(butir) // n_bab
        return set(teks[i * k:(i + 1) * k])
    return set(teks)


def _judul_dalam(subbab):
    """Judul bagian di dalam satu subbab (judul slide yang berbeda dari subbab, lalu subjudul Heading)."""
    hasil = []
    for sl in subbab["slide"]:
        if sl.get("judul") and kunci(sl["judul"]) != kunci(subbab["subbab"]):
            hasil.append(sl["judul"])
        for b in sl.get("blok", []):
            if b["jenis"] == "subjudul" and b.get("tingkat", 1) == 1:
                hasil.append(re.sub(r"\*\*", "", b["teks"]))
    lihat = set()
    return [x for x in hasil if not (kunci(x) in lihat or lihat.add(kunci(x)))]


def bab_soal(contoh, daftar, nomor=None):
    """Indeks Bab untuk setiap slide contoh soal (urut), atau None bila tidak ditemukan."""
    hasil, n = [], 0
    for sub in contoh["materi"]:
        for sl in sub["slide"]:
            if sl["tipe"] != "contoh_soal":
                hasil.append(None)
                continue
            n += 1
            j = None
            if nomor:
                j = next((i for i, x in nomor.items() if n in x), None)
            label = kunci(sl.get("label"))
            if j is None and label in daftar:
                j = daftar.index(label)
            if j is None:  # nama Bab di label atau judul soal ('Fungsi Linear' -> Fungsi), lalu label di nama Bab
                for teks in (label, kunci(sl.get("judul"))):
                    cocok = [(len(b), i) for i, b in enumerate(daftar) if b and re.search(rf"\b{re.escape(b)}\b", teks)]
                    if cocok:
                        j = max(cocok)[1]
                        break
            if j is None and label:
                j = next((i for i, b in enumerate(daftar) if re.search(rf"\b{re.escape(label)}\b", b)), None)
            hasil.append(j)
    return hasil


def bagi(materi, contoh, entri, awalan):
    """[(entri katalog, naskah)] untuk setiap entri Bab (materi dan "bagian": "soal") dari satu dokumen."""
    bab_entri = [m for m in entri if m.get("bagian") != "soal"]
    daftar = [kunci(m["bab"]) for m in bab_entri]
    n = len(daftar)
    # subbab -> Bab
    isi = {i: [] for i in range(n)}
    pengantar, rangkuman, baru = [], [], False
    for sub in materi["materi"]:
        k = kunci(sub["subbab"])
        if k in daftar:
            i = daftar.index(k)
            isi[i] += pengantar + [sub] if not isi[i] else [sub]
            baru = True
        elif k.startswith("rangkuman"):
            rangkuman.append(sub)
        else:
            if baru:
                pengantar, baru = [], False
            pengantar.append(sub)
    kosong = [bab_entri[i]["bab"] for i in range(n) if not isi[i]]
    if kosong:
        print(f"  PERINGATAN Bab tanpa subbab bernama sama di dokumen: {'; '.join(kosong)}")
    for sub in rangkuman:
        bagian = _rangkuman(sub, daftar)
        for i in range(n):
            if bagian[i]["slide"]:
                isi[i].append(bagian[i])

    nomor = {}
    for m in entri:
        if m.get("bagian") == "soal" and m.get("nomor_soal"):
            nomor[daftar.index(kunci(m["bab"].removeprefix("Contoh Soal ")))] = set(m["nomor_soal"])
    milik = bab_soal(contoh, daftar, nomor)
    slide_soal = [(sl, sub) for sub in contoh["materi"] for sl in sub["slide"]]
    tanpa = [sl["judul"] for (sl, _), j in zip(slide_soal, milik) if j is None]
    if tanpa:
        print(f"  PERINGATAN soal tanpa Bab (dimasukkan ke Bab pertama; isi 'nomor_soal' di katalog): {'; '.join(tanpa)}")

    hasil = []
    for m in entri:
        soal = m.get("bagian") == "soal"
        nama_bab = m["bab"].removeprefix("Contoh Soal ") if soal else m["bab"]
        i = daftar.index(kunci(nama_bab))
        nama = f"{m.get('awalan', awalan)} - {m['bab']}".replace("/", "-").replace(": ", " - ")
        if soal:
            n_ = copy.deepcopy(contoh)
            ambil, k = [], 0
            for sub in n_["materi"]:
                pilih = []
                for sl in sub["slide"]:
                    j = milik[k] if milik[k] is not None else 0
                    k += 1
                    if j == i:
                        if kunci(sl.get("label")) == daftar[i]:
                            sl.pop("label", None)  # nama Bab sudah jadi judul deck
                        pilih.append(sl)
                if pilih:
                    ambil.append(dict(sub, slide=pilih))
            n_["materi"] = ambil
            n_["meta"].update(judul=m["bab"], nama_file=nama)
            # nomor soal dokumen ('Contoh Soal 7 — …') diurutkan ulang di deck Bab
            semua = [sl for sub in ambil for sl in sub["slide"] if sl["tipe"] == "contoh_soal"]
            pola = re.compile(r"^((?:Contoh Soal|Contoh|Soal)\s+)\d+\b")
            if semua and all(pola.match(sl["judul"]) for sl in semua):
                for no, sl in enumerate(semua, 1):
                    sl["judul"] = pola.sub(rf"\g<1>{no}", sl["judul"])
            if not ambil:
                print(f"  PERINGATAN {m['id']}: tidak ada contoh soal untuk Bab {nama_bab}")
            hasil.append((m, n_))
            continue
        n_ = copy.deepcopy(materi)
        n_["materi"] = copy.deepcopy(isi[i])
        n_["meta"].update(judul=m["bab"], nama_file=nama)
        n_["tujuan"]["butir"] = _tujuan_bab(materi["tujuan"]["butir"], daftar, i, n)
        if n_["tujuan"].get("tambahan"):  # catatan 'tujuan di luar highlight' hanya untuk butir Bab ini
            asli = _tujuan_asli(materi["tujuan"]["butir"], daftar, i, n)
            n_["tujuan"]["tambahan"] = [t for t in n_["tujuan"]["tambahan"] if kunci(t["isi"]) in asli]
            if not n_["tujuan"]["tambahan"]:
                n_["tujuan"].pop("tambahan")
        subbab = [s for s in n_["materi"] if not kunci(s["subbab"]).startswith("rangkuman")]
        judul = [s["subbab"] for s in subbab] if len(subbab) > 1 else _judul_dalam(subbab[0]) if subbab else []
        n_["penutup"]["ringkasan"] = [{"judul": j, "teks": ""} for j in judul[:10]]
        hasil.append((m, n_))
    for i in range(n):
        if any(j == i for j in milik) and not any(m.get("bagian") == "soal" and kunci(m["bab"].removeprefix("Contoh Soal ")) == daftar[i]
                                                 for m in entri):
            print(f"  PERINGATAN Bab {bab_entri[i]['bab']} punya contoh soal tetapi tidak punya entri soal di katalog")
    return hasil
