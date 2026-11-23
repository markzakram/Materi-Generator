"""Aplikasi web Materi Generator (Flask), siap di-deploy ke Vercel.

Halaman:
  /               Dasbor tinjauan: daftar materi di konten/katalog.json, status, catatan, unduh PPTX.
  /materi/<id>    Editor: pratinjau slide dan form isi per slide, simpan draf, unduh PPTX.
  /unggah         Buat PPTX dari berkas naskah JSON sendiri.

Ekspor PDF tidak tersedia di sini karena butuh Microsoft PowerPoint; buat PDF di Windows
lokal dengan generator/buat_materi.py.

Variabel lingkungan:
  MATERI_SANDI          Sandi akses (HTTP Basic, nama pengguna bebas). Di Vercel wajib
                        diisi, kecuali MATERI_TANPA_SANDI=1.
  MATERI_TANPA_SANDI    Isi 1 hanya jika akses sudah dilindungi Vercel Authentication.
  MATERI_TEMPLATE_URL   URL unduhan template Opsi B (.pptx). Wajib jika berkas template
                        tidak ada di akar proyek (selalu begitu di Vercel).
  MATERI_TEMPLATE_GENERAL_URL
                        URL unduhan 'Template Materi General.pptx' untuk tampilan General
                        (opsional; tanpa ini pilihan General menampilkan pesan).
  MATERI_ASET_TOKEN     Token 'Authorization: Bearer ...' untuk URL template (opsional).
  KV_REST_API_URL/TOKEN Upstash Redis untuk draf dan status (diisi otomatis oleh integrasi
                        Upstash di Vercel). Tanpa ini, Vercel berjalan dalam mode baca saja.

Jalankan lokal:
  python app.py --template "path/ke/Template Referensi Opsi B - Ceria & Membulat.pptx" \
                --template-general "path/ke/Template Materi General.pptx"
"""
from __future__ import annotations

import argparse
import datetime
import hashlib
import hmac
import io
import json
import os
import re
import shutil
import sys
import tempfile
import threading
import urllib.error
import urllib.parse
import urllib.request
import zipfile

from flask import Flask, Response, abort, jsonify, request, send_file, send_from_directory

AKAR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(AKAR, "generator"))
import buat_materi as BM  # noqa: E402
import desain as D  # noqa: E402
import pratinjau as PR  # noqa: E402

import penyimpanan  # noqa: E402

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 4 * 1024 * 1024  # Vercel menolak request di atas 4,5 MB
app.json.sort_keys = False  # urutan status dan isi naskah dipertahankan

BATAS_RESPONS = 4_400_000  # byte; Vercel menolak response di atas 4,5 MB
CACHE = os.path.join(tempfile.gettempdir(), "materi-generator")
KONTEN = os.path.join(AKAR, "konten")
WEB = os.path.join(AKAR, "web")
TEMPLATE_LOKAL = dict(BM.TEMPLATE_TEMA)  # tema -> path template lokal
ENV_TEMPLATE = {"opsi-b": "MATERI_TEMPLATE_URL", "general": "MATERI_TEMPLATE_GENERAL_URL"}
NAMA_TEMA = {"opsi-b": "Opsi B (Ceria & Membulat)", "general": "General"}
CONTOH = os.path.join(KONTEN, "tiu_penalaran_logis.json")
FONT_WEB = {
    "Poppins-Regular.ttf": os.path.join(D.ASET, "fonts", "Poppins-Regular.ttf"),
    "Poppins-Bold.ttf": os.path.join(D.ASET, "fonts", "Poppins-Bold.ttf"),
}
STATUS = {"belum": "Belum ditinjau", "ditinjau": "Sedang ditinjau", "revisi": "Perlu revisi",
          "disetujui": "Disetujui"}
_kunci_unduh = threading.Lock()
_kunci_bangun = threading.Lock()  # tema disimpan sebagai keadaan modul desain: bangun satu per satu
SIMPAN = penyimpanan.buka(os.path.join(AKAR, "data"))


class AsetGagal(RuntimeError):
    pass


class TemaTidakDikenal(ValueError):
    pass


# ------------------------------------------------------------ akses ---
@app.before_request
def periksa_akses():
    sandi = os.environ.get("MATERI_SANDI", "")
    if not sandi:
        if os.environ.get("VERCEL") and os.environ.get("MATERI_TANPA_SANDI") != "1":
            return Response("MATERI_SANDI belum diatur di Environment Variables proyek Vercel.",
                            503, mimetype="text/plain")
        return None
    auth = request.authorization
    if auth and auth.password and hmac.compare_digest(auth.password.encode(), sandi.encode()):
        return None
    return Response("Perlu sandi untuk membuka Materi Generator.", 401,
                    {"WWW-Authenticate": 'Basic realm="Materi Generator", charset="UTF-8"'},
                    mimetype="text/plain")


@app.after_request
def tanpa_cache(resp):
    resp.headers.setdefault("Cache-Control", "no-store")
    resp.headers["X-Content-Type-Options"] = "nosniff"
    return resp


# ------------------------------------------------------------- aset ---
def _unduh(url, tujuan, periksa):
    req = urllib.request.Request(url, headers={"User-Agent": "materi-generator"})
    token = os.environ.get("MATERI_ASET_TOKEN", "").strip()
    if token:
        req.add_header("Authorization", f"Bearer {token}")
    os.makedirs(CACHE, exist_ok=True)
    sementara = tujuan + ".part"
    try:
        with urllib.request.urlopen(req, timeout=30) as r, open(sementara, "wb") as f:
            shutil.copyfileobj(r, f)
        periksa(sementara)
    except (urllib.error.URLError, OSError, ValueError) as e:
        if os.path.exists(sementara):
            os.remove(sementara)
        raise AsetGagal(str(e)) from e
    os.replace(sementara, tujuan)


def _cek_pptx(path):
    if not zipfile.is_zipfile(path):
        raise ValueError("berkas hasil unduhan bukan PPTX (cek URL dan token)")


def path_template(tema="opsi-b"):
    lokal = TEMPLATE_LOKAL.get(tema)
    if lokal and os.path.exists(lokal):
        return lokal
    env = ENV_TEMPLATE[tema]
    url = os.environ.get(env, "").strip()
    if not url:
        raise BM.TemplateTidakAda(
            f"Template tampilan {NAMA_TEMA[tema]} belum diatur. Isi {env} dengan URL unduhan "
            f"'{BM.NAMA_TEMPLATE_TEMA[tema]}'.")
    # nama cache mengikuti URL, sehingga URL baru otomatis diunduh ulang
    tujuan = os.path.join(CACHE, hashlib.sha256(url.encode()).hexdigest()[:16] + ".pptx")
    with _kunci_unduh:
        if not os.path.exists(tujuan):
            try:
                _unduh(url, tujuan, _cek_pptx)
            except AsetGagal as e:
                raise AsetGagal(f"Template gagal diunduh dari {env}: {e}") from e
    return tujuan


# ------------------------------------------------------ katalog & data ---
def _baca_json(path):
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def katalog():
    return _baca_json(os.path.join(KONTEN, "katalog.json"))


def entri(id_materi):
    for m in katalog()["materi"]:
        if m["id"] == id_materi:
            return m
    abort(404, description="Materi tidak ada di katalog.")


def _kunci(id_materi, jenis):
    return f"materi:{id_materi}:{jenis}"


def _muat(nilai):
    try:
        return json.loads(nilai) if nilai else None
    except ValueError:
        return None


def ambil_data(id_materi):
    """(draf, status) dari penyimpanan; masing-masing dict atau None."""
    draf, status = SIMPAN.ambil_banyak([_kunci(id_materi, "naskah"), _kunci(id_materi, "status")])
    return _muat(draf), _muat(status)


def naskah_aktif(id_materi):
    """(naskah, draf_atau_None): draf tersimpan bila ada, selain itu naskah di repo."""
    m = entri(id_materi)
    draf, _ = ambil_data(id_materi)
    if draf and isinstance(draf.get("naskah"), dict):
        return draf["naskah"], draf
    return _baca_json(os.path.join(KONTEN, m["naskah"])), None


def ringkas_naskah(n):
    materi = n.get("materi") or []
    isi = [s for sub in materi for s in sub.get("slide", [])]
    perlu = []
    meta = n.get("meta") or {}
    tanpa_pembuka = meta.get("bagian") == "soal" or meta.get("lewati_kosong")  # tidak punya pendahuluan sendiri
    if not tanpa_pembuka and not (n.get("pendahuluan") or {}).get("pemantik"):
        perlu.append("pertanyaan pemantik")
    if any(not b.get("level") for b in (n.get("tujuan") or {}).get("butir", [])):
        perlu.append("level Bloom")
    if any(not r.get("teks") for r in (n.get("penutup") or {}).get("ringkasan", [])):
        perlu.append("teks ringkasan penutup")
    return {"judul": (n.get("meta") or {}).get("judul", ""), "jenis": BM.jenis_deck(n),
            "slide_isi": sum(1 for s in isi if s.get("tipe") != "contoh_soal"),
            "contoh_soal": sum(1 for s in isi if s.get("tipe") == "contoh_soal"),
            "perlu_dilengkapi": perlu}


def _sekarang():
    return datetime.datetime.now(datetime.timezone.utc).isoformat(timespec="seconds")


def _oleh():
    return str((request.get_json(silent=True) or {}).get("oleh") or "").strip()[:60]


# ----------------------------------------------------------- respons ---
def _galat(pesan, kode):
    return jsonify({"galat": pesan}), kode


def _nama_aman(nama):
    nama = re.sub(r'[\\/:*?"<>|\x00-\x1f]+', "-", str(nama)).strip(" .-")
    return nama or "materi"


def _tema():
    """Tema dari ?tema= atau body JSON {"tema": ...}; bawaan opsi-b."""
    data = request.get_json(silent=True) if request.is_json else None
    tema = request.args.get("tema") or (data or {}).get("tema") or request.form.get("tema") or "opsi-b"
    if tema not in BM.TEMA:
        raise TemaTidakDikenal(f"Tema tidak dikenal: {tema} (pilihan: {', '.join(BM.TEMA)}).")
    return tema


def _bangun(konten, tema="opsi-b", folder=KONTEN):
    """(prs, pembangun) atau lempar galat yang sudah berpesan. Deck dibangun satu per satu."""
    if not isinstance(konten, dict) or not isinstance(konten.get("meta"), dict):
        raise ValueError("Naskah harus objek JSON dengan bagian 'meta' (lihat contoh naskah).")
    template = path_template(tema)
    with _kunci_bangun:
        return BM.buat_deck(konten, template, folder, tema)


def _tangani(fungsi):
    """Jalankan fungsi pembuat deck; ubah galat umum menjadi respons JSON berpesan."""
    try:
        return fungsi()
    except (BM.TemplateTidakAda, BM.TemplateTidakCocok) as e:
        return _galat(str(e), 503)
    except (AsetGagal, penyimpanan.PenyimpananGagal) as e:
        return _galat(str(e), 502)
    except TemaTidakDikenal as e:
        return _galat(str(e), 400)
    except KeyError as e:
        return _galat(f"Naskah tidak lengkap: kunci {e} tidak ditemukan.", 400)
    except (IndexError, TypeError, ValueError, AttributeError) as e:
        return _galat(f"Naskah tidak sesuai struktur ({type(e).__name__}: {e}).", 400)
    except Exception:  # noqa: BLE001
        app.logger.exception("Gagal membuat deck")
        return _galat("Terjadi galat saat membuat deck. Cek log fungsi di Vercel.", 500)


def _kirim_pptx(prs, b, konten, tema="opsi-b"):
    buf = io.BytesIO()
    with _kunci_bangun:
        prs.save(buf)
    ukuran = buf.tell()
    if ukuran > BATAS_RESPONS:
        return _galat(f"Deck {ukuran / 1e6:.1f} MB melebihi batas unduhan Vercel 4,5 MB. "
                      "Perkecil gambar atau buat deck di komputer lokal.", 413)
    buf.seek(0)
    resp = send_file(
        buf, as_attachment=True,
        download_name=_nama_aman(BM.nama_file(konten)) + (" (General)" if tema == "general" else "") + ".pptx",
        mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation")
    resp.headers["X-Jumlah-Slide"] = str(len(b.slides))
    resp.headers["X-Peringatan"] = urllib.parse.quote(json.dumps(b.peringatan, ensure_ascii=False))
    return resp


# ------------------------------------------------------------- halaman ---
@app.get("/")
def dasbor():
    return send_file(os.path.join(WEB, "dasbor.html"), mimetype="text/html")


@app.get("/materi/<id_materi>")
def halaman_editor(id_materi):
    entri(id_materi)
    return send_file(os.path.join(WEB, "editor.html"), mimetype="text/html")


@app.get("/unggah")
def halaman_unggah():
    return send_file(os.path.join(WEB, "unggah.html"), mimetype="text/html")


@app.get("/web/<nama>")
def berkas_web(nama):
    if nama not in ("gaya.css", "pratinjau.js"):
        abort(404)
    resp = send_from_directory(WEB, nama)
    resp.headers["Cache-Control"] = "private, max-age=300"
    return resp


@app.get("/font/<nama>")
def font(nama):
    path = FONT_WEB.get(nama)
    if not path:
        return _galat("Font tidak ada.", 404)
    resp = send_file(path, mimetype="font/ttf", max_age=86400)
    resp.headers["Cache-Control"] = "private, max-age=86400"
    return resp


@app.get("/media/<path:berkas>")
def media(berkas):
    return send_from_directory(os.path.join(KONTEN, "media"), berkas)


# ----------------------------------------------------------------- API ---
@app.get("/api/status")
def status():
    hasil = {"pdf": False, "penyimpanan": SIMPAN.jenis, "bisa_simpan": SIMPAN.bisa_tulis, "tema": {}}
    for tema in BM.TEMA:
        try:
            path_template(tema)
            hasil["tema"][tema] = {"nama": NAMA_TEMA[tema], "siap": True}
        except (BM.TemplateTidakAda, AsetGagal) as e:
            hasil["tema"][tema] = {"nama": NAMA_TEMA[tema], "siap": False, "pesan": str(e)}
    hasil["template"] = hasil["tema"]["opsi-b"]
    return jsonify(hasil)


@app.get("/api/katalog")
def api_katalog():
    kat = katalog()
    kunci = [_kunci(m["id"], j) for m in kat["materi"] for j in ("naskah", "status")]
    try:
        nilai = SIMPAN.ambil_banyak(kunci)
        galat = None
    except penyimpanan.PenyimpananGagal as e:
        nilai, galat = [None] * len(kunci), str(e)
    daftar = []
    for i, m in enumerate(kat["materi"]):
        draf, stat = _muat(nilai[2 * i]), _muat(nilai[2 * i + 1])
        naskah = draf["naskah"] if draf and isinstance(draf.get("naskah"), dict) else \
            _baca_json(os.path.join(KONTEN, m["naskah"]))
        daftar.append(dict(m, url=m.get("url") or f"https://docs.google.com/document/d/{m['doc_id']}/edit",
                           ringkas=ringkas_naskah(naskah),
                           draf={"diubah": draf.get("diubah"), "oleh": draf.get("oleh")} if draf else None,
                           status=stat or {"status": "belum"}))
    return jsonify({"judul": kat.get("judul"), "sumber_sheet": kat.get("sumber_sheet"), "materi": daftar,
                    "status_label": STATUS, "bisa_simpan": SIMPAN.bisa_tulis, "penyimpanan": SIMPAN.jenis,
                    "galat_penyimpanan": galat})


@app.get("/api/materi/<id_materi>")
def api_materi(id_materi):
    m = entri(id_materi)
    try:
        naskah, draf = naskah_aktif(id_materi)
        _, stat = ambil_data(id_materi)
    except penyimpanan.PenyimpananGagal as e:
        return _galat(str(e), 502)
    return jsonify({"katalog": dict(m, url=m.get("url") or f"https://docs.google.com/document/d/{m['doc_id']}/edit"),
                    "naskah": naskah, "draf": {"diubah": draf["diubah"], "oleh": draf.get("oleh")} if draf else None,
                    "status": stat or {"status": "belum"}, "status_label": STATUS,
                    "bisa_simpan": SIMPAN.bisa_tulis})


@app.put("/api/materi/<id_materi>/naskah")
def simpan_naskah(id_materi):
    entri(id_materi)
    data = request.get_json(silent=True) or {}
    naskah = data.get("naskah")
    if not isinstance(naskah, dict) or not isinstance(naskah.get("meta"), dict) \
            or not isinstance(naskah.get("materi"), list):
        return _galat("Naskah tidak valid: perlu bagian 'meta' dan 'materi'.", 400)
    try:
        lama, _ = ambil_data(id_materi)
        dasar = data.get("versi_dasar")
        if lama and dasar is not None and lama.get("diubah") != dasar:
            return _galat(f"Draf sudah diubah oleh {lama.get('oleh') or 'orang lain'} pada {lama.get('diubah')}. "
                          "Muat ulang halaman sebelum menyimpan agar suntingan mereka tidak tertimpa.", 409)
        baru = {"naskah": naskah, "diubah": _sekarang(), "oleh": _oleh()}
        SIMPAN.simpan(_kunci(id_materi, "naskah"), json.dumps(baru, ensure_ascii=False))
    except penyimpanan.PenyimpananGagal as e:
        return _galat(str(e), 503)
    return jsonify({"diubah": baru["diubah"], "oleh": baru["oleh"]})


@app.delete("/api/materi/<id_materi>/naskah")
def buang_draf(id_materi):
    entri(id_materi)
    try:
        SIMPAN.hapus(_kunci(id_materi, "naskah"))
    except penyimpanan.PenyimpananGagal as e:
        return _galat(str(e), 503)
    return jsonify({"ok": True})


@app.put("/api/materi/<id_materi>/status")
def simpan_status(id_materi):
    entri(id_materi)
    data = request.get_json(silent=True) or {}
    if data.get("status") not in STATUS:
        return _galat("Status tidak dikenal.", 400)
    nilai = {"status": data["status"], "peninjau": str(data.get("peninjau") or "").strip()[:60],
             "catatan": str(data.get("catatan") or "").strip()[:2000], "diubah": _sekarang(), "oleh": _oleh()}
    try:
        SIMPAN.simpan(_kunci(id_materi, "status"), json.dumps(nilai, ensure_ascii=False))
    except penyimpanan.PenyimpananGagal as e:
        return _galat(str(e), 503)
    return jsonify(nilai)


@app.post("/api/materi/<id_materi>/pratinjau")
def pratinjau(id_materi):
    entri(id_materi)
    data = request.get_json(silent=True) or {}

    def kerja():
        konten = data.get("naskah") if isinstance(data.get("naskah"), dict) else naskah_aktif(id_materi)[0]
        prs, b = _bangun(konten, _tema())
        with _kunci_bangun:
            hasil = PR.dump_deck(prs, b.asal)
        punya = set(data.get("punya") or [])
        hasil["gambar"] = {k: v for k, v in hasil["gambar"].items() if k not in punya}
        hasil["peringatan"] = b.peringatan
        return jsonify(hasil)

    return _tangani(kerja)


@app.route("/api/materi/<id_materi>/pptx", methods=["GET", "POST"])
def pptx_materi(id_materi):
    entri(id_materi)
    data = request.get_json(silent=True) or {}

    def kerja():
        konten = data.get("naskah") if isinstance(data.get("naskah"), dict) else naskah_aktif(id_materi)[0]
        tema = _tema()
        prs, b = _bangun(konten, tema)
        return _kirim_pptx(prs, b, konten, tema)

    return _tangani(kerja)


@app.get("/api/contoh")
def contoh():
    return send_file(CONTOH, mimetype="application/json", as_attachment=True,
                     download_name="contoh_naskah_materi.json")


def _baca_unggahan():
    berkas = request.files.get("naskah")
    try:
        if berkas is not None:
            return json.loads(berkas.read().decode("utf-8-sig"))
        if request.is_json:
            return json.loads(request.get_data().decode("utf-8-sig"))
    except UnicodeDecodeError:
        raise ValueError("Berkas naskah harus teks UTF-8.")
    except json.JSONDecodeError as e:
        raise ValueError(f"Naskah bukan JSON yang valid (baris {e.lineno}, kolom {e.colno}): {e.msg}.")
    raise ValueError("Kirim naskah sebagai berkas 'naskah' atau body application/json.")


@app.post("/api/unggah")
def unggah():
    try:
        konten = _baca_unggahan()
    except ValueError as e:
        return _galat(str(e), 400)

    def kerja():
        tema = _tema()
        prs, b = _bangun(konten, tema)
        return _kirim_pptx(prs, b, konten, tema)

    return _tangani(kerja)


@app.errorhandler(404)
def tidak_ada(e):
    if request.path.startswith("/api/"):
        return _galat(getattr(e, "description", "Tidak ditemukan."), 404)
    return Response("Halaman tidak ditemukan.", 404, mimetype="text/plain")


@app.errorhandler(413)
def terlalu_besar(_e):
    return _galat("Berkas atau naskah melebihi 4 MB.", 413)


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description="Jalankan Materi Generator web secara lokal.")
    ap.add_argument("--template", default=TEMPLATE_LOKAL["opsi-b"], help="Template Opsi B")
    ap.add_argument("--template-general", default=TEMPLATE_LOKAL["general"], help="Template Materi General")
    ap.add_argument("--port", type=int, default=5000)
    a = ap.parse_args()
    TEMPLATE_LOKAL.update({"opsi-b": a.template, "general": a.template_general})
    app.run(host="127.0.0.1", port=a.port)
