"""Aplikasi web Materi Generator (Flask), siap di-deploy ke Vercel.

Halaman unggah naskah materi (JSON) lalu unduh deck PPTX. Ekspor PDF tidak
tersedia di sini karena butuh Microsoft PowerPoint; buat PDF di Windows lokal
dengan generator/buat_materi.py.

Variabel lingkungan:
  MATERI_SANDI          Sandi akses (HTTP Basic, nama pengguna bebas). Di Vercel
                        wajib diisi, kecuali MATERI_TANPA_SANDI=1.
  MATERI_TANPA_SANDI    Isi 1 hanya jika akses sudah dilindungi Vercel Authentication.
  MATERI_TEMPLATE_URL   URL unduhan 'Template Desain Materi.pptx'. Wajib jika berkas
                        itu tidak ada di akar proyek (selalu begitu di Vercel).
  MATERI_COVER_URL      URL unduhan foto latar sampul (opsional).
  MATERI_ASET_TOKEN     Token yang dikirim sebagai 'Authorization: Bearer ...' saat
                        mengunduh dua URL di atas (opsional, untuk penyimpanan privat).

Jalankan lokal:
  python app.py --template "path/ke/Template Desain Materi.pptx"
"""
from __future__ import annotations

import argparse
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

from flask import Flask, Response, jsonify, request, send_file

AKAR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(AKAR, "generator"))
import buat_materi as BM  # noqa: E402
import desain as D  # noqa: E402

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 4 * 1024 * 1024  # Vercel menolak request di atas 4,5 MB

BATAS_RESPONS = 4_400_000  # byte; Vercel menolak response di atas 4,5 MB
CACHE = os.path.join(tempfile.gettempdir(), "materi-generator")
TEMPLATE_LOKAL = BM.TEMPLATE
COVER_LOKAL = D.LATAR_COVER
CONTOH = os.path.join(AKAR, "konten", "tiu_penalaran_logis.json")
FONT_WEB = {
    "Poppins-Regular.ttf": os.path.join(D.ASET, "fonts", "Poppins-Regular.ttf"),
    "Poppins-Bold.ttf": os.path.join(D.ASET, "fonts", "Poppins-Bold.ttf"),
}
_kunci_unduh = threading.Lock()


class AsetGagal(RuntimeError):
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


def _aset(lokal, env, akhiran, periksa):
    """Path aset: berkas lokal jika ada, jika tidak unduh dari URL di env (di-cache di /tmp)."""
    if lokal and os.path.exists(lokal):
        return lokal
    url = os.environ.get(env, "").strip()
    if not url:
        return None
    # nama cache mengikuti URL, sehingga URL baru otomatis diunduh ulang
    tujuan = os.path.join(CACHE, hashlib.sha256(url.encode()).hexdigest()[:16] + akhiran)
    with _kunci_unduh:
        if not os.path.exists(tujuan):
            _unduh(url, tujuan, periksa)
    return tujuan


def _cek_pptx(path):
    if not zipfile.is_zipfile(path):
        raise ValueError("berkas hasil unduhan bukan PPTX (cek URL dan token)")


def _cek_gambar(path):
    from PIL import Image

    with Image.open(path) as im:
        im.verify()


def path_template():
    try:
        p = _aset(TEMPLATE_LOKAL, "MATERI_TEMPLATE_URL", ".pptx", _cek_pptx)
    except AsetGagal as e:
        raise AsetGagal(f"Template gagal diunduh dari MATERI_TEMPLATE_URL: {e}") from e
    if not p:
        raise BM.TemplateTidakAda(
            "Template belum diatur. Isi MATERI_TEMPLATE_URL dengan URL unduhan "
            "'Template Desain Materi.pptx'.")
    return p


def path_cover(peringatan):
    try:
        return _aset(COVER_LOKAL, "MATERI_COVER_URL", ".img", _cek_gambar)
    except AsetGagal as e:
        peringatan.append(f"Foto sampul gagal diunduh ({e}); sampul dibuat tanpa foto.")
        return None


# ----------------------------------------------------------- respons ---
def _galat(pesan, kode):
    return jsonify({"galat": pesan}), kode


def _baca_naskah():
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


def _nama_aman(nama):
    nama = re.sub(r'[\\/:*?"<>|\x00-\x1f]+', "-", str(nama)).strip(" .-")
    return nama or "materi"


# ------------------------------------------------------------ rute ---
@app.get("/")
def beranda():
    return send_file(os.path.join(AKAR, "web", "index.html"), mimetype="text/html")


@app.get("/font/<nama>")
def font(nama):
    path = FONT_WEB.get(nama)
    if not path:
        return _galat("Font tidak ada.", 404)
    resp = send_file(path, mimetype="font/ttf", max_age=86400)
    resp.headers["Cache-Control"] = "private, max-age=86400"
    return resp


@app.get("/api/contoh")
def contoh():
    return send_file(CONTOH, mimetype="application/json", as_attachment=True,
                     download_name="contoh_naskah_materi.json")


@app.get("/api/status")
def status():
    hasil = {"pdf": False}
    try:
        path_template()
        hasil["template"] = {"siap": True}
    except (BM.TemplateTidakAda, AsetGagal) as e:
        hasil["template"] = {"siap": False, "pesan": str(e)}
    cover = []
    hasil["cover"] = bool(path_cover(cover))
    if cover:
        hasil["cover_pesan"] = cover[0]
    return jsonify(hasil)


@app.post("/api/materi")
def materi():
    try:
        konten = _baca_naskah()
    except ValueError as e:
        return _galat(str(e), 400)
    if not isinstance(konten, dict) or not isinstance(konten.get("meta"), dict):
        return _galat("Naskah harus objek JSON dengan bagian 'meta' (lihat contoh naskah).", 400)

    peringatan = []
    try:
        template = path_template()
        D.LATAR_COVER = path_cover(peringatan)
        prs, b = BM.buat_deck(konten, template)
        buf = io.BytesIO()
        prs.save(buf)
    except BM.TemplateTidakAda as e:
        return _galat(str(e), 503)
    except AsetGagal as e:
        return _galat(str(e), 502)
    except KeyError as e:
        return _galat(f"Naskah tidak lengkap: kunci {e} tidak ditemukan.", 400)
    except (IndexError, TypeError, ValueError, AttributeError) as e:
        return _galat(f"Naskah tidak sesuai struktur ({type(e).__name__}: {e}).", 400)
    except Exception:  # noqa: BLE001
        app.logger.exception("Gagal membuat deck")
        return _galat("Terjadi galat saat membuat deck. Cek log fungsi di Vercel.", 500)

    ukuran = buf.tell()
    if ukuran > BATAS_RESPONS:
        return _galat(f"Deck {ukuran / 1e6:.1f} MB melebihi batas unduhan Vercel 4,5 MB. "
                      "Perkecil foto sampul atau buat deck di komputer lokal.", 413)
    buf.seek(0)
    peringatan += b.peringatan
    resp = send_file(
        buf, as_attachment=True, download_name=_nama_aman(BM.nama_file(konten)) + ".pptx",
        mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation")
    resp.headers["X-Jumlah-Slide"] = str(len(b.slides))
    resp.headers["X-Peringatan"] = urllib.parse.quote(json.dumps(peringatan, ensure_ascii=False))
    return resp


@app.errorhandler(413)
def terlalu_besar(_e):
    return _galat("Berkas naskah melebihi 4 MB.", 413)


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description="Jalankan Materi Generator web secara lokal.")
    ap.add_argument("--template", default=TEMPLATE_LOKAL)
    ap.add_argument("--port", type=int, default=5000)
    a = ap.parse_args()
    TEMPLATE_LOKAL = a.template
    app.run(host="127.0.0.1", port=a.port)
