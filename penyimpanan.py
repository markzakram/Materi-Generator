"""Penyimpanan draf naskah dan status tinjauan untuk dasbor.

- Vercel: Upstash Redis lewat REST API. Variabel lingkungan diisi otomatis saat database
  Upstash disambungkan ke proyek (KV_REST_API_URL/KV_REST_API_TOKEN atau
  UPSTASH_REDIS_REST_URL/UPSTASH_REDIS_REST_TOKEN).
- Lokal: berkas JSON di folder data/ (diabaikan git).
- Vercel tanpa Upstash: mode baca saja (dasbor dan unduhan jalan, simpan ditolak).
"""
from __future__ import annotations

import json
import os
import re
import threading
import urllib.error
import urllib.request


class PenyimpananGagal(RuntimeError):
    pass


class _Upstash:
    jenis = "Upstash Redis"
    bisa_tulis = True

    def __init__(self, url, token):
        self.url, self.token = url.rstrip("/"), token

    def _perintah(self, *args):
        req = urllib.request.Request(self.url, data=json.dumps(list(args)).encode(), method="POST",
                                     headers={"Authorization": f"Bearer {self.token}",
                                              "Content-Type": "application/json"})
        try:
            with urllib.request.urlopen(req, timeout=10) as r:
                hasil = json.loads(r.read().decode())
        except (urllib.error.URLError, OSError, ValueError) as e:
            raise PenyimpananGagal(f"Upstash tidak bisa dihubungi: {e}") from e
        if "error" in hasil:
            raise PenyimpananGagal(f"Upstash menolak perintah: {hasil['error']}")
        return hasil.get("result")

    def ambil_banyak(self, kunci):
        if not kunci:
            return []
        return self._perintah("MGET", *kunci)

    def simpan(self, kunci, nilai):
        self._perintah("SET", kunci, nilai)

    def hapus(self, kunci):
        self._perintah("DEL", kunci)


class _Berkas:
    jenis = "berkas lokal"
    bisa_tulis = True

    def __init__(self, folder):
        self.folder = folder
        self._kunci = threading.Lock()

    def _path(self, kunci):
        return os.path.join(self.folder, re.sub(r"[^A-Za-z0-9_.-]", "_", kunci) + ".json")

    def ambil_banyak(self, kunci):
        hasil = []
        for k in kunci:
            try:
                with open(self._path(k), encoding="utf-8") as f:
                    hasil.append(f.read())
            except FileNotFoundError:
                hasil.append(None)
        return hasil

    def simpan(self, kunci, nilai):
        with self._kunci:
            os.makedirs(self.folder, exist_ok=True)
            sementara = self._path(kunci) + ".part"
            with open(sementara, "w", encoding="utf-8") as f:
                f.write(nilai)
            os.replace(sementara, self._path(kunci))

    def hapus(self, kunci):
        try:
            os.remove(self._path(kunci))
        except FileNotFoundError:
            pass


class _BacaSaja:
    jenis = "belum disambungkan"
    bisa_tulis = False

    def ambil_banyak(self, kunci):
        return [None] * len(kunci)

    def simpan(self, kunci, nilai):
        raise PenyimpananGagal("Penyimpanan belum disambungkan. Sambungkan database Upstash Redis "
                               "ke proyek Vercel (Storage > Upstash for Redis), lalu deploy ulang.")

    hapus = simpan


def buka(folder_lokal):
    url = os.environ.get("KV_REST_API_URL") or os.environ.get("UPSTASH_REDIS_REST_URL")
    token = os.environ.get("KV_REST_API_TOKEN") or os.environ.get("UPSTASH_REDIS_REST_TOKEN")
    if url and token:
        return _Upstash(url, token)
    if os.environ.get("VERCEL"):
        return _BacaSaja()  # sistem berkas Vercel tidak permanen
    return _Berkas(folder_lokal)
