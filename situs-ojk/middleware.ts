// Vercel Routing Middleware: semua halaman dan berkas unduhan situs ini dikunci dengan sandi
// (HTTP Basic, nama pengguna bebas), sama seperti aplikasi generator.
// Sandi diatur di env variable MATERI_SANDI proyek Vercel. Tanpa sandi itu situs menolak semua
// akses (503) supaya tidak pernah terbuka tanpa sengaja.
// Hanya berjalan di Vercel; server lokal (python -m http.server) tidak memakainya.

// Sama dengan next() dari @vercel/functions: lanjutkan ke berkas statis yang diminta.
const LANJUT = { "x-middleware-next": "1" };
const TEKS = "text/plain; charset=utf-8";

function samaPanjangTetap(a: Uint8Array, b: Uint8Array): boolean {
  let beda = a.length ^ b.length;
  for (let i = 0; i < Math.max(a.length, b.length); i++) beda |= (a[i] ?? 0) ^ (b[i] ?? 0);
  return beda === 0;
}

function sandiDariHeader(auth: string | null): Uint8Array | null {
  const cocok = /^Basic\s+(\S+)$/i.exec(auth || "");
  if (!cocok) return null;
  let biner: string;
  try {
    biner = atob(cocok[1]);
  } catch {
    return null;
  }
  const titikDua = biner.indexOf(":");
  if (titikDua < 0) return null;
  // atob menghasilkan byte mentah; sandi dikirim browser sebagai UTF-8 (charset="UTF-8").
  return Uint8Array.from(biner.slice(titikDua + 1), (c) => c.charCodeAt(0));
}

export default function middleware(request: Request): Response {
  const sandi = process.env.MATERI_SANDI;
  if (!sandi) {
    return new Response("MATERI_SANDI belum diatur di Environment Variables proyek Vercel.", {
      status: 503,
      headers: { "content-type": TEKS, "cache-control": "no-store" },
    });
  }
  const masukan = sandiDariHeader(request.headers.get("authorization"));
  if (masukan && samaPanjangTetap(masukan, new TextEncoder().encode(sandi))) {
    return new Response(null, { headers: LANJUT });
  }
  return new Response("Perlu sandi untuk membuka materi.", {
    status: 401,
    headers: {
      "www-authenticate": 'Basic realm="Materi", charset="UTF-8"',
      "content-type": TEKS,
      "cache-control": "no-store",
    },
  });
}
