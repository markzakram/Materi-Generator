// Gambar ulang slide dari data bentuk (generator/pratinjau.py) di browser.
(function () {
  const IN = 96; // px per inci sebelum diskalakan
  const PT = IN / 72;
  const RUMUS = "'Cambria Math', 'STIX Two Math', 'Latin Modern Math', 'Times New Roman', serif";
  const SVG = "http://www.w3.org/2000/svg";
  let nomorId = 0;

  const rgba = (c, a = 1) => {
    const n = parseInt(c, 16);
    return `rgba(${(n >> 16) & 255}, ${(n >> 8) & 255}, ${n & 255}, ${a})`;
  };

  function teks(tx) {
    const d = document.createElement("div");
    d.className = "tx";
    const [l, t, r, b] = tx.in;
    d.style.padding = `${t * IN}px ${r * IN}px ${b * IN}px ${l * IN}px`;
    d.style.justifyContent = tx.an === "ctr" ? "center" : tx.an === "b" ? "flex-end" : "flex-start";
    tx.p.forEach((p, i) => {
      const el = document.createElement("p");
      el.style.textAlign = { c: "center", r: "right", j: "justify" }[p.al] || "left";
      const besar = Math.max(...p.r.map((r) => r.pt), 10);
      el.style.lineHeight = p.ls ? `${p.ls * PT}px` : p.lp ? String(p.lp * 1.2) : "1.2";
      el.style.fontSize = `${besar * PT}px`;
      if (i < tx.p.length - 1 && p.sa) el.style.marginBottom = `${p.sa * PT}px`;
      for (const r of p.r) {
        const s = document.createElement("span");
        s.textContent = r.t;
        s.style.fontSize = `${r.pt * PT}px`;
        s.style.color = "#" + r.c;
        if (r.b || /bold/i.test(r.f || "")) s.style.fontWeight = "700";
        if (r.i) s.style.fontStyle = "italic";
        if (r.s) s.style.textDecoration = "line-through";
        s.style.fontFamily = r.f === "Cambria Math" ? RUMUS : "'Poppins', sans-serif";
        if (r.v) { s.style.verticalAlign = r.v; s.style.fontSize = `${r.pt * PT * 0.67}px`; s.style.lineHeight = "0"; }
        el.appendChild(s);
      }
      if (!p.r.length) el.innerHTML = "&nbsp;";
      d.appendChild(el);
    });
    return d;
  }

  // isian untuk div biasa (latar CSS)
  function isiCss(d, f, gambar) {
    if (!f) return;
    if (f.t === "solid") d.style.background = rgba(f.c, f.a);
    else if (f.t === "grad") d.style.background = `linear-gradient(${f.ang + 90}deg, ${f.s.map(([p, c, a]) => `${rgba(c, a)} ${p}%`).join(", ")})`;
    else if (f.t === "img" && gambar[f.src]) { d.style.backgroundImage = `url("${gambar[f.src]}")`; d.style.backgroundSize = "100% 100%"; }
  }

  // bentuk bebas: path SVG dengan isian warna, gradien, atau gambar
  function bebas(b, gambar) {
    const svg = document.createElementNS(SVG, "svg");
    svg.setAttribute("viewBox", `0 0 ${b.geo.w} ${b.geo.h}`);
    svg.setAttribute("preserveAspectRatio", "none");
    Object.assign(svg.style, { position: "absolute", inset: "0", width: "100%", height: "100%", overflow: "visible" });
    const defs = document.createElementNS(SVG, "defs");
    const path = document.createElementNS(SVG, "path");
    path.setAttribute("d", b.geo.d);
    const f = b.fill;
    if (!f) path.setAttribute("fill", "none");
    else if (f.t === "solid") { path.setAttribute("fill", "#" + f.c); path.setAttribute("fill-opacity", f.a); }
    else if (f.t === "grad") {
      const id = "g" + (++nomorId);
      const g = document.createElementNS(SVG, "linearGradient");
      g.id = id;
      const rad = (f.ang * Math.PI) / 180, dx = Math.cos(rad) / 2, dy = Math.sin(rad) / 2;
      Object.entries({ x1: 0.5 - dx, y1: 0.5 - dy, x2: 0.5 + dx, y2: 0.5 + dy }).forEach(([k, v]) => g.setAttribute(k, v));
      f.s.forEach(([p, c, a]) => { const s = document.createElementNS(SVG, "stop"); s.setAttribute("offset", p + "%"); s.setAttribute("stop-color", "#" + c); s.setAttribute("stop-opacity", a); g.appendChild(s); });
      defs.appendChild(g); path.setAttribute("fill", `url(#${id})`);
    } else if (f.t === "img" && gambar[f.src]) {
      const id = "p" + (++nomorId);
      const pat = document.createElementNS(SVG, "pattern");
      pat.id = id; pat.setAttribute("patternUnits", "userSpaceOnUse");
      pat.setAttribute("width", b.geo.w); pat.setAttribute("height", b.geo.h);
      const img = document.createElementNS(SVG, "image");
      img.setAttribute("href", gambar[f.src]); img.setAttribute("width", b.geo.w); img.setAttribute("height", b.geo.h);
      img.setAttribute("preserveAspectRatio", "none");
      pat.appendChild(img); defs.appendChild(pat); path.setAttribute("fill", `url(#${id})`);
    }
    if (b.ln) { path.setAttribute("stroke", "#" + b.ln.c); path.setAttribute("stroke-width", b.ln.w * PT); path.setAttribute("vector-effect", "non-scaling-stroke"); }
    svg.append(defs, path);
    return svg;
  }

  function bentuk(b, gambar) {
    const d = document.createElement("div");
    d.className = "b";
    Object.assign(d.style, { left: `${b.x * IN}px`, top: `${b.y * IN}px`, width: `${b.w * IN}px`, height: `${b.h * IN}px` });
    const ubah = [];
    if (b.rot) ubah.push(`rotate(${b.rot}deg)`);
    if (b.fh) ubah.push("scaleX(-1)");
    if (b.fv) ubah.push("scaleY(-1)");
    if (ubah.length) d.style.transform = ubah.join(" ");
    if (b.k === "img") {
      const img = new Image();
      img.alt = ""; img.decoding = "async";
      if (gambar[b.src]) img.src = gambar[b.src];
      d.appendChild(img);
      return d;
    }
    if (b.k === "tbl") {
      const t = document.createElement("table");
      t.className = "tbl";
      t.style.width = `${b.cw.reduce((a, c) => a + c, 0) * IN}px`;
      const cg = document.createElement("colgroup");
      b.cw.forEach((w) => { const c = document.createElement("col"); c.style.width = `${w * IN}px`; cg.appendChild(c); });
      t.appendChild(cg);
      b.c.forEach((baris, i) => {
        const tr = document.createElement("tr");
        tr.style.height = `${b.rh[i] * IN}px`;
        baris.forEach((sel) => {
          const td = document.createElement("td");
          if (sel.bg) td.style.background = "#" + sel.bg;
          if (sel.tx) td.appendChild(teks(sel.tx));
          tr.appendChild(td);
        });
        t.appendChild(tr);
      });
      d.appendChild(t);
      return d;
    }
    if (b.geo) {
      d.appendChild(bebas(b, gambar));
    } else {
      if (b.fill) isiCss(d, b.fill, gambar); else if (b.bg) d.style.background = "#" + b.bg;
      if (b.ln) d.style.border = `${Math.max(1, b.ln.w * PT)}px solid #${b.ln.c}`;
      if (b.r === "oval") d.style.borderRadius = "50%";
      else if (b.r === "panah") d.style.clipPath = "polygon(0 25%, 50% 25%, 50% 0, 100% 50%, 50% 100%, 50% 75%, 0 75%)";
      else if (typeof b.r === "number") d.style.borderRadius = `${b.r * IN}px`;
    }
    if (b.tx) d.appendChild(teks(b.tx));
    return d;
  }

  // Gambar satu slide ke dalam wadah; lebar wadah menentukan skala.
  function gambarSlide(wadah, data, i, gambar) {
    const sl = data.slide[i];
    const lebar = wadah.clientWidth || 320;
    const skala = lebar / (data.lebar * IN);
    wadah.style.height = `${data.tinggi * IN * skala}px`;
    wadah.replaceChildren();
    const s = document.createElement("div");
    s.className = "slide";
    Object.assign(s.style, { width: `${data.lebar * IN}px`, height: `${data.tinggi * IN}px`,
      transform: `scale(${skala})`, background: "#" + sl.bg });
    if (sl.bgimg && gambar[sl.bgimg]) { s.style.backgroundImage = `url("${gambar[sl.bgimg]}")`; s.style.backgroundSize = "100% 100%"; }
    for (const b of sl.b) s.appendChild(bentuk(b, gambar));
    wadah.appendChild(s);
  }

  window.Pratinjau = { gambarSlide };
})();
