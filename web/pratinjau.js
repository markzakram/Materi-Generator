// Gambar ulang slide dari data bentuk (generator/pratinjau.py) di browser.
(function () {
  const IN = 96; // px per inci sebelum diskalakan
  const PT = IN / 72;
  const RUMUS = "'Cambria Math', 'STIX Two Math', 'Latin Modern Math', 'Times New Roman', serif";

  function teks(tx) {
    const d = document.createElement("div");
    d.className = "tx";
    const [l, t, r, b] = tx.in;
    d.style.padding = `${t * IN}px ${r * IN}px ${b * IN}px ${l * IN}px`;
    d.style.justifyContent = tx.an === "ctr" ? "center" : tx.an === "b" ? "flex-end" : "flex-start";
    tx.p.forEach((p, i) => {
      const el = document.createElement("p");
      el.style.textAlign = p.al === "c" ? "center" : p.al === "r" ? "right" : "left";
      const besar = Math.max(...p.r.map((r) => r.pt), 10);
      el.style.lineHeight = p.ls ? `${p.ls * PT}px` : "1.2";
      el.style.fontSize = `${besar * PT}px`;
      if (i < tx.p.length - 1 && p.sa) el.style.marginBottom = `${p.sa * PT}px`;
      for (const r of p.r) {
        const s = document.createElement("span");
        s.textContent = r.t;
        s.style.fontSize = `${r.pt * PT}px`;
        s.style.color = "#" + r.c;
        if (r.b) s.style.fontWeight = "700";
        s.style.fontFamily = r.f === "Cambria Math" ? RUMUS : "'Poppins', sans-serif";
        el.appendChild(s);
      }
      if (!p.r.length) el.innerHTML = "&nbsp;";
      d.appendChild(el);
    });
    return d;
  }

  function bentuk(b, gambar) {
    const d = document.createElement("div");
    d.className = "b";
    Object.assign(d.style, { left: `${b.x * IN}px`, top: `${b.y * IN}px`, width: `${b.w * IN}px`, height: `${b.h * IN}px` });
    if (b.k === "img") {
      const img = new Image();
      img.alt = "";
      img.decoding = "async";
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
    if (b.bg) d.style.background = "#" + b.bg;
    if (b.r === "oval") d.style.borderRadius = "50%";
    else if (b.r === "panah") d.style.clipPath = "polygon(0 25%, 50% 25%, 50% 0, 100% 50%, 50% 100%, 50% 75%, 0 75%)";
    else if (typeof b.r === "number") d.style.borderRadius = `${b.r * IN}px`;
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
    for (const b of sl.b) s.appendChild(bentuk(b, gambar));
    wadah.appendChild(s);
  }

  window.Pratinjau = { gambarSlide };
})();
