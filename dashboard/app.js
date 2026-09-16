/* =========================================================
   ALEL Production — Live BI Dashboard (v2)
   ========================================================= */
(function () {
  "use strict";

  const SHEET_ID = "1HeKJ0XueH0WeAxLrJpNaNz_pvPuxRDElQmHcPd7QfOQ";
  const REFRESH_MS = 30000;
  const PALETTE = ["#22d3ee", "#6366f1", "#a855f7", "#34d399", "#f59e0b", "#ef4444", "#14b8a6", "#eab308", "#f472b6", "#38bdf8", "#fb923c", "#4ade80", "#c084fc", "#2dd4bf", "#facc15", "#f87171"];
  const $ = (s) => document.querySelector(s);

  /* ---------------- CSV ---------------- */
  function parseCSV(text) {
    const rows = []; let row = [], field = "", q = false;
    for (let i = 0; i < text.length; i++) {
      const c = text[i];
      if (q) { if (c === '"') { if (text[i + 1] === '"') { field += '"'; i++; } else q = false; } else field += c; }
      else if (c === '"') q = true;
      else if (c === ",") { row.push(field); field = ""; }
      else if (c === "\n") { row.push(field); rows.push(row); row = []; field = ""; }
      else if (c !== "\r") field += c;
    }
    if (field.length || row.length) { row.push(field); rows.push(row); }
    return rows;
  }
  const MONTHS = { jan: 0, feb: 1, mar: 2, apr: 3, may: 4, jun: 5, jul: 6, aug: 7, sep: 8, oct: 9, nov: 10, dec: 11 };
  function toDate(v) {
    if (v == null || v === "") return null;
    const s = String(v).trim();
    let m = s.match(/^(\d{1,2})-([A-Za-z]{3})-(\d{4})$/);
    if (m) return new Date(+m[3], MONTHS[m[2].toLowerCase()], +m[1]);
    m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (m) return new Date(+m[1], +m[2] - 1, +m[3]);
    if (/^\d+(\.\d+)?$/.test(s)) { const n = +s; if (n > 20000 && n < 80000) return new Date(Date.UTC(1899, 11, 30) + n * 86400000); }
    const d = new Date(s); return isNaN(d) ? null : d;
  }
  const num = (v) => { const n = parseFloat(String(v == null ? "" : v).replace(/[%,]/g, "")); return isNaN(n) ? 0 : n; };
  const dstr = (d) => (d ? `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}` : "");

  /* ---------------- Load ---------------- */
  async function fetchSheet(name) {
    const url = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/gviz/tq?tqx=out:csv&sheet=${encodeURIComponent(name)}&_=${Date.now()}`;
    const res = await fetch(url, { cache: "no-store" });
    if (!res.ok) throw new Error(`${name}: HTTP ${res.status}`);
    return parseCSV(await res.text());
  }
  function rowsToObjects(rows) {
    const header = (rows[0] || []).map((h) => String(h).trim());
    return rows.slice(1).filter((r) => r.some((c) => c !== "")).map((r) => { const o = {}; header.forEach((h, i) => { if (h) o[h] = r[i]; }); return o; });
  }

  const state = { prod: [], def: [], obSmv: {}, sig: "", loading: false, firstLoad: true, smvDim: "pg", tab: "dash",
    filters: { from: "", to: "", pg: "", section: "", line: "", item: "" },
    cross: { dim: "", value: "" } };
  const selects = {};

  async function loadObSmv() {
    try {
      const r = await fetch("ob-smv.json?" + Date.now());
      const arr = await r.json();
      const map = {};
      arr.forEach((o) => { map[o.n] = o.v; });
      state.obSmv = map;
    } catch (e) { /* OB not available — fall back to Target Productivity */ }
  }

  async function loadData() {
    if (state.loading) return;
    state.loading = true;
    setLive(true, "Updating…");
    try {
      const [p, d] = await Promise.all([fetchSheet("Production_Data"), fetchSheet("Defect_Data")]);
      const sig = `${p.length}|${d.length}|${p[p.length - 1] ? p[p.length - 1].join("~") : ""}|${d[d.length - 1] ? d[d.length - 1].join("~") : ""}`;
      const changed = sig !== state.sig;
      if (changed) {
        state.prod = normalizeProd(rowsToObjects(p));
        state.def = normalizeDef(rowsToObjects(d));
        state.sig = sig;
      }
      const now = new Date();
      setLive(false, `Live · ${now.toLocaleTimeString()}`);
      $("#lastUpdated").textContent = `Last updated ${now.toLocaleString()}`;
      if (changed) { if (state.firstLoad) initSlicers(); render(); if (!state.firstLoad) toast("New data loaded ✓"); }
      state.firstLoad = false;
      $("#overlay").classList.add("hide");
    } catch (e) {
      setLive(false, "Offline — retrying…");
      $("#loadMsg").textContent = "Could not reach the sheet. Retrying…";
    } finally { state.loading = false; $("#refreshBtn").classList.remove("spin"); }
  }

  function normalizeProd(o) {
    return o.map((r) => ({ date: toDate(r["Date"]), itemCode: String(r["Item Code"] || ""), item: String(r["Item Name"] || ""),
      line: String(r["Line Name"] || ""), manpower: num(r["Manpower"]), production: num(r["Production"]),
      productivity: num(r["Productivity"]), target: num(r["Target Productivity"]), achievement: num(r["Achievement %"]),
      pg: String(r["PG"] || ""), section: String(r["Section"] || ""), opMin: num(r["Operating Minutes"]) })).filter((r) => r.date);
  }
  function normalizeDef(o) {
    return o.map((r) => ({ date: toDate(r["Date"]), itemCode: String(r["Item Code"] || ""), item: String(r["Item"] || ""),
      line: String(r["Line"] || ""), pg: String(r["PG"] || ""), problem: String(r["Problem Type"] || ""), qty: num(r["Defective Qty"]) })).filter((r) => r.date || r.problem);
  }

  /* ---------------- Relationship (defect <-> production) ---------------- */
  function defMatches(r, ignoreDim) {
    const f = state.filters, c = state.cross;
    if (f.from && r.date && dstr(r.date) < f.from) return false;
    if (f.to && r.date && dstr(r.date) > f.to) return false;
    if (f.pg && r.pg !== f.pg) return false;
    if (f.line && r.line !== f.line) return false;
    if (f.item && r.item !== f.item) return false;
    if (c.dim && c.dim !== ignoreDim) {
      if (c.dim === "pg" && r.pg !== c.value) return false;
      if (c.dim === "line" && r.line !== c.value) return false;
      if (c.dim === "item" && r.item !== c.value) return false;
      if (c.dim === "defectType" && r.problem !== c.value) return false;
      if (c.dim === "date" && r.date && dstr(r.date) !== c.value) return false;
    }
    return true;
  }
  // item codes/lines affected by the active defect-type selection
  function defectAffectedItems() {
    const c = state.cross;
    if (c.dim !== "defectType") return null;
    const items = new Set(), lines = new Set();
    state.def.forEach((r) => {
      if (r.problem !== c.value) return;
      if (!defMatches(r, "defectType")) return;
      if (r.itemCode) items.add(r.itemCode);
      if (r.item) items.add(r.item);
      if (r.line) lines.add(r.line);
    });
    return { items, lines };
  }
  function prodMatches(r, ignoreDim, aff) {
    const f = state.filters, c = state.cross;
    if (f.from && dstr(r.date) < f.from) return false;
    if (f.to && dstr(r.date) > f.to) return false;
    if (f.pg && r.pg !== f.pg) return false;
    if (f.section && r.section !== f.section) return false;
    if (f.line && r.line !== f.line) return false;
    if (f.item && r.item !== f.item) return false;
    if (c.dim && c.dim !== ignoreDim) {
      if (c.dim === "pg" && r.pg !== c.value) return false;
      if (c.dim === "section" && r.section !== c.value) return false;
      if (c.dim === "line" && r.line !== c.value) return false;
      if (c.dim === "item" && r.item !== c.value) return false;
      if (c.dim === "date" && dstr(r.date) !== c.value) return false;
      if (c.dim === "defectType" && aff) { if (!(aff.items.has(r.itemCode) || aff.items.has(r.item))) return false; }
    }
    return true;
  }
  const filteredProd = (ignoreDim) => { const aff = defectAffectedItems(); return state.prod.filter((r) => prodMatches(r, ignoreDim, aff)); };
  const filteredDef = (ignoreDim) => state.def.filter((r) => defMatches(r, ignoreDim));

  /* ---------------- Aggregation ---------------- */
  const sum = (a, f) => a.reduce((x, r) => x + f(r), 0);
  const avg = (a, f) => (a.length ? sum(a, f) / a.length : 0);
  function groupSum(rows, key, val) { const m = new Map(); rows.forEach((r) => { const k = key(r) || "(blank)"; m.set(k, (m.get(k) || 0) + val(r)); }); return [...m.entries()].map(([k, v]) => ({ k, v })); }
  function groupAvg(rows, key, val) { const s = new Map(), c = new Map(); rows.forEach((r) => { const k = key(r) || "(blank)"; s.set(k, (s.get(k) || 0) + val(r)); c.set(k, (c.get(k) || 0) + 1); }); return [...s.entries()].map(([k, v]) => ({ k, v: v / c.get(k) })); }
  const sortByV = (a) => a.sort((x, y) => y.v - x.v);
  const nf = (n, d = 0) => Number(n).toLocaleString("en-US", { minimumFractionDigits: d, maximumFractionDigits: d });

  /* ---------------- KPI ---------------- */
  const ICON = {
    prod: '<path d="M3 13h4v8H3zm7-6h4v14h-4zm7 3h4v11h-4z"/>',
    ach: '<path d="M12 2 3 7v6c0 5 3.8 9.3 9 10 5.2-.7 9-5 9-10V7l-9-5zm-1 13-3-3 1.4-1.4L11 12.2l4.6-4.6L17 9l-6 6z"/>',
    prodv: '<path d="M13 2 3 14h7l-1 8 10-12h-7z"/>',
    mp: '<path d="M16 11a4 4 0 1 0-4-4 4 4 0 0 0 4 4zm-8 1a3 3 0 1 0-3-3 3 3 0 0 0 3 3zm0 2c-2.7 0-6 1.3-6 4v3h8v-3c0-1 .4-1.9 1.1-2.6A8.5 8.5 0 0 0 8 14z"/>',
    def: '<path d="M12 2 1 21h22L12 2zm1 15h-2v-2h2v2zm0-4h-2V9h2v4z"/>',
    rate: '<path d="M4 20h16v2H4zM6 10h3v8H6zm5-6h3v14h-3zm5 4h3v10h-3z"/>',
    hours: '<path d="M12 2a10 10 0 1 0 10 10A10 10 0 0 0 12 2zm1 11h5v-2h-3V7h-2z"/>',
    lines: '<path d="M3 3h8v8H3zm10 0h8v8h-8zM3 13h8v8H3zm10 0h8v8h-8z"/>',
  };
  function renderKpis(p, d) {
    const prod = sum(p, (r) => r.production), ach = avg(p, (r) => r.achievement), prodv = avg(p, (r) => r.productivity);
    const target = avg(p, (r) => r.target), mp = avg(p, (r) => r.manpower), defects = sum(d, (r) => r.qty);
    const rate = prod > 0 ? (defects / prod) * 100 : 0, hours = sum(p, (r) => r.opMin) / 60;
    const cards = [
      ["Total Production", nf(prod), `${nf(p.length)} entries`, "prod", "#22d3ee", "#0ea5e9"],
      ["Avg Achievement", nf(ach, 1) + "%", `vs target`, "ach", "#34d399", "#10b981"],
      ["Avg Productivity", nf(prodv, 2), `target ${nf(target, 1)}`, "prodv", "#a855f7", "#6366f1"],
      ["Avg Manpower", nf(mp, 1), `per entry`, "mp", "#14b8a6", "#0d9488"],
      ["Total Defects", nf(defects), `${nf(d.length)} logs`, "def", "#ef4444", "#f97316"],
      ["Defect Rate", nf(rate, 2) + "%", `of production`, "rate", "#f59e0b", "#ef4444"],
      ["Operating Hours", nf(hours, 1), `from minutes`, "hours", "#38bdf8", "#6366f1"],
      ["Active Lines", nf(new Set(p.map((r) => r.line)).size), `${nf(new Set(p.map((r) => r.item)).size)} items`, "lines", "#f472b6", "#a855f7"],
    ];
    $("#kpis").innerHTML = cards.map(([lab, val, sub, ic, c1, c2]) =>
      `<div class="kpi" style="--k1:${c1};--k2:${c2}">
         <div class="kpi-top"><span class="kpi-ic"><svg viewBox="0 0 24 24" fill="currentColor">${ICON[ic]}</svg></span><span class="lab">${lab}</span></div>
         <div class="val">${val}</div><div class="sub">${sub}</div>
       </div>`).join("");
  }

  /* ---------------- Charts ---------------- */
  const charts = {};
  const tc = () => { const cs = getComputedStyle(document.documentElement); return { text: cs.getPropertyValue("--muted").trim(), grid: cs.getPropertyValue("--border").trim() }; };
  function grad(ctx, c1, c2, horiz) {
    const chart = ctx.chart, area = chart.chartArea;
    if (!area) return c1;
    const g = horiz ? chart.ctx.createLinearGradient(area.left, 0, area.right, 0) : chart.ctx.createLinearGradient(0, area.bottom, 0, area.top);
    g.addColorStop(0, c1); g.addColorStop(1, c2); return g;
  }
  const gradFn = (c1, c2, horiz) => (ctx) => grad(ctx, c1, c2, horiz);
  // Chart.js renders canvas gradients as black on horizontal (indexAxis:"y") bars — use solid interpolated colours there
  function hex2rgb(h) { h = String(h).replace("#", ""); if (h.length === 3) h = h.split("").map((x) => x + x).join(""); return [parseInt(h.slice(0, 2), 16), parseInt(h.slice(2, 4), 16), parseInt(h.slice(4, 6), 16)]; }
  function lerpColor(c1, c2, t) { const a = hex2rgb(c1), b = hex2rgb(c2); t = Math.max(0, Math.min(1, t || 0)); return `rgb(${Math.round(a[0] + (b[0] - a[0]) * t)},${Math.round(a[1] + (b[1] - a[1]) * t)},${Math.round(a[2] + (b[2] - a[2]) * t)})`; }
  function baseOpts() {
    const t = tc();
    return {
      responsive: true, maintainAspectRatio: false, animation: { duration: 700 },
      plugins: { legend: { labels: { color: t.text, boxWidth: 12, font: { size: 11 }, usePointStyle: true } }, tooltip: { backgroundColor: "rgba(8,12,22,.95)", borderColor: t.grid, borderWidth: 1, padding: 10, titleColor: "#fff" } },
      scales: { x: { ticks: { color: t.text, font: { size: 10.5 }, maxRotation: 40, autoSkip: true }, grid: { color: t.grid, drawBorder: false } }, y: { ticks: { color: t.text, font: { size: 10.5 } }, grid: { color: t.grid, drawBorder: false }, beginAtZero: true } },
    };
  }
  function pieOpts() { const t = tc(); return { responsive: true, maintainAspectRatio: false, animation: { duration: 700 }, cutout: "58%", plugins: { legend: { position: "right", labels: { color: t.text, boxWidth: 12, font: { size: 11 }, usePointStyle: true } }, tooltip: { backgroundColor: "rgba(8,12,22,.95)", padding: 10 } } }; }
  function mk(id, type, data, options, onClick) {
    const ctx = $("#" + id); if (!ctx) return;
    if (charts[id]) charts[id].destroy();
    const cfg = { type, data, options };
    if (onClick) cfg.options.onClick = onClick;
    charts[id] = new Chart(ctx, cfg);
  }
  function toggleCross(dim, value) {
    if (state.cross.dim === dim && state.cross.value === value) state.cross = { dim: "", value: "" };
    else state.cross = { dim, value };
    render();
  }
  const clickBar = (dim) => (e, els, chart) => { if (els.length && chart) toggleCross(dim, String(chart.data.labels[els[0].index])); };
  const clickPie = (dim) => (e, els, chart) => { if (els.length && chart) toggleCross(dim, String(chart.data.labels[els[0].index])); };

  function renderCharts(p, d) {
    // ---- Production trend ----
    const byDate = groupSum(p, (r) => dstr(r.date), (r) => r.production).sort((a, b) => (a.k < b.k ? -1 : 1));
    const pd = { labels: byDate.map((x) => x.k), datasets: [{ label: "Production", data: byDate.map((x) => x.v), borderColor: "#22d3ee", backgroundColor: (c) => grad(c, "rgba(34,211,238,.05)", "rgba(34,211,238,.45)"), fill: true, tension: .35, pointRadius: 2, pointBackgroundColor: "#22d3ee", borderWidth: 2 }] };
    if (charts.cProdDate) { charts.cProdDate.data = pd; charts.cProdDate.update(); } else mk("cProdDate", "line", pd, baseOpts(), clickBar("date"));

    // ---- Achievement by PG ----
    const achPg = sortByV(groupAvg(p, (r) => r.pg, (r) => r.achievement));
    const ap = { labels: achPg.map((x) => x.k), datasets: [{ label: "Avg Achv %", data: achPg.map((x) => +x.v.toFixed(1)), backgroundColor: gradFn("#34d399", "#059669"), borderRadius: 7, maxBarThickness: 40 }] };
    if (charts.cAchPg) { charts.cAchPg.data = ap; charts.cAchPg.update(); } else mk("cAchPg", "bar", ap, baseOpts(), clickBar("pg"));

    // ---- Production by Section (doughnut) ----
    const sec = sortByV(groupSum(p, (r) => r.section, (r) => r.production));
    const sd = { labels: sec.map((x) => x.k), datasets: [{ data: sec.map((x) => x.v), backgroundColor: sec.map((_, i) => PALETTE[i % PALETTE.length]), borderWidth: 2, borderColor: "rgba(0,0,0,.25)", hoverOffset: 8 }] };
    if (charts.cSection) { charts.cSection.data = sd; charts.cSection.update(); } else mk("cSection", "doughnut", sd, pieOpts(), clickPie("section"));

    // ---- Production vs Target by PG ----
    const pv = groupAvg(p, (r) => r.pg, (r) => r.productivity).sort((a, b) => a.k.localeCompare(b.k));
    const tmap = Object.fromEntries(groupAvg(p, (r) => r.pg, (r) => r.target).map((x) => [x.k, x.v]));
    const pvdata = { labels: pv.map((x) => x.k), datasets: [
      { label: "Productivity", data: pv.map((x) => +x.v.toFixed(2)), backgroundColor: gradFn("#a855f7", "#6366f1"), borderRadius: 6, maxBarThickness: 26 },
      { label: "Target", data: pv.map((x) => +(tmap[x.k] || 0).toFixed(2)), backgroundColor: gradFn("#fbbf24", "#f59e0b"), borderRadius: 6, maxBarThickness: 26 } ] };
    if (charts.cProdVsTarget) { charts.cProdVsTarget.data = pvdata; charts.cProdVsTarget.update(); } else mk("cProdVsTarget", "bar", pvdata, baseOpts(), clickBar("pg"));

    // ---- Production by Line ----
    const ln = sortByV(groupSum(p, (r) => r.line, (r) => r.production)).slice(0, 12);
    const ldata = { labels: ln.map((x) => x.k), datasets: [{ label: "Production", data: ln.map((x) => x.v), backgroundColor: gradFn("#22d3ee", "#2563eb"), borderRadius: 6, maxBarThickness: 34 }] };
    if (charts.cLine) { charts.cLine.data = ldata; charts.cLine.update(); } else mk("cLine", "bar", ldata, baseOpts(), clickBar("line"));

    // ---- Defects by Problem Type (self-excluded so all bars remain) ----
    const dTypeRows = filteredDef("defectType");
    const dt = sortByV(groupSum(dTypeRows, (r) => r.problem, (r) => r.qty)).slice(0, 10);
    const sel = state.cross.dim === "defectType" ? state.cross.value : "";
    const dmax = Math.max(1, ...dt.map((x) => x.v));
    const dtColors = dt.map((x) => (sel ? (x.k === sel ? "#ef4444" : "rgba(239,68,68,.28)") : lerpColor("#f97316", "#ef4444", x.v / dmax)));
    const dtd = { labels: dt.map((x) => x.k), datasets: [{ label: "Defective Qty", data: dt.map((x) => x.v), backgroundColor: dtColors, borderRadius: 6, maxBarThickness: 26 }] };
    const dto = baseOpts(); dto.indexAxis = "y";
    if (charts.cDefectType) { charts.cDefectType.data = dtd; charts.cDefectType.update(); } else mk("cDefectType", "bar", dtd, dto, clickBar("defectType"));

    // ---- Defect share by PG ----
    const dpg = sortByV(groupSum(d, (r) => r.pg, (r) => r.qty));
    const dpgd = { labels: dpg.map((x) => x.k), datasets: [{ data: dpg.map((x) => x.v), backgroundColor: dpg.map((_, i) => PALETTE[(i + 3) % PALETTE.length]), borderWidth: 2, borderColor: "rgba(0,0,0,.25)", hoverOffset: 8 }] };
    if (charts.cDefectPg) { charts.cDefectPg.data = dpgd; charts.cDefectPg.update(); } else mk("cDefectPg", "doughnut", dpgd, pieOpts(), clickPie("pg"));

    // ---- Defect trend ----
    const dd = groupSum(d, (r) => dstr(r.date), (r) => r.qty).sort((a, b) => (a.k < b.k ? -1 : 1));
    const ddd = { labels: dd.map((x) => x.k), datasets: [{ label: "Defects", data: dd.map((x) => x.v), borderColor: "#f43f5e", backgroundColor: (c) => grad(c, "rgba(244,63,94,.05)", "rgba(244,63,94,.42)"), fill: true, tension: .35, pointRadius: 2, pointBackgroundColor: "#f43f5e", borderWidth: 2 }] };
    if (charts.cDefectDate) { charts.cDefectDate.data = ddd; charts.cDefectDate.update(); } else mk("cDefectDate", "line", ddd, baseOpts(), clickBar("date"));
  }

  /* ---------------- Table ---------------- */
  function renderTable(p, d) {
    const map = new Map();
    p.forEach((r) => { const k = r.item || r.itemCode; if (!map.has(k)) map.set(k, { item: k, pg: r.pg, prod: 0, ach: 0, n: 0, mp: 0 }); const o = map.get(k); o.prod += r.production; o.ach += r.achievement; o.mp += r.manpower; o.n++; });
    const dmap = new Map(); d.forEach((r) => { const k = r.item || r.itemCode; dmap.set(k, (dmap.get(k) || 0) + r.qty); });
    const rows = [...map.values()].map((o) => ({ ...o, ach: o.n ? o.ach / o.n : 0, mp: o.n ? o.mp / o.n : 0, def: dmap.get(o.item) || 0 })).sort((a, b) => b.prod - a.prod).slice(0, 15);
    $("#tblItems").innerHTML = rows.map((r, i) => `<tr data-item="${r.item.replace(/"/g, "&quot;")}"><td><span class="rank">${i + 1}</span></td><td>${r.item}</td><td>${r.pg}</td><td class="num">${nf(r.prod)}</td><td class="num">${nf(r.ach, 1)}%</td><td class="num">${nf(r.mp, 1)}</td><td class="num">${nf(r.def)}</td></tr>`).join("");
    document.querySelectorAll("#tblItems tr").forEach((tr) => tr.addEventListener("click", () => toggleCross("item", tr.dataset.item)));
  }

  /* ---------------- SMV (Standard from Operation Bulletin, Actual from Production sheet) ---------------- */
  function smvOf(rows) {
    let prod = 0, opMin = 0, stdMin = 0;
    rows.forEach((r) => {
      const minutes = r.opMin > 0 ? r.opMin : 60;
      prod += r.production;
      opMin += r.manpower * minutes;                                              // actual operator-minutes
      const ob = state.obSmv[r.item] ?? state.obSmv[r.item.toLowerCase()];       // Standard SMV from OB (case-insensitive)
      const stdSmv = ob != null ? ob : (r.target > 0 ? 60 / r.target : null);
      if (stdSmv != null) stdMin += r.production * stdSmv;                        // standard operator-minutes
    });
    return { prod, std: prod > 0 ? stdMin / prod : 0, act: prod > 0 ? opMin / prod : 0 };
  }
  function renderSmv(p) {
    const dim = state.smvDim;
    const key = dim === "item" ? (r) => r.item : dim === "pg" ? (r) => r.pg : (r) => r.section;
    const m = new Map();
    p.forEach((r) => { const k = key(r) || "(blank)"; if (!m.has(k)) m.set(k, []); m.get(k).push(r); });
    let rows = [...m.entries()].map(([k, rs]) => ({ k, ...smvOf(rs) }));
    if (dim === "item") rows.sort((a, b) => b.prod - a.prod); else rows.sort((a, b) => b.std - a.std);
    const top = rows.slice(0, dim === "item" ? 12 : 14);

    const data = { labels: top.map((x) => x.k), datasets: [
      { label: "Standard SMV", data: top.map((x) => +x.std.toFixed(3)), backgroundColor: gradFn("#34d399", "#0d9488"), borderRadius: 6, maxBarThickness: 26 },
      { label: "Actual SMV", data: top.map((x) => +x.act.toFixed(3)), backgroundColor: gradFn("#f59e0b", "#ef4444"), borderRadius: 6, maxBarThickness: 26 } ] };
    if (charts.cSmv) { charts.cSmv.data = data; charts.cSmv.update(); } else mk("cSmv", "bar", data, baseOpts());

    const all = smvOf(p), c = state.cross;
    let sel;
    if (c.dim && c.dim !== "date" && c.dim !== "defectType") sel = `${c.dim === "item" ? "Product" : c.dim === "pg" ? "Group" : c.dim === "section" ? "Section" : c.dim}: ${c.value}`;
    else if (state.filters.item) sel = `Product: ${state.filters.item}`;
    else if (state.filters.pg) sel = `Group: ${state.filters.pg}`;
    else if (state.filters.section) sel = `Section: ${state.filters.section}`;
    else sel = dim === "item" ? "All products" : dim === "pg" ? "All groups" : "All sections";

    $("#smvRun").innerHTML =
      `<span class="s sel">Running SMV for<b>${sel}</b></span>` +
      `<span class="s std">Standard SMV<b>${nf(all.std, 3)}</b></span>` +
      `<span class="s act">Actual SMV<b>${nf(all.act, 3)}</b></span>` +
      `<span class="s">Production<b>${nf(all.prod)}</b></span>`;
    $("#smvFoot").innerHTML = rows.length
      ? `<span class="s">Categories<b>${nf(rows.length)}</b></span><span class="s">Standard<b>Operation Bulletin</b></span><span class="s">Actual<b>Production ÷ Operator-minutes</b></span>`
      : `<span class="s" style="color:#ef4444">No data for this selection — clear a filter or pick another.</span>`;
  }

  /* ================= OPEX Analysis ================= */
  const OBSTD = (r) => { const ob = state.obSmv[r.item] ?? state.obSmv[r.item.toLowerCase()]; return ob != null ? ob : (r.target > 0 ? 60 / r.target : null); };
  const rMin = (r) => r.manpower * (r.opMin > 0 ? r.opMin : 60);

  function lineOpex(rows) {
    const m = new Map();
    rows.forEach((r) => {
      const k = r.line || "(blank)";
      if (!m.has(k)) m.set(k, { k, prod: 0, opMin: 0, stdMin: 0, mp: 0, n: 0, ach: 0 });
      const o = m.get(k), std = OBSTD(r);
      o.prod += r.production; o.opMin += rMin(r); o.mp += r.manpower; o.n++; o.ach += r.achievement;
      if (std != null) o.stdMin += r.production * std;
    });
    return [...m.values()].map((o) => {
      const std = o.prod > 0 ? o.stdMin / o.prod : 0;
      const act = o.prod > 0 ? o.opMin / o.prod : 0;
      const achievable = std > 0 ? o.opMin / std : o.prod;
      return { ...o, std, act, gap: act - std, eff: act > 0 ? (std / act) * 100 : 0,
        achievable, lost: Math.max(0, achievable - o.prod), lostMin: Math.max(0, o.opMin - o.stdMin),
        ach: o.n ? o.ach / o.n : 0, mpp: o.opMin > 0 ? o.prod / (o.opMin / 60) : 0 };
    });
  }
  function lineDefects(d) { const m = new Map(); d.forEach((r) => { const k = r.line || "(blank)"; m.set(k, (m.get(k) || 0) + r.qty); }); return m; }

  function opexTotals(p, d) {
    const all = smvOf(p), opMin = sum(p, rMin), def = sum(d, (r) => r.qty);
    const achievable = all.std > 0 ? opMin / all.std : all.prod;
    return { std: all.std, act: all.act, prod: all.prod, opMin, def, achievable,
      gap: all.act - all.std, eff: all.act > 0 ? (all.std / all.act) * 100 : 0,
      lost: Math.max(0, achievable - all.prod),
      ppm: all.prod > 0 ? (def / all.prod) * 1e6 : 0,
      wastePct: all.prod > 0 ? (def / all.prod) * 100 : 0,
      mpp: opMin > 0 ? all.prod / (opMin / 60) : 0,
      recoverMin: Math.max(0, opMin - all.std * all.prod) };
  }

  function renderOpexGoals(p, d) {
    const t = opexTotals(p, d);
    const cards = [
      ["SMV Gap", (t.gap >= 0 ? "+" : "") + nf(t.gap, 3), "min / pc over standard", "hours", "#f59e0b", "#ef4444", t.act > 0 ? Math.min(100, (t.gap / t.act) * 100) : 0],
      ["Line Efficiency", nf(t.eff, 1) + "%", "standard ÷ actual pace", "ach", "#34d399", "#0d9488", Math.min(100, t.eff)],
      ["Lost Capacity", nf(t.lost), "pcs recoverable at standard", "prodv", "#38bdf8", "#6366f1", 0],
      ["Wastage", nf(t.wastePct, 2) + "%", nf(t.ppm) + " PPM defective", "def", "#ef4444", "#f97316", Math.min(100, t.wastePct * 10)],
      ["Manpower Productivity", nf(t.mpp, 2), "pcs / operator-hour", "mp", "#a855f7", "#6366f1", 0],
      ["Recoverable Minutes", nf(t.recoverMin), "operator-min · this period", "rate", "#14b8a6", "#0d9488", 0],
    ];
    $("#opexGoals").innerHTML = cards.map(([lab, val, sub, ic, c1, c2, bar]) =>
      `<div class="opex-goal" style="--g1:${c1};--g2:${c2}">
         <div class="g-top"><span class="g-ic"><svg viewBox="0 0 24 24" fill="currentColor">${ICON[ic] || ICON.prod}</svg></span><span class="g-lab">${lab}</span></div>
         <div class="g-val">${val}</div><div class="g-sub">${sub}</div>
         ${bar > 0 ? `<div class="g-bar"><i style="width:${Math.max(2, Math.min(100, bar)).toFixed(0)}%"></i></div>` : ""}
       </div>`).join("");
  }

  function renderOpexFocus(p, d) {
    const t = opexTotals(p, d);
    const lo = lineOpex(p).filter((x) => x.prod > 0);
    const dm = lineDefects(d);
    const worstSmv = lo.slice().sort((a, b) => b.lostMin - a.lostMin).slice(0, 4);
    const worstPpm = lo.map((o) => ({ k: o.k, v: o.prod > 0 ? ((dm.get(o.k) || 0) / o.prod) * 1e6 : 0 })).sort((a, b) => b.v - a.v).slice(0, 4);
    const bestProd = lo.slice().sort((a, b) => b.lost - a.lost).slice(0, 4);
    const card = (title, color, big, sub, list) =>
      `<div class="focus-card" style="--fc:${color}">
         <h4><span class="dot"></span>${title}</h4>
         <div class="fc-val">${big}</div>
         <div class="fc-sub">${sub}</div>
         <ul>${list.length ? list.map(([a, b]) => `<li><span title="${a}">${a}</span><b>${b}</b></li>`).join("") : `<li><span>No data in this selection</span><b>—</b></li>`}</ul>
       </div>`;
    $("#opexFocus").innerHTML =
      card("Reduce line SMV", "#f59e0b", nf(t.gap, 3) + " min/pc",
        `Recoverable <b>${nf(t.recoverMin)}</b> operator-minutes ≈ <b>${nf(t.lost)}</b> pcs at standard pace.`,
        worstSmv.map((o) => [o.k, "−" + nf(o.lostMin) + " min"])) +
      card("Reduce wastage", "#ef4444", nf(t.ppm) + " PPM",
        `Defect rate <b>${nf(t.wastePct, 2)}%</b> · <b>${nf(t.def)}</b> defective pcs in this selection.`,
        worstPpm.map((o) => [o.k, nf(o.v) + " PPM"])) +
      card("Increase production", "#22d3ee", "+" + nf(t.lost) + " pcs",
        `Achievable <b>${nf(t.achievable)}</b> vs actual <b>${nf(t.prod)}</b> with the same manpower.`,
        bestProd.map((o) => [o.k, "+" + nf(o.lost) + " pcs"]));
  }

  function renderOpexCharts(p, d) {
    const lo = lineOpex(p).filter((x) => x.prod > 0);
    const dm = lineDefects(d);

    // SMV gap by line (worst first)
    const g = lo.slice().sort((a, b) => b.gap - a.gap).slice(0, 12);
    const gmax = Math.max(0.001, ...g.map((x) => Math.max(0, x.gap)));
    const gd = { labels: g.map((x) => x.k), datasets: [{ label: "SMV gap (min/pc)", data: g.map((x) => +x.gap.toFixed(3)), backgroundColor: g.map((x) => (x.gap > 0 ? lerpColor("#f59e0b", "#ef4444", x.gap / gmax) : "#34d399")), borderRadius: 6, maxBarThickness: 24 }] };
    const go = baseOpts(); go.indexAxis = "y";
    if (charts.oC_smvGapLine) { charts.oC_smvGapLine.data = gd; charts.oC_smvGapLine.update(); } else mk("oC_smvGapLine", "bar", gd, go, clickBar("line"));

    // efficiency by line (worst first)
    const e = lo.slice().sort((a, b) => a.eff - b.eff).slice(0, 12);
    const ed = { labels: e.map((x) => x.k), datasets: [{ label: "Efficiency %", data: e.map((x) => +x.eff.toFixed(1)), backgroundColor: gradFn("#34d399", "#059669"), borderRadius: 6, maxBarThickness: 26 }] };
    if (charts.oC_effLine) { charts.oC_effLine.data = ed; charts.oC_effLine.update(); } else mk("oC_effLine", "bar", ed, baseOpts(), clickBar("line"));

    // lost capacity by line
    const l = lo.filter((x) => x.lost > 0).sort((a, b) => b.lost - a.lost).slice(0, 12);
    const ld = { labels: l.map((x) => x.k), datasets: [{ label: "Lost pcs", data: l.map((x) => Math.round(x.lost)), backgroundColor: gradFn("#38bdf8", "#6366f1"), borderRadius: 6, maxBarThickness: 26 }] };
    if (charts.oC_lostLine) { charts.oC_lostLine.data = ld; charts.oC_lostLine.update(); } else mk("oC_lostLine", "bar", ld, baseOpts(), clickBar("line"));

    // defect PPM by line
    const pm = lo.map((o) => ({ k: o.k, v: o.prod > 0 ? ((dm.get(o.k) || 0) / o.prod) * 1e6 : 0 })).sort((a, b) => b.v - a.v).slice(0, 12);
    const pmd = { labels: pm.map((x) => x.k), datasets: [{ label: "Defect PPM", data: pm.map((x) => Math.round(x.v)), backgroundColor: gradFn("#ef4444", "#b91c1c"), borderRadius: 6, maxBarThickness: 26 }] };
    if (charts.oC_ppmLine) { charts.oC_ppmLine.data = pmd; charts.oC_ppmLine.update(); } else mk("oC_ppmLine", "bar", pmd, baseOpts(), clickBar("line"));

    // wastage pareto (bar + cumulative %)
    const dt = sortByV(groupSum(filteredDef("defectType"), (r) => r.problem, (r) => r.qty)).slice(0, 10);
    const tot = sum(dt, (x) => x.v);
    let cum = 0;
    const cumD = dt.map((x) => { cum += x.v; return tot ? +(cum / tot * 100).toFixed(1) : 0; });
    const sel = state.cross.dim === "defectType" ? state.cross.value : "";
    const par = { labels: dt.map((x) => x.k), datasets: [
      { type: "bar", label: "Defective qty", data: dt.map((x) => x.v), backgroundColor: dt.map((x) => (sel ? (x.k === sel ? "#ef4444" : "rgba(239,68,68,.28)") : lerpColor("#f97316", "#ef4444", x.v / Math.max(1, dt.length ? dt[0].v : 1)))), borderRadius: 6, maxBarThickness: 30, yAxisID: "y", order: 2 },
      { type: "line", label: "Cumulative %", data: cumD, borderColor: "#22d3ee", backgroundColor: "#22d3ee", tension: .3, pointRadius: 3, borderWidth: 2, yAxisID: "y1", order: 1 } ] };
    const po = baseOpts();
    po.scales.y1 = { position: "right", beginAtZero: true, max: 100, ticks: { color: tc().text, font: { size: 10.5 }, callback: (v) => v + "%" }, grid: { drawOnChartArea: false } };
    if (charts.oC_pareto) { charts.oC_pareto.data = par; charts.oC_pareto.update(); } else mk("oC_pareto", "bar", par, po, clickBar("defectType"));

    // production vs achievable (daily)
    const m = new Map();
    p.forEach((r) => { const k = dstr(r.date); if (!m.has(k)) m.set(k, { k, prod: 0, opMin: 0, stdMin: 0 }); const o = m.get(k), std = OBSTD(r); o.prod += r.production; o.opMin += rMin(r); if (std != null) o.stdMin += r.production * std; });
    const days = [...m.values()].map((o) => { const std = o.prod > 0 ? o.stdMin / o.prod : 0; return { k: o.k, prod: o.prod, ach: std > 0 ? o.opMin / std : o.prod }; }).sort((a, b) => (a.k < b.k ? -1 : 1));
    const pgd = { labels: days.map((x) => x.k), datasets: [
      { label: "Actual production", data: days.map((x) => x.prod), borderColor: "#22d3ee", backgroundColor: (c) => grad(c, "rgba(34,211,238,.04)", "rgba(34,211,238,.35)"), fill: true, tension: .35, pointRadius: 0, borderWidth: 2 },
      { label: "Achievable at standard SMV", data: days.map((x) => Math.round(x.ach)), borderColor: "#34d399", borderDash: [6, 4], backgroundColor: "transparent", fill: false, tension: .35, pointRadius: 0, borderWidth: 2 } ] };
    if (charts.oC_prodGap) { charts.oC_prodGap.data = pgd; charts.oC_prodGap.update(); } else mk("oC_prodGap", "line", pgd, baseOpts(), clickBar("date"));

    // manpower productivity by line
    const mpl = lo.slice().sort((a, b) => b.mpp - a.mpp).slice(0, 12);
    const mpd = { labels: mpl.map((x) => x.k), datasets: [{ label: "pcs / op-hr", data: mpl.map((x) => +x.mpp.toFixed(2)), backgroundColor: gradFn("#a855f7", "#6366f1"), borderRadius: 6, maxBarThickness: 26 }] };
    if (charts.oC_mpProd) { charts.oC_mpProd.data = mpd; charts.oC_mpProd.update(); } else mk("oC_mpProd", "bar", mpd, baseOpts(), clickBar("line"));
  }

  function renderOpexTable(p, d) {
    const lo = lineOpex(p), dm = lineDefects(d);
    const rows = lo.map((o) => {
      const def = dm.get(o.k) || 0;
      return { ...o, def, ppm: o.prod > 0 ? (def / o.prod) * 1e6 : 0, impact: o.lostMin + def * o.std };
    }).filter((o) => o.prod > 0).sort((a, b) => b.impact - a.impact).slice(0, 15);
    $("#tblOpex").innerHTML = rows.map((r, i) => {
      const cls = r.gap <= 0.02 ? "good" : r.gap <= 0.25 ? "warn" : "bad";
      return `<tr data-line="${String(r.k).replace(/"/g, "&quot;")}">
        <td><span class="rank">${i + 1}</span></td>
        <td>${r.k}</td>
        <td class="num">${nf(r.std, 3)}</td>
        <td class="num">${nf(r.act, 3)}</td>
        <td class="num"><span class="tag ${cls}">${r.gap >= 0 ? "+" : ""}${nf(r.gap, 3)}</span></td>
        <td class="num">${nf(r.eff, 1)}%</td>
        <td class="num">${nf(r.lost)}</td>
        <td class="num">${nf(r.ppm)}</td>
        <td class="num">${nf(r.impact)}</td>
      </tr>`;
    }).join("");
    document.querySelectorAll("#tblOpex tr").forEach((tr) => tr.addEventListener("click", () => toggleCross("line", tr.dataset.line)));
  }

  function renderOpex(p, d) { renderOpexGoals(p, d); renderOpexFocus(p, d); renderOpexCharts(p, d); renderOpexTable(p, d); }

  function renderTabs() {
    document.querySelectorAll(".tab").forEach((t) => t.classList.toggle("on", t.dataset.tab === state.tab));
    const vd = $("#viewDash"), vo = $("#viewOpex");
    if (vd) vd.hidden = state.tab !== "dash";
    if (vo) vo.hidden = state.tab !== "opex";
  }

  function render() {
    const p = filteredProd(), d = filteredDef();
    renderTabs();
    if (state.tab === "opex") renderOpex(p, d);
    else { renderKpis(p, d); renderCharts(p, d); renderSmv(p); renderTable(p, d); }
    renderChips();
  }

  /* ---------------- Slicers (searchable) ---------------- */
  function escHtml(s) { return String(s == null ? "" : s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;"); }
  function escAttr(s) { return escHtml(s).replace(/"/g, "&quot;"); }
  function createSearchSelect(host, onChange) {
    const S = { options: [], value: "" };
    host.classList.add("ss");
    host.innerHTML =
      '<div class="ss-ctl">' +
        '<input class="ss-inp" type="text" placeholder="All" autocomplete="off" spellcheck="false" />' +
        '<button class="ss-clr" type="button" aria-label="Clear" hidden>\u00d7</button>' +
        '<svg class="ss-car" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round"><polyline points="6 9 12 15 18 9"/></svg>' +
      '</div>' +
      '<div class="ss-pop" hidden><div class="ss-list"></div></div>';
    const inp = host.querySelector(".ss-inp"), clr = host.querySelector(".ss-clr"),
          pop = host.querySelector(".ss-pop"), list = host.querySelector(".ss-list");
    let cur = -1;
    function sync() {
      inp.value = S.value || "";
      inp.placeholder = S.value ? "" : "All";
      inp.classList.toggle("has", !!S.value);
      clr.hidden = !S.value;
    }
    function draw() {
      const q = inp.value.trim();
      const typed = q && q.toLowerCase() !== (S.value || "").toLowerCase();
      const items = typed ? S.options.filter((o) => o.toLowerCase().includes(q.toLowerCase())) : S.options;
      const CAP = 400, capped = items.slice(0, CAP);
      let html = `<div class="ss-opt${S.value ? "" : " on"}" data-v="">All</div>`;
      html += capped.map((o) => `<div class="ss-opt${o === S.value ? " on" : ""}" data-v="${escAttr(o)}" title="${escAttr(o)}">${escHtml(o)}</div>`).join("");
      if (!items.length) html = `<div class="ss-empty">No match for \u201c${escHtml(q)}\u201d</div>`;
      else if (items.length > CAP) html += `<div class="ss-more">${(items.length - CAP).toLocaleString()} more \u2014 keep typing to narrow</div>`;
      list.innerHTML = html;
      cur = -1;
      if (S.value && !typed) { const on = list.querySelector(".ss-opt.on"); if (on) on.scrollIntoView({ block: "nearest" }); }
    }
    function open() { host.classList.add("open"); pop.hidden = false; draw(); }
    function close() { host.classList.remove("open"); pop.hidden = true; }
    function pick(v) { S.value = v || ""; sync(); close(); if (onChange) onChange(S.value); }
    inp.addEventListener("focus", () => { open(); inp.select(); });
    inp.addEventListener("input", open);
    inp.addEventListener("keydown", (e) => {
      const opts = [...list.querySelectorAll(".ss-opt")];
      if (e.key === "ArrowDown" || e.key === "ArrowUp") {
        e.preventDefault(); if (pop.hidden) { open(); return; }
        cur = e.key === "ArrowDown" ? Math.min(opts.length - 1, cur + 1) : Math.max(0, cur - 1);
        opts.forEach((o, i) => o.classList.toggle("hl", i === cur));
        if (opts[cur]) opts[cur].scrollIntoView({ block: "nearest" });
      } else if (e.key === "Enter") {
        e.preventDefault(); const el = cur >= 0 ? opts[cur] : list.querySelector(".ss-opt"); if (el) pick(el.getAttribute("data-v"));
      } else if (e.key === "Escape") { close(); inp.blur(); }
    });
    inp.addEventListener("blur", () => setTimeout(() => { sync(); close(); }, 150));
    clr.addEventListener("mousedown", (e) => { e.preventDefault(); pick(""); });
    list.addEventListener("mousedown", (e) => { const o = e.target.closest(".ss-opt"); if (o) { e.preventDefault(); pick(o.getAttribute("data-v")); } });
    document.addEventListener("click", (e) => { if (!host.contains(e.target)) close(); });
    sync();
    return {
      get value() { return S.value; },
      get options() { return S.options; },
      setValue(v) { S.value = v || ""; sync(); },
      setOptions(vals) { S.options = vals.slice(); sync(); },
    };
  }

  function initSlicers() {
    const JUNK = /^(#n\/a|#ref!|#div\/0!|#value!|#name\?|#null!|n\/a|-|--)$/i;
    const uniq = (f) => [...new Set(state.prod.map(f).map((v) => String(v == null ? "" : v).trim()).filter((v) => v && !JUNK.test(v)))].sort();
    const read = () => {
      state.filters.from = $("#fFrom").value; state.filters.to = $("#fTo").value;
      state.filters.pg = selects.fPg.value; state.filters.section = selects.fSection.value;
      state.filters.line = selects.fLine.value; state.filters.item = selects.fItem.value;
      render();
    };
    selects.fPg = createSearchSelect($("#fPg"), read);
    selects.fSection = createSearchSelect($("#fSection"), read);
    selects.fLine = createSearchSelect($("#fLine"), read);
    selects.fItem = createSearchSelect($("#fItem"), read);
    selects.fPg.setOptions(uniq((r) => r.pg));
    selects.fSection.setOptions(uniq((r) => r.section));
    selects.fLine.setOptions(uniq((r) => r.line));
    selects.fItem.setOptions(uniq((r) => r.item));
    const dates = state.prod.map((r) => dstr(r.date)).filter(Boolean).sort();
    if (dates.length) { const from = $("#fFrom"), to = $("#fTo"); from.min = to.min = dates[0]; from.max = to.max = dates[dates.length - 1]; from.value = dates[0]; to.value = dates[dates.length - 1]; state.filters.from = dates[0]; state.filters.to = dates[dates.length - 1]; }
    ["#fFrom", "#fTo"].forEach((id) => $(id).addEventListener("change", read));
    document.querySelectorAll("#smvSeg button").forEach((b) => b.addEventListener("click", () => {
      document.querySelectorAll("#smvSeg button").forEach((x) => x.classList.remove("on"));
      b.classList.add("on"); state.smvDim = b.dataset.dim; render();
    }));
    document.querySelectorAll(".tab").forEach((t) => t.addEventListener("click", () => { state.tab = t.dataset.tab; render(); }));
  }
  function renderChips() {
    const chips = [], c = state.cross, f = state.filters;
    if (c.dim) chips.push({ label: `${c.dim === "defectType" ? "defect" : c.dim}: ${c.value}`, clear: () => (state.cross = { dim: "", value: "" }) });
    if (f.pg) chips.push({ label: `PG: ${f.pg}`, clear: () => { state.filters.pg = ""; selects.fPg.setValue(""); } });
    if (f.section) chips.push({ label: `Section: ${f.section}`, clear: () => { state.filters.section = ""; selects.fSection.setValue(""); } });
    if (f.line) chips.push({ label: `Line: ${f.line}`, clear: () => { state.filters.line = ""; selects.fLine.setValue(""); } });
    if (f.item) chips.push({ label: `Item: ${f.item}`, clear: () => { state.filters.item = ""; selects.fItem.setValue(""); } });
    $("#chips").innerHTML = chips.length
      ? chips.map((ch, i) => `<span class="chip">${ch.label}<button data-i="${i}">✕</button></span>`).join("") + `<button class="clear-btn" id="clearAll">Clear all</button>`
      : `<span style="color:var(--muted);font-size:12px">Click any chart bar / slice / table row to cross-filter — every visual updates together</span>`;
    document.querySelectorAll("#chips .chip button").forEach((b) => b.addEventListener("click", () => { chips[+b.dataset.i].clear(); render(); }));
    const ca = $("#clearAll"); if (ca) ca.addEventListener("click", () => { state.cross = { dim: "", value: "" }; state.filters = { ...state.filters, pg: "", section: "", line: "", item: "" }; ["fPg", "fSection", "fLine", "fItem"].forEach((k) => selects[k].setValue("")); render(); });
  }

  /* ---------------- Live / toast / theme ---------------- */
  function setLive(updating, text) { $("#livePill").classList.toggle("updating", updating); $("#liveText").textContent = text; }
  let toastT; function toast(msg) { const t = $("#toast"); t.textContent = msg; t.classList.add("show"); clearTimeout(toastT); toastT = setTimeout(() => t.classList.remove("show"), 2600); }

  const root = document.documentElement;
  root.setAttribute("data-theme", localStorage.getItem("alel-dash-theme") || "dark");
  $("#themeBtn").addEventListener("click", () => { const next = root.getAttribute("data-theme") === "light" ? "dark" : "light"; root.setAttribute("data-theme", next); localStorage.setItem("alel-dash-theme", next); Object.values(charts).forEach((c) => c.destroy()); Object.keys(charts).forEach((k) => delete charts[k]); render(); });
  $("#refreshBtn").addEventListener("click", () => { $("#refreshBtn").classList.add("spin"); loadData(); });

  /* ---------------- PWA ---------------- */
  if ("serviceWorker" in navigator) window.addEventListener("load", () => navigator.serviceWorker.register("sw.js").catch(() => {}));
  let deferred = null; const installBtn = $("#installBtn");
  window.addEventListener("beforeinstallprompt", (e) => { e.preventDefault(); deferred = e; installBtn.classList.add("show"); });
  installBtn.addEventListener("click", () => { if (deferred) { deferred.prompt(); deferred.userChoice.then(() => { deferred = null; installBtn.classList.remove("show"); }); return; } alert("To install:\n\n• Windows (Edge/Chrome): menu ⋮ → Apps → Install this site as an app\n• iPhone/iPad (Safari): Share ⬆ → Add to Home Screen\n• Android (Chrome): ⋮ → Install app"); });
  if (/iphone|ipad|ipod/i.test(navigator.userAgent) && !navigator.standalone) installBtn.classList.add("show");

  /* ---------------- Boot ---------------- */
  loadObSmv().then(() => loadData());
  setInterval(() => { if (!document.hidden) loadData(); }, REFRESH_MS);
  document.addEventListener("visibilitychange", () => { if (!document.hidden) loadData(); });
})();
