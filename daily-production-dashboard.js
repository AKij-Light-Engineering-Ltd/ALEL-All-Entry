/* =========================================================
   ALEL — Daily Production Dashboard  (OPEX / IE / Lean / CEO)
   Live from the master Google Sheet (public).
   ========================================================= */
(function () {
  "use strict";

  const SHEET_ID = "1vYNcYNcCog_AZgzSL4xs8WuyJVUCDxsmZzbblq9Xrtg";
  const SHEET = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/gviz/tq`;
  const TABS = {
    prod: { name: "Raw Data Production", end: "M" },
    qual: { name: "Raw data Quality", end: "K" },
  };
  const FAST_MS = 1000;     // change-signal poll (~100 bytes) — detects any add / edit / delete
  const FULL_MS = 60000;    // full reconcile every 60s as a safety net
  const $ = (s) => document.querySelector(s);
  const PALETTE = ["#22d3ee","#6366f1","#a855f7","#34d399","#f59e0b","#ef4444","#14b8a6","#eab308","#f472b6","#38bdf8","#fb923c","#4ade80","#c084fc","#2dd4bf","#facc15","#f87171"];

  /* ---------------- CSV ---------------- */
  function parseCSV(text) {
    const rows = []; let row = [], f = "", q = false;
    for (let i = 0; i < text.length; i++) {
      const c = text[i];
      if (q) { if (c === '"') { if (text[i + 1] === '"') { f += '"'; i++; } else q = false; } else f += c; }
      else if (c === '"') q = true;
      else if (c === ",") { row.push(f); f = ""; }
      else if (c === "\n") { row.push(f); rows.push(row); row = []; f = ""; }
      else if (c !== "\r") f += c;
    }
    if (f.length || row.length) { row.push(f); rows.push(row); }
    return rows;
  }
  const MON = { jan:0,feb:1,mar:2,apr:3,may:4,jun:5,jul:6,aug:7,sep:8,oct:9,nov:10,dec:11 };
  function toDate(v) {
    if (v == null || v === "") return null;
    const s = String(v).trim();
    let m = s.match(/^(\d{1,2})-([A-Za-z]{3})-(\d{4})$/);
    if (m) return new Date(+m[3], MON[m[2].toLowerCase()], +m[1]);
    m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (m) return new Date(+m[1], +m[2] - 1, +m[3]);
    const d = new Date(s); return isNaN(d) ? null : d;
  }
  const dstr = (d) => d ? `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}-${String(d.getDate()).padStart(2,"0")}` : "";
  const num = (v) => { const n = parseFloat(String(v == null ? "" : v).replace(/[%,]/g, "")); return isNaN(n) ? 0 : n; };

  /* hour slot "08 AM - 09 AM" -> 8 */
  function hourOf(slot) {
    const m = String(slot || "").match(/(\d{1,2})\s*(AM|PM)/i);
    if (!m) return -1;
    let h = +m[1] % 12;
    if (/PM/i.test(m[2])) h += 12;
    return h;
  }

  /* ---------------- Fetch ---------------- */
  async function fetchTab(name) {
    const res = await fetch(`${SHEET}?tqx=out:csv&sheet=${encodeURIComponent(name)}&_=${Date.now()}`, { cache: "no-store" });
    if (!res.ok) throw new Error(`${name}: HTTP ${res.status}`);
    return parseCSV(await res.text());
  }
  // fetch only rows after `afterRow` — tiny payload, used for live tail polling
  async function fetchTail(tab, afterRow) {
    const url = `${SHEET}?tqx=out:csv&sheet=${encodeURIComponent(tab.name)}&range=A${afterRow + 1}:${tab.end}${afterRow + 400}&_=${Date.now()}`;
    const res = await fetch(url, { cache: "no-store" });
    if (!res.ok) return [];
    const rows = parseCSV(await res.text());
    return rows.filter((r) => r.some((c) => String(c).trim() !== ""));
  }
  // tiny aggregate "change signal" — row count + key column sums.
  // Changes whenever a row is added, deleted, or any of those columns is edited.
  async function fetchSignal(tab, select) {
    const url = `${SHEET}?tqx=out:csv&sheet=${encodeURIComponent(tab.name)}&tq=${encodeURIComponent("select " + select)}&_=${Date.now()}`;
    const res = await fetch(url, { cache: "no-store" });
    if (!res.ok) return "err";
    const rows = parseCSV(await res.text());
    return (rows[1] || []).join("|") || "0";
  }
  function objsFrom(rows, header) {
    return rows.map((r) => { const o = {}; header.forEach((h, i) => { if (h) o[h] = r[i]; }); return o; });
  }
  function rowsToObjects(rows) {
    const h = (rows[0] || []).map((x) => String(x).trim());
    return rows.slice(1).filter((r) => r.some((c) => c !== "")).map((r) => { const o = {}; h.forEach((k, i) => { if (k) o[k] = r[i]; }); return o; });
  }

  /* ---------------- State ---------------- */
  const state = {
    prod: [], qual: [], headers: { prod: [], qual: [] }, lastRow: { prod: 0, qual: 0 },
    sig: "", loading: false, first: true, tab: "exec", fullAt: 0, fp: "", syncing: false,
    qFrom: "", qTo: "", pFrom: "", pTo: "",
    filters: { from: "", to: "", section: "", line: "", pg: "", item: "" },
    preset: 30,
  };
  const charts = {};

  function normalizeProd(o) {
    return o.map((r) => ({
      date: toDate(r["Date"]), code: String(r["Item Code"] || "").trim(), item: String(r["Item Name"] || "").trim(),
      line: String(r["Line Name"] || "").trim(), slot: String(r["Time"] || "").trim(), hour: hourOf(r["Time"]),
      mp: num(r["Manpower"]), out: num(r["Production"]), prod: num(r["Productivity"]),
      target: num(r["Target Productivity"]), achv: num(r["Achievement %"]),
      pg: String(r["PG"] || "").trim(), section: String(r["Section"] || "").trim(),
    })).filter((r) => r.date && r.line);
  }
  function normalizeQual(o) {
    return o.map((r) => ({
      date: toDate(r["Date"]), code: String(r["Item Code"] || "").trim(), item: String(r["Item"] || "").trim(),
      line: String(r["Line"] || "").trim(), pg: String(r["PG"] || "").trim(),
      problem: String(r["Problem Type"] || "").trim(), qty: num(r["Defective Qty"]),
    })).filter((r) => r.date && r.problem);
  }

  /* ---------------- Load (full) ---------------- */
  async function loadAll(silent) {
    if (state.loading) return;
    state.loading = true; setLive(true, "Syncing…");
    try {
      const [p, q] = await Promise.all([fetchTab(TABS.prod.name), fetchTab(TABS.qual.name)]);
      state.headers = { prod: (p[0] || []).map((x) => String(x).trim()), qual: (q[0] || []).map((x) => String(x).trim()) };
      state.prod = normalizeProd(rowsToObjects(p));
      state.qual = normalizeQual(rowsToObjects(q));
      state.lastRow = { prod: p.length, qual: q.length };   // last SHEET row (header included)
      state.fullAt = Date.now();
      // data-coverage ranges (quality tab usually starts later than production)
      const pd = state.prod.map((r) => dstr(r.date)).filter(Boolean).sort();
      const qd = state.qual.map((r) => dstr(r.date)).filter(Boolean).sort();
      state.pFrom = pd[0] || ""; state.pTo = pd[pd.length - 1] || "";
      state.qFrom = qd[0] || ""; state.qTo = qd[qd.length - 1] || "";
      state.fp = await signal();                            // baseline change-signal
      if (state.first) initFilters();
      render();
      if (!state.first && !silent) toast("Dashboard synced ✓");
      markLive();
      state.first = false;
      $("#overlay").classList.add("hide");
    } catch (e) {
      setLive(false, "Offline — retrying…");
      $("#loadMsg").textContent = "Could not reach the sheet. Retrying…";
    } finally { state.loading = false; $("#refreshBtn").classList.remove("spin"); }
  }

  /* ---------------- Real-time sync ----------------
     1s change-signal (~100 bytes: row count + key column sums) detects ANY
     add / edit / delete. On change we append new rows instantly, then
     reconcile the rest in the background.                                        */
  const signal = () => Promise.all([
    fetchSignal(TABS.prod, "count(A), sum(F), sum(G)"),   // rows, manpower, production
    fetchSignal(TABS.qual, "count(A), sum(I)"),           // rows, defective qty
  ]).then(([a, b]) => a + "~" + b);

  async function pollTail() {
    if (state.loading || state.first) return 0;
    try {
      const [np, nq] = await Promise.all([
        fetchTail(TABS.prod, state.lastRow.prod),
        fetchTail(TABS.qual, state.lastRow.qual),
      ]);
      let added = 0;
      if (np.length) { state.prod.push(...normalizeProd(objsFrom(np, state.headers.prod))); state.lastRow.prod += np.length; added += np.length; }
      if (nq.length) { state.qual.push(...normalizeQual(objsFrom(nq, state.headers.qual))); state.lastRow.qual += nq.length; added += nq.length; }
      return added;
    } catch (e) { return 0; }
  }

  let reconcileT = null;
  async function checkSignal() {
    if (document.hidden || state.first || state.syncing) return;
    try {
      const fp = await signal();
      if (!state.fp) { state.fp = fp; return; }
      if (fp === state.fp) return;
      state.fp = fp; state.syncing = true;
      const added = await pollTail();
      if (added) { render(); toast(`+${added} new entr${added === 1 ? "y" : "ies"} synced`); }
      markLive(added);
      clearTimeout(reconcileT);
      reconcileT = setTimeout(() => { state.syncing = false; loadAll(true); }, 1200);
    } catch (e) { state.syncing = false; }
  }

  function markLive(added) {
    const t = new Date().toLocaleTimeString();
    setLive(false, `Live · ${t}`);
    $("#lastUpdated").textContent = `${added ? `+${added} new · ` : ""}synced ${t} · real-time (${FAST_MS / 1000}s)`;
  }

  /* ---------------- Filters ---------------- */
  function inRange(r) {
    const f = state.filters;
    if (f.from && dstr(r.date) < f.from) return false;
    if (f.to && dstr(r.date) > f.to) return false;
    if (f.section && r.section !== f.section) return false;
    if (f.line && r.line !== f.line) return false;
    if (f.pg && r.pg !== f.pg) return false;
    if (f.item && (r.item !== f.item && r.code !== f.item)) return false;
    return true;
  }
  const P = () => state.prod.filter(inRange);
  const Q = () => state.qual.filter(inRange);

  /* ---- quality coverage ----------------------------------------------
     The Quality tab usually starts later than Production. Any wastage
     RATIO must use only the output produced inside that overlap, else the
     earlier months (production with no defect records) dilute the PPM.   */
  function alignedWindow() {
    const f = state.filters;
    const from = state.qFrom && (!f.from || f.from < state.qFrom) ? state.qFrom : (f.from || state.qFrom);
    const to = state.qTo && (!f.to || f.to > state.qTo) ? state.qTo : (f.to || state.qTo);
    return { from, to };
  }
  const isClamped = () => { const f = state.filters; return !!(state.qFrom && ((f.from && f.from < state.qFrom) || (f.to && f.to > state.qTo))); };
  const Pq = () => { const w = alignedWindow(); return state.prod.filter((r) => { const k = dstr(r.date); return k >= w.from && k <= w.to && inRange(r); }); };
  const MONA = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
  const fmtShort = (iso) => { if (!iso) return "—"; const p = iso.split("-"); return `${+p[2]}-${MONA[+p[1] - 1]}`; };

  function initFilters() {
    const uq = (f) => [...new Set(state.prod.map(f).filter(Boolean))].sort();
    fill("#fSection", uq((r) => r.section));
    fill("#fLine", uq((r) => r.line));
    fill("#fPg", uq((r) => r.pg));
    const dates = state.prod.map((r) => dstr(r.date)).filter(Boolean).sort();
    if (dates.length) {
      const a = $("#fFrom"), b = $("#fTo");
      a.min = b.min = dates[0]; a.max = b.max = dates[dates.length - 1];
      b.value = dates[dates.length - 1];
      a.value = addDays(dates[dates.length - 1], -30) < dates[0] ? dates[0] : addDays(dates[dates.length - 1], -30);
      state.filters.from = a.value; state.filters.to = b.value;
    }
    ["#fFrom","#fTo","#fSection","#fLine","#fPg"].forEach((id) => $(id).addEventListener("change", () => {
      state.filters.from = $("#fFrom").value; state.filters.to = $("#fTo").value;
      state.filters.section = $("#fSection").value; state.filters.line = $("#fLine").value;
      state.filters.pg = $("#fPg").value; state.filters.item = "";
      state.preset = -1; document.querySelectorAll(".preset").forEach((p) => p.classList.remove("on"));
      render();
    }));
    $("#presets").addEventListener("click", (e) => {
      const b = e.target.closest(".preset"); if (!b) return;
      document.querySelectorAll(".preset").forEach((p) => p.classList.remove("on")); b.classList.add("on");
      const dates = state.prod.map((r) => dstr(r.date)).filter(Boolean).sort();
      if (!dates.length) return;
      const last = dates[dates.length - 1];
      if (b.dataset.all) { $("#fFrom").value = dates[0]; }
      else { $("#fFrom").value = addDays(last, -(+b.dataset.d)); }
      $("#fTo").value = last;
      state.filters.from = $("#fFrom").value; state.filters.to = $("#fTo").value;
      state.preset = b.dataset.all ? 0 : +b.dataset.d;
      render();
    });
    document.querySelectorAll(".tab").forEach((t) => t.addEventListener("click", () => {
      state.tab = t.dataset.tab;
      document.querySelectorAll(".tab").forEach((x) => x.classList.toggle("on", x === t));
      ["exec","action","prod","qual","score"].forEach((k) => {
        const el = $("#v" + k.charAt(0).toUpperCase() + k.slice(1)); if (el) el.classList.toggle("hidden", k !== state.tab);
      });
      render();
    }));
  }
  function fill(sel, vals) {
    const el = $(sel), cur = el.value;
    el.innerHTML = `<option value="">All</option>` + vals.map((v) => `<option value="${esc(v)}">${esc(v)}</option>`).join("");
    if (vals.includes(cur)) el.value = cur;
  }
  function addDays(iso, n) { const d = new Date(iso + "T00:00:00"); d.setDate(d.getDate() + n); return dstr(d); }
  const esc = (s) => String(s == null ? "" : s).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;");
  const nf = (n, d = 0) => Number(n || 0).toLocaleString("en-US", { minimumFractionDigits: d, maximumFractionDigits: d });

  /* ---------------- Aggregation ---------------- */
  const sum = (a, f) => a.reduce((x, r) => x + f(r), 0);

  function agg(rows) {
    let out = 0, opH = 0, stdH = 0, mp = 0;
    rows.forEach((r) => {
      out += r.out; mp += r.mp; opH += r.mp;                       // each row = 1 hour
      if (r.target > 0) stdH += r.out / r.target;                  // hours at standard pace
    });
    const actualProd = opH > 0 ? out / opH : 0;
    const stdProd = stdH > 0 ? out / stdH : 0;
    const actSmv = actualProd > 0 ? 60 / actualProd : 0;
    const stdSmv = stdProd > 0 ? 60 / stdProd : 0;
    return {
      out, mp, opH, stdH, actualProd, stdProd, actSmv, stdSmv,
      achv: opH > 0 ? (stdH / opH) * 100 : 0,
      eff: opH > 0 ? (stdH / opH) * 100 : 0,
      gap: actSmv - stdSmv,
      lostPcs: Math.max(0, opH * stdProd - out),
      lostMin: Math.max(0, (opH - stdH) * 60),
      rows: rows.length,
    };
  }
  function byKey(rows, key) {
    const m = new Map();
    rows.forEach((r) => { const k = key(r) || "(blank)"; if (!m.has(k)) m.set(k, []); m.get(k).push(r); });
    return m;
  }
  function lineData(p, q, pq) {
    const pm = byKey(p, (r) => r.line);
    const qm = byKey(q, (r) => r.line);
    const am = byKey(pq || p, (r) => r.line);
    return [...pm.entries()].map(([line, rows]) => {
      const a = agg(rows);
      const dq = sum(qm.get(line) || [], (r) => r.qty);
      const baseOut = sum(am.get(line) || [], (r) => r.out) || a.out;   // aligned output for wastage
      const section = rows[0].section || "";
      return { line, section, ...a, defects: dq, ppm: baseOut > 0 ? (dq / baseOut) * 1e6 : 0, rate: baseOut > 0 ? (dq / baseOut) * 100 : 0 };
    }).filter((x) => x.out > 0).sort((a, b) => b.lostPcs - a.lostPcs);
  }

  /* ---------------- KPI ---------------- */
  const ICON = {
    out:'<path d="M3 13h4v8H3zm7-6h4v14h-4zm7 3h4v11h-4z"/>',
    achv:'<path d="M12 2 3 7v6c0 5 3.8 9.3 9 10 5.2-.7 9-5 9-10V7l-9-5zm-1 13-3-3 1.4-1.4L11 12.2l4.6-4.6L17 9l-6 6z"/>',
    eff:'<path d="M13 2 3 14h7l-1 8 10-12h-7z"/>',
    smv:'<path d="M12 2a10 10 0 1 0 10 10A10 10 0 0 0 12 2zm1 11h5v-2h-3V7h-2z"/>',
    lost:'<path d="M12 2 1 21h22L12 2zm1 15h-2v-2h2v2zm0-4h-2V9h2v4z"/>',
    def:'<path d="M12 2 4 6v6c0 5 3.4 9.4 8 10 4.6-.6 8-5 8-10V6z"/>',
    mp:'<path d="M16 11a4 4 0 1 0-4-4 4 4 0 0 0 4 4zm-8 1a3 3 0 1 0-3-3 3 3 0 0 0 3 3zm0 2c-2.7 0-6 1.3-6 4v3h8v-3c0-1 .4-1.9 1.1-2.6A8.5 8.5 0 0 0 8 14z"/>',
    lines:'<path d="M3 3h8v8H3zm10 0h8v8h-8zM3 13h8v8H3zm10 0h8v8h-8z"/>',
  };
  function renderKpis(p, q, pq) {
    const a = agg(p), def = sum(q, (r) => r.qty);
    const baseOut = sum(pq, (r) => r.out) || a.out;          // output inside the quality-coverage window
    const ppm = baseOut > 0 ? (def / baseOut) * 1e6 : 0;
    const rate = baseOut > 0 ? (def / baseOut) * 100 : 0;
    const clamped = isClamped();
    const lines = new Set(p.map((r) => r.line)).size;
    const cards = [
      ["Total Output", nf(a.out), `${nf(p.length)} entries`, "out", "#22d3ee", "#0ea5e9"],
      ["Avg Achievement", nf(a.achv, 1) + "%", a.achv >= 95 ? "on / above target" : a.achv >= 85 ? "slightly below target" : "below target", "achv", a.achv >= 95 ? "#34d399" : a.achv >= 85 ? "#f59e0b" : "#ef4444", a.achv >= 95 ? "#10b981" : a.achv >= 85 ? "#d97706" : "#b91c1c"],
      ["Line Efficiency", nf(a.eff, 1) + "%", `standard ÷ actual pace`, "eff", "#34d399", "#0d9488"],
      ["SMV Gap", (a.gap >= 0 ? "+" : "") + nf(a.gap, 3), "min/pc over standard", "smv", "#f59e0b", "#ef4444"],
      ["Lost Capacity", nf(a.lostPcs), "pcs recoverable", "lost", "#38bdf8", "#6366f1"],
      ["Wastage", nf(ppm), `${nf(def)} pcs · ${nf(rate, 2)}% rate${clamped ? ` · ⚠ quality data from ${fmtShort(state.qFrom)}` : ""}`, "def", "#ef4444", "#f97316"],
      ["Manpower Productivity", nf(a.actualProd, 2), `pcs / op-hr · std ${nf(a.stdProd, 2)}`, "mp", "#a855f7", "#6366f1"],
      ["Active Lines", nf(lines), `${new Set(p.map((r) => r.item)).size} items`, "lines", "#f472b6", "#a855f7"],
    ];
    $("#kpis").innerHTML = cards.map(([lab, val, sub, ic, c1, c2]) =>
      `<div class="kpi" style="--k1:${c1};--k2:${c2}">
        <div class="kpi-top"><span class="kpi-ic"><svg viewBox="0 0 24 24" fill="currentColor">${ICON[ic]}</svg></span><span class="lab">${lab}</span></div>
        <div class="val">${val}</div><div class="sub">${sub}</div>
      </div>`).join("");
  }

  /* ---------------- Insight / Action engine ---------------- */
  const FIX = [
    [/screw|fasten/i, "Check the screw feeder / torque driver setting and bin level; add a poka-yoke so a missing screw stops the cycle."],
    [/fitting|assembl|loose/i, "Check fixture & jig alignment on the station and retrain the operator on the standard method (one-point lesson)."],
    [/plastic|component|mould|mold/i, "Audit incoming moulded parts (dimensional + visual) and raise a supplier CAPA; verify machine parameters."],
    [/solder|metal base/i, "Verify soldering temperature profile and tip condition; run a skill check on the soldering operators."],
    [/cosmetic|scratch|spot|appearance|dirt|unclean/i, "Review handling & packing at the workstation; add protective film / separators and reinforce 5S."],
    [/bush|connecting|contact/i, "Check incoming quality of the part and spring/contact tension; raise a supplier CAPA."],
    [/nut/i, "Check the nut-runner torque and calibration; verify incoming nut dimensions."],
    [/switch|switching|function/i, "Add a functional test at end-of-line and check the switch contact assembly."],
    [/rivet/i, "Check riveting tooling, rivet length and press pressure; verify the anvil condition."],
    [/print|black|inject|ink/i, "Check the print head / inkjet nozzle, ink level and print position; verify the marking spec."],
  ];
  const fixFor = (t) => { for (const [re, s] of FIX) if (re.test(t)) return s; return "Run a 5-Why root-cause on this defect at the line and standardise the countermeasure."; };

  function buildInsights(p, q, pq) {
    const lines = lineData(p, q, pq);
    const out = [];
    if (!lines.length) return out;
    const totalOut = sum(lines, (x) => x.out);
    const avgPpm = totalOut > 0 ? (sum(lines, (x) => x.defects) / totalOut) * 1e6 : 0;

    // 1) lowest achievement (volume-weighted impact)
    lines.filter((x) => x.achv < 95).sort((a, b) => b.lostPcs - a.lostPcs).slice(0, 4).forEach((x) => {
      const sev = x.achv < 70 ? "bad" : x.achv < 85 ? "warn" : "info";
      out.push({
        sev, tag: "OUTPUT", score: x.lostPcs,
        title: `${x.line} — running at ${nf(x.achv, 1)}% of target`,
        why: `Produced <b>${nf(x.out)}</b> pcs in <b>${nf(x.opH)}</b> operator-hours against a standard of <b>${nf(x.stdProd, 2)}</b> pcs/op-hr. That is <b>${nf(x.lostPcs)}</b> pcs of lost output this period.`,
        action: "Run a method study + line balancing on this line; verify the standard manpower allocation and remove the biggest stoppage in the shift. Set a daily target board and review at the 9 AM meeting.",
        impact: [["Lost output", nf(x.lostPcs) + " pcs"], ["Efficiency", nf(x.eff, 1) + "%"], ["Actual SMV", nf(x.actSmv, 3)]],
      });
    });

    // 2) SMV gap
    lines.filter((x) => x.gap > 0.05).sort((a, b) => b.lostMin - a.lostMin).slice(0, 2).forEach((x) => {
      out.push({
        sev: x.gap > 0.5 ? "bad" : "warn", tag: "SMV / METHOD", score: x.lostMin * 1.2,
        title: `${x.line} — SMV gap ${nf(x.gap, 3)} min/pc`,
        why: `Standard SMV <b>${nf(x.stdSmv, 3)}</b> vs actual <b>${nf(x.actSmv, 3)}</b> min/pc — <b>${nf(x.lostMin)}</b> operator-minutes above standard.`,
        action: "Break the operation into elements and time-study it. Attack the longest element first: improve the fixture, pre-kit parts at the station, and remove any double handling.",
        impact: [["Extra minutes", nf(x.lostMin) + " min"], ["Efficiency", nf(x.eff, 1) + "%"], ["Std SMV", nf(x.stdSmv, 3)]],
      });
    });

    // 3) wastage
    lines.filter((x) => x.defects > 0).sort((a, b) => b.defects - a.defects).slice(0, 3).forEach((x) => {
      const top = topDefectFor(x.line, q);
      out.push({
        sev: x.ppm > avgPpm * 1.5 ? "bad" : "warn", tag: "WASTAGE", score: x.defects * 60,
        title: `${x.line} — ${nf(x.defects)} defective pcs (${nf(x.ppm)} PPM)`,
        why: top ? `Biggest contributor: <b>${esc(top.problem)}</b> with <b>${nf(top.qty)}</b> pcs (${nf(top.share, 0)}% of this line's rejects).` : `Wastage is ${nf(x.rate, 2)}% of output.`,
        action: top ? fixFor(top.problem) : "Run a 5-Why on the top defect and standardise the countermeasure.",
        impact: [["Defects", nf(x.defects) + " pcs"], ["PPM", nf(x.ppm)], ["Rate", nf(x.rate, 2) + "%"]],
      });
    });

    // 4) worst hour
    const byHour = new Map();
    p.forEach((r) => { if (r.hour < 0) return; if (!byHour.has(r.hour)) byHour.set(r.hour, []); byHour.get(r.hour).push(r); });
    const hours = [...byHour.entries()].map(([h, rows]) => ({ h, ...agg(rows) })).filter((x) => x.out > 0).sort((a, b) => a.achv - b.achv);
    if (hours.length > 2 && hours[0].achv < hours[hours.length - 1].achv - 8) {
      const w = hours[0], b = hours[hours.length - 1];
      out.push({
        sev: "warn", tag: "SHIFT PATTERN", score: (b.achv - w.achv) * 100,
        title: `Weak hour: ${fmtHour(w.h)} at ${nf(w.achv, 1)}%`,
        why: `Best hour <b>${fmtHour(b.h)}</b> achieves <b>${nf(b.achv, 1)}%</b> but <b>${fmtHour(w.h)}</b> only <b>${nf(w.achv, 1)}%</b> — a <b>${nf(b.achv - w.achv, 1)}</b> point swing inside the same day.`,
        action: "Check what is different in that hour — start-up/shift change, material feeding, tea/ prayer break overlap or fatigue. Fix the routine and re-measure.",
        impact: [["Weak hour", nf(w.achv, 1) + "%"], ["Best hour", nf(b.achv, 1) + "%"]],
      });
    }

    // 5) best performer (replicate)
    const best = lines.filter((x) => x.achv >= 95 && x.out > totalOut * 0.01).sort((a, b) => b.achv - a.achv)[0];
    if (best) out.push({
      sev: "good", tag: "REPLICATE", score: -1,
      title: `${best.line} — best practice at ${nf(best.achv, 1)}%`,
      why: `Running <b>${nf(best.achv, 1)}%</b> of target with an actual SMV of only <b>${nf(best.actSmv, 3)}</b> min/pc.`,
      action: "Capture what this line does differently (fixture, method, feeding, supervision) and horizontal-deploy it to the weak lines above.",
      impact: [["Achievement", nf(best.achv, 1) + "%"], ["Efficiency", nf(best.eff, 1) + "%"]],
    });

    return out.sort((a, b) => (b.sev === "good" ? -1e9 : b.score) - (a.sev === "good" ? -1e9 : a.score));
  }
  function topDefectFor(line, q) {
    const rows = q.filter((r) => r.line === line);
    if (!rows.length) return null;
    const m = new Map();
    rows.forEach((r) => m.set(r.problem, (m.get(r.problem) || 0) + r.qty));
    const tot = sum([...m.values()], (v) => v) || 1;
    const [problem, qty] = [...m.entries()].sort((a, b) => b[1] - a[1])[0];
    return { problem, qty, share: (qty / tot) * 100 };
  }
  const fmtHour = (h) => `${String(h % 12 === 0 ? 12 : h % 12).padStart(2, "0")} ${h < 12 ? "AM" : "PM"}`;

  function renderActions(p, q, pq) {
    const items = buildInsights(p, q, pq);
    const bad = items.filter((i) => i.sev === "bad").length;
    $("#acCount").textContent = `${items.length} finding${items.length === 1 ? "" : "s"}`;
    $("#tabAlerts").textContent = bad || items.filter((i) => i.sev !== "good").length;
    $("#acList").innerHTML = items.length ? items.map((it, i) => {
      const c = it.sev === "bad" ? "var(--bad)" : it.sev === "warn" ? "var(--warn)" : it.sev === "good" ? "var(--good)" : "var(--info)";
      return `<div class="ac-item" style="--sev:${c}">
        <div class="rk">${i + 1}</div>
        <div class="ac-body">
          <div class="ac-title">${esc(it.title)}<span class="ac-tag">${it.tag}</span></div>
          <div class="ac-why">${it.why}</div>
          <div class="ac-do"><b>→ What to do:</b> ${esc(it.action)}</div>
          <div class="ac-impact">${it.impact.map(([k, v]) => `<span>${k}<b>${v}</b></span>`).join("")}</div>
        </div>
      </div>`;
    }).join("") : `<div class="ac-empty">No issues found for this selection — all lines are at or above target. 👏</div>`;
  }

  /* ---------------- Charts ---------------- */
  const tc = () => { const cs = getComputedStyle(document.documentElement); return { text: cs.getPropertyValue("--muted").trim(), grid: cs.getPropertyValue("--border").trim() }; };
  function baseOpts(extra) {
    const t = tc();
    return Object.assign({
      responsive: true, maintainAspectRatio: false, animation: { duration: 600 },
      plugins: { legend: { labels: { color: t.text, boxWidth: 12, font: { size: 11 }, usePointStyle: true } }, tooltip: { backgroundColor: "rgba(8,12,22,.95)", borderColor: t.grid, borderWidth: 1, padding: 10, titleColor: "#fff" } },
      scales: { x: { ticks: { color: t.text, font: { size: 10 }, maxRotation: 42, autoSkip: true }, grid: { color: t.grid, drawBorder: false } }, y: { ticks: { color: t.text, font: { size: 10 } }, grid: { color: t.grid, drawBorder: false }, beginAtZero: true } },
    }, extra || {});
  }
  const lerp = (a, b, t) => { const h = (x) => { x = x.replace("#",""); if (x.length===3) x = x.split("").map(c=>c+c).join(""); return [parseInt(x.slice(0,2),16),parseInt(x.slice(2,4),16),parseInt(x.slice(4,6),16)]; };
    const A = h(a), B = h(b); t = Math.max(0, Math.min(1, t || 0));
    return `rgb(${Math.round(A[0]+(B[0]-A[0])*t)},${Math.round(A[1]+(B[1]-A[1])*t)},${Math.round(A[2]+(B[2]-A[2])*t)})`; };
  const statusColor = (v, good, warn) => v >= good ? "#34d399" : v >= warn ? "#f59e0b" : "#ef4444";
  function mk(id, type, data, options) {
    const el = $("#" + id); if (!el) return;
    if (charts[id]) { charts[id].data = data; charts[id].options = options; charts[id].update(); }
    else charts[id] = new Chart(el, { type, data, options });
  }
  function grad(c1, c2) { return (ctx) => { const a = ctx.chart.chartArea; if (!a) return c1; const g = ctx.chart.ctx.createLinearGradient(0, a.bottom, 0, a.top); g.addColorStop(0, c1); g.addColorStop(1, c2); return g; }; }

  function renderCharts(p, q, lines) {
    // trend
    const byDate = new Map();
    p.forEach((r) => { const k = dstr(r.date); if (!byDate.has(k)) byDate.set(k, []); byDate.get(k).push(r); });
    const days = [...byDate.entries()].map(([k, rows]) => { const a = agg(rows); return { k, ...a }; }).sort((a, b) => (a.k < b.k ? -1 : 1));
    mk("cTrend", "line", { labels: days.map((d) => d.k.slice(5)), datasets: [
      { label: "Actual output", data: days.map((d) => d.out), borderColor: "#22d3ee", backgroundColor: grad("rgba(34,211,238,.04)", "rgba(34,211,238,.35)"), fill: true, tension: .35, pointRadius: 0, borderWidth: 2 },
      { label: "Target output", data: days.map((d) => Math.round(d.stdH * d.stdProd)), borderColor: "#34d399", borderDash: [6,4], backgroundColor: "transparent", fill: false, tension: .35, pointRadius: 0, borderWidth: 2 },
    ] }, baseOpts());

    // section doughnut
    const secM = new Map();
    p.forEach((r) => secM.set(r.section, (secM.get(r.section) || 0) + r.out));
    const sec = [...secM.entries()].map(([k, v]) => ({ k, v })).sort((a, b) => b.v - a.v);
    mk("cSection", "doughnut", { labels: sec.map((x) => x.k), datasets: [{ data: sec.map((x) => x.v), backgroundColor: sec.map((_, i) => PALETTE[i % PALETTE.length]), borderWidth: 2, borderColor: "rgba(0,0,0,.25)", hoverOffset: 8 }] },
      { responsive: true, maintainAspectRatio: false, cutout: "58%", plugins: { legend: { position: "right", labels: { color: tc().text, boxWidth: 12, font: { size: 11 }, usePointStyle: true } } } });

    // achievement by line
    const ach = lines.slice().sort((a, b) => a.achv - b.achv).slice(0, 14);
    mk("cAchLine", "bar", { labels: ach.map((x) => x.line), datasets: [{ label: "Achievement %", data: ach.map((x) => +x.achv.toFixed(1)), backgroundColor: ach.map((x) => statusColor(x.achv, 95, 85)), borderRadius: 6, maxBarThickness: 26 }] }, baseOpts());

    // top defect
    const dt = new Map();
    q.forEach((r) => dt.set(r.problem, (dt.get(r.problem) || 0) + r.qty));
    const top = [...dt.entries()].map(([k, v]) => ({ k, v })).sort((a, b) => b.v - a.v).slice(0, 8);
    const o1 = baseOpts(); o1.indexAxis = "y";
    mk("cTopDefect", "bar", { labels: top.map((x) => x.k), datasets: [{ label: "Defective qty", data: top.map((x) => x.v), backgroundColor: top.map((x) => lerp("#f97316", "#ef4444", x.v / (top[0] ? top[0].v : 1))), borderRadius: 6, maxBarThickness: 20 }] }, o1);

    // SMV gap by line
    const gap = lines.slice().sort((a, b) => b.gap - a.gap).slice(0, 14);
    const gmax = Math.max(0.001, ...gap.map((x) => Math.max(0, x.gap)));
    const o2 = baseOpts(); o2.indexAxis = "y";
    mk("cSmvGap", "bar", { labels: gap.map((x) => x.line), datasets: [{ label: "SMV gap (min/pc)", data: gap.map((x) => +x.gap.toFixed(3)), backgroundColor: gap.map((x) => x.gap > 0 ? lerp("#f59e0b", "#ef4444", x.gap / gmax) : "#34d399"), borderRadius: 6, maxBarThickness: 20 }] }, o2);

    // efficiency
    const eff = lines.slice().sort((a, b) => a.eff - b.eff).slice(0, 14);
    mk("cEff", "bar", { labels: eff.map((x) => x.line), datasets: [{ label: "Efficiency %", data: eff.map((x) => +x.eff.toFixed(1)), backgroundColor: eff.map((x) => statusColor(x.eff, 95, 85)), borderRadius: 6, maxBarThickness: 24 }] }, baseOpts());

    // lost capacity
    const lost = lines.filter((x) => x.lostPcs > 0).slice(0, 14);
    mk("cLost", "bar", { labels: lost.map((x) => x.line), datasets: [{ label: "Lost pcs", data: lost.map((x) => Math.round(x.lostPcs)), backgroundColor: grad("#38bdf8", "#6366f1"), borderRadius: 6, maxBarThickness: 24 }] }, baseOpts());

    // manpower productivity
    const mpp = lines.slice().sort((a, b) => b.actualProd - a.actualProd).slice(0, 14);
    mk("cMpp", "bar", { labels: mpp.map((x) => x.line), datasets: [
      { label: "Actual pcs/op-hr", data: mpp.map((x) => +x.actualProd.toFixed(2)), backgroundColor: grad("#a855f7", "#6366f1"), borderRadius: 6, maxBarThickness: 22 },
      { label: "Standard", data: mpp.map((x) => +x.stdProd.toFixed(2)), backgroundColor: "rgba(255,255,255,.18)", borderRadius: 6, maxBarThickness: 22 },
    ] }, baseOpts());

    // hourly
    const hm = new Map();
    p.forEach((r) => { if (r.hour < 0) return; if (!hm.has(r.hour)) hm.set(r.hour, []); hm.get(r.hour).push(r); });
    const hrs = [...hm.entries()].map(([h, rows]) => ({ h, ...agg(rows) })).sort((a, b) => a.h - b.h);
    mk("cHourly", "bar", { labels: hrs.map((x) => fmtHour(x.h)), datasets: [{ label: "Output", data: hrs.map((x) => x.out), backgroundColor: hrs.map((x) => statusColor(x.achv, 95, 85)), borderRadius: 6, maxBarThickness: 30 }] }, baseOpts());

    // pareto
    const dtAll = [...dt.entries()].map(([k, v]) => ({ k, v })).sort((a, b) => b.v - a.v).slice(0, 10);
    const tot = sum(dtAll, (x) => x.v); let cum = 0;
    const cumArr = dtAll.map((x) => { cum += x.v; return tot ? +(cum / tot * 100).toFixed(1) : 0; });
    const po = baseOpts(); po.scales.y1 = { position: "right", beginAtZero: true, max: 100, ticks: { color: tc().text, font: { size: 10 }, callback: (v) => v + "%" }, grid: { drawOnChartArea: false } };
    mk("cPareto", "bar", { labels: dtAll.map((x) => x.k), datasets: [
      { type: "bar", label: "Defective qty", data: dtAll.map((x) => x.v), backgroundColor: dtAll.map((x) => lerp("#f97316", "#ef4444", x.v / (dtAll[0] ? dtAll[0].v : 1))), borderRadius: 6, maxBarThickness: 30, yAxisID: "y", order: 2 },
      { type: "line", label: "Cumulative %", data: cumArr, borderColor: "#22d3ee", backgroundColor: "#22d3ee", tension: .3, pointRadius: 3, borderWidth: 2, yAxisID: "y1", order: 1 },
    ] }, po);

    // ppm by line
    const ppm = lines.slice().sort((a, b) => b.ppm - a.ppm).slice(0, 14);
    mk("cPpm", "bar", { labels: ppm.map((x) => x.line), datasets: [{ label: "Defect PPM", data: ppm.map((x) => Math.round(x.ppm)), backgroundColor: grad("#ef4444", "#b91c1c"), borderRadius: 6, maxBarThickness: 24 }] }, baseOpts());

    // defect trend
    const dm = new Map();
    q.forEach((r) => { const k = dstr(r.date); dm.set(k, (dm.get(k) || 0) + r.qty); });
    const ddays = [...dm.entries()].sort((a, b) => (a[0] < b[0] ? -1 : 1));
    mk("cDefTrend", "line", { labels: ddays.map((x) => x[0].slice(5)), datasets: [{ label: "Defects", data: ddays.map((x) => x[1]), borderColor: "#f43f5e", backgroundColor: grad("rgba(244,63,94,.05)", "rgba(244,63,94,.4)"), fill: true, tension: .35, pointRadius: 0, borderWidth: 2 }] }, baseOpts());

    // defect matrix (stacked bar: line x problem)
    const probs = [...dt.entries()].map(([k, v]) => ({ k, v })).sort((a, b) => b.v - a.v).slice(0, 6).map((x) => x.k);
    const lm = new Map();
    q.forEach((r) => { if (!probs.includes(r.problem)) return; if (!lm.has(r.line)) lm.set(r.line, {}); lm.get(r.line)[r.problem] = (lm.get(r.line)[r.problem] || 0) + r.qty; });
    const larr = [...lm.entries()].map(([k, o]) => ({ k, tot: sum(Object.values(o), (v) => v), o })).sort((a, b) => b.tot - a.tot).slice(0, 10);
    mk("cMatrix", "bar", { labels: larr.map((x) => x.k), datasets: probs.map((pr, i) => ({ label: pr, data: larr.map((x) => x.o[pr] || 0), backgroundColor: PALETTE[i % PALETTE.length], borderRadius: 4, maxBarThickness: 34 })) }, baseOpts({ scales: Object.assign(baseOpts().scales, { x: Object.assign(baseOpts().scales.x, { stacked: true }), y: Object.assign(baseOpts().scales.y, { stacked: true }) }) }));
  }

  /* ---------------- Tables ---------------- */
  function renderTables(lines, p, q, pq) {
    const dq = new Map();
    q.forEach((r) => dq.set(r.line, (dq.get(r.line) || 0) + r.qty));
    const aq = new Map();
    (pq || p).forEach((r) => { const k = r.item || r.code; aq.set(k, (aq.get(k) || 0) + r.out); });
    $("#tblLine").innerHTML = lines.map((x, i) => {
      const st = x.achv >= 95 ? ["good", "On target"] : x.achv >= 85 ? ["warn", "Watch"] : ["bad", "Action"];
      return `<tr data-line="${esc(x.line)}">
        <td><span class="rank">${i + 1}</span></td>
        <td>${esc(x.line)}</td><td>${esc(x.section)}</td>
        <td class="num">${nf(x.out)}</td>
        <td class="num"><span class="tag ${st[0]}">${nf(x.achv, 1)}%</span></td>
        <td class="num">${nf(x.eff, 1)}%</td>
        <td class="num">${nf(x.stdSmv, 3)}</td><td class="num">${nf(x.actSmv, 3)}</td>
        <td class="num">${(x.gap >= 0 ? "+" : "") + nf(x.gap, 3)}</td>
        <td class="num">${nf(x.lostPcs)}</td>
        <td class="num">${nf(x.defects)}</td>
        <td class="num">${nf(x.ppm)}</td>
        <td><span class="tag ${st[0]}">${st[1]}</span></td>
      </tr>`;
    }).join("");
    document.querySelectorAll("#tblLine tr").forEach((tr) => tr.addEventListener("click", () => { state.filters.line = tr.dataset.line; $("#fLine").value = tr.dataset.line; render(); toast("Filtered: " + tr.dataset.line); }));

    const im = new Map();
    p.forEach((r) => { const k = r.item || r.code; if (!im.has(k)) im.set(k, { item: k, pg: r.pg, out: 0, stdH: 0 }); const o = im.get(k); o.out += r.out; if (r.target > 0) o.stdH += r.out / r.target; });
    const dqi = new Map(); q.forEach((r) => { const k = r.item || r.code; dqi.set(k, (dqi.get(k) || 0) + r.qty); });
    const irows = [...im.values()].map((o) => { const base = aq.get(o.item) || o.out; return { ...o, achv: o.stdH > 0 ? (o.out / o.stdH) * 100 : 0, def: dqi.get(o.item) || 0, ppm: base > 0 ? ((dqi.get(o.item) || 0) / base) * 1e6 : 0 }; }).sort((a, b) => b.out - a.out).slice(0, 15);
    $("#tblItem").innerHTML = irows.map((r, i) => `<tr data-item="${esc(r.item)}">
      <td><span class="rank">${i + 1}</span></td><td>${esc(r.item)}</td><td>${esc(r.pg)}</td>
      <td class="num">${nf(r.out)}</td><td class="num">${nf(r.achv, 1)}%</td><td class="num">${nf(r.def)}</td><td class="num">${nf(r.ppm)}</td>
    </tr>`).join("");
    document.querySelectorAll("#tblItem tr").forEach((tr) => tr.addEventListener("click", () => { state.filters.item = tr.dataset.item; render(); toast("Filtered: " + tr.dataset.item); }));
  }

  /* ---------------- Render ---------------- */
  function render() {
    const p = P(), q = Q(), pq = Pq(), lines = lineData(p, q, pq);
    renderKpis(p, q, pq);
    renderActions(p, q, pq);
    renderCharts(p, q, lines);
    renderTables(lines, p, q, pq);
    renderCoverage();
  }

  /* data-coverage banner (Quality tab) */
  function renderCoverage() {
    const el = $("#covNote"); if (!el) return;
    if (!state.qFrom) { el.className = "cov-note warn"; el.innerHTML = "⚠ No quality data loaded — wastage figures are unavailable."; return; }
    const w = alignedWindow(), clamped = isClamped();
    el.className = "cov-note" + (clamped ? " warn" : "");
    el.innerHTML = clamped
      ? `⚠ <b>Wastage is aligned to the overlap window.</b> Production covers <b>${fmtShort(state.pFrom)} → ${fmtShort(state.pTo)}</b>, but quality records only exist from <b>${fmtShort(state.qFrom)}</b> → <b>${fmtShort(state.qTo)}</b>. Wastage PPM is therefore calculated on <b>${fmtShort(w.from)} → ${fmtShort(w.to)}</b> (${nf(sum(Pq(), (r) => r.out))} pcs) so the ratio isn't diluted by earlier months that have no defect records.`
      : `Quality data covers <b>${fmtShort(state.qFrom)} → ${fmtShort(state.qTo)}</b> — fully overlapping the selected period. Wastage PPM is calculated on the actual output in this window.`;
  }

  /* ---------------- misc ---------------- */
  function setLive(up, text) { $("#livePill").classList.toggle("updating", up); $("#liveText").textContent = text; }
  let tt; function toast(m) { const t = $("#toast"); t.textContent = m; t.classList.add("show"); clearTimeout(tt); tt = setTimeout(() => t.classList.remove("show"), 2400); }

  const root = document.documentElement;
  root.setAttribute("data-theme", localStorage.getItem("alel-dpd-theme") || "dark");
  $("#themeBtn").addEventListener("click", () => {
    const n = root.getAttribute("data-theme") === "light" ? "dark" : "light";
    root.setAttribute("data-theme", n); localStorage.setItem("alel-dpd-theme", n);
    Object.values(charts).forEach((c) => c.destroy()); Object.keys(charts).forEach((k) => delete charts[k]);
    render();
  });
  $("#refreshBtn").addEventListener("click", () => { $("#refreshBtn").classList.add("spin"); loadAll(); });

  /* real-time sync: 1s change-signal + 60s full reconcile + on focus/online */
  loadAll();
  setInterval(checkSignal, FAST_MS);
  setInterval(() => { if (!document.hidden) loadAll(true); }, FULL_MS);
  document.addEventListener("visibilitychange", () => {
    if (!document.hidden) { checkSignal(); if (Date.now() - state.fullAt > FULL_MS) loadAll(true); }
  });
  window.addEventListener("online", () => loadAll(true));
})();
