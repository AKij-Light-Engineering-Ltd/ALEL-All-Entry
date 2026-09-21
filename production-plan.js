/* =========================================================
   ALEL — Production Plan  (daily build plan per SKU)
   Reads the planning workbook and computes, for every SKU,
   how much must be produced each remaining working day.
   ========================================================= */
(function () {
  "use strict";

  const SHEET_ID = "1LYws_M4TANxF0O0N3lJ7RixKxsEQSOsTJfa4h5V8pro";
  const SHEET = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/gviz/tq`;
  const GROUP_TAB = "SKU Plan 3M";
  const SNAPSHOT = "plan-snapshot.json";
  // Optional live proxy (Apps Script) — works even when the workbook is private.
  // Paste the /exec URL from apps-script/PlanProxy.gs here.
  const PROXY_URL = "";
  const OFF_DAYS = [5];                     // Friday = weekly off
  const REFRESH_MS = 60000;
  const MONA = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
  const $ = (s) => document.querySelector(s);

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
  const num = (v) => { const n = parseFloat(String(v == null ? "" : v).replace(/[%,]/g, "")); return isNaN(n) ? 0 : n; };
  const xlDate = (n) => { const d = new Date(Date.UTC(1899, 11, 30) + n * 86400000); return d.toISOString().slice(0, 10); };
  const dstr = (d) => d ? `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}` : "";
  const esc = (s) => String(s == null ? "" : s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  const nf = (n, d = 0) => Number(n || 0).toLocaleString("en-US", { minimumFractionDigits: d, maximumFractionDigits: d });
  const sum = (a, f) => a.reduce((x, r) => x + f(r), 0);

  /* ---------------- fetch ---------------- */
  async function tryTab(name) {
    try {
      const res = await fetch(`${SHEET}?tqx=out:csv&sheet=${encodeURIComponent(name)}&_=${Date.now()}`, { cache: "no-store" });
      if (!res.ok) return null;
      const txt = await res.text();
      if (txt.trim().startsWith("<")) return null;
      return parseCSV(txt);
    } catch (e) { return null; }
  }
  function tabCandidates(d) {
    const m = MONA[d.getMonth()], y = String(d.getFullYear()).slice(2);
    return [`${m}${y}-Forecast vs Production`, `${m}${y}-Forecast vs production`, `${m}-Forecast vs Production`, `${m}-Forecast vs production`, `${m}${y} Forecast`, `${m} Forecast`];
  }

  /* ---------------- working days ---------------- */
  const isWorking = (d) => !OFF_DAYS.includes(d.getDay());
  function workDaysInMonth(y, m) { const last = new Date(y, m + 1, 0).getDate(); let n = 0; for (let d = 1; d <= last; d++) if (isWorking(new Date(y, m, d))) n++; return n; }
  function workDaysBetween(y, m, from, to) { let n = 0; for (let d = from; d <= to; d++) if (isWorking(new Date(y, m, d))) n++; return n; }

  /* ---------------- state ---------------- */
  const state = {
    months: [], tab: "", rows: [], groups: {}, loading: false, first: true,
    tabView: "today", search: "", group: "", month: "", asOf: "", daysLeft: 0, userDays: null,
  };
  const charts = {};

  /* ---------------- parse ---------------- */
  function parsePlan(rows) {
    const hdr = (rows[0] || []).map((x) => String(x == null ? "" : x).trim());
    const r1 = (rows[1] || []).map((x) => String(x == null ? "" : x).trim());
    const find = (re, not) => hdr.findIndex((h) => re.test(h) && !(not && not.test(h)));
    const iCode = find(/^code$/i), iName = find(/item\s*name/i);
    const iFc = find(/forecast/i, /value/i), iOpen = find(/opening/i, /value/i);
    const iReq = find(/^req\.?\s*qty$/i), iProd = find(/total\s*production/i), iDel = find(/delivered\s*quantity/i);
    const dateCols = [];
    r1.forEach((v, i) => { const n = Number(v); if (n > 40000 && n < 60000) dateCols.push({ i, date: xlDate(n) }); });
    const out = [];
    for (let r = 2; r < rows.length; r++) {
      const row = rows[r] || [];
      const name = String(row[iName] || "").trim();
      if (!name || /^total/i.test(name)) continue;
      const daily = {};
      dateCols.forEach((c) => { const q = num(row[c.i]); if (q) daily[c.date] = q; });
      out.push({
        code: String(row[iCode] || "").trim(), name,
        fc: num(row[iFc]), open: num(row[iOpen]), req: num(row[iReq]),
        prod: num(row[iProd]), del: num(row[iDel]), daily,
      });
    }
    return { rows: out, dateCols };
  }

  /* ---------------- load (proxy → public sheet → snapshot) ---------------- */
  async function load() {
    if (state.loading) return;
    state.loading = true; setLive(true, "Loading plan…");
    try {
      let data = null, source = "";
      const now = new Date();
      const months = []; for (let b = 0; b <= 4; b++) months.push(new Date(now.getFullYear(), now.getMonth() - b, 1));

      /* 1) live Apps Script proxy — works with a private workbook */
      if (PROXY_URL) {
        try {
          const r = await fetch(`${PROXY_URL}?action=plan&_=${Date.now()}`, { cache: "no-store" });
          const j = await r.json();
          if (j && j.ok && j.rows && j.rows.length) { data = { tab: j.tab, rows: j.rows, dateCols: (j.dateCols || []).map((d) => ({ date: d })), groups: j.groups || {} }; source = "live · proxy"; }
        } catch (e) { /* fall through */ }
      }

      /* 2) public sheet (gviz) */
      if (!data) {
        for (const mo of months) {
          for (const cand of tabCandidates(mo)) {
            const rows = await tryTab(cand);
            if (rows && rows.length > 5) {
              const parsed = parsePlan(rows);
              const gm = await tryTab(GROUP_TAB);
              const groups = {};
              if (gm) gm.slice(4).forEach((r) => { const c = String(r[1] || "").trim(); const g = String(r[7] || "").trim(); if (c && g) groups[c] = g; });
              data = { tab: cand, rows: parsed.rows, dateCols: parsed.dateCols, groups };
              source = "live · sheet public";
              break;
            }
          }
          if (data) break;
        }
      }

      /* 3) bundled snapshot (always works) */
      if (!data) {
        const r = await fetch(`${SNAPSHOT}?_=${Date.now()}`, { cache: "no-store" });
        const j = await r.json();
        data = {
          tab: j.tab,
          rows: j.rows.map((x) => ({ code: x.c, name: x.n, fc: x.f, open: x.o, req: x.q, prod: x.p, del: x.d, daily: x.y || {} })),
          dateCols: (j.dateCols || []).map((d) => ({ date: d })),
          groups: j.groups || {},
        };
        source = "snapshot · " + new Date(j.generatedAt).toLocaleString();
      }

      state.months = months; state.tab = data.tab;
      const mm = (data.tab.match(/([A-Za-z]{3})\s*(\d{2})/) || []);
      const mi = MONA.findIndex((m) => m.toLowerCase() === String(mm[1] || "").toLowerCase());
      state.month = mi >= 0 ? `${2000 + Number(mm[2])}-${String(mi + 1).padStart(2, "0")}` : dstr(new Date()).slice(0, 7);
      state.rows = data.rows; state.dateCols = data.dateCols; state.groups = data.groups;
      state.rows.forEach((r) => { if (!r.group) r.group = state.groups[r.code] || guessGroup(r.name); });
      state.source = source;
      if (state.first) initFilters();
      render();
      const t = new Date().toLocaleTimeString();
      setLive(false, `Plan · ${t}`);
      $("#lastUpdated").textContent = `${source} · tab “${state.tab}” · ${state.rows.length} SKUs · updated ${t}`;
      $("#overlay").classList.add("hide");
      state.first = false;
    } catch (e) {
      setLive(false, "Cannot read the plan");
      $("#loadMsg").innerHTML = "Could not load the plan data.";
    } finally { state.loading = false; $("#refreshBtn").classList.remove("spin"); }
  }

  const GROUP_RX = [
    [/LUMIGOLD/i,"GSS Lumigold"],[/AURORA/i,"GSS Aurora"],[/ONYX/i,"GSS Onyx"],[/CHERRY/i,"GSS Cherry"],
    [/SAFFRON/i,"GSS Saffron"],[/ORCHID/i,"GSS Orchid"],[/PIANO/i,"Piano"],[/\bMCB\b/i,"MCB"],
    [/PVC TAPE/i,"PVC Tape"],[/AC\/DC/i,"AC/DC Bulb"],[/BRACKET TUBE/i,"Bracket Tube"],
    [/HIGH WATT/i,"High Watt LED"],[/PANEL/i,"Panel Light"],[/FLOOD LIGHT/i,"Flood Light"],
    [/EXHAUST FAN/i,"Exhaust Fan"],[/CEILING ROSE/i,"Ceiling Rose"],[/HOLDER/i,"Holder"],
    [/PLUG/i,"Plug"],[/DISTRIBUTION BOX|\bDB\b/i,"DB"],[/EXTENSION SOCKET/i,"Extension Socket"],
    [/CONCEALED BOX/i,"Concealed Box"],[/RICE COOKER/i,"Rice Cooker"],[/GAS STOVE/i,"Gas Stove"],
    [/KETTLE/i,"Electric Kettle"],[/INDUCTION/i,"Induction Cooker"],[/IRON/i,"Electric Iron"],
    [/LED BULB/i,"Regular LED Bulb"],[/CEILING/i,"Ceiling Rose"],[/SOCKET/i,"Socket"],
  ];
  function guessGroup(n) { for (const [re, g] of GROUP_RX) if (re.test(n)) return g; return "Other"; }

  /* ---------------- plan computation ---------------- */
  function planRows() {
    const [y, m] = state.month.split("-").map(Number);
    const mi = m - 1, lastDay = new Date(y, m, 0).getDate();
    const asOfD = state.asOf ? new Date(state.asOf + "T00:00:00") : new Date(y, mi, Math.min(new Date().getDate(), lastDay));
    const asOfDay = Math.min(Math.max(1, asOfD.getDate()), lastDay);
    const totalWD = workDaysInMonth(y, mi);
    const elapsedWD = workDaysBetween(y, mi, 1, asOfDay);
    const autoLeft = workDaysBetween(y, mi, asOfDay + 1, lastDay);
    const daysLeft = state.userDays != null ? state.userDays : autoLeft;
    const rows = state.rows
      .filter((r) => (r.fc > 0 || r.prod > 0 || r.req > 0) && (!state.group || r.group === state.group))
      .filter((r) => !state.search || r.name.toLowerCase().includes(state.search) || r.code.includes(state.search))
      .map((r) => {
        const req = r.req > 0 ? r.req : Math.max(0, r.fc - r.open);
        const balance = Math.max(0, req - r.prod);
        const expected = totalWD > 0 ? (req * elapsedWD) / totalWD : 0;
        const achv = req > 0 ? (r.prod / req) * 100 : (r.fc > 0 ? 100 : 0);
        const target = daysLeft > 0 ? Math.ceil(balance / daysLeft) : balance;
        let status = "done";
        if (balance > 0) status = r.prod >= expected ? "ontrack" : (r.prod < expected * 0.7 ? "critical" : "behind");
        return { ...r, req, balance, expected, achv, target, status };
      });
    return { rows, daysLeft, totalWD, elapsedWD, asOfDay, lastDay, autoLeft };
  }

  /* ---------------- KPI ---------------- */
  const ICON = {
    fc:'<path d="M3 13h4v8H3zm7-6h4v14h-4zm7 3h4v11h-4z"/>',
    req:'<path d="M12 2 3 7v6c0 5 3.8 9.3 9 10 5.2-.7 9-5 9-10V7l-9-5zm-1 13-3-3 1.4-1.4L11 12.2l4.6-4.6L17 9l-6 6z"/>',
    prod:'<path d="M13 2 3 14h7l-1 8 10-12h-7z"/>',
    bal:'<path d="M12 2 1 21h22L12 2zm1 15h-2v-2h2v2zm0-4h-2V9h2v4z"/>',
    days:'<path d="M12 2a10 10 0 1 0 10 10A10 10 0 0 0 12 2zm1 11h5v-2h-3V7h-2z"/>',
    tgt:'<path d="M12 2 2 7v6c0 5.2 4.3 9.4 10 10 5.7-.6 10-4.8 10-10V7l-10-5zm0 6a4 4 0 1 1-4 4 4 4 0 0 1 4-4zm0 2a2 2 0 1 0 2 2 2 2 0 0 0-2-2z"/>',
  };
  function renderKpis(P) {
    const rows = P.rows;
    const fc = sum(rows, (r) => r.fc), req = sum(rows, (r) => r.req), prod = sum(rows, (r) => r.prod);
    const bal = sum(rows, (r) => r.balance), tgt = sum(rows, (r) => r.target);
    const achv = req > 0 ? (prod / req) * 100 : 0;
    const behind = rows.filter((r) => r.status === "behind" || r.status === "critical").length;
    const cards = [
      ["Month Forecast", nf(fc), `${rows.length} SKUs in plan`, "fc", "#22d3ee", "#0ea5e9"],
      ["Required Production", nf(req), `forecast − opening stock`, "req", "#a855f7", "#6366f1"],
      ["Produced to Date", nf(prod), `${nf(achv, 1)}% of requirement`, "prod", achv >= 90 ? "#34d399" : achv >= 70 ? "#f59e0b" : "#ef4444", achv >= 90 ? "#10b981" : achv >= 70 ? "#d97706" : "#b91c1c"],
      ["Balance to Produce", nf(bal), `${rows.filter((r) => r.balance > 0).length} SKUs pending`, "bal", "#f59e0b", "#ef4444"],
      ["Working Days Left", nf(P.daysLeft), `${P.elapsedWD} of ${P.totalWD} elapsed`, "days", "#38bdf8", "#6366f1"],
      ["Today's Total Target", nf(tgt), behind ? `${behind} SKUs behind schedule` : "all on track", "tgt", behind ? "#ef4444" : "#34d399", behind ? "#b91c1c" : "#0d9488"],
    ];
    $("#kpis").innerHTML = cards.map(([lab, val, sub, ic, c1, c2]) =>
      `<div class="kpi" style="--k1:${c1};--k2:${c2}">
        <div class="kpi-top"><span class="kpi-ic"><svg viewBox="0 0 24 24" fill="currentColor">${ICON[ic]}</svg></span><span class="lab">${lab}</span></div>
        <div class="val">${val}</div><div class="sub">${sub}</div>
      </div>`).join("");
    const note = $("#banner");
    note.className = "note" + (behind ? " warn" : "");
    note.innerHTML = behind
      ? `⚠ <b>${behind} SKU${behind === 1 ? "" : "s"} behind schedule.</b> Daily targets below already include the catch-up quantity so the monthly requirement is still met — prioritise the red rows.`
      : `✅ <b>Every SKU is on or ahead of schedule.</b> The daily targets below keep the plan on track to finish the month's requirement.`;
  }

  /* ---------------- tables ---------------- */
  function statusTag(s) {
    return s === "done" ? `<span class="tag good">Completed</span>`
      : s === "ontrack" ? `<span class="tag good">On track</span>`
      : s === "behind" ? `<span class="tag warn">Behind</span>`
      : `<span class="tag bad">Critical</span>`;
  }
  function renderToday(P) {
    const rows = P.rows.filter((r) => r.balance > 0).sort((a, b) => b.target - a.target).slice(0, 60);
    $("#tblToday").innerHTML = rows.length ? rows.map((r, i) => `<tr>
      <td><span class="rank">${i + 1}</span></td>
      <td>${esc(r.name)}</td><td>${esc(r.group)}</td>
      <td class="num">${nf(r.fc)}</td><td class="num">${nf(r.req)}</td><td class="num">${nf(r.prod)}</td>
      <td class="num">${nf(r.balance)}</td>
      <td class="num"><span class="tgt">${nf(r.target)}</span></td>
      <td>${statusTag(r.status)}</td>
    </tr>`).join("") : `<tr><td colspan="9" style="padding:22px;color:var(--muted)">🎉 Nothing left to build — every SKU in this selection has met its requirement.</td></tr>`;
  }
  function renderSku(P) {
    const rows = P.rows.slice().sort((a, b) => b.balance - a.balance || b.fc - a.fc);
    $("#tblSku").innerHTML = rows.map((r, i) => `<tr>
      <td><span class="rank">${i + 1}</span></td>
      <td>${esc(r.name)}</td><td>${esc(r.group)}</td>
      <td class="num">${nf(r.fc)}</td><td class="num">${nf(r.open)}</td><td class="num">${nf(r.req)}</td>
      <td class="num">${nf(r.prod)}</td>
      <td class="num"><span class="tag ${r.achv >= 100 ? "good" : r.achv >= 70 ? "warn" : "bad"}">${nf(Math.min(r.achv, 999), 0)}%</span></td>
      <td class="num">${nf(r.balance)}</td>
      <td class="num"><span class="tgt">${nf(r.target)}</span></td>
      <td>${statusTag(r.status)}</td>
    </tr>`).join("");
  }
  function groupRollup(P) {
    const m = new Map();
    P.rows.forEach((r) => {
      if (!m.has(r.group)) m.set(r.group, { group: r.group, n: 0, fc: 0, req: 0, prod: 0, bal: 0, tgt: 0 });
      const o = m.get(r.group); o.n++; o.fc += r.fc; o.req += r.req; o.prod += r.prod; o.bal += r.balance; o.tgt += r.target;
    });
    return [...m.values()].map((o) => ({ ...o, achv: o.req > 0 ? (o.prod / o.req) * 100 : 0 })).sort((a, b) => b.bal - a.bal || b.fc - a.fc);
  }
  function renderGroup(P) {
    const g = groupRollup(P);
    $("#tblGroup").innerHTML = g.map((r, i) => `<tr data-g="${esc(r.group)}">
      <td><span class="rank">${i + 1}</span></td><td>${esc(r.group)}</td><td class="num">${r.n}</td>
      <td class="num">${nf(r.fc)}</td><td class="num">${nf(r.req)}</td><td class="num">${nf(r.prod)}</td>
      <td class="num"><span class="tag ${r.achv >= 90 ? "good" : r.achv >= 60 ? "warn" : "bad"}">${nf(r.achv, 0)}%</span></td>
      <td class="num">${nf(r.bal)}</td><td class="num"><span class="tgt">${nf(r.tgt)}</span></td>
    </tr>`).join("");
    document.querySelectorAll("#tblGroup tr").forEach((tr) => tr.addEventListener("click", () => {
      state.group = tr.dataset.g; $("#fGroup").value = state.group; render();
    }));
    const top = g.slice(0, 14);
    mk("cGroup", "bar", { labels: top.map((x) => x.group), datasets: [
      { label: "Required", data: top.map((x) => x.req), backgroundColor: "rgba(255,255,255,.18)", borderRadius: 6, maxBarThickness: 24 },
      { label: "Produced", data: top.map((x) => x.prod), backgroundColor: "#22d3ee", borderRadius: 6, maxBarThickness: 24 },
      { label: "Balance", data: top.map((x) => x.bal), backgroundColor: "#f59e0b", borderRadius: 6, maxBarThickness: 24 },
    ] }, baseOpts());
    const pie = g.filter((x) => x.bal > 0).slice(0, 12);
    mk("cGroupPie", "doughnut", { labels: pie.map((x) => x.group), datasets: [{ data: pie.map((x) => x.bal), backgroundColor: PALETTE, borderWidth: 2, borderColor: "rgba(0,0,0,.25)", hoverOffset: 8 }] }, pieOpts());
  }
  function renderProgress(P) {
    const dc = state.dateCols || [];
    const totals = dc.map((c) => sum(state.rows, (r) => r.daily[c.date] || 0));
    const cum = []; let run = 0;
    totals.forEach((v) => { run += v; cum.push(run); });
    const grandReq = sum(state.rows, (r) => r.req);
    const planLine = dc.map((_, i) => Math.round((grandReq * (i + 1)) / dc.length));
    mk("cCum", "line", { labels: dc.map((c) => c.date.slice(5)), datasets: [
      { label: "Actual (cumulative)", data: cum, borderColor: "#22d3ee", backgroundColor: grad("rgba(34,211,238,.04)", "rgba(34,211,238,.35)"), fill: true, tension: .3, pointRadius: 0, borderWidth: 2 },
      { label: "Required pace", data: planLine, borderColor: "#34d399", borderDash: [6, 4], backgroundColor: "transparent", fill: false, tension: .3, pointRadius: 0, borderWidth: 2 },
    ] }, baseOpts());
    mk("cDaily", "bar", { labels: dc.map((c) => c.date.slice(5)), datasets: [{ label: "Output", data: totals, backgroundColor: "#6366f1", borderRadius: 5, maxBarThickness: 26 }] }, baseOpts());
    const g = groupRollup(P).slice(0, 12);
    mk("cAchv", "bar", { labels: g.map((x) => x.group), datasets: [{ label: "Achievement %", data: g.map((x) => +x.achv.toFixed(1)), backgroundColor: g.map((x) => (x.achv >= 90 ? "#34d399" : x.achv >= 60 ? "#f59e0b" : "#ef4444")), borderRadius: 6, maxBarThickness: 26 }] }, baseOpts());
  }

  /* ---------------- charts helpers ---------------- */
  const PALETTE = ["#22d3ee","#6366f1","#a855f7","#34d399","#f59e0b","#ef4444","#14b8a6","#eab308","#f472b6","#38bdf8","#fb923c","#4ade80"];
  const tc = () => { const cs = getComputedStyle(document.documentElement); return { text: cs.getPropertyValue("--muted").trim(), grid: cs.getPropertyValue("--border").trim() }; };
  function baseOpts() { const t = tc(); return { responsive: true, maintainAspectRatio: false, animation: { duration: 600 },
    plugins: { legend: { labels: { color: t.text, boxWidth: 12, font: { size: 11 }, usePointStyle: true } }, tooltip: { backgroundColor: "rgba(8,12,22,.95)", borderColor: t.grid, borderWidth: 1, padding: 10, titleColor: "#fff" } },
    scales: { x: { ticks: { color: t.text, font: { size: 10 }, maxRotation: 42, autoSkip: true }, grid: { color: t.grid, drawBorder: false } }, y: { ticks: { color: t.text, font: { size: 10 } }, grid: { color: t.grid, drawBorder: false }, beginAtZero: true } } }; }
  function pieOpts() { const t = tc(); return { responsive: true, maintainAspectRatio: false, cutout: "58%", plugins: { legend: { position: "right", labels: { color: t.text, boxWidth: 12, font: { size: 11 }, usePointStyle: true } } } }; }
  function grad(c1, c2) { return (ctx) => { const a = ctx.chart.chartArea; if (!a) return c1; const g = ctx.chart.ctx.createLinearGradient(0, a.bottom, 0, a.top); g.addColorStop(0, c1); g.addColorStop(1, c2); return g; }; }
  function mk(id, type, data, options) { const el = $("#" + id); if (!el) return; if (charts[id]) { charts[id].data = data; charts[id].options = options; charts[id].update(); } else charts[id] = new Chart(el, { type, data, options }); }

  /* ---------------- filters ---------------- */
  function initFilters() {
    const gsel = $("#fGroup"); const cur = gsel.value;
    const groups = [...new Set(state.rows.map((r) => r.group).filter(Boolean))].sort();
    gsel.innerHTML = `<option value="">All</option>` + groups.map((g) => `<option value="${esc(g)}">${esc(g)}</option>`).join("");
    if (groups.includes(cur)) gsel.value = cur;
    const [y, m] = state.month.split("-").map(Number);
    const lastDay = new Date(y, m, 0).getDate();
    const a = $("#fAsOf"); a.min = `${state.month}-01`; a.max = `${state.month}-${String(lastDay).padStart(2, "0")}`;
    if (!a.value) { const t = new Date(); a.value = (t.getFullYear() === y && t.getMonth() + 1 === m) ? dstr(t) : a.max; }
    state.asOf = a.value;
    const P = planRows(); $("#fDays").value = P.daysLeft;
    gsel.onchange = () => { state.group = gsel.value; render(); };
    a.onchange = () => { state.asOf = a.value; state.userDays = null; render(); };
    $("#fDays").oninput = () => { const v = parseInt($("#fDays").value, 10); state.userDays = isNaN(v) ? null : v; render(); };
    const si = $("#skuSearch"), sc = $("#skuClear");
    si.oninput = () => { state.search = si.value.trim().toLowerCase(); sc.hidden = !si.value; render(); };
    si.onkeydown = (e) => { if (e.key === "Escape") { si.value = ""; state.search = ""; sc.hidden = true; render(); } };
    sc.onclick = () => { si.value = ""; state.search = ""; sc.hidden = true; render(); si.focus(); };
    document.querySelectorAll(".tab").forEach((t) => t.onclick = () => {
      state.tabView = t.dataset.tab;
      document.querySelectorAll(".tab").forEach((x) => x.classList.toggle("on", x === t));
      ["today","sku","group","progress"].forEach((k) => { const el = $("#v" + k.charAt(0).toUpperCase() + k.slice(1)); if (el) el.classList.toggle("hidden", k !== state.tabView); });
      render();
    });
  }

  /* ---------------- render ---------------- */
  function render() {
    const P = planRows();
    renderKpis(P);
    if (state.tabView === "today") renderToday(P);
    else if (state.tabView === "sku") renderSku(P);
    else if (state.tabView === "group") renderGroup(P);
    else renderProgress(P);
  }

  function setLive(up, text) { $("#livePill").classList.toggle("updating", up); $("#liveText").textContent = text; }
  const root = document.documentElement;
  root.setAttribute("data-theme", localStorage.getItem("alel-plan-theme") || "dark");
  $("#themeBtn").onclick = () => { const n = root.getAttribute("data-theme") === "light" ? "dark" : "light"; root.setAttribute("data-theme", n); localStorage.setItem("alel-plan-theme", n); Object.values(charts).forEach((c) => c.destroy()); Object.keys(charts).forEach((k) => delete charts[k]); render(); };
  $("#refreshBtn").onclick = () => { $("#refreshBtn").classList.add("spin"); load(); };

  load();
  setInterval(() => { if (!document.hidden) load(); }, REFRESH_MS);
})();
