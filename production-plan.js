/* =========================================================
   ALEL — Production Planning Dashboard
   Build plan · time & manpower · capacity load · adherence
   Data: planning workbook (private, via Apps Script proxy)
         + productivity targets (public master workbook)
   ========================================================= */
(function () {
  "use strict";

  const SHEET_ID = "1LYws_M4TANxF0O0N3lJ7RixKxsEQSOsTJfa4h5V8pro";
  const SHEET = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/gviz/tq`;
  const GROUP_TAB = "SKU Plan 3M";
  const SNAPSHOT = "plan-snapshot.json";
  const PROXY_URL = "https://script.google.com/macros/s/AKfycbxP5ZKxXIiHlzYtvdYkfkWdByP1W_n0y9t1S52V6OQuZn9N4WoPMC59OcHSdcR2cW7Fqw/exec";
  // productivity targets + line mapping (public master workbook)
  const TP_SHEET_ID = "1HeKJ0XueH0WeAxLrJpNaNz_pvPuxRDElQmHcPd7QfOQ";
  const TP_SHEET = `https://docs.google.com/spreadsheets/d/${TP_SHEET_ID}/gviz/tq`;
  const TP_TAB = "Master Data";
  const OFF_DAYS = [5];                     // Friday = weekly off
  const REFRESH_MS = 60000;
  const MONA = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
  const $ = (s) => document.querySelector(s);
  const PALETTE = ["#22d3ee","#6366f1","#a855f7","#34d399","#f59e0b","#ef4444","#14b8a6","#eab308","#f472b6","#38bdf8","#fb923c","#4ade80"];

  /* ---------------- utils ---------------- */
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
  const xlDate = (n) => new Date(Date.UTC(1899, 11, 30) + n * 86400000).toISOString().slice(0, 10);
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
  /* productivity target (pcs/op-hr) from Master Data,
     item → line/table mapping derived from actual production entries */
  async function fetchMaster() {
    const out = { tp: {}, line: {} };
    try {
      const res = await fetch(`${TP_SHEET}?tqx=out:csv&sheet=${encodeURIComponent(TP_TAB)}&_=${Date.now()}`, { cache: "no-store" });
      if (res.ok) {
        const rows = parseCSV(await res.text());
        rows.slice(1).forEach((r) => { const c = String(r[0] || "").trim(); const tp = num(r[2]); if (c && tp > 0) out.tp[c] = tp; });
      }
    } catch (e) { /* optional */ }
    try {
      const res = await fetch(`${TP_SHEET}?tqx=out:csv&sheet=${encodeURIComponent("Production_Data")}&_=${Date.now()}`, { cache: "no-store" });
      if (res.ok) {
        const rows = parseCSV(await res.text());
        const counts = {};
        rows.slice(1).forEach((r) => {
          const c = String(r[1] || "").trim(), ln = String(r[3] || "").trim();
          if (!c || !ln) return;
          (counts[c] = counts[c] || {})[ln] = (counts[c][ln] || 0) + 1;
        });
        Object.keys(counts).forEach((c) => {
          const best = Object.entries(counts[c]).sort((a, b) => b[1] - a[1])[0];
          if (best) out.line[c] = best[0];
        });
      }
    } catch (e) { /* optional */ }
    return out;
  }

  /* ---------------- working days ---------------- */
  const isWorking = (d) => !OFF_DAYS.includes(d.getDay());
  function workDaysInMonth(y, m) { const last = new Date(y, m + 1, 0).getDate(); let n = 0; for (let d = 1; d <= last; d++) if (isWorking(new Date(y, m, d))) n++; return n; }
  function workDaysBetween(y, m, from, to) { let n = 0; for (let d = from; d <= to; d++) if (isWorking(new Date(y, m, d))) n++; return n; }

  /* ---------------- state ---------------- */
  const state = {
    rows: [], dateCols: [], groups: {}, targets: {}, lines: {}, month: "", asOf: "", userDays: null,
    shiftHours: 8, availMp: null, tabView: "brief", search: "", group: "", tab: "", source: "",
    loading: false, first: true,
  };
  const charts = {};

  /* ---------------- parse ---------------- */
  function parsePlan(rows) {
    const hdr = (rows[0] || []).map((x) => String(x == null ? "" : x).trim());
    const find = (re, not) => hdr.findIndex((h) => re.test(h) && !(not && not.test(h)));
    const iCode = find(/^code$/i), iName = find(/item\s*name/i);
    const iFc = find(/forecast/i, /value/i), iOpen = find(/opening/i, /value/i);
    const iReq = find(/^req\.?\s*qty$/i), iProd = find(/total\s*production/i), iDel = find(/delivered\s*quantity/i);
    let dateCols = [], dateIdx = [];
    for (let s = 1; s <= Math.min(3, rows.length - 1); s++) {
      const cand = rows[s] || [], cols = [], idx = [];
      cand.forEach((v, i) => { const n = num(v); if (n > 40000 && n < 60000) { cols.push(xlDate(n)); idx.push(i); } });
      if (cols.length > dateCols.length) { dateCols = cols; dateIdx = idx; }
    }
    const out = [];
    for (let r = 2; r < rows.length; r++) {
      const row = rows[r] || [];
      const name = String(row[iName] || "").trim();
      if (!name || /^total/i.test(name)) continue;
      const daily = {};
      dateIdx.forEach((ci, k) => { const q = num(row[ci]); if (q) daily[dateCols[k]] = q; });
      out.push({ code: String(row[iCode] || "").trim(), name, fc: num(row[iFc]), open: num(row[iOpen]), req: num(row[iReq]), prod: num(row[iProd]), del: num(row[iDel]), daily });
    }
    return { rows: out, dateCols };
  }

  /* ---------------- load ---------------- */
  async function load() {
    if (state.loading) return;
    state.loading = true; setLive(true, "Loading plan…");
    try {
      let data = null, source = "";
      const now = new Date();
      const months = []; for (let b = 0; b <= 4; b++) months.push(new Date(now.getFullYear(), now.getMonth() - b, 1));

      if (PROXY_URL) {
        try {
          const r = await fetch(`${PROXY_URL}?action=plan&_=${Date.now()}`, { cache: "no-store" });
          const j = await r.json();
          if (j && j.ok && j.rows && j.rows.length) { data = { tab: j.tab, rows: j.rows, dateCols: (j.dateCols || []).map((d) => ({ date: d })), groups: j.groups || {} }; source = "live · proxy"; }
        } catch (e) { /* fall through */ }
      }
      if (!data) {
        for (const mo of months) {
          for (const cand of tabCandidates(mo)) {
            const rows = await tryTab(cand);
            if (rows && rows.length > 5) {
              const parsed = parsePlan(rows), gm = await tryTab(GROUP_TAB), groups = {};
              if (gm) gm.slice(4).forEach((r) => { const c = String(r[1] || "").trim(); const g = String(r[7] || "").trim(); if (c && g) groups[c] = g; });
              data = { tab: cand, rows: parsed.rows, dateCols: parsed.dateCols, groups }; source = "live · sheet public"; break;
            }
          }
          if (data) break;
        }
      }
      if (!data) {
        const j = await (await fetch(`${SNAPSHOT}?_=${Date.now()}`, { cache: "no-store" })).json();
        data = { tab: j.tab, rows: j.rows.map((x) => ({ code: x.c, name: x.n, fc: x.f, open: x.o, req: x.q, prod: x.p, del: x.d, daily: x.y || {} })), dateCols: (j.dateCols || []).map((d) => ({ date: d })), groups: j.groups || {} };
        source = "snapshot · " + new Date(j.generatedAt).toLocaleString();
      }

      state.tab = data.tab; state.rows = data.rows; state.dateCols = data.dateCols; state.groups = data.groups;
      const mm = (data.tab.match(/([A-Za-z]{3})\s*(\d{2})/) || []);
      const mi = MONA.findIndex((m) => m.toLowerCase() === String(mm[1] || "").toLowerCase());
      state.month = mi >= 0 ? `${2000 + Number(mm[2])}-${String(mi + 1).padStart(2, "0")}` : dstr(new Date()).slice(0, 7);

      const master = await fetchMaster();
      state.targets = master.tp; state.lines = master.line;
      state.rows.forEach((r) => {
        if (!r.group) r.group = state.groups[r.code] || guessGroup(r.name);
        if (!r.line) r.line = state.lines[r.code] || "Unassigned";
      });
      state.source = source;
      if (state.first) initFilters();
      render();
      const t = new Date().toLocaleTimeString();
      setLive(false, `Plan · ${t}`);
      $("#lastUpdated").textContent = `${source} · “${state.tab}” · ${state.rows.length} SKUs · ${Object.keys(state.targets).length} productivity targets · updated ${t}`;
      $("#overlay").classList.add("hide");
      state.first = false;
    } catch (e) {
      setLive(false, "Cannot load the plan");
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
    [/LED BULB/i,"Regular LED Bulb"],[/SOCKET/i,"Socket"],
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
    const shiftH = state.shiftHours > 0 ? state.shiftHours : 8;

    const rows = state.rows
      .filter((r) => (r.fc > 0 || r.prod > 0 || r.req > 0) && (!state.group || r.group === state.group))
      .filter((r) => !state.search || r.name.toLowerCase().includes(state.search) || r.code.includes(state.search))
      .map((r) => {
        const req = r.req > 0 ? r.req : Math.max(0, r.fc - r.open);
        const balance = Math.max(0, req - r.prod);
        const expected = totalWD > 0 ? (req * elapsedWD) / totalWD : 0;
        const achv = req > 0 ? (r.prod / req) * 100 : (r.fc > 0 ? 100 : 0);
        const target = daysLeft > 0 ? Math.ceil(balance / daysLeft) : balance;
        const tp = state.targets[r.code] || 0;
        const opHours = tp > 0 ? target / tp : 0;                 // operator-hours per day for this SKU
        const hoursToClear = tp > 0 ? balance / tp : 0;           // total effort left
        const mpNeeded = tp > 0 ? Math.ceil(opHours / shiftH) : 0;
        let status = "done";
        if (balance > 0) status = r.prod >= expected ? "ontrack" : (r.prod < expected * 0.7 ? "critical" : "behind");
        return { ...r, req, balance, expected, achv, target, tp, opHours, hoursToClear, mpNeeded, status, line: r.line || "Unassigned" };
      });
    return { rows, daysLeft, totalWD, elapsedWD, asOfDay, lastDay, autoLeft, shiftH };
  }

  /* ---------------- KPI ---------------- */
  const ICON = {
    fc:'<path d="M3 13h4v8H3zm7-6h4v14h-4zm7 3h4v11h-4z"/>',
    bal:'<path d="M12 2 1 21h22L12 2zm1 15h-2v-2h2v2zm0-4h-2V9h2v4z"/>',
    tgt:'<path d="M12 2 2 7v6c0 5.2 4.3 9.4 10 10 5.7-.6 10-4.8 10-10V7l-10-5zm0 6a4 4 0 1 1-4 4 4 4 0 0 1 4-4zm0 2a2 2 0 1 0 2 2 2 2 0 0 0-2-2z"/>',
    hrs:'<path d="M12 2a10 10 0 1 0 10 10A10 10 0 0 0 12 2zm1 11h5v-2h-3V7h-2z"/>',
    mp:'<path d="M16 11a4 4 0 1 0-4-4 4 4 0 0 0 4 4zm-8 1a3 3 0 1 0-3-3 3 3 0 0 0 3 3zm0 2c-2.7 0-6 1.3-6 4v3h8v-3c0-1 .4-1.9 1.1-2.6A8.5 8.5 0 0 0 8 14z"/>',
    days:'<path d="M19 3h-1V1h-2v2H8V1H6v2H5a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2V5a2 2 0 0 0-2-2zm0 16H5V9h14z"/>',
  };
  function renderKpis(P) {
    const rows = P.rows;
    const fc = sum(rows, (r) => r.fc), req = sum(rows, (r) => r.req), prod = sum(rows, (r) => r.prod);
    const bal = sum(rows, (r) => r.balance), tgt = sum(rows, (r) => r.target);
    const opH = sum(rows, (r) => r.opHours), mp = sum(rows, (r) => r.mpNeeded);
    const achv = req > 0 ? (prod / req) * 100 : 0;
    const behind = rows.filter((r) => r.status === "behind" || r.status === "critical").length;
    const cards = [
      ["Today's Target", nf(tgt), `pcs across ${rows.filter((r) => r.balance > 0).length} SKUs`, "tgt", "#22d3ee", "#0ea5e9"],
      ["Operator-Hours", nf(opH, 1), `at ${P.shiftH}h shift`, "hrs", "#a855f7", "#6366f1"],
      ["Manpower Needed", nf(mp), `for today's plan`, "mp", "#38bdf8", "#6366f1"],
      ["Balance to Produce", nf(bal), `${nf(req)} required · ${nf(prod)} done`, "bal", "#f59e0b", "#ef4444"],
      ["Month Forecast", nf(fc), `${rows.length} SKUs in plan`, "fc", "#14b8a6", "#0d9488"],
      ["Working Days Left", nf(P.daysLeft), `${P.elapsedWD} of ${P.totalWD} elapsed`, "days", behind ? "#ef4444" : "#34d399", behind ? "#b91c1c" : "#0d9488"],
    ];
    $("#kpis").innerHTML = cards.map(([lab, val, sub, ic, c1, c2]) =>
      `<div class="kpi" style="--k1:${c1};--k2:${c2}">
        <div class="kpi-top"><span class="kpi-ic"><svg viewBox="0 0 24 24" fill="currentColor">${ICON[ic]}</svg></span><span class="lab">${lab}</span></div>
        <div class="val">${val}</div><div class="sub">${sub}</div>
      </div>`).join("");
    $("#tabPending").textContent = rows.filter((r) => r.balance > 0).length;
  }

  /* ---------------- Morning brief ---------------- */
  const SEV = { critical: 3, behind: 2, ontrack: 1, done: 0 };
  const statusTag = (s) => s === "done" ? `<span class="tag good">Completed</span>`
    : s === "ontrack" ? `<span class="tag good">On track</span>`
    : s === "behind" ? `<span class="tag warn">Behind</span>` : `<span class="tag bad">Critical</span>`;

  function priorityRows(P) {
    return P.rows.filter((r) => r.balance > 0)
      .sort((a, b) => (SEV[b.status] - SEV[a.status]) || (b.opHours - a.opHours) || (b.balance - a.balance));
  }
  function renderBrief(P) {
    const rows = P.rows, pri = priorityRows(P);
    const tgt = sum(rows, (r) => r.target), opH = sum(rows, (r) => r.opHours), mp = sum(rows, (r) => r.mpNeeded);
    const crit = rows.filter((r) => r.status === "critical").length;
    const behind = rows.filter((r) => r.status === "behind").length;
    const ok = rows.filter((r) => r.status === "ontrack" || r.status === "done").length;
    const noTp = rows.filter((r) => r.balance > 0 && !r.tp).length;
    const [y, m] = state.month.split("-").map(Number);
    const dayName = ["Sunday","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday"][new Date(state.asOf + "T00:00:00").getDay()];
    const tone = crit ? "bad" : behind ? "warn" : "good";
    const toneTxt = crit ? `${crit} critical · ${behind} behind` : behind ? `${behind} behind schedule` : "all on track";

    let capTxt = "";
    if (state.availMp) {
      const avail = state.availMp * P.shiftH;
      const gap = opH - avail;
      capTxt = gap > 0
        ? `Short by <b>${nf(gap, 1)}</b> op-hrs — need ~<b>${nf(Math.ceil(opH / P.shiftH))}</b> people, have ${nf(state.availMp)}`
        : `Surplus <b>${nf(-gap, 1)}</b> op-hrs — capacity covers today's plan`;
    }

    $("#heroBrief").innerHTML = `
      <div class="hb-top">
        <div>
          <div class="hb-date">${dayName}, ${state.asOf}</div>
          <div class="hb-sub">${state.tab} · ${P.elapsedWD} of ${P.totalWD} working days elapsed · <b>${P.daysLeft} working days left</b></div>
        </div>
        <span class="hb-pill ${tone}">${toneTxt}</span>
      </div>
      <div class="hb-grid">
        <div class="hb-cell"><div class="l">Build today</div><div class="v">${nf(tgt)}</div><div class="s">pieces</div></div>
        <div class="hb-cell"><div class="l">Operator-hours</div><div class="v">${nf(opH, 1)}</div><div class="s">at ${P.shiftH}h shift</div></div>
        <div class="hb-cell"><div class="l">Manpower needed</div><div class="v">${nf(mp)}</div><div class="s">operators</div></div>
        <div class="hb-cell"><div class="l">Items to build</div><div class="v">${nf(pri.length)}</div><div class="s">of ${nf(rows.length)} in plan</div></div>
        <div class="hb-cell"><div class="l">Balance left</div><div class="v">${nf(sum(rows, (r) => r.balance))}</div><div class="s">pieces this month</div></div>
        <div class="hb-cell ${crit ? "bad" : behind ? "warn" : "good"}"><div class="l">At risk</div><div class="v">${nf(crit + behind)}</div><div class="s">${nf(ok)} on track</div></div>
      </div>`;

    $("#tblPriority").innerHTML = pri.slice(0, 25).map((r, i) => `<tr>
      <td><span class="rank">${i + 1}</span></td><td>${esc(r.name)}</td><td>${esc(r.group)}</td>
      <td class="num">${nf(r.balance)}</td>
      <td class="num">${r.tp ? nf(r.tp) : "—"}</td>
      <td class="num"><span class="tgt">${nf(r.target)}</span></td>
      <td class="num"><span class="hrs">${r.opHours ? nf(r.opHours, 1) : "—"}</span></td>
      <td class="num">${r.mpNeeded ? nf(r.mpNeeded) : "—"}</td>
      <td>${statusTag(r.status)}</td>
    </tr>`).join("") || `<tr><td colspan="9" style="padding:22px;color:var(--muted)">🎉 Nothing left to build in this selection.</td></tr>`;

    /* alerts */
    const alerts = [];
    if (crit) {
      const worst = rows.filter((r) => r.status === "critical").sort((a, b) => b.hoursToClear - a.hoursToClear).slice(0, 3);
      alerts.push(["bad", "🔴", `${crit} SKU${crit === 1 ? "" : "s"} critically behind`, `Biggest effort: ${worst.map((x) => `<b>${esc(x.name)}</b> (${nf(x.hoursToClear, 0)} op-hrs left)`).join(", ")}. Their daily targets already include catch-up.`]);
    }
    if (behind) alerts.push(["warn", "🟠", `${behind} SKU${behind === 1 ? "" : "s"} behind schedule`, "Running under the required pace. Prioritise them above the on-track items in the queue."]);
    if (noTp) {
      const names = rows.filter((r) => r.balance > 0 && !r.tp).slice(0, 4).map((r) => r.name);
      alerts.push(["warn", "📋", `${noTp} item${noTp === 1 ? "" : "s"} without a productivity target`, `No pcs/op-hr standard found in Master Data, so time cannot be planned: ${names.map(esc).join(", ")}${noTp > 4 ? "…" : ""}`]);
    }
    if (state.availMp) {
      const avail = state.availMp * P.shiftH, gap = opH - avail;
      if (gap > 0) alerts.push(["bad", "⏱️", "Capacity shortfall today", `Today's plan needs <b>${nf(opH, 1)}</b> operator-hours but you have <b>${nf(avail, 1)}</b>. Short by <b>${nf(gap, 1)}</b> — either add manpower, extend the shift, or defer the lowest-priority SKUs.`]);
      else alerts.push(["good", "✅", "Capacity is sufficient", `Today's plan needs <b>${nf(opH, 1)}</b> op-hrs; available <b>${nf(avail, 1)}</b> — a surplus of <b>${nf(-gap, 1)}</b>.`]);
    } else {
      alerts.push(["info", "💡", "Enter available manpower", `Put your operator count in the filter above and the dashboard will check whether today's plan fits your capacity.`]);
    }
    const biggest = pri[0];
    if (biggest) alerts.push(["info", "🎯", "Start with this", `<b>${esc(biggest.name)}</b> — ${nf(biggest.balance)} pcs left, ${nf(biggest.target)} needed today${biggest.opHours ? `, about <b>${nf(biggest.opHours, 1)}</b> operator-hours` : ""}.`]);
    const lines = lineRollup(P);
    if (lines.length) { const top = lines[0]; alerts.push(["info", "🏭", "Heaviest line today", `<b>${esc(top.line)}</b> carries <b>${nf(top.opH, 1)}</b> operator-hours across ${top.n} SKUs — the likely constraint.`]); }
    if (!alerts.length) alerts.push(["good", "✅", "Nothing to flag", "Every SKU is on or ahead of schedule."]);

    $("#alertList").innerHTML = alerts.map(([sev, ic, t, d]) => {
      const c = sev === "bad" ? "var(--bad)" : sev === "warn" ? "var(--warn)" : sev === "good" ? "var(--good)" : "var(--info)";
      return `<div class="alert" style="--sev:${c}"><div class="ai">${ic}</div><div><div class="at">${t}</div><div class="ad">${d}</div></div></div>`;
    }).join("");
  }

  /* ---------------- line rollup ---------------- */
  function lineRollup(P) {
    const m = new Map();
    P.rows.forEach((r) => {
      const k = r.line || "Unassigned";
      if (!m.has(k)) m.set(k, { line: k, n: 0, tgt: 0, opH: 0, mp: 0, bal: 0, hoursToClear: 0 });
      const o = m.get(k);
      o.n++; o.tgt += r.target; o.opH += r.opHours; o.mp += r.mpNeeded; o.bal += r.balance; o.hoursToClear += r.hoursToClear;
    });
    return [...m.values()].filter((x) => x.n).sort((a, b) => b.opH - a.opH);
  }

  /* ---------------- tables ---------------- */
  function renderToday(P) {
    const rows = P.rows.filter((r) => r.balance > 0).sort((a, b) => b.target - a.target).slice(0, 80);
    $("#tblToday").innerHTML = rows.map((r, i) => `<tr>
      <td><span class="rank">${i + 1}</span></td><td>${esc(r.name)}</td><td>${esc(r.group)}</td>
      <td class="num">${nf(r.fc)}</td><td class="num">${nf(r.req)}</td><td class="num">${nf(r.prod)}</td>
      <td class="num">${nf(r.balance)}</td>
      <td class="num">${r.tp ? nf(r.tp) : "—"}</td>
      <td class="num"><span class="tgt">${nf(r.target)}</span></td>
      <td class="num"><span class="hrs">${r.opHours ? nf(r.opHours, 1) : "—"}</span></td>
      <td class="num">${r.mpNeeded ? nf(r.mpNeeded) : "—"}</td>
      <td>${statusTag(r.status)}</td>
    </tr>`).join("") || `<tr><td colspan="12" style="padding:22px;color:var(--muted)">🎉 Nothing left to build in this selection.</td></tr>`;
  }
  function renderSku(P) {
    const rows = P.rows.slice().sort((a, b) => b.balance - a.balance || b.fc - a.fc);
    $("#tblSku").innerHTML = rows.map((r, i) => `<tr>
      <td><span class="rank">${i + 1}</span></td><td>${esc(r.name)}</td><td>${esc(r.group)}</td>
      <td class="num">${nf(r.fc)}</td><td class="num">${nf(r.open)}</td><td class="num">${nf(r.req)}</td>
      <td class="num">${nf(r.prod)}</td>
      <td class="num"><span class="tag ${r.achv >= 100 ? "good" : r.achv >= 70 ? "warn" : "bad"}">${nf(Math.min(r.achv, 999))}%</span></td>
      <td class="num">${nf(r.balance)}</td>
      <td class="num">${r.tp ? nf(r.tp) : "—"}</td>
      <td class="num"><span class="tgt">${nf(r.target)}</span></td>
      <td class="num"><span class="hrs">${r.hoursToClear ? nf(r.hoursToClear, 1) : "—"}</span></td>
      <td class="num">${r.mpNeeded ? nf(r.mpNeeded) : "—"}</td>
      <td>${statusTag(r.status)}</td>
    </tr>`).join("");
  }
  function groupRollup(P) {
    const m = new Map();
    P.rows.forEach((r) => {
      if (!m.has(r.group)) m.set(r.group, { group: r.group, n: 0, fc: 0, req: 0, prod: 0, bal: 0, tgt: 0, opH: 0, hoursToClear: 0 });
      const o = m.get(r.group);
      o.n++; o.fc += r.fc; o.req += r.req; o.prod += r.prod; o.bal += r.balance; o.tgt += r.target; o.opH += r.opHours; o.hoursToClear += r.hoursToClear;
    });
    return [...m.values()].map((o) => ({ ...o, achv: o.req > 0 ? (o.prod / o.req) * 100 : 0 })).sort((a, b) => b.bal - a.bal || b.fc - a.fc);
  }
  function renderGroup(P) {
    const g = groupRollup(P);
    $("#tblGroup").innerHTML = g.map((r, i) => `<tr data-g="${esc(r.group)}">
      <td><span class="rank">${i + 1}</span></td><td>${esc(r.group)}</td><td class="num">${r.n}</td>
      <td class="num">${nf(r.fc)}</td><td class="num">${nf(r.req)}</td><td class="num">${nf(r.prod)}</td>
      <td class="num"><span class="tag ${r.achv >= 90 ? "good" : r.achv >= 60 ? "warn" : "bad"}">${nf(r.achv)}%</span></td>
      <td class="num">${nf(r.bal)}</td><td class="num"><span class="tgt">${nf(r.tgt)}</span></td>
      <td class="num"><span class="hrs">${nf(r.opH, 1)}</span></td>
    </tr>`).join("");
    document.querySelectorAll("#tblGroup tr").forEach((tr) => tr.addEventListener("click", () => { state.group = tr.dataset.g; $("#fGroup").value = state.group; render(); }));
    const top = g.slice(0, 14);
    mk("cGroup", "bar", { labels: top.map((x) => x.group), datasets: [
      { label: "Required", data: top.map((x) => x.req), backgroundColor: "rgba(255,255,255,.18)", borderRadius: 6, maxBarThickness: 24 },
      { label: "Produced", data: top.map((x) => x.prod), backgroundColor: "#22d3ee", borderRadius: 6, maxBarThickness: 24 },
      { label: "Balance", data: top.map((x) => x.bal), backgroundColor: "#f59e0b", borderRadius: 6, maxBarThickness: 24 },
    ] }, baseOpts());
    const pie = g.filter((x) => x.bal > 0).slice(0, 12);
    mk("cGroupPie", "doughnut", { labels: pie.map((x) => x.group), datasets: [{ data: pie.map((x) => x.bal), backgroundColor: PALETTE, borderWidth: 2, borderColor: "rgba(0,0,0,.25)", hoverOffset: 8 }] }, pieOpts());
  }

  /* ---------------- capacity ---------------- */
  function renderCapacity(P) {
    const L = lineRollup(P);
    const top = L.slice(0, 14);
    const avail = state.availMp ? state.availMp * P.shiftH : 0;
    mk("cLineLoad", "bar", { labels: top.map((x) => x.line), datasets: [
      { label: "Op-hrs needed today", data: top.map((x) => +x.opH.toFixed(1)), backgroundColor: "#a855f7", borderRadius: 6, maxBarThickness: 22 },
      { label: "Op-hrs to clear balance", data: top.map((x) => +x.hoursToClear.toFixed(0)), backgroundColor: "rgba(34,211,238,.55)", borderRadius: 6, maxBarThickness: 22 },
    ] }, baseOpts());
    const pie = L.filter((x) => x.opH > 0).slice(0, 12);
    mk("cLoadPie", "doughnut", { labels: pie.map((x) => x.line), datasets: [{ data: pie.map((x) => +x.opH.toFixed(1)), backgroundColor: PALETTE, borderWidth: 2, borderColor: "rgba(0,0,0,.25)", hoverOffset: 8 }] }, pieOpts());
    const g = groupRollup(P).filter((x) => x.hoursToClear > 0).sort((a, b) => b.hoursToClear - a.hoursToClear).slice(0, 12);
    mk("cGroupHours", "bar", { labels: g.map((x) => x.group), datasets: [{ label: "Op-hrs to clear", data: g.map((x) => +x.hoursToClear.toFixed(0)), backgroundColor: "#f59e0b", borderRadius: 6, maxBarThickness: 26 }] }, baseOpts());
    const eff = P.rows.filter((r) => r.hoursToClear > 0).sort((a, b) => b.hoursToClear - a.hoursToClear).slice(0, 10);
    mk("cTopEffort", "bar", { labels: eff.map((x) => x.name.length > 26 ? x.name.slice(0, 25) + "…" : x.name), datasets: [{ label: "Op-hrs to clear", data: eff.map((x) => +x.hoursToClear.toFixed(1)), backgroundColor: "#38bdf8", borderRadius: 6, maxBarThickness: 22 }] }, baseOpts({ indexAxis: "y" }));

    const maxH = Math.max(1, ...L.map((x) => x.opH));
    $("#tblLine").innerHTML = L.map((x, i) => {
      const load = x.opH / maxH;
      const cls = load > 0.75 ? "bad" : load > 0.4 ? "warn" : "good";
      return `<tr>
        <td><span class="rank">${i + 1}</span></td><td>${esc(x.line)}</td><td class="num">${x.n}</td>
        <td class="num"><span class="tgt">${nf(x.tgt)}</span></td>
        <td class="num"><span class="hrs">${nf(x.opH, 1)}</span></td>
        <td class="num">${nf(x.mp)}</td>
        <td class="num">${nf(x.bal)}</td>
        <td class="num">${nf(x.hoursToClear, 0)}</td>
        <td><span class="tag ${cls}">${load > 0.75 ? "Heavy" : load > 0.4 ? "Medium" : "Light"}</span></td>
      </tr>`;
    }).join("") || `<tr><td colspan="9" style="padding:22px;color:var(--muted)">No line data — add a “Line Name” column in Master Data to enable line loading.</td></tr>`;
  }

  /* ---------------- progress ---------------- */
  function renderProgress(P) {
    const dc = state.dateCols || [];
    const totals = dc.map((c) => sum(state.rows, (r) => r.daily[c.date] || 0));
    let run = 0; const cum = totals.map((v) => (run += v));
    const grandReq = sum(state.rows, (r) => r.req);
    const planLine = dc.map((_, i) => Math.round((grandReq * (i + 1)) / Math.max(1, dc.length)));
    mk("cCum", "line", { labels: dc.map((c) => c.date.slice(5)), datasets: [
      { label: "Actual (cumulative)", data: cum, borderColor: "#22d3ee", backgroundColor: grad("rgba(34,211,238,.04)", "rgba(34,211,238,.35)"), fill: true, tension: .3, pointRadius: 0, borderWidth: 2 },
      { label: "Required pace", data: planLine, borderColor: "#34d399", borderDash: [6, 4], backgroundColor: "transparent", fill: false, tension: .3, pointRadius: 0, borderWidth: 2 },
    ] }, baseOpts());
    mk("cDaily", "bar", { labels: dc.map((c) => c.date.slice(5)), datasets: [{ label: "Output", data: totals, backgroundColor: "#6366f1", borderRadius: 5, maxBarThickness: 26 }] }, baseOpts());
    const g = groupRollup(P).slice(0, 12);
    mk("cAchv", "bar", { labels: g.map((x) => x.group), datasets: [{ label: "Achievement %", data: g.map((x) => +x.achv.toFixed(1)), backgroundColor: g.map((x) => (x.achv >= 90 ? "#34d399" : x.achv >= 60 ? "#f59e0b" : "#ef4444")), borderRadius: 6, maxBarThickness: 26 }] }, baseOpts());
  }

  /* ---------------- chart helpers ---------------- */
  const tc = () => { const cs = getComputedStyle(document.documentElement); return { text: cs.getPropertyValue("--muted").trim(), grid: cs.getPropertyValue("--border").trim() }; };
  function baseOpts(extra) { const t = tc(); return Object.assign({ responsive: true, maintainAspectRatio: false, animation: { duration: 600 },
    plugins: { legend: { labels: { color: t.text, boxWidth: 12, font: { size: 11 }, usePointStyle: true } }, tooltip: { backgroundColor: "rgba(8,12,22,.95)", borderColor: t.grid, borderWidth: 1, padding: 10, titleColor: "#fff" } },
    scales: { x: { ticks: { color: t.text, font: { size: 10 }, maxRotation: 42, autoSkip: true }, grid: { color: t.grid, drawBorder: false } }, y: { ticks: { color: t.text, font: { size: 10 } }, grid: { color: t.grid, drawBorder: false }, beginAtZero: true } } }, extra || {}); }
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
    $("#fDays").value = planRows().daysLeft;
    gsel.onchange = () => { state.group = gsel.value; render(); };
    a.onchange = () => { state.asOf = a.value; state.userDays = null; $("#fDays").value = planRows().daysLeft; render(); };
    $("#fDays").oninput = () => { const v = parseInt($("#fDays").value, 10); state.userDays = isNaN(v) ? null : v; render(); };
    $("#fShift").oninput = () => { const v = parseFloat($("#fShift").value); state.shiftHours = isNaN(v) || v <= 0 ? 8 : v; render(); };
    $("#fMp").oninput = () => { const v = parseFloat($("#fMp").value); state.availMp = isNaN(v) || v <= 0 ? null : v; render(); };
    const si = $("#skuSearch"), sc = $("#skuClear");
    si.oninput = () => { state.search = si.value.trim().toLowerCase(); sc.hidden = !si.value; render(); };
    si.onkeydown = (e) => { if (e.key === "Escape") { si.value = ""; state.search = ""; sc.hidden = true; render(); } };
    sc.onclick = () => { si.value = ""; state.search = ""; sc.hidden = true; render(); si.focus(); };
    document.querySelectorAll(".tab").forEach((t) => t.onclick = () => {
      state.tabView = t.dataset.tab;
      document.querySelectorAll(".tab").forEach((x) => x.classList.toggle("on", x === t));
      ["brief","today","capacity","sku","group","progress"].forEach((k) => { const el = $("#v" + k.charAt(0).toUpperCase() + k.slice(1)); if (el) el.classList.toggle("hidden", k !== state.tabView); });
      render();
    });
  }

  /* ---------------- render ---------------- */
  function render() {
    const P = planRows();
    renderKpis(P);
    if (state.tabView === "brief") renderBrief(P);
    else if (state.tabView === "today") renderToday(P);
    else if (state.tabView === "capacity") renderCapacity(P);
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
