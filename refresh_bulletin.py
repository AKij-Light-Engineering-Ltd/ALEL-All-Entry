# -*- coding: utf-8 -*-
"""
ALEL Operation Bulletin - One-click refresh tool  (MODERN UI)
Downloads the Google Sheet workbook (all tabs), extracts every product
operation study + product image, and regenerates a modern searchable HTML.

Run directly:  python refresh_bulletin.py
or double-click: Refresh Operation Bulletin.bat
"""
import os
import re
import sys
import io
import json
import base64
import zipfile
import urllib.request
import datetime

SHEET_ID = "1GpALH7TsJ0eW-Ko8ejzV8oQF073dhe7Eb9930o0qvSo"
HERE = os.path.dirname(os.path.abspath(__file__))
XLSX_TMP = os.path.join(HERE, "_latest_capacity_study.xlsx")
OUT_HTML = os.path.join(HERE, "ALEL_Operation_Bulletin_Searchable.html")

MAX_SIDE = 520  # px for embedded product images


def log(msg):
    print(msg)


# --------------------------------------------------------------------------
# 1) Download workbook
# --------------------------------------------------------------------------
def download():
    url = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=xlsx"
    log("[1/4] Downloading Google Sheet (all tabs) ...")
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"})
    last = None
    for attempt in range(1, 4):
        try:
            with urllib.request.urlopen(req, timeout=300) as r:
                data = r.read()
            if len(data) < 100000:
                raise RuntimeError("Unexpected small response: %d bytes" % len(data))
            with open(XLSX_TMP, "wb") as f:
                f.write(data)
            log("      saved %.1f MB" % (len(data) / 1048576))
            return
        except Exception as e:
            last = e
            log("      attempt %d failed: %s" % (attempt, e))
    raise SystemExit("Download failed after 3 tries: %s" % last)


# --------------------------------------------------------------------------
# 2) Parse workbook
# --------------------------------------------------------------------------
def num_try(v):
    if v is None:
        return None
    s = str(v).strip().replace(",", "")
    if not s or s in ("#N/A", "#REF!", "#DIV/0!", "N/A", "None"):
        return None
    try:
        return float(s)
    except ValueError:
        return s


def colname(i):
    s = ""
    i += 1
    while i > 0:
        i, rem = divmod(i - 1, 26)
        s = chr(65 + rem) + s
    return s


NAME_RE = re.compile(r"(?i)^\s*product\s*name\s*[:=\-]?\s*(.*?)\s*$")
GROUP_RE = re.compile(r"(?i)^\s*product\s*group\s*[:=\-]?\s*(.*?)\s*$")
CODE_RE = re.compile(r"(?i)^\s*code\s*[:=\-]?\s*(.*?)\s*$")
SMV_RE = re.compile(r"(?i)^\s*total\s*smv\s*[:=\-]?\s*(.*?)\s*$")
MP_RE = re.compile(r"(?i)^\s*total\s*manpower\s*[:=\-]?\s*(.*?)\s*$")
HC_RE = re.compile(r"(?i)^\s*hourly\s*capacity\s*[:=\-]?\s*(.*?)\s*$")
PROD_RE = re.compile(r"(?i)^\s*productivity\s*[:=\-]?\s*(.*?)\s*$")


def parse_sheet(ws, sno2img):
    rows = []
    for row in ws.iter_rows():
        rows.append([None if c.value is None else str(c.value).strip() for c in row])

    header = {}
    for r in rows[:7]:
        for i, v in enumerate(r):
            if not v:
                continue

            def nxt():
                for j in range(i + 1, len(r)):
                    if r[j]:
                        return r[j]
                return None

            m = NAME_RE.match(v)
            if m and "name" not in header:
                header["name"] = (m.group(1) or nxt() or "")
                continue
            m = GROUP_RE.match(v)
            if m and "group" not in header:
                header["group"] = (m.group(1) or nxt() or "")
                continue
            m = CODE_RE.match(v)
            if m and "code" not in header:
                header["code"] = (m.group(1) or nxt() or "")
                continue
            m = SMV_RE.match(v)
            if m and "smv" not in header:
                header["smv"] = num_try(m.group(1) or nxt())
                continue
            m = MP_RE.match(v)
            if m and "mp" not in header:
                header["mp"] = num_try(m.group(1) or nxt())
                continue
            m = HC_RE.match(v)
            if m and "hc" not in header:
                header["hc"] = num_try(m.group(1) or nxt())
                continue
            m = PROD_RE.match(v)
            if m and "prod" not in header:
                header["prod"] = num_try(m.group(1) or nxt())
                continue

    meta = {"sheet": ws.title, "header": header, "ops": [], "alloc": [], "image": sno2img.get(ws.title)}

    hdr_idx = None
    for idx, r in enumerate(rows):
        for v in r:
            if v and v.strip().lower() == "operation / process":
                hdr_idx = idx
                break
        if hdr_idx is not None:
            break
    if hdr_idx is None:
        return meta

    maxcol = 0
    for i, v in enumerate(rows[hdr_idx]):
        if v:
            maxcol = i

    for r in rows[hdr_idx + 1:]:
        if not r or r[0] is None or str(r[0]).strip() == "":
            break
        try:
            float(str(r[0]).replace(",", ""))
        except ValueError:
            break
        op = {}
        for i in range(maxcol + 1):
            op[colname(i)] = r[i] if i < len(r) else None
        meta["ops"].append(op)

    alloc = []
    started = False
    for r in rows[hdr_idx + 1 + len(meta["ops"]):]:
        cells = [x for x in r if x]
        if not cells:
            if started:
                break
            continue
        matched = False
        for cc in cells:
            mm = re.fullmatch(r"\s*(.+?)\s*-\s*(\d+(?:\.\d+)?)\s*", cc)
            if mm:
                alloc.append({"k": mm.group(1).strip(), "v": mm.group(2)})
                matched = True
        if matched:
            started = True
        else:
            break
    meta["alloc"] = alloc
    return meta


def parse_all():
    log("[2/4] Parsing workbook tabs ...")
    import openpyxl

    z = zipfile.ZipFile(XLSX_TMP)
    wbxml = z.read("xl/workbook.xml").decode("utf-8")
    wbrels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    sheet_ids = re.findall(r'<sheet [^>]*name="([^"]+)"[^>]*r:id="(rId\d+)"', wbxml)
    rid2tgt = dict(re.findall(r'Id="(rId\d+)"[^>]*Target="([^"]+)"', wbrels))
    sno2img = {}
    for nm, rid in sheet_ids:
        m = re.search(r"sheet(\d+)\.xml", rid2tgt.get(rid, ""))
        if not m:
            continue
        sno = m.group(1)
        try:
            srels = z.read(f"xl/worksheets/_rels/sheet{sno}.xml.rels").decode("utf-8")
        except KeyError:
            continue
        dm = re.search(r'Target="\.\./drawings/(drawing\d+\.xml)"', srels)
        if not dm:
            continue
        dnum = re.search(r"drawing(\d+)", dm.group(1)).group(1)
        try:
            drels = z.read(f"xl/drawings/_rels/drawing{dnum}.xml.rels").decode("utf-8")
        except KeyError:
            continue
        imgs = re.findall(r'Target="\.\./media/(image\d+\.png)"', drels)
        if imgs:
            sno2img[nm] = imgs[0]

    media_bytes = {}
    for name in set(sno2img.values()):
        try:
            media_bytes[name] = z.read(f"xl/media/{name}")
        except KeyError:
            pass
    z.close()

    wb = openpyxl.load_workbook(XLSX_TMP, read_only=True, data_only=True)
    results = []
    for nm in wb.sheetnames:
        m = parse_sheet(wb[nm], sno2img)
        if m["ops"] and (m["header"].get("name") or m["header"].get("code")):
            results.append(m)
    wb.close()
    log("      %d product bulletins found" % len(results))
    return results, media_bytes


# --------------------------------------------------------------------------
# 3) Generate HTML  (MODERN)
# --------------------------------------------------------------------------
def image_uri(name, media_bytes, cache):
    if name in cache:
        return cache[name]
    from PIL import Image
    im = Image.open(io.BytesIO(media_bytes[name]))
    if im.mode in ("RGBA", "P", "LA"):
        im = im.convert("RGBA")
        bg = Image.new("RGB", im.size, (255, 255, 255))
        bg.paste(im, mask=im.split()[-1])
        im = bg
    else:
        im = im.convert("RGB")
    im.thumbnail((MAX_SIDE, MAX_SIDE), Image.LANCZOS)
    buf = io.BytesIO()
    im.save(buf, format="JPEG", quality=84, optimize=True)
    uri = "data:image/jpeg;base64," + base64.b64encode(buf.getvalue()).decode("ascii")
    cache[name] = uri
    return uri


def esc(s):
    if s is None:
        return ""
    return str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;")


def num(v):
    if v is None:
        return None
    try:
        return float(str(v).replace(",", ""))
    except Exception:
        return None


def fmt_sl(v):
    n = num(v)
    if n is None:
        return esc(v)
    return str(int(n)) if n == int(n) else str(n)


def fmt_hr(v):
    n = num(v)
    return "&ndash;" if n is None else f"{n:,.0f}"


def fmt_dec(v, d=4):
    n = num(v)
    if n is None:
        return "&ndash;"
    return f"{n:,.{d}f}".rstrip("0").rstrip(".")


def fmt_pct(v, d=0):
    n = num(v)
    if n is None:
        return "&ndash;"
    return f"{n * 100:.{d}f}%"


def fmt_time(v):
    n = num(v)
    if n is None:
        return ""
    return str(int(n)) if n == int(n) else f"{n:.1f}"


ICONS = {
    "smv": '<svg viewBox="0 0 24 24" width="20" height="20" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="9"/><path d="M12 7v5l3 3"/></svg>',
    "mp": '<svg viewBox="0 0 24 24" width="20" height="20" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M17 21v-2a4 4 0 0 0-4-4H7a4 4 0 0 0-4 4v2"/><circle cx="10" cy="7" r="4"/><path d="M23 21v-2a4 4 0 0 0-3-3.87"/><path d="M16 3.13a4 4 0 0 1 0 7.75"/></svg>',
    "hc": '<svg viewBox="0 0 24 24" width="20" height="20" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M13 2L3 14h7l-1 8 10-12h-7l1-8z"/></svg>',
    "prod": '<svg viewBox="0 0 24 24" width="20" height="20" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="9"/><circle cx="12" cy="12" r="5"/><circle cx="12" cy="12" r="1"/></svg>',
    "ops": '<svg viewBox="0 0 24 24" width="20" height="20" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M8 6h13M8 12h13M8 18h13"/><circle cx="3.5" cy="6" r="1.4" fill="currentColor" stroke="none"/><circle cx="3.5" cy="12" r="1.4" fill="currentColor" stroke="none"/><circle cx="3.5" cy="18" r="1.4" fill="currentColor" stroke="none"/></svg>',
    "clock": '<svg viewBox="0 0 24 24" width="14" height="14" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="9"/><path d="M12 7v5l3 2"/></svg>',
    "code": '<svg viewBox="0 0 24 24" width="14" height="14" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round"><polyline points="16 18 22 12 16 6"/><polyline points="8 6 2 12 8 18"/></svg>',
    "folder": '<svg viewBox="0 0 24 24" width="14" height="14" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round"><path d="M22 19a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h5l2 3h9a2 2 0 0 1 2 2z"/></svg>',
    "sheet": '<svg viewBox="0 0 24 24" width="14" height="14" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/></svg>',
    "cal": '<svg viewBox="0 0 24 24" width="14" height="14" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="4" width="18" height="18" rx="2"/><line x1="16" y1="2" x2="16" y2="6"/><line x1="8" y1="2" x2="8" y2="6"/><line x1="3" y1="10" x2="21" y2="10"/></svg>',
    "search": '<svg viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>',
    "refresh": '<svg viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round"><path d="M21 12a9 9 0 1 1-2.64-6.36"/><polyline points="21 3 21 9 15 9"/></svg>',
}



def build_html(results, media_bytes, out_path=None):
    log("[3/4] Generating modern HTML ...")
    img_cache = {}
    today = datetime.date.today().strftime("%d-%b-%Y")
    bullets = []
    for idx, r in enumerate(results, start=1):
        h = r["header"]
        ops = r["ops"]
        name = h.get("name") or r["sheet"]
        code = str(h.get("code") or "").strip()
        group = h.get("group") or ""
        img = image_uri(r["image"], media_bytes, img_cache) if r.get("image") else None

        smv, mp, hc, prod = (num(h.get(k)) for k in ("smv", "mp", "hc", "prod"))
        sum_smv = sum(num(o.get("P")) or 0 for o in ops)

        rows = []
        for o in ops:
            rows.append("<tr>"
                        f'<td class="n">{fmt_sl(o.get("A"))}</td>'
                        f'<td class="l op-name">{esc(o.get("B"))}</td>'
                        f'<td>{esc(o.get("C")) or "&ndash;"}</td>'
                        f'<td class="n">{fmt_sl(o.get("D"))}</td>'
                        f'<td class="n">{fmt_sl(o.get("E"))}</td>'
                        f'<td class="n">{fmt_sl(o.get("F"))}</td>'
                        + "".join(f'<td class="n">{fmt_time(o.get(k)) or "&ndash;"}</td>' for k in "GHIJK")
                        + f'<td class="n">{fmt_dec(o.get("L"), 3)}</td>'
                        + f'<td class="n">{fmt_pct(o.get("M"))}</td>'
                        + f'<td class="n">{fmt_dec(o.get("N"), 4)}</td>'
                        + f'<td class="n">{fmt_pct(o.get("O"))}</td>'
                        + f'<td class="n smv-cell">{fmt_dec(o.get("P"), 5)}</td>'
                        + f'<td class="n">{fmt_hr(o.get("Q"))}</td>'
                        + "</tr>")
        rows.append('<tr class="total">'
                    '<td class="l" colspan="6">TOTAL (SMV per unit)</td>'
                    '<td class="n" colspan="5"></td>'
                    '<td class="n"></td><td class="n"></td><td class="n"></td><td class="n"></td>'
                    f'<td class="n smv-cell">{fmt_dec(sum_smv, 5)}</td><td class="n"></td></tr>')

        smv_eff = smv if smv is not None else (sum_smv if sum_smv else None)
        sum_mp_ops = sum(num(o.get("F")) or 0 for o in ops)
        mp_eff = mp if mp is not None else (sum_mp_ops if sum_mp_ops else None)

        if smv_eff:
            calc_prod = 60.0 / smv_eff
            calc_hc = (mp_eff * calc_prod) if mp_eff else None
        else:
            calc_prod = prod
            calc_hc = hc

        if mp_eff:
            mp_card = (f'<div class="kpi mp live"><div class="kpi-top">'
                       f'<span class="kpi-ic">{ICONS["mp"]}</span>'
                       f'<span class="kpi-lab">Manpower</span><span class="live-tag">edit</span></div>'
                       f'<div class="mp-ctrl">'
                       f'<button type="button" class="mp-btn" data-adj="-1" aria-label="minus">&minus;</button>'
                       f'<input type="number" class="mp-input" id="mp-{idx}" value="{fmt_sl(mp_eff)}" '
                       f'min="1" step="1" data-base="{fmt_sl(mp_eff)}" inputmode="numeric">'
                       f'<button type="button" class="mp-btn" data-adj="1" aria-label="plus">+</button>'
                       f'</div>'
                       f'<div class="kpi-unit">operators &middot; '
                       f'<button type="button" class="mp-reset">reset</button></div></div>')
        else:
            mp_card = (f'<div class="kpi mp"><div class="kpi-top">'
                       f'<span class="kpi-ic">{ICONS["mp"]}</span>'
                       f'<span class="kpi-lab">Manpower</span></div>'
                       f'<div class="kpi-val">&ndash;</div>'
                       f'<div class="kpi-unit">operators</div></div>')

        kpis = "".join([
            f'<div class="kpi smv"><div class="kpi-top"><span class="kpi-ic">{ICONS["smv"]}</span>'
            f'<span class="kpi-lab">Total SMV</span></div>'
            f'<div class="kpi-val">{fmt_dec(smv_eff, 5)}</div>'
            f'<div class="kpi-unit">min / pc</div></div>',
            mp_card,
            f'<div class="kpi hc live"><div class="kpi-top"><span class="kpi-ic">{ICONS["hc"]}</span>'
            f'<span class="kpi-lab">Hourly Capacity</span><span class="live-tag">auto</span></div>'
            f'<div class="kpi-val hc-val">{fmt_hr(calc_hc)}</div>'
            f'<div class="kpi-unit">pcs / hr</div></div>',
            f'<div class="kpi prod live"><div class="kpi-top"><span class="kpi-ic">{ICONS["prod"]}</span>'
            f'<span class="kpi-lab">Productivity</span><span class="live-tag">auto</span></div>'
            f'<div class="kpi-val prod-val">{fmt_dec(calc_prod, 2) if calc_prod else "&ndash;"}</div>'
            f'<div class="kpi-unit">pcs / op-hr</div></div>',
            f'<div class="kpi ops"><div class="kpi-top"><span class="kpi-ic">{ICONS["ops"]}</span>'
            f'<span class="kpi-lab">Operations</span></div>'
            f'<div class="kpi-val">{len(ops)}</div><div class="kpi-unit">steps</div></div>',
        ])

        img_html = (f'<div class="photo"><img src="{img}" alt=""><span class="photo-tag">PRODUCT</span></div>'
                    if img else '<div class="photo empty"><span class="ph-em">No product image</span></div>')

        alloc = r["alloc"]
        alloc_html = ""
        if alloc:
            items = "".join(f'<li><b>{esc(a["v"])}</b><span>{esc(a["k"])}</span></li>' for a in alloc)
            alloc_html = (f'<div class="alloc"><div class="block-h">'
                          f'<span class="block-title">Manpower Allocation</span></div><ul>{items}</ul></div>')

        bullets.append(
            f'<section class="bulletin" data-name="{esc(name.lower())}" data-code="{esc(code.lower())}" '
            f'data-group="{esc(group.lower())}" data-sheet="{esc(r["sheet"].lower())}" '
            f'data-smv="{smv_eff if smv_eff else 0}" data-mp="{mp_eff if mp_eff else 0}">'
            f'<div class="b-head">'
            f'<div class="b-brand">'
            f'<div class="b-eyebrow">ALEL INDUSTRIES LIMITED &middot; Industrial Engineering</div>'
            f'<h2 class="b-title">Operation Bulletin</h2>'
            f'<div class="b-sub">Work study &amp; capacity standard &mdash; Standard Minute Value (SMV)</div>'
            f'</div>'
            f'<div class="b-chips">'
            f'<span class="pill ob">OB-{idx:04d}</span>'
            f'<span class="pill date">{today}</span>'
            f'<span class="pill rev">Rev 00</span>'
            f'</div></div>'
            f'<div class="b-body">'
            f'{img_html}'
            f'<div class="b-info">'
            f'<div class="kv"><span class="kv-i">{ICONS["folder"]} Product Name / SKU</span><span class="kv-v">{esc(name)}</span></div>'
            f'<div class="kv"><span class="kv-i">{ICONS["code"]} Item Code</span><span class="kv-v code">{code or "&ndash;"}</span></div>'
            f'<div class="kv"><span class="kv-i">{ICONS["sheet"]} Product Group</span><span class="kv-v"><span class="g-tag">{esc(group) or "&ndash;"}</span></span></div>'
            f'<div class="kv"><span class="kv-i">{ICONS["sheet"]} Worksheet</span><span class="kv-v dim">{esc(r["sheet"])}</span></div>'
            f'<div class="kpis">{kpis}</div>'
            f'</div></div>'
            f'<div class="ops-wrap">'
            f'<div class="block-h"><span class="block-title">Operation Sequence &amp; Time Study</span>'
            f'<span class="ops-count">{len(ops)} operations</span></div>'
            f'<div class="tw"><table class="ops">'
            f'<thead><tr>'
            f'<th rowspan="2" class="r">SL</th><th rowspan="2" class="l">Operation / Process</th><th rowspan="2">Machine / Manual</th>'
            f'<th rowspan="2">Pcs/Unit</th><th rowspan="2">Pcs/Cycle</th><th rowspan="2">MP</th>'
            f'<th colspan="5">Observed Time (sec)</th>'
            f'<th rowspan="2">Cycle Time<br>(min/unit)</th><th rowspan="2">Rating</th>'
            f'<th rowspan="2">Basic Time</th><th rowspan="2">Allow</th><th rowspan="2">SMV<br>(min)</th>'
            f'<th rowspan="2">Hourly Cap.</th>'
            f'</tr><tr><th>T1</th><th>T2</th><th>T3</th><th>T4</th><th>T5</th></tr>'
            f'</thead><tbody>{"".join(rows)}</tbody></table></div>'
            f'</div>'
            f'{alloc_html}'
            f'<div class="signoff"><div><span>Prepared By</span></div><div><span>Checked By</span></div><div><span>Approved By</span><em>I.E. / Production</em></div></div>'
            f'</section>')

    groups = {}
    for r in results:
        g = (r["header"].get("group") or "").strip().lower()
        if g:
            groups[g] = groups.get(g, 0) + 1
    chip_html = '<button class="chip all on" data-g="all">All</button>' + "".join(
        f'<button class="chip" data-g="{esc(g)}">{g[0].upper() + g[1:]}<i>{groups[g]}</i></button>'
        for g in sorted(groups))

    html_doc = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>ALEL Operation Bulletin &mdash; Searchable (SKU / Item Code)</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&family=Sora:wght@600;700;800&display=swap" rel="stylesheet">
<style>
  :root{
    --bg:#0b1020;
    --bg2:#111a30;
    --card:#ffffff;
    --ink:#0f172a;
    --muted:#5b6b81;
    --navy:#0f2b46;
    --teal:#0e9f7a;
    --teal-d:#0b7a5f;
    --amber:#e8a13c;
    --line:#e6ebf2;
    --line2:#d6deea;
    --violet:#7c5cf0;
    --soft:#f4f7fb;
    --shadow:0 18px 50px -18px rgba(4,12,30,.35);
    --r:16px;
  }
  *{box-sizing:border-box;margin:0;padding:0;}
  html{scroll-behavior:smooth;}
  body{font-family:'Inter',system-ui,'Segoe UI',Arial,sans-serif;color:var(--ink);
       background:
         radial-gradient(1100px 520px at 85% -10%, rgba(124,92,240,.22), transparent 60%),
         radial-gradient(900px 480px at -10% 0%, rgba(14,159,122,.20), transparent 55%),
         linear-gradient(160deg,var(--bg),var(--bg2)); background-attachment:fixed;
       padding-bottom:80px;}
  .num{font-variant-numeric:tabular-nums;}
  ::selection{background:rgba(14,159,122,.25);}
  ::-webkit-scrollbar{width:11px;height:11px;}
  ::-webkit-scrollbar-thumb{background:#2a3a56;border-radius:8px;}
  ::-webkit-scrollbar-track{background:transparent;}

  /* ---------- top bar ---------- */
  .topbar{position:sticky;top:0;z-index:90;backdrop-filter:blur(14px);-webkit-backdrop-filter:blur(14px);
          background:rgba(8,13,27,.72);border-bottom:1px solid rgba(255,255,255,.08);}
  .topbar-inner{max-width:1240px;margin:0 auto;padding:12px 20px;display:flex;gap:14px;
                align-items:center;flex-wrap:wrap;}
  .brand-h{display:flex;align-items:center;gap:11px;color:#fff;}
  .logo{width:42px;height:42px;border-radius:12px;display:grid;place-items:center;flex:none;
        background:linear-gradient(135deg,#12b886,#0e9f7a 55%,#0b7a5f);
        box-shadow:0 8px 20px -6px rgba(14,159,122,.55);}
  .brand-h h1{font-family:'Sora',sans-serif;font-size:17px;line-height:1.1;}
  .brand-h h1 small{display:block;font-weight:500;font-size:10.5px;letter-spacing:.14em;color:#9fb3d0;
                    text-transform:uppercase;margin-top:3px;}
  .searchbox{flex:1 1 300px;display:flex;align-items:center;gap:10px;background:rgba(255,255,255,.08);
             border:1px solid rgba(255,255,255,.14);border-radius:14px;padding:9px 14px;color:#cdd9ea;
             transition:.2s;}
  .searchbox:focus-within{border-color:rgba(14,159,122,.7);background:rgba(255,255,255,.12);
                          box-shadow:0 0 0 4px rgba(14,159,122,.18);}
  .searchbox input{flex:1;background:transparent;border:0;outline:0;color:#fff;font-size:14.5px;min-width:0;font-family:inherit;}
  .searchbox input::placeholder{color:#7d8ea8;}
  .searchbox .clr{background:none;border:0;color:#9fb3d0;font-size:17px;cursor:pointer;line-height:1;display:none;}
  .kbd{font-size:10px;border:1px solid rgba(255,255,255,.25);border-radius:5px;padding:2px 6px;color:#9fb3d0;}
  .stat{color:#e6ecf5;font-size:12.5px;white-space:nowrap;}
  .tbtns{display:flex;gap:8px;align-items:center;}
  .tbtn{font-family:inherit;display:inline-flex;align-items:center;gap:7px;background:rgba(255,255,255,.1);
        color:#fff;border:1px solid rgba(255,255,255,.22);padding:8px 13px;border-radius:12px;
        font-size:12.5px;font-weight:700;cursor:pointer;transition:.18s;white-space:nowrap;}
  .tbtn:hover{background:rgba(255,255,255,.2);transform:translateY(-1px);}
  .tbtn svg{flex:none;}
  .tbtn.refresh{background:linear-gradient(135deg,#12b886,#0e9f7a);border-color:transparent;
                box-shadow:0 6px 16px -6px rgba(14,159,122,.6);}
  .tbtn.refresh:hover{filter:brightness(1.08);}
  .tbtn.on{background:linear-gradient(135deg,#ffb02e,#ff6b35);border-color:transparent;}
  .auto-dot{width:8px;height:8px;border-radius:50%;background:currentColor;opacity:.4;animation:pulse 1.6s infinite;}
  .tbtn.on .auto-dot{opacity:1;background:#fff;}
  @keyframes pulse{0%,100%{transform:scale(1);opacity:.5;}50%{transform:scale(1.5);opacity:1;}}
  .last-check{font-size:10.5px;color:#8fa0bb;display:flex;align-items:center;gap:6px;}
  .last-check.ok{color:#4adeb0;}
  .stat b{color:#4adeb0;font-size:16px;font-family:'Sora',sans-serif;}
  .refresh-note{width:100%;max-width:1240px;margin:-6px auto 0;padding:0 20px 10px;color:#8fa0bb;font-size:11px;}
  .refresh-note b{color:#c7d4e8;}
  .chips{max-width:1240px;margin:2px auto 0;padding:0 20px 12px;display:flex;gap:8px;flex-wrap:wrap;}
  .chip{font-family:inherit;display:inline-flex;align-items:center;gap:7px;background:rgba(255,255,255,.06);
        color:#cfd9ea;border:1px solid rgba(255,255,255,.13);padding:7px 13px;border-radius:999px;
        font-size:12.5px;font-weight:600;cursor:pointer;transition:.18s;}
  .chip i{font-style:normal;background:rgba(255,255,255,.14);border-radius:999px;padding:1px 8px;font-size:11px;}
  .chip:hover{background:rgba(255,255,255,.13);transform:translateY(-1px);}
  .chip.on{background:linear-gradient(135deg,#12b886,#0e9f7a);border-color:transparent;color:#fff;
           box-shadow:0 8px 18px -6px rgba(14,159,122,.6);}
  .chip.on i{background:rgba(0,0,0,.2);}

  main{max-width:1240px;margin:18px auto 0;padding:0 20px;}
  .toolbar-row{display:flex;justify-content:space-between;align-items:center;gap:12px;color:#9fb3d0;
               font-size:13px;margin-bottom:14px;flex-wrap:wrap;}
  .toolbar-row .lt{color:#e7eef9;font-weight:600;}
  .toolbar-row .lt span{color:#4adeb0;}

  /* ---------- bulletin card ---------- */
  .bulletin{background:var(--card);border-radius:var(--r);margin:0 0 26px;overflow:hidden;
            box-shadow:var(--shadow);border:1px solid rgba(255,255,255,.5);opacity:0;
            transform:translateY(14px);animation:rise .5s ease forwards;}
  .bulletin.hidden{display:none;}
  @keyframes rise{to{opacity:1;transform:none;}}
  .b-head{display:flex;justify-content:space-between;align-items:flex-start;gap:16px;flex-wrap:wrap;
          color:#fff;padding:20px 24px;position:relative;overflow:hidden;
          background:linear-gradient(120deg,#0b1c38 0%,#12314f 45%,#0e7a63 120%);}
  .b-head::after{content:"";position:absolute;inset:0;pointer-events:none;opacity:.5;
    background:radial-gradient(600px 200px at 92% 0%,rgba(255,255,255,.12),transparent 60%);}
  .b-eyebrow{font-size:10.5px;letter-spacing:.18em;text-transform:uppercase;color:#9fd8c6;}
  .b-title{font-family:'Sora',sans-serif;font-size:clamp(19px,2.4vw,26px);letter-spacing:.5px;margin:2px 0 3px;}
  .b-sub{font-size:11.5px;color:#c6d6ea;opacity:.9;}
  .b-chips{display:flex;gap:8px;flex-wrap:wrap;align-items:center;position:relative;z-index:2;}
  .pill{font-size:11px;font-weight:700;letter-spacing:.04em;padding:5px 11px;border-radius:999px;
        border:1px solid rgba(255,255,255,.35);background:rgba(255,255,255,.1);backdrop-filter:blur(4px);}
  .pill.ob{background:#ffd166;color:#5b3a00;border-color:transparent;}
  .pill.date,.pill.rev{color:#e9f0f9;}
  .b-body{display:flex;gap:20px;padding:20px 24px;flex-wrap:wrap;}
  .photo{flex:0 0 190px;width:190px;height:210px;border:1px solid var(--line);border-radius:14px;
         padding:10px;background:linear-gradient(180deg,#fff,var(--soft));position:relative;
         display:flex;align-items:center;justify-content:center;}
  .photo img{max-width:100%;max-height:100%;object-fit:contain;filter:drop-shadow(0 8px 16px rgba(20,40,70,.12));}
  .photo-tag{position:absolute;bottom:9px;left:50%;transform:translateX(-50%);font-size:8.5px;
             letter-spacing:.2em;color:#8fa0b8;background:#eef2f8;border-radius:99px;padding:2px 9px;}
  .photo.empty{border-style:dashed;color:#9fb0c4;}
  .b-info{flex:1 1 420px;min-width:300px;}
  .kv{display:grid;grid-template-columns:230px 1fr;gap:10px;padding:7px 2px;border-bottom:1px solid var(--soft);}
  .kv-i{display:inline-flex;align-items:center;gap:8px;color:var(--muted);font-size:12px;font-weight:600;}
  .kv-i svg{color:var(--teal);}
  .kv-v{font-size:14px;font-weight:700;word-break:break-word;}
  .kv-v.code{color:#1d5fd0;}
  .kv-v.dim{font-weight:500;color:var(--muted);font-size:12.5px;}
  .g-tag{display:inline-block;background:linear-gradient(135deg,#eafaf4,#d9f5ea);color:#0b7a5f;
         border-radius:8px;padding:3px 10px;font-size:12px;font-weight:700;border:1px solid #bfe9d9;}
  .kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(124px,1fr));gap:10px;margin-top:14px;}
  .kpi{background:linear-gradient(180deg,#fff,#fafcfe);border:1px solid var(--line);border-radius:13px;
       padding:11px 13px;transition:.18s;}
  .kpi:hover{transform:translateY(-2px);box-shadow:0 10px 22px -12px rgba(10,30,60,.35);}
  .kpi-top{display:flex;align-items:center;gap:8px;}
  .kpi-ic{width:30px;height:30px;border-radius:9px;display:grid;place-items:center;color:#fff;}
  .kpi-lab{font-size:10px;font-weight:800;text-transform:uppercase;letter-spacing:.06em;color:var(--muted);}
  .kpi-val{font-family:'Sora',sans-serif;font-size:19px;font-weight:800;margin-top:7px;letter-spacing:-.01em;}
  .kpi-unit{font-size:10px;color:#94a3b5;margin-top:1px;}
  .kpi.smv .kpi-ic{background:linear-gradient(135deg,#12b886,#0b7a5f);}
  .kpi.smv .kpi-val{color:#0b7a5f;}
  .kpi.mp .kpi-ic{background:linear-gradient(135deg,#5b8def,#3b62c7);}
  .kpi.mp .kpi-val{color:#3b62c7;}
  .kpi.hc .kpi-ic{background:linear-gradient(135deg,#f0a83c,#d97f17);}
  .kpi.hc .kpi-val{color:#c57412;}
  .kpi.prod .kpi-ic{background:linear-gradient(135deg,#8a6bf0,#6a45d9);}
  .kpi.prod .kpi-val{color:#6a45d9;}
  .kpi.ops .kpi-ic{background:linear-gradient(135deg,#e85f7e,#d1365a);}
  .kpi.ops .kpi-val{color:#d1365a;}
  .live-tag{font-size:8.5px;font-weight:800;text-transform:uppercase;letter-spacing:.06em;
            color:#fff;background:#ff6b35;border-radius:99px;padding:1px 7px;margin-left:auto;}
  .live-tag{background:linear-gradient(135deg,#ffb02e,#ff6b35);}
  .kpi.live{border:1px dashed #c9d6e4;background:linear-gradient(180deg,#fff,#fffdf8);}
  .mp-ctrl{display:flex;align-items:stretch;gap:6px;margin-top:8px;}
  .mp-btn{flex:0 0 30px;height:32px;border-radius:9px;border:1px solid var(--line2);background:#fff;
          color:var(--navy);font-size:18px;font-weight:700;cursor:pointer;line-height:1;
          transition:.15s;font-family:inherit;display:grid;place-items:center;}
  .mp-btn:hover{background:var(--navy);color:#fff;border-color:var(--navy);}
  .mp-input{flex:1;min-width:0;height:32px;border:1px solid var(--line2);border-radius:9px;text-align:center;
            font-family:'Sora',sans-serif;font-size:16px;font-weight:800;color:#3b62c7;outline:none;
            background:#f8fbff;-moz-appearance:textfield;appearance:textfield;}
  .mp-input::-webkit-outer-spin-button,.mp-input::-webkit-inner-spin-button{-webkit-appearance:none;margin:0;}
  .mp-input:focus{border-color:#3b62c7;box-shadow:0 0 0 3px rgba(59,98,199,.18);}
  .mp-reset{font-family:inherit;font-size:9px;color:#7c8ba0;background:none;border:none;cursor:pointer;
            text-decoration:underline dotted;margin-left:2px;}
  .mp-reset:hover{color:#3b62c7;}
  .kpi.mp .kpi-unit{line-height:1.6;}
  .ops-wrap{padding:0 24px 6px;}
  .alloc{margin:0 24px 4px;}
  .block-h{display:flex;justify-content:space-between;align-items:center;gap:10px;margin:10px 0 10px;flex-wrap:wrap;}
  .block-title{font-size:12px;font-weight:800;text-transform:uppercase;letter-spacing:.1em;color:var(--navy);
               display:flex;align-items:center;gap:8px;}
  .block-title::before{content:"";width:4px;height:15px;border-radius:3px;
                       background:linear-gradient(180deg,#12b886,#0b7a5f);}
  .ops-count{font-size:11px;font-weight:700;color:var(--teal-d);background:#e7f8f2;border:1px solid #c5ecdd;
             padding:3px 10px;border-radius:999px;}
  .tw{overflow-x:auto;border:1px solid var(--line);border-radius:12px;}
  table.ops{border-collapse:collapse;width:100%;font-size:11.8px;background:#fff;}
  table.ops th,table.ops td{border-bottom:1px solid var(--line);border-right:1px solid var(--line);
        padding:7px 8px;text-align:center;white-space:nowrap;}
  table.ops th:last-child,table.ops td:last-child{border-right:0;}
  table.ops thead th{background:var(--navy);color:#dbe7f7;font-weight:600;font-size:9.8px;
        text-transform:uppercase;letter-spacing:.03em;line-height:1.3;position:relative;}
  table.ops thead tr:first-child th{color:#fff;}
  table.ops thead tr:first-child th.l{text-align:left;}
  table.ops tbody td.l{text-align:left;}
  table.ops tbody tr:hover{background:#f0f7ff;}
  table.ops tbody td{color:#33445c;}
  table.ops td.op-name{min-width:180px;white-space:normal;font-weight:500;}
  table.ops td.smv-cell{color:var(--teal-d);font-weight:800;background:#f0faf6;}
  tr.total td{background:linear-gradient(180deg,#0f2b46,#163a5c);color:#fff;font-weight:700;}
  tr.total td.smv-cell{color:#ffd166;background:transparent;font-size:13px;}
  .alloc ul{display:grid;grid-template-columns:repeat(auto-fill,minmax(220px,1fr));gap:8px 16px;
            list-style:none;padding:2px 0 12px;}
  .alloc li{display:flex;gap:9px;align-items:center;background:var(--soft);border:1px solid var(--line);
            border-radius:10px;padding:6px 11px;font-size:12px;}
  .alloc li b{min-width:30px;text-align:center;background:#fff;border:1px solid var(--line2);color:#0b7a5f;
              border-radius:8px;padding:2px 7px;font-size:12px;}
  .alloc li span{font-weight:500;color:#33445c;}
  .signoff{display:grid;grid-template-columns:repeat(3,1fr);gap:20px;padding:26px 24px 24px;margin-top:6px;}
  .signoff div{text-align:center;border-top:2px dotted #b6c2d2;padding-top:9px;position:relative;}
  .signoff span{font-size:11.5px;font-weight:700;color:#34465e;text-transform:uppercase;letter-spacing:.05em;}
  .signoff em{display:block;font-style:normal;font-size:9.5px;color:#94a3b5;margin-top:3px;}
  .msg{display:none;text-align:center;color:#cbd7e8;padding:70px 20px;font-size:16px;}
  .msg b{color:#fff;}
  .footer-note{max-width:1240px;margin:8px auto 0;padding:0 20px;color:#6f83a3;font-size:11.5px;text-align:center;}

  /* ============================================================
     OPERATION CARDS (mobile-friendly alt to the wide table)
     ============================================================ */
  .opcards{display:none;}
  @media (max-width:820px){
    .opcards{display:block;}
    .tw{display:none;}
    .oc-card{background:#fff;border:1px solid var(--line);border-radius:12px;margin:0 0 9px;
             overflow:hidden;}
    .oc-top{display:flex;align-items:center;gap:10px;padding:9px 12px;
            background:linear-gradient(180deg,#fff,var(--soft));cursor:pointer;}
    .oc-no{flex:none;width:26px;height:26px;border-radius:8px;display:grid;place-items:center;
           background:var(--navy);color:#fff;font-weight:800;font-size:12.5px;}
    .oc-title{flex:1;font-size:13px;font-weight:700;line-height:1.35;color:var(--ink);}
    .oc-smv{flex:none;text-align:right;}
    .oc-smv b{display:block;font-family:'Sora',sans-serif;font-size:15px;color:var(--teal-d);}
    .oc-smv span{font-size:8.5px;color:#8fa0b8;text-transform:uppercase;letter-spacing:.05em;}
    .oc-chips{display:flex;flex-wrap:wrap;gap:6px;padding:0 12px 9px;}
    .oc-chip{font-size:10.5px;font-weight:600;color:#42536a;background:#eef2f8;border:1px solid var(--line);
             border-radius:99px;padding:3px 9px;}
    .oc-chip b{color:var(--navy);}
    .oc-body{display:none;padding:2px 12px 11px;border-top:1px dashed var(--line);}
    .oc-card.open .oc-body{display:block;}
    .oc-body .ocg{display:grid;grid-template-columns:1fr 1fr 1fr;gap:6px;margin-top:9px;}
    .oc-item{background:var(--soft);border:1px solid var(--line);border-radius:9px;padding:5px 8px;}
    .oc-item span{display:block;font-size:8.5px;color:#8fa0b8;text-transform:uppercase;letter-spacing:.04em;}
    .oc-item b{font-size:12.5px;color:#24364c;}
    .oc-item.wide{grid-column:1/-1;}
    .oc-caret{margin-left:auto;color:#7c8ba0;flex:none;transition:.2s;}
    .oc-card.open .oc-caret{transform:rotate(180deg);}
  }

  /* ---------- mobile ---------- */
  @media (max-width:820px){
    .topbar-inner{padding:10px 12px;gap:10px;}
    .logo{width:38px;height:38px;border-radius:11px;}
    .brand-h h1{font-size:14.5px;}
    .brand-h h1 small{font-size:9.5px;}
    .searchbox{flex:1 1 100%;order:5;padding:10px 13px;border-radius:12px;}
    .searchbox input{font-size:16px;padding:4px 0;}   /* 16px prevents iOS zoom */
    .kbd{display:none;}
    .stat{font-size:11.5px;}
    .stat b{font-size:14px;}
    .tbtns{gap:7px;}
    .tbtn{padding:9px 12px;font-size:12px;border-radius:11px;}
    .tbtn svg{width:17px;height:17px;}
    .refresh-note{font-size:10px;padding:0 14px 9px;line-height:1.5;}
    .chips{padding:0 12px 10px;gap:6px;}
    .chip{padding:8px 12px;font-size:12px;}
    main{padding:0 10px;margin:12px auto 0;}
    .toolbar-row{display:none;}
    .bulletin{margin:0 0 14px;border-radius:14px;}
    .b-head{padding:14px 14px;gap:10px;}
    .b-eyebrow{font-size:9px;}
    .b-title{font-size:19px;}
    .b-sub{font-size:10.5px;line-height:1.4;}
    .b-chips{gap:6px;}
    .pill{font-size:10px;padding:4px 9px;}
    .b-body{padding:12px 14px;gap:12px;}
    .photo{flex:0 0 92px;width:92px;height:104px;border-radius:11px;padding:6px;}
    .photo-tag{display:none;}
    .b-info{flex:1 1 200px;min-width:0;}
    .kv{grid-template-columns:1fr;gap:1px;padding:5px 0;}
    .kv-i{font-size:10.5px;}
    .kv-i svg{width:12px;height:12px;}
    .kv-v{font-size:13.5px;line-height:1.35;}
    .g-tag{font-size:11px;padding:2px 8px;}
    .kpis{grid-template-columns:repeat(2,1fr);gap:8px;margin-top:11px;}
    .kpi{padding:9px 10px;border-radius:12px;}
    .kpi-ic{width:27px;height:27px;border-radius:8px;}
    .kpi-lab{font-size:9px;}
    .kpi-val{font-size:17px;margin-top:6px;}
    .mp-ctrl{margin-top:7px;gap:5px;}
    .mp-btn{flex:0 0 38px;height:38px;font-size:20px;border-radius:10px;}
    .mp-input{height:38px;font-size:17px;}
    .kpi.ops .kpi-val{font-size:15px;}
    .ops-wrap{padding:0 14px 6px;}
    .alloc{margin:0 14px 4px;}
    .block-h{margin:8px 0;}
    .block-title{font-size:11px;}
    .ops-count{font-size:10px;}
    .alloc ul{grid-template-columns:repeat(auto-fill,minmax(140px,1fr));gap:6px;}
    .alloc li{font-size:11px;padding:5px 9px;}
    .signoff{padding:18px 14px 18px;gap:16px;}
    .signoff div{font-size:10px;}
    .ops-wrap .tw{display:none;}          /* cards already shown */
    .msg{padding:40px 14px;}
  }
  @media (max-width:360px){
    .b-title{font-size:17px;}
    .kpis{grid-template-columns:1fr 1fr;}
  }
  /* print */
  @media print{
    body{background:#fff;padding:0;}
    .topbar,.toolbar-row,.refresh-note,.footer-note{display:none !important;}
    main{max-width:100%;padding:0;}
    .bulletin{box-shadow:none;margin:0 0 10px;border-radius:0;animation:none;opacity:1;transform:none;
              break-after:page;}
    .b-head{border-radius:0;}
    .tw{display:block !important;overflow:visible;border:0;}
    .opcards{display:none !important;}
    table.ops th{position:static;}
    .kpi:hover{transform:none;box-shadow:none;}
    @page{size:A4 landscape;margin:8mm;}
  }
</style>
</head>
<body>
<div class="topbar">
  <div class="topbar-inner">
    <div class="brand-h">
      <div class="logo">&#9889;</div>
      <h1>ALEL Operation Bulletin<small>ENOVAR &middot; Work Study / SMV &middot; Searchable</small></h1>
    </div>
    <div class="searchbox">
      %ICON_SEARCH%
      <input id="q" type="text" placeholder="Search SKU name or item code… e.g. MCB, 14414, BULB" autocomplete="off">
      <button class="clr" id="clr" title="Clear">&times;</button>
      <span class="kbd">/</span>
    </div>
    <div class="stat"><b id="cnt">0</b> / %TOTAL% bulletins</div>
    <div class="tbtns">
      <button class="tbtn auto" id="autoBtn" type="button"><span class="auto-dot"></span><span id="autoLab">Auto 1m</span></button>
      <button class="tbtn refresh" id="refreshBtn" type="button">%ICON_REFRESH%<span>Refresh Data</span></button>
    </div>
  </div>
  <div class="chips" id="chips">%CHIPS%</div>
  <div class="refresh-note">Data refreshed from Google Sheets on <b>%REFRESHED%</b> &middot; Click category chip or type to filter &middot; Manpower &#8722;/+ edits update Hourly Capacity &amp; Productivity live &middot; <b>Ctrl+P</b> prints each bulletin cleanly</div>
</div>
<main>
  <div class="toolbar-row"><div class="lt">Live results&nbsp;&nbsp;<span id="cnt2"></span></div><div id="hint" style="color:#9fb3d0;"></div></div>
  %BULLETINS%
  <div class="msg" id="msg"><b>No bulletins found.</b><br>Try another SKU, item code or category.</div>
</main>
<div class="footer-note">ALEL Industries Limited &middot; Industrial Engineering &middot; Operation Bulletin v2.0 &middot; For internal use</div>
<script>
(function(){
  var bullets = Array.prototype.slice.call(document.querySelectorAll('.bulletin'));
  var activeGroup = '';
  var qEl=document.getElementById('q'), cntEl=document.getElementById('cnt'),
      cnt2=document.getElementById('cnt2'), msg=document.getElementById('msg'),
      chipsEl=document.getElementById('chips'), clr=document.getElementById('clr');
  bullets.forEach(function(b,i){ b.style.animationDelay=(Math.min(i,30)*14)+'ms'; });

  function apply(){
    var t=(qEl.value||'').trim().toLowerCase();
    var n=0;
    bullets.forEach(function(b){
      var show=true;
      if(activeGroup && activeGroup!=='all' && b.getAttribute('data-group')!==activeGroup) show=false;
      if(show && t){
        var nm=b.getAttribute('data-name')||'';
        var cd=b.getAttribute('data-code')||'';
        var sh=b.getAttribute('data-sheet')||'';
        show = nm.indexOf(t)>-1 || cd.indexOf(t)>-1 || sh.indexOf(t)>-1;
      }
      b.classList.toggle('hidden',!show);
      if(show){n++; b.style.animationDelay='0ms'; b.style.animation='none'; void b.offsetWidth;
               b.style.animation='rise .45s ease forwards';}
    });
    cntEl.textContent=n; if(cnt2) cnt2.textContent=n+' shown';
    msg.style.display=n?'none':'block';
    clr.style.display=qEl.value?'block':'none';
    var h=document.getElementById('hint');
    if(t){var hs=document.querySelectorAll('.bulletin:not(.hidden) .b-title');
          h.textContent=hs.length? 'Showing '+hs.length+' result'+(hs.length>1?'s':''):'';}
    else h.textContent='';
  }
  chipsEl.addEventListener('click',function(e){
    var b=e.target.closest('.chip'); if(!b)return;
    chipsEl.querySelectorAll('.chip').forEach(function(c){c.classList.remove('on');});
    b.classList.add('on');
    activeGroup=b.getAttribute('data-g');
    apply();
  });
  qEl.addEventListener('input',apply);
  clr.addEventListener('click',function(){qEl.value='';apply();qEl.focus();});
  document.addEventListener('keydown',function(e){
    if(e.key==='/' && document.activeElement!==qEl){e.preventDefault();qEl.focus();}
    if(e.key==='Escape' && document.activeElement===qEl){qEl.value='';apply();qEl.blur();}
  });
  apply();

  /* ---------- live manpower -> Hourly Capacity & Productivity ---------- */
  function fmtNum(x,d){
    if(x===null||x===undefined||isNaN(x)) return '&ndash;';
    return x.toLocaleString('en-US',{minimumFractionDigits:d,maximumFractionDigits:d});
  }
  function recalc(card){
    var sec = card.closest('.bulletin');
    var smv = parseFloat(sec.getAttribute('data-smv'));
    var inp = card.querySelector('.mp-input');
    var mp = parseFloat(inp.value);
    if(isNaN(mp) || mp<1) mp=1;
    inp.value = mp;
    if(!smv || smv<=0){          // no SMV -> cannot compute, restore base
      inp.value = inp.getAttribute('data-base');
      return;
    }
    sec.setAttribute('data-mp', mp);
    var prod = 60 / smv;                 // pcs / operator-hour
    var hc   = mp * prod;                // pcs / line-hour
    sec.querySelector('.hc-val').innerHTML = fmtNum(hc,0);
    sec.querySelector('.prod-val').innerHTML = fmtNum(prod,2);
  }
  function bindMP(card){
    var inp = card.querySelector('.mp-input');
    var btns = card.querySelectorAll('.mp-btn');
    inp.addEventListener('change',function(){ recalc(card); });
    inp.addEventListener('keyup',function(){ if(inp.value) recalc(card); });
    btns.forEach(function(btn){
      btn.addEventListener('click',function(){
        var v = parseFloat(inp.value);
        if(isNaN(v)) v=1;
        v += parseInt(btn.getAttribute('data-adj'),10);
        if(v<1) v=1;
        inp.value = v;
        recalc(card);
      });
    });
    var rst = card.querySelector('.mp-reset');
    if(rst){ rst.addEventListener('click',function(){
        inp.value = inp.getAttribute('data-base');
        recalc(card);
    }); }
  }
  document.querySelectorAll('.kpi.mp .mp-input').forEach(function(i){ bindMP(i.closest('.kpi')); });

  /* ---------- mobile operation cards (built from the table) ---------- */
  function buildOpCards(){
    document.querySelectorAll('.bulletin .ops-wrap').forEach(function(wrap){
      if(wrap.querySelector('.opcards')) return;          // only once
      var box = document.createElement('div');
      box.className = 'opcards';
      var table = wrap.querySelector('table.ops');
      if(!table) return;
      table.querySelectorAll('tbody tr:not(.total)').forEach(function(tr){
        var td = tr.querySelectorAll('td');
        function txt(i){ return td[i] ? td[i].textContent.trim() : ''; }
        var n   = txt(0);
        var op  = txt(1);
        var mac = txt(2);
        var smv = txt(15);
        var mp  = txt(5);
        var ct  = txt(11);
        var rt  = txt(12);
        var t1=txt(6),t2=txt(7),t3=txt(8),t4=txt(9),t5=txt(10);
        var pu=txt(3), pc=txt(4), bt=txt(13), al=txt(14), hc=txt(16);
        var card = document.createElement('div');
        card.className = 'oc-card';
        card.innerHTML =
          '<div class="oc-top"><span class="oc-no">'+n+'</span>'+
          '<div class="oc-title">'+op+'</div>'+
          '<div class="oc-smv"><b>'+smv+'</b><span>SMV min</span></div>'+
          '<svg class="oc-caret" viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round"><polyline points="6 9 12 15 18 9"/></svg></div>'+
          '<div class="oc-chips">'+
            (mac && mac!=='&ndash;' ? '<span class="oc-chip"><b>'+mac+'</b></span>' : '')+
            (mp  && mp !=='&ndash;' ? '<span class="oc-chip">MP <b>'+mp+'</b></span>' : '')+
            (ct  && ct !=='&ndash;' ? '<span class="oc-chip">Cycle <b>'+ct+' min</b></span>' : '')+
            (rt  && rt !=='&ndash;' ? '<span class="oc-chip">Rating <b>'+rt+'</b></span>' : '')+
          '</div>'+
          '<div class="oc-body"><div class="ocg">'+
            '<div class="oc-item"><span>Pcs/Unit</span><b>'+pu+'</b></div>'+
            '<div class="oc-item"><span>Pcs/Cycle</span><b>'+pc+'</b></div>'+
            '<div class="oc-item"><span>Hourly Cap.</span><b>'+hc+'</b></div>'+
            '<div class="oc-item"><span>T1 (sec)</span><b>'+t1+'</b></div>'+
            '<div class="oc-item"><span>T2 (sec)</span><b>'+t2+'</b></div>'+
            '<div class="oc-item"><span>T3 (sec)</span><b>'+t3+'</b></div>'+
            '<div class="oc-item"><span>T4 (sec)</span><b>'+t4+'</b></div>'+
            '<div class="oc-item"><span>T5 (sec)</span><b>'+t5+'</b></div>'+
            '<div class="oc-item"><span>Basic Time</span><b>'+bt+'</b></div>'+
            '<div class="oc-item"><span>Allowance</span><b>'+al+'</b></div>'+
          '</div></div>';
        card.querySelector('.oc-top').addEventListener('click', function(){
          card.classList.toggle('open');
        });
        box.appendChild(card);
      });
      // append after the scrollable table wrapper, inside the ops-wrap
      var tw = wrap.querySelector('.tw');
      if(tw && tw.nextSibling) wrap.insertBefore(box, tw.nextSibling);
      else wrap.appendChild(box);
    });
  }
  buildOpCards();

  /* ---------- refresh / auto-reload ---------- */
  var autoTimer = null;
  var autoBtn = document.getElementById('autoBtn');
  var autoLab = document.getElementById('autoLab');
  function setAuto(on){
    if(autoTimer){ clearInterval(autoTimer); autoTimer=null; }
    autoBtn.classList.toggle('on', on);
    autoLab.textContent = on ? 'Auto 1m ON' : 'Auto 1m';
    if(on){ autoTimer = setInterval(function(){ location.reload(); }, 60000); }
    try{ localStorage.setItem('alel_auto', on?'1':'0'); }catch(e){}
  }
  autoBtn.addEventListener('click', function(){
    setAuto(!autoBtn.classList.contains('on'));
  });
  document.getElementById('refreshBtn').addEventListener('click', function(){
    location.reload();
  });
  var savedAuto = null;
  try{ savedAuto = localStorage.getItem('alel_auto'); }catch(e){}
  if(savedAuto === '1') setAuto(true);
  document.addEventListener('visibilitychange', function(){
    if(!document.hidden) setAuto(autoBtn.classList.contains('on'));
  });

})();
</script>
</body>
</html>"""

    html_doc = (html_doc.replace("%BULLETINS%", "".join(bullets))
                        .replace("%CHIPS%", chip_html)
                        .replace("%ICON_SEARCH%", ICONS["search"])
                        .replace("%ICON_REFRESH%", ICONS["refresh"])
                        .replace("%TOTAL%", str(len(results)))
                        .replace("%REFRESHED%", today))
    dest = out_path or OUT_HTML
    with open(dest, "w", encoding="utf-8") as f:
        f.write(html_doc)
    log("      written: " + dest)
    log("      size MB: " + str(round(os.path.getsize(dest) / 1048576, 2)))


# --------------------------------------------------------------------------
def main():
    import argparse
    ap = argparse.ArgumentParser(description="ALEL Operation Bulletin generator")
    ap.add_argument("--output", default=None, help="output html path (default: ALEL_Operation_Bulletin_Searchable.html in this folder)")
    ap.add_argument("--sheet-id", default=None, help="override google sheet id")
    args = ap.parse_args()

    global SHEET_ID, XLSX_TMP
    if args.sheet_id:
        SHEET_ID = args.sheet_id
        XLSX_TMP = os.path.join(HERE, "_latest_capacity_study.xlsx")

    try:
        download()
        results, media = parse_all()
        if len(results) < 50:
            raise SystemExit("Expected 100+ bulletins, got %d - workbook looks wrong." % len(results))
        build_html(results, media, out_path=args.output)
        log("[4/4] DONE. Open the HTML and search by SKU name or item code.")
    except SystemExit as e:
        log(str(e))
        sys.exit(1)
    except Exception as e:
        log("ERROR: " + str(e))
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()
