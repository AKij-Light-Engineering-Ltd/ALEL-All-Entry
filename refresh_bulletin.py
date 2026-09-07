# -*- coding: utf-8 -*-
"""
ALEL Operation Bulletin - One-click refresh tool  (MODERN UI v3)
Downloads the Google Sheet workbook, extracts every product operation study
+ product image + Section (PG LIST col E), and regenerates a modern
searchable HTML with an approval workflow (Prepared -> Checked -> Approved).

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
import hashlib
import urllib.request
import datetime

SHEET_ID = "1GpALH7TsJ0eW-Ko8ejzV8oQF073dhe7Eb9930o0qvSo"
HERE = os.path.dirname(os.path.abspath(__file__))
XLSX_TMP = os.path.join(HERE, "_latest_capacity_study.xlsx")
OUT_HTML = os.path.join(HERE, "ALEL_Operation_Bulletin_Searchable.html")

MAX_SIDE = 520  # px for embedded product images

# ---------------- Approval workflow configuration ----------------
PREPARED_BY = {"name": "Anoy Kumar Das", "role": "Operational Excellence / I.E."}
APPROVED_BY = {"name": "Md. Moshfequr Rahman", "role": "Plant Head", "email": "head.plant@akijlightengineering.com"}
CHECKER_BY_SECTION = {
    "GSS": {"name": "Jhumour Rani", "email": "jhumour@akijlightengineering.com"},
    "LED": {"name": "Kajal Kanti", "email": "kajal04@akijlightengineering.com"},
    "HAP": {"name": "Abdullah Al-Mamun", "email": "almamun@akijlightengineering.com"},
}
STORE_KEY = "alel_approvals_v1"

# ---------------- Login users (login by EMAIL) ----------------
# role: admin | prepared | checker | approver
# section (for checker): GSS | LED | HAP
USERS = [
    {"id": "md.marufhossain@akijlightengineering.com", "name": "Md. Maruf Hossain",        "role": "admin",     "section": "", "pass": "admin123"},
    {"id": "anoy@akijlightengineering.com",             "name": "Anoy Kumar Das",          "role": "prepared",  "section": "", "pass": "anoy123"},
    {"id": "jhumour@akijlightengineering.com",          "name": "Jhumour Rani",            "role": "checker",   "section": "GSS", "pass": "gss123"},
    {"id": "kajal04@akijlightengineering.com",          "name": "Kajal Kanti",             "role": "checker",   "section": "LED", "pass": "led123"},
    {"id": "almamun@akijlightengineering.com",          "name": "Abdullah Al-Mamun",       "role": "checker",   "section": "HAP", "pass": "hap123"},
    {"id": "head.plant@akijlightengineering.com",       "name": "Md. Moshfequr Rahman",    "role": "approver",  "section": "", "pass": "head123"},
]
USER_SESSION_KEY = "alel_login_v1"
# Prepared signers (displayed list)
PREPARED_NAMES = ["Md. Maruf Hossain", "Anoy Kumar Das"]
ORG_LABEL = "Operational Excellence"   # replaces "Industrial Engineering"


def login_user_db():
    db = []
    for u in USERS:
        h = hashlib.sha256(u["pass"].encode("utf-8")).hexdigest()
        db.append({"id": u["id"], "name": u["name"], "role": u["role"],
                   "section": u["section"], "hash": h})
    return db


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

    # ---- Section map from PG LIST: col A=Code, col E=Section ----
    code2sec = {}
    try:
        pgl = wb["PG LIST"]
        for row in pgl.iter_rows(min_row=2):
            if len(row) < 5:
                continue
            code = row[0].value
            sec = row[4].value
            if code is None or sec is None:
                continue
            c = str(code).strip()
            if c.endswith(".0"):
                c = c[:-2]
            s = str(sec).strip()
            if c and s:
                code2sec[c] = s
    except Exception as e:
        log("      (PG LIST section read skipped: %s)" % e)

    results = []
    for nm in wb.sheetnames:
        m = parse_sheet(wb[nm], sno2img)
        if m["ops"] and (m["header"].get("name") or m["header"].get("code")):
            hc = str(m["header"].get("code") or "").strip()
            if hc.endswith(".0"):
                hc = hc[:-2]
            m["section"] = code2sec.get(hc, "")
            results.append(m)
    wb.close()
    log("      %d product bulletins found" % len(results))
    return results, media_bytes


# --------------------------------------------------------------------------
# 3) HTML generation helpers
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


IC = {
    "smv": '<svg viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="9"/><path d="M12 7v5l3 3"/></svg>',
    "mp": '<svg viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M17 21v-2a4 4 0 0 0-4-4H7a4 4 0 0 0-4 4v2"/><circle cx="10" cy="7" r="4"/><path d="M23 21v-2a4 4 0 0 0-3-3.87"/><path d="M16 3.13a4 4 0 0 1 0 7.75"/></svg>',
    "hc": '<svg viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M13 2L3 14h7l-1 8 10-12h-7l1-8z"/></svg>',
    "prod": '<svg viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="9"/><circle cx="12" cy="12" r="5"/><circle cx="12" cy="12" r="1"/></svg>',
    "ops": '<svg viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M8 6h13M8 12h13M8 18h13"/><circle cx="3.5" cy="6" r="1.4" fill="currentColor" stroke="none"/><circle cx="3.5" cy="12" r="1.4" fill="currentColor" stroke="none"/><circle cx="3.5" cy="18" r="1.4" fill="currentColor" stroke="none"/></svg>',
    "search": '<svg viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>',
    "refresh": '<svg viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round"><path d="M21 12a9 9 0 1 1-2.64-6.36"/><polyline points="21 3 21 9 15 9"/></svg>',
    "ck": '<svg viewBox="0 0 24 24" width="17" height="17" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M9 11l3 3L22 4"/><path d="M21 12v7a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11"/></svg>',
    "ap": '<svg viewBox="0 0 24 24" width="17" height="17" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"/><polyline points="22 4 12 14.01 9 11.01"/></svg>',
    "pp": '<svg viewBox="0 0 24 24" width="17" height="17" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/></svg>',
    "folder": '<svg viewBox="0 0 24 24" width="13" height="13" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M22 19a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h5l2 3h9a2 2 0 0 1 2 2z"/></svg>',
    "code": '<svg viewBox="0 0 24 24" width="13" height="13" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><polyline points="16 18 22 12 16 6"/><polyline points="8 6 2 12 8 18"/></svg>',
}


def build_html(results, media_bytes, out_path=None):
    log("[3/4] Generating modern HTML ...")
    img_cache = {}
    today = datetime.date.today().strftime("%d-%b-%Y")
    bullets = []
    catalog = []
    for idx, r in enumerate(results, start=1):
        h = r["header"]
        ops = r["ops"]
        name = h.get("name") or r["sheet"]
        code = str(h.get("code") or "").strip()
        if code.endswith(".0"):
            code = code[:-2]
        group = h.get("group") or ""
        section = (r.get("section") or "").upper()
        img = image_uri(r["image"], media_bytes, img_cache) if r.get("image") else None

        smv, mp, hc, prod = (num(h.get(k)) for k in ("smv", "mp", "hc", "prod"))
        sum_smv = sum(num(o.get("P")) or 0 for o in ops)
        smv_eff = smv if smv is not None else (sum_smv if sum_smv else None)
        sum_mp_ops = sum(num(o.get("F")) or 0 for o in ops)
        mp_eff = mp if mp is not None else (sum_mp_ops if sum_mp_ops else None)
        calc_prod = (60.0 / smv_eff) if smv_eff else prod
        calc_hc = (mp_eff * calc_prod) if (mp_eff and smv_eff) else hc

        cat = {
            "c": code, "n": name, "s": section, "g": group.lower(), "sh": r["sheet"].lower(),
            "p": PREPARED_BY["name"], "ck": CHECKER_BY_SECTION.get(section, {}).get("name", ""),
            "ap": APPROVED_BY["name"],
        }
        catalog.append(cat)

        # ---------- KPI cards ----------
        mp_ctrl = ""
        if mp_eff:
            mp_ctrl = (
                '<div class="mp-ctrl">'
                '<button type="button" class="mp-btn" data-adj="-1" aria-label="minus">&minus;</button>'
                f'<input type="number" class="mp-input" id="mp-{idx}" value="{fmt_sl(mp_eff)}" min="1" '
                f'step="1" data-base="{fmt_sl(mp_eff)}" inputmode="numeric">'
                '<button type="button" class="mp-btn" data-adj="1" aria-label="plus">+</button>'
                '</div>'
                '<button type="button" class="mp-reset">reset</button>')
        else:
            mp_ctrl = '<div class="kpi-na">&ndash;</div>'

        kpi_smv = fmt_dec(smv_eff, 5) if smv_eff is not None else "&ndash;"
        kpi_hc = fmt_hr(calc_hc)
        kpi_prod = fmt_dec(calc_prod, 2) if calc_prod else "&ndash;"

        kpis = (
            f'<div class="kpi k-smv"><div class="k-top">{IC["smv"]}<span>Total SMV</span></div>'
            f'<div class="k-val">{kpi_smv}</div><div class="k-unit">min / pc</div></div>'
            f'<div class="kpi k-mp"><div class="k-top">{IC["mp"]}<span>Manpower</span>'
            f'<em class="live">edit</em></div>{mp_ctrl}'
            f'<div class="k-unit">operators</div></div>'
            f'<div class="kpi k-hc"><div class="k-top">{IC["hc"]}<span>Hourly Capacity</span>'
            f'<em class="live">auto</em></div><div class="k-val hc-val">{kpi_hc}</div>'
            f'<div class="k-unit">pcs / hr</div></div>'
            f'<div class="kpi k-prod"><div class="k-top">{IC["prod"]}<span>Productivity</span>'
            f'<em class="live">auto</em></div><div class="k-val prod-val">{kpi_prod}</div>'
            f'<div class="k-unit">pcs / op-hr</div></div>'
            f'<div class="kpi k-ops"><div class="k-top">{IC["ops"]}<span>Operations</span></div>'
            f'<div class="k-val">{len(ops)}</div><div class="k-unit">steps</div></div>'
        )

        photo = (f'<div class="photo"><img src="{img}" alt=""><span class="ptag">PRODUCT</span></div>'
                 if img else '<div class="photo empty"><span>No product image</span></div>')

        # ---------- operation rows (kept for print/table) ----------
        tro = []
        for o in ops:
            tro.append(
                "<tr>"
                f'<td class="c">{fmt_sl(o.get("A"))}</td>'
                f'<td class="l">{esc(o.get("B"))}</td>'
                f'<td>{esc(o.get("C")) or "&ndash;"}</td>'
                f'<td class="c">{fmt_sl(o.get("D"))}</td>'
                f'<td class="c">{fmt_sl(o.get("E"))}</td>'
                f'<td class="c">{fmt_sl(o.get("F"))}</td>'
                + "".join(f'<td class="c">{fmt_time(o.get(k)) or "&ndash;"}</td>' for k in "GHIJK")
                + f'<td class="c">{fmt_dec(o.get("L"), 3)}</td>'
                + f'<td class="c">{fmt_pct(o.get("M"))}</td>'
                + f'<td class="c">{fmt_dec(o.get("N"), 4)}</td>'
                + f'<td class="c">{fmt_pct(o.get("O"))}</td>'
                + f'<td class="c smv">{fmt_dec(o.get("P"), 5)}</td>'
                + f'<td class="c">{fmt_hr(o.get("Q"))}</td>'
                + "</tr>")
        tro.append('<tr class="total"><td class="l" colspan="6">TOTAL (SMV per unit)</td>'
                   '<td class="c" colspan="5"></td><td class="c"></td><td class="c"></td>'
                   '<td class="c"></td><td class="c"></td>'
                   f'<td class="c smv">{fmt_dec(sum_smv, 5)}</td><td class="c"></td></tr>')

        # ---------- alloc ----------
        alloc_html = ""
        if r["alloc"]:
            lis = "".join(f'<li><b>{esc(a["v"])}</b><span>{esc(a["k"])}</span></li>' for a in r["alloc"])
            alloc_html = (f'<div class="alloc"><h4 class="sec-h">Manpower Allocation</h4>'
                          f'<ul class="alloc-l">{lis}</ul></div>')

        ck = CHECKER_BY_SECTION.get(section)
        ck_name = ck["name"] if ck else "&mdash;"
        ck_email = ck.get("email", "") if ck else ""
        sec_tag = section or "&ndash;"

        # ops compact JSON for client-side cards (raw text; final esc() will encode for the attribute)
        ops_json = []
        for o in ops:
            def rv(v):
                return "" if v is None else str(v)
            ops_json.append([
                fmt_sl(o.get("A")), rv(o.get("B")), rv(o.get("C")) or "",
                fmt_sl(o.get("D")), fmt_sl(o.get("E")), fmt_sl(o.get("F")),
                fmt_time(o.get("G")), fmt_time(o.get("H")), fmt_time(o.get("I")),
                fmt_time(o.get("J")), fmt_time(o.get("K")),
                fmt_dec(o.get("L"), 3), fmt_pct(o.get("M")), fmt_dec(o.get("N"), 4),
                fmt_pct(o.get("O")), fmt_dec(o.get("P"), 5), fmt_hr(o.get("Q"))
            ])
        ops_data = esc(json.dumps(ops_json, ensure_ascii=False))

        bullets.append(
            f'<section class="bulletin" id="ob-{idx}" data-name="{esc(name.lower())}" '
            f'data-code="{esc(code.lower())}" data-group="{esc(group.lower())}" '
            f'data-sec="{esc(section.lower())}" data-sheet="{esc(r["sheet"].lower())}" '
            f'data-smv="{smv_eff if smv_eff else 0}" data-mp="{mp_eff if mp_eff else 0}" '
            f'data-ops="{ops_data}">'

            f'<div class="b-head">'
            f'<div class="b-brand">'
            f'<div class="b-eyebrow">ALEL INDUSTRIES LIMITED &middot; {ORG_LABEL}</div>'
            f'<h2 class="b-title">{esc(name)}</h2>'
            f'<div class="b-sub">Operation Bulletin &middot; Work Study &amp; Capacity (SMV)</div>'
            f'</div>'
            f'<div class="b-badges"><span class="obn">OB-{idx:04d}</span>'
            f'<span class="secp sec-{esc(section.lower())}">{esc(sec_tag)}</span>'
            f'<span class="aprv" data-role="status">Pending</span>'
            f'</div></div>'

            f'<div class="kv-row">'
            f'{photo}'
            f'<div class="kv-wrap">'
            f'<div class="kv"><span>{IC["folder"]} Product Name / SKU</span><b>{esc(name)}</b></div>'
            f'<div class="kv"><span>{IC["code"]} Item Code</span><b class="code">{esc(code) or "&ndash;"}</b></div>'
            f'<div class="kv"><span>Section</span><b>{esc(sec_tag)}</b></div>'
            f'<div class="kv"><span>Product Group</span><b>{esc(group) or "&ndash;"}</b></div>'
            f'<div class="kv"><span>Work Sheet</span><b class="dim">{esc(r["sheet"])}</b></div>'
            f'</div>'
            f'</div>'

            f'<div class="kpis">{kpis}</div>'

            f'<div class="ops-sec">'
            f'<div class="ops-head"><h4 class="sec-h">Operation Sequence &amp; Time Study</h4>'
            f'<span class="ocnt">{len(ops)} operations</span></div>'
            f'<div class="oplist"></div>'
            f'<div class="tw" hidden><table class="ops"><thead><tr>'
            f'<th rowspan="2">SL</th><th rowspan="2" class="l">Operation / Process</th>'
            f'<th rowspan="2">Machine / Manual</th><th rowspan="2">Pcs/Unit</th><th rowspan="2">Pcs/Cycle</th>'
            f'<th rowspan="2">MP</th><th colspan="5">Observed Time (sec)</th>'
            f'<th rowspan="2">Cycle Time<br>(min)</th><th rowspan="2">Rating</th><th rowspan="2">Basic Time</th>'
            f'<th rowspan="2">Allow</th><th rowspan="2">SMV</th><th rowspan="2">Hourly</th>'
            f'</tr><tr><th>T1</th><th>T2</th><th>T3</th><th>T4</th><th>T5</th></tr>'
            f'</thead><tbody>{"".join(tro)}</tbody></table></div>'
            f'</div>'

            f'{alloc_html}'

            f'<div class="wf">'
            f'<div class="wf-step wf-prep" data-role="prepared">'
            f'<div class="wf-ic">{IC["pp"]}</div>'
            f'<div class="wf-tx"><div class="wf-lab">Prepared By</div>'
            f'<div class="wf-name">{" &amp; ".join(esc(n) for n in PREPARED_NAMES)}</div>'
            f'<div class="wf-role">{esc(PREPARED_BY["role"])}</div></div>'
            f'<div class="wf-st"><span class="st-pill" data-role="pill">Auto</span></div>'
            f'</div>'
            f'<div class="wf-arrow">&rarr;</div>'
            f'<div class="wf-step wf-check" data-role="checked">'
            f'<div class="wf-ic">{IC["ck"]}</div>'
            f'<div class="wf-tx"><div class="wf-lab">Checked By</div>'
            f'<div class="wf-name">{esc(ck_name)}</div>'
            f'<div class="wf-role">{esc(sec_tag)} Section</div>'
            f'<div class="wf-mail">{esc(ck_email)}</div></div>'
            f'<div class="wf-st"><span class="st-pill" data-role="pill">Pending</span></div>'
            f'</div>'
            f'<div class="wf-arrow">&rarr;</div>'
            f'<div class="wf-step wf-approve" data-role="approved">'
            f'<div class="wf-ic">{IC["ap"]}</div>'
            f'<div class="wf-tx"><div class="wf-lab">Approved By</div>'
            f'<div class="wf-name">{esc(APPROVED_BY["name"])}</div>'
            f'<div class="wf-role">{esc(APPROVED_BY["role"])}</div>'
            f'<div class="wf-mail">{esc(APPROVED_BY.get("email",""))}</div></div>'
            f'<div class="wf-st"><span class="st-pill" data-role="pill">Pending</span></div>'
            f'</div>'
            f'</div>'
            f'</section>')

    # ---------- chips ----------
    sec_counts = {"GSS": 0, "LED": 0, "HAP": 0, "OTHER": 0}
    for r in results:
        s = (r.get("section") or "").strip().upper()
        key = s if s in sec_counts else "OTHER"
        sec_counts[key] += 1
    chip_html = ""
    order = ["GSS", "LED", "HAP", "OTHER"]
    chip_html = '<button class="chip on" data-sec="all">All Bulletins<i>%TOTAL%</i></button>'
    for k in order:
        if sec_counts.get(k):
            lbl = {"OTHER": "Other"}.get(k, k)
            chip_html += f'<button class="chip" data-sec="{esc(k.lower())}">{esc(lbl)} Section<i>{sec_counts[k]}</i></button>'

    people = []
    people.append({"k": PREPARED_BY["name"], "l": PREPARED_BY["name"] + " (Prepared)"})
    for sec in sorted(CHECKER_BY_SECTION):
        people.append({"k": CHECKER_BY_SECTION[sec]["name"], "l": CHECKER_BY_SECTION[sec]["name"] + " (" + sec + " Check)"})
    people.append({"k": APPROVED_BY["name"], "l": APPROVED_BY["name"] + " (Approve)"})

    # JSON for <script type="application/json"> blocks: only escape "</script"
    def js_json(obj):
        return json.dumps(obj, ensure_ascii=False).replace("</", "<\\/")
    cat_json = js_json(catalog)
    people_json = js_json(people)
    users_json = js_json(login_user_db())

    html_doc = HTML_TEMPLATE.replace("%BULLETINS%", "".join(bullets)) \
                            .replace("%CHIPS%", chip_html) \
                            .replace("%TOTAL%", str(len(results))) \
                            .replace("%REFRESHED%", today) \
                            .replace("%CATALOG%", cat_json) \
                            .replace("%PEOPLE%", people_json) \
                            .replace("%USERS%", users_json) \
                            .replace("%ICON_SEARCH%", IC["search"]) \
                            .replace("%ICON_REFRESH%", IC["refresh"])
    dest = out_path or OUT_HTML
    with open(dest, "w", encoding="utf-8") as f:
        f.write(html_doc)
    log("      written: " + dest)
    log("      size MB: " + str(round(os.path.getsize(dest) / 1048576, 2)))


def main():
    import argparse
    ap = argparse.ArgumentParser(description="ALEL Operation Bulletin generator")
    ap.add_argument("--output", default=None)
    ap.add_argument("--sheet-id", default=None)
    args = ap.parse_args()
    global SHEET_ID, XLSX_TMP
    if args.sheet_id:
        SHEET_ID = args.sheet_id
        XLSX_TMP = os.path.join(HERE, "_latest_capacity_study.xlsx")
    try:
        download()
        results, media = parse_all()
        if len(results) < 50:
            raise SystemExit("Expected 100+ bulletins, got %d" % len(results))
        build_html(results, media, out_path=args.output)
        log("[4/4] DONE.")
    except SystemExit as e:
        log(str(e))
        sys.exit(1)
    except Exception as e:
        log("ERROR: " + str(e))
        import traceback
        traceback.print_exc()
        sys.exit(1)


HTML_TEMPLATE = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
<title>ALEL Operation Bulletin &mdash; Searchable + Approval</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&family=Sora:wght@600;700;800&display=swap" rel="stylesheet">
<style>
  :root{--bg:#0b1020;--bg2:#111a30;--card:#fff;--ink:#0f172a;--mut:#5b6b81;--navy:#0f2b46;
        --teal:#0e9f7a;--teal-d:#0b7a5f;--amber:#e8a13c;--blue:#3b62c7;--red:#d1365a;--violet:#6a45d9;
        --line:#e6ebf2;--line2:#d3deeb;--soft:#f4f7fb;--shadow:0 18px 50px -18px rgba(4,12,30,.35);--r:16px;}
  *{box-sizing:border-box;margin:0;padding:0;}
  html,body{overflow-x:hidden;} html{scroll-behavior:smooth;}
  body{font-family:'Inter',system-ui,'Segoe UI',Arial,sans-serif;color:var(--ink);
       background:radial-gradient(1100px 520px at 85% -10%,rgba(124,92,240,.22),transparent 60%),
                  radial-gradient(900px 480px at -10% 0%,rgba(14,159,122,.20),transparent 55%),
                  linear-gradient(160deg,var(--bg),var(--bg2));background-attachment:fixed;padding-bottom:70px;}
  ::selection{background:rgba(14,159,122,.25);}
  ::-webkit-scrollbar{width:10px;height:10px;} ::-webkit-scrollbar-thumb{background:#2a3a56;border-radius:8px;}
  button{font-family:inherit;}

  /* ---------- topbar ---------- */
  .topbar{position:sticky;top:0;z-index:100;backdrop-filter:blur(14px);-webkit-backdrop-filter:blur(14px);
          background:rgba(8,13,27,.8);border-bottom:1px solid rgba(255,255,255,.08);transition:transform .28s;}
  body.hide-top .topbar{transform:translateY(-105%);}
  .tb-in{max-width:1280px;margin:0 auto;padding:10px 16px 0;display:grid;gap:8px;grid-template-columns:1fr;}
  .tb-row1{display:flex;align-items:center;gap:10px;flex-wrap:wrap;}
  .brand-h{display:flex;align-items:center;gap:10px;color:#fff;min-width:0;}
  .logo-link{display:inline-flex;text-decoration:none;border-radius:12px;}
  .logo-link:hover .logo{transform:scale(1.05);box-shadow:0 10px 24px -6px rgba(14,159,122,.8);}
  .logo{width:38px;height:38px;border-radius:11px;flex:none;display:grid;place-items:center;cursor:pointer;transition:.2s;
        background:linear-gradient(135deg,#12b886,#0b7a5f);box-shadow:0 8px 20px -6px rgba(14,159,122,.55);}
  .brand-h h1{font-family:'Sora',sans-serif;font-size:15px;line-height:1.1;min-width:0;}
  .brand-h h1 small{display:block;font-weight:500;font-size:9.5px;letter-spacing:.14em;color:#9fb3d0;text-transform:uppercase;}
  .who{display:flex;align-items:center;gap:8px;margin-left:auto;color:#cfd9ea;}
  .who-user{font-size:11.5px;font-weight:700;color:#fff;background:rgba(255,255,255,.1);border:1px solid rgba(255,255,255,.2);
            border-radius:10px;padding:6px 10px;white-space:nowrap;max-width:230px;overflow:hidden;text-overflow:ellipsis;}
  .logout{background:rgba(255,255,255,.1);color:#ffd1a8;border:1px solid rgba(255,255,255,.25);border-radius:10px;
          padding:6px 10px;font-size:11px;font-weight:700;cursor:pointer;}
  .logout:hover{background:#c0392b;color:#fff;border-color:#c0392b;}
  .tbtns{display:flex;gap:6px;}
  .tbtn{display:inline-flex;align-items:center;gap:6px;background:rgba(255,255,255,.1);color:#fff;
        border:1px solid rgba(255,255,255,.2);border-radius:11px;padding:8px 11px;font-size:12px;font-weight:700;
        cursor:pointer;white-space:nowrap;}
  .tbtn:hover{background:rgba(255,255,255,.2);}
  .tbtn.on{background:linear-gradient(135deg,#ffb02e,#ff6b35);border-color:transparent;}
  .tbtn.refresh{background:linear-gradient(135deg,#12b886,#0e9f7a);border-color:transparent;}
  .auto-dot{width:8px;height:8px;border-radius:50%;background:#fff;opacity:.5;animation:pulse 1.6s infinite;}
  .tbtn.on .auto-dot{opacity:1;}
  @keyframes pulse{50%{transform:scale(1.6);opacity:.3;}}
  .tb-row2{position:relative;display:flex;align-items:center;gap:8px;}
  .searchbox{flex:1;display:flex;align-items:center;gap:8px;background:rgba(255,255,255,.08);
             border:1px solid rgba(255,255,255,.15);border-radius:12px;padding:9px 12px;color:#cdd9ea;}
  .searchbox:focus-within{border-color:rgba(14,159,122,.7);box-shadow:0 0 0 4px rgba(14,159,122,.18);}
  .searchbox input{flex:1;background:transparent;border:0;outline:0;color:#fff;font-size:15px;min-width:0;font-family:inherit;}
  .searchbox input::placeholder{color:#7d8ea8;}
  .clr{display:none;background:none;border:0;color:#9fb3d0;font-size:18px;cursor:pointer;line-height:1;}
  .sugg{position:absolute;left:0;right:0;top:calc(100% + 6px);background:#0e1830;border:1px solid rgba(255,255,255,.15);
        border-radius:14px;overflow:hidden;display:none;box-shadow:0 20px 50px -12px rgba(0,0,0,.7);z-index:110;}
  .sugg .sg{display:flex;gap:10px;align-items:center;padding:9px 12px;color:#dbe6f5;font-size:12.5px;cursor:pointer;
            border-bottom:1px solid rgba(255,255,255,.05);}
  .sugg .sg:hover{background:rgba(255,255,255,.07);}
  .sugg .sg b{color:#fff;font-weight:700;}
  .sugg .sg .sc{color:#4adeb0;font-variant-numeric:tabular-nums;}
  .sugg .sg .ss{margin-left:auto;font-size:10px;color:#8fa0bb;border:1px solid rgba(255,255,255,.25);border-radius:99px;padding:1px 8px;}
  .tb-note{color:#8fa0bb;font-size:10.5px;padding:0 2px 8px;}
  .tb-note b{color:#c7d4e8;}

  /* chips */
  .chips{display:flex;gap:6px;overflow-x:auto;-webkit-overflow-scrolling:touch;padding:4px 16px 12px;scrollbar-width:none;}
  .chips::-webkit-scrollbar{display:none;}
  .chip{flex:0 0 auto;display:inline-flex;align-items:center;gap:6px;background:rgba(255,255,255,.07);color:#cfd9ea;
        border:1px solid rgba(255,255,255,.14);padding:7px 12px;border-radius:999px;font-size:12px;font-weight:600;cursor:pointer;}
  .chip i{font-style:normal;background:rgba(255,255,255,.14);border-radius:99px;padding:1px 7px;font-size:10px;}
  .chip.on{background:linear-gradient(135deg,#12b886,#0e9f7a);border-color:transparent;color:#fff;}

  main{max-width:1280px;margin:6px auto 0;padding:0 14px;}
  .statbar{display:flex;gap:10px;align-items:center;flex-wrap:wrap;color:#9fb3d0;font-size:12px;margin:0 0 12px;}
  .statbar .lt{color:#e7eef9;font-weight:700;} .statbar .lt span{color:#4adeb0;}
  .statbar .fld{display:flex;gap:6px;margin-left:auto;}
  .statbar .fld button{background:rgba(255,255,255,.1);color:#cfd9ea;border:1px solid rgba(255,255,255,.15);
        border-radius:999px;padding:6px 12px;font-size:11.5px;font-weight:700;cursor:pointer;}
  .statbar .fld button.on{background:linear-gradient(135deg,#ffb02e,#ff6b35);color:#fff;border-color:transparent;}

  /* ---------- bulletin card ---------- */
  .bulletin{background:var(--card);border-radius:var(--r);margin:0 0 20px;overflow:hidden;box-shadow:var(--shadow);
            border:1px solid rgba(255,255,255,.5);}
  .bulletin.hidden{display:none;}
  .b-head{display:flex;justify-content:space-between;align-items:flex-start;gap:12px;flex-wrap:wrap;color:#fff;
          padding:16px 20px;background:linear-gradient(120deg,#0b1c38,#12314f 55%,#0e7a63);position:relative;overflow:hidden;}
  .b-eyebrow{font-size:10px;letter-spacing:.16em;text-transform:uppercase;color:#9fd8c6;}
  .b-title{font-family:'Sora',sans-serif;font-size:20px;margin:2px 0 2px;line-height:1.2;word-break:break-word;}
  .b-sub{font-size:11px;color:#c6d6ea;}
  .b-badges{display:flex;gap:7px;flex-wrap:wrap;align-items:flex-start;}
  .obn,.secp,.aprv{font-size:10.5px;font-weight:800;padding:5px 10px;border-radius:999px;letter-spacing:.03em;}
  .obn{background:#ffd166;color:#5b3a00;}
  .secp{background:rgba(255,255,255,.15);border:1px solid rgba(255,255,255,.3);color:#fff;}
  .secp.sec-gss{background:rgba(59,98,199,.35);}
  .secp.sec-led{background:rgba(232,161,60,.3);}
  .secp.sec-hap{background:rgba(14,159,122,.35);}
  .aprv{color:#fff;border:1px solid rgba(255,255,255,.35);background:rgba(255,255,255,.12);}
  .aprv.a-done{background:linear-gradient(135deg,#12b886,#0b7a5f);border-color:transparent;}
  .aprv.a-pend{background:linear-gradient(135deg,#ffb02e,#ff6b35);border-color:transparent;}
  .aprv.a-prep{background:linear-gradient(135deg,#5b8def,#3b62c7);border-color:transparent;}
  .kv-row{display:flex;gap:16px;padding:14px 20px 6px;align-items:flex-start;}
  .photo{flex:0 0 110px;width:110px;height:118px;border:1px solid var(--line);border-radius:12px;padding:6px;
         background:linear-gradient(180deg,#fff,var(--soft));position:relative;display:flex;align-items:center;justify-content:center;}
  .photo img{max-width:100%;max-height:100%;object-fit:contain;filter:drop-shadow(0 6px 12px rgba(20,40,70,.12));}
  .photo .ptag{position:absolute;bottom:4px;font-size:7.5px;letter-spacing:.16em;color:#8fa0b8;}
  .photo.empty{border-style:dashed;color:#9fb0c4;font-size:11px;text-align:center;}
  .kv-wrap{flex:1;min-width:0;display:grid;gap:4px;}
  .kv{display:flex;flex-wrap:wrap;gap:6px;padding:5px 2px;border-bottom:1px dashed var(--line);align-items:baseline;}
  .kv span{color:var(--mut);font-size:11px;font-weight:700;min-width:150px;display:inline-flex;gap:6px;align-items:center;}
  .kv span svg{color:var(--teal);}
  .kv b{font-size:13.5px;word-break:break-word;}
  .kv b.code{color:var(--blue);}
  .kv b.dim{color:var(--mut);font-weight:500;}

  /* kpis */
  .kpis{display:grid;grid-template-columns:repeat(5,1fr);gap:10px;padding:6px 20px 4px;}
  .kpi{background:linear-gradient(180deg,#fff,#fafcfe);border:1px solid var(--line);border-radius:13px;padding:10px 12px;}
  .k-top{display:flex;align-items:center;gap:7px;color:var(--mut);font-size:10px;font-weight:800;text-transform:uppercase;letter-spacing:.04em;}
  .k-top svg{color:#fff;}
  .k-top{color:#fff;} .k-top span{color:var(--mut);}
  .kpi .k-top .k-ic{width:28px;height:28px;border-radius:8px;display:grid;place-items:center;color:#fff;}
  .k-smv .k-top .k-ic{background:linear-gradient(135deg,#12b886,#0b7a5f);} .k-smv .k-val{color:#0b7a5f;}
  .k-mp .k-top .k-ic{background:linear-gradient(135deg,#5b8def,#3b62c7);} .k-mp .k-val{color:#3b62c7;}
  .k-hc .k-top .k-ic{background:linear-gradient(135deg,#f0a83c,#d97f17);} .k-hc .k-val{color:#c57412;}
  .k-prod .k-top .k-ic{background:linear-gradient(135deg,#8a6bf0,#6a45d9);} .k-prod .k-val{color:#6a45d9;}
  .k-ops .k-top .k-ic{background:linear-gradient(135deg,#e85f7e,#d1365a);} .k-ops .k-val{color:#d1365a;}
  .live{font-style:normal;font-size:8px;color:#ff6b35;background:#fff1ea;border-radius:99px;padding:1px 6px;margin-left:auto;}
  .k-val{font-family:'Sora',sans-serif;font-size:21px;font-weight:800;margin-top:6px;font-variant-numeric:tabular-nums;}
  .k-unit{font-size:10px;color:#94a3b5;margin-top:1px;}
  .mp-ctrl{display:flex;align-items:center;gap:5px;margin-top:7px;}
  .mp-btn{flex:0 0 34px;height:34px;border-radius:9px;border:1px solid var(--line2);background:#fff;color:var(--navy);
          font-size:19px;font-weight:800;cursor:pointer;display:grid;place-items:center;}
  .mp-btn:active{transform:scale(.92);}
  .mp-input{flex:1;min-width:0;height:34px;border:1px solid var(--line2);border-radius:9px;text-align:center;
            font-family:'Sora',sans-serif;font-size:16px;font-weight:800;color:#3b62c7;outline:none;background:#f8fbff;
            -moz-appearance:textfield;appearance:textfield;}
  .mp-input::-webkit-inner-spin-button,.mp-input::-webkit-outer-spin-button{-webkit-appearance:none;margin:0;}
  .mp-reset{font-size:9px;color:#7c8ba0;background:none;border:none;cursor:pointer;text-decoration:underline dotted;margin-top:4px;}
  .kpi-na{font-size:20px;color:#b6c2d2;margin-top:8px;}

  /* ops */
  .ops-sec{padding:8px 20px 6px;}
  .ops-head{display:flex;justify-content:space-between;align-items:center;gap:10px;margin:6px 0 10px;}
  .sec-h{font-size:12px;font-weight:800;text-transform:uppercase;letter-spacing:.08em;color:var(--navy);
         display:flex;align-items:center;gap:8px;}
  .sec-h::before{content:"";width:4px;height:15px;border-radius:3px;background:linear-gradient(180deg,#12b886,#0b7a5f);}
  .ocnt{font-size:11px;font-weight:700;color:var(--teal-d);background:#e7f8f2;border:1px solid #c5ecdd;padding:3px 10px;border-radius:99px;white-space:nowrap;}
  .oplist{display:grid;gap:8px;}
  .opc{display:grid;grid-template-columns:auto 1fr auto auto;gap:11px;align-items:start;
       background:#fff;border:1px solid var(--line);border-radius:13px;padding:10px 12px;
       transition:.16s;cursor:pointer;}
  .opc:hover{border-color:#9fd8c6;box-shadow:0 8px 18px -14px rgba(11,122,95,.5);}
  .opc.open{background:#f4fcf9;border-color:#12b886;}
  .opc-no{width:30px;height:30px;border-radius:9px;display:grid;place-items:center;color:#fff;
          font-weight:800;font-size:13px;background:linear-gradient(135deg,#0f2b46,#1d4567);}
  .opc-main{min-width:0;}
  .opc-name{font-size:13.5px;font-weight:700;line-height:1.35;}
  .opc-meta{display:flex;flex-wrap:wrap;gap:5px 12px;margin-top:5px;color:var(--mut);font-size:11px;}
  .opc-meta b{color:var(--ink);}
  .opc-right{text-align:right;}
  .opc-smv{color:var(--teal-d);font-family:'Sora',sans-serif;font-weight:800;font-size:15px;white-space:nowrap;}
  .opc-smv small{display:block;font-size:8px;color:#8fa0b8;text-transform:uppercase;font-family:'Inter';}
  .caret{color:#7c8ba0;transition:.2s;align-self:center;}
  .opc.open .caret{transform:rotate(180deg);}
  .opc-detail{display:none;}
  .opc.open .opc-detail{display:grid;grid-template-columns:repeat(auto-fill,minmax(86px,1fr));gap:6px;margin-top:8px;}
  .od{background:#fff;border:1px solid var(--line);border-radius:9px;padding:5px 8px;}
  .od span{display:block;font-size:8px;color:#8fa0b8;text-transform:uppercase;letter-spacing:.03em;}
  .od b{font-size:12px;}
  .alloc{padding:0 20px 6px;}
  .alloc-l{list-style:none;display:grid;grid-template-columns:repeat(auto-fill,minmax(210px,1fr));gap:6px;}
  .alloc-l li{display:flex;gap:8px;align-items:center;background:var(--soft);border:1px solid var(--line);
              border-radius:9px;padding:5px 9px;font-size:11.5px;}
  .alloc-l li b{min-width:28px;text-align:center;background:#fff;border:1px solid var(--line2);color:#0b7a5f;border-radius:7px;padding:1px 6px;}
  .alloc-l li span{color:#33445c;font-weight:500;}

  /* table (print) */
  .tw{margin-top:10px;}
  table.ops{border-collapse:collapse;width:100%;font-size:11px;}
  table.ops th,table.ops td{border:1px solid var(--line);padding:5px 7px;text-align:center;}
  table.ops thead th{background:var(--navy);color:#fff;}
  table.ops td.l{text-align:left;}
  td.smv{color:var(--teal-d);font-weight:800;}
  tr.total td{background:linear-gradient(180deg,#0f2b46,#163a5c);color:#fff;font-weight:700;}

  /* workflow / approval */
  .wf{display:grid;grid-template-columns:1fr auto 1fr auto 1fr;gap:6px;align-items:stretch;padding:12px 20px 16px;
      background:linear-gradient(180deg,#fbfdff,#f2f6fb);border-top:1px solid var(--line);}
  .wf-step{display:flex;gap:10px;align-items:flex-start;background:#fff;border:1px solid var(--line);
           border-radius:12px;padding:10px 12px;position:relative;min-width:0;}
  .wf-ic{flex:0 0 34px;width:34px;height:34px;border-radius:10px;display:grid;place-items:center;color:#fff;
         background:linear-gradient(135deg,#0f2b46,#1d4567);}
  .wf-tx{min-width:0;}
  .wf-lab{font-size:9.5px;text-transform:uppercase;letter-spacing:.06em;color:var(--mut);font-weight:800;}
  .wf-name{font-size:13px;font-weight:800;color:var(--ink);line-height:1.25;word-break:break-word;}
  .wf-role,.wf-mail{font-size:10px;color:var(--mut);word-break:break-word;}
  .wf-mail{color:#7c8ba0;font-size:9.5px;}
  .wf-st{margin-left:auto;align-self:center;}
  .st-pill{font-size:9px;font-weight:800;border-radius:99px;padding:3px 9px;white-space:nowrap;display:inline-block;
           background:#eef1f6;color:#5b6b81;}
  .wf-step.st-active .st-pill{background:#fff0db;color:#b45309;}
  .wf-step.st-done .st-pill{background:#d8f6ec;color:#0b7a5f;}
  .wf-step.st-done .wf-ic{background:linear-gradient(135deg,#12b886,#0b7a5f);}
  .wf-step.st-active .wf-ic{background:linear-gradient(135deg,#ffb02e,#e08b1d);}
  .wf-arrow{display:flex;align-items:center;color:#9fb0c0;font-size:20px;}
  .wf-step .act{grid-column:1/-1;}
  .wf-btn{width:100%;display:inline-flex;justify-content:center;align-items:center;gap:7px;padding:9px;
          border:0;border-radius:10px;font-weight:800;font-size:12.5px;cursor:pointer;color:#fff;}
  .wf-btn.ck{background:linear-gradient(135deg,#5b8def,#3b62c7);}
  .wf-btn.ap{background:linear-gradient(135deg,#12b886,#0b7a5f);}
  .wf-btn.ghost{background:#fff;color:#5b6b81;border:1px solid var(--line2);font-weight:700;}
  .wf-btn:disabled{background:#ccd6e2;cursor:not-allowed;}
  .wf-hint{margin-top:7px;font-size:10px;color:#b45309;font-weight:700;}
  .wf-step.st-active{background:#fffbef;border-color:#e8c37a;border-style:dashed;}
  .wf-step.st-done{background:#f0fbf7;border-color:#bde7d8;}
  .wf-signed{display:inline-flex;align-items:center;gap:5px;font-size:11px;color:#0b7a5f;font-weight:700;}
  .wf-signed svg{color:#0b7a5f;}

  .to-top{position:fixed;right:14px;bottom:16px;z-index:80;width:46px;height:46px;border-radius:14px;border:0;
          display:none;align-items:center;justify-content:center;color:#fff;cursor:pointer;
          background:linear-gradient(135deg,#12b886,#0e9f7a);box-shadow:0 10px 24px -8px rgba(14,159,122,.7);}
  .to-top.show{display:flex;}
  .msg{display:none;text-align:center;color:#cbd7e8;padding:50px 20px;}
  .foot{max-width:1280px;margin:10px auto;padding:0 16px;color:#6f83a3;font-size:11px;text-align:center;}

  /* ---------- login overlay ---------- */
  .login{position:fixed;inset:0;z-index:500;display:flex;align-items:center;justify-content:center;
         padding:16px;background:radial-gradient(900px 500px at 80% -10%,rgba(124,92,240,.28),transparent 55%),
         radial-gradient(700px 400px at -10% 110%,rgba(14,159,122,.25),transparent 55%),
         linear-gradient(160deg,#0b1020,#111a30);}
  .login-card{width:100%;max-width:400px;background:#fff;border-radius:20px;padding:28px 26px;
              box-shadow:0 30px 80px -20px rgba(0,0,0,.6);animation:rise .4s ease;}
  .login-logo{width:58px;height:58px;border-radius:16px;margin:0 auto 12px;display:grid;place-items:center;
              font-size:28px;background:linear-gradient(135deg,#12b886,#0b7a5f);box-shadow:0 12px 26px -8px rgba(14,159,122,.6);}
  .login-card h2{font-family:'Sora',sans-serif;text-align:center;font-size:20px;color:var(--ink);}
  .login-card > p{text-align:center;color:var(--mut);font-size:12.5px;margin:4px 0 4px;}
  .login-f{display:grid;gap:6px;margin-top:6px;}
  .login-f label{font-size:11px;font-weight:700;color:var(--mut);text-transform:uppercase;letter-spacing:.05em;}
  .login-f input{padding:11px 12px;border:1px solid var(--line2);border-radius:11px;font-size:14px;outline:none;font-family:inherit;}
  .login-f input:focus{border-color:var(--teal);box-shadow:0 0 0 3px rgba(14,159,122,.15);}
  .login-err{color:#c0392b;font-size:12px;font-weight:700;min-height:16px;}
  .login-f button#loginBtn{padding:12px;border:0;border-radius:11px;background:linear-gradient(135deg,#12b886,#0b7a5f);
        color:#fff;font-weight:800;font-size:14px;cursor:pointer;margin-top:4px;}
  .login-f button#loginBtn:hover{filter:brightness(1.06);}
  .hint-btn{background:none;border:none;color:#7c8ba0;font-size:11px;cursor:pointer;text-decoration:underline dotted;margin-top:6px;}
  .login-links{display:flex;justify-content:space-between;gap:8px;margin-top:2px;}
  .login-links .hint-btn{background:none;border:none;color:#3b62c7;font-size:11.5px;font-weight:600;cursor:pointer;padding:4px 0;}
  .login-links .hint-btn:hover{text-decoration:underline;}
  .panel-txt{font-size:12px;color:var(--mut);line-height:1.5;margin-bottom:6px;}
  .ok-msg{display:none;margin-top:12px;background:#e7f8f2;border:1px solid #bde7d8;color:#0b7a5f;
          border-radius:11px;padding:10px 12px;font-size:12.5px;line-height:1.5;white-space:pre-line;text-align:center;}
  .ok-msg.show{display:block;}
  .adm-wrap{position:fixed;inset:0;z-index:600;display:flex;align-items:flex-start;justify-content:center;
            padding:40px 14px;background:rgba(5,10,20,.7);backdrop-filter:blur(4px);overflow:auto;}
  .adm-box{width:100%;max-width:680px;background:#fff;border-radius:18px;overflow:hidden;box-shadow:0 30px 80px -20px rgba(0,0,0,.7);}
  .adm-head{display:flex;justify-content:space-between;align-items:center;padding:14px 18px;
            background:linear-gradient(135deg,#0f2b46,#163a5c);color:#fff;}
  .adm-head h3{margin:0;}
  .adm-head button{background:none;border:0;color:#fff;font-size:24px;cursor:pointer;line-height:1;}
  .adm-body{padding:16px 18px;max-height:70vh;overflow:auto;}
  .adm-body h3{font-size:13px;text-transform:uppercase;letter-spacing:.06em;color:var(--navy);margin:16px 0 8px;border-bottom:2px solid var(--navy);padding-bottom:4px;}
  .adm-body h3:first-child{margin-top:0;}
  .adm-body .empty{color:var(--mut);font-size:12.5px;padding:4px 0;}
  .req-list,.user-list{list-style:none;}
  .req-list li,.user-list li{display:flex;flex-wrap:wrap;gap:6px 10px;align-items:center;padding:8px 10px;
        border:1px solid var(--line);border-radius:11px;margin-bottom:6px;background:#fff;}
  .req-list b,.user-list b{min-width:150px;}
  .req-list .rq{font-size:10.5px;font-weight:700;color:#fff;background:var(--blue);border-radius:99px;padding:2px 9px;}
  .req-list .rw,.user-list code{font-size:10.5px;color:#8fa0b8;}
  .user-list code{margin-left:auto;}
  .rq-act{margin-left:auto;display:flex;gap:6px;}
  .mini{font-size:11px;font-weight:700;padding:5px 11px;border-radius:8px;border:1px solid var(--line2);background:#fff;cursor:pointer;}
  .mini.ok{background:linear-gradient(135deg,#12b886,#0b7a5f);color:#fff;border:0;}
  .mini.no{color:#b3392c;border-color:#e7b3ac;}
  .adm-note{font-size:10.5px;color:#8fa0bb;margin:12px 0 0;padding:0 18px 14px;}
  .tbtn.admin{background:rgba(255,255,255,.12);}
  .tbtn.admin.has{border-color:#ffd166;color:#ffd166;animation:glow 1.6s infinite;}
  @keyframes glow{50%{box-shadow:0 0 0 4px rgba(255,209,102,.18);}}

  @media (min-width:821px){
    .opc:hover{transform:translateX(2px);}
    .kpis{grid-template-columns:repeat(5,1fr);}
  }
  @media (max-width:820px){
    .tb-in{padding:8px 10px 0;}
    .brand-h h1{font-size:13px;}
    .who{margin-left:0;}
    .who label{display:none;}
    .who select{max-width:150px;font-size:11px;padding:6px;}
    .tbtn span.lb{display:none;}
    .searchbox input{font-size:16px;}
    .b-head{padding:13px 14px;}
    .b-title{font-size:17px;}
    .kv-row{padding:12px 14px 4px;}
    .photo{flex:0 0 84px;width:84px;height:96px;}
    .kpis{grid-template-columns:1fr 1fr;padding:6px 14px 2px;}
    .k-ops,.kpi:last-child{grid-column:auto;}
    .ops-sec,.alloc{padding-left:14px;padding-right:14px;}
    .wf{grid-template-columns:1fr;padding:12px 14px;}
    .wf-arrow{display:none;}
    .wf-step{width:100%;}
    .b-badges{width:100%;}
    .chips{padding-left:12px;padding-right:12px;}
    main{padding:0 8px;}
    .statbar{padding:0 4px;}
  }
  @media print{
    body{background:#fff;padding:0;}
    .topbar,.statbar,.to-top,.foot{display:none !important;}
    .bulletin{break-after:page;box-shadow:none;border-radius:0;}
    .oplist{display:none;}
    .tw{display:block !important;}
    .wf{break-inside:avoid;}
    @page{size:A4 portrait;margin:8mm;}
  }
  @media (max-width:820px) and (min-width:0px){
    .kpis{grid-template-columns:repeat(2,1fr);}
    .kpis .kpi:nth-child(5){grid-column:1/3;}
  }
</style>
</head>
<body>
<div class="topbar">
  <div class="tb-in">
    <div class="tb-row1">
      <div class="brand-h">
        <a class="logo-link" href="https://akij-light-engineering-ltd.github.io/ALEL-All-Entry/" title="Back to ALEL Home">
          <div class="logo">&#127968;</div>
        </a>
        <h1>ALEL Operation Bulletin<small>Live SMV &amp; Digital Approval</small></h1>
      </div>
      <div class="who" id="whoBox">
        <span class="who-user" id="whoUser"></span>
        <button class="logout" id="logoutBtn" type="button">Logout</button>
      </div>
      <div class="tbtns">
        <button class="tbtn" id="autoBtn" type="button"><span class="auto-dot"></span><span class="lb" id="autoLab">Auto 1m</span></button>
        <button class="tbtn refresh" id="refreshBtn" type="button">%ICON_REFRESH%<span class="lb">Refresh</span></button>
      </div>
      <button class="tbtn admin" id="adminBtn" type="button" style="display:none">&#128110; Admin<i></i></button>
    </div>
    <div class="tb-row2">
      <div class="searchbox">%ICON_SEARCH%
        <input id="q" type="text" placeholder="Search SKU name or item code… e.g. MCB, 14414, BULB" autocomplete="off">
        <button class="clr" id="clr" title="Clear">&times;</button>
      </div>
      <div class="sugg" id="sugg"></div>
    </div>
    <div class="tb-note">Approval: <b>Prepared (IE)</b> &rarr; <b>Checked</b> (section head) &rarr; <b>Approved</b> (Plant Head) &middot; login as your role to act &middot; Refreshed <b>%REFRESHED%</b></div>
  </div>
  <div class="chips" id="chips">%CHIPS%</div>
</div>

<!-- login overlay -->
<div class="login" id="loginOverlay">
  <div class="login-card">
    <div class="login-logo">&#9889;</div>
    <h2>ALEL Operation Bulletin</h2>
    <p id="loginSub">Sign in with your office email</p>
    <div class="login-f" id="loginForm">
      <label>Email</label>
      <input id="loginId" type="email" autocomplete="username" placeholder="yourname@akijlightengineering.com">
      <label>Password</label>
      <input id="loginPass" type="password" autocomplete="current-password" placeholder="Password">
      <div class="login-err" id="loginErr"></div>
      <button id="loginBtn" type="button">Sign In</button>
      <div class="login-links">
        <button class="hint-btn" id="forgotBtn" type="button">Forgot password?</button>
        <button class="hint-btn" id="requestBtn" type="button">New user? Request access</button>
      </div>
    </div>
    <div class="login-f" id="forgotPanel" style="display:none">
      <p class="panel-txt">Enter your office email and we will send a reset link to the administrator. (Offline demo: the request appears in the Admin panel.)</p>
      <input id="forgotEmail" type="email" placeholder="yourname@akijlightengineering.com">
      <div class="login-err" id="forgotErr"></div>
      <button id="forgotSend" type="button">Send Reset Request</button>
      <button class="hint-btn" id="backLogin1" type="button">&larr; Back to login</button>
    </div>
    <div class="login-f" id="requestPanel" style="display:none">
      <p class="panel-txt">Fill in your details. The Admin will approve your account.</p>
      <input id="reqName" type="text" placeholder="Full name">
      <input id="reqEmail" type="email" placeholder="Office email">
      <select id="reqSection">
        <option value="">Select your section</option>
        <option value="GSS">GSS</option>
        <option value="LED">LED</option>
        <option value="HAP">HAP</option>
      </select>
      <div class="login-err" id="reqErr"></div>
      <button id="reqSend" type="button">Send Access Request</button>
      <button class="hint-btn" id="backLogin2" type="button">&larr; Back to login</button>
    </div>
    <div class="ok-msg" id="loginOk"></div>
  </div>
</div>
<main>
  <div class="statbar">
    <div class="lt"><span id="cnt">0</span> shown &middot; Pending <span id="cntPend">0</span> &middot; Approved <span id="cntApr">0</span></div>
    <div class="fld" id="viewFld">
      <button class="on" data-v="all">All</button>
      <button data-v="pend">Pending</button>
      <button data-v="appr">Approved Only</button>
    </div>
  </div>
  %BULLETINS%
  <div class="msg" id="msg"><b>No bulletins found.</b></div>
</main>
<button class="to-top" id="toTop" type="button" aria-label="Back to top">
  <svg viewBox="0 0 24 24" width="20" height="20" fill="none" stroke="currentColor" stroke-width="2.6" stroke-linecap="round" stroke-linejoin="round"><polyline points="18 15 12 9 6 15"/></svg>
</button>
<!-- admin dialog -->
<div class="adm-wrap" id="adminDlg" style="display:none">
  <div class="adm-box">
    <div class="adm-head"><h3>&#128110; Admin Panel</h3><button id="admClose" type="button">&times;</button></div>
    <div class="adm-body" id="adminBody"></div>
    <p class="adm-note">Offline demo: users &amp; requests are saved in this browser only. For a real shared system (email login, cross-device requests) a Google Apps Script / backend is needed.</p>
  </div>
</div>
<div class="foot">ALEL Industries Limited &middot; Operational Excellence &middot; Operation Bulletin</div>
<script id="data-cat" type="application/json">%CATALOG%</script>
<script id="data-people" type="application/json">%PEOPLE%</script>
<script id="data-users" type="application/json">%USERS%</script>
<script>
(function(){
  var CATALOG = JSON.parse(document.getElementById('data-cat').textContent || '[]');
  var PEOPLE = JSON.parse(document.getElementById('data-people').textContent || '[]');
  var USERS = JSON.parse(document.getElementById('data-users').textContent || '[]');
  var STORE = {};
  try{ STORE = JSON.parse(localStorage.getItem('alel_approvals_v1')||'{}'); }catch(e){}
  var SESSION = null;   // {id,name,role,section}
  try{ var s=localStorage.getItem('alel_login_v1'); if(s){ SESSION=JSON.parse(s); } }catch(e){}

  var bullets = Array.prototype.slice.call(document.querySelectorAll('.bulletin'));
  var qEl=document.getElementById('q'), clr=document.getElementById('clr'),
      sugg=document.getElementById('sugg'), cntEl=document.getElementById('cnt'),
      cntP=document.getElementById('cntPend'), cntA=document.getElementById('cntApr'),
      chipsEl=document.getElementById('chips'), msg=document.getElementById('msg'),
      whoBox=document.getElementById('whoBox'), whoUser=document.getElementById('whoUser'),
      logoutBtn=document.getElementById('logoutBtn');
  var activeSec='all', activeView='all';
  var codeOf = {}; bullets.forEach(function(b){ codeOf[b.getAttribute('data-code')]=b; });

  /* ---------- login ---------- */
  function sha256Hex(str){
    // synchronous-ish using a tiny sha-256 via SubtleCrypto when https available
    return new Promise(function(resolve){
      if(window.crypto && crypto.subtle){
        crypto.subtle.digest('SHA-256', new TextEncoder().encode(str)).then(function(buf){
          resolve(Array.from(new Uint8Array(buf)).map(function(b){ return ('0'+b.toString(16)).slice(-2); }).join(''));
        }).catch(function(){ resolve(fallbackHash(str)); });
      } else resolve(fallbackHash(str));
    });
  }
  function fallbackHash(s){ // FNV-1a hex (used only if no WebCrypto)
    var h=2166136261; for(var i=0;i<s.length;i++){ h^=s.charCodeAt(i); h=Math.imul(h,16777619); }
    return ('00000000'+(h>>>0).toString(16)).slice(-8);
  }
  var overlay=document.getElementById('loginOverlay'),
      loginId=document.getElementById('loginId'), loginPass=document.getElementById('loginPass'),
      loginErr=document.getElementById('loginErr'), loginBtn=document.getElementById('loginBtn');
  var forgotPanel=document.getElementById('forgotPanel'), forgotEmail=document.getElementById('forgotEmail'),
      forgotErr=document.getElementById('forgotErr'), forgotSend=document.getElementById('forgotSend'),
      forgotBtn=document.getElementById('forgotBtn'),
      requestPanel=document.getElementById('requestPanel'), requestBtn=document.getElementById('requestBtn'),
      reqName=document.getElementById('reqName'), reqEmail=document.getElementById('reqEmail'),
      reqSection=document.getElementById('reqSection'), reqErr=document.getElementById('reqErr'),
      reqSend=document.getElementById('reqSend'),
      loginForm=document.getElementById('loginForm'), loginOk=document.getElementById('loginOk'),
      loginSub=document.getElementById('loginSub');
  // extra users/requests persisted locally (offline only; a real server would share them)
  var EXUSERS = {}; try{ EXUSERS=JSON.parse(localStorage.getItem('alel_extra_users')||'{}'); }catch(e){}
  var REQS = []; try{ REQS=JSON.parse(localStorage.getItem('alel_access_requests')||'[]'); }catch(e){}
  function allUsers(){ return USERS.concat(Object.keys(EXUSERS).map(function(k){ return {id:k,name:EXUSERS[k].name,role:EXUSERS[k].role,section:EXUSERS[k].section,hash:EXUSERS[k].hash}; })); }
  function findUser(em){ var ex=EXUSERS[em]; if(ex) return {id:em,name:ex.name,role:ex.role,section:ex.section,hash:ex.hash}; return USERS.filter(function(x){return x.id===em;})[0]||null; }

  function showLogin(){ overlay.style.display='flex'; document.body.style.overflow='hidden'; }
  function hideLogin(){ overlay.style.display='none'; document.body.style.overflow=''; }
  function showPanel(which){
    loginForm.style.display = which==='login'?'grid':'none';
    forgotPanel.style.display = which==='forgot'?'grid':'none';
    requestPanel.style.display = which==='request'?'grid':'none';
    loginSub.textContent = which==='login' ? 'Sign in with your office email'
                        : which==='forgot' ? 'Reset your password' : 'Request access';
    loginOk.textContent=''; loginErr.textContent='';
    loginOk.className='ok-msg';
  }
  function applySession(){
    if(SESSION){ hideLogin(); whoUser.textContent = SESSION.name + ' \u00b7 ' + SESSION.role; }
    else { whoUser.textContent=''; }
    if(adminBtn){ adminBtn.style.display = (SESSION && SESSION.role==='admin') ? 'inline-flex' : 'none'; }
    if(SESSION && SESSION.role==='admin'){ renderAdminPanel(); }
    refreshWF();
  }
  logoutBtn.addEventListener('click', function(){
    SESSION=null; try{ localStorage.removeItem('alel_login_v1'); }catch(e){}
    loginPass.value=''; loginErr.textContent=''; applySession(); showPanel('login'); showLogin();
  });
  function showOk(msg){ loginOk.textContent=msg; loginOk.className='ok-msg show'; }
  // ---- tab switching ----
  forgotBtn.addEventListener('click', function(){ showPanel('forgot'); forgotEmail.focus(); });
  requestBtn.addEventListener('click', function(){ showPanel('request'); reqName.focus(); });
  document.getElementById('backLogin1').addEventListener('click', function(){ showPanel('login'); });
  document.getElementById('backLogin2').addEventListener('click', function(){ showPanel('login'); });
  // ---- login ----
  loginBtn.addEventListener('click', doLogin);
  [loginId, loginPass].forEach(function(i){ i.addEventListener('keydown', function(e){ if(e.key==='Enter') doLogin(); }); });
  function doLogin(){
    var email=(loginId.value||'').trim().toLowerCase();
    var pw=loginPass.value||'';
    if(!email||!pw){ loginErr.textContent='Enter email and password'; return; }
    sha256Hex(pw).then(function(h){
      var u=findUser(email);
      if(!u){ loginErr.textContent='No account with this email. Use \u201cRequest access\u201d.'; return; }
      if(u.hash!==h){ loginErr.textContent='Wrong password. Use \u201cForgot password?\u201d'; return; }
      SESSION={id:u.id,name:u.name,role:u.role,section:u.section};
      try{ localStorage.setItem('alel_login_v1', JSON.stringify(SESSION)); }catch(e){}
      loginErr.textContent=''; loginId.value=''; loginPass.value='';
      activeView = SESSION.role==='approver' ? 'appr' : 'all';
      setView(activeView);
      applySession();
    });
  }
  // ---- forgot password ----
  forgotSend.addEventListener('click', function(){
    var em=(forgotEmail.value||'').trim().toLowerCase();
    if(!em){ forgotErr.textContent='Enter your email'; return; }
    var u=allUsers().filter(function(x){ return x.id===em; })[0];
    if(!u){ forgotErr.textContent='No account with this email.'; return; }
    forgotErr.textContent='';
    var req={type:'reset', email:em, when:new Date().toLocaleString()};
    REQS.push(req); try{ localStorage.setItem('alel_access_requests', JSON.stringify(REQS)); }catch(e){}
    showOk('Reset request sent to the Administrator.\\n(Offline demo: open this page as Admin \u2192 \u201cAdmin panel\u201d to reset the password.)');
    forgotEmail.value='';
  });
  // ---- request access ----
  reqSend.addEventListener('click', function(){
    var nm=(reqName.value||'').trim(), em=(reqEmail.value||'').trim().toLowerCase(), sec=reqSection.value;
    if(!nm||!em){ reqErr.textContent='Enter your name and email'; return; }
    if(allUsers().filter(function(x){ return x.id===em; }).length){ reqErr.textContent='This email already has an account.'; return; }
    reqErr.textContent='';
    REQS.push({type:'access', name:nm, email:em, section:sec, when:new Date().toLocaleString()});
    try{ localStorage.setItem('alel_access_requests', JSON.stringify(REQS)); }catch(e){}
    showOk('Access request sent to the Administrator.\\n(Offline demo: the Admin can approve it in the \u201cAdmin panel\u201d.)');
    reqName.value=''; reqEmail.value=''; reqSection.value='';
  });
  // ---- admin panel (visible only to admin) ----
  var adminBtn=document.getElementById('adminBtn');
  function renderAdminPanel(){
    if(!adminBtn) return;
    var pending=REQS.filter(function(r){ return !r.done; }).length;
    adminBtn.classList.toggle('has', pending>0);
    adminBtn.querySelector('i').textContent = pending ? pending : '';
  }
  function openAdmin(){
    var dlg=document.getElementById('adminDlg'); if(!dlg) return;
    var body=document.getElementById('adminBody'); body.innerHTML='';
    var h='<h3>Pending requests ('+REQS.filter(function(r){return !r.done;}).length+')</h3>';
    if(!REQS.filter(function(r){return !r.done;}).length){ h+='<p class="empty">No pending requests.</p>'; }
    h+='<ul class="req-list">';
    REQS.forEach(function(r,i){
      h+='<li data-i="'+i+'"><b>'+(r.name||r.email)+'</b> <span class="rq">'+(r.type==='access'?('Access · '+r.section):'Password reset')+'</span> <span class="rw">'+r.when+'</span>'
        +'<div class="rq-act">'
        +(r.type==='access'
          ? '<button class="mini ok" data-act="grant">Approve</button>'
          : '<button class="mini ok" data-act="grantpass">Reset pass</button>')
        +'<button class="mini no" data-act="deny">Deny</button></div></li>';
    });
    h+='</ul><h3>Users</h3><ul class="user-list">';
    allUsers().forEach(function(u){
      h+='<li><b>'+u.name+'</b> <span>'+u.role+(u.section?' · '+u.section:'')+'</span><code>'+u.id+'</code>'
        +'<div class="rq-act"><button class="mini" data-act="resetpass" data-em="'+u.id+'">Reset password</button></div></li>';
    });
    h+='</ul>';
    body.innerHTML=h;
    dlg.style.display='flex';
    body.querySelectorAll('[data-act]').forEach(function(btn){
      btn.addEventListener('click', function(){
        var li=btn.closest('li'); var idx=li?li.getAttribute('data-i'):null; var em=btn.getAttribute('data-em');
        if(btn.getAttribute('data-act')==='grant'||btn.getAttribute('data-act')==='grantpass'){
          var r=REQS[idx];
          if(r.type==='access'){ grantAccess(r.email, r.name, r.section); }
          else { resetUser(r.email); }
          r.done=true; saveReqs(); openAdmin();
        } else if(btn.getAttribute('data-act')==='deny'){
          REQS.splice(idx,1); saveReqs(); openAdmin();
        } else if(btn.getAttribute('data-act')==='resetpass'){
          resetUser(em); openAdmin();
        }
      });
    });
  }
  function saveReqs(){ try{ localStorage.setItem('alel_access_requests', JSON.stringify(REQS)); }catch(e){} }
  function grantAccess(em, nm, sec){
    var role = (sec==='GSS'||sec==='LED'||sec==='HAP') ? 'checker' : 'prepared';
    var pass=makePass();
    sha256Hex(pass).then(function(h){
      EXUSERS[em]={name:nm, role:role, section:sec||'', hash:h};
      try{ localStorage.setItem('alel_extra_users', JSON.stringify(EXUSERS)); }catch(e){}
      alert('Account created for '+em+'.\\nRole: '+role+(sec?' ('+sec+')':'')+'\\nLogin password: '+pass+'\\n(Share this with the user.)');
    });
  }
  function resetUser(em){
    var pass=makePass();
    sha256Hex(pass).then(function(h){
      var u=USERS.filter(function(x){return x.id===em;})[0];
      if(u){ EXUSERS[em]={name:u.name, role:u.role, section:u.section, hash:h}; }
      else if(EXUSERS[em]){ EXUSERS[em].hash=h; }
      try{ localStorage.setItem('alel_extra_users', JSON.stringify(EXUSERS)); }catch(e){}
      alert('Password reset for '+em+'.\\nNew temporary password: '+pass);
    });
  }
  function makePass(){ return 'Al@'+(Math.floor(1000+Math.random()*9000)); }
  function closeAdmin(){ var d=document.getElementById('adminDlg'); if(d) d.style.display='none'; }
  if(adminBtn) adminBtn.addEventListener('click', openAdmin);
  var admClose=document.getElementById('admClose'); if(admClose) admClose.addEventListener('click', closeAdmin);
  // bootstrap
  if(!SESSION){ showLogin(); showPanel('login'); }
  else { applySession(); }

  /* ---------- role helpers ---------- */
  function canSee(b){
    if(!SESSION) return false;
    if(SESSION.role==='admin'||SESSION.role==='prepared') return true;
    if(SESSION.role==='approver') return true;            // sees all, but filters to pending in view
    if(SESSION.role==='checker'){
      return b.getAttribute('data-sec').toUpperCase()===SESSION.section.toUpperCase();
    }
    return false;
  }
  function isMyTurn(b, stage){ // stage 'p'|'c'|'a'
    if(!SESSION) return false;
    if(SESSION.role==='admin') return true;
    if(stage==='p') return SESSION.role==='prepared';
    if(stage==='c') return SESSION.role==='checker' && b.getAttribute('data-sec').toUpperCase()===SESSION.section.toUpperCase();
    if(stage==='a') return SESSION.role==='approver';
    return false;
  }

  /* ---------- search suggestions ---------- */
  var cur=0;
  qEl.addEventListener('input', function(){ apply(); buildSugg(); });
  qEl.addEventListener('focus', buildSugg);
  qEl.addEventListener('keydown', function(e){
    var items=sugg.querySelectorAll('.sg');
    if(!sugg.style.display||sugg.style.display==='none') return;
    if(e.key==='ArrowDown'&&items.length){ e.preventDefault(); cur=(cur+1)%items.length; hilite(items); }
    else if(e.key==='ArrowUp'&&items.length){ e.preventDefault(); cur=(cur-1+items.length)%items.length; hilite(items); }
    else if(e.key==='Enter'&&items.length){ e.preventDefault(); items[cur].click(); }
  });
  function hilite(items){ items.forEach(function(it,i){ it.style.background=i===cur?'rgba(255,255,255,.09)':''; }); }
  function buildSugg(){
    var t=(qEl.value||'').trim().toLowerCase();
    if(!t){ sugg.style.display='none'; return; }
    var hits=CATALOG.filter(function(c){ return c.n.toLowerCase().indexOf(t)>-1 || c.c.indexOf(t)>-1 || c.s.toLowerCase().indexOf(t)>-1; }).slice(0,8);
    if(!hits.length){ sugg.style.display='none'; return; }
    sugg.innerHTML='';
    cur=0;
    hits.forEach(function(c){
      var d=document.createElement('div'); d.className='sg';
      d.innerHTML='<b>'+c.n.replace(/</g,'&lt;')+'</b><span class="sc">'+c.c+'</span><span class="ss">'+c.s+'</span>';
      d.addEventListener('mousedown', function(ev){ ev.preventDefault(); qEl.value=c.n; sugg.style.display='none';
        var b=codeOf[c.c]; if(b){ activeSec='all'; syncChips(); setView('all'); apply(); b.scrollIntoView({behavior:'smooth',block:'start'}); b.style.animation='hl 1.2s ease'; setTimeout(function(){b.style.animation='';},1300);} });
      sugg.appendChild(d);
    });
    hilite(sugg.querySelectorAll('.sg'));
    sugg.style.display='block';
  }
  document.addEventListener('click', function(e){ if(!e.target.closest('.tb-row2')) sugg.style.display='none'; });
  clr.addEventListener('click', function(){ qEl.value=''; apply(); sugg.style.display='none'; qEl.focus(); });

  /* ---------- filters ---------- */
  function setView(v){ activeView=v;
    document.querySelectorAll('#viewFld button').forEach(function(b){ b.classList.toggle('on', b.getAttribute('data-v')===v); });
    apply(); }
  document.getElementById('viewFld').addEventListener('click', function(e){ var b=e.target.closest('button'); if(b) setView(b.getAttribute('data-v')); });
  function syncChips(){ chipsEl.querySelectorAll('.chip').forEach(function(c){ c.classList.toggle('on', c.getAttribute('data-sec')===activeSec); }); }
  chipsEl.addEventListener('click', function(e){
    var c=e.target.closest('.chip'); if(!c)return;
    activeSec=c.getAttribute('data-sec'); syncChips(); apply();
  });

  // default: prepared (Anoy) is auto-signed for every bulletin -> p starts 1
  function stOf(b){
    var code=b.getAttribute('data-code');
    var st=STORE[code]||{};
    return {code:code,
            p:(st.p===undefined?1:st.p),
            c:(st.c===undefined?0:st.c),
            a:(st.a===undefined?0:st.a)};
  }
  function setState(code,field,val){
    var st=STORE[code]||{};
    st[field]=val===undefined?1:val; STORE[code]=st;
    try{ localStorage.setItem('alel_approvals_v1', JSON.stringify(STORE)); }catch(e){}
  }

  function apply(){
    var t=(qEl.value||'').trim().toLowerCase();
    var n=0, np=0, na=0;
    bullets.forEach(function(b){
      var show = SESSION ? canSee(b) : false;
      if(show && activeSec!=='all' && b.getAttribute('data-sec')!==activeSec) show=false;
      var s=stOf(b);
      if(show && activeView==='pend' && s.a===1) show=false;
      if(show && activeView==='appr' && !(s.a===1)) show=false;
      if(show && t){
        var nm=b.getAttribute('data-name')||'', cd=b.getAttribute('data-code')||'',
            sc=b.getAttribute('data-sec')||'', sh=b.getAttribute('data-sheet')||'';
        show = nm.indexOf(t)>-1||cd.indexOf(t)>-1||sc.indexOf(t)>-1||sh.indexOf(t)>-1;
      }
      b.classList.toggle('hidden',!show);
      if(show){ n++; if(s.a===0) np++; else na++; }
    });
    cntEl.textContent=n; cntP.textContent=np; cntA.textContent=na;
    msg.style.display=n?'none':'block';
    clr.style.display=qEl.value?'block':'none';
    buildSugg();
  }

  /* ---------- operation cards (rendered from data-ops JSON) ---------- */
  function escH(s){ return String(s==null?'':s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;'); }
  function buildOps(){
    bullets.forEach(function(b){
      var box=b.querySelector('.oplist'); if(!box || box.childElementCount) return;
      var raw=b.getAttribute('data-ops'); if(!raw) return;
      var ops=[]; try{ ops=JSON.parse(raw); }catch(e){ return; }
      ops.forEach(function(o){
        var card=document.createElement('div'); card.className='opc';
        card.innerHTML =
          '<div class="opc-no">'+escH(o[0])+'</div>'+
          '<div class="opc-main">'+
            '<div class="opc-name">'+escH(o[1])+'</div>'+
            '<div class="opc-meta">'+
              '<span>Machine <b>'+escH(o[2]||'&ndash;')+'</b></span>'+
              '<span>Pcs/Cycle <b>'+escH(o[4])+'</b></span>'+
              '<span>MP <b>'+escH(o[5])+'</b></span>'+
              '<span>Cycle <b>'+escH(o[11])+' min</b></span>'+
              '<span>Rating <b>'+escH(o[12])+'</b></span>'+
            '</div>'+
            '<div class="opc-detail">'+
              '<div class="od"><span>Pcs/Unit</span><b>'+escH(o[3])+'</b></div>'+
              '<div class="od"><span>T1</span><b>'+escH(o[6])+'</b></div>'+
              '<div class="od"><span>T2</span><b>'+escH(o[7])+'</b></div>'+
              '<div class="od"><span>T3</span><b>'+escH(o[8])+'</b></div>'+
              '<div class="od"><span>T4</span><b>'+escH(o[9])+'</b></div>'+
              '<div class="od"><span>T5</span><b>'+escH(o[10])+'</b></div>'+
              '<div class="od"><span>Basic Time</span><b>'+escH(o[13])+'</b></div>'+
              '<div class="od"><span>Allowance</span><b>'+escH(o[14])+'</b></div>'+
              '<div class="od"><span>Hourly Cap.</span><b>'+escH(o[16])+'</b></div>'+
            '</div>'+
          '</div>'+
          '<div class="opc-right"><div class="opc-smv">'+escH(o[15])+'<small>SMV min</small></div></div>'+
          '<svg class="caret" viewBox="0 0 24 24" width="15" height="15" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round"><polyline points="6 9 12 15 18 9"/></svg>';
        card.addEventListener('click', function(){ card.classList.toggle('open'); });
        box.appendChild(card);
      });
    });
  }
  buildOps();

  /* ---------- workflow buttons ---------- */
  function refreshWF(){
    bullets.forEach(function(b){
      var s=stOf(b);
      var sec=(b.getAttribute('data-sec')||'').toUpperCase();
      var ck=({GSS:'Jhumour Rani',LED:'Kajal Kanti',HAP:'Abdullah Al-Mamun'})[sec]||'';
      var stP=b.querySelector('.wf-step[data-role=prepared]');
      var stC=b.querySelector('.wf-step[data-role=checked]');
      var stA=b.querySelector('.wf-step[data-role=approved]');
      // remove old buttons + hints
      [stP,stC,stA].forEach(function(st){ if(!st)return; var h=st.querySelector('.wf-hint'); if(h)h.remove(); var ab=st.querySelector('.wf-btn'); if(ab)ab.remove(); });
      // ---- prepared step ----
      stP.classList.toggle('st-done', !!s.p);
      var pillP=stP.querySelector('[data-role=pill]');
      pillP.textContent = s.p?'Signed':'Pending';
      pillP.className='st-pill';
      // ---- checked step ----
      var pillC=stC.querySelector('[data-role=pill]');
      stC.classList.remove('st-done','st-active');
      if(s.c){ stC.classList.add('st-done'); pillC.textContent='Checked'; }
      else { stC.classList.add('st-active'); pillC.textContent=s.p?'Pending Check':'Awaiting Prep'; }
      // ---- approved step ----
      var pillA=stA.querySelector('[data-role=pill]');
      stA.classList.remove('st-done','st-active');
      if(s.a){ stA.classList.add('st-done'); pillA.textContent='Approved'; }
      else { stA.classList.add('st-active'); pillA.textContent=s.c?'Pending Approve':'Awaiting Check'; }

      // header badge
      var ap=b.querySelector('.aprv');
      ap.className='aprv '+(s.a?'a-done':s.c?'a-pend':s.p?'a-prep':'');
      ap.textContent=s.a?'Approved':s.c?'With Approver':s.p?'With Checker':'Preparing';

      // ---- action buttons gated by logged-in role ----
      if(!SESSION){ return; }
      if(s.a){
        if(isMyTurn(b,'a')){ addBtn(b, stA, function(){ setState(s.code,'a',0); refreshWF(); }, 'Undo Approval', 'ghost'); }
        return;
      }
      if(!s.p){
        if(isMyTurn(b,'p')){ addBtn(b, stP, function(){ setState(s.code,'p',1); refreshWF(); }, 'Sign as Prepared', 'ck'); }
        return;
      }
      if(!s.c){
        if(isMyTurn(b,'c') && ck){ addBtn(b, stC, function(){ setState(s.code,'c',1); refreshWF(); }, 'Check & Send to Approve', 'ck'); }
        else { addHint(stC, nextWho(ck)); }
        return;
      }
      if(isMyTurn(b,'a')){ addBtn(b, stA, function(){ setState(s.code,'a',1); refreshWF(); }, 'Approve Final', 'ap'); }
      else { addHint(stA, nextWho('Md. Moshfequr Rahman')); }
    });
    apply();
  }
  function nextWho(name){
    var el=document.createElement('div'); el.className='wf-hint';
    el.textContent='Waiting for '+name;
    return el;
  }
  function addHint(step,el){ if(step) step.appendChild(el); }
  function addBtn(b,step,cb,label,cls){
    var btn=document.createElement('button');
    btn.className='wf-btn '+(cls||''); btn.textContent=label;
    btn.addEventListener('click',function(){ cb(); });
    step.appendChild(btn);
  }
  function clearOne(b,field){
    var s=stOf(b); setState(s.code,field,0); refreshWF();
  }
  refreshWF();

  /* ---------- manpower live ---------- */
  function fmtNum(x,d){ if(x===null||x===undefined||isNaN(x)) return '&ndash;';
    return x.toLocaleString('en-US',{minimumFractionDigits:d,maximumFractionDigits:d}); }
  function recalc(kpi){
    var b=kpi.closest('.bulletin'); var smv=parseFloat(b.getAttribute('data-smv'));
    var inp=kpi.querySelector('.mp-input'); var mp=parseFloat(inp.value);
    if(isNaN(mp)||mp<1) mp=1; inp.value=mp;
    if(!smv||smv<=0){ inp.value=inp.getAttribute('data-base'); return; }
    b.setAttribute('data-mp',mp);
    var prod=60/smv, hc=mp*prod;
    b.querySelector('.hc-val').textContent=fmtNum(hc,0);
    b.querySelector('.prod-val').textContent=fmtNum(prod,2);
  }
  function bindMP(kpi){
    var inp=kpi.querySelector('.mp-input');
    inp.addEventListener('change',function(){recalc(kpi);});
    inp.addEventListener('keyup',function(){if(inp.value)recalc(kpi);});
    kpi.querySelectorAll('.mp-btn').forEach(function(btn){ btn.addEventListener('click',function(){
      var v=parseFloat(inp.value); if(isNaN(v))v=1; v+=parseInt(btn.getAttribute('data-adj'),10); if(v<1)v=1;
      inp.value=v; recalc(kpi); }); });
    var rs=kpi.querySelector('.mp-reset'); if(rs) rs.addEventListener('click',function(){inp.value=inp.getAttribute('data-base'); recalc(kpi);});
  }
  document.querySelectorAll('.k-mp').forEach(bindMP);

  /* ---------- scroll hide + to-top ---------- */
  var lastY=window.pageYOffset||0, toTop=document.getElementById('toTop');
  window.addEventListener('scroll',function(){
    var y=window.pageYOffset||0;
    if(y>90 && y>lastY && window.innerWidth<=820) document.body.classList.add('hide-top');
    else if(y<lastY||y<10) document.body.classList.remove('hide-top');
    toTop.classList.toggle('show', y>500);
    lastY=y;
  },{passive:true});
  toTop.addEventListener('click',function(){window.scrollTo({top:0,behavior:'smooth'});});

  /* auto 1m */
  var autoT=null, autoBtn=document.getElementById('autoBtn'), autoLab=document.getElementById('autoLab');
  function setAuto(on){ if(autoT){clearInterval(autoT);autoT=null;} autoBtn.classList.toggle('on',on);
    autoLab.textContent=on?'ON':'Auto 1m';
    if(on) autoT=setInterval(function(){location.reload();},60000);
    try{localStorage.setItem('alel_auto',on?'1':'0');}catch(e){} }
  autoBtn.addEventListener('click',function(){setAuto(!autoBtn.classList.contains('on'));});
  document.getElementById('refreshBtn').addEventListener('click',function(){location.reload();});
  var sa=''; try{sa=localStorage.getItem('alel_auto')||'';}catch(e){}
  if(sa==='1') setAuto(true);
})();
</script>
</body>
</html>"""

if __name__ == "__main__":
    main()
