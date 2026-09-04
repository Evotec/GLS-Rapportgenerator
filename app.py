# GLS Evotec Distance by hours rapportgenerator v37
# Pure Python / standard library only. No pip, no openpyxl required.

from http.server import ThreadingHTTPServer, BaseHTTPRequestHandler
from urllib.parse import urlparse, parse_qs, urlencode
from urllib.request import Request, urlopen
import csv
import io
import json
import os
import re
import traceback
import webbrowser
import zipfile
import xml.etree.ElementTree as ET

APP_TITLE = "GLS Rapportgenerator"
VERSION = "v37"
PORT = 8765

LAST_ROWS = []
LAST_HEADERS = []
LAST_LOG = []

NS = {
    "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "rel": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}

HTML = r'''<!doctype html>
<html lang="da">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>GLS Rapportgenerator</title>
<style>
:root{--bg:#f4f7fb;--card:#fff;--ink:#14213d;--muted:#5d6b82;--line:#d9e2ef;--blue:#2563eb;--dark:#0f172a}
*{box-sizing:border-box}body{margin:0;background:var(--bg);font-family:Arial,Helvetica,sans-serif;color:var(--ink)}
header{background:var(--dark);color:#fff;padding:16px 24px;font-size:22px;font-weight:700}
.wrap{padding:22px}.card{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:18px;margin-bottom:18px;box-shadow:0 1px 2px rgba(15,23,42,.04)}
.grid{display:grid;grid-template-columns:185px 1fr;gap:12px 16px;align-items:center;max-width:980px}
label{font-weight:700}input,select,textarea{width:100%;padding:10px;border:1px solid #bfd0e5;border-radius:8px;font-size:14px;background:white}
button{border:0;border-radius:8px;padding:11px 16px;font-weight:700;cursor:pointer;margin-right:8px}.primary{background:var(--blue);color:white}.secondary{background:#334155;color:white}.ghost{background:#e2e8f0;color:#0f172a}
.help{color:var(--muted);font-size:13px;margin-top:12px}.warn{background:#fff7ed;border:1px solid #fdba74;color:#7c2d12;padding:10px;border-radius:8px;margin-bottom:16px}
#log{background:#020617;color:#d1e7ff;font-family:Consolas,monospace;white-space:pre-wrap;max-height:270px;overflow:auto;padding:14px;border-radius:10px;font-size:13px}
.pills{display:flex;gap:10px;flex-wrap:wrap;margin-bottom:10px}.pill{background:#e0e7ff;color:#3730a3;border-radius:20px;padding:7px 12px;font-size:13px}
.tablewrap{overflow:auto;max-height:520px;border:1px solid var(--line);border-radius:10px;background:white}
table{border-collapse:collapse;width:max-content;min-width:100%}th,td{border:1px solid #e5e7eb;padding:6px 8px;font-size:12px;white-space:nowrap}th{background:#f1f5f9;position:sticky;top:0;z-index:2}
details{max-width:980px;margin-top:10px}summary{cursor:pointer;font-weight:700;color:#334155}
</style>
</head>
<body>
<header>GLS Rapportgenerator <small style="font-size:14px;color:#cbd5e1">v37</small></header>
<div class="wrap">
  <div class="card">
    <div class="warn"><b>Anbefalet workflow:</b> Eksportér Evotecs <b>Distance by hours</b> Excel og brug den som kilde. Programmet bevarer distance-tallene og tilføjer kun metadata fra Evotec GPS API. CSV downloades med dansk Excel-format: semikolon og komma som decimaltegn.</div>
    <form id="form">
      <div class="grid">
        <label>Evotec Excel-rapport</label><input type="file" name="xlsx" id="xlsx" accept=".xlsx" required>
        <label>Evotec API-nøgle</label><input type="password" name="api_key" id="api_key" placeholder="Indsæt API-nøgle">
      </div>
      <details>
        <summary>Avanceret / teknikerindstillinger</summary>
        <div class="grid" style="margin-top:12px">
          <label>Evotec GPS adresse</label><input name="base_url" id="base_url" value="https://gps.evotec.dk">
          <label>Auth mode</label><select name="auth_mode" id="auth_mode"><option value="query">Query parameter</option><option value="bearer">Bearer token</option></select>
          <label>API-key parameter</label><input name="api_param" id="api_param" value="key">
          <label>Group lookup mode</label><select name="group_mode" id="group_mode"><option value="smart" selected>Smart - sikker group-scan</option><option value="none">Ingen group-scan</option></select>
          <label>Max group-kald</label><input name="max_group_calls" id="max_group_calls" value="300">
          <label>API timeout sek.</label><input name="timeout" id="timeout" value="4">
          <label>Manuel metadata-mapping</label><textarea name="manual" id="manual" rows="4" placeholder="Én pr. linje, fx:&#10;DX92718=DK0026-København / Bojon Logistics Service | DK0026-København | Bojon Logistics Service"></textarea>
          <label></label><button type="button" class="ghost" onclick="clearKey()">Slet gemt API-nøgle</button>
        </div>
      </details>
      <div style="margin-top:16px">
        <button type="submit" class="primary">Lav rapport</button>
        <button type="button" id="download" class="secondary" disabled onclick="location.href='/download.csv'">Download CSV</button>
        <button type="button" class="ghost" onclick="clearLog()">Ryd log</button>
      </div>
      <div class="help">API-nøglen gemmes kun lokalt i browseren på denne PC. Programmet kræver ikke pip/openpyxl.</div>
    </form>
  </div>
  <div class="card"><div id="log">Klar.</div></div>
  <div class="card" id="resultCard" style="display:none">
    <div class="pills"><span class="pill" id="rowsPill"></span><span class="pill" id="metaPill"></span><span class="pill" id="kmPill"></span></div>
    <div class="tablewrap"><table id="tbl"></table></div>
  </div>
</div>
<script>
const STORAGE_PREFIX='gls_evotec_';
const ids=['api_key','base_url','auth_mode','api_param','group_mode','max_group_calls','timeout','manual'];
for(const id of ids){
  let v=localStorage.getItem(STORAGE_PREFIX+id);
  if(v===null && id==='api_key') v=localStorage.getItem('gls_api_key');
  if(v!==null) document.getElementById(id).value=v;
}
if(!localStorage.getItem(STORAGE_PREFIX+'base_url')) document.getElementById('base_url').value='https://gps.evotec.dk';
document.getElementById('api_key').addEventListener('change',()=>localStorage.setItem(STORAGE_PREFIX+'api_key',document.getElementById('api_key').value));
for(const id of ids.filter(x=>x!='api_key')) document.getElementById(id).addEventListener('change',()=>localStorage.setItem(STORAGE_PREFIX+id,document.getElementById(id).value));
function clearKey(){localStorage.removeItem(STORAGE_PREFIX+'api_key');localStorage.removeItem('gls_api_key');document.getElementById('api_key').value='';}
function clearLog(){document.getElementById('log').textContent='Klar.';}
document.getElementById('form').addEventListener('submit',async(e)=>{
  e.preventDefault();
  document.getElementById('download').disabled=true;
  document.getElementById('resultCard').style.display='none';
  document.getElementById('log').textContent='Læser Excel og henter Evotec metadata...';
  const fd=new FormData(e.target);
  try{
    const res=await fetch('/enrich',{method:'POST',body:fd});
    const data=await res.json();
    document.getElementById('log').textContent=data.log||'';
    if(!data.ok){alert('Fejl: '+(data.error||'Ukendt fejl'));return;}
    render(data.headers,data.rows,data.stats);
    document.getElementById('download').disabled=false;
  }catch(err){document.getElementById('log').textContent='FEJL: '+err;alert(err);}
});
function render(headers,rows,stats){
  document.getElementById('resultCard').style.display='block';
  document.getElementById('rowsPill').textContent='Preview-rækker: '+rows.length;
  document.getElementById('metaPill').textContent='Matchede metadata: '+stats.matched+'/'+stats.source_rows;
  document.getElementById('kmPill').textContent='Total km: '+stats.total_km;
  let html='<thead><tr>'+headers.map(h=>'<th>'+esc(h)+'</th>').join('')+'</tr></thead><tbody>';
  for(const r of rows.slice(0,500)) html+='<tr>'+headers.map(h=>'<td>'+esc(r[h]??'')+'</td>').join('')+'</tr>';
  html+='</tbody>';document.getElementById('tbl').innerHTML=html;
}
function esc(s){return String(s).replace(/[&<>"']/g,m=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[m]));}
</script>
</body>
</html>'''

def log(msg):
    LAST_LOG.append(str(msg))

def clean(s):
    if s is None:
        return ""
    if isinstance(s, float) and s.is_integer():
        return str(int(s))
    return str(s).strip()

def norm_plate(s):
    return re.sub(r"[^A-Z0-9]", "", clean(s).upper())

def is_numeric_text(s):
    return bool(re.fullmatch(r"-?\d+(?:[.,]\d+)?", clean(s)))

def csv_value(value):
    """Format numeric values for Danish Excel CSV.

    Excel XML gives decimals with dot, e.g. 77.300000000000.
    Danish Excel needs comma decimals in semicolon separated CSV.
    """
    s = clean(value)
    if not s:
        return ""
    # Do not convert id-ish values with leading zeroes
    if re.fullmatch(r"0\d+", s):
        return s
    if is_numeric_text(s):
        try:
            n = float(s.replace(",", "."))
            out = f"{n:.10f}".rstrip("0").rstrip(".")
            return out.replace(".", ",")
        except Exception:
            return s
    return s

def xlsx_shared_strings(z):
    out = []
    try:
        root = ET.fromstring(z.read("xl/sharedStrings.xml"))
    except KeyError:
        return out
    for si in root.findall("main:si", NS):
        out.append("".join(t.text or "" for t in si.iter("{%s}t" % NS["main"])))
    return out

def col_to_idx(cell_ref):
    letters = "".join(ch for ch in cell_ref if ch.isalpha())
    n = 0
    for ch in letters:
        n = n * 26 + ord(ch.upper()) - 64
    return n - 1

def read_xlsx_rows(data):
    with zipfile.ZipFile(io.BytesIO(data)) as z:
        shared = xlsx_shared_strings(z)
        sheet_name = "xl/worksheets/sheet1.xml"
        root = ET.fromstring(z.read(sheet_name))
        rows = []
        for row in root.findall(".//main:sheetData/main:row", NS):
            vals = []
            for c in row.findall("main:c", NS):
                idx = col_to_idx(c.attrib.get("r", "A1"))
                while len(vals) <= idx:
                    vals.append("")
                typ = c.attrib.get("t")
                v = c.find("main:v", NS)
                is_el = c.find("main:is", NS)
                value = ""
                if typ == "s" and v is not None:
                    try:
                        value = shared[int(v.text)]
                    except Exception:
                        value = v.text or ""
                elif typ == "inlineStr" and is_el is not None:
                    value = "".join(t.text or "" for t in is_el.iter("{%s}t" % NS["main"]))
                elif v is not None:
                    value = v.text or ""
                vals[idx] = clean(value)
            rows.append(vals)
        return rows

def detect_table(rows):
    header_i = None
    for i, r in enumerate(rows[:50]):
        joined = "|".join(clean(x).lower() for x in r)
        if "licence plate" in joined or "license plate" in joined:
            header_i = i
            break
    if header_i is None:
        raise ValueError("Kunne ikke finde header-række med Licence plate i Excel-filen.")
    headers = [clean(x) for x in rows[header_i]]
    while headers and headers[-1] == "":
        headers.pop()

    data_rows = []
    total_row = None
    for r in rows[header_i + 1:]:
        r = (r + [""] * len(headers))[:len(headers)]
        first = clean(r[0])
        if not any(clean(x) for x in r):
            continue
        if first.lower() == "total":
            total_row = r
            break
        data_rows.append(r)
    return headers, data_rows, total_row, header_i + 1

def api_get(base_url, endpoint, api_key, auth_mode="query", api_param="key", timeout=4, params=None):
    params = params or {}
    url = base_url.rstrip("/") + "/api/v1/" + endpoint.lstrip("/")
    headers = {"User-Agent": "GLS-Rapportgenerator-v37", "Accept": "application/json"}
    q = dict(params)
    if auth_mode == "bearer":
        headers["Authorization"] = "Bearer " + api_key
    else:
        q[api_param or "key"] = api_key
    if q:
        url += "?" + urlencode(q, doseq=True)
    req = Request(url, headers=headers)
    with urlopen(req, timeout=timeout) as resp:
        raw = resp.read()
    return json.loads(raw.decode("utf-8", "replace"))

def walk(obj):
    if isinstance(obj, dict):
        yield obj
        for v in obj.values():
            yield from walk(v)
    elif isinstance(obj, list):
        for x in obj:
            yield from walk(x)

def find_list(obj):
    if isinstance(obj, list):
        return obj
    if isinstance(obj, dict):
        for k in ("data", "units", "items", "result", "results", "list", "groups"):
            v = obj.get(k)
            if isinstance(v, list):
                return v
            if isinstance(v, dict):
                inner = find_list(v)
                if inner:
                    return inner
        vals = list(obj.values())
        if vals and all(isinstance(x, dict) for x in vals):
            return vals
        for v in obj.values():
            if isinstance(v, (list, dict)):
                inner = find_list(v)
                if inner:
                    return inner
    return []

def first_value(d, *names):
    if not isinstance(d, dict):
        return ""
    for n in names:
        if d.get(n) not in (None, ""):
            return d.get(n)
    return ""

def unit_fields(u):
    plate = first_value(u, "number", "numberplate", "number_plate", "license_plate", "licence_plate", "plate", "nr_plate", "registration_number", "reg_number", "label")
    uid = first_value(u, "id", "unit_id", "unitId")
    fuel = first_value(u, "fuel_type", "fuelType", "fuel", "engine_type", "engineType")
    client = first_value(u, "client", "client_id", "clientId", "client_name", "customer", "customer_name", "owner", "company")
    title = first_value(u, "title", "vehicle_title", "name", "label")
    make = first_value(u, "make", "brand")
    model = first_value(u, "model")
    make_model = (clean(make) + " " + clean(model)).strip() if (make or model) else clean(first_value(u, "make_model", "makeModel"))

    keys = set()
    for k in ("id", "unit_id", "unitId", "object_id", "objectId", "device_id", "deviceId", "imei"):
        v = first_value(u, k)
        if v not in (None, ""):
            keys.add(clean(v))
    if plate:
        keys.add(norm_plate(plate))

    return {
        "plate": norm_plate(plate),
        "unit_id": clean(uid),
        "fuel": clean(fuel),
        "client": clean(client),
        "title": clean(title),
        "make_model": make_model,
        "keys": {k for k in keys if k},
        "raw": u,
    }

def load_units(config):
    log("Henter units fra Evotec GPS API...")
    obj = api_get(config["base_url"], "unit/list.json", config["api_key"], config["auth_mode"], config["api_param"], config["timeout"])
    arr = find_list(obj)
    by_plate = {}
    by_key = {}
    for u in arr:
        if not isinstance(u, dict):
            continue
        f = unit_fields(u)
        if f["plate"]:
            by_plate[f["plate"]] = f
        for k in f["keys"]:
            by_key[k] = f
    log(f"Units: {len(by_plate)} nummerplader / {len(arr)} unit-objekter")
    return by_plate, by_key

def group_name(g):
    return clean(first_value(g, "name", "title", "label", "group_name", "groupName"))

def group_id(g):
    return clean(first_value(g, "id", "group_id", "groupId"))

def keys_from_unit_obj(o):
    if not isinstance(o, dict):
        return set()
    keys = set()
    for k in ("id", "unit_id", "unitId", "object_id", "objectId", "device_id", "deviceId", "imei", "vehicle_id", "vehicleId", "car_id", "carId"):
        v = first_value(o, k)
        if v not in (None, ""):
            keys.add(clean(v))
    for k in ("number", "numberplate", "number_plate", "license_plate", "licence_plate", "plate", "nr_plate", "registration_number", "reg_number"):
        v = first_value(o, k)
        if v not in (None, ""):
            keys.add(norm_plate(v))
    for k in ("unit", "vehicle", "object", "device"):
        if isinstance(o.get(k), dict):
            keys |= keys_from_unit_obj(o[k])
    return {x for x in keys if x}

def load_groups(config, target_keys, by_key):
    if config.get("group_mode") == "none":
        return {}
    log("Henter grupper/depot/client fra Evotec GPS API...")
    try:
        obj = api_get(config["base_url"], "unit_groups/list.json", config["api_key"], config["auth_mode"], config["api_param"], config["timeout"])
        groups = find_list(obj)
    except Exception as e:
        log("Kunne ikke hente grupper: " + str(e))
        return {}

    log(f"Grupper hentet: {len(groups)}")
    unit_to_groups = {}
    max_calls = int(config.get("max_group_calls") or 300)
    calls = relations = skipped = 0

    for g in groups[:max_calls]:
        gid = group_id(g)
        gname = group_name(g)
        if not gid or not gname:
            continue
        calls += 1
        try:
            res = api_get(config["base_url"], "unit_groups/list_units.json", config["api_key"], config["auth_mode"], config["api_param"], config["timeout"], {"id": gid})
        except Exception as e:
            if calls <= 3:
                log(f"list_units?id={gid} fejl: {e}")
            continue

        items = find_list(res)
        keys = set()
        for item in items:
            if isinstance(item, (str, int, float)):
                keys.add(clean(item))
            elif isinstance(item, dict):
                keys |= keys_from_unit_obj(item)
                for kk in ("unit_ids", "unitIds", "units_ids"):
                    if isinstance(item.get(kk), list):
                        keys |= {clean(x) for x in item[kk] if clean(x)}

        if not keys:
            for d in walk(res):
                keys |= keys_from_unit_obj(d)
                for kk in ("unit_ids", "unitIds", "units_ids"):
                    if isinstance(d.get(kk), list):
                        keys |= {clean(x) for x in d[kk] if clean(x)}

        target_hits = len(keys & target_keys)
        if target_hits > max(400, len(target_keys) // 2):
            skipped += 1
            continue

        for k in keys:
            if k in by_key or k in target_keys:
                unit_to_groups.setdefault(k, set()).add(gname)
                relations += 1

        if calls in (1, 2, 3, 25, 50, 75, 100, 125, 150, 175, 200, 225, 250, 274, 300):
            log(f"Group-scan status: {calls}/{min(max_calls, len(groups))}, relationer={relations}, senest={gname}, units={len(keys)}, skipped={skipped}")

    log(f"Units/group-keys med grupper: {len(unit_to_groups)}; relationer={relations}; brede svar sprunget over={skipped}")
    return unit_to_groups

def leaf_groups(groups):
    gs = sorted({clean(g) for g in groups if clean(g)})
    leaves = []
    for g in gs:
        if not any(other.startswith(g + " / ") for other in gs):
            leaves.append(g)
    return leaves

def split_depot_client(groups):
    leaves = leaf_groups(groups)
    depots = []
    clients = []
    for g in leaves:
        parts = [p.strip() for p in g.split("/")]
        if parts and parts[0] and parts[0] not in depots:
            depots.append(parts[0])
        if len(parts) > 1 and parts[-1] and parts[-1] not in clients:
            clients.append(parts[-1])
    return ", ".join(leaves), ", ".join(depots), ", ".join(clients)

def parse_manual(text):
    out = {}
    for line in (text or "").splitlines():
        line = line.strip()
        if not line or "=" not in line:
            continue
        k, v = line.split("=", 1)
        parts = [p.strip() for p in v.split("|")]
        out[norm_plate(k)] = {
            "groups": parts[0] if len(parts) > 0 else "",
            "depot": parts[1] if len(parts) > 1 else "",
            "client": parts[2] if len(parts) > 2 else "",
        }
    return out

def enrich(data, config):
    LAST_LOG.clear()
    headers, raw_rows, total_row, header_no = detect_table(read_xlsx_rows(data))
    log(f"Læser Excel: header række {header_no}; kolonner: {len(headers)}; data-rækker: {len(raw_rows)}")
    lower = [h.lower().strip() for h in headers]

    def idx(names):
        for n in names:
            if n.lower() in lower:
                return lower.index(n.lower())
        return -1

    plate_i = idx(["Licence plate", "License plate"])
    make_i = idx(["Make / Model", "Make /Model", "Make/Model"])
    title_i = idx(["Vehicle title"])
    if plate_i < 0:
        raise ValueError("Licence plate-kolonne ikke fundet.")

    target_plates = {norm_plate(r[plate_i]) for r in raw_rows if plate_i < len(r) and norm_plate(r[plate_i])}
    by_plate, by_key = load_units(config)

    target_keys = set(target_plates)
    for p in target_plates:
        if p in by_plate:
            target_keys |= by_plate[p]["keys"]

    unit_groups = load_groups(config, target_keys, by_key)
    manual = parse_manual(config.get("manual", ""))

    meta_headers = ["Groups", "Depot", "Client", "Fuel type", "Unit ID"]
    existing_meta = {h.lower() for h in meta_headers}
    base_headers = [h for h in headers if h.lower() not in existing_meta]

    lower_base = [h.lower().strip() for h in base_headers]
    def base_idx(names):
        for n in names:
            if n.lower() in lower_base:
                return lower_base.index(n.lower())
        return -1

    insert_after = max([x for x in [base_idx(["Vehicle title"]), base_idx(["Make / Model", "Make /Model", "Make/Model"]), base_idx(["Licence plate", "License plate"])] if x >= 0])
    new_headers = base_headers[:insert_after + 1] + meta_headers + base_headers[insert_after + 1:]

    out = []
    matched = groups_api = groups_empty = client_count = 0
    for r in raw_rows:
        rowdict = {headers[i]: clean(r[i]) if i < len(r) else "" for i in range(len(headers))}
        plate = norm_plate(r[plate_i])
        u = by_plate.get(plate, {})
        groups = []
        if u:
            matched += 1
            for k in u.get("keys", []):
                groups.extend(list(unit_groups.get(k, [])))

        gtxt, depot, client_from_group = split_depot_client(groups)
        if plate in manual:
            m = manual[plate]
            gtxt = m.get("groups") or gtxt
            depot = m.get("depot") or depot
            client_from_group = m.get("client") or client_from_group

        client = client_from_group or (u.get("client", "") if u else "")
        fuel = u.get("fuel", "") if u else ""
        uid = u.get("unit_id", "") if u else ""

        if gtxt:
            groups_api += 1
        else:
            groups_empty += 1
        if client:
            client_count += 1

        meta = {"Groups": gtxt, "Depot": depot, "Client": client, "Fuel type": fuel, "Unit ID": uid}
        newrow = {}
        for h in base_headers[:insert_after + 1]:
            newrow[h] = rowdict.get(h, "")
        for h in meta_headers:
            newrow[h] = meta[h]
        for h in base_headers[insert_after + 1:]:
            newrow[h] = rowdict.get(h, "")
        out.append(newrow)

    total_km = ""
    total_col = next((h for h in base_headers if "total" in h.lower() and "km" in h.lower()), None)
    if total_col:
        s = 0.0
        for r in out:
            try:
                s += float(clean(r.get(total_col, "")).replace(",", "."))
            except Exception:
                pass
        total_km = str(round(s, 1)).replace(".", ",")

    log(f"Matchede nummerplader til Evotec units: {matched}/{len(raw_rows)}")
    log(f"Metadata udfyldt: Groups API={groups_api}, Groups tom={groups_empty}, Client={client_count}")
    if out:
        sm = out[0]
        log(f"Sample metadata: plate={sm.get(headers[plate_i])}, groups={sm.get('Groups')}, depot={sm.get('Depot')}, client={sm.get('Client')}")
    log("Færdig. Distance-tal kommer direkte fra Evotec Excel-eksport og er ikke genberegnet.")
    log("CSV downloades med dansk talformat: semikolon og komma som decimaltegn.")
    return new_headers, out, {"matched": matched, "source_rows": len(raw_rows), "total_km": total_km}

def parse_multipart_form(handler):
    content_type = handler.headers.get("Content-Type", "")
    length = int(handler.headers.get("Content-Length", "0") or "0")
    body = handler.rfile.read(length)
    fields = {}
    files = {}

    m = re.search(r"boundary=(?P<b>[^;]+)", content_type)
    if not m:
        parsed = parse_qs(body.decode("utf-8", "replace"), keep_blank_values=True)
        for k, v in parsed.items():
            fields[k] = v[0] if v else ""
        return fields, files

    boundary = m.group("b").strip().strip('"')
    marker = ("--" + boundary).encode("utf-8")
    for part in body.split(marker):
        if not part or part in (b"--", b"--\r\n"):
            continue
        if part.startswith(b"\r\n"):
            part = part[2:]
        if part.endswith(b"--\r\n"):
            part = part[:-4]
        elif part.endswith(b"--"):
            part = part[:-2]
        if part.endswith(b"\r\n"):
            part = part[:-2]

        header_blob, sep, payload = part.partition(b"\r\n\r\n")
        if not sep:
            continue
        header_text = header_blob.decode("utf-8", "replace")
        cd_line = ""
        for line in header_text.split("\r\n"):
            if line.lower().startswith("content-disposition:"):
                cd_line = line
                break
        if not cd_line:
            continue

        name_m = re.search(r'name="([^"]+)"', cd_line)
        if not name_m:
            continue
        name = name_m.group(1)
        filename_m = re.search(r'filename="([^"]*)"', cd_line)
        if filename_m:
            files[name] = {"filename": filename_m.group(1), "content": payload}
        else:
            fields[name] = payload.decode("utf-8", "replace")
    return fields, files

class Handler(BaseHTTPRequestHandler):
    def log_message(self, format, *args):
        return

    def send_json(self, status, obj):
        data = json.dumps(obj, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        try:
            self.wfile.write(data)
        except (BrokenPipeError, ConnectionAbortedError):
            pass

    def do_GET(self):
        global LAST_ROWS, LAST_HEADERS
        path = urlparse(self.path).path
        if path == "/download.csv":
            buf = io.StringIO(newline="")
            writer = csv.DictWriter(buf, fieldnames=LAST_HEADERS, extrasaction="ignore", delimiter=";", quoting=csv.QUOTE_MINIMAL)
            writer.writeheader()
            for row in LAST_ROWS:
                writer.writerow({h: csv_value(row.get(h, "")) for h in LAST_HEADERS})
            data = buf.getvalue().encode("utf-8-sig")
            self.send_response(200)
            self.send_header("Content-Type", "text/csv; charset=utf-8")
            self.send_header("Content-Disposition", 'attachment; filename="GLS_Distance_by_hours_enriched.csv"')
            self.send_header("Content-Length", str(len(data)))
            self.end_headers()
            self.wfile.write(data)
            return

        data = HTML.encode("utf-8")
        self.send_response(200)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def do_POST(self):
        global LAST_ROWS, LAST_HEADERS
        if urlparse(self.path).path != "/enrich":
            self.send_json(404, {"ok": False, "error": "Not found"})
            return
        try:
            form, uploaded = parse_multipart_form(self)
            fileitem = uploaded.get("xlsx")
            if not fileitem or not fileitem.get("content"):
                raise ValueError("Excel-fil mangler.")

            config = {
                "api_key": form.get("api_key", "").strip(),
                "base_url": form.get("base_url", "https://gps.evotec.dk").strip() or "https://gps.evotec.dk",
                "auth_mode": form.get("auth_mode", "query"),
                "api_param": form.get("api_param", "key").strip() or "key",
                "group_mode": form.get("group_mode", "smart"),
                "max_group_calls": form.get("max_group_calls", "300"),
                "timeout": int(float(form.get("timeout", "4") or "4")),
                "manual": form.get("manual", ""),
            }
            if not config["api_key"]:
                raise ValueError("API-nøgle mangler.")

            headers, rows, stats = enrich(fileitem["content"], config)
            LAST_HEADERS = headers
            LAST_ROWS = rows
            self.send_json(200, {"ok": True, "headers": headers, "rows": rows[:1000], "stats": stats, "log": "\n".join(LAST_LOG)})
        except Exception as exc:
            tb = traceback.format_exc()
            try:
                with open("last_error.txt", "w", encoding="utf-8") as f:
                    f.write(tb)
            except Exception:
                pass
            self.send_json(500, {"ok": False, "error": str(exc), "log": "\n".join(LAST_LOG) + "\nFEJL:\n" + tb})

def main():
    os.chdir(os.path.dirname(os.path.abspath(__file__)))
    srv = ThreadingHTTPServer(("127.0.0.1", PORT), Handler)
    url = f"http://127.0.0.1:{PORT}"
    print(f"{APP_TITLE} {VERSION} kører på {url}")
    webbrowser.open(url)
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        pass

if __name__ == "__main__":
    main()
