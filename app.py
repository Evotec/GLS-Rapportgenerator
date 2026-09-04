# GLS Evotec Rapportgenerator
# Pure Python / standard library only. No pip packages required.

from http.server import ThreadingHTTPServer, BaseHTTPRequestHandler
from urllib.parse import urlparse, parse_qs, urlencode
from urllib.request import Request, urlopen
from urllib.error import HTTPError, URLError
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
VERSION = "github-v1"
PORT = 8765
LAST_HEADERS = []
LAST_ROWS = []
LAST_LOG = []

NS = {"x": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}

HTML = r'''<!doctype html><html lang="da"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>GLS Rapportgenerator</title>
<style>
body{font-family:Arial,Helvetica,sans-serif;background:#f4f7fb;color:#14213d;margin:0}header{background:#0f172a;color:white;padding:16px 24px;font-size:22px;font-weight:700}.wrap{padding:22px}.card{background:white;border:1px solid #d9e2ef;border-radius:12px;padding:18px;margin-bottom:18px;box-shadow:0 1px 2px #0001}.grid{display:grid;grid-template-columns:190px 1fr;gap:12px 16px;max-width:980px;align-items:center}label{font-weight:700}input,select,textarea{width:100%;padding:10px;border:1px solid #bfd0e5;border-radius:8px;font-size:14px;box-sizing:border-box}button{border:0;border-radius:8px;padding:11px 16px;font-weight:700;cursor:pointer;margin-right:8px}.primary{background:#2563eb;color:white}.secondary{background:#334155;color:white}.ghost{background:#e2e8f0;color:#0f172a}.warn{background:#fff7ed;border:1px solid #fdba74;color:#7c2d12;padding:10px;border-radius:8px;margin-bottom:16px}.help{color:#5d6b82;font-size:13px;margin-top:12px}#log{background:#020617;color:#d1e7ff;font-family:Consolas,monospace;white-space:pre-wrap;max-height:270px;overflow:auto;padding:14px;border-radius:10px;font-size:13px}.tablewrap{overflow:auto;max-height:520px;border:1px solid #d9e2ef;border-radius:10px;background:white}table{border-collapse:collapse;width:max-content;min-width:100%}th,td{border:1px solid #e5e7eb;padding:6px 8px;font-size:12px;white-space:nowrap}th{background:#f1f5f9;position:sticky;top:0}.pill{display:inline-block;background:#e0e7ff;color:#3730a3;border-radius:20px;padding:7px 12px;margin-right:8px;margin-bottom:10px;font-size:13px}details{max-width:980px;margin-top:10px}summary{cursor:pointer;font-weight:700;color:#334155}
</style></head><body>
<header>GLS Rapportgenerator <small style="font-size:14px;color:#cbd5e1">Evotec</small></header>
<div class="wrap"><div class="card"><div class="warn"><b>Workflow:</b> Eksportér <b>Distance by hours</b> fra Evotec GPS/Mapon. Programmet bevarer alle distance-tal og tilføjer kun metadata fra Evotec API.</div>
<form id="form"><div class="grid">
<label>Evotec Excel-rapport</label><input type="file" name="xlsx" accept=".xlsx" required>
<label>Evotec API-nøgle</label><input type="password" id="api_key" name="api_key" placeholder="Indsæt API-nøgle" required>
</div><details><summary>Avanceret / teknikerindstillinger</summary><div class="grid" style="margin-top:12px">
<label>Evotec GPS adresse</label><input id="base_url" name="base_url" value="https://gps.evotec.dk">
<label>API-key parameter</label><input id="api_param" name="api_param" value="key">
<label>Auth mode</label><select id="auth_mode" name="auth_mode"><option value="query">Query parameter</option><option value="bearer">Bearer token</option></select>
<label>Max group-kald</label><input id="max_group_calls" name="max_group_calls" value="300">
<label>API timeout sek.</label><input id="timeout" name="timeout" value="4">
<label>Manuel mapping</label><textarea id="manual" name="manual" rows="4" placeholder="Valgfrit. Én pr. linje:&#10;DX92718=DK0026-København / Bojon Logistics Service | DK0026-København | Bojon Logistics Service"></textarea>
<label></label><button type="button" class="ghost" onclick="clearKey()">Slet gemt API-nøgle</button>
</div></details><div style="margin-top:16px"><button class="primary" type="submit">Lav rapport</button><button id="download" class="secondary" type="button" disabled onclick="location.href='/download.csv'">Download CSV</button><button class="ghost" type="button" onclick="document.getElementById('log').textContent='Klar.'">Ryd log</button></div><div class="help">API-nøglen gemmes kun lokalt i denne browser på denne PC. Den sendes ikke til GitHub.</div></form></div>
<div class="card"><div id="log">Klar.</div></div><div id="result" class="card" style="display:none"><div><span class="pill" id="rows"></span><span class="pill" id="matched"></span><span class="pill" id="km"></span></div><div class="tablewrap"><table id="tbl"></table></div></div></div>
<script>
const P='gls_evotec_';
for(const id of ['api_key','base_url','api_param','auth_mode','max_group_calls','timeout','manual']){const v=localStorage.getItem(P+id); if(v!==null) document.getElementById(id).value=v;}
if(!localStorage.getItem(P+'base_url')) document.getElementById('base_url').value='https://gps.evotec.dk';
document.getElementById('form').addEventListener('change',e=>{if(e.target.id) localStorage.setItem(P+e.target.id,e.target.value);});
function clearKey(){localStorage.removeItem(P+'api_key');document.getElementById('api_key').value='';}
document.getElementById('form').addEventListener('submit',async e=>{e.preventDefault();document.getElementById('download').disabled=true;document.getElementById('result').style.display='none';document.getElementById('log').textContent='Læser Excel og henter Evotec metadata...';const fd=new FormData(e.target);try{const r=await fetch('/enrich',{method:'POST',body:fd});const d=await r.json();document.getElementById('log').textContent=d.log||'';if(!d.ok){alert('Fejl: '+(d.error||'Ukendt fejl'));return;}render(d.headers,d.rows,d.stats);document.getElementById('download').disabled=false;}catch(err){document.getElementById('log').textContent='FEJL: '+err;alert(err);}});
function esc(s){return String(s??'').replace(/[&<>"']/g,m=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[m]));}
function render(headers,rows,stats){document.getElementById('result').style.display='block';document.getElementById('rows').textContent='Rækker: '+stats.source_rows;document.getElementById('matched').textContent='Matchede: '+stats.matched+'/'+stats.source_rows;document.getElementById('km').textContent='Total km: '+(stats.total_km||'');let h='<thead><tr>'+headers.map(x=>'<th>'+esc(x)+'</th>').join('')+'</tr></thead><tbody>';for(const row of rows.slice(0,500)){h+='<tr>'+headers.map(x=>'<td>'+esc(row[x])+'</td>').join('')+'</tr>'}h+='</tbody>';document.getElementById('tbl').innerHTML=h;}
</script></body></html>'''

def log(msg):
    LAST_LOG.append(str(msg))

def clean(v):
    if v is None:
        return ""
    return str(v).strip()

def norm_plate(v):
    return re.sub(r"[^A-Z0-9]", "", clean(v).upper())

def is_plate_like(v):
    s = norm_plate(v)
    return 4 <= len(s) <= 12 and any(c.isalpha() for c in s) and any(c.isdigit() for c in s)

def col_to_idx(ref):
    letters = ''.join(c for c in ref if c.isalpha())
    n = 0
    for c in letters:
        n = n * 26 + ord(c.upper()) - 64
    return n - 1

def xlsx_rows(data):
    with zipfile.ZipFile(io.BytesIO(data)) as z:
        shared = []
        try:
            root = ET.fromstring(z.read('xl/sharedStrings.xml'))
            for si in root.findall('x:si', NS):
                shared.append(''.join(t.text or '' for t in si.iter('{%s}t' % NS['x'])))
        except KeyError:
            pass
        sheet = 'xl/worksheets/sheet1.xml'
        root = ET.fromstring(z.read(sheet))
        out = []
        for row in root.findall('.//x:sheetData/x:row', NS):
            vals = []
            for c in row.findall('x:c', NS):
                idx = col_to_idx(c.attrib.get('r', 'A1'))
                while len(vals) <= idx:
                    vals.append('')
                typ = c.attrib.get('t')
                v = c.find('x:v', NS)
                is_el = c.find('x:is', NS)
                val = ''
                if typ == 's' and v is not None:
                    try:
                        val = shared[int(v.text)]
                    except Exception:
                        val = v.text or ''
                elif typ == 'inlineStr' and is_el is not None:
                    val = ''.join(t.text or '' for t in is_el.iter('{%s}t' % NS['x']))
                elif v is not None:
                    val = v.text or ''
                vals[idx] = clean(val)
            while vals and vals[-1] == '':
                vals.pop()
            out.append(vals)
        return out

def detect_table(rows):
    header_i = None
    for i, r in enumerate(rows):
        low = [clean(x).lower() for x in r]
        if 'licence plate' in low or 'license plate' in low:
            header_i = i
            break
    if header_i is None:
        raise ValueError('Kunne ikke finde header-rækken med Licence plate.')
    headers = [clean(x) for x in rows[header_i]]
    data = []
    for r in rows[header_i + 1:]:
        if not any(clean(x) for x in r):
            continue
        first = clean(r[0]).lower() if r else ''
        if first.startswith('total'):
            continue
        rr = r + [''] * (len(headers) - len(r))
        data.append(rr[:len(headers)])
    return headers, data, header_i + 1

def flatten(obj):
    if isinstance(obj, dict):
        yield obj
        for v in obj.values():
            yield from flatten(v)
    elif isinstance(obj, list):
        for v in obj:
            yield from flatten(v)

def api_get(config, endpoint, params=None):
    params = dict(params or {})
    base = config['base_url'].rstrip('/')
    headers = {'Accept': 'application/json'}
    if config.get('auth_mode') == 'bearer':
        headers['Authorization'] = 'Bearer ' + config['api_key']
    else:
        params[config.get('api_param') or 'key'] = config['api_key']
    url = base + endpoint
    if params:
        url += '?' + urlencode(params)
    req = Request(url, headers=headers)
    with urlopen(req, timeout=config.get('timeout', 4)) as resp:
        raw = resp.read().decode('utf-8', 'replace')
    try:
        return json.loads(raw)
    except Exception:
        return {'raw': raw}

def val(d, names):
    for n in names:
        if isinstance(d, dict) and n in d and d[n] not in (None, ''):
            return d[n]
    return ''

def load_units(config):
    log('Henter units fra Evotec API...')
    data = api_get(config, '/api/v1/unit/list.json')
    by_plate = {}
    by_key = {}
    count = 0
    for d in flatten(data):
        uid = val(d, ['id','unit_id','unitId'])
        plate = val(d, ['licence_plate','license_plate','numberplate','number_plate','plate','label','registration_number','reg_number'])
        if not uid and not plate:
            continue
        if not uid or not (is_plate_like(plate) or val(d, ['fuel_type','fuelType','make','model','title','name'])):
            continue
        plate_n = norm_plate(plate)
        keys = set()
        if uid:
            keys.add(str(uid))
        if plate_n:
            keys.add(plate_n)
        imei = val(d, ['imei','device_imei','deviceImei','object_id','objectId','device_id','deviceId'])
        if imei:
            keys.add(str(imei))
        make = clean(val(d, ['make','brand']))
        model = clean(val(d, ['model']))
        make_model = clean((make + ' ' + model).strip() or val(d, ['make_model','makeModel','title','name']))
        fuel = clean(val(d, ['fuel_type','fuelType','fuel','engine_type','engineType']))
        client = clean(val(d, ['client','customer','client_name','customer_name']))
        rec = {'unit_id': clean(uid), 'plate': plate_n, 'make_model': make_model, 'fuel': fuel, 'client': client, 'keys': keys}
        count += 1
        for k in keys:
            by_key.setdefault(k, rec)
        if plate_n:
            by_plate.setdefault(plate_n, rec)
    log(f'Units: {len(by_plate)} nummerplader / {count} unit-objekter')
    return by_plate, by_key

def load_group_list(config):
    data = api_get(config, '/api/v1/unit_groups/list.json')
    groups = []
    seen = set()
    for d in flatten(data):
        gid = val(d, ['id','group_id','groupId'])
        name = val(d, ['name','title','label'])
        if gid and name and (gid, name) not in seen:
            seen.add((gid, name))
            groups.append({'id': clean(gid), 'name': clean(name)})
    log(f'Grupper hentet: {len(groups)}')
    return groups

def unit_keys_from_item(d):
    keys = set()
    for n in ['id','unit_id','unitId','unit','unitID','vehicle_id','vehicleId','car_id','carId']:
        v = val(d, [n])
        if v:
            keys.add(str(v))
    for n in ['licence_plate','license_plate','numberplate','number_plate','plate','label','registration_number','reg_number']:
        v = val(d, [n])
        if is_plate_like(v):
            keys.add(norm_plate(v))
    for n in ['imei','device_imei','deviceImei','object_id','objectId','device_id','deviceId']:
        v = val(d, [n])
        if v:
            keys.add(str(v))
    return keys

def load_groups(config, target_keys):
    if not target_keys:
        return {}
    groups = load_group_list(config)
    max_calls = int(config.get('max_group_calls') or 300)
    unit_to_groups = {}
    relations = 0
    skipped = 0
    log('Scanner groups via unit_groups/list_units.json?id=<group id> ...')
    for i, g in enumerate(groups[:max_calls], start=1):
        try:
            data = api_get(config, '/api/v1/unit_groups/list_units.json', {'id': g['id']})
        except Exception as e:
            if i <= 5:
                log(f'Group-kald fejlede for {g["name"]}: {e}')
            continue
        found = set()
        for d in flatten(data):
            found |= unit_keys_from_item(d)
        hits = found & target_keys
        if len(hits) > max(250, len(target_keys) * 0.35):
            skipped += 1
            continue
        for k in hits:
            unit_to_groups.setdefault(k, set()).add(g['name'])
            relations += 1
        if i <= 3 or i % 25 == 0:
            log(f'Group-scan status: {i}/{min(len(groups), max_calls)}, relationer={relations}, senest={g["name"]}, units={len(hits)}, skipped={skipped}')
    log(f'Units/group-keys med grupper: {len(unit_to_groups)}; relationer={relations}; brede svar sprunget over={skipped}')
    return unit_to_groups

def leaf_groups(groups):
    gs = sorted({clean(g) for g in groups if clean(g)})
    leaves = []
    for g in gs:
        if not any(o.startswith(g + ' / ') for o in gs):
            leaves.append(g)
    return leaves

def split_group_path(groups):
    leaves = leaf_groups(groups)
    depot = ''
    client = ''
    if leaves:
        parts = [p.strip() for p in leaves[0].split('/')]
        depot = parts[0] if parts else ''
        client = parts[-1] if len(parts) > 1 else ''
    return ', '.join(leaves), depot, client

def parse_manual(text):
    out = {}
    for line in (text or '').splitlines():
        if '=' not in line:
            continue
        k, v = line.split('=', 1)
        parts = [p.strip() for p in v.split('|')]
        out[norm_plate(k)] = {'groups': parts[0] if len(parts)>0 else '', 'depot': parts[1] if len(parts)>1 else '', 'client': parts[2] if len(parts)>2 else ''}
    return out

def enrich(xlsx_data, config):
    LAST_LOG.clear()
    headers, rows, header_no = detect_table(xlsx_rows(xlsx_data))
    log(f'Læser Excel: header række {header_no}; kolonner: {len(headers)}; data-rækker: {len(rows)}')
    lower = [h.lower().strip() for h in headers]
    def idx(*names):
        for n in names:
            if n.lower() in lower:
                return lower.index(n.lower())
        return -1
    plate_i = idx('Licence plate','License plate')
    make_i = idx('Make / Model','Make /Model','Make/Model')
    title_i = idx('Vehicle title')
    if plate_i < 0:
        raise ValueError('Licence plate-kolonne ikke fundet.')
    by_plate, by_key = load_units(config)
    target_keys = set()
    for r in rows:
        p = norm_plate(r[plate_i] if plate_i < len(r) else '')
        if p:
            target_keys.add(p)
            if p in by_plate:
                target_keys |= by_plate[p]['keys']
    group_map = load_groups(config, target_keys)
    manual = parse_manual(config.get('manual',''))
    meta_headers = ['Groups','Depot','Client','Fuel type','Unit ID']
    insert_after = max([x for x in [title_i, make_i, plate_i] if x >= 0])
    new_headers = headers[:insert_after+1] + meta_headers + headers[insert_after+1:]
    out = []
    matched = groups_api = groups_empty = client_count = 0
    for r in rows:
        src = {headers[i]: clean(r[i]) if i < len(r) else '' for i in range(len(headers))}
        plate = norm_plate(r[plate_i] if plate_i < len(r) else '')
        u = by_plate.get(plate)
        groups = []
        if u:
            matched += 1
            for k in u['keys']:
                groups.extend(group_map.get(k, []))
        gtxt, depot, client_from_group = split_group_path(groups)
        if plate in manual:
            m = manual[plate]
            gtxt = m.get('groups') or gtxt
            depot = m.get('depot') or depot
            client_from_group = m.get('client') or client_from_group
        client = (client_from_group or (u.get('client','') if u else ''))
        fuel = u.get('fuel','') if u else ''
        uid = u.get('unit_id','') if u else ''
        if gtxt: groups_api += 1
        else: groups_empty += 1
        if client: client_count += 1
        dst = {}
        for h in headers[:insert_after+1]: dst[h] = src[h]
        dst.update({'Groups': gtxt, 'Depot': depot, 'Client': client, 'Fuel type': fuel, 'Unit ID': uid})
        for h in headers[insert_after+1:]: dst[h] = src[h]
        out.append(dst)
    total_km = ''
    total_col = next((h for h in headers if 'total' in h.lower() and 'km' in h.lower()), '')
    if total_col:
        s = 0.0
        for r in out:
            try: s += float(str(r.get(total_col,'')).replace(',','.'))
            except Exception: pass
        total_km = round(s, 1)
    log(f'Matchede nummerplader til Evotec units: {matched}/{len(rows)}')
    log(f'Metadata udfyldt: Groups API={groups_api}, Groups tom={groups_empty}, Client={client_count}')
    if out:
        sm = out[0]
        log(f"Sample metadata: plate={sm.get(headers[plate_i])}, groups={sm.get('Groups')}, depot={sm.get('Depot')}, client={sm.get('Client')}")
    log('Færdig. Distance-tal kommer direkte fra Evotec Excel-eksport og er ikke genberegnet.')
    return new_headers, out, {'matched': matched, 'source_rows': len(rows), 'total_km': total_km}

def parse_multipart(handler):
    ctype = handler.headers.get('Content-Type','')
    length = int(handler.headers.get('Content-Length','0') or '0')
    body = handler.rfile.read(length)
    fields, files = {}, {}
    m = re.search(r'boundary=([^;]+)', ctype)
    if not m:
        for k, v in parse_qs(body.decode('utf-8','replace'), keep_blank_values=True).items():
            fields[k] = v[0] if v else ''
        return fields, files
    boundary = ('--' + m.group(1).strip().strip('"')).encode()
    for part in body.split(boundary):
        if not part or part in (b'--', b'--\r\n'):
            continue
        part = part.strip(b'\r\n')
        if part.endswith(b'--'):
            part = part[:-2]
        head, sep, payload = part.partition(b'\r\n\r\n')
        if not sep:
            continue
        htxt = head.decode('utf-8','replace')
        nm = re.search(r'name="([^"]+)"', htxt)
        if not nm:
            continue
        name = nm.group(1)
        fm = re.search(r'filename="([^"]*)"', htxt)
        if fm:
            files[name] = {'filename': fm.group(1), 'content': payload.rstrip(b'\r\n')}
        else:
            fields[name] = payload.decode('utf-8','replace').rstrip('\r\n')
    return fields, files

class Handler(BaseHTTPRequestHandler):
    def log_message(self, fmt, *args):
        return
    def send_json(self, status, obj):
        data = json.dumps(obj, ensure_ascii=False).encode('utf-8')
        self.send_response(status)
        self.send_header('Content-Type','application/json; charset=utf-8')
        self.send_header('Content-Length', str(len(data)))
        self.end_headers()
        try:
            self.wfile.write(data)
        except (BrokenPipeError, ConnectionAbortedError):
            pass
    def do_GET(self):
        global LAST_HEADERS, LAST_ROWS
        if urlparse(self.path).path == '/download.csv':
            buf = io.StringIO(newline='')
            w = csv.DictWriter(buf, fieldnames=LAST_HEADERS, extrasaction='ignore', delimiter=';')
            w.writeheader(); w.writerows(LAST_ROWS)
            data = buf.getvalue().encode('utf-8-sig')
            self.send_response(200)
            self.send_header('Content-Type','text/csv; charset=utf-8')
            self.send_header('Content-Disposition','attachment; filename="GLS_Distance_by_hours_enriched.csv"')
            self.send_header('Content-Length', str(len(data)))
            self.end_headers(); self.wfile.write(data); return
        data = HTML.encode('utf-8')
        self.send_response(200); self.send_header('Content-Type','text/html; charset=utf-8'); self.send_header('Content-Length', str(len(data))); self.end_headers(); self.wfile.write(data)
    def do_POST(self):
        global LAST_HEADERS, LAST_ROWS
        if urlparse(self.path).path != '/enrich':
            self.send_json(404, {'ok': False, 'error': 'Not found'}); return
        try:
            fields, files = parse_multipart(self)
            if 'xlsx' not in files or not files['xlsx']['content']:
                raise ValueError('Excel-fil mangler.')
            config = {
                'api_key': fields.get('api_key','').strip(),
                'base_url': fields.get('base_url','https://gps.evotec.dk').strip() or 'https://gps.evotec.dk',
                'api_param': fields.get('api_param','key').strip() or 'key',
                'auth_mode': fields.get('auth_mode','query'),
                'max_group_calls': fields.get('max_group_calls','300'),
                'timeout': int(float(fields.get('timeout','4') or '4')),
                'manual': fields.get('manual',''),
            }
            if not config['api_key']:
                raise ValueError('API-nøgle mangler.')
            headers, rows, stats = enrich(files['xlsx']['content'], config)
            LAST_HEADERS, LAST_ROWS = headers, rows
            self.send_json(200, {'ok': True, 'headers': headers, 'rows': rows[:1000], 'stats': stats, 'log': '\n'.join(LAST_LOG)})
        except Exception as exc:
            tb = traceback.format_exc()
            try:
                with open('last_error.txt','w',encoding='utf-8') as f: f.write(tb)
            except Exception:
                pass
            self.send_json(500, {'ok': False, 'error': str(exc), 'log': '\n'.join(LAST_LOG) + '\nFEJL:\n' + tb})

def main():
    os.chdir(os.path.dirname(os.path.abspath(__file__)))
    srv = ThreadingHTTPServer(('127.0.0.1', PORT), Handler)
    url = f'http://127.0.0.1:{PORT}'
    print(f'{APP_TITLE} {VERSION} kører på {url}')
    webbrowser.open(url)
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        pass

if __name__ == '__main__':
    main()
