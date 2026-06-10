"""
NeoProprio — Extracteur d'opérations certifiées
Deux outils dans une seule application, séparés par navigation :
  • NF Habitat  (base Cerqual / Qlik)        -> routes /api/nf/...
  • Prestaterre BEE (API interne du site)    -> routes /api/presta/...
"""

import time, json, re, io, os, threading
from datetime import datetime
from flask import Flask, render_template_string, jsonify, request, send_file, Response
import requests

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

app = Flask(__name__)

# ════════════════════════════════════════════════════════════════════
#  DRIVER PARTAGÉ
# ════════════════════════════════════════════════════════════════════
import tempfile, shutil, subprocess

def _kill_zombies():
    """Tue d'éventuels processus Chrome restés d'une extraction précédente (conteneur)."""
    for name in ("chrome", "chromium", "chromium-browser", "chromedriver"):
        try:
            subprocess.run(["pkill", "-9", "-f", name], check=False,
                           stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=5)
        except Exception:
            pass


def get_driver():
    _kill_zombies()
    # Profil temporaire UNIQUE par lancement : évite le verrou SingletonLock
    # qui faisait planter la 2e extraction en production ("not connected to DevTools").
    profile = tempfile.mkdtemp(prefix="np-chrome-")

    opts = Options()
    for arg in [
        "--headless=new", "--no-sandbox", "--disable-dev-shm-usage", "--disable-gpu",
        "--disable-software-rasterizer", "--disable-extensions", "--disable-background-networking",
        "--disable-renderer-backgrounding", "--disable-backgrounding-occluded-windows",
        "--disable-background-timer-throttling", "--disable-features=Translate,BackForwardCache",
        "--no-first-run", "--no-default-browser-check", "--mute-audio",
        "--blink-settings=imagesEnabled=false",  # économie mémoire (données = texte uniquement)
        "--window-size=1600,900", "--lang=fr-FR",
        f"--user-data-dir={profile}",
    ]:
        opts.add_argument(arg)
    opts.add_experimental_option("excludeSwitches", ["enable-automation"])

    if os.path.exists("/root/.nix-profile/bin/chromium"):
        opts.binary_location = "/root/.nix-profile/bin/chromium"
        driver = webdriver.Chrome(service=Service("/root/.nix-profile/bin/chromedriver"), options=opts)
    else:
        from webdriver_manager.chrome import ChromeDriverManager
        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=opts)

    driver._profile_dir = profile  # pour le nettoyage en fin d'extraction
    return driver


def _quit_driver(driver):
    """Ferme proprement le navigateur et supprime son profil temporaire."""
    if not driver:
        return
    prof = getattr(driver, "_profile_dir", None)
    try:
        driver.quit()
    except Exception:
        pass
    if prof:
        shutil.rmtree(prof, ignore_errors=True)


def extract_year(value):
    if not value:
        return ""
    m = re.search(r"(?:19|20)\d{2}", str(value))
    return m.group(0) if m else ""


# ════════════════════════════════════════════════════════════════════
#  OUTIL 1 — NF HABITAT (Qlik / Cerqual)
# ════════════════════════════════════════════════════════════════════
state_nf = {"status": "idle", "message": "", "progress": 0, "data": [], "total": 0}

QLIK_APP = "ed735054-1aec-4957-ad1b-c531be3a90bd"
QLIK_URI = "https://qlik-public.nf-habitat.fr"
NF_URL   = "https://www.nf-habitat.fr/moteur-de-recherche-des-operations-certifiees-nf-habitat/"
CERQUAL_PDF = "https://api.cerqual-pro.net/v1/qualitel_site_service/certificats/{ref}"

NF_SETUP_JS = """
window.__qlikReady = false; window.__qlikError = null;
(async () => {
  try {
    const r = await fetch('""" + QLIK_URI + """/qps/csrftoken', {credentials:'include'});
    const csrf = r.headers.get('qlik-csrf-token') || '';
    const schema = await (await fetch('https://unpkg.com/enigma.js@2.14.0/schemas/12.2015.0.json')).json();
    const wsUrl = '""" + QLIK_URI.replace('https','wss') + """/app/""" + QLIK_APP + """' + (csrf ? '?qlik-csrf-token='+csrf : '');
    const session = window.enigma.create({url: wsUrl, schema: schema});
    const global = await session.open();
    window.__qlikApp = await global.openDoc('""" + QLIK_APP + """');
    window.__qlikReady = true;
  } catch(e) { window.__qlikError = e.toString(); }
})();
"""

def nf_build_js(filters):
    fields = ["Numéro de contrat", "Nom commercial", "Code postal", "Ville",
              "Département", "Statut de certification", "Promoteur Cabinet",
              "Promoteur Groupe", "URL Promoteur", "Date enregistrement"]
    parts = []
    if filters.get("marque"):
        parts.append("[Marque]={'" + filters["marque"] + "'}")
    if filters.get("type"):
        parts.append("[Type de logement]={'" + filters["type"] + "'}")
    if filters.get("region"):
        parts.append("[Région]={'" + filters["region"] + "'}")
    statut = filters.get("statut", "both")
    if statut == "both":
        parts.append('[Statut de certification]={"Certifiée","En cours d\u2019évaluation"}')
    elif statut == "certifiee":
        parts.append('[Statut de certification]={"Certifiée"}')
    elif statut == "encours":
        parts.append('[Statut de certification]={"En cours d\u2019évaluation"}')
    set_expr = "<" + ", ".join(parts) + ">" if parts else "<>"
    fields_json = json.dumps(fields, ensure_ascii=False)
    set_json    = json.dumps(set_expr, ensure_ascii=False)
    return """
const cb = arguments[arguments.length - 1];
(async () => {
  try {
    const app = window.__qlikApp;
    const fields = """ + fields_json + """;
    const setExpr = """ + set_json + """;
    const dims = fields.map(f => ({qDef: {qFieldDefs: [f]}, qNullSuppression: false}));
    const obj = await app.createSessionObject({
      qInfo: {qType: 'hypercube'},
      qHyperCubeDef: {
        qInitialDataFetch: [{qHeight: 900, qWidth: 11}],
        qDimensions: dims,
        qMeasures: [{qDef: {qDef: "sum({" + setExpr + "} 1)"}, qLabel: "s"}],
        qSuppressZero: true, qSuppressMissing: false, qMode: 'S', qStateName: '$'
      }
    });
    window.__listObj = obj;
    const layout = await obj.getLayout();
    const size = layout.qHyperCube.qSize;
    cb({totalRows: size.qcy,
        firstPage: layout.qHyperCube.qDataPages[0].qMatrix.map(r => r.map(c => c.qText))});
  } catch(e) { cb({error: e.toString()}); }
})();
"""

NF_PAGE_JS = """
const top = arguments[0], h = arguments[1], cb = arguments[arguments.length-1];
(async () => {
  try {
    const pages = await window.__listObj.getHyperCubeData('/qHyperCubeDef',
      [{qTop: top, qLeft: 0, qHeight: h, qWidth: 11}]);
    cb(pages[0].qMatrix.map(r => r.map(c => c.qText)));
  } catch(e) { cb({error: e.toString()}); }
})();
"""

def run_nf(filters):
    state_nf.update({"status": "running", "message": "Ouverture du navigateur...", "progress": 5, "data": [], "total": 0})
    driver = None
    try:
        driver = get_driver()
        driver.set_script_timeout(90)
        state_nf.update({"message": "Connexion au site NF Habitat...", "progress": 15})
        driver.get(NF_URL)
        time.sleep(6)
        try:
            driver.execute_script("document.querySelectorAll('[class*=axeptio],[id*=axeptio]').forEach(e=>e.remove());")
        except Exception:
            pass

        state_nf.update({"message": "Connexion à la base Cerqual...", "progress": 25})
        driver.execute_script(NF_SETUP_JS)
        for _ in range(40):
            if driver.execute_script("return window.__qlikReady;"):
                break
            err = driver.execute_script("return window.__qlikError;")
            if err:
                state_nf.update({"status": "error", "message": f"Erreur Qlik : {err}"}); return
            time.sleep(1)

        state_nf.update({"message": "Récupération des données...", "progress": 40})
        res = driver.execute_async_script(nf_build_js(filters))
        if isinstance(res, dict) and "error" in res:
            state_nf.update({"status": "error", "message": res["error"]}); return

        total = res["totalRows"]
        state_nf.update({"message": f"{total} opérations trouvées — chargement...", "progress": 50, "total": total})

        matrix = res["firstPage"]
        fetched = len(matrix)
        while fetched < total:
            page = driver.execute_async_script(NF_PAGE_JS, fetched, min(900, total - fetched))
            if isinstance(page, dict) and "error" in page:
                break
            matrix.extend(page)
            fetched = len(matrix)
            state_nf.update({"message": f"{fetched}/{total} récupérées...", "progress": 50 + int((fetched / total) * 40)})

        seen, records = set(), []
        for r in matrix:
            num = r[0] if r else ""
            if num in seen:
                continue
            seen.add(num)
            statut = r[5] if len(r) > 5 else ""
            date_val = r[9] if len(r) > 9 else ""
            pdf_url = CERQUAL_PDF.format(ref=num) if statut == "Certifiée" and num else ""
            records.append({
                "reference": num,
                "nom": r[1] if len(r) > 1 else "",
                "cp": r[2] if len(r) > 2 else "",
                "ville": r[3] if len(r) > 3 else "",
                "departement": r[4] if len(r) > 4 else "",
                "statut": statut,
                "promoteur": r[6] if (len(r) > 6 and r[6] not in ("", "-")) else (r[7] if len(r) > 7 else ""),
                "groupe": r[7] if len(r) > 7 else "",
                "url_promoteur": r[8] if len(r) > 8 else "",
                "date": date_val,
                "year": extract_year(date_val),
                "pdf": pdf_url,
            })

        state_nf.update({"status": "done", "message": f"{len(records)} opérations extraites", "progress": 100, "data": records})
    except Exception as e:
        state_nf.update({"status": "error", "message": str(e)})
    finally:
        _quit_driver(driver)


@app.route("/api/nf/extract", methods=["POST"])
def nf_extract():
    if state_nf["status"] == "running":
        return jsonify({"error": "Extraction déjà en cours"}), 400
    t = threading.Thread(target=run_nf, args=(request.json or {},)); t.daemon = True; t.start()
    return jsonify({"ok": True})


@app.route("/api/nf/status")
def nf_status():
    return jsonify({"status": state_nf["status"], "message": state_nf["message"], "progress": state_nf["progress"],
                    "total": state_nf["total"], "count": len(state_nf["data"]), "preview": state_nf["data"][:20]})


@app.route("/api/nf/pdf/<reference>")
def nf_pdf(reference):
    if not reference or not re.fullmatch(r"[A-Za-z0-9_-]+", reference):
        return "Référence invalide", 400
    try:
        r = requests.get(CERQUAL_PDF.format(ref=reference), timeout=30, headers={"User-Agent": "Mozilla/5.0"})
    except Exception as e:
        return f"Erreur de connexion au certificat : {e}", 502
    if r.status_code != 200 or not r.content:
        return f"Certificat introuvable ({r.status_code})", 404
    return Response(r.content, mimetype="application/pdf",
                    headers={"Content-Disposition": f'inline; filename="{reference}.pdf"',
                             "Cache-Control": "public, max-age=3600"})


@app.route("/api/nf/download")
def nf_download():
    if not state_nf["data"]:
        return "Aucune donnée", 400
    rows = state_nf["data"]
    year_filter = (request.args.get("year") or "").strip()
    if year_filter.isdigit():
        ymin = int(year_filter)
        rows = [r for r in rows if (r.get("year") or "").isdigit() and int(r["year"]) >= ymin]
    if not rows:
        return "Aucune donnée pour ce filtre", 400

    host = request.host
    scheme = "http" if ("localhost" in host or "127.0.0.1" in host) else "https"
    base = f"{scheme}://{host}/"

    wb = Workbook(); ws = wb.active; ws.title = "Opérations"
    headers = ["Référence", "Nom opération", "Code postal", "Ville", "Département", "Statut",
               "Promoteur", "Groupe promoteur", "Site web promoteur", "Date enregistrement", "Année", "Certificat PDF"]
    ws.append(headers)
    header_fill = PatternFill("solid", fgColor="0A7E57")
    header_font = Font(color="FFFFFF", bold=True, size=11)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    center_s = Alignment(horizontal="center", vertical="center")
    left = Alignment(horizontal="left", vertical="center")
    thin = Side(style="thin", color="D9E4D9")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    fill_alt = PatternFill("solid", fgColor="F1F8E9")
    fill_certif = PatternFill("solid", fgColor="C8E6C9")
    fill_cours = PatternFill("solid", fgColor="FFF3C4")
    link_font = Font(color="1565C0", underline="single")
    for c, h in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=c, value=h)
        cell.fill, cell.font, cell.alignment, cell.border = header_fill, header_font, center, border
    ws.row_dimensions[1].height = 30
    for i, rec in enumerate(rows, start=2):
        values = [rec.get("reference", ""), rec.get("nom", ""), rec.get("cp", ""), rec.get("ville", ""),
                  rec.get("departement", ""), rec.get("statut", ""), rec.get("promoteur", ""), rec.get("groupe", ""),
                  rec.get("url_promoteur", ""), rec.get("date", ""), rec.get("year", ""), ""]
        for c, v in enumerate(values, start=1):
            cell = ws.cell(row=i, column=c, value=v)
            cell.border = border
            cell.alignment = center_s if c in (1, 3, 5, 11) else left
            if i % 2 == 0:
                cell.fill = fill_alt
        st = ws.cell(row=i, column=6)
        if rec.get("statut") == "Certifiée":
            st.fill = fill_certif
        elif rec.get("statut"):
            st.fill = fill_cours
        st.alignment = center_s
        pdf_cell = ws.cell(row=i, column=12)
        if rec.get("pdf") and rec.get("reference"):
            pdf_cell.value = "Voir le certificat"
            pdf_cell.hyperlink = f"{base}api/nf/pdf/{rec['reference']}"
            pdf_cell.font = link_font
        else:
            pdf_cell.value = "—"
        pdf_cell.alignment = center_s
    for c, w in enumerate([16, 36, 11, 18, 16, 14, 26, 24, 30, 18, 8, 20], start=1):
        ws.column_dimensions[get_column_letter(c)].width = w
    ws.auto_filter.ref = ws.dimensions
    ws.freeze_panes = "A2"
    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    stamp = datetime.now().strftime("%Y%m%d_%H%M")
    suffix = f"_des_{year_filter}" if year_filter.isdigit() else ""
    return send_file(buf, as_attachment=True, download_name=f"NF_Habitat{suffix}_{stamp}.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# ════════════════════════════════════════════════════════════════════
#  OUTIL 2 — PRESTATERRE BEE (API interne du site)
# ════════════════════════════════════════════════════════════════════
state_presta = {"status": "idle", "message": "", "progress": 0, "data": [], "total": 0}

PRESTA_URL = "https://www.prestaterre.eu/operations-certifiees"

# JS unique : boucle ENTIÈRE côté navigateur, un seul execute_async_script.
# Évite les 172 allers-retours Python<>Chrome qui provoquaient le timeout DevTools.
PRESTA_ALL_JS = """
const cb      = arguments[arguments.length - 1];
const filters = arguments[0];
(async () => {
  try {
    if (typeof postAPI !== 'function' || typeof postCertifiedOperations === 'undefined') {
      cb({error: "API du site indisponible"}); return;
    }
    const allRows = [];
    let page = 1, lastId = "", totalPages = 1, guard = 0;
    while (guard++ < 600) {
      const payload = {
        page, lastId,
        navFlag: page > 1 ? "next" : "",
        department:           filters.department           || [],
        city:                 filters.city                 || "",
        operationName:        filters.operationName        || "",
        repositoryAndVersion: filters.repositoryAndVersion || [],
        mentionsAndLevels:    filters.mentionsAndLevels    || []
      };
      let resp;
      try { resp = await postAPI(postCertifiedOperations, payload); }
      catch(e) { break; }
      if (!resp || !resp.data || resp.data.length === 0) break;
      totalPages = parseInt(resp.totalPages) || 1;
      lastId     = resp.lastId || "";
      for (const row of resp.data) {
        const out = [];
        for (const k in row) { if (k !== 'latitude' && k !== 'longitude') out.push(row[k]); }
        allRows.push(out);
      }
      if (page >= totalPages) break;
      page++;
    }
    cb({rows: allRows, totalPages});
  } catch(e) { cb({error: e.toString()}); }
})();
"""


def clean_html(value):
    if value is None:
        return ""
    s = str(value)
    s = re.sub(r"<\s*br\s*/?\s*>", " / ", s, flags=re.IGNORECASE)
    s = re.sub(r"<[^>]+>", "", s)
    s = re.sub(r"\s*/\s*/\s*", " / ", s)
    return s.strip(" /").strip()


def run_presta(filters):
    state_presta.update({"status": "running", "message": "Ouverture du navigateur...",
                         "progress": 5, "data": [], "total": 0})
    driver = None
    try:
        driver = get_driver()
        # Timeout généreux : jusqu'à 10 min pour parcourir toutes les pages en JS
        driver.set_script_timeout(600)

        # Bloque les ressources lourdes/inutiles via CDP pour économiser la mémoire
        # (sinon Chrome est tué par OOM sur Railway pendant la longue extraction).
        # On garde jQuery + le JS du site (sur prestaterre.eu) qui fournit l'API.
        try:
            driver.execute_cdp_cmd("Network.enable", {})
            driver.execute_cdp_cmd("Network.setBlockedURLs", {"urls": [
                "*googletagmanager.com*", "*google-analytics.com*", "*doubleclick*",
                "*hs-scripts.com*", "*hs-analytics*", "*hsforms*", "*hscollectedforms*", "*hubspot*",
                "*axept.io*", "*axeptio*",
                "*tile.openstreetmap.org*", "*unpkg.com*",          # carte Leaflet + markercluster
                "*.png", "*.jpg", "*.jpeg", "*.gif", "*.webp",      # images inutiles
            ]})
        except Exception:
            pass

        state_presta.update({"message": "Connexion au site Prestaterre...", "progress": 15})
        driver.get(PRESTA_URL)
        time.sleep(4)
        try:
            driver.execute_script(
                "document.querySelectorAll('[class*=axeptio],[id*=axeptio]').forEach(e=>e.remove());")
        except Exception:
            pass

        state_presta.update({"message": "Initialisation de l'API du site...", "progress": 25})
        ready = False
        for _ in range(40):
            ready = driver.execute_script(
                "return (typeof postAPI==='function' && typeof postCertifiedOperations!=='undefined');")
            if ready:
                break
            time.sleep(0.5)
        if not ready:
            state_presta.update({"status": "error",
                                  "message": "Impossible d'accéder à l'API du site (postAPI non chargé)."}); return

        state_presta.update({"message": "Récupération de toutes les pages en cours...", "progress": 35})

        payload = {
            "department":           filters.get("department", []) or [],
            "city":                 (filters.get("city", "") or "").strip(),
            "operationName":        (filters.get("operationName", "") or "").strip(),
            "repositoryAndVersion": filters.get("repositoryAndVersion", []) or [],
            "mentionsAndLevels":    filters.get("mentionsAndLevels", []) or [],
        }

        # Un seul appel — le JS boucle en interne sur toutes les pages
        res = driver.execute_async_script(PRESTA_ALL_JS, payload)

        if isinstance(res, dict) and res.get("error"):
            state_presta.update({"status": "error", "message": res["error"]}); return

        raw_rows = res.get("rows", [])
        state_presta.update({"message": f"Nettoyage de {len(raw_rows)} lignes...", "progress": 88})

        seen, deduped = set(), []
        for r in raw_rows:
            vals = [clean_html(x) for x in r]
            while len(vals) < 6:
                vals.append("")
            key = (vals[0], vals[1], vals[2], vals[3])
            if key not in seen:
                seen.add(key)
                deduped.append({
                    "departement":  vals[0],
                    "ville":        vals[1],
                    "operation":    vals[2],
                    "referentiel":  vals[3],
                    "mentions":     vals[4],
                    "fin_validite": vals[5],
                })

        state_presta.update({"status": "done",
                             "message": f"{len(deduped)} opérations extraites",
                             "progress": 100, "data": deduped, "total": len(deduped)})
    except Exception as e:
        state_presta.update({"status": "error", "message": str(e)})
    finally:
        _quit_driver(driver)


@app.route("/api/presta/extract", methods=["POST"])
def presta_extract():
    if state_presta["status"] == "running":
        return jsonify({"error": "Extraction déjà en cours"}), 400
    t = threading.Thread(target=run_presta, args=(request.json or {},)); t.daemon = True; t.start()
    return jsonify({"ok": True})


@app.route("/api/presta/status")
def presta_status():
    return jsonify({"status": state_presta["status"], "message": state_presta["message"], "progress": state_presta["progress"],
                    "total": state_presta["total"], "count": len(state_presta["data"]), "preview": state_presta["data"][:20]})


@app.route("/api/presta/download")
def presta_download():
    if not state_presta["data"]:
        return "Aucune donnée", 400
    rows = state_presta["data"]
    wb = Workbook(); ws = wb.active; ws.title = "Opérations BEE"
    headers = ["Département", "Ville", "Opération", "Référentiel et Version", "Mentions et Niveaux", "Fin de validité"]
    ws.append(headers)
    header_fill = PatternFill("solid", fgColor="0A7E57")
    header_font = Font(color="FFFFFF", bold=True, size=11)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left = Alignment(horizontal="left", vertical="center", wrap_text=True)
    thin = Side(style="thin", color="C8E6C9")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    fill_alt = PatternFill("solid", fgColor="F9FBE7")
    for c, h in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=c, value=h)
        cell.fill, cell.font, cell.alignment, cell.border = header_fill, header_font, center, border
    ws.row_dimensions[1].height = 28
    for i, rec in enumerate(rows, start=2):
        values = [rec.get("departement", ""), rec.get("ville", ""), rec.get("operation", ""),
                  rec.get("referentiel", ""), rec.get("mentions", ""), rec.get("fin_validite", "")]
        for c, v in enumerate(values, start=1):
            cell = ws.cell(row=i, column=c, value=v)
            cell.border = border
            cell.alignment = center if c in (1, 6) else left
            if i % 2 == 0:
                cell.fill = fill_alt
    for c, w in enumerate([18, 22, 44, 36, 40, 16], start=1):
        ws.column_dimensions[get_column_letter(c)].width = w
    ws.auto_filter.ref = ws.dimensions
    ws.freeze_panes = "A2"
    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    stamp = datetime.now().strftime("%Y%m%d_%H%M")
    return send_file(buf, as_attachment=True, download_name=f"Prestaterre_BEE_{stamp}.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# Listes Prestaterre (identiques au site)
REFERENTIELS = ["BEE Logement Neuf", "BEE Logement Rénovation", "BEE Tertiaire Neuf",
                "BEE Tertiaire Rénovation", "BEE Tertiaire Exploitation"]
MENTIONS = ["Bâtiment Performance Énergétique", "BEE+", "Habitat qualité", "E+C-",
            "Effinergie", "BBCA", "RT 2012", "HPE", "RE2020"]
DEPARTEMENTS = ["01 - Ain","02 - Aisne","05 - Hautes-Alpes","06 - Alpes-Maritimes","07 - Ardèche",
    "08 - Ardennes","10 - Aube","12 - Aveyron","13 - Bouches-du-Rhône","14 - Calvados","16 - Charente",
    "17 - Charente-Maritime","18 - Cher","20 - Corse","21 - Côte-d'Or","22 - Côtes-d'Armor","24 - Dordogne",
    "25 - Doubs","26 - Drôme","27 - Eure","28 - Eure-et-Loir","29 - Finistère","2A - Corse-du-Sud",
    "2B - Haute-Corse","30 - Gard","31 - Haute-Garonne","33 - Gironde","34 - Hérault","35 - Ile-et-Vilaine",
    "36 - Indre","37 - Indre-et-Loire","38 - Isère","39 - Jura","40 - Landes","41 - Loir-et-Cher","42 - Loire",
    "44 - Loire-Atlantique","45 - Loiret","47 - Lot-et-Garonne","49 - Maine-et-Loire","50 - Manche","51 - Marne",
    "52 - Haute-Marne","54 - Meurthe-et-Moselle","56 - Morbihan","57 - Moselle","59 - Nord","60 - Oise",
    "62 - Pas-de-Calais","63 - Puy-de-Dôme","64 - Pyrénées-Atlantiques","67 - Bas-Rhin","68 - Haut-Rhin",
    "69 - Rhône","70 - Haute-Saône","71 - Saône-et-Loire","72 - Sarthe","73 - Savoie","74 - Haute-Savoie",
    "75 - Paris","76 - Seine-Maritime","77 - Seine-et-Marne","78 - Yvelines","79 - Deux-Sèvres","80 - Somme",
    "81 - Tarn","83 - Var","84 - Vaucluse","85 - Vendée","86 - Vienne","88 - Vosges","89 - Yonne","91 - Essonne",
    "92 - Hauts-de-Seine","93 - Seine-Saint-Denis","94 - Val-de-Marne","95 - Val-d'Oise"]

def _checkboxes(group_id, items):
    return "".join(
        f'<label class="ms-opt"><input type="checkbox" value="{v}" onchange="msLabel(\'{group_id}\')"/>'
        f'<span>{v}</span></label>' for v in items)


# ════════════════════════════════════════════════════════════════════
#  ROUTE PAGE
# ════════════════════════════════════════════════════════════════════
@app.route("/")
def index():
    return render_template_string(HTML)


LOGO_SRC = "data:image/png;base64,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"

HTML = '''<!DOCTYPE html>
<html lang="fr"><head>
<meta charset="UTF-8"/><meta name="viewport" content="width=device-width,initial-scale=1"/>
<title>NeoProprio — Extracteur d'opérations certifiées</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=Space+Grotesk:wght@500;600;700&display=swap" rel="stylesheet">
<style>
:root{
  --bg:#F4F7F5; --surface:#FFFFFF; --surface-2:#F4F7F5; --inset:#F0F4F2;
  --border:#E0E8E3; --border-soft:#EAF0EC;
  --text:#17241F; --muted:#5F6F68; --faint:#9AA6A0;
  --green:#0A7E57; --green-hover:#08694A; --green-soft:#E6F4EC;
  --accent:#0A7E57; --accent-soft:#E6F4EC;
  --danger:#C0392B; --warn:#B7791F;
  --radius:14px; --radius-sm:10px;
  --shadow:0 4px 18px -8px rgba(10,100,70,.12);
}
*{box-sizing:border-box;margin:0;padding:0}
html{scroll-behavior:smooth}
body{font-family:'Inter',system-ui,sans-serif;color:var(--text);min-height:100vh;
  background:var(--bg);-webkit-font-smoothing:antialiased;letter-spacing:.005em}

/* ── Header ── */
.topbar{position:sticky;top:0;z-index:50;
  background:var(--green);padding:14px 28px;
  display:flex;align-items:center;justify-content:space-between;
  box-shadow:0 2px 12px rgba(0,0,0,.12)}
.brand{display:flex;align-items:center;gap:14px}
.logo-text{font-family:'Space Grotesk',sans-serif;font-size:1.35rem;font-weight:700;letter-spacing:-.3px;display:flex;line-height:1}
.logo-neo{color:#fff}
.logo-propri{color:rgba(255,255,255,.5)}
.brand-sep{width:1px;height:22px;background:rgba(255,255,255,.25)}
.brand-tag{font-size:.8rem;color:rgba(255,255,255,.72);font-weight:500;letter-spacing:.01em}
.header-meta{display:flex;align-items:center;gap:8px;font-size:.74rem;
  color:rgba(255,255,255,.65);font-weight:500}
.dot{width:7px;height:7px;border-radius:50%;background:#7DFFCA}

/* ── Tab bar (sous le header) ── */
.tabbar{background:var(--surface);border-bottom:1px solid var(--border);padding:10px 28px}
.tabs{display:flex;gap:6px}
.tab{display:flex;align-items:center;gap:10px;padding:9px 18px;
  border-radius:var(--radius-sm);cursor:pointer;border:1px solid transparent;
  background:none;font-family:'Inter',sans-serif;transition:background .18s,border-color .18s}
.tab .tab-ic{display:grid;place-items:center;width:28px;height:28px;border-radius:7px;
  background:var(--inset);border:1px solid var(--border);flex-shrink:0;transition:.18s}
.tab .tab-ic svg{width:15px;height:15px;stroke:var(--muted)}
.tab b{display:block;font-family:'Space Grotesk',sans-serif;font-size:.88rem;font-weight:600;
  color:var(--muted);line-height:1.2;transition:color .18s}
.tab small{display:block;font-size:.68rem;font-weight:500;color:var(--faint);
  letter-spacing:.02em;transition:color .18s}
.tab:hover:not(.active){background:var(--inset)}
.tab.active{background:var(--green);border-color:var(--green)}
.tab.active .tab-ic{background:rgba(255,255,255,.2);border-color:transparent}
.tab.active .tab-ic svg{stroke:#fff}
.tab.active b,.tab.active small{color:#fff}

/* ── Shell ── */
.shell{max-width:1180px;margin:0 auto;padding:30px 22px 80px}

/* ── Panels ── */
.panel{display:none;animation:fade .35s ease}
.panel.active{display:block}
@keyframes fade{from{opacity:0;transform:translateY(6px)}to{opacity:1;transform:none}}

/* ── Cards ── */
.card{background:var(--surface);border:1px solid var(--border);border-radius:var(--radius);
  padding:26px 28px;margin-bottom:20px;box-shadow:var(--shadow)}
.eyebrow{display:flex;align-items:center;gap:9px;margin-bottom:20px}
.eyebrow .bar{width:18px;height:2px;background:var(--accent);border-radius:2px}
.eyebrow span{font-family:'Space Grotesk',sans-serif;font-size:.74rem;font-weight:600;
  letter-spacing:.16em;text-transform:uppercase;color:var(--accent)}
.eyebrow .src{margin-left:auto;font-size:.72rem;color:var(--faint);font-weight:500}

/* ── Form ── */
.grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(215px,1fr));gap:16px}
.field > label{display:block;font-size:.7rem;font-weight:600;color:var(--muted);
  text-transform:uppercase;letter-spacing:.08em;margin-bottom:7px}
.field input,.field select{width:100%;padding:11px 13px;background:var(--inset);color:var(--text);
  border:1px solid var(--border);border-radius:var(--radius-sm);font-family:inherit;font-size:.9rem;
  transition:border-color .2s,box-shadow .2s}
.field input::placeholder{color:var(--faint)}
.field input:focus,.field select:focus{outline:none;border-color:var(--accent);
  box-shadow:0 0 0 3px var(--accent-soft);background:#fff}
.field select{appearance:none;cursor:pointer;
  background-image:url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='12' height='12' fill='none' stroke='%235F6F68' stroke-width='2'%3E%3Cpath d='M2 4l4 4 4-4'/%3E%3C/svg%3E");
  background-repeat:no-repeat;background-position:right 13px center;padding-right:34px}
.field select option{background:#fff;color:var(--text)}

/* multi-select */
.ms{position:relative}
.ms-btn{width:100%;padding:11px 13px;background:var(--inset);border:1px solid var(--border);
  border-radius:var(--radius-sm);font-size:.9rem;cursor:pointer;display:flex;
  justify-content:space-between;align-items:center;gap:8px;user-select:none;
  color:var(--text);transition:border-color .2s}
.ms-btn:hover{border-color:var(--accent)}
.ms-btn .lab{white-space:nowrap;overflow:hidden;text-overflow:ellipsis;color:var(--faint)}
.ms-btn .lab.has{color:var(--text);font-weight:500}
.ms-btn .chev{color:var(--faint);transition:transform .2s}
.ms.open .ms-btn .chev{transform:rotate(180deg)}
.ms.open .ms-pop{display:block}
.ms-pop{display:none;position:absolute;z-index:40;top:calc(100% + 6px);left:0;right:0;
  background:#fff;border:1px solid var(--border);border-radius:var(--radius-sm);
  box-shadow:0 16px 36px -14px rgba(10,100,70,.2);max-height:262px;overflow-y:auto;padding:6px}
.ms-opt{display:flex;align-items:flex-start;gap:11px;padding:9px 10px;font-size:.875rem;
  line-height:1.35;border-radius:7px;cursor:pointer;color:var(--text);
  text-transform:none;letter-spacing:normal;font-weight:400}
.ms-opt:hover{background:var(--accent-soft)}
.ms-opt input{appearance:none;width:17px;height:17px;margin-top:1px;border:1.5px solid var(--border);
  border-radius:5px;background:#fff;cursor:pointer;flex-shrink:0;position:relative;transition:.15s}
.ms-opt input:checked{background:var(--accent);border-color:var(--accent)}
.ms-opt input:checked::after{content:"";position:absolute;left:5px;top:1.5px;width:4px;height:8px;
  border:solid #fff;border-width:0 2px 2px 0;transform:rotate(45deg)}

/* ── Buttons ── */
.actions{display:flex;gap:12px;flex-wrap:wrap;margin-top:22px;align-items:center}
.btn{padding:12px 22px;border-radius:var(--radius-sm);border:1px solid transparent;
  font-family:'Space Grotesk',sans-serif;font-size:.9rem;font-weight:600;cursor:pointer;
  display:inline-flex;align-items:center;gap:9px;
  transition:transform .12s,box-shadow .2s,background .2s,opacity .2s}
.btn:active{transform:translateY(1px)}
.btn svg{width:16px;height:16px}
.btn-primary{background:var(--green);color:#fff;box-shadow:0 8px 20px -10px var(--green)}
.btn-primary:hover{background:var(--green-hover);box-shadow:0 12px 24px -8px var(--green)}
.btn-primary:disabled{opacity:.5;cursor:not-allowed;box-shadow:none}
.btn-ghost{background:#fff;color:var(--muted);border-color:var(--border)}
.btn-ghost:hover{color:var(--text);border-color:var(--accent)}
.btn-export{background:#fff;color:var(--text);border-color:var(--border)}
.btn-export:hover:not(:disabled){border-color:var(--accent);color:var(--accent)}
.btn-export:disabled{opacity:.5;cursor:not-allowed}

/* ── Status + progress ── */
.status{display:none;margin-top:16px;padding:11px 15px;border-radius:var(--radius-sm);
  font-size:.86rem;font-weight:500;border:1px solid transparent}
.status.show{display:flex;align-items:center;gap:10px}
.status.loading{background:#FCF6E3;color:var(--warn);border-color:#F0E2B8}
.status.done{background:var(--green-soft);color:var(--green);border-color:#BFE6D0}
.status.error{background:#FBEAE8;color:var(--danger);border-color:#F1C9C3}
.spinner{width:14px;height:14px;border:2px solid currentColor;border-right-color:transparent;
  border-radius:50%;animation:spin .7s linear infinite;flex-shrink:0}
@keyframes spin{to{transform:rotate(360deg)}}
.progress{height:5px;background:var(--inset);border-radius:4px;margin-top:14px;display:none;overflow:hidden}
.progress.show{display:block}
.bar{height:100%;width:0;border-radius:4px;background:linear-gradient(90deg,var(--green-hover),var(--green));
  transition:width .4s ease}

/* ── Stats ── */
.stats{display:flex;gap:14px;flex-wrap:wrap;margin-bottom:18px}
.stat{flex:1;min-width:130px;background:var(--inset);border:1px solid var(--border);
  border-radius:var(--radius-sm);padding:14px 16px}
.stat .n{font-family:'Space Grotesk',sans-serif;font-size:1.7rem;font-weight:700;
  line-height:1;color:var(--text)}
.stat .n.accent{color:var(--accent)}
.stat .l{font-size:.72rem;color:var(--muted);margin-top:6px;letter-spacing:.04em}

/* ── Table ── */
.tablewrap{overflow:auto;max-height:540px;border:1px solid var(--border);border-radius:var(--radius-sm)}
table{width:100%;border-collapse:collapse;font-size:.84rem}
thead th{position:sticky;top:0;z-index:1;background:var(--inset);color:var(--muted);text-align:left;
  font-weight:600;font-size:.72rem;letter-spacing:.06em;text-transform:uppercase;
  padding:12px 14px;white-space:nowrap;border-bottom:2px solid var(--accent)}
tbody td{padding:11px 14px;border-bottom:1px solid var(--border-soft);
  vertical-align:top;color:var(--text)}
tbody tr:nth-child(even){background:rgba(10,126,87,.025)}
tbody tr:hover{background:var(--accent-soft)}
tbody tr:last-child td{border-bottom:none}
code.ref{font-family:ui-monospace,'SF Mono',Menlo,monospace;font-size:.78rem;
  color:var(--accent);font-weight:600}
.badge{display:inline-block;padding:3px 9px;border-radius:20px;font-size:.72rem;font-weight:600;white-space:nowrap}
.badge.ok{background:var(--green-soft);color:var(--green)}
.badge.wait{background:#FCF6E3;color:var(--warn)}
.badge.ref{background:var(--accent-soft);color:var(--accent);font-weight:500}
a.cert{color:var(--accent);font-size:.8rem;text-decoration:none;white-space:nowrap;font-weight:600}
a.cert:hover{text-decoration:underline}
tfoot td{padding:12px 14px;color:var(--faint);font-size:.78rem;font-style:italic}
.muted-cell{color:var(--faint)}

@media(max-width:760px){
  .brand-tag{display:none}
  .tab small{display:none}.tab{padding:9px 13px}
  .shell{padding:24px 14px 60px}.card{padding:20px 18px}
  .tabbar{padding:10px 14px}
}
@media(prefers-reduced-motion:reduce){*{animation:none!important;transition:none!important}}
</style></head>
<body>

<div class="topbar">
  <div class="brand">
    <div class="logo-text"><span class="logo-neo">neo</span><span class="logo-propri">proprio</span></div>
    <div class="brand-sep"></div>
    <span class="brand-tag">Extracteur d'opérations certifiées</span>
  </div>
  <div class="header-meta"><span class="dot"></span> Données en direct</div>
</div>

<div class="tabbar">
  <div class="tabs" id="tabs">
    <button class="tab active" data-tool="nf" onclick="switchTab('nf')">
      <span class="tab-ic"><svg viewBox="0 0 24 24" fill="none" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M3 11l9-7 9 7v9a1 1 0 0 1-1 1H4a1 1 0 0 1-1-1z"/><path d="M9 21v-6h6v6"/></svg></span>
      <span><b>NF Habitat</b><small>Cerqual · Qualitel</small></span>
    </button>
    <button class="tab" data-tool="presta" onclick="switchTab('presta')">
      <span class="tab-ic"><svg viewBox="0 0 24 24" fill="none" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 2l8 4v6c0 5-3.5 8-8 10-4.5-2-8-5-8-10V6z"/><path d="M9 12l2 2 4-4"/></svg></span>
      <span><b>Prestaterre BEE</b><small>Labels BEE</small></span>
    </button>
  </div>
</div>

<div class="shell">

  <!-- ============ PANEL NF HABITAT ============ -->
  <section class="panel active" data-tool="nf" id="panel-nf">
    <div class="card">
      <div class="eyebrow"><span class="bar"></span><span>Filtres</span><span class="src">nf-habitat.fr</span></div>
      <div class="grid">
        <div class="field"><label>Région</label>
          <select id="nf-region">
            <option value="">Toutes</option>
            <option value="Ile-de-France" selected>Île-de-France</option>
            <option value="Auvergne-Rhône-Alpes">Auvergne-Rhône-Alpes</option>
            <option value="Bretagne">Bretagne</option>
            <option value="Grand Est">Grand Est</option>
            <option value="Hauts-de-France">Hauts-de-France</option>
            <option value="Normandie">Normandie</option>
            <option value="Nouvelle-Aquitaine">Nouvelle-Aquitaine</option>
            <option value="Occitanie">Occitanie</option>
            <option value="Pays de la Loire">Pays de la Loire</option>
            <option value="Provence-Alpes-Côte d&#8217;Azur">Provence-Alpes-Côte d&#8217;Azur</option>
          </select></div>
        <div class="field"><label>Type de logement</label>
          <select id="nf-type">
            <option value="">Tous</option>
            <option value="Collectif" selected>Collectif</option>
            <option value="Individuel">Individuel</option>
          </select></div>
        <div class="field"><label>Certification</label>
          <select id="nf-marque">
            <option value="">Toutes</option>
            <option value="NF Habitat HQE" selected>NF Habitat HQE</option>
            <option value="NF Habitat">NF Habitat</option>
          </select></div>
        <div class="field"><label>Statut</label>
          <select id="nf-statut">
            <option value="both" selected>Certifiée + En cours</option>
            <option value="certifiee">Certifiée uniquement</option>
            <option value="encours">En cours uniquement</option>
          </select></div>
        <div class="field"><label>Année (à partir de)</label>
          <select id="nf-year">
            <option value="" selected>Toutes les années</option>
            <option>2027</option><option>2026</option><option>2025</option><option>2024</option>
            <option>2023</option><option>2022</option><option>2021</option><option>2020</option>
            <option>2019</option><option>2018</option><option>2017</option><option>2016</option>
            <option>2015</option><option>2014</option><option>2013</option><option>2012</option>
          </select></div>
        <div class="field"><label>Recherche libre</label>
          <input type="text" id="nf-search" placeholder="nom, ville, promoteur..."/></div>
      </div>
      <div class="actions">
        <button class="btn btn-primary" id="nf-go" onclick="nfRun()">
          <svg viewBox="0 0 24 24" fill="currentColor"><path d="M8 5v14l11-7z"/></svg> Lancer l'extraction</button>
        <button class="btn btn-ghost" onclick="nfReset()">Réinitialiser</button>
        <button class="btn btn-export" id="nf-dl" disabled onclick="nfDownload()">
          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 3v12m0 0l-4-4m4 4l4-4M4 19h16"/></svg> Télécharger l'Excel</button>
      </div>
      <div class="progress" id="nf-pw"><div class="bar" id="nf-pb"></div></div>
      <div class="status" id="nf-status"></div>
    </div>
    <div class="card" id="nf-results" style="display:none">
      <div class="eyebrow"><span class="bar"></span><span>Résultats</span><span class="src" id="nf-subtitle"></span></div>
      <div class="stats" id="nf-stats"></div>
      <div class="tablewrap"><table>
        <thead><tr><th>Référence</th><th>Nom opération</th><th>Promoteur</th><th>Statut</th>
          <th>CP</th><th>Ville</th><th>Département</th><th>Année</th><th>Certificat</th></tr></thead>
        <tbody id="nf-tbody"></tbody>
        <tfoot><tr><td colspan="9" id="nf-tfoot"></td></tr></tfoot>
      </table></div>
    </div>
  </section>

  <!-- ============ PANEL PRESTATERRE ============ -->
  <section class="panel" data-tool="presta" id="panel-presta">
    <div class="card">
      <div class="eyebrow"><span class="bar"></span><span>Filtres</span><span class="src">prestaterre.eu</span></div>
      <div class="grid">
        <div class="field"><label>Nom de l'opération</label>
          <input type="text" id="presta-op" placeholder="Nom de l'opération"/></div>
        <div class="field"><label>Référentiel et version</label>
          <div class="ms" id="presta-ms-ref">
            <div class="ms-btn" onclick="msToggle('presta-ms-ref')">
              <span class="lab" data-ph="Référentiel et version">Référentiel et version</span><span class="chev">▾</span></div>
            <div class="ms-pop">__REF__</div></div></div>
        <div class="field"><label>Mentions et niveaux</label>
          <div class="ms" id="presta-ms-ment">
            <div class="ms-btn" onclick="msToggle('presta-ms-ment')">
              <span class="lab" data-ph="Mentions et niveaux">Mentions et niveaux</span><span class="chev">▾</span></div>
            <div class="ms-pop">__MENT__</div></div></div>
        <div class="field"><label>Département</label>
          <div class="ms" id="presta-ms-dept">
            <div class="ms-btn" onclick="msToggle('presta-ms-dept')">
              <span class="lab" data-ph="Département">Département</span><span class="chev">▾</span></div>
            <div class="ms-pop">__DEPT__</div></div></div>
        <div class="field"><label>Ville</label>
          <input type="text" id="presta-ville" placeholder="Ville"/></div>
      </div>
      <div class="actions">
        <button class="btn btn-primary" id="presta-go" onclick="prestaRun()">
          <svg viewBox="0 0 24 24" fill="currentColor"><path d="M8 5v14l11-7z"/></svg> Lancer l'extraction</button>
        <button class="btn btn-ghost" onclick="prestaReset()">Réinitialiser</button>
        <button class="btn btn-export" id="presta-dl" disabled onclick="prestaDownload()">
          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 3v12m0 0l-4-4m4 4l4-4M4 19h16"/></svg> Télécharger l'Excel</button>
      </div>
      <div class="progress" id="presta-pw"><div class="bar" id="presta-pb"></div></div>
      <div class="status" id="presta-status"></div>
    </div>
    <div class="card" id="presta-results" style="display:none">
      <div class="eyebrow"><span class="bar"></span><span>Résultats</span><span class="src" id="presta-subtitle"></span></div>
      <div class="stats" id="presta-stats"></div>
      <div class="tablewrap"><table>
        <thead><tr><th>Département</th><th>Ville</th><th>Opération</th>
          <th>Référentiel</th><th>Mentions / Niveaux</th><th>Fin de validité</th></tr></thead>
        <tbody id="presta-tbody"></tbody>
        <tfoot><tr><td colspan="6" id="presta-tfoot"></td></tr></tfoot>
      </table></div>
    </div>
  </section>

</div>

<script>
const TOOLS = {nf:{timer:null,allData:[],total:0}, presta:{timer:null,allData:[],total:0}};
function el(k,s){ return document.getElementById(k+'-'+s); }

/* ---- Tabs ---- */
function switchTab(k){
  document.querySelectorAll('.tab').forEach(t=>t.classList.toggle('active',t.dataset.tool===k));
  document.querySelectorAll('.panel').forEach(p=>p.classList.toggle('active',p.dataset.tool===k));
}

/* ---- Multi-select ---- */
function msToggle(id){
  const e=document.getElementById(id);
  document.querySelectorAll('.ms.open').forEach(x=>{if(x.id!==id)x.classList.remove('open')});
  e.classList.toggle('open');
}
document.addEventListener('click',e=>{
  if(!e.target.closest('.ms')) document.querySelectorAll('.ms.open').forEach(x=>x.classList.remove('open'));
});
function msValues(id){ return [...document.querySelectorAll('#'+id+' .ms-pop input:checked')].map(c=>c.value); }
function msLabel(id){
  const v=msValues(id), lab=document.querySelector('#'+id+' .lab');
  if(v.length){ lab.textContent=v.length+' sélectionné'+(v.length>1?'s':''); lab.classList.add('has'); }
  else { lab.textContent=lab.getAttribute('data-ph'); lab.classList.remove('has'); }
}

/* ---- Generic runner ---- */
function setStatus(k,type,msg){
  const e=el(k,'status'); e.className='status show '+type;
  e.innerHTML=(type==='loading'?'<span class="spinner"></span>':'')+'<span>'+msg+'</span>';
}
function setProgress(k,pct){
  el(k,'pw').classList.add('show'); el(k,'pb').style.width=pct+'%';
  if(pct>=100) setTimeout(()=>el(k,'pw').classList.remove('show'),900);
}
function launch(k,filters,afficher){
  el(k,'go').disabled=true; el(k,'dl').disabled=true;
  el(k,'results').style.display='none';
  TOOLS[k].allData=[]; TOOLS[k].total=0;
  fetch('/api/'+k+'/extract',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(filters)})
    .then(()=>{ TOOLS[k].timer=setInterval(()=>poll(k,afficher),1200); });
}
function poll(k,afficher){
  fetch('/api/'+k+'/status').then(r=>r.json()).then(d=>{
    setStatus(k, d.status==='running'?'loading':d.status==='done'?'done':'error', d.message);
    setProgress(k,d.progress);
    if(d.status==='done'||d.status==='error'){
      clearInterval(TOOLS[k].timer); el(k,'go').disabled=false;
      if(d.status==='done'){ TOOLS[k].allData=d.preview; TOOLS[k].total=d.count; afficher(); el(k,'dl').disabled=false; }
    }
  });
}

/* ---- NF Habitat ---- */
function nfRun(){
  launch('nf',{region:el('nf','region').value,type:el('nf','type').value,
    marque:el('nf','marque').value,statut:el('nf','statut').value}, nfAfficher);
}
function nfAfficher(){
  const all=TOOLS.nf.allData, tot=TOOLS.nf.total;
  const q=el('nf','search').value.toLowerCase().trim(), yr=el('nf','year').value;
  let data=all;
  if(yr) data=data.filter(r=>r.year && parseInt(r.year)>=parseInt(yr));
  if(q) data=data.filter(r=>(r.nom+r.ville+r.promoteur+r.cp).toLowerCase().includes(q));
  const certif=all.filter(r=>r.statut==='Certifiée').length;
  const cours=all.filter(r=>r.statut && r.statut.includes('cours')).length;
  el('nf','stats').innerHTML=
    `<div class="stat"><div class="n accent">${tot}</div><div class="l">Total opérations</div></div>
     <div class="stat"><div class="n">${certif}</div><div class="l">Certifiées</div></div>
     <div class="stat"><div class="n">${cours}</div><div class="l">En cours d'évaluation</div></div>`;
  el('nf','subtitle').textContent=tot+' opérations'+(yr?(' · ≥ '+yr):'');
  el('nf','tbody').innerHTML=data.map(r=>`<tr>
    <td><code class="ref">${r.reference||'—'}</code></td>
    <td><strong>${r.nom||'—'}</strong></td>
    <td>${r.promoteur||'<span class="muted-cell">—</span>'}</td>
    <td>${r.statut==='Certifiée'?'<span class="badge ok">Certifiée</span>':'<span class="badge wait">En cours</span>'}</td>
    <td>${r.cp||'—'}</td><td>${r.ville||'—'}</td><td>${r.departement||'—'}</td>
    <td>${r.year||'<span class="muted-cell">—</span>'}</td>
    <td>${r.pdf?`<a class="cert" href="/api/nf/pdf/${r.reference}" target="_blank" rel="noopener">Voir le certificat</a>`:'<span class="muted-cell">—</span>'}</td>
  </tr>`).join('');
  el('nf','tfoot').textContent = yr
    ? `Filtre année ≥ ${yr} appliqué à l'aperçu — l'export Excel l'applique sur la totalité des ${tot} opérations.`
    : (tot>20?`Aperçu des 20 premières lignes sur ${tot} — téléchargez l'Excel pour tout voir.`:'');
  el('nf','results').style.display='block';
}
function nfDownload(){
  const y=el('nf','year').value;
  window.location.href='/api/nf/download'+(y?('?year='+encodeURIComponent(y)):'');
}
function nfReset(){
  el('nf','region').value='Ile-de-France'; el('nf','type').value='Collectif';
  el('nf','marque').value='NF Habitat HQE'; el('nf','statut').value='both';
  el('nf','year').value=''; el('nf','search').value='';
  el('nf','results').style.display='none'; el('nf','status').className='status';
  el('nf','dl').disabled=true; TOOLS.nf.allData=[]; TOOLS.nf.total=0;
}
el('nf','search').addEventListener('input',()=>{if(TOOLS.nf.allData.length)nfAfficher();});
el('nf','year').addEventListener('change',()=>{if(TOOLS.nf.allData.length)nfAfficher();});

/* ---- Prestaterre ---- */
function prestaRun(){
  launch('presta',{
    operationName:el('presta','op').value.trim(), city:el('presta','ville').value.trim(),
    repositoryAndVersion:msValues('presta-ms-ref'), mentionsAndLevels:msValues('presta-ms-ment'),
    department:msValues('presta-ms-dept')}, prestaAfficher);
}
function prestaAfficher(){
  const all=TOOLS.presta.allData, tot=TOOLS.presta.total;
  el('presta','stats').innerHTML=
    `<div class="stat"><div class="n accent">${tot}</div><div class="l">Total opérations</div></div>
     <div class="stat"><div class="n">${all.length}</div><div class="l">Affichées dans l'aperçu</div></div>`;
  el('presta','subtitle').textContent=tot+' opérations';
  el('presta','tbody').innerHTML=all.map(r=>`<tr>
    <td>${r.departement||'—'}</td>
    <td><strong>${r.ville||'—'}</strong></td>
    <td>${r.operation||'—'}</td>
    <td><span class="badge ref">${r.referentiel||'—'}</span></td>
    <td>${r.mentions||'<span class="muted-cell">—</span>'}</td>
    <td>${r.fin_validite||'<span class="muted-cell">—</span>'}</td>
  </tr>`).join('');
  el('presta','tfoot').textContent = tot>20?`Aperçu des 20 premières lignes sur ${tot} — téléchargez l'Excel pour tout voir.`:'';
  el('presta','results').style.display='block';
}
function prestaDownload(){ window.location.href='/api/presta/download'; }
function prestaReset(){
  el('presta','op').value=''; el('presta','ville').value='';
  ['presta-ms-ref','presta-ms-ment','presta-ms-dept'].forEach(id=>{
    document.querySelectorAll('#'+id+' input:checked').forEach(c=>c.checked=false); msLabel(id);
  });
  el('presta','results').style.display='none'; el('presta','status').className='status';
  el('presta','dl').disabled=true; TOOLS.presta.allData=[]; TOOLS.presta.total=0;
}
</script>
</body></html>'''

HTML = (HTML.replace("__REF__",  _checkboxes("presta-ms-ref",  REFERENTIELS))
            .replace("__MENT__", _checkboxes("presta-ms-ment", MENTIONS))
            .replace("__DEPT__", _checkboxes("presta-ms-dept", DEPARTEMENTS))
            .replace("__LOGO_SRC__", LOGO_SRC))


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    print("=" * 52)
    print("  NeoProprio — Extracteur d'operations certifiees")
    print(f"  http://localhost:{port}")
    print("=" * 52)
    if port == 5000:
        import webbrowser
        threading.Timer(1.5, lambda: webbrowser.open(f"http://localhost:{port}")).start()
    app.run(host="0.0.0.0", port=port, debug=False, threaded=True)