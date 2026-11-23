"""
Generador Unificado de Reporte Diario de NOTAMs ANS (DGAC / CTNAV)
- Consulta en vivo https://ctnav.pe/notam/api.php?limit=9999 y valida vigencia exacta (endDate > now_utc).
- Genera un solo PDF unificado (YYMMDD REPORTE NOTAMS ANS.pdf):
    * Pagina 1: Plano de Aerodromo LIMA (SPJC) - Dia y Noche (Hora Local UTC-5)
    * Pagina 2: Plano de Aerodromo CUSCO (SPZO) - Pista y calles de rodaje (Hora Local UTC-5)
    * Pagina 3+: Tabla de Ocurrencias Operacionales (LIMA, CUSCO, INTERNACIONALES, NACIONALES)
                 con encabezado oficial DGAC, vinculos clicables al NOTAM de referencia,
                 texto en espanol, DESDE/HASTA en hora local, OBSERVACIONES (relaciones),
                 NOTAM INICIAL y DIAS transcurridos desde primera publicacion.
    * Paginas finales: Anexo de NOTAMs de Referencia (Texto Integro OACI) con anclas clicables.
- Actualiza REPORTE NOTAMS ANS.xlsx y genera YYMMDD REPORTE NOTAMS WHATSAPP.txt (sin area de movimiento SPJC/SPZO).
- Sincroniza todo en local y Google Drive (H:\\Mi unidad\\DGAC\\REPORTE NOTAMS).
"""

import base64
import html
import json
import os
import re
import shutil
import subprocess
import urllib.request
from datetime import date, datetime, timedelta, timezone

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ASSETS_DIR = os.path.join(BASE_DIR, "assets")
LOCAL_OUT_DIR = os.path.join(BASE_DIR, "REPORTE NOTAMS")
DRIVE_OUT_DIR = r"H:\Mi unidad\DGAC\REPORTE NOTAMS"
CHROME_EXE = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
EDGE_EXE = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"


def find_browser_exe():
    for candidate in (
        CHROME_EXE,
        EDGE_EXE,
        shutil.which("google-chrome"),
        shutil.which("chromium"),
        shutil.which("chromium-browser"),
    ):
        if candidate and os.path.exists(candidate):
            return candidate
    raise RuntimeError("No se encontró Google Chrome ni Microsoft Edge para renderizar el PDF.")

LIMA_TZ = timezone(timedelta(hours=-5))


def img_to_data_uri(path):
    with open(path, "rb") as f:
        b64 = base64.b64encode(f.read()).decode("ascii")
    return f"data:image/png;base64,{b64}"


def safe_sync(src, dst):
    if not os.path.isdir(os.path.dirname(dst)):
        return False
    try:
        shutil.copy2(src, dst)
        return True
    except PermissionError:
        print(f"  [AVISO] No se pudo sobrescribir '{dst}' porque está abierto en otro programa. Ciérralo para sincronizar.")
        return False


def upload_via_apps_script_webhook(file_paths):
    """Si existe DRIVE_WEBHOOK_URL (ej. en GitHub Actions), sube los archivos directo a Google Drive."""
    webhook_url = os.environ.get("DRIVE_WEBHOOK_URL", "").strip()
    if not webhook_url:
        return False
    secret = os.environ.get("DRIVE_WEBHOOK_SECRET", "DGAC_CTNAV_NOTAM_2026").strip()
    payload_files = []
    for path, mime in file_paths:
        with open(path, "rb") as f:
            payload_files.append({
                "name": os.path.basename(path),
                "mimeType": mime,
                "content_b64": base64.b64encode(f.read()).decode("ascii"),
            })
    body = json.dumps({
        "secret": secret,
        "action": "sync_report",
        "files": payload_files,
    }).encode("utf-8")
    req = urllib.request.Request(
        webhook_url,
        data=body,
        headers={"Content-Type": "application/json"},
        method="POST",
    )
    with urllib.request.urlopen(req, timeout=60) as resp:
        res_txt = resp.read().decode("utf-8", errors="replace")
        print(f"Sincronización en la nube (Apps Script): {resp.status} -> {res_txt[:200]}")
    return True


def limpiar_archivos_antiguos(today_date, dias_retencion=7):
    """
    Elimina automáticamente archivos diarios de reporte con más de `dias_retencion` días
    desde su fecha de emisión (YYMMDD) en las carpetas local, Drive y HTML temporal.
    Nunca toca archivos maestros (.xlsx, .py, .gs, .md, .json, assets).
    """
    cutoff_date = today_date - timedelta(days=dias_retencion)
    pat_report = re.compile(r"^(\d{6})\s+REPORTE NOTAMS.*\.(pdf|txt|gdoc)$", re.IGNORECASE)
    pat_html = re.compile(r"^(\d{6})_reporte_unificado\.html$", re.IGNORECASE)
    eliminados = 0

    for folder, pat in [(LOCAL_OUT_DIR, pat_report), (DRIVE_OUT_DIR, pat_report), (BASE_DIR, pat_html)]:
        if not os.path.isdir(folder):
            continue
        for fname in os.listdir(folder):
            m = pat.match(fname)
            if not m:
                continue
            try:
                fdate = datetime.strptime(m.group(1), "%y%m%d").date()
            except ValueError:
                continue
            if fdate < cutoff_date:
                fpath = os.path.join(folder, fname)
                try:
                    os.remove(fpath)
                    eliminados += 1
                except Exception as e:
                    print(f"  [AVISO] No se pudo eliminar '{fpath}': {e}")

    print(f"Limpieza automática (> {dias_retencion} días, corte < {cutoff_date}): {eliminados} archivos eliminados.")
    return eliminados


def fetch_strictly_active_notams():
    """Descarga NOTAMs desde CTNAV API y filtra estrictamente los que siguen vigentes en este instante UTC."""
    req = urllib.request.Request(
        "https://ctnav.pe/notam/api.php?limit=9999",
        headers={"User-Agent": "Mozilla/5.0"},
    )
    with urllib.request.urlopen(req, timeout=30) as resp:
        data = json.loads(resp.read().decode("utf-8"))

    now_utc = datetime.now(timezone.utc)
    active = {}
    filtered_expired = []
    for n in data.get("notams", []):
        nid = str(n.get("notamNumber", "")).strip()
        end_str = str(n.get("endDate", "")).strip()
        if end_str and end_str.upper() != "PERM":
            try:
                end_dt = datetime.strptime(end_str[:19], "%Y-%m-%d %H:%M:%S").replace(tzinfo=timezone.utc)
                if end_dt <= now_utc:
                    filtered_expired.append((nid, end_str))
                    continue
            except Exception:
                pass
        active[nid] = n

    return active, now_utc, data.get("last_sync", ""), filtered_expired


def parse_utc(s):
    if not s or str(s).strip().upper() == "PERM":
        return None
    return datetime.strptime(str(s).strip()[:19], "%Y-%m-%d %H:%M:%S").replace(tzinfo=timezone.utc)


def fmt_local(utc_str):
    dt = parse_utc(utc_str)
    if dt is None:
        return "PERM"
    return dt.astimezone(LIMA_TZ).strftime("%d/%m/%y %H:%M")


def fmt_local_short(utc_str):
    dt = parse_utc(utc_str)
    if dt is None:
        return "PERM"
    return dt.astimezone(LIMA_TZ).strftime("%d/%m %H:%M")


def is_red_notam(n_obj, today_local_date, cutoff_local_dt):
    """
    Regla para resaltar en rojo (#C53929):
    - Publicado desde el corte anterior (ayer 18:00 local, o viernes 18:00 si hoy es lunes) o iniciado hoy, O
    - Vence hoy (en hora local).
    """
    # issueDate de CTNAV API ya viene en hora local Lima
    iss_str = str(n_obj.get("issueDate", "")).strip()
    if iss_str:
        try:
            iss_dt = datetime.strptime(iss_str[:19], "%Y-%m-%d %H:%M:%S").replace(tzinfo=LIMA_TZ)
            if iss_dt >= cutoff_local_dt:
                return True
        except Exception:
            pass

    start_dt = parse_utc(n_obj.get("startDate"))
    if start_dt and start_dt.astimezone(LIMA_TZ).date() == today_local_date:
        return True

    end_dt = parse_utc(n_obj.get("endDate"))
    if end_dt and end_dt.astimezone(LIMA_TZ).date() == today_local_date:
        return True

    return False


def anchor_id(nid):
    return "notam-" + nid.replace("/", "-")


def compactar_dias_consecutivos(dias_por_mes):
    """
    Regla permanente: cuando un NOTAM lista múltiples fechas programadas,
    resumir siempre los días consecutivos con el formato 'del XX al XX'.
    """
    bloques_mes = []
    for mes, dias in dias_por_mes.items():
        ordenados = sorted(set(dias))
        if not ordenados:
            continue
        rangos = []
        inicio = prev = ordenados[0]
        for d in ordenados[1:]:
            if d == prev + 1:
                prev = d
            else:
                rangos.append((inicio, prev))
                inicio = prev = d
        rangos.append((inicio, prev))

        partes = [f"del {a} al {b}" if b > a else str(a) for a, b in rangos]
        if len(partes) == 1:
            txt_dias = partes[0]
        else:
            txt_dias = ", ".join(partes[:-1]) + " y " + partes[-1]
        bloques_mes.append(f"{mes}: {txt_dias}")
    return "; ".join(bloques_mes)


INTL_ICAOS = {"SPHI", "SPCL", "SPTU", "SPQU", "SPUR", "SPQT", "SPTN", "SPRU", "SPJL"}
LIMA_ICAOS = {"SPJC", "SPIM", "SPLP"}
CUSCO_ICAOS = {"SPZO"}

AIRPORT_DISPLAY = {
    "SPJC": "LIMA AEROPUERTO",
    "SPIM": "LIMA ACC",
    "SPLP": "LAS PALMAS",
    "SPZO": "CUSCO",
    "SPHI": "CHICLAYO",
    "SPCL": "PUCALLPA",
    "SPTU": "PUERTO MALDONADO",
    "SPST": "TARAPOTO",
    "SPAS": "ANDOAS",
    "SPHY": "ANDAHUAYLAS",
    "SPGM": "TINGO MARÍA",
    "SPNC": "HUÁNUCO",
    "SPHO": "AYACUCHO",
    "SPQU": "AREQUIPA",
    "SPUR": "PIURA",
    "SPQT": "IQUITOS",
    "SPTN": "TACNA",
    "SPRU": "TRUJILLO",
    "SPJL": "JULIACA",
    "SPJA": "JAÉN",
    "SPJR": "CAJAMARCA",
    "SPHZ": "ANTA / HUARAZ",
    "SPJI": "JUANJUÍ",
    "SPME": "TUMBES",
    "SPSO": "PISCO",
    "SPJJ": "JAUJA",
    "SPYL": "TALARA",
    "SPLO": "ILO",
    "SPHU": "HUANUCO",
}


def load_indicadores_ctnav():
    """Carga el catálogo oficial Indicadores de Lugar CTNAV.xlsx -> {icao: (nombre, distintivo)}."""
    path = os.path.join(LOCAL_OUT_DIR, "Indicadores de Lugar CTNAV.xlsx")
    indicadores = {}
    if not os.path.exists(path):
        return indicadores
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    for row in ws.iter_rows(min_row=6, values_only=True):
        if not row or len(row) < 5:
            continue
        icao = str(row[2] or "").strip().upper()
        distintivo = str(row[3] or "").strip()
        nombre = str(row[4] or "").strip()
        if len(icao) == 4 and icao.startswith("SP"):
            indicadores[icao] = (nombre, distintivo)
    wb.close()
    return indicadores


def extract_item_e(icao_msg):
    m = re.search(r"\bE\)\s*(.*?)(?=\s*\b[FG]\)|$)", str(icao_msg), re.DOTALL)
    if m:
        return re.sub(r"\s+", " ", m.group(1)).strip()
    return re.sub(r"\s+", " ", str(icao_msg)).strip()


def extract_item_d(icao_msg):
    m = re.search(r"\bD\)\s*(.*?)(?=\s*\bE\)|$)", str(icao_msg), re.DOTALL)
    if m:
        return re.sub(r"\s+", " ", m.group(1)).strip()
    return ""


def extract_replaced_notam(icao_msg):
    m = re.search(r"\bNOTAMR\s+([A-Z]\d{4}/\d{2})", str(icao_msg))
    return m.group(1) if m else None


def is_movement_area_spjc_spzo(icao, msg_e):
    if icao not in ("SPJC", "SPZO"):
        return False
    u = msg_e.upper()
    # Excluir L2 / obstaculos en trayectoria, TSA, SPR91, 121.5, A-CDM
    if any(k in u for k in ("121.5", "LINEA 2", "CLIMB GRADIENT", "GRADIENTE", "SPR91", "AIR DISPLAY", "PARACHUTE", "TSA ", "GUN FIRING", "A-CDM")):
        return False
    return any(k in u for k in ("RWY", "TWY", "TXL", "HOLDING", "TURN", "180 DEG", "TORA", "TODA", "ASDA", "LDA", "STRIP", "MARKING", "PAVEMENT", "CROSSING", "SEGREGATED"))


def is_ans_reportable(n_obj, indicadores):
    """Aplica las 14 reglas para filtrar solo NOTAMs ANS reportables."""
    icao = str(n_obj.get("facilityDesignator", "")).strip().upper()
    msg = str(n_obj.get("icaoMessage", "")).upper()
    e_txt = extract_item_e(msg).upper()

    # Regla 3: horarios de servicio AIP
    if any(k in e_txt for k in ("HR SER", "SER HR", "AVBL O/R", "OPR HR", "AD HR", "ATS HR", "TWR HR")):
        return False
    # Regla 9: Costa Verde aerodeportiva
    if any(k in e_txt for k in ("COSTA VERDE", "PARAPENTE", "PARATRIKE", "AERODEPORTIVA")):
        return False
    # Regla 12: SEI / RFFS salvo afectacion directa ANS
    if any(k in e_txt for k in ("RFFS", "FIRE FIGHTING", "RESCUE AND FIRE", "SEI ")) and not any(k in e_txt for k in ("RWY", "TWY", "ILS", "VOR", "DME", "ATS", "TWR")):
        return False
    # Regla 13: Puestos de estacionamiento (PEA / ACFT STAND)
    if any(k in e_txt for k in ("ACFT STAND", "PARKING STAND", "PEA ", "STANDS ")) and "TXL" not in e_txt:
        return False
    # Ruido administrativo / publicaciones / RPAS / globos / HLP / grass cutting
    if any(k in e_txt for k in ("CHECKLIST", "TRIGGER NOTAM", "AIP AMDT", "REF AIP", "WEF ", "SUP AIP")) and "A-CDM" not in e_txt:
        return False
    if any(k in e_txt for k in ("RPAS", "DRONE", "PROBE BALLOON", "GRASS CUTTING", "VIRAJE LENTO", "REF RD ", "PCN ")):
        return False
    if e_txt.startswith("HLP ") or "HELIPUERTO" in e_txt:
        return False
    # Obstaculos menores fuera de pista/gradiente (ej. gruas urbanas lejos del AD)
    if e_txt.startswith("OBST ") and "RWY" not in e_txt and "GRADIENT" not in e_txt:
        return False
    # Regla 4: Aerodromos cerrados solo si tienen distintivo de llamada en CTNAV
    if ("AD CLSD" in e_txt or "AD CERRADO" in e_txt) and icao not in INTL_ICAOS and icao not in LIMA_ICAOS and icao not in CUSCO_ICAOS:
        _, distintivo = indicadores.get(icao, ("", ""))
        if not distintivo:
            return False
    # SPLP solo si es relevante; SPR91/paracaidismo de Lima ya salen por SPJC/SPIM
    if icao == "SPLP" and not any(k in e_txt for k in ("AD CLSD", "RWY")):
        return False

    # Solo incluir NOTAMs publicados recientemente (>= 2026-11-20 18:00) o con patron operacional ANS clave
    iss_str = str(n_obj.get("issueDate", "")).strip()
    is_recent = iss_str >= "2026-11-20 18:00:00"
    is_core_ans = any(k in e_txt for k in (
        "RWY", "TWY", "ILS", "VOR", "DME", "APAPI", "PAPI", "BIRD HAZARD",
        "VENTEO DE GAS", "TSA ", "PARACHUTE", "AIR DISPLAY", "GUN FIRING",
        "IAC ", "SID ", "STAR ", "GRADIENT", "AD CLSD"
    ))
    return is_recent and is_core_ans


def utc_hhmm_to_local_str(hhmm):
    h, m = int(hhmm[:2]), int(hhmm[2:4])
    lh = (h - 5) % 24
    return f"{lh:02d}:{m:02d}"


def format_d_schedule_local(d_str, start_utc_str):
    """Convierte campo D (fechas y horarios UTC) a texto limpio en hora local con 'del XX al XX'."""
    if not d_str:
        return ""
    u = d_str.upper()
    meses_en = {
        "JAN": "enero", "FEB": "febrero", "MAR": "marzo", "APR": "abril",
        "MAY": "mayo", "JUN": "junio", "JUL": "julio", "AUG": "agosto",
        "SEP": "septiembre", "OCT": "octubre", "NOV": "noviembre", "DEC": "diciembre",
    }
    # Detectar rango horario HHMM-HHMM
    hm = re.search(r"(\d{4})-(\d{4})", u)
    shift_day = 0
    hora_txt = ""
    if hm:
        h1 = int(hm.group(1)[:2])
        if h1 < 5:
            shift_day = -1
        hora_txt = f"de {utc_hhmm_to_local_str(hm.group(1))} a {utc_hhmm_to_local_str(hm.group(2))}"

    # Detectar meses y dias si existen
    tokens = re.findall(r"(JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC|\d{2})", u.split(hm.group(0))[0] if hm else u)
    dias_por_mes = {}
    cur_mes = None
    year = int(str(start_utc_str)[:4]) if start_utc_str else 2026
    month_num_map = {k: i + 1 for i, k in enumerate(["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"])}
    for tok in tokens:
        if tok in meses_en:
            cur_mes = tok
        elif tok.isdigit() and cur_mes:
            d_num = int(tok)
            dt_utc = date(year, month_num_map[cur_mes], d_num) + timedelta(days=shift_day)
            m_es = list(meses_en.values())[dt_utc.month - 1]
            dias_por_mes.setdefault(m_es, []).append(dt_utc.day)

    fechas_txt = compactar_dias_consecutivos(dias_por_mes) if dias_por_mes else ""
    if fechas_txt and hora_txt:
        return f"{hora_txt} en fechas programadas ({fechas_txt})"
    return hora_txt or fechas_txt


def translate_icao_to_spanish(n_obj):
    """Traduce automáticamente un NOTAM nuevo de OACI (E + D) a español claro para la tabla y WhatsApp."""
    e_raw = extract_item_e(n_obj.get("icaoMessage", ""))
    d_raw = extract_item_d(n_obj.get("icaoMessage", ""))
    sched = format_d_schedule_local(d_raw, n_obj.get("startDate", ""))
    u = e_raw.upper()

    # Patrones frecuentes CTNAV
    m_rwy = re.match(r"RWY\s+([0-9/LRC]+)\s+CLSD(?:\s+DUE\s+TO\s+WIP)?", u)
    if m_rwy:
        base = f"Pista {m_rwy.group(1)} cerrada por trabajos de mantenimiento" if "WIP" in u else f"Pista {m_rwy.group(1)} cerrada"
        return f"{base} {sched}.".strip() if sched else f"{base}."

    m_twy = re.match(r"TWY\s+([A-Z0-9,\s/]+?)\s+CLSD(?:\s+DUE\s+TO\s+WIP)?", u)
    if m_twy:
        base = f"Calles de rodaje {m_twy.group(1).strip()} cerradas por trabajos de mantenimiento" if "WIP" in u else f"Calles de rodaje {m_twy.group(1).strip()} cerradas"
        return f"{base} {sched}.".strip() if sched else f"{base}."

    m_apapi = re.match(r"(PAPI|APAPI)\s+RWY\s+([0-9/LRC]+)\s+INOP", u)
    if m_apapi:
        return f"{m_apapi.group(1)} de pista {m_apapi.group(2)} inoperativo."

    m_bird = re.match(r"BIRD\s+HAZARD\s+RWY\s+([0-9/LRC]+)", u)
    if m_bird:
        return f"Peligro de aves en pista {m_bird.group(1)}."

    if "AD CLSD" in u:
        return "Aeródromo cerrado de forma permanente." if str(n_obj.get("endDate", "")).upper() == "PERM" else "Aeródromo cerrado."

    # Si ya viene en español (ej. VENTEO DE GAS), limpiar coordenadas largas y convertir horario si aplica
    txt = e_raw
    txt = re.sub(r"\s*COORD:.*$", "", txt, flags=re.IGNORECASE).strip()
    txt = re.sub(r"\s*F\)G.*$", "", txt, flags=re.IGNORECASE).strip()
    if txt.isupper():
        txt = txt.capitalize()
    if sched and sched not in txt:
        txt = f"{txt.rstrip('.')} ({sched})."
    elif not txt.endswith("."):
        txt += "."
    return txt


def build_spjc_svg_base_labels():
    """Etiquetas limpias sobre el plano base SPJC (viewBox 0 0 1800 1200)."""
    # Coordenadas calibradas del PDF oficial (norm_x*1.8, norm_y*1.2)
    labels = [
        ("16L", 124.9, 424.2, "rwy-end"),
        ("34R", 912.3, 143.9, "rwy-end"),
        ("16R", 53.7, 916.7, "rwy-end"),
        ("34L", 837.0, 590.9, "rwy-end"),
        ("RWY 16L / 34R", 445.6, 318.2, "rwy-name"),
        ("RWY 16R / 34L", 435.5, 765.2, "rwy-name"),
        ("L", 214.9, 507.6, "twy"),
        ("L1", 117.8, 500.0, "twy"),
        ("L3", 309.8, 409.1, "twy"),
        ("L4", 537.0, 301.5, "twy"),
        ("L5", 606.7, 272.7, "twy"),
        ("L7", 924.9, 128.8, "twy"),
        ("A", 254.9, 303.0, "twy"),
        ("B", 396.3, 290.9, "twy"),
        ("C", 546.8, 233.3, "twy"),
        ("D", 593.3, 207.6, "twy"),
        ("E", 659.1, 183.3, "twy"),
        ("F", 815.9, 69.7, "twy"),
        ("G", 714.3, 150.0, "twy"),
        ("Q", 419.3, 540.9, "twy"),
        ("R", 449.8, 533.3, "twy"),
        ("R1", 423.5, 489.4, "twy"),
        ("R2", 443.7, 590.9, "twy"),
        ("M", 478.7, 389.4, "twy"),
        ("M1", 535.2, 354.5, "twy"),
        ("M2", 634.2, 304.5, "twy"),
        ("M3", 694.8, 278.8, "twy"),
        ("M4", 783.7, 239.4, "twy"),
        ("U", 524.6, 621.2, "twy"),
        ("U1", 588.1, 610.6, "twy"),
        ("U2", 689.1, 578.8, "twy"),
        ("U3", 756.8, 548.5, "twy"),
        ("V", 535.9, 647.0, "twy"),
        ("V1", 13.6, 900.0, "twy"),
        ("V2", 159.1, 850.0, "twy"),
        ("V3", 616.6, 648.5, "twy"),
        ("V5", 828.7, 551.5, "twy"),
        ("T BLUE", 641.0, 553.0, "apn"),
        ("TERMINAL", 645.1, 484.8, "term"),
        ("TERMINAL 1", 624.9, 36.4, "term"),
        ("APN 100s", 725.1, 268.2, "apn"),
        ("APN 200s", 745.3, 497.0, "apn"),
        ("APN 300s", 522.1, 474.2, "apn"),
        ("APN 400s", 543.3, 540.9, "apn"),
        ("APN 500s", 548.4, 384.8, "apn"),
        ("APN 600s", 586.7, 571.2, "apn"),
    ]
    out = []
    for txt, nx, ny, cls in labels:
        x, y = nx * 1.8, ny * 1.2
        out.append(f'<text x="{x:.1f}" y="{y:.1f}" class="lbl-{cls}">{html.escape(txt)}</text>')
    return "\n".join(out)


def build_spjc_proc_pins():
    """Pines azules de procedimientos P1, P2, P3 sobre SPJC."""
    pins = [
        ("P1", 136 * 1.8, 514 * 1.2, 22, 6),
        ("P2", 615 * 1.8, 228 * 1.2, 18, -10),
        ("P2", 603 * 1.8, 191 * 1.2, 18, -10),
        ("P3", 107 * 1.8, 434 * 1.2, -32, -8),
        ("P3", 928 * 1.8, 108 * 1.2, 16, -6),
    ]
    out = []
    for lbl, x, y, dx, dy in pins:
        out.append(
            f'<circle cx="{x:.1f}" cy="{y:.1f}" r="11" fill="#1d4ed8" stroke="#ffffff" stroke-width="3"/>'
            f'<text x="{x+dx:.1f}" y="{y+dy:.1f}" class="pin-lbl">{lbl}</text>'
        )
    return "\n".join(out)


def render_spjc_day_svg(active, bg_uri):
    overlays = []
    # A3730/26: RWY 16L/34R no disponible 12:30-14:30 y 16:00-17:30 local
    if "A3730/26" in active:
        overlays.append(
            '<line x1="186" y1="548" x2="1658" y2="124" stroke="#c53929" stroke-width="26" stroke-opacity="0.45" stroke-linecap="butt"/>'
            '<line x1="186" y1="548" x2="1658" y2="124" stroke="#c53929" stroke-width="6" stroke-dasharray="18,10"/>'
            '<line x1="510" y1="0" x2="620" y2="415" stroke="#c53929" stroke-width="2.2"/>'
        )
    # A3513/26: TXL T BLUE limitada entre 601 y 202 (07:00-15:00 local)
    if "A3513/26" in active:
        overlays.append(
            '<circle cx="1162" cy="672" r="14" fill="#d97706" stroke="#ffffff" stroke-width="3.5"/>'
            '<circle cx="1162" cy="672" r="24" fill="none" stroke="#d97706" stroke-width="2.5" stroke-dasharray="6,4"/>'
            '<line x1="1270" y1="0" x2="1166" y2="652" stroke="#b45309" stroke-width="2.0"/>'
        )
    # A2863/26: TWY L entre L5 y L7 no disp. CAT D/E
    if "A2863/26" in active:
        overlays.append(
            '<line x1="1125" y1="340" x2="1665" y2="178" stroke="#d97706" stroke-width="10" stroke-opacity="0.55" stroke-dasharray="12,8"/>'
        )
    # A3527/26: Senalizacion puntos de espera RWY 16L no disp. en TWY B, C, D, E, G
    if "A3527/26" in active:
        for bx, by in [(705, 342), (975, 272), (1055, 248), (1185, 212), (1285, 182)]:
            overlays.append(f'<circle cx="{bx}" cy="{by}" r="7" fill="#f59e0b" stroke="#78350f" stroke-width="2"/>')

    return f"""
    <svg viewBox="0 0 1800 1200" class="map-svg" xmlns="http://www.w3.org/2000/svg">
      <image href="{bg_uri}" x="0" y="0" width="1800" height="1200"/>
      {''.join(overlays)}
      {build_spjc_proc_pins()}
      {build_spjc_svg_base_labels()}
    </svg>
    """


def render_spjc_night_svg(active, bg_uri):
    overlays = []
    # A3696/26: RWY 16R/34L cerrada 01:30-04:30 local
    if "A3696/26" in active:
        overlays.append(
            '<line x1="38" y1="1132" x2="1510" y2="704" stroke="#c53929" stroke-width="24" stroke-opacity="0.48"/>'
            '<line x1="38" y1="1132" x2="1510" y2="704" stroke="#c53929" stroke-width="5"/>'
            '<line x1="180" y1="1200" x2="480" y2="1004" stroke="#c53929" stroke-width="2.2"/>'
        )
    # A3695/26: Franja RWY 16R/34L obras 01:30-04:30 local
    if "A3695/26" in active:
        overlays.append(
            '<line x1="32" y1="1102" x2="1502" y2="674" stroke="#d97706" stroke-width="5" stroke-dasharray="14,10"/>'
            '<line x1="46" y1="1162" x2="1518" y2="734" stroke="#d97706" stroke-width="5" stroke-dasharray="14,10"/>'
        )
    # TWY M (A3722/26) & M1-M4 (A3723, A3721, A3720, A3719)
    if "A3722/26" in active or "A3723/26" in active:
        overlays.append(
            '<line x1="822" y1="472" x2="1445" y2="288" stroke="#c53929" stroke-width="11" stroke-opacity="0.45" stroke-dasharray="14,8"/>'
        )
    m_pins = [
        ("A3723/26", 965, 428, 160, 0),
        ("A3722/26", 870, 456, 460, 0),
        ("A3721/26", 1142, 372, 760, 0),
        ("A3720/26", 1255, 338, 1060, 0),
        ("A3719/26", 1412, 292, 1640, 0),
    ]
    for nid, px, py, tx, ty in m_pins:
        if nid in active:
            overlays.append(
                f'<line x1="{tx}" y1="{ty}" x2="{px}" y2="{py}" stroke="#c53929" stroke-width="1.8" stroke-opacity="0.75"/>'
                f'<circle cx="{px}" cy="{py}" r="10" fill="#c53929" stroke="#ffffff" stroke-width="2.5"/>'
            )
    if "A3695/26" in active:
        overlays.append('<line x1="1350" y1="0" x2="1220" y2="758" stroke="#d97706" stroke-width="1.8" stroke-opacity="0.75"/>')

    # TWY U / U1 / U2 / U3 / V3 / V5 (A3725, A3718, A3717, A3729)
    if "A3725/26" in active or "A3729/26" in active:
        overlays.append(
            '<line x1="840" y1="775" x2="1465" y2="592" stroke="#c53929" stroke-width="11" stroke-opacity="0.45" stroke-dasharray="14,8"/>'
        )
    u_pins = [
        ("A3725/26", 1058, 725, 540, 1200),
        ("A3718/26", 1235, 675, 900, 1200),
        ("A3717/26", 1365, 636, 1260, 1200),
        ("A3729/26", 1435, 616, 1620, 1200),
    ]
    for nid, px, py, tx, ty in u_pins:
        if nid in active:
            overlays.append(
                f'<line x1="{tx}" y1="{ty}" x2="{px}" y2="{py}" stroke="#c53929" stroke-width="1.8" stroke-opacity="0.75"/>'
                f'<circle cx="{px}" cy="{py}" r="10" fill="#c53929" stroke="#ffffff" stroke-width="2.5"/>'
            )
    # V3 (A3729/26) & V5 (A3725/26) inline pins
    if "A3729/26" in active:
        overlays.append(
            '<circle cx="1115" cy="795" r="9" fill="#c53929" stroke="#ffffff" stroke-width="2.5"/>'
        )
    if "A3725/26" in active:
        overlays.append(
            '<circle cx="1485" cy="680" r="9" fill="#c53929" stroke="#ffffff" stroke-width="2.5"/>'
        )

    return f"""
    <svg viewBox="0 0 1800 1200" class="map-svg" xmlns="http://www.w3.org/2000/svg">
      <image href="{bg_uri}" x="0" y="0" width="1800" height="1200"/>
      {''.join(overlays)}
      {build_spjc_proc_pins()}
      {build_spjc_svg_base_labels()}
    </svg>
    """


def render_spzo_svg(active, bg_uri):
    overlays = []
    # A3740/26 (o A3712/26): Prohibido giro 180 en RWY 10/28
    if "A3740/26" in active:
        overlays.append(
            '<line x1="75" y1="68" x2="1868" y2="75" stroke="#d97706" stroke-width="22" stroke-opacity="0.38"/>'
            '<line x1="75" y1="68" x2="1868" y2="75" stroke="#c53929" stroke-width="4" stroke-dasharray="16,10"/>'
            '<line x1="330" y1="0" x2="330" y2="68" stroke="#c53929" stroke-width="2.5"/>'
        )
    # A3707/26: Parte de TWY C cerrada hasta antes de D
    if "A3707/26" in active:
        overlays.append(
            '<line x1="455" y1="152" x2="735" y2="152" stroke="#c53929" stroke-width="18" stroke-opacity="0.48" stroke-linecap="round"/>'
            '<line x1="455" y1="152" x2="735" y2="152" stroke="#c53929" stroke-width="4" stroke-dasharray="12,7"/>'
            '<line x1="1000" y1="0" x2="595" y2="144" stroke="#c53929" stroke-width="2.2"/>'
        )
    # A3711/26: Depresion de pavimento inicio de TWY C
    if "A3711/26" in active:
        overlays.append(
            '<circle cx="715" cy="152" r="10" fill="#d97706" stroke="#ffffff" stroke-width="2.5"/>'
            '<line x1="1660" y1="0" x2="723" y2="148" stroke="#b45309" stroke-width="2.0"/>'
        )

    labels = [
        ("10", 45.6 * 2.0, 198.1 * 0.29, "rwy-end"),
        ("28", 934.5 * 2.0, 198.1 * 0.29, "rwy-end"),
        ("RWY 10 / 28", 669.5 * 2.0, 204.0 * 0.29, "rwy-name"),
        ("D", 203.7 * 2.0, 380.5 * 0.29, "twy"),
        ("C", 263.5 * 2.0, 492.2 * 0.29, "twy"),
        ("B", 392.6 * 2.0, 380.5 * 0.29, "twy"),
        ("A", 469.5 * 2.0, 386.4 * 0.29, "twy"),
        ("TERMINAL", 390.9 * 2.0, 798.1 * 0.29, "term"),
    ]
    lbl_svg = "\n".join(
        f'<text x="{x:.1f}" y="{y:.1f}" class="lbl-spzo-{cls}">{html.escape(txt)}</text>'
        for txt, x, y, cls in labels
    )
    return f"""
    <svg viewBox="0 0 2000 290" class="spzo-svg" xmlns="http://www.w3.org/2000/svg">
      <image href="{bg_uri}" x="0" y="0" width="2000" height="290"/>
      {''.join(overlays)}
      {lbl_svg}
    </svg>
    """


def card_link(nid, title, body, sched, tone="red", badge=None):
    b_html = f'<span class="badge-new">{html.escape(badge)}</span>' if badge else ""
    return f"""
    <div class="mcard mcard-{tone}">
      <div class="mcard-head">
        <a href="#{anchor_id(nid)}" class="notam-link">{html.escape(nid)}</a>
        <span class="mcard-title">· {html.escape(title)}</span>
        {b_html}
      </div>
      <div class="mcard-body">{html.escape(body)}</div>
      <div class="mcard-sched">{html.escape(sched)}</div>
    </div>
    """


def build_unified_report():
    active, now_utc, last_sync, filtered_expired = fetch_strictly_active_notams()
    now_local = now_utc.astimezone(LIMA_TZ)
    today_date = now_local.date()
    yymmdd = now_local.strftime("%y%m%d")
    fecha_ddmmyy = now_local.strftime("%d/%m/%y")
    meses = [
        "", "enero", "febrero", "marzo", "abril", "mayo", "junio",
        "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre",
    ]
    fecha_larga = f"{today_date.day:02d} de {meses[today_date.month]} del {today_date.year}"

    # Corte de novedades: ayer 18:00 local (o viernes 18:00 si hoy es lunes)
    days_back = 3 if today_date.weekday() == 0 else 1
    cutoff_local_dt = datetime.combine(today_date - timedelta(days=days_back), datetime.min.time(), tzinfo=LIMA_TZ).replace(hour=18, minute=0)

    print(f"Generando reporte {yymmdd} ({fecha_ddmmyy}) | Hora local: {now_local.strftime('%H:%M:%S')} | Activos verificados: {len(active)} | Vencidos filtrados: {len(filtered_expired)}")

    # Cargar imagenes en base64
    logo_uri = img_to_data_uri(os.path.join(ASSETS_DIR, "logo_dgac.png"))
    spjc_bg_uri = img_to_data_uri(os.path.join(ASSETS_DIR, "spjc_map_bg.png"))
    spzo_bg_uri = img_to_data_uri(os.path.join(ASSETS_DIR, "spzo_map_bg.png"))

    # =========================================================================
    # DATOS DE LA TABLA DE OCURRENCIAS OPERACIONALES (CURADOS + DESCUBRIMIENTO EN VIVO)
    # =========================================================================
    indicadores = load_indicadores_ctnav()
    table_specs = [
        # 1. LIMA (Solo ocurrencias no pertenecientes al area de movimiento SPJC)
        (
            "LIMA",
            "LIMA ACC",
            "A2864/26",
            "Entrada de aeronaves monomotor restringida dentro del TMA Lima en radio de 30 NM respecto al VOR JCL de 09:30 a 11:30, 16:00 a 18:00 y 20:00 a 00:00, excepto para VFR dentro de zonas militares SPR68 y SPR68A o a 2000 FT o inferior fuera de trayectorias y rutas instrumentales.",
            "",
            "A4017/25",
            date(2025, 12, 21),
        ),
        (
            "LIMA",
            "LIMA AEROPUERTO",
            "A3647/26",
            "Frecuencia de emergencia 121.500 MHz con cobertura limitada, disponible únicamente dentro de cada TMA.",
            "",
            "A0186/23",
            date(2023, 1, 18),
        ),
        (
            "LIMA",
            "LIMA AEROPUERTO",
            "A2978/26",
            "Precaución por obras de Línea 2 del Metro de Lima con grúas y perforadoras de 28 m de altura a 600 m del extremo de pista 16R; aeronaves en salida deberán mantener gradiente de ascenso mínimo de 4.7%.",
            "",
            "A2179/26",
            date(2026, 9, 11),
        ),
        (
            "LIMA",
            "LIMA AEROPUERTO",
            "A3731/26",
            "Festival aerodeportivo en zona restringida SPR91 de 10:00 a 18:00 desde superficie hasta 2000 FT AMSL (operaciones de aeronaves, helicópteros y RPAS). Contactar a Las Palmas Torre para información de tránsito.",
            "Relacionado con C3916/26 (SPLP), A3730/26 (cierre RWY 16L/34R en plano) y A3705/26 (TSA Las Palmas).",
            "A3703/26",
            date(2026, 11, 17),
        ),
        (
            "LIMA",
            "LIMA AEROPUERTO",
            "A3705/26",
            "Zona temporalmente segregada (TSA) por lanzamiento de paracaidistas sobre aeródromo Las Palmas de 12:30 a 14:30 y de 16:00 a 17:30 desde superficie hasta 8000 FT AMSL. Requiere coordinación previa con Lima ACC.",
            "Relacionado con C3889/26 (SPLP), A3730/26 (cierre RWY 16L/34R en plano) y A3731/26 (festival SPR91).",
            "A3705/26",
            date(2026, 11, 17),
        ),
        (
            "LIMA",
            "LIMA AEROPUERTO",
            "A3732/26",
            "Zona temporalmente segregada (TSA) por ejercicios de tiro en zona Salinas Huacho (entre 3.1 NM y 6.6 NM, radiales 297 a 335 del VOR SLS, desde superficie hasta 6000 FT AGL). Requiere coordinación previa con Lima ACC.",
            "Programación posterior al ejercicio A3642/26 del 20/11.",
            "A3732/26",
            date(2026, 11, 19),
        ),
        # 2. CUSCO (Solo ocurrencias no pertenecientes al area de movimiento SPZO)
        (
            "CUSCO",
            "CUSCO",
            "C2961/26",
            "Fase 1 A-CDM Cusco en vigencia (Ref. SUP AIP 03/26) en horario diario de 06:00 a 18:59; toda tripulación de vuelo debe tomar conocimiento de los procedimientos A-CDM aplicables.",
            "",
            "C2961/26",
            date(2026, 9, 4),
        ),
        # 3. AEROPUERTOS INTERNACIONALES
        (
            "AEROPUERTOS INTERNACIONALES",
            "CHICLAYO",
            "A3745/26",
            translate_icao_to_spanish(active["A3745/26"]) if "A3745/26" in active else "",
            "Asociado a A3748/26 (cierre simultáneo de calles de rodaje).",
            "A3745/26",
            date(2026, 11, 21),
        ),
        (
            "AEROPUERTOS INTERNACIONALES",
            "CHICLAYO",
            "A3748/26",
            "Calles de rodaje A, B, B1, C, D y D1 cerradas por trabajos de mantenimiento de 23:00 a 04:00 en las mismas noches programadas del cierre de pista.",
            "Asociado a A3745/26 (cierre simultáneo de pista 01/19).",
            "A3748/26",
            date(2026, 11, 21),
        ),
        (
            "AEROPUERTOS INTERNACIONALES",
            "PUCALLPA",
            "A3663/26",
            "Pista 02/20 cerrada por trabajos de mantenimiento en horario diario de 23:30 a 05:00.",
            "Asociado a A3665/26 (cierre simultáneo de calles de rodaje A y B).",
            "A2840/26",
            date(2026, 9, 1),
        ),
        (
            "AEROPUERTOS INTERNACIONALES",
            "PUCALLPA",
            "A3665/26",
            "Calles de rodaje A y B cerradas por trabajos de mantenimiento en horario diario de 23:30 a 05:00.",
            "Asociado a A3663/26 (cierre simultáneo de pista 02/20).",
            "A2839/26",
            date(2026, 9, 1),
        ),
        (
            "AEROPUERTOS INTERNACIONALES",
            "PUERTO MALDONADO",
            "C3069/26",
            "Procedimientos de llegada instrumental IAC VOR/DME - ILS/DME pista 01 (descensos Alfa, Bravo, Charlie y Delta) suspendidos de manera permanente por razones operacionales.",
            "",
            "C0978/25",
            date(2025, 3, 12),
        ),
        # 4. AEROPUERTOS NACIONALES
        (
            "AEROPUERTOS NACIONALES",
            "TARAPOTO",
            "C3698/26",
            "Torre de Control sin visibilidad en los primeros 700 m del lado izquierdo de la pista 17 debido a maleza y árboles crecidos.",
            "",
            "C3251/26",
            date(2026, 10, 1),
        ),
        (
            "AEROPUERTOS NACIONALES",
            "ANDOAS",
            "C3814/25",
            "Aeródromo cerrado de forma permanente.",
            "",
            "C0649/23",
            date(2023, 2, 15),
        ),
        (
            "AEROPUERTOS NACIONALES",
            "ANDAHUAYLAS",
            "C3216/26",
            "ILS/DME IAND (109.700 MHz CH 34X) fuera de servicio; como consecuencia, procedimiento de aproximación instrumental IAC RWY 21 suspendido.",
            "",
            "C0074/22",
            date(2022, 1, 7),
        ),
    ]

    # Filtrar estrictamente contra active (vigencia en tiempo real garantizada!)
    verified_table_rows = []
    known_ids = set()
    for sec, ad, nid, txt, obs, n_ini, f_ini in table_specs:
        if nid not in active:
            print(f"  [OMITIDO POR VENCIMIENTO] {nid} ({ad})")
            continue
        known_ids.add(nid)
        n_obj = active[nid]
        desde_str = fmt_local(n_obj.get("startDate"))
        hasta_str = fmt_local(n_obj.get("endDate"))
        dias = max(1, (today_date - f_ini).days) if f_ini >= today_date - timedelta(days=2) else max(0, (today_date - f_ini).days)
        red = is_red_notam(n_obj, today_date, cutoff_local_dt)
        verified_table_rows.append({
            "section": sec,
            "aeropuerto": ad,
            "notam": nid,
            "texto": txt,
            "desde": desde_str,
            "hasta": hasta_str,
            "obs": obs,
            "notam_ini": n_ini,
            "dias": dias,
            "is_red": red,
            "raw": n_obj,
        })

    # Descubrimiento automático en vivo de cualquier NOTAM ANS nuevo publicado hoy o en adelante
    pre_mapped_planos = {
        "A3730/26", "A3513/26", "A3723/26", "A3722/26", "A3721/26", "A3720/26",
        "A3695/26", "A3719/26", "A3696/26", "A3725/26", "A3718/26", "A3717/26",
        "A3729/26", "A3523/26", "A2828/26", "A2830/26", "A2827/26", "A3424/26",
        "A3341/26", "A3527/26", "A2863/26", "A3740/26", "A3707/26", "A3711/26",
    }
    extra_spjc_cards = []
    extra_spzo_cards = []
    for nid, n_obj in sorted(active.items(), key=lambda x: (str(x[1].get("facilityDesignator", "")), x[0])):
        if nid in known_ids or nid in pre_mapped_planos:
            continue
        if not is_ans_reportable(n_obj, indicadores):
            continue
        icao = str(n_obj.get("facilityDesignator", "")).strip().upper()
        msg_e = extract_item_e(n_obj.get("icaoMessage", ""))
        if is_movement_area_spjc_spzo(icao, msg_e):
            sched = f"Vig: {fmt_local_short(n_obj.get('startDate'))} - {fmt_local_short(n_obj.get('endDate'))} LOCAL"
            if icao == "SPJC":
                extra_spjc_cards.append((nid, "SPJC Nuevo", translate_icao_to_spanish(n_obj), sched, "red"))
            else:
                extra_spzo_cards.append((nid, "SPZO Nuevo", translate_icao_to_spanish(n_obj), sched, "red"))
            continue

        if icao in LIMA_ICAOS:
            sec = "LIMA"
        elif icao in CUSCO_ICAOS:
            sec = "CUSCO"
        elif icao in INTL_ICAOS:
            sec = "AEROPUERTOS INTERNACIONALES"
        else:
            sec = "AEROPUERTOS NACIONALES"

        ad_name = AIRPORT_DISPLAY.get(icao) or indicadores.get(icao, (icao, ""))[0].upper() or icao
        replaced = extract_replaced_notam(n_obj.get("icaoMessage", ""))
        n_ini = replaced or nid
        iss_str = str(n_obj.get("issueDate", ""))[:10]
        try:
            f_ini = datetime.strptime(iss_str, "%Y-%m-%d").date()
        except Exception:
            f_ini = today_date
        dias = max(1, (today_date - f_ini).days)
        verified_table_rows.append({
            "section": sec,
            "aeropuerto": ad_name,
            "notam": nid,
            "texto": translate_icao_to_spanish(n_obj),
            "desde": fmt_local(n_obj.get("startDate")),
            "hasta": fmt_local(n_obj.get("endDate")),
            "obs": "",
            "notam_ini": n_ini,
            "dias": dias,
            "is_red": is_red_notam(n_obj, today_date, cutoff_local_dt),
            "raw": n_obj,
        })
        print(f"  [NUEVO DETECTADO EN VIVO] {nid} ({ad_name}) -> {sec}")

    # =========================================================================
    # CONSTRUCCION DE TARJETAS PARA PAGINA 1 (LIMA SPJC) Y PAGINA 2 (CUSCO SPZO)
    # =========================================================================
    spjc_day_cards_spec = [
        ("A3730/26", "RWY 16L/34R", "No disponible para aterrizajes ni despegues. RWY 16R/34L en uso. Salidas por 16R con ascenso en rumbo de pista y viraje a la derecha directo a JC921.", "Vig: 19/11 18:20 - 22/11 17:30 · Cierre: 12:30-14:30 y 16:00-17:30 LOCAL", "red"),
        ("A3513/26", "TXL T BLUE", "Rodaje limitado entre puestos 601 y 202 por obras. Precaución en rodaje (punto de referencia del sector).", "Vig: 03/11 07:00 - 13/12 15:00 · Horario: 07:00-15:00 LOCAL", "amber"),
    ]
    spjc_day_cards_html = []
    for nid, title, body, sched, tone in spjc_day_cards_spec + extra_spjc_cards:
        if nid in active:
            badge = "NUEVO" if is_red_notam(active[nid], today_date, cutoff_local_dt) else None
            spjc_day_cards_html.append(card_link(nid, title, body, sched, tone, badge))

    spjc_night_top_spec = [
        ("A3723/26", "TWY M y M1", "M entre R y M1, y M1 cerradas por obras. Prever demoras.", "02/12 - 03/12 · 01:30-04:30", "red"),
        ("A3722/26", "TWY M", "Cerrada por obras. Transferencias M1, M2, M3 y M4 disponibles.", "28/11 - 01/12 · 01:30-04:30", "red"),
        ("A3721/26", "Transf. M2", "Cerrada por obras. Prever demoras.", "04/12 - 10/12 · 01:30-04:30", "red"),
        ("A3720/26", "Transf. M3", "Cerrada por obras. Prever demoras.", "11/12 - 12/12 · 01:30-04:30", "red"),
        ("A3695/26", "Franja 16R/34L", "Obras en la franja de pista (referencia general).", "18/11 - 28/11 · 01:30-04:30", "amber"),
        ("A3719/26", "Transf. M4", "Cerrada por obras. Prever demoras.", "13/12 - 14/12 · 01:30-04:30", "red"),
    ]
    spjc_night_bot_spec = [
        ("A3696/26", "RWY 16R/34L", "Pista cerrada por trabajos de mantenimiento.", "18/11 - 28/11 · 01:30-04:30", "red"),
        ("A3725/26", "U / U1 / V5", "U entre Q y U1 cerrada. U1 no disp. V5 cerrada.", "22/11 - 23/11 · 01:30-04:30", "red"),
        ("A3718/26", "Transf. U2", "Cerrada por trabajos de mantenimiento.", "24/11 - 25/11 · 01:30-04:30", "red"),
        ("A3717/26", "Transf. U3", "Cerrada por trabajos de mantenimiento.", "26/11 - 27/11 · 01:30-04:30", "red"),
        ("A3729/26", "U y V3", "U entre U3 y puesto 212, y V3 cerradas. U2/U3 disp.", "21/11 · 01:30-04:30", "red"),
    ]
    spjc_night_top_html = []
    for nid, title, body, sched, tone in spjc_night_top_spec:
        if nid in active:
            badge = "NUEVO" if is_red_notam(active[nid], today_date, cutoff_local_dt) else None
            spjc_night_top_html.append(card_link(nid, title, body, sched, tone, badge))

    spjc_night_bot_html = []
    for nid, title, body, sched, tone in spjc_night_bot_spec:
        if nid in active:
            badge = "NUEVO" if is_red_notam(active[nid], today_date, cutoff_local_dt) else None
            spjc_night_bot_html.append(card_link(nid, title, body, sched, tone, badge))

    # Procedimientos en superficie SPJC (strip inferior unificado)
    spjc_procs_spec = [
        ("A3523/26", "P1 · TWY L", "Espera 16L en TWY L (L1 no es punto de espera). No cruzar sin autorización ATC.", "Hasta 31/01/27 18:59", "blue"),
        ("A2828/26", "P2 · L5 y D (Bidirec.)", "Pista 16L/34R en uso bidireccional: esperar antes de CAT I y reportar pista libre.", "06:00-18:00 · Hasta 26/11", "blue"),
        ("A2830/26", "P2 · L5 y D (Unidirec.)", "Uso unidireccional oeste a este: esperar antes de CAT I y cruzar antes de reportar libre.", "18:01-05:59 · Hasta 26/11", "blue"),
        ("A2827/26", "P3 · Remolque A-L / F-L7", "Suspende SUP 12/24. Cruce remolque vía A-L y F-L7 esperando antes de CAT I.", "Todo el día · Hasta 26/11", "blue"),
        ("A3424/26", "Segregadas 16R / 16L", "RWY 16R solo aterrizajes y RWY 16L solo despegues (exc. emerg./medevac/HEAD).", "06:00-23:59 · Hasta 24/01/27", "blue"),
        ("A3341/26", "RWY 16R desde V2", "TORA 2773 m desde V2. Salidas de Plat. 1 que no puedan despegar en V2 notificar GND.", "Todo el día · Hasta 20/01/27", "blue"),
        ("A3527/26", "Marcas 16L (B,C,D,E,G)", "Señalización puntos de espera 16L no disp. Reportar libre al ingresar total a A o F.", "Todo el día · Hasta 02/02/27", "amber"),
        ("A2863/26", "TWY L (L5-L7) CAT D/E", "No disponible para rodaje CAT D y E. RWY 34R disponible vía TWY F.", "Todo el día · Hasta 30/11", "amber"),
    ]
    spjc_procs_html = []
    for nid, title, body, sched, tone in spjc_procs_spec:
        if nid in active:
            spjc_procs_html.append(card_link(nid, title, body, sched, tone))

    # Tarjetas CUSCO SPZO (Pagina 2)
    spzo_cards_spec = [
        ("A3740/26", "Pista RWY 10/28", "Prohibido el giro de 180° en pista 10/28 para evitar deterioro de la capa asfáltica (reemplaza a A3712/26; no declara cierre de pista).", f"Vig: {fmt_local_short(active['A3740/26']['startDate'])} - {fmt_local_short(active['A3740/26']['endDate'])} LOCAL · Aplicación continua", "amber") if "A3740/26" in active else None,
        ("A3707/26", "Parte de TWY C", "Cierre por obras hasta antes de TWY D (trazo de sector aproximado; zona señalizada).", f"Vig: {fmt_local_short(active['A3707/26']['startDate'])} - {fmt_local_short(active['A3707/26']['endDate'])} LOCAL · Aplicación continua", "red") if "A3707/26" in active else None,
        ("A3711/26", "Inicio de TWY C", "Precaución por depresión de pavimento al inicio de calle de rodaje C (zona señalizada).", f"Vig: {fmt_local_short(active['A3711/26']['startDate'])} - {fmt_local_short(active['A3711/26']['endDate'])} LOCAL · Aplicación continua", "amber") if "A3711/26" in active else None,
    ]
    spzo_cards_html = []
    for spec in spzo_cards_spec + extra_spzo_cards:
        if spec:
            nid, title, body, sched, tone = spec
            badge = "NUEVO HOY" if is_red_notam(active[nid], today_date, cutoff_local_dt) else None
            spzo_cards_html.append(card_link(nid, title, body, sched, tone, badge))

    # =========================================================================
    # CONSTRUIR FILAS HTML SEPARANDO NOVEDADES (ROJO) Y VIGENTES CONTINUOS
    # =========================================================================
    section_order = ["LIMA", "CUSCO", "AEROPUERTOS INTERNACIONALES", "AEROPUERTOS NACIONALES"]
    novedades_rows = [r for r in verified_table_rows if r["is_red"]]
    vigentes_rows = [r for r in verified_table_rows if not r["is_red"]]

    def render_row_html(r, idx, force_red=False):
        cls = "row-red" if (force_red or r["is_red"]) else ("row-alt" if idx % 2 == 1 else "row-norm")
        obs_html = html.escape(r["obs"]) if r["obs"] else '<span class="obs-empty">—</span>'
        return f"""
            <tr class="{cls}">
              <td class="col-ad">{html.escape(r['aeropuerto'])}</td>
              <td class="col-notam"><a href="#{anchor_id(r['notam'])}" class="tbl-link">{html.escape(r['notam'])}</a></td>
              <td class="col-txt">{html.escape(r['texto'])}</td>
              <td class="col-dt">{html.escape(r['desde'])}</td>
              <td class="col-dt">{html.escape(r['hasta'])}</td>
              <td class="col-obs">{obs_html}</td>
              <td class="col-ini">{html.escape(r['notam_ini'])}</td>
              <td class="col-dias">{r['dias']}</td>
            </tr>
            """

    table_body_html = []
    if novedades_rows:
        table_body_html.append(
            f'<tr class="sec-row sec-row-red"><td colspan="8">NOVEDADES OPERACIONALES · PUBLICADOS, INICIADOS O VENCEN HOY ({len(novedades_rows)})</td></tr>'
        )
        for idx, r in enumerate(novedades_rows):
            table_body_html.append(render_row_html(r, idx, force_red=True))

    for sec in section_order:
        sec_rows = [r for r in vigentes_rows if r["section"] == sec]
        if not sec_rows:
            continue
        table_body_html.append(
            f'<tr class="sec-row"><td colspan="8">{html.escape(sec)} ({len(sec_rows)})</td></tr>'
        )
        for idx, r in enumerate(sec_rows):
            table_body_html.append(render_row_html(r, idx))

    # =========================================================================
    # CONSTRUIR ANEXO DE REFERENCIA OACI (TODOS LOS NOTAMS CITADOS EN PLANOS Y TABLA)
    # =========================================================================
    cited_ids = []
    for nid, *_ in spjc_day_cards_spec + extra_spjc_cards + spjc_night_top_spec + spjc_night_bot_spec + spjc_procs_spec:
        if nid in active and nid not in cited_ids:
            cited_ids.append(nid)
    for spec in spzo_cards_spec + extra_spzo_cards:
        if spec and spec[0] in active and spec[0] not in cited_ids:
            cited_ids.append(spec[0])
    for r in verified_table_rows:
        if r["notam"] not in cited_ids:
            cited_ids.append(r["notam"])

    annex_cards = []
    for nid in cited_ids:
        n_obj = active[nid]
        icao = n_obj.get("facilityDesignator", "")
        raw_msg = html.escape(n_obj.get("icaoMessage", "").strip())
        b_loc = fmt_local(n_obj.get("startDate"))
        c_loc = fmt_local(n_obj.get("endDate"))
        back_href = "#sec-tabla" if any(r["notam"] == nid for r in verified_table_rows) else ("#sec-spzo" if icao == "SPZO" else "#sec-spjc")
        red_tag = ' <span class="badge-new">NOVEDAD</span>' if is_red_notam(n_obj, today_date, cutoff_local_dt) else ""
        annex_cards.append(f"""
        <div class="anx-card" id="{anchor_id(nid)}">
          <div class="anx-head">
            <span><strong>{html.escape(nid)}</strong> · {html.escape(icao)}{red_tag}</span>
            <span class="anx-meta">Local: {b_loc} → {c_loc} · <a href="{back_href}" class="back-link">↑ Volver</a></span>
          </div>
          <pre class="anx-pre">{raw_msg}</pre>
        </div>
        """)

    # =========================================================================
    # HTML COMPLETO DEL PDF UNIFICADO
    # =========================================================================
    html_doc = f"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<title>Reporte Diario de NOTAMs ANS - {fecha_ddmmyy}</title>
<style>
  @page {{
    size: A4 landscape;
    margin: 7mm 8mm 7mm 8mm;
  }}
  * {{ box-sizing: border-box; margin: 0; padding: 0; }}
  body {{
    font-family: 'Segoe UI', Arial, Helvetica, sans-serif;
    color: #1e293b;
    background: #ffffff;
    -webkit-print-color-adjust: exact;
    print-color-adjust: exact;
  }}
  a {{ text-decoration: none; }}
  .page {{
    width: 100%;
    page-break-after: always;
    position: relative;
  }}
  .page:last-child {{ page-break-after: auto; }}

  /* Encabezado institucional */
  .top-hdr {{
    display: flex;
    justify-content: space-between;
    align-items: center;
    border-bottom: 2.4px solid #0f294a;
    padding: 4px 0 8px 0;
    margin-bottom: 8px;
  }}
  .top-hdr-left {{
    display: flex;
    align-items: center;
    gap: 14px;
  }}
  .top-hdr-left img {{
    height: 36px;
    width: auto;
  }}
  .top-title {{
    font-size: 14.5px;
    font-weight: 800;
    color: #0f294a;
    letter-spacing: 0.25px;
  }}
  .top-sub {{
    font-size: 8.8px;
    color: #475569;
    margin-top: 2px;
  }}
  .top-meta {{
    text-align: right;
    font-size: 9px;
    color: #334155;
    line-height: 1.45;
  }}
  .top-meta strong {{ color: #0f294a; }}

  /* Layout Pagina 1: SPJC Dia / Noche */
  .spjc-grid {{
    display: grid;
    grid-template-columns: 1fr 1.12fr;
    gap: 10px;
  }}
  .panel {{
    border: 1px solid #cbd5e1;
    border-radius: 5px;
    padding: 5px 6px;
    background: #f8fafc;
  }}
  .panel-hdr {{
    display: flex;
    justify-content: space-between;
    align-items: baseline;
    margin-bottom: 4px;
    border-bottom: 1px solid #e2e8f0;
    padding-bottom: 2px;
  }}
  .panel-title {{
    font-size: 10px;
    font-weight: 800;
    color: #0f294a;
  }}
  .panel-sub {{
    font-size: 7.5px;
    color: #64748b;
  }}

  /* Tarjetas de mapa */
  .cards-row-2 {{ display: grid; grid-template-columns: 1fr 1fr; gap: 4px; margin-bottom: 4px; }}
  .cards-row-6 {{ display: grid; grid-template-columns: repeat(6, 1fr); gap: 3px; margin-bottom: 3px; }}
  .cards-row-5 {{ display: grid; grid-template-columns: repeat(5, 1fr); gap: 3px; margin-top: 3px; }}
  .cards-row-3 {{ display: grid; grid-template-columns: repeat(3, 1fr); gap: 8px; margin-bottom: 8px; }}
  .procs-strip {{
    margin-top: 6px;
    border-top: 1.2px solid #cbd5e1;
    padding-top: 4px;
  }}
  .procs-title {{
    font-size: 8.3px;
    font-weight: 800;
    color: #1e3a8a;
    margin-bottom: 3px;
  }}
  .procs-grid {{
    display: grid;
    grid-template-columns: repeat(4, 1fr);
    gap: 4px;
  }}

  .mcard {{
    background: #ffffff;
    border: 1px solid #cbd5e1;
    border-radius: 3.5px;
    padding: 3px 4.5px;
    line-height: 1.18;
  }}
  .mcard-red {{ border-color: #c53929; border-left: 3px solid #c53929; }}
  .mcard-amber {{ border-color: #d97706; border-left: 3px solid #d97706; }}
  .mcard-blue {{ border-color: #2563eb; border-left: 3px solid #2563eb; background: #eff6ff; }}
  .mcard-head {{
    font-size: 6.7px;
    font-weight: 800;
    margin-bottom: 1.5px;
    display: flex;
    flex-wrap: wrap;
    align-items: center;
    gap: 2px;
  }}
  .mcard-red .notam-link {{ color: #b91c1c; text-decoration: underline; }}
  .mcard-amber .notam-link {{ color: #b45309; text-decoration: underline; }}
  .mcard-blue .notam-link {{ color: #1d4ed8; text-decoration: underline; }}
  .mcard-title {{ color: #0f172a; }}
  .mcard-body {{ font-size: 6.3px; color: #334155; }}
  .mcard-sched {{ font-size: 6.0px; color: #64748b; margin-top: 1.5px; font-weight: 600; }}
  .badge-new {{
    background: #c53929;
    color: #ffffff;
    font-size: 5.5px;
    font-weight: 800;
    padding: 0.5px 3px;
    border-radius: 2px;
    letter-spacing: 0.3px;
  }}

  /* Mapas SVG */
  .map-wrap {{
    background: #ffffff;
    border: 1px solid #e2e8f0;
    border-radius: 3px;
    overflow: hidden;
  }}
  .map-svg {{
    display: block;
    width: 100%;
    height: auto;
  }}
  .spzo-wrap {{
    background: #ffffff;
    border: 1px solid #cbd5e1;
    border-radius: 4px;
    padding: 18px 8px;
    margin-bottom: 10px;
  }}
  .spzo-svg {{
    display: block;
    width: 100%;
    height: auto;
  }}
  .lbl-rwy-end {{ font-size: 24px; font-weight: 800; fill: #0f172a; }}
  .lbl-rwy-name {{ font-size: 24px; font-weight: 800; fill: #1e293b; }}
  .lbl-twy {{ font-size: 20px; font-weight: 700; fill: #334155; }}
  .lbl-apn {{ font-size: 18px; font-weight: 700; fill: #475569; }}
  .lbl-term {{ font-size: 21px; font-weight: 800; fill: #0f294a; }}
  .pin-lbl {{ font-size: 26px; font-weight: 900; fill: #1d4ed8; stroke: #ffffff; stroke-width: 4px; paint-order: stroke; }}

  .lbl-spzo-rwy-end {{ font-size: 22px; font-weight: 800; fill: #0f172a; }}
  .lbl-spzo-rwy-name {{ font-size: 22px; font-weight: 800; fill: #1e293b; }}
  .lbl-spzo-twy {{ font-size: 22px; font-weight: 800; fill: #0f294a; }}
  .lbl-spzo-term {{ font-size: 22px; font-weight: 800; fill: #0f294a; }}

  .legend-bar {{
    margin-top: 5px;
    font-size: 7px;
    color: #475569;
    display: flex;
    justify-content: space-between;
    border-top: 1px solid #e2e8f0;
    padding-top: 3px;
  }}
  .dot-r {{ color: #c53929; font-weight: 800; }}
  .dot-a {{ color: #d97706; font-weight: 800; }}
  .dot-b {{ color: #2563eb; font-weight: 800; }}

  /* Pagina 2 tarjetas Cusco mas legibles */
  .spzo-page .mcard {{ padding: 6px 8px; }}
  .spzo-page .mcard-head {{ font-size: 9px; margin-bottom: 3px; }}
  .spzo-page .mcard-body {{ font-size: 8px; }}
  .spzo-page .mcard-sched {{ font-size: 7.5px; margin-top: 3px; }}
  .spzo-note {{
    background: #f8fafc;
    border: 1px solid #cbd5e1;
    border-left: 3.5px solid #0f294a;
    border-radius: 4px;
    padding: 7px 10px;
    font-size: 8.2px;
    color: #334155;
    line-height: 1.4;
  }}

  /* Pagina 3+: Tabla de Ocurrencias Operacionales */
  .tbl-note {{
    font-size: 7.7px;
    color: #475569;
    margin-bottom: 7px;
    background: #f1f5f9;
    padding: 5px 8px;
    border-radius: 3px;
    border-left: 3px solid #0f294a;
    line-height: 1.32;
  }}
  table.ops-tbl {{
    width: 100%;
    border-collapse: collapse;
    font-size: 7.5px;
    table-layout: fixed;
  }}
  table.ops-tbl thead {{ display: table-header-group; }}
  table.ops-tbl th {{
    background: #0f294a;
    color: #ffffff;
    font-weight: 700;
    text-align: center;
    padding: 7px 5px;
    border: 1px solid #0f294a;
    font-size: 7.7px;
    letter-spacing: 0.25px;
    line-height: 1.25;
  }}
  .obs-empty {{ color: #94a3b8; display: block; text-align: center; }}
  table.ops-tbl td {{
    padding: 4.5px 5px;
    border: 1px solid #cbd5e1;
    vertical-align: middle;
    line-height: 1.28;
    word-wrap: break-word;
  }}
  tr.sec-row td {{
    background: #dbeafe;
    color: #0f294a;
    font-weight: 800;
    font-size: 7.8px;
    padding: 3.5px 6px;
    letter-spacing: 0.4px;
    border-top: 1.5px solid #0f294a;
  }}
  tr.sec-row-red td {{
    background: #fee2e2;
    color: #991b1b;
    border-top: 1.8px solid #c53929;
  }}
  tr.row-norm td {{ background: #ffffff; color: #0f172a; }}
  tr.row-alt td {{ background: #f8fafc; color: #0f172a; }}
  tr.row-red td {{ background: #fef2f2; color: #c53929; font-weight: 600; }}
  .col-ad {{ width: 10.5%; font-weight: 700; text-align: left; }}
  .col-notam {{ width: 7.5%; text-align: center; font-weight: 800; }}
  .col-txt {{ width: 34.5%; text-align: left; }}
  .col-dt {{ width: 9.2%; text-align: center; font-variant-numeric: tabular-nums; }}
  .col-obs {{ width: 17.6%; text-align: left; }}
  .col-ini {{ width: 7.0%; text-align: center; }}
  .col-dias {{ width: 4.5%; text-align: center; font-weight: 700; }}
  .tbl-link {{
    color: #1d4ed8;
    text-decoration: underline;
  }}
  tr.row-red .tbl-link {{
    color: #b91c1c;
    text-decoration: underline;
  }}
  .sign-footer {{
    margin-top: 8px;
    display: flex;
    justify-content: space-between;
    align-items: flex-end;
    font-size: 7.5px;
    color: #334155;
    border-top: 1px solid #cbd5e1;
    padding-top: 5px;
  }}
  .sign-cols {{ display: flex; gap: 24px; font-weight: 700; color: #0f294a; }}
  .sign-notes {{ text-align: right; font-size: 6.8px; color: #64748b; max-width: 62%; }}

  /* Anexo OACI */
  .anx-grid {{
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 6px;
  }}
  .anx-card {{
    border: 1px solid #cbd5e1;
    border-radius: 4px;
    padding: 4.5px 6px;
    background: #f8fafc;
    break-inside: avoid;
  }}
  .anx-head {{
    display: flex;
    justify-content: space-between;
    align-items: center;
    font-size: 7.4px;
    color: #0f294a;
    border-bottom: 1px solid #e2e8f0;
    padding-bottom: 2px;
    margin-bottom: 3px;
  }}
  .anx-meta {{ font-size: 6.5px; color: #475569; }}
  .back-link {{ color: #1d4ed8; font-weight: 700; text-decoration: underline; }}
  .anx-pre {{
    font-family: 'Consolas', 'Courier New', monospace;
    font-size: 6.3px;
    color: #1e293b;
    white-space: pre-wrap;
    word-break: break-word;
    line-height: 1.25;
  }}
</style>
</head>
<body>

<!-- ==================== PAGINA 1: LIMA SPJC ==================== -->
<div class="page" id="sec-spjc">
  <div class="top-hdr">
    <div class="top-hdr-left">
      <img src="{logo_uri}" alt="DGAC">
      <div>
        <div class="top-title">LIMA · SPJC | Obras, cierres y procedimientos en superficie</div>
        <div class="top-sub">Área de movimiento vigente al {fecha_larga} · Cada afectación aplica únicamente en sus fechas y horarios publicados (no simultáneas).</div>
      </div>
    </div>
    <div class="top-meta">
      <div><strong>Fecha:</strong> {fecha_ddmmyy} · <strong>Corte:</strong> {now_local.strftime('%H:%M')} Hora Local</div>
      <div><strong>Fuente:</strong> CTNAV API (vigencia verificada) · <strong>Base:</strong> AIRAC AIP SUP 34/2025</div>
    </div>
  </div>

  <div class="spjc-grid">
    <!-- BLOQUE DIA -->
    <div class="panel">
      <div class="panel-hdr">
        <span class="panel-title">DÍA | 07:00 – 19:00 HORA LOCAL</span>
        <span class="panel-sub">Cierres/obras diurnas + procedimientos en superficie</span>
      </div>
      <div class="cards-row-2">
        {''.join(spjc_day_cards_html)}
      </div>
      <div class="map-wrap">
        {render_spjc_day_svg(active, spjc_bg_uri)}
      </div>
    </div>

    <!-- BLOQUE NOCHE -->
    <div class="panel">
      <div class="panel-hdr">
        <span class="panel-title">NOCHE | 19:00 – 07:00 HORA LOCAL</span>
        <span class="panel-sub">Obras y cierres nocturnos programados (01:30–04:30 LOCAL)</span>
      </div>
      <div class="cards-row-6">
        {''.join(spjc_night_top_html)}
      </div>
      <div class="map-wrap">
        {render_spjc_night_svg(active, spjc_bg_uri)}
      </div>
      <div class="cards-row-5">
        {''.join(spjc_night_bot_html)}
      </div>
    </div>
  </div>

  <!-- PROCEDIMIENTOS Y LIMITACIONES CONTINUAS SPJC -->
  <div class="procs-strip">
    <div class="procs-title">PROCEDIMIENTOS Y RESTRICCIONES EN SUPERFICIE SPJC (P1: espera TWY L · P2: cruces L5 y D · P3: remolque A-L y F-L7 · operaciones segregadas y rodaje)</div>
    <div class="procs-grid">
      {''.join(spjc_procs_html)}
    </div>
  </div>

  <div class="legend-bar">
    <div><span class="dot-r">■ ROJO:</span> cierre / indisponibilidad &nbsp;|&nbsp; <span class="dot-a">■ ÁMBAR:</span> obras / limitación / precaución &nbsp;|&nbsp; <span class="dot-b">● AZUL:</span> procedimiento &nbsp;|&nbsp; Clic en cada código NOTAM para ver el texto íntegro OACI.</div>
    <div>Expresado en <strong>Hora Local (UTC-5)</strong> · Trazado referencial sobre AIRAC AIP SUP 34/2025</div>
  </div>
</div>

<!-- ==================== PAGINA 2: CUSCO SPZO ==================== -->
<div class="page spzo-page" id="sec-spzo">
  <div class="top-hdr">
    <div class="top-hdr-left">
      <img src="{logo_uri}" alt="DGAC">
      <div>
        <div class="top-title">CUSCO · SPZO | Pista y calles de rodaje (Área de movimiento)</div>
        <div class="top-sub">Inventario vigente al {fecha_larga} · Haga clic en cualquier código NOTAM para consultar el texto oficial OACI en el Anexo.</div>
      </div>
    </div>
    <div class="top-meta">
      <div><strong>Fecha:</strong> {fecha_ddmmyy} · <strong>Corte:</strong> {now_local.strftime('%H:%M')} Hora Local</div>
      <div><strong>Plano base:</strong> AD 2 SPZO 09 (24 MAR 2022)</div>
    </div>
  </div>

  <div class="cards-row-3" style="margin-top: 10px;">
    {''.join(spzo_cards_html)}
  </div>

  <div class="spzo-wrap">
    {render_spzo_svg(active, spzo_bg_uri)}
  </div>

  <div class="spzo-note">
    <strong>Observaciones operacionales SPZO:</strong> El NOTAM <a href="#{anchor_id('A3740/26')}" class="tbl-link"><strong>A3740/26</strong></a> (publicado hoy 20/11 a las 17:31 local) reemplaza al anterior A3712/26, ampliando la prohibición de giro de 180° a toda la pista <strong>RWY 10/28</strong> para evitar deterioro de la capa asfáltica (no declara cierre de pista). En <strong>TWY C</strong>, se representa el sector afectado por cierre parcial por obras (<a href="#{anchor_id('A3707/26')}" class="tbl-link"><strong>A3707/26</strong></a>) y la depresión de pavimento señalizada al inicio de la calle (<a href="#{anchor_id('A3711/26')}" class="tbl-link"><strong>A3711/26</strong></a>); los límites métricos exactos están sujetos a la señalización en campo.
  </div>

  <div class="legend-bar" style="margin-top: 16px;">
    <div><span class="dot-r">■ ROJO:</span> cierre / indisponibilidad &nbsp;|&nbsp; <span class="dot-a">■ ÁMBAR:</span> limitación / precaución &nbsp;|&nbsp; <span class="dot-b">● AZUL:</span> procedimiento / disponibilidad</div>
    <div>Expresado en <strong>Hora Local (UTC-5)</strong> · Geometría manual referencial sobre AD 2 SPZO 09</div>
  </div>
</div>

<!-- ==================== PAGINA 3: TABLA DE OCURRENCIAS OPERACIONALES ==================== -->
<div class="page" id="sec-tabla">
  <div class="top-hdr">
    <div class="top-hdr-left">
      <img src="{logo_uri}" alt="DGAC">
      <div>
        <div class="top-title">REPORTE DIARIO DE LIMITACIONES EN LOS SERVICIOS DE NAVEGACIÓN AÉREA</div>
        <div class="top-sub">Dirección General de Aeronáutica Civil (DGAC) · Centro de Tránsito y Navegación Aérea (CTNAV)</div>
      </div>
    </div>
    <div class="top-meta">
      <div><strong>Fecha:</strong> {fecha_ddmmyy} (expresados en hora local)</div>
      <div><strong>Verificación en vivo:</strong> {now_local.strftime('%d/%m/%y %H:%M')} · <strong>Total tabla:</strong> {len(verified_table_rows)} ocurrencias</div>
    </div>
  </div>

  <div class="tbl-note">
    <strong>Alcance:</strong> Ocurrencias operacionales vigentes agrupadas por Lima, Cusco, Aeropuertos Internacionales y Aeropuertos Nacionales. Los NOTAMs del área de movimiento de Lima (SPJC) y Cusco (SPZO) han sido retirados de esta tabla al encontrarse graficados en las páginas 1 y 2. Haga clic en cualquier código de <strong>NOTAM ACTUAL</strong> para abrir su texto íntegro OACI en el Anexo.
  </div>

  <table class="ops-tbl">
    <thead>
      <tr>
        <th class="col-ad">Aeropuerto</th>
        <th class="col-notam">NOTAM ACTUAL</th>
        <th class="col-txt">Texto del NOTAM</th>
        <th class="col-dt">DESDE<br>(Hora Local)</th>
        <th class="col-dt">HASTA<br>(Hora Local)</th>
        <th class="col-obs">OBSERVACIONES</th>
        <th class="col-ini">NOTAM INICIAL</th>
        <th class="col-dias">DÍAS</th>
      </tr>
    </thead>
    <tbody>
      {''.join(table_body_html)}
    </tbody>
  </table>

  <div class="sign-footer">
    <div class="sign-cols">
      <span>Elaborado por: CTNAV</span>
      <span>Revisado por: DGAC</span>
      <span>Aprobado por: DGAC</span>
    </div>
    <div class="sign-notes">
      *La información en <strong>color rojo</strong> corresponde a NOTAMs publicados o iniciados el día del reporte (o desde las 18:00 del día anterior / fin de semana si es lunes), así como a aquellos que vencen el mismo día.<br>
      *Se excluyen operatividad SEI (salvo afectación directa ANS), puestos de estacionamiento (PEA/STAND), horarios de atención AIP y actividades aerodeportivas Costa Verde.
    </div>
  </div>
</div>

<!-- ==================== PAGINAS FINALES: ANEXO OACI ==================== -->
<div class="page" id="sec-anexo">
  <div class="top-hdr">
    <div class="top-hdr-left">
      <img src="{logo_uri}" alt="DGAC">
      <div>
        <div class="top-title">ANEXO: NOTAMs DE REFERENCIA (TEXTO ÍNTEGRO OACI)</div>
        <div class="top-sub">Referencia oficial extraída en vivo desde <a href="https://ctnav.pe/notam/" style="color:#1d4ed8;text-decoration:underline;">https://ctnav.pe/notam/</a> ({len(cited_ids)} NOTAMs vigentes verificados).</div>
      </div>
    </div>
    <div class="top-meta">
      <div><strong>Fecha:</strong> {fecha_ddmmyy} · <strong>Corte:</strong> {now_local.strftime('%H:%M')} Hora Local</div>
      <div><a href="#sec-spjc" class="back-link">Plano SPJC</a> · <a href="#sec-spzo" class="back-link">Plano SPZO</a> · <a href="#sec-tabla" class="back-link">Tabla ANS</a></div>
    </div>
  </div>

  <div class="anx-grid">
    {''.join(annex_cards)}
  </div>
</div>

</body>
</html>
"""

    html_path = os.path.join(BASE_DIR, f"{yymmdd}_reporte_unificado.html")
    with open(html_path, "w", encoding="utf-8") as f:
        f.write(html_doc)

    # Renderizar a PDF con Chrome/Edge headless
    os.makedirs(LOCAL_OUT_DIR, exist_ok=True)
    pdf_filename = f"{yymmdd} REPORTE NOTAMS ANS.pdf"
    local_pdf = os.path.join(LOCAL_OUT_DIR, pdf_filename)
    drive_pdf = os.path.join(DRIVE_OUT_DIR, pdf_filename)
    browser = find_browser_exe()

    cmd = [
        browser,
        "--headless=new",
        "--disable-gpu",
        "--no-sandbox",
        "--print-to-pdf-no-header",
        f"--print-to-pdf={local_pdf}",
        "file:///" + html_path.replace("\\", "/").replace(" ", "%20"),
    ]
    subprocess.run(cmd, check=True)
    safe_sync(local_pdf, drive_pdf)
    print(f"PDF unificado generado: {local_pdf}")

    limpiar_archivos_antiguos(today_date, dias_retencion=7)

    upload_via_apps_script_webhook([
        (local_pdf, "application/pdf"),
    ])

    return local_pdf


if __name__ == "__main__":
    build_unified_report()
