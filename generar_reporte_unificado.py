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
    secret = os.environ.get("DRIVE_WEBHOOK_SECRET", "").strip() or "DGAC_CTNAV_NOTAM_2026"
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
        res_json = json.loads(res_txt)
        if not res_json.get("ok"):
            raise RuntimeError(f"Error al subir a Google Drive: {res_txt}")
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


def is_excluded_non_ans(e_txt, icao, indicadores):
    """Filtros estrictos de exclusión (Costa Verde, SEI/RFFS, stands/TXL, horarios AIP, RPAS, globos, HLP, etc.)."""
    u = e_txt.upper()
    if any(k in u for k in ("HR SER", "SER HR", "AVBL O/R", "OPR HR", "AD HR", "ATS HR", "TWR HR")):
        return True
    if any(k in u for k in ("COSTA VERDE", "PARAPENTE", "PARATRIKE", "PARAMOTOR", "ALA DELTA", "AERODEPORTIVA")):
        return True
    if any(k in u for k in ("RFFS", "FIRE FIGHTING", "RESCUE AND FIRE", "SEI ")) and not re.search(r"\b(RWY|TWY|ILS|VOR|DME|ATS|TWR)\b", u):
        return True
    if any(k in u for k in ("ACFT STAND", "PARKING STAND", "PEA ", "STANDS ", "TXL ")):
        return True
    if any(k in u for k in ("CHECKLIST", "TRIGGER NOTAM", "AIP AMDT", "REF AIP", "WEF ", "SUP AIP", "HTTPS://")) and "A-CDM" not in u and "ACDM" not in u:
        return True
    if any(k in u for k in ("RPAS", "DRONE", "PROBE BALLOON", "GRASS CUTTING", "VIRAJE LENTO", "REF RD ", "PCN ")):
        return True
    if u.startswith("HLP ") or "HELIPUERTO" in u:
        return True
    if u.startswith("OBST ") and not re.search(r"\b(RWY|GRADIENT)\b", u):
        return True
    if ("AD CLSD" in u or "AD CERRADO" in u) and icao not in INTL_ICAOS and icao not in LIMA_ICAOS and icao not in CUSCO_ICAOS:
        _, distintivo = indicadores.get(icao, ("", ""))
        if not distintivo:
            return True
    if icao == "SPLP" and not re.search(r"\b(AD CLSD|RWY)\b", u):
        return True
    return False


def is_movement_area_spjc_spzo(icao, msg_e):
    """Solo verdadero si refiere estrictamente a pistas (RWY), calles de rodaje (TWY) o transferencias (GATE TRANSFER) en SPJC/SPZO."""
    if icao not in ("SPJC", "SPZO"):
        return False
    u = msg_e.upper()
    if any(k in u for k in (
        "121.5", "LINEA 2", "CLIMB GRADIENT", "GRADIENTE", "SPR91", "AIR DISPLAY",
        "PARACHUTE", "TSA ", "GUN FIRING", "GUN-FIRING", "A-CDM", "ACDM", "GNSS",
        "TXL ", "ACFT STAND", "AERODEPORTIVA", "PARAPENTE", "COSTA VERDE"
    )):
        return False
    return bool(re.search(r"\b(RWY|TWY|GATE\s+TRANSFER|TORA|ASDA|LDA)\b", u))


def is_spjc_night_notam(n_obj):
    """Determina si un NOTAM de SPJC aplica en bloque NOCHE (19:00-07:00 hora local)."""
    d_raw = extract_item_d(n_obj.get("icaoMessage", "")).upper()
    hm = re.search(r"(\d{4})\s*-\s*(\d{4})", d_raw)
    if hm:
        lh = (int(hm.group(1)[:2]) - 5) % 24
        return lh >= 19 or lh < 7
    st = parse_utc(n_obj.get("startDate"))
    en = parse_utc(n_obj.get("endDate"))
    if st and en:
        dur_h = (en - st).total_seconds() / 3600.0
        if dur_h <= 12:
            lh = st.astimezone(LIMA_TZ).hour
            return lh >= 19 or lh < 7
    return False


def is_ans_reportable(n_obj, indicadores):
    """Aplica las 14 reglas para filtrar solo NOTAMs ANS reportables."""
    icao = str(n_obj.get("facilityDesignator", "")).strip().upper()
    msg = str(n_obj.get("icaoMessage", "")).upper()
    e_txt = extract_item_e(msg).upper()

    if is_excluded_non_ans(e_txt, icao, indicadores):
        return False

    # Solo incluir NOTAMs publicados recientemente (>= 2026-11-20 18:00) y con patron operacional ANS clave
    iss_str = str(n_obj.get("issueDate", "")).strip()
    is_recent = iss_str >= "2026-11-20 18:00:00"
    is_core_ans = bool(re.search(
        r"\b(RWY|TWY|ILS|VOR|DME|GNSS|APAPI|PAPI|BIRD HAZARD|VENTEO DE GAS|TSA|PARACHUTE|AIR DISPLAY|GUN FIRING|GUN-FIRING|IAC|SID|STAR|GRADIENT|AD CLSD)\b",
        e_txt,
    ))
    return is_recent and is_core_ans


def extract_local_slot_spjc(n_obj):
    """Extrae la franja horaria diaria en hora local (ej. '01:30 – 04:30 HL') para agrupar NOTAMs con el mismo horario."""
    d_raw = extract_item_d(n_obj.get("icaoMessage", "")).upper()
    hm = re.search(r"(\d{4})\s*-\s*(\d{4})", d_raw)
    if hm:
        return f"{utc_hhmm_to_local_str(hm.group(1))} – {utc_hhmm_to_local_str(hm.group(2))} HL"
    st = parse_utc(n_obj.get("startDate"))
    en = parse_utc(n_obj.get("endDate"))
    if st and en and (en - st).total_seconds() <= 12 * 3600:
        return f"{st.astimezone(LIMA_TZ).strftime('%H:%M')} – {en.astimezone(LIMA_TZ).strftime('%H:%M')} HL"
    return "Horario publicado"


def classify_spjc_day_surface(n_obj):
    """
    Clasifica dinámicamente cualquier NOTAM diurno del área de movimiento de SPJC:
    - Cierres (CLSD): tono 'red', trazo de toda la superficie cerrada.
    - Obras / precaución (WIP / CTN): tono 'orange', puntos naranjas aproximados.
    - Procedimientos: tono 'blue', puntos azules aproximados.
    """
    u = extract_item_e(n_obj.get("icaoMessage", "")).upper()
    st_s = parse_utc(n_obj.get("startDate"))
    en_s = parse_utc(n_obj.get("endDate"))
    if st_s and en_s:
        d1 = st_s.astimezone(LIMA_TZ).strftime("%d/%m")
        d2 = en_s.astimezone(LIMA_TZ).strftime("%d/%m")
        f_rng = d1 if d1 == d2 else f"{d1}–{d2}"
    else:
        f_rng = ""
    slot = extract_local_slot_spjc(n_obj)

    if "CLSD" in u:
        if "16L/34R" in u:
            return ("Cierres", "RWY 16L/34R cerrada", f_rng, slot, "red", "rwy_16l", [])
        if "16R/34L" in u:
            return ("Cierres", "RWY 16R/34L cerrada", f_rng, slot, "red", "rwy_16r", [])
        return ("Cierres", translate_icao_to_spanish(n_obj), f_rng, slot, "red", "custom_clsd", [(900, 500)])

    if "TWY V" in u and "V1" in u and "TWY Q" in u:
        return (
            "Obras en superficie",
            "Cerca de TWY V (entre V1 e int. TWY Q) · zona señalizada",
            f_rng,
            slot,
            "orange",
            "pts",
            [(160, 996), (450, 912), (740, 828)],
        )
    if "STRIP" in u and "16R" in u:
        lado = "lado derecho de RWY 16R" if "RIGHT SIDE" in u else "RWY 16R/34L"
        pts = [(380, 1060), (760, 950), (1140, 839)] if "RIGHT SIDE" in u else [(280, 1036), (760, 896), (1240, 756)]
        return (
            "Obras en superficie",
            f"Franja {lado} · zona señalizada",
            f_rng,
            slot,
            "orange",
            "pts",
            pts,
        )
    if "WIP" in u or "OBRA" in u or "DEPRESSION" in u or "CTN" in u:
        pts = [(620, 422)] if "16L" in u else ([(650, 954)] if "16R" in u else [(950, 580)])
        return ("Obras en superficie", translate_icao_to_spanish(n_obj), f_rng, slot, "orange", "pts", pts)

    return ("Procedimientos", translate_icao_to_spanish(n_obj), f_rng, slot, "blue", "pts", [(900, 520)])


def classify_spjc_night_surface(n_obj):
    """
    Clasifica cualquier NOTAM nocturno de SPJC:
    - Cierres (CLSD / NOT AVBL): tono 'red' y geometría de superficie completa cerrada.
    - Obras (WIP sin cierre): tono 'orange' y puntos naranjas aproximados.
    """
    u = extract_item_e(n_obj.get("icaoMessage", "")).upper()
    st_s = parse_utc(n_obj.get("startDate"))
    en_s = parse_utc(n_obj.get("endDate"))
    if st_s and en_s:
        d1 = st_s.astimezone(LIMA_TZ).strftime("%d/%m")
        d2 = en_s.astimezone(LIMA_TZ).strftime("%d/%m")
        f_rng = d1 if d1 == d2 else f"{d1}–{d2}"
    else:
        f_rng = ""
    slot = extract_local_slot_spjc(n_obj)

    if "GATE TRANSFER U2" in u:
        return ("Puertas U", "Puerta U2", f_rng, slot, "red", "seg", [(1220, 628, 1235, 676)])
    if "GATE TRANSFER U3" in u:
        return ("Puertas U", "Puerta U3", f_rng, slot, "red", "seg", [(1350, 588, 1365, 636)])
    if "GATE TRANSFER U1" in u or ("V5" in u and "U1" in u):
        return ("Puertas U", "U / U1 / V5", f_rng, slot, "red", "seg", [(840, 775, 1058, 728), (1043, 680, 1058, 728)])
    if "V3" in u:
        return ("Puertas U", "TWY U y V3", f_rng, slot, "red", "seg", [(1105, 750, 1122, 804)])
    if "RWY 16R/34L CLSD" in u:
        return ("RWY / Franja 16R", "RWY 16R/34L cerrada", f_rng, slot, "red", "rwy_16r", [(38, 1132, 1510, 704)])
    if "STRIP RWY 16R/34L" in u:
        return ("RWY / Franja 16R", "Franja 16R/34L (obras)", f_rng, slot, "orange", "pts", [(280, 1036), (760, 896), (1240, 756)])
    if "GATE TRANSFER M1" in u:
        segs = [(965, 428, 980, 476)]
        if "BTN TWY R" in u:
            segs.insert(0, (810, 476, 965, 428))
        return ("Puertas M1–M4", "TWY M (R–M1) + M1", f_rng, slot, "red", "seg", segs)
    if "GATE TRANSFER M2" in u:
        return ("Puertas M1–M4", "Puerta M2", f_rng, slot, "red", "seg", [(1142, 372, 1157, 420)])
    if "GATE TRANSFER M3" in u:
        return ("Puertas M1–M4", "Puerta M3", f_rng, slot, "red", "seg", [(1255, 338, 1270, 386)])
    if "GATE TRANSFER M4" in u:
        return ("Puertas M1–M4", "Puerta M4", f_rng, slot, "red", "seg", [(1412, 292, 1427, 340)])
    if "BTN TWY R TO TWY M1" in u:
        return ("Tramos TWY M", "Tramo R–M1", f_rng, slot, "red", "seg", [(810, 476, 965, 428)])
    if "BTN TWY M1 TO TWY M2" in u:
        return ("Tramos TWY M", "Tramo M1–M2", f_rng, slot, "red", "seg", [(965, 428, 1142, 372)])
    if "BTN TWY M2 TO TWY M3" in u:
        return ("Tramos TWY M", "Tramo M2–M3", f_rng, slot, "red", "seg", [(1142, 372, 1255, 338)])
    if "BTN TWY M3 TO TWY M4" in u:
        return ("Tramos TWY M", "Tramo M3–M4", f_rng, slot, "red", "seg", [(1255, 338, 1412, 292)])
    if "WIP" in u and "CLSD" not in u:
        return ("RWY / Franja 16R", translate_icao_to_spanish(n_obj), f_rng, slot, "orange", "pts", [(950, 580)])
    if "TWY M" in u:
        return ("Tramos TWY M", "TWY M", f_rng, slot, "red", "seg", [(810, 476, 1412, 292)])
    return ("Puertas U", "TWY U / V", f_rng, slot, "red", "seg", [(840, 775, 1365, 636)])


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


ACRONYM_TOKENS = [
    "A-CDM", "APAPI", "PAPI", "DME", "VOR", "ILS", "NDB", "GNSS",
    "TWR", "TWY", "RWY", "ACFT", "TSA", "TRA", "PJE", "ATC", "ATS",
    "FPL", "TMA", "CTR", "ATZ", "AWY", "ACC", "AIP", "SUP", "IAC",
    "SID", "STAR", "VFR", "IFR", "AMSL", "AGL", "TORA", "TODA",
    "ASDA", "LDA", "CAT", "GND", "NOTAM", "UTC", "NM", "FT", "MHz",
    "U/S", "AD",
]


def preserve_acronyms(txt):
    """Mantiene en MAYÚSCULAS todos los acrónimos aeronáuticos y canales (ej. CH 90X)."""
    txt = re.sub(r"(\d+(?:\.\d+)?)\s*(nm|ft)\b", lambda m: f"{m.group(1)} {m.group(2).upper()}", txt, flags=re.IGNORECASE)
    for acr in ACRONYM_TOKENS:
        txt = re.sub(rf"(?<![A-Za-z0-9]){re.escape(acr)}(?![A-Za-z0-9])", acr, txt, flags=re.IGNORECASE)
    txt = re.sub(r"\bch\s*([0-9]+[xy])\b", lambda m: f"CH {m.group(1).upper()}", txt, flags=re.IGNORECASE)
    txt = re.sub(r"\bmhz\b", "MHz", txt, flags=re.IGNORECASE)
    txt = re.sub(r"\b(TWY|RWY|CTR|TMA|ATZ|AWY)\s+([a-z0-9/]+)\b", lambda m: f"{m.group(1).upper()} {m.group(2).upper()}", txt, flags=re.IGNORECASE)
    for tok in ("RNAV", "RNP", "AR", "FAF", "IAF", "IF", "MAPT", "JCL", "LPA", "SLS", "IQT", "SPQT", "GRU42"):
        txt = re.sub(rf"(?<![A-Za-z0-9]){tok}(?![A-Za-z0-9])", tok, txt, flags=re.IGNORECASE)
    for proper in ("Bergerie", "Iquitos", "Lima", "Cusco", "Piura", "Pisco", "Chiclayo", "Pucallpa", "Tarapoto"):
        txt = re.sub(rf"\b{proper.lower()}\b", proper, txt, flags=re.IGNORECASE)
    txt = re.sub(r"\b([A-Za-z]{1,2}\d{3})\b", lambda m: m.group(1).upper(), txt)
    return txt


def translate_icao_to_spanish(n_obj):
    """Traduce automáticamente un NOTAM nuevo de OACI (E + D) a español claro con acrónimos en mayúsculas."""
    e_raw = extract_item_e(n_obj.get("icaoMessage", ""))
    d_raw = extract_item_d(n_obj.get("icaoMessage", ""))
    sched = format_d_schedule_local(d_raw, n_obj.get("startDate", ""))
    u = e_raw.upper()

    m_rwy = re.match(r"RWY\s+([0-9/LRC]+)\s+CLSD(?:\s+DUE\s+TO\s+WIP)?", u)
    if m_rwy:
        base = f"RWY {m_rwy.group(1)} cerrada por trabajos de mantenimiento" if "WIP" in u else f"RWY {m_rwy.group(1)} cerrada"
        return preserve_acronyms(f"{base} {sched}.".strip() if sched else f"{base}.")

    m_twy = re.match(r"(?:TWY|GATE\s+TRANSFER)\s+([A-Z0-9,\s/\-]+?)\s+CLSD(?:\s+DUE\s+TO\s+WIP)?", u)
    if m_twy:
        seg = m_twy.group(1).strip().replace(" BTN ", " entre ").replace(" TO ", " y ")
        base = f"TWY {seg} cerrada por trabajos de mantenimiento" if "WIP" in u else f"TWY {seg} cerrada"
        return preserve_acronyms(f"{base} {sched}.".strip() if sched else f"{base}.")

    m_navaid = re.match(r"(DME|VOR|ILS|ILS/DME|VOR/DME)\s*(?:AND\s+)?(CH\s*\d+[XY])?\s*U/S", u)
    if m_navaid:
        ch = f" ({m_navaid.group(2).strip()})" if m_navaid.group(2) else ""
        return preserve_acronyms(f"{m_navaid.group(1)}{ch} fuera de servicio.")

    m_apapi = re.match(r"(PAPI|APAPI)\s+RWY\s+([0-9/LRC]+)\s+INOP", u)
    if m_apapi:
        return preserve_acronyms(f"{m_apapi.group(1)} de RWY {m_apapi.group(2)} inoperativo.")

    m_bird = re.match(r"BIRD\s+HAZARD\s+RWY\s+([0-9/LRC]+)", u)
    if m_bird:
        return preserve_acronyms(f"Peligro de aves en RWY {m_bird.group(1)}.")

    if "GNSS SIGNAL DEGRADATION" in u:
        m_dist = re.search(r"(\d+\s*NM)\s+AND\s+BELOW\s+(\d+\s*FT)\s+FROM\s+([A-Z0-9]+)", u)
        zona = f" dentro de {m_dist.group(1)} y por debajo de {m_dist.group(2)} respecto a {m_dist.group(3)}" if m_dist else ""
        return preserve_acronyms(f"Degradación de señal GNSS reportada{zona}; procedimientos RNAV/RNP/AR afectados. Verificar posición con medios alternos y prever demoras.")

    if "TWR WO VIS" in u:
        m_vis = re.match(r"TWR\s+WO\s+VIS\s+TO\s+(.*?)(?:\s+DUE\s+TO\s+(.*))?$", u)
        if m_vis:
            zonas = re.sub(r",\s*", ", ", m_vis.group(1)).replace(" INT ", ", intersección ").replace(" AND ", " y ")
            causa = m_vis.group(2) or ""
            causa = re.sub(r"ACFT\s+PARKED\s+ON\s+STAND\s+NR\s*(\d+)", r"ACFT estacionada en puesto \1", causa, flags=re.IGNORECASE)
            txt_vis = f"TWR sin visibilidad hacia {zonas}" + (f" debido a {causa}." if causa else ".")
            return preserve_acronyms(txt_vis)

    if "AD CLSD" in u:
        return "AD cerrado de forma permanente." if str(n_obj.get("endDate", "")).upper() == "PERM" else "AD cerrado."

    txt = e_raw
    txt = re.sub(r"\s*COORD:.*$", "", txt, flags=re.IGNORECASE).strip()
    txt = re.sub(r"\s*RDO:.*$", "", txt, flags=re.IGNORECASE).strip()
    txt = re.sub(r"\s*F\)[^\n]*$", "", txt, flags=re.IGNORECASE).strip()
    txt = re.sub(r"\.([A-Za-z])", r". \1", txt)
    abbrev_es = [
        (r"^TRA\s+DEBIDO\s+A\b", "Área temporalmente reservada (TRA) debido a"),
        (r"^TSA\s+DEBIDO\s+A\b", "Área temporalmente segregada (TSA) debido a"),
        (r"\bPOSS\b", "Posible"),
        (r"\bPERDIDA\b", "pérdida"),
        (r"\bSGL\b", "de señal"),
        (r"\bAPCH\b", "aproximación"),
        (r"\bBTN\b", "entre"),
        (r"\bRDL\b", "radiales"),
        (r"\bFM\s+([A-Za-z]{3})\s+VOR\b", r"del VOR \1"),
        (r"\bPREVIA\s+COOR\b", "coordinación previa"),
        (r"\bCOOR\b", "coordinación"),
        (r"\bAEREOS\b", "aéreos"),
        (r"\bREQ\b", "solicitar"),
        (r"\bTORRE\b", "TWR"),
        (r"\bAERODROMO\b|\bAERÓDROMO\b", "AD"),
        (r"\bAERONAVES\b", "ACFT"),
    ]
    if txt.isupper():
        txt = txt.capitalize()
    for pat, rep in abbrev_es:
        txt = re.sub(pat, rep, txt, flags=re.IGNORECASE)
    txt = re.sub(r"^posible\b", "Posible", txt, flags=re.IGNORECASE)
    txt = re.sub(r"^área\b", "Área", txt, flags=re.IGNORECASE)
    txt = re.sub(r"\.\s+([a-záéíóúñ])", lambda m: ". " + m.group(1).upper(), txt)
    txt = preserve_acronyms(txt)
    if sched and sched not in txt:
        txt = f"{txt.rstrip('.')} ({sched})."
    elif not txt.endswith("."):
        txt += "."
    return txt


def build_spjc_svg_base_labels():
    """Etiquetas limpias sobre el plano base oficial SPJC (viewBox 0 0 1800 1200)."""
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


def build_spjc_proc_pins(active):
    """Puntos azules aproximados de todos los procedimientos activos (P1 a P7) sobre el plano oficial SPJC."""
    out = []
    if "A3523/26" in active:
        x, y = 136 * 1.8, 514 * 1.2
        out.append(
            f'<circle cx="{x:.1f}" cy="{y:.1f}" r="11" fill="#1d4ed8" stroke="#ffffff" stroke-width="3"/>'
            f'<text x="{x+20:.1f}" y="{y+6:.1f}" class="pin-lbl">P1</text>'
        )
    if "A2828/26" in active or "A2830/26" in active:
        for nx, ny in [(615, 228), (603, 191)]:
            x, y = nx * 1.8, ny * 1.2
            out.append(
                f'<circle cx="{x:.1f}" cy="{y:.1f}" r="11" fill="#1d4ed8" stroke="#ffffff" stroke-width="3"/>'
                f'<text x="{x+18:.1f}" y="{y-10:.1f}" class="pin-lbl">P2</text>'
            )
    if "A2827/26" in active:
        for nx, ny, dx, dy in [(107, 434, -32, -8), (928, 108, 16, -6)]:
            x, y = nx * 1.8, ny * 1.2
            out.append(
                f'<circle cx="{x:.1f}" cy="{y:.1f}" r="11" fill="#1d4ed8" stroke="#ffffff" stroke-width="3"/>'
                f'<text x="{x+dx:.1f}" y="{y+dy:.1f}" class="pin-lbl">P3</text>'
            )
    if "A3424/26" in active:
        for x, y, dx, dy in [(620, 422, 16, -12), (650, 954, 16, 26)]:
            out.append(
                f'<circle cx="{x}" cy="{y}" r="11" fill="#1d4ed8" stroke="#ffffff" stroke-width="3"/>'
                f'<text x="{x+dx}" y="{y+dy}" class="pin-lbl">P4</text>'
            )
    if "A3341/26" in active:
        x, y = 335, 1046
        out.append(
            f'<circle cx="{x}" cy="{y}" r="11" fill="#1d4ed8" stroke="#ffffff" stroke-width="3"/>'
            f'<text x="{x+16}" y="{y+26}" class="pin-lbl">P5</text>'
        )
    if "A3527/26" in active:
        for bx, by in [(705, 342), (975, 272), (1055, 248), (1185, 212), (1285, 182)]:
            out.append(f'<circle cx="{bx}" cy="{by}" r="8" fill="#1d4ed8" stroke="#ffffff" stroke-width="2.5"/>')
        out.append('<text x="718" y="366" class="pin-lbl">P6</text>')
    if "A2863/26" in active:
        x, y = 1395, 259
        out.append(
            f'<circle cx="{x}" cy="{y}" r="11" fill="#1d4ed8" stroke="#ffffff" stroke-width="3"/>'
            f'<text x="{x+16}" y="{y+8}" class="pin-lbl">P7</text>'
        )
    return "\n".join(out)


def render_spjc_day_canvas(active, bg_uri, day_items, today_date, cutoff_local_dt):
    """
    Plano oficial SPJC Día siguiendo convención estricta:
    - Cierres: toda la superficie cerrada en ROJO sólido.
    - Obras: puntos aproximados en NARANJA.
    - Procedimientos: puntos aproximados en AZUL (P1–P7).
    """
    overlays = []
    for nid, cat, zone_lbl, f_rng, slot, tone, geom_type, coords in day_items:
        if geom_type == "rwy_16l":
            overlays.append(
                '<line x1="186" y1="548" x2="1658" y2="124" stroke="#c53929" stroke-width="28" stroke-opacity="0.65" stroke-linecap="butt"/>'
                '<line x1="186" y1="548" x2="1658" y2="124" stroke="#991b1b" stroke-width="5"/>'
            )
        elif geom_type == "rwy_16r":
            overlays.append(
                '<line x1="38" y1="1132" x2="1510" y2="704" stroke="#c53929" stroke-width="28" stroke-opacity="0.65" stroke-linecap="butt"/>'
                '<line x1="38" y1="1132" x2="1510" y2="704" stroke="#991b1b" stroke-width="5"/>'
            )
        elif geom_type == "pts":
            col = "#ea580c" if tone == "orange" else ("#1d4ed8" if tone == "blue" else "#c53929")
            for px, py in coords:
                overlays.append(f'<circle cx="{px}" cy="{py}" r="11" fill="{col}" stroke="#ffffff" stroke-width="3"/>')

    if day_items:
        by_slot = {}
        for it in day_items:
            by_slot.setdefault((it[0] if False else it[4], it[1], it[5]), []).append(it)

        # Agrupar por (slot, tone) igual que en el bloque nocturno
        slot_groups = {}
        for it in day_items:
            slot_groups.setdefault(it[4], []).append(it)

        blocks = []
        for slot_lbl, s_items in slot_groups.items():
            all_orange = all(it[5] == "orange" for it in s_items)
            all_blue = all(it[5] == "blue" for it in s_items)
            box_cls = "uleg-box uleg-box-orange" if all_orange else ("uleg-box uleg-box-blue" if all_blue else "uleg-box")
            hdr_cls = "uleg-hdr uleg-hdr-orange" if all_orange else ("uleg-hdr uleg-hdr-blue" if all_blue else "uleg-hdr")
            hdr_prefix = "OBRAS DIURNAS UNIFICADAS" if all_orange else "AFECTACIONES DIURNAS UNIFICADAS"

            rows_html = []
            for nid, cat, zone_lbl, f_rng, _, tone, _, _ in s_items:
                is_new = is_red_notam(active[nid], today_date, cutoff_local_dt)
                b_new = ' <span class="badge-new">N</span>' if is_new else ""
                chip_cls = "uchip-orange" if tone == "orange" else ("uchip-blue" if tone == "blue" else "uchip-red")
                dt_part = f' <span class="uchip-dt">({html.escape(f_rng)})</span>' if f_rng else ""
                rows_html.append(
                    f'<span class="uchip {chip_cls}"><a href="#{anchor_id(nid)}" class="notam-link">{html.escape(nid)}</a>{b_new} '
                    f'{html.escape(zone_lbl)}{dt_part}</span>'
                )
            blocks.append(f"""
            <div class="{box_cls}">
              <div class="{hdr_cls}">
                <span>{hdr_prefix} · HORARIO COMÚN: <strong>{html.escape(slot_lbl)}</strong></span>
                <span class="uleg-subnote">Puntos naranjas aproximados en el plano</span>
              </div>
              <table class="uleg-tbl"><tbody>
                <tr><td class="uleg-cat">Obras / Superficie</td><td class="uleg-chips">{" ".join(rows_html)}</td></tr>
              </tbody></table>
            </div>
            """)
        legend_html = "".join(blocks)
    else:
        legend_html = """
        <div class="uleg-box uleg-box-blue">
          <div class="uleg-hdr uleg-hdr-blue">ESTADO DIURNO (07:00 – 19:00 HL)</div>
          <div class="uleg-empty">Sin cierres ni obras temporales diurnas activas hoy. Aplican en el plano oficial los puntos azules <strong>P1–P7</strong> detallados en la banda inferior de procedimientos.</div>
        </div>
        """

    return f"""
    <div class="canvas-wrap">
      <svg viewBox="0 15 1800 1165" class="canvas-svg" xmlns="http://www.w3.org/2000/svg">
        <image href="{bg_uri}" x="0" y="0" width="1800" height="1200"/>
        {''.join(overlays)}
        {build_spjc_proc_pins(active)}
        {build_spjc_svg_base_labels()}
      </svg>
    </div>
    {legend_html}
    """


def render_spjc_night_canvas(active, bg_uri, night_items, today_date, cutoff_local_dt):
    """
    Plano oficial SPJC Noche siguiendo convención estricta:
    - Cierres: abarcan toda la superficie cerrada en ROJO sólido (RWY, tramos TWY M y puertas de transferencia).
    - Obras: puntos aproximados en NARANJA (ej. franja RWY 16R/34L).
    - Procedimientos: puntos aproximados en AZUL (P1–P7).
    """
    overlays = []
    drawn_rwy_16r = False

    # Primero dibujar superficies cerradas en rojo sólido, luego puntos naranjas de obras encima
    for nid, cat, zone_lbl, f_rng, slot, tone, geom_type, coords in night_items:
        if geom_type == "rwy_16r" and not drawn_rwy_16r:
            drawn_rwy_16r = True
            overlays.append(
                '<line x1="38" y1="1132" x2="1510" y2="704" stroke="#c53929" stroke-width="28" stroke-opacity="0.65" stroke-linecap="butt"/>'
                '<line x1="38" y1="1132" x2="1510" y2="704" stroke="#991b1b" stroke-width="5"/>'
            )
        elif geom_type == "seg":
            for x1, y1, x2, y2 in coords:
                overlays.append(
                    f'<line x1="{x1}" y1="{y1}" x2="{x2}" y2="{y2}" stroke="#c53929" stroke-width="18" stroke-opacity="0.78" stroke-linecap="butt"/>'
                )

    for nid, cat, zone_lbl, f_rng, slot, tone, geom_type, coords in night_items:
        if geom_type == "pts":
            col = "#ea580c" if tone == "orange" else "#1d4ed8"
            for px, py in coords:
                overlays.append(f'<circle cx="{px}" cy="{py}" r="11" fill="{col}" stroke="#ffffff" stroke-width="3"/>')

    # Agrupar NOTAMs nocturnos por franja horaria común (ej. '01:30 – 04:30 HL')
    by_slot = {}
    for item in night_items:
        by_slot.setdefault(item[4], []).append(item)

    slot_blocks_html = []
    cat_order = ["RWY / Franja 16R", "Tramos TWY M", "Puertas M1–M4", "Puertas U"]
    for slot_lbl, s_items in by_slot.items():
        rows_html = []
        for cat in cat_order:
            c_items = sorted([it for it in s_items if it[1] == cat], key=lambda x: x[7][0][0] if x[7] else 0)
            if not c_items:
                continue
            chips = []
            for nid, _, zone_lbl, f_rng, _, tone, _, _ in c_items:
                is_new = is_red_notam(active[nid], today_date, cutoff_local_dt)
                b_new = '<span class="badge-new">N</span>' if is_new else ""
                cls_chip = "uchip-orange" if tone in ("amber", "orange") else ("uchip-blue" if tone == "blue" else "uchip-red")
                dt_part = f' <span class="uchip-dt">({html.escape(f_rng)})</span>' if f_rng else ""
                chips.append(
                    f'<span class="uchip {cls_chip}"><a href="#{anchor_id(nid)}" class="notam-link">{html.escape(nid)}</a>{b_new} '
                    f'{html.escape(zone_lbl)}{dt_part}</span>'
                )
            rows_html.append(
                f'<tr><td class="uleg-cat">{html.escape(cat)}</td><td class="uleg-chips">{" ".join(chips)}</td></tr>'
            )
        slot_blocks_html.append(f"""
        <div class="uleg-box">
          <div class="uleg-hdr">
            <span>CIERRES Y OBRAS UNIFICADOS · HORARIO COMÚN: <strong>{html.escape(slot_lbl)}</strong></span>
            <span class="uleg-subnote">Alternas: TWY L disp. (sector M) · TWY V disp. (sector U) · ACFT prever demoras</span>
          </div>
          <table class="uleg-tbl"><tbody>{''.join(rows_html)}</tbody></table>
        </div>
        """)

    return f"""
    <div class="canvas-wrap">
      <svg viewBox="0 15 1800 1165" class="canvas-svg" xmlns="http://www.w3.org/2000/svg">
        <image href="{bg_uri}" x="0" y="0" width="1800" height="1200"/>
        {''.join(overlays)}
        {build_spjc_proc_pins(active)}
        {build_spjc_svg_base_labels()}
      </svg>
    </div>
    {''.join(slot_blocks_html)}
    """


def render_spzo_canvas(active, bg_uri, spzo_specs, today_date, cutoff_local_dt):
    """
    Plano oficial AD 2 SPZO 09 de CUSCO · SPZO:
    - Cierres (ej. A3707/26): toda la superficie cerrada en ROJO sólido.
    - Obras / precaución pavimento (ej. A3711/26): punto aproximado en NARANJA.
    - Procedimientos / restricciones operativas (ej. A3740/26 giro 180°): puntos aproximados en AZUL.
    """
    overlays = []
    if "A3707/26" in active:
        overlays.append(
            '<line x1="455" y1="152" x2="735" y2="152" stroke="#c53929" stroke-width="20" stroke-opacity="0.72" stroke-linecap="butt"/>'
        )
    if "A3740/26" in active:
        for bx, by in [(180, 70), (970, 72), (1760, 74)]:
            overlays.append(f'<circle cx="{bx}" cy="{by}" r="11" fill="#1d4ed8" stroke="#ffffff" stroke-width="2.8"/>')
    if "A3711/26" in active:
        overlays.append(
            '<circle cx="527" cy="148" r="11" fill="#ea580c" stroke="#ffffff" stroke-width="2.8"/>'
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

    by_sched = {}
    for spec in spzo_specs:
        if spec and spec[0] in active:
            nid, title, body, sched, tone = spec
            end_loc = fmt_local_short(active[nid].get("endDate"))
            grp_key = f"Aplicación continua (24 h) · Vigencia común hasta {end_loc} HL"
            by_sched.setdefault(grp_key, []).append(spec)

    leg_blocks = []
    for grp_key, items in by_sched.items():
        rows = []
        for nid, title, body, sched, tone in items:
            b_new = ' <span class="badge-new">NUEVO</span>' if is_red_notam(active[nid], today_date, cutoff_local_dt) else ""
            desde_loc = fmt_local_short(active[nid].get("startDate"))
            link_col = "#b91c1c" if tone == "red" else ("#1d4ed8" if tone == "blue" else "#c2410c")
            rows.append(
                f'<tr><td class="uleg-cat" style="width:19%;"><a href="#{anchor_id(nid)}" style="color:{link_col}; text-decoration:underline; font-size:8.5px; font-weight:800;">{html.escape(nid)}</a>{b_new} · <strong>{html.escape(title)}</strong></td>'
                f'<td class="uleg-chips" style="font-size:8.2px; padding:4.5px 8px;">{html.escape(body)} <span class="uchip-dt" style="font-size:7.6px;">(Desde {desde_loc} HL)</span></td></tr>'
            )
        leg_blocks.append(f"""
        <div class="uleg-box uleg-box-blue" style="margin-bottom:10px;">
          <div class="uleg-hdr uleg-hdr-blue" style="font-size:8.5px; padding:5px 8px;">
            <span>LEYENDA UNIFICADA SPZO · <strong>{html.escape(grp_key)}</strong></span>
          </div>
          <table class="uleg-tbl"><tbody>{''.join(rows)}</tbody></table>
        </div>
        """)

    return f"""
    <div class="spzo-map-wrap">
      <svg viewBox="0 0 2000 290" class="spzo-svg" xmlns="http://www.w3.org/2000/svg">
        <image href="{bg_uri}" x="0" y="0" width="2000" height="290"/>
        {''.join(overlays)}
        {lbl_svg}
      </svg>
    </div>
    {''.join(leg_blocks)}
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
            "Entrada de ACFT monomotor restringida dentro del TMA Lima en radio de 30 NM respecto al VOR JCL de 09:30 a 11:30, 16:00 a 18:00 y 20:00 a 00:00, excepto para VFR dentro de zonas militares SPR68 y SPR68A o a 2000 FT o inferior fuera de trayectorias y rutas instrumentales.",
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
            "Precaución por obras de Línea 2 del Metro de Lima con grúas y perforadoras de 28 m de altura a 600 m del extremo de RWY 16R; ACFT en salida deberán mantener gradiente de ascenso mínimo de 4.7%.",
            "",
            "A2179/26",
            date(2026, 9, 11),
        ),
        (
            "LIMA",
            "LIMA AEROPUERTO",
            "A3731/26",
            "Festival aerodeportivo en zona restringida SPR91 de 10:00 a 18:00 desde superficie hasta 2000 FT AMSL (operaciones de ACFT, helicópteros y RPAS). Contactar a Las Palmas TWR para información de tránsito.",
            "Relacionado con C3916/26 (SPLP), A3730/26 (cierre RWY 16L/34R en plano) y A3705/26 (TSA Las Palmas).",
            "A3703/26",
            date(2026, 11, 17),
        ),
        (
            "LIMA",
            "LIMA AEROPUERTO",
            "A3705/26",
            "Área temporalmente segregada (TSA) por lanzamiento de paracaidistas sobre AD Las Palmas de 12:30 a 14:30 y de 16:00 a 17:30 desde superficie hasta 8000 FT AMSL. Requiere coordinación previa con Lima ACC.",
            "Relacionado con C3889/26 (SPLP), A3730/26 (cierre RWY 16L/34R en plano) y A3731/26 (festival SPR91).",
            "A3705/26",
            date(2026, 11, 17),
        ),
        (
            "LIMA",
            "LIMA AEROPUERTO",
            "A3732/26",
            "Área temporalmente segregada (TSA) por ejercicios de tiro en zona Salinas Huacho (entre 3.1 NM y 6.6 NM, radiales 297 a 335 del VOR SLS, desde superficie hasta 6000 FT AGL). Requiere coordinación previa con Lima ACC.",
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
            "Asociado a A3748/26 (cierre simultáneo de TWY).",
            "A3745/26",
            date(2026, 11, 21),
        ),
        (
            "AEROPUERTOS INTERNACIONALES",
            "CHICLAYO",
            "A3748/26",
            "TWY A, B, B1, C, D y D1 cerradas por trabajos de mantenimiento de 23:00 a 04:00 en las mismas noches programadas del cierre de RWY.",
            "Asociado a A3745/26 (cierre simultáneo de RWY 01/19).",
            "A3748/26",
            date(2026, 11, 21),
        ),
        (
            "AEROPUERTOS INTERNACIONALES",
            "PUCALLPA",
            "A3663/26",
            "RWY 02/20 cerrada por trabajos de mantenimiento en horario diario de 23:30 a 05:00.",
            "Asociado a A3665/26 (cierre simultáneo de TWY A y B).",
            "A2840/26",
            date(2026, 9, 1),
        ),
        (
            "AEROPUERTOS INTERNACIONALES",
            "PUCALLPA",
            "A3665/26",
            "TWY A y B cerradas por trabajos de mantenimiento en horario diario de 23:30 a 05:00.",
            "Asociado a A3663/26 (cierre simultáneo de RWY 02/20).",
            "A2839/26",
            date(2026, 9, 1),
        ),
        (
            "AEROPUERTOS INTERNACIONALES",
            "PUERTO MALDONADO",
            "C3069/26",
            "Procedimientos de llegada instrumental IAC VOR/DME - ILS/DME RWY 01 (descensos Alfa, Bravo, Charlie y Delta) suspendidos de manera permanente por razones operacionales.",
            "",
            "C0978/25",
            date(2025, 3, 12),
        ),
        # 4. AEROPUERTOS NACIONALES
        (
            "AEROPUERTOS NACIONALES",
            "TARAPOTO",
            "C3698/26",
            "TWR sin visibilidad en los primeros 700 m del lado izquierdo de la RWY 17 debido a maleza y árboles crecidos.",
            "",
            "C3251/26",
            date(2026, 10, 1),
        ),
        (
            "AEROPUERTOS NACIONALES",
            "ANDOAS",
            "C3814/25",
            "AD cerrado de forma permanente.",
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
    pre_mapped_procs_and_spzo = {
        "A3730/26", "A3523/26", "A2828/26", "A2830/26", "A2827/26", "A3424/26",
        "A3341/26", "A3527/26", "A2863/26", "A3740/26", "A3707/26", "A3711/26",
    }
    spjc_day_items = []
    if "A3730/26" in active:
        cat, zlbl, frng, slt, tn, gtype, crds = classify_spjc_day_surface(active["A3730/26"])
        spjc_day_items.append(("A3730/26", cat, zlbl, frng, slt, tn, gtype, crds))

    extra_spjc_day = []
    spjc_night_items = []
    extra_spzo_cards = []
    for nid, n_obj in sorted(active.items(), key=lambda x: (str(x[1].get("facilityDesignator", "")), x[0])):
        if nid in known_ids or nid in pre_mapped_procs_and_spzo:
            continue
        icao = str(n_obj.get("facilityDesignator", "")).strip().upper()
        msg_e = extract_item_e(n_obj.get("icaoMessage", ""))
        # Primero aplicar SIEMPRE las reglas de exclusión (Costa Verde, stands/TXL, SEI, RPAS, etc.)
        if is_excluded_non_ans(msg_e, icao, indicadores):
            continue
        if icao == "SPJC" and is_movement_area_spjc_spzo(icao, msg_e) and is_spjc_night_notam(n_obj):
            cat, zone_lbl, f_rng, slot, tone, geom_type, coords = classify_spjc_night_surface(n_obj)
            spjc_night_items.append((nid, cat, zone_lbl, f_rng, slot, tone, geom_type, coords))
            continue
        if not is_ans_reportable(n_obj, indicadores):
            continue
        if is_movement_area_spjc_spzo(icao, msg_e):
            if icao == "SPJC":
                cat, zone_lbl, f_rng, slot, tone, geom_type, coords = classify_spjc_day_surface(n_obj)
                extra_spjc_day.append((nid, cat, zone_lbl, f_rng, slot, tone, geom_type, coords))
            else:
                sched = f"Vig: {fmt_local_short(n_obj.get('startDate'))} - {fmt_local_short(n_obj.get('endDate'))} LOCAL"
                u_spzo = msg_e.upper()
                tn_spzo = "red" if "CLSD" in u_spzo else ("orange" if ("WIP" in u_spzo or "DEPRESSION" in u_spzo or "CTN" in u_spzo) else "blue")
                extra_spzo_cards.append((nid, "RWY / TWY", translate_icao_to_spanish(n_obj), sched, tn_spzo))
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
        ("A3730/26", "RWY 16L/34R", "No disponible para aterrizajes ni despegues. RWY 16R/34L en uso. Salidas por 16R con ascenso en rumbo de RWY y viraje a la derecha directo a JC921.", "Vig: 19/11 18:20 - 22/11 17:30 · Cierre: 12:30-14:30 y 16:00-17:30 LOCAL", "red"),
    ]

    # Procedimientos en superficie SPJC (strip inferior unificado, todos en AZUL con puntos P1-P7 en el mapa)
    spjc_procs_spec = [
        ("A3523/26", "P1 · TWY L", "Espera 16L en TWY L (L1 no es punto de espera). No cruzar sin autorización ATC.", "Todo el día · Hasta 31/01/27", "blue"),
        ("A2828/26", "P2 · L5 y D (Bidirec.)", "RWY 16L/34R en uso bidireccional: esperar antes de CAT I y reportar RWY libre.", "06:00-18:00 · Hasta 26/11", "blue"),
        ("A2830/26", "P2 · L5 y D (Unidirec.)", "Uso unidireccional oeste a este: esperar antes de CAT I y cruzar antes de reportar libre.", "18:01-05:59 · Hasta 26/11", "blue"),
        ("A2827/26", "P3 · Remolque A-L / F-L7", "Suspende SUP 12/24. Cruce remolque vía A-L y F-L7 esperando antes de CAT I.", "Todo el día · Hasta 26/11", "blue"),
        ("A3424/26", "P4 · Segregadas 16R / 16L", "RWY 16R solo aterrizajes y RWY 16L solo despegues (exc. emerg./medevac/HEAD).", "06:00-23:59 · Hasta 24/01/27", "blue"),
        ("A3341/26", "P5 · RWY 16R desde V2", "TORA 2773 m desde V2. Salidas de Plat. 1 que no puedan despegar en V2 notificar GND.", "Todo el día · Hasta 20/01/27", "blue"),
        ("A3527/26", "P6 · Marcas 16L (B,C,D,E,G)", "Señalización puntos de espera 16L no disp. Reportar libre al ingresar total a A o F.", "Todo el día · Hasta 02/02/27", "blue"),
        ("A2863/26", "P7 · TWY L (L5-L7) CAT D/E", "No disponible para rodaje CAT D y E. RWY 34R disponible vía TWY F.", "Todo el día · Hasta 30/11", "blue"),
    ]
    spjc_procs_html = []
    for nid, title, body, sched, tone in spjc_procs_spec:
        if nid in active:
            spjc_procs_html.append(card_link(nid, title, body, sched, tone))

    # Tarjetas CUSCO SPZO (Pagina 2: Azul=procedimiento, Rojo=cierre superficie, Naranja=obras/precaución)
    spzo_cards_spec = [
        ("A3740/26", "RWY 10/28 (Procedimiento)", "Prohibido el giro de 180° en RWY 10/28 para evitar deterioro de la capa asfáltica (reemplaza a A3712/26; no declara cierre de RWY).", f"Vig: {fmt_local_short(active['A3740/26']['startDate'])} - {fmt_local_short(active['A3740/26']['endDate'])} LOCAL · Aplicación continua", "blue") if "A3740/26" in active else None,
        ("A3707/26", "Parte de TWY C (Cierre)", "Cierre por obras hasta antes de TWY D (superficie cerrada en rojo).", f"Vig: {fmt_local_short(active['A3707/26']['startDate'])} - {fmt_local_short(active['A3707/26']['endDate'])} LOCAL · Aplicación continua", "red") if "A3707/26" in active else None,
        ("A3711/26", "Inicio de TWY C (Obras / Precaución)", "Precaución por depresión de pavimento al inicio de TWY C (zona señalizada).", f"Vig: {fmt_local_short(active['A3711/26']['startDate'])} - {fmt_local_short(active['A3711/26']['endDate'])} LOCAL · Aplicación continua", "orange") if "A3711/26" in active else None,
    ]
    spzo_obs_html = (
        "Representación gráfica sobre el plano oficial AD 2 SPZO 09 · Las afectaciones vigentes de igual horario/régimen se agrupan en la leyenda unificada superior sin duplicar información."
        if any(s and s[0] in active for s in spzo_cards_spec + extra_spzo_cards)
        else "Sin afectaciones activas en el área de movimiento SPZO."
    )
    spzo_cards_html = []
    for spec in spzo_cards_spec + extra_spzo_cards:
        if spec:
            nid, title, body, sched, tone = spec
            badge = "NUEVO HOY" if is_red_notam(active[nid], today_date, cutoff_local_dt) else None
            spzo_cards_html.append(card_link(nid, title, body, sched, tone, badge))

    # =========================================================================
    # UNIFICAR FILAS DEL MISMO AERODROMO QUE COMPARTEN EL MISMO HORARIO Y VIGENCIA
    # =========================================================================
    def is_spanish_notam_row(row):
        if str(row["notam"]).upper().startswith("C"):
            return True
        e_up = extract_item_e(row["raw"].get("icaoMessage", "")).upper()
        return bool(re.search(r"\b(DEBIDO|EJERCICIOS|PARACAIDISMO|ESPACIOS|REQUIERE|ZONA\s+[A-Z0-9\-]+\s+BTN|FUERA\s+DE\s+SERVICIO|PERDIDA)\b", e_up))

    def unify_same_schedule_rows(rows):
        grouped = []
        idx_map = {}
        for r in rows:
            d_sched = re.sub(r"\s+", " ", extract_item_d(r["raw"].get("icaoMessage", "")).upper()).strip()
            key = (r["section"], r["aeropuerto"], r["is_red"], r["desde"], r["hasta"], d_sched)
            if key not in idx_map:
                nr = dict(r)
                nr["notams"] = [r["notam"]]
                nr["notam_inis"] = [r["notam_ini"]]
                nr["has_es"] = is_spanish_notam_row(r)
                idx_map[key] = len(grouped)
                grouped.append(nr)
            else:
                g = grouped[idx_map[key]]
                g["notams"].append(r["notam"])
                if r["notam_ini"] not in g["notam_inis"]:
                    g["notam_inis"].append(r["notam_ini"])
                g["dias"] = max(g["dias"], r["dias"])
                r_is_es = is_spanish_notam_row(r)
                # Par bilingüe (ej. Serie A en inglés + Serie C en español para el mismo evento): conservar SOLO el texto en español
                prefixes = {str(n)[0].upper() for n in g["notams"] if n}
                if ("A" in prefixes and "C" in prefixes) or (r_is_es != g.get("has_es", False)):
                    if r_is_es and not g.get("has_es", False):
                        g["texto"] = r["texto"]
                        g["has_es"] = True
                else:
                    # Fusionar limpiamente cierre de RWY + cierre de TWY sin repetir fechas/horarios
                    m1 = re.match(r"^(RWY\s+[0-9/LRC]+)\s+cerrada\s+(por trabajos de mantenimiento.*)$", g["texto"], re.IGNORECASE)
                    m2 = re.match(r"^(TWY\s+[A-Z0-9,\s/y]+?)\s+cerradas?\b", r["texto"], re.IGNORECASE)
                    if m1 and m2:
                        g["texto"] = f"{m1.group(1)} y {m2.group(1)} cerradas {m1.group(2)}"
                    elif r["texto"].strip().lower() != g["texto"].strip().lower():
                        g["texto"] = g["texto"].rstrip(".") + "; asimismo, " + r["texto"]
                # Limpiar referencias cruzadas redundantes entre los mismos NOTAMs ya unificados
                obs_parts = []
                for o in (g["obs"], r["obs"]):
                    if o and not any(nid_m in o for nid_m in g["notams"]):
                        obs_parts.append(o)
                g["obs"] = " ".join(obs_parts)
        return grouped

    unified_table_rows = unify_same_schedule_rows(verified_table_rows)

    # =========================================================================
    # CONSTRUIR FILAS HTML SEPARANDO NOVEDADES (ROJO) Y VIGENTES CONTINUOS
    # =========================================================================
    section_order = ["LIMA", "CUSCO", "AEROPUERTOS INTERNACIONALES", "AEROPUERTOS NACIONALES"]
    novedades_rows = [r for r in unified_table_rows if r["is_red"]]
    vigentes_rows = [r for r in unified_table_rows if not r["is_red"]]

    def render_row_html(r, idx, force_red=False):
        cls = "row-red" if (force_red or r["is_red"]) else ("row-alt" if idx % 2 == 1 else "row-norm")
        obs_html = html.escape(r["obs"]) if r["obs"] else '<span class="obs-empty">—</span>'
        links_html = "<br>".join(
            f'<a href="#{anchor_id(nid)}" class="tbl-link">{html.escape(nid)}</a>'
            for nid in r.get("notams", [r["notam"]])
        )
        inis_html = "<br>".join(html.escape(ni) for ni in r.get("notam_inis", [r["notam_ini"]]))
        return f"""
            <tr class="{cls}">
              <td class="col-ad">{html.escape(r['aeropuerto'])}</td>
              <td class="col-notam">{links_html}</td>
              <td class="col-txt">{html.escape(r['texto'])}</td>
              <td class="col-dt">{html.escape(r['desde'])}</td>
              <td class="col-dt">{html.escape(r['hasta'])}</td>
              <td class="col-obs">{obs_html}</td>
              <td class="col-ini">{inis_html}</td>
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
    for nid, *_ in spjc_day_items + extra_spjc_day + spjc_night_items + spjc_procs_spec:
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
    height: 34px;
    width: auto;
    flex-shrink: 0;
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
    grid-template-columns: 1fr 1.04fr;
    gap: 8px;
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
  .cards-row-6 {{ display: grid; grid-template-columns: repeat(5, 1fr); gap: 2.5px; margin-bottom: 2.5px; }}
  .cards-row-5 {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(90px, 1fr)); gap: 2.5px; margin-top: 2.5px; }}
  .cards-row-3 {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(220px, 1fr)); gap: 8px; margin-bottom: 8px; }}
  .procs-strip {{
    margin-top: 4px;
    border-top: 1.2px solid #cbd5e1;
    padding-top: 3px;
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
    padding: 2px 3.5px;
    line-height: 1.13;
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

  /* Planos oficiales SVG + etiquetas limpias */
  .canvas-wrap {{
    position: relative;
    width: 100%;
    background: #ffffff;
    border: 1px solid #cbd5e1;
    border-radius: 4px;
    overflow: hidden;
  }}
  .canvas-svg {{
    display: block;
    width: 100%;
    height: auto;
    max-height: 295px;
  }}
  .spzo-map-wrap {{
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

  /* Leyenda unificada por horario debajo de cada plano esquemático */
  .uleg-box {{
    margin-top: 4px;
    border: 1px solid #cbd5e1;
    border-radius: 4px;
    background: #ffffff;
    overflow: hidden;
  }}
  .uleg-box-blue {{
    border-color: #bfdbfe;
    background: #f8fafc;
  }}
  .uleg-box-orange {{
    border-color: #fdba74;
    background: #fffaf5;
  }}
  .uleg-hdr {{
    background: #fef2f2;
    border-bottom: 1px solid #fecaca;
    color: #991b1b;
    font-size: 7.2px;
    font-weight: 800;
    padding: 3px 6px;
    display: flex;
    justify-content: space-between;
    align-items: center;
    gap: 6px;
    flex-wrap: wrap;
  }}
  .uleg-hdr-blue {{
    background: #eff6ff;
    border-bottom-color: #bfdbfe;
    color: #1e3a8a;
  }}
  .uleg-hdr-orange {{
    background: #ffedd5;
    border-bottom-color: #fdba74;
    color: #c2410c;
  }}
  .uleg-subnote {{
    font-size: 6.6px;
    font-weight: 600;
    color: #475569;
  }}
  .uleg-tbl {{
    width: 100%;
    border-collapse: collapse;
    font-size: 7.0px;
  }}
  .uleg-tbl td {{
    border-bottom: 1px solid #f1f5f9;
    padding: 2.5px 5px;
    vertical-align: middle;
    line-height: 1.28;
  }}
  .uleg-tbl tr:last-child td {{ border-bottom: none; }}
  .uleg-cat {{
    width: 21%;
    font-weight: 800;
    color: #0f294a;
    background: #f8fafc;
    border-right: 1px solid #e2e8f0;
    white-space: nowrap;
  }}
  .uleg-chips {{
    color: #1e293b;
  }}
  .uchip {{
    display: inline-flex;
    align-items: center;
    gap: 2.5px;
    padding: 1.2px 4.5px;
    margin: 1px 3px 1px 0;
    border-radius: 3px;
    font-size: 6.9px;
    font-weight: 700;
    color: #0f172a;
    border: 1px solid #cbd5e1;
    background: #ffffff;
  }}
  .uchip-red {{
    border-color: #fca5a5;
    border-left: 2.5px solid #c53929;
    background: #fffafa;
  }}
  .uchip-red .notam-link {{ color: #b91c1c; text-decoration: underline; font-weight: 800; }}
  .uchip-amber, .uchip-orange {{
    border-color: #fdba74;
    border-left: 2.5px solid #ea580c;
    background: #fff7ed;
  }}
  .uchip-amber .notam-link, .uchip-orange .notam-link {{ color: #c2410c; text-decoration: underline; font-weight: 800; }}
  .uchip-blue {{
    border-color: #93c5fd;
    border-left: 2.5px solid #2563eb;
    background: #eff6ff;
  }}
  .uchip-blue .notam-link {{ color: #1d4ed8; text-decoration: underline; font-weight: 800; }}
  .uchip-dt {{
    font-weight: 600;
    color: #475569;
    font-size: 6.6px;
  }}
  .uleg-list {{
    padding: 4px 6px;
    font-size: 7.1px;
    line-height: 1.35;
  }}
  .uleg-item .notam-link {{ color: #b91c1c; font-weight: 800; text-decoration: underline; }}
  .uleg-empty {{
    padding: 6px 8px;
    font-size: 7.3px;
    color: #334155;
    line-height: 1.4;
  }}

  .legend-bar {{
    margin-bottom: 5px;
    font-size: 8.0px;
    font-weight: 600;
    color: #0f172a;
    background: #f1f5f9;
    border: 1px solid #cbd5e1;
    border-left: 3.5px solid #0f294a;
    border-radius: 4px;
    padding: 3.5px 8px;
    display: flex;
    justify-content: space-between;
    align-items: center;
  }}
  .dot-r {{ color: #c53929; font-weight: 800; }}
  .dot-a {{ color: #ea580c; font-weight: 800; }}
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
    padding: 3.2px 4.5px;
    border: 1px solid #cbd5e1;
    vertical-align: middle;
    line-height: 1.24;
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
    margin-top: 5px;
    display: flex;
    justify-content: space-between;
    align-items: flex-end;
    font-size: 7.5px;
    color: #334155;
    border-top: 1px solid #cbd5e1;
    padding-top: 4px;
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

  <div class="legend-bar">
    <div><span class="dot-r">■ ROJO:</span> cierre (superficie cerrada completa) &nbsp;|&nbsp; <span class="dot-a">● NARANJA:</span> obras / precaución (puntos aprox.) &nbsp;|&nbsp; <span class="dot-b">● AZUL:</span> procedimientos (puntos aprox.) &nbsp;|&nbsp; Clic en cada código NOTAM para texto OACI.</div>
    <div>Expresado en <strong>Hora Local (UTC-5)</strong> · Trazado referencial sobre AIRAC AIP SUP 34/2025</div>
  </div>

  <div class="spjc-grid">
    <!-- BLOQUE DIA -->
    <div class="panel">
      <div class="panel-hdr">
        <span class="panel-title">DÍA | 07:00 – 19:00 HORA LOCAL</span>
        <span class="panel-sub">Pistas, calles de rodaje, obras y procedimientos diurnos</span>
      </div>
      {render_spjc_day_canvas(active, spjc_bg_uri, spjc_day_items + extra_spjc_day, today_date, cutoff_local_dt)}
    </div>

    <!-- BLOQUE NOCHE -->
    <div class="panel">
      <div class="panel-hdr">
        <span class="panel-title">NOCHE | 19:00 – 07:00 HORA LOCAL</span>
        <span class="panel-sub">Obras y cierres nocturnos unificados por horario</span>
      </div>
      {render_spjc_night_canvas(active, spjc_bg_uri, spjc_night_items, today_date, cutoff_local_dt)}
    </div>
  </div>

  <!-- PROCEDIMIENTOS Y LIMITACIONES CONTINUAS SPJC -->
  <div class="procs-strip">
    <div class="procs-title">PROCEDIMIENTOS EN SUPERFICIE SPJC · PUNTOS AZULES EN EL PLANO (P1: TWY L · P2: L5 y D · P3: remolque · P4: segregadas · P5: V2 · P6: marcas 16L · P7: TWY L CAT D/E)</div>
    <div class="procs-grid">
      {''.join(spjc_procs_html)}
    </div>
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

  <div class="legend-bar" style="margin-bottom: 10px;">
    <div><span class="dot-r">■ ROJO:</span> cierre (superficie cerrada completa) &nbsp;|&nbsp; <span class="dot-a">● NARANJA:</span> obras / precaución (puntos aprox.) &nbsp;|&nbsp; <span class="dot-b">● AZUL:</span> procedimientos (puntos aprox.)</div>
    <div>Expresado en <strong>Hora Local (UTC-5)</strong> · Geometría manual referencial sobre AD 2 SPZO 09</div>
  </div>

  {render_spzo_canvas(active, spzo_bg_uri, spzo_cards_spec + extra_spzo_cards, today_date, cutoff_local_dt)}

  <div class="spzo-note">
    <strong>Observaciones operacionales SPZO:</strong> {spzo_obs_html}
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
      <div><strong>Verificación en vivo:</strong> {now_local.strftime('%d/%m/%y %H:%M')} · <strong>Total tabla:</strong> {len(unified_table_rows)} filas ({len(verified_table_rows)} NOTAMs)</div>
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
