#!/usr/bin/env python3
"""
wompi_sync.py — Extractor unificado Wompi VP
  Módulo A (Rollos): MySQL lineacom_analitica
                     v_rollos_tablero_wompi (todas las filas + columnas)
                     JOIN wompi_tablero_rollos_calculos por codigo_sitio
                     → data_tablero_rollos.json.gz → GitHub
  Módulo B (VP):     SharePoint Excel → data.json → GitHub + Correo operativo

Uso:
  python wompi_sync.py                         # ejecuta ambos módulos
  python wompi_sync.py --only-rollos           # solo Módulo A
  python wompi_sync.py --only-vp               # solo Módulo B
  python wompi_sync.py --no-push               # sin subida a GitHub
  python wompi_sync.py --no-mail               # sin correo (Módulo B)
  python wompi_sync.py --loop                  # repite cada INTERVAL_HOURS horas
"""

import os, sys, json, re, base64, gzip, hashlib, logging, smtplib, time, threading
from datetime import datetime, timedelta
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import formataddr
from pathlib import Path
from typing import Any, Dict, List, Optional

import requests
import pymysql, pymysql.cursors
import pandas as pd
from sqlalchemy import create_engine, text, inspect as sa_inspect

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
log = logging.getLogger("wompi_sync")

# ══════════════════════════════════════════════════════════════════
#  CONFIG COMPARTIDA
# ══════════════════════════════════════════════════════════════════
GITHUB_TOKEN  = os.getenv("GITHUB_TOKEN",  "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")
GITHUB_REPO   = os.getenv("GITHUB_REPO",   "segoloo/TablerosWompi")
GITHUB_BRANCH = os.getenv("GITHUB_BRANCH", "main")
INTERVAL_HOURS = int(os.getenv("INTERVAL_HOURS", "4"))
QUERY_TIMEOUT  = int(os.getenv("QUERY_TIMEOUT",  "180"))

# ── Módulo A: MySQL / Rollos ──────────────────────────────────────
DB_HOST     = os.getenv("DB_HOST",     "100.99.250.115")
DB_PORT     = int(os.getenv("DB_PORT", "3306"))
DB_USER     = os.getenv("DB_USER",     "root")
DB_PASSWORD = os.getenv("DB_PASSWORD", "An4l1t1c4l1n34*")
DB_NAME     = os.getenv("DB_NAME",     "lineacom_analitica")

# ── Módulo B: SharePoint / VP ─────────────────────────────────────
CLIENT_ID     = os.getenv("SP_CLIENT_ID",     "637ffa8d-43fc-485c-ba3d-e3ee2da6d6d6")
CLIENT_SECRET = os.getenv("SP_CLIENT_SECRET", "F4.8Q~P.DqObpdzw8wZmjNqCXVfjkEftNicoraIX")
TENANT_ID     = os.getenv("SP_TENANT_ID",     "af1a17b2-5d34-4f58-8b6c-6b94c6cd87ea")
SITE_ID       = os.getenv("SP_SITE_ID",       "6bc4f4bb-4479-45c0-a26b-3030893bc1c1")
SHEET_NAME    = os.getenv("SP_SHEET",         "Hoja1")
OAUTH_TOKEN_URL = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
GRAPH_BASE      = "https://graph.microsoft.com/v1.0"
OUTPUT_VP       = Path("data.json")

SMTP_SERVER    = "smtp-mail.outlook.com"
SMTP_PORT      = 587
EMAIL_USER     = "analitica@lineacom.co"
EMAIL_PASSWORD = "Linea.2024*"
TO_EMAILS      = ["mayra.llinares@wompi.com", "sebastian.gomez@lineacom.co", "lini.hernandez@lineacom.co", "tatiana.garcia@lineacom.co", "daniela.correa@lineacom.co"]
CC_EMAILS: List[str] = []

# ── Paleta Wompi 2025 ─────────────────────────────────────────────
C_NEGRO_CIB   = "#2C2A29"; C_BLANCO     = "#FAFAFA"; C_VERDE_MENTA = "#B0F2AE"
C_AZUL_CIELO  = "#99D1FC"; C_VERDE_SELVA= "#00825A"; C_VERDE_LIMA  = "#DFFF61"
C_CARD_DARK   = "#1E1C1B"; C_MUTED      = "#6B6967"; C_DANGER      = "#FF5C5C"
C_WARNING     = "#FFC04D"

# ── Constantes VP ────────────────────────────────────────────────
VT_EXACT   = "VISITA DATAFONO+KIT POP+CAPACITACION"
OPLG_EXACT = "ENVIO DATAFONO+KIT POP"
NOV_EXACT_COLS  = ["NOVEDADES","NOVEDAD","CAUSAL INCU","CAUSAL INC","RESPONSABLE INCUMPLIMIENTO","CAUSAL INCUMPLIMIENTO"]
NOV_KEY_INCLUDES = ["novedad","novedades","causal","responsable incump"]
NOV_KEY_EXCLUDES = ["estado","fecha","solicitud","comercio","guia","transpo","datafon","serial","departam","ciudad","tipolog","cumple","referencia","tipo de sol","id com"]
DATE_COLS = ["FECHA DE SOLICITUD","FECHA LIMITE DE ENTREGA","FECHA ENTREGA AL COMERCIO","FECHA VISITA TECNICA","FECHA DE ENTREGA"]


# ══════════════════════════════════════════════════════════════════
#  HELPERS COMUNES
# ══════════════════════════════════════════════════════════════════
def push_github(content_bytes: bytes, file_path: str, message: str, is_binary: bool = False):
    """Sube un archivo a GitHub. content_bytes puede ser texto o binario."""
    if not GITHUB_TOKEN:
        log.warning("GITHUB_TOKEN no configurado — se omite push.")
        return
    api_url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{file_path}"
    headers = {"Authorization": f"token {GITHUB_TOKEN}", "Accept": "application/vnd.github.v3+json"}
    sha = None
    r = requests.get(api_url, headers=headers, params={"ref": GITHUB_BRANCH}, timeout=30)
    if r.status_code == 200:
        sha = r.json().get("sha")
    body: Dict[str, Any] = {"message": message, "content": base64.b64encode(content_bytes).decode("ascii"), "branch": GITHUB_BRANCH}
    if sha:
        body["sha"] = sha
    r2 = requests.put(api_url, headers=headers, json=body, timeout=60)
    if r2.status_code in (200, 201):
        log.info(f"✅ {file_path} subido a GitHub ({GITHUB_REPO}/{GITHUB_BRANCH})")
    else:
        raise RuntimeError(f"GitHub PUT {r2.status_code}: {r2.text[:300]}")


# ══════════════════════════════════════════════════════════════════
#  MÓDULO A — ROLLOS (MySQL)
# ══════════════════════════════════════════════════════════════════
def _get_conn():
    return pymysql.connect(
        host=DB_HOST, port=DB_PORT, user=DB_USER, password=DB_PASSWORD, database=DB_NAME,
        charset="utf8mb4", cursorclass=pymysql.cursors.DictCursor,
        connect_timeout=20, read_timeout=QUERY_TIMEOUT + 10, write_timeout=30,
    )

def fetch(sql: str, label: str = "query") -> List[Dict]:
    result, error, conn_ref = [], [None], [None]
    def _run():
        try:
            c = _get_conn(); conn_ref[0] = c
            with c.cursor() as cur:
                cur.execute(sql); result.extend(cur.fetchall())
            c.close()
        except Exception as e:
            error[0] = e
    t0 = time.time()
    th = threading.Thread(target=_run, daemon=True); th.start(); th.join(timeout=QUERY_TIMEOUT)
    if th.is_alive():
        try:
            if conn_ref[0]: conn_ref[0].close()
        except Exception: pass
        log.warning(f"  ⚠ {label} TIMEOUT ({QUERY_TIMEOUT}s) — devolviendo vacío"); return []
    if error[0]:
        log.warning(f"  ⚠ {label} ERROR: {error[0]} — devolviendo vacío"); return []
    log.info(f"  ✓ {label}: {len(result)} filas en {time.time()-t0:.1f}s")
    return result

def _safe_val_mysql(v: Any) -> Any:
    if v is None: return ""
    if isinstance(v, datetime): return v.strftime("%Y-%m-%d %H:%M:%S")
    if hasattr(v, "isoformat"): return str(v)
    if isinstance(v, (int, float)): return v
    return str(v)

def safe_rows(rows: List[Dict]) -> List[Dict]:
    return [{k: _safe_val_mysql(v) for k, v in r.items()} for r in rows]

# ══════════════════════════════════════════════════════════════════
#  MÓDULO A — TABLERO ROLLOS COMPLETO (data_tablero_rollos.json.gz)
#
#  Fuente principal : v_rollos_tablero_wompi  → TODAS las filas y columnas
#  JOIN enriquecedor: wompi_tablero_rollos_calculos → por codigo_sitio
#                     (se resuelve codigo_sitio via v_rollos_trimestral_liquidar)
#
#  Resultado: una fila por cada fila de v_rollos_tablero_wompi,
#             con TODAS sus columnas originales más las 20 columnas de
#             wompi_tablero_rollos_calculos prefijadas con "cal_".
# ══════════════════════════════════════════════════════════════════

OUTPUT_TABLERO_ROLLOS = Path("data_tablero_rollos.json.gz")

# DESCONTINUADO — mantenido solo como referencia; ya no se genera
OUTPUT_ROLLOS = Path("data_rollos.json.gz")

# Query que trae TODAS las columnas originales de v_rollos_tablero_wompi
# sin renombrar, para preservar los nombres exactos de la vista
SQL_TABLERO_FULL = """
SELECT
    tarea,
    codigo_operacion,
    codigo_material,
    Cantidad,
    guia_raw,
    guia,
    transportadora,
    codigo_ubicacion_destino,
    codigo_ubicacion_origen,
    nombre_ubicacion_destino,
    nombre_ubicacion_origen,
    nombre_material,
    nombre_plantilla_tarea,
    nombre_plantilla_operacion,
    fecha_confirmacion,
    codigo_area_trabajo,
    nombre_sitio,
    Latitud,
    Longitud,
    nit,
    red_asociada,
    telefono,
    departamento,
    Ciudad,
    ciudad_sede,
    tipologia,
    tipologia_cb,
    zona_coordinador,
    direccion,
    tarea_fecha_fin,
    subproyecto,
    plan_inicio,
    plan_fin,
    proyecto,
    estado_tarea,
    flujo
FROM v_rollos_tablero_wompi"""

# Query auxiliar: tarea → codigo_sitio desde v_rollos_trimestral_liquidar
# Trae también nit y nombre_sitio para fallbacks de join
SQL_LIQ_SITIO = """
SELECT
    tarea              AS tarea,
    codigo_sitio       AS codigo_sitio,
    nit                AS nit,
    nombre_sitio       AS nombre_sitio
FROM v_rollos_trimestral_liquidar"""

# Tabla de cálculos de inventario (JOIN por codigo_sitio)
SQL_CALCULOS = """
SELECT
    id, tarea, codigo_mo, codigo_sitio,
    estado_punto, promedio_mensual, rollos_promedio_mes,
    periodo_abast_e5, valor_busqueda, rollos_periodo_abast_e5,
    rollos_anio_e5, punto_reorden, fecha_apertura_final,
    fecha_abst_1, rollos_entregados_mig_apert, trx_desde_migra_apert,
    rollos_consumidos_migr_apert, saldo_rollos, saldo_dias, saldo
FROM wompi_tablero_rollos_calculos"""


def build_tablero_rollos_json() -> Dict[str, Any]:
    """
    Construye el payload para data_tablero_rollos.json.gz:
    - Todas las filas y columnas de v_rollos_tablero_wompi
    - Enriquecidas con las 20 columnas de wompi_tablero_rollos_calculos (prefijo cal_)
      mediante join en cascada (4 niveles) para maximizar cobertura:
        N1: tarea exacta → codigo_sitio (via v_rollos_trimestral_liquidar)
        N2: codigo_ubicacion_destino == codigo_sitio en calculos
        N3: nit del tablero == valor_busqueda en calculos
        N4: nombre_sitio normalizado == nombre_sitio normalizado en calculos
    """
    t0 = time.time()
    log.info("═"*60 + "\n  BUILD TABLERO ROLLOS COMPLETO (data_tablero_rollos.json.gz)\n" + "═"*60)

    # ── Fetch de las 3 fuentes ────────────────────────────────────────
    tablero_full  = safe_rows(fetch(SQL_TABLERO_FULL, "v_rollos_tablero_wompi [full]"))
    liq_sitio     = safe_rows(fetch(SQL_LIQ_SITIO,   "v_rollos_trimestral_liquidar [sitio]"))
    calculos_rows = safe_rows(fetch(SQL_CALCULOS,     "wompi_tablero_rollos_calculos [tablero]"))

    log.info(f"  → tablero_full: {len(tablero_full)} filas | liq_sitio: {len(liq_sitio)} | calculos: {len(calculos_rows)}")

    # ── Helpers de normalización ──────────────────────────────────────
    def _norm(v: Any) -> str:
        """Normaliza un valor a string limpio para comparaciones fuzzy."""
        import unicodedata
        s = str(v or "").strip().upper()
        s = unicodedata.normalize("NFD", s)
        s = "".join(c for c in s if unicodedata.category(c) != "Mn")  # quitar tildes
        s = re.sub(r"\s+", " ", s)  # colapsar espacios
        return s

    # ── Lookup N1: tarea → codigo_sitio (via liquidar) ───────────────
    tarea_sitio_map: Dict[str, str] = {}
    # También construir nit→sitio y nombre→sitio desde liquidar para N3/N4
    nit_sitio_liq:    Dict[str, str] = {}
    nombre_sitio_liq: Dict[str, str] = {}
    for r in liq_sitio:
        t   = str(r.get("tarea")        or "").strip()
        cod = str(r.get("codigo_sitio") or "").strip()
        nit = str(r.get("nit")          or "").strip()
        nom = _norm(r.get("nombre_sitio"))
        if t and cod and t not in tarea_sitio_map:
            tarea_sitio_map[t] = cod
        if nit and cod and nit not in nit_sitio_liq:
            nit_sitio_liq[nit] = cod
        if nom and cod and nom not in nombre_sitio_liq:
            nombre_sitio_liq[nom] = cod

    # ── Lookup codigo_sitio → calculos (registro más reciente por id) ─
    calculos_sorted = sorted(
        calculos_rows,
        key=lambda x: int(x.get("id") or 0) if str(x.get("id") or "").isdigit() else 0,
        reverse=True,
    )
    calculos_map: Dict[str, Dict] = {}
    # Lookup N2: codigo_mo → calculos (codigo_mo suele ser el codigo del punto/MO)
    calculos_by_mo:     Dict[str, Dict] = {}
    # Lookup N3: valor_busqueda (NIT) → calculos
    calculos_by_vb:     Dict[str, Dict] = {}
    # Lookup N4: nombre_sitio normalizado → calculos
    calculos_by_nombre: Dict[str, Dict] = {}

    for r in calculos_sorted:
        cod = str(r.get("codigo_sitio") or "").strip()
        mo  = str(r.get("codigo_mo")    or "").strip()
        vb  = str(r.get("valor_busqueda") or "").strip()
        # nombre_sitio no está en calculos directamente; lo resolveremos vía liquidar

        if cod and cod not in calculos_map:
            calculos_map[cod] = r
        if mo and mo not in calculos_by_mo:
            calculos_by_mo[mo] = r
        if vb and vb not in calculos_by_vb:
            calculos_by_vb[vb] = r

    # N4: construir lookup nombre_sitio_norm → calculos usando liquidar como puente
    # (liquidar tiene nombre_sitio y codigo_sitio; calculos tiene codigo_sitio)
    for r in liq_sitio:
        nom = _norm(r.get("nombre_sitio"))
        cod = str(r.get("codigo_sitio") or "").strip()
        if nom and cod and nom not in calculos_by_nombre:
            cal = calculos_map.get(cod)
            if cal:
                calculos_by_nombre[nom] = cal

    log.info(f"  → Lookups construidos: "
             f"N1(tarea)={len(tarea_sitio_map):,} "
             f"N2(cod_mo)={len(calculos_by_mo):,} "
             f"N3(nit/vb)={len(calculos_by_vb):,} "
             f"N4(nombre)={len(calculos_by_nombre):,}")

    # ── Columnas de calculos que se adjuntan (todas las 20) ───────────
    CAL_COLS = [
        "id", "tarea", "codigo_mo", "codigo_sitio",
        "estado_punto", "promedio_mensual", "rollos_promedio_mes",
        "periodo_abast_e5", "valor_busqueda", "rollos_periodo_abast_e5",
        "rollos_anio_e5", "punto_reorden", "fecha_apertura_final",
        "fecha_abst_1", "rollos_entregados_mig_apert", "trx_desde_migra_apert",
        "rollos_consumidos_migr_apert", "saldo_rollos", "saldo_dias", "saldo",
    ]

    # ── Construir filas enriquecidas ──────────────────────────────────
    filas: List[Dict] = []
    nivel_hits = {1: 0, 2: 0, 3: 0, 4: 0, 0: 0}  # 0 = sin match

    for r in tablero_full:
        tarea          = str(r.get("tarea")                    or "").strip()
        cod_ub_dest    = str(r.get("codigo_ubicacion_destino") or "").strip()
        nit_tab        = str(r.get("nit")                      or "").strip()
        nombre_tab_n   = _norm(r.get("nombre_sitio"))

        cod_sitio = ""
        cal       = {}
        nivel     = 0

        # N1: tarea → codigo_sitio (via liquidar) → calculos
        if tarea in tarea_sitio_map:
            cs = tarea_sitio_map[tarea]
            if cs in calculos_map:
                cod_sitio, cal, nivel = cs, calculos_map[cs], 1

        # N2: codigo_ubicacion_destino == codigo_sitio en calculos
        if not cal and cod_ub_dest and cod_ub_dest in calculos_map:
            cod_sitio, cal, nivel = cod_ub_dest, calculos_map[cod_ub_dest], 2

        # N2b: codigo_ubicacion_destino == codigo_mo en calculos
        if not cal and cod_ub_dest and cod_ub_dest in calculos_by_mo:
            c = calculos_by_mo[cod_ub_dest]
            cod_sitio = str(c.get("codigo_sitio") or "").strip()
            cal, nivel = c, 2

        # N3: nit del tablero == valor_busqueda en calculos
        if not cal and nit_tab and nit_tab in calculos_by_vb:
            c = calculos_by_vb[nit_tab]
            cod_sitio = str(c.get("codigo_sitio") or "").strip()
            cal, nivel = c, 3

        # N4: nombre_sitio normalizado
        if not cal and nombre_tab_n and nombre_tab_n in calculos_by_nombre:
            c = calculos_by_nombre[nombre_tab_n]
            cod_sitio = str(c.get("codigo_sitio") or "").strip()
            cal, nivel = c, 4

        nivel_hits[nivel] += 1

        fila: Dict = {}
        # Todas las columnas originales de v_rollos_tablero_wompi
        for col, val in r.items():
            fila[col] = val
        # codigo_sitio resuelto + nivel de resolución (para diagnóstico)
        fila["codigo_sitio"]    = cod_sitio
        fila["join_nivel"]      = nivel  # 1-4 = nivel que dio match, 0 = sin match
        # Columnas de wompi_tablero_rollos_calculos (prefijo cal_)
        for col in CAL_COLS:
            val = cal.get(col, "")
            if isinstance(val, datetime):
                val = val.strftime("%Y-%m-%d %H:%M:%S")
            fila[f"cal_{col}"] = val

        filas.append(fila)

    total      = len(tablero_full)
    con_cal    = sum(1 for n, c in nivel_hits.items() if n > 0 for _ in range(c))
    con_cal    = total - nivel_hits[0]
    pct_sitio  = round(con_cal / total * 100, 1) if total else 0
    pct_cal    = pct_sitio  # en esta lógica cal sigue a cod_sitio

    log.info(f"  ✓ Filas procesadas   : {total:,}")
    log.info(f"  ✓ Con calculos       : {con_cal:,} / {total:,} ({pct_sitio}%)")
    log.info(f"  ✓ Desglose por nivel :")
    log.info(f"       N1 tarea→sitio  : {nivel_hits[1]:,}")
    log.info(f"       N2 cod_ub_dest  : {nivel_hits[2]:,}")
    log.info(f"       N3 nit/val_busq : {nivel_hits[3]:,}")
    log.info(f"       N4 nombre_sitio : {nivel_hits[4]:,}")
    log.info(f"       Sin match       : {nivel_hits[0]:,}")
    log.info(f"  ✓ Columnas por fila  : {len(filas[0]) if filas else 0} "
             f"(36 tablero + codigo_sitio + join_nivel + 20 cal_*)")
    log.info(f"  ✓ Build completado en {time.time()-t0:.1f}s")

    return {
        "generado"         : datetime.now().strftime("%d/%m/%Y %H:%M"),
        "fuente"           : "v_rollos_tablero_wompi JOIN wompi_tablero_rollos_calculos (join cascada N1-N4)",
        "total_filas"      : len(filas),
        "pct_join_sitio"   : pct_sitio,
        "pct_join_calculos": pct_cal,
        "join_niveles"     : {
            "n1_tarea"     : nivel_hits[1],
            "n2_cod_ub"    : nivel_hits[2],
            "n3_nit"       : nivel_hits[3],
            "n4_nombre"    : nivel_hits[4],
            "sin_match"    : nivel_hits[0],
        },
        "filas": filas,
    }


def run_rollos(no_push: bool = False):
    start = datetime.now()
    log.info("═"*60 + f"\n  SYNC ROLLOS WOMPI VP — {start.strftime('%Y-%m-%d %H:%M:%S')}\n" + "═"*60)

    # ── Único JSON de rollos: tablero completo ────────────────────────
    # Todas las filas de v_rollos_tablero_wompi + JOIN a wompi_tablero_rollos_calculos
    payload_t    = build_tablero_rollos_json()
    json_bytes_t = json.dumps(payload_t, ensure_ascii=False, separators=(",", ":")).encode("utf-8")
    gz_bytes_t   = gzip.compress(json_bytes_t, compresslevel=9)
    OUTPUT_TABLERO_ROLLOS.write_bytes(gz_bytes_t)
    log.info(f"data_tablero_rollos.json.gz → {payload_t['total_filas']} filas | "
             f"{len(json_bytes_t)//1024}KB JSON → {len(gz_bytes_t)//1024}KB gz")
    if not no_push:
        push_github(gz_bytes_t, "data_tablero_rollos.json.gz",
                    f"sync: tablero rollos {datetime.now().strftime('%Y-%m-%d %H:%M')}", is_binary=True)
    else:
        log.info("--no-push activo: se omite subida a GitHub.")

    log.info(f"✅ Rollos completado en {(datetime.now()-start).total_seconds():.1f}s")


# ══════════════════════════════════════════════════════════════════
#  MÓDULO B — VP (SharePoint + Correo)
# ══════════════════════════════════════════════════════════════════
class TokenCache:
    def __init__(self): self._token: Optional[str] = None; self._expiry: Optional[datetime] = None
    def get(self) -> str:
        if self._token and self._expiry and datetime.now() < self._expiry: return self._token
        log.info("Obteniendo token Graph API...")
        r = requests.post(OAUTH_TOKEN_URL, data={"client_id": CLIENT_ID, "client_secret": CLIENT_SECRET,
            "scope": "https://graph.microsoft.com/.default", "grant_type": "client_credentials"}, timeout=60)
        if r.status_code != 200: raise RuntimeError(f"Error token Graph API: {r.status_code} — {r.text[:200]}")
        j = r.json(); self._token = j["access_token"]; self._expiry = datetime.now() + timedelta(seconds=3500)
        log.info("Token obtenido ✅"); return self._token

_token_cache = TokenCache()


def find_workbook_item_id() -> str:
    token = _token_cache.get(); headers = {"Authorization": f"Bearer {token}"}
    for kw in ["Comodato", "VP", "Operaci"]:
        url = f"{GRAPH_BASE}/sites/{SITE_ID}/drive/root/search(q='{requests.utils.quote(kw, safe='')}')"
        r = requests.get(url, headers=headers, timeout=60)
        if r.status_code == 200:
            for item in r.json().get("value", []):
                if any(k.lower() in item.get("name","").lower() for k in ["Comodato","VP_Comodato"]):
                    log.info(f"Archivo encontrado: {item['name']}"); return item["id"]
    r2 = requests.get(f"{GRAPH_BASE}/sites/{SITE_ID}/drive/root/children", headers=headers, timeout=60)
    if r2.status_code == 200:
        for item in r2.json().get("value", []):
            if any(k.lower() in item.get("name","").lower() for k in ["Comodato","VP"]):
                log.info(f"Archivo encontrado (raíz): {item['name']}"); return item["id"]
    raise RuntimeError("No se encontró 'Operación VP_Comodato.xlsx' en SharePoint.")


def read_sheet(item_id: str) -> pd.DataFrame:
    token = _token_cache.get(); headers = {"Authorization": f"Bearer {token}"}
    url = f"{GRAPH_BASE}/sites/{SITE_ID}/drive/items/{item_id}/workbook/worksheets/{requests.utils.quote(SHEET_NAME, safe='')}/usedRange"
    log.info(f"Leyendo hoja '{SHEET_NAME}'...")
    r = requests.get(url, headers=headers, timeout=120)
    if r.status_code != 200: raise RuntimeError(f"Error leyendo hoja '{SHEET_NAME}': {r.status_code}\n{r.text[:400]}")
    values = r.json().get("values", [])
    if not values or len(values) < 2: raise ValueError("La hoja está vacía o sin datos.")
    cols = [str(x).strip() if x else f"col_{i}" for i, x in enumerate(values[0])]
    rows = []
    for row in values[1:]:
        norm = [c if c is not None else "" for c in row]
        if len(norm) < len(cols): norm += [""] * (len(cols) - len(norm))
        rows.append(norm[:len(cols)])
    df = pd.DataFrame(rows, columns=cols)
    log.info(f"{len(df)} filas leídas ✅"); return df


def _parse_date_str(val: Any) -> str:
    try:
        if pd.isna(val): return ""
    except Exception: pass
    s = str(val).strip()
    if s in ("", "nan", "NaN", "None"): return ""
    s_stripped = s.replace(".", "", 1)
    if s_stripped.isdigit():
        try: return (datetime(1899, 12, 30) + timedelta(days=int(float(s)))).strftime("%d/%m/%Y")
        except Exception: return s
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%m/%d/%Y", "%d/%m/%y", "%Y-%m-%dT%H:%M:%S", "%Y-%m-%dT%H:%M"):
        try: return datetime.strptime(s, fmt).strftime("%d/%m/%Y")
        except ValueError: pass
    try: return pd.to_datetime(s, dayfirst=True).strftime("%d/%m/%Y")
    except Exception: return s


def clean_df(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy(); df.columns = [str(c).strip() for c in df.columns]
    for col in df.columns:
        if df[col].dtype == object:
            df[col] = df[col].astype(str).str.strip().replace({"nan": "", "None": "", "NaN": ""})
    for col in DATE_COLS:
        if col in df.columns: df[col] = df[col].apply(_parse_date_str)
    return df.dropna(how="all").loc[~(df == "").all(axis=1)]


def _safe_val_sp(v: Any) -> Any:
    try:
        if pd.isna(v): return ""
    except Exception: pass
    return v if isinstance(v, (int, float)) else str(v)

def df_to_rows(df: pd.DataFrame) -> List[Dict]: return [{c: _safe_val_sp(row[c]) for c in df.columns} for _, row in df.iterrows()]


def build_vp_json(df: pd.DataFrame) -> Dict[str, Any]:
    df_clean = clean_df(df); rows = df_to_rows(df_clean)
    return {"generado": datetime.now().strftime("%d/%m/%Y %H:%M"), "filas": len(rows), "columnas": list(df_clean.columns), "rows": rows}


# ── Helpers KPI correo VP ─────────────────────────────────────────
def get_col(row: dict, *keys) -> str:
    for k in keys:
        if k in row and str(row[k]).strip(): return str(row[k]).strip()
        for rk in row:
            if rk.upper() == k.upper() and str(row[rk]).strip(): return str(row[rk]).strip()
    return ""

def parse_date(s) -> Optional[datetime]:
    if not s: return None
    s = str(s).strip()
    if not s or s.upper() in ("NAN", ""): return None
    for pattern, fmt in [
        (r"^(\d{1,2})/(\d{1,2})/(\d{4})$", lambda m: datetime(int(m.group(3)), int(m.group(2)), int(m.group(1)))),
        (r"^(\d{4})-(\d{2})-(\d{2})", lambda m: datetime(int(m.group(1)), int(m.group(2)), int(m.group(3)))),
        (r"^(\d{1,2})-(\d{1,2})-(\d{4})$", lambda m: datetime(int(m.group(3)), int(m.group(2)), int(m.group(1)))),
    ]:
        m = re.match(pattern, s)
        if m:
            try:
                d = fmt(m)
                if 2000 < d.year < 2100: return d
            except ValueError: pass
    s_stripped = s.replace(".", "", 1)
    if s_stripped.isdigit():
        try:
            d = datetime(1899, 12, 30) + timedelta(days=int(float(s)))
            if 2000 < d.year < 2100: return d
        except Exception: pass
    for fmt in ("%d/%m/%y", "%m/%d/%Y"):
        try:
            d = datetime.strptime(s, fmt)
            if 2000 < d.year < 2100: return d
        except ValueError: pass
    return None

def get_fecha_limite(row: dict) -> Optional[datetime]:
    for col in ["FECHA LIMITE DE ENTREGA","fecha limite de entrega","FECHA LÍMITE DE ENTREGA","fecha límite de entrega","FECHA LIMITE","fecha limite"]:
        d = parse_date(get_col(row, col))
        if d: return d
    for k in row:
        kl = k.lower()
        if ("limite" in kl or "límite" in kl) and "solicitud" not in kl and "comercio" not in kl:
            d = parse_date(str(row[k] or ""))
            if d: return d
    return None

def find_novedad(row: dict) -> Optional[Dict[str, str]]:
    for col in NOV_EXACT_COLS:
        v = get_col(row, col).strip()
        if v and v != "0" and v.upper() != "NAN": return {"col": col, "val": v}
    for k in row:
        kl = k.lower()
        if any(n in kl for n in NOV_KEY_INCLUDES) and not any(x in kl for x in NOV_KEY_EXCLUDES):
            v = str(row[k] or "").strip()
            if v and v != "0" and v.upper() != "NAN": return {"col": k, "val": v}
    return None

def is_devolucion(row: dict) -> bool:
    for k, v in row.items():
        vu = str(v or "").upper()
        if any(x in vu for x in ["DEVOLUCI","DEVOLUCIÓN","DEVUELTO","REMITENTE"]): return True
        if any(x in k.upper() for x in ["DEVOLUCI","DEVOLUCIÓN"]) and vu and vu not in ("","0","NAN"): return True
    return False


def compute_kpis_for_email(df: pd.DataFrame) -> Dict[str, Any]:
    all_rows = [{str(k).strip(): v for k, v in row.items()} for _, row in df.iterrows()]
    data = [r for r in all_rows if get_col(r, "ESTADO DATAFONO", "estado datafono").upper() != "CANCELADO"]
    cancelados = len(all_rows) - len(data); total = len(data)

    ec: Dict[str, int] = {}
    for r in data:
        e = get_col(r, "ESTADO DATAFONO", "estado datafono").upper() or "SIN ESTADO"
        ec[e] = ec.get(e, 0) + 1

    entregados      = ec.get("ENTREGADO", 0)
    en_transito     = ec.get("EN TRANSITO", 0) + ec.get("EN TRÁNSITO", 0)
    en_alistamiento = ec.get("EN ALISTAMIENTO", 0)
    devueltos       = len([r for r in data if is_devolucion(r)])
    n_alistados     = total - en_alistamiento

    def tipo_rows(tipo): return [r for r in data if get_col(r,"TIPO DE SOLICITUD FACTURACIÓN","TIPO DE SOLICITUD FACTURACION","tipo de solicitud facturacion").upper() == tipo.upper()]
    vt_rows = tipo_rows(VT_EXACT); ol_rows = tipo_rows(OPLG_EXACT)
    def ent(rows): return len([r for r in rows if get_col(r,"ESTADO DATAFONO","estado datafono").upper() == "ENTREGADO"])
    ent_vt = ent(vt_rows); ent_ol = ent(ol_rows)
    programados_vt = len([r for r in vt_rows if get_col(r,"ESTADO DATAFONO","estado datafono").upper() == "VISITA PROGRAMADA"])

    ent_df       = [r for r in data if get_col(r,"ESTADO DATAFONO","estado datafono").upper() == "ENTREGADO"]
    cumple_oport = len([r for r in ent_df if get_col(r,"CUMPLE ANS","cumple ans","Cumple Ans").upper() == "SI"])
    pct_oport    = round(cumple_oport / len(ent_df) * 100) if ent_df else 0
    pct_calidad  = round((entregados - devueltos) / entregados * 100) if entregados else 100
    pct = lambda a, b: round(a / b * 100) if b else 0

    today = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)

    def get_pending(r): return get_col(r,"ESTADO DATAFONO","estado datafono").upper() not in ("ENTREGADO","CANCELADO")

    vencen_hoy_rows, vencidas_rows, backlog_24h_rows, backlog_48h_rows = [], [], [], []
    guias_sin_cambios_rows, intentos_fallidos_rows = [], []

    for r in data:
        est = get_col(r, "ESTADO DATAFONO", "estado datafono").upper()
        lim = get_fecha_limite(r)
        if est in ("ENTREGADO","CANCELADO"): pass
        elif lim is not None:
            lim_day = lim.replace(hour=0, minute=0, second=0, microsecond=0)
            if lim_day == today: vencen_hoy_rows.append(r)
            if lim < today: vencidas_rows.append(r)
            lim_end = lim.replace(hour=23, minute=59, second=59, microsecond=999999)
            if lim_end >= today and lim <= today + timedelta(hours=24): backlog_24h_rows.append(r)
            if lim_end >= today and lim <= today + timedelta(hours=48): backlog_48h_rows.append(r)
            dias = (today - lim_day).days
            if lim_day <= today and dias >= 1 and find_novedad(r): guias_sin_cambios_rows.append({**r, "_dias_sin_cambios": dias})
        if est in ("EN TRANSITO","EN TRÁNSITO") and find_novedad(r): intentos_fallidos_rows.append(r)

    def row_info(r): 
        lim = get_fecha_limite(r)
        return {"comercio": get_col(r,"Nombre del comercio","nombre del comercio","NOMBRE DEL COMERCIO") or "—",
                "id_sitio": get_col(r,"ID Comercio","id comercio") or "—",
                "guia": get_col(r,"NÚMERO DE GUIA","NUMERO DE GUIA","numero de guia") or "—",
                "transportadora": get_col(r,"TRANSPORTADORA","Transportadora","transportadora") or "—",
                "estado": get_col(r,"ESTADO DATAFONO","estado datafono") or "—",
                "fecha_limite": lim.strftime("%d/%m/%Y") if lim else "—",
                "tipo": get_col(r,"TIPO DE SOLICITUD FACTURACIÓN","TIPO DE SOLICITUD FACTURACION","tipo de solicitud facturacion") or "—",
                "novedad": (find_novedad(r) or {}).get("val","") or "—"}

    return {
        "fecha_informe": datetime.now().strftime("%d/%m/%Y %H:%M"), "total": total, "cancelados": cancelados,
        "entregados": entregados, "en_transito": en_transito, "en_alistamiento": en_alistamiento,
        "n_alistados": n_alistados, "devueltos": devueltos, "total_vt": len(vt_rows), "ent_vt": ent_vt,
        "programados_vt": programados_vt, "total_ol": len(ol_rows), "ent_ol": ent_ol,
        "pct_entregado": pct(entregados,total), "pct_transito": pct(en_transito,total),
        "pct_alistamiento": pct(en_alistamiento,total), "pct_n_alistados": pct(n_alistados,total),
        "pct_vt": pct(ent_vt,len(vt_rows)), "pct_ol": pct(ent_ol,len(ol_rows)),
        "pct_oport": pct_oport, "pct_calidad": pct_calidad,
        "vencen_hoy": len(vencen_hoy_rows), "vencen_hoy_data": [row_info(r) for r in vencen_hoy_rows],
        "vencidas": len(vencidas_rows), "vencidas_data": [row_info(r) for r in vencidas_rows],
        "backlog_24h": len(backlog_24h_rows), "backlog_24h_data": [row_info(r) for r in backlog_24h_rows],
        "backlog_48h": len(backlog_48h_rows), "backlog_48h_data": [row_info(r) for r in backlog_48h_rows],
        "guias_sin_cambios": len(guias_sin_cambios_rows),
        "guias_sin_cambios_data": [{**row_info(r), "dias": r.get("_dias_sin_cambios",0)} for r in guias_sin_cambios_rows],
        "intentos_fallidos": len(intentos_fallidos_rows), "intentos_fallidos_data": [row_info(r) for r in intentos_fallidos_rows],
    }


# ── HTML Correo VP ────────────────────────────────────────────────
def _truncar(texto: str, max_chars: int = 60) -> str:
    t = str(texto or "—").strip()
    return t if len(t) <= max_chars else t[:max_chars] + "…"

def _rows_table_html(rows: list, extra_cols=None, max_rows=20) -> str:
    if not rows:
        return ('<table width="100%" cellpadding="0" cellspacing="0" border="0">'
                '<tr><td align="center" style="padding:18px 12px;background-color:#1C1F1C;border:1px dashed #2E3C2E">'
                '<p style="margin:0;color:#4A7A4A;font-size:20px">&#10003;</p>'
                '<p style="margin:6px 0 0;color:#6B6967;font-size:12px;font-family:Arial,Helvetica,sans-serif">Sin registros en este indicador</p>'
                '</td></tr></table>')

    base_cols = [("comercio","Comercio"),("guia","Gu&#237;a"),("transportadora","Transp."),("estado","Estado"),("fecha_limite","L&#237;mite")]
    extra_raw = extra_cols or [("novedad","Novedad")]
    cols = base_cols + [(e[0], e[1]) if len(e) >= 2 else (e[0], e[0]) for e in extra_raw]
    visible = rows[:max_rows]; omitted = len(rows) - len(visible)

    th = "".join(f'<th style="padding:7px 10px;text-align:left;background-color:#162116;color:#B0F2AE;font-weight:700;font-size:11px;font-family:Arial,Helvetica,sans-serif;border-bottom:2px solid #00825A;white-space:nowrap">{lb}</th>' for _,lb in cols)

    body_rows = ""
    for i, r in enumerate(visible):
        bg = "#191C19" if i % 2 == 0 else "#1D201D"
        cells = ""
        for key, _ in cols:
            raw = r.get(key, "—") or "—"
            if key == "estado":
                eu = str(raw).upper()
                c = ("#B0F2AE" if "ENTREGADO" in eu else "#99D1FC" if "TRANSITO" in eu else "#FFC04D" if ("ALISTAMIENTO" in eu or "PROGRAMADA" in eu) else "#FF5C5C")
                short = (str(raw)[:13]+"…") if len(str(raw))>14 else str(raw)
                cell_v = f'<span style="color:{c};font-size:10px;font-weight:700;font-family:Arial,Helvetica,sans-serif">{short}</span>'
            elif key == "dias":
                d_val = r.get("dias", 0); col = "#FF5C5C" if d_val >= 7 else "#FFC04D" if d_val >= 3 else "#99D1FC"
                cell_v = f'<span style="color:{col};font-weight:700;font-size:11px;font-family:Arial,Helvetica,sans-serif">{d_val}d</span>'
            elif key == "novedad":
                tf = str(raw).strip(); tc = _truncar(tf, 50)
                cell_v = '<span style="color:#4A4845;font-size:11px;font-family:Arial,Helvetica,sans-serif">&#8212;</span>' if tf in ("—","") else f'<span style="color:#FFC04D;font-size:10px;font-family:Arial,Helvetica,sans-serif">{tc}</span>'
            elif key == "comercio":
                cell_v = f'<span style="color:#FAFAFA;font-size:11px;font-weight:600;font-family:Arial,Helvetica,sans-serif">{_truncar(raw,22)}</span>'
            elif key == "guia":
                cell_v = f'<span style="color:#B0F2AE;font-size:11px;font-weight:600;font-family:Courier New,Courier,monospace">{_truncar(raw,16)}</span>'
            elif key == "transportadora":
                cell_v = f'<span style="color:#99D1FC;font-size:11px;font-family:Arial,Helvetica,sans-serif">{_truncar(raw,11)}</span>'
            else:
                cell_v = f'<span style="color:#DEDBD8;font-size:11px;font-family:Arial,Helvetica,sans-serif">{_truncar(raw,14)}</span>'
            cells += f'<td style="padding:7px 10px;border-bottom:1px solid #232623;vertical-align:middle">{cell_v}</td>'
        body_rows += f'<tr style="background-color:{bg}">{cells}</tr>'

    extra_note = ""
    if omitted > 0:
        extra_note = (f'<tr><td colspan="{len(cols)}" style="text-align:center;padding:10px;color:#6B6967;font-size:11px;background-color:#141614;border-top:1px solid #232623;font-family:Arial,Helvetica,sans-serif">'
                      f'&#8230;y <strong style="color:#B0F2AE">{omitted}</strong> registros m&#225;s. '
                      f'<a href="https://segoloo.github.io/TablerosWompi/" style="color:#B0F2AE;text-decoration:none;font-weight:600">Ver todos &#8594;</a></td></tr>')

    return (f'<table width="100%" cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse;border:1px solid #232623">'
            f'<thead><tr>{th}</tr></thead><tbody>{body_rows}{extra_note}</tbody></table>')

def _kpi_card_html(value,label,color,icon,sub=""):
    sub_html = f'<p style="margin:2px 0 0;font-size:9px;color:{color};font-family:Arial,Helvetica,sans-serif">{sub}</p>' if sub else ""
    return (f'<td style="width:14%;padding:4px;vertical-align:top"><table width="100%" cellpadding="0" cellspacing="0" border="0" '
            f'style="background-color:#1A1F1A;border-top:3px solid {color}"><tr><td align="center" style="padding:12px 8px 10px">'
            f'<p style="margin:0 0 5px;font-size:18px;line-height:1">{icon}</p>'
            f'<p style="margin:0 0 4px;font-size:19px;font-weight:bold;color:{color};font-family:Georgia,serif;line-height:1">{value}</p>'
            f'<p style="margin:0;font-size:9px;color:#6B6967;text-transform:uppercase;letter-spacing:0.8px;font-family:Arial,Helvetica,sans-serif">{label}</p>'
            f'{sub_html}</td></tr></table></td>')

def _progress_bar(pct,color):
    filled = max(0,min(100,pct)); empty = 100-filled
    return (f'<table width="100%" cellpadding="0" cellspacing="0" border="0" style="margin-top:5px"><tr>'
            f'<td width="{filled}%" style="height:5px;background-color:{color};font-size:1px;line-height:1px">&nbsp;</td>'
            + (f'<td width="{empty}%" style="height:5px;background-color:#232623;font-size:1px;line-height:1px">&nbsp;</td>' if empty > 0 else "")
            + f'</tr></table>')

def _alert_section_html(title,subtitle,color,icon,count,table_html,alert=False):
    border = "#FF5C5C" if alert and count>0 else color
    hdr_bg = "#1E1210" if alert and count>0 else "#141A14"
    count_bg = "#FF5C5C" if alert and count>0 else color
    count_fg = "#FAFAFA" if alert and count>0 else "#0A0A0A"
    return (f'<table width="100%" cellpadding="0" cellspacing="0" border="0" style="margin:16px 0;border:1px solid {border}">'
            f'<tr style="background-color:{hdr_bg}"><td style="padding:12px 16px;border-bottom:1px solid {border}">'
            f'<span style="font-size:15px;vertical-align:middle">{icon}</span>&nbsp;'
            f'<span style="font-size:13px;font-weight:bold;color:{color};vertical-align:middle;font-family:Arial,Helvetica,sans-serif">{title}</span>'
            f'<br><span style="font-size:10px;color:#6B6967;font-family:Arial,Helvetica,sans-serif">{subtitle}</span></td>'
            f'<td align="right" style="padding:12px 16px;border-bottom:1px solid {border};white-space:nowrap;vertical-align:middle">'
            f'<span style="background-color:{count_bg};color:{count_fg};font-weight:bold;font-size:16px;padding:4px 12px;font-family:Georgia,serif">{count}</span>'
            f'</td></tr><tr><td colspan="2" style="padding:12px 16px">{table_html}</td></tr></table>')

def build_html_email(kpis: Dict[str, Any]) -> str:
    fecha = kpis["fecha_informe"]; dashboard_url = "https://segoloo.github.io/TablerosWompi/"
    total_alert = kpis["vencen_hoy"] + kpis["vencidas"] + kpis["guias_sin_cambios"] + kpis["intentos_fallidos"]
    alert_color = C_DANGER if total_alert>5 else C_WARNING if total_alert>0 else C_VERDE_MENTA
    alert_bg    = "#1E1210" if total_alert>5 else "#1E1A10" if total_alert>0 else "#101E10"
    alert_icon  = "🚨" if total_alert>5 else "⚠️" if total_alert>0 else "✅"
    alert_msg   = (f"ALERTA: {total_alert} registros requieren atención inmediata" if total_alert>5 else
                   f"Atención: {total_alert} registros en seguimiento" if total_alert>0 else "Todo en orden — sin alertas activas")

    s_hoy      = _alert_section_html("Gu&#237;as que vencen hoy","Sin entregar con l&#237;mite hoy",C_WARNING,"&#9200;",kpis["vencen_hoy"],_rows_table_html(kpis["vencen_hoy_data"]),alert=kpis["vencen_hoy"]>0)
    s_vencidas = _alert_section_html("Vencidas ANS","Fuera de plazo (VT + OPLG)",C_DANGER,"&#128680;",kpis["vencidas"],_rows_table_html(kpis["vencidas_data"]),alert=kpis["vencidas"]>0)
    s_bl24     = _alert_section_html("Backlog por vencer (24h)","Gu&#237;as cuyo l&#237;mite vence en las pr&#243;ximas 24 h",C_AZUL_CIELO,"&#9889;",kpis["backlog_24h"],_rows_table_html(kpis["backlog_24h_data"]))
    s_bl48     = _alert_section_html("Backlog por vencer (48h)","Gu&#237;as cuyo l&#237;mite vence en las pr&#243;ximas 48 h",C_AZUL_CIELO,"&#128197;",kpis["backlog_48h"],_rows_table_html(kpis["backlog_48h_data"]))
    s_stalled  = _alert_section_html("Gu&#237;as sin cambios &gt;24h","Vencidas ANS con novedad registrada",C_WARNING,"&#128274;",kpis["guias_sin_cambios"],_rows_table_html(kpis["guias_sin_cambios_data"],extra_cols=[("dias","D&#237;as"),("novedad","Novedad")]),alert=kpis["guias_sin_cambios"]>0)
    s_fallidos = _alert_section_html("Gu&#237;as con intento fallido","En tr&#225;nsito con novedad de intento fallido",C_DANGER,"&#10060;",kpis["intentos_fallidos"],_rows_table_html(kpis["intentos_fallidos_data"],extra_cols=[("novedad","Novedad")]),alert=kpis["intentos_fallidos"]>0)

    return f"""<!DOCTYPE html>
<html lang="es" xmlns="http://www.w3.org/1999/xhtml">
<head>
  <meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Reporte Operativo Wompi VP</title>
  <style type="text/css">
    body,table,td,p,a{{-ms-text-size-adjust:100%;-webkit-text-size-adjust:100%;}}
    table,td{{mso-table-lspace:0pt;mso-table-rspace:0pt;}}
    body{{margin:0!important;padding:0!important;background-color:#0C0F0C;}}
  </style>
</head>
<body style="margin:0;padding:0;background-color:#0C0F0C;font-family:Arial,Helvetica,sans-serif">
<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="background-color:#0C0F0C">
<tr><td align="center" style="padding:20px 10px">
  <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="620" style="max-width:620px;width:100%">
    <tr>
      <td width="33%" height="4" style="background-color:{C_VERDE_LIMA};font-size:1px;line-height:4px">&nbsp;</td>
      <td width="34%" height="4" style="background-color:{C_VERDE_MENTA};font-size:1px;line-height:4px">&nbsp;</td>
      <td width="33%" height="4" style="background-color:{C_AZUL_CIELO};font-size:1px;line-height:4px">&nbsp;</td>
    </tr>
    <tr>
      <td colspan="3" bgcolor="{C_NEGRO_CIB}" style="background-color:{C_NEGRO_CIB}">
        <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%">
          <tr>
            <td style="padding:22px 24px 14px;vertical-align:top">
              <p style="margin:0 0 3px;font-size:10px;font-weight:bold;color:{C_VERDE_LIMA};letter-spacing:2px;text-transform:uppercase;font-family:Arial,Helvetica,sans-serif">Reporte Operativo</p>
              <p style="margin:0 0 4px;font-size:28px;font-weight:bold;color:{C_VERDE_MENTA};letter-spacing:-0.5px;line-height:1.1;font-family:Georgia,Times New Roman,serif">Wompi VP</p>
              <p style="margin:0;font-size:11px;color:#5A5856;font-family:Arial,Helvetica,sans-serif">Tracking Venta Presente &middot; By Sebasti&#225;n G&#243;mez L&#243;pez</p>
            </td>
            <td align="right" style="padding:22px 24px 14px;vertical-align:top">
              <table role="presentation" cellpadding="0" cellspacing="0" border="0"><tr>
                <td bgcolor="#1A1816" style="background-color:#1A1816;border:1px solid #3A3835;padding:10px 14px;text-align:center">
                  <p style="margin:0 0 3px;font-size:9px;color:#5A5856;text-transform:uppercase;letter-spacing:1px;font-family:Arial,Helvetica,sans-serif">Generado</p>
                  <p style="margin:0;font-size:12px;font-weight:bold;color:{C_VERDE_MENTA};font-family:Arial,Helvetica,sans-serif">{fecha}</p>
                </td>
              </tr></table>
            </td>
          </tr>
          <tr>
            <td colspan="2" bgcolor="{C_VERDE_MENTA}" style="background-color:{C_VERDE_MENTA};padding:8px 24px">
              <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%"><tr>
                <td style="font-size:11px;font-weight:bold;color:{C_NEGRO_CIB};text-transform:uppercase;letter-spacing:0.5px;font-family:Arial,Helvetica,sans-serif">Informe Automatizado ANS</td>
                <td align="right" style="font-size:11px;font-weight:bold;color:{C_VERDE_SELVA};font-family:Arial,Helvetica,sans-serif">{kpis["total"]} registros activos</td>
              </tr></table>
            </td>
          </tr>
        </table>
      </td>
    </tr>
    <tr>
      <td colspan="3" bgcolor="#0F120F" style="background-color:#0F120F;padding:20px 20px 24px">
        <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="margin-bottom:16px"><tr>
          <td bgcolor="{alert_bg}" style="background-color:{alert_bg};border-left:4px solid {alert_color};padding:12px 16px">
            <p style="margin:0;font-size:13px;font-weight:bold;color:{alert_color};font-family:Arial,Helvetica,sans-serif">{alert_icon}&nbsp;&nbsp;{alert_msg}</p>
          </td>
        </tr></table>
        <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="margin-bottom:20px"><tr>
          <td bgcolor="{C_VERDE_SELVA}" align="center" style="background-color:{C_VERDE_SELVA};padding:16px 20px">
            <p style="margin:0 0 10px;font-size:11px;font-weight:bold;color:{C_NEGRO_CIB};font-family:Arial,Helvetica,sans-serif">Para filtros, hist&#243;rico y an&#225;lisis completo:</p>
            <table role="presentation" cellpadding="0" cellspacing="0" border="0"><tr>
              <td bgcolor="{C_NEGRO_CIB}" style="background-color:{C_NEGRO_CIB};padding:10px 26px">
                <a href="{dashboard_url}" style="color:{C_VERDE_MENTA};text-decoration:none;font-weight:bold;font-size:12px;font-family:Arial,Helvetica,sans-serif;letter-spacing:0.3px">&#128279;&nbsp; Abrir Dashboard Wompi VP &nbsp;&#8594;</a>
              </td>
            </tr></table>
          </td>
        </tr></table>
        <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="margin-bottom:10px;border-bottom:1px solid #1E251E"><tr>
          <td style="padding-bottom:8px"><p style="margin:0;font-size:10px;color:#4A7A4A;text-transform:uppercase;letter-spacing:1.5px;font-weight:bold;font-family:Arial,Helvetica,sans-serif">Alertas Operativas</p></td>
        </tr></table>
        {s_hoy}{s_vencidas}{s_bl24}{s_bl48}{s_stalled}{s_fallidos}
      </td>
    </tr>
    <tr>
      <td width="33%" height="2" style="background-color:{C_VERDE_LIMA};font-size:1px;line-height:2px">&nbsp;</td>
      <td width="34%" height="2" style="background-color:{C_VERDE_MENTA};font-size:1px;line-height:2px">&nbsp;</td>
      <td width="33%" height="2" style="background-color:{C_AZUL_CIELO};font-size:1px;line-height:2px">&nbsp;</td>
    </tr>
    <tr>
      <td colspan="3" bgcolor="{C_NEGRO_CIB}" style="background-color:{C_NEGRO_CIB}">
        <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%"><tr>
          <td style="padding:14px 22px;vertical-align:middle">
            <span style="font-size:14px;font-weight:bold;color:{C_VERDE_MENTA};font-family:Georgia,Times New Roman,serif">W Wompi</span>
            <span style="font-size:10px;color:#3C3A38;font-family:Arial,Helvetica,sans-serif"> &middot; powered by Lineacom</span>
          </td>
          <td align="right" style="padding:14px 22px;vertical-align:middle">
            <p style="margin:0;font-size:9px;color:#3C3A38;font-family:Arial,Helvetica,sans-serif">{fecha} &middot; generado autom&#225;ticamente</p>
            <p style="margin:2px 0 0;font-size:9px;font-weight:bold;color:{C_VERDE_LIMA};letter-spacing:1px;font-family:Arial,Helvetica,sans-serif">CONFIDENCIAL</p>
          </td>
        </tr></table>
      </td>
    </tr>
  </table>
</td></tr>
</table>
</body>
</html>"""


def build_plain_text(kpis: Dict[str, Any]) -> str:
    return "\n".join([
        f"REPORTE OPERATIVO WOMPI VP — {kpis['fecha_informe']}", "="*60, "",
        "ALERTAS OPERATIVAS",
        f"  ⏰ Vencen Hoy           : {kpis['vencen_hoy']}",
        f"  🚨 Vencidas ANS         : {kpis['vencidas']}",
        f"  ⚡ Backlog 24h          : {kpis['backlog_24h']}",
        f"  📅 Backlog 48h          : {kpis['backlog_48h']}",
        f"  🔒 Sin cambios >24h     : {kpis['guias_sin_cambios']}",
        f"  ❌ Intentos fallidos    : {kpis['intentos_fallidos']}", "",
        "Para ver el detalle completo ingrese a:", "https://segoloo.github.io/TablerosWompi/", "",
        "Equipo Analítica TI — Lineacom",
    ])


def enviar_correo(kpis: Dict[str, Any]):
    total_alert = kpis["vencen_hoy"] + kpis["vencidas"] + kpis["guias_sin_cambios"] + kpis["intentos_fallidos"]
    prefix = "🚨" if total_alert>5 else "⚠️" if total_alert>0 else "✅"
    msg = MIMEMultipart("mixed")
    msg["From"]    = formataddr(("Analítica TI · Lineacom", EMAIL_USER))
    msg["To"]      = ", ".join(TO_EMAILS)
    msg["Subject"] = f"{prefix} Reporte ANS Wompi VP — {kpis['fecha_informe']}"
    if CC_EMAILS: msg["Cc"] = ", ".join(CC_EMAILS)
    alt = MIMEMultipart("alternative")
    alt.attach(MIMEText(build_plain_text(kpis), "plain", "utf-8"))
    alt.attach(MIMEText(build_html_email(kpis), "html",  "utf-8"))
    msg.attach(alt)
    server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
    server.starttls(); server.login(EMAIL_USER, EMAIL_PASSWORD)
    server.sendmail(EMAIL_USER, list(dict.fromkeys(TO_EMAILS + CC_EMAILS)), msg.as_string())
    server.quit()
    log.info("✅ Correo enviado exitosamente.")


def run_vp(no_push: bool = False, no_mail: bool = False):
    start = datetime.now()
    log.info("═"*60 + f"\n  SYNC WOMPI VP — {start.strftime('%Y-%m-%d %H:%M:%S')}\n" + "═"*60)
    _token_cache.get()
    item_id = find_workbook_item_id()
    df      = read_sheet(item_id)
    log.info(f"  Columnas detectadas: {list(df.columns)[:8]} ...")
    payload = build_vp_json(df)
    content = json.dumps(payload, ensure_ascii=False, separators=(",", ":"))
    OUTPUT_VP.write_text(content, encoding="utf-8")
    log.info(f"data.json escrito: {payload['filas']} filas, {len(content.encode())//1024:.0f}KB")
    if not no_push:
        push_github(content.encode("utf-8"), "data.json", f"sync: tracking VP {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    else:
        log.info("--no-push activo: se omite subida a GitHub.")
    if not no_mail:
        log.info("Calculando KPIs para correo...")
        kpis = compute_kpis_for_email(clean_df(df))
        for k, label in [("vencen_hoy","⏰ Vencen Hoy"),("vencidas","🚨 Vencidas ANS"),("backlog_24h","⚡ Backlog 24h"),("backlog_48h","📅 Backlog 48h"),("guias_sin_cambios","🔒 Sin cambios"),("intentos_fallidos","❌ Intentos fallidos")]:
            log.info(f"   {label}: {kpis[k]}")
        enviar_correo(kpis)
        log.info(f"✅ Correo enviado a: {', '.join(TO_EMAILS)}")
    else:
        log.info("--no-mail activo: se omite envío de correo.")
    log.info(f"✅ VP completado en {(datetime.now()-start).total_seconds():.1f}s — {payload['filas']} registros")


# ══════════════════════════════════════════════════════════════════
#  MÓDULO C — STOCK WOMPI FILTRADO (Sytex → stock_wompi_filtrado.json.gz)
#
#  Origen  : API Sytex (datareport export)
#  Filtros : códigos de material, seriales excluidos, posiciones,
#            atributos, fusión de bodegas por ciudad
#  Enriq.  : fecha de confirmación + tipo_de_operacion (reuso) desde MySQL
#  Destino : stock_wompi_filtrado.json.gz → GitHub segoloo/TablerosWompi
# ══════════════════════════════════════════════════════════════════

# ── Credenciales Sytex ────────────────────────────────────────────
SYTEX_USER = "monitoreo.ti@lineacom.co"
SYTEX_PASS = "G0cjiGisqcGZsBcYiBOFa5jZ"

STOCK_ENDPOINT = "https://app.sytex.io/api/datareport/01212d5f-3d96-4097-89fc-515bb0617694/export/?org_id=115"

OPERACION_MATERIAL_TABLE       = "operacion_material_items_areas_dw"
OPERACION_MATERIAL_AREAS_TABLE = "operacion_material_areas_dw"

OUTPUT_STOCK = Path("stock_wompi_filtrado.json.gz")

CODIGOS_FILTRO = {
    "MA-1554", "MA-1614", "MA-1616", "MA-1553", "MA-1615",
    "MA-1881", "MA-1997", "MA-2003", "MA-2004", "MA-2005",
    "MA-2006", "MA-2007", "MA-1617", "MA-2300", "MA-2301",
    "MA-2302", "MA-2303", "MA-2309", "MA-2310", "MA-2311",
    "MA-2312", "MA-2313", "MA-2314", "MA-2315", "MA-2316",
    "MA-2371",
}

SERIALES_EXCLUIR = {
    "KU403616","KU403620","KU404444","KU404484","KU403655",
    "KU403865","KU403905","KU403910","KU403928","KU405030","KU405594",
    "KU403873","KU403938","KU407472",
    "KU404451","KU403661","KU403882","KU403923","KU403927","KU403941",
    "KU403807","KU405982","KU408110","KU408934",
    "KU403606","KU403866",
    "12341234","234234234234",
    "SIMP00001","SIMP00002","SIMP00003",
    "SIMVP00001","SIMVP00002","SIMVP00003","SIMVP00004",
}

POSICIONES_EXCLUIR = {
    "BAJA POR SUSTITUCION 21/11/2025",
    "PRUEBAS API VP",
}

ATRIBUTOS_EXCLUIR = {
    "WOMPI PRUEBA",
    "WOMPI PRUEBA, WOMPI",
    "WOMPI PRUEBA, WOMPI, WOMPI VP",
}

BODEGA_MERGE_MAP = {
    "ALMACEN WOMPI MEDELLIN":                "ALMACEN WOMPI MEDELLIN",
    "ALMACEN WOMPI ALISTAMIENTO MEDELLIN":   "ALMACEN WOMPI MEDELLIN",
    "ALMACEN WOMPI VP MEDELLIN | VENTA":     "ALMACEN WOMPI MEDELLIN",
    "ALMACEN WOMPI VP MEDELLIN | ALQUILER":  "ALMACEN WOMPI MEDELLIN",
    "ALMACEN WOMPI VP ALISTAMIENTO MEDELLIN":"ALMACEN WOMPI MEDELLIN",
    "ALMACEN WOMPI BOGOTA":                  "ALMACEN WOMPI BOGOTA",
    "ALMACEN WOMPI VP BOGOTA | VENTA":       "ALMACEN WOMPI BOGOTA",
    "ALMACEN WOMPI VP BOGOTA | ALQUILER":    "ALMACEN WOMPI BOGOTA",
    "ALMACEN WOMPI BUCARAMANGA":             "ALMACEN WOMPI BUCARAMANGA",
    "ALMACEN WOMPI VP BUCARAMANGA | VENTA":  "ALMACEN WOMPI BUCARAMANGA",
    "ALMACEN WOMPI VP BUCARAMANGA | ALQUILER":"ALMACEN WOMPI BUCARAMANGA",
    "ALMACEN WOMPI CALI":                    "ALMACEN WOMPI CALI",
    "ALMACEN WOMPI VP CALI | VENTA":         "ALMACEN WOMPI CALI",
    "ALMACEN WOMPI VP CALI | ALQUILER":      "ALMACEN WOMPI CALI",
    "ALMACEN WOMPI PEREIRA":                 "ALMACEN WOMPI PEREIRA",
    "ALMACEN WOMPI VP PEREIRA | VENTA":      "ALMACEN WOMPI PEREIRA",
    "ALMACEN WOMPI VP PEREIRA | ALQUILER":   "ALMACEN WOMPI PEREIRA",
    "ALMACEN WOMPI BARRANQUILLA":            "ALMACEN WOMPI BARRANQUILLA",
    "ALMACEN WOMPI VP BARRANQUILLA | VENTA": "ALMACEN WOMPI BARRANQUILLA",
    "ALMACEN WOMPI VP BARRANQUILLA | ALQUILER":"ALMACEN WOMPI BARRANQUILLA",
    "ALMACEN WOMPI MONTERIA":                "ALMACEN WOMPI MONTERIA",
    "ALMACEN WOMPI VP MONTERIA | VENTA":     "ALMACEN WOMPI MONTERIA",
    "ALMACEN WOMPI VP MONTERIA | ALQUILER":  "ALMACEN WOMPI MONTERIA",
    "ALMACEN WOMPI VILLAVICENCIO":           "ALMACEN WOMPI VILLAVICENCIO",
    "ALMACEN WOMPI CUCUTA":                  "ALMACEN WOMPI CUCUTA",
    "ALMACEN WOMPI NEIVA":                   "ALMACEN WOMPI NEIVA",
    "ALMACEN WOMPI IBAGUE":                  "ALMACEN WOMPI IBAGUE",
    "ALMACEN WOMPI TUNJA":                   "ALMACEN WOMPI TUNJA",
    "ALMACEN WOMPI SANTA MARTA":             "ALMACEN WOMPI SANTA MARTA",
    "ALMACEN WOMPI VALLEDUPAR":              "ALMACEN WOMPI VALLEDUPAR",
    "ALMACEN WOMPI CARTAGENA":               "ALMACEN WOMPI CARTAGENA",
    "ALMACEN WOMPI FLORENCIA":               "ALMACEN WOMPI FLORENCIA",
    "ALMACEN WOMPI POPAYAN":                 "ALMACEN WOMPI POPAYAN",
    "ALMACEN WOMPI MANIZALES":               "ALMACEN WOMPI MANIZALES",
    "ALMACEN WOMPI YOPAL":                   "ALMACEN WOMPI YOPAL",
    "ALMACEN WOMPI APARTADO":                "ALMACEN WOMPI APARTADO",
    "ALMACEN WOMPI PASTO":                   "ALMACEN WOMPI PASTO",
    "ALMACEN WOMPI SINCELEJO":               "ALMACEN WOMPI SINCELEJO",
    "ALMACEN WOMPI ARMENIA":                 "ALMACEN WOMPI ARMENIA",
    "ALMACEN BAJAS WOMPI":                   "ALMACEN BAJAS WOMPI",
    "ALMACEN INGENICO - PROVEEDOR WOMPI":    "ALMACEN INGENICO - PROVEEDOR WOMPI",
}

COLUMNAS_SALIDA_STOCK = [
    "Código", "Nombre", "Tipo", "Cantidad", "Código de ubicación",
    "Códigos de area de trabajo", "Descripción", "Email del creador",
    "Email del último editor", "Fecha de creación", "Fecha de la última edición",
    "ID", "Nombre del cliente", "Número de serie", "Posición en depósito",
    "Subtipo", "Tipo de ubicación", "Unidad de medida", "Nombre de la ubicación",
    "Reusable", "Atributos", "Comentarios", "tipo_de_operacion",
]


def _norm_col(s: str) -> str:
    return (s.strip().lower()
            .replace("ó","o").replace("é","e").replace("á","a")
            .replace("í","i").replace("ú","u"))


def _fmt_fecha_stock(serie: "pd.Series") -> "pd.Series":
    fechas = pd.to_datetime(serie, errors="coerce")
    mask_tz = fechas.notna()
    fechas[mask_tz] = fechas[mask_tz] - pd.Timedelta(hours=5)
    return fechas.dt.strftime("%Y-%m-%dT%H:%M:%S")


def _download_stock() -> bytes:
    s = requests.Session()
    s.auth = (SYTEX_USER, SYTEX_PASS)
    s.headers.update({"User-Agent": "lineacom-collector/1.0", "Accept": "*/*"})
    log.info("Descargando stock_wompi desde Sytex...")
    with s.get(STOCK_ENDPOINT, timeout=(30, 600), stream=True) as r:
        r.raise_for_status()
        chunks, total = [], 0
        for chunk in r.iter_content(chunk_size=512 * 1024):
            chunks.append(chunk); total += len(chunk)
        raw = b"".join(chunks)
    log.info(f"  ✓ Descargado: {len(raw)/1024/1024:.1f} MB")
    return raw


def _raw_to_df(raw: bytes) -> "pd.DataFrame":
    import io
    try:
        if raw.strip().startswith((b"{", b"[")):
            d = json.loads(raw.decode("utf-8"))
            if isinstance(d, dict) and isinstance(d.get("data"), list):
                return pd.DataFrame(d["data"])
            return pd.DataFrame(d)
    except Exception:
        pass
    try:
        return pd.read_csv(io.BytesIO(raw), encoding="utf-8", low_memory=False)
    except Exception:
        return pd.read_csv(io.BytesIO(raw), encoding="latin-1", sep=None, engine="python", low_memory=False)


def _enrich_tipo_operacion_stock(df: "pd.DataFrame", conn) -> "pd.DataFrame":
    df = df.copy()
    if "tipo_de_operacion" not in df.columns:
        df["tipo_de_operacion"] = None

    insp = sa_inspect(conn)
    table_names = insp.get_table_names(schema=DB_NAME)
    if OPERACION_MATERIAL_TABLE not in table_names:
        log.warning(f"Tabla {OPERACION_MATERIAL_TABLE} no existe. Se omite tipo_de_operacion.")
        return df

    operacion_cols = {col["name"] for col in insp.get_columns(OPERACION_MATERIAL_TABLE, schema=DB_NAME)}
    serie_col = next((c for c in ["Número_de_serie","numero_de_serie","numero_serie","serie","serial"] if c in operacion_cols), None)
    tipo_col  = next((c for c in ["Tipo_de_operación","tipo_de_operacion","tipo_operacion","Tipo_de_movimiento","tipo_de_movimiento"] if c in operacion_cols), None)

    if not serie_col or not tipo_col:
        log.warning(f"Columnas de serie/tipo no encontradas en {OPERACION_MATERIAL_TABLE}. Se omite tipo_de_operacion.")
        return df

    col_serie_df = next((c for c in df.columns if _norm_col(c) == "numero de serie"), None)
    if not col_serie_df:
        log.warning("No se encontró columna 'Número de serie' en el DataFrame. Se omite tipo_de_operacion.")
        return df

    df["numero_de_serie"] = df[col_serie_df].astype(str).str.strip().replace("nan", None)
    numeros_validos = [s for s in df["numero_de_serie"].dropna().unique() if s]
    if not numeros_validos:
        df = df.drop(columns=["numero_de_serie"], errors="ignore")
        return df

    ns_str = ", ".join(f"'{s}'" for s in numeros_validos)
    query = text(f"""
        SELECT `{serie_col}` AS numero_serie_operacion,
               MAX(CASE WHEN LOWER(TRIM(`{tipo_col}`)) = 'return' THEN 1 ELSE 0 END) AS es_reusado
        FROM `{OPERACION_MATERIAL_TABLE}`
        WHERE `{serie_col}` IN ({ns_str})
          AND `{serie_col}` IS NOT NULL AND `{serie_col}` <> ''
          AND `{tipo_col}` IS NOT NULL AND `{tipo_col}` <> ''
        GROUP BY `{serie_col}`
    """)
    reuso_df = pd.read_sql(query, conn)
    if reuso_df.empty:
        df = df.drop(columns=["numero_de_serie"], errors="ignore")
        return df

    df = df.merge(reuso_df, left_on="numero_de_serie", right_on="numero_serie_operacion", how="left")
    df["es_reusado"] = df["es_reusado"].fillna(0).astype(int)
    df["tipo_de_operacion"] = df["tipo_de_operacion"].astype(object)
    df.loc[df["es_reusado"] == 1, "tipo_de_operacion"] = "return"
    df.loc[df["es_reusado"] == 0, "tipo_de_operacion"] = None
    cant = (df["es_reusado"] == 1).sum()
    log.info(f"  ✓ Marcados {cant} equipos como REUSADOS (tipo_de_operacion='return')")
    df = df.drop(columns=["numero_serie_operacion","es_reusado","numero_de_serie"], errors="ignore")
    return df


def _enrich_fecha_confirmacion_stock(df: "pd.DataFrame") -> "pd.DataFrame":
    log.info("Enriqueciendo 'Fecha de la última edición' desde MySQL...")
    df = df.copy()
    df["Fecha de la última edición"] = None

    try:
        engine = create_engine(
            f"mysql+pymysql://{DB_USER}:{DB_PASSWORD}@{DB_HOST}:{DB_PORT}/{DB_NAME}",
            connect_args={"connect_timeout": 30},
        )
        conn = engine.connect()
    except Exception as e:
        log.warning(f"No se pudo conectar a MySQL para fechas: {e}")
        return df

    try:
        insp = sa_inspect(engine)
        table_names = insp.get_table_names(schema=DB_NAME)
        for tbl in [OPERACION_MATERIAL_TABLE, OPERACION_MATERIAL_AREAS_TABLE]:
            if tbl not in table_names:
                log.warning(f"Tabla '{tbl}' no encontrada. Se omite enriquecimiento de fechas.")
                return df

        items_cols = {col["name"] for col in insp.get_columns(OPERACION_MATERIAL_TABLE, schema=DB_NAME)}
        areas_cols = {col["name"] for col in insp.get_columns(OPERACION_MATERIAL_AREAS_TABLE, schema=DB_NAME)}

        serie_col           = next((c for c in ["Número_de_serie","numero_de_serie","numero_serie","serial"] if c in items_cols), None)
        codigo_material_col = next((c for c in ["Código_de_material","codigo_de_material"] if c in items_cols), None)
        pos_deposito_col    = next((c for c in ["Posición_en_depósito_de_origen","Posicion_en_deposito_de_origen"] if c in items_cols), None)
        cod_op_col          = next((c for c in ["Código_de_operación_de_materiales","Codigo_de_operacion_de_materiales"] if c in items_cols), None)
        cod_ubic_destino_col= next((c for c in ["Código_de_ubicación_de_destino","Codigo_de_ubicacion_de_destino"] if c in items_cols), None)
        fecha_col_areas     = next((c for c in ["Fecha_de_confirmación","Fecha_de_confirmacion","fecha_de_confirmacion"] if c in areas_cols), None)
        cod_op_areas_col    = next((c for c in ["Código_de_operación_de_materiales","Codigo_de_operacion_de_materiales","Código_de_operación","Codigo_de_operacion"] if c in areas_cols), None)

        col_serie_df  = next((c for c in df.columns if _norm_col(c) == "numero de serie"), None)
        col_codigo_df = next((c for c in df.columns if _norm_col(c) == "codigo"), None)
        col_pos_df    = next((c for c in df.columns if _norm_col(c) == "posicion en deposito"), None)
        col_ubic_df   = next((c for c in df.columns if _norm_col(c) == "codigo de ubicacion"), None)

        # ── Estrategia 1: materiales CON serie ───────────────────────
        if serie_col and col_serie_df and fecha_col_areas and cod_op_col and cod_op_areas_col:
            mask_con = df[col_serie_df].notna() & (df[col_serie_df].astype(str).str.strip() != "")
            df_con = df[mask_con]
            if not df_con.empty:
                numeros = [str(s).strip() for s in df_con[col_serie_df].dropna().unique() if str(s).strip().lower() != "nan"]
                if numeros:
                    ph = ", ".join(f"'{s.replace(chr(39), chr(39)+chr(39))}'" for s in numeros)
                    q = text(f"""
                        WITH ranked AS (
                            SELECT i.`{serie_col}` AS serial_op,
                                   a.`{fecha_col_areas}` AS fecha_raw,
                                   ROW_NUMBER() OVER (PARTITION BY i.`{serie_col}` ORDER BY a.`{fecha_col_areas}` DESC) AS rn
                            FROM `{OPERACION_MATERIAL_TABLE}` i
                            JOIN `{OPERACION_MATERIAL_AREAS_TABLE}` a ON i.`{cod_op_col}` = a.`{cod_op_areas_col}`
                            WHERE i.`{serie_col}` IN ({ph}) AND a.`{fecha_col_areas}` IS NOT NULL
                        )
                        SELECT serial_op, fecha_raw FROM ranked WHERE rn = 1
                    """)
                    op_df = pd.read_sql(q, conn)
                    if not op_df.empty:
                        op_df["fecha_calc"] = _fmt_fecha_stock(op_df["fecha_raw"])
                        mapa = op_df.set_index("serial_op")["fecha_calc"].to_dict()
                        idx = df.index[mask_con]
                        mapped = df.loc[idx, col_serie_df].astype(str).str.strip().map(mapa)
                        df.loc[idx[mapped.notna()], "Fecha de la última edición"] = mapped[mapped.notna()].values
                        log.info(f"  ✓ Fechas por serie aplicadas: {mapped.notna().sum():,}/{len(df_con):,}")

        # ── Estrategia 2: materiales SIN serie ───────────────────────
        if col_serie_df:
            mask_sin = df[col_serie_df].isna() | (df[col_serie_df].astype(str).str.strip() == "")
        else:
            mask_sin = pd.Series([True]*len(df), index=df.index)

        df_sin = df[mask_sin]
        if (not df_sin.empty and codigo_material_col and col_codigo_df
                and fecha_col_areas and cod_op_col and cod_op_areas_col):
            codigos = [str(c).strip() for c in df_sin[col_codigo_df].dropna().unique() if str(c).strip()]
            if codigos:
                ph2 = ", ".join(f"'{c.replace(chr(39), chr(39)+chr(39))}'" for c in codigos)
                sel = [f"i.`{codigo_material_col}` AS codigo_material"]
                grp = [f"i.`{codigo_material_col}`"]
                if cod_ubic_destino_col:
                    sel.append(f"i.`{cod_ubic_destino_col}` AS cod_ubic_destino"); grp.append(f"i.`{cod_ubic_destino_col}`")
                if pos_deposito_col:
                    sel.append(f"i.`{pos_deposito_col}` AS pos_deposito"); grp.append(f"i.`{pos_deposito_col}`")
                q2 = text(f"""
                    SELECT {', '.join(sel)}, MAX(a.`{fecha_col_areas}`) AS fecha_raw
                    FROM `{OPERACION_MATERIAL_TABLE}` i
                    JOIN `{OPERACION_MATERIAL_AREAS_TABLE}` a ON i.`{cod_op_col}` = a.`{cod_op_areas_col}`
                    WHERE i.`{codigo_material_col}` IN ({ph2})
                      AND (i.`{serie_col}` IS NULL OR TRIM(i.`{serie_col}`) = '')
                      AND a.`{fecha_col_areas}` IS NOT NULL
                    GROUP BY {', '.join(grp)}
                """)
                op_gen = pd.read_sql(q2, conn)
                if not op_gen.empty:
                    op_gen["fecha_calc"] = _fmt_fecha_stock(op_gen["fecha_raw"])
                    lk = [col_codigo_df]; rk = ["codigo_material"]; mc = ["codigo_material","fecha_calc"]
                    if cod_ubic_destino_col and col_ubic_df and "cod_ubic_destino" in op_gen.columns:
                        lk.insert(1, col_ubic_df); rk.insert(1, "cod_ubic_destino"); mc.append("cod_ubic_destino")
                    if pos_deposito_col and col_pos_df and "pos_deposito" in op_gen.columns:
                        lk.append(col_pos_df); rk.append("pos_deposito"); mc.append("pos_deposito")
                    tmp = df.loc[mask_sin, list(dict.fromkeys(lk))].copy(); tmp["__idx"] = tmp.index
                    merged = tmp.merge(op_gen[list(dict.fromkeys(mc))], left_on=lk, right_on=rk, how="left")
                    # fallbacks
                    sin = merged["fecha_calc"].isna()
                    if sin.any() and len(lk) > 2:
                        fb = op_gen.groupby(rk[:2], dropna=False)["fecha_calc"].max().reset_index().rename(columns={"fecha_calc":"fecha_calc_fb"})
                        fb2 = tmp[sin.values].merge(fb, left_on=lk[:2], right_on=rk[:2], how="left")
                        merged.loc[sin, "fecha_calc"] = fb2["fecha_calc_fb"].values
                        sin = merged["fecha_calc"].isna()
                    if sin.any():
                        fb3 = op_gen.groupby("codigo_material", dropna=False)["fecha_calc"].max().reset_index().rename(columns={"fecha_calc":"fecha_calc_fb2"})
                        fb4 = tmp[sin.values].merge(fb3, left_on=col_codigo_df, right_on="codigo_material", how="left")
                        merged.loc[sin, "fecha_calc"] = fb4["fecha_calc_fb2"].values
                    ok = merged["fecha_calc"].notna()
                    df.loc[merged.loc[ok,"__idx"].values, "Fecha de la última edición"] = merged.loc[ok,"fecha_calc"].values
                    log.info(f"  ✓ Fechas sin serie aplicadas: {ok.sum():,}/{len(df_sin):,}")

        total_fechas = df["Fecha de la última edición"].notna().sum()
        log.info(f"  ✓ Total fechas enriquecidas: {total_fechas:,}/{len(df):,} ({total_fechas/len(df)*100:.1f}%)")

    except Exception as e:
        log.warning(f"Error enriqueciendo fechas: {e}")
        import traceback; traceback.print_exc()
    finally:
        conn.close(); engine.dispose()

    return df


def run_stock(no_push: bool = False):
    start = datetime.now()
    log.info("═"*60 + f"\n  SYNC STOCK WOMPI FILTRADO — {start.strftime('%Y-%m-%d %H:%M:%S')}\n" + "═"*60)

    # ── Descarga ──────────────────────────────────────────────────
    raw      = _download_stock()
    df_stock = _raw_to_df(raw)
    log.info(f"  Filas descargadas: {len(df_stock):,} × {len(df_stock.columns)} columnas")

    # ── Filtro por código de material ─────────────────────────────
    col_codigo = next((c for c in df_stock.columns if _norm_col(c) == "codigo"), None)
    if col_codigo is None:
        raise RuntimeError(f"No se encontró columna 'Código'. Columnas disponibles: {df_stock.columns.tolist()}")
    df_f = df_stock[df_stock[col_codigo].astype(str).str.strip().isin(CODIGOS_FILTRO)].copy()
    log.info(f"  Filtro códigos: {len(df_stock):,} → {len(df_f):,} filas")

    # ── Filtro por número de serie ────────────────────────────────
    col_serie = next((c for c in df_f.columns if _norm_col(c) == "numero de serie"), None)
    if col_serie:
        antes = len(df_f)
        df_f = df_f[~df_f[col_serie].astype(str).str.strip().str.upper().isin({s.upper() for s in SERIALES_EXCLUIR})].copy()
        log.info(f"  Seriales excluidos: {antes - len(df_f):,} filas")

    # ── Filtro por posición en depósito ───────────────────────────
    col_pos = next((c for c in df_f.columns if _norm_col(c) == "posicion en deposito"), None)
    if col_pos:
        antes = len(df_f)
        df_f = df_f[~df_f[col_pos].astype(str).str.strip().isin(POSICIONES_EXCLUIR)].copy()
        log.info(f"  Posiciones excluidas: {antes - len(df_f):,} filas")

    # ── Filtro por atributos ──────────────────────────────────────
    col_attr = next((c for c in df_f.columns if _norm_col(c) == "atributos"), None)
    if col_attr:
        antes = len(df_f)
        df_f = df_f[~df_f[col_attr].astype(str).str.strip().isin(ATRIBUTOS_EXCLUIR)].copy()
        log.info(f"  Atributos excluidos: {antes - len(df_f):,} filas")

    # ── Fusión de bodegas ─────────────────────────────────────────
    col_ubic_nombre = next((c for c in df_f.columns if _norm_col(c) == "nombre de la ubicacion"), None)
    if col_ubic_nombre:
        df_f[col_ubic_nombre] = df_f[col_ubic_nombre].astype(str).str.strip().map(lambda v: BODEGA_MERGE_MAP.get(v, v))
        log.info("  Bodegas fusionadas por ciudad")

    # ── Seleccionar columnas de salida ────────────────────────────
    cols_presentes = [c for c in COLUMNAS_SALIDA_STOCK if c in df_f.columns]
    cols_faltantes = [c for c in COLUMNAS_SALIDA_STOCK if c not in df_f.columns]
    if cols_faltantes:
        log.warning(f"  Columnas no encontradas (se omiten): {cols_faltantes}")
    df_salida = df_f[cols_presentes].copy()

    # ── Enriquecimiento: fechas de confirmación ───────────────────
    df_salida = _enrich_fecha_confirmacion_stock(df_salida)

    # ── Enriquecimiento: tipo_de_operacion (reuso) ────────────────
    try:
        engine_r = create_engine(
            f"mysql+pymysql://{DB_USER}:{DB_PASSWORD}@{DB_HOST}:{DB_PORT}/{DB_NAME}",
            connect_args={"connect_timeout": 30},
        )
        with engine_r.connect() as conn_r:
            df_salida = _enrich_tipo_operacion_stock(df_salida, conn_r)
        engine_r.dispose()
    except Exception as e:
        log.warning(f"No se pudo conectar a MySQL para tipo_de_operacion: {e}")

    # ── Reordenar columnas finales ────────────────────────────────
    cols_final = [c for c in COLUMNAS_SALIDA_STOCK if c in df_salida.columns]
    df_salida = df_salida[cols_final]

    # ── Serializar a JSON.GZ ──────────────────────────────────────
    json_bytes = df_salida.to_json(orient="records", force_ascii=False, date_format="iso").encode("utf-8")
    gz_bytes   = gzip.compress(json_bytes, compresslevel=9)
    OUTPUT_STOCK.write_bytes(gz_bytes)
    log.info(f"stock_wompi_filtrado.json.gz → {len(df_salida):,} filas | "
             f"{len(json_bytes)//1024}KB JSON → {len(gz_bytes)//1024}KB gz")

    # ── Subir a GitHub ────────────────────────────────────────────
    if not no_push:
        push_github(gz_bytes, "stock_wompi_filtrado.json.gz",
                    f"sync: stock wompi filtrado {datetime.now().strftime('%Y-%m-%d %H:%M')}", is_binary=True)
    else:
        log.info("--no-push activo: se omite subida a GitHub.")

    log.info(f"✅ Stock completado en {(datetime.now()-start).total_seconds():.1f}s — {len(df_salida):,} registros")


# ══════════════════════════════════════════════════════════════════
#  MAIN
# ══════════════════════════════════════════════════════════════════
def run_once(no_push=False, no_mail=False, only_rollos=False, only_vp=False,
             only_stock=False, no_stock=False):
    if only_rollos:
        run_rollos(no_push=no_push)
    elif only_vp:
        run_vp(no_push=no_push, no_mail=no_mail)
    elif only_stock:
        run_stock(no_push=no_push)
    else:
        run_rollos(no_push=no_push)
        run_vp(no_push=no_push, no_mail=no_mail)
        if not no_stock:
            run_stock(no_push=no_push)


def main():
    no_push      = "--no-push"      in sys.argv
    no_mail      = "--no-mail"      in sys.argv
    loop         = "--loop"         in sys.argv
    only_rollos  = "--only-rollos"  in sys.argv
    only_vp      = "--only-vp"      in sys.argv
    only_stock   = "--only-stock"   in sys.argv
    no_stock     = "--no-stock"     in sys.argv

    if loop:
        log.info(f"Modo loop cada {INTERVAL_HOURS}h. Ctrl+C para detener.")
        while True:
            try:
                run_once(no_push=no_push, no_mail=no_mail, only_rollos=only_rollos,
                         only_vp=only_vp, only_stock=only_stock, no_stock=no_stock)
            except Exception as e:
                log.error(f"Ciclo fallido: {e}"); import traceback; traceback.print_exc()
            time.sleep(INTERVAL_HOURS * 3600)
    else:
        try:
            run_once(no_push=no_push, no_mail=no_mail, only_rollos=only_rollos,
                     only_vp=only_vp, only_stock=only_stock, no_stock=no_stock)
        except Exception as e:
            log.error(f"❌ Error: {e}"); import traceback; traceback.print_exc(); sys.exit(1)


if __name__ == "__main__":
    main()