#!/usr/bin/env python3
"""
╔══════════════════════════════════════════════════════════════════╗
║   data.py — Extractor SharePoint → data.json → GitHub           ║
║             + Correo operativo Wompi VP                         ║
║   LINEACOM · Dashboard Tracking VP Wompi                        ║
║                                                                  ║
║   Responsabilidad:                                               ║
║     1. Extrae datos crudos del Excel en SharePoint               ║
║     2. Serializa data.json (sin KPIs pre-calculados)             ║
║     3. Sube data.json a GitHub Pages                             ║
║     4. Calcula KPIs y envía correo operativo                     ║
║                                                                  ║
║   NOTA: dashboard.js recalcula todos los KPIs desde las filas    ║
║         crudas del data.json — no depende de cálculos de aquí.   ║
║                                                                  ║
║   Uso:                                                           ║
║     python data.py              # ejecutar una vez               ║
║     python data.py --loop       # cada N horas                   ║
║     python data.py --no-push    # genera JSON sin subir          ║
║     python data.py --no-mail    # sin envío de correo            ║
╚══════════════════════════════════════════════════════════════════╝
"""

import os
import sys
import json
import re
import time
import base64
import hashlib
import logging
import smtplib
from datetime import datetime, timedelta
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import formataddr
from pathlib import Path
from typing import Any, Dict, List, Optional

import requests
import pandas as pd

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s"
)
log = logging.getLogger("data_wompi_vp")


# ══════════════════════════════════════════════════════════════════
#  CONFIGURACIÓN — edita o usa variables de entorno
# ══════════════════════════════════════════════════════════════════

# ── Graph API / SharePoint ────────────────────────────────────────
CLIENT_ID     = os.getenv("SP_CLIENT_ID",     "637ffa8d-43fc-485c-ba3d-e3ee2da6d6d6")
CLIENT_SECRET = os.getenv("SP_CLIENT_SECRET", "F4.8Q~P.DqObpdzw8wZmjNqCXVfjkEftNicoraIX")
TENANT_ID     = os.getenv("SP_TENANT_ID",     "af1a17b2-5d34-4f58-8b6c-6b94c6cd87ea")
SITE_ID       = os.getenv("SP_SITE_ID",       "6bc4f4bb-4479-45c0-a26b-3030893bc1c1")
SHEET_NAME    = os.getenv("SP_SHEET",         "Hoja1")

OAUTH_TOKEN_URL = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
GRAPH_BASE      = "https://graph.microsoft.com/v1.0"

# ── GitHub ────────────────────────────────────────────────────────
GITHUB_TOKEN     = os.getenv("GITHUB_TOKEN",     "")
GITHUB_REPO      = os.getenv("GITHUB_REPO",      "segoloo/TablerosWompi")
GITHUB_BRANCH    = os.getenv("GITHUB_BRANCH",    "main")
GITHUB_FILE_PATH = os.getenv("GITHUB_FILE_PATH", "data.json")

# ── Correo ────────────────────────────────────────────────────────
SMTP_SERVER    = "smtp-mail.outlook.com"
SMTP_PORT      = 587
EMAIL_USER     = "analitica@lineacom.co"
EMAIL_PASSWORD = "Linea.2024*"
TO_EMAILS      = ["", "sebastian.gomez@lineacom.co"]
CC_EMAILS: List[str] = []

# ── Loop ──────────────────────────────────────────────────────────
INTERVAL_HOURS = int(os.getenv("INTERVAL_HOURS", "4"))

# ── Output local ─────────────────────────────────────────────────
OUTPUT_JSON = Path("data.json")


# ══════════════════════════════════════════════════════════════════
#  PALETA WOMPI 2025
# ══════════════════════════════════════════════════════════════════
C_NEGRO_CIB   = "#2C2A29"
C_BLANCO      = "#FAFAFA"
C_VERDE_MENTA = "#B0F2AE"
C_AZUL_CIELO  = "#99D1FC"
C_VERDE_SELVA = "#00825A"
C_VERDE_LIMA  = "#DFFF61"
C_CARD_DARK   = "#1E1C1B"
C_MUTED       = "#6B6967"
C_DANGER      = "#FF5C5C"
C_WARNING     = "#FFC04D"


# ══════════════════════════════════════════════════════════════════
#  CONSTANTES VT / OPLG  (idénticas a dashboard.js)
# ══════════════════════════════════════════════════════════════════
VT_EXACT   = "VISITA DATAFONO+KIT POP+CAPACITACION"
OPLG_EXACT = "ENVIO DATAFONO+KIT POP"

# Columnas que contienen novedades (alineado con dashboard.js)
NOV_EXACT_COLS = [
    "NOVEDADES", "NOVEDAD", "CAUSAL INCU", "CAUSAL INC",
    "RESPONSABLE INCUMPLIMIENTO", "CAUSAL INCUMPLIMIENTO",
]
NOV_KEY_INCLUDES = ["novedad", "novedades", "causal", "responsable incump"]
NOV_KEY_EXCLUDES = [
    "estado", "fecha", "solicitud", "comercio", "guia", "transpo",
    "datafon", "serial", "departam", "ciudad", "tipolog", "cumple",
    "referencia", "tipo de sol", "id com",
]

# Columnas de fecha que se normalizan a dd/mm/yyyy
DATE_COLS = [
    "FECHA DE SOLICITUD",
    "FECHA LIMITE DE ENTREGA",
    "FECHA ENTREGA AL COMERCIO",
    "FECHA VISITA TECNICA",
    "FECHA DE ENTREGA",
]


# ══════════════════════════════════════════════════════════════════
#  TOKEN CACHE
# ══════════════════════════════════════════════════════════════════
class TokenCache:
    def __init__(self):
        self._token: Optional[str]  = None
        self._expiry: Optional[datetime] = None

    def get(self) -> str:
        now = datetime.now()
        if self._token and self._expiry and now < self._expiry:
            return self._token
        log.info("Obteniendo token Graph API...")
        data = {
            "client_id":     CLIENT_ID,
            "client_secret": CLIENT_SECRET,
            "scope":         "https://graph.microsoft.com/.default",
            "grant_type":    "client_credentials",
        }
        r = requests.post(OAUTH_TOKEN_URL, data=data, timeout=60)
        if r.status_code != 200:
            raise RuntimeError(f"Error token Graph API: {r.status_code} — {r.text[:200]}")
        j = r.json()
        self._token  = j["access_token"]
        self._expiry = now + timedelta(seconds=3500)
        log.info("Token obtenido ✅")
        return self._token


_token_cache = TokenCache()


# ══════════════════════════════════════════════════════════════════
#  SHAREPOINT — buscar y leer el libro
# ══════════════════════════════════════════════════════════════════
def find_workbook_item_id() -> str:
    """Busca el archivo Excel de Operación VP_Comodato en SharePoint."""
    token   = _token_cache.get()
    headers = {"Authorization": f"Bearer {token}"}

    for kw in ["Comodato", "VP", "Operaci"]:
        url = f"{GRAPH_BASE}/sites/{SITE_ID}/drive/root/search(q='{requests.utils.quote(kw, safe='')}' )"
        r   = requests.get(url, headers=headers, timeout=60)
        if r.status_code == 200:
            for item in r.json().get("value", []):
                name = item.get("name", "")
                if any(k.lower() in name.lower() for k in ["Comodato", "VP_Comodato"]):
                    log.info(f"Archivo encontrado: {name}")
                    return item["id"]

    r2 = requests.get(f"{GRAPH_BASE}/sites/{SITE_ID}/drive/root/children",
                      headers=headers, timeout=60)
    if r2.status_code == 200:
        for item in r2.json().get("value", []):
            if any(k.lower() in item.get("name", "").lower() for k in ["Comodato", "VP"]):
                log.info(f"Archivo encontrado (raíz): {item['name']}")
                return item["id"]

    raise RuntimeError(
        "No se encontró el archivo 'Operación VP_Comodato.xlsx' en SharePoint. "
        "Verifica SITE_ID y que el archivo exista en la raíz del drive."
    )


def read_sheet(item_id: str) -> pd.DataFrame:
    """Lee la hoja Excel desde SharePoint via Graph API."""
    token   = _token_cache.get()
    headers = {"Authorization": f"Bearer {token}"}
    sheet_enc = requests.utils.quote(SHEET_NAME, safe="")
    url = (f"{GRAPH_BASE}/sites/{SITE_ID}/drive/items/{item_id}"
           f"/workbook/worksheets/{sheet_enc}/usedRange")

    log.info(f"Leyendo hoja '{SHEET_NAME}'...")
    r = requests.get(url, headers=headers, timeout=120)
    if r.status_code != 200:
        raise RuntimeError(
            f"Error leyendo hoja '{SHEET_NAME}': {r.status_code}\n{r.text[:400]}"
        )

    values = r.json().get("values", [])
    if not values or len(values) < 2:
        raise ValueError("La hoja está vacía o sin datos.")

    cols = [str(x).strip() if x else f"col_{i}" for i, x in enumerate(values[0])]
    rows = []
    for row in values[1:]:
        norm = [c if c is not None else "" for c in row]
        if len(norm) < len(cols):
            norm += [""] * (len(cols) - len(norm))
        rows.append(norm[:len(cols)])

    df = pd.DataFrame(rows, columns=cols)
    log.info(f"{len(df)} filas leídas ✅")
    return df


# ══════════════════════════════════════════════════════════════════
#  LIMPIEZA Y SERIALIZACIÓN — sin calcular KPIs
# ══════════════════════════════════════════════════════════════════
def _parse_date_str(val: Any) -> str:
    """Normaliza cualquier valor de fecha a dd/mm/yyyy o ''."""
    try:
        if pd.isna(val):
            return ""
    except Exception:
        pass
    s = str(val).strip()
    if s in ("", "nan", "NaN", "None"):
        return ""
    # Excel serial number
    s_stripped = s.replace(".", "", 1)
    if s_stripped.isdigit():
        try:
            d = datetime(1899, 12, 30) + timedelta(days=int(float(s)))
            return d.strftime("%d/%m/%Y")
        except Exception:
            return s
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%m/%d/%Y",
                "%d/%m/%y", "%Y-%m-%dT%H:%M:%S", "%Y-%m-%dT%H:%M"):
        try:
            return datetime.strptime(s, fmt).strftime("%d/%m/%Y")
        except ValueError:
            pass
    try:
        return pd.to_datetime(s, dayfirst=True).strftime("%d/%m/%Y")
    except Exception:
        return s


def clean_df(df: pd.DataFrame) -> pd.DataFrame:
    """Limpieza básica: normaliza texto y fechas."""
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]

    for col in df.columns:
        if df[col].dtype == object:
            df[col] = df[col].astype(str).str.strip()
            df[col] = df[col].replace({"nan": "", "None": "", "NaN": ""})

    for col in DATE_COLS:
        if col in df.columns:
            df[col] = df[col].apply(_parse_date_str)

    df = df.dropna(how="all")
    df = df[~(df == "").all(axis=1)]
    return df


def _safe_val(v: Any) -> Any:
    try:
        if pd.isna(v):
            return ""
    except Exception:
        pass
    if isinstance(v, (int, float)):
        return v
    return str(v)


def df_to_rows(df: pd.DataFrame) -> List[Dict[str, Any]]:
    cols = list(df.columns)
    return [{c: _safe_val(row[c]) for c in cols} for _, row in df.iterrows()]


# ══════════════════════════════════════════════════════════════════
#  GENERAR data.json — SOLO datos crudos
#  dashboard.js calcula todos los KPIs a partir de estas filas.
# ══════════════════════════════════════════════════════════════════
def build_data_json(df: pd.DataFrame) -> Dict[str, Any]:
    df_clean = clean_df(df)
    rows     = df_to_rows(df_clean)
    return {
        "generado":  datetime.now().strftime("%d/%m/%Y %H:%M"),
        "filas":     len(rows),
        "columnas":  list(df_clean.columns),
        "rows":      rows,
    }


def write_json(payload: Dict[str, Any]) -> str:
    content = json.dumps(payload, ensure_ascii=False, separators=(",", ":"))
    OUTPUT_JSON.write_text(content, encoding="utf-8")
    size_kb = len(content.encode()) / 1024
    log.info(f"data.json escrito: {len(payload['rows'])} filas, {size_kb:.1f} KB")
    return content


# ══════════════════════════════════════════════════════════════════
#  GITHUB PUSH
# ══════════════════════════════════════════════════════════════════
def push_to_github(content: str):
    if not GITHUB_TOKEN:
        log.warning("GITHUB_TOKEN no configurado — se omite el push.")
        return

    api_url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{GITHUB_FILE_PATH}"
    headers = {
        "Authorization": f"token {GITHUB_TOKEN}",
        "Accept":        "application/vnd.github.v3+json",
    }

    sha: Optional[str] = None
    r = requests.get(api_url, headers=headers, params={"ref": GITHUB_BRANCH}, timeout=30)
    if r.status_code == 200:
        sha = r.json().get("sha")
    elif r.status_code not in (404,):
        log.warning(f"GitHub GET {r.status_code}: {r.text[:200]}")

    content_b64 = base64.b64encode(content.encode("utf-8")).decode("ascii")
    content_hash = hashlib.sha256(content.encode()).hexdigest()[:8]
    log.info(f"{'Actualizando' if sha else 'Creando'} archivo en GitHub (hash: {content_hash})...")

    body: Dict[str, Any] = {
        "message": f"sync: tracking VP {datetime.now().strftime('%Y-%m-%d %H:%M')}",
        "content": content_b64,
        "branch":  GITHUB_BRANCH,
    }
    if sha:
        body["sha"] = sha

    r2 = requests.put(api_url, headers=headers, json=body, timeout=60)
    if r2.status_code in (200, 201):
        log.info(f"✅ data.json subido a GitHub ({GITHUB_REPO}/{GITHUB_BRANCH})")
    else:
        raise RuntimeError(f"Error GitHub PUT {r2.status_code}: {r2.text[:300]}")


# ══════════════════════════════════════════════════════════════════
#  HELPERS PARSEO para cálculo de KPIs del correo
#  (idénticos a dashboard.js — solo se usan para el correo)
# ══════════════════════════════════════════════════════════════════
def get_col(row: dict, *keys) -> str:
    for k in keys:
        if k in row and row[k] is not None and str(row[k]).strip():
            return str(row[k]).strip()
        k_up = k.upper()
        for rk in row:
            if rk.upper() == k_up and row[rk] is not None and str(row[rk]).strip():
                return str(row[rk]).strip()
    return ""


def parse_date(s) -> Optional[datetime]:
    if not s:
        return None
    s = str(s).strip()
    if not s or s.upper() in ("NAN", ""):
        return None
    m = re.match(r"^(\d{1,2})/(\d{1,2})/(\d{4})$", s)
    if m:
        try:
            d = datetime(int(m.group(3)), int(m.group(2)), int(m.group(1)))
            if 2000 < d.year < 2100:
                return d
        except ValueError:
            pass
    m = re.match(r"^(\d{4})-(\d{2})-(\d{2})", s)
    if m:
        try:
            d = datetime(int(m.group(1)), int(m.group(2)), int(m.group(3)))
            if 2000 < d.year < 2100:
                return d
        except ValueError:
            pass
    m = re.match(r"^(\d{1,2})-(\d{1,2})-(\d{4})$", s)
    if m:
        try:
            d = datetime(int(m.group(3)), int(m.group(2)), int(m.group(1)))
            if 2000 < d.year < 2100:
                return d
        except ValueError:
            pass
    s_stripped = s.replace(".", "", 1)
    if s_stripped.isdigit():
        try:
            d = datetime(1899, 12, 30) + timedelta(days=int(float(s)))
            if 2000 < d.year < 2100:
                return d
        except Exception:
            pass
    for fmt in ("%d/%m/%y", "%m/%d/%Y"):
        try:
            d = datetime.strptime(s, fmt)
            if 2000 < d.year < 2100:
                return d
        except ValueError:
            pass
    return None


def get_fecha_limite(row: dict) -> Optional[datetime]:
    exact_cols = [
        "FECHA LIMITE DE ENTREGA", "fecha limite de entrega",
        "FECHA LÍMITE DE ENTREGA", "fecha límite de entrega",
        "FECHA LIMITE", "fecha limite",
    ]
    for col in exact_cols:
        v = get_col(row, col)
        if v:
            d = parse_date(v)
            if d:
                return d
    for k in row:
        kl = k.lower()
        if ("limite" in kl or "límite" in kl) and "solicitud" not in kl and "comercio" not in kl:
            d = parse_date(str(row[k] or ""))
            if d:
                return d
    return None


def find_novedad(row: dict) -> Optional[Dict[str, str]]:
    for col in NOV_EXACT_COLS:
        v = get_col(row, col).strip()
        if v and v != "0" and v.upper() != "NAN":
            return {"col": col, "val": v}
    for k in row:
        kl = k.lower()
        is_nov = any(n in kl for n in NOV_KEY_INCLUDES)
        is_exc = any(x in kl for x in NOV_KEY_EXCLUDES)
        if is_nov and not is_exc:
            v = str(row[k] or "").strip()
            if v and v != "0" and v.upper() != "NAN":
                return {"col": k, "val": v}
    return None


def is_devolucion(row: dict) -> bool:
    for k, v in row.items():
        k_up = k.upper()
        v_up = str(v or "").upper()
        if any(x in v_up for x in ["DEVOLUCI", "DEVOLUCION", "DEVOLUCIÓN", "DEVUELTO", "REMITENTE"]):
            return True
        if any(x in k_up for x in ["DEVOLUCI", "DEVOLUCION", "DEVOLUCIÓN"]):
            if v_up and v_up not in ("", "0", "NAN"):
                return True
    return False


# ══════════════════════════════════════════════════════════════════
#  CÁLCULO DE KPIs para el correo
#  (se calcula sobre el df crudo, mismo que el dashboard.js)
# ══════════════════════════════════════════════════════════════════
def compute_kpis_for_email(df: pd.DataFrame) -> Dict[str, Any]:
    all_rows = [{str(k).strip(): v for k, v in row.items()} for _, row in df.iterrows()]

    # 1. Excluir cancelados
    data = [r for r in all_rows
            if get_col(r, "ESTADO DATAFONO", "estado datafono").upper() != "CANCELADO"]
    cancelados = len(all_rows) - len(data)
    total = len(data)

    # 2. Conteo por estado
    ec: Dict[str, int] = {}
    for r in data:
        e = get_col(r, "ESTADO DATAFONO", "estado datafono").upper() or "SIN ESTADO"
        ec[e] = ec.get(e, 0) + 1

    entregados      = ec.get("ENTREGADO", 0)
    en_transito     = ec.get("EN TRANSITO", 0) + ec.get("EN TRÁNSITO", 0)
    en_alistamiento = ec.get("EN ALISTAMIENTO", 0)
    devueltos       = len([r for r in data if is_devolucion(r)])
    n_alistados     = total - en_alistamiento

    # 3. VT y OPLG
    vt_rows = [r for r in data
               if get_col(r, "TIPO DE SOLICITUD FACTURACIÓN",
                           "TIPO DE SOLICITUD FACTURACION",
                           "tipo de solicitud facturacion").upper() == VT_EXACT.upper()]
    ol_rows = [r for r in data
               if get_col(r, "TIPO DE SOLICITUD FACTURACIÓN",
                           "TIPO DE SOLICITUD FACTURACION",
                           "tipo de solicitud facturacion").upper() == OPLG_EXACT.upper()]

    ent_vt        = len([r for r in vt_rows if get_col(r, "ESTADO DATAFONO", "estado datafono").upper() == "ENTREGADO"])
    ent_ol        = len([r for r in ol_rows if get_col(r, "ESTADO DATAFONO", "estado datafono").upper() == "ENTREGADO"])
    programados_vt = len([r for r in vt_rows if get_col(r, "ESTADO DATAFONO", "estado datafono").upper() == "VISITA PROGRAMADA"])

    # 4. ANS Oportunidad
    ent_df       = [r for r in data if get_col(r, "ESTADO DATAFONO", "estado datafono").upper() == "ENTREGADO"]
    cumple_oport = len([r for r in ent_df if get_col(r, "CUMPLE ANS", "cumple ans", "Cumple Ans").upper() == "SI"])
    pct_oport    = round(cumple_oport / len(ent_df) * 100) if ent_df else 0
    pct_calidad  = round((entregados - devueltos) / entregados * 100) if entregados else 100

    def pct(a, b):
        return round(a / b * 100) if b else 0

    today = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)

    # 5. Vencen hoy — igual que dashboard.js: limDay === today
    vencen_hoy_rows = []
    for r in data:
        est = get_col(r, "ESTADO DATAFONO", "estado datafono").upper()
        if est in ("ENTREGADO", "CANCELADO"):
            continue
        lim = get_fecha_limite(r)
        if lim is None:
            continue
        lim_day = lim.replace(hour=0, minute=0, second=0, microsecond=0)
        if lim_day == today:
            vencen_hoy_rows.append(r)

    # 6. Vencidas ANS — igual que dashboard.js: lim < today
    vencidas_rows = []
    for r in data:
        est = get_col(r, "ESTADO DATAFONO", "estado datafono").upper()
        if est in ("ENTREGADO", "CANCELADO"):
            continue
        lim = get_fecha_limite(r)
        if lim is not None and lim < today:
            vencidas_rows.append(r)

    # 7. Backlog 24h — igual que dashboard.js renderBacklog(hrs=24)
    cutoff_24h = today + timedelta(hours=24)
    backlog_24h_rows = []
    for r in data:
        est = get_col(r, "ESTADO DATAFONO", "estado datafono").upper()
        if est in ("ENTREGADO", "CANCELADO"):
            continue
        lim = get_fecha_limite(r)
        if lim is None:
            continue
        lim_day_end = lim.replace(hour=23, minute=59, second=59, microsecond=999999)
        if lim_day_end >= today and lim <= cutoff_24h:
            backlog_24h_rows.append(r)

    # 8. Backlog 48h
    cutoff_48h = today + timedelta(hours=48)
    backlog_48h_rows = []
    for r in data:
        est = get_col(r, "ESTADO DATAFONO", "estado datafono").upper()
        if est in ("ENTREGADO", "CANCELADO"):
            continue
        lim = get_fecha_limite(r)
        if lim is None:
            continue
        lim_day_end = lim.replace(hour=23, minute=59, second=59, microsecond=999999)
        if lim_day_end >= today and lim <= cutoff_48h:
            backlog_48h_rows.append(r)

    # 9. Guías sin cambios >24h — igual que dashboard.js renderStalledGuias():
    #    limDay <= today (vencida, incluyendo hoy) Y findNovedad != null
    #    CORRECCIÓN: lim_day > today descarta las futuras; lim_day <= today las incluye.
    guias_sin_cambios_rows = []
    for r in data:
        est = get_col(r, "ESTADO DATAFONO", "estado datafono").upper()
        if est in ("ENTREGADO", "CANCELADO"):
            continue
        lim = get_fecha_limite(r)
        if lim is None:
            continue
        lim_day = lim.replace(hour=0, minute=0, second=0, microsecond=0)
        if lim_day > today:           # ← CORRECTO: excluir solo las FUTURAS
            continue
        dias = (today - lim_day).days
        if dias >= 1 and find_novedad(r) is not None:
            guias_sin_cambios_rows.append({**r, "_dias_sin_cambios": dias})

    # 10. Intentos fallidos — igual que dashboard.js renderFallidos():
    #     est == EN TRANSITO y findNovedad != null
    intentos_fallidos_rows = []
    for r in data:
        est = get_col(r, "ESTADO DATAFONO", "estado datafono").upper()
        if est not in ("EN TRANSITO", "EN TRÁNSITO"):
            continue
        if find_novedad(r) is not None:
            intentos_fallidos_rows.append(r)

    def row_info(r: dict) -> dict:
        lim = get_fecha_limite(r)
        return {
            "comercio":       get_col(r, "Nombre del comercio", "nombre del comercio", "NOMBRE DEL COMERCIO") or "—",
            "id_sitio":       get_col(r, "ID Comercio", "id comercio") or "—",
            "guia":           get_col(r, "NÚMERO DE GUIA", "NUMERO DE GUIA", "numero de guia") or "—",
            "transportadora": get_col(r, "TRANSPORTADORA", "Transportadora", "transportadora") or "—",
            "estado":         get_col(r, "ESTADO DATAFONO", "estado datafono") or "—",
            "fecha_limite":   lim.strftime("%d/%m/%Y") if lim else "—",
            "tipo":           get_col(r, "TIPO DE SOLICITUD FACTURACIÓN",
                                      "TIPO DE SOLICITUD FACTURACION",
                                      "tipo de solicitud facturacion") or "—",
            "novedad":        (find_novedad(r) or {}).get("val", "") or "—",
        }

    return {
        "fecha_informe":          datetime.now().strftime("%d/%m/%Y %H:%M"),
        "total":                  total,
        "cancelados":             cancelados,
        "entregados":             entregados,
        "en_transito":            en_transito,
        "en_alistamiento":        en_alistamiento,
        "n_alistados":            n_alistados,
        "devueltos":              devueltos,
        "total_vt":               len(vt_rows),
        "ent_vt":                 ent_vt,
        "programados_vt":         programados_vt,
        "total_ol":               len(ol_rows),
        "ent_ol":                 ent_ol,
        "pct_entregado":          pct(entregados, total),
        "pct_transito":           pct(en_transito, total),
        "pct_alistamiento":       pct(en_alistamiento, total),
        "pct_n_alistados":        pct(n_alistados, total),
        "pct_vt":                 pct(ent_vt, len(vt_rows)),
        "pct_ol":                 pct(ent_ol, len(ol_rows)),
        "pct_oport":              pct_oport,
        "pct_calidad":            pct_calidad,
        "vencen_hoy":             len(vencen_hoy_rows),
        "vencen_hoy_data":        [row_info(r) for r in vencen_hoy_rows],
        "vencidas":               len(vencidas_rows),
        "vencidas_data":          [row_info(r) for r in vencidas_rows],
        "backlog_24h":            len(backlog_24h_rows),
        "backlog_24h_data":       [row_info(r) for r in backlog_24h_rows],
        "backlog_48h":            len(backlog_48h_rows),
        "backlog_48h_data":       [row_info(r) for r in backlog_48h_rows],
        "guias_sin_cambios":      len(guias_sin_cambios_rows),
        "guias_sin_cambios_data": [
            {**row_info(r), "dias": r.get("_dias_sin_cambios", 0)}
            for r in guias_sin_cambios_rows
        ],
        "intentos_fallidos":      len(intentos_fallidos_rows),
        "intentos_fallidos_data": [row_info(r) for r in intentos_fallidos_rows],
    }


# ══════════════════════════════════════════════════════════════════
#  HTML CORREO — Wompi Brand 2025
# ══════════════════════════════════════════════════════════════════
def _truncar(texto: str, max_chars: int = 60) -> str:
    t = str(texto or "—").strip()
    if not t:
        return "—"
    return t if len(t) <= max_chars else t[:max_chars] + "…"


def _rows_table_html(rows: list, extra_cols=None, max_rows=20) -> str:
    if not rows:
        return (
            '<table width="100%" cellpadding="0" cellspacing="0" border="0">'
            '<tr><td align="center" style="padding:18px 12px;'
            'background-color:#1C1F1C;border:1px dashed #2E3C2E">'
            '<p style="margin:0;color:#4A7A4A;font-size:20px">&#10003;</p>'
            '<p style="margin:6px 0 0;color:#6B6967;font-size:12px;'
            'font-family:Arial,Helvetica,sans-serif">'
            'Sin registros en este indicador</p>'
            '</td></tr></table>'
        )

    base_cols = [
        ("comercio",       "Comercio"),
        ("guia",           "Gu&#237;a"),
        ("transportadora", "Transp."),
        ("estado",         "Estado"),
        ("fecha_limite",   "L&#237;mite"),
    ]
    extra_raw = extra_cols or [("novedad", "Novedad")]
    extra = [(item[0], item[1]) if len(item) >= 2 else (item[0], item[0]) for item in extra_raw]
    cols  = base_cols + extra

    visible = rows[:max_rows]
    omitted = len(rows) - len(visible)

    th = "".join(
        f'<th style="padding:7px 10px;text-align:left;'
        f'background-color:#162116;color:#B0F2AE;font-weight:700;font-size:11px;'
        f'font-family:Arial,Helvetica,sans-serif;'
        f'border-bottom:2px solid #00825A;white-space:nowrap">{label}</th>'
        for _, label in cols
    )

    body_rows = ""
    for i, r in enumerate(visible):
        bg = "#191C19" if i % 2 == 0 else "#1D201D"
        cells = ""
        for key, _ in cols:
            raw = r.get(key, "—") or "—"
            if key == "estado":
                eu = str(raw).upper()
                c  = ("#B0F2AE" if "ENTREGADO"    in eu else
                      "#99D1FC"  if "TRANSITO"     in eu else
                      "#FFC04D"  if ("ALISTAMIENTO" in eu or "PROGRAMADA" in eu) else
                      "#FF5C5C")
                short = (str(raw)[:13] + "…") if len(str(raw)) > 14 else str(raw)
                cell_v = (f'<span style="color:{c};font-size:10px;font-weight:700;'
                          f'font-family:Arial,Helvetica,sans-serif">{short}</span>')
            elif key == "dias":
                d_val = r.get("dias", 0)
                col   = "#FF5C5C" if d_val >= 7 else "#FFC04D" if d_val >= 3 else "#99D1FC"
                cell_v = (f'<span style="color:{col};font-weight:700;font-size:11px;'
                          f'font-family:Arial,Helvetica,sans-serif">{d_val}d</span>')
            elif key == "novedad":
                texto_full  = str(raw).strip()
                texto_corto = _truncar(texto_full, 50)
                if texto_full in ("—", ""):
                    cell_v = '<span style="color:#4A4845;font-size:11px;font-family:Arial,Helvetica,sans-serif">&#8212;</span>'
                else:
                    cell_v = (f'<span style="color:#FFC04D;font-size:10px;'
                              f'font-family:Arial,Helvetica,sans-serif">{texto_corto}</span>')
            elif key == "comercio":
                cell_v = (f'<span style="color:#FAFAFA;font-size:11px;font-weight:600;'
                          f'font-family:Arial,Helvetica,sans-serif">{_truncar(raw, 22)}</span>')
            elif key == "guia":
                cell_v = (f'<span style="color:#B0F2AE;font-size:11px;font-weight:600;'
                          f'font-family:Courier New,Courier,monospace">{_truncar(raw, 16)}</span>')
            elif key == "transportadora":
                cell_v = (f'<span style="color:#99D1FC;font-size:11px;'
                          f'font-family:Arial,Helvetica,sans-serif">{_truncar(raw, 11)}</span>')
            else:
                cell_v = (f'<span style="color:#DEDBD8;font-size:11px;'
                          f'font-family:Arial,Helvetica,sans-serif">{_truncar(raw, 14)}</span>')

            cells += (f'<td style="padding:7px 10px;'
                      f'border-bottom:1px solid #232623;vertical-align:middle">'
                      f'{cell_v}</td>')
        body_rows += f'<tr style="background-color:{bg}">{cells}</tr>'

    extra_note = ""
    if omitted > 0:
        extra_note = (
            f'<tr><td colspan="{len(cols)}" style="text-align:center;padding:10px;'
            f'color:#6B6967;font-size:11px;background-color:#141614;'
            f'border-top:1px solid #232623;font-family:Arial,Helvetica,sans-serif">'
            f'&#8230;y <strong style="color:#B0F2AE">{omitted}</strong> registros m&#225;s. '
            f'<a href="https://segoloo.github.io/TablerosWompi/" '
            f'style="color:#B0F2AE;text-decoration:none;font-weight:600">Ver todos &#8594;</a>'
            f'</td></tr>'
        )

    return (
        '<table width="100%" cellpadding="0" cellspacing="0" border="0" '
        'style="border-collapse:collapse;border:1px solid #232623">'
        f'<thead><tr>{th}</tr></thead>'
        f'<tbody>{body_rows}{extra_note}</tbody>'
        '</table>'
    )


def _kpi_card_html(value: str, label: str, color: str, icon: str, sub: str = "") -> str:
    sub_html = (f'<p style="margin:2px 0 0;font-size:9px;color:{color};'
                f'font-family:Arial,Helvetica,sans-serif">{sub}</p>') if sub else ''
    return (
        f'<td style="width:14%;padding:4px;vertical-align:top">'
        f'<table width="100%" cellpadding="0" cellspacing="0" border="0" '
        f'style="background-color:#1A1F1A;border-top:3px solid {color}">'
        f'<tr><td align="center" style="padding:12px 8px 10px">'
        f'<p style="margin:0 0 5px;font-size:18px;line-height:1">{icon}</p>'
        f'<p style="margin:0 0 4px;font-size:19px;font-weight:bold;color:{color};'
        f'font-family:Georgia,serif;line-height:1">{value}</p>'
        f'<p style="margin:0;font-size:9px;color:#6B6967;text-transform:uppercase;'
        f'letter-spacing:0.8px;font-family:Arial,Helvetica,sans-serif">{label}</p>'
        f'{sub_html}'
        f'</td></tr></table></td>'
    )


def _progress_bar(pct: int, color: str) -> str:
    filled = max(0, min(100, pct))
    empty  = 100 - filled
    return (
        f'<table width="100%" cellpadding="0" cellspacing="0" border="0" style="margin-top:5px">'
        f'<tr>'
        f'<td width="{filled}%" style="height:5px;background-color:{color};font-size:1px;line-height:1px">&nbsp;</td>'
        + (f'<td width="{empty}%" style="height:5px;background-color:#232623;font-size:1px;line-height:1px">&nbsp;</td>' if empty > 0 else '') +
        f'</tr></table>'
    )


def _tipo_row(label: str, ejec: int, total: int, pct: int, color: str) -> str:
    return (
        f'<tr>'
        f'<td style="padding:8px 12px;color:#DEDBD8;font-size:12px;font-weight:bold;'
        f'font-family:Arial,Helvetica,sans-serif">{label}</td>'
        f'<td style="padding:8px 6px;color:{color};font-size:13px;font-weight:bold;'
        f'font-family:Georgia,serif;text-align:center">{ejec}</td>'
        f'<td style="padding:8px 6px;color:#6B6967;font-size:11px;text-align:center;'
        f'font-family:Arial,Helvetica,sans-serif">/ {total}</td>'
        f'<td style="padding:8px 12px">'
        + _progress_bar(pct, color) +
        f'<p style="margin:3px 0 0;font-size:10px;color:{color};text-align:right;'
        f'font-family:Arial,Helvetica,sans-serif">{pct}%</p>'
        f'</td></tr>'
    )


def _alert_section_html(title: str, subtitle: str, color: str, icon: str,
                        count: int, table_html: str, alert: bool = False) -> str:
    border   = "#FF5C5C" if alert and count > 0 else color
    hdr_bg   = "#1E1210" if alert and count > 0 else "#141A14"
    count_bg = "#FF5C5C" if alert and count > 0 else color
    count_fg = "#FAFAFA" if alert and count > 0 else "#0A0A0A"
    return (
        f'<table width="100%" cellpadding="0" cellspacing="0" border="0" '
        f'style="margin:16px 0;border:1px solid {border}">'
        f'<tr style="background-color:{hdr_bg}">'
        f'<td style="padding:12px 16px;border-bottom:1px solid {border}">'
        f'<span style="font-size:15px;vertical-align:middle">{icon}</span>'
        f'&nbsp;<span style="font-size:13px;font-weight:bold;color:{color};'
        f'vertical-align:middle;font-family:Arial,Helvetica,sans-serif">{title}</span>'
        f'<br><span style="font-size:10px;color:#6B6967;'
        f'font-family:Arial,Helvetica,sans-serif">{subtitle}</span>'
        f'</td>'
        f'<td align="right" style="padding:12px 16px;border-bottom:1px solid {border};'
        f'white-space:nowrap;vertical-align:middle">'
        f'<span style="background-color:{count_bg};color:{count_fg};font-weight:bold;'
        f'font-size:16px;padding:4px 12px;font-family:Georgia,serif">{count}</span>'
        f'</td></tr>'
        f'<tr><td colspan="2" style="padding:12px 16px">{table_html}</td></tr>'
        f'</table>'
    )


def build_html_email(kpis: Dict[str, Any]) -> str:
    fecha         = kpis["fecha_informe"]
    dashboard_url = "https://segoloo.github.io/TablerosWompi/"

    total_alert = (kpis["vencen_hoy"] + kpis["vencidas"] +
                   kpis["guias_sin_cambios"] + kpis["intentos_fallidos"])
    if total_alert > 5:
        alert_color = C_DANGER
        alert_bg    = "#1E1210"
        alert_icon  = "🚨"
        alert_msg   = f"ALERTA: {total_alert} registros requieren atención inmediata"
    elif total_alert > 0:
        alert_color = C_WARNING
        alert_bg    = "#1E1A10"
        alert_icon  = "⚠️"
        alert_msg   = f"Atención: {total_alert} registros en seguimiento"
    else:
        alert_color = C_VERDE_MENTA
        alert_bg    = "#101E10"
        alert_icon  = "✅"
        alert_msg   = "Todo en orden — sin alertas activas"

    s_hoy = _alert_section_html(
        "Gu&#237;as que vencen hoy", "Sin entregar con l&#237;mite hoy",
        C_WARNING, "&#9200;", kpis["vencen_hoy"],
        _rows_table_html(kpis["vencen_hoy_data"]),
        alert=kpis["vencen_hoy"] > 0)

    s_vencidas = _alert_section_html(
        "Vencidas ANS", "Fuera de plazo (VT + OPLG)",
        C_DANGER, "&#128680;", kpis["vencidas"],
        _rows_table_html(kpis["vencidas_data"]),
        alert=kpis["vencidas"] > 0)

    s_bl24 = _alert_section_html(
        "Backlog por vencer (24h)", "Gu&#237;as cuyo l&#237;mite vence en las pr&#243;ximas 24 h",
        C_AZUL_CIELO, "&#9889;", kpis["backlog_24h"],
        _rows_table_html(kpis["backlog_24h_data"]))

    s_bl48 = _alert_section_html(
        "Backlog por vencer (48h)", "Gu&#237;as cuyo l&#237;mite vence en las pr&#243;ximas 48 h",
        C_AZUL_CIELO, "&#128197;", kpis["backlog_48h"],
        _rows_table_html(kpis["backlog_48h_data"]))

    s_stalled = _alert_section_html(
        "Gu&#237;as sin cambios &gt;24h", "Vencidas ANS con novedad registrada",
        C_WARNING, "&#128274;", kpis["guias_sin_cambios"],
        _rows_table_html(kpis["guias_sin_cambios_data"],
                         extra_cols=[("dias", "D&#237;as"), ("novedad", "Novedad")]),
        alert=kpis["guias_sin_cambios"] > 0)

    s_fallidos = _alert_section_html(
        "Gu&#237;as con intento fallido", "En tr&#225;nsito con novedad de intento fallido",
        C_DANGER, "&#10060;", kpis["intentos_fallidos"],
        _rows_table_html(kpis["intentos_fallidos_data"], extra_cols=[("novedad", "Novedad")]),
        alert=kpis["intentos_fallidos"] > 0)

    tipo_rows = (
        _tipo_row("Visita T&#233;cnica (VT)",
                  kpis["ent_vt"], kpis["total_vt"], kpis["pct_vt"], C_VERDE_MENTA) +
        _tipo_row("Op. Log&#237;stico (OPLG)",
                  kpis["ent_ol"], kpis["total_ol"], kpis["pct_ol"], C_AZUL_CIELO)
    )

    return f"""<!DOCTYPE html>
<html lang="es" xmlns="http://www.w3.org/1999/xhtml">
<head>
  <meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Reporte Operativo Wompi VP</title>
  <style type="text/css">
    body, table, td, p, a {{ -ms-text-size-adjust:100%; -webkit-text-size-adjust:100%; }}
    table, td {{ mso-table-lspace:0pt; mso-table-rspace:0pt; }}
    body {{ margin:0!important; padding:0!important; background-color:#0C0F0C; }}
  </style>
</head>
<body style="margin:0;padding:0;background-color:#0C0F0C;font-family:Arial,Helvetica,sans-serif">

<table role="presentation" cellpadding="0" cellspacing="0" border="0"
       width="100%" style="background-color:#0C0F0C">
<tr><td align="center" style="padding:20px 10px">

  <table role="presentation" cellpadding="0" cellspacing="0" border="0"
         width="620" style="max-width:620px;width:100%">

    <!-- Barra tricolor superior -->
    <tr>
      <td width="33%" height="4" style="background-color:{C_VERDE_LIMA};font-size:1px;line-height:4px">&nbsp;</td>
      <td width="34%" height="4" style="background-color:{C_VERDE_MENTA};font-size:1px;line-height:4px">&nbsp;</td>
      <td width="33%" height="4" style="background-color:{C_AZUL_CIELO};font-size:1px;line-height:4px">&nbsp;</td>
    </tr>

    <!-- Header -->
    <tr>
      <td colspan="3" bgcolor="{C_NEGRO_CIB}" style="background-color:{C_NEGRO_CIB}">
        <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%">
          <tr>
            <td style="padding:22px 24px 14px;vertical-align:top">
              <p style="margin:0 0 3px 0;font-size:10px;font-weight:bold;color:{C_VERDE_LIMA};
                         letter-spacing:2px;text-transform:uppercase;font-family:Arial,Helvetica,sans-serif">Reporte Operativo</p>
              <p style="margin:0 0 4px 0;font-size:28px;font-weight:bold;color:{C_VERDE_MENTA};
                         letter-spacing:-0.5px;line-height:1.1;font-family:Georgia,Times New Roman,serif">Wompi VP</p>
              <p style="margin:0;font-size:11px;color:#5A5856;font-family:Arial,Helvetica,sans-serif">
                Tracking Venta Presente &middot; By Sebasti&#225;n G&#243;mez L&#243;pez
              </p>
            </td>
            <td align="right" style="padding:22px 24px 14px;vertical-align:top">
              <table role="presentation" cellpadding="0" cellspacing="0" border="0">
                <tr>
                  <td bgcolor="#1A1816" style="background-color:#1A1816;border:1px solid #3A3835;
                      padding:10px 14px;text-align:center">
                    <p style="margin:0 0 3px 0;font-size:9px;color:#5A5856;text-transform:uppercase;
                               letter-spacing:1px;font-family:Arial,Helvetica,sans-serif">Generado</p>
                    <p style="margin:0;font-size:12px;font-weight:bold;color:{C_VERDE_MENTA};
                               font-family:Arial,Helvetica,sans-serif">{fecha}</p>
                  </td>
                </tr>
              </table>
            </td>
          </tr>
          <tr>
            <td colspan="2" bgcolor="{C_VERDE_MENTA}" style="background-color:{C_VERDE_MENTA};padding:8px 24px">
              <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%">
                <tr>
                  <td style="font-size:11px;font-weight:bold;color:{C_NEGRO_CIB};
                              text-transform:uppercase;letter-spacing:0.5px;font-family:Arial,Helvetica,sans-serif">
                    Informe Automatizado ANS
                  </td>
                  <td align="right" style="font-size:11px;font-weight:bold;color:{C_VERDE_SELVA};
                              font-family:Arial,Helvetica,sans-serif">
                    {kpis["total"]} registros activos
                  </td>
                </tr>
              </table>
            </td>
          </tr>
        </table>
      </td>
    </tr>

    <!-- Cuerpo -->
    <tr>
      <td colspan="3" bgcolor="#0F120F" style="background-color:#0F120F;padding:20px 20px 24px">

        <!-- Banner estado global -->
        <table role="presentation" cellpadding="0" cellspacing="0" border="0"
               width="100%" style="margin-bottom:16px">
          <tr>
            <td bgcolor="{alert_bg}" style="background-color:{alert_bg};
                border-left:4px solid {alert_color};padding:12px 16px">
              <p style="margin:0;font-size:13px;font-weight:bold;color:{alert_color};
                          font-family:Arial,Helvetica,sans-serif">
                {alert_icon}&nbsp;&nbsp;{alert_msg}
              </p>
            </td>
          </tr>
        </table>

        <!-- CTA Dashboard -->
        <table role="presentation" cellpadding="0" cellspacing="0" border="0"
               width="100%" style="margin-bottom:20px">
          <tr>
            <td bgcolor="{C_VERDE_SELVA}" align="center"
                style="background-color:{C_VERDE_SELVA};padding:16px 20px">
              <p style="margin:0 0 10px 0;font-size:11px;font-weight:bold;
                          color:{C_NEGRO_CIB};font-family:Arial,Helvetica,sans-serif">
                Para filtros, hist&#243;rico y an&#225;lisis completo:
              </p>
              <table role="presentation" cellpadding="0" cellspacing="0" border="0">
                <tr>
                  <td bgcolor="{C_NEGRO_CIB}" style="background-color:{C_NEGRO_CIB};padding:10px 26px">
                    <a href="{dashboard_url}"
                       style="color:{C_VERDE_MENTA};text-decoration:none;font-weight:bold;
                              font-size:12px;font-family:Arial,Helvetica,sans-serif;letter-spacing:0.3px">
                      &#128279;&nbsp; Abrir Dashboard Wompi VP &nbsp;&#8594;
                    </a>
                  </td>
                </tr>
              </table>
            </td>
          </tr>
        </table>

        <!-- Separador Alertas -->
        <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%"
               style="margin-bottom:10px;border-bottom:1px solid #1E251E">
          <tr><td style="padding-bottom:8px">
            <p style="margin:0;font-size:10px;color:#4A7A4A;text-transform:uppercase;
                       letter-spacing:1.5px;font-weight:bold;font-family:Arial,Helvetica,sans-serif">
              Alertas Operativas
            </p>
          </td></tr>
        </table>

        {s_hoy}
        {s_vencidas}
        {s_bl24}
        {s_bl48}
        {s_stalled}
        {s_fallidos}

      </td>
    </tr>

    <!-- Barra tricolor footer -->
    <tr>
      <td width="33%" height="2" style="background-color:{C_VERDE_LIMA};font-size:1px;line-height:2px">&nbsp;</td>
      <td width="34%" height="2" style="background-color:{C_VERDE_MENTA};font-size:1px;line-height:2px">&nbsp;</td>
      <td width="33%" height="2" style="background-color:{C_AZUL_CIELO};font-size:1px;line-height:2px">&nbsp;</td>
    </tr>

    <!-- Footer -->
    <tr>
      <td colspan="3" bgcolor="{C_NEGRO_CIB}" style="background-color:{C_NEGRO_CIB}">
        <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%">
          <tr>
            <td style="padding:14px 22px;vertical-align:middle">
              <span style="font-size:14px;font-weight:bold;color:{C_VERDE_MENTA};
                           font-family:Georgia,Times New Roman,serif">W Wompi</span>
              <span style="font-size:10px;color:#3C3A38;
                           font-family:Arial,Helvetica,sans-serif"> &middot; powered by Lineacom</span>
            </td>
            <td align="right" style="padding:14px 22px;vertical-align:middle">
              <p style="margin:0;font-size:9px;color:#3C3A38;font-family:Arial,Helvetica,sans-serif">
                {fecha} &middot; generado autom&#225;ticamente
              </p>
              <p style="margin:2px 0 0 0;font-size:9px;font-weight:bold;
                          color:{C_VERDE_LIMA};letter-spacing:1px;font-family:Arial,Helvetica,sans-serif">
                CONFIDENCIAL
              </p>
            </td>
          </tr>
        </table>
      </td>
    </tr>

  </table>

</td></tr>
</table>

</body>
</html>"""


def build_plain_text(kpis: Dict[str, Any]) -> str:
    return "\n".join([
        f"REPORTE OPERATIVO WOMPI VP — {kpis['fecha_informe']}",
        "=" * 60,
        "",
        "ALERTAS OPERATIVAS",
        f"  ⏰ Vencen Hoy           : {kpis['vencen_hoy']}",
        f"  🚨 Vencidas ANS         : {kpis['vencidas']}",
        f"  ⚡ Backlog 24h          : {kpis['backlog_24h']}",
        f"  📅 Backlog 48h          : {kpis['backlog_48h']}",
        f"  🔒 Sin cambios >24h     : {kpis['guias_sin_cambios']}",
        f"  ❌ Intentos fallidos    : {kpis['intentos_fallidos']}",
        "",
        "Para ver el detalle completo ingrese a:",
        "https://segoloo.github.io/TablerosWompi/",
        "",
        "Equipo Analítica TI — Lineacom",
    ])


# ══════════════════════════════════════════════════════════════════
#  CONSTRUCCIÓN Y ENVÍO DE CORREO
# ══════════════════════════════════════════════════════════════════
def construir_correo(kpis: Dict[str, Any]) -> MIMEMultipart:
    total_alert = (kpis["vencen_hoy"] + kpis["vencidas"] +
                   kpis["guias_sin_cambios"] + kpis["intentos_fallidos"])
    prefix = "🚨" if total_alert > 5 else "⚠️" if total_alert > 0 else "✅"
    asunto = f"{prefix} Reporte ANS Wompi VP — {kpis['fecha_informe']}"

    msg = MIMEMultipart("mixed")
    msg["From"]    = formataddr(("Analítica TI · Lineacom", EMAIL_USER))
    msg["To"]      = ", ".join(TO_EMAILS)
    if CC_EMAILS:
        msg["Cc"]  = ", ".join(CC_EMAILS)
    msg["Subject"] = asunto

    alt = MIMEMultipart("alternative")
    alt.attach(MIMEText(build_plain_text(kpis), "plain", "utf-8"))
    alt.attach(MIMEText(build_html_email(kpis), "html",  "utf-8"))
    msg.attach(alt)
    return msg


def enviar_correo(msg: MIMEMultipart):
    all_rx = list(dict.fromkeys(TO_EMAILS + CC_EMAILS))
    server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
    server.starttls()
    server.login(EMAIL_USER, EMAIL_PASSWORD)
    server.sendmail(EMAIL_USER, all_rx, msg.as_string())
    server.quit()
    log.info("✅ Correo enviado exitosamente.")


# ══════════════════════════════════════════════════════════════════
#  MAIN PIPELINE
# ══════════════════════════════════════════════════════════════════
def run_once(no_push: bool = False, no_mail: bool = False):
    start = datetime.now()
    log.info("═" * 60)
    log.info(f"  SYNC WOMPI VP — {start.strftime('%Y-%m-%d %H:%M:%S')}")
    log.info("═" * 60)

    _token_cache.get()
    item_id = find_workbook_item_id()
    df      = read_sheet(item_id)
    log.info(f"  Columnas detectadas: {list(df.columns)[:8]} ...")

    # ── Generar y guardar data.json (datos crudos para el dashboard) ──
    payload = build_data_json(df)
    content = write_json(payload)

    if not no_push:
        push_to_github(content)
    else:
        log.info("--no-push activo: se omite subida a GitHub.")

    # ── Calcular KPIs y enviar correo ────────────────────────────────
    if not no_mail:
        log.info("Calculando KPIs para correo...")
        df_clean = clean_df(df)
        kpis = compute_kpis_for_email(df_clean)

        log.info(f"   ⏰ Vencen Hoy        : {kpis['vencen_hoy']}")
        log.info(f"   🚨 Vencidas ANS      : {kpis['vencidas']}")
        log.info(f"   ⚡ Backlog 24h       : {kpis['backlog_24h']}")
        log.info(f"   📅 Backlog 48h       : {kpis['backlog_48h']}")
        log.info(f"   🔒 Sin cambios >24h  : {kpis['guias_sin_cambios']}")
        log.info(f"   ❌ Intentos fallidos : {kpis['intentos_fallidos']}")

        msg = construir_correo(kpis)
        enviar_correo(msg)
        log.info(f"✅ Correo enviado a: {', '.join(TO_EMAILS)}")
    else:
        log.info("--no-mail activo: se omite envío de correo.")

    elapsed = (datetime.now() - start).total_seconds()
    log.info(f"✅ Ciclo completado en {elapsed:.1f}s — {payload['filas']} registros")


def main():
    no_push = "--no-push" in sys.argv
    no_mail = "--no-mail" in sys.argv
    loop    = "--loop"    in sys.argv

    if loop:
        log.info(f"Modo loop cada {INTERVAL_HOURS}h. Ctrl+C para detener.")
        while True:
            try:
                run_once(no_push=no_push, no_mail=no_mail)
            except Exception as e:
                log.error(f"Ciclo fallido: {e}. Reintentando en {INTERVAL_HOURS}h...")
                import traceback
                traceback.print_exc()
            time.sleep(INTERVAL_HOURS * 3600)
    else:
        try:
            run_once(no_push=no_push, no_mail=no_mail)
        except Exception as e:
            log.error(f"❌ Error: {e}")
            import traceback
            traceback.print_exc()
            sys.exit(1)


if __name__ == "__main__":
    main()