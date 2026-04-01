#!/usr/bin/env python3
"""
╔══════════════════════════════════════════════════════════════════╗
║   sync_wompi_vp.py — Extractor SharePoint → data.json → GitHub  ║
║   LINEACOM · Dashboard Tracking VP Wompi                         ║
║                                                                  ║
║   Uso:                                                           ║
║     python sync_wompi_vp.py              # ejecutar una vez      ║
║     python sync_wompi_vp.py --loop       # cada N horas          ║
║     python sync_wompi_vp.py --no-push    # genera JSON sin subir ║
╚══════════════════════════════════════════════════════════════════╝
"""

import os
import sys
import gzip
import json
import time
import base64
import logging
import hashlib
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any, Dict, List, Optional

import requests
import pandas as pd

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s"
)
log = logging.getLogger("sync_wompi_vp")


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
GITHUB_TOKEN     = os.getenv("GITHUB_TOKEN",     "")     # <-- pon tu PAT aquí
GITHUB_REPO      = os.getenv("GITHUB_REPO",      "TuOrg/TuRepo")
GITHUB_BRANCH    = os.getenv("GITHUB_BRANCH",    "main")
GITHUB_FILE_PATH = os.getenv("GITHUB_FILE_PATH", "data.json")

# ── Loop ──────────────────────────────────────────────────────────
INTERVAL_HOURS = int(os.getenv("INTERVAL_HOURS", "4"))

# ── Output local ─────────────────────────────────────────────────
OUTPUT_JSON = Path("data.json")


# ══════════════════════════════════════════════════════════════════
#  TOKEN CACHE
# ══════════════════════════════════════════════════════════════════
class TokenCache:
    def __init__(self):
        self._token: Optional[str] = None
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
SEARCH_KEYWORDS = ["Comodato", "VP", "Operaci", "Wompi"]   # términos para encontrar el archivo


def find_workbook_item_id() -> str:
    """Busca el archivo Excel de Operación VP_Comodato en SharePoint."""
    token   = _token_cache.get()
    headers = {"Authorization": f"Bearer {token}"}

    # Búsqueda por nombre
    for kw in SEARCH_KEYWORDS:
        url = f"{GRAPH_BASE}/sites/{SITE_ID}/drive/root/search(q='{requests.utils.quote(kw, safe='')}' )"
        r   = requests.get(url, headers=headers, timeout=60)
        if r.status_code == 200:
            for item in r.json().get("value", []):
                name = item.get("name", "")
                if any(k.lower() in name.lower() for k in ["Comodato", "VP_Comodato"]):
                    log.info(f"Archivo encontrado: {name}")
                    return item["id"]

    # Fallback: listar raíz
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
#  LIMPIEZA Y NORMALIZACIÓN
# ══════════════════════════════════════════════════════════════════
DATE_COLS = [
    "FECHA DE SOLICITUD",
    "FECHA LIMITE DE ENTREGA",
    "FECHA ENTREGA AL COMERCIO",
    "FECHA VISITA TECNICA",
    "FECHA DE ENTREGA",
]


def _parse_date_str(val: Any) -> str:
    """Normaliza cualquier valor de fecha a ISO dd/mm/yyyy o ''."""
    if pd.isna(val) or str(val).strip() in ("", "nan", "NaN", "None"):
        return ""
    s = str(val).strip()
    # Excel serial number
    if s.replace(".", "", 1).isdigit():
        try:
            d = datetime(1899, 12, 30) + timedelta(days=int(float(s)))
            return d.strftime("%d/%m/%Y")
        except Exception:
            return s
    # Try common formats
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
    """Limpieza general y normalización de fechas."""
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]

    # Normalizar texto
    for col in df.columns:
        if df[col].dtype == object:
            df[col] = df[col].astype(str).str.strip()
            df[col] = df[col].replace({"nan": "", "None": "", "NaN": ""})

    # Normalizar fechas
    for col in DATE_COLS:
        if col in df.columns:
            df[col] = df[col].apply(_parse_date_str)

    # Eliminar filas completamente vacías
    df = df.dropna(how="all")
    df = df[~(df == "").all(axis=1)]

    return df


# ══════════════════════════════════════════════════════════════════
#  KPI SUMMARY (pre-calculado para carga rápida en el frontend)
# ══════════════════════════════════════════════════════════════════
def compute_summary(df: pd.DataFrame) -> Dict[str, Any]:
    """Calcula KPIs resumen para incluir en data.json."""
    d = df.copy()
    d.columns = [str(c).strip().upper() for c in d.columns]
    for c in d.columns:
        if d[c].dtype == object:
            d[c] = d[c].astype(str).str.strip().str.upper()

    # Excluir cancelados para métricas
    if "ESTADO DATAFONO" in d.columns:
        dA = d[d["ESTADO DATAFONO"] != "CANCELADO"].copy()
    else:
        dA = d.copy()

    total = len(dA)
    ec    = dA["ESTADO DATAFONO"].value_counts().to_dict() if "ESTADO DATAFONO" in dA.columns else {}

    entregados      = ec.get("ENTREGADO", 0)
    en_transito     = ec.get("EN TRANSITO", ec.get("EN TRÁNSITO", 0))
    en_alistamiento = ec.get("EN ALISTAMIENTO", 0)
    devueltos       = sum(v for k, v in ec.items() if "DEVOLU" in k or "REMIT" in k)
    cancelados      = len(d) - total

    # Visita Técnica
    COL_TF = "TIPO DE SOLICITUD FACTURACIÓN"
    if COL_TF in dA.columns:
        vt = dA[dA[COL_TF].str.contains("VISITA", na=False)]
        ol = dA[dA[COL_TF].str.contains("ENVIO|ENVÍO", na=False)]
    else:
        vt = ol = pd.DataFrame()

    entVT = len(vt[vt["ESTADO DATAFONO"] == "ENTREGADO"]) if len(vt) else 0
    entOL = len(ol[ol["ESTADO DATAFONO"] == "ENTREGADO"]) if len(ol) else 0

    # ANS
    entDf = dA[dA["ESTADO DATAFONO"] == "ENTREGADO"] if "ESTADO DATAFONO" in dA.columns else pd.DataFrame()
    cumple = len(entDf[entDf["CUMPLE ANS"] == "SI"]) if "CUMPLE ANS" in entDf.columns and len(entDf) else 0
    pct_oport   = round(cumple / len(entDf) * 100) if len(entDf) else 0
    pct_calidad = round((entregados - devueltos) / entregados * 100) if entregados else 100

    # Solicitudes por día (para gráfica de tendencia)
    daily: Dict[str, int] = {}
    if "FECHA DE SOLICITUD" in dA.columns:
        for v in dA["FECHA DE SOLICITUD"]:
            vs = str(v).strip()
            if not vs or vs in ("", "NAN"):
                continue
            try:
                for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y"):
                    try:
                        d_obj = datetime.strptime(vs, fmt)
                        key   = d_obj.strftime("%Y-%m-%d")
                        daily[key] = daily.get(key, 0) + 1
                        break
                    except ValueError:
                        pass
            except Exception:
                pass

    return {
        "total":           total,
        "entregados":      entregados,
        "en_transito":     en_transito,
        "en_alistamiento": en_alistamiento,
        "devueltos":       devueltos,
        "cancelados":      cancelados,
        "total_vt":        len(vt),
        "entregados_vt":   entVT,
        "total_ol":        len(ol),
        "entregados_ol":   entOL,
        "pct_entregado":   round(entregados / total * 100, 1) if total else 0,
        "pct_transito":    round(en_transito / total * 100, 1) if total else 0,
        "pct_oportunidad": pct_oport,
        "pct_calidad":     pct_calidad,
        "pct_vt":          round(entVT / len(vt) * 100) if len(vt) else 0,
        "pct_ol":          round(entOL / len(ol) * 100) if len(ol) else 0,
        "daily":           dict(sorted(daily.items())),
    }


# ══════════════════════════════════════════════════════════════════
#  SERIALIZACIÓN SAFE
# ══════════════════════════════════════════════════════════════════
def _safe_val(v: Any) -> Any:
    """Convierte cualquier valor a algo serializable en JSON."""
    if pd.isna(v) if not isinstance(v, (list, dict, bool)) else False:
        return ""
    if isinstance(v, (int, float)):
        if pd.isna(v):
            return ""
        return v
    return str(v)


def df_to_rows(df: pd.DataFrame) -> List[Dict[str, Any]]:
    """Convierte DataFrame a lista de dicts con valores seguros."""
    rows = []
    cols = list(df.columns)
    for _, row in df.iterrows():
        rows.append({c: _safe_val(row[c]) for c in cols})
    return rows


# ══════════════════════════════════════════════════════════════════
#  GENERAR data.json
# ══════════════════════════════════════════════════════════════════
def build_data_json(df: pd.DataFrame) -> Dict[str, Any]:
    """Construye el payload completo para data.json."""
    df_clean = clean_df(df)
    summary  = compute_summary(df_clean)
    rows     = df_to_rows(df_clean)

    return {
        "generado":  datetime.now().strftime("%d/%m/%Y %H:%M"),
        "filas":     len(rows),
        "columnas":  list(df_clean.columns),
        "summary":   summary,
        "rows":      rows,
    }


def write_json(payload: Dict[str, Any]) -> str:
    """Escribe data.json localmente y retorna el string JSON."""
    content = json.dumps(payload, ensure_ascii=False, separators=(",", ":"))
    OUTPUT_JSON.write_text(content, encoding="utf-8")
    size_kb = len(content.encode()) / 1024
    log.info(f"data.json escrito: {len(payload['rows'])} filas, {size_kb:.1f} KB")
    return content


# ══════════════════════════════════════════════════════════════════
#  GITHUB PUSH
# ══════════════════════════════════════════════════════════════════
def push_to_github(content: str):
    """Sube data.json al repositorio de GitHub Pages."""
    if not GITHUB_TOKEN:
        log.warning("GITHUB_TOKEN no configurado — se omite el push.")
        return

    api_url = f"https://api.github.com/repos/{GITHUB_REPO}/contents/{GITHUB_FILE_PATH}"
    headers = {
        "Authorization": f"token {GITHUB_TOKEN}",
        "Accept":        "application/vnd.github.v3+json",
    }

    # Obtener SHA actual (para update)
    sha: Optional[str] = None
    r = requests.get(api_url, headers=headers, params={"ref": GITHUB_BRANCH}, timeout=30)
    if r.status_code == 200:
        sha = r.json().get("sha")
    elif r.status_code not in (404,):
        log.warning(f"GitHub GET {r.status_code}: {r.text[:200]}")

    # Verificar si el contenido cambió
    content_b64 = base64.b64encode(content.encode("utf-8")).decode("ascii")
    if sha:
        content_hash = hashlib.sha256(content.encode()).hexdigest()[:8]
        log.info(f"Actualizando archivo en GitHub (hash: {content_hash})...")
    else:
        log.info("Creando archivo en GitHub...")

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
#  MAIN PIPELINE
# ══════════════════════════════════════════════════════════════════
def run_once(no_push: bool = False):
    start = datetime.now()
    log.info("═" * 60)
    log.info(f"  SYNC WOMPI VP — {start.strftime('%Y-%m-%d %H:%M:%S')}")
    log.info("═" * 60)

    try:
        # 1. Token
        _token_cache.get()

        # 2. Buscar archivo
        item_id = find_workbook_item_id()

        # 3. Leer hoja
        df = read_sheet(item_id)
        log.info(f"  Columnas detectadas: {list(df.columns)[:8]} ...")

        # 4. Construir JSON
        payload = build_data_json(df)

        # 5. Escribir localmente
        content = write_json(payload)

        # 6. Push a GitHub
        if not no_push:
            push_to_github(content)
        else:
            log.info("--no-push activo: se omite subida a GitHub.")

        elapsed = (datetime.now() - start).total_seconds()
        log.info(f"✅ Ciclo completado en {elapsed:.1f}s — {payload['filas']} registros")

    except Exception as e:
        log.error(f"❌ Error en ciclo: {e}")
        raise


def main():
    no_push  = "--no-push" in sys.argv
    loop     = "--loop"    in sys.argv

    if loop:
        log.info(f"Modo loop cada {INTERVAL_HOURS}h. Ctrl+C para detener.")
        while True:
            try:
                run_once(no_push=no_push)
            except Exception as e:
                log.error(f"Ciclo fallido: {e}. Reintentando en {INTERVAL_HOURS}h...")
            time.sleep(INTERVAL_HOURS * 3600)
    else:
        run_once(no_push=no_push)


if __name__ == "__main__":
    main()