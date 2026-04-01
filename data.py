#!/usr/bin/env python3
"""
╔══════════════════════════════════════════════════════════════════╗
║   sync_wompi_vp.py — Extractor SharePoint → data.json → GitHub  ║
║   LINEACOM · Dashboard Tracking VP Wompi                         ║
║                                                                  ║
║   Responsabilidad: SOLO extrae los datos crudos del Excel        ║
║   y los serializa en data.json. Todos los cálculos de KPIs       ║
║   se realizan en el frontend (dashboard.js).                     ║
║                                                                  ║
║   Uso:                                                           ║
║     python data.py              # ejecutar una vez               ║
║     python data.py --loop       # cada N horas                   ║
║     python data.py --no-push    # genera JSON sin subir          ║
╚══════════════════════════════════════════════════════════════════╝
"""

import os
import sys
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
GITHUB_TOKEN     = os.getenv("GITHUB_TOKEN",     "")
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
SEARCH_KEYWORDS = ["Comodato", "VP", "Operaci", "Wompi"]


def find_workbook_item_id() -> str:
    """Busca el archivo Excel de Operación VP_Comodato en SharePoint."""
    token   = _token_cache.get()
    headers = {"Authorization": f"Bearer {token}"}

    for kw in SEARCH_KEYWORDS:
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
#  LIMPIEZA BÁSICA — solo normaliza texto y fechas, SIN calcular KPIs
# ══════════════════════════════════════════════════════════════════
DATE_COLS = [
    "FECHA DE SOLICITUD",
    "FECHA LIMITE DE ENTREGA",
    "FECHA ENTREGA AL COMERCIO",
    "FECHA VISITA TECNICA",
    "FECHA DE ENTREGA",
]


def _parse_date_str(val: Any) -> str:
    """Normaliza cualquier valor de fecha a dd/mm/yyyy o ''."""
    if pd.isna(val) if not isinstance(val, (list, dict, bool)) else False:
        return ""
    s = str(val).strip()
    if s in ("", "nan", "NaN", "None"):
        return ""
    # Excel serial number
    if s.replace(".", "", 1).isdigit():
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

    # Eliminar filas completamente vacías
    df = df.dropna(how="all")
    df = df[~(df == "").all(axis=1)]

    return df


# ══════════════════════════════════════════════════════════════════
#  SERIALIZACIÓN SAFE
# ══════════════════════════════════════════════════════════════════
def _safe_val(v: Any) -> Any:
    if isinstance(v, (list, dict, bool)):
        return v
    try:
        if pd.isna(v):
            return ""
    except Exception:
        pass
    if isinstance(v, (int, float)):
        return v
    return str(v)


def df_to_rows(df: pd.DataFrame) -> List[Dict[str, Any]]:
    rows = []
    cols = list(df.columns)
    for _, row in df.iterrows():
        rows.append({c: _safe_val(row[c]) for c in cols})
    return rows


# ══════════════════════════════════════════════════════════════════
#  GENERAR data.json  — SOLO datos crudos, sin summary pre-calculado
# ══════════════════════════════════════════════════════════════════
def build_data_json(df: pd.DataFrame) -> Dict[str, Any]:
    """
    Construye el payload mínimo para data.json.
    NO pre-calcula KPIs — el frontend (dashboard.js) se encarga de todo.
    """
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
        _token_cache.get()
        item_id = find_workbook_item_id()
        df      = read_sheet(item_id)
        log.info(f"  Columnas detectadas: {list(df.columns)[:8]} ...")

        payload = build_data_json(df)
        content = write_json(payload)

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
    no_push = "--no-push" in sys.argv
    loop    = "--loop"    in sys.argv

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