"""
github_sync.py
Persistencia de reglas mediante un Gist privado de GitHub.

"Apaño" pensado para el plan gratuito de Render: el archivo SQLite se
pierde en cada redeploy porque el disco no es persistente. Como respaldo,
las reglas se guardan también como JSON en un Gist de GitHub, y se
restauran desde ahí cada vez que la app arranca.

Variables de entorno necesarias (configúralas en Render → Settings →
Environment):
  GITHUB_TOKEN  → Personal Access Token con permiso "gist"
  GIST_ID       → ID del Gist donde se guarda el respaldo (ver README)

Si no están configuradas, la app funciona igual pero sin persistencia
entre redeploys — is_configured() devuelve False y el resto de funciones
no hacen nada (fallan de forma silenciosa, nunca rompen la app).
"""

import json
import os

try:
    import requests
except ImportError:
    requests = None

GITHUB_TOKEN  = os.environ.get("GITHUB_TOKEN")
GIST_ID       = os.environ.get("GIST_ID")
GIST_FILENAME = "style_rules_backup.json"
GITHUB_API    = "https://api.github.com"
TIMEOUT       = 10  # segundos


def is_configured() -> bool:
    """True si hay token, gist id y la librería requests disponibles."""
    return bool(GITHUB_TOKEN and GIST_ID and requests is not None)


def _headers():
    return {
        "Authorization": f"Bearer {GITHUB_TOKEN}",
        "Accept": "application/vnd.github+json",
    }


def fetch_rules_from_gist():
    """
    Descarga las reglas guardadas en el Gist.
    Devuelve una lista de dicts (una por regla) o None si falla o no
    está configurado. Nunca lanza excepciones — el fallo es silencioso
    para no romper el arranque de la app.
    """
    if not is_configured():
        return None
    try:
        resp = requests.get(
            f"{GITHUB_API}/gists/{GIST_ID}", headers=_headers(), timeout=TIMEOUT
        )
        resp.raise_for_status()
        data = resp.json()
        file_info = data.get("files", {}).get(GIST_FILENAME)
        if not file_info:
            return None
        content = file_info.get("content", "[]")
        return json.loads(content)
    except Exception as e:
        print(f"⚠️  No se pudo leer el respaldo de GitHub: {e}")
        return None


def push_rules_to_gist(rules) -> bool:
    """
    Sube la lista completa de reglas (lista de dicts) al Gist,
    sobrescribiendo el contenido anterior por completo.
    Devuelve True si tuvo éxito, False en cualquier otro caso.
    """
    if not is_configured():
        return False
    try:
        payload = {
            "files": {
                GIST_FILENAME: {
                    "content": json.dumps(rules, indent=2, ensure_ascii=False, default=str)
                }
            }
        }
        resp = requests.patch(
            f"{GITHUB_API}/gists/{GIST_ID}", headers=_headers(),
            json=payload, timeout=TIMEOUT,
        )
        resp.raise_for_status()
        return True
    except Exception as e:
        print(f"⚠️  No se pudo guardar el respaldo en GitHub: {e}")
        return False
