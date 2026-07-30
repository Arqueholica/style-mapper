"""
database.py
Gestiona la base de datos SQLite del proyecto.
Contiene dos tablas:
  - known_styles  : estilos reconocidos por OxygenAuthor (se carga del XML)
  - style_rules   : reglas de mapeo aprendidas por el equipo

Persistencia entre redeploys de Render (plan gratuito, sin disco
persistente): las reglas se respaldan en un Gist de GitHub mediante
github_sync.py. Al arrancar, se restauran desde ahí si está configurado;
cada vez que se guarda o elimina una regla, se vuelve a subir el conjunto
completo. Si no está configurado, todo funciona igual pero sin persistencia
entre redespliegues — ver github_sync.py para las variables de entorno.
"""

import sqlite3
import os
from xml_parser import parse_oxygen_styles

try:
    import github_sync
except ImportError:
    github_sync = None

DB_PATH = "style_mapper.db"


def get_connection():
    """Devuelve una conexión a la base de datos con rows como diccionarios."""
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def initialize_database():
    """
    Crea las tablas si no existen y rellena known_styles desde el XML.
    Se llama una vez al arrancar la app.
    """
    conn = get_connection()
    cursor = conn.cursor()

    # Tabla de estilos conocidos por Oxygen (fuente de verdad)
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS known_styles (
            id                INTEGER PRIMARY KEY AUTOINCREMENT,
            element_type      TEXT    NOT NULL,
            style_name        TEXT    NOT NULL,
            target_html       TEXT    NOT NULL,
            creates_new_block BOOLEAN DEFAULT FALSE,
            UNIQUE(element_type, style_name)
        )
    """)

    # Tabla de reglas de mapeo creadas por el equipo
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS style_rules (
            id             INTEGER PRIMARY KEY AUTOINCREMENT,
            source_style   TEXT    NOT NULL,
            source_element TEXT    NOT NULL,
            target_style   TEXT    NOT NULL,
            confidence     REAL    DEFAULT 1.0,
            rule_origin    TEXT    DEFAULT 'human',
            created_by     TEXT    DEFAULT 'unknown',
            created_at     TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            times_applied  INTEGER DEFAULT 0,
            is_active      BOOLEAN DEFAULT TRUE,
            UNIQUE(source_style, source_element)
        )
    """)

    conn.commit()

    # Rellenar known_styles desde el XML (si aún no tiene datos)
    cursor.execute("SELECT COUNT(*) as n FROM known_styles")
    row = cursor.fetchone()
    if row["n"] == 0:
        styles = parse_oxygen_styles()
        for s in styles:
            cursor.execute("""
                INSERT OR IGNORE INTO known_styles
                    (element_type, style_name, target_html, creates_new_block)
                VALUES (?, ?, ?, ?)
            """, (s["element_type"], s["style_name"],
                  s["target_html"], s["creates_new_block"]))
        conn.commit()
        print(f"✅ Base de datos inicializada con {len(styles)} estilos de Oxygen.")

    conn.close()
    restore_rules_from_github()
    load_seed_rules()


# ──────────────────────────────────────────────
# Consultas sobre known_styles
# ──────────────────────────────────────────────

def get_known_styles():
    """Devuelve todos los estilos conocidos como lista de dicts."""
    conn = get_connection()
    rows = conn.execute(
        "SELECT * FROM known_styles ORDER BY style_name"
    ).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def get_known_style_names():
    """Devuelve un set con los nombres de estilos reconocidos por Oxygen."""
    return {s["style_name"] for s in get_known_styles()}


# ──────────────────────────────────────────────
# Consultas sobre style_rules
# ──────────────────────────────────────────────

def get_style_rule(source_style, source_element):
    """
    Busca si ya existe una regla para este estilo.
    Devuelve el dict de la regla o None.
    """
    conn = get_connection()
    row = conn.execute("""
        SELECT * FROM style_rules
        WHERE source_style = ? AND source_element = ? AND is_active = TRUE
    """, (source_style, source_element)).fetchone()
    conn.close()
    return dict(row) if row else None


def save_style_rule(source_style, source_element, target_style,
                    origin="human", created_by="usuario"):
    """
    Guarda o actualiza una regla de mapeo.
    Si ya existe, incrementa times_applied y actualiza el target.
    """
    conn = get_connection()
    conn.execute("""
        INSERT INTO style_rules
            (source_style, source_element, target_style, rule_origin, created_by)
        VALUES (?, ?, ?, ?, ?)
        ON CONFLICT(source_style, source_element) DO UPDATE SET
            target_style  = excluded.target_style,
            rule_origin   = excluded.rule_origin,
            times_applied = times_applied + 1
    """, (source_style, source_element, target_style, origin, created_by))
    conn.commit()
    conn.close()
    _sync_rules_to_github()


def get_all_rules():
    """Devuelve todas las reglas activas, ordenadas por uso."""
    conn = get_connection()
    rows = conn.execute("""
        SELECT * FROM style_rules
        WHERE is_active = TRUE
        ORDER BY times_applied DESC, created_at DESC
    """).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def delete_rule(rule_id):
    """Desactiva una regla (no la borra físicamente)."""
    conn = get_connection()
    conn.execute(
        "UPDATE style_rules SET is_active = FALSE WHERE id = ?", (rule_id,)
    )
    conn.commit()
    conn.close()
    _sync_rules_to_github()


# ──────────────────────────────────────────────
# Reglas iniciales confirmadas con documentos reales
# ──────────────────────────────────────────────

SEED_RULES = [
    # ── Confirmadas con análisis de archivos InDesign (.idml) ──────────────
    # Convención "legacy" numerada, confirmada con M-1917 (Light-Structure)
    ("103 Heading 1-Rev",    "p", "Title",      0.85, "seed_confirmed"),
    ("103.5 Heading 1",      "p", "Heading 1",  0.90, "seed_confirmed"),
    ("104 Heading 2",        "p", "Heading 2",  0.90, "seed_confirmed"),
    ("105 Heading 3",        "p", "Heading 3",  0.90, "seed_confirmed"),
    ("106 Heading 4",        "p", "Heading 4",  0.85, "seed_confirmed"),
    ("200 Body Text",        "p", "Body Text",  0.90, "seed_confirmed"),
    ("400 Table Text",       "p", "Table Paragraph", 0.85, "seed_confirmed"),
    ("402 Table Header Rev", "p", "Table Header",    0.85, "seed_confirmed"),
    # Nombres propios encontrados en M-1291 (sin usar la convención numerada)
    ("Subhead",              "p", "Heading 2",  0.70, "seed_suggested"),
    ("Subheads P1",          "p", "Heading 2",  0.65, "seed_suggested"),
    ("Body Copy P1",         "p", "Body Text",  0.75, "seed_suggested"),
    ("Body",                 "p", "Body Text",  0.80, "seed_confirmed"),
    # Catálogo "moderno" (New Paragraph Styles) — confirmado en CCMS/N8
    ("Table Body",           "p", "Table Paragraph",   0.85, "seed_confirmed"),
    ("Table Heading Left aligned", "p", "Table Header", 0.80, "seed_confirmed"),

    # Resto de la convención numerada legacy, confirmadas con M-1917
    ("203 Procedure",             "p", "Body Text",     0.80, "seed_confirmed"),
    ("201.5 Checkbox",            "p", "List Paragraph", 0.80, "seed_confirmed"),
    ("201 Bullet",                "p", "List Paragraph", 0.85, "seed_confirmed"),
    ("201.3 Indented Bullet",     "p", "List Paragraph", 0.75, "seed_suggested"),
    ("206 Warning Signal Word",   "p", "Body Text",     0.70, "seed_suggested"),
    ("207 Note",                  "p", "Body Text",     0.70, "seed_suggested"),
    ("205 Footnote",              "p", "footnote text", 0.75, "seed_confirmed"),
    ("300 Caption",               "p", "Caption",       0.80, "seed_confirmed"),
    ("504 Contact Info",          "p", "Body Text",     0.65, "seed_suggested"),
    ("503 Copyright, Code, Patents", "p", "footnote text", 0.70, "seed_suggested"),

    # Catálogo "moderno" (New Paragraph Styles) — resto de nombres confirmados
    ("Figure number callout",     "p", "Caption",        0.75, "seed_confirmed"),
    ("Table - Bullets",           "p", "Table Paragraph", 0.75, "seed_suggested"),
    ("Table - Numbered start",    "p", "Table Paragraph", 0.75, "seed_suggested"),
    ("Table - Numbered cont",     "p", "Table Paragraph", 0.75, "seed_suggested"),
    ("Footer & page#",            "p", "footnote text",  0.60, "seed_suggested"),
    ("Note-Tip-Caution-Warning",  "p", "Body Text",      0.70, "seed_suggested"),
    ("NTWC within body list",     "p", "Body Text",      0.65, "seed_suggested"),
    ("Body Copy - Numbered start","p", "List Paragraph", 0.75, "seed_suggested"),
    ("Body Copy - Numbered cont", "p", "List Paragraph", 0.75, "seed_suggested"),
    ("For Professionals",         "p", "Body Text",      0.60, "seed_suggested"),
    ("For Recipients",            "p", "Body Text",      0.60, "seed_suggested"),

    # Plantilla de brochure de casos de estudio (World Class Leaders)
    ("Project Name",              "p", "Heading 2",      0.75, "seed_suggested"),
    ("Project Location_2",        "p", "Body Text",      0.65, "seed_suggested"),
    ("Project Bullet Points",     "p", "List Paragraph", 0.70, "seed_suggested"),

    # Restantes confirmados en el barrido final de los 14 archivos
    ("Body Copy -  Plain list",   "p", "List Paragraph", 0.75, "seed_confirmed"),
    ("NTCW bullet",               "p", "List Paragraph", 0.70, "seed_suggested"),
    ("Footnotes",                 "p", "footnote text",  0.75, "seed_confirmed"),
    ("Figure Caption",            "p", "Caption",        0.80, "seed_confirmed"),
    ("Blue Subhead",              "p", "Heading 2",      0.65, "seed_suggested"),
    ("Quote new",                 "p", "Quote",          0.65, "seed_suggested"),
    ("New Byline",                "p", "Body Text",      0.60, "seed_suggested"),
    ("Main Head",                 "p", "Heading 1",      0.70, "seed_suggested"),
    ("209 Half Col Tab w Dot Leader", "p", "Body Text",  0.60, "seed_suggested"),
    ("Numbered Paragraphs",       "p", "List Paragraph", 0.70, "seed_suggested"),
    ("Bold copy P1",              "p", "Body Text",      0.60, "seed_suggested"),
    ("Lower Case Paragrphs",      "p", "Body Text",      0.60, "seed_suggested"),
    # ── Confirmadas con 25 pares de documentos (2,760 cambios analizados) ──

    # Estilos desconocidos que mapean a Heading 1
    # (is_likely_heading filtra: solo si el párrafo parece título de sección)
    ("Default",             "p", "Heading 1", 0.95, "seed_confirmed"),
    ("paragraph",           "p", "Heading 1", 0.90, "seed_confirmed"),
    ("Body Text",           "p", "Heading 1", 0.85, "seed_confirmed"),
    ("Title",               "p", "Heading 1", 0.80, "seed_confirmed"),
    ("PARAGRAPH",           "p", "Heading 1", 0.75, "seed_confirmed"),

    # Estilos que mapean a Heading 2
    ("Heading 1 No Number", "p", "Heading 2", 0.85, "seed_confirmed"),
    ("Table Caption",       "p", "Heading 2", 0.70, "seed_confirmed"),
    ("Heading 9",           "p", "Heading 2", 0.85, "seed_confirmed"),

    # Heading 7/8 → Body Text (documentos Light-Structure y SSCP)
    ("Heading 7",           "p", "Body Text", 0.80, "seed_confirmed"),
    ("Heading 8",           "p", "Body Text", 0.75, "seed_confirmed"),

    # Estilos de tabla → Table Paragraph
    # (la lógica de tabla en analyze_document cubre Normal→Table Paragraph)
    ("Table Text",          "p", "Table Paragraph", 0.88, "seed_confirmed"),
    ("Table Header 2",      "p", "Table Header",    0.85, "seed_confirmed"),
    ("Table Header Text",   "p", "Table Header",    0.75, "seed_confirmed"),
    ("Numbered List 1",     "p", "Table Paragraph", 0.75, "seed_confirmed"),

    # Estilos FrameMaker
    ("pf0",                 "p", "Body Text", 0.80, "seed_confirmed"),
    ("pf1",                 "p", "Body Text", 0.80, "seed_suggested"),

    # Otros estilos personalizados
    ("Style1",              "p", "Normal",    0.70, "seed_confirmed"),
    ("Body Text 3 FDA",     "p", "Body Text", 0.80, "seed_suggested"),
    ("Normal (Web)",        "p", "Normal",    0.85, "seed_confirmed"),
    ("List Bullet 2",       "p", "List Paragraph", 0.75, "seed_suggested"),

    # Estilo de carácter
    ("eop",                 "r", "Default Paragraph Font", 0.50, "seed_suggested"),
]


def restore_rules_from_github():
    """
    Restaura las reglas guardadas en el Gist de respaldo (si está
    configurado). Se llama una vez al arrancar, antes de cargar las
    reglas semilla, para que las reglas del equipo persistan entre
    redespliegues de Render.
    """
    if github_sync is None or not github_sync.is_configured():
        return

    remote_rules = github_sync.fetch_rules_from_gist()
    if not remote_rules:
        return

    conn = get_connection()
    for r in remote_rules:
        try:
            conn.execute("""
                INSERT INTO style_rules
                    (source_style, source_element, target_style, confidence,
                     rule_origin, created_by, created_at, times_applied, is_active)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(source_style, source_element) DO UPDATE SET
                    target_style  = excluded.target_style,
                    confidence    = excluded.confidence,
                    rule_origin   = excluded.rule_origin,
                    times_applied = excluded.times_applied,
                    is_active     = excluded.is_active
            """, (
                r.get("source_style"), r.get("source_element"), r.get("target_style"),
                r.get("confidence", 1.0), r.get("rule_origin", "human"),
                r.get("created_by", "usuario"), r.get("created_at"),
                r.get("times_applied", 0), r.get("is_active", True),
            ))
        except Exception as e:
            print(f"⚠️  Regla del respaldo omitida por error: {e}")
    conn.commit()
    conn.close()
    print(f"✅ {len(remote_rules)} reglas restauradas desde el respaldo de GitHub.")


def _sync_rules_to_github():
    """
    Sube el conjunto completo de reglas activas al Gist de respaldo.
    Se llama tras cada guardado o borrado. Falla en silencio si no está
    configurado — nunca bloquea la operación local.
    """
    if github_sync is None or not github_sync.is_configured():
        return
    github_sync.push_rules_to_gist(get_all_rules())


def load_seed_rules():
    """
    Carga las reglas iniciales en la base de datos con INSERT OR IGNORE,
    por lo que es seguro llamarla siempre (no duplica ni sobrescribe
    reglas ya existentes, vengan del uso del equipo o restauradas desde
    GitHub). Esto garantiza que cualquier regla semilla nueva añadida en
    el código se incorpore aunque el respaldo de GitHub ya tenga datos.
    """
    conn = get_connection()
    cursor = conn.cursor()
    for source_style, source_element, target_style, confidence, origin in SEED_RULES:
        cursor.execute("""
            INSERT OR IGNORE INTO style_rules
                (source_style, source_element, target_style,
                 confidence, rule_origin, created_by)
            VALUES (?, ?, ?, ?, ?, ?)
        """, (source_style, source_element, target_style,
              confidence, origin, "sistema"))
    conn.commit()
    conn.close()
