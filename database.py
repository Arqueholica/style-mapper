"""
database.py
Gestiona la base de datos SQLite del proyecto.
Contiene dos tablas:
  - known_styles  : estilos reconocidos por OxygenAuthor (se carga del XML)
  - style_rules   : reglas de mapeo aprendidas por el equipo
"""

import sqlite3
import os
from xml_parser import parse_oxygen_styles

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


# ──────────────────────────────────────────────
# Reglas iniciales confirmadas con documentos reales
# ──────────────────────────────────────────────

SEED_RULES = [
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


def load_seed_rules():
    """
    Carga las reglas iniciales en la base de datos.
    Solo inserta las que no existen todavía (INSERT OR IGNORE).
    Se llama una vez desde initialize_database().
    """
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT COUNT(*) as n FROM style_rules")
    if cursor.fetchone()["n"] == 0:
        for source_style, source_element, target_style, confidence, origin in SEED_RULES:
            cursor.execute("""
                INSERT OR IGNORE INTO style_rules
                    (source_style, source_element, target_style,
                     confidence, rule_origin, created_by)
                VALUES (?, ?, ?, ?, ?, ?)
            """, (source_style, source_element, target_style,
                  confidence, origin, "sistema"))
        conn.commit()
        print(f"✅ {len(SEED_RULES)} reglas iniciales cargadas.")
    conn.close()
