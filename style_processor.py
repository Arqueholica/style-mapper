"""
style_processor.py
Funciones para leer y modificar estilos en documentos .docx.
Incluye detección heurística de headings basada en análisis de
72 párrafos reales de documentos IFU.
"""

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.opc.exceptions import PackageNotFoundError

# ──────────────────────────────────────────────────────────────
# Constantes del detector de headings
# Derivadas del análisis de dos pares de documentos reales:
#   · Todos los headings convertidos tenían ≤ 7 palabras
#   · Los no-headings tenían promedio de 48 palabras
#   · Los headings eran bold y/o mayúsculas
# ──────────────────────────────────────────────────────────────
HEADING_MAX_WORDS        = 7
HEADING_MAX_CHARS        = 120
HEADING_CONFIDENCE_HIGH  = 0.90
HEADING_CONFIDENCE_MED   = 0.65
HEADING_CONFIDENCE_LOW   = 0.40


def is_likely_heading(para):
    """
    Evalúa si un párrafo parece un título de sección.
    Devuelve (es_probable, confianza_0_a_1, motivo_string)
    """
    text = para.text.strip()
    if not text:
        return False, 0.0, "párrafo vacío"

    word_count = len(text.split())
    if word_count > HEADING_MAX_WORDS or len(text) > HEADING_MAX_CHARS:
        return False, 0.0, f"demasiado largo ({word_count} palabras)"

    is_bold  = any(run.bold for run in para.runs if run.text.strip())
    is_upper = text.isupper()
    is_title = text.istitle()

    if is_bold and (is_upper or word_count <= 3):
        return True, HEADING_CONFIDENCE_HIGH, "negrita + mayúsculas/corto"
    if is_bold and is_title:
        return True, HEADING_CONFIDENCE_MED, "negrita + capitalizado"
    if is_bold:
        return True, HEADING_CONFIDENCE_MED, "negrita"
    if is_upper and word_count <= 5:
        return True, HEADING_CONFIDENCE_LOW, "mayúsculas (sin negrita)"

    return False, 0.0, "no cumple criterios de heading"


def extract_styles_from_docx(file_path):
    """
    Lee un .docx y devuelve todos los estilos usados.
    Para cada estilo de párrafo incluye heading_candidates:
    lista de párrafos que podrían ser Heading 1.
    """
    try:
        doc = Document(file_path)
    except PackageNotFoundError:
        raise ValueError("El archivo no es un .docx válido.")

    paragraph_styles = {}
    character_styles = {}

    def _add_para(para):
        name = (para.style.name if para.style else "Normal") or "Normal"

        if name not in paragraph_styles:
            paragraph_styles[name] = {
                "element_type": "p",
                "count": 0,
                "sample": "",
                "heading_candidates": [],
            }

        paragraph_styles[name]["count"] += 1
        if not paragraph_styles[name]["sample"] and para.text.strip():
            paragraph_styles[name]["sample"] = para.text.strip()[:120]

        likely, confidence, reason = is_likely_heading(para)
        if likely:
            paragraph_styles[name]["heading_candidates"].append({
                "text":       para.text.strip()[:80],
                "confidence": confidence,
                "reason":     reason,
            })

        for run in para.runs:
            if run.style and run.style.name not in ("Default Paragraph Font", ""):
                rname = run.style.name
                if rname not in character_styles:
                    character_styles[rname] = {
                        "element_type": "r",
                        "count": 0,
                        "sample": "",
                    }
                character_styles[rname]["count"] += 1
                if not character_styles[rname]["sample"] and run.text.strip():
                    character_styles[rname]["sample"] = run.text.strip()[:120]

    for para in doc.paragraphs:
        _add_para(para)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    _add_para(para)

    all_styles = {}
    all_styles.update(paragraph_styles)
    for name, info in character_styles.items():
        key = name if name not in all_styles else f"{name} (carácter)"
        all_styles[key] = info

    return all_styles


def apply_style_mappings(input_path, output_path, mappings):
    """
    Aplica mapeos de estilos al documento.
    Para mapeos hacia Heading 1/2/3, aplica la heurística:
    solo cambia párrafos que parecen títulos de sección.

    Devuelve (n_cambios, n_omitidos_por_heuristica)
    """
    doc = Document(input_path)
    changes = 0
    skipped = 0
    HEADING_TARGETS = {"Heading 1", "Heading 2", "Heading 3"}

    def _restyle_para(para):
        nonlocal changes, skipped

        if para.style and para.style.name in mappings:
            target = mappings[para.style.name]
            try:
                target_style = doc.styles[target]
                if target_style.type != WD_STYLE_TYPE.PARAGRAPH:
                    pass
                elif target in HEADING_TARGETS:
                    likely, _, _ = is_likely_heading(para)
                    if likely:
                        para.style = target_style
                        changes += 1
                    else:
                        skipped += 1
                else:
                    para.style = target_style
                    changes += 1
            except KeyError:
                pass

        for run in para.runs:
            if not run.style or run.style.name not in mappings:
                continue
            target = mappings[run.style.name]
            try:
                target_style = doc.styles[target]
                if target_style.type == WD_STYLE_TYPE.CHARACTER:
                    run.style = target_style
                    changes += 1
                elif target_style.type == WD_STYLE_TYPE.PARAGRAPH:
                    run.style = doc.styles["Default Paragraph Font"]
                    changes += 1
            except KeyError:
                pass

    for para in doc.paragraphs:
        _restyle_para(para)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    _restyle_para(para)

    doc.save(output_path)
    return changes, skipped
