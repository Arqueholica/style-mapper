"""
style_processor.py
Funciones para leer y modificar estilos en documentos .docx.
Incluye detección heurística de headings y revisión párrafo a párrafo.
"""

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.opc.exceptions import PackageNotFoundError
from docx.text.paragraph import Paragraph as DocxParagraph

# ──────────────────────────────────────────────────────────────
# Constantes del detector de headings
# Derivadas del análisis de 72 párrafos reales de documentos IFU
# ──────────────────────────────────────────────────────────────
HEADING_MAX_WORDS       = 7
HEADING_MAX_CHARS       = 120
HEADING_CONFIDENCE_HIGH = 0.90
HEADING_CONFIDENCE_MED  = 0.65
HEADING_CONFIDENCE_LOW  = 0.40

HEADING_TARGETS = {"Heading 1", "Heading 2", "Heading 3"}

STATUS_LABELS = {
    "no_change": "✅ Sin cambio",
    "auto":      "🔵 Automático",
    "review":    "🟡 Revisar",
    "unknown":   "⚠️ Sin mapeo",
}


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


def _get_all_paras(doc):
    """
    Devuelve todos los párrafos del documento en orden,
    incluyendo los que están dentro de tablas.
    """
    paras = []
    para_index = {p._element: p for p in doc.paragraphs}

    def traverse(element):
        for child in element:
            tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
            if tag == 'p':
                if child in para_index:
                    paras.append(para_index[child])
                else:
                    paras.append(DocxParagraph(child, doc))
            elif tag in ('tbl', 'tr', 'tc', 'sdt', 'sdtContent', 'body'):
                traverse(child)

    traverse(doc.element.body)
    return paras


def analyze_document(file_path, rules_dict, known_style_names):
    """
    Analiza todos los párrafos del documento y devuelve una lista de dicts
    con la información necesaria para la revisión párrafo a párrafo.

    rules_dict       : {source_style: (target_style, confidence)}
    known_style_names: set de nombres reconocidos por Oxygen

    Estados posibles:
      no_change → estilo ya reconocido por Oxygen, no hace falta cambiar
      auto      → cambio con alta confianza, se aplica sin revisión
      review    → cambio sugerido pero con confianza media o estilo sin regla
      unknown   → sin mapeo y no parece heading, requiere decisión manual
    """
    try:
        doc = Document(file_path)
    except PackageNotFoundError:
        raise ValueError("El archivo no es un .docx válido.")

    all_paras = _get_all_paras(doc)
    results   = []

    for i, para in enumerate(all_paras):
        style_name = (para.style.name if para.style else "Normal") or "Normal"
        text       = para.text.strip()
        display    = (text[:100] if text else "(vacío)")

        if style_name in known_style_names:
            suggested  = style_name
            confidence = 1.0
            status     = "no_change"

        elif style_name in rules_dict:
            target, rule_conf = rules_dict[style_name]

            if target in HEADING_TARGETS:
                likely, heur_conf, _ = is_likely_heading(para)
                if likely:
                    suggested  = target
                    confidence = heur_conf
                    status     = "auto" if heur_conf >= 0.85 else "review"
                else:
                    suggested  = style_name
                    confidence = 0.0
                    status     = "no_change"
            else:
                suggested  = target
                confidence = rule_conf
                status     = "auto" if rule_conf >= 0.80 else "review"

        else:
            # Sin regla: comprobar igualmente si parece un heading
            likely, heur_conf, _ = is_likely_heading(para)
            if likely:
                # Sugerir Heading 1 pero siempre como "revisar" (nunca automático)
                suggested  = "Heading 1"
                confidence = heur_conf
                status     = "review"
            else:
                suggested  = style_name
                confidence = 0.0
                status     = "unknown"

        results.append({
            "idx":            i,
            "text":           display,
            "original_style": style_name,
            "suggested_style":suggested,
            "final_style":    suggested,
            "confidence_pct": round(confidence * 100),
            "status":         status,
            "status_label":   STATUS_LABELS[status],
        })

    return results


def apply_paragraph_decisions(input_path, output_path, decisions):
    """
    Aplica los estilos finales decididos por el usuario, párrafo a párrafo.

    decisions: lista de dicts con {idx, final_style, original_style}

    Devuelve (n_cambios, n_errores)
    """
    doc     = Document(input_path)
    paras   = _get_all_paras(doc)
    changes = 0
    errors  = 0

    for d in decisions:
        idx            = d["idx"]
        final_style    = d["final_style"]
        original_style = d["original_style"]

        if final_style == original_style or final_style == "(sin cambio)":
            continue
        if idx >= len(paras):
            continue

        para = paras[idx]
        try:
            target_style = doc.styles[final_style]
            if target_style.type == WD_STYLE_TYPE.PARAGRAPH:
                para.style = target_style
                changes += 1
        except KeyError:
            errors += 1

    doc.save(output_path)
    return changes, errors


def extract_styles_from_docx(file_path):
    """Devuelve resumen de estilos usados en el documento (función legacy)."""
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
                "element_type": "p", "count": 0, "sample": "",
                "heading_candidates": [],
            }
        paragraph_styles[name]["count"] += 1
        if not paragraph_styles[name]["sample"] and para.text.strip():
            paragraph_styles[name]["sample"] = para.text.strip()[:120]

        likely, conf, reason = is_likely_heading(para)
        if likely:
            paragraph_styles[name]["heading_candidates"].append(
                {"text": para.text.strip()[:80], "confidence": conf, "reason": reason}
            )

        for run in para.runs:
            if run.style and run.style.name not in ("Default Paragraph Font", ""):
                rname = run.style.name
                if rname not in character_styles:
                    character_styles[rname] = {"element_type": "r", "count": 0, "sample": ""}
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
