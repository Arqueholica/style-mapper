"""
style_processor.py
Funciones para leer y modificar estilos en documentos .docx.
Heurísticas validadas con 12 pares de documentos reales (1,305 cambios analizados).
"""

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.opc.exceptions import PackageNotFoundError
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph as DocxParagraph

# ── Constantes ────────────────────────────────────────────────────────────────
HEADING_MAX_WORDS       = 7
HEADING_MAX_CHARS       = 120
HEADING_CONFIDENCE_HIGH = 0.90
HEADING_CONFIDENCE_MED  = 0.65
HEADING_CONFIDENCE_LOW  = 0.40

# Estilos que pueden contener headings mal aplicados (no bold = probablemente Normal)
DEMOTABLE_HEADINGS = {"Heading 1", "Heading 2", "Heading 3", "Heading 4",
                      "Heading 5", "Heading 6", "Heading 7", "Title"}

# Heading targets para los que aplica la heurística
HEADING_TARGETS = {"Heading 1", "Heading 2", "Heading 3"}

# Estilos no explícitos en el XML de Oxygen pero aceptables como destino.
# Cuando un párrafo tiene estos estilos y no parece heading ni lista,
# se deja sin cambio en lugar de marcarse como "unknown".
TOLERATED_STYLES = {
    "Normal",       # Estilo por defecto de Word — Oxygen lo procesa como body text
    "annotation text",  # Estilo de anotación — se deja como está
    "toc 2",        # Tabla de contenidos nivel 2 — se deja como está
}

STATUS_LABELS = {
    "no_change": "✅ Sin cambio",
    "auto":      "🔵 Automático",
    "review":    "🟡 Revisar",
    "unknown":   "⚠️ Sin mapeo",
}


# ── Funciones de detección ────────────────────────────────────────────────────

def _is_bold(para) -> bool:
    """
    True si al menos un run con texto tiene bold EXPLÍCITO (run.bold is True).
    Usamos bold explícito intencionalmente para is_misapplied_heading:
    los headings mal aplicados (ej. body text con estilo Heading en ICF)
    nunca tienen bold explícito — solo los headings reales lo tienen.
    """
    return any(run.bold is True for run in para.runs if run.text.strip())


def _word_count(para) -> int:
    return len(para.text.split())


def has_list_formatting(para) -> bool:
    """
    True si el párrafo tiene numeración/viñeta en el XML (numPr),
    independientemente del estilo asignado.
    Identifica párrafos que deberían ser List Paragraph aunque tengan otro estilo.
    """
    pPr = para._element.find(qn('w:pPr'))
    if pPr is not None:
        return pPr.find(qn('w:numPr')) is not None
    return False


def is_likely_heading(para) -> tuple:
    """
    Evalúa si un párrafo parece un título de sección (candidato a Heading 1/2).
    Devuelve (es_probable, confianza_0_a_1, motivo_string).

    Criterio validado con 101 casos reales Normal→Heading 1:
      - ≤ 7 palabras
      - Bold y/o mayúsculas
    """
    text = para.text.strip()
    if not text:
        return False, 0.0, "párrafo vacío"

    wc = _word_count(para)
    if wc > HEADING_MAX_WORDS or len(text) > HEADING_MAX_CHARS:
        return False, 0.0, f"demasiado largo ({wc} palabras)"

    bold    = _is_bold(para)
    upper   = text.isupper()
    title   = text.istitle()

    if bold and (upper or wc <= 3):
        return True, HEADING_CONFIDENCE_HIGH, "negrita + mayúsculas/corto"
    if bold and title:
        return True, HEADING_CONFIDENCE_MED,  "negrita + capitalizado"
    if bold:
        return True, HEADING_CONFIDENCE_MED,  "negrita"
    if upper and wc <= 5:
        return True, HEADING_CONFIDENCE_LOW,  "mayúsculas (sin negrita)"

    return False, 0.0, "no cumple criterios de heading"


def is_misapplied_heading(para) -> tuple:
    """
    Detecta headings aplicados por error: tienen estilo Heading pero no están
    en negrita, lo que indica que son cuerpo de texto o párrafos vacíos.

    Validado con 842 casos reales (Heading 1/2 → Normal):
      - 0% de los casos que se convirtieron a Normal estaban en negrita.

    Devuelve (es_mal_aplicado, confianza, motivo).
    """
    text  = para.text.strip()
    bold  = _is_bold(para)
    empty = not text

    if bold:
        # Tiene negrita → probablemente es un heading real
        return False, 0.0, "tiene negrita — heading correcto"

    if empty:
        return True, 0.85, "heading vacío sin negrita"

    wc = _word_count(para)
    if wc > 10:
        return True, 0.90, f"heading sin negrita y largo ({wc} palabras)"

    # Corto pero sin negrita: confianza media — depende del contexto
    return True, 0.65, f"heading sin negrita ({wc} palabras)"


# ── Traversal de párrafos ─────────────────────────────────────────────────────

def _get_all_paras(doc):
    """
    Devuelve todos los párrafos en orden de documento,
    incluyendo los que están dentro de tablas.
    """
    paras      = []
    para_index = {p._element: p for p in doc.paragraphs}

    def traverse(element):
        for child in element:
            tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
            if tag == 'p':
                paras.append(
                    para_index.get(child) or DocxParagraph(child, doc)
                )
            elif tag in ('tbl', 'tr', 'tc', 'sdt', 'sdtContent', 'body'):
                traverse(child)

    traverse(doc.element.body)
    return paras


# ── Análisis principal ────────────────────────────────────────────────────────

def analyze_document(file_path, rules_dict, known_style_names):
    """
    Analiza todos los párrafos y devuelve lista de dicts para la revisión.

    rules_dict        : {source_style: (target_style, confidence)}
    known_style_names : set de nombres reconocidos por Oxygen

    Lógica de clasificación (en orden de prioridad):

    1. ¿Tiene list formatting (numPr en XML)?
       → Sugerir List Paragraph si no lo tiene ya

    2. ¿Es un heading conocido (H1-H7, Title) SIN negrita?
       → Sugerir Normal para revisión (heading mal aplicado)

    3. ¿Está en known_style_names?
       → Sin cambio

    4. ¿Tiene regla en rules_dict?
       → Aplicar regla (con heurística si el target es Heading)

    5. ¿No tiene regla pero parece heading?
       → Sugerir Heading 1 para revisión

    6. Sin mapeo
       → Marcar como unknown
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
        display    = text[:100] if text else "(vacío)"

        # ── 1. Heading mal aplicado (sin negrita) ─────────────────────
        if style_name in DEMOTABLE_HEADINGS:
            misapplied, conf, reason = is_misapplied_heading(para)
            if misapplied:
                results.append(_make_result(
                    i, display, style_name,
                    suggested="Normal",
                    confidence=conf,
                    status="review",
                    note=reason,
                ))
                continue
            # Heading con negrita → sin cambio
            results.append(_make_result(
                i, display, style_name,
                suggested=style_name,
                confidence=1.0,
                status="no_change",
            ))
            continue

        # ── 3. Estilo reconocido por Oxygen ───────────────────────────
        if style_name in known_style_names:
            results.append(_make_result(
                i, display, style_name,
                suggested=style_name,
                confidence=1.0,
                status="no_change",
            ))
            continue

        # ── 4. Regla guardada ─────────────────────────────────────────
        if style_name in rules_dict:
            target, rule_conf = rules_dict[style_name]

            if target in HEADING_TARGETS:
                likely, heur_conf, reason = is_likely_heading(para)
                if likely:
                    results.append(_make_result(
                        i, display, style_name,
                        suggested=target,
                        confidence=heur_conf,
                        status="auto" if heur_conf >= 0.85 else "review",
                        note=reason,
                    ))
                else:
                    # Tiene regla de heading pero no parece heading → sin cambio
                    results.append(_make_result(
                        i, display, style_name,
                        suggested=style_name,
                        confidence=0.0,
                        status="no_change",
                    ))
            else:
                results.append(_make_result(
                    i, display, style_name,
                    suggested=target,
                    confidence=rule_conf,
                    status="auto" if rule_conf >= 0.80 else "review",
                ))
            continue

        # ── 5. Sin regla — ¿parece heading? ──────────────────────────
        likely, heur_conf, reason = is_likely_heading(para)
        if likely:
            results.append(_make_result(
                i, display, style_name,
                suggested="Heading 1",
                confidence=heur_conf,
                status="review",   # Siempre revisar si no hay regla
                note=reason,
            ))
            continue

        # ── 6. Estilo tolerable — dejar sin cambio ────────────────────
        if style_name in TOLERATED_STYLES:
            results.append(_make_result(
                i, display, style_name,
                suggested=style_name,
                confidence=1.0,
                status="no_change",
            ))
            continue

        # ── 7. Sin mapeo ──────────────────────────────────────────────
        results.append(_make_result(
            i, display, style_name,
            suggested=style_name,
            confidence=0.0,
            status="unknown",
        ))

    return results


def _make_result(idx, display, original_style,
                 suggested, confidence, status, note=""):
    return {
        "idx":            idx,
        "text":           display,
        "original_style": original_style,
        "suggested_style":suggested,
        "final_style":    suggested,
        "confidence_pct": round(confidence * 100),
        "status":         status,
        "status_label":   STATUS_LABELS[status],
        "note":           note,
    }


# ── Aplicar decisiones ────────────────────────────────────────────────────────

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

        if final_style in (original_style, "(sin cambio)") or idx >= len(paras):
            continue

        try:
            target_style = doc.styles[final_style]
            if target_style.type == WD_STYLE_TYPE.PARAGRAPH:
                paras[idx].style = target_style
                changes += 1
        except KeyError:
            errors += 1

    doc.save(output_path)
    return changes, errors


# ── Función legacy ────────────────────────────────────────────────────────────

def extract_styles_from_docx(file_path):
    """Devuelve resumen de estilos usados (para compatibilidad)."""
    try:
        doc = Document(file_path)
    except PackageNotFoundError:
        raise ValueError("El archivo no es un .docx válido.")

    styles = {}
    def _add(para):
        name = (para.style.name if para.style else "Normal") or "Normal"
        if name not in styles:
            styles[name] = {"element_type": "p", "count": 0,
                            "sample": "", "heading_candidates": []}
        styles[name]["count"] += 1
        if not styles[name]["sample"] and para.text.strip():
            styles[name]["sample"] = para.text.strip()[:120]
        likely, conf, reason = is_likely_heading(para)
        if likely:
            styles[name]["heading_candidates"].append(
                {"text": para.text.strip()[:80], "confidence": conf, "reason": reason}
            )

    for para in doc.paragraphs:
        _add(para)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    _add(para)
    return styles
