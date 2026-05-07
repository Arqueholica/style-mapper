"""
style_processor.py
Funciones para leer y modificar estilos en documentos .docx.
Usa la librería python-docx.
"""

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.opc.exceptions import PackageNotFoundError


def extract_styles_from_docx(file_path):
    """
    Lee un .docx y devuelve todos los estilos que usa,
    con su tipo, cuántas veces aparece y un texto de ejemplo.

    Devuelve un dict con este formato:
    {
        "Heading 1": {"element_type": "p", "count": 3, "sample": "Introducción"},
        "Strong":    {"element_type": "r", "count": 7, "sample": "importante"},
        ...
    }
    """
    try:
        doc = Document(file_path)
    except PackageNotFoundError:
        raise ValueError("El archivo no es un .docx válido.")

    paragraph_styles = {}
    character_styles = {}

    def _add_para(para):
        """Registra el estilo de un párrafo y sus runs."""
        name = (para.style.name if para.style else "Normal") or "Normal"
        if name not in paragraph_styles:
            paragraph_styles[name] = {
                "element_type": "p",
                "count": 0,
                "sample": (para.text[:120] if para.text.strip() else "(párrafo vacío)"),
            }
        paragraph_styles[name]["count"] += 1

        # Estilos de carácter en los runs
        for run in para.runs:
            if run.style and run.style.name not in ("Default Paragraph Font", ""):
                rname = run.style.name
                if rname not in character_styles:
                    character_styles[rname] = {
                        "element_type": "r",
                        "count": 0,
                        "sample": (run.text[:120] if run.text.strip() else "(run vacío)"),
                    }
                character_styles[rname]["count"] += 1

    # Párrafos del cuerpo principal
    for para in doc.paragraphs:
        _add_para(para)

    # Párrafos dentro de tablas
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    _add_para(para)

    # Unir ambos dicts (los estilos de carácter van después)
    all_styles = {}
    all_styles.update(paragraph_styles)
    # Si un nombre de estilo existe como párrafo Y como carácter,
    # el de carácter se añade con sufijo " (r)" para distinguirlo
    for name, info in character_styles.items():
        key = name if name not in all_styles else f"{name} (carácter)"
        all_styles[key] = info

    return all_styles


def apply_style_mappings(input_path, output_path, mappings):
    """
    Crea una copia del documento con los estilos renombrados según `mappings`.

    mappings: dict  {nombre_estilo_origen: nombre_estilo_destino}
              Ejemplo: {"Pa9": "Body Text", "Table Header Text": "Heading 2"}

    Devuelve el número de cambios realizados.
    """
    doc = Document(input_path)
    changes = 0

    def _restyle_para(para):
        nonlocal changes

        # ── Estilo de párrafo ─────────────────────────────────────────
        if para.style and para.style.name in mappings:
            target = mappings[para.style.name]
            try:
                target_style = doc.styles[target]
                # Solo asignar si el estilo destino es de tipo PÁRRAFO
                if target_style.type == WD_STYLE_TYPE.PARAGRAPH:
                    para.style = target_style
                    changes += 1
            except KeyError:
                pass  # El estilo no existe en el documento, se ignora

        # ── Estilos de carácter en los runs ───────────────────────────
        for run in para.runs:
            if not run.style or run.style.name not in mappings:
                continue
            target = mappings[run.style.name]
            try:
                target_style = doc.styles[target]
                if target_style.type == WD_STYLE_TYPE.CHARACTER:
                    # El destino es un estilo de carácter → asignar directo
                    run.style = target_style
                    changes += 1
                elif target_style.type == WD_STYLE_TYPE.PARAGRAPH:
                    # El destino es un estilo de párrafo → no se puede
                    # asignar a un run; en su lugar eliminamos el estilo
                    # de carácter para que herede el del párrafo
                    run.style = doc.styles["Default Paragraph Font"]
                    changes += 1
            except KeyError:
                pass  # El estilo no existe en el documento, se ignora

    for para in doc.paragraphs:
        _restyle_para(para)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    _restyle_para(para)

    doc.save(output_path)
    return changes
