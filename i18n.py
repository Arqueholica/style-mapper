"""
i18n.py
Sistema de traducción ES/EN para Style Mapper.
El idioma activo se guarda en st.session_state.language ('es' por defecto).
"""

import streamlit as st

# ── Textos de la interfaz principal ────────────────────────────────────────
TEXT = {
    "es": {
        # Cabecera / navegación
        "tagline": "Style Mapper · Preparación de documentos Word para conversión a DITA",
        "nav_label": "Ir a",
        "nav_procesar": "🏠 Procesar documento",
        "nav_reglas": "📋 Reglas guardadas",
        "nav_estilos": "ℹ️ Estilos de Oxygen",
        "help_open": "❓ Ayuda",
        "help_close": "✖️ Cerrar",
        "lang_toggle_label": "🌐 English",

        # Página 1 — Procesar documento
        "page1_title": "📄 Style Mapper para DITA",
        "page1_caption": "Prepara tus documentos Word para la conversión a DITA en OxygenAuthor.",
        "step1_header": "Paso 1 · Sube tu documento",
        "step1_caption": "Acepta archivos .docx. Si el documento viene de Word en otro idioma, "
                         "los estilos estándar se detectan y renombran automáticamente.",
        "step1_uploader_label": "Selecciona un archivo .docx",
        "step1_info_no_file": "Sube un archivo .docx para empezar.",
        "step1_success": "cargado correctamente.",

        "lang_norm_expander": "Se normalizaron {n} nombre(s) de estilo (documento en otro idioma)",
        "lang_norm_caption": "Estos estilos tenían el nombre traducido en Word (p. ej. español) pero "
                             "corresponden a estilos estándar. Se renombraron al nombre en inglés "
                             "que reconoce OxygenAuthor.",

        "step2_header": "Paso 2 · Revisión de estilos",
        "step2_caption": "Cada párrafo se clasifica según si su estilo necesita corregirse. Los "
                         "cambios automáticos de alta confianza se aplican solos; el resto espera "
                         "tu confirmación en la columna 'Estilo final'.",

        "table_toggle_label": "Convertir {n} párrafo(s) de cuerpo en tablas → Table Paragraph",
        "table_toggle_recommended": " ✅ Recomendado",
        "table_toggle_help": "Activa esta opción si el documento es un formulario SSCP, FMR/SSP "
                             "u otro documento con tablas densas de contenido.",

        "hierarchy_expander": "⚠️ {n} posible(s) salto(s) de jerarquía en headings",
        "hierarchy_caption": "Un salto de nivel (p.ej. Heading 3 justo después de Heading 1, sin "
                             "pasar por Heading 2) puede generar una estructura de topics/subtopics "
                             "incorrecta al convertir a DITA. Revisa si es intencionado.",
        "hierarchy_line": "Nivel {from_level} → **Nivel {to_level}** — *{text}*",

        "metric_no_change": "✅ Sin cambio",
        "metric_auto": "🔵 Cambio automático",
        "metric_review": "🟡 Revisar",
        "metric_unknown": "⚠️ Sin mapeo",
        "metric_review_delta": "{n} párrafos",
        "metric_unknown_delta": "{n} pendientes",

        "all_auto_success": "✅ Todos los cambios ({n}) son automáticos con alta confianza. "
                            "Puedes procesar directamente o revisar la tabla completa.",

        "show_all_toggle": "Ver documento completo ({n} párrafos)",
        "show_all_help": "Por defecto se muestran solo los párrafos que requieren atención.",
        "show_all_caption": "Mostrando todos los párrafos · {n} filas",
        "filtered_caption": "Mostrando {n} párrafos con cambios o que necesitan atención. "
                            "Activa 'Ver documento completo' para ver todos.",
        "no_attention_needed": "No hay párrafos que requieran atención. Todos los estilos son "
                               "reconocidos o quedan sin cambio.",

        "dropdown_hint": "💡 La columna **'Estilo final ✏️'** es un desplegable — haz clic en "
                         "cualquier celda para ver las opciones y elegir un estilo distinto.",

        "col_estado": "Estado",
        "col_texto": "Texto del párrafo",
        "col_estilo_original": "Estilo original",
        "col_estilo_final": "Estilo final ✏️",
        "col_estilo_final_help": "Haz clic para abrir el desplegable y elegir el estilo correcto.",
        "col_confianza": "Confianza",
        "opt_no_change": "(sin cambio)",

        "rules_saved_info": "💾 {n} regla(s) nueva(s) guardadas para futuros documentos.",

        "step3_header": "Paso 3 · Generar documento",
        "step3_caption": "Aplica todas las decisiones (automáticas + las confirmadas o corregidas "
                         "arriba) y genera un documento nuevo listo para el Batch Convert de "
                         "OxygenAuthor. El original nunca se modifica.",
        "process_button": "🔄 Procesar y descargar documento",
        "process_success": "✅ Listo. **{n}** párrafos modificados.",
        "process_errors": " ({n} errores — estilos no encontrados en el documento.)",
        "process_visual_note": "ℹ️ El cambio es de **estilo** (lo que necesita OxygenAuthor), no "
                               "necesariamente de aspecto visual. Si el texto ya estaba en negrita "
                               "antes del cambio, puede verse igual en Word — para comprobarlo, "
                               "haz clic en el párrafo y mira qué estilo aparece resaltado en el "
                               "panel de Estilos.",
        "download_button": "⬇️ Descargar documento procesado",

        # Página 2 — Reglas guardadas
        "page2_title": "📋 Reglas de mapeo guardadas",
        "page2_caption": "Historial de correcciones de estilo confirmadas por el equipo. Se "
                         "aplican automáticamente en futuros documentos cuando su confianza es alta.",
        "page2_no_rules": "Todavía no hay reglas. Procesa un documento para empezar.",
        "page2_total": "Total de reglas activas: **{n}**",
        "col_id": "ID", "col_origen": "Estilo origen", "col_tipo": "Tipo",
        "col_destino": "Estilo destino", "col_veces": "Veces aplicado",
        "col_rule_origin": "Origen", "col_creado": "Creado",
        "delete_rule_subheader": "Eliminar una regla",
        "delete_rule_label": "ID de la regla a eliminar:",
        "delete_rule_button": "🗑️ Eliminar regla",
        "delete_rule_success": "Regla {n} eliminada.",

        # Página 3 — Estilos de Oxygen
        "page3_title": "ℹ️ Estilos reconocidos por OxygenAuthor",
        "page3_caption": "Tabla de referencia cargada desde `stylesWordToDita.xml`. Solo estos "
                         "estilos son reconocidos directamente por el Batch Converter.",
        "page3_error_no_styles": "No se encontraron estilos. Verifica que existe "
                                 "`data/stylesWordToDita.xml`.",
        "page3_search": "🔍 Buscar estilo…",
        "col_nombre_estilo": "Nombre del estilo",
        "col_elemento_html": "Elemento HTML/DITA",
        "col_crea_bloque": "Crea bloque nuevo",
        "page3_total": "Total: {n} estilos",

        # Persistencia GitHub (backup)
        "backup_active": "🔄 Respaldo automático en GitHub activo — las reglas persistirán "
                         "aunque Render redespliegue la app.",
        "backup_inactive": "⚠️ Sin respaldo configurado — las reglas guardadas se perderán si "
                           "Render redespliega la app. Configura GITHUB_TOKEN y GIST_ID para "
                           "activar la persistencia.",
    },

    "en": {
        "tagline": "Style Mapper · Preparing Word documents for DITA conversion",
        "nav_label": "Go to",
        "nav_procesar": "🏠 Process document",
        "nav_reglas": "📋 Saved rules",
        "nav_estilos": "ℹ️ Oxygen styles",
        "help_open": "❓ Help",
        "help_close": "✖️ Close",
        "lang_toggle_label": "🌐 Español",

        "page1_title": "📄 Style Mapper for DITA",
        "page1_caption": "Prepare your Word documents for DITA conversion in OxygenAuthor.",
        "step1_header": "Step 1 · Upload your document",
        "step1_caption": "Accepts .docx files. If the document comes from Word in another "
                         "language, standard styles are detected and renamed automatically.",
        "step1_uploader_label": "Select a .docx file",
        "step1_info_no_file": "Upload a .docx file to get started.",
        "step1_success": "uploaded successfully.",

        "lang_norm_expander": "{n} style name(s) were normalized (document in another language)",
        "lang_norm_caption": "These styles had a translated name in Word (e.g. Spanish) but "
                             "correspond to standard styles. They were renamed to the English "
                             "name that OxygenAuthor recognizes.",

        "step2_header": "Step 2 · Style review",
        "step2_caption": "Each paragraph is classified by whether its style needs correcting. "
                         "High-confidence automatic changes are applied on their own; the rest "
                         "wait for your confirmation in the 'Final style' column.",

        "table_toggle_label": "Convert {n} table body paragraph(s) → Table Paragraph",
        "table_toggle_recommended": " ✅ Recommended",
        "table_toggle_help": "Enable this if the document is an SSCP, FMR/SSP form, or another "
                             "document with dense table content.",

        "hierarchy_expander": "⚠️ {n} possible heading hierarchy jump(s)",
        "hierarchy_caption": "A level jump (e.g. Heading 3 right after Heading 1, skipping "
                             "Heading 2) can produce an incorrect topic/subtopic structure when "
                             "converting to DITA. Check whether it's intentional.",
        "hierarchy_line": "Level {from_level} → **Level {to_level}** — *{text}*",

        "metric_no_change": "✅ No change",
        "metric_auto": "🔵 Automatic change",
        "metric_review": "🟡 Review",
        "metric_unknown": "⚠️ Unmapped",
        "metric_review_delta": "{n} paragraphs",
        "metric_unknown_delta": "{n} pending",

        "all_auto_success": "✅ All changes ({n}) are automatic with high confidence. You can "
                            "process directly or review the full table.",

        "show_all_toggle": "View full document ({n} paragraphs)",
        "show_all_help": "By default only paragraphs that need attention are shown.",
        "show_all_caption": "Showing all paragraphs · {n} rows",
        "filtered_caption": "Showing {n} paragraphs with changes or that need attention. "
                            "Toggle 'View full document' to see them all.",
        "no_attention_needed": "No paragraphs need attention. All styles are recognized or "
                               "left unchanged.",

        "dropdown_hint": "💡 The **'Final style ✏️'** column is a dropdown — click any cell to "
                         "see the options and pick a different style.",

        "col_estado": "Status",
        "col_texto": "Paragraph text",
        "col_estilo_original": "Original style",
        "col_estilo_final": "Final style ✏️",
        "col_estilo_final_help": "Click to open the dropdown and pick the correct style.",
        "col_confianza": "Confidence",
        "opt_no_change": "(no change)",

        "rules_saved_info": "💾 {n} new rule(s) saved for future documents.",

        "step3_header": "Step 3 · Generate document",
        "step3_caption": "Applies all decisions (automatic + those confirmed or corrected above) "
                         "and generates a new document ready for OxygenAuthor's Batch Convert. "
                         "The original is never modified.",
        "process_button": "🔄 Process and download document",
        "process_success": "✅ Done. **{n}** paragraphs modified.",
        "process_errors": " ({n} errors — styles not found in the document.)",
        "process_visual_note": "ℹ️ The change is a **style** change (what OxygenAuthor needs), "
                               "not necessarily a visual one. If the text was already bold before "
                               "the change, it may look the same in Word — to verify it, click the "
                               "paragraph and check which style is highlighted in the Styles panel.",
        "download_button": "⬇️ Download processed document",

        "page2_title": "📋 Saved mapping rules",
        "page2_caption": "History of style corrections confirmed by the team. They're applied "
                         "automatically to future documents when their confidence is high.",
        "page2_no_rules": "No rules yet. Process a document to get started.",
        "page2_total": "Total active rules: **{n}**",
        "col_id": "ID", "col_origen": "Source style", "col_tipo": "Type",
        "col_destino": "Target style", "col_veces": "Times applied",
        "col_rule_origin": "Origin", "col_creado": "Created",
        "delete_rule_subheader": "Delete a rule",
        "delete_rule_label": "ID of the rule to delete:",
        "delete_rule_button": "🗑️ Delete rule",
        "delete_rule_success": "Rule {n} deleted.",

        "page3_title": "ℹ️ Styles recognized by OxygenAuthor",
        "page3_caption": "Reference table loaded from `stylesWordToDita.xml`. Only these styles "
                         "are directly recognized by the Batch Converter.",
        "page3_error_no_styles": "No styles found. Verify that `data/stylesWordToDita.xml` exists.",
        "page3_search": "🔍 Search style…",
        "col_nombre_estilo": "Style name",
        "col_elemento_html": "HTML/DITA element",
        "col_crea_bloque": "Creates new block",
        "page3_total": "Total: {n} styles",

        "backup_active": "🔄 Automatic GitHub backup is active — rules will persist even if "
                         "Render redeploys the app.",
        "backup_inactive": "⚠️ No backup configured — saved rules will be lost if Render "
                           "redeploys the app. Set GITHUB_TOKEN and GIST_ID to enable persistence.",
    },
}


def get_lang() -> str:
    return st.session_state.get("language", "es")


def t(key: str, **kwargs) -> str:
    """Devuelve el texto traducido para la clave dada, en el idioma activo."""
    lang = get_lang()
    template = TEXT.get(lang, TEXT["es"]).get(key, TEXT["es"].get(key, key))
    if kwargs:
        try:
            return template.format(**kwargs)
        except (KeyError, IndexError):
            return template
    return template


def language_toggle():
    """
    Dibuja el interruptor ES/EN. Debe llamarse una vez, normalmente junto
    a la cabecera. Al cambiar, fuerza un rerun para refrescar toda la UI.
    """
    if "language" not in st.session_state:
        st.session_state.language = "es"

    current = st.session_state.language
    new_lang_label = "English" if current == "es" else "Español"
    if st.button(f"🌐 {new_lang_label}", key="vst_lang_toggle", use_container_width=True):
        st.session_state.language = "en" if current == "es" else "es"
        st.rerun()


# ── Estado (status) ────────────────────────────────────────────────────────
def status_label(status_code: str) -> str:
    mapping = {
        "no_change": t("metric_no_change"),
        "auto":      t("metric_auto"),
        "review":    t("metric_review"),
        "unknown":   t("metric_unknown"),
    }
    return mapping.get(status_code, status_code)


# ── Motivos (reason codes) de style_processor.py ───────────────────────────
REASON_TEMPLATES = {
    "es": {
        "empty": "párrafo vacío",
        "too_long": "demasiado largo ({n} palabras)",
        "bold_question_faq": "negrita + pregunta (formato FAQ)",
        "bold_caps_short": "negrita + mayúsculas/corto",
        "bold_titlecase": "negrita + capitalizado",
        "bold": "negrita",
        "caps_no_bold": "mayúsculas (sin negrita)",
        "not_heading_like": "no cumple criterios de heading",
        "no_explicit_bold": "sin negrita explícita",
        "too_long_subheading": "demasiado largo para sub-heading ({n} palabras)",
        "bold_short_words": "negrita ({n} palabras)",
        "bold_long_words": "negrita (largo: {n} palabras)",
        "has_bold_correct_heading": "tiene negrita — heading correcto",
        "empty_heading_no_bold": "heading vacío sin negrita",
        "heading_no_bold_long": "heading sin negrita y largo ({n} palabras)",
        "heading_no_bold_short": "heading sin negrita ({n} palabras)",
        "exact_match_case": "coincide salvo mayúsculas/espacios",
        "name_similarity": "similitud de nombre",
        "no_reasonable_match": "sin coincidencia razonable",
        "table_context": "en tabla, estilo '{n}' → Table Paragraph",
    },
    "en": {
        "empty": "empty paragraph",
        "too_long": "too long ({n} words)",
        "bold_question_faq": "bold + question (FAQ format)",
        "bold_caps_short": "bold + uppercase/short",
        "bold_titlecase": "bold + title case",
        "bold": "bold",
        "caps_no_bold": "uppercase (no bold)",
        "not_heading_like": "doesn't meet heading criteria",
        "no_explicit_bold": "no explicit bold",
        "too_long_subheading": "too long for a sub-heading ({n} words)",
        "bold_short_words": "bold ({n} words)",
        "bold_long_words": "bold (long: {n} words)",
        "has_bold_correct_heading": "has bold — correct heading",
        "empty_heading_no_bold": "empty heading without bold",
        "heading_no_bold_long": "heading without bold and long ({n} words)",
        "heading_no_bold_short": "heading without bold ({n} words)",
        "exact_match_case": "matches except for casing/spacing",
        "name_similarity": "name similarity",
        "no_reasonable_match": "no reasonable match",
        "table_context": "in table, style '{n}' → Table Paragraph",
    },
}


def format_reason(code_str: str) -> str:
    """
    Traduce un código de motivo generado por style_processor.py
    (p.ej. 'too_long|13' o 'bold_caps_short') al idioma activo.
    """
    if not code_str:
        return ""
    lang = get_lang()
    if "|" in code_str:
        code, value = code_str.split("|", 1)
    else:
        code, value = code_str, None
    template = REASON_TEMPLATES.get(lang, {}).get(code)
    if template is None:
        return code_str  # Código desconocido: devolver tal cual (no debería pasar)
    return template.format(n=value) if value is not None else template
