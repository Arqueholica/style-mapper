"""
theme.py
Estilos visuales, cabecera de marca y panel de ayuda para Style Mapper.
Paleta inspirada en GlobalLink Vasont (azul marino + acento amarillo).
"""

import streamlit as st

# ── Paleta de colores ──────────────────────────────────────────────────────
# Ajusta estos valores si tienes la paleta exacta de la marca (hex codes).
NAVY_DARK     = "#0F2A4E"   # Barra superior principal
NAVY_MEDIUM   = "#1B3A6B"   # Degradado / elementos secundarios
NAVY_LIGHT    = "#2C5490"   # Hover / bordes activos
ACCENT_YELLOW = "#F5C518"   # Acento — 'Tools' en el logo, resaltados
WHITE         = "#FFFFFF"
GRAY_BG       = "#F4F6F9"   # Fondo general de la app
GRAY_BORDER   = "#E1E5EA"
TEXT_MUTED    = "#C9D6E8"


def inject_custom_css():
    """Inyecta CSS global: paleta corporativa, tarjetas modernas y ajustes
    responsive para pantallas estrechas."""
    st.markdown(f"""
    <style>
        .stApp {{
            background-color: {GRAY_BG};
        }}

        /* ── Cabecera personalizada ──────────────────────────────────── */
        .vst-header {{
            display: flex;
            align-items: center;
            justify-content: space-between;
            background: linear-gradient(90deg, {NAVY_DARK} 0%, {NAVY_MEDIUM} 100%);
            padding: 16px 28px;
            border-radius: 12px;
            margin-bottom: 1.2rem;
            box-shadow: 0 2px 12px rgba(15, 42, 78, 0.25);
        }}
        .vst-logo-row {{
            display: flex;
            align-items: baseline;
            gap: 7px;
        }}
        .vst-logo-main {{
            color: {WHITE};
            font-size: 1.55rem;
            font-weight: 700;
            letter-spacing: 0.2px;
            font-family: 'Segoe UI', 'Helvetica Neue', Arial, sans-serif;
        }}
        .vst-logo-accent {{
            color: {ACCENT_YELLOW};
            font-size: 1.55rem;
            font-weight: 700;
            font-family: 'Segoe UI', 'Helvetica Neue', Arial, sans-serif;
        }}
        .vst-tagline {{
            color: {TEXT_MUTED};
            font-size: 0.8rem;
            margin-top: 1px;
            font-family: 'Segoe UI', 'Helvetica Neue', Arial, sans-serif;
        }}

        /* ── Botón de ayuda: estilo redondeado tipo "chip" ───────────── */
        div[data-testid="column"]:has(button[kind="secondary"]) button {{
            border-radius: 20px;
            border: 1px solid {NAVY_LIGHT};
        }}

        /* ── Tarjetas / expanders ─────────────────────────────────────── */
        div[data-testid="stExpander"] {{
            border: 1px solid {GRAY_BORDER};
            border-radius: 10px;
            background: {WHITE};
        }}

        /* ── Métricas como tarjetas ──────────────────────────────────── */
        div[data-testid="stMetric"] {{
            background: {WHITE};
            border: 1px solid {GRAY_BORDER};
            border-radius: 10px;
            padding: 12px 16px;
            box-shadow: 0 1px 4px rgba(0,0,0,0.04);
        }}

        /* ── Botón primario con el azul de marca ─────────────────────── */
        button[kind="primary"] {{
            background-color: {NAVY_MEDIUM} !important;
            border-radius: 8px !important;
            border: none !important;
        }}
        button[kind="primary"]:hover {{
            background-color: {NAVY_DARK} !important;
        }}

        /* ── Panel de ayuda (columna derecha) ────────────────────────── */
        .vst-help-panel {{
            background: {WHITE};
            border: 1px solid {GRAY_BORDER};
            border-left: 4px solid {ACCENT_YELLOW};
            border-radius: 10px;
            padding: 18px 20px;
            margin-bottom: 14px;
        }}
        .vst-help-title {{
            color: {NAVY_DARK};
            font-weight: 700;
            font-size: 1.0rem;
            margin-bottom: 6px;
        }}

        /* ── Responsive: pantallas estrechas (móvil / tablet) ────────── */
        @media (max-width: 700px) {{
            .vst-header {{
                flex-direction: column;
                align-items: flex-start;
                gap: 10px;
                padding: 14px 18px;
            }}
            .vst-logo-main, .vst-logo-accent {{
                font-size: 1.25rem;
            }}
        }}
    </style>
    """, unsafe_allow_html=True)


def render_header():
    """
    Dibuja la cabecera de marca ('Vasont Inspire Tools') con el interruptor
    de idioma y el botón de ayuda a la derecha. El estado de apertura de la
    ayuda se guarda en st.session_state.help_open y se lee desde app.py
    para decidir el layout.
    """
    import i18n

    if "help_open" not in st.session_state:
        st.session_state.help_open = False
    if "language" not in st.session_state:
        st.session_state.language = "es"

    col_logo, col_lang, col_btn = st.columns([5, 1, 1], vertical_alignment="center")

    with col_logo:
        st.markdown(
            f"""
            <div class="vst-header">
                <div>
                    <div class="vst-logo-row">
                        <span class="vst-logo-main">Vasont Inspire</span>
                        <span class="vst-logo-accent">Tools</span>
                    </div>
                    <div class="vst-tagline">
                        {i18n.t("tagline")}
                    </div>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with col_lang:
        i18n.language_toggle()

    with col_btn:
        label = i18n.t("help_close") if st.session_state.help_open else i18n.t("help_open")
        if st.button(label, use_container_width=True, key="vst_help_toggle"):
            st.session_state.help_open = not st.session_state.help_open
            st.rerun()


# ── Contenido de ayuda por sección (bilingüe) ──────────────────────────────
HELP_SECTIONS = {
    "es": {
        "procesar": [
            (
                "🏠 Procesar documento — visión general",
                "Esta pantalla prepara un documento Word para que OxygenAuthor "
                "pueda convertirlo correctamente a DITA. El proceso tiene tres "
                "pasos: subir el archivo, revisar los estilos sugeridos, y "
                "generar el documento corregido para descargar."
            ),
            (
                "Paso 1 · Sube tu documento",
                "Acepta archivos .docx. Si el documento fue creado o editado en "
                "Word en otro idioma (por ejemplo español), la app detecta y "
                "renombra automáticamente los estilos estándar traducidos "
                "(ej. 'Título 1' → 'Heading 1') para que sean reconocibles."
            ),
            (
                "Convertir párrafos de tablas",
                "Si el documento tiene mucho contenido dentro de tablas (formularios "
                "SSCP, FMR/SSP), este interruptor sugiere convertir esos párrafos "
                "al estilo 'Table Paragraph'. Está desactivado por defecto porque "
                "no todos los documentos con tablas lo necesitan — actívalo solo "
                "si reconoces ese tipo de documento."
            ),
            (
                "Saltos de jerarquía en headings",
                "Si aparece este aviso, significa que la estructura de títulos "
                "tiene un salto (p. ej. un Heading 3 justo después de un Heading 1, "
                "sin pasar por Heading 2). Esto puede generar una jerarquía de "
                "topics incorrecta en DITA — revisa si es intencionado."
            ),
            (
                "Paso 2 · Revisión de estilos",
                "Cada párrafo se clasifica en uno de cuatro estados:\n\n"
                "- ✅ **Sin cambio** — el estilo ya es válido para OxygenAuthor.\n"
                "- 🔵 **Automático** — se aplicará un cambio de alta confianza.\n"
                "- 🟡 **Revisar** — hay una sugerencia, pero necesita tu confirmación.\n"
                "- ⚠️ **Sin mapeo** — estilo desconocido sin sugerencia clara.\n\n"
                "Por defecto solo se muestran los párrafos que necesitan atención. "
                "Activa 'Ver documento completo' para revisar todos."
            ),
            (
                "La columna 'Estilo final'",
                "Es un desplegable editable. Haz clic en cualquier celda para "
                "abrir las opciones y elegir el estilo correcto si la sugerencia "
                "no es la adecuada. Cada corrección que hagas se guarda como "
                "regla para futuros documentos — la app aprende con el uso."
            ),
            (
                "Paso 3 · Generar documento",
                "Al pulsar 'Procesar y descargar', se aplican todas las "
                "decisiones (automáticas + las que confirmaste o corregiste en "
                "la tabla) y se genera un documento nuevo, listo para el Batch "
                "Convert de OxygenAuthor. El original nunca se modifica.\n\n"
                "⚠️ Importante: el cambio es de **nombre de estilo**, no siempre "
                "de aspecto visual. Si el texto ya estaba en negrita, puede "
                "verse igual en Word aunque el estilo haya cambiado correctamente "
                "por debajo. Compruébalo en el panel de Estilos de Word."
            ),
        ],
        "reglas": [
            (
                "📋 Reglas guardadas",
                "Aquí se listan todas las correcciones de estilo que el equipo "
                "ha confirmado a lo largo del tiempo. Cada regla dice: 'cuando "
                "aparezca este estilo, sugiere este otro'. Las reglas con más "
                "usos son más fiables — se aplican automáticamente si su "
                "confianza es alta."
            ),
            (
                "Eliminar una regla",
                "Si una regla resulta ser incorrecta, puedes desactivarla por su "
                "ID. No se borra el historial, solo deja de aplicarse a partir "
                "de ahora."
            ),
            (
                "Persistencia de las reglas",
                "Las reglas se respaldan automáticamente en GitHub (si está "
                "configurado) para que no se pierdan cuando Render redespliegue "
                "la app. Verás un aviso arriba indicando si el respaldo está "
                "activo."
            ),
        ],
        "estilos": [
            (
                "ℹ️ Estilos de Oxygen",
                "Esta es la lista de referencia de estilos que el Batch "
                "Converter de OxygenAuthor reconoce directamente, cargada desde "
                "el archivo de configuración `stylesWordToDita.xml`. Cualquier "
                "estilo de Word que no aparezca aquí necesita una regla de "
                "mapeo para convertirse correctamente."
            ),
        ],
    },
    "en": {
        "procesar": [
            (
                "🏠 Process document — overview",
                "This screen prepares a Word document so OxygenAuthor can "
                "correctly convert it to DITA. The process has three steps: "
                "upload the file, review the suggested styles, and generate "
                "the corrected document for download."
            ),
            (
                "Step 1 · Upload your document",
                "Accepts .docx files. If the document was created or edited "
                "in Word in another language (e.g. Spanish), the app detects "
                "and automatically renames the translated standard styles "
                "(e.g. 'Título 1' → 'Heading 1') so they're recognizable."
            ),
            (
                "Convert table paragraphs",
                "If the document has a lot of content inside tables (SSCP, "
                "FMR/SSP forms), this toggle suggests converting those "
                "paragraphs to the 'Table Paragraph' style. It's off by "
                "default because not every document with tables needs it — "
                "only enable it if you recognize that type of document."
            ),
            (
                "Heading hierarchy jumps",
                "If this warning appears, it means the heading structure has "
                "a jump (e.g. a Heading 3 right after a Heading 1, skipping "
                "Heading 2). This can produce an incorrect topic hierarchy "
                "in DITA — check whether it's intentional."
            ),
            (
                "Step 2 · Style review",
                "Each paragraph is classified into one of four states:\n\n"
                "- ✅ **No change** — the style is already valid for OxygenAuthor.\n"
                "- 🔵 **Automatic** — a high-confidence change will be applied.\n"
                "- 🟡 **Review** — there's a suggestion, but it needs your confirmation.\n"
                "- ⚠️ **Unmapped** — unknown style with no clear suggestion.\n\n"
                "By default only paragraphs that need attention are shown. "
                "Toggle 'View full document' to review them all."
            ),
            (
                "The 'Final style' column",
                "It's an editable dropdown. Click any cell to open the "
                "options and pick the right style if the suggestion isn't "
                "correct. Every correction you make is saved as a rule for "
                "future documents — the app learns as it's used."
            ),
            (
                "Step 3 · Generate document",
                "Clicking 'Process and download' applies all decisions "
                "(automatic + those you confirmed or corrected in the table) "
                "and generates a new document, ready for OxygenAuthor's Batch "
                "Convert. The original is never modified.\n\n"
                "⚠️ Important: the change is a **style name** change, not "
                "always a visual one. If the text was already bold, it may "
                "look the same in Word even though the style changed "
                "correctly underneath. Verify it in Word's Styles panel."
            ),
        ],
        "reglas": [
            (
                "📋 Saved rules",
                "This lists all the style corrections the team has confirmed "
                "over time. Each rule says: 'when this style appears, suggest "
                "this other one'. Rules with more uses are more reliable — "
                "they're applied automatically when their confidence is high."
            ),
            (
                "Deleting a rule",
                "If a rule turns out to be incorrect, you can deactivate it "
                "by its ID. The history isn't erased, it just stops being "
                "applied from now on."
            ),
            (
                "Rule persistence",
                "Rules are automatically backed up to GitHub (if configured) "
                "so they aren't lost when Render redeploys the app. You'll "
                "see a notice above indicating whether the backup is active."
            ),
        ],
        "estilos": [
            (
                "ℹ️ Oxygen styles",
                "This is the reference list of styles that OxygenAuthor's "
                "Batch Converter recognizes directly, loaded from the "
                "`stylesWordToDita.xml` configuration file. Any Word style "
                "not appearing here needs a mapping rule to convert correctly."
            ),
        ],
    },
}


def render_help_panel(page_key: str):
    """
    Dibuja el panel de ayuda para la página actual dentro de la columna
    derecha, en el idioma activo. page_key debe ser una clave de
    HELP_SECTIONS[lang] ('procesar', 'reglas' o 'estilos').
    """
    import i18n

    lang = i18n.get_lang()
    st.markdown(f"### {i18n.t('help_open')}" if lang == "en" else "### ❓ Ayuda")
    st.caption(
        "Step-by-step explanation of this screen." if lang == "en"
        else "Explicación de esta pantalla, paso a paso."
    )
    st.divider()

    sections = HELP_SECTIONS.get(lang, HELP_SECTIONS["es"]).get(page_key, [])
    for title, body in sections:
        st.markdown(
            f"""
            <div class="vst-help-panel">
                <div class="vst-help-title">{title}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        st.markdown(body)
        st.write("")
