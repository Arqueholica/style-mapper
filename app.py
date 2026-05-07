"""
app.py
Interfaz principal de Style Mapper para DITA.
Ejecutar con:  streamlit run app.py
"""

import os
import tempfile

import pandas as pd
import streamlit as st

from database import (
    delete_rule,
    get_all_rules,
    get_known_style_names,
    get_known_styles,
    get_style_rule,
    initialize_database,
    save_style_rule,
)
from style_processor import apply_style_mappings, extract_styles_from_docx

# ──────────────────────────────────────────────
# Configuración general
# ──────────────────────────────────────────────
initialize_database()

st.set_page_config(
    page_title="Style Mapper para DITA",
    page_icon="📄",
    layout="wide",
)

st.title("📄 Style Mapper para DITA")
st.caption(
    "Prepara tus documentos Word para la conversión a DITA en OxygenAuthor. "
    "Detecta estilos no reconocidos, define el mapeo y descarga el documento listo."
)

# ──────────────────────────────────────────────
# Navegación lateral
# ──────────────────────────────────────────────
page = st.sidebar.radio(
    "Ir a",
    ["🏠 Procesar documento", "📋 Reglas guardadas", "ℹ️ Estilos de Oxygen"],
)

# ══════════════════════════════════════════════
# PÁGINA 1 — Procesar documento
# ══════════════════════════════════════════════
if page == "🏠 Procesar documento":

    st.header("Paso 1 · Sube tu documento")
    uploaded = st.file_uploader(
        "Selecciona un archivo .docx", type=["docx"], label_visibility="collapsed"
    )

    if not uploaded:
        st.info("Sube un archivo .docx para empezar.")
        st.stop()

    # Guardar en archivo temporal
    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
        tmp.write(uploaded.read())
        tmp_path = tmp.name

    st.success(f"✅ Archivo cargado: **{uploaded.name}**")

    # ── Extraer estilos ──────────────────────────
    with st.spinner("Analizando estilos del documento…"):
        try:
            styles_found = extract_styles_from_docx(tmp_path)
        except ValueError as e:
            st.error(str(e))
            os.unlink(tmp_path)
            st.stop()

    known_names   = get_known_style_names()
    known_list    = sorted([s["style_name"] for s in get_known_styles()])

    # Clasificar cada estilo encontrado
    auto_ok   = {}   # Reconocido directamente por Oxygen
    has_rule  = {}   # Tiene regla guardada
    unknown   = {}   # Sin mapeo → necesita decisión del usuario

    for sname, info in styles_found.items():
        # Quitar el sufijo " (carácter)" que añadimos para distinguir
        clean_name = sname.replace(" (carácter)", "")
        if clean_name in known_names:
            auto_ok[sname] = info
        else:
            rule = get_style_rule(clean_name, info["element_type"])
            if rule:
                has_rule[sname] = {**info, "rule": rule}
            else:
                unknown[sname] = info

    # ── Resumen ─────────────────────────────────
    st.header("Paso 2 · Revisión de estilos encontrados")
    c1, c2, c3 = st.columns(3)
    c1.metric("✅ Reconocidos por Oxygen",      len(auto_ok))
    c2.metric("🔵 Con regla guardada",           len(has_rule))
    c3.metric("⚠️ Sin mapeo (requieren acción)", len(unknown))

    # ── Estilos sin mapeo ────────────────────────
    new_mappings = {}   # {nombre_estilo: estilo_destino} decidido ahora

    if unknown:
        st.subheader("⚠️ Estilos que necesitan mapeo")
        st.info(
            "Estos estilos no son reconocidos por OxygenAuthor. "
            "Elige a qué estilo estándar deben convertirse."
        )

        for sname, info in unknown.items():
            clean_name = sname.replace(" (carácter)", "")
            label = (
                f"**{sname}** &nbsp;·&nbsp; tipo: `{info['element_type']}` "
                f"&nbsp;·&nbsp; aparece **{info['count']}** vez/veces"
            )
            with st.expander(label):
                st.caption(f"Ejemplo de texto: *{info['sample']}*")
                selected = st.selectbox(
                    "Mapear a:",
                    options=["(no mapear)"] + known_list,
                    key=f"sel_{sname}",
                )
                if selected != "(no mapear)":
                    new_mappings[clean_name] = selected
    else:
        st.success("✅ Todos los estilos tienen mapeo. Puedes procesar el documento.")

    # ── Reglas ya guardadas ──────────────────────
    if has_rule:
        with st.expander(
            f"🔵 {len(has_rule)} estilo(s) con regla guardada "
            "(se aplicarán automáticamente)"
        ):
            for sname, info in has_rule.items():
                r = info["rule"]
                st.write(
                    f"• **{sname}** → `{r['target_style']}` "
                    f"*(aplicado {r['times_applied']} veces antes)*"
                )

    # ── Procesar y descargar ─────────────────────
    st.header("Paso 3 · Generar documento")

    if st.button("🔄 Procesar y descargar documento", type="primary"):

        # Construir el diccionario completo de mapeos
        all_mappings = {}

        # 1. Reglas ya guardadas
        for sname, info in has_rule.items():
            clean = sname.replace(" (carácter)", "")
            all_mappings[clean] = info["rule"]["target_style"]

        # 2. Decisiones nuevas del usuario
        for clean_name, target in new_mappings.items():
            all_mappings[clean_name] = target
            element_type = unknown.get(clean_name, {}).get("element_type", "p")
            save_style_rule(clean_name, element_type, target)

        output_path = tmp_path.replace(".docx", "_PROCESADO.docx")

        with st.spinner("Aplicando cambios de estilo…"):
            n_changes = apply_style_mappings(tmp_path, output_path, all_mappings)

        with open(output_path, "rb") as f:
            output_bytes = f.read()

        out_name = uploaded.name.replace(".docx", "_ESTILOS_CORREGIDOS.docx")

        st.success(f"✅ Listo. Se realizaron **{n_changes}** cambios de estilo.")
        st.download_button(
            label="⬇️ Descargar documento procesado",
            data=output_bytes,
            file_name=out_name,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )

        # Limpiar archivos temporales
        try:
            os.unlink(tmp_path)
            os.unlink(output_path)
        except OSError:
            pass

# ══════════════════════════════════════════════
# PÁGINA 2 — Reglas guardadas
# ══════════════════════════════════════════════
elif page == "📋 Reglas guardadas":

    st.header("📋 Reglas de mapeo guardadas")

    rules = get_all_rules()

    if not rules:
        st.info(
            "Todavía no hay reglas guardadas. "
            "Procesa un documento y define mapeos para que aparezcan aquí."
        )
        st.stop()

    st.write(f"Total de reglas activas: **{len(rules)}**")

    df = pd.DataFrame(rules)[
        ["id", "source_style", "source_element", "target_style",
         "times_applied", "rule_origin", "created_at"]
    ]
    df.columns = [
        "ID", "Estilo origen", "Tipo", "Estilo destino",
        "Veces aplicado", "Origen", "Creado"
    ]

    st.dataframe(df, use_container_width=True, hide_index=True)

    st.divider()
    st.subheader("Eliminar una regla")
    rule_id = st.number_input(
        "ID de la regla a eliminar (ver tabla de arriba):",
        min_value=1, step=1
    )
    if st.button("🗑️ Eliminar regla", type="secondary"):
        delete_rule(int(rule_id))
        st.success(f"Regla {rule_id} eliminada.")
        st.rerun()

# ══════════════════════════════════════════════
# PÁGINA 3 — Estilos de Oxygen
# ══════════════════════════════════════════════
elif page == "ℹ️ Estilos de Oxygen":

    st.header("ℹ️ Estilos reconocidos por OxygenAuthor")
    st.caption(
        "Esta es la tabla de referencia cargada desde tu archivo "
        "`stylesWordToDita.xml`. Son los únicos estilos que el "
        "Batch Converter de Oxygen reconoce directamente."
    )

    known = get_known_styles()
    if not known:
        st.error("No se encontraron estilos. Verifica que existe `data/stylesWordToDita.xml`.")
        st.stop()

    df = pd.DataFrame(known)[
        ["element_type", "style_name", "target_html", "creates_new_block"]
    ]
    df.columns = ["Tipo", "Nombre del estilo", "Elemento HTML/DITA", "Crea bloque nuevo"]

    search = st.text_input("🔍 Buscar estilo…")
    if search:
        df = df[df["Nombre del estilo"].str.contains(search, case=False, na=False)]

    st.dataframe(df, use_container_width=True, hide_index=True)
    st.caption(f"Total: {len(df)} estilos")
