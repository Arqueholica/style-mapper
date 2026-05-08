"""
app.py — Style Mapper para DITA
Interfaz principal con revisión párrafo a párrafo.
"""

import os
import tempfile

import pandas as pd
import streamlit as st

from database import (
    delete_rule, get_all_rules, get_known_style_names,
    get_known_styles, get_style_rule, initialize_database, save_style_rule,
)
from style_processor import (
    analyze_document, apply_paragraph_decisions, STATUS_LABELS,
)

# ── Inicialización ────────────────────────────────────────────────────────────
initialize_database()

st.set_page_config(
    page_title="Style Mapper para DITA",
    page_icon="📄",
    layout="wide",
)

# ── Navegación ────────────────────────────────────────────────────────────────
page = st.sidebar.radio(
    "Ir a",
    ["🏠 Procesar documento", "📋 Reglas guardadas", "ℹ️ Estilos de Oxygen"],
)

# ══════════════════════════════════════════════════════════════════════════════
# PÁGINA 1 — Procesar documento
# ══════════════════════════════════════════════════════════════════════════════
if page == "🏠 Procesar documento":

    st.title("📄 Style Mapper para DITA")
    st.caption("Prepara tus documentos Word para la conversión a DITA en OxygenAuthor.")

    # ── Paso 1: subir archivo ─────────────────────────────────────────────────
    st.header("Paso 1 · Sube tu documento")
    uploaded = st.file_uploader(
        "Selecciona un archivo .docx", type=["docx"], label_visibility="collapsed"
    )

    if not uploaded:
        st.info("Sube un archivo .docx para empezar.")
        st.stop()

    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
        tmp.write(uploaded.read())
        tmp_path = tmp.name

    st.success(f"✅ **{uploaded.name}** cargado correctamente.")

    # ── Paso 2: analizar ──────────────────────────────────────────────────────
    st.header("Paso 2 · Revisión de estilos")

    known_names = get_known_style_names()
    known_list  = sorted([s["style_name"] for s in get_known_styles()])

    # Construir rules_dict desde la base de datos: {style: (target, confidence)}
    all_rules  = get_all_rules()
    rules_dict = {
        r["source_style"]: (r["target_style"], float(r["confidence"]))
        for r in all_rules
    }

    with st.spinner("Analizando párrafos…"):
        try:
            paragraphs = analyze_document(tmp_path, rules_dict, known_names)
        except ValueError as e:
            st.error(str(e))
            os.unlink(tmp_path)
            st.stop()

    # ── Resumen ───────────────────────────────────────────────────────────────
    total    = len(paragraphs)
    n_auto   = sum(1 for p in paragraphs if p["status"] == "auto")
    n_review = sum(1 for p in paragraphs if p["status"] == "review")
    n_unkn   = sum(1 for p in paragraphs if p["status"] == "unknown")
    n_ok     = sum(1 for p in paragraphs if p["status"] == "no_change")

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("✅ Sin cambio",         n_ok)
    c2.metric("🔵 Cambio automático",  n_auto)
    c3.metric("🟡 Revisar",            n_review,
              delta=f"{n_review} párrafos" if n_review else None,
              delta_color="off")
    c4.metric("⚠️ Sin mapeo",          n_unkn,
              delta=f"{n_unkn} pendientes" if n_unkn else None,
              delta_color="inverse")

    # ── Tabla de revisión ─────────────────────────────────────────────────────
    needs_attention = n_review + n_unkn

    if needs_attention == 0 and n_auto > 0:
        st.success(
            f"✅ Todos los cambios ({n_auto}) son automáticos con alta confianza. "
            "Puedes procesar directamente o revisar la tabla completa."
        )

    # Estado del toggle de vista completa
    show_all = st.toggle(
        f"Ver documento completo ({total} párrafos)",
        value=False,
        help="Por defecto se muestran solo los párrafos que requieren atención."
    )

    # Filtrar filas según la vista
    if show_all:
        display_rows = paragraphs
        st.caption(f"Mostrando todos los párrafos · {total} filas")
    else:
        display_rows = [
            p for p in paragraphs
            if p["status"] in ("review", "unknown", "auto")
        ]
        if not display_rows:
            st.info("No hay párrafos que requieran atención. Todos los estilos son reconocidos o quedan sin cambio.")
        else:
            st.caption(
                f"Mostrando {len(display_rows)} párrafos con cambios o que necesitan atención. "
                "Activa 'Ver documento completo' para ver todos."
            )

    # Construir DataFrame para el editor
    df = pd.DataFrame([
        {
            "Estado":           p["status_label"],
            "Texto":            p["text"],
            "Estilo original":  p["original_style"],
            "Estilo final":     p["final_style"],
            "Confianza":        p["confidence_pct"],
            "_idx":             p["idx"],
            "_original":        p["original_style"],
        }
        for p in display_rows
    ])

    if not df.empty:
        edited_df = st.data_editor(
            df,
            column_config={
                "Estado": st.column_config.TextColumn(
                    "Estado", width="small", disabled=True
                ),
                "Texto": st.column_config.TextColumn(
                    "Texto del párrafo", width="large", disabled=True
                ),
                "Estilo original": st.column_config.TextColumn(
                    "Estilo original", width="medium", disabled=True
                ),
                "Estilo final": st.column_config.SelectboxColumn(
                    "Estilo final ✏️",
                    width="medium",
                    options=["(sin cambio)"] + known_list,
                    required=True,
                ),
                "Confianza": st.column_config.ProgressColumn(
                    "Confianza", width="small",
                    min_value=0, max_value=100, format="%d%%",
                ),
                "_idx":      None,   # columna oculta
                "_original": None,   # columna oculta
            },
            hide_index=True,
            use_container_width=True,
            height=min(600, 50 + len(df) * 35),
            key="review_table",
        )
    else:
        edited_df = df

    # ── Guardar nuevas reglas desde la tabla ──────────────────────────────────
    if not edited_df.empty:
        changed_rules = edited_df[
            (edited_df["Estilo final"] != "(sin cambio)") &
            (edited_df["Estilo final"] != edited_df["Estilo original"])
        ]
        if not changed_rules.empty:
            new_rule_count = 0
            for _, row in changed_rules.iterrows():
                orig  = row["Estilo original"]
                final = row["Estilo final"]
                existing = get_style_rule(orig, "p")
                if not existing or existing["target_style"] != final:
                    save_style_rule(orig, "p", final, origin="human")
                    new_rule_count += 1
            if new_rule_count:
                st.info(
                    f"💾 {new_rule_count} regla(s) nueva(s) guardadas para futuros documentos."
                )

    # ── Paso 3: procesar ──────────────────────────────────────────────────────
    st.header("Paso 3 · Generar documento")

    if st.button("🔄 Procesar y descargar documento", type="primary"):

        # Construir lista de decisiones desde la tabla editada
        # Combinar: filas visibles (editadas) + filas ocultas (si vista filtrada)
        decisions_from_table = {}
        if not edited_df.empty:
            for _, row in edited_df.iterrows():
                final = row["Estilo final"]
                if final == "(sin cambio)":
                    final = row["Estilo original"]
                decisions_from_table[int(row["_idx"])] = {
                    "idx":            int(row["_idx"]),
                    "final_style":    final,
                    "original_style": row["_original"],
                }

        # Para párrafos no visibles en la tabla, usar la sugerencia automática
        all_decisions = []
        for p in paragraphs:
            if p["idx"] in decisions_from_table:
                all_decisions.append(decisions_from_table[p["idx"]])
            else:
                # Fuera de la tabla visible: aplicar sugerencia si era auto
                final = p["suggested_style"] if p["status"] == "auto" else p["original_style"]
                all_decisions.append({
                    "idx":            p["idx"],
                    "final_style":    final,
                    "original_style": p["original_style"],
                })

        output_path = tmp_path.replace(".docx", "_PROCESADO.docx")

        with st.spinner("Aplicando cambios…"):
            n_changes, n_errors = apply_paragraph_decisions(
                tmp_path, output_path, all_decisions
            )

        with open(output_path, "rb") as f:
            output_bytes = f.read()

        out_name = uploaded.name.replace(".docx", "_ESTILOS_CORREGIDOS.docx")

        msg = f"✅ Listo. **{n_changes}** párrafos modificados."
        if n_errors:
            msg += f" ({n_errors} errores — estilos no encontrados en el documento.)"
        st.success(msg)

        st.download_button(
            label="⬇️ Descargar documento procesado",
            data=output_bytes,
            file_name=out_name,
            mime=(
                "application/vnd.openxmlformats-officedocument"
                ".wordprocessingml.document"
            ),
        )

        try:
            os.unlink(tmp_path)
            os.unlink(output_path)
        except OSError:
            pass

# ══════════════════════════════════════════════════════════════════════════════
# PÁGINA 2 — Reglas guardadas
# ══════════════════════════════════════════════════════════════════════════════
elif page == "📋 Reglas guardadas":

    st.title("📋 Reglas de mapeo guardadas")

    rules = get_all_rules()

    if not rules:
        st.info("Todavía no hay reglas. Procesa un documento para empezar.")
        st.stop()

    st.write(f"Total de reglas activas: **{len(rules)}**")

    df = pd.DataFrame(rules)[[
        "id", "source_style", "source_element", "target_style",
        "confidence", "times_applied", "rule_origin", "created_at"
    ]]
    df.columns = [
        "ID", "Estilo origen", "Tipo", "Estilo destino",
        "Confianza", "Veces aplicado", "Origen", "Creado"
    ]
    df["Confianza"] = (df["Confianza"] * 100).round(0).astype(int).astype(str) + "%"

    st.dataframe(df, use_container_width=True, hide_index=True)

    st.divider()
    st.subheader("Eliminar una regla")
    rule_id = st.number_input("ID de la regla a eliminar:", min_value=1, step=1)
    if st.button("🗑️ Eliminar regla", type="secondary"):
        delete_rule(int(rule_id))
        st.success(f"Regla {rule_id} eliminada.")
        st.rerun()

# ══════════════════════════════════════════════════════════════════════════════
# PÁGINA 3 — Estilos de Oxygen
# ══════════════════════════════════════════════════════════════════════════════
elif page == "ℹ️ Estilos de Oxygen":

    st.title("ℹ️ Estilos reconocidos por OxygenAuthor")
    st.caption(
        "Tabla de referencia cargada desde `stylesWordToDita.xml`. "
        "Solo estos estilos son reconocidos directamente por el Batch Converter."
    )

    known = get_known_styles()
    if not known:
        st.error("No se encontraron estilos. Verifica que existe `data/stylesWordToDita.xml`.")
        st.stop()

    df = pd.DataFrame(known)[["element_type", "style_name", "target_html", "creates_new_block"]]
    df.columns = ["Tipo", "Nombre del estilo", "Elemento HTML/DITA", "Crea bloque nuevo"]

    search = st.text_input("🔍 Buscar estilo…")
    if search:
        df = df[df["Nombre del estilo"].str.contains(search, case=False, na=False)]

    st.dataframe(df, use_container_width=True, hide_index=True)
    st.caption(f"Total: {len(df)} estilos")
