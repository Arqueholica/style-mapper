"""
app.py — Vasont Inspire Tools · Style Mapper para DITA
Interfaz principal con revisión párrafo a párrafo.
"""

import os
import tempfile

import pandas as pd
import streamlit as st

import i18n
import theme
from database import (
    delete_rule, get_all_rules, get_known_style_names,
    get_known_styles, get_style_rule, initialize_database, save_style_rule,
)
from style_processor import (
    analyze_document, apply_paragraph_decisions, normalize_document_file,
    count_table_body_paragraphs, check_heading_hierarchy,
)

try:
    import github_sync
except ImportError:
    github_sync = None

# ── Inicialización ────────────────────────────────────────────────────────────
initialize_database()

st.set_page_config(
    page_title="Vasont Inspire Tools · Style Mapper",
    page_icon="🧰",
    layout="wide",
)

theme.inject_custom_css()
theme.render_header()

# ── Navegación ────────────────────────────────────────────────────────────────
page = st.sidebar.radio(
    i18n.t("nav_label"),
    [i18n.t("nav_procesar"), i18n.t("nav_reglas"), i18n.t("nav_estilos")],
)

# Clave de ayuda correspondiente a la página actual (ver theme.HELP_SECTIONS)
_HELP_KEY_BY_PAGE = {
    i18n.t("nav_procesar"): "procesar",
    i18n.t("nav_reglas"):   "reglas",
    i18n.t("nav_estilos"):  "estilos",
}

# ── Layout: columna principal + columna de ayuda (si está abierta) ───────────
# La columna de ayuda "empuja" el contenido principal en vez de superponerse:
# al abrirla, el área principal se estrecha para dejar sitio a la ayuda a la
# derecha. Se renderiza AQUÍ, antes del contenido principal, porque varias
# pantallas usan st.stop() — si la ayuda se dibujara después, se quedaría
# vacía justo en esos casos.
if st.session_state.get("help_open"):
    main_col, help_col = st.columns([3, 1], gap="large")
    with help_col:
        _help_key = _HELP_KEY_BY_PAGE.get(page, "procesar")
        theme.render_help_panel(_help_key)
else:
    main_col = st.container()

with main_col:
# ══════════════════════════════════════════════════════════════════════════════
# PÁGINA 1 — Procesar documento
# ══════════════════════════════════════════════════════════════════════════════
    if page == i18n.t("nav_procesar"):

        st.title(i18n.t("page1_title"))
        st.caption(i18n.t("page1_caption"))

        # ── Paso 1: subir archivo ─────────────────────────────────────────────
        st.header(i18n.t("step1_header"))
        st.caption(i18n.t("step1_caption"))
        uploaded = st.file_uploader(
            i18n.t("step1_uploader_label"), type=["docx", "idml"], label_visibility="collapsed"
        )

        if not uploaded:
            st.info(i18n.t("step1_info_no_file"))
            st.stop()

        is_idml = uploaded.name.lower().endswith(".idml")
        suffix = ".idml" if is_idml else ".docx"

        with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
            # IMPORTANTE: usar getvalue(), NO read(). Streamlit re-ejecuta TODO
            # el script en cada interacción (toggle, edición de tabla, clic en
            # botón). El objeto 'uploaded' es el MISMO en cada re-ejecución, y
            # .read() agota el buffer la primera vez — las siguientes veces
            # devolvería bytes vacíos, corrompiendo silenciosamente el archivo
            # que se procesa al pulsar "Procesar y descargar documento".
            # .getvalue() no mueve el puntero de lectura, así que es seguro
            # llamarlo en cada re-ejecución sin perder el contenido.
            tmp.write(uploaded.getvalue())
            uploaded_path = tmp.name

        st.success(f"✅ **{uploaded.name}** {i18n.t('step1_success')}")

        if is_idml:
            # ── Adaptador de InDesign ──────────────────────────────────────
            # Transcribe el contenido del .idml a un .docx sintético: mismo
            # texto, con el nombre de estilo que tenía en InDesign (creado al
            # vuelo si no existe como estilo real de Word). A partir de aquí,
            # el documento entra en el mismo pipeline que cualquier .docx
            # subido directamente — sin ningún cambio en la lógica existente.
            with st.spinner("…"):
                idml_synthetic_path = uploaded_path.replace(".idml", "_SYNTH.docx")
                try:
                    from idml_handler import build_synthetic_docx
                    n_transcribed = build_synthetic_docx(uploaded_path, idml_synthetic_path)
                except Exception as e:
                    st.error(f"No se pudo leer el archivo InDesign: {e}")
                    st.stop()
            st.caption(
                f"🎨 Documento InDesign — {n_transcribed} párrafos transcritos a un "
                f"DOCX intermedio antes de analizar los estilos."
            )
            raw_path = idml_synthetic_path
            try:
                os.unlink(uploaded_path)
            except OSError:
                pass
        else:
            raw_path = uploaded_path

        # ── Normalización de idioma ───────────────────────────────────────────
        # Word en español (u otros idiomas) traduce el NOMBRE VISIBLE de los
        # estilos estándar (ej. "Título 1" en vez de "Heading 1"), aunque su
        # identificador interno se mantenga en inglés. Se normaliza antes de
        # analizar para que las reglas y OxygenAuthor los reconozcan.
        normalized_path = raw_path.replace(".docx", "_NORM.docx")
        renamed_styles = normalize_document_file(raw_path, normalized_path)
        tmp_path = normalized_path
        try:
            os.unlink(raw_path)
        except OSError:
            pass

        if renamed_styles:
            with st.expander(
                "🌐 " + i18n.t("lang_norm_expander", n=len(renamed_styles)),
                expanded=True,
            ):
                st.caption(i18n.t("lang_norm_caption"))
                for old_name, new_name, style_id in renamed_styles:
                    st.write(f"• `{old_name}` → **{new_name}**")

        # ── Paso 2: analizar ──────────────────────────────────────────────────
        st.header(i18n.t("step2_header"))
        st.caption(i18n.t("step2_caption"))

        known_names = get_known_style_names()
        known_list  = sorted([s["style_name"] for s in get_known_styles()])

        # Construir rules_dict desde la base de datos: {style: (target, confidence)}
        all_rules  = get_all_rules()
        rules_dict = {
            r["source_style"]: (r["target_style"], float(r["confidence"]))
            for r in all_rules
        }

        with st.spinner("…"):
            try:
                table_info = count_table_body_paragraphs(tmp_path)
                paragraphs = analyze_document(tmp_path, rules_dict, known_names)
            except ValueError as e:
                st.error(str(e))
                os.unlink(tmp_path)
                st.stop()

        # ── Toggle de conversión de tablas ───────────────────────────────
        apply_table = False
        if table_info["normal_in_table"] > 0:
            n_tbl = table_info["normal_in_table"]
            hint  = i18n.t("table_toggle_recommended") if table_info["suggest_conversion"] else ""
            apply_table = st.toggle(
                i18n.t("table_toggle_label", n=n_tbl) + hint,
                value=table_info["suggest_conversion"],
                help=i18n.t("table_toggle_help"),
            )
            if apply_table:
                with st.spinner("…"):
                    paragraphs = analyze_document(tmp_path, rules_dict, known_names,
                                                  apply_table_context=True)

        # ── Validación de jerarquía de headings ──────────────────────────
        # Informativo: detecta saltos de nivel en la estructura de headings
        # sugerida (p.ej. Heading 3 justo después de Heading 1). No bloquea
        # el flujo — el usuario decide si es intencionado o hay que corregirlo.
        hierarchy_issues = check_heading_hierarchy(paragraphs)
        if hierarchy_issues:
            with st.expander(
                i18n.t("hierarchy_expander", n=len(hierarchy_issues)),
                expanded=False,
            ):
                st.caption(i18n.t("hierarchy_caption"))
                for issue in hierarchy_issues:
                    st.write(
                        "• " + i18n.t(
                            "hierarchy_line",
                            from_level=issue["from_level"],
                            to_level=issue["to_level"],
                            text=issue["text"][:70],
                        )
                    )

        # ── Resumen ───────────────────────────────────────────────────────────
        total    = len(paragraphs)
        n_auto   = sum(1 for p in paragraphs if p["status"] == "auto")
        n_review = sum(1 for p in paragraphs if p["status"] == "review")
        n_unkn   = sum(1 for p in paragraphs if p["status"] == "unknown")
        n_ok     = sum(1 for p in paragraphs if p["status"] == "no_change")

        c1, c2, c3, c4 = st.columns(4)
        c1.metric(i18n.t("metric_no_change"), n_ok)
        c2.metric(i18n.t("metric_auto"),      n_auto)
        c3.metric(i18n.t("metric_review"),    n_review,
                  delta=i18n.t("metric_review_delta", n=n_review) if n_review else None,
                  delta_color="off")
        c4.metric(i18n.t("metric_unknown"),   n_unkn,
                  delta=i18n.t("metric_unknown_delta", n=n_unkn) if n_unkn else None,
                  delta_color="inverse")

        # ── Tabla de revisión ─────────────────────────────────────────────────
        needs_attention = n_review + n_unkn

        if needs_attention == 0 and n_auto > 0:
            st.success(i18n.t("all_auto_success", n=n_auto))

        # Estado del toggle de vista completa
        show_all = st.toggle(
            i18n.t("show_all_toggle", n=total),
            value=False,
            help=i18n.t("show_all_help"),
        )

        # Filtrar filas según la vista
        if show_all:
            display_rows = paragraphs
            st.caption(i18n.t("show_all_caption", n=total))
        else:
            display_rows = [
                p for p in paragraphs
                if p["status"] in ("review", "unknown", "auto")
            ]
            if not display_rows:
                st.info(i18n.t("no_attention_needed"))
            else:
                st.caption(i18n.t("filtered_caption", n=len(display_rows)))

        # Construir DataFrame para el editor
        df = pd.DataFrame([
            {
                "Estado":           i18n.status_label(p["status"]),
                "Texto":            p["text"],
                "Estilo original":  p["original_style"],
                "Estilo final":     p["final_style"],
                "Confianza":        p["confidence_pct"],
                "_idx":             p["idx"],
                "_original":        p["original_style"],
            }
            for p in display_rows
        ])

        # El desplegable debe incluir siempre los valores actuales de cada fila,
        # aunque no estén en la lista de estilos de Oxygen (p.ej. estilos
        # personalizados sin mapeo). Si no se incluyen, Streamlit puede bloquear
        # la edición de esa celda porque el valor actual no está en las opciones.
        extra_values = set()
        for p in display_rows:
            extra_values.add(p["original_style"])
            extra_values.add(p["suggested_style"])
        dropdown_options = sorted(set(known_list) | extra_values)

        st.caption(i18n.t("dropdown_hint"))

        if not df.empty:
            edited_df = st.data_editor(
                df,
                column_config={
                    "Estado": st.column_config.TextColumn(
                        i18n.t("col_estado"), width="small", disabled=True
                    ),
                    "Texto": st.column_config.TextColumn(
                        i18n.t("col_texto"), width="large", disabled=True
                    ),
                    "Estilo original": st.column_config.TextColumn(
                        i18n.t("col_estilo_original"), width="medium", disabled=True
                    ),
                    "Estilo final": st.column_config.SelectboxColumn(
                        i18n.t("col_estilo_final"),
                        width="medium",
                        options=[i18n.t("opt_no_change")] + dropdown_options,
                        required=True,
                        help=i18n.t("col_estilo_final_help"),
                    ),
                    "Confianza": st.column_config.ProgressColumn(
                        i18n.t("col_confianza"), width="small",
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

        # ── Guardar nuevas reglas desde la tabla ──────────────────────────────
        no_change_label = i18n.t("opt_no_change")
        if not edited_df.empty:
            changed_rules = edited_df[
                (edited_df["Estilo final"] != no_change_label) &
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
                    st.info(i18n.t("rules_saved_info", n=new_rule_count))

        # ── Paso 3: procesar ──────────────────────────────────────────────────
        st.header(i18n.t("step3_header"))
        st.caption(i18n.t("step3_caption"))

        if st.button(i18n.t("process_button"), type="primary"):

            # Construir lista de decisiones desde la tabla editada
            # Combinar: filas visibles (editadas) + filas ocultas (si vista filtrada)
            decisions_from_table = {}
            if not edited_df.empty:
                for _, row in edited_df.iterrows():
                    final = row["Estilo final"]
                    if final == no_change_label:
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

            with st.spinner("…"):
                n_changes, n_errors = apply_paragraph_decisions(
                    tmp_path, output_path, all_decisions
                )

            with open(output_path, "rb") as f:
                output_bytes = f.read()

            base_name = uploaded.name.rsplit(".", 1)[0]
            out_name = f"{base_name}_ESTILOS_CORREGIDOS.docx"

            msg = i18n.t("process_success", n=n_changes)
            if n_errors:
                msg += i18n.t("process_errors", n=n_errors)
            st.success(msg)
            st.caption(i18n.t("process_visual_note"))

            st.download_button(
                label=i18n.t("download_button"),
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
    elif page == i18n.t("nav_reglas"):

        st.title(i18n.t("page2_title"))
        st.caption(i18n.t("page2_caption"))

        # ── Indicador de persistencia (GitHub backup) ─────────────────────────
        if github_sync and github_sync.is_configured():
            st.success(i18n.t("backup_active"))
        else:
            st.warning(i18n.t("backup_inactive"))

        rules = get_all_rules()

        if not rules:
            st.info(i18n.t("page2_no_rules"))
            st.stop()

        st.write(i18n.t("page2_total", n=len(rules)))

        df = pd.DataFrame(rules)[[
            "id", "source_style", "source_element", "target_style",
            "confidence", "times_applied", "rule_origin", "created_at"
        ]]
        df.columns = [
            i18n.t("col_id"), i18n.t("col_origen"), i18n.t("col_tipo"),
            i18n.t("col_destino"), i18n.t("col_confianza"), i18n.t("col_veces"),
            i18n.t("col_rule_origin"), i18n.t("col_creado"),
        ]
        conf_col = i18n.t("col_confianza")
        df[conf_col] = (df[conf_col] * 100).round(0).astype(int).astype(str) + "%"

        st.dataframe(df, use_container_width=True, hide_index=True)

        st.divider()
        st.subheader(i18n.t("delete_rule_subheader"))
        rule_id = st.number_input(i18n.t("delete_rule_label"), min_value=1, step=1)
        if st.button(i18n.t("delete_rule_button"), type="secondary"):
            delete_rule(int(rule_id))
            st.success(i18n.t("delete_rule_success", n=rule_id))
            st.rerun()

    # ══════════════════════════════════════════════════════════════════════════════
    # PÁGINA 3 — Estilos de Oxygen
    # ══════════════════════════════════════════════════════════════════════════════
    elif page == i18n.t("nav_estilos"):

        st.title(i18n.t("page3_title"))
        st.caption(i18n.t("page3_caption"))

        known = get_known_styles()
        if not known:
            st.error(i18n.t("page3_error_no_styles"))
            st.stop()

        df = pd.DataFrame(known)[["element_type", "style_name", "target_html", "creates_new_block"]]
        df.columns = [
            i18n.t("col_tipo"), i18n.t("col_nombre_estilo"),
            i18n.t("col_elemento_html"), i18n.t("col_crea_bloque"),
        ]

        search = st.text_input(i18n.t("page3_search"))
        if search:
            name_col = i18n.t("col_nombre_estilo")
            df = df[df[name_col].str.contains(search, case=False, na=False)]

        st.dataframe(df, use_container_width=True, hide_index=True)
        st.caption(i18n.t("page3_total", n=len(df)))
