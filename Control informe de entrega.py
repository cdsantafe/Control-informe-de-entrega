# ════════════════════════════════════════════════════════════════════════
    # SECCIONES ADICIONALES ABAJO
    # ════════════════════════════════════════════════════════════════════════
    st.markdown("---")
    st.subheader("📌 Resumen Operativo Adicional")

    col_izq, col_der = st.columns(2)

    with col_izq:
        st.markdown("### 🪞 Patentes Duplicadas en el Informe")
        if df_dups.empty:
            st.success("No se detectaron patentes duplicadas en la misma vuelta.")
        else:
            st.warning(f"Se encontraron {len(df_dups)} patente(s) duplicadas.")
            st.dataframe(df_dups, hide_index=True)

    with col_der:
        st.markdown("### 🟢 Patentes Disponibles Sin Utilizar (100% Libres)")
        if df_sin_usar.empty:
            st.info("Todas las unidades disponibles fueron utilizadas al menos en una planilla.")
        else:
            st.success(f"{len(df_sin_usar)} unidades disponibles no fueron asignadas en todo el día.")
            st.dataframe(df_sin_usar, hide_index=True)
