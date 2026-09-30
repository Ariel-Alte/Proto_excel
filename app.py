"""
Extractor PISE (escaneados) -> Excel digitalizado
==================================================
Sube un PISE en PDF (Bureau Veritas), le hace OCR, detecta ítems/tolerancias
y genera un Excel con la misma estructura que PISE-SGBV-022_Rev03_Digitalizado.xlsx.

Correr localmente:
    streamlit run app.py

Deploy en Streamlit Cloud: subir esta carpeta a un repo con requirements.txt
y packages.txt (quedan más abajo) y apuntar la app a app.py.
"""
import io
from datetime import date

import pandas as pd
import streamlit as st

from pise_lib import pdf_a_texto_por_pagina, parsear_catalogo, generar_excel

st.set_page_config(page_title="Extractor PISE → Excel", layout="wide")
st.title("📋 Extractor PISE (escaneado) → Excel")
st.caption(
    "Pensado para PISE de Bureau Veritas escaneados (sin texto embebido), como el "
    "PISE-SGBV-022. Hace OCR, propone un catálogo de ítems y te deja corregirlo antes "
    "de armar el Excel — el OCR nunca es perfecto sobre un escaneo de fotocopiadora."
)

if "catalogo" not in st.session_state:
    st.session_state.catalogo = None
if "titulo" not in st.session_state:
    st.session_state.titulo = "Protocolo digitalizado"

# --------------------------------------------------------------- 1) subir PDF
st.header("1) Subir el PDF")
archivo = st.file_uploader("PISE en PDF", type=["pdf"])

col1, col2 = st.columns(2)
with col1:
    idioma_ocr = st.selectbox(
        "Idioma para OCR", ["spa+eng", "spa", "eng"], index=0,
        help="Si la app tira error de idioma 'spa' no instalado, probá 'eng' "
             "(revisá igual el catálogo, van a salir más errores de tildes).",
    )
with col2:
    titulo = st.text_input("Título del protocolo (para la hoja LEEME)",
                           value=st.session_state.titulo)

if archivo and st.button("🔍 Extraer ítems (OCR)", type="primary"):
    with st.spinner("Leyendo el PDF y haciendo OCR — puede tardar según la cantidad de páginas..."):
        paginas = pdf_a_texto_por_pagina(archivo.getvalue(), idioma=idioma_ocr)
        catalogo = parsear_catalogo(paginas)
    if catalogo.empty:
        st.error(
            "No se detectaron ítems. Puede ser que el OCR no reconoció los códigos "
            "(1.1, 2.3, etc.) por baja calidad del escaneo, o que el formato de "
            "numeración es distinto al esperado. Revisá el texto crudo más abajo."
        )
        with st.expander("Ver texto OCR crudo, página por página"):
            for i, t in enumerate(paginas, start=1):
                st.text_area(f"Página {i}", t, height=200)
    else:
        st.session_state.catalogo = catalogo
        st.session_state.titulo = titulo
        st.success(f"Se detectaron {len(catalogo)} ítems. Revisalos en la tabla de abajo.")

# --------------------------------------------------------------- 2) revisar
if st.session_state.catalogo is not None:
    st.header("2) Revisar y corregir el catálogo detectado")
    st.markdown(
        "- **tipo**: `VISUAL` (OK/X/N/A) o `MEDICION` (número + mín/máx).\n"
        "- **alcance**: cuántas veces se repite el ítem en el coche — "
        "`rueda` (una fila por rueda), `eje` (una por eje), `bogie` (una por bogie) "
        "o `general` (una sola vez, para todo el coche).\n"
        "- Agregá o borrá filas con los botones de la tabla si el OCR se comió o "
        "duplicó algún ítem."
    )
    catalogo_editado = st.data_editor(
        st.session_state.catalogo,
        num_rows="dynamic",
        use_container_width=True,
        height=420,
        column_config={
            "tipo": st.column_config.SelectboxColumn(options=["VISUAL", "MEDICION"]),
            "alcance": st.column_config.SelectboxColumn(
                options=["rueda", "eje", "bogie", "general"]),
            "minimo": st.column_config.NumberColumn(format="%.2f"),
            "maximo": st.column_config.NumberColumn(format="%.2f"),
        },
        key="editor_catalogo",
    )

    st.header("3) Multiplicidades del coche")
    st.caption(
        "Con esto se calcula cuántas filas de relevamiento salen de cada ítem "
        "según su alcance (p. ej. 2 bogies × 1 eje × 2 ruedas = 4 ruedas por bogie)."
    )
    c1, c2, c3 = st.columns(3)
    n_bogies = c1.number_input("Bogies por coche", min_value=1, max_value=4, value=2)
    n_ejes = c2.number_input("Ejes por bogie", min_value=1, max_value=4, value=1)
    n_ruedas = c3.number_input("Ruedas por eje", min_value=1, max_value=4, value=2)

    total_filas = 0
    for _, r in catalogo_editado.iterrows():
        alc = str(r.get("alcance", "bogie")).lower()
        if alc == "general":
            total_filas += 1
        elif alc == "bogie":
            total_filas += n_bogies
        elif alc == "eje":
            total_filas += n_bogies * n_ejes
        else:
            total_filas += n_bogies * n_ejes * n_ruedas
    st.info(f"El Excel va a tener **{total_filas} filas** en RELEVAMIENTO.")

    st.header("4) Generar el Excel")
    nombre_salida = st.text_input(
        "Nombre del archivo",
        value=f"PISE_Digitalizado_{date.today().isoformat()}.xlsx",
    )
    if st.button("📥 Generar Excel", type="primary"):
        buffer_path = "/tmp/_pise_generado.xlsx"
        generar_excel(
            catalogo_editado.rename(columns=str.lower),
            buffer_path,
            titulo=st.session_state.titulo,
            n_bogies=int(n_bogies),
            n_ejes_por_bogie=int(n_ejes),
            n_ruedas_por_eje=int(n_ruedas),
        )
        with open(buffer_path, "rb") as f:
            datos = f.read()
        st.success("Listo. Descargalo acá abajo:")
        st.download_button(
            "⬇️ Descargar Excel", data=datos, file_name=nombre_salida,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

st.divider()
with st.expander("¿Cómo funciona por adentro?"):
    st.markdown(
        "1. `pdftotext` intenta sacar texto embebido de cada página.\n"
        "2. Si una página sale casi vacía (PDF escaneado), se rasteriza con "
        "`pdftoppm` a 300 dpi y se le corre `pytesseract` (Tesseract OCR).\n"
        "3. Se buscan líneas con patrón de sección (`1. PARES MONTADOS...`) y de "
        "ítem (`1.1 MEDIR...`), y se arma la tarea completa juntando líneas de "
        "continuación.\n"
        "4. Regex busca patrones de tolerancia (`≥`, `≤`, `MÍN`, `MÁX`, `+/-`, "
        "rangos `a ≤ X ≤ b`) para clasificar el ítem como VISUAL o MEDICIÓN y "
        "sacar mínimo/máximo/unidad.\n"
        "5. Todo eso queda en una tabla editable — el OCR se equivoca con tildes, "
        "ceros/oes y saltos de línea raros, así que esta revisión es el paso que "
        "garantiza que el Excel final sea correcto.\n"
        "6. Con el catálogo aprobado + las multiplicidades del coche se genera el "
        "Excel: CATALOGO (fuente única), RELEVAMIENTO (una fila por ítem × "
        "posición, con fórmulas de ESTADO), RESUMEN (conteo por sección) y "
        "CABECERA (datos generales)."
    )
