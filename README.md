# Extractor PISE (escaneado) → Excel

App de Streamlit para digitalizar PISE de Bureau Veritas que llegan como PDF
**escaneado** (sin texto embebido, como el PISE-SGBV-022). Sube el PDF, hace
OCR, arma un catálogo de ítems/tolerancias, lo deja editar en pantalla y
genera un Excel con la misma estructura que
`PISE-SGBV-022_Rev03_Digitalizado.xlsx` (CABECERA, CATALOGO, RELEVAMIENTO,
RESUMEN, LEEME).

## Archivos
- `app.py` — interfaz Streamlit.
- `pise_lib.py` — OCR, parseo por regex y generación del Excel (sin depender de Streamlit; se puede reusar en otro script o en el pipeline del ISMR).
- `requirements.txt` — dependencias Python.
- `packages.txt` — dependencias de sistema para Streamlit Cloud (Tesseract + poppler). **Sin esto no anda el OCR en Streamlit Cloud.**

## Correr local
```bash
pip install -r requirements.txt
sudo apt install tesseract-ocr tesseract-ocr-spa poppler-utils   # Linux
streamlit run app.py
```

## Límite importante: tablas densas
Tesseract OCR lee bien párrafos de texto corrido, pero **las tablas con
muchas columnas angostas** (como "1. PARES MONTADOS RUEDAS", con 4 columnas
de rueda por bogie) tienden a mezclarse: el motor a veces junta celdas de
filas distintas en una sola línea. Por eso el catálogo que arma la app es un
**punto de partida para revisar**, no un resultado listo para usar sin mirar.
Si notás que faltan ítems o que el texto de la tarea se ve cortado/mezclado,
conviene:
1. Mirar el texto OCR crudo (expander de la app) para esa página.
2. Corregir a mano en la tabla editable (agregar la fila que faltó, separar
   una tarea que quedó pegada a otra).

Para escaneos de mejor calidad (300+ dpi, sin sello/firma tapando texto) la
detección mejora bastante.

## Próximos pasos posibles
- Sumar reconocimiento de checkboxes/manuscrito para PISE ya completados
  (esto es harina de otro costal: hoy la app está pensada para digitalizar
  la **estructura** del protocolo, no para leer resultados escritos a mano).
- Exportar el catálogo directo a la tabla `ref_sistemas`/`inspecciones_estaticas`
  de la base SQLite del proyecto ISMR en lugar de (o además de) al Excel.
