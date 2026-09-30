"""
Lógica compartida del extractor PISE:
  1) PDF -> texto (texto embebido si existe; si no, OCR con Tesseract)
  2) texto -> catálogo de ítems (sección, código, tarea, tipo, mín, máx, unidad)
  3) catálogo + multiplicidades (bogies/ejes/ruedas) -> Excel con la misma
     estructura que PISE-SGBV-022_Rev03_Digitalizado.xlsx

No importa streamlit acá: así se puede probar por consola o reusar en otro
proyecto (por ejemplo, para cargar directo a la base SQLite del ISMR).
"""
from __future__ import annotations
import re
import subprocess
import tempfile
from dataclasses import dataclass, field
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule

# --------------------------------------------------------------- PDF -> texto

def pdf_a_texto_por_pagina(pdf_bytes: bytes, idioma: str = "spa+eng") -> list[str]:
    """Devuelve una lista con el texto de cada página.

    Primero intenta `pdftotext` (rápido, exacto) por si el PDF tiene texto
    embebido. Si una página sale vacía, la rasteriza con `pdftoppm` y le hace
    OCR con Tesseract — es el caso de los PISE escaneados por fotocopiadora.
    """
    import pytesseract
    from PIL import Image

    with tempfile.TemporaryDirectory() as tmp:
        tmp = Path(tmp)
        pdf_path = tmp / "in.pdf"
        pdf_path.write_bytes(pdf_bytes)

        # 1) texto embebido, página por página
        textos = []
        salida = tmp / "texto"
        subprocess.run(["pdftotext", "-layout", str(pdf_path), str(salida) + ".txt"],
                        check=False, capture_output=True)
        # pdftotext no separa por página con -layout de forma simple; usamos
        # el flag -f/-l para pedir página por página y así detectar cuáles
        # están vacías (sin texto -> escaneadas).
        n_paginas = _contar_paginas(pdf_path)
        for i in range(1, n_paginas + 1):
            r = subprocess.run(
                ["pdftotext", "-layout", "-f", str(i), "-l", str(i), str(pdf_path), "-"],
                capture_output=True, text=True,
            )
            textos.append(r.stdout or "")

        # 2) OCR para las páginas sin texto (o con muy poco texto = ruido)
        for i, t in enumerate(textos, start=1):
            if len(t.strip()) >= 40:
                continue
            img_prefix = tmp / f"pag_{i}"
            subprocess.run(
                ["pdftoppm", "-png", "-r", "300", "-f", str(i), "-l", str(i),
                 str(pdf_path), str(img_prefix)],
                check=False, capture_output=True,
            )
            candidatos = sorted(tmp.glob(f"pag_{i}*.png"))
            if not candidatos:
                continue
            img = Image.open(candidatos[0])
            try:
                ocr = pytesseract.image_to_string(img, lang=idioma)
            except Exception:
                # si falta el paquete de idioma "spa", cae a inglés
                ocr = pytesseract.image_to_string(img, lang="eng")
            textos[i - 1] = ocr
        return textos


def _contar_paginas(pdf_path: Path) -> int:
    r = subprocess.run(["pdfinfo", str(pdf_path)], capture_output=True, text=True)
    for linea in r.stdout.splitlines():
        if linea.startswith("Pages:"):
            return int(linea.split(":")[1].strip())
    return 1


# --------------------------------------------------------------- parseo

RE_SECCION = re.compile(
    r"^\s*(\d{1,2})\.\s*([A-ZÁÉÍÓÚÑ][A-ZÁÉÍÓÚÑ0-9°\.\,\s/()\-]{4,})$"
)
RE_ITEM = re.compile(r"^\s*(\d{1,2}\.\d{1,2}(?:\.\d{1,2})?)[\.\)\s]+(.{6,})$")
RE_NUM = r"[-+]?\d{1,4}(?:[.,]\d{1,3})?"
PATRONES_TOLERANCIA = [
    # "1600 ≤ Ai ≤ 1604 MM"  /  "49 ≤ Pd + Pi ≤ 63.6 MM"
    (re.compile(rf"({RE_NUM})\s*[≤<]=?\s*.{{0,25}}?\s*[≤<]=?\s*({RE_NUM})\s*(MM|V|%)?", re.I),
     lambda m: (float(m.group(1).replace(",", ".")), float(m.group(2).replace(",", ".")), m.group(3))),
    # "≥ 25 MM" / ">= 25 MM"
    (re.compile(rf"[≥>]=?\s*({RE_NUM})\s*(MM|V|%)?", re.I),
     lambda m: (float(m.group(1).replace(",", ".")), None, m.group(2))),
    # "≤ 34 MM" / "MÁX 5 MM" / "MÁX: 5 MM"
    (re.compile(rf"(?:[≤<]=?|M[ÁA]X\.?:?)\s*({RE_NUM})\s*(MM|V|%)?", re.I),
     lambda m: (None, float(m.group(1).replace(",", ".")), m.group(2))),
    # "MÍN 42 MM" / "MÍN: 42 MM"
    (re.compile(rf"M[ÍI]N\.?:?\s*({RE_NUM})\s*(MM|V|%)?", re.I),
     lambda m: (float(m.group(1).replace(",", ".")), None, m.group(2))),
    # "60 MM +/-3 MM" / "60 MM +/- 3 MM"
    (re.compile(rf"({RE_NUM})\s*(MM|V|%)?\s*\+/?-\s*({RE_NUM})", re.I),
     lambda m: (float(m.group(1).replace(",", ".")) - float(m.group(3).replace(",", ".")),
                float(m.group(1).replace(",", ".")) + float(m.group(3).replace(",", ".")),
                m.group(2))),
]
PALABRAS_MEDICION = ("MEDIR", "CALCULAR", "CONTROLAR COTA", "CONTROLAR ENTRECENTRO",
                     "CONTROLAR HUELGO", "CONTROLAR TENSIÓN", "CONTROLAR TENSION")
PALABRAS_ALCANCE = {
    "RUEDA": "rueda", "RUEDAS": "rueda",
    "EJE": "eje", "EJES": "eje",
    "BOGIE": "bogie", "BOGIES": "bogie",
}


def _detectar_tolerancia(desc: str):
    for patron, extraer in PATRONES_TOLERANCIA:
        m = patron.search(desc)
        if m:
            mn, mx, u = extraer(m)
            return mn, mx, (u or "mm").lower() if (mn is not None or mx is not None) else None
    return None, None, None


def _detectar_tipo(desc: str) -> str:
    d = desc.upper()
    if any(p in d for p in PALABRAS_MEDICION):
        return "MEDICION"
    return "VISUAL"


def _detectar_alcance(desc: str, ultimo_alcance: str) -> str:
    d = desc.upper()
    for palabra, alcance in PALABRAS_ALCANCE.items():
        if re.search(rf"\b{palabra}\b", d):
            return alcance
    return ultimo_alcance  # hereda de la sección/ítem anterior si no hay pista


def parsear_catalogo(paginas_texto: list[str]) -> pd.DataFrame:
    """Convierte el texto OCR/embebido en una tabla de ítems.

    Heurística deliberadamente simple y auditable (nada de IA/ML): busca
    líneas que empiecen con un código tipo '2.12.1', junta líneas siguientes
    que no sean un nuevo ítem/sección (continuación de la misma frase), y les
    aplica regex de tolerancia. Pensada para revisarse y corregirse a mano
    en la tabla editable de la app antes de generar el Excel.
    """
    filas = []
    seccion_actual = ""
    alcance_actual = "bogie"
    item_actual = None

    def cerrar_item():
        if item_actual is None:
            return
        desc = " ".join(item_actual["partes"]).strip()
        desc = re.sub(r"\s+", " ", desc)
        mn, mx, u = _detectar_tolerancia(desc)
        tipo = _detectar_tipo(desc)
        alcance = _detectar_alcance(desc, item_actual["alcance_heredado"])
        filas.append(dict(
            seccion=item_actual["seccion"], item=item_actual["codigo"],
            tarea=desc, tipo=tipo, minimo=mn, maximo=mx,
            unidad=(u or ("mm" if tipo == "MEDICION" else "")),
            alcance=alcance,
        ))

    for texto in paginas_texto:
        for linea_cruda in texto.splitlines():
            linea = linea_cruda.strip()
            if not linea:
                continue
            m_sec = RE_SECCION.match(linea)
            if m_sec and len(m_sec.group(2).split()) >= 2:
                cerrar_item()
                item_actual = None
                seccion_actual = f"{m_sec.group(1)}. {m_sec.group(2).strip().rstrip('.')}"
                continue
            m_item = RE_ITEM.match(linea)
            if m_item:
                cerrar_item()
                item_actual = dict(
                    seccion=seccion_actual, codigo=m_item.group(1),
                    partes=[m_item.group(2)],
                    alcance_heredado=alcance_actual,
                )
                continue
            if item_actual is not None and not linea[0].isdigit():
                # continuación de la frase del ítem anterior (el OCR corta líneas)
                item_actual["partes"].append(linea)
        cerrar_item()
        item_actual = None
    df = pd.DataFrame(filas, columns=[
        "seccion", "item", "tarea", "tipo", "minimo", "maximo", "unidad", "alcance"
    ])
    return df.drop_duplicates(subset=["item", "tarea"]).reset_index(drop=True)


# --------------------------------------------------------------- Excel

ARIAL = "Arial"
F_NORMAL = Font(name=ARIAL, size=10)
F_NEGRITA = Font(name=ARIAL, size=10, bold=True)
F_TITULO = Font(name=ARIAL, size=14, bold=True, color="1F3864")
F_HEADER = Font(name=ARIAL, size=10, bold=True, color="FFFFFF")
FILL_HEADER = PatternFill("solid", fgColor="1F3864")
FILL_INPUT = PatternFill("solid", fgColor="FFF2CC")
FILL_BLOQUE = PatternFill("solid", fgColor="D9E1F2")
FILL_OK = PatternFill("solid", fgColor="C6EFCE")
FILL_MAL = PatternFill("solid", fgColor="FFC7CE")
FILL_NA = PatternFill("solid", fgColor="EDEDED")
FINO = Side(style="thin", color="A6A6A6")
BORDE = Border(left=FINO, right=FINO, top=FINO, bottom=FINO)
WRAP = Alignment(wrap_text=True, vertical="top")
CENTRO = Alignment(horizontal="center", vertical="top", wrap_text=True)


def _encabezado(ws, fila, textos, anchos=None):
    for i, t in enumerate(textos, start=1):
        c = ws.cell(row=fila, column=i, value=t)
        c.font, c.fill, c.border = F_HEADER, FILL_HEADER, BORDE
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    if anchos:
        for i, a in enumerate(anchos, start=1):
            ws.column_dimensions[chr(64 + i)].width = a


def _posiciones(alcance: str, n_bogies: int, n_ejes_por_bogie: int, n_ruedas_por_eje: int):
    """Expande un ítem con alcance rueda/eje/bogie/general en filas de
    relevamiento (bogie, eje, posición), según las multiplicidades del coche."""
    alcance = (alcance or "bogie").lower()
    filas = []
    if alcance == "general":
        return [("Coche", "", "General")]
    for b in range(1, n_bogies + 1):
        if alcance == "bogie":
            filas.append((f"Bogie {b}", "", "Bogie completo"))
            continue
        eje_ini = (b - 1) * n_ejes_por_bogie + 1
        for e in range(eje_ini, eje_ini + n_ejes_por_bogie):
            if alcance == "eje":
                filas.append((f"Bogie {b}", f"Eje {e}", "Eje completo"))
                continue
            r_ini = (e - 1) * n_ruedas_por_eje + 1
            lados = ["D", "I"] if n_ruedas_por_eje == 2 else [str(k) for k in range(1, n_ruedas_por_eje + 1)]
            for k, lado in enumerate(lados):
                filas.append((f"Bogie {b}", f"Eje {e}", f"Rueda {r_ini + k} ({lado})"))
    return filas


def generar_excel(
    catalogo_df: pd.DataFrame,
    ruta_salida: str,
    titulo: str = "Protocolo digitalizado",
    n_bogies: int = 2,
    n_ejes_por_bogie: int = 1,
    n_ruedas_por_eje: int = 2,
    campos_cabecera: dict[str, list[str]] | None = None,
):
    """Genera el Excel con la misma estructura del ejemplo:
    LEEME, CABECERA, RELEVAMIENTO, RESUMEN, CATALOGO.

    catalogo_df necesita las columnas: seccion, item, tarea, tipo, minimo,
    maximo, unidad, alcance (alcance en {rueda, eje, bogie, general}).
    """
    wb = Workbook()

    # ---- LEEME
    ws = wb.active
    ws.title = "LEEME"
    ws["A1"] = titulo
    ws["A1"].font = F_TITULO
    lineas = [
        "Generado automáticamente a partir de un PISE escaneado (OCR + revisión manual).",
        "",
        "1) CABECERA: completar los datos del coche (celdas amarillas).",
        "2) RELEVAMIENTO: completar solo la columna RESULTADO (amarilla) por posición.",
        "     • Ítems VISUALES: OK / X / N/A. Si es X, detallar en OBSERVACIÓN.",
        "     • Ítems de MEDICIÓN: cargar el número medido. ESTADO se calcula solo.",
        "3) RESUMEN: avance y conteo por sección (fórmulas, no tocar).",
        "4) CATALOGO: ítems, tolerancias y alcance tal como se revisaron antes de generar.",
        "",
        "IMPORTANTE: este catálogo salió de OCR sobre un escaneo y de la revisión que hiciste",
        "en la app antes de descargar el archivo. Si algo quedó mal interpretado, se corrige",
        "en la hoja CATALOGO y no hace falta volver a correr el OCR.",
    ]
    for i, t in enumerate(lineas, start=3):
        c = ws.cell(row=i, column=1, value=t)
        c.font = F_NORMAL
    ws.column_dimensions["A"].width = 115

    # ---- CABECERA
    wc = wb.create_sheet("CABECERA")
    wc["A1"] = "DATOS DEL COCHE E INFORME"
    wc["A1"].font = F_TITULO
    _encabezado(wc, 3, ["Bloque", "Campo", "Valor"], [26, 46, 55])
    campos_cabecera = campos_cabecera or {
        "Encabezado": ["LÍNEA", "LUGAR", "MARCA", "INSPECCIÓN Nº", "VEHÍCULO",
                       "INFORME Nº", "FECHA DE INSP."],
        "Cierre": ["C. REGISTROS FOTOGRÁFICOS (referencia)", "D. CONCLUSIONES",
                   "E. RECOMENDACIONES FINALES"],
    }
    fila = 4
    for bloque, lista in campos_cabecera.items():
        for campo in lista:
            a, b, v = (wc.cell(row=fila, column=k) for k in (1, 2, 3))
            a.value, b.value = bloque, campo
            for x in (a, b, v):
                x.font, x.border, x.alignment = F_NORMAL, BORDE, WRAP
            a.fill = FILL_BLOQUE
            v.fill = FILL_INPUT
            fila += 1
    wc.freeze_panes = "A4"

    # ---- CATALOGO
    wk = wb.create_sheet("CATALOGO")
    wk["A1"] = "CATÁLOGO DE ÍTEMS Y TOLERANCIAS (revisado antes de generar)"
    wk["A1"].font = F_TITULO
    _encabezado(wk, 4, ["Sección", "Ítem", "Tarea a realizar", "Tipo", "Mín", "Máx",
                        "Unidad", "Alcance"], [30, 8, 80, 11, 9, 9, 8, 10])
    for i, row in enumerate(catalogo_df.itertuples(index=False), start=5):
        vals = [row.seccion, row.item, row.tarea, row.tipo, row.minimo, row.maximo,
                row.unidad, row.alcance]
        for j, v in enumerate(vals, start=1):
            c = wk.cell(row=i, column=j, value=v)
            c.font, c.border, c.alignment = F_NORMAL, BORDE, WRAP
    ult_cat = 4 + len(catalogo_df)
    wk.freeze_panes = "D5"
    wk.auto_filter.ref = f"A4:H{ult_cat}"

    # ---- RELEVAMIENTO
    wr = wb.create_sheet("RELEVAMIENTO", 2)
    wr["A1"] = "RELEVAMIENTO — REGISTROS DE CONTROL"
    wr["A1"].font = F_TITULO
    wr["A2"] = "Completar solo la columna RESULTADO (amarilla)."
    wr["A2"].font = F_NORMAL
    _encabezado(wr, 4, ["N°", "Sección", "Ítem", "Tarea a realizar", "Tipo",
                        "Bogie", "Eje", "Posición", "RESULTADO", "Mín", "Máx",
                        "Unidad", "ESTADO", "Observación"],
                [6, 30, 8, 70, 11, 12, 8, 16, 14, 9, 9, 8, 17, 40])
    rng = lambda col: f"CATALOGO!${col}$5:${col}${ult_cat}"
    dv = DataValidation(type="list", formula1='"OK,X,N/A"', allow_blank=True)
    wr.add_data_validation(dv)

    n = 0
    for row in catalogo_df.itertuples(index=False):
        posiciones = _posiciones(row.alcance, n_bogies, n_ejes_por_bogie, n_ruedas_por_eje)
        for bogie, eje, pos in posiciones:
            n += 1
            f = 4 + n
            match = f"MATCH($C{f},{rng('B')},0)"
            vals = [n, row.seccion, row.item, row.tarea, row.tipo, bogie, eje, pos, None,
                    f'=IF(INDEX({rng("E")},{match})="","",INDEX({rng("E")},{match}))',
                    f'=IF(INDEX({rng("F")},{match})="","",INDEX({rng("F")},{match}))',
                    f'=IF(INDEX({rng("G")},{match})="","",INDEX({rng("G")},{match}))',
                    (f'=IF(I{f}="","",IF(I{f}="N/A","N/A",IF(E{f}="MEDICION",'
                     f'IF(ISNUMBER(I{f}),IF(AND(J{f}="",K{f}=""),"CARGADO",'
                     f'IF(AND(OR(J{f}="",I{f}>=J{f}),OR(K{f}="",I{f}<=K{f})),"OK","FUERA DE RANGO")),"REVISAR"),'
                     f'IF(OR(I{f}="OK",I{f}="X"),I{f},"REVISAR"))))'),
                    None]
            for j, v in enumerate(vals, start=1):
                c = wr.cell(row=f, column=j, value=v)
                c.font, c.border = F_NORMAL, BORDE
                c.alignment = CENTRO if j in (1, 3, 5, 6, 7, 9, 10, 11, 12, 13) else WRAP
            wr.cell(row=f, column=9).fill = FILL_INPUT
            if row.tipo == "VISUAL":
                dv.add(wr.cell(row=f, column=9))
    ult_rel = 4 + n
    wr.freeze_panes = "D5"
    wr.auto_filter.ref = f"A4:N{ult_rel}"
    zona = f"M5:M{ult_rel}"
    wr.conditional_formatting.add(zona, CellIsRule(operator="equal", formula=['"OK"'], fill=FILL_OK))
    for val in ('"X"', '"FUERA DE RANGO"', '"REVISAR"'):
        wr.conditional_formatting.add(zona, CellIsRule(operator="equal", formula=[val], fill=FILL_MAL))
    wr.conditional_formatting.add(zona, CellIsRule(operator="equal", formula=['"N/A"'], fill=FILL_NA))

    # ---- RESUMEN
    wm = wb.create_sheet("RESUMEN", 3)
    wm["A1"] = "RESUMEN POR SECCIÓN"
    wm["A1"].font = F_TITULO
    _encabezado(wm, 3, ["Sección", "Puntos totales", "Cargados", "OK", "X",
                        "Fuera de rango", "N/A", "Pendientes", "Avance"],
                [34, 14, 12, 10, 10, 15, 10, 12, 10])
    secciones = list(dict.fromkeys(catalogo_df["seccion"]))  # orden de aparición, sin duplicar
    RS = f"RELEVAMIENTO!$B$5:$B${ult_rel}"
    RI = f"RELEVAMIENTO!$I$5:$I${ult_rel}"
    RM = f"RELEVAMIENTO!$M$5:$M${ult_rel}"
    for i, nombre in enumerate(secciones, start=4):
        fs = [nombre, f"=COUNTIF({RS},$A{i})", f'=COUNTIFS({RS},$A{i},{RI},"<>")',
              f'=COUNTIFS({RS},$A{i},{RM},"OK")', f'=COUNTIFS({RS},$A{i},{RM},"X")',
              f'=COUNTIFS({RS},$A{i},{RM},"FUERA DE RANGO")',
              f'=COUNTIFS({RS},$A{i},{RM},"N/A")', f"=B{i}-C{i}",
              f"=IF(B{i}=0,0,C{i}/B{i})"]
        for j, v in enumerate(fs, start=1):
            c = wm.cell(row=i, column=j, value=v)
            c.font, c.border = F_NORMAL, BORDE
            c.alignment = Alignment(horizontal="left" if j == 1 else "center")
        wm.cell(row=i, column=9).number_format = "0%"
    tot = 4 + len(secciones)
    wm.cell(row=tot, column=1, value="TOTAL")
    for j in range(2, 9):
        L = chr(64 + j)
        wm.cell(row=tot, column=j, value=f"=SUM({L}4:{L}{tot - 1})")
    wm.cell(row=tot, column=9, value=f"=IF(B{tot}=0,0,C{tot}/B{tot})").number_format = "0%"
    for j in range(1, 10):
        c = wm.cell(row=tot, column=j)
        c.font, c.border, c.fill = F_NEGRITA, BORDE, FILL_BLOQUE
        c.alignment = Alignment(horizontal="left" if j == 1 else "center")

    wb.save(ruta_salida)
    return ruta_salida
