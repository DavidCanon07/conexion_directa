"""
exportador.py
Escribe DataFrames a .xlsx con formato legible: encabezado resaltado,
ancho de columna automático, encabezado congelado y formato de moneda en
columnas "monto"/"valor".

Usa xlsxwriter.Workbook(..., {"constant_memory": True}) en vez de
openpyxl: con hojas de cientos de miles/millones de filas, escribir
celda por celda es el cuello de botella real de todo el programa, más
lento que leer y parsear el archivo plano de origen. Se comparó
directamente contra openpyxl.Workbook(write_only=True) (el motor
anterior, ya optimizado con muestreo de anchos e indices_formato) con un
benchmark sobre datos sintéticos del mismo shape real (55 columnas,
~1.75M filas): XlsxWriter fue 2.52x más rápido (348s vs 878s). El modo
constant_memory transmite cada fila directamente a disco sin mantener el
worksheet completo en memoria — igual que write_only en openpyxl — a
costa de la misma restricción: las celdas de una fila deben escribirse en
orden de columna estrictamente creciente, sin poder volver atrás.

Las columnas "fecha" YA NO reciben number_format: siguen siendo texto
(lector.py nunca las convierte a un tipo fecha real), así que ese formato
no tenía ningún efecto visual en Excel — solo costaba tiempo de escritura
sin beneficio. Solo "monto"/"valor" reciben formato de moneda.

Cualquier hoja tipo DataFrame que supere LIMITE_FILAS_HOJA filas se
particiona automáticamente en varias hojas consecutivas antes de
escribirse (ver _particionar_hoja_grande) — Excel tiene un límite real de
1,048,576 filas por hoja; sin esto, un resultado de negocio de más de un
millón de filas no cabría en una sola hoja y el .xlsx quedaría corrupto o
Excel se saturaría al abrirlo.

Se habilita use_zip64() en cada workbook: sin esto, XlsxWriter falla con
"Filesize would require ZIP64 extensions" en cuanto el .xlsx comprimido
supera ~4GB, algo real a esta escala (varios millones de filas x 55
columnas).
"""

import math
from pathlib import Path

import pandas as pd
import xlsxwriter

from utils import logger, manejar_errores

FORMATO_MONEDA = "#,##0.00"
LIMITE_FILAS_HOJA = 1_000_000
_FILAS_MUESTRA_ANCHO = 2000  # suficiente para estimar el ancho de columna sin recorrer hojas de millones de filas


def _clasificar_columnas(df: pd.DataFrame) -> dict:
    """
    Determina, por nombre de columna, qué columnas necesitan formato de
    moneda (nombre contiene "monto"/"valor"). Las columnas "fecha" no
    reciben formato: siguen siendo texto, así que un number_format de
    fecha no tendría efecto visual en Excel, solo costaría tiempo de
    escritura sin beneficio.
    """
    formatos = {}
    for col_name in df.columns:
        nombre = str(col_name).lower()
        if "monto" in nombre or "valor" in nombre:
            formatos[col_name] = FORMATO_MONEDA
    return formatos


def _anchos_columnas(df: pd.DataFrame) -> list:
    # Ancho estimado con una MUESTRA de las primeras filas, no con el
    # DataFrame completo: con hojas de cientos de miles o millones de
    # filas (ej. archivos consolidados de 3-4 planos de ~1M filas cada
    # uno), recorrer las 55 columnas completas solo para un ancho visual
    # aproximado es un costo real. El ancho es cosmético, no afecta datos.
    muestra = df.head(_FILAS_MUESTRA_ANCHO)
    anchos = []
    for col_name in df.columns:
        # fillna("") antes de astype(str): en pandas 3.x, astype(str) sobre
        # NaN en una columna float (ej. MONTO-1 cuando pd.to_numeric(errors=
        # "coerce") no pudo convertir el valor) NO produce la cadena "nan"
        # como en pandas < 3 — deja el NaN como float, y map(len) revienta
        # con "object of type 'float' has no len()".
        largo_max = muestra[col_name].fillna("").astype(str).map(len).max() if len(muestra) else 0
        largo_max = max(largo_max, len(str(col_name)))
        anchos.append(min(largo_max + 3, 40))
    return anchos


def _particionar_hoja_grande(nombre_hoja: str, df: pd.DataFrame) -> dict:
    """
    Si df supera LIMITE_FILAS_HOJA filas, lo divide en tantas hojas
    consecutivas de hasta LIMITE_FILAS_HOJA filas cada una como haga
    falta (num_partes = ceil(filas / LIMITE_FILAS_HOJA)), sin reordenar
    filas — solo bloques secuenciales — para que Excel (límite real:
    1,048,576 filas por hoja, encabezado incluido) pueda abrir y reflejar
    la salida completa sin saturarse. Las partes se nombran
    "{nombre_hoja} (1)", "{nombre_hoja} (2)", etc., recortadas a 31
    caracteres (límite de Excel para nombres de hoja) como resguardo,
    aunque con los nombres de hoja actuales del proyecto nunca se acerca
    a ese límite. Si df no supera LIMITE_FILAS_HOJA, se devuelve sin
    cambios bajo su nombre original.
    """
    if len(df) <= LIMITE_FILAS_HOJA:
        return {nombre_hoja: df}

    num_partes = math.ceil(len(df) / LIMITE_FILAS_HOJA)
    partes = {}
    for i in range(num_partes):
        inicio = i * LIMITE_FILAS_HOJA
        fin = inicio + LIMITE_FILAS_HOJA
        nombre_parte = f"{nombre_hoja} ({i + 1})"[:31]
        partes[nombre_parte] = df.iloc[inicio:fin]

    nombres_partes = ", ".join(f"'{n}'" for n in partes)
    mensaje = (
        f"   [i]  '{nombre_hoja}' superó {LIMITE_FILAS_HOJA:,} filas "
        f"({len(df):,}) — dividida en {num_partes} hojas: {nombres_partes}."
    )
    print(mensaje)
    logger.info(f"Hoja particionada: '{nombre_hoja}' ({len(df)} filas) -> {num_partes} partes")

    return partes


def _escribir_hoja_dataframe(wb, nombre_hoja: str, df: pd.DataFrame, fmt_header, fmt_moneda):
    ws = wb.add_worksheet(nombre_hoja)

    formatos = _clasificar_columnas(df)
    anchos = _anchos_columnas(df)

    for col_idx, ancho in enumerate(anchos):
        ws.set_column(col_idx, col_idx, ancho)
    ws.freeze_panes(1, 0)

    columnas = list(df.columns)
    ws.write_row(0, 0, columnas, fmt_header)

    formato_por_columna = [formatos.get(col) for col in columnas]
    # Posiciones que sí necesitan formato de moneda — normalmente 1-2 de
    # 55 (MONTO-1, a veces MONTO-2). En modo constant_memory las celdas de
    # una fila se escriben en orden de columna estrictamente creciente,
    # así que cada fila se parte en tramos sin formato (write_row) y
    # celdas puntuales con formato (write), en vez de recorrer las 55
    # columnas por fila — la diferencia real con hojas de varios millones
    # de filas (ej. consolidados de 3-4 archivos planos de ~1M filas cada
    # uno).
    indices_formato = [i for i, formato in enumerate(formato_por_columna) if formato]

    # NaN (ej. MONTO-1 no numérico tras pd.to_numeric(errors="coerce"))
    # se pasa como None para que quede como celda vacía.
    valores = df.astype(object).where(df.notna(), None).values.tolist()
    ncols = len(columnas)

    if not indices_formato:
        for fila_idx, fila in enumerate(valores, start=1):
            ws.write_row(fila_idx, 0, fila)
        return

    for fila_idx, fila in enumerate(valores, start=1):
        col = 0
        for idx in indices_formato:
            if idx > col:
                ws.write_row(fila_idx, col, fila[col:idx])
            ws.write(fila_idx, idx, fila[idx], fmt_moneda)
            col = idx + 1
        if col < ncols:
            ws.write_row(fila_idx, col, fila[col:])


def _escribir_hoja_libre(wb, nombre_hoja: str, celdas: list, fmt_negrita):
    """
    Escribe una hoja de posición libre (no tabular), a partir de una lista
    de celdas: [{"fila": int, "columna": "D", "valor": ..., "negrita": bool,
    "formato_numero": str|None}, ...]. Se usa para replicar layouts fijos
    (ej. la plantilla de VALOR CINTA) en vez de un DataFrame con encabezado.

    constant_memory no permite direccionar celdas sueltas fuera de orden,
    así que se arma una grilla completa desde la columna A hasta la última
    columna usada (dejando huecos en None) y se manda fila por fila con
    write_row — las columnas conservan su posición real (D, E, F...) para
    que la plantilla replicada no cambie de posición.
    """
    ws = wb.add_worksheet(nombre_hoja)
    if not celdas:
        return

    columnas_usadas = {c["columna"] for c in celdas}
    for col in columnas_usadas:
        col_idx = _col_letra_a_indice(col)
        ws.set_column(col_idx, col_idx, 24)

    max_fila = max(c["fila"] for c in celdas)
    max_col = max(_col_letra_a_indice(c["columna"]) for c in celdas) + 1

    grid = [[None] * max_col for _ in range(max_fila)]
    grid_fmt = [[None] * max_col for _ in range(max_fila)]
    for celda_def in celdas:
        col_idx = _col_letra_a_indice(celda_def["columna"])
        fila_idx = celda_def["fila"] - 1
        grid[fila_idx][col_idx] = celda_def["valor"]
        if celda_def.get("negrita"):
            grid_fmt[fila_idx][col_idx] = fmt_negrita

    for fila_idx, fila in enumerate(grid):
        formatos_fila = grid_fmt[fila_idx]
        if not any(formatos_fila):
            ws.write_row(fila_idx, 0, fila)
            continue
        for col_idx, valor in enumerate(fila):
            if valor is None and formatos_fila[col_idx] is None:
                continue
            ws.write(fila_idx, col_idx, valor, formatos_fila[col_idx])


def _col_letra_a_indice(letra: str) -> int:
    """Convierte una letra de columna Excel ("A", "D", "AA"...) a índice 0-based."""
    indice = 0
    for char in letra:
        indice = indice * 26 + (ord(char.upper()) - ord("A") + 1)
    return indice - 1


@manejar_errores
def guardar_con_formato(path: Path, hojas: dict) -> bool:
    """
    hojas: {"NombreHoja": dataframe_o_lista_de_celdas, ...}
    Cada hoja puede ser un DataFrame (tabla con encabezado, formato
    automático de moneda) o una lista de celdas de posición libre
    (ver _escribir_hoja_libre), para layouts fijos que no son tabulares.
    Si algún DataFrame viene vacío (0 filas), igual se escribe la hoja
    con encabezados en lugar de fallar, pero se avisa explícitamente
    en consola y en el log — así el usuario no pasa por alto que un
    filtro no encontró resultados. Si algún DataFrame supera
    LIMITE_FILAS_HOJA filas, se particiona automáticamente en varias
    hojas antes de escribir (ver _particionar_hoja_grande) — así ninguna
    hoja de salida excede el límite real de Excel (1,048,576 filas).

    Devuelve True si el archivo se generó correctamente. Si algo falla
    (ej. el archivo está abierto en Excel), @manejar_errores captura el
    error y esta función devuelve None — el llamador debe revisar el
    resultado antes de asumir éxito, para no mostrar un mensaje de
    "generado" después de un fallo real.
    """
    hojas_vacias = [
        nombre for nombre, contenido in hojas.items()
        if isinstance(contenido, pd.DataFrame) and len(contenido) == 0
    ]

    # Particionar antes de escribir: una hoja particionada nunca está
    # vacía (supera LIMITE_FILAS_HOJA filas), así que no hay conflicto
    # con hojas_vacias, calculada arriba sobre los nombres originales.
    hojas_expandidas = {}
    for nombre_hoja, contenido in hojas.items():
        if isinstance(contenido, pd.DataFrame):
            hojas_expandidas.update(_particionar_hoja_grande(nombre_hoja, contenido))
        else:
            hojas_expandidas[nombre_hoja] = contenido

    wb = xlsxwriter.Workbook(str(path), {"constant_memory": True})
    wb.use_zip64()  # el .xlsx comprimido puede superar 4GB a esta escala (varios millones de filas x 55 columnas)

    fmt_header = wb.add_format({
        "bold": True,
        "font_color": "FFFFFF",
        "bg_color": "1F4E78",
        "align": "center",
    })
    fmt_moneda = wb.add_format({"num_format": FORMATO_MONEDA})
    fmt_negrita = wb.add_format({"bold": True})

    for nombre_hoja, contenido in hojas_expandidas.items():
        if isinstance(contenido, pd.DataFrame):
            _escribir_hoja_dataframe(wb, nombre_hoja, contenido, fmt_header, fmt_moneda)
        else:
            _escribir_hoja_libre(wb, nombre_hoja, contenido, fmt_negrita)
    wb.close()

    print(f"[OK] Generado: {path}")
    for nombre_hoja in hojas_vacias:
        mensaje = f"   [!]  '{nombre_hoja}' en {path.name} quedó sin filas — revisa el filtro aplicado."
        print(mensaje)
        logger.warning(f"Hoja vacía: {nombre_hoja} en {path}")

    return True
