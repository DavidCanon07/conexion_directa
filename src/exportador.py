"""
exportador.py
Escribe DataFrames a .xlsx con formato legible: encabezado resaltado,
ancho de columna automático, encabezado congelado y formato numérico/
fecha según el nombre del campo.

Las hojas tabulares (DataFrame) se escriben fila por fila con
ws.append() en vez de DataFrame.to_excel(engine="openpyxl"): to_excel
escribe celda por celda (ws.cell(row, col, value) por cada una de las
~55 columnas x cada fila), que con cientos de miles de filas es
sensiblemente más lento que anexar la fila completa de una vez. El
ancho de columna y el formato de moneda/fecha se calculan en el mismo
recorrido de escritura en vez de en pasadas adicionales sobre todo el
DataFrame.
"""

from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.cell import WriteOnlyCell
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from utils import logger, manejar_errores

_ANCHO_MAXIMO_COLUMNA = 40
_FILAS_MUESTRA_ANCHO = 2000  # suficiente para estimar el ancho de columnas de texto sin recorrer todo el DataFrame


def _calcular_formato_y_ancho(df: pd.DataFrame) -> tuple:
    """
    Determina, por columna, el formato numérico (moneda/fecha por
    substring del nombre) y el ancho de columna. Compartido entre el
    escritor normal y el write_only: ambos necesitan lo mismo antes de
    escribir el encabezado.
    """
    formato_por_columna = {}
    ancho_por_columna = {}
    muestra = df.head(_FILAS_MUESTRA_ANCHO)

    for col_idx, nombre in enumerate(df.columns, start=1):
        nombre_min = str(nombre).lower()
        if "monto" in nombre_min or "valor" in nombre_min:
            formato_por_columna[col_idx] = "#,##0.00"
            ancho_por_columna[col_idx] = 18
        elif "fecha" in nombre_min:
            formato_por_columna[col_idx] = "DD/MM/YYYY"
            ancho_por_columna[col_idx] = 12
        else:
            largo_muestra = muestra[nombre].astype(str).map(len).max() if len(muestra) else 0
            ancho_por_columna[col_idx] = min(max(largo_muestra, len(str(nombre))) + 3, _ANCHO_MAXIMO_COLUMNA)

    return formato_por_columna, ancho_por_columna


def _formatear_hoja(ws, df: pd.DataFrame):
    """
    Escribe encabezado (resaltado + centrado), congela la fila 1 y
    escribe los datos fila por fila. El ancho de columna se estima con
    una muestra de las primeras filas (no con el DataFrame completo:
    con cientos de miles de filas, astype(str).map(len) sobre cada
    columna es un recorrido Python completo adicional y solo afecta el
    ancho visual, no los datos ni los filtros). Las columnas de
    moneda/fecha reciben un ancho fijo razonable en vez de muestreado,
    ya que su formato numérico ya fija su longitud típica.
    """
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF")
    header_alignment = Alignment(horizontal="center")

    columnas = list(df.columns)
    formato_por_columna, ancho_por_columna = _calcular_formato_y_ancho(df)

    ws.append(columnas)
    for col_idx in range(1, len(columnas) + 1):
        celda = ws.cell(row=1, column=col_idx)
        celda.fill = header_fill
        celda.font = header_font
        celda.alignment = header_alignment
        ws.column_dimensions[get_column_letter(col_idx)].width = ancho_por_columna[col_idx]
    ws.freeze_panes = "A2"

    # Excel/openpyxl no aceptan NaN/NaT; to_excel los convertía a celda
    # vacía automáticamente, aquí se reemplazan explícitamente por None.
    datos = df.where(pd.notnull(df), None)

    if not formato_por_columna:
        for fila in datos.itertuples(index=False, name=None):
            ws.append(fila)
        return

    fila_excel = 1
    for fila in datos.itertuples(index=False, name=None):
        ws.append(fila)
        fila_excel += 1
        for col_idx, formato in formato_por_columna.items():
            ws.cell(row=fila_excel, column=col_idx).number_format = formato


def _escribir_dataframe_write_only(ws, df: pd.DataFrame):
    """
    Equivalente a _formatear_hoja pero para una hoja del Workbook en
    modo write_only (openpyxl.Workbook(write_only=True)): ese modo
    transmite cada fila directo al XML de salida sin retener los
    objetos Cell en memoria, en vez de acumularlos todos como hace un
    Workbook normal. Con archivos de hasta 1 millón de filas concentradas
    en pocas hojas grandes (ej. "TIPO 50", "TIPO 12-15"), retener todas
    las celdas en memoria degrada el rendimiento de forma no lineal
    (se midió: una hoja con ~3x más filas tardó en escribirse ~5x más
    con Workbook normal) — write_only evita esa degradación.

    Contrapartida: en modo write_only no se puede volver a tocar una
    celda después de escribir su fila (ya se serializó), así que el
    encabezado y el formato de moneda/fecha por celda deben construirse
    ANTES de cada ws.append(), con WriteOnlyCell en vez de asignar
    ws.cell(row, col).atributo después.
    """
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF")
    header_alignment = Alignment(horizontal="center")

    columnas = list(df.columns)
    formato_por_columna, ancho_por_columna = _calcular_formato_y_ancho(df)

    # En modo write_only, freeze_panes solo se serializa si se asigna
    # ANTES del primer ws.append() (comportamiento verificado en
    # openpyxl 3.1.5) — después del primer append queda ignorado
    # silenciosamente, sin error.
    ws.freeze_panes = "A2"
    for col_idx, ancho in ancho_por_columna.items():
        ws.column_dimensions[get_column_letter(col_idx)].width = ancho

    encabezado = []
    for nombre in columnas:
        celda = WriteOnlyCell(ws, value=nombre)
        celda.fill = header_fill
        celda.font = header_font
        celda.alignment = header_alignment
        encabezado.append(celda)
    ws.append(encabezado)

    datos = df.where(pd.notnull(df), None)

    if not formato_por_columna:
        for fila in datos.itertuples(index=False, name=None):
            ws.append(fila)
        return

    for fila in datos.itertuples(index=False, name=None):
        fila_salida = list(fila)
        for col_idx, formato in formato_por_columna.items():
            celda = WriteOnlyCell(ws, value=fila_salida[col_idx - 1])
            celda.number_format = formato
            fila_salida[col_idx - 1] = celda
        ws.append(fila_salida)


def _escribir_hoja_libre(ws, celdas: list):
    """
    Escribe una hoja de posición libre (no tabular), a partir de una lista
    de celdas: [{"fila": int, "columna": "D", "valor": ..., "negrita": bool,
    "formato_numero": str|None}, ...]. Se usa para replicar layouts fijos
    (ej. la plantilla de VALOR CINTA) en vez de un DataFrame con encabezado.
    """
    for celda_def in celdas:
        celda = ws[f"{celda_def['columna']}{celda_def['fila']}"]
        celda.value = celda_def["valor"]
        if celda_def.get("negrita"):
            celda.font = Font(bold=True)
        formato = celda_def.get("formato_numero")
        if formato:
            celda.number_format = formato

    columnas_usadas = {c["columna"] for c in celdas}
    for col in columnas_usadas:
        ws.column_dimensions[col].width = 24


@manejar_errores
def guardar_con_formato(path: Path, hojas: dict) -> bool:
    """
    hojas: {"NombreHoja": dataframe_o_lista_de_celdas, ...}
    Cada hoja puede ser un DataFrame (tabla con encabezado, formato
    automático de moneda/fecha) o una lista de celdas de posición libre
    (ver _escribir_hoja_libre), para layouts fijos que no son tabulares.
    Si algún DataFrame viene vacío (0 filas), igual se escribe la hoja
    con encabezados en lugar de fallar, pero se avisa explícitamente
    en consola y en el log — así el usuario no pasa por alto que un
    filtro no encontró resultados.

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

    # Workbook manual (no pd.ExcelWriter). Si TODAS las hojas son
    # DataFrame (caso de Archivo 1 y Archivo 2, que nunca traen la hoja
    # de posición libre) se usa Workbook(write_only=True): transmite
    # cada fila al XML sin retenerla en memoria, y escala linealmente
    # con hojas de cientos de miles de filas (ver _escribir_dataframe_write_only).
    # Si alguna hoja es de posición libre (Archivo 3, "Hoja1") se necesita
    # asignación aleatoria de celdas (ws["D3"] = ...), incompatible con
    # write_only, así que se usa un Workbook normal para todo el libro.
    todas_dataframes = all(isinstance(contenido, pd.DataFrame) for contenido in hojas.values())

    if todas_dataframes:
        wb = Workbook(write_only=True)
        for nombre_hoja, contenido in hojas.items():
            ws = wb.create_sheet(title=nombre_hoja)
            _escribir_dataframe_write_only(ws, contenido)
    else:
        wb = Workbook()
        wb.remove(wb.active)
        for nombre_hoja, contenido in hojas.items():
            ws = wb.create_sheet(title=nombre_hoja)
            if isinstance(contenido, pd.DataFrame):
                _formatear_hoja(ws, contenido)
            else:
                _escribir_hoja_libre(ws, contenido)
    wb.save(path)

    print(f"[OK] Generado: {path}")
    for nombre_hoja in hojas_vacias:
        mensaje = f"   [!]  '{nombre_hoja}' en {path.name} quedó sin filas — revisa el filtro aplicado."
        print(mensaje)
        logger.warning(f"Hoja vacía: {nombre_hoja} en {path}")

    return True