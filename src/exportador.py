"""
exportador.py
Escribe DataFrames a .xlsx con formato legible: encabezado resaltado,
ancho de columna automático, encabezado congelado y formato numérico/
fecha según el nombre del campo.
"""

from pathlib import Path

import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from utils import logger, manejar_errores


def _formatear_hoja(ws, df: pd.DataFrame):
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF")

    columnas_moneda = []
    columnas_fecha = []

    for col_idx, col_name in enumerate(df.columns, start=1):
        celda = ws.cell(row=1, column=col_idx)
        celda.fill = header_fill
        celda.font = header_font
        celda.alignment = Alignment(horizontal="center")

        largo_max = max(
            df[col_name].astype(str).map(len).max() if len(df) else 0,
            len(str(col_name)),
        )
        ws.column_dimensions[get_column_letter(col_idx)].width = min(largo_max + 3, 40)

        nombre = str(col_name).lower()
        if "monto" in nombre or "valor" in nombre:
            columnas_moneda.append(col_idx)
        elif "fecha" in nombre:
            columnas_fecha.append(col_idx)

    ws.freeze_panes = "A2"

    # Aplicar formato numérico/fecha solo en las columnas que lo necesitan
    # (no celda por celda en todas las columnas) — con miles de filas, iterar
    # las 55 columnas por fila es el cuello de botella real de este proceso.
    for col_idx in columnas_moneda:
        for (celda,) in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=col_idx, max_col=col_idx):
            celda.number_format = "#,##0.00"

    for col_idx in columnas_fecha:
        for (celda,) in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=col_idx, max_col=col_idx):
            celda.number_format = "DD/MM/YYYY"


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

    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        for nombre_hoja, contenido in hojas.items():
            if isinstance(contenido, pd.DataFrame):
                contenido.to_excel(writer, sheet_name=nombre_hoja, index=False)
                ws = writer.sheets[nombre_hoja]
                _formatear_hoja(ws, contenido)
            else:
                ws = writer.book.create_sheet(nombre_hoja)
                _escribir_hoja_libre(ws, contenido)

    print(f"[OK] Generado: {path}")
    for nombre_hoja in hojas_vacias:
        mensaje = f"   [!]  '{nombre_hoja}' en {path.name} quedó sin filas — revisa el filtro aplicado."
        print(mensaje)
        logger.warning(f"Hoja vacía: {nombre_hoja} en {path}")

    return True