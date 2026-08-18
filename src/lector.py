"""
lector.py
Lee el archivo plano y construye el DataFrame base usando el LAYOUT.

Este DataFrame es la única fuente de verdad: de aquí se derivan
tanto los filtros de negocio (archivo 1 y 2) como las tablas
dinámicas sobre el universo completo (archivo 3).
"""

import io
from pathlib import Path

import pandas as pd

from config import ENCODING, LAYOUT
from utils import manejar_errores

_NOMBRES_LAYOUT = [campo["nombre"] for campo in LAYOUT]
_COLSPECS_LAYOUT = [(campo["inicio"] - 1, campo["inicio"] - 1 + campo["longitud"]) for campo in LAYOUT]


@manejar_errores
def leer_lineas(ruta: Path) -> list:
    """
    Lee el archivo plano y devuelve la lista de líneas no vacías,
    sin el salto de línea. No se aplica ningún filtro de largo.
    """
    with open(ruta, "r", encoding=ENCODING, errors="replace") as f:
        return [linea.rstrip("\n\r") for linea in f if linea.strip()]


@manejar_errores
def construir_dataframe(lineas: list) -> pd.DataFrame:
    """
    Extrae cada campo de LAYOUT por posición de caracter usando el
    parser de ancho fijo de pandas (pd.read_fwf, acelerado en C) en vez
    de un bucle Python campo por campo por línea — con archivos de
    800 mil a 1 millón de filas ese bucle (55 campos x fila) es el
    cuello de botella real del proceso completo. El resultado es
    idéntico al slicing manual: mismas posiciones (colspecs = inicio-1,
    inicio-1+longitud, igual que antes), todo como texto (dtype=str,
    sin inferencia de tipo) y stripeado campo por campo — las
    conversiones de tipo (fecha, numérico) se siguen haciendo
    explícitamente en reglas.py.
    """
    if not lineas:
        return pd.DataFrame(columns=_NOMBRES_LAYOUT)

    buffer = io.StringIO("\n".join(lineas))
    df = pd.read_fwf(
        buffer,
        colspecs=_COLSPECS_LAYOUT,
        names=_NOMBRES_LAYOUT,
        dtype=str,
        header=None,
    )
    df = df.fillna("")
    for nombre in _NOMBRES_LAYOUT:
        df[nombre] = df[nombre].str.strip()
    return df
