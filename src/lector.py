"""
lector.py
Lee el archivo plano y construye el DataFrame base usando el LAYOUT.

Este DataFrame es la única fuente de verdad: de aquí se derivan
tanto los filtros de negocio (archivo 1 y 2) como las tablas
dinámicas sobre el universo completo (archivo 3).
"""

from pathlib import Path

import pandas as pd

from config import ENCODING, LAYOUT
from utils import manejar_errores


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
    Recorre LAYOUT y extrae cada campo por posición de caracter.
    Todo se trae como texto (str) y se limpia (strip); las conversiones
    de tipo (fecha, numérico) se hacen explícitamente en reglas.py,
    campo por campo, para mantener control total sobre el formato real.
    """
    registros = []
    for linea in lineas:
        registro = {
            campo["nombre"]: linea[campo["inicio"] - 1: campo["inicio"] - 1 + campo["longitud"]].strip()
            for campo in LAYOUT
        }
        registros.append(registro)

    df = pd.DataFrame(registros)
    return df
