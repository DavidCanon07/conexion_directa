"""
utils.py
Utilidades generales: logging y decorador de manejo de errores.

Punto 5 del requerimiento: ningún fallo de usuario (archivo no encontrado,
archivo abierto en Excel, formato inválido, etc.) debe tumbar el programa.
Todo pasa por manejar_errores, que registra en log y regresa al menú.
"""

import functools
import logging
import os
from datetime import datetime

from config import CARPETA_LOGS

logging.basicConfig(
    filename=CARPETA_LOGS / f"ejecucion_{datetime.now():%Y%m%d}.log",
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
    encoding="utf-8",
)
logger = logging.getLogger(__name__)


def manejar_errores(func):
    """
    Decorador: envuelve cualquier función del flujo principal.
    Captura errores comunes de forma específica (mensaje claro al usuario)
    y cualquier otro error de forma genérica, sin detener la ejecución.
    """
    @functools.wraps(func)
    def wrapper(*args, **kwargs):
        try:
            return func(*args, **kwargs)
        except FileNotFoundError as e:
            print(f"\n[!] No se encontró el archivo: {e}")
            logger.error(f"FileNotFoundError en {func.__name__}: {e}")
        except PermissionError as e:
            print(f"\n[!] Sin permisos de acceso (¿el archivo está abierto en Excel?): {e}")
            logger.error(f"PermissionError en {func.__name__}: {e}")
        except ValueError as e:
            print(f"\n[!] Error de datos o configuración: {e}")
            logger.error(f"ValueError en {func.__name__}: {e}")
        except Exception as e:
            print(f"\n[!] Ocurrió un error inesperado: {e}")
            logger.exception(f"Error inesperado en {func.__name__}")
        return None
    return wrapper


def limpiar_pantalla():
    os.system("cls" if os.name == "nt" else "clear")


def pausar():
    input("\nPresiona ENTER para continuar...")
