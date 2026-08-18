"""
main.py
Menú interactivo principal.

Flujo:
1. Seleccionar archivo(s) plano(s) — uno o varios (ej. consolidado de
fin de semana: sábado+domingo+lunes, o +martes si hay festivo)
2. Generar Archivo 1: leer archivo(s), construir DataFrame base
etiquetado por día, aplicar reglas de Archivo 1, generar su Excel
3. Generar Archivo 2: ídem, con reglas de Archivo 2
4. Generar Archivo 3: ídem, con reglas de Archivo 3
5. Salir

Cada opción de generación relee los archivos planos y reconstruye el
DataFrame base desde cero (no se cachea entre opciones).

Todo el flujo pasa por manejar_errores en cada módulo: un fallo puntual
(archivo abierto en Excel, ruta inválida, etc.) no debe tumbar el menú.
"""

import sys
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent / "src"))

from config import ARCHIVO_1, ARCHIVO_2, ARCHIVO_3  # noqa: E402
from exportador import guardar_con_formato  # noqa: E402
from lector import construir_dataframe, leer_lineas  # noqa: E402
from reglas import generar_archivo_1, generar_archivo_2, generar_archivo_3  # noqa: E402
from utils import limpiar_pantalla, pausar  # noqa: E402
from validador import extraer_dia, seleccionar_archivos  # noqa: E402

# Estado de la sesión (simple, sin necesidad de clases para esto — KISS)
estado = {"rutas": [], "df": None}


def mostrar_menu():
    limpiar_pantalla()
    print("=" * 55)
    print(" PROCESAMIENTO DE ARCHIVO PLANO — CONEXION DIRECTA")
    print("=" * 55)
    if estado["rutas"]:
        print(f" Archivos actuales: {len(estado['rutas'])} seleccionado(s)")
        for ruta in estado["rutas"]:
            print(f"   - {ruta.name}")
    else:
        print(" Archivos actuales: (ninguno seleccionado)")
    print("-" * 55)
    print(" 1. Seleccionar archivo(s) plano(s)")
    print(" 2. Generar Archivo 1 (SOPORTES REDEBAN VISA)")
    print(" 3. Generar Archivo 2 (RECH EP DEPO)")
    print(" 4. Generar Archivo 3 (DDMMYY) - VALORES CINTA")
    print(" 5. Salir")
    print("=" * 55)


def opcion_seleccionar():
    rutas = seleccionar_archivos()
    if rutas:
        estado["rutas"] = rutas
        estado["df"] = None
        print(f"\n{len(rutas)} archivo(s) seleccionado(s):")
        for ruta in rutas:
            print(f"   - {ruta.name}")
    pausar()


def _construir_dataframe_base():
    """
    Lee los archivos planos seleccionados y construye el DataFrame base
    combinado, etiquetando cada fila con su día de origen (columna
    "__dia", extraída del nombre de archivo vía extraer_dia). Cada
    opción de generación relee y reconstruye desde cero (sin caché) para
    no trabajar con datos desactualizados. Devuelve None si algo falla o
    no hay archivos seleccionados.
    """
    if not estado["rutas"]:
        print("\nPrimero selecciona uno o varios archivos (opción 1).")
        pausar()
        return None

    partes = []
    for ruta in estado["rutas"]:
        dia = extraer_dia(ruta)
        if dia is None:
            print(f"\nNo se pudo determinar el día para {ruta.name}. Se omite este archivo.")
            continue

        print(f"\nLeyendo archivo {ruta.name} (día {dia})...")
        lineas = leer_lineas(ruta)
        if not lineas:
            print(f"El archivo {ruta.name} no tiene líneas para procesar. Se omite.")
            continue

        df_parcial = construir_dataframe(lineas)
        if df_parcial is None:
            print(f"No se pudo construir la tabla para {ruta.name}. Se omite.")
            continue

        df_parcial["__dia"] = dia
        partes.append(df_parcial)
        print(f"   {len(df_parcial)} filas asignadas al día {dia}.")

    if not partes:
        print("\nNingún archivo pudo procesarse.")
        pausar()
        return None

    print("\nConstruyendo tabla base combinada...")
    df = pd.concat(partes, ignore_index=True)
    estado["df"] = df
    print(f"Tabla base construida: {len(df)} filas, {df['__dia'].nunique()} día(s).")
    return df


def opcion_generar_archivo_1():
    df = _construir_dataframe_base()
    if df is None:
        return

    print("\nAplicando reglas de negocio (Archivo 1)...")
    hojas = generar_archivo_1(df)
    if hojas is None:
        pausar()
        return

    print("\nGenerando archivo Excel...")
    if not guardar_con_formato(ARCHIVO_1, hojas):
        pausar()
        return

    print("\n[OK] Archivo 1 generado.")
    pausar()


def opcion_generar_archivo_2():
    df = _construir_dataframe_base()
    if df is None:
        return

    print("\nAplicando reglas de negocio (Archivo 2)...")
    hojas = generar_archivo_2(df)
    if hojas is None:
        pausar()
        return

    print("\nGenerando archivo Excel...")
    if not guardar_con_formato(ARCHIVO_2, hojas):
        pausar()
        return

    print("\n[OK] Archivo 2 generado.")
    pausar()


def opcion_generar_archivo_3():
    df = _construir_dataframe_base()
    if df is None:
        return

    print("\nAplicando reglas de negocio (Archivo 3)...")
    hojas = generar_archivo_3(df)  # sobre el universo completo, no sobre 1 ni 2
    if hojas is None:
        pausar()
        return

    print("\nGenerando archivo Excel...")
    if not guardar_con_formato(ARCHIVO_3, hojas):
        pausar()
        return

    print("\n[OK] Archivo 3 generado.")
    pausar()


def main():
    acciones = {
        "1": opcion_seleccionar,
        "2": opcion_generar_archivo_1,
        "3": opcion_generar_archivo_2,
        "4": opcion_generar_archivo_3,
    }

    while True:
        mostrar_menu()
        opcion = input("Selecciona una opción: ").strip()

        if opcion == "5":
            print("\nSaliendo...")
            break

        accion = acciones.get(opcion)
        if accion:
            accion()
        else:
            print("\nOpción no válida.")
            pausar()


if __name__ == "__main__":
    main()
