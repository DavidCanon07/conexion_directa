"""
validador.py
Selección de los archivos planos y extracción de su día de origen.
"""

import re
from pathlib import Path
from typing import List

from utils import manejar_errores

# Patrón de nombre real de producción: GOF.GRB.FM14.FYYMMDD.txt
PATRON_DIA = re.compile(r"F(\d{2})(\d{2})(\d{2})", re.IGNORECASE)


@manejar_errores
def seleccionar_archivos() -> List[Path]:
    """
    Abre un diálogo gráfico para elegir uno o varios archivos planos
    (Ctrl/Shift+clic para selección múltiple — pensado para consolidar
    varios días, ej. sábado+domingo+lunes de un fin de semana).
    Si no hay entorno gráfico disponible, pide las rutas por consola
    (separadas por coma).
    """
    rutas = []
    try:
        import tkinter as tk
        from tkinter import filedialog

        root = tk.Tk()
        root.withdraw()
        seleccion = filedialog.askopenfilenames(
            title="Selecciona uno o varios archivos planos (.txt)",
            filetypes=[("Archivos de texto", "*.txt"), ("Todos los archivos", "*.*")],
        )
        root.destroy()
        rutas = list(seleccion)
    except Exception:
        texto = input("Ingresa las rutas completas de los archivos planos (separadas por coma): ").strip()
        rutas = [r.strip() for r in texto.split(",") if r.strip()]

    if not rutas:
        print("No se seleccionó ningún archivo.")
        return []

    rutas_path = []
    for ruta in rutas:
        ruta_path = Path(ruta)
        if not ruta_path.exists():
            raise FileNotFoundError(ruta_path)
        rutas_path.append(ruta_path)

    return rutas_path


@manejar_errores
def extraer_dia(ruta: Path) -> str:
    """
    Extrae el día (2 dígitos) del nombre del archivo plano, según el
    patrón de producción "GOF.GRB.FM14.FYYMMDD.txt" (F + año + mes + día).
    Si el nombre no calza con el patrón, se le pregunta al usuario qué
    día corresponde a ese archivo específico, en vez de asumir un valor
    — así ningún archivo queda mal etiquetado silenciosamente.
    """
    coincidencia = PATRON_DIA.search(ruta.stem)
    if coincidencia:
        return coincidencia.group(3)

    print(f"\n[!] No se pudo identificar el día en el nombre del archivo: {ruta.name}")
    while True:
        dia = input(f"Ingresa el día (1-2 dígitos) que corresponde a '{ruta.name}': ").strip()
        if dia.isdigit() and 1 <= len(dia) <= 2:
            return dia.zfill(2)
        print("Valor inválido. Ingresa solo números (ej. 9 o 09).")
