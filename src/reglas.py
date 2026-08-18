"""
reglas.py
Reglas de negocio del proceso.

Archivo 1 y Archivo 2: se construyen con filtros explícitos sobre el
DataFrame base (df completo, sin filtrar).

Archivo 3: calcula el VALOR CINTA (nacional e internacional) mediante una
fórmula de negocio fija sobre el DataFrame base completo — replica las
fórmulas reales de la hoja "TABLA" de ESTRUC CINTA POS DEPOS FINAL V3.xlsx
(celdas F67 y F81). NO depende de los archivos 1 o 2.

Cada función recibe el df base y devuelve un diccionario
{"NombreHoja": dataframe_resultado}, listo para exportar. El valor de
cada hoja normalmente es un DataFrame; Archivo 3 usa en cambio una lista
de celdas de posición libre (ver _escribir_hoja_libre en exportador.py)
porque replica un layout fijo, no una tabla.
Así cada archivo de salida puede tener 1 o varias hojas sin tocar
el resto del programa.

Selección múltiple de archivos (consolidado de fin de semana): el df
base puede traer una columna "__dia" (agregada en main.py, un valor por
archivo de origen). Las hojas "TIPO 01" (Archivo 1), "RECH EP"
(Archivo 2) y "Detalle Dinamicas" (Archivo 3) se generan una vez por
cada día presente en esa columna (ej. "TIPO 01 13", "TIPO 01 14"); las
demás hojas ("TIPO 50", "TIPO 12-15", "Hoja1" de VALOR CINTA) se
mantienen combinadas con todos los días juntos. Si "__dia" no está
presente (df sin etiquetar), se genera una sola hoja sin sufijo.
"""

import pandas as pd

from utils import manejar_errores


def _segregar_por_dia(df_base: pd.DataFrame, hoja_filtrada: pd.DataFrame, nombre_base: str) -> dict:
    """
    Segrega una hoja ya filtrada en una hoja por día ("NombreBase DD"),
    usando los días presentes en df_base (no solo los que sobrevivieron
    el filtro), para que un día sin resultados igual genere su hoja
    vacía con la advertencia ya existente de exportador.py. Si df_base
    no tiene columna "__dia" (sin selección múltiple/etiquetado),
    devuelve una sola hoja sin sufijo, con el comportamiento de siempre.
    """
    if "__dia" not in df_base.columns:
        return {nombre_base: hoja_filtrada}

    dias = sorted(df_base["__dia"].dropna().unique())
    return {
        f"{nombre_base} {dia}": hoja_filtrada[hoja_filtrada["__dia"] == dia].drop(columns="__dia", errors="ignore")
        for dia in dias
    }


# ─────────────────────────────────────────────────────────────
# ARCHIVO 1 — filtros de negocio (2 hojas)
# ─────────────────────────────────────────────────────────────
@manejar_errores
def generar_archivo_1(df: pd.DataFrame) -> dict:
    df = df.copy()
    if "MONTO-1" in df.columns:
        # MONTO-1 viene como texto de ancho fijo (13 dígitos) con 2 decimales
        # implícitos en los últimos 2 dígitos, ej. "0000001464000" -> 14640.00.
        df["MONTO-1"] = pd.to_numeric(df["MONTO-1"], errors="coerce") / 100

    #
    #condiciones para la hoja 1
    cond_red = df.get("RED-LOGICA", pd.Series(dtype=str)).isin(["VISN", "VISA"])
    cond_tipo_50 = df.get("TIPO-REGISTRO", pd.Series(dtype=str)).eq("50")
    
    #condiciones para la hoja 2
    cond_tipo_01 = df.get("TIPO-REGISTRO", pd.Series(dtype=str)).eq("01")
    cond_FIID_SPONSOR = df.get("FIID SPONSOR", pd.Series(dtype=str)).eq("0013")
    cond_TIPO_DE_MENSAJE = df.get("TIPO DE MENSAJE", pd.Series(dtype=str)).isin(["0210", "0420"])
    
    cond_COD_TIPO_TRANS = df.get("COD-TIPO-TRANS", pd.Series(dtype=str)).isin(["10", "14", "15"])
    cond_CODIGO_RESP = df.get(" CODIGO-RESP", pd.Series(dtype=str)).between("000", "009")

    hoja1 = df[
        cond_red
        & cond_tipo_50
    ].drop(columns="__dia", errors="ignore")

    hoja2 = df[
        cond_red
        & cond_tipo_01 & cond_FIID_SPONSOR & cond_TIPO_DE_MENSAJE & cond_COD_TIPO_TRANS & cond_CODIGO_RESP
    ]

    # MONTO-1 queda en negativo para mensajes de reverso (TIPO DE MENSAJE
    # "0420") y para el código de transacción "14" — unión de condiciones
    # con -abs() para que una fila que cumpla ambas no se niegue dos veces.
    hoja2 = hoja2.copy()
    cond_negativo = (
        hoja2.get("TIPO DE MENSAJE", pd.Series(dtype=str)).eq("0420")
        | hoja2.get("COD-TIPO-TRANS", pd.Series(dtype=str)).eq("14")
    )
    hoja2.loc[cond_negativo, "MONTO-1"] = -hoja2.loc[cond_negativo, "MONTO-1"].abs()

    hoja2 = hoja2.sort_values("COD-TIPO-TRANS", kind="stable")

    hojas = {"TIPO 50": hoja1}
    hojas.update(_segregar_por_dia(df, hoja2, "TIPO 01"))
    return hojas


# ─────────────────────────────────────────────────────────────
# ARCHIVO 2 — filtros de negocio (2 hojas)
# ─────────────────────────────────────────────────────────────
@manejar_errores
def generar_archivo_2(df: pd.DataFrame) -> dict:
    df = df.copy()
    if "MONTO-1" in df.columns:
        # MONTO-1 viene como texto de ancho fijo (13 dígitos) con 2 decimales
        # implícitos en los últimos 2 dígitos, ej. "0000001464000" -> 14640.00.
        df["MONTO-1"] = pd.to_numeric(df["MONTO-1"], errors="coerce") / 100

    #condiciones para la hoja "TIPO 12-15"
    cond_tipo_12_15 = df.get("TIPO-REGISTRO", pd.Series(dtype=str)).isin(["12", "15"])

    #condiciones para la hoja "RECH EP"
    cond_tipo_01_20 = df.get("TIPO-REGISTRO", pd.Series(dtype=str)).isin(["01", "20"])
    cond_FIID_SPONSOR = df.get("FIID SPONSOR", pd.Series(dtype=str)).eq("0013")
    cond_TIPO_DE_MENSAJE = df.get("TIPO DE MENSAJE", pd.Series(dtype=str)).isin(["0210", "0420"])
    cond_COD_TIPO_TRANS = df.get("COD-TIPO-TRANS", pd.Series(dtype=str)).isin(["10", "14", "15"])
    cond_CODIGO_RESP = df.get(" CODIGO-RESP", pd.Series(dtype=str)).between("000", "009")

    hoja1 = df[cond_tipo_12_15].drop(columns="__dia", errors="ignore")

    hoja2 = df[
        cond_tipo_01_20
        & cond_FIID_SPONSOR & cond_TIPO_DE_MENSAJE & cond_COD_TIPO_TRANS & cond_CODIGO_RESP
    ]

    hojas = {"TIPO 12-15": hoja1}
    hojas.update(_segregar_por_dia(df, hoja2, "RECH EP"))
    return hojas


def _calcular_valor_cinta(df: pd.DataFrame) -> dict:
    """
    Calcula los componentes del VALOR CINTA (nacional e internacional)
    sobre el df recibido (puede ser el universo completo combinado o el
    subconjunto de un solo día). Replica exactamente las fórmulas
    GETPIVOTDATA de la hoja "TABLA" del libro de estructura (celdas F67
    y F81), traducidas a máscaras booleanas sobre MONTO-1.

    Cada término corresponde a una de las 5 tablas dinámicas reales de esa
    hoja (verificado reconstruyendo los filtros desde el XML del libro y
    contrastando contra los valores cacheados en F67/F81):
      pivot5 ($A$3)  -> TIPO-REGISTRO in [12,15], sin más filtros
      pivot4 ($E$4)  -> FIID=0013, TIPO-REGISTRO in [01,21], TIPO DE MENSAJE=0210,
                        COD-TIPO-TRANS in [10,15], CODIGO-RESP in [000,001]
      pivot1 ($E$56) -> FIID=0013, TIPO-REGISTRO in [01,21], TIPO DE MENSAJE=0420,
                        COD-TIPO-TRANS=14 (sin filtro de CODIGO-RESP ni de indicador)
      pivot3 ($E$27) -> FIID=0013, TIPO-REGISTRO in [01,21], TIPO DE MENSAJE=0420,
                        COD-TIPO-TRANS in [10,14,15], CODIGO-RESP in [000,001]
      pivot2 ($E$42) -> FIID=0013, TIPO-REGISTRO in [01,21], TIPO DE MENSAJE=0210,
                        COD-TIPO-TRANS=14, CODIGO-RESP in [000,001]
    """
    monto = df.get("MONTO-1", pd.Series(dtype=float))
    tipo_reg = df.get("TIPO-REGISTRO", pd.Series(dtype=str))
    fiid = df.get("FIID SPONSOR", pd.Series(dtype=str))
    tipo_msg = df.get("TIPO DE MENSAJE", pd.Series(dtype=str))
    cod_tipo = df.get("COD-TIPO-TRANS", pd.Series(dtype=str))
    codigo_resp = df.get(" CODIGO-RESP", pd.Series(dtype=str))
    indicador = df.get("INDICADOR INTER/NACIONAL", pd.Series(dtype=str))

    cond_fiid = fiid.eq("0013")
    cond_resp = codigo_resp.isin(["000", "001"])
    cond_reg_01_21 = tipo_reg.isin(["01", "21"])
    cond_msg_0210 = tipo_msg.eq("0210")
    cond_msg_0420 = tipo_msg.eq("0420")

    pivot5_total = monto[tipo_reg.isin(["12", "15"])].sum()

    base_pivot4 = cond_fiid & cond_reg_01_21 & cond_msg_0210 & cod_tipo.isin(["10", "15"]) & cond_resp
    pivot4_total_n = monto[base_pivot4 & indicador.eq("N")].sum()
    pivot4_total_i = monto[base_pivot4 & indicador.eq("I")].sum()

    pivot1_total = monto[cond_fiid & cond_reg_01_21 & cond_msg_0420 & cod_tipo.eq("14")].sum()

    base_pivot3 = cond_fiid & cond_reg_01_21 & cond_msg_0420 & cod_tipo.isin(["10", "14", "15"]) & cond_resp
    pivot3_total_n = monto[base_pivot3 & indicador.eq("N")].sum()
    pivot3_total_i = monto[base_pivot3 & indicador.eq("I")].sum()

    base_pivot2 = cond_fiid & cond_reg_01_21 & cond_msg_0210 & cod_tipo.eq("14") & cond_resp
    pivot2_total_n = monto[base_pivot2 & indicador.eq("N")].sum()
    pivot2_total_i = monto[base_pivot2 & indicador.eq("I")].sum()

    valor_cinta_nacional = pivot5_total + pivot4_total_n + pivot1_total - pivot3_total_n - pivot2_total_n
    valor_cinta_internacional = pivot4_total_i - pivot3_total_i - pivot2_total_i

    return {
        "pivot5_total": pivot5_total,
        "pivot4_total_n": pivot4_total_n,
        "pivot1_total": pivot1_total,
        "pivot3_total_n": pivot3_total_n,
        "pivot2_total_n": pivot2_total_n,
        "valor_cinta_nacional": valor_cinta_nacional,
        "pivot4_total_i": pivot4_total_i,
        "pivot3_total_i": pivot3_total_i,
        "pivot2_total_i": pivot2_total_i,
        "valor_cinta_internacional": valor_cinta_internacional,
    }


def _layout_valor_cinta(v: dict) -> list:
    """
    Layout que replica tal cual el rango D66:K85 de la hoja "TABLA" del
    libro de estructura (ESTRUCTURA 1828 / ARCHIVO GENERAL 0030), compactado
    para empezar en la fila 1 (en el original arranca en la fila 66 porque
    queda debajo de las 5 tablas dinámicas de esa hoja, que aquí no existen).
    Los códigos HA32/HA26 y sus números de cuenta son fijos y se copian tal
    cual; las celdas que en el original dependían de digitación manual
    externa (NETO CONTABILIDAD, NETO, sumas HA32/HA26) quedan en blanco.
    Recibe el dict devuelto por _calcular_valor_cinta.
    """
    return [
        {"fila": 1, "columna": "G", "valor": "ESTRUCTURA 1828", "negrita": True},

        {"fila": 2, "columna": "E", "valor": "VALOR CINTA", "negrita": True},
        {"fila": 2, "columna": "F", "valor": v["valor_cinta_nacional"], "formato_numero": "#,##0.00"},
        {"fila": 2, "columna": "I", "valor": "PMD PRIMEROS ORIGINAL COMPRA"},

        {"fila": 3, "columna": "D", "valor": "HA32"},
        {"fila": 3, "columna": "E", "valor": "196095001828"},
        {"fila": 3, "columna": "G", "valor": "=+G2-H2", "formato_numero": "#,##0.00"},
        {"fila": 4, "columna": "D", "valor": "HA32"},
        {"fila": 4, "columna": "E", "valor": 299095241},
        {"fila": 4, "columna": "G", "valor": "=+F2+G3", "formato_numero": "#,##0.00"},
        {"fila": 5, "columna": "E", "valor": "NETO CONTABILIDAD"},
        {"fila": 5, "columna": "F", "valor": "=+F3-F4", "formato_numero": "#,##0.00"},
        {"fila": 6, "columna": "E", "valor": "NETO "},
        {"fila": 6, "columna": "F", "valor": "=+F5-F2", "formato_numero": "#,##0.00"},
        {"fila": 7, "columna": "D", "valor": "HA32"},
        {"fila": 7, "columna": "E", "valor": 299095837},
        {"fila": 7, "columna": "G", "valor": "dia actual"},
        {"fila": 8, "columna": "D", "valor": "HA26"},
        {"fila": 8, "columna": "E", "valor": "196095001215"},
        {"fila": 9, "columna": "D", "valor": "HA26"},
        {"fila": 9, "columna": "E", "valor": 299095237},

        {"fila": 12, "columna": "D", "valor": "ARCHIVO GENERAL 0030", "negrita": True},
        {"fila": 13, "columna": "F", "valor": "=+F6+F7+F8+F9+F12", "formato_numero": "#,##0.00"},

        {"fila": 16, "columna": "E", "valor": "VALOR CINTA", "negrita": True},
        {"fila": 16, "columna": "F", "valor": v["valor_cinta_internacional"], "formato_numero": "#,##0.00"},

        {"fila": 17, "columna": "D", "valor": "HA32"},
        {"fila": 17, "columna": "E", "valor": "196095001829"},
        {"fila": 18, "columna": "D", "valor": "HA32"},
        {"fila": 18, "columna": "E", "valor": 299095240},
        {"fila": 19, "columna": "E", "valor": "NETO CONTABILIDAD"},
        {"fila": 19, "columna": "F", "valor": "=+F17-F18", "formato_numero": "#,##0.00"},
        {"fila": 20, "columna": "E", "valor": "NETO "},
        {"fila": 20, "columna": "F", "valor": "=+F19-F16", "formato_numero": "#,##0.00"},
    ]


def _detalle_dinamicas(v: dict) -> pd.DataFrame:
    """
    Desglose auditable de las 5 tablas dinámicas (total general y los
    subtotales por COD-TIPO-TRANS/Indicador), para poder verificar cómo
    se llega a cada VALOR CINTA sin abrir el Excel original. Recibe el
    dict devuelto por _calcular_valor_cinta.
    """
    return pd.DataFrame([
        {"Concepto": "TIPO-REGISTRO 12/15 (sin más filtros)", "Valor": v["pivot5_total"]},
        {"Concepto": "TIPO-REG 01/21, MENSAJE 0210, COD-TIPO-TRANS 10/15, RESP 000-001, Indicador N", "Valor": v["pivot4_total_n"]},
        {"Concepto": "TIPO-REG 01/21, MENSAJE 0420, COD-TIPO-TRANS 14 (todos los indicadores, sin filtro de RESP)", "Valor": v["pivot1_total"]},
        {"Concepto": "menos: TIPO-REG 01/21, MENSAJE 0420, COD-TIPO-TRANS 10/14/15, RESP 000-001, Indicador N", "Valor": -v["pivot3_total_n"]},
        {"Concepto": "menos: TIPO-REG 01/21, MENSAJE 0210, COD-TIPO-TRANS 14, RESP 000-001, Indicador N", "Valor": -v["pivot2_total_n"]},
        {"Concepto": "VALOR CINTA NACIONAL", "Valor": v["valor_cinta_nacional"]},
        {"Concepto": "TIPO-REG 01/21, MENSAJE 0210, COD-TIPO-TRANS 10/15, RESP 000-001, Indicador I", "Valor": v["pivot4_total_i"]},
        {"Concepto": "menos: TIPO-REG 01/21, MENSAJE 0420, COD-TIPO-TRANS 10/14/15, RESP 000-001, Indicador I", "Valor": -v["pivot3_total_i"]},
        {"Concepto": "menos: TIPO-REG 01/21, MENSAJE 0210, COD-TIPO-TRANS 14, RESP 000-001, Indicador I", "Valor": -v["pivot2_total_i"]},
        {"Concepto": "VALOR CINTA INTERNACIONAL", "Valor": v["valor_cinta_internacional"]},
    ])


# ─────────────────────────────────────────────────────────────
# ARCHIVO 3 — VALOR CINTA (nacional e internacional)
# ─────────────────────────────────────────────────────────────
@manejar_errores
def generar_archivo_3(df: pd.DataFrame) -> dict:
    """
    "Hoja1" calcula el VALOR CINTA sobre TODO el df base combinado (suma
    de todos los días seleccionados) — nunca se segrega por día.
    "Detalle Dinamicas" sí se genera una hoja por día (columna "__dia"),
    para poder revisar el desglose de cada fecha por separado.
    """
    df = df.copy()
    if "MONTO-1" in df.columns:
        # MONTO-1 viene como texto de ancho fijo (13 dígitos) con 2 decimales
        # implícitos en los últimos 2 dígitos, ej. "0000001464000" -> 14640.00.
        df["MONTO-1"] = pd.to_numeric(df["MONTO-1"], errors="coerce") / 100

    hojas = {"Hoja1": _layout_valor_cinta(_calcular_valor_cinta(df))}

    if "__dia" in df.columns:
        for dia in sorted(df["__dia"].dropna().unique()):
            subconjunto = df[df["__dia"] == dia]
            hojas[f"Detalle Dinamicas {dia}"] = _detalle_dinamicas(_calcular_valor_cinta(subconjunto))
    else:
        hojas["Detalle Dinamicas"] = _detalle_dinamicas(_calcular_valor_cinta(df))

    return hojas
