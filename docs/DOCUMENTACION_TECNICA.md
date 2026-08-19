# Documentación técnica — Conexión Directa

Herramienta CLI de menú interactivo (sin interfaz gráfica propia, solo un diálogo nativo de Windows para elegir archivos) que procesa un archivo plano bancario de ancho fijo con transacciones REDEBAN VISA y genera 3 reportes Excel con reglas de negocio de conciliación.

Ruta del proyecto: `c:\Users\dacanonm\OneDrive - Indra\Documentos\GitHub\conexion directa`

Estructura de carpetas: `main.py` y `orquestador.bat` (punto de entrada)
viven en la raíz; los 6 módulos de negocio (`src/config.py`, `src/lector.py`,
`src/utils.py`, `src/validador.py`, `src/reglas.py`, `src/exportador.py`) viven en `src/`;
esta documentación y `MANUAL_USUARIO.md` viven en `docs/`. `main.py`
agrega `src/` a `sys.path` antes de importar, así que los imports internos
(`from config import ...`, etc.) no cambian de forma solo por la ubicación
del archivo.

## 1. Arquitectura general

Bucle de menú (`while True` en `main.py`) sin clases, con un único estado de sesión en memoria:

```python
estado = {"rutas": [], "df": None}
```

Grafo de dependencias entre módulos (todos dentro de `src/`, salvo
`main.py` que está en la raíz):

```
main.py
 ├─ src/config.py      (ARCHIVO_1, ARCHIVO_2, ARCHIVO_3, LAYOUT)
 ├─ src/lector.py      (construir_dataframe, leer_lineas) → config.py, utils.py
 ├─ src/validador.py   (extraer_dia, seleccionar_archivos)  → utils.py
 ├─ src/reglas.py      (generar_archivo_1/2/3)              → utils.py
 ├─ src/exportador.py  (guardar_con_formato)                → utils.py
 └─ src/utils.py       (manejar_errores, logger, limpiar_pantalla, pausar) → config.py (CARPETA_LOGS)
```

`src/config.py` es la base (no importa nada del proyecto). Todos los demás módulos dependen de `utils.manejar_errores`.

Pipeline de datos, disparado en cada una de las opciones 2/3/4 del menú:

```
archivo(s) plano(s) .txt (ancho fijo)
        │  validador.seleccionar_archivos()   → List[Path] (diálogo Tk multi-selección)
        │  validador.extraer_dia(ruta)        → día "DD" por archivo
        ▼
lector.leer_lineas(ruta)            → líneas de texto crudas
        ▼
lector.construir_dataframe(lineas)  → DataFrame (1 fila por línea, columnas = LAYOUT, todo str)
        │  main.py: df_parcial["__dia"] = dia   (por cada archivo)
        │  pd.concat(partes, ignore_index=True) → DataFrame base combinado
        ▼
reglas.generar_archivo_N(df_base)   → dict {"NombreHoja": DataFrame | lista_de_celdas}
        ▼
exportador.guardar_con_formato(ARCHIVO_N, hojas) → .xlsx en salidas/, devuelve True/None
```

Decisiones de diseño clave:

- **Sin caché entre opciones del menú**: cada vez que se elige generar el Archivo 1, 2 o 3, `_construir_dataframe_base()` (`main.py:65-110`) relee los archivos planos desde disco y reconstruye el DataFrame desde cero. Es deliberado, para no trabajar con datos desactualizados si el usuario reemplazó el archivo plano entre generaciones — el costo es releer archivos de ~130+ MB cada vez.
- **Archivo 3 es independiente de Archivo 1 y 2**: siempre parte del `df` base sin filtrar, nunca de los resultados de `generar_archivo_1`/`generar_archivo_2`.

## 2. Estructura del archivo plano de entrada (LAYOUT)

Definido en `src/config.py:24-83` como una lista de 55 dicts `{"nombre", "inicio", "longitud"}`. `inicio` es 1-indexado (posición del carácter) y `longitud` es la cantidad de caracteres que ocupa el campo (no una posición final).

Fórmula de parseo (`src/lector.py`, dentro de `construir_dataframe`): equivale exactamente a un `MID(texto, inicio, longitud)` de Excel — mismas posiciones, `inicio - 1` como offset 0-indexado. Desde la optimización de rendimiento para archivos de 800 mil a 1 millón de filas, la extracción ya no es un bucle Python campo por campo por línea (`linea[inicio-1:inicio-1+longitud].strip()` dentro de un `for` sobre `lineas`, con 55 slices/dict-inserts por fila) sino slicing vectorizado de pandas (`Series.str.slice`) sobre todas las líneas a la vez, una columna completa por campo:

```python
serie = pd.Series(lineas)
columnas = {
    campo["nombre"]: serie.str.slice(
        campo["inicio"] - 1, campo["inicio"] - 1 + campo["longitud"]
    ).str.strip()
    for campo in LAYOUT
}
df = pd.DataFrame(columnas)
```

`Series.str.slice` reemplaza el costo dominante del pipeline completo (el bucle Python de 55 campos × N filas, con ~50 objetos dict creados por línea) por 55 operaciones vectorizadas — una por campo, sobre la columna completa — dejando que pandas construya el DataFrame columna por columna en vez de a partir de una lista de dicts fila por fila. Semántica idéntica al slicing manual byte a byte: mismas posiciones, `colspecs` solapados entre sí siguen funcionando igual (ver nota de `DISPOSITIVO`/`PLANO` en la sección de campos), y una línea más corta que lo esperado produce campo vacío (`str.slice` sobre una cadena corta devuelve lo que alcance, igual que el slicing nativo de Python — nunca `NaN`, así que no hace falta ningún `fillna` especial). También maneja `lineas == []` sin caso especial: `pd.Series([])` seguido de `.str.slice(...).str.strip()` produce columnas vacías sin error. Todo se guarda como `str` — **`src/lector.py` no convierte tipos**; eso se hace explícitamente campo a campo en `src/reglas.py`, para mantener control total sobre el formato real.

### Los 55 campos

| # | nombre (literal en LAYOUT) | inicio | longitud |
|---|---|---|---|
| 1 | TIPO-REGISTRO | 1 | 2 |
| 2 | RED-LOGICA | 3 | 4 |
| 3 | FIID-AUTORIZA | 7 | 4 |
| 4 | NUMERO-TARJETA | 11 | 19 |
| 5 | RED-LOGICA-ALMACEN | 30 | 4 |
| 6 | FIID SPONSOR | 34 | 4 |
| 7 | CODIGO ALMACEN | 38 | 19 |
| 8 | CODIGO DATAFONO | 57 | 16 |
| 9 | TIPO DE MENSAJE | 73 | 4 |
| 10 | ORIGINA | 77 | 1 |
| 11 | RESPONDE | 78 | 1 |
| 12 | FECHA-TRANSACCION | 79 | 8 |
| 13 | HORA TRANSACCION | 87 | 8 |
| 14 | FECHA-POSTEO | 95 | 8 |
| 15 | NUMERO-SECUENCIA | 103 | 12 |
| 16 | UBICACIÓN DATAFONO | 115 | 25 |
| 17 | NOMBRE-ALMACEN | 140 | 22 |
| 18 | CIUDAD | 162 | 13 |
| 19 | DEPARTAMENTO | 175 | 2 |
| 20 | PAIS | 177 | 2 |
| 21 | COD-TIPO-TRANS | 179 | 2 |
| 22 | COD-TIPO-TARJETA | 181 | 1 |
| 23 | COD-TIPO-CTA | 182 | 2 |
| 24 | NUMERO - CUENTA | 184 | 19 |
| 25 | " CODIGO-RESP" (espacio inicial literal) | 203 | 3 |
| 26 | MONTO-1 | 206 | 13 |
| 27 | MONTO-2 | 219 | 13 |
| 28 | FECHA-VENCE-TARJETA | 232 | 6 |
| 29 | CODIGO-EMPRESA-PSP | 238 | 4 |
| 30 | NUMERO-FACTURA-PSP | 242 | 30 |
| 31 | ORIGEN-PSP | 272 | 1 |
| 32 | NUMERO-SEGUIMIENTO | 273 | 6 |
| 33 | NUMERO -APROBACION | 279 | 8 |
| 34 | DRAFT-CAPTURE-FLAG | 287 | 1 |
| 35 | CODIGO- REVERSO | 288 | 2 |
| 36 | MONEDA | 290 | 3 |
| 37 | NUMEROS - CUOTAS (espacio final literal) | 293 | 2 |
| 38 | COMISION - FCERA | 295 | 8 |
| 39 | COMISION-ADMIN | 303 | 4 |
| 40 | POR-RETENCION | 307 | 4 |
| 41 | POR- BASE-RETENCION | 311 | 4 |
| 42 | LIQUIDA-RETENCION | 315 | 10 |
| 43 | COMISION - FINANCIERA - AUTORIZADOR | 325 | 8 |
| 44 | COMISION -  FINANCIERA - ADQUIRIENTE (doble espacio literal) | 333 | 8 |
| 45 | LIQUIDA-IVA | 341 | 12 |
| 46 | MODO-INGRESO-POS | 353 | 3 |
| 47 | CODIGO-DE-SERVICIO-TARJETA | 356 | 3 |
| 48 | POR-RETEICA | 359 | 6 |
| 49 | LIQUIDA-RETEICA | 365 | 10 |
| 50 | FILLER-1 | 375 | 26 |
| 51 | NTLF-COMISION-FIN-EMP-ADICIONAL | 401 | 8 |
| 52 | FILLER-2 | 409 | 14 |
| 53 | PLANO | 423 | 125 |
| 54 | INDICADOR INTER/NACIONAL | 548 | 1 |
| 55 | DISPOSITIVO | 519 | 2 |

### Casos especiales del LAYOUT

- **Campos FILLER** (`FILLER-1`, `FILLER-2`) no se usan en ninguna regla de negocio; solo ocupan espacio para que el resto de posiciones cuadre con el spec del banco.
- **`MONTO-1`** (13 caracteres) trae 2 decimales implícitos en los últimos 2 dígitos. La conversión se repite de forma casi idéntica en las tres funciones públicas de `src/reglas.py`:
  ```python
  df["MONTO-1"] = pd.to_numeric(df["MONTO-1"], errors="coerce") / 100
  ```
  Ejemplo: `"0000001464000"` → `14640.00`.
- **Nombres de campo con espacios inconsistentes son literales e intencionales** — deben citarse tal cual en el código, no normalizarse: `" CODIGO-RESP"` (espacio inicial), `"NUMEROS - CUOTAS "` (espacio final), `"COMISION -  FINANCIERA - ADQUIRIENTE"` (doble espacio interno).
- **`DISPOSITIVO`** (inicio 519, longitud 2) cae numéricamente dentro del rango que ocupa `PLANO` (423-547). Es intencional según comentario explícito en `src/config.py:78-80` — cada campo se extrae de forma independiente por sus propias coordenadas, no hay conflicto técnico.
- `leer_lineas` solo descarta líneas vacías; no valida que cada línea tenga el ancho total esperado. Una línea más corta produce menos caracteres o cadena vacía en los campos finales, sin lanzar error.

## 3. Selección múltiple de archivos y etiquetado por día

**`validador.seleccionar_archivos()`** abre `tkinter.filedialog.askopenfilenames` (soporta Ctrl/Shift+clic) filtrado a `*.txt`. Si tkinter no está disponible o falla, cae en un fallback de consola que pide rutas separadas por coma. Valida existencia de cada ruta (`FileNotFoundError` si no) y devuelve `List[Path]`.

**Patrón de día** — `PATRON_DIA = re.compile(r"F(\d{2})(\d{2})(\d{2})", re.IGNORECASE)`, diseñado para `GOF.GRB.FM14.FYYMMDD.txt`. `extraer_dia(ruta)` busca el patrón sobre `ruta.stem` con `.search()` y devuelve el **tercer grupo capturado** (DD), descartando año y mes. Ejemplo: `GOF.GRB.FM14.F260813.txt` → grupos `("26","08","13")` → `"13"`.

**Fallback interactivo** si el nombre no calza: imprime advertencia y entra en un bucle pidiendo el día por consola (1-2 dígitos), validado con `dia.isdigit() and 1 <= len(dia) <= 2` y normalizado con `.zfill(2)`. Nunca asume un valor por defecto — ningún archivo queda mal etiquetado en silencio.

**Propagación de `__dia`**: ocurre en `main.py:_construir_dataframe_base()`, no en `src/validador.py` ni `src/lector.py`. Por cada ruta: se extrae el día, se leen las líneas, se construye el DataFrame parcial, se le asigna `df_parcial["__dia"] = dia` a todas sus filas, y se acumula. Al final: `df = pd.concat(partes, ignore_index=True)`. Si un archivo falla en cualquier paso, se omite con mensaje y se continúa con los demás.

### Segregación por día vs. combinado (regla central de `src/reglas.py`)

`_segregar_por_dia(df_base, hoja_filtrada, nombre_base)` genera una hoja `"NombreBase DD"` por cada día **presente en `df_base`** (no solo los que sobrevivieron el filtro de negocio), para que un día sin resultados igual produzca su hoja vacía con advertencia, en vez de desaparecer silenciosamente. Si `df_base` no tiene columna `__dia`, devuelve una sola hoja sin sufijo.

| Se segrega por día (una hoja por día) | Se mantiene combinada (todos los días juntos) |
|---|---|
| "TIPO 01" (Archivo 1) | "TIPO 50" (Archivo 1) |
| "RECH EP" (Archivo 2) | "TIPO 12-15" (Archivo 2) |
| "Detalle Dinamicas" (Archivo 3) | "Hoja1" / VALOR CINTA (Archivo 3) |

El sufijo de día se aplica **siempre**, incluso si solo se seleccionó un archivo — comportamiento consistente, no hay un "modo archivo único" separado.

## 4. Reglas de negocio exactas (`src/reglas.py`)

### Archivo 1 — `generar_archivo_1(df)` → hojas "TIPO 50" + "TIPO 01 \<día\>"

Conversión inicial: `MONTO-1` a numérico/100.

**Hoja "TIPO 50"** (combinada):
```python
cond_red = RED-LOGICA isin ["VISN", "VISA"]
cond_tipo_50 = TIPO-REGISTRO == "50"
hoja1 = df[cond_red & cond_tipo_50]
```

**Hoja "TIPO 01 \<día\>"** (segregada):
```python
cond_red            = RED-LOGICA isin ["VISN", "VISA"]      # también aplica aquí
cond_tipo_01         = TIPO-REGISTRO == "01"
cond_FIID_SPONSOR    = FIID SPONSOR == "0013"
cond_TIPO_DE_MENSAJE = TIPO DE MENSAJE isin ["0210", "0420"]
cond_COD_TIPO_TRANS  = COD-TIPO-TRANS isin ["10", "14", "15"]
cond_CODIGO_RESP     = " CODIGO-RESP" between "000" y "009"  (comparación de string)
hoja2 = df[cond_red & cond_tipo_01 & cond_FIID_SPONSOR & cond_TIPO_DE_MENSAJE & cond_COD_TIPO_TRANS & cond_CODIGO_RESP]
```

**Negativización de MONTO-1** (solo en hoja2): si el mensaje es un reverso (`TIPO DE MENSAJE == "0420"`) **o** `COD-TIPO-TRANS == "14"`, el monto se fuerza a negativo con `-abs()`:
```python
cond_negativo = (TIPO_DE_MENSAJE.eq("0420")) | (COD_TIPO_TRANS.eq("14"))
hoja2.loc[cond_negativo, "MONTO-1"] = -hoja2.loc[cond_negativo, "MONTO-1"].abs()
```
Se usa `-abs()` (no negación doble) para que una fila que cumpla ambas condiciones no quede positiva otra vez.

**Orden final**: `hoja2.sort_values("COD-TIPO-TRANS", kind="stable")` — merge sort estable, preserva el orden relativo original en empates.

### Archivo 2 — `generar_archivo_2(df)` → hojas "TIPO 12-15" + "RECH EP \<día\>"

Misma conversión inicial de MONTO-1.

**Hoja "TIPO 12-15"** (combinada), sin más filtros:
```python
hoja1 = df[TIPO-REGISTRO isin ["12", "15"]]
```

**Hoja "RECH EP \<día\>"** (segregada):
```python
cond_tipo_01_20      = TIPO-REGISTRO isin ["01", "20"]        # nota: "20", no "21"
cond_FIID_SPONSOR    = FIID SPONSOR == "0013"
cond_TIPO_DE_MENSAJE = TIPO DE MENSAJE isin ["0210", "0420"]
cond_COD_TIPO_TRANS  = COD-TIPO-TRANS isin ["10", "14", "15"]
cond_CODIGO_RESP     = " CODIGO-RESP" between "000" y "009"
hoja2 = df[cond_tipo_01_20 & cond_FIID_SPONSOR & cond_TIPO_DE_MENSAJE & cond_COD_TIPO_TRANS & cond_CODIGO_RESP]
```
A diferencia de Archivo 1: no filtra `RED-LOGICA`, no hay negativización de MONTO-1, no hay `sort_values` explícito.

> **Nota de auditoría**: Archivo 2 usa `TIPO-REGISTRO isin ["01","20"]`, mientras que los pivots de Archivo 3 (VALOR CINTA) usan `isin ["01","21"]` — códigos "20" y "21" son literalmente distintos. Está así en ambos bloques del código fuente; conviene confirmar con negocio si es intencional o un typo heredado del Excel original.

### Archivo 3 — `generar_archivo_3(df)` → "Hoja1" (layout libre) + "Detalle Dinamicas \<día\>" (tabular)

Replica las fórmulas `GETPIVOTDATA` reales de la hoja "TABLA" del libro de estructura del banco (celdas F67 y F81 del original). Siempre trabaja sobre el `df` base **sin filtrar**, independiente de Archivo 1/2.

`_calcular_valor_cinta(df)` traduce 5 tablas dinámicas del Excel original a máscaras booleanas sobre `MONTO-1`:

| Término | Condiciones | Nace de |
|---|---|---|
| `pivot5_total` | `TIPO-REGISTRO isin [12,15]` — sin más filtros | Tabla dinámica en `$A$3` |
| `pivot4_total_n` / `pivot4_total_i` | `FIID=0013 & TIPO-REGISTRO isin[01,21] & TIPO DE MENSAJE=0210 & COD-TIPO-TRANS isin[10,15] & CODIGO-RESP isin[000,001]`, separado por `INDICADOR INTER/NACIONAL` (N / I) | Tabla dinámica en `$E$4` |
| `pivot1_total` | `FIID=0013 & TIPO-REGISTRO isin[01,21] & TIPO DE MENSAJE=0420 & COD-TIPO-TRANS=14` — **sin** filtro de CODIGO-RESP ni de indicador (suma N+I+blanco todo junto) | Tabla dinámica en `$E$56` |
| `pivot3_total_n` / `pivot3_total_i` | `FIID=0013 & TIPO-REGISTRO isin[01,21] & TIPO DE MENSAJE=0420 & COD-TIPO-TRANS isin[10,14,15] & CODIGO-RESP isin[000,001]` | Tabla dinámica en `$E$27` |
| `pivot2_total_n` / `pivot2_total_i` | `FIID=0013 & TIPO-REGISTRO isin[01,21] & TIPO DE MENSAJE=0210 & COD-TIPO-TRANS=14 & CODIGO-RESP isin[000,001]` | Tabla dinámica en `$E$42` |

Fórmulas finales:
```python
valor_cinta_nacional      = pivot5_total + pivot4_total_n + pivot1_total - pivot3_total_n - pivot2_total_n
valor_cinta_internacional = pivot4_total_i - pivot3_total_i - pivot2_total_i
```

Este mapeo fue **verificado exacto** contrastando contra los valores cacheados en las celdas del Excel original (F67 = 70,787,889,889.00, F81 = 4,201,399,110.00), reconstruyendo los 916,357 registros del pivot cache real del libro `ejemplo estructura.xlsx`. Antes de esa verificación faltaban los filtros de `TIPO-REGISTRO isin[01,21]` y de `TIPO DE MENSAJE` en todos los términos, lo cual inflaba el valor nacional en ~16,171 millones y arrojaba signo/magnitud equivocados en el internacional — corregido en esta sesión.

`_layout_valor_cinta(v)` genera la hoja "Hoja1" como lista de celdas de posición libre (`{"fila","columna","valor","negrita"?,"formato_numero"?}`), replicando el rango D66:K85 del libro original compactado a partir de la fila 1: rótulos, códigos de cuenta fijos HA32/HA26 (`196095001828`, `299095241`, `299095837`, `196095001215`, `299095237`, `196095001829`, `299095240`), y fórmulas Excel literales encadenadas (`=+G2-H2`, `=+F2+G3`, `=+F3-F4`, `=+F5-F2`, `=+F6+F7+F8+F9+F12`, `=+F17-F18`, `=+F19-F16`). Las celdas que en el original dependían de digitación manual externa (NETO CONTABILIDAD, NETO, sumas HA32/HA26) quedan en blanco a propósito.

`_detalle_dinamicas(v)` produce un DataFrame `["Concepto","Valor"]` con el desglose auditable de los 10 componentes (5 nacionales + 5 internacionales, incluyendo cada total), describiendo el filtro real aplicado en cada concepto.

En `generar_archivo_3`: "Hoja1" se calcula una sola vez sobre todo `df` (nunca se segrega por día); "Detalle Dinamicas" se genera una vez por cada día en `__dia`.

## 5. Manejo de errores

`utils.manejar_errores` es un decorador (`functools.wraps`) que envuelve la función objetivo en `try/except` con 4 ramas, de más a menos específica, todas devolviendo `None` en vez de propagar la excepción:

| Excepción | Mensaje al usuario | Log |
|---|---|---|
| `FileNotFoundError` | `[!] No se encontró el archivo: {e}` | `logger.error` |
| `PermissionError` | `[!] Sin permisos de acceso (¿el archivo está abierto en Excel?): {e}` | `logger.error` |
| `ValueError` | `[!] Error de datos o configuración: {e}` | `logger.error` |
| `Exception` genérico | `[!] Ocurrió un error inesperado: {e}` | `logger.exception` (traceback completo) |

Aplicado a las funciones públicas: `leer_lineas`, `construir_dataframe`, `generar_archivo_1/2/3`, `guardar_con_formato`, `seleccionar_archivos`, `extraer_dia`. Las funciones internas de `src/reglas.py` (`_segregar_por_dia`, `_calcular_valor_cinta`, `_layout_valor_cinta`, `_detalle_dinamicas`) no llevan el decorador — su fallo se propaga a la función pública que sí lo tiene.

**Señal de éxito/fracaso explícita**: `guardar_con_formato` devuelve `True` solo si `wb.save(path)` completó sin excepción; si `manejar_errores` capturó algo (ej. `PermissionError` por archivo abierto en Excel), nunca llega al `return True` y el decorador devuelve `None`. `main.py` usa `if not guardar_con_formato(...)` en las 3 opciones de generación para distinguir "se guardó" de "falló silenciosamente" y no mostrar un falso `[OK]` tras un fallo real.

## 6. Generación de Excel (`src/exportador.py`)

`guardar_con_formato(path, hojas)` crea un único `openpyxl.Workbook(write_only=True)` (no `pd.ExcelWriter`, no un `Workbook` normal) para **todas** las hojas, sean `pd.DataFrame` (`_escribir_hoja_dataframe`) o de posición libre (`_escribir_hoja_libre`). El modo `write_only` transmite cada fila directamente al XML de salida sin mantener el worksheet completo en memoria — mucho más rápido y liviano con hojas de cientos de miles de filas que `DataFrame.to_excel(engine="openpyxl")`, que asigna celda por celda (`ws.cell(row, col, value)` una vez por cada una de las ~55 columnas × cada fila).

La contrapartida de `write_only` es que **no admite acceso aleatorio a celdas** (`ws["D3"] = valor`, usado antes para la hoja de posición libre "Hoja1" de Archivo 3) ni volver a tocar una fila ya escrita — se serializa en cuanto se hace `ws.append()`. Por eso:
- El encabezado y el formato de moneda/fecha se arman **antes** de cada `append`, envolviendo el valor en `openpyxl.cell.WriteOnlyCell` (con `.fill`/`.font`/`.alignment` para el encabezado, `.number_format` para moneda/fecha) en vez de asignarlos después.
- `freeze_panes` y `column_dimensions` también deben fijarse **antes** del primer `append` de cada hoja — verificado en la versión de openpyxl usada en este proyecto: asignados después del primer `append`, se ignoran silenciosamente (sin error, pero no se guardan en el `.xlsx`).
- **`_escribir_hoja_libre`** arma una grilla completa en memoria (`list[list]`, de la fila 1 a la última fila usada, columna A a la última columna usada, con `None` en las posiciones sin celda definida) a partir de la lista de celdas de posición libre, y la vuelca fila por fila con `ws.append(fila)` — las columnas conservan su letra real (D, E, F...) para que la plantilla replicada (ver `_layout_valor_cinta` en `src/reglas.py`) no cambie de posición. Esto permite que hasta "Hoja1" (no tabular) también use `write_only`, sin necesitar un `Workbook` normal como camino alternativo.

**`_escribir_hoja_dataframe`**: `_clasificar_columnas(df)` determina, por nombre de columna, qué formato de número aplica (`"monto"`/`"valor"` en el nombre, minúsculas → `#,##0.00`; `"fecha"` → `DD/MM/YYYY`). `_anchos_columnas(df)` calcula el ancho de cada columna con una **muestra** de las primeras `_FILAS_MUESTRA_ANCHO` (2000) filas (`df.head(2000)`), no con el DataFrame completo — con hojas de cientos de miles o millones de filas (ej. un consolidado de 3-4 archivos planos de ~1M filas cada uno), recorrer las 55 columnas completas solo para un ancho visual aproximado es un costo real; el ancho es cosmético, nunca afecta los datos. El `fillna("")` antes de `astype(str)` es necesario porque, en pandas 3.x (la versión instalada en este entorno), `astype(str)` sobre un `NaN` de una columna `float64` (ej. `MONTO-1` cuando `pd.to_numeric(..., errors="coerce")` no pudo convertir el valor) **ya no produce la cadena `"nan"`** como en pandas < 3 — deja el `NaN` como `float`, y `.map(len)` revienta con `TypeError: object of type 'float' has no len()`; esto rompía la generación de cualquier hoja con al menos una fila cuyo `MONTO-1` no se pudo convertir, un caso real y común (filas con el campo en blanco). Los valores de fila se preparan con `df.astype(object).where(df.notna(), None).values.tolist()` (`NaN`/`NaT` → `None`, porque openpyxl no acepta `NaN` como valor de celda) y se escriben con `ws.append(fila)`. Solo se envuelven en `WriteOnlyCell` las celdas de las columnas que sí necesitan `number_format` (`indices_formato`, precalculado una sola vez por hoja) — normalmente `MONTO-1` y las columnas cuyo nombre contiene `"fecha"` (`FECHA-TRANSACCION`, `FECHA-POSTEO`, `FECHA-VENCE-TARJETA`); el resto de las ~55 columnas de cada fila se escribe sin tocar. Antes se recorrían las 55 columnas de cada fila para decidir esto (`zip(fila, formato_por_columna)`); con hojas de varios millones de filas eso es una cantidad significativa de iteraciones Python-level desperdiciadas.

> **Nota de rendimiento pendiente de decidir con el usuario**: las columnas `"fecha"` reciben `number_format = "DD/MM/YYYY"` aunque `lector.py`/`reglas.py` nunca convierten esos campos a un tipo fecha real (siguen siendo texto, ej. `"20260813"`) — Excel ignora `number_format` sobre celdas de texto, así que ese formato hoy no tiene ningún efecto visual, pero igual cuesta construir un `WriteOnlyCell` por cada una de esas celdas. Se detectó al medir el benchmark de la sección 6c (~1.75M filas combinadas en Archivo 2 tardaron 690s en escribirse) pero no se eliminó porque cambia semántica fuera del alcance aprobado — si se confirma que ningún flujo futuro va a convertir esos campos a fecha real, se podría excluir las columnas `"fecha"` de `indices_formato` cuando su dtype siga siendo texto, para no pagar ese costo sin beneficio visual.

**Partición de hojas de más de 1,000,000 de filas**: Excel tiene un límite real de 1,048,576 filas por hoja (encabezado incluido). `_particionar_hoja_grande(nombre_hoja, df)` se aplica a cada hoja `pd.DataFrame` justo antes de escribirse (dentro de `guardar_con_formato`, antes de crear el `Workbook`): si `len(df) > LIMITE_FILAS_HOJA` (constante = 1,000,000), la divide en `math.ceil(len(df) / LIMITE_FILAS_HOJA)` hojas consecutivas de hasta 1,000,000 filas cada una, **sin reordenar filas** (bloques secuenciales vía `df.iloc[inicio:fin]`), nombradas `"{nombre_hoja} (1)"`, `"{nombre_hoja} (2)"`, etc. — no siempre 2 partes: si una hoja llegara a superar 2,000,000 de filas, se generarían 3 o más. Los nombres de parte se recortan a 31 caracteres (límite de Excel para nombres de hoja) como resguardo, aunque con los nombres de hoja actuales del proyecto (ej. `"TIPO 50"`, `"RECH EP 13"`) nunca se acerca a ese límite. Cada parte se escribe como una hoja independiente completa (encabezado, formato de moneda/fecha, ancho de columna, `freeze_panes`), igual que cualquier otra hoja. Esto se hace de forma genérica en `exportador.py` (no en `src/reglas.py`) para las hojas de **cualquiera** de los 3 archivos, aunque en la práctica solo "TIPO 50" (Archivo 1) y "TIPO 12-15" (Archivo 2) —las combinadas que nunca se segregan por día— tienen probabilidad real de acercarse al límite. Cuando esto ocurre se avisa en consola y en el log (`logger.info`), igual que el aviso ya existente de hoja vacía. Las hojas de posición libre (`_escribir_hoja_libre`) nunca se particionan — son fijas y pequeñas por diseño.

**Hojas vacías**: se escriben igual (con encabezados), pero se advierte en consola y en el log por cada una — no aborta la generación.

**Nombres de archivos de salida** (`src/config.py:96-98`):
```python
ARCHIVO_1 = "salidas/SOPORTES REDEBAN VISA {DDMMYYYY}.xlsx"
ARCHIVO_2 = "salidas/RECH EP DEPO {DDMMYYYY}.xlsx"
ARCHIVO_3 = "salidas/{DDMMYY}.xlsx"   # sin prefijo, año a 2 dígitos
```
Usan la fecha del **sistema al momento de ejecutar**, no la fecha del archivo plano. `CARPETA_SALIDA`/`CARPETA_LOGS` se crean automáticamente y son relativas al directorio de trabajo — por eso `orquestador.bat` hace `cd /d "%~dp0"` antes de invocar `python main.py`.

## 6b. Rendimiento con archivos grandes (800 mil - 1 millón de filas)

El pipeline fue optimizado para que un archivo plano de 800 mil a 1 millón de filas se procese en minutos, no en decenas de minutos, **sin tocar ninguna regla de negocio ni filtro** — los cambios son puramente de cómo se parsea el archivo y cómo se escribe el Excel, no de qué filas o valores calcula el programa. Esta optimización se hizo en dos rondas independientes (una local, otra subida directamente al remoto) que llegaron a la misma idea general por caminos distintos; se resolvió el conflicto de fusión quedándose con la versión del remoto por ser más simple, y corrigiendo en el proceso un bug real de compatibilidad con pandas 3.x que esa versión tenía (ver sección 6).

Los dos cuellos de botella reales, identificados y corregidos:

1. **Parseo del archivo plano** (`src/lector.py:construir_dataframe`): antes era un bucle Python que, por cada línea, hacía 55 slices + `.strip()` + inserciones en un dict — con 1 millón de filas, 55 millones de operaciones a nivel de bytecode Python. Se reemplazó por slicing vectorizado de pandas (`Series.str.slice` sobre la columna completa de líneas, una vez por campo). Ver sección 2.
2. **Escritura del Excel** (`src/exportador.py:guardar_con_formato`): antes usaba `DataFrame.to_excel(engine="openpyxl")`, que asigna celda por celda (`ws.cell(row, col, value)` por cada una de las ~55 columnas × cada fila), más un recorrido `astype(str).map(len)` sobre columnas completas solo para calcular el ancho de columna, más un segundo recorrido completo con `ws.iter_rows()` para aplicar el formato de moneda/fecha. Se reemplazó por `openpyxl.Workbook(write_only=True)` con escritura fila-por-fila vía `ws.append()`, que no retiene los objetos `Cell` en memoria y escala mejor que un `Workbook` normal con hojas grandes. Ver sección 6.

**Benchmark** (1,000,000 de filas sintéticas, distribución pareja entre categorías — máquina de esta sesión, no representa necesariamente el hardware de producción; medido con la versión final del código, después de fusionar ambas rondas de optimización y corregir el bug de `_anchos_columnas`):

| Paso | Tiempo |
|---|---|
| `construir_dataframe` (parseo, `Series.str.slice`) | ~16s |
| `generar_archivo_1` (reglas) + escritura ("TIPO 50": 111,207 filas, "TIPO 01": 83,492 filas) | ~1s + ~59s |
| `generar_archivo_2` (reglas) + escritura ("TIPO 12-15": 333,705 filas, "RECH EP": 250,438 filas) | ~1s + ~179s |
| `generar_archivo_3` (reglas) + escritura (VALOR CINTA, tablas pequeñas) | ~2s + <1s |
| **Total (parseo + los 3 archivos)** | **~259s ≈ 4.3 min** |

Con archivos reales de producción, donde las categorías de `TIPO-REGISTRO` no se reparten parejo entre 6 valores (una sola categoría puede concentrar la mayoría de las filas), la mejora de `write_only` es la que más importa: su ventaja crece con el tamaño de la hoja en vez de degradarse, a diferencia de un `Workbook` normal que retiene todas las celdas en memoria.

## 6c. Indicador de tiempo y rendimiento con consolidados de 3-4 archivos (2-4 millones de filas combinadas)

**Indicador de tiempo en consola** (`main.py`): `_construir_dataframe_base()` imprime la hora (`HH:MM:SS`) al iniciar la lectura de los archivos seleccionados y el tiempo transcurrido (`time.perf_counter()`) al terminar de construir el DataFrame combinado. Cada `opcion_generar_archivo_N()` imprime la hora justo antes de `"Aplicando reglas de negocio..."` — el momento exacto en que empieza el filtrado — y el mensaje final `"[OK] Archivo N generado."` ahora incluye el tiempo total transcurrido (reglas + escritura). Así el usuario ve en pantalla, sin abrir el log, cuánto está tardando cada corrida mientras procesa consolidados grandes.

**Dos optimizaciones adicionales en `src/exportador.py`** (más allá de las de la sección 6b), motivadas por el escenario real de consolidar hasta 3-4 archivos planos de ~900 mil a 1 millón de filas cada uno (2-4 millones de filas combinadas por corrida):

1. `_anchos_columnas` pasó de recorrer el DataFrame completo a una muestra de 2000 filas (ver sección 6).
2. El bucle de escritura de `_escribir_hoja_dataframe` pasó de recorrer las 55 columnas de cada fila a solo las columnas en `indices_formato` (normalmente `MONTO-1` + las columnas `"fecha"`, ver sección 6).

**Benchmark** (3 archivos sintéticos de 1,000,000 de filas cada uno, 3,000,000 combinadas, distribución pareja entre categorías — máquina de esta sesión):

| Paso | Tiempo |
|---|---|
| `construir_dataframe` (parseo de las 3,000,000 de filas combinadas) | ~59s |
| `generar_archivo_1` (reglas) | ~7s |
| `guardar_con_formato` Archivo 1 ("TIPO 50": 333,581 filas, "TIPO 01": 250,214 filas) | ~198s |
| `generar_archivo_2` (reglas) | ~9s |
| `guardar_con_formato` Archivo 2 ("TIPO 12-15": 999,737 filas, "RECH EP": 750,588 filas) | ~690s |

El costo por fila escrita se mantiene estable entre Archivo 1 (~584 mil filas de salida en 198s ≈ 0.34 ms/fila) y Archivo 2 (~1.75 millones de filas de salida en 690s ≈ 0.39 ms/fila) — es decir, la escritura escala de forma aproximadamente lineal con el volumen de salida tras estas dos optimizaciones, sin la degradación no lineal que tenía un `Workbook` normal (ver sección 6b). El tiempo de Archivo 2 es alto en términos absolutos (~11.5 min) simplemente porque su volumen de salida (~1.75M filas) casi triplica el de Archivo 1 (~584 mil) en este escenario sintético — no es una regresión de la optimización, es proporcional al tamaño real de la hoja.

**Optimización adicional identificada, no aplicada** (fuera del alcance acordado con el usuario): las columnas `"fecha"` reciben `WriteOnlyCell` con `number_format = "DD/MM/YYYY"` en cada fila aunque su valor siga siendo texto sin convertir (ver nota en sección 6) — Excel ignora ese formato sobre texto, así que hoy no tiene efecto visual, pero sigue costando tiempo de escritura. Si se confirma que ningún flujo futuro necesita ese formato mientras las fechas sigan sin convertirse a tipo fecha real, excluir esas columnas de `indices_formato` recortaría aún más el tiempo de escritura de Archivo 1 y Archivo 2.

## 7. Dependencias externas

| Paquete | Usado en | Propósito |
|---|---|---|
| `pandas` | `main.py`, `src/lector.py`, `src/reglas.py`, `src/exportador.py` | DataFrame, `concat`, `to_numeric`, `Series.str.slice` |
| `openpyxl` | `src/exportador.py` | `Workbook(write_only=True)`, `WriteOnlyCell`, estilos manuales |
| `tkinter` (stdlib, import opcional con fallback) | `src/validador.py` | Diálogo de selección múltiple de archivos |

Resto: módulos estándar (`pathlib`, `datetime`, `re`, `typing`, `functools`, `logging`, `os`). No existe `requirements.txt`, `pyproject.toml` ni `.gitignore` en la raíz — las dependencias (`pandas`, `openpyxl`) deben instalarse manualmente en el entorno Python que ejecuta el programa (`tkinter` viene con la instalación estándar de Python en Windows).

## 8. Convenciones y decisiones de diseño no obvias

- **Sin caché entre opciones** y **Archivo 3 independiente de 1/2**: ver sección 1.
- **`RED-LOGICA` solo se filtra en Archivo 1**, no en Archivo 2 ni en los pivots de Archivo 3 — asimetría real del negocio (está en el propio nombre del reporte: "SOPORTES REDEBAN VISA" vs "RECH EP DEPO"), no un olvido.
- **`kind="stable"`** en el único `sort_values` del proyecto (Archivo 1, hoja TIPO 01): preserva el orden relativo original de filas con el mismo `COD-TIPO-TRANS`.
- **Todo se guarda como string en `src/lector.py`**; las conversiones de tipo (solo `MONTO-1`, no hay conversión de fechas en ningún punto) se repiten casi idénticas en las tres funciones públicas de `src/reglas.py` en vez de centralizarse — redundancia menor, no afecta el comportamiento.
- **`_segregar_por_dia` itera los días de `df_base`, no de la hoja ya filtrada** — para que un día sin resultados de negocio aún produzca una hoja vacía con advertencia, en vez de desaparecer sin explicación.
- **Formato moneda/fecha por heurística de nombre de columna**: cualquier columna futura que contenga `"monto"`/`"valor"`/`"fecha"` en su nombre (minúsculas) hereda el formato automáticamente, sin tocar `src/exportador.py`.
- **`_layout_valor_cinta` deja celdas en blanco a propósito** para valores que en el Excel original requerían digitación manual externa — el programa no intenta inventarlos ni calcularlos porque están fuera del alcance de las fórmulas `GETPIVOTDATA` reconstruidas.
- Los archivos `ejemplo_datos.xlsx`, `ejemplo estructura.xlsx` y `ESTRUC CINTA POS DEPOS FINAL V3.xlsx` en la raíz **no se abren por código** — son insumos de referencia humana que se usaron para escribir el LAYOUT y las máscaras booleanas de VALOR CINTA, no se leen en tiempo de ejecución.

## Archivos del proyecto

| Archivo | Responsabilidad |
|---|---|
| `main.py` | Menú interactivo, orquestación, estado de sesión |
| `src/config.py` | LAYOUT (55 campos), rutas de carpetas y de archivos de salida |
| `src/lector.py` | Lectura del archivo plano y construcción del DataFrame crudo (todo string) |
| `src/validador.py` | Selección de archivo(s) plano(s), extracción/confirmación del día |
| `src/reglas.py` | Reglas de negocio de los 3 archivos de salida |
| `src/exportador.py` | Escritura y formato de los archivos Excel de salida |
| `src/utils.py` | Decorador `manejar_errores`, logging, utilidades de consola |
| `orquestador.bat` | Punto de entrada para el usuario final (doble clic) |
| `CLAUDE.md` | Guía de orientación para agentes/desarrolladores que trabajen en el repo |
| `docs/DOCUMENTACION_TECNICA.md` | Este documento |
| `docs/MANUAL_USUARIO.md` | Manual de usuario final en español |
