# conexión directa

Herramienta de línea de comandos que procesa archivos planos de ancho fijo
con transacciones de tarjeta (formato bancario `GOF.GRB.FM14.FYYMMDD.txt`)
y genera tres reportes Excel independientes: soportes REDEBAN VISA,
rechazos EP DEPO y valores cinta.

Es el proyecto principal de esta familia; **"conexión directa conciso"**
es un hermano más pequeño que comparte el mismo layout de archivo plano y
arquitectura, pero genera un único libro de Excel en vez de tres.

## Qué hace

1. Lee uno o varios archivos planos (selección múltiple pensada para
   consolidar un cierre de fin de semana: sábado + domingo + lunes, o +
   martes si hay festivo, en una sola corrida).
2. Extrae los 55 campos del layout de ancho fijo por posición de carácter.
3. Genera, a elección del usuario, cualquiera de estos tres reportes
   (independientes entre sí, se pueden generar en cualquier orden):
   - **Archivo 1 — SOPORTES REDEBAN VISA**: hoja combinada "TIPO 50" +
     una hoja "TIPO 01 <día>" por cada día seleccionado.
   - **Archivo 2 — RECH EP DEPO**: hoja combinada "TIPO 12-15" + una hoja
     "RECH EP <día>" por cada día seleccionado.
   - **Archivo 3 — VALORES CINTA**: hoja fija "Hoja1" con la fórmula de
     VALOR CINTA (nacional + internacional) sobre el universo completo
     combinado, más una hoja "Detalle Dinámicas <día>" por cada día.
4. Escribe cada reporte en `.xlsx` con encabezado resaltado, ancho de
   columna automático, fila de encabezado congelada y formato de moneda
   en las columnas de monto. Cualquier hoja que supere 1,000,000 de filas
   se particiona automáticamente en varias hojas para que Excel pueda
   abrirla sin saturarse.

## Requisitos

No hay manifiesto de dependencias (no hay `requirements.txt` ni
`pyproject.toml`). Dependencias a instalar manualmente:

```
pip install pandas xlsxwriter
```

(`tkinter`, usado para el diálogo de selección de archivos, viene incluido
en la instalación estándar de Python en Windows.)

## Uso

```
python main.py
```

En Windows también se puede ejecutar con doble clic sobre
`orquestador.bat`, que se posiciona en la carpeta del script y corre
`python main.py`.

El programa muestra un menú de texto:

1. **Seleccionar archivo(s) plano(s)** — abre un diálogo de selección
   múltiple (Ctrl/Shift+clic); si no hay entorno gráfico disponible, cae a
   un prompt de consola separado por comas.
2. **Generar Archivo 1 (SOPORTES REDEBAN VISA)**
3. **Generar Archivo 2 (RECH EP DEPO)**
4. **Generar Archivo 3 (VALORES CINTA)**
5. **Salir**

Cada opción de generación relee los archivos seleccionados desde disco y
reconstruye la tabla base en cada corrida (sin caché entre opciones), para
que el reporte siempre refleje el contenido actual del archivo aunque se
haya reemplazado a mitad de sesión. Los archivos de salida se guardan en
`salidas/`, con la fecha del sistema en el nombre.

## Estructura del proyecto

```
main.py              punto de entrada, menú interactivo
orquestador.bat       lanzador para Windows
src/
  config.py           layout de campos (55 posiciones), rutas de salida
  lector.py            parseo del archivo plano a DataFrame (vectorizado)
  reglas.py            reglas de negocio: Archivo 1, Archivo 2, Archivo 3
  exportador.py         escritura del .xlsx con formato (XlsxWriter)
  validador.py           selección de archivos, extracción del día
  utils.py               manejo de errores y logging
docs/
  DOCUMENTACION_TECNICA.md   layout completo, reglas exactas, benchmarks
  MANUAL_USUARIO.md           manual de usuario en español
```

No hay suite de pruebas automatizadas, linter ni build configurados en
este repositorio.

## Documentación

- [`docs/DOCUMENTACION_TECNICA.md`](docs/DOCUMENTACION_TECNICA.md) — tabla
  completa del layout, condiciones exactas detrás de cada hoja de los tres
  reportes, derivación de la fórmula de VALOR CINTA y benchmarks de
  rendimiento.
- [`docs/MANUAL_USUARIO.md`](docs/MANUAL_USUARIO.md) — manual de usuario en
  español, con recorrido del menú, tabla de mensajes de error y el flujo
  recomendado de consolidación de fin de semana.
- [`CLAUDE.md`](CLAUDE.md) — guía técnica para trabajar en el código
  (arquitectura, convenciones, decisiones de diseño no obvias).
