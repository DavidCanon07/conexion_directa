# Manual de usuario — Conexión Directa

## ¿Qué hace este programa?

Toma el archivo plano diario que llega del banco/REDEBAN con las transacciones de tarjetas VISA (compras, reversos, rechazos) y lo convierte automáticamente en 3 reportes Excel listos para usar: uno de soportes REDEBAN VISA, uno de rechazos/depósitos EP, y uno con el "valor cinta" (el valor total nacional e internacional de la cinta procesada, usado para conciliar contra contabilidad). Reemplaza el trabajo manual de aplicar filtros y fórmulas en Excel sobre el archivo plano.

## Requisitos previos

- El equipo debe tener **Python instalado**, con las librerías `pandas` y `openpyxl` (esto lo prepara una sola vez el área de sistemas/soporte técnico, no es algo que deba hacer el usuario en cada uso).
- Debe tener **Microsoft Excel** (u otro programa compatible) instalado para poder abrir los 3 reportes generados.
- Debe conservarse la carpeta completa del proyecto en el equipo (con todos sus archivos).

## Cómo se ejecuta

Hacer **doble clic en `orquestador.bat`**. No hace falta abrir una consola ni escribir comandos: ese archivo se encarga de ubicarse en la carpeta correcta y arrancar el programa. Al terminar (cuando se elige la opción "Salir"), la ventana queda abierta para poder leer los últimos mensajes antes de cerrarla.

> Si al hacer doble clic la ventana se abre y se cierra sola muy rápido, probablemente falta instalar Python en ese equipo — avisar a soporte técnico.

## El menú principal

Al abrir el programa se ve esto:

```
=======================================================
 PROCESAMIENTO DE ARCHIVO PLANO — CONEXION DIRECTA
=======================================================
 Archivos actuales: (ninguno seleccionado)
-------------------------------------------------------
 1. Seleccionar archivo(s) plano(s)
 2. Generar Archivo 1 (SOPORTES REDEBAN VISA)
 3. Generar Archivo 2 (RECH EP DEPO)
 4. Generar Archivo 3 (DDMMYY) - VALORES CINTA
 5. Salir
=======================================================
Selecciona una opción:
```

Una vez se seleccionan archivos, la línea "Archivos actuales" cambia y lista cada archivo elegido debajo — es la única confirmación visual de qué está cargado en la sesión.

**Orden de uso normal**: primero la opción **1** (siempre, incluso si se va a repetir el proceso), luego 2, 3 y 4 en el orden que se quiera — no dependen entre sí, cada una vuelve a leer el/los archivo(s) plano(s) desde cero. Al final, opción **5** para salir.

Si se escribe algo que no es 1-5, aparece `Opción no válida.` y se vuelve al menú.

## Seleccionar el o los archivos planos (opción 1)

Al elegir la opción 1 se abre la ventana estándar de Windows para elegir archivos, filtrada a archivos `.txt`.

- **Para elegir varios archivos a la vez**: mantener presionado **Ctrl** (para archivos sueltos) o **Shift** (para un rango consecutivo) mientras se hace clic — igual que en cualquier ventana de Windows.
- **¿Para qué sirve elegir varios?** El caso típico es consolidar el cierre de **fin de semana** en una sola corrida: seleccionar juntos el archivo del sábado, domingo y lunes (y también el martes si hubo festivo de por medio), en vez de procesar cada día por separado. El programa junta todas las transacciones en una sola tabla interna, pero mantiene registro de a qué día pertenece cada una.
- Si se cancela el diálogo sin elegir nada, el programa avisa "No se seleccionó ningún archivo." y no cambia nada.

### Nombre esperado del archivo

El programa espera el patrón `GOF.GRB.FM14.FYYMMDD.txt` (año-mes-día, 2 dígitos cada uno), por ejemplo `GOF.GRB.FM14.F260813.txt` (archivo del día 13) o `GOF.GRB.FM14.F260814.txt` (archivo del día 14). De ese nombre, el programa solo usa los **2 últimos dígitos** (el día) para etiquetar las transacciones de ese archivo.

**Si el nombre del archivo no sigue ese patrón**, el programa no lo descarta ni asume nada por su cuenta — en su lugar **pregunta el día por consola**, archivo por archivo:

```
[!] No se pudo identificar el día en el nombre del archivo: <nombre del archivo>
Ingresa el día (1-2 dígitos) que corresponde a '<nombre del archivo>':
```

- Se debe escribir solo el número del día (por ejemplo `9` o `09`, ambos son válidos).
- Si se escribe algo que no es un número o tiene más de 2 dígitos, el programa insiste: `Valor inválido. Ingresa solo números (ej. 9 o 09).` y vuelve a preguntar hasta recibir una respuesta válida.
- Esto pasa una vez por cada archivo cuyo nombre no calce con el patrón, mostrando siempre el nombre exacto del archivo en cuestión.

## Los 3 reportes que se generan

Los 3 Excel se guardan siempre en la carpeta **`salidas`** (dentro de la carpeta del programa, se crea sola si no existe). El nombre de archivo incluye la fecha **del día en que se ejecuta el programa** (no la fecha de las transacciones).

### Archivo 1 — `SOPORTES REDEBAN VISA DDMMAAAA.xlsx` (opción 2)

- **Hoja "TIPO 50"**: transacciones de la red VISA/VISN — hoja combinada con todos los días seleccionados juntos.
- **Hoja(s) "TIPO 01 \<día\>"** (una por cada día procesado): transacciones de compra/reverso VISA/VISN válidas para soporte. Las transacciones que son **reversos** quedan con su valor en **negativo**, para que al sumar la hoja el reverso se reste automáticamente del total.

### Archivo 2 — `RECH EP DEPO DDMMAAAA.xlsx` (opción 3)

- **Hoja "TIPO 12-15"**: registros complementarios/de depósito, combinados de todos los días juntos.
- **Hoja(s) "RECH EP \<día\>"** (una por día): transacciones asociadas a rechazos/depósitos EP.

### Archivo 3 — `DDMMYY.xlsx` (opción 4, ej. `140826.xlsx`)

- **Hoja "Hoja1"**: la planilla de **VALOR CINTA** — muestra el valor cinta nacional y el valor cinta internacional calculados sobre el total combinado de todos los archivos/días seleccionados juntos (siempre el consolidado completo, nunca por día separado).
- **Hoja(s) "Detalle Dinamicas \<día\>"** (una por día): el desglose de cómo se llega a cada valor cinta, concepto por concepto, para poder verificar la cuenta sin necesidad de abrir el Excel original.

> El Archivo 3 siempre se calcula sobre **todo** el archivo plano, sin importar si ya se generaron o no los Archivos 1 y 2 — no depende de ellos.

## Errores comunes y qué hacer

El programa **nunca se cierra solo** por un error — siempre vuelve al menú. Todos los mensajes de error empiezan con `[!]`.

| Mensaje | Causa probable | Qué hacer |
|---|---|---|
| `[!] No se encontró el archivo: <ruta>` | El archivo plano fue movido, renombrado o borrado después de seleccionarlo | Volver a la opción 1 y elegirlo desde su ubicación actual |
| `[!] Sin permisos de acceso (¿el archivo está abierto en Excel?): ...` | El Excel de salida ya existe y **está abierto en Excel** en ese momento | Cerrar ese Excel y volver a generar el reporte |
| `[!] Error de datos o configuración: ...` | Algo en el contenido del archivo plano no es el esperado | Confirmar que el archivo seleccionado es el correcto; si persiste, avisar a soporte técnico |
| `[!] Ocurrió un error inesperado: ...` | Cualquier fallo no anticipado | Anotar el mensaje completo y reportarlo a soporte técnico; revisar la carpeta `logs` |
| `Opción no válida.` | Se escribió algo fuera de 1-5 | Escribir un número del 1 al 5 |
| `Primero selecciona uno o varios archivos (opción 1).` | Se intentó generar un reporte sin haber usado antes la opción 1 | Usar primero la opción 1 |
| `No se seleccionó ningún archivo.` | Se canceló el diálogo de selección | Repetir la opción 1 y elegir uno o más archivos |
| `El archivo <nombre> no tiene líneas para procesar. Se omite.` | El archivo plano está vacío o dañado | Confirmar que es el archivo correcto y que no esté vacío |
| `No se pudo construir la tabla para <archivo>. Se omite.` | Error al interpretar las líneas de ese archivo | Verificar el formato del archivo; contactar soporte si persiste |
| `Ningún archivo pudo procesarse.` | Todos los archivos seleccionados fallaron | Revisar los mensajes previos y corregir la selección |
| `'<Hoja>' en <archivo>.xlsx quedó sin filas — revisa el filtro aplicado.` | El reporte se generó bien, pero esa hoja/día no tuvo transacciones que cumplieran la condición | **No es un error crítico** — el Excel se genera igual, solo esa hoja queda sin filas. Confirma que no hubo transacciones de ese tipo ese día; si se esperaban, revisar que se seleccionó el archivo correcto |
| Pregunta por el día de un archivo | El nombre del archivo no sigue el patrón esperado | Escribir el día (1-2 dígitos) cuando se pida |

Cuando todo sale bien, se ve `[OK] Generado: <ruta del archivo>` y luego `[OK] Archivo 1/2/3 generado.`.

Cada ejecución queda registrada en un archivo de log diario (`logs/ejecucion_YYYYMMDD.log`) — normalmente no hace falta abrirlo, pero es útil para soporte técnico si hay que reportar un problema.

## Flujo paso a paso recomendado (consolidado de fin de semana)

1. Doble clic en `orquestador.bat`.
2. Elegir opción **1**. En el explorador de Windows, ubicar los archivos planos del banco y seleccionar con Ctrl+clic los de sábado, domingo y lunes (y martes si hubo festivo). Confirmar.
   - Si algún nombre no sigue el patrón esperado, responder en consola el día que corresponde cuando se pregunte.
3. Verificar en pantalla que aparecen listados todos los archivos esperados bajo "Archivos actuales".
4. Elegir opción **2** (Archivo 1) y esperar `[OK] Archivo 1 generado.`.
5. Elegir opción **3** (Archivo 2) y esperar `[OK] Archivo 2 generado.`.
6. Elegir opción **4** (Archivo 3) y esperar `[OK] Archivo 3 generado.`.
7. Elegir opción **5** para salir.
8. Abrir los 3 Excel generados en la carpeta `salidas` para revisión.
9. Si alguna hoja aparece marcada como "sin filas" durante el proceso, revisar si es esperado (no hubo transacciones de ese tipo ese día) o si falta seleccionar algún archivo.
