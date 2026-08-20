# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

A small interactive CLI tool that parses a fixed-width flat file (bank card
transaction records) and produces three formatted Excel reports. There is no
package manager manifest (no requirements.txt/pyproject.toml) — dependencies
are `pandas` and `XlsxWriter` (`pip install xlsxwriter`), installed into
whatever Python environment runs the script. `openpyxl` was the export
engine until it was replaced by XlsxWriter for performance (see
`exportador.py`'s module docstring) — it is no longer a dependency.

Two deeper reference docs live in `docs/` and are worth reading before
making non-trivial changes:
- **`docs/DOCUMENTACION_TECNICA.md`** — field-by-field LAYOUT table, the exact
  boolean conditions behind every sheet in all three reports, and the
  VALOR CINTA pivot-formula derivation (Archivo 3), including a noted
  discrepancy worth confirming with the business (Archivo 2 filters
  `TIPO-REGISTRO isin ["01","20"]` while Archivo 3's pivots use
  `isin ["01","21"]`).
- **`docs/MANUAL_USUARIO.md`** — end-user manual in Spanish (menu walkthrough,
  error messages table, recommended weekend-consolidation flow).

`main.py` and `orquestador.bat` (the entry points) live at the repo root;
the six business-logic modules (`config.py`, `lector.py`, `utils.py`,
`validador.py`, `reglas.py`, `exportador.py`) live in `src/`. `main.py`
adds `src/` to `sys.path` before importing, so the internal import
statements (`from config import ...`, etc.) are unaffected by the module
files' location.

## Running

```
python main.py
```

or on Windows, double-click / run `orquestador.bat`, which `cd`s into the
script's own directory and runs `python main.py`.

The program is a text menu (see `main.py`):
1. Select one or more flat files to process (opens a Tk file dialog with
   multi-select — Ctrl/Shift+click; falls back to a comma-separated console
   prompt if no GUI is available). Multi-select exists to consolidate a
   weekend close (Saturday+Sunday+Monday, +Tuesday if there's a holiday)
   into one run.
2. Generar Archivo 1 (SOPORTES REDEBAN VISA)
3. Generar Archivo 2 (RECH EP DEPO)
4. Generar Archivo 3 (VALORES CINTA)
5. Exit.

Options 2/3/4 are independent — the user can trigger them in any order, and
each one calls `_construir_dataframe_base()` in `main.py`, which **rereads
the selected flat file(s) from disk and rebuilds the DataFrame from scratch
every time** (no caching across menu options), so a report always reflects
the current file contents even if it was swapped out mid-session. This is
deliberate, not an oversight — see `main.py`'s module docstring.

There is no test suite, linter, or build step configured in this repo.

## Architecture

Data flows in one direction through five modules, each with a single
responsibility. All five (plus `utils.py`) live in `src/`:

```
src/config.py  →  src/lector.py  →  src/reglas.py  →  src/exportador.py
   (layout)          (parse)          (filter/pivot)      (write .xlsx)
```

- **`config.py`** — `LAYOUT`: the fixed-width field spec (name, 1-indexed
  start position and field length — `inicio`, `longitud`) describing the
  flat file format.
  This is the source of truth for every field name used elsewhere in the
  codebase. Also defines output paths (`ARCHIVO_1/2/3` in `salidas/`,
  timestamped by day) and encoding.
- **`lector.py`** — reads non-empty lines from the flat file and slices each
  line per `LAYOUT` into a dict, producing one flat `pandas.DataFrame`
  (`construir_dataframe`). Every field is kept as a stripped string; no type
  conversion happens here.
- **`reglas.py`** — all business logic. Three functions
  (`generar_archivo_1`, `generar_archivo_2`, `generar_archivo_3`) each take
  the *same* base DataFrame and return a `{"NombreHoja": contenido}` dict of
  sheets (`contenido` is normally a DataFrame; Archivo 3's `"Hoja1"` is a
  list of free-position cells, see `_escribir_hoja_libre` in
  `exportador.py`). Archivo 1 applies boolean-mask filters and produces
  `"TIPO 50"` (combined) + `"TIPO 01 <dia>"` (one sheet per day). Archivo 2
  is the same pattern for `"TIPO 12-15"` (combined) + `"RECH EP <dia>"`.
  Archivo 3's `"Hoja1"` computes the fixed VALOR CINTA formula (nacional +
  internacional) over the **entire combined universe** — it never depends
  on the filtered results of 1 or 2, and is never split by day — while
  `"Detalle Dinamicas <dia>"` gives the same breakdown per day. See "Multi-file
  selection" below for how day-splitting works.
- **`exportador.py`** — `guardar_con_formato(path, hojas)` writes a sheets
  dict to `.xlsx` using XlsxWriter (`constant_memory=True`). Each hoja can
  be a DataFrame (header styling, frozen header row, auto column width,
  and currency formatting inferred from column name substrings —
  `"monto"`/`"valor"` only; `"fecha"` columns get no number format, see
  "Key conventions" below) or a list of free-position cells
  (`_escribir_hoja_libre`, for fixed layouts that aren't tabular). Writes
  empty DataFrame sheets (headers only) rather than failing when a filter
  matches zero rows, and warns in console + log.
- **`validador.py`** — file picker (`seleccionar_archivos`, plural — Tk
  multi-select dialog with a comma-separated console fallback) and
  `extraer_dia(ruta)`, which parses the day out of the production filename
  pattern `GOF.GRB.FM14.FYYMMDD.txt` (regex `F(\d{2})(\d{2})(\d{2})`,
  keeping the day group). If a filename doesn't match, it interactively
  asks the user which day that specific file belongs to — it never guesses
  silently.
- **`utils.py`** — `manejar_errores` decorator and `logger`. Every
  I/O-facing function in the other modules is wrapped in `manejar_errores`:
  it catches `FileNotFoundError`, `PermissionError`, `ValueError`, and any
  other `Exception`, prints a Spanish user-facing message, logs to
  `logs/ejecucion_YYYYMMDD.log`, and returns `None` instead of raising — a
  single bad file/row must never crash the menu loop in `main.py`.
- **`main.py`** — the menu loop and `estado` dict (`{"rutas", "df"}`, note
  `rutas` is a list) holding session state. Orchestrates the pipeline
  above; contains no business logic itself.

## Multi-file selection (weekend consolidation)

`_construir_dataframe_base()` in `main.py` loops over every selected path,
calls `validador.extraer_dia` to get its day, builds that file's DataFrame
via the normal single-file `lector.py` functions, tags every row of that
partial DataFrame with a `"__dia"` column, then concatenates all partial
DataFrames into one combined `df` — the rest of the pipeline (`reglas.py`)
is unaware of file boundaries, only of the `"__dia"` column.

- Sheets that must be split per day (`"TIPO 01"`, `"RECH EP"`, `"Detalle
  Dinamicas"`) always carry the day suffix, **even when only one file is
  selected** — there is no separate "single file" naming mode, so behavior
  is consistent regardless of how many files were picked.
- Sheets that stay combined across all days (`"TIPO 50"`, `"TIPO 12-15"`,
  Archivo 3's `"Hoja1"`) drop the `"__dia"` column before being returned —
  it's an internal grouping key, never a real output column.
- `reglas.py`'s `_segregar_por_dia(df_base, hoja_filtrada, nombre_base)`
  is the shared helper for the day-split sheets; it iterates the days
  present in `df_base` (not just the days that survived the filter), so a
  day with zero matching rows still gets its own (empty, warned-about)
  sheet instead of silently disappearing.

## Key conventions to preserve when editing

- **`LAYOUT` stores `inicio` (1-indexed start position) and `longitud`
  (field length in characters)** — not an end position — matching what the
  bank's flat-file spec Excel actually exposes (`MID(texto, inicio,
  longitud)`). `lector.py` computes the 0-indexed slice internally as
  `linea[inicio-1 : inicio-1+longitud]` — don't adjust indices anywhere
  else to compensate.
- **All fields come out of `lector.py` as stripped strings.** Do type
  conversion (dates, numerics) explicitly inside `reglas.py`, per field —
  this is a deliberate choice to keep full control over real-world data
  quirks, not an oversight.
- **Archivo 3 always reads the unfiltered base DataFrame**, never the output
  of `generar_archivo_1`/`generar_archivo_2`. Keep the three generator
  functions independent.
- **Never let a function raise past the menu loop.** Wrap new I/O or
  user-input-facing functions in `@manejar_errores` rather than adding local
  try/except blocks, so failures log and return to the menu instead of
  crashing.
- All three generator functions in `reglas.py` are implemented against real
  business rules and real `LAYOUT` field names (`RED-LOGICA`,
  `TIPO-REGISTRO`, `FIID SPONSOR`, `COD-TIPO-TRANS`, `INDICADOR
  INTER/NACIONAL`, etc.) — there is no placeholder logic left.
- **`RED-LOGICA` (VISN/VISA) is only filtered in Archivo 1**, never in
  Archivo 2 or in Archivo 3's pivots — a real business asymmetry (it's in
  the report names themselves: "SOPORTES REDEBAN VISA" vs "RECH EP DEPO"),
  not a bug to "fix" for consistency.
- Field names with irregular spacing in `LAYOUT` are literal and must be
  quoted exactly as-is, never normalized: `" CODIGO-RESP"` (leading space),
  `"NUMEROS - CUOTAS "` (trailing space), `"COMISION -  FINANCIERA -
  ADQUIRIENTE"` (double internal space).
- **Any output sheet over 1,000,000 rows gets auto-partitioned** into
  `"{nombre} (1)"`, `"{nombre} (2)"`, etc. (`_particionar_hoja_grande` in
  `src/exportador.py`, applied inside `guardar_con_formato` to every
  DataFrame sheet from any of the 3 archivos) — Excel's real limit is
  1,048,576 rows/sheet. This is a generic export-layer safeguard, not a
  business rule — don't move it into `reglas.py`.
- **The export engine is XlsxWriter (`constant_memory=True`), not
  openpyxl** — chosen after a benchmark on synthetic data of the real
  shape (55 columns, ~1.75M rows) showed it 2.52x faster than openpyxl's
  `write_only` mode (348s vs 878s). `constant_memory` streams rows to
  disk and enforces writing each row's cells in strictly increasing
  column order — `_escribir_hoja_dataframe` exploits `indices_formato`
  (precomputed once per sheet, normally just `MONTO-1`) to write each row
  in `write_row()` chunks for the unformatted stretch and a single
  `write()` call for the formatted cell, instead of iterating all ~55
  columns per row or creating a wrapper object per formatted cell.
  `use_zip64()` is enabled on every workbook — without it XlsxWriter
  raises `FileSizeError` once the compressed `.xlsx` passes ~4GB, which is
  real at multi-million-row scale. `_anchos_columnas` similarly samples
  the first 2000 rows (`_FILAS_MUESTRA_ANCHO`) instead of scanning the
  full DataFrame.
- **`"fecha"`-named columns get no `number_format`** — confirmed dead
  weight: `lector.py`/`reglas.py` never convert those fields to a real
  date type, so a date `number_format` on a text cell has zero visual
  effect in Excel, but still cost real write time under the old openpyxl
  implementation. Only `"monto"`/`"valor"` columns get formatted
  (currency). Don't reintroduce date formatting without first converting
  the column to an actual date/datetime value.
- **`main.py` prints a `HH:MM:SS` timestamp when reading starts and again
  right before "Aplicando reglas de negocio..."** (the moment filtering
  begins), plus elapsed time (`time.perf_counter()`) at the end of parsing
  and of each `opcion_generar_archivo_N()` — visible in the console
  without opening the log, for tracking how long large consolidated runs
  take.
- Large `*.txt` files at the repo root (`GOF.GRB.FM14.F260813.txt`,
  `GOF.GRB.FM14.F260814.txt` — real production naming, days 13/14) are
  sample flat-file inputs for manual testing, ~130+ MB each — don't read
  them in full; slice or `head` a few lines instead. To exercise
  `extraer_dia`'s interactive fallback (filenames that don't match the
  `GOF.GRB.FM14.FYYMMDD.txt` pattern), copy one of these to an arbitrary
  name rather than relying on fixture files, since none are currently
  checked into the repo.
