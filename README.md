# Comparación SIE — Marca vs Mercado (misma molécula)

Comparativa de performance de las marcas de Siegfried Argentina contra su mercado (productos no-SIE con la **misma molécula**), por canal y por periodo.

## Entregables

- [`Comparativa_Marcas.html`](Comparativa_Marcas.html) — vista web responsive (desktop + mobile) con filtro por segmento y búsqueda por marca.
- [`Comparativa_Marcas_MAT_YTD.xlsx`](Comparativa_Marcas_MAT_YTD.xlsx) — Excel con segmentador (celda `B2`), fórmulas vivas (`SUMIFS` + `IFERROR`) y formato condicional.
- [`build_analisis.py`](build_analisis.py) — script que reconstruye ambos entregables a partir del extract IQVIA.

## Indicadores

Para tres bloques de tiempo (MAT, YTD y MES) se calcula:

- **U Marca Ant / Act**: unidades de Siegfried (marca propia).
- **U Mercado Ant / Act**: unidades de competidores con la misma molécula.
- **MS% Ant / Act**: `U Marca / (U Marca + U Mercado)`.
- **IE** (Índice de Evolución):

  ```
  IE = ((Marca Act / Marca Ant) / (Universo Mol Act / Universo Mol Ant)) × 100
  ```

  donde `Universo Mol = U Marca + U Mercado` (la misma molécula).
  - `IE > 100`: la marca crece más rápido que su mercado (gana share).
  - `IE < 100`: la marca crece más lento que su mercado (pierde share).
- **Var pp**: `MS% Act − MS% Ant` (puntos porcentuales).

## Bloques de periodo

| Bloque | Ant            | Act            | Significado                       |
|--------|----------------|----------------|-----------------------------------|
| MAT    | MAT Abr-2025   | MAT Abr-2026   | 12 meses móviles cerrados a abril |
| YTD    | Ene–Abr 2025   | Ene–Abr 2026   | acumulado del año hasta abril     |
| MES    | Abr 2025       | Abr 2026       | mes actual vs mismo mes año anterior |

## Segmentador

Tres opciones de canal (en `Comparativa!B2` o en los botones del HTML):

- **TOTAL** — Ético + Popular
- **ÉTICO** — sólo ético
- **POPULAR** — sólo OTC/popular

## Marcas en el análisis (49)

Familias en promoción 2026 + extras detectadas en la Grilla:

`ACEMUK, AIREAL, HEXALER CORT, ALIDIAL, ALIDIAL L, MOMETAX, BACTRIM, ACANTEX, DECADRON, MACROMAX, CEFALEXINA, CALCIO BASE, ISIS FREE, TRIP D3, DELTROX, ENTRESTO, DILATREND, ROXOLAN, ROXOLAN PLUS, SINTROM, EMPAX, SILTRAN, METGLUCON, TERLOC, DIOVAN, TELPRES, DAURAN, PIXABAN, LEVITAL, VALIUM, PGB, QTP, MADOPAR, ACNECLIN, CLOBESOL, MICOMAZOL, MICROSONA, PALDAR, ROACCUTAN, MAGNUS, TETRALGIN, LURAP, BREXIL, VALQUIR, SIDERBLUT, NEBILET, EXFORGE, ISIS NAT, GYNODERM`.

## Cómo reconstruir

1. Colocar el extract IQVIA (`AR_PM_FV_Standard_*.xlsx`, con la columna `Market (E/OTC)`) en la misma carpeta.
2. Instalar dependencias:
   ```bash
   python -m pip install pandas openpyxl
   ```
3. Ejecutar:
   ```bash
   python build_analisis.py
   ```

> El extract IQVIA fuente **no** se versiona (ver `.gitignore`).

## Estructura del Excel

| Hoja          | Contenido                                                                 |
|---------------|---------------------------------------------------------------------------|
| `Comparativa` | Vista principal con segmentador en `B2` y 3 bloques (MAT / YTD / MES).    |
| `Datos`       | Tabla larga: una fila por (Marca × Segmento) con unidades por periodo.    |
| `Notas`       | Definiciones de los indicadores.                                          |

La hoja `Comparativa` usa `SUMIFS(Datos!..., Marca, Segmento=$B$2)` para traer las unidades; al cambiar `B2` (TOTAL/ETICO/POPULAR) toda la tabla recalcula.
