"""
Comparativa Marca vs Mercado (misma molécula, no-SIE).
Bloques: MAT (Abr-2026 vs Abr-2025), YTD (Ene-Abr 2026 vs 2025), MES (Abr-2026 vs Abr-2025).
Segmentador: TOTAL (ético+popular) | ÉTICO | POPULAR.

IE (índice de evolución) = ((Marca Act / Marca Ant) / (Mercado Act / Mercado Ant)) × 100
   IE > 100: la marca crece más que el mercado (gana share)
   IE < 100: la marca crece menos que el mercado (pierde share)

Inputs : AR_PM_FV_Standard_May-22-2026 (5).xlsx
Outputs: Comparativa_Marcas_MAT_YTD.xlsx | Comparativa_Marcas.html (responsive)
"""

import os, json
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import CellIsRule
from openpyxl.worksheet.datavalidation import DataValidation

BASE = r"C:\Users\camarinaro\Downloads\Analisis fran"
SRC  = os.path.join(BASE, "AR_PM_FV_Standard_May-22-2026 (5).xlsx")
OUT_XLSX = os.path.join(BASE, "Comparativa_Marcas_MAT_YTD.xlsx")
OUT_HTML = os.path.join(BASE, "Comparativa_Marcas.html")

COL_MAT_ACT = "Units\nMAT M 2026 April"
COL_MAT_ANT = "Units\nMAT M 2025 April"
COL_YTD_ACT = "Units\nYTD Apr 2026"
COL_YTD_ANT = "Units\nYTD Apr 2025"
COL_MES_ACT = "Units\nApr 2026"
COL_MES_ANT = "Units\nApr 2025"
COL_TRIM_ACT = "Units\nFeb  to  Apr 2026"      # Feb-Apr 2026
COL_TRIM_ANT = "Units\nFeb  to  Apr 2025"      # Feb-Apr 2025
COL_SEM_ACT  = "Units\nNov 2025 to  Apr 2026"  # Nov 2025-Apr 2026
COL_SEM_ANT  = "Units\nNov 2024 to  Apr 2025"  # Nov 2024-Apr 2025
COL_SEG     = "Market (E/OTC)\n"

BRANDS = {
    "ACEMUK":        ["ACEMUK (SIE)", "ACEMUK VL (SIE)", "ACEMUK BIOTIC DUO (SIE)",
                      "ACEMUK DIA Y NOCHE (SIE)", "ACEMUK L (SIE)", "ACEMUK GRIP (SIE)"],
    "AIREAL":        ["AIREAL (SIE)", "AIREAL PLUS (SIE)"],
    "HEXALER CORT":  ["HEXALER CORT (SIE)"],
    "ALIDIAL":       ["ALIDIAL (SIE)"],
    "ALIDIAL L":     ["ALIDIAL L (SIE)"],
    "MOMETAX":       ["MOMETAX (SIE)"],
    "BACTRIM":       ["BACTRIM (SIE)", "BACTRIM BALSAMICO (SIE)",
                      "BACTRIM FORTE (SIE)", "BACTRIM JARABE (SIE)"],
    "ACANTEX":       ["ACANTEX (SIE)"],
    "DECADRON":      ["DECADRON (SIE)", "DECADRON SHOCK (SIE)",
                      "DECADRON. (SIE)", "DUO-DECADRON (SIE)"],
    "MACROMAX":      ["MACROMAX (SIE)", "MACROMAX PEDIATR (SIE)"],
    "CEFALEXINA":    ["CEFALEXINA ARGENTI (SIE)"],
    "CALCIO BASE":   ["CALCIO BASE (SIE)", "CALCIO BASE D (SIE)",
                      "CALCIO BASE D3 (SIE)", "CALCIO BASE EFER (SIE)",
                      "CALCIO CIT DUPO D3 (SIE)"],
    "ISIS FREE":     ["ISIS FREE S/ESTROG (SIE)"],
    "TRIP D3":       ["TRIP D3 (SIE)", "TRIP D3 PLUS (SIE)"],
    "DELTROX":       ["DELTROX NF (SIE)"],
    "ENTRESTO":      ["ENTRESTO (SIE)"],
    "DILATREND":     ["DILATREND (SIE)", "DILATREND AP (SIE)", "DILATREND D (SIE)"],
    "ROXOLAN":       ["ROXOLAN (SIE)"],
    "ROXOLAN PLUS":  ["ROXOLAN PLUS (SIE)"],
    "SINTROM":       ["SINTROM (SIE)"],
    "EMPAX":         ["EMPAX (SIE)", "EMPAX MET (SIE)"],
    "SILTRAN":       ["SILTRAN (SIE)", "SILTRAN MET (SIE)"],
    "METGLUCON":     ["METGLUCON (SIE)", "METGLUCON AP (SIE)", "METGLUCON DUO (SIE)"],
    "TERLOC":        ["TERLOC (SIE)", "TERLOC DUO (SIE)"],
    "DIOVAN":        ["DIOVAN (SIE)", "DIOVAN-D (SIE)", "DIOVAN IC (SIE)"],
    "TELPRES":       ["TELPRES (SIE)"],
    "DAURAN":        ["DAURAN (SIE)"],
    "PIXABAN":       ["PIXABAN (SIE)"],
    "LEVITAL":       ["LEVITAL (SIE)"],
    "VALIUM":        ["VALIUM (SIE)"],
    "PGB":           ["PGB (SIE)"],
    "QTP":           ["QTP (SIE)"],
    "MADOPAR":       ["MADOPAR (SIE)", "MADOPAR HBS (SIE)"],
    "ACNECLIN":      ["ACNECLIN (SIE)", "ACNECLIN PBA (SIE)"],
    "CLOBESOL":      ["CLOBESOL (SIE)"],
    "MICOMAZOL":     ["MICOMAZOL (SIE)", "MICOMAZOL B (SIE)"],
    "MICROSONA":     ["MICROSONA (SIE)", "MICROSONA BB (SIE)", "MICROSONA C (SIE)"],
    "PALDAR":        ["PALDAR (SIE)", "PALDAR H (SIE)"],
    "ROACCUTAN":     ["ROACCUTAN (SIE)"],
    "MAGNUS":        ["MAGNUS (SIE)", "MAGNUS 36 (SIE)"],
    "TETRALGIN":     ["TETRALGIN (SIE)", "TETRALGIN APC (SIE)", "TETRALGIN NOVO (SIE)"],
    # Extras detectadas en la Grilla 2026
    "LURAP":         ["LURAP (SIE)"],
    "BREXIL":        ["BREXIL (SIE)"],
    "VALQUIR":       ["VALQUIR (SIE)"],
    "SIDERBLUT":     ["SIDERBLUT (SIE)", "SIDERBLUT FOLICO (SIE)", "SIDERBLUT COMPLEX (SIE)",
                      "SIDERBLUT GOTAS (SIE)", "SIDERBLUT IM (SIE)", "SIDERBLUT POLI (SIE)"],
    "NEBILET":       ["NEBILET (SIE)", "NEBILET D (SIE)"],
    "EXFORGE":       ["EXFORGE (SIE)", "EXFORGE D (SIE)"],
    "ISIS NAT":      ["ISIS NAT (SIE)"],
    "GYNODERM":      ["GYNODERM (SIE)", "GYNODERM CARE GEL (SIE)", "GYNODERM HYALU (SIE)"],
}

SEGMENTS = ["TOTAL", "ETICO", "POPULAR"]

# ---- LOAD ------------------------------------------------------------------
print("Cargando datos...")
df = pd.read_excel(SRC, sheet_name="DS_AR_PM_FV_Standard")
df["Molecules Long\n"] = df["Molecules Long\n"].astype(str).str.strip()
df["Product\n"]        = df["Product\n"].astype(str).str.strip()
df["Manufacturer\n"]   = df["Manufacturer\n"].astype(str).str.strip()
df[COL_SEG]            = df[COL_SEG].astype(str).str.strip().str.upper()
for c in [COL_MAT_ACT, COL_MAT_ANT, COL_YTD_ACT, COL_YTD_ANT,
          COL_MES_ACT, COL_MES_ANT, COL_TRIM_ACT, COL_TRIM_ANT,
          COL_SEM_ACT, COL_SEM_ANT]:
    df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)

def aggregate(bdf, mdf, seg):
    bsel = bdf if seg=="TOTAL" else bdf[bdf[COL_SEG]==seg]
    msel = mdf if seg=="TOTAL" else mdf[mdf[COL_SEG]==seg]
    return {
        "BMA": float(bsel[COL_MAT_ANT].sum()), "BMC": float(bsel[COL_MAT_ACT].sum()),
        "MMA": float(msel[COL_MAT_ANT].sum()), "MMC": float(msel[COL_MAT_ACT].sum()),
        "BYA": float(bsel[COL_YTD_ANT].sum()), "BYC": float(bsel[COL_YTD_ACT].sum()),
        "MYA": float(msel[COL_YTD_ANT].sum()), "MYC": float(msel[COL_YTD_ACT].sum()),
        "BPA": float(bsel[COL_MES_ANT].sum()), "BPC": float(bsel[COL_MES_ACT].sum()),
        "MPA": float(msel[COL_MES_ANT].sum()), "MPC": float(msel[COL_MES_ACT].sum()),
        "BTA": float(bsel[COL_TRIM_ANT].sum()), "BTC": float(bsel[COL_TRIM_ACT].sum()),
        "MTA": float(msel[COL_TRIM_ANT].sum()), "MTC": float(msel[COL_TRIM_ACT].sum()),
        "BSA": float(bsel[COL_SEM_ANT].sum()), "BSC": float(bsel[COL_SEM_ACT].sum()),
        "MSAS": float(msel[COL_SEM_ANT].sum()), "MSCS": float(msel[COL_SEM_ACT].sum()),
    }

rows = []
for brand, prods in BRANDS.items():
    sub = df[(df["Manufacturer\n"]=="SIEGFRIED") & (df["Product\n"].isin(prods))]
    molecules = sorted(set(sub["Molecules Long\n"].unique()))
    mkt = df[(df["Manufacturer\n"]!="SIEGFRIED") & (df["Molecules Long\n"].isin(molecules))]
    for seg in SEGMENTS:
        agg = aggregate(sub, mkt, seg)
        rows.append({"Marca": brand, "Segmento": seg,
                     "Moleculas": " | ".join(molecules) if molecules else "",
                     **agg})
datos = pd.DataFrame(rows)
brand_order = list(BRANDS.keys())

# ---- EXCEL -----------------------------------------------------------------
print("Escribiendo Excel...")
wb = Workbook()

# Hoja Datos
ws_d = wb.active
ws_d.title = "Datos"
DATOS_HDR = ["Marca","Segmento","Moleculas",
             "BMA","BMC","MMA","MMC",
             "BYA","BYC","MYA","MYC",
             "BPA","BPC","MPA","MPC"]
for j,h in enumerate(DATOS_HDR, start=1):
    c = ws_d.cell(row=1, column=j, value=h)
    c.font = Font(bold=True, color="FFFFFF")
    c.fill = PatternFill(start_color="305496", end_color="305496", fill_type="solid")
DATOS_ROW0 = 2
for i,r in enumerate(datos.itertuples(index=False), start=DATOS_ROW0):
    ws_d.cell(row=i, column=1, value=r.Marca)
    ws_d.cell(row=i, column=2, value=r.Segmento)
    ws_d.cell(row=i, column=3, value=r.Moleculas)
    for k,col in enumerate(["BMA","BMC","MMA","MMC","BYA","BYC","MYA","MYC","BPA","BPC","MPA","MPC"], start=4):
        ws_d.cell(row=i, column=k, value=getattr(r, col))
DATOS_ROWN = DATOS_ROW0 + len(datos) - 1
for col,w in zip("ABCDEFGHIJKLMNO",[16,12,55]+[12]*12):
    ws_d.column_dimensions[col].width = w
ws_d.freeze_panes = "A2"

# Hoja Comparativa
ws = wb.create_sheet("Comparativa", index=0)

HEADERS = [
    "Marca","Moleculas",
    "U Marca MAT Ant","U Marca MAT Act","U Mercado MAT Ant","U Mercado MAT Act",
    "MS% MAT Ant","MS% MAT Act","IE MAT","Var pp MAT",
    "U Marca YTD Ant","U Marca YTD Act","U Mercado YTD Ant","U Mercado YTD Act",
    "MS% YTD Ant","MS% YTD Act","IE YTD","Var pp YTD",
    "U Marca MES Ant","U Marca MES Act","U Mercado MES Ant","U Mercado MES Act",
    "MS% MES Ant","MS% MES Act","IE MES","Var pp MES",
]

THIN  = Side(border_style="thin", color="A6A6A6")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HFONT = Font(bold=True, color="FFFFFF", size=10)
HFILL = PatternFill(start_color="305496", end_color="305496", fill_type="solid")
MAT_BANNER = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
YTD_BANNER = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
MES_BANNER = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)

# Título
ws.cell(row=1, column=1, value="Comparativa Marca vs Mercado (misma molécula, no-SIE)").font = Font(bold=True, size=12)
ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(HEADERS))

# Segmentador
ws.cell(row=2, column=1, value="Segmento:").font = Font(bold=True, size=11)
ws.cell(row=2, column=1).alignment = Alignment(horizontal="right")
seg_cell = ws.cell(row=2, column=2, value="TOTAL")
seg_cell.font = Font(bold=True, size=12, color="FFFFFF")
seg_cell.fill = PatternFill(start_color="C65911", end_color="C65911", fill_type="solid")
seg_cell.alignment = CENTER
seg_cell.border = BORDER
dv = DataValidation(type="list", formula1='"TOTAL,ETICO,POPULAR"', allow_blank=False)
dv.prompt = "Elegí TOTAL, ETICO o POPULAR"
dv.promptTitle = "Segmento"
ws.add_data_validation(dv)
dv.add("B2")
ws.cell(row=2, column=4, value="(B2: TOTAL = Ético+Popular | ETICO sólo | POPULAR sólo)").font = Font(italic=True, color="666666", size=9)
ws.merge_cells(start_row=2, start_column=4, end_row=2, end_column=len(HEADERS))

# Banners (3 bloques)
ws.cell(row=3, column=3, value="BLOQUE MAT (Apr 2026 vs Apr 2025)").font = Font(bold=True)
ws.merge_cells(start_row=3, start_column=3, end_row=3, end_column=10)
ws.cell(row=3, column=3).fill = MAT_BANNER; ws.cell(row=3, column=3).alignment = CENTER

ws.cell(row=3, column=11, value="BLOQUE YTD (Ene-Abr 2026 vs 2025)").font = Font(bold=True)
ws.merge_cells(start_row=3, start_column=11, end_row=3, end_column=18)
ws.cell(row=3, column=11).fill = YTD_BANNER; ws.cell(row=3, column=11).alignment = CENTER

ws.cell(row=3, column=19, value="BLOQUE MES (Abr 2026 vs Abr 2025)").font = Font(bold=True)
ws.merge_cells(start_row=3, start_column=19, end_row=3, end_column=26)
ws.cell(row=3, column=19).fill = MES_BANNER; ws.cell(row=3, column=19).alignment = CENTER

# Header row
HEADER_ROW = 4
for i,h in enumerate(HEADERS, start=1):
    c = ws.cell(row=HEADER_ROW, column=i, value=h)
    c.font = HFONT; c.fill = HFILL; c.alignment = CENTER; c.border = BORDER

def rng(col): return f"Datos!${col}${DATOS_ROW0}:${col}${DATOS_ROWN}"
DR_MARCA   = rng("A")
DR_SEG     = rng("B")
# columns D..O en Datos
COLMAP = {"BMA":"D","BMC":"E","MMA":"F","MMC":"G",
          "BYA":"H","BYC":"I","MYA":"J","MYC":"K",
          "BPA":"L","BPC":"M","MPA":"N","MPC":"O"}

START_ROW = HEADER_ROW + 1
for i, brand in enumerate(brand_order, start=START_ROW):
    mol = datos[(datos.Marca==brand) & (datos.Segmento=="TOTAL")]["Moleculas"].iloc[0]
    ws.cell(row=i, column=1, value=brand).font = Font(bold=True)
    ws.cell(row=i, column=2, value=mol)

    def sumifs_for(key, row=i):
        col = COLMAP[key]
        return f'=SUMIFS({rng(col)},{DR_MARCA},$A{row},{DR_SEG},$B$2)'

    # MAT block
    ws.cell(row=i, column=3, value=sumifs_for("BMA"))
    ws.cell(row=i, column=4, value=sumifs_for("BMC"))
    ws.cell(row=i, column=5, value=sumifs_for("MMA"))
    ws.cell(row=i, column=6, value=sumifs_for("MMC"))
    # Totales mol
    Tant_mat = f"(C{i}+E{i})"; Tact_mat = f"(D{i}+F{i})"
    ws.cell(row=i, column=7, value=f"=IFERROR(C{i}/{Tant_mat},0)")        # MS% Ant
    ws.cell(row=i, column=8, value=f"=IFERROR(D{i}/{Tact_mat},0)")        # MS% Act
    # IE = (BC/BA) / (TAC/TAN) * 100   (si BA=0 o TAN=0 → 0)
    ws.cell(row=i, column=9, value=f"=IFERROR((D{i}/C{i})/({Tact_mat}/{Tant_mat})*100,0)")
    ws.cell(row=i, column=10, value=f"=(H{i}-G{i})*100")                  # Var pp

    # YTD block
    ws.cell(row=i, column=11, value=sumifs_for("BYA"))
    ws.cell(row=i, column=12, value=sumifs_for("BYC"))
    ws.cell(row=i, column=13, value=sumifs_for("MYA"))
    ws.cell(row=i, column=14, value=sumifs_for("MYC"))
    Tant_ytd = f"(K{i}+M{i})"; Tact_ytd = f"(L{i}+N{i})"
    ws.cell(row=i, column=15, value=f"=IFERROR(K{i}/{Tant_ytd},0)")
    ws.cell(row=i, column=16, value=f"=IFERROR(L{i}/{Tact_ytd},0)")
    ws.cell(row=i, column=17, value=f"=IFERROR((L{i}/K{i})/({Tact_ytd}/{Tant_ytd})*100,0)")
    ws.cell(row=i, column=18, value=f"=(P{i}-O{i})*100")

    # MES block
    ws.cell(row=i, column=19, value=sumifs_for("BPA"))
    ws.cell(row=i, column=20, value=sumifs_for("BPC"))
    ws.cell(row=i, column=21, value=sumifs_for("MPA"))
    ws.cell(row=i, column=22, value=sumifs_for("MPC"))
    Tant_mes = f"(S{i}+U{i})"; Tact_mes = f"(T{i}+V{i})"
    ws.cell(row=i, column=23, value=f"=IFERROR(S{i}/{Tant_mes},0)")
    ws.cell(row=i, column=24, value=f"=IFERROR(T{i}/{Tact_mes},0)")
    ws.cell(row=i, column=25, value=f"=IFERROR((T{i}/S{i})/({Tact_mes}/{Tant_mes})*100,0)")
    ws.cell(row=i, column=26, value=f"=(X{i}-W{i})*100")

    for col in range(1, len(HEADERS)+1):
        cc = ws.cell(row=i, column=col); cc.border = BORDER
        if col in (3,4,5,6,11,12,13,14,19,20,21,22):
            cc.number_format = "#,##0"
        elif col in (7,8,15,16,23,24):
            cc.number_format = "0.0%"
        elif col in (9,17,25):
            cc.number_format = "0.0"
        elif col in (10,18,26):
            cc.number_format = "+0.00 \"pp\";-0.00 \"pp\";0.00 \"pp\""

widths = [16, 60,
          11,11,12,12, 9,9,8,12,
          11,11,12,12, 9,9,8,12,
          11,11,12,12, 9,9,8,12]
for i,w in enumerate(widths, start=1):
    ws.column_dimensions[get_column_letter(i)].width = w
ws.freeze_panes = "C5"

last_row = START_ROW + len(brand_order) - 1

# CF: Var pp + IE
green = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
red   = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
for col_letter in ("J","R","Z"):
    rng2 = f"{col_letter}{START_ROW}:{col_letter}{last_row}"
    ws.conditional_formatting.add(rng2, CellIsRule(operator="greaterThan", formula=["0"], fill=green))
    ws.conditional_formatting.add(rng2, CellIsRule(operator="lessThan",    formula=["0"], fill=red))
for col_letter in ("I","Q","Y"):
    rng2 = f"{col_letter}{START_ROW}:{col_letter}{last_row}"
    ws.conditional_formatting.add(rng2, CellIsRule(operator="greaterThan", formula=["100"], fill=green))
    ws.conditional_formatting.add(rng2, CellIsRule(operator="lessThan",    formula=["100"], fill=red))

# Total Cartera
total_row = last_row + 2
ws.cell(row=total_row, column=1, value="TOTAL CARTERA").font = Font(bold=True)
sum_cols = [3,4,5,6,11,12,13,14,19,20,21,22]
for col in sum_cols:
    L = get_column_letter(col)
    ws.cell(row=total_row, column=col, value=f"=SUM({L}{START_ROW}:{L}{last_row})").number_format = "#,##0"
# Recomputar derivados en la fila TOTAL
i = total_row
Tant_mat = f"(C{i}+E{i})"; Tact_mat = f"(D{i}+F{i})"
ws.cell(row=i, column=7, value=f"=IFERROR(C{i}/{Tant_mat},0)").number_format = "0.0%"
ws.cell(row=i, column=8, value=f"=IFERROR(D{i}/{Tact_mat},0)").number_format = "0.0%"
ws.cell(row=i, column=9, value=f"=IFERROR((D{i}/C{i})/({Tact_mat}/{Tant_mat})*100,0)").number_format = "0.0"
ws.cell(row=i, column=10, value=f"=(H{i}-G{i})*100").number_format = "+0.00 \"pp\";-0.00 \"pp\";0.00 \"pp\""
Tant_ytd = f"(K{i}+M{i})"; Tact_ytd = f"(L{i}+N{i})"
ws.cell(row=i, column=15, value=f"=IFERROR(K{i}/{Tant_ytd},0)").number_format = "0.0%"
ws.cell(row=i, column=16, value=f"=IFERROR(L{i}/{Tact_ytd},0)").number_format = "0.0%"
ws.cell(row=i, column=17, value=f"=IFERROR((L{i}/K{i})/({Tact_ytd}/{Tant_ytd})*100,0)").number_format = "0.0"
ws.cell(row=i, column=18, value=f"=(P{i}-O{i})*100").number_format = "+0.00 \"pp\";-0.00 \"pp\";0.00 \"pp\""
Tant_mes = f"(S{i}+U{i})"; Tact_mes = f"(T{i}+V{i})"
ws.cell(row=i, column=23, value=f"=IFERROR(S{i}/{Tant_mes},0)").number_format = "0.0%"
ws.cell(row=i, column=24, value=f"=IFERROR(T{i}/{Tact_mes},0)").number_format = "0.0%"
ws.cell(row=i, column=25, value=f"=IFERROR((T{i}/S{i})/({Tact_mes}/{Tant_mes})*100,0)").number_format = "0.0"
ws.cell(row=i, column=26, value=f"=(X{i}-W{i})*100").number_format = "+0.00 \"pp\";-0.00 \"pp\";0.00 \"pp\""
for col in range(1, len(HEADERS)+1):
    cc = ws.cell(row=total_row, column=col)
    cc.fill = PatternFill(start_color="FFE699", end_color="FFE699", fill_type="solid")
    cc.font = Font(bold=True); cc.border = BORDER

# Notas
ws2 = wb.create_sheet("Notas")
for row in [
    ["Definiciones"],
    [""],
    ["Segmento (B2)",  "TOTAL = Ético + Popular | ETICO | POPULAR"],
    ["MAT",            "Apr 2026 vs Apr 2025 (MAT mensual)"],
    ["YTD",            "Ene-Abr 2026 vs Ene-Abr 2025"],
    ["MES",            "Abr 2026 vs Abr 2025 (mes actual vs mismo mes año anterior)"],
    ["Mercado",        "Productos con la misma molécula de la marca, de manufacturers ≠ SIEGFRIED, filtrados por segmento"],
    ["MS%",            "Unidades de la marca / (marca + mercado)"],
    ["IE",             "((Marca Act / Marca Ant) / (Mercado total Act / Mercado total Ant)) × 100. >100 gana share; <100 pierde share"],
    ["Var pp",         "MS% Act − MS% Ant"],
]:
    ws2.append(row)
ws2.column_dimensions["A"].width = 20
ws2.column_dimensions["B"].width = 110
ws2.cell(row=1, column=1).font = Font(bold=True, size=12)

wb.save(OUT_XLSX)
print(f"Excel guardado: {OUT_XLSX}")

# ---- HTML responsive -------------------------------------------------------
print("Escribiendo HTML...")
js_data = {}
for seg in SEGMENTS:
    sub = datos[datos.Segmento==seg].set_index("Marca")
    js_data[seg] = {}
    for brand in brand_order:
        r = sub.loc[brand]
        js_data[seg][brand] = {
            "ba": r.BMA, "bc": r.BMC, "ma": r.MMA, "mc": r.MMC,
            "ya": r.BYA, "yc": r.BYC, "ymka": r.MYA, "ymkc": r.MYC,
            "pa": r.BPA, "pc": r.BPC, "mpa": r.MPA, "mpc": r.MPC,
            "ta": r.BTA, "tc": r.BTC, "mta": r.MTA, "mtc": r.MTC,
            "sa": r.BSA, "sc": r.BSC, "msa6": r.MSAS, "msc6": r.MSCS,
        }
js_brands = json.dumps(brand_order, ensure_ascii=False)
js_payload = json.dumps(js_data, ensure_ascii=False)

html = """<!doctype html>
<html lang="es">
<head>
<meta charset="utf-8"/>
<meta name="viewport" content="width=device-width, initial-scale=1, maximum-scale=2"/>
<title>Comparativa Marca vs Mercado</title>
<style>
  :root {
    --mat: #4472C4; --ytd: #C65911; --mes: #548235;
    --trim: #7030A0; --sem: #2E75B6;
    --bg: #f4f6fb; --card: #ffffff; --txt: #222; --muted: #6b7280;
    --line: #d0d0d0; --hdr: #305496;
  }
  * { box-sizing: border-box; }
  body { font-family: 'Segoe UI', Roboto, system-ui, Arial, sans-serif;
         margin: 0; padding: 16px; color: var(--txt); background: var(--bg);
         -webkit-text-size-adjust: 100%; }
  h1 { font-size: 18px; margin: 0 0 4px 0; }
  .sub { color: var(--muted); font-size: 12px; margin-bottom: 12px; }

  .controls {
    display: flex; flex-wrap: wrap; gap: 12px; align-items: center;
    background: var(--card); padding: 10px 12px; border-radius: 10px;
    box-shadow: 0 1px 2px rgba(0,0,0,.05); margin-bottom: 12px;
    position: sticky; top: 0; z-index: 10;
  }
  .controls label { font-weight: 600; font-size: 13px; }
  .seg-btns { display: flex; flex-wrap: wrap; }
  .seg-btns button {
    padding: 8px 14px; border: 1px solid #888; background: #fff; cursor: pointer;
    font-weight: 600; font-size: 13px; min-height: 38px;
  }
  .seg-btns button.active { background: var(--hdr); color: #fff; border-color: var(--hdr); }
  .seg-btns button:first-child { border-radius: 6px 0 0 6px; }
  .seg-btns button:last-child  { border-radius: 0 6px 6px 0; }
  .blk-toggle { display: inline-flex; gap: 4px; }
  .blk-toggle label { font-size: 12px; padding: 6px 8px; border: 1px solid #ccc; border-radius: 4px;
                       cursor: pointer; user-select: none; background: #fff; }
  .blk-toggle input { vertical-align: middle; margin-right: 4px; }
  .search { flex: 1 1 160px; min-width: 140px; padding: 8px 10px; border-radius: 6px;
            border: 1px solid #ccc; font-size: 13px; }

  /* ----- TABLA (siempre, también en mobile) ----- */
  .tbl-wrap { background: var(--card); border-radius: 10px; overflow: auto;
              box-shadow: 0 1px 3px rgba(0,0,0,.06); max-height: 78vh;
              -webkit-overflow-scrolling: touch; }
  table { border-collapse: separate; border-spacing: 0; width: 100%; font-size: 12px; min-width: 1800px; }
  th, td { border: 1px solid var(--line); padding: 6px 8px; }
  th { background: var(--hdr); color: #fff; text-align: center; font-weight: 600; }
  /* Header sticky en 2 niveles (la 2da fila pegada debajo de la 1ra) */
  thead tr:nth-child(1) th { position: sticky; top: 0; z-index: 3; }
  thead tr:nth-child(2) th { position: sticky; top: 30px; z-index: 3; }
  /* Marca sticky en horizontal */
  thead tr:nth-child(1) th:first-child { left: 0; z-index: 5; background: var(--hdr); }
  td.brand { font-weight: 700; background: #f7f9fc; white-space: nowrap;
             position: sticky; left: 0; z-index: 2;
             border-right: 2px solid #c0c0c0; }
  th.mat { background: var(--mat); } th.ytd { background: var(--ytd); }
  th.mes { background: var(--mes); } th.trim { background: var(--trim); }
  th.sem { background: var(--sem); }
  /* Sort */
  th[data-sort] { cursor: pointer; user-select: none; }
  th[data-sort]:hover { filter: brightness(1.12); }
  th[data-sort]::after { content: ' \\2195'; opacity: 0.45; font-size: 10px; }
  th[data-sort].asc::after  { content: ' \\25B2'; opacity: 1; }
  th[data-sort].desc::after { content: ' \\25BC'; opacity: 1; }
  td.num { text-align: right; font-variant-numeric: tabular-nums; white-space: nowrap; }
  tr:nth-child(even) td:not(.brand) { background: #f9fbfd; }
  td.good { background: #e6f4ea !important; color: #0a6e0a; }
  td.bad  { background: #fdecec !important; color: #9c0006; }
  td.neutral { color: #555; }
  tr.total td { background: #FFE699 !important; font-weight: 800; }
  tr.total td.brand { background: #FFE699 !important; }

  /* En mobile achicamos un poco padding/fuente pero seguimos tabla */
  @media (max-width: 760px) {
    body { padding: 8px; }
    h1 { font-size: 16px; }
    .tbl-wrap { max-height: 70vh; border-radius: 8px; }
    table { font-size: 11px; min-width: 1800px; }
    th, td { padding: 5px 6px; }
    thead tr:nth-child(2) th { top: 28px; }
    .controls { padding: 8px; gap: 8px; }
    .seg-btns button { padding: 6px 10px; font-size: 12px; min-height: 34px; }
    .blk-toggle label { font-size: 11px; padding: 5px 6px; }
    .search { font-size: 12px; padding: 6px 8px; }
    .btn-export { font-size: 12px; padding: 6px 10px; min-height: 34px; }
  }

  /* Botones export */
  .btn-export { padding: 8px 12px; border: 1px solid #888; background: #fff; cursor: pointer;
                font-weight: 600; font-size: 13px; border-radius: 6px; min-height: 38px; }
  .btn-export:hover { background: #f0f0f0; }
  .btn-export[disabled] { opacity: .6; cursor: progress; }
  /* Modo export: desactiva sticky + scroll para render limpio */
  .exporting .tbl-wrap { overflow: visible !important; max-height: none !important; }
  .exporting thead tr:nth-child(1) th,
  .exporting thead tr:nth-child(2) th,
  .exporting td.brand { position: static !important; }
  .exporting .controls { display: none !important; }
  .exporting #exportArea { display: block !important; width: max-content; min-width: 1800px; background: #fff; padding: 8px; }
  .exporting .summary { margin-bottom: 14px; }
  .exporting .sum-card { box-shadow: none; border: 1px solid #d0d0d0; }

  .hidden { display: none !important; }

  /* Resumen */
  .summary {
    display: grid; grid-template-columns: repeat(auto-fit, minmax(240px, 1fr));
    gap: 10px; margin-bottom: 12px;
  }
  .sum-card {
    background: var(--card); border-radius: 10px; padding: 10px 12px;
    box-shadow: 0 1px 2px rgba(0,0,0,.06); border-left: 4px solid #888;
  }
  .sum-card.sum-m { border-left-color: var(--mat); }
  .sum-card.sum-y { border-left-color: var(--ytd); }
  .sum-card.sum-p { border-left-color: var(--mes); }
  .sum-card.sum-t { border-left-color: var(--trim); }
  .sum-card.sum-s { border-left-color: var(--sem); }
  .sum-h { display: flex; justify-content: space-between; align-items: baseline; font-weight: 700; font-size: 13px; margin-bottom: 4px; }
  .sum-desc { font-size: 11px; color: var(--muted); font-weight: 500; }
  .sum-line { display: flex; flex-wrap: wrap; gap: 6px; margin-bottom: 4px; }
  .chip { display: inline-flex; align-items: center; gap: 4px; padding: 3px 8px; border-radius: 12px; font-size: 12px; }
  .chip b { font-size: 13px; }
  .chip-pos { background: #e6f4ea; color: #0a6e0a; }
  .chip-neg { background: #fdecec; color: #9c0006; }
  .chip-flat{ background: #eee; color: #444; }
  .sum-sub { font-size: 11px; color: var(--muted); }
  .sum-sub b { color: var(--txt); }
  @media (max-width: 760px) {
    .summary { grid-template-columns: 1fr 1fr; gap: 6px; }
    .sum-card { padding: 8px; }
    .sum-h { font-size: 12px; }
    .chip { font-size: 11px; padding: 2px 6px; }
  }
</style>
</head>
<body>
  <h1>Comparativa Marca vs Mercado (misma molécula, no-SIE)</h1>
  <div class="sub">Fuente: AR_PM_FV_Standard May-22-2026</div>

  <div class="controls">
    <label>Segmento:</label>
    <div class="seg-btns" id="segBtns">
      <button data-seg="TOTAL" class="active">TOTAL</button>
      <button data-seg="ETICO">ÉTICO</button>
      <button data-seg="POPULAR">POPULAR</button>
    </div>
    <span class="blk-toggle">
      <label><input type="checkbox" data-blk="mat" checked> MAT</label>
      <label><input type="checkbox" data-blk="ytd" checked> YTD</label>
      <label><input type="checkbox" data-blk="mes" checked> MES</label>
      <label><input type="checkbox" data-blk="trim" checked> TRIM</label>
      <label><input type="checkbox" data-blk="sem" checked> SEM</label>
    </span>
    <input class="search" id="q" placeholder="Buscar marca..."/>
    <button class="btn-export" id="btnPdf" type="button">Exportar PDF</button>
    <button class="btn-export" id="btnPng" type="button">Exportar imagen</button>
  </div>

  <div id="exportArea">
  <div id="summary" class="summary"></div>

  <div class="tbl-wrap">
    <table id="tbl">
      <thead>
        <tr>
          <th rowspan="2" data-sort="brand">Marca</th>
          <th class="mat blk-mat" colspan="6">MAT (Apr 2026 vs Apr 2025)</th>
          <th class="ytd blk-ytd" colspan="6">YTD (Ene-Abr 26 vs 25)</th>
          <th class="mes blk-mes" colspan="6">MES (Apr 2026 vs Apr 2025)</th>
          <th class="trim blk-trim" colspan="6">TRIM (Feb-Abr 26 vs 25)</th>
          <th class="sem blk-sem" colspan="6">SEM (Nov 25-Abr 26 vs Nov 24-Abr 25)</th>
        </tr>
        <tr>
          <th class="mat blk-mat" data-sort="m.ba">U Ant</th><th class="mat blk-mat" data-sort="m.bc">U Act</th>
          <th class="mat blk-mat" data-sort="m.msa">MS% Ant</th><th class="mat blk-mat" data-sort="m.msc">MS% Act</th>
          <th class="mat blk-mat" data-sort="m.ie">IE</th><th class="mat blk-mat" data-sort="m.pp">Var pp</th>
          <th class="ytd blk-ytd" data-sort="y.ba">U Ant</th><th class="ytd blk-ytd" data-sort="y.bc">U Act</th>
          <th class="ytd blk-ytd" data-sort="y.msa">MS% Ant</th><th class="ytd blk-ytd" data-sort="y.msc">MS% Act</th>
          <th class="ytd blk-ytd" data-sort="y.ie">IE</th><th class="ytd blk-ytd" data-sort="y.pp">Var pp</th>
          <th class="mes blk-mes" data-sort="p.ba">U Ant</th><th class="mes blk-mes" data-sort="p.bc">U Act</th>
          <th class="mes blk-mes" data-sort="p.msa">MS% Ant</th><th class="mes blk-mes" data-sort="p.msc">MS% Act</th>
          <th class="mes blk-mes" data-sort="p.ie">IE</th><th class="mes blk-mes" data-sort="p.pp">Var pp</th>
          <th class="trim blk-trim" data-sort="t.ba">U Ant</th><th class="trim blk-trim" data-sort="t.bc">U Act</th>
          <th class="trim blk-trim" data-sort="t.msa">MS% Ant</th><th class="trim blk-trim" data-sort="t.msc">MS% Act</th>
          <th class="trim blk-trim" data-sort="t.ie">IE</th><th class="trim blk-trim" data-sort="t.pp">Var pp</th>
          <th class="sem blk-sem" data-sort="s.ba">U Ant</th><th class="sem blk-sem" data-sort="s.bc">U Act</th>
          <th class="sem blk-sem" data-sort="s.msa">MS% Ant</th><th class="sem blk-sem" data-sort="s.msc">MS% Act</th>
          <th class="sem blk-sem" data-sort="s.ie">IE</th><th class="sem blk-sem" data-sort="s.pp">Var pp</th>
        </tr>
      </thead>
      <tbody id="body"></tbody>
    </table>
  </div>
  </div><!-- /exportArea -->

<script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js"></script>
<script>
const BRANDS = __BRANDS__;
const DATA = __DATA__;

function fmtNum(x){ return Math.round(x).toLocaleString('es-AR'); }
function fmtPct(x){ return (x*100).toFixed(1).replace('.', ',') + '%'; }
function fmtIE(x){ return x.toFixed(1).replace('.', ','); }
function fmtPP(x){ const s = x>0?'+':''; return s + x.toFixed(2).replace('.', ',') + ' pp'; }

function computeBlock(ba, bc, ma, mc){
  const tant = ba + ma, tact = bc + mc;
  const msa = tant>0 ? ba/tant : 0;
  const msc = tact>0 ? bc/tact : 0;
  let ie = 0;
  if (ba>0 && tant>0 && tact>0) ie = (bc/ba) / (tact/tant) * 100;
  const pp = (msc - msa) * 100;
  return {ba,bc,msa,msc,ie,pp};
}

function ieClass(ie, ba){ if (ba===0) return 'neutral'; return ie>=100 ? 'good' : 'bad'; }
function ppClass(pp){ return pp>0 ? 'good' : (pp<0 ? 'bad' : 'neutral'); }

function buildRow(brand, d){
  return {
    brand,
    d,
    m: computeBlock(d.ba, d.bc, d.ma, d.mc),
    y: computeBlock(d.ya, d.yc, d.ymka, d.ymkc),
    p: computeBlock(d.pa, d.pc, d.mpa, d.mpc),
    t: computeBlock(d.ta, d.tc, d.mta, d.mtc),
    s: computeBlock(d.sa, d.sc, d.msa6, d.msc6),
  };
}

function rowHTML(r, isTotal){
  const cls = isTotal ? 'total' : '';
  return `
    <tr data-brand="${r.brand}" class="${cls}">
      <td class="brand">${r.brand}</td>
      <td class="num blk-mat">${fmtNum(r.m.ba)}</td><td class="num blk-mat">${fmtNum(r.m.bc)}</td>
      <td class="num blk-mat">${fmtPct(r.m.msa)}</td><td class="num blk-mat">${fmtPct(r.m.msc)}</td>
      <td class="num blk-mat ${ieClass(r.m.ie,r.m.ba)}">${fmtIE(r.m.ie)}</td>
      <td class="num blk-mat ${ppClass(r.m.pp)}">${fmtPP(r.m.pp)}</td>
      <td class="num blk-ytd">${fmtNum(r.y.ba)}</td><td class="num blk-ytd">${fmtNum(r.y.bc)}</td>
      <td class="num blk-ytd">${fmtPct(r.y.msa)}</td><td class="num blk-ytd">${fmtPct(r.y.msc)}</td>
      <td class="num blk-ytd ${ieClass(r.y.ie,r.y.ba)}">${fmtIE(r.y.ie)}</td>
      <td class="num blk-ytd ${ppClass(r.y.pp)}">${fmtPP(r.y.pp)}</td>
      <td class="num blk-mes">${fmtNum(r.p.ba)}</td><td class="num blk-mes">${fmtNum(r.p.bc)}</td>
      <td class="num blk-mes">${fmtPct(r.p.msa)}</td><td class="num blk-mes">${fmtPct(r.p.msc)}</td>
      <td class="num blk-mes ${ieClass(r.p.ie,r.p.ba)}">${fmtIE(r.p.ie)}</td>
      <td class="num blk-mes ${ppClass(r.p.pp)}">${fmtPP(r.p.pp)}</td>
      <td class="num blk-trim">${fmtNum(r.t.ba)}</td><td class="num blk-trim">${fmtNum(r.t.bc)}</td>
      <td class="num blk-trim">${fmtPct(r.t.msa)}</td><td class="num blk-trim">${fmtPct(r.t.msc)}</td>
      <td class="num blk-trim ${ieClass(r.t.ie,r.t.ba)}">${fmtIE(r.t.ie)}</td>
      <td class="num blk-trim ${ppClass(r.t.pp)}">${fmtPP(r.t.pp)}</td>
      <td class="num blk-sem">${fmtNum(r.s.ba)}</td><td class="num blk-sem">${fmtNum(r.s.bc)}</td>
      <td class="num blk-sem">${fmtPct(r.s.msa)}</td><td class="num blk-sem">${fmtPct(r.s.msc)}</td>
      <td class="num blk-sem ${ieClass(r.s.ie,r.s.ba)}">${fmtIE(r.s.ie)}</td>
      <td class="num blk-sem ${ppClass(r.s.pp)}">${fmtPP(r.s.pp)}</td>
    </tr>`;
}

let sortState = { key: null, dir: null };

function getSortValue(row, key){
  if (key === 'brand') return row.brand;
  const [blk, fld] = key.split('.');
  return row[blk][fld];
}

function renderDesktop(seg){
  const body = document.getElementById('body');
  body.innerHTML = '';
  const rows = [];
  for(const brand of BRANDS){
    const d = DATA[seg][brand]; if(!d) continue;
    if (seg === 'POPULAR' && (d.ba + d.bc + d.ya + d.yc + d.pa + d.pc) === 0) continue;
    rows.push(buildRow(brand, d));
  }
  if (sortState.key){
    const k = sortState.key, dir = sortState.dir === 'asc' ? 1 : -1;
    rows.sort((a,b)=>{
      const va = getSortValue(a, k), vb = getSortValue(b, k);
      if (typeof va === 'string') return dir * va.localeCompare(vb);
      return dir * ((va||0) - (vb||0));
    });
  }
  body.insertAdjacentHTML('beforeend', rows.map(r=>rowHTML(r,false)).join(''));
  document.querySelectorAll('th[data-sort]').forEach(th=>{
    th.classList.remove('asc','desc');
    if (sortState.key === th.dataset.sort && sortState.dir) th.classList.add(sortState.dir);
  });
  renderSummary(rows);
}

function renderSummary(rows){
  const blocks = [
    {key:'m', name:'mat',  label:'MAT',  desc:'Abr-26 vs Abr-25 (12m)'},
    {key:'y', name:'ytd',  label:'YTD',  desc:'Ene-Abr 26 vs 25'},
    {key:'p', name:'mes',  label:'MES',  desc:'Abr 26 vs Abr 25'},
    {key:'t', name:'trim', label:'TRIM', desc:'Feb-Abr 26 vs 25'},
    {key:'s', name:'sem',  label:'SEM',  desc:'Nov 25-Abr 26 vs Nov 24-Abr 25'},
  ];
  const html = blocks.map(b=>{
    let pos = 0, neg = 0, flat = 0, nuevas = 0, negCreceU = 0;
    for (const r of rows){
      const x = r[b.key];
      if (x.ba === 0 && x.bc === 0) continue;
      if (x.ba === 0) { nuevas++; continue; }
      if (x.ie > 100) pos++;
      else if (x.ie < 100) {
        neg++;
        if (x.bc > x.ba) negCreceU++;
      } else flat++;
    }
    return `
      <div class="sum-card sum-${b.key} blk-${b.name}">
        <div class="sum-h"><span>${b.label}</span><span class="sum-desc">${b.desc}</span></div>
        <div class="sum-line">
          <span class="chip chip-pos">IE &gt; 100: <b>${pos}</b></span>
          <span class="chip chip-neg">IE &lt; 100: <b>${neg}</b></span>
          ${flat?`<span class="chip chip-flat">IE = 100: <b>${flat}</b></span>`:''}
        </div>
        <div class="sum-sub">De las ${neg} con IE&lt;100, <b>${negCreceU}</b> crecen en U (mercado creció más)
          ${nuevas?` · Nuevas (sin base): <b>${nuevas}</b>`:''}
        </div>
      </div>`;
  }).join('');
  document.getElementById('summary').innerHTML = html;
}

function applyBlockVisibility(){
  ['mat','ytd','mes','trim','sem'].forEach(k=>{
    const cb = document.querySelector('input[data-blk="'+k+'"]');
    if (!cb) return;
    const on = cb.checked;
    document.querySelectorAll('.blk-'+k).forEach(el=> el.classList.toggle('hidden', !on));
  });
}

function applyFilter(){
  const q = (document.getElementById('q').value || '').toUpperCase().trim();
  document.querySelectorAll('#body tr').forEach(el=>{
    const b = (el.dataset.brand || '').toUpperCase();
    el.classList.toggle('hidden', q && !b.includes(q));
  });
}

function render(seg){
  renderDesktop(seg);
  applyBlockVisibility();
  applyFilter();
}

document.querySelectorAll('#segBtns button').forEach(btn=>{
  btn.addEventListener('click', ()=>{
    document.querySelectorAll('#segBtns button').forEach(b=>b.classList.remove('active'));
    btn.classList.add('active');
    render(btn.dataset.seg);
  });
});
document.querySelectorAll('input[data-blk]').forEach(cb=>{
  cb.addEventListener('change', applyBlockVisibility);
});
document.getElementById('q').addEventListener('input', applyFilter);

/* ---------- Sort por click en headers ---------- */
document.querySelectorAll('th[data-sort]').forEach(th=>{
  th.addEventListener('click', ()=>{
    const k = th.dataset.sort;
    if (sortState.key === k){
      if (sortState.dir === 'asc') sortState.dir = 'desc';
      else if (sortState.dir === 'desc') { sortState.key = null; sortState.dir = null; }
      else sortState.dir = 'asc';
    } else {
      sortState.key = k;
      sortState.dir = (k === 'brand') ? 'asc' : 'desc';
    }
    const seg = document.querySelector('#segBtns .active').dataset.seg;
    render(seg);
  });
});

/* ---------- Export PDF / PNG ---------- */
async function captureTable(){
  document.body.classList.add('exporting');
  await new Promise(r => requestAnimationFrame(()=>requestAnimationFrame(r)));
  const target = document.querySelector('#exportArea');
  const table = document.querySelector('#tbl');
  const canvas = await html2canvas(target, {
    scale: 2, backgroundColor: '#ffffff', useCORS: true,
    windowWidth: Math.max(1400, table.scrollWidth + 40),
  });
  document.body.classList.remove('exporting');
  return canvas;
}

async function exportPDF(){
  const btn = document.getElementById('btnPdf');
  btn.disabled = true; const old = btn.textContent; btn.textContent = 'Generando PDF...';
  try {
    const canvas = await captureTable();
    const { jsPDF } = window.jspdf;
    /* Hoja a la medida del contenido para que TODO entre legible en una página */
    const w_mm = 420; /* A3 width */
    const ratio = canvas.height / canvas.width;
    const h_mm = w_mm * ratio + 12;
    const pdf = new jsPDF({ orientation: w_mm >= h_mm ? 'landscape' : 'portrait',
                            unit: 'mm', format: [w_mm, h_mm], compress: true });
    pdf.setFontSize(11); pdf.setTextColor(60);
    pdf.text('Comparativa Marca vs Mercado — ' + (document.querySelector('#segBtns .active')?.textContent || ''), 6, 8);
    pdf.addImage(canvas.toDataURL('image/jpeg', 0.92), 'JPEG', 6, 12, w_mm - 12, h_mm - 14);
    pdf.save('Comparativa_Marcas.pdf');
  } catch(e){ alert('No se pudo generar el PDF: ' + e.message); }
  finally { btn.disabled = false; btn.textContent = old; }
}

async function exportPNG(){
  const btn = document.getElementById('btnPng');
  btn.disabled = true; const old = btn.textContent; btn.textContent = 'Generando imagen...';
  try {
    const canvas = await captureTable();
    const a = document.createElement('a');
    a.href = canvas.toDataURL('image/png');
    a.download = 'Comparativa_Marcas.png';
    a.click();
  } catch(e){ alert('No se pudo generar la imagen: ' + e.message); }
  finally { btn.disabled = false; btn.textContent = old; }
}

document.getElementById('btnPdf').addEventListener('click', exportPDF);
document.getElementById('btnPng').addEventListener('click', exportPNG);

render('TOTAL');
</script>
</body>
</html>
"""
html = html.replace("__BRANDS__", js_brands).replace("__DATA__", js_payload)
with open(OUT_HTML, "w", encoding="utf-8") as f:
    f.write(html)
print(f"HTML guardado: {OUT_HTML}")
print("OK")
