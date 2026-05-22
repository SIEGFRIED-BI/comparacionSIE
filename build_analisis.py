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
for c in [COL_MAT_ACT, COL_MAT_ANT, COL_YTD_ACT, COL_YTD_ANT, COL_MES_ACT, COL_MES_ANT]:
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

  /* ----- DESKTOP / TABLET (table) ----- */
  .desktop { display: block; }
  .mobile  { display: none; }

  .tbl-wrap { background: var(--card); border-radius: 10px; overflow: auto;
              box-shadow: 0 1px 3px rgba(0,0,0,.06); }
  table { border-collapse: collapse; width: 100%; font-size: 12px; min-width: 1100px; }
  th, td { border: 1px solid var(--line); padding: 6px 8px; }
  th { background: var(--hdr); color: #fff; text-align: center; font-weight: 600;
       position: sticky; top: 0; z-index: 2; }
  th.mat { background: var(--mat); } th.ytd { background: var(--ytd); } th.mes { background: var(--mes); }
  td.brand { font-weight: 700; background: #f7f9fc; white-space: nowrap; position: sticky; left: 0; z-index: 1;
             border-right: 2px solid #c0c0c0; }
  td.num { text-align: right; font-variant-numeric: tabular-nums; white-space: nowrap; }
  tr:nth-child(even) td:not(.brand) { background: #f9fbfd; }
  td.good { background: #e6f4ea !important; color: #0a6e0a; }
  td.bad  { background: #fdecec !important; color: #9c0006; }
  td.neutral { color: #555; }
  tr.total td { background: #FFE699 !important; font-weight: 800; }

  /* ----- MOBILE (cards) ----- */
  @media (max-width: 760px) {
    body { padding: 10px; }
    .desktop { display: none; }
    .mobile  { display: block; }
  }
  .cards { display: grid; gap: 10px; }
  .card {
    background: var(--card); border-radius: 12px; padding: 12px;
    box-shadow: 0 1px 3px rgba(0,0,0,.08);
  }
  .card .brand { font-size: 15px; font-weight: 700; margin-bottom: 6px; display: flex; justify-content: space-between; align-items: baseline; }
  .card .brand .mol { font-size: 11px; font-weight: 400; color: var(--muted); }
  .blocks { display: grid; grid-template-columns: 1fr; gap: 8px; }
  .block { border-radius: 8px; padding: 8px 10px; }
  .block.mat { background: #eef3fc; } .block.ytd { background: #fdeede; } .block.mes { background: #ebf3e3; }
  .block .label { display: flex; justify-content: space-between; align-items: baseline; font-weight: 700; font-size: 13px; }
  .block .label .periodo { font-size: 11px; color: var(--muted); font-weight: 500; }
  .grid {
    display: grid; grid-template-columns: 1fr 1fr; gap: 4px 10px;
    font-size: 12px; margin-top: 6px;
  }
  .grid .k { color: var(--muted); }
  .grid .v { text-align: right; font-variant-numeric: tabular-nums; font-weight: 600; }
  .grid .v.good { color: #0a6e0a; }
  .grid .v.bad  { color: #9c0006; }
  .card.total { background: #fff9e6; border: 1px solid #f1c40f; }
  .hidden { display: none !important; }

  .footer { color: var(--muted); font-size: 11px; margin-top: 14px; text-align: center; }
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
    </span>
    <input class="search" id="q" placeholder="Buscar marca..."/>
  </div>

  <div class="desktop">
    <div class="tbl-wrap">
      <table id="tbl">
        <thead>
          <tr>
            <th rowspan="2">Marca</th>
            <th class="mat" colspan="6">MAT (Apr 2026 vs Apr 2025)</th>
            <th class="ytd" colspan="6">YTD (Ene-Abr 26 vs 25)</th>
            <th class="mes" colspan="6">MES (Apr 2026 vs Apr 2025)</th>
          </tr>
          <tr>
            <th class="mat">U Ant</th><th class="mat">U Act</th>
            <th class="mat">MS% Ant</th><th class="mat">MS% Act</th>
            <th class="mat">IE</th><th class="mat">Var pp</th>
            <th class="ytd">U Ant</th><th class="ytd">U Act</th>
            <th class="ytd">MS% Ant</th><th class="ytd">MS% Act</th>
            <th class="ytd">IE</th><th class="ytd">Var pp</th>
            <th class="mes">U Ant</th><th class="mes">U Act</th>
            <th class="mes">MS% Ant</th><th class="mes">MS% Act</th>
            <th class="mes">IE</th><th class="mes">Var pp</th>
          </tr>
        </thead>
        <tbody id="body"></tbody>
      </table>
    </div>
  </div>

  <div class="mobile">
    <div class="cards" id="cards"></div>
  </div>

  <div class="footer">v3 · IE = ((Marca Act / Marca Ant) / (Mercado Act / Mercado Ant)) × 100</div>

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

function renderDesktop(seg){
  const body = document.getElementById('body');
  body.innerHTML = '';
  let T = {ba:0,bc:0,ma:0,mc:0,ya:0,yc:0,ymka:0,ymkc:0,pa:0,pc:0,mpa:0,mpc:0};
  for(const brand of BRANDS){
    const d = DATA[seg][brand]; if(!d) continue;
    T.ba+=d.ba; T.bc+=d.bc; T.ma+=d.ma; T.mc+=d.mc;
    T.ya+=d.ya; T.yc+=d.yc; T.ymka+=d.ymka; T.ymkc+=d.ymkc;
    T.pa+=d.pa; T.pc+=d.pc; T.mpa+=d.mpa; T.mpc+=d.mpc;
    const m = computeBlock(d.ba, d.bc, d.ma, d.mc);
    const y = computeBlock(d.ya, d.yc, d.ymka, d.ymkc);
    const p = computeBlock(d.pa, d.pc, d.mpa, d.mpc);
    body.insertAdjacentHTML('beforeend', `
      <tr data-brand="${brand}">
        <td class="brand">${brand}</td>
        <td class="num blk-mat">${fmtNum(d.ba)}</td>
        <td class="num blk-mat">${fmtNum(d.bc)}</td>
        <td class="num blk-mat">${fmtPct(m.msa)}</td>
        <td class="num blk-mat">${fmtPct(m.msc)}</td>
        <td class="num blk-mat ${ieClass(m.ie,d.ba)}">${fmtIE(m.ie)}</td>
        <td class="num blk-mat ${ppClass(m.pp)}">${fmtPP(m.pp)}</td>
        <td class="num blk-ytd">${fmtNum(d.ya)}</td>
        <td class="num blk-ytd">${fmtNum(d.yc)}</td>
        <td class="num blk-ytd">${fmtPct(y.msa)}</td>
        <td class="num blk-ytd">${fmtPct(y.msc)}</td>
        <td class="num blk-ytd ${ieClass(y.ie,d.ya)}">${fmtIE(y.ie)}</td>
        <td class="num blk-ytd ${ppClass(y.pp)}">${fmtPP(y.pp)}</td>
        <td class="num blk-mes">${fmtNum(d.pa)}</td>
        <td class="num blk-mes">${fmtNum(d.pc)}</td>
        <td class="num blk-mes">${fmtPct(p.msa)}</td>
        <td class="num blk-mes">${fmtPct(p.msc)}</td>
        <td class="num blk-mes ${ieClass(p.ie,d.pa)}">${fmtIE(p.ie)}</td>
        <td class="num blk-mes ${ppClass(p.pp)}">${fmtPP(p.pp)}</td>
      </tr>`);
  }
  const m = computeBlock(T.ba, T.bc, T.ma, T.mc);
  const y = computeBlock(T.ya, T.yc, T.ymka, T.ymkc);
  const p = computeBlock(T.pa, T.pc, T.mpa, T.mpc);
  body.insertAdjacentHTML('beforeend', `
    <tr class="total">
      <td class="brand">TOTAL CARTERA</td>
      <td class="num blk-mat">${fmtNum(T.ba)}</td><td class="num blk-mat">${fmtNum(T.bc)}</td>
      <td class="num blk-mat">${fmtPct(m.msa)}</td><td class="num blk-mat">${fmtPct(m.msc)}</td>
      <td class="num blk-mat ${ieClass(m.ie,T.ba)}">${fmtIE(m.ie)}</td>
      <td class="num blk-mat ${ppClass(m.pp)}">${fmtPP(m.pp)}</td>
      <td class="num blk-ytd">${fmtNum(T.ya)}</td><td class="num blk-ytd">${fmtNum(T.yc)}</td>
      <td class="num blk-ytd">${fmtPct(y.msa)}</td><td class="num blk-ytd">${fmtPct(y.msc)}</td>
      <td class="num blk-ytd ${ieClass(y.ie,T.ya)}">${fmtIE(y.ie)}</td>
      <td class="num blk-ytd ${ppClass(y.pp)}">${fmtPP(y.pp)}</td>
      <td class="num blk-mes">${fmtNum(T.pa)}</td><td class="num blk-mes">${fmtNum(T.pc)}</td>
      <td class="num blk-mes">${fmtPct(p.msa)}</td><td class="num blk-mes">${fmtPct(p.msc)}</td>
      <td class="num blk-mes ${ieClass(p.ie,T.pa)}">${fmtIE(p.ie)}</td>
      <td class="num blk-mes ${ppClass(p.pp)}">${fmtPP(p.pp)}</td>
    </tr>`);
}

function renderMobile(seg){
  const cont = document.getElementById('cards');
  cont.innerHTML = '';
  let T = {ba:0,bc:0,ma:0,mc:0,ya:0,yc:0,ymka:0,ymkc:0,pa:0,pc:0,mpa:0,mpc:0};
  for(const brand of BRANDS){
    const d = DATA[seg][brand]; if(!d) continue;
    T.ba+=d.ba; T.bc+=d.bc; T.ma+=d.ma; T.mc+=d.mc;
    T.ya+=d.ya; T.yc+=d.yc; T.ymka+=d.ymka; T.ymkc+=d.ymkc;
    T.pa+=d.pa; T.pc+=d.pc; T.mpa+=d.mpa; T.mpc+=d.mpc;
    const m = computeBlock(d.ba, d.bc, d.ma, d.mc);
    const y = computeBlock(d.ya, d.yc, d.ymka, d.ymkc);
    const p = computeBlock(d.pa, d.pc, d.mpa, d.mpc);
    cont.insertAdjacentHTML('beforeend', `
      <div class="card" data-brand="${brand}">
        <div class="brand"><span>${brand}</span></div>
        <div class="blocks">
          ${blockHTML('mat', 'MAT', 'Abr-26 vs Abr-25 (12m)', d.ba, d.bc, m)}
          ${blockHTML('ytd', 'YTD', 'Ene-Abr 26 vs 25', d.ya, d.yc, y)}
          ${blockHTML('mes', 'MES', 'Abr 2026 vs Abr 2025', d.pa, d.pc, p)}
        </div>
      </div>`);
  }
  const m = computeBlock(T.ba, T.bc, T.ma, T.mc);
  const y = computeBlock(T.ya, T.yc, T.ymka, T.ymkc);
  const p = computeBlock(T.pa, T.pc, T.mpa, T.mpc);
  cont.insertAdjacentHTML('beforeend', `
    <div class="card total" data-brand="__total__">
      <div class="brand"><span>TOTAL CARTERA</span></div>
      <div class="blocks">
        ${blockHTML('mat', 'MAT', 'Abr-26 vs Abr-25 (12m)', T.ba, T.bc, m)}
        ${blockHTML('ytd', 'YTD', 'Ene-Abr 26 vs 25', T.ya, T.yc, y)}
        ${blockHTML('mes', 'MES', 'Abr 2026 vs Abr 2025', T.pa, T.pc, p)}
      </div>
    </div>`);
}

function blockHTML(klass, lbl, per, ba, bc, c){
  const ieCls = ieClass(c.ie, ba);
  const ppCls = ppClass(c.pp);
  return `
    <div class="block ${klass} blk-${klass}">
      <div class="label"><span>${lbl}</span><span class="periodo">${per}</span></div>
      <div class="grid">
        <span class="k">U Marca Ant</span><span class="v">${fmtNum(ba)}</span>
        <span class="k">U Marca Act</span><span class="v">${fmtNum(bc)}</span>
        <span class="k">MS% Ant</span><span class="v">${fmtPct(c.msa)}</span>
        <span class="k">MS% Act</span><span class="v">${fmtPct(c.msc)}</span>
        <span class="k">IE</span><span class="v ${ieCls}">${fmtIE(c.ie)}</span>
        <span class="k">Var pp</span><span class="v ${ppCls}">${fmtPP(c.pp)}</span>
      </div>
    </div>`;
}

function applyBlockVisibility(){
  ['mat','ytd','mes'].forEach(k=>{
    const on = document.querySelector('input[data-blk="'+k+'"]').checked;
    // desktop: ocultar columnas con clase blk-{k} y los th del bloque
    document.querySelectorAll('.blk-'+k).forEach(el=> el.classList.toggle('hidden', !on));
    document.querySelectorAll('th.'+k).forEach(el=> el.classList.toggle('hidden', !on));
  });
}

function applyFilter(){
  const q = (document.getElementById('q').value || '').toUpperCase().trim();
  document.querySelectorAll('#body tr, #cards .card').forEach(el=>{
    const b = (el.dataset.brand || '').toUpperCase();
    if (b === '__TOTAL__') return;
    el.classList.toggle('hidden', q && !b.includes(q));
  });
}

function render(seg){
  renderDesktop(seg);
  renderMobile(seg);
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
