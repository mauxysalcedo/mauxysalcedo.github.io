import numpy as np
import os
os.chdir("/home/user/mauxysalcedo.github.io/proyecto_final")
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.chart import LineChart, BarChart, Reference
from openpyxl.worksheet.datavalidation import DataValidation
import modelo as M

S = M.S
wb = Workbook()
NAVY = "0B1F3A"; GOLD = "F2B705"
HF = Font(bold=True, color="FFFFFF"); HFILL = PatternFill("solid", fgColor=NAVY)
INP = Font(color="0000FF"); BOLD = Font(bold=True)
thin = Side(style="thin", color="BBBBBB"); BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
GREEN = PatternFill("solid", bgColor="C6EFCE", fgColor="C6EFCE")
YEL = PatternFill("solid", bgColor="FFEB9C", fgColor="FFEB9C")
RED = PatternFill("solid", bgColor="FFC7CE", fgColor="FFC7CE")

def hdr(ws, row, vals, col=1):
    for i, v in enumerate(vals):
        c = ws.cell(row=row, column=col + i, value=v); c.font = HF; c.fill = HFILL
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True); c.border = BOX
def widths(ws, ws_w):
    for i, w in enumerate(ws_w, 1): ws.column_dimensions[L(i)].width = w
def title(ws, t, sub=None):
    ws["A1"] = t; ws["A1"].font = Font(bold=True, size=14, color=NAVY)
    if sub: ws["A2"] = sub; ws["A2"].font = Font(italic=True, color="666666")

# ------------------------------------------------------------ LÉAME
ws = wb.active; ws.title = "Leame"
title(ws, "Proyecto Final - Cementos Argos (cemento gris, Colombia)", "Consultoría simulada · Prospectiva Tecnológica & Producción Inteligente · Universidad Sergio Arboleda Barranquilla · Equipo: Brandon Afanador, Natalia Aragón, Mary Barrera y María Auxiliadora Salcedo Pérez")
lines = [
 "Este libro contiene todos los cálculos del proyecto. Flujo: Supuestos → Historico → Pronostico → Capacidad → Agregado → MPS → Tablero.",
 "Convención de color: AZUL = dato de entrada o supuesto (editable). NEGRO = fórmula. Cambie un supuesto en la hoja 'Supuestos' y todo se recalcula.",
 "Unidades: miles de toneladas (kt) de cemento por mes, salvo que se indique otra cosa. Dinero en pesos colombianos (COP).",
 "Fuente de la demanda histórica: serie oficial de despachos nacionales de cemento gris del DANE (Estadísticas de Cemento Gris, Anexo A1, boletín de septiembre de 2026; las cifras recientes son provisionales). La hoja 'Historico' contiene esa serie (celdas azules).",
 "Los datos de la hoja 'Tablero' (valores 'reales') son SIMULADOS para demostrar el funcionamiento del tablero.",
 "Verificar antes de entregar: ver columna 'Verificar' en la hoja Supuestos.",
]
for i, t in enumerate(lines, 4):
    ws.cell(row=i, column=1, value=t).alignment = Alignment(wrap_text=True, vertical="top"); ws.row_dimensions[i].height = 45
ws.column_dimensions["A"].width = 130

# ------------------------------------------------------------ SUPUESTOS
sp = wb.create_sheet("Supuestos"); title(sp, "Supuestos y datos de entrada", "Todas las demás hojas leen de aquí. Azul = editable.")
hdr(sp, 3, ["Parámetro", "Símbolo", "Valor", "Unidad", "Tipo", "Fuente / justificación", "Verificar"])
rows = [
 ("Despachos nacionales de cemento gris 2024", "nac24", S["nac_2024"], "kt", "Dato público", "DANE, Estadísticas de Cemento Gris (ECG), Anexo A1, boletín sep-2026", "No"),
 ("Despachos nacionales de cemento gris 2025", "nac25", S["nac_2025"], "kt", "Dato público", "DANE, ECG, Anexo A1, boletín sep-2026 (provisional)", "No"),
 ("Cemento Argos regional Colombia 2025", "arg25", S["argos_col_2025"], "kt", "Dato público", "Informe de Gestión Grupo Argos 2025 (5,0 Mt). Otra fuente reporta 5,4 Mt: alcances distintos", "Sí"),
 ("Participación de mercado de Argos (calculada)", "share", "=C6/C5", "%", "Calculado", "Argos Colombia / despachos nacionales 2025", "Sí"),
 ("Capacidad de diseño del horno (clínker)", "clk_dis", S["clinker_design"], "kt/año", "SUPUESTO", "Red Colombia; la capacidad total de Argos (14,4 Mt cemento) incluye otras regionales. Valor propio del equipo", "Sí"),
 ("Disponibilidad del horno", "disp", S["disp_horno"], "%", "SUPUESTO", "Paros programados (2-4 semanas) + no programados; rango típico del sector 85-90 %", "Sí"),
 ("Factor clínker (t clínker / t cemento)", "fclk", S["factor_clinker"], "t/t", "SUPUESTO", "Consistente con la mezcla de productos (0,746 ponderado); verificar con reporte de sostenibilidad de Argos", "Sí"),
 ("Capacidad de diseño de molienda de cemento", "mol_dis", S["mill_design"], "kt/año", "SUPUESTO", "Holgura de molienda sobre el horno (típica en cementeras)", "Sí"),
 ("Eficiencia de molienda", "mol_ef", S["mill_eff"], "%", "SUPUESTO", "Rango típico 88-92 %", "No"),
 ("Capacidad de diseño de ensacado", "bag_dis", S["bag_design"], "kt/año", "SUPUESTO", "No es información pública", "Declarar"),
 ("Eficiencia de ensacado", "bag_ef", S["bag_eff"], "%", "SUPUESTO", "Rango típico 80-88 % (paros por cambio de saco, mantenimiento)", "No"),
 ("Proporción en saco (UG + estructural)", "saco", "=C23+C24", "%", "Calculado", "Suma de mezcla en saco; el canal minorista predomina en Colombia", "Sí"),
 ("Precio promedio de venta", "precio", S["precio"], "COP/t", "SUPUESTO", "Orden de magnitud del cemento en Colombia; calibrar con ingresos/toneladas de Argos Colombia", "Sí"),
 ("Costo variable", "cvar", S["costo_var"], "COP/t", "SUPUESTO", "Energía, clínker, adiciones, empaque, flete variable; ~59 % del precio", "Sí"),
 ("Costo de producción (referencia de mantener inventario)", "cprod", S["costo_prod"], "COP/t", "SUPUESTO", "Costo de ventas / toneladas vendidas (Argos Colombia)", "Sí"),
 ("Costo anual de mantener inventario", "h", S["inv_anual"], "% anual", "SUPUESTO", "Rango típico de libros de texto 20-25 % (citar bibliografía del curso)", "No"),
 ("Sobrecosto sobre capacidad regular", "extra", S["sobrecosto_extra"], "COP/t", "SUPUESTO", "Compra de clínker / molienda adicional (~0,75 t x ~80.000 COP/t de sobreprecio)", "Declarar"),
 ("Costo de cambiar el ritmo de producción", "cambio", S["costo_cambio"], "COP/t cambiada", "SUPUESTO", "Arranques/paradas del horno y ajuste de turnos", "Declarar"),
 ("Capacidad adicional con recursos extra", "extra_cap", S["extra_cap"], "%", "SUPUESTO", "Máximo extra por compra de clínker y turnos", "Declarar"),
 ("Inventario inicial de cemento terminado", "inv0", S["inv_ini"], "kt", "SUPUESTO", "~10 días de demanda a diciembre de 2026", "Declarar"),
 ("Vida útil / máximo inventario terminado", "maxinv", S["max_inv_meses"], "meses", "SUPUESTO", "El cemento en saco se endurece con humedad; máximo ~1,5 meses de demanda", "Sí"),
 ("Mezcla: Uso General en saco", "mix_ug", S["mix"]["UG saco"], "%", "SUPUESTO", "Referencia clave del MPS", "Declarar"),
 ("Mezcla: Estructural en saco", "mix_est", S["mix"]["Estructural saco"], "%", "SUPUESTO", "", "Declarar"),
 ("Mezcla: Granel", "mix_gr", S["mix"]["Granel"], "%", "SUPUESTO", "Concreteras y obras grandes", "Declarar"),
 ("Mejora de disponibilidad del horno (explotar)", "dpp", M.EXPL_PP, "pp", "Escenario", "Recomendación 2", "No"),
 ("Factor clínker objetivo (ESG)", "fclk_n", M.FC_NEW, "t/t", "Escenario", "Recomendación 4", "No"),
 ("t CO2 por t de clínker", "co2", M.CO2_CLK, "t/t", "Referencia", "Valor de referencia sectorial", "Sí"),
]
R = {}
for i, (a, sym, v, u, t, f, ver) in enumerate(rows, 4):
    sp.cell(row=i, column=1, value=a); sp.cell(row=i, column=2, value=sym)
    c = sp.cell(row=i, column=3, value=v); sp.cell(row=i, column=4, value=u); sp.cell(row=i, column=5, value=t)
    sp.cell(row=i, column=6, value=f); sp.cell(row=i, column=7, value=ver)
    if not (isinstance(v, str) and v.startswith("=")): c.font = INP
    if u in ("%", "% anual", "pp") and not isinstance(v, str): c.number_format = "0.0%"
    elif sym == "share": c.number_format = "0.0%"
    R[sym] = f"Supuestos!$C${i}"
# corregir referencias de las dos fórmulas calculadas
sp["C7"] = f"={R['arg25'].split('!')[1]}/{R['nac25'].split('!')[1]}"; sp["C7"].number_format = "0.0%"
sp["C15"] = f"={R['mix_ug'].split('!')[1]}+{R['mix_est'].split('!')[1]}"; sp["C15"].number_format = "0.0%"
widths(sp, [48, 11, 14, 15, 13, 95, 10])
sp.freeze_panes = "A4"

# ------------------------------------------------------------ HISTÓRICO
hs = wb.create_sheet("Historico"); title(hs, "Demanda histórica (kt/mes)", "Serie oficial del DANE (ECG, Anexo A1, boletín sep-2026, cifras provisionales). Demanda Argos = despachos nacionales x participación de mercado.")
hdr(hs, 3, ["Mes", "Despachos nacionales (kt)", "Demanda Argos Colombia (kt)", "Observación"])
obs = {}
for i, d in enumerate(M.H_LABEL):
    obs[i] = "DANE ECG (provisional)" if d.year == 2026 else "DANE ECG"
for i, d in enumerate(M.H_LABEL):
    r = 4 + i
    hs.cell(row=r, column=1, value=d).number_format = "mmm-yy"
    c = hs.cell(row=r, column=2, value=round(float(M.HIST_NAC[i]), 2)); c.font = INP
    hs.cell(row=r, column=3, value=f"=B{r}*{R['share']}").number_format = "0.0"
    hs.cell(row=r, column=4, value=obs[i])
HN = 4 + len(M.H_LABEL) - 1       # última fila histórico
hs.cell(row=HN + 2, column=1, value="Control 2024 (debe = DANE)"); hs.cell(row=HN + 2, column=2, value="=SUM(B4:B15)")
hs.cell(row=HN + 3, column=1, value="Control 2025 (debe = DANE)"); hs.cell(row=HN + 3, column=2, value="=SUM(B16:B27)")
hs.cell(row=HN + 4, column=1, value="Control ene-jul 2026 (debe = DANE)"); hs.cell(row=HN + 4, column=2, value="=SUM(B28:B34)")
for r in range(HN + 2, HN + 5): hs.cell(row=r, column=2).number_format = "#,##0.0"
widths(hs, [30, 24, 28, 34])

# ------------------------------------------------------------ PRONÓSTICO
pr = wb.create_sheet("Pronostico"); title(pr, "Pronóstico de demanda Argos Colombia (kt/mes)", "Se comparan 3 métodos sobre la misma ventana de evaluación (ene-25 a sep-26). Se elige el de menor MAPE.")
hdr(pr, 3, ["Mes", "Real", "Prom. móvil 3m", "Suav. exp. simple", "Holt-Winters", "|e|/real MA", "|e|/real SES", "|e|/real HW", "|e| MA", "|e| SES", "|e| HW", "e HW"])
n = len(M.H_LABEL)
for i in range(n):
    r = 4 + i
    pr.cell(row=r, column=1, value=M.H_LABEL[i]).number_format = "mmm-yy"
    pr.cell(row=r, column=2, value=f"=Historico!C{4+i}").number_format = "0.0"
    for col, arr in ((3, M.f_ma), (4, M.f_ses), (5, M.f_hw)):
        v = arr[i]
        if not np.isnan(v):
            c = pr.cell(row=r, column=col, value=round(float(v), 2)); c.font = INP; c.number_format = "0.0"
    if i >= 13:
        for k, col in enumerate((3, 4, 5)):
            pr.cell(row=r, column=6 + k, value=f"=ABS(B{r}-{L(col)}{r})/B{r}").number_format = "0.0%"
            pr.cell(row=r, column=9 + k, value=f"=ABS(B{r}-{L(col)}{r})").number_format = "0.0"
        pr.cell(row=r, column=12, value=f"=B{r}-E{r}").number_format = "0.0"
e0, e1 = 4 + 13, 4 + n - 1
for j in range(14):
    r = 4 + n + j
    pr.cell(row=r, column=1, value=M.FUT_LABEL[j]).number_format = "mmm-yy"
    c = pr.cell(row=r, column=5, value=round(float(M.FUT_ARG[j]), 2)); c.font = INP; c.number_format = "0.0"
    pr.cell(row=r, column=2, value="Pronóstico")
FUT0 = 4 + n                 # fila oct-26
PLAN0 = FUT0 + 2             # fila dic-26 (primer mes del plan)
sm = e1 + 16
pr.cell(row=sm - 1, column=1, value="Medidas de error (ventana de evaluación)").font = BOLD
hdr(pr, sm, ["Método", "MAD (kt)", "MAPE", "Señal de rastreo"])
for k, (nm, col, ac, ap) in enumerate((("Promedio móvil 3 meses", "C", "I", "F"), ("Suavización exponencial simple", "D", "J", "G"), ("Holt-Winters", "E", "K", "H"))):
    r = sm + 1 + k
    pr.cell(row=r, column=1, value=nm)
    pr.cell(row=r, column=2, value=f"=AVERAGE({ac}{e0}:{ac}{e1})").number_format = "0.0"
    pr.cell(row=r, column=3, value=f"=AVERAGE({ap}{e0}:{ap}{e1})").number_format = "0.0%"
    pr.cell(row=r, column=4, value=f"=(SUM(B{e0}:B{e1})-SUM({col}{e0}:{col}{e1}))/B{r}").number_format = "0.00"
pr.cell(row=sm + 5, column=1, value="Método elegido:").font = BOLD
pr.cell(row=sm + 5, column=2, value=f'=INDEX(A{sm+1}:A{sm+3},MATCH(MIN(C{sm+1}:C{sm+3}),C{sm+1}:C{sm+3},0))')
SMROW = sm
pr.cell(row=sm + 7, column=1, value="Nota: los parámetros (α, β, γ) se eligieron por búsqueda en cuadrícula minimizando el MAPE; la serie tiene solo ~2,75 temporadas, por lo que el MAPE in-sample es optimista.").alignment = Alignment(wrap_text=True)
widths(pr, [30, 12, 16, 18, 14, 12, 12, 12, 10, 10, 10, 10])
ch = LineChart(); ch.title = "Demanda real y pronóstico (kt/mes)"; ch.height = 8; ch.width = 20
for col in (2, 5): ch.add_data(Reference(pr, min_col=col, min_row=3, max_row=4 + n + 13), titles_from_data=True)
ch.set_categories(Reference(pr, min_col=1, min_row=4, max_row=4 + n + 13)); pr.add_chart(ch, "N3")

# ------------------------------------------------------------ CAPACIDAD
cp = wb.create_sheet("Capacidad"); title(cp, "Análisis de capacidad (kt/mes)", "Capacidad diseñada, efectiva y utilización por etapa frente a la demanda pronosticada (dic-26 a nov-27).")
hdr(cp, 3, ["Etapa", "Diseñada anual", "Diseñada mensual", "Factor (disp./efic.)", "Efectiva mensual", "Requerimiento por t de cemento", "Requerimiento medio", "Requerimiento pico", "Utilización media", "Utilización pico", "Cemento soportado (kt/mes)"])
etapas = [("Horno (clínker)", R["clk_dis"], R["disp"], R["fclk"]), ("Molienda", R["mol_dis"], R["mol_ef"], "1"), ("Ensacado", R["bag_dis"], R["bag_ef"], R["saco"])]
d0, d1 = PLAN0, PLAN0 + 11
for k, (nm, dis, ef, req) in enumerate(etapas):
    r = 4 + k
    cp.cell(row=r, column=1, value=nm); cp.cell(row=r, column=2, value=f"={dis}"); cp.cell(row=r, column=3, value=f"=B{r}/12")
    cp.cell(row=r, column=4, value=f"={ef}").number_format = "0%"; cp.cell(row=r, column=5, value=f"=C{r}*D{r}")
    cp.cell(row=r, column=6, value=f"={req}")
    cp.cell(row=r, column=7, value=f"=AVERAGE(Pronostico!E{d0}:E{d1})*F{r}"); cp.cell(row=r, column=8, value=f"=MAX(Pronostico!E{d0}:E{d1})*F{r}")
    cp.cell(row=r, column=9, value=f"=G{r}/E{r}").number_format = "0.0%"; cp.cell(row=r, column=10, value=f"=H{r}/E{r}").number_format = "0.0%"
    cp.cell(row=r, column=11, value=f"=E{r}/F{r}")
    for c in (2, 3, 5, 6, 7, 8, 11): cp.cell(row=r, column=c).number_format = "#,##0.0"
cp["A8"] = "Etapa más cargada (mayor utilización media):"; cp["A8"].font = BOLD
cp["E8"] = "=INDEX(A4:A6,MATCH(MAX(I4:I6),I4:I6,0))"
cp["A9"] = "Capacidad regular de cemento del sistema (kt/mes) = mínimo de la columna K:"; cp["E9"] = "=MIN(K4:K6)"; cp["E9"].number_format = "0.0"
cp["A10"] = "Capacidad con recursos extra (kt/mes):"; cp["E10"] = f"=E9*(1+{R['extra_cap']})"; cp["E10"].number_format = "0.0"
cp["A12"] = "Utilización mensual por etapa"; cp["A12"].font = BOLD
hdr(cp, 13, ["Mes", "Demanda Argos (kt)", "Horno", "Molienda", "Ensacado", "Brecha vs cap. regular (kt)"])
for j in range(12):
    r = 14 + j
    cp.cell(row=r, column=1, value=f"=Pronostico!A{PLAN0+j}").number_format = "mmm-yy"
    cp.cell(row=r, column=2, value=f"=Pronostico!E{PLAN0+j}").number_format = "0.0"
    for k in range(3):
        cp.cell(row=r, column=3 + k, value=f"=$B{r}*$F${4+k}/$E${4+k}").number_format = "0.0%"
    cp.cell(row=r, column=6, value=f"=MAX(0,B{r}-$E$9)").number_format = "0.0"
cp.conditional_formatting.add("C14:E25", CellIsRule(operator="greaterThan", formula=["1"], fill=RED))
cp.conditional_formatting.add("C14:E25", CellIsRule(operator="between", formula=["0.9", "1"], fill=YEL))
widths(cp, [52, 16, 16, 18, 16, 20, 18, 16, 16, 16, 22])
cp.column_dimensions["A"].width = 52

# ------------------------------------------------------------ PLAN AGREGADO
ag = wb.create_sheet("Agregado"); title(ag, "Plan agregado (kt de cemento/mes) - 3 estrategias", "Producción (azul) es la decisión; inventario, faltantes y costos se calculan. Costos en millones de COP.")
ag["A3"] = "Costo mantener inventario (COP/t-mes):"; ag["D3"] = f"={R['cprod']}*{R['h']}/12"; ag["D3"].number_format = "#,##0"
ag["A4"] = "Capacidad regular (kt/mes):"; ag["D4"] = "=Capacidad!E9"; ag["D4"].number_format = "0.0"
ag["F4"] = "Capacidad con extra:"; ag["H4"] = "=Capacidad!E10"; ag["H4"].number_format = "0.0"
blocks = {}
def block(r0, name, P):
    ag.cell(row=r0, column=1, value=name).font = Font(bold=True, color=NAVY, size=12)
    hdr(ag, r0 + 1, ["Mes", "Demanda", "Producción", "Inv. final", "Faltante", "Sobre cap. regular", "Cambio de ritmo", "C. producción", "C. sobrecosto", "C. inventario", "C. cambio", "C. faltante", "Total"])
    for t in range(12):
        r = r0 + 2 + t
        ag.cell(row=r, column=1, value=f"=Pronostico!A{PLAN0+t}").number_format = "mmm-yy"
        ag.cell(row=r, column=2, value=f"=Pronostico!E{PLAN0+t}").number_format = "0.0"
        c = ag.cell(row=r, column=3, value=round(float(P[t]), 1)); c.font = INP
        prev = R["inv0"] if t == 0 else f"D{r-1}"
        ag.cell(row=r, column=4, value=f"=MAX(0,{prev}+C{r}-B{r})")
        ag.cell(row=r, column=5, value=f"=MAX(0,B{r}-({prev}+C{r}))")
        ag.cell(row=r, column=6, value=f"=MAX(0,C{r}-$D$4)")
        ag.cell(row=r, column=7, value="=0" if t == 0 else f"=ABS(C{r}-C{r-1})")
        ag.cell(row=r, column=8, value=f"=C{r}*1000*{R['cvar']}/1000000")
        ag.cell(row=r, column=9, value=f"=F{r}*1000*{R['extra']}/1000000")
        ag.cell(row=r, column=10, value=f"=D{r}*1000*$D$3/1000000")
        ag.cell(row=r, column=11, value=f"=G{r}*1000*{R['cambio']}/1000000")
        ag.cell(row=r, column=12, value=f"=E{r}*1000*{R['precio']}/1000000")
        ag.cell(row=r, column=13, value=f"=SUM(H{r}:L{r})")
        for cc in range(2, 14): ag.cell(row=r, column=cc).number_format = "#,##0.0"
    rt = r0 + 14
    ag.cell(row=rt, column=1, value="Total").font = BOLD
    for cc in (2, 3, 5, 6, 7, 8, 9, 10, 11, 12, 13):
        x = ag.cell(row=rt, column=cc, value=f"=SUM({L(cc)}{r0+2}:{L(cc)}{r0+13})"); x.font = BOLD; x.number_format = "#,##0.0"
    ag.cell(row=rt + 1, column=1, value="Control: inventario final >= inicial").font = Font(italic=True)
    ag.cell(row=rt + 1, column=4, value=f'=IF(D{r0+13}>={R["inv0"]}-0.001,"OK","REVISAR")')
    blocks[name] = (r0, rt)
rr = 6
for p in M.PLANES:
    block(rr, p["nombre"], p["P"]); rr += 18
# resumen
ag.cell(row=rr, column=1, value="Comparación de estrategias (millones de COP)").font = Font(bold=True, color=NAVY, size=12)
hdr(ag, rr + 1, ["Estrategia", "Producción", "Sobrecosto", "Inventario", "Cambio de ritmo", "Faltante", "TOTAL", "Faltante (kt)"])
tot_cells = []
for k, (nm, (r0, rt)) in enumerate(blocks.items()):
    r = rr + 2 + k
    ag.cell(row=r, column=1, value=nm)
    for cc, src in zip(range(2, 8), ("H", "I", "J", "K", "L", "M")):
        ag.cell(row=r, column=cc, value=f"={src}{rt}").number_format = "#,##0"
    ag.cell(row=r, column=8, value=f"=E{rt}").number_format = "0.0"
    tot_cells.append(f"G{r}")
RS = rr + 2
ag.cell(row=RS + 4, column=1, value="Estrategia de menor costo:").font = BOLD
ag.cell(row=RS + 4, column=2, value=f"=INDEX(A{RS}:A{RS+2},MATCH(MIN(G{RS}:G{RS+2}),G{RS}:G{RS+2},0))")
ag.cell(row=RS + 5, column=1, value="Índice elegido:"); ag.cell(row=RS + 5, column=2, value=f"=MATCH(MIN(G{RS}:G{RS+2}),G{RS}:G{RS+2},0)")
ag.cell(row=RS + 6, column=1, value="Ahorro de la mixta vs nivelación (MM COP):"); ag.cell(row=RS + 6, column=2, value=f"=G{RS}-G{RS+2}").number_format = "#,##0"
ag.cell(row=RS + 7, column=1, value="Ahorro de la mixta vs persecución (MM COP):"); ag.cell(row=RS + 7, column=2, value=f"=G{RS+1}-G{RS+2}").number_format = "#,##0"
# plan elegido
rp = RS + 10
ag.cell(row=rp, column=1, value="PLAN ELEGIDO (alimenta el MPS)").font = Font(bold=True, color=NAVY, size=12)
hdr(ag, rp + 1, ["Mes", "Demanda (kt)", "Producción plan (kt)"])
bl = list(blocks.values())
for t in range(12):
    r = rp + 2 + t
    ag.cell(row=r, column=1, value=f"=Pronostico!A{PLAN0+t}").number_format = "mmm-yy"
    ag.cell(row=r, column=2, value=f"=Pronostico!E{PLAN0+t}").number_format = "0.0"
    ag.cell(row=r, column=3, value=f"=CHOOSE($B${RS+5},C{bl[0][0]+2+t},C{bl[1][0]+2+t},C{bl[2][0]+2+t})").number_format = "0.0"
PE0 = rp + 2
widths(ag, [46, 12, 13, 12, 12, 16, 16, 14, 14, 14, 12, 12, 12])
bc = BarChart(); bc.title = "Costo total por estrategia (MM COP)"; bc.height = 7; bc.width = 14
bc.add_data(Reference(ag, min_col=7, min_row=RS - 1, max_row=RS + 2), titles_from_data=True); bc.set_categories(Reference(ag, min_col=1, min_row=RS, max_row=RS + 2)); ag.add_chart(bc, "J" + str(rr))

# ------------------------------------------------------------ MPS
mp = wb.create_sheet("MPS"); title(mp, f"MPS y ATP - {M.REF}", "12 semanas (dic-26 a feb-27, meses de 4 semanas). Campaña de empaque quincenal. Unidades: toneladas.")
mp["A3"] = "Participación de la referencia en la producción:"; mp["E3"] = f"={R['mix_ug']}"; mp["E3"].number_format = "0%"
mp["A4"] = "Inventario inicial de la referencia (t):"; mp["E4"] = f"={R['inv0']}*1000*E3"; mp["E4"].number_format = "#,##0"
hdr(mp, 6, ["Semana", "Mes", "Pronóstico semanal", "Pedidos firmes (simulados)", "MPS (producción)", "Inv. proyectado", "ATP"])
for w in range(12):
    r = 7 + w; mo = w // 4
    mp.cell(row=r, column=1, value=w + 1); mp.cell(row=r, column=2, value=mo + 1)
    mp.cell(row=r, column=3, value=f"=Agregado!B{PE0+mo}*1000*$E$3/4").number_format = "#,##0"
    c = mp.cell(row=r, column=4, value=float(M.PEDIDOS[w])); c.font = INP; c.number_format = "#,##0"
    if w % 2 == 0:
        mp.cell(row=r, column=5, value=f"=Agregado!C{PE0+mo}*1000*$E$3/2").number_format = "#,##0"
    else:
        mp.cell(row=r, column=5, value=0).number_format = "#,##0"
    prev = "$E$4" if w == 0 else f"F{r-1}"
    mp.cell(row=r, column=6, value=f"={prev}+E{r}-MAX(C{r},D{r})").number_format = "#,##0"
# ATP: semanas con MPS (1,3,5,..): ATP = (inv inicial si primera) + MPS - pedidos hasta próximo MPS
for k, w in enumerate(range(0, 12, 2)):
    r = 7 + w
    base = f"$E$4+E{r}" if k == 0 else f"E{r}"
    mp.cell(row=r, column=7, value=f"={base}-SUM(D{r}:D{r+1})").number_format = "#,##0"
mp.cell(row=20, column=1, value="Control de coherencia con el plan agregado").font = Font(bold=True, color=NAVY, size=12)
hdr(mp, 21, ["Mes", "MPS del mes (t)", "Plan agregado x participación (t)", "¿Cuadra?"])
for mo in range(3):
    r = 22 + mo
    mp.cell(row=r, column=1, value=mo + 1)
    mp.cell(row=r, column=2, value=f"=SUM(E{7+4*mo}:E{10+4*mo})").number_format = "#,##0"
    mp.cell(row=r, column=3, value=f"=Agregado!C{PE0+mo}*1000*$E$3").number_format = "#,##0"
    mp.cell(row=r, column=4, value=f'=IF(ABS(B{r}-C{r})<1,"OK","REVISAR")')
mp.cell(row=26, column=1, value="Regla ATP: primera campaña = inventario inicial + MPS - pedidos hasta la siguiente campaña; las demás = MPS - pedidos hasta la siguiente. Inventario proyectado usa el mayor entre pronóstico y pedidos.").alignment = Alignment(wrap_text=True)
widths(mp, [14, 10, 20, 24, 18, 16, 14])
lc = LineChart(); lc.title = "Inventario proyectado (t)"; lc.height = 7; lc.width = 14
lc.add_data(Reference(mp, min_col=6, min_row=6, max_row=18), titles_from_data=True); lc.set_categories(Reference(mp, min_col=1, min_row=7, max_row=18)); mp.add_chart(lc, "I6")

# ------------------------------------------------------------ ESCENARIOS (en la hoja Capacidad)
cp["A28"] = "Escenarios de mejora (respaldan las recomendaciones)"; cp["A28"].font = BOLD
cp["A29"] = "Cemento adicional anual por +2 pp de disponibilidad del horno (kt):"; cp["E29"] = f"=C4*{R['dpp']}*12/{R['fclk']}"; cp["E29"].number_format = "#,##0.0"
cp["A30"] = "Clínker liberado por bajar el factor clínker a 0,72 (kt/año):"; cp["E30"] = f"=SUM(Pronostico!E{PLAN0}:E{PLAN0+11})*({R['fclk']}-{R['fclk_n']})"; cp["E30"].number_format = "#,##0.0"
cp["A31"] = "CO2 evitado (kt/año):"; cp["E31"] = f"=E30*{R['co2']}"; cp["E31"].number_format = "#,##0.0"

# ------------------------------------------------------------ TABLERO
tb = wb.create_sheet("Tablero"); title(tb, "Tablero de indicadores - Cementos Argos Colombia", "Valores 'reales' SIMULADOS (plan ± variación aleatoria con semilla fija) para demostrar el funcionamiento. Seleccione el mes en la celda amarilla.")
rng = np.random.default_rng(11)
hdr(tb, 6, ["Mes", "Demanda real (kt)", "Producción real (kt)", "Disp. horno real", "Factor clínker real", "Pronóstico (kt)", "Plan (kt)",
            "Inv. final real (kt)", "Ventas (kt)", "MAPE mes", "Utilización horno", "Cumplimiento MPS", "Fill rate", "Factor clínker"])
fc_trend = np.linspace(0.752, 0.738, 12)
for t in range(12):
    r = 7 + t
    tb.cell(row=r, column=1, value=f"=Pronostico!A{PLAN0+t}").number_format = "mmm-yy"
    tb.cell(row=r, column=2, value=round(float(M.DEM[t] * (1 + rng.normal(0, 0.035))), 1)).font = INP
    tb.cell(row=r, column=3, value=round(float(M.MEJOR_PLAN["P"][t] * (1 + rng.normal(-0.005, 0.02))), 1)).font = INP
    tb.cell(row=r, column=4, value=round(float(0.88 + rng.normal(0, 0.012)), 3)).font = INP
    tb.cell(row=r, column=5, value=round(float(fc_trend[t] + rng.normal(0, 0.003)), 3)).font = INP
    tb.cell(row=r, column=6, value=f"=Pronostico!E{PLAN0+t}")
    tb.cell(row=r, column=7, value=f"=Agregado!C{PE0+t}")
    prev = R["inv0"] if t == 0 else f"H{r-1}"
    tb.cell(row=r, column=8, value=f"=MAX(0,{prev}+C{r}-B{r})")
    tb.cell(row=r, column=9, value=f"=MIN(B{r},{prev}+C{r})")
    tb.cell(row=r, column=10, value=f"=ABS(B{r}-F{r})/B{r}")
    tb.cell(row=r, column=11, value=f"=C{r}*E{r}/(Capacidad!$C$4*D{r})")
    tb.cell(row=r, column=12, value=f"=C{r}/G{r}")
    tb.cell(row=r, column=13, value=f"=I{r}/B{r}")
    tb.cell(row=r, column=14, value=f"=E{r}")
    for c in (2, 3, 6, 7, 8, 9): tb.cell(row=r, column=c).number_format = "0.0"
    for c in (4, 10, 11, 12, 13): tb.cell(row=r, column=c).number_format = "0.0%"
    for c in (5, 14): tb.cell(row=r, column=c).number_format = "0.000"
tb["A3"] = "Mes seleccionado:"; tb["A3"].font = BOLD
tb["B3"] = "=A18"; tb["B3"].number_format = "mmm-yy"; tb["B3"].fill = PatternFill("solid", fgColor="FFF2CC")
dv = DataValidation(type="list", formula1="=$A$7:$A$18", allow_blank=False); tb.add_data_validation(dv); dv.add("B3")
tb["B3"].value = "=A7"
# Tarjetas de KPI
hdr(tb, 21, ["Indicador", "Valor (mes)", "Promedio 12 m", "Meta", "Umbral amarillo", "Estado", "Sección origen"])
kpis = [
 ("MAPE del pronóstico", "J", "<=", 0.05, 0.08, "0.0%", "C2 Pronóstico"),
 ("Utilización del horno", "K", "band", 0.85, 0.80, "0.0%", "C3 Capacidad"),
 ("Cumplimiento del MPS", "L", ">=", 0.97, 0.93, "0.0%", "C4 / C5"),
 ("Fill rate (nivel de servicio)", "M", ">=", 0.98, 0.95, "0.0%", "C5"),
 ("Factor clínker (ESG)", "N", "<=", 0.74, 0.75, "0.000", "ESG"),
]
for k, (nm, col, op, meta, amb, fmt, src) in enumerate(kpis):
    r = 22 + k
    tb.cell(row=r, column=1, value=nm)
    tb.cell(row=r, column=2, value=f"=INDEX({col}7:{col}18,MATCH($B$3,$A$7:$A$18,0))").number_format = fmt
    tb.cell(row=r, column=3, value=f"=AVERAGE({col}7:{col}18)").number_format = fmt
    tb.cell(row=r, column=4, value=meta).number_format = fmt; tb.cell(row=r, column=4).font = INP
    tb.cell(row=r, column=5, value=amb).number_format = fmt; tb.cell(row=r, column=5).font = INP
    if op == "<=":
        f = f'=IF(B{r}<=D{r},"VERDE",IF(B{r}<=E{r},"AMARILLO","ROJO"))'
    elif op == ">=":
        f = f'=IF(B{r}>=D{r},"VERDE",IF(B{r}>=E{r},"AMARILLO","ROJO"))'
    else:   # banda: >= meta (90%) verde hasta 100 %; sobre 100% rojo (sobreutilización)
        f = f'=IF(AND(B{r}>=D{r},B{r}<=1),"VERDE",IF(AND(B{r}>=E{r},B{r}<=1.05),"AMARILLO","ROJO"))'
    tb.cell(row=r, column=6, value=f); tb.cell(row=r, column=7, value=src)
tb.conditional_formatting.add("F22:F26", CellIsRule(operator="equal", formula=['"VERDE"'], fill=GREEN))
tb.conditional_formatting.add("F22:F26", CellIsRule(operator="equal", formula=['"AMARILLO"'], fill=YEL))
tb.conditional_formatting.add("F22:F26", CellIsRule(operator="equal", formula=['"ROJO"'], fill=RED))
tb["A28"] = "Metas: MAPE ≤5 % · utilización del horno 85-100 % (el plan produce por debajo de la capacidad en los meses de baja demanda) · cumplimiento MPS ≥97 % · fill rate ≥98 % · factor clínker ≤0,74 (objetivo ESG; base 0,75)."
tb["A28"].font = Font(italic=True, color="666666")
for col, ttl, anchor in (("K", "Utilización del horno", "I20"), ("M", "Fill rate", "I36")):
    c = LineChart(); c.title = ttl; c.height = 7; c.width = 14
    c.add_data(Reference(tb, min_col=11 if col == "K" else 13, min_row=6, max_row=18), titles_from_data=True)
    c.set_categories(Reference(tb, min_col=1, min_row=7, max_row=18)); tb.add_chart(c, anchor)
c = BarChart(); c.title = "Producción real vs plan (kt)"; c.height = 7; c.width = 14
for col in (3, 7): c.add_data(Reference(tb, min_col=col, min_row=6, max_row=18), titles_from_data=True)
c.set_categories(Reference(tb, min_col=1, min_row=7, max_row=18)); tb.add_chart(c, "A32")
widths(tb, [36, 16, 18, 14, 16, 16, 16, 16, 12, 12, 16, 16, 12, 14])

wb.save("Proyecto_Argos_Modelo_y_Tablero.xlsx")
print("ok", RS, PE0, PLAN0, SMROW)
