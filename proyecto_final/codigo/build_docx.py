import sys, json, numpy as np
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT, WD_TAB_LEADER, WD_BREAK
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import os
import modelo as M
BASE = "/home/user/mauxysalcedo.github.io/proyecto_final"
os.chdir(BASE)

S = M.S
PAGES = json.load(open("toc_pages.json")) if len(sys.argv) > 1 and sys.argv[1] == "final" else {}
NAVY = RGBColor(0x0B, 0x1F, 0x3A)

def f0(x): return f"{x:,.0f}".replace(",", ".")
def f1(x): return f"{x:,.1f}".replace(",", "X").replace(".", ",").replace("X", ".")
def pc(x, d=1): return f"{x*100:.{d}f} %".replace(".", ",")
def mm(x): return f0(x / 1e6)          # COP -> millones

d = Document()
sec = d.sections[0]
sec.page_width, sec.page_height = Cm(21.59), Cm(27.94)
FONT = "Times New Roman"
st = d.styles["Normal"]; st.font.name = FONT; st.font.size = Pt(12)
st.element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
st.paragraph_format.space_after = Pt(0); st.paragraph_format.space_before = Pt(0); st.paragraph_format.line_spacing = 2.0
for nm in ("Heading 1", "Heading 2", "Heading 3", "List Bullet"):
    h = d.styles[nm]; h.font.name = FONT; h.font.size = Pt(12); h.font.color.rgb = RGBColor(0, 0, 0)
    h.element.rPr.rFonts.set(qn("w:eastAsia"), FONT); h.element.rPr.rFonts.set(qn("w:ascii"), FONT); h.element.rPr.rFonts.set(qn("w:hAnsi"), FONT)
    h.paragraph_format.line_spacing = 2.0; h.paragraph_format.space_before = Pt(0); h.paragraph_format.space_after = Pt(0)
for nm in ("Heading 1", "Heading 2", "Heading 3"):
    h = d.styles[nm]; h.font.bold = True; h.paragraph_format.keep_with_next = True
d.styles["Heading 1"].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
d.styles["Heading 2"].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
sec.left_margin = sec.right_margin = sec.top_margin = sec.bottom_margin = Cm(2.54)
FI = Cm(1.27)   # sangría de primera línea de 0,5 in

def para(text, bold=False, italic=False, align=None, size=None, color=None, after=None, style=None, indent=None):
    p = d.add_paragraph(style=style)
    parts = text.split("**")
    for i, t in enumerate(parts):
        if not t: continue
        r = p.add_run(t); r.bold = bold or (i % 2 == 1); r.italic = italic
        if size: r.font.size = Pt(size)
    if align == "c": p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if indent is None: indent = (style is None and align != "c" and size is None)
    if indent: p.paragraph_format.first_line_indent = FI
    if after is not None: p.paragraph_format.space_after = Pt(after)
    return p

def bullets(items):
    for t in items:
        p = para(t, style="List Bullet")

def cell_border(cell, edge, sz=8):
    tcPr = cell._tc.get_or_add_tcPr(); bd = tcPr.find(qn("w:tcBorders"))
    if bd is None: bd = OxmlElement("w:tcBorders"); tcPr.append(bd)
    e = OxmlElement(f"w:{edge}"); e.set(qn("w:val"), "single"); e.set(qn("w:sz"), str(sz)); e.set(qn("w:color"), "000000"); bd.append(e)

def single(p, after=0, keep=False):
    p.paragraph_format.line_spacing = 1.0; p.paragraph_format.space_after = Pt(after); p.paragraph_format.first_line_indent = None
    if keep: p.paragraph_format.keep_with_next = True

def caption_above(caption):
    num, ttl = caption.split(". ", 1)
    a = para(num, bold=True, size=12); single(a, 0, True)
    b = para(ttl, italic=True, size=12); single(b, 4, True)

def nota_below(texto):
    p = d.add_paragraph(); r = p.add_run("Nota. "); r.italic = True; r.font.size = Pt(10)
    r2 = p.add_run(texto); r2.font.size = Pt(10); single(p, 10)

NOTAS = {
 "Cifras públicas": "Datos del Departamento Administrativo Nacional de Estadística [DANE] (2026) y de Grupo Argos (2026). La participación de mercado fue calculada por el equipo.",
 "MPS simplificado": "Los pedidos firmes son simulados. Elaboración propia con el modelo de Excel adjunto.",
 "Datos simulados": "Datos simulados con el modelo de Excel adjunto (hoja Tablero).",
 "Supuestos consolidados": "Elaboración propia. La columna Verificar indica qué debe contrastarse con fuentes antes de la entrega.",
 "Demanda histórica": "Serie histórica reconstruida con base en datos del DANE (2026). Elaboración propia.",
}
def nota_para(caption):
    for k, v in NOTAS.items():
        if k in caption: return v
    return "Elaboración propia con el modelo de Excel adjunto."

def table(headers, rows, widths=None, size=10, caption=None):
    if caption: caption_above(caption)
    t = d.add_table(rows=1, cols=len(headers)); t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, h in enumerate(headers):
        c = t.rows[0].cells[i]; c.text = ""; p = c.paragraphs[0]; r = p.add_run(h); r.bold = True; r.font.size = Pt(size)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER; single(p)
        cell_border(c, "top"); cell_border(c, "bottom")
    for k, row in enumerate(rows):
        cells = t.add_row().cells
        for i, v in enumerate(row):
            cells[i].text = ""; p = cells[i].paragraphs[0]; r = p.add_run(str(v)); r.font.size = Pt(size); single(p)
            if i > 0 and len(str(v)) < 16: p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            if k == len(rows) - 1: cell_border(cells[i], "bottom")
    if widths:
        t.autofit = False
        for i, w in enumerate(widths): t.columns[i].width = Cm(w)
        for row in t.rows:
            for i, w in enumerate(widths): row.cells[i].width = Cm(w)
    for ri, row in enumerate(t.rows):
        trPr = row._tr.get_or_add_trPr(); cs = OxmlElement("w:cantSplit"); trPr.append(cs)
        if ri < len(t.rows) - 1:
            for c in row.cells:
                for p_ in c.paragraphs: p_.paragraph_format.keep_with_next = True
    if caption: nota_below(nota_para(caption))
    return t

def fig(path, width, caption):
    caption_above(caption)
    p = d.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER; single(p, 0, True)
    p.add_run().add_picture(path, width=Cm(width))
    nota_below(nota_para(caption))

def h1(t, newpage=False):
    if newpage: d.add_page_break()
    return d.add_heading(t, 1)

# --------------------------------------------------------------- pie de página
def add_header():
    hp = sec.header.paragraphs[0]; hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    for tag, txt in (("begin", None), (None, "PAGE"), ("end", None)):
        run = hp.add_run(); run.font.size = Pt(12); run.font.name = FONT
        if tag:
            e = OxmlElement("w:fldChar"); e.set(qn("w:fldCharType"), tag); run._r.append(e)
        else:
            e = OxmlElement("w:instrText"); e.set(qn("xml:space"), "preserve"); e.text = txt; run._r.append(e)
add_header()

# --------------------------------------------------------------- datos
P = M.MEJOR_PLAN["P"]; A, B, Cc = M.PLANES
rA = A["rows"]; rC = Cc["rows"]
inv_max = max(r["inv"] for r in rC); inv_max_m = rC[[r["inv"] for r in rC].index(inv_max)]["mes"]
ES = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"]
mes = lambda dt: f"{ES[dt.month-1]}-{str(dt.year)[2:]}"
DEMs = M.DEM.sum()
cap_year = M.CAP_CEM_REG * 12
yrs = np.log(cap_year / DEMs) / np.log(1.05)
ahorro_AC = (A["total"] - Cc["total"]) / 1e6; ahorro_BC = (B["total"] - Cc["total"]) / 1e6
esc_disp = (M.ESC_BASE["total"] - M.ESC_DISP["total"]) / 1e6; esc_fc = (M.ESC_BASE["total"] - M.ESC_FC["total"]) / 1e6
clk_def = np.maximum(0, M.CLK_REQ - M.CLK_CAP).sum(); val_mix = clk_def * 1000 * (M.TOC["UG saco"]["T_clk"] - M.TOC["Granel"]["T_clk"]) / 1e6
months_over = [mes(M.PLAN_LABEL[i]) for i in range(12) if M.CLK_REQ[i] > M.CLK_CAP]
from openpyxl import load_workbook
tb = load_workbook("Proyecto_Argos_Modelo_y_Tablero.xlsx", data_only=True)["Tablero"]
kpi = {tb.cell(row=r, column=1).value: (tb.cell(row=r, column=3).value, tb.cell(row=r, column=6).value) for r in range(22, 27)}

def sens(cambio, inv):
    old = (S["costo_cambio"], S["inv_anual"], M.h_mes)
    S["costo_cambio"] = cambio; S["inv_anual"] = inv; M.h_mes = M.ci * inv / 12
    try: pl = [M.plan_nivelado(), M.plan_persecucion(), M.plan_mixto()]
    finally: S["costo_cambio"], S["inv_anual"], M.h_mes = old
    return pl
SENS = []
for cb, iv in ((0, 0.22), (20000, 0.22), (60000, 0.22), (20000, 0.12), (20000, 0.35)):
    pl = sens(cb, iv); tots = [p["total"] for p in pl]; k = int(np.argmin(tots))
    SENS.append([f0(cb), pc(iv, 0), "A. Nivelación" if k == 0 else "B. Persecución" if k == 1 else "C. Mixta", mm(tots[0]), mm(tots[1]), mm(tots[2])])
# =============================================================== PORTADA (formato APA 7 para trabajos estudiantiles)
for _ in range(4): para("", align="c")
para("Consultoría simulada de producción inteligente: diagnóstico y plan de mejora del sistema de producción de cemento gris de Cementos Argos en Colombia", bold=True, align="c")
para("", align="c")
for ln in ("Brandon Afanador, Natalia Aragón, Mary Barrera y María Auxiliadora Salcedo Pérez",
           "Programa de Administración de Empresas, Universidad Sergio Arboleda, Barranquilla",
           "Prospectiva Tecnológica & Producción Inteligente",
           "Natasha Charris",
           "Noviembre de 2026"):
    para(ln, align="c")
d.add_page_break()

# =============================================================== TABLA DE CONTENIDO
d.add_heading("Tabla de contenido", 1)
TOC = ["Introducción y alcance", "Contextualización de la empresa", "Pronóstico de demanda", "Análisis de capacidad",
       "Plan agregado", "Plan Maestro de Producción (MPS) y ATP", "Tablero de indicadores",
       "Recomendaciones", "Conclusiones", "Anexo A. Tabla consolidada de supuestos", "Referencias"]
for t in TOC:
    p = d.add_paragraph(); p.paragraph_format.tab_stops.add_tab_stop(Cm(16.4), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)
    p.paragraph_format.line_spacing = 1.5
    r = p.add_run(t + "\t" + str(PAGES.get(t, "0")))
d.add_page_break()

# =============================================================== 1 INTRO
h1(TOC[0])
para("Este documento presenta una consultoría simulada: nuestro equipo asume el papel de consultores externos que diagnostican el sistema de producción de **Cementos Argos en Colombia** y formulan recomendaciones sustentadas con las herramientas del curso. La empresa se eligió porque es un emisor público que reporta sus resultados, y porque el DANE publica cada mes los despachos nacionales de cemento gris, lo que permite apoyar el pronóstico en datos reales y no solo en supuestos.", align="j")
para("**Alcance.** Se analiza una sola familia (cemento gris) en la operación colombiana, con un horizonte de planeación de 12 meses (diciembre de 2026 a noviembre de 2027). No se incluyen concreto, agregados ni operaciones fuera de Colombia. Se prefirió un análisis sencillo y coherente a uno complejo: una unidad de medida (toneladas), una familia de producto y un flujo de cálculo único.", align="j")
para("**Cómo se conectan los componentes.** Un libro de Excel (anexo) concentra todos los supuestos en una sola hoja y el resto de los cálculos se alimenta de ella. El pronóstico fija la demanda; la demanda se compara con la capacidad por etapa; la capacidad y la demanda alimentan el plan agregado; el plan agregado se desagrega en el MPS; y el tablero mide lo que las secciones anteriores decidieron.", align="j")
fig("fig/f7_flujo.png", 12.5, "Figura 1. Flujo de coherencia entre los componentes del proyecto.")
para("**Transparencia sobre los datos.** Distinguimos tres tipos de cifras:", after=2)
bullets([
 "**Datos públicos:** despachos nacionales anuales del Departamento Administrativo Nacional de Estadística [DANE] (2026), a saber, 12.109,6 kt en 2024 y 12.708,0 kt en 2025), el acumulado enero-julio de 2026 (7.555,2 kt) y los despachos de junio y septiembre de 2026; y las 5,0 millones de toneladas de cemento de la regional Colombia reportadas por Grupo Argos para 2025.",
 "**Datos reconstruidos:** no pudimos descargar la serie mensual oficial del DANE desde nuestro entorno de trabajo. Reconstruimos los 33 meses (ene-24 a sep-26) con una estacionalidad supuesta, calibrada para que los totales coincidan con las cifras públicas anteriores. La hoja 'Historico' del Excel permite reemplazarla por la serie oficial.",
 "**Supuestos:** capacidades por etapa, costos y precios, que no son información pública. Se declaran, se justifican y se usan de forma idéntica en todas las secciones (Anexo A).",
])
para("**Uso de inteligencia artificial.** Se utilizó una herramienta de IA generativa como apoyo para la búsqueda de fuentes, el cálculo y la redacción. Las decisiones y supuestos fueron revisados por el equipo, conforme a la política de integridad académica del curso.", align="j")

# =============================================================== 2 CONTEXTO
h1(TOC[1])
para("**La empresa.** Cementos Argos es una cementera colombiana con operaciones en Colombia, el Caribe y Centroamérica. En 2025 reportó despachos de cemento del orden de 9,3 millones de toneladas (Grupo Argos, 2026) y una capacidad instalada de cemento de 14,4 millones de toneladas considerando todas sus regionales (Argos, 2025); ambas cifras deben contrastarse con el reporte integrado vigente. Para este proyecto solo se analiza la regional Colombia, que Grupo Argos (2026) reporta con 5,0 millones de toneladas de cemento en 2025. Otra fuente de prensa cita 5,4 millones, probablemente por un alcance distinto; usamos 5,0 por provenir del informe oficial y lo marcamos como dato a verificar.", align="j")
para(f"**El mercado.** Según el Departamento Administrativo Nacional de Estadística [DANE] (2026), los despachos nacionales de cemento gris fueron de {f1(S['nac_2025'])} kt en 2025, un {pc(S['nac_2025']/S['nac_2024']-1)} más que en 2024 ({f1(S['nac_2024'])} kt), y el acumulado de enero a julio de 2026 creció 4,3 % frente al mismo periodo de 2025. La participación de Argos en Colombia, calculada por nosotros como 5.000 / {f1(S['nac_2025'])}, es de **{pc(S['share'])}**.", align="j")
para("**Productos y tipo de sistema productivo.** Consideramos tres productos: cemento de uso general (UG) en saco, cemento estructural en saco y cemento a granel (45 %, 20 % y 35 % de las toneladas, supuesto). El sistema es de **proceso continuo** en la etapa de horno (el clínker se produce de forma ininterrumpida y se detiene solo por mantenimiento) y de **producción por lotes o campañas** en la molienda y el ensacado, donde se cambia de tipo de cemento y de presentación.", align="j")
ctx = [["Despachos nacionales de cemento gris 2024 (kt)", f1(S["nac_2024"]), "DANE (2026)"],
       ["Despachos nacionales de cemento gris 2025 (kt)", f1(S["nac_2025"]), "DANE (2026)"],
       ["Crecimiento 2025 frente a 2024", pc(S["nac_2025"]/S["nac_2024"]-1), "Calculado"],
       ["Acumulado ene–jul 2026: 7.555,2 kt frente a 7.245,5 kt (2025)", "+4,3 %", "DANE (2026)"],
       ["Cemento Argos regional Colombia 2025 (kt)", "5.000,0", "Grupo Argos (2026)"],
       ["Participación de mercado de Argos (calculada)", pc(S["share"]), "Calculada"]]
table(["Indicador de contexto", "Valor", "Fuente"], ctx, [9.5, 3, 4], size=9, caption="Tabla 1. Cifras públicas que anclan el análisis.")
para("**Problema de consultoría.** La demanda del cemento es estacional y crece cerca de 5 % al año, mientras que el horno, el activo más costoso y menos flexible, tiene una capacidad fija. La pregunta que guía el proyecto es: **¿cómo debe planear y programar Argos su producción de cemento gris en Colombia durante los próximos 12 meses para cumplir la demanda al menor costo, identificando y protegiendo la etapa que limita el sistema?** Cada componente responde una parte: cuánto se demandará, cuánto se puede producir, cómo repartir la producción en el tiempo, qué referencia programar semana a semana, dónde está el cuello de botella y cómo se medirá el resultado.", align="j")
fig("fig/f8_proceso.png", 12.5, "Figura 2. Proceso productivo simplificado. En rojo, la etapa que el análisis de capacidad identifica como la más cargada (cuello de botella).")
para("**Proceso.** La materia prima (caliza y arcilla) se tritura y se cocina en el horno para producir clínker. El clínker se muela con yeso y adiciones para obtener cemento, que se empaca en sacos o se despacha a granel. Usar más adiciones reduce el contenido de clínker por tonelada (factor clínker), que es el insumo más intensivo en energía y emisiones. Por eso esta variable aparece más adelante como indicador de sostenibilidad (ESG).", align="j")

# =============================================================== 3 PRONÓSTICO
h1(TOC[2])
para(f"**Producto y datos.** Se pronostica la demanda mensual de cemento gris de Argos Colombia, calculada como despachos nacionales multiplicados por la participación de mercado ({pc(S['share'])}). Se asume que la participación se mantiene constante en el horizonte. La serie cubre enero de 2024 a septiembre de 2026 (33 meses; ver nota de reconstrucción en la introducción). No se incluye 2020 ni años anteriores: se evitan así los efectos de la pandemia sin necesidad de ajustes.", align="j")
para("**Métodos comparados.** Se compararon tres métodos vistos en el curso, evaluados sobre la misma ventana (enero de 2025 a septiembre de 2026) con las medidas MAD, MAPE y señal de rastreo (Heizer et al., 2020; Hyndman y Athanasopoulos, 2021):", after=2)
bullets(["Promedio móvil de 3 meses (no captura estacionalidad ni tendencia).", f"Suavización exponencial simple con α = {M.SES_A:.1f} (captura el nivel, no la estacionalidad).", f"Holt-Winters multiplicativo con α = {M.HW_A:.1f}, β = {M.HW_B:.2f} y γ = {M.HW_G:.1f} (nivel, tendencia y estacionalidad de 12 meses), con parámetros elegidos por búsqueda en cuadrícula."])
rows = [[r[0].split(" (")[0], f1(r[1]), pc(r[2] / 100), f"{r[3]:.2f}".replace(".", ",")] for r in M.RES]
table(["Método", "MAD (kt)", "MAPE", "Señal de rastreo"], rows, [7, 3, 3, 3.5], caption="Tabla 2. Medidas de error de los tres métodos (ventana común de evaluación).")
para(f"**Selección.** Holt-Winters tiene el menor error (MAPE {pc(M.MEJOR[2]/100)}) y la única señal de rastreo cercana a cero ({M.MEJOR[3]:.2f}). Los otros dos métodos tienen señales de rastreo de 7,5 y 7,6, muy fuera del rango habitual de ±4: subestiman sistemáticamente una demanda con tendencia y estacionalidad. **Advertencia:** la serie tiene solo unas 2,75 temporadas y el error es *in-sample* (los parámetros se ajustaron con los mismos datos), por lo que el MAPE real a futuro será mayor. Por eso el tablero de indicadores (véase Tablero de indicadores) incluye el MAPE como indicador de seguimiento.".replace(".", ".") , align="j")
para("**Construcción del modelo elegido.** Holt-Winters multiplicativo (Hyndman y Athanasopoulos, 2021) actualiza tres componentes en cada mes t: el nivel L, la tendencia T y un índice estacional S de 12 meses (Winters, 1960). El pronóstico de h meses adelante es (L + h·T) multiplicado por el índice estacional del mes correspondiente. Con γ = 0,7 el índice estacional se adapta rápido, lo que es razonable para una serie con una caída marcada en diciembre y enero, mientras que α = 0,1 y β = 0,10 hacen que nivel y tendencia cambien con lentitud y no reaccionen en exceso al ruido mensual. Los pronósticos de octubre y noviembre de 2026 se calculan, pero quedan fuera del horizonte de planeación (que empieza en diciembre).", align="j")
fig("fig/f1_pronostico.png", 12.5, "Figura 3. Demanda histórica de Argos Colombia, ajuste de Holt-Winters y pronóstico.")
fr = []
for i in range(6):
    fr.append([mes(M.PLAN_LABEL[i]), f1(M.DEM[i]), mes(M.PLAN_LABEL[i+6]), f1(M.DEM[i+6])])
table(["Mes", "Demanda (kt)", "Mes", "Demanda (kt)"], fr, [3, 3.5, 3, 3.5], size=9, caption="Tabla 3. Demanda pronosticada de Argos Colombia en el horizonte de planeación.")
para(f"**Resultado.** La demanda pronosticada para el horizonte de planeación suma **{f1(DEMs)} kt** (promedio de {f1(DEMs/12)} kt/mes), con un mínimo de {f1(M.DEM.min())} kt en {mes(M.PLAN_LABEL[int(M.DEM.argmin())])} y un máximo de {f1(M.DEM.max())} kt en {mes(M.PLAN_LABEL[int(M.DEM.argmax())])}. Esto implica un crecimiento de aproximadamente 5 % anual frente a las 5.000 kt de 2025, coherente con el crecimiento nacional de 4,9 % en 2025 y 4,3 % en 2026. Esta serie es el insumo de los análisis de capacidad, plan agregado y MPS.", align="j")

# =============================================================== 4 CAPACIDAD
h1(TOC[3])
para("**Etapas y supuestos.** Se analizan tres etapas: horno (clínker), molienda y ensacado. Las capacidades de diseño no son información pública y se declaran como supuestos del equipo (Anexo A). Para cada etapa, la capacidad efectiva es la capacidad de diseño multiplicada por un factor de disponibilidad o eficiencia (88 % en el horno por paros programados y no programados, 90 % en molienda y 85 % en ensacado). El requerimiento de cada etapa se obtiene de la demanda pronosticada: el horno necesita 0,75 t de clínker por tonelada de cemento, la molienda procesa el 100 % y el ensacado el 65 % (la fracción en saco).", align="j")
cr = []
for k in M.REQ:
    cr.append([k, f0(M.CAPD[k] * 12), f1(M.CAPD[k]), f1(M.CAPE[k]), f1(M.REQ[k].mean()), f1(M.REQ[k].max()), pc(M.UTIL_AVG[k]), pc(M.UTIL_MAX[k])])
table(["Etapa", "Diseño anual (kt)", "Diseño mensual", "Efectiva mensual", "Requerida media", "Requerida pico", "Utiliz. media", "Utiliz. pico"], cr, [3.3, 2, 2, 2, 2, 2, 1.9, 1.9], size=8.5, caption="Tabla 4. Capacidad diseñada, efectiva y utilización por etapa (kt/mes; dic-26 a nov-27).")
para("**Definiciones.** Se usan las definiciones habituales de capacidad (Heizer et al., 2020). *Capacidad diseñada*: producción máxima teórica de la etapa. *Capacidad efectiva*: la diseñada por la disponibilidad o eficiencia (descuenta paros y pérdidas de ritmo). *Utilización*: requerimiento dividido entre capacidad efectiva; indica qué tan cargada está la etapa y, por encima de 100 %, que la demanda no cabe. *Eficiencia*: relación entre la capacidad efectiva y la diseñada (88 %, 90 % y 85 % en las tres etapas).".replace("*",""), align="j")
fig("fig/f3_util.png", 12.5, "Figura 4. Utilización mensual de la capacidad efectiva por etapa. Sobre la línea roja la demanda excede la capacidad.")
para(f"**Resultado.** El **horno** es la etapa con mayor utilización ({pc(M.UTIL_AVG['Horno (clínker)'])} en promedio) y la única que supera el 100 % en los meses pico ({', '.join(months_over)}, con un máximo de {pc(M.UTIL_MAX['Horno (clínker)'])}). La molienda y el ensacado tienen holgura. En términos de cemento, el sistema puede producir **{f1(M.CAP_CEM_REG)} kt/mes** con recursos regulares, limitado por el horno. Con recursos extra (compra de clínker, molienda adicional) se asume una capacidad máxima de {f1(M.CAP_CEM_EXT)} kt/mes (+8 %). Solo {f1(np.maximum(0, M.DEM - M.CAP_CEM_REG).sum())} kt de cemento (en dos o tres meses) exceden la capacidad regular.", align="j")
para(f"**Estrategia de capacidad.** Recomendamos una estrategia de **ajuste (*match*) con mejoras incrementales**, y no adelantarse con una nueva línea de horno. La brecha en 2027 es pequeña y se concentra en pocos meses, de modo que no justifica una inversión grande; se cubre con inventario previo y con mejoras en el horno. El riesgo es el crecimiento: con 5 % anual, la demanda media alcanzaría la capacidad regular anual ({f0(cap_year)} kt) en aproximadamente {yrs:.1f} años después del horizonte, es decir, hacia 2029. Ese es el punto en que la decisión de capacidad debe replantearse con una evaluación financiera de inversión.", align="j")
para("**Nota metodológica.** El resultado de que el horno sea la etapa más cargada depende de las capacidades supuestas. Es la configuración típica de una cementera (el horno es el activo más caro y menos flexible), pero debe confirmarse con datos internos o con una entrevista a la empresa en una consultoría real.", align="j", italic=True)

# =============================================================== 5 PLAN AGREGADO
h1(TOC[4])
para("**Unidad y costos.** El plan se expresa en kt de cemento por mes. Se comparan tres estrategias, todas sujetas a que la producción no supere la capacidad con recursos extra (528,0 kt/mes) y a que el inventario final no sea inferior al inicial (150 kt, equivalente a un mes de demanda). Los costos relevantes (supuestos del Anexo A) son:", after=2)
bullets([
 "Producción: costo variable de 330.000 COP/t.",
 "Mantener inventario: 22 % anual sobre 400.000 COP/t, es decir, 7.333 COP/t por mes.",
 "Sobrecosto sobre la capacidad regular (clínker comprado y molienda adicional): 60.000 COP/t.",
 "Cambiar el ritmo de producción (arranques y paradas, turnos): 20.000 COP por tonelada de cambio entre meses.",
 "Faltante: se pierde el ingreso completo (560.000 COP/t); el costo de producir esa tonelada no se incurre.",
])
para("**Estrategias.** Las tres estrategias son las clásicas de la planeación agregada (Heizer et al., 2020). (A) **Nivelación:** producir un volumen constante igual a la demanda media y absorber la estacionalidad con inventario. (B) **Persecución:** producir cada mes lo que se demanda. (C) **Mixta:** dos niveles de producción, uno hasta cierto mes y otro después, elegidos por optimización simple para minimizar el costo total.", align="j")
fig("fig/f4_agregado.png", 13.5, "Figura 5. Producción de cada estrategia frente a la demanda (izquierda) y ahorro frente a la nivelación (derecha).")
def costrow(p): return [p["nombre"], mm(p["prod"]), mm(p["extra"]), mm(p["inv"]), mm(p["cambio"]), mm(p["falt"]), mm(p["total"])]
table(["Estrategia", "Producción", "Sobrecosto", "Inventario", "Cambio de ritmo", "Faltante", "TOTAL"], [costrow(p) for p in M.PLANES], [4.6, 2.2, 1.9, 1.9, 2.2, 1.7, 2.3], size=8.5, caption="Tabla 5. Costos relevantes por estrategia (millones de COP, horizonte de 12 meses).")
pm = [f0(x) for x in P]
para(f"**Resultado.** La estrategia **mixta** es la de menor costo ({mm(Cc['total'])} millones de COP): produce {f1(P[0])} kt/mes de {mes(M.PLAN_LABEL[0])} a {mes(M.PLAN_LABEL[3])} y {f1(P[4])} kt/mes de {mes(M.PLAN_LABEL[4])} a {mes(M.PLAN_LABEL[11])}. Cumple toda la demanda (fill rate de 100 %) sin comprar clínker, porque mantiene la producción dentro de la capacidad regular y cubre los picos de septiembre a noviembre con el inventario acumulado (máximo de {f0(inv_max)} kt en {mes(inv_max_m)}, que representa solo {inv_max/M.DEM[M.PLAN_LABEL.index(inv_max_m)]:.1f} meses de demanda, muy por debajo del límite de 1,5 meses por vida útil del cemento en saco). Ahorra {f0(ahorro_AC)} millones de COP frente a la nivelación y {f0(ahorro_BC)} millones frente a la persecución. La persecución obliga a mover el horno cada mes, y la nivelación acumula más inventario.", align="j")
pr = []
for r_ in rC:
    clk = r_["prod"] * S["factor_clinker"]
    pr.append([mes(r_["mes"]), f1(r_["dem"]), f1(r_["prod"]), f1(r_["inv"]), f1(r_["falt"]), pc(clk / M.CLK_CAP)])
table(["Mes", "Demanda (kt)", "Producción (kt)", "Inventario final (kt)", "Faltante (kt)", "Utilización del horno"], pr, [2.2, 2.8, 3, 3.2, 2.5, 3], size=8.5, caption="Tabla 6. Plan agregado elegido (estrategia mixta), mes a mes.")
para("**Sensibilidad.** Como los costos de cambiar el ritmo y de mantener inventario son supuestos, se recalculó el plan con otros valores para ver si la recomendación cambia:", align="j", after=2)
table(["Costo de cambio (COP/t)", "Inventario (% anual)", "Mejor estrategia", "A (MM COP)", "B (MM COP)", "C (MM COP)"], SENS, [3.4, 2.8, 3.2, 2.7, 2.7, 2.7], size=8.5, caption="Tabla 7. Sensibilidad del ranking de estrategias.")
para("La estrategia mixta es la de menor costo en los cinco escenarios probados. Incluso si cambiar el ritmo del horno no costara nada, la mixta sigue siendo la mejor (la ventaja sobre la persecución se reduce a unos 1.800 millones de COP, frente a unos 3.900 en el caso base) porque, además, mantiene menos inventario. Cuando el costo de cambiar el ritmo es alto (60.000 COP/t), la persecución pasa a ser la peor opción, incluso peor que la nivelación. El costo de mantener inventario cambia el nivel de los costos, pero no el orden: con 12 % anual, nivelación y persecución quedan casi empatadas, y con 35 % la mixta amplía su ventaja. La recomendación es, por tanto, robusta frente a estos dos supuestos.", align="j")
para("**Lectura crítica.** Las tres estrategias cuestan casi lo mismo (la diferencia es menor al 0,5 %), porque el costo de producción domina y la demanda cabe en la capacidad. El ahorro de la mixta, aunque pequeño en términos porcentuales, equivale a miles de millones de pesos al año y, sobre todo, reduce la dependencia de comprar clínker en los picos. La ventaja de la mixta se mantiene al variar los costos de cambio de ritmo y de inventario (ver la sensibilidad más abajo), aunque su tamaño sí depende de ellos.", align="j")

# =============================================================== 6 MPS
h1(TOC[5])
para(f"**Referencia y datos.** Se elaboró el MPS de la referencia clave **{M.REF}**, que representa el {pc(M.REF_SHARE, 0)} de la producción. Cubre 12 semanas (tres meses de cuatro semanas: dic-26, ene-27 y feb-27), con campañas de empaque quincenales (tamaño de lote = mitad de la producción mensual de la referencia). La producción mensual proviene del plan agregado elegido; el pronóstico semanal es la demanda pronosticada de la referencia dividida en cuatro semanas; los pedidos firmes son **simulados** (declarados como tales) y decrecen hacia el futuro, como ocurre en la práctica. El inventario inicial de la referencia es {f0(M.INV0)} t (45 % de las 150 kt iniciales).", align="j")
hd = ["Semana", "Pronóstico (t)", "Pedidos firmes (t)", "MPS (t)", "Inv. proyectado (t)", "ATP (t)"]
rows = []
for i, r in enumerate(M.MPS_ROWS):
    rows.append([r[0], f0(r[1]), f0(r[2]), f0(r[3]) if r[3] else "—", f0(r[4]), f0(M.ATP[i]) if M.ATP[i] is not None else "—"])
table(hd, rows, [1.8, 2.8, 3, 2.6, 3.2, 2.4], size=8.5, caption=f"Tabla 8. MPS simplificado y ATP de {M.REF}.")
para("**Cálculos.** Se sigue la lógica estándar de la planeación maestra de producción (Vollmann et al., 2005). El inventario proyectado de cada semana es el inventario anterior más el MPS menos el mayor entre el pronóstico y los pedidos firmes. El **ATP** (disponible para prometer) se calcula en cada semana con campaña: la primera campaña es el inventario inicial más el MPS menos los pedidos firmes hasta la siguiente campaña; las demás son el MPS menos los pedidos firmes hasta la siguiente. El ATP indica cuánto se puede prometer a nuevos clientes sin alterar los pedidos ya comprometidos.", align="j")
fig("fig/f5_mps.png", 12, "Figura 6. MPS, pedidos firmes e inventario proyectado de la referencia.")
para(f"**Coherencia con el plan agregado.** La suma del MPS de cada mes es igual a la producción del plan agregado multiplicada por la participación de la referencia: {f0(P[0]*1000*M.REF_SHARE)} t en diciembre, enero y febrero (celda de control en el Excel). El ATP de la tercera semana es bajo ({f0(M.ATP[2])} t): casi toda la campaña de la semana 3 ya está comprometida, por lo que un pedido nuevo grande debería ofrecerse con la campaña de la semana 5, que tiene un ATP de {f0(M.ATP[4])} t.", align="j")

# =============================================================== 8 TABLERO
h1(TOC[6])
para("El tablero (anexo en Excel: hoja 'Tablero') contiene cinco indicadores, cada uno derivado de una sección anterior, con metas y umbrales de semáforo. Los valores 'reales' de los 12 meses son **simulados**: se generaron como el plan o pronóstico multiplicado por una variación aleatoria con semilla fija (desviaciones de 3,5 % en la demanda, 2 % en la producción y 1,2 puntos en la disponibilidad). El tablero incluye un selector de mes (celda amarilla) que actualiza los valores y el semáforo.", align="j")
ind = [
 ["MAPE del pronóstico", "≤ 5 %", "≤ 8 %", "Pronóstico", pc(kpi["MAPE del pronóstico"][0]), kpi["MAPE del pronóstico"][1]],
 ["Utilización del horno", "85–100 %", "80–100 %", "Capacidad", pc(kpi["Utilización del horno"][0]), kpi["Utilización del horno"][1]],
 ["Cumplimiento del MPS (producido / plan)", "≥ 97 %", "≥ 93 %", "Plan agregado y MPS", pc(kpi["Cumplimiento del MPS"][0]), kpi["Cumplimiento del MPS"][1]],
 ["Fill rate (nivel de servicio)", "≥ 98 %", "≥ 95 %", "MPS", pc(kpi["Fill rate (nivel de servicio)"][0]), kpi["Fill rate (nivel de servicio)"][1]],
 ["Factor clínker (ESG)", "≤ 0,74", "≤ 0,75", "Contexto y capacidad", f"{kpi['Factor clínker (ESG)'][0]:.3f}".replace(".", ","), kpi["Factor clínker (ESG)"][1]],
]
table(["Indicador", "Meta (verde)", "Amarillo", "Origen", "Promedio simulado", "Estado"], ind, [5.2, 2.2, 2.2, 2.8, 2.6, 2], size=8.5, caption="Tabla 9. Indicadores, metas y resultado simulado (promedio de 12 meses).")
fig("fig/f9_tablero.png", 12.5, "Figura 7. Evolución de los cinco indicadores en los 12 meses simulados (líneas verdes: metas).")
para("**Justificación de las metas.** El MAPE de 5 % es exigente pero coherente con el error de la ventana de evaluación. La utilización mínima de 85 % refleja que el plan produce por debajo de la capacidad en los meses de baja demanda; superar 100 % indica sobreutilización y riesgo de paros no programados. El factor clínker de 0,74 es una meta de reducción gradual frente a la base de 0,75, y es el indicador ESG: bajarlo reduce emisiones de CO₂ y libera capacidad del horno.", align="j")

# =============================================================== 9 RECOMENDACIONES
h1(TOC[7])
para("Las recomendaciones se priorizan por su valor financiero estimado con el modelo y por su efecto en sostenibilidad. Cada una indica el hallazgo del que sale, la acción y el indicador de seguimiento. Los valores no son aditivos (los escenarios se evalúan por separado sobre el mismo plan) y no incluyen el costo de implementarlas, que no fue estimado por falta de datos.", align="j")
rec = [
 ["1", "Adoptar el plan agregado mixto", f"Plan agregado: es la estrategia de menor costo y no requiere comprar clínker.", f"Producir {f1(P[0])} kt/mes hasta marzo y {f1(P[4])} kt/mes de abril a noviembre, acumulando inventario previo al pico.", f"{f0(ahorro_AC)} (vs. nivelación) y {f0(ahorro_BC)} (vs. persecución)", "Cumplimiento MPS; fill rate"],
 ["2", "Reducir el factor clínker de 0,75 a 0,72 (ESG)", f"Capacidad: el horno es la etapa más cargada y el clínker es su insumo escaso y de mayores emisiones.", f"Aumentar el uso de adiciones, verificando que los productos cumplan sus normas técnicas.", f"{f0(esc_fc)} en el plan; {f0(M.lib_clk)} kt de clínker y unas {f0(M.CO2_AHORRO)} kt de CO₂ evitados al año", "Factor clínker"],
 ["3", "Aumentar la disponibilidad del horno en 2 puntos porcentuales", "Capacidad: el horno supera el 100 % de utilización en el pico.", "Mantenimiento predictivo con sensores y analítica; programar los paros anuales en los meses de menor demanda (dic–mar).", f"{f0(esc_disp)} en el plan; {f0(M.cem_extra_anual)} kt/año de capacidad", "Utilización y disponibilidad del horno"],
]
table(["N.º", "Recomendación", "Hallazgo", "Acción", "Impacto estimado (MM COP)", "Indicador"], rec, [0.8, 3.2, 3.6, 4.2, 3.2, 2.2], size=8, caption="Tabla 10. Recomendaciones priorizadas.")
para("**Industria 4.0/5.0 y sostenibilidad.** Las recomendaciones 2 y 3 son las que conectan con el enfoque del curso. El mantenimiento predictivo del horno (sensores de vibración, temperatura y consumo energético, con analítica de datos) es una aplicación directa de la Industria 4.0: reemplaza paros no programados por intervenciones planeadas, lo que mueve la disponibilidad hacia la meta. La reducción del factor clínker es la palanca ESG con efecto operativo: cada tonelada de clínker evitada reduce el consumo de combustible y las emisiones de proceso, y libera tiempo de horno, que es la etapa más cargada. El enfoque 5.0 se refleja en que las decisiones se apoyan en el criterio del equipo de planeación (por ejemplo, el ajuste del plan cuando cambia el pronóstico) con datos del tablero, y no se delegan a un algoritmo opaco.", align="j")
para("**Criterio de priorización.** La recomendación 1 es la de mayor valor y la más inmediata (es una decisión de programación, sin inversión). La 2 combina ahorro y beneficio ESG, pero requiere validar normas y aceptación del mercado. La 3 requiere inversión en sensores, cuyo retorno debe evaluarse.", align="j")
imp = [
 ["Diciembre de 2026", "Aprobar el plan mixto y fijar la producción de dic–mar (404,0 kt/mes)", "Planeación de operaciones", "Cumplimiento MPS"],
 ["Dic. 2026 – mar. 2027", "Programar el paro anual del horno en el mes de menor demanda; iniciar piloto de mantenimiento predictivo", "Mantenimiento", "Disponibilidad del horno"],
 ["Abril de 2027", "Subir producción a 488,0 kt/mes para acumular inventario previo al pico", "Planeación y comercial", "Fill rate; utilización del horno"],
 ["Abril – septiembre 2027", "Pruebas de adiciones para llevar el factor clínker de 0,75 a 0,72", "Calidad y producción", "Factor clínker"],
 ["Octubre – noviembre 2027", "Operar el pico con el inventario acumulado; medir el resultado frente al modelo", "Planeación", "Todos los indicadores"],
]
table(["Cuándo", "Acción", "Responsable (rol)", "Indicador"], imp, [3.2, 7.2, 3.2, 3.2], size=8.5, caption="Tabla 11. Hoja de ruta de implementación (horizonte dic-26 a nov-27).")
rk = [
 ["La demanda real supera el pronóstico en el pico", "El inventario no alcanza y aparecen faltantes", "Revisar el pronóstico cada mes (MAPE y señal de rastreo); activar compra de clínker (sobrecosto de 60.000 COP/t)"],
 ["La disponibilidad del horno es menor a 88 %", "Se pierde la capacidad del pico", "Reforzar el inventario previo y activar la compra de clínker en el pico"],
 ["Las adiciones no cumplen la norma o la aceptación del mercado", "No se logra el factor clínker de 0,72", "Pilotear por producto y por planta antes de generalizar"],
 ["Los supuestos de capacidad y costos difieren de la realidad", "Cambia el ranking de estrategias o la etapa más cargada", "Reemplazar los supuestos del Excel por datos internos y recalcular"],
]
table(["Riesgo", "Consecuencia", "Respuesta"], rk, [5.5, 4.5, 7], size=8.5, caption="Tabla 12. Riesgos principales y respuesta.")
para("**Limitaciones.** (i) La serie mensual histórica está reconstruida y debe sustituirse por la oficial del DANE; (ii) las capacidades, precios y costos son supuestos; (iii) la participación de mercado se asume constante; (iv) no se evaluó el costo de implementar las recomendaciones; (v) los datos del tablero son simulados.", align="j")

# =============================================================== 10 CONCLUSIONES
h1(TOC[8])
bullets([
 f"La demanda de Argos Colombia crecería cerca de 5 % anual y suma {f0(DEMs)} kt en el horizonte; Holt-Winters fue el método más preciso (MAPE {pc(M.MEJOR[2]/100)} en la ventana de evaluación), con la salvedad de que la serie es corta y reconstruida.",
 f"El horno de clínker es la etapa más cargada ({pc(M.UTIL_AVG['Horno (clínker)'])} de utilización media) y la única que supera el 100 % en los meses pico; en los meses de baja demanda el límite es el mercado.",
 f"Una estrategia mixta de dos niveles de producción minimiza el costo ({mm(Cc['total'])} millones de COP), cumple el 100 % de la demanda y evita comprar clínker; el ahorro frente a las alternativas es de {f0(ahorro_BC)} a {f0(ahorro_AC)} millones de COP.",
 "El MPS de la referencia UG en saco es coherente con el plan agregado y el ATP muestra cuánta capacidad queda para nuevos pedidos en cada campaña.",
 "Reducir el factor clínker conecta sostenibilidad y productividad: libera capacidad en el horno y reduce emisiones.",
 "Antes de usar estos resultados en una decisión real, se deben reemplazar los supuestos por datos internos de la empresa; el modelo de Excel permite hacerlo sin rehacer los cálculos.",
])

# =============================================================== ANEXO A
h1(TOC[9], newpage=True)
sup = [
 ["Despachos nacionales 2024 / 2025", "12.109,6 / 12.708,0 kt", "Dato DANE", "Sí"], ["Acumulado ene–jul 2026; jun; sep", "7.555,2; 1.061,6; 1.194,4 kt", "Dato DANE", "Sí"],
 ["Cemento Argos Colombia 2025", "5.000 kt", "Grupo Argos (2026)", "Sí"], ["Participación de mercado", pc(S["share"]), "Calculada", "Sí"],
 ["Capacidad de diseño del horno", "5.000 kt clínker/año", "Supuesto", "Sí"], ["Disponibilidad del horno", "88 %", "Supuesto (rango 85–90 %)", "Sí"],
 ["Factor clínker", "0,75 t/t (0,746 ponderado)", "Supuesto", "Sí"], ["Capacidad de molienda / eficiencia", "6.800 kt/año / 90 %", "Supuesto", "Sí"],
 ["Capacidad de ensacado / eficiencia", "4.800 kt/año / 85 %", "Supuesto", "Declarar"], ["Mezcla UG / estructural / granel", "45 % / 20 % / 35 %", "Supuesto", "Declarar"], ["Precio y costo variable promedio", "560.000 / 330.000 COP/t", "Supuesto", "Sí"],
 ["Costo de mantener inventario", "22 % anual sobre 400.000 COP/t", "Supuesto (rango 20–25 %)", "No"], ["Sobrecosto sobre capacidad regular", "60.000 COP/t", "Supuesto", "Declarar"],
 ["Costo de cambiar el ritmo", "20.000 COP/t de cambio", "Supuesto", "Declarar"], ["Capacidad extra / inventario inicial", "+8 % / 150 kt", "Supuesto", "Declarar"],
 ["Vida útil del cemento en saco", "1,5 meses de demanda", "Supuesto", "Sí"], ["t CO₂ por t de clínker", "0,85", "Referencia sectorial", "Sí"],
 ["Serie mensual histórica", "Reconstruida (calibrada con totales DANE)", "Supuesto", "Reemplazar por ECG oficial"],
 ["Pedidos firmes del MPS y datos del tablero", "Simulados", "Simulación", "—"],
]
table(["Supuesto o dato", "Valor", "Tipo / fuente", "Verificar"], sup, [5.8, 4.8, 4.2, 2.6], size=8.5, caption="Tabla A1. Supuestos consolidados (la columna 'Verificar' indica qué debe contrastarse con fuentes antes de la entrega).")

# =============================================================== REFERENCIAS
h1(TOC[10], newpage=True)
refs = [
 "Argos. (2025). *Reporte integrado Cementos Argos 2024*. https://files.grupoargos.com/uploads-grupo-argos/2025/03/Reporte-Integrado-cementos-argos-2024.pdf",
 "Departamento Administrativo Nacional de Estadística. (2026). *Estadísticas de cemento gris (ECG): boletines técnicos de diciembre de 2025, julio de 2026 y septiembre de 2026*. https://www.dane.gov.co/index.php/en/statistics-by-topic-1/construction/estadisticas-de-cemento-gris",
 "Grupo Argos. (2026). *Informe de gestión Grupo Argos 2025*. https://files.grupoargos.com/uploads-grupo-argos/2026/03/InformeGestion_GrupoArgos_2025.pdf",
 "Heizer, J., Render, B., & Munson, C. (2020). *Operations management: Sustainability and supply chain management* (13.ª ed.). Pearson.",
 "Hyndman, R. J., & Athanasopoulos, G. (2021). *Forecasting: Principles and practice* (3.ª ed.). OTexts. https://otexts.com/fpp3/",
 "Vollmann, T. E., Berry, W. L., Whybark, D. C., & Jacobs, F. R. (2005). *Manufacturing planning and control for supply chain management* (5.ª ed.). McGraw-Hill.",
 "Winters, P. R. (1960). Forecasting sales by exponentially weighted moving averages. *Management Science, 6*(3), 324–342. https://doi.org/10.1287/mnsc.6.3.324",
]
for r in refs:
    p = d.add_paragraph(); p.paragraph_format.left_indent = Cm(1.27); p.paragraph_format.first_line_indent = Cm(-1.27)
    parts = r.split("*")
    for i, t in enumerate(parts):
        if t: run = p.add_run(t); run.italic = (i % 2 == 1)

d.save("Documento_Proyecto_Final_Argos.docx")
print("docx ok")
