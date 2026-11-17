import os
os.chdir("/home/user/mauxysalcedo.github.io/proyecto_final")
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
import numpy as np, modelo as M
from openpyxl import load_workbook
S=M.S
NAVY=RGBColor(0x0B,0x1F,0x3A); GOLD=RGBColor(0xF2,0xB7,0x05); WHITE=RGBColor(255,255,255); GRAY=RGBColor(0x55,0x5F,0x70); LIGHT=RGBColor(0xEE,0xF1,0xF6); RED=RGBColor(0xC0,0x39,0x2B); GREEN=RGBColor(0x2E,0x8B,0x57)
def f0(x): return f"{x:,.0f}".replace(",",".")
def f1(x): return f"{x:,.1f}".replace(",","X").replace(".",",").replace("X",".")
def pc(x,d=1): return f"{x*100:.{d}f} %".replace(".",",")
mm=lambda x:f0(x/1e6)
prs=Presentation(); prs.slide_width=Inches(13.333); prs.slide_height=Inches(7.5)
blank=prs.slide_layouts[6]
A,B,C=M.PLANES; P=M.MEJOR_PLAN["P"]
ES=["ene","feb","mar","abr","may","jun","jul","ago","sep","oct","nov","dic"]; mes=lambda d:f"{ES[d.month-1]}-{str(d.year)[2:]}"
months_over=[mes(M.PLAN_LABEL[i]) for i in range(12) if M.CLK_REQ[i]>M.CLK_CAP]
clk_def=np.maximum(0,M.CLK_REQ-M.CLK_CAP).sum()
esc_disp=(M.ESC_BASE["total"]-M.ESC_DISP["total"])/1e6; esc_fc=(M.ESC_BASE["total"]-M.ESC_FC["total"])/1e6
tb=load_workbook("Proyecto_Argos_Modelo_y_Tablero.xlsx",data_only=True)["Tablero"]
kpi={tb.cell(row=r,column=1).value:(tb.cell(row=r,column=3).value,tb.cell(row=r,column=6).value) for r in range(22,27)}
num=[0]
def slide(titulo, notas="", seccion=""):
    s=prs.slides.add_slide(blank); num[0]+=1
    bar=s.shapes.add_shape(MSO_SHAPE.RECTANGLE,0,0,prs.slide_width,Inches(1.05)); bar.fill.solid(); bar.fill.fore_color.rgb=NAVY; bar.line.fill.background()
    acc=s.shapes.add_shape(MSO_SHAPE.RECTANGLE,0,Inches(1.05),prs.slide_width,Inches(0.07)); acc.fill.solid(); acc.fill.fore_color.rgb=GOLD; acc.line.fill.background()
    tf=s.shapes.add_textbox(Inches(0.6),Inches(0.18),Inches(12),Inches(0.8)).text_frame; tf.word_wrap=True
    r=tf.paragraphs[0].add_run(); r.text=titulo; r.font.size=Pt(28); r.font.bold=True; r.font.color.rgb=WHITE; r.font.name="Georgia"
    ft=s.shapes.add_textbox(Inches(0.6),Inches(7.05),Inches(12),Inches(0.35)).text_frame
    r=ft.paragraphs[0].add_run(); r.text=f"Cementos Argos · Proyecto Final · {num[0]}"; r.font.size=Pt(10); r.font.color.rgb=GRAY
    if notas: s.notes_slide.notes_text_frame.text=notas
    return s
def text(s,x,y,w,h,items,size=18,bullet=True,color=NAVY,bold_first=False):
    tf=s.shapes.add_textbox(Inches(x),Inches(y),Inches(w),Inches(h)).text_frame; tf.word_wrap=True
    for i,t in enumerate(items):
        p=tf.paragraphs[0] if i==0 else tf.add_paragraph(); p.space_after=Pt(8)
        parts=(("• " if bullet else "")+t).split("**")
        for j,pt in enumerate(parts):
            if not pt: continue
            r=p.add_run(); r.text=pt; r.font.size=Pt(size); r.font.bold=(j%2==1); r.font.color.rgb=color; r.font.name="Calibri"
    return tf
def img(s,path,x,y,w=None,h=None):
    return s.shapes.add_picture(path,Inches(x),Inches(y),width=Inches(w) if w else None,height=Inches(h) if h else None)
def tabla(s,x,y,w,headers,rows,colw=None,size=12,rowh=0.36):
    sh=s.shapes.add_table(len(rows)+1,len(headers),Inches(x),Inches(y),Inches(w),Inches(rowh*(len(rows)+1))); t=sh.table
    for i,hh in enumerate(headers):
        c=t.cell(0,i); c.text=hh; c.fill.solid(); c.fill.fore_color.rgb=NAVY
        for p in c.text_frame.paragraphs:
            p.alignment=PP_ALIGN.CENTER
            for r in p.runs: r.font.size=Pt(size); r.font.bold=True; r.font.color.rgb=WHITE
    for k,row in enumerate(rows):
        for i,v in enumerate(row):
            c=t.cell(k+1,i); c.text=str(v); c.fill.solid(); c.fill.fore_color.rgb=LIGHT if k%2 else WHITE
            for p in c.text_frame.paragraphs:
                p.alignment=PP_ALIGN.LEFT if i==0 else PP_ALIGN.CENTER
                for r in p.runs: r.font.size=Pt(size); r.font.color.rgb=NAVY
    if colw:
        for i,cw in enumerate(colw): t.columns[i].width=Inches(cw)
    return t
def callout(s,x,y,w,h,txt,fill=GOLD,color=NAVY,size=16):
    b=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Inches(x),Inches(y),Inches(w),Inches(h)); b.fill.solid(); b.fill.fore_color.rgb=fill; b.line.fill.background()
    tf=b.text_frame; tf.word_wrap=True; tf.vertical_anchor=MSO_ANCHOR.MIDDLE
    parts=txt.split("**")
    p=tf.paragraphs[0]
    for j,pt in enumerate(parts):
        if not pt: continue
        r=p.add_run(); r.text=pt; r.font.size=Pt(size); r.font.bold=(j%2==1); r.font.color.rgb=color; r.font.name="Calibri"
    return b

# 1 portada
s=prs.slides.add_slide(blank); num[0]+=1
bg=s.shapes.add_shape(MSO_SHAPE.RECTANGLE,0,0,prs.slide_width,prs.slide_height); bg.fill.solid(); bg.fill.fore_color.rgb=NAVY; bg.line.fill.background()
g=s.shapes.add_shape(MSO_SHAPE.RECTANGLE,Inches(0.8),Inches(3.55),Inches(2.2),Inches(0.08)); g.fill.solid(); g.fill.fore_color.rgb=GOLD; g.line.fill.background()
text(s,0.8,1.2,11.5,1.2,["Consultoría de producción inteligente"],size=22,bullet=False,color=GOLD)
text(s,0.8,1.9,11.7,1.6,["Cementos Argos: cómo planear y programar la producción de cemento gris en Colombia"],size=38,bullet=False,color=WHITE)
text(s,0.8,3.85,11.5,2.6,["Prospectiva Tecnológica & Producción Inteligente · 2026-03","Universidad Sergio Arboleda — Barranquilla","Docente: Natasha Charris","Equipo: Brandon Afanador, Natalia Aragón, Mary Barrera y María Auxiliadora Salcedo Pérez"],size=18,bullet=False,color=RGBColor(0xDD,0xE3,0xEE))
s.notes_slide.notes_text_frame.text="Integrante 1 (30 s): saludo, equipo y qué haremos: diagnóstico de la producción de cemento gris de Argos Colombia y recomendaciones."

# 2 agenda
s=slide("Agenda","Integrante 1 (30 s).")
text(s,0.8,1.5,6,5,["**1.** La empresa y el problema","**2.** Pronóstico de demanda","**3.** Capacidad","**4.** Plan agregado","**5.** MPS y ATP","**6.** Tablero de indicadores","**7.** Recomendaciones y conclusiones"],size=24,bullet=False)
callout(s,7.6,2.0,5,3.2,"**Hilo conductor:** cada análisis alimenta al siguiente y todos apuntan a una decisión: cómo producir al menor costo sin dejar de cumplir la demanda.",size=20)

# 3 empresa
s=slide("La empresa: Argos en Colombia","Integrante 1 (1 min). Datos públicos: DANE y Grupo Argos. La participación es calculada por nosotros.")
text(s,0.6,1.4,6.4,5.4,["**Cementera colombiana** con operaciones en Colombia, el Caribe y Centroamérica","**Alcance:** cemento gris, regional Colombia, 12 meses (dic-26 a nov-27)","**Sistema productivo:** proceso continuo en el horno de clínker; lotes o campañas en molienda y ensacado","**Productos (supuesto):** uso general en saco 45 %, estructural en saco 20 %, granel 35 %"],size=19)
tabla(s,7.2,1.6,5.6,["Dato público","Valor"],[["Despachos nacionales 2024",f1(S["nac_2024"])+" kt"],["Despachos nacionales 2025",f1(S["nac_2025"])+" kt"],["Crecimiento 2025",pc(S["nac_2025"]/S["nac_2024"]-1)],["Argos Colombia 2025","5.000 kt"],["Participación (calculada)",pc(S["share"])]],[3.6,2.0],size=14,rowh=0.5)

# 4 problema
s=slide("El problema de consultoría","Integrante 1. Explicar la pregunta guía y mostrar el proceso.")
img(s,"fig/f8_proceso.png",0.6,1.4,w=12.1)
callout(s,0.9,3.9,11.5,1.5,"**¿Cómo debe planear y programar Argos su producción de cemento gris durante los próximos 12 meses para cumplir la demanda al menor costo, identificando la etapa que más limita el sistema?**",size=20)
text(s,0.9,5.7,11.5,1.2,["La demanda es estacional y crece cerca de 5 % al año; el horno, el activo más caro y menos flexible, tiene capacidad fija."],size=17,bullet=False,color=GRAY)

# 5 coherencia
s=slide("Cómo se conectan los componentes","Integrante 1 (1 min). Resaltar que todo sale de una hoja de supuestos del Excel.")
img(s,"fig/f7_flujo.png",0.5,1.6,w=12.3)
text(s,0.8,3.9,11.8,2.8,["**Una sola hoja de supuestos** alimenta todo el modelo en Excel: si cambia un supuesto, se recalcula todo.","**Datos públicos** (DANE, Grupo Argos) anclan la demanda; los **supuestos** (capacidades, costos) están declarados y justificados.","**Transparencia:** la serie mensual histórica fue reconstruida con totales públicos; los datos del tablero son simulados."],size=19)

# 6 pronóstico 1
s=slide("Pronóstico (1/2): elegimos Holt-Winters","Integrante 2 (1 min). Tres métodos, misma ventana. Señal de rastreo de los otros dos fuera de ±4.")
img(s,"fig/f2_mape.png",0.5,1.4,w=6.2)
tabla(s,7.0,1.6,5.8,["Método","MAD (kt)","MAPE","Señal"],[[["Promedio móvil 3 m","Suav. exp. simple","Holt-Winters"][i],f1(r[1]),pc(r[2]/100),f"{r[3]:.2f}".replace(".",",")] for i,r in enumerate(M.RES)],[2.9,0.9,0.9,1.1],size=12,rowh=0.55)
callout(s,0.6,5.0,12.1,1.5,f"**Holt-Winters:** MAPE {pc(M.MEJOR[2]/100)} y señal de rastreo ≈ 0. Los otros dos métodos subestiman sistemáticamente (señal ≈ 7,5). Salvedad: serie corta y error in-sample; por eso el MAPE entra al tablero.",size=17)

# 7 pronóstico 2
s=slide("Pronóstico (2/2): demanda de dic-26 a nov-27","Integrante 2 (1 min).")
img(s,"fig/f1_pronostico.png",0.5,1.35,w=8.3)
text(s,9.0,1.6,3.9,5,[f"**{f0(M.DEM.sum())} kt** en 12 meses",f"Mínimo: {f1(M.DEM.min())} kt ({mes(M.PLAN_LABEL[int(M.DEM.argmin())])})",f"Máximo: {f1(M.DEM.max())} kt ({mes(M.PLAN_LABEL[int(M.DEM.argmax())])})","Crecimiento ≈ 5 % anual, coherente con el DANE (4,9 % en 2025; 4,3 % ene–jul 2026)"],size=17)

# 8 capacidad 1
s=slide("Capacidad (1/2): el horno es la etapa más cargada","Integrante 3 (1 min). Capacidades diseñadas son supuestos declarados.")
rows=[[k,f1(M.CAPD[k]),f1(M.CAPE[k]),f1(M.REQ[k].mean()),pc(M.UTIL_AVG[k]),pc(M.UTIL_MAX[k])] for k in M.REQ]
tabla(s,0.6,1.5,12.1,["Etapa (kt/mes)","Diseñada","Efectiva","Requerida media","Utilización media","Utilización pico"],rows,[3.4,1.8,1.8,1.9,1.8,1.4],size=15,rowh=0.6)
callout(s,0.6,4.2,12.1,1.3,f"**Etapa más cargada: horno de clínker** ({pc(M.UTIL_AVG['Horno (clínker)'])} de utilización media; supera 100 % en el pico). Capacidad regular del sistema: **{f1(M.CAP_CEM_REG)} kt/mes** de cemento.",size=19)
text(s,0.6,5.7,12.1,1.2,["Efectiva = diseñada × disponibilidad o eficiencia (88 % horno, 90 % molienda, 85 % ensacado). Supuestos del equipo, a validar con datos internos."],size=14,bullet=False,color=GRAY)

# 9 capacidad 2
s=slide("Capacidad (2/2): estrategia de ajuste","Integrante 3 (1 min).")
img(s,"fig/f3_util.png",0.5,1.4,w=7.6)
text(s,8.3,1.5,4.6,5.3,[f"Solo {f1(np.maximum(0,M.DEM-M.CAP_CEM_REG).sum())} kt de cemento exceden la capacidad regular","**Recomendación:** estrategia de **ajuste (match)**, no adelantarse con un nuevo horno","La brecha es pequeña y concentrada: se cubre con inventario previo y mejoras en el horno","**Alerta:** con +5 % anual, la demanda alcanzaría la capacidad regular hacia 2029"],size=17)

# 10 plan agregado 1
s=slide("Plan agregado (1/2): tres estrategias","Integrante 4 (1 min). Mostrar la lógica de cada estrategia.")
img(s,"fig/f4_agregado.png",0.4,1.35,w=8.4)
text(s,9.0,1.5,4.0,5.3,["**A. Nivelación:** producción constante; la estacionalidad se absorbe con inventario","**B. Persecución:** producir lo que se demanda cada mes","**C. Mixta:** dos niveles de producción, optimizados","Costos: producción, inventario (22 % anual), cambio de ritmo, sobrecosto sobre capacidad regular y faltantes"],size=16)

# 11 plan agregado 2
s=slide("Plan agregado (2/2): la estrategia mixta gana","Integrante 4 (1 min). Sensibilidad: gana en los 5 escenarios probados; la diferencia relativa es pequeña (<0,5 %), pero equivale a miles de millones.")
def crow(p): return [p["nombre"].split(" (")[0],mm(p["prod"]),mm(p["inv"]),mm(p["cambio"]),mm(p["falt"]),mm(p["total"])]
tabla(s,0.6,1.5,12.1,["Estrategia (MM COP)","Producción","Inventario","Cambio ritmo","Faltante","TOTAL"],[crow(p) for p in M.PLANES],[3.6,1.9,1.7,1.7,1.4,1.8],size=14,rowh=0.55)
callout(s,0.6,3.9,5.9,2.6,f"**Mixta:** {f1(P[0])} kt/mes hasta mar-27 y {f1(P[4])} kt/mes de abr a nov. Cumple el 100 % de la demanda sin comprar clínker.",size=18)
callout(s,6.8,3.9,5.9,2.6,f"**Ahorro:** {f0((A['total']-C['total'])/1e6)} MM COP vs. nivelación y {f0((B['total']-C['total'])/1e6)} MM vs. persecución. Gana en los 5 escenarios de sensibilidad probados.",fill=LIGHT,size=18)

# 12 MPS
s=slide("MPS y ATP: Cemento UG en saco 50 kg","Integrante 4 (1 min). Pedidos firmes simulados y declarados. Control: suma mensual del MPS = plan agregado × 45 %.")
rows=[[r[0],f0(r[1]),f0(r[2]),f0(r[3]) if r[3] else "—",f0(r[4]),f0(M.ATP[i]) if M.ATP[i] is not None else "—"] for i,r in enumerate(M.MPS_ROWS)]
tabla(s,0.5,1.4,7.4,["Sem.","Pronóstico","Pedidos","MPS","Inv. proy.","ATP"],rows,[0.7,1.3,1.3,1.3,1.5,1.3],size=11,rowh=0.37)
text(s,8.2,1.5,4.8,5.3,["Campañas **quincenales** de 90.900 t (180% del plan mensual × 45 %)".replace("180% del plan mensual × 45 %","= mitad del plan mensual de la referencia"),"**ATP** = MPS (+ inventario inicial) − pedidos hasta la próxima campaña",f"Semana 3: ATP de solo {f0(M.ATP[2])} t → ofrecer pedidos nuevos grandes en la semana 5 ({f0(M.ATP[4])} t)","**Control:** suma mensual del MPS = plan agregado × participación (cuadra)"],size=16)

# 15 tablero
s=slide("Tablero de indicadores (datos simulados)","Integrante 5 (1 min). Mostrar el Excel: selector de mes, semáforos y gráficos.")
img(s,"fig/f9_tablero.png",0.4,1.3,w=7.9)
rows=[["MAPE","≤5 %",pc(kpi["MAPE del pronóstico"][0]),kpi["MAPE del pronóstico"][1]],["Utiliz. horno","85–100 %",pc(kpi["Utilización del horno"][0]),kpi["Utilización del horno"][1]],["Cumpl. MPS","≥97 %",pc(kpi["Cumplimiento del MPS"][0]),kpi["Cumplimiento del MPS"][1]],["Fill rate","≥98 %",pc(kpi["Fill rate (nivel de servicio)"][0]),kpi["Fill rate (nivel de servicio)"][1]],["Factor clínker (ESG)","≤0,74",f"{kpi['Factor clínker (ESG)'][0]:.3f}".replace(".",","),kpi["Factor clínker (ESG)"][1]]]
tabla(s,8.5,1.5,4.5,["Indicador","Meta","Prom.","Estado"],rows,[1.7,0.9,0.9,1.0],size=11,rowh=0.5)
text(s,8.5,4.6,4.5,2,["5 indicadores, cada uno ligado a un componente","Selector de mes y semáforos en el Excel"],size=14,color=GRAY)

# 16 recomendaciones
s=slide("Recomendaciones priorizadas","Integrante 1 o 2 (1,5 min). Valores no aditivos; no incluyen costo de implementación.")
rows=[["1","Adoptar el plan mixto",f"{f0((A['total']-C['total'])/1e6)} / {f0((B['total']-C['total'])/1e6)}","Cumpl. MPS; fill rate"],["2","Bajar factor clínker a 0,72 (ESG)",f"{f0(esc_fc)} + {f0(M.CO2_AHORRO)} kt CO₂/año","Factor clínker"],["3","+2 pp de disponibilidad del horno",f"{f0(esc_disp)}","Utilización / disponibilidad"]]
tabla(s,0.5,1.5,12.3,["N.º","Recomendación","Impacto (MM COP)","Indicador"],rows,[0.7,5.0,3.6,3.0],size=15,rowh=0.7)
text(s,0.5,5.0,12.3,1.8,["**Industria 4.0:** mantenimiento predictivo del horno (sensores y analítica).","**ESG:** menos clínker = menos emisiones y más capacidad en el horno."],size=17)

# 17 supuestos y limitaciones
s=slide("Supuestos y limitaciones","Integrante 3 o 4 (1 min). Ser transparentes: es lo que pide la docente.")
text(s,0.6,1.4,6,5.5,["**Datos públicos:** DANE y Grupo Argos (totales anuales, acumulado 2026, Argos Colombia 2025)","**Reconstruido:** serie mensual histórica, calibrada con esos totales","**Supuestos declarados:** capacidades, precios, costos, mezcla","**Simulado:** pedidos del MPS y datos del tablero"],size=18)
text(s,6.9,1.4,6,5.5,["**Limitaciones:**","La serie es corta (el error del pronóstico es in-sample)","La etapa más cargada depende de las capacidades supuestas","No se estimó el costo de implementar las recomendaciones","Participación de mercado constante"],size=18)

# 18 conclusiones
s=slide("Conclusiones","Integrante 5 (1 min) y cierre. Abrir a preguntas.")
text(s,0.6,1.4,12.2,4.4,[f"Demanda de **{f0(M.DEM.sum())} kt** (≈ +5 % anual); Holt-Winters, el método más preciso","El **horno de clínker** es la etapa más cargada y solo supera su capacidad en el pico","La **estrategia mixta** minimiza el costo y cumple el 100 % de la demanda sin comprar clínker","MPS coherente con el plan agregado; el ATP indica cuánto se puede prometer en cada campaña","Reducir el factor clínker une **productividad y sostenibilidad**"],size=20)
callout(s,0.6,5.7,12.2,1.0,"Gracias. ¿Preguntas?",size=26)

def src(slide,txt):
    tf=slide.shapes.add_textbox(Inches(0.6),Inches(6.72),Inches(12.1),Inches(0.35)).text_frame; tf.word_wrap=True
    r=tf.paragraphs[0].add_run(); r.text=txt; r.font.size=Pt(11); r.font.italic=True; r.font.color.rgb=GRAY
SL=prs.slides
src(SL[2],"Fuentes: Departamento Administrativo Nacional de Estadística [DANE] (2026); Grupo Argos (2026). Participación calculada por el equipo.")
src(SL[5],"Medidas de error y método: Heizer et al. (2020); Hyndman y Athanasopoulos (2021); Winters (1960).")
src(SL[7],"Definiciones de capacidad: Heizer et al. (2020). Capacidades de diseño: supuestos del equipo.")
src(SL[9],"Estrategias de planeación agregada: Heizer et al. (2020). Costos: supuestos del equipo.")
src(SL[11],"Lógica de MPS y ATP: Vollmann et al. (2005). Pedidos firmes simulados.")
s=slide("Referencias")
refs=[("Argos. (2025). ","Reporte integrado Cementos Argos 2024",". https://files.grupoargos.com/uploads-grupo-argos/2025/03/Reporte-Integrado-cementos-argos-2024.pdf"),
("Departamento Administrativo Nacional de Estadística. (2026). ","Estadísticas de cemento gris (ECG): boletines técnicos",". https://www.dane.gov.co/index.php/en/statistics-by-topic-1/construction/estadisticas-de-cemento-gris"),
("Grupo Argos. (2026). ","Informe de gestión Grupo Argos 2025",". https://files.grupoargos.com/uploads-grupo-argos/2026/03/InformeGestion_GrupoArgos_2025.pdf"),
("Heizer, J., Render, B., & Munson, C. (2020). ","Operations management: Sustainability and supply chain management"," (13.ª ed.). Pearson."),
("Hyndman, R. J., & Athanasopoulos, G. (2021). ","Forecasting: Principles and practice"," (3.ª ed.). OTexts. https://otexts.com/fpp3/"),
("Vollmann, T. E., Berry, W. L., Whybark, D. C., & Jacobs, F. R. (2005). ","Manufacturing planning and control for supply chain management"," (5.ª ed.). McGraw-Hill."),
("Winters, P. R. (1960). Forecasting sales by exponentially weighted moving averages. ","Management Science, 6","(3), 324–342. https://doi.org/10.1287/mnsc.6.3.324")]
tf=s.shapes.add_textbox(Inches(0.6),Inches(1.35),Inches(12.1),Inches(5.6)).text_frame; tf.word_wrap=True
for i,(a_,b_,c_) in enumerate(refs):
    p=tf.paragraphs[0] if i==0 else tf.add_paragraph(); p.space_after=Pt(6)
    for t,it in ((a_,False),(b_,True),(c_,False)):
        r=p.add_run(); r.text=t; r.font.size=Pt(13); r.italic=it; r.font.color.rgb=NAVY; r.font.name="Times New Roman"
import re as _re
for sl in prs.slides:
    if sl.has_notes_slide:
        tf_=sl.notes_slide.notes_text_frame
        t_=_re.sub(r"^Integrante [\d o]+\s*","",tf_.text); t_=_re.sub(r"^\(","(",t_); tf_.text=t_
prs.save("Presentacion_Proyecto_Final_Argos.pptx"); print("slides",num[0])
