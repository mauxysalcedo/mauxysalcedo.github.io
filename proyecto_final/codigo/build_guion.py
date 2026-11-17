from docx import Document
from docx.shared import Pt, Cm
import modelo as M
import numpy as np
S=M.S; P=M.MEJOR_PLAN["P"]; PA,PB,PC=M.PLANES
def f0(x): return f"{x:,.0f}".replace(",",".")
def f1(x): return f"{x:,.1f}".replace(",","X").replace(".",",").replace("X",".")
def pc(x,d=1): return f"{x*100:.{d}f} %".replace(".",",")
ES=["ene","feb","mar","abr","may","jun","jul","ago","sep","oct","nov","dic"]; mes=lambda d:f"{ES[d.month-1]}-{str(d.year)[2:]}"
k1=int(sum(1 for x in P if abs(x-P[0])<1e-9)); sav_A=(PA["total"]-PC["total"])/1e6; sav_B=(PB["total"]-PC["total"])/1e6
rel=max(PA["total"]-PC["total"],PB["total"]-PC["total"])/PC["total"]*100
mapes=[r[2] for r in M.RES]; trk=[r[3] for r in M.RES]
d=Document(); st=d.styles["Normal"]; st.font.name="Calibri"; st.font.size=Pt(11)
for s in d.sections: s.left_margin=s.right_margin=Cm(2.2); s.top_margin=s.bottom_margin=Cm(2)
d.add_heading("Guion del pitch y preparación de la sustentación",1)
d.add_paragraph("Equipo: Brandon Afanador, Natalia Aragón, Mary Barrera y María Auxiliadora Salcedo Pérez. Docente: Natasha Charris.")
d.add_paragraph("Duración objetivo: unos 12 minutos (la guía permite de 8 a 20), más unos 5 minutos de preguntas. Los cuatro integrantes deben hablar.")
d.add_heading("Bloques y tiempos sugeridos",2)
d.add_paragraph("El reparto de quién presenta cada bloque lo define el equipo; aquí solo se sugiere cómo dividir las diapositivas y el tiempo. Los cuatro integrantes deben hablar.")
t=d.add_table(rows=1,cols=3); t.style="Table Grid"
for i,h in enumerate(["Diapositivas","Tiempo","Mensaje central"]): t.rows[0].cells[i].text=h
for r in [("1–5","3 min","Quiénes somos, la empresa, el problema y cómo se conectan los análisis"),("6–9","3 min","Holt-Winters por menor error; demanda de "+f0(M.DEM.sum())+" kt; el horno es la etapa más cargada; estrategia de ajuste"),("10–12","3 min","Tres estrategias, la mixta gana; MPS y ATP de la referencia UG"),("13","1,5 min","Tablero de indicadores"),("14–17","2,5 min","Recomendaciones, supuestos y limitaciones, conclusiones y referencias")]:
    c=t.add_row().cells
    for i,v in enumerate(r): c[i].text=v
d.add_paragraph()
d.add_heading("Consejos",2)
for x in ["No lean las diapositivas: cuenten la historia 'demanda → capacidad → plan → MPS → tablero'.","Digan desde el inicio qué es dato público, qué es supuesto y qué es simulado. La docente pide transparencia, y eso evita que una pregunta los descoloque.","Cada integrante debe poder explicar cualquier cifra del Excel: ábranlo y recorran las hojas en orden antes de la sustentación.","Practiquen con cronómetro al menos una vez completa."]:
    d.add_paragraph(x,style="List Bullet")
d.add_heading("Preguntas probables y cómo responderlas",2)
qa=[
("¿De dónde salen los datos de demanda?","La serie mensual de despachos nacionales es la oficial del DANE (Estadísticas de Cemento Gris, Anexo A1 del boletín de septiembre de 2026; las cifras de 2026 son provisionales). La demanda de Argos es esa serie por su participación de mercado ("+pc(S["share"])+"), calculada con las 5,0 millones de toneladas de la regional Colombia que reporta Grupo Argos para 2025."),
("¿Por qué Holt-Winters y no otro método?",f"Tiene el menor MAPE ({pc(mapes[2]/100)} frente a {pc(mapes[1]/100)} y {pc(mapes[0]/100)}) y su señal de rastreo está cerca de cero. Los otros dos tienen señales de rastreo de {str(round(trk[0],1)).replace('.',',')} y {str(round(trk[1],1)).replace('.',',')}: subestiman sistemáticamente porque no modelan tendencia ni estacionalidad. Reconocemos que la serie es corta (33 meses) y el error es in-sample."),
("¿Por qué el horno es la etapa más cargada?",f"Es la etapa con mayor utilización media ({pc(M.UTIL_AVG['Horno (clínker)'])}) y la única que supera el 100 % en el pico ({pc(M.UTIL_MAX['Horno (clínker)'])}). Aclaramos que depende de capacidades que son supuestos nuestros; con datos internos podría cambiar, y el Excel lo permite."),
("Si el horno es la etapa más cargada, ¿por qué el plan produce por debajo de su capacidad al inicio?",f"Porque de {mes(M.PLAN_LABEL[0])} a {mes(M.PLAN_LABEL[k1-1])} el límite es el mercado: la demanda es menor que la capacidad. Producir más solo acumularía inventario costoso. A partir de {mes(M.PLAN_LABEL[k1])} el plan sube a {f1(P[k1])} kt/mes, casi la capacidad regular ({f1(M.CAP_CEM_REG)} kt/mes), y acumula inventario para el pico."),
("Las tres estrategias cuestan casi lo mismo. ¿Por qué importa?",f"La diferencia relativa es de {str(round(rel,1)).replace('.',',')} % o menos, pero son {f0(sav_B)} millones de COP frente a la persecución y {f0(sav_A)} frente a la nivelación, y la mixta evita comprar clínker en el pico. Probamos 5 escenarios de costos y la mixta gana en todos."),
("¿Qué es el ATP y cómo lo calcularon?","Es la cantidad disponible para prometer a nuevos pedidos sin afectar los comprometidos. En la primera campaña es el inventario inicial más el MPS menos los pedidos firmes hasta la siguiente campaña; en las demás, el MPS menos los pedidos hasta la siguiente. Los pedidos firmes son simulados."),
("¿Cuánto cuesta implementar las recomendaciones?","No lo estimamos por falta de datos; lo declaramos como limitación. Los valores que mostramos son ahorros del modelo y no son aditivos. El siguiente paso sería evaluar el costo de los sensores y de las pruebas de adiciones."),
("¿Qué tiene que ver esto con sostenibilidad e Industria 4.0?",f"Reducir el factor clínker de 0,75 a 0,72 evita unas {f0(M.CO2_AHORRO)} mil toneladas de CO₂ al año (con un factor de referencia de 0,85 t CO₂/t clínker, a verificar) y libera capacidad del horno. El mantenimiento predictivo con sensores y analítica es la aplicación de Industria 4.0."),
("¿Los datos del tablero son reales?","No, son simulados (el plan o el pronóstico más una variación aleatoria con semilla fija) para demostrar cómo funciona el tablero, como pide la guía."),
]
for q,a in qa:
    p=d.add_paragraph(); r=p.add_run("P: "+q); r.bold=True
    d.add_paragraph("R: "+a)
d.add_heading("Lista de verificación antes de entregar",2)
for x in ["Confirmar en el Reporte Integrado vigente de Argos: despachos de cemento en Colombia (5,0 frente a 5,4 millones de t), capacidad instalada y factor clínker.","La serie de demanda es la oficial del DANE (Anexo A1, boletín de septiembre de 2026). Las cifras de 2026 son provisionales y pueden ser revisadas por el DANE; si publican un boletín nuevo antes de entregar, actualicen la hoja 'Historico' del Excel.","Revisar los supuestos marcados 'Sí' en la columna 'Verificar' del Anexo A (hoja 'Supuestos').","Ajustar la bibliografía a los textos guía del curso.","En Word, el índice tiene números de página fijos calculados en LibreOffice: confirmar que coincidan al abrir el archivo."]:
    d.add_paragraph(x,style="List Bullet")
d.save("/home/user/mauxysalcedo.github.io/proyecto_final/Guion_Pitch_y_Preguntas.docx")
