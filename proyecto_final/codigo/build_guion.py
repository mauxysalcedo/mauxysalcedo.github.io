from docx import Document
from docx.shared import Pt, Cm
import modelo as M
d=Document(); st=d.styles["Normal"]; st.font.name="Calibri"; st.font.size=Pt(11)
for s in d.sections: s.left_margin=s.right_margin=Cm(2.2); s.top_margin=s.bottom_margin=Cm(2)
d.add_heading("Guion del pitch y preparación de la sustentación",1)
d.add_paragraph("Duración objetivo: unos 12 minutos (la guía permite de 8 a 20), más unos 5 minutos de preguntas. Los cuatro integrantes deben hablar.")
d.add_heading("Reparto y tiempos",2)
t=d.add_table(rows=1,cols=4); t.style="Table Grid"
for i,h in enumerate(["Integrante","Diapositivas","Tiempo","Mensaje central"]): t.rows[0].cells[i].text=h
for r in [("Natalia Aragón","1–5 y 16","3,5 min","Quiénes somos, la empresa, el problema, cómo se conectan los análisis y las recomendaciones"),("Mary Barrera","6–9","3 min","Holt-Winters por menor error; demanda de 5.518 kt; el horno es la etapa más cargada; estrategia de ajuste"),("Brandon Afanador","10–12 y 17","3 min","Tres estrategias, la mixta gana; MPS y ATP de la referencia UG; supuestos y limitaciones"),("María Auxiliadora Salcedo","13–15 y 18","3 min","La restricción y cómo explotarla; tablero de indicadores; conclusiones")]:
    c=t.add_row().cells
    for i,v in enumerate(r): c[i].text=v
d.add_paragraph()
d.add_heading("Consejos",2)
for x in ["No lean las diapositivas: cuenten la historia 'demanda → capacidad → plan → MPS → restricción → tablero'.","Digan desde el inicio qué es dato público, qué es supuesto y qué es simulado. La docente pide transparencia, y eso evita que una pregunta los descoloque.","Cada integrante debe poder explicar cualquier cifra del Excel: ábranlo y recorran las hojas en orden antes de la sustentación.","Practiquen con cronómetro al menos una vez completa."]:
    d.add_paragraph(x,style="List Bullet")
d.add_heading("Preguntas probables y cómo responderlas",2)
qa=[
("¿De dónde salen los datos de demanda?","Los totales anuales y varios puntos de 2026 son del DANE (ECG). La serie mensual la reconstruimos con una estacionalidad supuesta, calibrada para que los totales coincidan con los públicos, porque no pudimos descargar la serie mensual oficial. En el Excel, la hoja 'Historico' se reemplaza por la serie oficial y todo se recalcula."),
("¿Por qué Holt-Winters y no otro método?","Tiene el menor MAPE (2,2 % frente a 4,3 % y 5,4 %) y su señal de rastreo está cerca de cero. Los otros dos tienen señales de rastreo de 7,5: subestiman sistemáticamente porque no modelan tendencia ni estacionalidad. Reconocemos que la serie es corta y el error es in-sample."),
("¿Por qué el horno es la restricción?","Es la etapa con mayor utilización media (94,1 %) y la única que supera el 100 % en el pico (102,6 %). Aclaramos que depende de capacidades que son supuestos nuestros; con datos internos podría cambiar, y el Excel lo permite."),
("Si el horno es la restricción, ¿por qué el plan produce por debajo de su capacidad al inicio?","Porque en dic–mar la restricción es el mercado: la demanda es menor que la capacidad. Producir más solo acumularía inventario costoso. A partir de abril el plan sube a 488 kt/mes, casi la capacidad regular, y acumula inventario para el pico."),
("Las tres estrategias cuestan casi lo mismo. ¿Por qué importa?","La diferencia relativa es menor al 0,5 %, pero son 3.887 millones de COP frente a la persecución y 7.723 frente a la nivelación, y la mixta evita comprar clínker en el pico. Probamos 5 escenarios de costos y la mixta gana en todos."),
("¿Qué es el ATP y cómo lo calcularon?","Es la cantidad disponible para prometer a nuevos pedidos sin afectar los comprometidos. En la primera campaña es el inventario inicial más el MPS menos los pedidos firmes hasta la siguiente campaña; en las demás, el MPS menos los pedidos hasta la siguiente. Los pedidos firmes son simulados."),
("¿Cómo aplicaron TOC?","Aplicamos los dos primeros pasos: identificar (horno) y explotar. Explotar significa priorizar el producto de mayor trúput por tonelada de clínker (UG saco: 319.444 COP/t de clínker frente a 236.111 del granel) y proteger la disponibilidad del horno (+2 pp con mantenimiento predictivo)."),
("¿Cuánto cuesta implementar las recomendaciones?","No lo estimamos por falta de datos; lo declaramos como limitación. Los valores que mostramos son ahorros del modelo y no son aditivos. El siguiente paso sería evaluar el costo de los sensores y de las pruebas de adiciones."),
("¿Qué tiene que ver esto con sostenibilidad e Industria 4.0?","Reducir el factor clínker de 0,75 a 0,72 evita unas 141 mil toneladas de CO₂ al año (con un factor de referencia de 0,85 t CO₂/t clínker, a verificar) y libera capacidad del horno. El mantenimiento predictivo con sensores y analítica es la aplicación de Industria 4.0."),
("¿Los datos del tablero son reales?","No, son simulados (el plan o el pronóstico más una variación aleatoria con semilla fija) para demostrar cómo funciona el tablero, como pide la guía."),
]
for q,a in qa:
    p=d.add_paragraph(); r=p.add_run("P: "+q); r.bold=True
    d.add_paragraph("R: "+a)
d.add_heading("Lista de verificación antes de entregar",2)
for x in ["Reemplazar la serie reconstruida (hoja 'Historico') por la serie oficial de la ECG del DANE y revisar que los pronósticos y el plan sigan cuadrando.","Confirmar en el Reporte Integrado vigente de Argos: despachos de cemento en Colombia (5,0 frente a 5,4 millones de t), capacidad instalada y factor clínker.","Contrastar con el DANE: 12.109,6 kt (2024), 12.708,0 kt (2025), 7.555,2 kt (ene–jul 2026), 1.061,6 kt (jun-26) y 1.194,4 kt (sep-26). Estas cifras se tomaron de resultados de búsqueda y no se abrió el boletín.","Revisar los supuestos marcados 'Sí' en la columna 'Verificar' del Anexo A (hoja 'Supuestos').","Ajustar la bibliografía a los textos guía del curso.","En Word, el índice tiene números de página fijos calculados en LibreOffice: confirmar que coincidan al abrir el archivo."]:
    d.add_paragraph(x,style="List Bullet")
d.save("Guion_Pitch_y_Preguntas.docx")
