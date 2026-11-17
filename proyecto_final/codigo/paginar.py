import subprocess, json, re
n=int(re.search(r'Pages:\s+(\d+)',subprocess.run(['pdfinfo','Documento_Proyecto_Final_Argos.pdf'],capture_output=True,text=True).stdout).group(1))
TOC=["1. Introducción y alcance","2. Contextualización de la empresa","3. Pronóstico de demanda","4. Análisis de capacidad","5. Plan agregado","6. Plan Maestro de Producción (MPS) y ATP","7. Diagnóstico de restricción (TOC)","8. Tablero de indicadores","9. Recomendaciones","10. Conclusiones","Anexo A. Tabla consolidada de supuestos","Bibliografía"]
pages={}
for p in range(3,n+1):
    t=subprocess.run(['pdftotext','-f',str(p),'-l',str(p),'-layout','Documento_Proyecto_Final_Argos.pdf','-'],capture_output=True,text=True).stdout
    for h in TOC:
        if h not in pages and re.search(r'^\s*'+re.escape(h)+r'\s*$',t,re.M): pages[h]=p
print(n,pages); json.dump(pages,open('toc_pages.json','w'))
