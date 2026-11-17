import subprocess, json, re, sys
PDF=sys.argv[1]; OUT="/home/user/mauxysalcedo.github.io/proyecto_final/toc_pages.json"
n=int(re.search(r'Pages:\s+(\d+)',subprocess.run(['pdfinfo',PDF],capture_output=True,text=True).stdout).group(1))
TOC=["Introducción y alcance","Contextualización de la empresa","Pronóstico de demanda","Análisis de capacidad","Plan agregado","Plan Maestro de Producción (MPS) y ATP","Diagnóstico de restricción (TOC)","Tablero de indicadores","Recomendaciones","Conclusiones","Anexo A. Tabla consolidada de supuestos","Referencias"]
pages={}
for p in range(3,n+1):
    t=subprocess.run(['pdftotext','-f',str(p),'-l',str(p),'-layout',PDF,'-'],capture_output=True,text=True).stdout
    for h in TOC:
        if h not in pages and re.search(r'^\s*'+re.escape(h)+r'\s*$',t,re.M): pages[h]=p
print(n,pages); json.dump(pages,open(OUT,'w'))
