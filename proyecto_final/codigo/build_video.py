"""Video explicativo del proyecto (voz espeak-ng + diapositivas PIL + ffmpeg)."""
import os, re, subprocess, sys, textwrap
import numpy as np
from PIL import Image, ImageDraw, ImageFont
BASE = "/home/user/mauxysalcedo.github.io/proyecto_final"
OUT = sys.argv[1]            # carpeta de trabajo
os.chdir(BASE); sys.path.insert(0, BASE + "/codigo")
import modelo as M
from openpyxl import load_workbook

S = M.S; P = M.MEJOR_PLAN["P"]; PA, PB, PC = M.PLANES
def f0(x): return f"{x:,.0f}".replace(",", ".")
def f1(x): return f"{x:,.1f}".replace(",", "X").replace(".", ",").replace("X", ".")
def pc(x, d=1): return f"{x*100:.{d}f} %".replace(".", ",")
ES = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"]
mes = lambda d: f"{ES[d.month-1]}-{str(d.year)[2:]}"
k1 = int(sum(1 for x in P if abs(x - P[0]) < 1e-9))
sav_A = (PA["total"] - PC["total"]) / 1e6; sav_B = (PB["total"] - PC["total"]) / 1e6
g25 = S["nac_2025"] / S["nac_2024"] - 1
g_impl = (M.DEM.sum() / S["argos_col_2025"]) ** (1 / 1.95) - 1
esc_disp = (M.ESC_BASE["total"] - M.ESC_DISP["total"]) / 1e6; esc_fc = (M.ESC_BASE["total"] - M.ESC_FC["total"]) / 1e6
tb = load_workbook(BASE + "/Proyecto_Argos_Modelo_y_Tablero.xlsx", data_only=True)["Tablero"]
# --- números para locución (sin separadores de miles)
def w(x): return f"{x:.0f}"
def wc(x, d=1): return f"{x:.{d}f}".replace(".", " coma ")
mape = [r[2] for r in M.RES]; trk = [r[3] for r in M.RES]
hu = M.UTIL_AVG["Horno (clínker)"] * 100; mu = M.UTIL_AVG["Molienda"] * 100; eu = M.UTIL_AVG["Ensacado"] * 100; hp = M.UTIL_MAX["Horno (clínker)"] * 100
dem = M.DEM.sum(); mn = M.DEM.min(); mx = M.DEM.max()

NAVY = (11, 31, 58); GOLD = (242, 183, 5); GRAY = (85, 95, 112); LIGHT = (238, 241, 246); RED = (192, 57, 43); GREEN = (46, 139, 87); WHITE = (255, 255, 255)
FT = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSerif-Bold.ttf", 54)
FB = ImageFont.truetype("/usr/share/fonts/truetype/crosextra/Carlito-Regular.ttf", 44)
FBB = ImageFont.truetype("/usr/share/fonts/truetype/crosextra/Carlito-Bold.ttf", 44)
FS = ImageFont.truetype("/usr/share/fonts/truetype/crosextra/Carlito-Regular.ttf", 38)
W, H = 1920, 1080

def wrap(d, text, font, maxw):
    words = text.split(); lines = []; cur = ""
    for wd in words:
        t = (cur + " " + wd).strip()
        if d.textlength(t, font=font) <= maxw: cur = t
        else: lines.append(cur); cur = wd
    if cur: lines.append(cur)
    return lines

def draw_rich(d, x, y, text, font, fontb, maxw, color=NAVY, lh=58):
    """texto con **negrita** y salto de línea automático; devuelve y final"""
    toks = re.split(r"(\*\*.*?\*\*)", text); words = []
    for t in toks:
        b = t.startswith("**")
        for wd in t.strip("*").split(" "):
            if wd: words.append((wd, b))
    cx = x
    for wd, b in words:
        f = fontb if b else font; ww = d.textlength(wd + " ", font=f)
        if cx + ww - x > maxw: cx = x; y += lh
        d.text((cx, y), wd, font=f, fill=color); cx += ww
    return y + lh

FB2 = ImageFont.truetype("/usr/share/fonts/truetype/crosextra/Carlito-Regular.ttf", 56)
FBB2 = ImageFont.truetype("/usr/share/fonts/truetype/crosextra/Carlito-Bold.ttf", 56)
def slide(title, bullets, fig=None, boxes=None, sub=None):
    im = Image.new("RGB", (W, H), WHITE); d = ImageDraw.Draw(im)
    d.rectangle([0, 0, W, 130], fill=NAVY); d.rectangle([0, 130, W, 140], fill=GOLD)
    d.text((70, 33), title, font=FT, fill=WHITE)
    wide = False
    if fig:
        fi = Image.open(fig).convert("RGB"); wide = fi.width / fi.height > 2.6
    side = bool(fig) and not wide
    big = not fig and not boxes
    f1_, f2_, lh = (FB2, FBB2, 74) if big else (FB, FBB, 58)
    textw = 800 if side else 1780
    y = 200 if big else 190
    for b in bullets:
        d.ellipse([70, y + 22, 92, y + 44], fill=GOLD)
        y = draw_rich(d, 115, y, b, f1_, f2_, textw - 70, lh=lh) + (26 if not big else 22)
    if fig:
        if wide:
            mw, mh = 1740, 330
            r = min(mw / fi.width, mh / fi.height); fi = fi.resize((int(fi.width * r), int(fi.height * r)))
            im.paste(fi, ((W - fi.width) // 2, max(y + 20, 470)))
        else:
            mw, mh = 1000, 640
            r = min(mw / fi.width, mh / fi.height); fi = fi.resize((int(fi.width * r), int(fi.height * r)))
            im.paste(fi, (W - 50 - fi.width, 175 + (mh - fi.height) // 2))
    if boxes:
        n = len(boxes); bw = (1780 - 40 * (n - 1)) // n; x = 70
        for (t, body, col) in boxes:
            d.rounded_rectangle([x, 200, x + bw, 780], 24, fill=col)
            d.text((x + 30, 235), t, font=FBB, fill=WHITE if col != GOLD else NAVY)
            yy = 325
            for ln in wrap(d, body, ImageFont.truetype("/usr/share/fonts/truetype/crosextra/Carlito-Regular.ttf", 44), bw - 60):
                d.text((x + 30, yy), ln, font=ImageFont.truetype("/usr/share/fonts/truetype/crosextra/Carlito-Regular.ttf", 44), fill=WHITE if col != GOLD else NAVY); yy += 58
            x += bw + 40
    return im

def with_sub(im, text):
    im = im.copy(); ov = Image.new("RGBA", (W, H), (0, 0, 0, 0)); d = ImageDraw.Draw(ov)
    d.rectangle([0, 840, W, H], fill=(11, 31, 58, 235))
    im = Image.alpha_composite(im.convert("RGBA"), ov).convert("RGB"); d = ImageDraw.Draw(im)
    lines = wrap(d, text, FS, 1760)[:4]; y = 860 + (210 - len(lines) * 48) // 2
    for ln in lines:
        d.text((80, y), ln, font=FS, fill=WHITE); y += 48
    return im

F = "fig/"
scenes = [
 ("El proyecto en una frase",
  ["**Empresa:** Cementos Argos, cemento gris en Colombia", "**Pregunta:** ¿cuánto producir cada mes para cumplir la demanda al menor costo?", "**Cinco pasos en orden:** pronóstico, capacidad, plan agregado, MPS y tablero"],
  F + "f7_flujo.png", None,
  "Hola. Este video explica el proyecto en pocos minutos, para que lo entiendan bien y puedan explicarlo. La idea completa cabe en una frase. Cementos Argos vende más cemento en unos meses que en otros, y su horno tiene un límite de producción. El proyecto responde cuánto debe producir cada mes para cumplir la demanda gastando lo menos posible. Para responder, hacemos cinco cosas en orden. Primero, pronosticamos la demanda. Segundo, miramos cuánto se puede producir. Tercero, decidimos cuánto producir cada mes. Cuarto, bajamos ese plan a semanas para un producto. Y quinto, armamos un tablero para medir los resultados. Cada paso usa el resultado del anterior, y eso es lo que hace coherente el trabajo."),
 ("La empresa y los datos",
  [f"**Mercado nacional 2025:** {f1(S['nac_2025'])} kt (DANE)", "**Argos Colombia 2025:** 5.000 kt (Grupo Argos)", f"**Participación:** 5.000 / {f1(S['nac_2025'])} = **{pc(S['share'])}**", f"**La demanda crece** cerca de {pc(g_impl,0)} al año", "1 kt = mil toneladas"],
  None, None,
  f"Argos es una cementera colombiana. El DANE publica cada mes cuánto cemento gris se despacha en todo el país. En dos mil veinticinco fueron doce mil novecientas once kilotoneladas. Una kilotonelada es mil toneladas, o sea que son casi trece millones de toneladas. Grupo Argos reporta que su operación en Colombia vendió cinco millones de toneladas. Si dividimos cinco millones entre el total nacional, Argos tiene alrededor del {wc(S['share']*100)} por ciento del mercado. Con ese porcentaje convertimos la demanda nacional en la demanda de Argos. La demanda crece cerca de {w(g_impl*100)} por ciento al año."),
 ("Dato, supuesto o simulado",
  [], None,
  [("DATOS PÚBLICOS", "Serie mensual del DANE (cemento gris) y las 5 millones de toneladas de Argos Colombia.", NAVY),
   ("SUPUESTOS", "Capacidades de las máquinas, precios y costos. No son públicos: los declaramos y justificamos.", GOLD),
   ("SIMULADOS", "Pedidos del plan maestro y datos 'reales' del tablero. Se inventaron para mostrar cómo funciona.", GRAY)],
  "Hay una distinción que deben tener clara. Hay tres tipos de cifras. Los datos públicos son la serie del DANE y las cinco millones de toneladas de Argos en Colombia. Los supuestos son las capacidades de las máquinas, los precios y los costos. La empresa no los publica así, entonces los declaramos y los justificamos. Y los simulados son los pedidos del plan maestro y los datos reales del tablero, que inventamos solo para mostrar cómo funcionan. Cuando les pregunten, digan siempre cuál es cuál. Eso es justo lo que pidió la profesora."),
 ("Paso 1: Pronóstico",
  ["**Qué es:** estimar cuánto cemento se pedirá en los próximos 12 meses", "**Se compararon 3 métodos** con 33 meses de datos reales", f"**Gana Holt-Winters:** error de {pc(mape[2]/100)}", "**MAPE** = error promedio en porcentaje"],
  F + "f2_mape.png", None,
  f"Pronosticar es estimar cuánto cemento se pedirá en los próximos doce meses. Probamos tres métodos. El promedio móvil promedia los últimos tres meses. Es simple, pero no entiende que hay meses buenos y meses malos. La suavización exponencial da más peso a lo más reciente, pero tampoco entiende los ciclos. Holt Winters sí. Separa la demanda en tres partes: el nivel, la tendencia, que es si la demanda sube o baja con el tiempo, y la estacionalidad, que es el patrón de meses altos y bajos, como enero, que siempre es flojo. Para comparar usamos el mape, que es el error promedio en porcentaje. Holt Winters se equivoca en promedio un {wc(mape[2])} por ciento. Los otros dos, alrededor de seis. También miramos la señal de rastreo. Si es muy grande, el método se equivoca siempre hacia el mismo lado, y eso les pasa a los otros dos. Un aviso honesto: la serie es corta, así que el error real a futuro puede ser mayor."),
 ("Resultado del pronóstico",
  [f"**{f0(dem)} kt** de demanda en 12 meses (dic-26 a nov-27)", f"**Mínimo:** {f0(mn)} kt en {mes(M.PLAN_LABEL[int(M.DEM.argmin())])}", f"**Máximo:** {f0(mx)} kt en {mes(M.PLAN_LABEL[int(M.DEM.argmax())])}", f"**Crecimiento** ≈ {pc(g_impl,0)} anual, coherente con el DANE"],
  F + "f1_pronostico.png", None,
  f"Este es el resultado. La demanda de Argos para los próximos doce meses suma unas {w(dem)} kilotoneladas. El mes más bajo es enero de dos mil veintisiete, con {w(mn)} kilotoneladas, y el más alto es octubre de dos mil veintisiete, con {w(mx)}. En la gráfica, la línea oscura es lo que ya pasó, y la verde es el pronóstico. Fíjense que sube y baja siguiendo el patrón de los años anteriores. Este número es el punto de partida de todo lo demás."),
 ("Paso 2: Capacidad",
  ["**Diseñada:** el máximo teórico de la máquina", "**Efectiva:** lo que realmente se logra, descontando paros", "**Utilización** = lo que necesito / lo que puedo producir", "**Tres etapas:** horno, molienda y ensacado"],
  F + "f3_util.png", None,
  "Ahora preguntamos: ¿cuánto se puede producir? Piensen en la cocina de un restaurante. La capacidad diseñada es lo máximo que la cocina podría sacar en teoría. La capacidad efectiva es lo que realmente saca, porque hay paradas y pérdidas. Y la utilización es qué tan llena está la cocina: lo que necesito dividido entre lo que puedo. Si pasa de cien por ciento, la demanda no cabe. En la fábrica hay tres etapas. El horno cocina la materia prima y produce clínker. La molienda lo muele con otros materiales para hacer cemento. Y el ensacado lo empaca. Cada etapa tiene su propia capacidad."),
 ("El horno es la etapa más cargada",
  [f"**Horno:** {pc(hu/100)} de utilización media", f"**Molienda:** {pc(mu/100)} · **Ensacado:** {pc(eu/100)}", f"**Octubre de 2027:** el horno llega a **{pc(hp/100)}**", f"El sistema produce **{f1(M.CAP_CEM_REG)} kt/mes** con recursos normales", "**Ojo:** las capacidades son supuestos nuestros"],
  None, None,
  f"El resultado es que el horno es la etapa más cargada. Su utilización media es de {wc(hu)} por ciento, mientras que la molienda tiene {wc(mu)} y el ensacado {wc(eu)}. Y en octubre de dos mil veintisiete el horno llega a {wc(hp)} por ciento: ahí la demanda ya no cabe. En términos de cemento, el sistema puede producir {wc(M.CAP_CEM_REG)} kilotoneladas al mes con recursos normales. Pero atención, porque esto hay que decirlo con transparencia: las capacidades de las máquinas son supuestos nuestros, porque la empresa no las publica. Si la capacidad real del horno fuera mayor, la etapa más cargada podría ser otra. Por eso es el supuesto más importante de verificar. Por ahora, la brecha es pequeña, así que recomendamos ajustarse con mejoras, y no construir un horno nuevo."),
 ("Paso 3: Plan agregado",
  ["**A. Nivelación:** producir siempre lo mismo", "**B. Persecución:** producir cada mes lo que se vende", "**C. Mixta:** dos niveles de producción", "**Costos:** producir, inventario, cambiar el ritmo, comprar clínker extra, faltantes"],
  F + "f4_agregado.png", None,
  "El plan agregado decide cuánto producir cada mes, sin entrar todavía en cada producto. Comparamos tres estrategias. La nivelación es producir siempre lo mismo, y absorber los meses altos con inventario guardado. La persecución es producir cada mes exactamente lo que se vende. Y la mixta usa dos niveles: produce menos al principio y más después. Cada estrategia tiene costos. Producir cuesta. Guardar inventario cuesta, porque el cemento guardado es plata parada. Cambiar el ritmo del horno cuesta, porque encenderlo y apagarlo es caro. Comprar clínker extra cuesta. Y quedarse sin cemento para vender es lo más caro, porque se pierde el ingreso."),
 ("Plan agregado: gana la mixta",
  [f"**Mixta:** {f1(P[0])} kt/mes hasta {mes(M.PLAN_LABEL[k1-1])} y {f1(P[k1])} kt/mes de {mes(M.PLAN_LABEL[k1])} a {mes(M.PLAN_LABEL[11])}", f"**Costo:** {f0(PC['total']/1e6)} millones de COP (mínimo)", f"**Ahorro:** {f0(sav_A)} MM vs. nivelar y {f0(sav_B)} MM vs. perseguir", "**Gana en los 5 escenarios** de sensibilidad probados", "Cumple el 100 % de la demanda sin comprar clínker"],
  None, None,
  f"La estrategia mixta es la más barata. Produce {w(P[0])} kilotoneladas al mes hasta abril de dos mil veintisiete, y luego {w(P[k1])} kilotoneladas al mes hasta noviembre. Así guarda inventario antes de los meses de más demanda, sin mover el horno cada mes. Cumple toda la demanda y no necesita comprar clínker. Ahorra unos {w(sav_A)} millones de pesos frente a nivelar, y unos {w(sav_B)} millones frente a perseguir la demanda. Una advertencia sincera: en porcentaje, la diferencia entre las tres estrategias es menor al medio por ciento, porque el costo de producir domina. Pero son miles de millones de pesos. Y probamos cinco escenarios cambiando los costos de inventario y de cambiar el ritmo, y la mixta ganó en todos. Eso es lo que hace robusta la recomendación."),
 ("Paso 4: MPS y ATP",
  ["**Una referencia:** cemento uso general en saco (45 %)", "**12 semanas**, campañas quincenales de **90.900 t**", "**ATP** = cuánto puedo prometer a clientes nuevos", f"**Ejemplo:** semana 3, ATP de solo {f0(M.ATP[2])} t"],
  F + "f5_mps.png", None,
  f"Ahora bajamos el plan mensual a semanas, para un solo producto: el cemento de uso general en saco. Esto se llama plan maestro de producción, o MPS. Se produce en campañas cada dos semanas, de noventa mil novecientas toneladas cada una. La suma de las campañas de cada mes es igual al plan agregado multiplicado por la participación del producto: esa es la coherencia entre los dos pasos. Aquí aparece el ATP, que significa disponible para prometer. Es la cantidad que puedo ofrecer a un cliente nuevo sin romper los pedidos que ya tengo comprometidos. Por ejemplo, en la semana tres el ATP es de solo {w(M.ATP[2])} toneladas, porque casi toda la campaña ya está vendida. Un pedido grande nuevo debería ofrecerse con la campaña de la semana cinco. Los pedidos firmes de este ejemplo son simulados."),
 ("Paso 5: Tablero de indicadores",
  ["**MAPE** del pronóstico (meta ≤ 5 %)", "**Utilización del horno** (85–100 %)", "**Cumplimiento del MPS** (≥ 97 %)", "**Fill rate** o nivel de servicio (≥ 98 %)", "**Factor clínker**, indicador de sostenibilidad (≤ 0,74)"],
  F + "f9_tablero.png", None,
  "El tablero tiene cinco indicadores, y cada uno viene de un paso anterior. El mape mide qué tan bien pronosticamos. La utilización del horno mide qué tan cargada está la etapa más importante. El cumplimiento del MPS compara lo que se produjo con lo que se planeó. El fill rate es el porcentaje de la demanda que se pudo cumplir. Y el factor clínker es cuánto clínker lleva cada tonelada de cemento, que es el indicador de sostenibilidad. Cada uno tiene una meta y un semáforo: verde, amarillo o rojo. Recuerden que los valores del tablero son simulados, solo para demostrar cómo funcionaría. En el Excel pueden elegir un mes y ver cómo cambian los semáforos."),
 ("Recomendaciones",
  [f"**1. Adoptar el plan mixto:** {f0(sav_A)} y {f0(sav_B)} MM COP", f"**2. Bajar el factor clínker** de 0,75 a 0,72: {f0(esc_fc)} MM COP y unas {f0(M.CO2_AHORRO)} kt de CO₂ menos al año", f"**3. Subir 2 puntos la disponibilidad del horno** con mantenimiento predictivo: {f0(esc_disp)} MM COP", "No incluyen el costo de implementarlas"],
  None, None,
  f"Con todo esto, hacemos tres recomendaciones. Primera, adoptar el plan mixto. Es la de mayor valor y no requiere inversión: es una decisión de programación. Segunda, reducir el factor clínker de cero coma setenta y cinco a cero coma setenta y dos, usando más adiciones. Eso ahorra unos {w(esc_fc)} millones de pesos, y evita unas {w(M.CO2_AHORRO)} mil toneladas de dióxido de carbono al año: aquí conectamos con sostenibilidad. Tercera, subir dos puntos la disponibilidad del horno con mantenimiento predictivo, que usa sensores y análisis de datos para evitar paradas sorpresa: aquí conectamos con la industria cuatro punto cero. Estos valores no se suman, porque cada escenario se calcula por separado. Y no incluimos el costo de implementarlas, porque no teníamos datos."),
 ("Limitaciones: lo que hay que decir",
  ["Las capacidades, precios y costos **son supuestos**", "Las cifras de 2026 del DANE son **provisionales**", "La serie es **corta** (33 meses): el error real puede ser mayor", "La participación de mercado se asume **constante**", "Los datos del tablero y los pedidos del MPS son **simulados**"],
  None, None,
  "Digan ustedes mismos las limitaciones, antes de que se las pregunten. Eso se ve como seriedad, no como debilidad. Las capacidades, los precios y los costos son supuestos. Las cifras de dos mil veintiséis del DANE son provisionales. La serie es corta, así que el error del pronóstico a futuro puede ser mayor. Asumimos que la participación de mercado de Argos no cambia. Y los datos del tablero y los pedidos del plan maestro son simulados. Si cambian la capacidad del horno en el Excel, todo se recalcula, y eso es parte de la fortaleza del trabajo."),
 ("Cómo explicarlo en un minuto",
  [f"Argos tiene demanda estacional y un horno con límite", f"Pronosticamos **{f0(dem)} kt** con Holt-Winters (error {pc(mape[2]/100)})", f"El **horno** es la etapa más cargada ({pc(hu/100)})", f"La **estrategia mixta** cuesta menos: ≈ {f0(sav_A)} y {f0(sav_B)} MM COP de ahorro", "Bajamos el plan a un MPS y medimos con **5 indicadores**"],
  None, None,
  f"Para cerrar, así pueden explicar todo en un minuto. Argos tiene una demanda que sube y baja durante el año y un horno con un límite. Pronosticamos unas {w(dem)} kilotoneladas con Holt Winters, con un error de {wc(mape[2])} por ciento. El horno es la etapa más cargada. La estrategia mixta, que produce en dos niveles, es la más barata y ahorra unos {w(sav_A)} millones de pesos frente a nivelar. Bajamos el plan a semanas con un plan maestro y lo medimos con cinco indicadores. Y recomendamos adoptar el plan mixto, reducir el factor clínker y mejorar la disponibilidad del horno. Ahora abran el Excel, cambien un supuesto y vean qué se mueve. Practiquen explicar cada paso en un minuto sin mirar. Mucho éxito."),
]

os.makedirs(OUT + "/wav", exist_ok=True); os.makedirs(OUT + "/img", exist_ok=True)
segs = []   # (imagen, wav, duración)
idx = 0
for si, (title, bullets, fig, boxes, narr) in enumerate(scenes):
    base = slide(title, bullets, fig, boxes)
    sentences = [t.strip() for t in re.split(r"(?<=[.?!:])\s+", narr) if t.strip()]
    for sent in sentences:
        idx += 1; wav = f"{OUT}/wav/s{idx:03d}.wav"
        subprocess.run(["espeak-ng", "-v", "es-419", "-s", "172", "-p", "42", "-w", wav, sent], check=True)
        dur = float(subprocess.run(["ffprobe", "-v", "error", "-show_entries", "format=duration", "-of", "csv=p=0", wav], capture_output=True, text=True).stdout)
        img = f"{OUT}/img/s{idx:03d}.png"; with_sub(base, sent).save(img)
        segs.append((img, wav, dur + 0.20))
# lista concat + audio con pausas
with open(f"{OUT}/list.txt", "w") as fl:
    for img, wav, dur in segs: fl.write(f"file '{img}'\nduration {dur:.3f}\n")
    fl.write(f"file '{segs[-1][0]}'\n")
with open(f"{OUT}/alist.txt", "w") as fa:
    for img, wav, dur in segs: fa.write(f"file '{wav}'\n")
# audio: cada wav seguido de 0.30 s de silencio
subprocess.run(["ffmpeg", "-y", "-loglevel", "error", "-f", "lavfi", "-i", "anullsrc=r=22050:cl=mono", "-t", "0.20", f"{OUT}/sil.wav"], check=True)
with open(f"{OUT}/alist.txt", "w") as fa:
    for img, wav, dur in segs: fa.write(f"file '{wav}'\nfile '{OUT}/sil.wav'\n")
subprocess.run(["ffmpeg", "-y", "-loglevel", "error", "-f", "concat", "-safe", "0", "-i", f"{OUT}/alist.txt", "-ar", "22050", "-ac", "1", f"{OUT}/audio.wav"], check=True)
subprocess.run(["ffmpeg", "-y", "-loglevel", "error", "-f", "concat", "-safe", "0", "-i", f"{OUT}/list.txt", "-i", f"{OUT}/audio.wav",
                "-vf", "scale=1280:720,fps=10,format=yuv420p", "-c:v", "libx264", "-preset", "veryfast", "-crf", "30", "-tune", "stillimage",
                "-c:a", "aac", "-b:a", "96k", "-shortest", f"{OUT}/Video_Explicacion_Proyecto_Argos.mp4"], check=True)
tot = sum(s[2] for s in segs); print("segmentos", len(segs), "duración s", round(tot), "min", round(tot / 60, 1))
