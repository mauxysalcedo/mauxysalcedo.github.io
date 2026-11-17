"""Modelo numérico del proyecto: Cementos Argos - cemento gris Colombia.
Todos los supuestos viven en SUP; el resto de cálculos se alimenta de ellos.
"""
import numpy as np
from datetime import date

# ------------------------------------------------------------------ SUPUESTOS
SUP = dict(
    # --- datos públicos (verificar) ---
    nac_2024=12247.2, nac_2025=12911.2,     # DANE ECG, miles t despachos nacionales
    nac_ene_jul_2026=7555.3, nac_jun_2026=1061.6, nac_sep_2026=1194.4,
    argos_col_2025=5000.0,                  # miles t cemento regional Colombia (Informe de Gestión Grupo Argos 2025)
    # --- supuestos de capacidad (declarados) ---
    clinker_design=5000.0,    # miles t clínker/año (horno, todas las plantas Colombia) SUPUESTO
    disp_horno=0.88,          # disponibilidad (paros programados y no programados)
    factor_clinker=0.75,      # t clínker / t cemento
    bag_design=4800.0,        # miles t/año capacidad de empaque (ensacado) SUPUESTO
    bag_eff=0.85,             # eficiencia ensacadora (OEE simplificado)
    saco_share=0.65,          # proporción en saco
    mill_design=6800.0,       # miles t/año molienda cemento SUPUESTO
    mill_eff=0.90,
    # --- costos (COP por tonelada) SUPUESTOS ---
    precio=560_000, costo_var=330_000, costo_prod=400_000,
    inv_anual=0.22,           # costo de mantener inventario (% anual del costo)
    sobrecosto_extra=60_000,  # COP/t sobre capacidad regular (compra de clínker y molienda adicional: ~0,75 t x ~80.000 COP/t de sobreprecio)
    costo_cambio=20_000,      # COP por t de cambio de nivel mes a mes (arranques/paradas del horno, ajuste de turnos)
    extra_cap=0.08,           # capacidad adicional con turno extra (+8 %)
    max_inv_meses=1.5,        # vida útil/almacenaje máx. de cemento terminado (meses de demanda)
    inv_ini=150.0,            # miles t cemento terminado inicial (Dic-2026)
    # --- mezcla de productos Argos Colombia ---
    mix={"UG saco": 0.45, "Estructural saco": 0.20, "Granel": 0.35},
    precio_prod={"UG saco": 560_000, "Estructural saco": 600_000, "Granel": 470_000},
    cvar_prod={"UG saco": 330_000, "Estructural saco": 355_000, "Granel": 300_000},
    fclk_prod={"UG saco": 0.72, "Estructural saco": 0.85, "Granel": 0.72},
)
S = SUP
S["share"] = S["argos_col_2025"] / S["nac_2025"]          # participación de mercado
S["saco_share"] = S["mix"]["UG saco"] + S["mix"]["Estructural saco"]

# ---------------------------------------------------- C2 serie histórica (DANE ECG, oficial)
import json, os
_d = json.load(open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "datos_dane_ecg_2024_2026.json")))["serie"]
HIST_NAC = np.array([x[2] for x in _d])                    # miles t, ene-24 .. sep-26
H_LABEL = [date(x[0], x[1], 1) for x in _d]
SEAS = None
HIST_ARG = HIST_NAC * S["share"]                 # demanda Argos Colombia

def mape(a, f): return float(np.mean(np.abs((a - f) / a)) * 100)
def mad(a, f): return float(np.mean(np.abs(a - f)))

def ma(y, k=3):
    f = np.full(len(y), np.nan)
    for t in range(k, len(y)): f[t] = y[t - k:t].mean()
    return f

def ses(y, a):
    f = np.full(len(y), np.nan); f[1] = y[0]
    for t in range(2, len(y)): f[t] = f[t - 1] + a * (y[t - 1] - f[t - 1])
    return f

def hw_fit(y, m, a, b, g, h=0):
    """Holt-Winters multiplicativo. Devuelve pronósticos 1-paso (in-sample desde 2 temporadas) y h futuros."""
    n = len(y)
    L = y[:m].mean(); T = (y[m:2 * m].mean() - y[:m].mean()) / m
    sea = list(y[:m] / L)
    f1 = np.full(n, np.nan)
    for t in range(m, n):
        f1[t] = (L + T) * sea[t - m]
        Ln = a * (y[t] / sea[t - m]) + (1 - a) * (L + T)
        T = b * (Ln - L) + (1 - b) * T
        sea.append(g * (y[t] / Ln) + (1 - g) * sea[t - m]); L = Ln
    fut = [(L + (i + 1) * T) * sea[n - m + (i % m)] for i in range(h)]
    return f1, np.array(fut)

Y = HIST_ARG.copy()
n = len(Y); m = 12
EVAL = slice(m + 1, n)                      # evaluación común (ene-25 en adelante)
best = None
for a in np.arange(0.1, 0.9, 0.1):
    for b in (0.05, 0.1, 0.2):
        for g in np.arange(0.1, 0.9, 0.2):
            f1, _ = hw_fit(Y, m, a, b, g)
            e = mape(Y[EVAL], f1[EVAL])
            if best is None or e < best[0]: best = (e, a, b, g)
HW_A, HW_B, HW_G = best[1:]
f_hw, FUT_ARG = hw_fit(Y, m, HW_A, HW_B, HW_G, h=14)   # oct-26 .. nov-27
f_ma = ma(Y, 3)
best_s = min(((mape(Y[EVAL], ses(Y, a)[EVAL]), a) for a in np.arange(0.1, 0.95, 0.1)))
SES_A = best_s[1]; f_ses = ses(Y, SES_A)

def tracking(a, f):
    e = (a - f)[EVAL]; return float(np.sum(e) / np.mean(np.abs(e)))

METODOS = {
    "Promedio móvil 3 meses": (f_ma, ),
    f"Suavización exponencial simple (α={SES_A:.1f})": (f_ses, ),
    f"Holt-Winters multiplicativo (α={HW_A:.1f}, β={HW_B:.2f}, γ={HW_G:.1f})": (f_hw, ),
}
RES = []
for k, (f,) in METODOS.items():
    RES.append((k, mad(Y[EVAL], f[EVAL]), mape(Y[EVAL], f[EVAL]), tracking(Y, f)))
MEJOR = min(RES, key=lambda r: r[2])

FUT_ARG = np.round(FUT_ARG, 2)
FUT_LABEL = [date(2026 + ((9 + i) // 12), (9 + i) % 12 + 1, 1) for i in range(14)]
# horizonte de planeación: dic-26 .. nov-27 (índices 2..13 del futuro)
PLAN_LABEL = FUT_LABEL[2:14]
DEM = FUT_ARG[2:14].copy()                  # miles t, demanda Argos Colombia

# ------------------------------------------------------------ C3 capacidad
def cap():
    c = {}
    c["horno_dis_mes"] = S["clinker_design"] / 12
    c["horno_efec_mes"] = c["horno_dis_mes"] * S["disp_horno"]
    c["horno_equiv_cem"] = c["horno_efec_mes"] / S["factor_clinker"]    # miles t cemento/mes
    c["mol_dis_mes"] = S["mill_design"] / 12
    c["mol_efec_mes"] = c["mol_dis_mes"] * S["mill_eff"]
    c["bag_dis_mes"] = S["bag_design"] / 12
    c["bag_efec_mes"] = c["bag_dis_mes"] * S["bag_eff"]
    c["bag_equiv_cem"] = c["bag_efec_mes"] / S["saco_share"]            # cemento total que soporta el ensacado
    return c
C = cap()
REQ = {
    "Horno (clínker)": DEM * S["factor_clinker"], "Molienda": DEM.copy(),
    "Ensacado": DEM * S["saco_share"],
}
CAPE = {"Horno (clínker)": C["horno_efec_mes"], "Molienda": C["mol_efec_mes"], "Ensacado": C["bag_efec_mes"]}
CAPD = {"Horno (clínker)": C["horno_dis_mes"], "Molienda": C["mol_dis_mes"], "Ensacado": C["bag_dis_mes"]}
UTIL_AVG = {k: REQ[k].mean() / CAPE[k] for k in REQ}
UTIL_MAX = {k: REQ[k].max() / CAPE[k] for k in REQ}
RESTR = max(UTIL_AVG, key=UTIL_AVG.get)
# cemento equivalente que soporta la restricción (kt/mes)
CAP_CEM_REG = min(C["horno_equiv_cem"], C["mol_efec_mes"], C["bag_equiv_cem"])
CAP_CEM_EXT = CAP_CEM_REG * (1 + S["extra_cap"])

# ---------------------------------------------------------- C4 plan agregado
ci = S["costo_prod"]; h_mes = ci * S["inv_anual"] / 12       # COP/t-mes
margen = S["precio"] - S["costo_var"]
def evalplan(P, nombre):
    P = np.round(np.asarray(P, dtype=float), 1)
    inv = S["inv_ini"]; rows = []; prev = P[0]
    tot = dict(prod=0, extra=0, inv=0, cambio=0, falt=0)
    for t in range(12):
        d = DEM[t]
        disp = inv + P[t]
        venta = min(d, disp); falt = d - venta
        inv = disp - venta
        extra = max(0, P[t] - CAP_CEM_REG)
        cambio = abs(P[t] - (P[t - 1] if t else prev))
        rows.append(dict(mes=PLAN_LABEL[t], dem=d, prod=P[t], inv=inv, falt=falt, extra=extra))
        tot["prod"] += P[t] * 1000 * S["costo_var"]; tot["extra"] += extra * 1000 * S["sobrecosto_extra"]
        tot["inv"] += inv * 1000 * h_mes; tot["cambio"] += cambio * 1000 * S["costo_cambio"]
        tot["falt"] += falt * 1000 * S["precio"]   # ingreso perdido (el costo de producir esa t no se incurre)
    tot["total"] = sum(tot.values())
    return dict(nombre=nombre, P=np.array(P), rows=rows, **tot)

def plan_nivelado():
    # producción constante que cubre la demanda anual y deja el inventario final = inicial
    P = np.full(12, DEM.sum() / 12)
    return evalplan(np.minimum(P, CAP_CEM_EXT), "A. Nivelación")

def plan_persecucion():
    P = np.minimum(DEM, CAP_CEM_EXT)
    return evalplan(P, "B. Persecución")

def plan_mixto():
    # dos niveles de producción (a hasta el mes k, b después) que minimizan el costo total
    # y terminan el año con al menos el inventario inicial (inventario de seguridad)
    best = None
    for k in range(1, 12):
        for a in np.arange(400, CAP_CEM_EXT + 1, 4):
            for b in np.arange(400, CAP_CEM_EXT + 1, 4):
                P = np.array([a] * k + [b] * (12 - k))
                pl = evalplan(P, "C. Mixta (dos niveles)")
                if pl["rows"][-1]["inv"] + 1e-9 < S["inv_ini"]: continue
                if best is None or pl["total"] < best["total"]: best = pl
    return best

def planes_con_cap(cap_reg):
    global CAP_CEM_REG, CAP_CEM_EXT
    o = (CAP_CEM_REG, CAP_CEM_EXT)
    CAP_CEM_REG = cap_reg; CAP_CEM_EXT = cap_reg * (1 + S["extra_cap"])
    try: r = [plan_nivelado(), plan_persecucion(), plan_mixto()]
    finally: CAP_CEM_REG, CAP_CEM_EXT = o
    return r

PLANES = [plan_nivelado(), plan_persecucion(), plan_mixto()]
MEJOR_PLAN = min(PLANES, key=lambda p: p["total"])

# --------------------------------------------------- C5 MPS (UG saco, 12 sem)
REF = "Cemento Uso General saco 50 kg"
REF_SHARE = S["mix"]["UG saco"]
MPS_MESES = 3                                  # dic-26, ene-27, feb-27 (4 sem c/u)
agg = MEJOR_PLAN["P"][:MPS_MESES]
ref_mes = agg * REF_SHARE * 1000               # t/mes de la referencia
pron_sem = np.repeat(DEM[:MPS_MESES] * REF_SHARE * 1000 / 4, 4)   # pronóstico semanal (t)
lote = np.repeat(ref_mes / 2, 2)               # campaña quincenal (2 por mes)
MPS_PROD = np.zeros(12)
for i in range(6): MPS_PROD[2 * i] = lote[i]   # semanas 1,3,5,7,9,11
rng2 = np.random.default_rng(7)
PEDIDOS = np.round(pron_sem * np.array([1.05, 1.02, 0.95, 0.90, 0.80, 0.60, 0.45, 0.35, 0.25, 0.15, 0.10, 0.05])
                   * (1 + rng2.normal(0, 0.02, 12)), 0)        # pedidos firmes (simulados)
INV0 = S["inv_ini"] * 1000 * REF_SHARE
def mps_table():
    onhand = INV0; rows = []
    for w in range(12):
        # saldo proyectado = anterior + MPS - max(pronóstico, pedidos)
        req = max(pron_sem[w], PEDIDOS[w])
        oh = onhand + MPS_PROD[w] - req
        rows.append([w + 1, pron_sem[w], PEDIDOS[w], MPS_PROD[w], oh])
        onhand = oh
    # ATP
    atp = [None] * 12
    mps_w = [w for w in range(12) if MPS_PROD[w] > 0]
    for k, w in enumerate(mps_w):
        nxt = mps_w[k + 1] if k + 1 < len(mps_w) else 12
        base = (INV0 if k == 0 else 0) + MPS_PROD[w]
        atp[w] = base - PEDIDOS[w:nxt].sum()
    return rows, atp
MPS_ROWS, ATP = mps_table()

# ------------------------------------------------------------ C6 TOC
def toc():
    res = {}
    for p in S["mix"]:
        T = S["precio_prod"][p] - S["cvar_prod"][p]                    # trúput por t
        T_clk = T / S["fclk_prod"][p]                                  # por t de clínker (horno)
        res[p] = dict(T=T, T_clk=T_clk, share=S["mix"][p])
    return res
TOC = toc()
# demanda de clínker vs capacidad horno con la mezcla
CLK_REQ = DEM * sum(S["mix"][p] * S["fclk_prod"][p] for p in S["mix"])
CLK_CAP = C["horno_efec_mes"]
# Explotar: +2 pp de disponibilidad del horno
EXPL_PP = 0.02
clk_extra_mes = S["clinker_design"] / 12 * EXPL_PP                       # kt clínker/mes
cem_extra_anual = clk_extra_mes * 12 / S["factor_clinker"]
best_prod = max(TOC, key=lambda p: TOC[p]["T_clk"])
# Valor de explotar: kt cemento adicionales × T medio
T_med = sum(S["mix"][p] * TOC[p]["T"] for p in S["mix"])
VAL_EXPL = cem_extra_anual * 1000 * T_med
# Sustituir factor clínker 0.75 -> 0.72 (ESG + libera horno)
FC_NEW = 0.72
lib_clk = DEM.sum() * (S["factor_clinker"] - FC_NEW)                     # kt clínker liberadas / año
lib_cem = lib_clk / FC_NEW                                               # kt cemento extra posibles

# Escenarios de recomendación (C8)
def _sc(cap_reg):
    pl = min(planes_con_cap(cap_reg), key=lambda p: p["total"]); return pl
ESC_BASE = MEJOR_PLAN
ESC_DISP = _sc(CAP_CEM_REG * (S["disp_horno"] + EXPL_PP) / S["disp_horno"])        # +2 pp disponibilidad
ESC_FC = _sc(C["horno_efec_mes"] / FC_NEW)                                          # factor clínker 0,72
CO2_CLK = 0.85   # t CO2 / t clínker (referencia sectorial, VERIFICAR)
CO2_AHORRO = lib_clk * CO2_CLK                                                      # kt CO2/año

if __name__ == "__main__":
    print("share", round(S["share"], 4))
    print("hist nac total 2024/25", HIST_NAC[:12].sum(), HIST_NAC[12:24].sum(), "2026 ene-jul", HIST_NAC[24:31].sum(), "sep", HIST_NAC[32])
    for r in RES: print(r)
    print("mejor", MEJOR[0])
    print("DEM", np.round(DEM, 1), DEM.sum())
    print("CAP", {k: round(v, 1) for k, v in C.items()})
    print("UTIL avg", {k: round(v, 3) for k, v in UTIL_AVG.items()}, "max", {k: round(v, 3) for k, v in UTIL_MAX.items()})
    print("RESTR", RESTR, "CAP_CEM_REG", CAP_CEM_REG, "EXT", CAP_CEM_EXT)
    for p in PLANES:
        print(p["nombre"], {k: round(p[k] / 1e6) for k in ("prod", "extra", "inv", "cambio", "falt", "total")})
        print("  P", np.round(p["P"], 0), "falt", round(sum(r["falt"] for r in p["rows"]), 1))
    print("MEJOR", MEJOR_PLAN["nombre"])
    print("MPS", [(r[0], round(r[1]), r[2], round(r[3]), round(r[4])) for r in MPS_ROWS])
    print("ATP", ATP)
    print("TOC", {k: (round(v["T"]), round(v["T_clk"])) for k, v in TOC.items()}, "mejor", best_prod)
    print("ESC base/disp/fc", [round(e["total"]/1e6) for e in (ESC_BASE, ESC_DISP, ESC_FC)], [e["nombre"] for e in (ESC_BASE, ESC_DISP, ESC_FC)], "co2", CO2_AHORRO)
    print("  extra base/disp/fc", [round(sum(r["extra"] for r in e["rows"]),1) for e in (ESC_BASE, ESC_DISP, ESC_FC)])
    print("CLK_REQ max/cap", CLK_REQ.max(), CLK_CAP, "VAL_EXPL M", VAL_EXPL / 1e6, "lib_cem", lib_cem)
