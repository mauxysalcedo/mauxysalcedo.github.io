import numpy as np, matplotlib, os
os.chdir("/home/user/mauxysalcedo.github.io/proyecto_final")
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import modelo as M
NAVY="#0B1F3A"; GOLD="#F2B705"; GRAY="#8A94A6"; RED="#C0392B"; GREEN="#2E8B57"
plt.rcParams.update({"font.size":10,"axes.spines.top":False,"axes.spines.right":False,"axes.edgecolor":"#999","figure.dpi":150})
ES=["ene","feb","mar","abr","may","jun","jul","ago","sep","oct","nov","dic"]
fmt=lambda d:f"{ES[d.month-1]}-{d.strftime('%y')}"
import matplotlib.dates as mdates
from matplotlib.ticker import FuncFormatter
ESF=FuncFormatter(lambda x,p:fmt(mdates.num2date(x)))
# 1 pronóstico
fig,ax=plt.subplots(figsize=(8,3.6))
ax.plot(M.H_LABEL,M.HIST_ARG,color=NAVY,lw=2,label="Demanda Argos (reconstruida)")
ax.plot(M.H_LABEL[13:],M.f_hw[13:],color=GOLD,lw=1.6,ls="--",label="Holt-Winters (1 paso)")
ax.plot(M.FUT_LABEL,M.FUT_ARG,color=GREEN,lw=2,label="Pronóstico oct-26 a nov-27")
ax.axvspan(M.PLAN_LABEL[0],M.PLAN_LABEL[-1],color=GOLD,alpha=.12,label="Horizonte de planeación")
ax.xaxis.set_major_formatter(ESF); ax.set_ylabel("kt de cemento/mes"); ax.legend(frameon=False,fontsize=8,ncol=2,loc="upper left"); ax.set_ylim(300,560)
fig.tight_layout(); fig.savefig("fig/f1_pronostico.png"); plt.close()
# 2 errores
fig,ax=plt.subplots(figsize=(6,3))
names=["Prom. móvil\n3 meses","Susv. exp.\nsimple","Holt-Winters"]; names[1]="Suav. exp.\nsimple"
vals=[r[2] for r in M.RES]
b=ax.bar(names,vals,color=[GRAY,GRAY,NAVY])
for r,v in zip(b,vals): ax.text(r.get_x()+r.get_width()/2,v+.1,f"{v:.1f}%",ha="center")
ax.set_ylabel("MAPE (%)"); ax.set_ylim(0,max(vals)*1.25); fig.tight_layout(); fig.savefig("fig/f2_mape.png"); plt.close()
# 3 utilización
fig,ax=plt.subplots(figsize=(8,3.6))
x=np.arange(12); w=.27
for i,(k,c) in enumerate(zip(M.REQ,(NAVY,GOLD,GRAY))):
    ax.bar(x+(i-1)*w,M.REQ[k]/M.CAPE[k]*100,w,label=k,color=c)
ax.axhline(100,color=RED,lw=1.5,ls="--"); ax.text(11.6,101,"100 %",color=RED,ha="right",fontsize=8)
ax.set_xticks(x); ax.set_xticklabels([fmt(d) for d in M.PLAN_LABEL],rotation=45); ax.set_ylim(70,108)
ax.set_ylabel("Utilización de la capacidad efectiva (%)"); ax.legend(frameon=False,ncol=3,fontsize=8,loc="upper left")
fig.tight_layout(); fig.savefig("fig/f3_util.png"); plt.close()
# 4 plan agregado
fig,axs=plt.subplots(1,2,figsize=(9,3.5),gridspec_kw={"width_ratios":[1.7,1]})
ax=axs[0]; ax.plot(M.PLAN_LABEL,M.DEM,color=NAVY,lw=2,label="Demanda")
for p,c,ls in zip(M.PLANES,(GRAY,GOLD,GREEN),(":","--","-")): ax.step(M.PLAN_LABEL,p["P"],where="mid",color=c,ls=ls,lw=2,label=p["nombre"].split(" (")[0])
ax.axhline(M.CAP_CEM_REG,color=RED,lw=1,ls="--"); ax.text(M.PLAN_LABEL[0],M.CAP_CEM_REG+3,"Capacidad regular 488,9",color=RED,fontsize=8)
ax.xaxis.set_major_formatter(ESF); ax.set_ylabel("kt/mes"); ax.legend(frameon=False,fontsize=7,ncol=2,loc="lower right"); ax.tick_params(axis="x",rotation=45)
ax=axs[1]; tot=[(p["total"]-M.PLANES[0]["total"])/1e6 for p in M.PLANES]
b=ax.bar(["A","B","C"],tot,color=[GRAY,GOLD,GREEN])
for r,v in zip(b,tot): ax.text(r.get_x()+r.get_width()/2,v-250 if v<0 else 30,f"{v:,.0f}",ha="center",va="top" if v<0 else "bottom",fontsize=8)
ax.axhline(0,color="#666",lw=.8); ax.set_ylim(-9500,1500); ax.set_ylabel("Costo vs. nivelación (MM COP)")
fig.tight_layout(); fig.savefig("fig/f4_agregado.png"); plt.close()
# 5 MPS
rows=M.MPS_ROWS; wk=[r[0] for r in rows]
fig,ax=plt.subplots(figsize=(8,3.3))
ax.bar(wk,[r[3]/1000 for r in rows],color=GOLD,label="MPS (campaña)"); ax.plot(wk,[r[2]/1000 for r in rows],color=NAVY,marker="o",label="Pedidos firmes")
ax.plot(wk,[r[4]/1000 for r in rows],color=GREEN,lw=2,label="Inventario proyectado"); ax.set_xlabel("Semana"); ax.set_ylabel("miles de t"); ax.legend(frameon=False,fontsize=8)
ax.set_xticks(wk); fig.tight_layout(); fig.savefig("fig/f5_mps.png"); plt.close()
# 6 TOC
fig,axs=plt.subplots(1,2,figsize=(8.5,3.2))
ax=axs[0]; ps=list(M.TOC); ax.bar(ps,[M.TOC[p]["T_clk"]/1000 for p in ps],color=[GREEN,GOLD,GRAY])
for i,p in enumerate(ps): ax.text(i,M.TOC[p]["T_clk"]/1000,f"{M.TOC[p]['T_clk']/1000:.0f}",ha="center",va="bottom",fontsize=8)
ax.set_ylabel("Trúput por t de clínker (miles COP)"); ax.tick_params(axis="x",labelsize=8)
ax=axs[1]; ax.bar(range(12),M.CLK_REQ,color=NAVY); ax.axhline(M.CLK_CAP,color=RED,ls="--"); ax.text(0,M.CLK_CAP+2,"Capacidad horno 366,7",color=RED,fontsize=8)
ax.set_ylim(280,390); ax.set_xticks(range(12)); ax.set_xticklabels([fmt(d) for d in M.PLAN_LABEL],rotation=60,fontsize=7); ax.set_ylabel("Clínker requerido (kt/mes)")
fig.tight_layout(); fig.savefig("fig/f6_toc.png"); plt.close()
# 7 flujo de coherencia
fig,ax=plt.subplots(figsize=(9,1.7)); ax.axis("off")
steps=["Supuestos","Pronóstico\nde demanda","Capacidad","Plan\nagregado","MPS / ATP","Tablero de\nindicadores"]
for i,s in enumerate(steps):
    ax.text(i*1.5,0,s,ha="center",va="center",fontsize=9,color="white",fontweight="bold",bbox=dict(boxstyle="round,pad=.5",fc=NAVY if i else GOLD,ec="none"))
    if i<len(steps)-1: ax.annotate("",xy=(i*1.5+.72,0),xytext=(i*1.5+.78,0),arrowprops=dict(arrowstyle="-|>",color=GRAY))
ax.set_xlim(-.8,8.3); ax.set_ylim(-.6,.6); fig.tight_layout(); fig.savefig("fig/f7_flujo.png"); plt.close()
# 8 diagrama de proceso
fig,ax=plt.subplots(figsize=(9,1.8)); ax.axis("off")
st=["Cantera y\nmateria prima","Horno de\nclínker","Molienda de\ncemento\n(+ adiciones)","Ensacado\n(saco 50 kg)","Despacho\nsaco / granel"]
for i,s in enumerate(st):
    hot=i==1
    ax.text(i*2,0,s,ha="center",va="center",fontsize=8.5,color="white" if hot else NAVY,fontweight="bold",bbox=dict(boxstyle="round,pad=.5",fc=RED if hot else "#E8ECF3",ec="none"))
    if i<len(st)-1: ax.annotate("",xy=(i*2+1.0,0),xytext=(i*2+.78,0),arrowprops=dict(arrowstyle="-|>",color=GRAY))
ax.text(2,-.62,"Etapa más cargada (cuello de botella)",color=RED,ha="center",fontsize=8,style="italic")
ax.set_xlim(-1,9); ax.set_ylim(-.9,.7); fig.tight_layout(); fig.savefig("fig/f8_proceso.png"); plt.close()
# 9 tablero (captura)
from openpyxl import load_workbook
wb=load_workbook("Proyecto_Argos_Modelo_y_Tablero.xlsx",data_only=True); tb=wb["Tablero"]
meses=[tb.cell(row=r,column=1).value for r in range(7,19)]
def col(c): return [tb.cell(row=r,column=c).value for r in range(7,19)]
fig,axs=plt.subplots(2,2,figsize=(9,5))
axs[0,0].plot(range(12),[v*100 for v in col(10)],color=NAVY,marker="o"); axs[0,0].axhline(5,color=GREEN,ls="--"); axs[0,0].set_title("MAPE del pronóstico (%) · meta ≤5",fontsize=9)
axs[0,1].plot(range(12),[v*100 for v in col(11)],color=NAVY,marker="o"); axs[0,1].axhline(85,color=GREEN,ls="--"); axs[0,1].axhline(100,color=RED,ls=":"); axs[0,1].set_title("Utilización del horno (%) · meta 85-100",fontsize=9)
axs[1,0].plot(range(12),[v*100 for v in col(12)],color=NAVY,marker="o",label="MPS"); axs[1,0].plot(range(12),[v*100 for v in col(13)],color=GOLD,marker="s",label="Fill rate"); axs[1,0].axhline(97,color=GREEN,ls="--"); axs[1,0].legend(frameon=False,fontsize=7); axs[1,0].set_title("Cumplimiento MPS y fill rate (%) · meta ≥97/98",fontsize=9)
axs[1,1].plot(range(12),col(14),color=NAVY,marker="o"); axs[1,1].axhline(.74,color=GREEN,ls="--"); axs[1,1].set_title("Factor clínker (ESG) · meta ≤0,74",fontsize=9)
for a in axs.ravel(): a.set_xticks(range(12)); a.set_xticklabels([ES[m.month-1] for m in meses],fontsize=7)
fig.tight_layout(); fig.savefig("fig/f9_tablero.png"); plt.close()
print("figs ok")
