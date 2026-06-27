#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Render the six BVD SitRep enhancement figures from a JSON data file.

All styling (WHO palette, 300 DPI, French number formatting, sizes) is fixed here so every
SitRep's figures look identical; you only supply the day's numbers. Adapt this script rather than
hand-rolling new styles for a one-off — cross-day consistency is the point.

Usage:
    python make_figures.py --example > day.json      # print a template to fill in
    python make_figures.py --data day.json --outdir fig/
"""
import argparse, json, sys
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np

BLUE='#0072BC'; DARK='#1A3A5C'; RED='#C8132A'; GREEN='#2E7D32'; AMBER='#F4A81D'; GREY='#D9D9D9'; ORANGE='#E65100'
plt.rcParams.update({'font.family':'DejaVu Sans','font.size':10,'axes.edgecolor':'#888','axes.linewidth':0.8})

def fr(x, dp=1):
    return (f'{x:.{dp}f}').replace('.', ',')

EXAMPLE = {
  "n": 41, "date": "24 juin 2026",
  "days": ["21 juin","22 juin","23 juin","24 juin"],
  "anchors_note": "Ancres vérifiées SitReps #038–#041",
  "daily": {"cases":[45,46,24,37], "deaths":[9,10,14,13], "recoveries":[None,3,7,16]},
  "cumulative": {"national":[1048,1094,1118,1155], "ituri":[954,997,1020,1054],
                 "nk":[91,94,95,98], "sk":[3,3,3,3], "cfr":[25.5,25.3,26.0,26.3]},
  "hz_daily": {"days":["22 juin","23 juin","24 juin"],
               "zones":{"Bunia":[11,8,10],"Rwampara":[10,0,12],"Mongbwalu":[8,4,4],
                        "Nyankunde":[0,0,5],"Mangala":[3,3,0],"Lita":[3,2,0],"Autres ZS":[11,7,6]},
               "national":[46,24,37]},
  "contacts": {"days":["22 juin","23 juin","24 juin"],
               "ituri":[72.7,77.9,78.1],"nk":[71.6,73.3,82.6],"national":[72.8,77.1,79.2],
               "under":[8278,8346,9305],"target":95},
  "nk_profile": {"zones":["Katwa","Butembo","Beni","Oicha","Kyondo","Kalunguta","Musienene","Vuhovi","Goma","Masereka","Mabalako"],
                 "cases":[33,31,20,3,3,2,2,1,1,1,1],"deaths":[20,13,12,2,2,1,2,1,0,0,0],
                 "cfr":[60.6,41.9,60.0,66.7,66.7,50.0,100.0,100.0,0.0,0.0,0.0],"mean_cfr":54.1},
  "lab": {"provinces":["Ituri","Nord-Kivu","Sud-Kivu"],"analysed":[134,97,6],"positives":[34,3,0],
          "backlog":[0,20,0],"positivity":[25.0,3.1,0.0],
          "nk_backlog_days":["22 juin","23 juin","24 juin"],"nk_backlog":[8,37,20]}
}

def save(fig, path):
    fig.savefig(path, dpi=300, bbox_inches='tight', facecolor='white'); plt.close(fig)

def fig1(d, out):
    days=d["days"]; x=np.arange(len(days)); dd=d["daily"]; w=0.26
    fig,ax=plt.subplots(figsize=(10,4.8))
    b1=ax.bar(x-w,dd["cases"],w,label='Cas confirmés',color=BLUE)
    b2=ax.bar(x,dd["deaths"],w,label='Décès',color=RED)
    rv=[0 if v is None else v for v in dd["recoveries"]]; b3=ax.bar(x+w,rv,w,label='Guérisons',color=GREEN)
    for b in b1: ax.annotate(f'{int(b.get_height())}',(b.get_x()+b.get_width()/2,b.get_height()),ha='center',va='bottom',fontsize=9,color=DARK,fontweight='bold')
    for b in b2: ax.annotate(f'{int(b.get_height())}',(b.get_x()+b.get_width()/2,b.get_height()),ha='center',va='bottom',fontsize=8,color=RED)
    for i,b in enumerate(b3):
        lab='ND' if dd["recoveries"][i] is None else f'{int(b.get_height())}'
        ax.annotate(lab,(b.get_x()+b.get_width()/2,b.get_height()),ha='center',va='bottom',fontsize=8,color=GREEN)
    ax.set_xticks(x); ax.set_xticklabels(days); ax.set_ylabel('Nombre par jour')
    top=max(dd["cases"]+dd["deaths"]+rv); ax.set_ylim(0,top*1.18)
    ax.legend(loc='upper right',frameon=False,ncol=3,fontsize=9)
    ax.set_title(f'Cas confirmés, décès et guérisons par jour — RDC, {days[0]}–{days[-1]}',fontsize=11,color=DARK,fontweight='bold')
    ax.annotate('⚠ Variation journalière à interpréter avec prudence (délais de notification).',xy=(0.01,0.97),xycoords='axes fraction',fontsize=8,color='#666',style='italic',va='top')
    for s in ['top','right']: ax.spines[s].set_visible(False)
    save(fig,f'{out}/fig1.png')

def fig2(d, out):
    days=d["days"]; x=np.arange(len(days)); c=d["cumulative"]
    fig,ax=plt.subplots(figsize=(10,4.8))
    ax.plot(x,c["national"],'-o',color=DARK,lw=2.4,label='Total RDC',zorder=5)
    ax.fill_between(x,c["ituri"],color=BLUE,alpha=0.12); ax.plot(x,c["ituri"],'-o',color=BLUE,lw=2,label='Ituri')
    ax.plot(x,c["nk"],'-s',color=RED,lw=1.8,label='Nord-Kivu'); ax.plot(x,c["sk"],'-^',color=AMBER,lw=1.5,label='Sud-Kivu')
    ax.annotate(f'{c["national"][-1]}',(x[-1],c["national"][-1]),xytext=(0,8),textcoords='offset points',ha='center',fontsize=10,fontweight='bold',color=DARK)
    ax.set_xticks(x); ax.set_xticklabels(days); ax.set_ylabel('Cas confirmés cumulés'); ax.set_ylim(0,max(c["national"])*1.15)
    ax2=ax.twinx(); ax2.plot(x,c["cfr"],'--D',color=ORANGE,lw=1.8,label='TL national (%)')
    ax2.set_ylabel('Taux de létalité (%)',color=ORANGE); ax2.tick_params(axis='y',colors=ORANGE); ax2.set_ylim(0,max(40,max(c["cfr"])*1.3))
    ax2.annotate(f'{fr(c["cfr"][-1])}%',(x[-1],c["cfr"][-1]),xytext=(-4,8),textcoords='offset points',ha='right',fontsize=9,color=ORANGE,fontweight='bold')
    l1,la1=ax.get_legend_handles_labels(); l2,la2=ax2.get_legend_handles_labels()
    ax.legend(l1+l2,la1+la2,loc='center left',frameon=False,fontsize=9)
    ax.set_title(f'Trajectoire cumulée des cas confirmés et taux de létalité — RDC, {days[0]}–{days[-1]}',fontsize=11,color=DARK,fontweight='bold')
    ax.spines['top'].set_visible(False); save(fig,f'{out}/fig2.png')

def fig3(d, out):
    h=d["hz_daily"]; days=h["days"]; x=np.arange(len(days))
    cols=['#003F7F',BLUE,'#4DA6E8','#87CEEB',AMBER,'#9CCB3B',GREY]
    fig,ax=plt.subplots(figsize=(10,4.8)); bottom=np.zeros(len(days))
    for (z,vals),col in zip(h["zones"].items(),cols+[GREY]*9):
        ax.bar(x,vals,0.55,bottom=bottom,label=z,color=col); bottom+=np.array(vals,dtype=float)
    ax.plot(x,h["national"],'--o',color=DARK,lw=2,label='Total national')
    for i,v in enumerate(h["national"]): ax.annotate(f'{v}',(x[i],v),xytext=(0,7),textcoords='offset points',ha='center',fontsize=10,fontweight='bold',color=DARK)
    ax.set_xticks(x); ax.set_xticklabels(days); ax.set_ylabel('Nouveaux cas confirmés'); ax.set_ylim(0,max(h["national"])*1.5)
    ax.legend(loc='upper right',frameon=False,ncol=2,fontsize=8)
    ax.set_title(f'Cas confirmés journaliers par zone de santé vs total national — RDC, {days[0]}–{days[-1]}',fontsize=11,color=DARK,fontweight='bold')
    for s in ['top','right']: ax.spines[s].set_visible(False)
    save(fig,f'{out}/fig3.png')

def fig4(d, out):
    c=d["contacts"]; days=c["days"]; x=np.arange(len(days))
    fig,ax=plt.subplots(figsize=(10,4.8))
    ax.bar(x,c["under"],0.5,color=GREY,alpha=0.5,label='Contacts sous suivi (n)',zorder=1)
    ax.set_ylabel('Contacts sous suivi (n)'); ax.set_ylim(0,max(c["under"])*1.3)
    ax2=ax.twinx()
    ax2.plot(x,c["ituri"],'-o',color=BLUE,lw=2,label='Ituri',zorder=5)
    ax2.plot(x,c["nk"],'-s',color=RED,lw=2,label='Nord-Kivu',zorder=5)
    ax2.plot(x,c["national"],'--D',color=DARK,lw=2.2,label='National',zorder=6)
    ax2.axhline(c["target"],ls=':',color=GREEN,lw=1.6); ax2.annotate(f'Cible {c["target"]}%',(0,c["target"]),xytext=(2,4),textcoords='offset points',color=GREEN,fontsize=8)
    gap=c["target"]-c["national"][-1]
    ax2.set_ylabel('Taux de suivi (%)'); ax2.set_ylim(40,100)
    ax2.annotate(f'{fr(c["national"][-1])}%',(x[-1],c["national"][-1]),xytext=(-4,-14),textcoords='offset points',ha='right',fontsize=9,color=DARK,fontweight='bold')
    ax2.annotate(f'↕ −{fr(gap)}pp vs cible',(x[-1],c["national"][-1]),xytext=(-6,18),textcoords='offset points',ha='right',fontsize=8,color='#666')
    ax.set_xticks(x); ax.set_xticklabels(days)
    l1,la1=ax.get_legend_handles_labels(); l2,la2=ax2.get_legend_handles_labels()
    ax2.legend(l1+l2,la1+la2,loc='lower left',frameon=False,fontsize=8,ncol=2)
    ax.set_title(f'Taux de suivi des contacts et contacts sous suivi — RDC, {days[0]}–{days[-1]}',fontsize=11,color=DARK,fontweight='bold')
    ax.spines['top'].set_visible(False); ax2.spines['top'].set_visible(False); save(fig,f'{out}/fig4.png')

def fig5(d, out):
    p=d["nk_profile"]; z=p["zones"]; xx=np.arange(len(z)); w=0.42
    fig,ax=plt.subplots(figsize=(11,5.0))
    ax.bar(xx-w/2,p["cases"],w,label='Cas confirmés',color=BLUE); ax.bar(xx+w/2,p["deaths"],w,label='Décès',color=RED)
    ax.set_xticks(xx); ax.set_xticklabels(z,rotation=35,ha='right'); ax.set_ylabel('Nombre (cumulés)'); ax.set_ylim(0,max(p["cases"])*1.25)
    for i,v in enumerate(p["cases"]): ax.annotate(f'{v}',(xx[i]-w/2,v),ha='center',va='bottom',fontsize=8,color=DARK)
    ax2=ax.twinx(); ax2.plot(xx,p["cfr"],'D',color=ORANGE,ms=6,label='TL (%)')
    ax2.axhline(p["mean_cfr"],ls=':',color=ORANGE,lw=1.4); ax2.annotate(f'TL NK moyen {fr(p["mean_cfr"])}%',(len(z)-1,p["mean_cfr"]),xytext=(-4,6),textcoords='offset points',ha='right',color=ORANGE,fontsize=8)
    ax2.set_ylabel('Taux de létalité (%)',color=ORANGE); ax2.set_ylim(0,110); ax2.tick_params(axis='y',colors=ORANGE)
    l1,la1=ax.get_legend_handles_labels(); l2,la2=ax2.get_legend_handles_labels()
    ax.legend(l1+l2,la1+la2,loc='upper right',frameon=False,fontsize=9)
    ax.set_title(f'Profil Nord-Kivu : cas confirmés, décès et TL par zone de santé — {d["date"]}',fontsize=11,color=DARK,fontweight='bold')
    ax.spines['top'].set_visible(False); save(fig,f'{out}/fig5.png')

def fig6(d, out):
    L=d["lab"]; provs=L["provinces"]; xx=np.arange(len(provs)); w=0.26
    fig,(axL,axR)=plt.subplots(1,2,figsize=(11,4.8),gridspec_kw={'width_ratios':[1.4,1]})
    axL.bar(xx-w,L["analysed"],w,label='Analysés',color=BLUE); axL.bar(xx,L["positives"],w,label='Positifs',color=DARK); axL.bar(xx+w,L["backlog"],w,label='Backlog',color=AMBER)
    for i in range(len(provs)):
        axL.annotate(f'{L["analysed"][i]}',(xx[i]-w,L["analysed"][i]),ha='center',va='bottom',fontsize=8,color=DARK)
        if L["backlog"][i]: axL.annotate(f'{L["backlog"][i]}',(xx[i]+w,L["backlog"][i]),ha='center',va='bottom',fontsize=9,color=AMBER,fontweight='bold')
    axL.set_xticks(xx); axL.set_xticklabels(provs); axL.set_ylabel('Échantillons'); axL.set_ylim(0,max(L["analysed"])*1.2); axL.legend(loc='upper right',frameon=False,fontsize=9)
    ax2=axL.twinx(); ax2.plot(xx,L["positivity"],'D--',color=ORANGE,lw=1.6); ax2.set_ylabel('Positivité (%)',color=ORANGE); ax2.set_ylim(0,max(50,max(L["positivity"])*1.5)); ax2.tick_params(axis='y',colors=ORANGE)
    for i,v in enumerate(L["positivity"]): ax2.annotate(f'{fr(v)}%',(xx[i],v),xytext=(0,6),textcoords='offset points',ha='center',fontsize=8,color=ORANGE)
    axL.set_title('Capacité diagnostique et backlog par province',fontsize=10,color=DARK,fontweight='bold'); axL.spines['top'].set_visible(False)
    bd=L["nk_backlog_days"]; axR.bar(np.arange(len(bd)),L["nk_backlog"],0.5,color=AMBER)
    for i,v in enumerate(L["nk_backlog"]): axR.annotate(f'{v}',(i,v),ha='center',va='bottom',fontsize=9,color=DARK,fontweight='bold')
    axR.set_xticks(np.arange(len(bd))); axR.set_xticklabels(bd); axR.set_ylabel('Échantillons en attente (NK)'); axR.set_ylim(0,max(L["nk_backlog"])*1.25)
    axR.set_title('Backlog laboratoire Nord-Kivu',fontsize=10,color=DARK,fontweight='bold')
    for s in ['top','right']: axR.spines[s].set_visible(False)
    save(fig,f'{out}/fig6.png')

def main():
    ap=argparse.ArgumentParser(description="Render the six BVD SitRep figures from a JSON data file.")
    ap.add_argument('--data'); ap.add_argument('--outdir',default='fig')
    ap.add_argument('--example',action='store_true',help='print a template JSON and exit')
    a=ap.parse_args()
    if a.example:
        print(json.dumps(EXAMPLE,ensure_ascii=False,indent=2)); return
    if not a.data: ap.error('--data is required (or use --example)')
    import os; os.makedirs(a.outdir,exist_ok=True)
    d=json.load(open(a.data,encoding='utf-8'))
    figs=[("Figure 1 (cas/décès/guérisons)",fig1),("Figure 2 (trajectoire+TL)",fig2),
          ("Figure 3 (cas par ZS)",fig3),("Figure 4 (suivi contacts)",fig4),
          ("Figure 5 (profil Nord-Kivu)",fig5),("Figure 6 (laboratoire)",fig6)]
    for i,(label,fn) in enumerate(figs,1):
        fn(d,a.outdir)
        print(f"  [{i}/6] {label} -> {a.outdir}/", flush=True)
    print(f"6 figures written to {a.outdir}/", flush=True)

if __name__=='__main__': main()
