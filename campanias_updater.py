# -*- coding: utf-8 -*-
"""
Motor unificado del COMITÉ SEMANAL DE CAMPAÑAS (deck consolidado).
Cada unidad = portada + resumen + detalles + diagnóstico.
- Unidades de campaña (SBP, EFECTIVA, BANBIF, BBVA OUT/HIBRIDO, DINERS, UNICEF): Excel CALL_PRODUCTOS.
- Sección Convenios: Excel de Convenios (se delega en convenios_updater).
Config-driven: para agregar/renombrar detalles se edita DETAILS.
"""
import io, re, zipfile
from collections import Counter
import numpy as np, pandas as pd
import matplotlib; matplotlib.use("Agg"); import matplotlib.pyplot as plt
from matplotlib.patches import Patch
from PIL import Image
from defusedxml.minidom import parseString
import summary_panel
import convenios_updater as CV

# ---- estilo ----
TEAL,CYAN="#009090","#00CCCC"; NAVY,NAVY_L="#245490","#7FA3CC"; SLATE,GRAY="#3C5460","#787878"
MES_ABBR={1:"ENE",2:"FEB",3:"MAR",4:"ABR",5:"MAY",6:"JUN",7:"JUL",8:"AGO",9:"SEP",10:"OCT",11:"NOV",12:"DIC"}
MES_FULL={1:"Enero",2:"Febrero",3:"Marzo",4:"Abril",5:"Mayo",6:"Junio",7:"Julio",8:"Agosto",9:"Septiembre",10:"Octubre",11:"Noviembre",12:"Diciembre"}
EMU=914400
PCT_RE=re.compile(r"^\s*\d+([.,]\d+)?\s*%\s*$")
DATE_RE=re.compile(r"al\s+\d{1,2}\s+de\s+[A-Za-zÁÉÍÓÚáéíóú]+",re.I)
CIERRE_RE=re.compile(r"cierre:?\s+[A-Za-zÁÉÍÓÚáéíóú]+\s+\d{4}",re.I)
MONTH_RE=re.compile(r"(Enero|Febrero|Marzo|Abril|Mayo|Junio|Julio|Agosto|Septiembre|Setiembre|Octubre|Noviembre|Diciembre)\s+\d{4}",re.I)

# ===================== CONFIG detalle: (title_kw, unidad, subproductos, productos, titulo_grafico, display) =====================
DETAILS=[
 ("sanna","SANNA",["POLIZAS"],["SEGUROS"],"PÓLIZAS","SANNA Pólizas"),
 ("scotiabank tarjetas","SBP",["OUT"],["TARJETAS"],"APROBADAS","Scotiabank Tarjetas"),
 ("efectiva (dormidos","EFECTIVA",["DORMIDOS","NUEVOS","RECURRENTES"],["PRESTAMOS"],"DESEMBOLSADO","Efectiva (Dor+Nue+Rec)"),
 ("fonocompras","EFECTIVA",["FONOCOMPRAS"],["FONOCOMPRAS"],"DESEMBOLSADO","Fonocompras"),
 ("banbif prestamos","BANBIF",["NACIONAL"],["PRESTAMOS"],"DESEMBOLSADO","Banbif Préstamos"),
 ("tc telemarketing out","BBVA",["OUT"],["TARJETAS"],"FORMALIZADAS","TC Telemarketing OUT"),
 ("pld out prestamos","BBVA",["OUT"],["PRESTAMOS"],"OPERACIONES DESEMBOLSADAS","PLD OUT (Préstamos)"),
 ("operaciones out","BBVA",["OUT"],["OPERACIONES"],"FORMALIZADAS","Operaciones OUT"),
 ("portafolio","BBVA",["OUT"],["PORTAFOLIO"],"DESEMBOLSADO","Portafolio"),
 ("tc hibrido","BBVA",["HIBRIDO"],["TARJETAS"],"FORMALIZADAS","TC Híbrido (Tarjetas)"),
 ("operaciones digital","BBVA",["HIBRIDO"],["OPERACIONES"],"FORMALIZADAS","Operaciones Digital"),
 ("pld digital","BBVA",["HIBRIDO"],["PRESTAMOS"],"DESEMBOLSADO","PLD Digital"),
 ("tc respaldada","BBVA",["HIBRIDO"],["TARJETAS RESPALDA"],"FORMALIZADAS","TC Respaldada"),
 ("tarjetas de credito","DINERS",["-"],["TARJETAS"],"ACTIVADAS","Diners Tarjetas"),
 # UNICEF (subproductos; se usan también para clonar la sección)
 ("digital","UNICEF",["DIGITAL"],["DONACIONES"],"DONACIONES","UNICEF Digital"),
 ("extracash","UNICEF",["EXTRACASH"],["RETENCIONES"],"RETENCIONES","Extracash"),
 ("fidelizacion","UNICEF",["FIDELIZACION"],["RETENCIONES"],"RETENCIONES","Fidelización"),
 ("saving","UNICEF",["SAVING"],["RETENCIONES"],"RETENCIONES","Saving"),
 ("upgrade linea","UNICEF",["UPGRADE LINEA"],["RETENCIONES"],"RETENCIONES","Upgrade Línea"),
]
# ================================================================================================================

def kfmt(v):
    v=float(v)
    if v>=1_000_000: return f"{v/1e6:.2f}M"
    if v>=100_000:   return f"{v/1e3:.1f}K"
    if v>=1000:      return f"{v/1e3:.2f}K"
    return f"{v:.0f}"
def fmt_pct(f):
    v=float(f)*100; return f"{v:.0f}%" if v>=100 else f"{v:.1f}%"
def _norm(s): return str(s).strip().upper()

def load(excel_bytes):
    xls=pd.ExcelFile(io.BytesIO(excel_bytes)); need={"UNIDAD","PRODUCTO","SUBPRODUCTO","PERIODO","TIPO","MONTO"}
    df=None
    for sh in xls.sheet_names:
        t=xls.parse(sh); t=t.rename(columns={c:re.sub(r"\s+","",str(c).strip().upper()) for c in t.columns})
        if need.issubset(t.columns): df=t; break
    if df is None: raise ValueError("Falta hoja con UNIDAD/PRODUCTO/SUBPRODUCTO/PERIODO/TIPO/MONTO")
    for c in ("UNIDAD","PRODUCTO","SUBPRODUCTO","TIPO"): df[c]=df[c].map(_norm)
    df=df[df.PERIODO.notna()]; df["PERIODO"]=df.PERIODO.astype(int)
    df["MONTO"]=pd.to_numeric(df.MONTO,errors="coerce").fillna(0)
    dia=18
    if "DIA" in df.columns:
        v=pd.to_numeric(df.DIA,errors="coerce").dropna()
        if len(v): dia=int(v.mode().iloc[0])
    return df,int(df.PERIODO.max()),dia

def series(df,unidad,subs,prods,current):
    d=df[(df.UNIDAD==unidad)&(df.SUBPRODUCTO.isin([_norm(s) for s in subs]))&(df.PRODUCTO.isin([_norm(p) for p in prods]))]
    periods=sorted(d.PERIODO.unique())
    def m(tipo,per): return d[(d.TIPO==tipo)&(d.PERIODO==per)]["MONTO"].sum()
    avance=[m("AVANCE",p) for p in periods]
    pr_cur=m("PROYECCION",current); has_proj=pr_cur>0
    cierre=[(pr_cur if (p==current and has_proj) else m("CIERRE",p)) for p in periods]
    cur=periods.index(current) if current in periods else None
    av=m("AVANCE",current); ma=m("META_AVANCE",current)
    kpi={"avance":av,"meta_avance":ma,"meta_mes":m("META_MES",current),
         "proyeccion":(pr_cur if has_proj else m("CIERRE",current)),
         "logro":(av/ma if ma else np.nan)}
    return {"periods":periods,"labels":[MES_ABBR[int(str(p)[4:6])] for p in periods],
            "avance":avance,"cierre":cierre,"cur":cur,"kpi":kpi,"proj":has_proj,"has_data":len(periods)>0}

def build_chart_series(s,title,size_px,out):
    labels,avance,cierre,cur=s["labels"],s["avance"],s["cierre"],s["cur"]
    n=len(labels); xs=np.arange(n); W,H=size_px; dpi=100
    fig,ax=plt.subplots(figsize=(W/dpi,H/dpi),dpi=dpi); fig.patch.set_alpha(0); ax.set_facecolor("none")
    cmax=max(cierre) or 1; amax=max(avance) or 1; ax.set_ylim(0,cmax*1.34)
    barh=[a*((cmax*0.42)/amax) for a in avance]
    bcol=[CYAN if (cur is not None and i==cur) else TEAL for i in range(n)]
    ax.bar(xs,barh,width=0.42,color=bcol,zorder=3)
    for x,h,a in zip(xs,barh,avance): ax.text(x,h+cmax*0.02,kfmt(a),ha="center",va="bottom",color=SLATE,fontsize=40,fontweight="bold")
    proj=s.get("proj",True)
    if cur is not None and cur>0 and proj:
        ax.plot(xs[:cur+1],cierre[:cur+1],color=NAVY,lw=6,solid_capstyle="round",zorder=4)
        ax.plot(xs[cur-1:cur+1],cierre[cur-1:cur+1],color=NAVY_L,lw=6,ls=(0,(4,3)),zorder=5)
        ax.plot(xs[:cur],cierre[:cur],"o",color=NAVY,ms=11,zorder=6); ax.plot([xs[cur]],[cierre[cur]],"o",color=NAVY_L,ms=11,zorder=6)
    elif n>0:
        ax.plot(xs,cierre,color=NAVY,lw=6,marker="o",ms=11,solid_capstyle="round",zorder=4)
    for i,(x,c) in enumerate(zip(xs,cierre)):
        col=(NAVY_L if (cur is not None and i==cur and proj) else NAVY)
        ax.text(x,c+cmax*0.035,kfmt(c),ha="center",va="bottom",color=col,fontsize=42,fontweight="bold")
    ax.set_xticks(xs); ax.set_xticklabels(labels,color=GRAY,fontsize=42,fontweight="bold")
    ax.tick_params(axis="x",length=0,pad=14); ax.set_yticks([])
    for sp in ("top","left","right"): ax.spines[sp].set_visible(False)
    ax.spines["bottom"].set_color("#B0B0B0"); ax.spines["bottom"].set_linewidth(2); ax.set_xlim(-0.6,n-0.4)
    ax.text(-0.55,cmax*1.30,title,color=SLATE,fontsize=78,fontweight="bold",ha="left",va="top")
    ax.legend(handles=[Patch(color=SLATE,label="AVANCE"),Patch(color=NAVY,label="CIERRE")],loc="upper right",ncol=2,
              frameon=False,fontsize=46,handlelength=1.1,columnspacing=1.6,bbox_to_anchor=(1.0,0.99),labelcolor=[SLATE,NAVY])
    plt.subplots_adjust(left=0.01,right=0.99,top=0.99,bottom=0.09)
    fig.savefig(out,dpi=dpi,transparent=True); plt.close(fig)

# ---- helpers XML ----
def _off(sp):
    o=sp.getElementsByTagName("a:off")
    if not o: return (None,None)
    try: return (int(o[0].getAttribute("x"))/EMU,int(o[0].getAttribute("y"))/EMU)
    except: return (None,None)
def _txt(sp): return " ".join(t.firstChild.nodeValue for t in sp.getElementsByTagName("a:t") if t.firstChild)
def _set(t,v):
    if t.firstChild: t.firstChild.nodeValue=v
def _fix_dates(doc,mes,year,dia):
    for t in doc.getElementsByTagName("a:t"):
        if not t.firstChild: continue
        s=t.firstChild.nodeValue
        if DATE_RE.search(s): _set(t,DATE_RE.sub(f"al {dia} de {mes}",s)); s=t.firstChild.nodeValue
        if CIERRE_RE.search(s): _set(t,CIERRE_RE.sub(f"Cierre {mes} {year}",s)); s=t.firstChild.nodeValue
        if MONTH_RE.search(s): _set(t,MONTH_RE.sub(f"{mes} {year}",s))
def _title_kw_match(doc):
    for sp in doc.getElementsByTagName("p:sp"):
        tt=_txt(sp).lower()
        if "informaci" in tt or "meta avance" in tt or "meta:" in tt:
            hit,best=None,-1
            for i,e in enumerate(DETAILS):
                if e[0] in tt and len(e[0])>best: hit,best=i,len(e[0])
            if hit is not None: return hit
    # fallback: buscar en toda la lamina
    allt=" ".join(_txt(sp).lower() for sp in doc.getElementsByTagName("p:sp"))
    hit,best=None,-1
    for i,e in enumerate(DETAILS):
        if e[0] in allt and len(e[0])>best: hit,best=i,len(e[0])
    return hit
def _frame_aspect(slide_xml,rels_xml,image):
    rid2img=dict(re.findall(r'Id="([^"]+)"[^>]*Target="[^"]*media/([^"]+)"',rels_xml))
    doc=parseString(slide_xml)
    for pic in doc.getElementsByTagName("p:pic"):
        b=pic.getElementsByTagName("a:blip")
        if not b or rid2img.get(b[0].getAttribute("r:embed"))!=image: continue
        for x in pic.getElementsByTagName("a:xfrm"):
            e=x.getElementsByTagName("a:ext")
            if e:
                cx,cy=int(e[0].getAttribute("cx")),int(e[0].getAttribute("cy"))
                if cy: return cx/cy
    return None
def _strip_srcrect(doc,rels_xml,image):
    rid2img=dict(re.findall(r'Id="([^"]+)"[^>]*Target="[^"]*media/([^"]+)"',rels_xml))
    for pic in doc.getElementsByTagName("p:pic"):
        b=pic.getElementsByTagName("a:blip")
        if not b or rid2img.get(b[0].getAttribute("r:embed"))!=image: continue
        for sr in pic.getElementsByTagName("a:srcRect"): sr.parentNode.removeChild(sr)
def edit_detail_text(doc,kpi):
    lg=kpi.get("logro",np.nan)
    for t in doc.getElementsByTagName("a:t"):
        if not t.firstChild: continue
        s=t.firstChild.nodeValue; low=s.strip().lower()
        if low.startswith("meta avance"): _set(t,f"META AVANCE: {kfmt(kpi['meta_avance'])}")
        elif low.startswith("meta:") or low.startswith("meta "):
            if kpi["meta_mes"]: _set(t,f"META: {kfmt(kpi['meta_mes'])}")
        elif low.startswith("avance:"): _set(t,f"AVANCE: {kfmt(kpi['avance'])}")
        elif PCT_RE.match(s) and pd.notna(lg): _set(t,fmt_pct(lg)); summary_panel.color_run(t, summary_panel.pct_color(lg*100))
        elif ((re.match(r"^\s*EN\s+\w", s, re.I) and len(s.strip())<22) or s.strip().lower() in ("bajo la meta","cerca de la meta","meta alcanzada","bajo meta","cerca de meta","meta alcanzada")) and pd.notna(lg):
            _set(t, summary_panel.estado_text(lg*100)); summary_panel.color_run(t, summary_panel.pct_color(lg*100))
    if pd.notna(lg): summary_panel.recolor_pill(doc, lg*100)

def _is_donut(img_bytes):
    """La dona tiene azules saturados (R<110, B>150) de las familias; los asesores no."""
    try:
        im=Image.open(io.BytesIO(img_bytes)).convert("RGB").resize((60,60)); px=im.load()
    except: return False
    blue=sum(1 for x in range(60) for y in range(60)
             if px[x,y][0]<110 and px[x,y][2]>150)
    return blue>120

def _pick_chart(files,sid,slide_imgs,logos):
    """imagen del grafico: la ancha (aspect>=1.85) o la compuesta recortada."""
    cands=[i for i in slide_imgs.get(sid,[]) if i not in logos]
    wide=[]; comp=[]
    for im in cands:
        try: w,h=Image.open(io.BytesIO(files[f"ppt/media/{im}"])).size
        except: continue
        if w/h>=1.85: wide.append((im,(w,h)))
        elif 1.0<=w/h<1.3 and (w<1100): comp.append((im,(w,h)))  # 983x899 compuesta
    if wide: return max(wide,key=lambda t:t[1][0]*t[1][1])[0]
    if comp: return comp[0][0]
    return None

def update_presentation(template_bytes, camp_excel, conv_excel=None, dia_override=None):
    df,current,dia=load(camp_excel)
    if dia_override:  # día de corte elegido en la app -> todas las fechas "avance hasta el día X"
        try: dia=int(dia_override)
        except: pass
    year,month=int(str(current)[:4]),int(str(current)[4:6]); mes=MES_FULL[month]
    convdf=convcur=None
    if conv_excel is not None:
        convdf,convcur,_=CV.load_data(conv_excel)

    zin=zipfile.ZipFile(io.BytesIO(template_bytes)); files={n:zin.read(n) for n in zin.namelist()}; zin.close()
    refs=Counter(); slide_imgs={}
    for n in list(files):
        m=re.match(r"ppt/slides/_rels/(slide\d+)\.xml\.rels",n)
        if m:
            imgs=re.findall(r"media/(image[0-9]+\.\w+)",files[n].decode("utf-8")); slide_imgs[m.group(1)]=imgs
            for im in imgs: refs[im]+=1
    logo=refs.most_common(1)[0][0] if refs else None
    # logos = imágenes compartidas por muchas láminas (los charts se usan en 1 sola).
    # Así, aunque cambie el balance (p.ej. swap de logo a MF), nunca se sobreescribe un logo.
    LOGOS=set(im for im,c in refs.items() if c>=4)
    if logo: LOGOS.add(logo)
    warnings=[]

    # ordenar por ORDEN DE PRESENTACIÓN (sldIdLst -> rels -> archivo), no por nombre
    def _present_order():
        try:
            pres=files["ppt/presentation.xml"].decode("utf-8")
            prels=files["ppt/_rels/presentation.xml.rels"].decode("utf-8")
            rid2t=dict(re.findall(r'Id="([^"]+)"[^>]*Target="([^"]+)"', prels))
            ids=re.findall(r'<p:sldId[^>]*r:id="([^"]+)"', pres)
            out=[]
            for rid in ids:
                tgt=rid2t.get(rid,"")
                nm=tgt.split("/")[-1]
                p=f"ppt/slides/{nm}"
                if p in files: out.append(p)
            return out
        except Exception:
            return []
    slide_files=_present_order() or sorted([n for n in files if re.match(r"ppt/slides/slide\d+\.xml$",n)],
                       key=lambda x:int(re.findall(r"\d+",x.split("/")[-1])[0]))
    # pass 1: clasificar y guardar kpis de detalle por resumen
    order=[]; groups={}; cur_group=None
    for sf in slide_files:
        sid=sf.split("/")[-1][:-4]; num=int(re.findall(r"\d+",sid)[0])
        xml=files[sf].decode("utf-8"); low=xml.lower(); doc=parseString(xml)
        is_conv=("convenios" in low)
        if "resumen general" in low:
            cur_group=num; groups[num]={"conv":is_conv,"rows":[]}
            order.append((sf,sid,num,"resumen",is_conv,doc,cur_group)); continue
        if "meta avance" in low:
            order.append((sf,sid,num,"detail",is_conv,doc,cur_group)); continue
        order.append((sf,sid,num,"other",is_conv,doc,cur_group))

    # pass 2: procesar
    for sf,sid,num,kind,is_conv,doc,grp in order:
        rels=files.get(f"ppt/slides/_rels/{sid}.xml.rels",b"").decode("utf-8")
        # fechas siempre
        cdia = (CV_dia if False else dia)
        if kind=="detail":
            if is_conv and convdf is not None:
                bank,plaza=CV._title_bank_plaza(doc, convdf)
                _ok = ((convdf.UNIDAD==bank).any() if str(plaza).upper()=="GENERAL"
                       else ((convdf.UNIDAD==bank)&(convdf.PLAZA==plaza)).any())
                if bank and plaza and _ok:
                    s=CV.plaza_series(convdf,plaza,convcur,bank=bank); CV.edit_detail_text(doc,s["kpi"]); _fix_dates(doc,mes,year,dia)
                    fam=CV.plaza_familia(convdf,plaza,convcur,bank=bank); aso=CV.plaza_asesores(convdf,plaza,convcur,bank=bank)
                    rid2img=dict(re.findall(r'Id="([^"]+)"[^>]*Target="[^"]*media/([^"]+)"', rels))
                    frames={}
                    for pic in doc.getElementsByTagName("p:pic"):
                        b=pic.getElementsByTagName("a:blip"); img=rid2img.get(b[0].getAttribute("r:embed")) if b else None
                        if not img: continue
                        xf=pic.getElementsByTagName("a:xfrm")
                        if not xf: continue
                        ext=xf[0].getElementsByTagName("a:ext")
                        if not ext: continue
                        cx=int(ext[0].getAttribute("cx")); cy=int(ext[0].getAttribute("cy"))
                        if cy: frames[img]=(cx/EMU, cx/cy)
                    files[sf]=doc.toxml().encode("utf-8")
                    for im in [i for i in slide_imgs.get(sid,[]) if i not in LOGOS]:
                        if im not in frames: continue
                        w_in,asp=frames[im]
                        try:
                            if asp>=1.85:                                   # Avances del ámbito
                                buf=io.BytesIO(); CV.chart_avances(s,plaza,(2400,round(2400/asp)),buf,dia)
                            else:                                            # dona o asesores según CONTENIDO original
                                Wp=max(600,round(w_in*300)); buf=io.BytesIO()
                                if _is_donut(files[f"ppt/media/{im}"]):
                                    CV.chart_familia(fam,MES_ABBR[month],(Wp,round(Wp/asp)),buf)
                                else:
                                    CV.chart_asesores(aso,(Wp,round(Wp/asp)),buf)
                            files[f"ppt/media/{im}"]=buf.getvalue()
                        except Exception as _e:
                            warnings.append(f"{bank} · {plaza}: no se pudo regenerar un gráfico ({type(_e).__name__}: {_e}); se conservó la imagen original.")
                else:
                    _fix_dates(doc,mes,year,dia); files[sf]=doc.toxml().encode("utf-8")
                continue
            idx=_title_kw_match(doc)
            if idx is None:
                _fix_dates(doc,mes,year,dia); files[sf]=doc.toxml().encode("utf-8"); continue
            kw,unidad,subs,prods,ctitle,disp=DETAILS[idx]
            s=series(df,unidad,subs,prods,current)
            edit_detail_text(doc,s["kpi"]); _fix_dates(doc,mes,year,dia)
            im=_pick_chart(files,sid,slide_imgs,LOGOS)
            if im and s["has_data"]:
                _strip_srcrect(doc,rels,im)
            files[sf]=doc.toxml().encode("utf-8")
            if im and s["has_data"]:
                px=Image.open(io.BytesIO(files[f"ppt/media/{im}"])).size
                asp=_frame_aspect(files[sf].decode("utf-8"),rels,im) or (px[0]/px[1])
                Wp=4750; buf=io.BytesIO(); build_chart_series(s,ctitle,(Wp,max(600,round(Wp/asp))),buf)
                files[f"ppt/media/{im}"]=buf.getvalue()
                if px[0]/px[1]<1.6:
                    warnings.append(f"{disp}: imagen compuesta (panel KPIs no está en el Excel); se regeneró solo el gráfico principal.")
            # registrar en el grupo/resumen (por orden de presentación)
            if grp is not None and grp in groups:
                g=groups[grp]
                if not g["conv"]:
                    lg=s["kpi"]["logro"]
                    g["rows"].append({"name":disp,"key":disp,"avance":s["kpi"]["avance"],
                                      "logro":(float(lg) if pd.notna(lg) else None),
                                      "proy":(s["kpi"]["proyeccion"] if pd.notna(s["kpi"]["proyeccion"]) else 0),
                                      "meta_avance":s["kpi"]["meta_avance"]})
        else:
            _fix_dates(doc,mes,year,dia); files[sf]=doc.toxml().encode("utf-8")

    # pass 3: resúmenes
    for sf,sid,num,kind,is_conv,doc,grp in order:
        if kind!="resumen": continue
        if is_conv and convdf is not None:
            tt=" ".join(t.firstChild.nodeValue for t in doc.getElementsByTagName("a:t") if t.firstChild).lower()
            bank=next((b.upper() for b in ("bbva","bcp","sbp") if b in tt), "BBVA")
            dfb=convdf[convdf.UNIDAD==bank] if "UNIDAD" in convdf.columns else convdf
            rows=[]; logros=[]; pt=0.0
            for key in sorted(dfb.PLAZA.dropna().unique()):
                k=CV.plaza_series(convdf,key,convcur,bank=bank)["kpi"]; lg=k["logro"]
                rows.append({"name":str(key).title(),"avance":k["avance"],"logro":(float(lg) if pd.notna(lg) else None),"proy":k["proyeccion"],"meta_avance":k["meta_avance"]})
                if pd.notna(lg): logros.append(lg)
                pt+=k["proyeccion"]
            avg=summary_panel.agg_frac(rows)   # % agregado (avance/meta)
            summary_panel.redesign_summary(files,num,avg,rows,pt)   # card4 = Proyección total
        else:
            g=groups.get(num,{"rows":[]}); rows=g["rows"]
            if not rows: continue
            avg=summary_panel.agg_frac(rows)   # % agregado (avance/meta)
            proy_total=sum((r["proy"] or 0) for r in rows)
            card4=None   # 1 sola campaña -> Proyección total (sin S/, suele ser conteo de TC)
            if len(rows)<=1:
                card4=("PROYECCIÓN TOTAL", kfmt(proy_total), "cierre estimado", "#E08A1E", "#103A5E")
            if len(rows)>1:
                cand=[r for r in rows if r["logro"] is not None]
                if cand:
                    w=min(cand,key=lambda r:r["logro"]); p=w["logro"]*100
                    acc="#C0392B" if p<90 else "#E08A1E"
                    short=re.sub(r"\s*\(.*?\)","",w["name"]).strip()[:20]
                    card4=("A REFORZAR",(f"{p:.0f}%" if p>=100 else f"{p:.1f}%"),short,acc,acc)
            summary_panel.redesign_summary(files,num,avg,rows,proy_total,card4)

    # pasada final: fechas -> mes/día en TODAS las láminas (portadas, diags, etc.)
    for n in list(files):
        if not re.match(r"ppt/slides/slide\d+\.xml$", n): continue
        xml=files[n].decode("utf-8")
        if not (DATE_RE.search(xml) or CIERRE_RE.search(xml)): continue
        d2=parseString(xml); _fix_dates(d2,mes,year,dia); files[n]=d2.toxml().encode("utf-8")

    out=io.BytesIO()
    with zipfile.ZipFile(out,"w",zipfile.ZIP_DEFLATED) as z:
        for n,d in files.items(): z.writestr(n,d)
    out.seek(0)
    return out.getvalue(),{"periodo":current,"mes":f"{mes} {year}","dia":dia},warnings

def cur_group_for(num,groups):
    prev=[g for g in groups if g<=num]
    return max(prev) if prev else None
