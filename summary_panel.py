# -*- coding: utf-8 -*-
"""
Rediseno de la lamina de RESUMEN (nivel directorio), compartido por ambos reportes.
Genera un panel de alta resolucion (heroe + 4 cards + tabla grande) con la paleta
del deck y lo coloca sobre el cuerpo de la lamina, sin tocar cabecera, logo ni pie.
La tabla tiene mas presencia (mitad inferior a lo ancho, filas altas, barra de avance).
"""
import io, re
import matplotlib; matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch, Rectangle
from defusedxml.minidom import parseString

NAVY="#123A5E"; NAVY2="#1C4E80"; GREEN="#8CA61E"; ORANGE="#E08A1E"; BLUE="#103A5E"
GREY="#7A8A99"; CARDBG="#F5F8FB"; ALT="#F7F9FB"; BARBG="#E4E9EF"; REDC="#C0392B"
EMU=914400

def _logro_color(p):  return GREEN if p>=99.9 else (ORANGE if p>=90 else REDC)
def agg_frac(rows):
    """% agregado = avance total / meta-a-la-fecha total (coincide con la lámina GENERAL).
    Incluye filas con avance 0 (usa meta_avance real; si falta, back-calcula avance/logro)."""
    ta=sum((r.get("avance") or 0) for r in rows)
    def _ma(r):
        if r.get("meta_avance") is not None: return r["meta_avance"] or 0
        if r.get("logro") and r.get("avance") is not None and r["logro"]>0: return r["avance"]/r["logro"]
        return 0
    tm=sum(_ma(r) for r in rows)
    return (ta/tm) if tm>0 else float("nan")
def kfmt(v):
    v=float(v)
    if v>=1_000_000: return f"{v/1e6:.2f}M"
    if v>=100_000:   return f"{v/1e3:.1f}K"
    if v>=1000:      return f"{v/1e3:.2f}K"
    return f"{v:.0f}"

# ---------- lectura del contexto existente ----------
def _shapes(doc, Win):
    out=[]
    for sp in doc.getElementsByTagName("p:sp"):
        off=sp.getElementsByTagName("a:off")
        x=y=None
        if off:
            try: x=int(off[0].getAttribute("x"))/EMU; y=int(off[0].getAttribute("y"))/EMU
            except: pass
        txt=" ".join(t.firstChild.nodeValue for t in sp.getElementsByTagName("a:t") if t.firstChild)
        out.append((x,y,txt))
    return out

def _val_below(shapes, keywords):
    """valor (numero) de la card cuyo label contiene alguna keyword."""
    for x,y,txt in shapes:
        up=txt.strip().upper()
        if x is None or y is None: continue
        if any(k in up for k in keywords):
            best=None; bd=9
            for x2,y2,t2 in shapes:
                if x2 is None or y2 is None or not t2.strip(): continue
                if abs(x2-x)<0.6 and 0.05<(y2-y)<0.8 and re.search(r"\d", t2):
                    if (y2-y)<bd: best,bd=t2.strip(),(y2-y)
            if best is not None: return best
    return None

def read_table_staffing(slide_xml):
    """Lee Presupuesto y Conectados por fila desde la tabla original (no vienen del Excel)."""
    doc=parseString(slide_xml); sh=_shapes(doc,None)
    colx={}
    for x,y,txt in sh:
        if x is None: continue
        h=txt.strip().lower()
        if h=="presupuesto": colx["pres"]=x
        elif h=="conectados": colx["con"]=x
    def val(cx,ry):
        if cx is None: return None
        best=None; bdx=9
        for x2,y2,t2 in sh:
            if x2 is None or y2 is None or not t2.strip(): continue
            if abs(x2-cx)<0.6 and abs(y2-ry)<0.18 and re.search(r"\d",t2):
                if abs(x2-cx)<bdx: best,bdx=t2.strip(),abs(x2-cx)
        return best
    out={}
    for x,y,txt in sh:
        if x is None or y is None or x>=3.5 or y<=3.5: continue
        name=txt.strip()
        if not name or "campa" in name.lower(): continue
        out[name.lower()]={"presupuesto":val(colx.get("pres"),y),
                           "conectados":val(colx.get("con"),y)}
    return out

def read_context(slide_xml):
    doc=parseString(slide_xml)
    sh=_shapes(doc, None)
    status=""; subtitle=""
    for x,y,txt in sh:
        t=txt.strip()
        if re.match(r"^EN\s+\w+", t, re.I) and len(t)<30 and not status: status=t
        if re.search(r"(gerent|l[ií]der)", t, re.I) and not subtitle: subtitle=t
    activos=_val_below(sh, ["ACTIVOS","CONECTADOS","DOTACION"])
    sel=_val_below(sh, ["SELECC"]); ind=_val_below(sh, ["INDUCC"])
    def num(s):
        try: return int(re.sub(r"[^\d]","",s))
        except: return None
    pend=None
    if sel is not None or ind is not None:
        pend=(num(sel) or 0)+(num(ind) or 0)
    return {"status":status or "EN SEGUIMIENTO", "subtitle":subtitle or "",
            "activos":activos or "-", "pend_ingreso":(str(pend) if pend is not None else "-")}

# ---------- render del panel ----------
def render_panel(region_in, avg_frac, ctx, cumplen_txt, proy_total, rows, card4=None, hdr_color=NAVY):
    from matplotlib.patches import Wedge, Circle
    W,H=region_in
    fig=plt.figure(figsize=(W,H),dpi=300); ax=fig.add_axes([0,0,1,1])
    ax.set_xlim(0,W); ax.set_ylim(0,H); ax.invert_yaxis(); ax.axis("off")
    ax.add_patch(Rectangle((0,0),W,H,fc="white"))
    def rbox(x,y,w,h,fc,r=0.06):
        ax.add_patch(FancyBboxPatch((x,y),w,h,boxstyle=f"round,pad=0,rounding_size={r}",fc=fc,ec="none",mutation_aspect=1))
    def T(x,y,s,sz,c,w="normal",ha="left",va="center",st="normal"):
        ax.text(x,y,s,fontsize=sz,color=c,fontweight=w,ha=ha,va=va,style=st)

    top_h=H*0.47
    # ----- Avance promedio: numero grande + barra + estado (umbral) -----
    def nivel(p):
        if p>=99.9: return ("META ALCANZADA", GREEN, "#EAF0D2")
        if p>=90:   return ("CERCA DE META", ORANGE, "#FBEED6")
        return ("BAJO META", REDC, "#F7DDDA")
    nv_name, nv_col, nv_bg = nivel(avg_frac*100 if avg_frac==avg_frac else 0)
    hx=0.34
    T(hx,top_h*0.16,"AVANCE PROMEDIO",8.4,GREY,"bold")
    T(hx,top_h*0.41,((f"{avg_frac*100:.0f}%" if avg_frac>=1 else f"{avg_frac*100:.1f}%") if avg_frac==avg_frac else "-"),41,nv_col,"bold",va="center")
    _tav=sum((r.get("avance") or 0) for r in rows)
    # meta a la fecha total = suma de la meta_avance REAL de cada fila (incluye filas con avance 0,
    # p.ej. una plaza sin ventas todavía). Solo si no viene meta_avance se back-calcula avance/logro.
    def _row_ma(r):
        if r.get("meta_avance") is not None: return r["meta_avance"] or 0
        if r.get("logro") and r.get("avance") is not None and r["logro"]>0: return r["avance"]/r["logro"]
        return 0
    _tma=sum(_row_ma(r) for r in rows)
    T(hx,top_h*0.55,f"Meta: {kfmt(_tma)}  |  Avance: {kfmt(_tav)}",8.3,GREY,"bold")
    bx=hx; bw=2.02; by=top_h*0.63
    ax.add_patch(FancyBboxPatch((bx,by),bw,0.16,boxstyle="round,pad=0,rounding_size=0.08",fc=BARBG,ec="none",mutation_aspect=1))
    frp=max(0.03,min(1.0,avg_frac if avg_frac==avg_frac else 0))
    ax.add_patch(FancyBboxPatch((bx,by),bw*frp,0.16,boxstyle="round,pad=0,rounding_size=0.08",fc=nv_col,ec="none",mutation_aspect=1))
    pt=top_h*0.76
    rbox(hx,pt,2.02,0.27,nv_bg,0.13); ax.add_patch(Circle((hx+0.20,pt+0.135),0.058,color=nv_col))
    T(hx+0.36,pt+0.135,nv_name,8.2,nv_col,"bold")
    if ctx.get("subtitle"): T(hx,pt+0.45,ctx["subtitle"],7.6,GREY,st="italic")
    ax.add_patch(Rectangle((2.62,0.16),0.013,top_h*0.80,fc="#E1E7EE"))

    # ----- 4 cards -----
    def kf(v): return kfmt(v)
    if card4 is None:
        card4=("PROYECCIÓN TOTAL",f"S/ {kf(proy_total)}","cierre estimado",ORANGE,BLUE)
    if len(card4)==4: card4=tuple(card4)+(BLUE,)
    cards=[("ACTIVOS TOTALES",str(ctx.get("activos","-")),"colaboradores",NAVY,BLUE),
           ("PEND. INGRESO",str(ctx.get("pend_ingreso","-")),"selección + inducción",NAVY2,BLUE),
           ("CUMPLEN META",cumplen_txt,"unidades ≥ 100%",GREEN,BLUE),
           card4]
    edit_fields=[]   # campos manuales -> texto nativo editable (no se dibujan en la imagen)
    def dflt(v):     # valor por defecto editable cuando falta el dato manual
        s=str(v).strip(); return s if (s and s not in ("-","None")) else "0"
    x0=2.86; gap=0.13; cw=(W-x0-0.20-gap*3)/4; cy=0.16; ch=top_h*0.66
    for i,(lab,val,sub,acc,vcol) in enumerate(cards):
        x=x0+i*(cw+gap); rbox(x,cy,cw,ch,CARDBG,0.07)
        ax.add_patch(FancyBboxPatch((x+0.12,cy+0.13),cw-0.24,0.065,boxstyle="round,pad=0,rounding_size=0.03",fc=acc,ec="none",mutation_aspect=1))
        T(x+0.16,cy+ch*0.30,lab,7.6,GREY,"bold")
        if i in (0,1):   # Activos totales / Pend. ingreso -> editable
            edit_fields.append((x+0.14,cy+ch*0.58-0.19,cw-0.26,0.38,dflt(val),16.5,vcol.lstrip("#"),"l"))
        else:
            T(x+0.16,cy+ch*0.58,str(val),16.5,vcol,"bold")
        T(x+0.16,cy+ch*0.86,sub,7,GREY,st="italic")

    # ----- TABLA (con fila TOTAL) -----
    tx=0.20; tw=W-0.40; ty=top_h+0.04
    avail=H-ty-0.05; hh=0.42; nr=len(rows)+1; rh=(avail-hh)/max(nr,1)
    cols=["SUB-CAMPAÑA","PRESUP.","CONECT.","AVANCE","% LOGRO AVANCE","PROYECCIÓN","META MES"]
    cxf=[0.02,0.275,0.385,0.505,0.645,0.80,0.955]; cx=[tx+f*tw for f in cxf]
    al=["left","center","center","center","center","right","right"]
    ax.add_patch(FancyBboxPatch((tx,ty),tw,hh,boxstyle="round,pad=0,rounding_size=0.04",fc=hdr_color,ec="none",mutation_aspect=1))
    for c,x,a in zip(cols,cx,al): T(x,ty+hh/2,c,9,"white","bold",ha=a)
    def lc(p): return GREEN if p>=99.9 else (ORANGE if p>=90 else REDC)
    def lc_light(p): return "#C4DA6E" if p>=99.9 else ("#F0BC80" if p>=90 else "#E6A39C")
    def toint(v):
        try: return int(re.sub(r"[^\d-]","",str(v)))
        except: return 0
    y=ty+hh; sump=sumc=suma=sumpr=summm=0
    for i,r in enumerate(rows):
        if i%2==0: ax.add_patch(Rectangle((tx,y),tw,rh,fc=ALT))
        T(cx[0],y+rh/2,r["name"],10,BLUE,"bold")
        # Presupuesto y Conectados -> editables (manuales)
        edit_fields.append((cx[1]-0.6,y+rh/2-0.15,1.2,0.30,dflt(r.get("presupuesto")),9.3,"4A5A6A","ctr"))
        edit_fields.append((cx[2]-0.6,y+rh/2-0.15,1.2,0.30,dflt(r.get("conectados")),9.3,"4A5A6A","ctr"))
        if r.get("presupuesto"): sump+=toint(r["presupuesto"])
        if r.get("conectados"):  sumc+=toint(r["conectados"])
        suma+=r["avance"] or 0; sumpr+=r["proy"] or 0
        bx=cx[3]-0.62; bw=1.24
        ax.add_patch(FancyBboxPatch((bx,y+rh/2-0.082),bw,0.164,boxstyle="round,pad=0,rounding_size=0.08",fc=BARBG,ec="none",mutation_aspect=1))
        p=(r["logro"] or 0)*100; frb=max(0.05,min(1.0,(r["logro"] or 0)))
        ax.add_patch(FancyBboxPatch((bx,y+rh/2-0.082),bw*frb,0.164,boxstyle="round,pad=0,rounding_size=0.08",fc=lc_light(p),ec="none",mutation_aspect=1))
        T(cx[3],y+rh/2,kfmt(r["avance"]),8.4,"#20303F","bold",ha="center")
        if r["logro"] is not None:
            rbox(cx[4]-0.33,y+rh/2-0.105,0.66,0.21,lc(p),0.10)
            T(cx[4],y+rh/2,(f"{p:.0f}%" if p>=100 else f"{p:.1f}%"),8.8,"white","bold",ha="center")
        else: T(cx[4],y+rh/2,"-",9,GREY,"bold",ha="center")
        T(cx[5],y+rh/2,kfmt(r["proy"]),9.3,GREY,ha="right",st="italic")
        mm=r.get("meta_mes") or 0; summm+=mm
        T(cx[6],y+rh/2,kfmt(mm),9.0,"#4A5A6A","bold",ha="right")
        y+=rh
    # TOTAL
    ax.add_patch(Rectangle((tx,y),tw,rh,fc="#EAF0F6"))
    T(cx[0],y+rh/2,"TOTAL / PROMEDIO",9.3,NAVY,"bold")
    edit_fields.append((cx[1]-0.6,y+rh/2-0.15,1.2,0.30,str(sump),9.3,"123A5E","ctr"))
    edit_fields.append((cx[2]-0.6,y+rh/2-0.15,1.2,0.30,str(sumc),9.3,"123A5E","ctr"))
    T(cx[3],y+rh/2,kfmt(suma),9.0,NAVY,"bold",ha="center")
    if avg_frac==avg_frac:
        rbox(cx[4]-0.33,y+rh/2-0.105,0.66,0.21,(GREEN if (avg_frac==avg_frac and avg_frac>=0.999) else (ORANGE if (avg_frac==avg_frac and avg_frac>=0.90) else REDC)),0.10)
        T(cx[4],y+rh/2,(f"{avg_frac*100:.0f}%" if avg_frac>=1 else f"{avg_frac*100:.1f}%"),8.8,"white","bold",ha="center")
    T(cx[5],y+rh/2,kfmt(sumpr),9.3,NAVY,"bold",ha="right")
    T(cx[6],y+rh/2,kfmt(summm),9.3,NAVY,"bold",ha="right")
    ax.add_patch(FancyBboxPatch((tx,ty),tw,hh+rh*nr,boxstyle="round,pad=0,rounding_size=0.04",fill=False,ec="#D8E0E8",lw=1.1,mutation_aspect=1))
    buf=io.BytesIO(); fig.savefig(buf,dpi=300,facecolor="white"); plt.close(fig); buf.seek(0)
    return buf.getvalue(), edit_fields

def body_region(files, slide_num):
    pres=files["ppt/presentation.xml"].decode()
    m=re.search(r'sldSz\s+cx="(\d+)"\s+cy="(\d+)"',pres)
    Wemu,Hemu=int(m.group(1)),int(m.group(2))
    doc=parseString(files[f"ppt/slides/slide{slide_num}.xml"].decode())
    hb=0; ft=Hemu
    for sp in doc.getElementsByTagName("p:sp"):
        off=sp.getElementsByTagName("a:off"); ext=sp.getElementsByTagName("a:ext")
        txt=" ".join(t.firstChild.nodeValue for t in sp.getElementsByTagName("a:t") if t.firstChild).lower()
        if off and ext:
            try:
                y=int(off[0].getAttribute("y")); cx=int(ext[0].getAttribute("cx")); cy=int(ext[0].getAttribute("cy"))
            except: continue
            if y < Hemu*0.15 and cx > Wemu*0.5: hb=max(hb,y+cy)
            if "tu meta es nuestra meta" in txt: ft=min(ft,y)
    if hb < int(0.35*EMU): hb=int(0.77*EMU)          # cabecera agrupada/no detectada -> alto estandar
    if ft >= Hemu: ft=Hemu-int(0.26*EMU)                 # pie no detectado -> margen estandar
    top=hb+int(0.015*EMU); bottom=ft-int(0.015*EMU)
    return (0,top,Wemu,bottom-top)

def insert_panel(files, slide_num, region_emu, png_bytes):
    nums=[int(re.search(r'image(\d+)\.',n).group(1)) for n in files if re.match(r'ppt/media/image\d+\.',n)]
    idx=(max(nums)+1) if nums else 1
    files[f"ppt/media/image{idx}.png"]=png_bytes
    relp=f"ppt/slides/_rels/slide{slide_num}.xml.rels"; rels=files[relp].decode()
    rids=[int(m) for m in re.findall(r'Id="rId(\d+)"',rels)]; rid=f"rId{(max(rids)+1) if rids else 1}"
    rels=rels.replace("</Relationships>",
        f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image{idx}.png"/></Relationships>')
    files[relp]=rels.encode()
    sp=f"ppt/slides/slide{slide_num}.xml"; xml=files[sp].decode()
    ids=[int(m) for m in re.findall(r'id="(\d+)"',xml)]; nid=(max(ids)+1) if ids else 100
    x,y,cx,cy=region_emu
    pic=(f'<p:pic><p:nvPicPr><p:cNvPr id="{nid}" name="Panel Resumen"/>'
         f'<p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>'
         f'<p:blipFill><a:blip r:embed="{rid}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>'
         f'<p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
         f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>')
    files[sp]=xml.replace("</p:spTree>", pic+"</p:spTree>").encode()

def insert_text_fields(files, slide_num, region_emu, fields):
    """Inserta cuadros de texto NATIVOS (editables) sobre el panel para los datos manuales."""
    if not fields: return
    rx,ry,_,_=region_emu
    sp=f"ppt/slides/slide{slide_num}.xml"; xml=files[sp].decode("utf-8")
    ids=[int(m) for m in re.findall(r'id="(\d+)"',xml)]; nid=(max(ids)+1) if ids else 500
    blocks=[]
    for (xin,yin,win,hin,text,size,color,algn) in fields:
        x=int(rx+xin*EMU); y=int(ry+yin*EMU); cx=int(win*EMU); cy=int(hin*EMU)
        sz=int(round(size*100)); t=(text or "").replace("&","&amp;").replace("<","&lt;").replace(">","&gt;")
        blocks.append(
            f'<p:sp><p:nvSpPr><p:cNvPr id="{nid}" name="dato_manual"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr/></p:nvSpPr>'
            f'<p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
            f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>'
            f'<p:txBody><a:bodyPr wrap="none" lIns="0" tIns="0" rIns="0" bIns="0" anchor="ctr"/><a:lstStyle/>'
            f'<a:p><a:pPr algn="{algn}"/><a:r><a:rPr lang="es-PE" sz="{sz}" b="1" dirty="0">'
            f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill><a:latin typeface="Arial"/></a:rPr>'
            f'<a:t>{t}</a:t></a:r></a:p></p:txBody></p:sp>')
        nid+=1
    files[sp]=xml.replace("</p:spTree>","".join(blocks)+"</p:spTree>").encode("utf-8")

def pct_color(p):
    return "#8CA61E" if p>=99.9 else ("#E08A1E" if p>=90 else "#C0392B")
def estado_text(p):
    return "META ALCANZADA" if p>=99.9 else ("CERCA DE LA META" if p>=90 else "BAJO META")
_PILL_LIGHT={"#8CA61E":"#EAF0D2","#E08A1E":"#FBEED6","#C0392B":"#F7DDDA"}
def _set_fill(sp, hexhex):
    spPr=sp.getElementsByTagName("p:spPr")
    if not spPr: return
    for ch in spPr[0].childNodes:
        if ch.nodeType==1 and ch.tagName=="a:solidFill":
            clrs=ch.getElementsByTagName("a:srgbClr")
            if clrs: clrs[0].setAttribute("val",hexhex.lstrip("#"))
            return
    doc=sp.ownerDocument; sf=doc.createElement("a:solidFill"); clr=doc.createElement("a:srgbClr")
    clr.setAttribute("val",hexhex.lstrip("#")); sf.appendChild(clr); spPr[0].appendChild(sf)
def _set_line(sp, hexhex):
    spPr=sp.getElementsByTagName("p:spPr")
    if not spPr: return
    lns=[c for c in spPr[0].childNodes if c.nodeType==1 and c.tagName=="a:ln"]
    if not lns: return
    ln=lns[0]
    for sf in [c for c in ln.childNodes if c.nodeType==1 and c.tagName in ("a:solidFill","a:noFill")]:
        ln.removeChild(sf)
    doc=sp.ownerDocument; sf=doc.createElement("a:solidFill"); clr=doc.createElement("a:srgbClr")
    clr.setAttribute("val",hexhex.lstrip("#")); sf.appendChild(clr); ln.insertBefore(sf, ln.firstChild)
def recolor_pill(doc, p):
    """Pinta el pill de estado (top-left) para que TODO coincida con el % grande:
    punto y texto en color fuerte, fondo en color claro, y el texto = ESTADO
    (BAJO META / CERCA DE LA META / META ALCANZADA). Se ancla al punto (elipse)
    para ser robusto ante variaciones de posición y de texto ('META: X' o estado)."""
    strong=pct_color(p); light=_PILL_LIGHT[strong]; est=estado_text(p); E=914400
    def geom(sp):
        xfs=sp.getElementsByTagName("a:xfrm")
        if not xfs: return None
        xf=xfs[0]
        offs=[c for c in xf.childNodes if c.nodeType==1 and c.tagName=="a:off"]
        exts=[c for c in xf.childNodes if c.nodeType==1 and c.tagName=="a:ext"]
        if not offs or not offs[0].getAttribute("x"): return None
        try:
            x=int(offs[0].getAttribute("x"))/E; y=int(offs[0].getAttribute("y"))/E
            w=int(exts[0].getAttribute("cx"))/E if exts and exts[0].getAttribute("cx") else 0.0
            h=int(exts[0].getAttribute("cy"))/E if exts and exts[0].getAttribute("cy") else 0.0
        except: return None
        return x,y,w,h
    def prst(sp):
        g=sp.getElementsByTagName("a:prstGeom"); return g[0].getAttribute("prst") if g else ""
    def txt(sp):
        return "".join(t.firstChild.nodeValue for t in sp.getElementsByTagName("a:t") if t.firstChild).strip()
    def paint(sp, hexhex):
        for c in (sp.getElementsByTagName("p:spPr")[0].getElementsByTagName("a:srgbClr")
                  if sp.getElementsByTagName("p:spPr") else []): c.setAttribute("val",hexhex.lstrip("#"))
    sps=[(sp,geom(sp)) for sp in doc.getElementsByTagName("p:sp")]
    sps=[(sp,g) for sp,g in sps if g]
    # 1) localizar el punto (elipse pequeña, top-left)
    dot=None
    for sp,g in sps:
        x,y,w,h=g
        if prst(sp)=="ellipse" and x<2.85 and y<1.8 and w<0.5 and h<0.5: dot=(sp,g); break
    if not dot: return
    dsp,(dx,dy,dw,dh)=dot; dcx=dx+dw/2; dcy=dy+dh/2
    paint(dsp, strong)                                              # punto -> fuerte
    # 2) fondo: rect SIN texto que contiene el centro del punto -> claro
    for sp,g in sps:
        x,y,w,h=g
        if prst(sp) in ("rect","roundRect") and not txt(sp) and w<6 and x<=dcx<=x+w and y<=dcy<=y+h:
            paint(sp, light)
    # 3) texto: rect CON texto en la misma banda-y del punto -> ESTADO + color fuerte
    best=None
    for sp,g in sps:
        x,y,w,h=g
        if prst(sp) in ("rect","roundRect") and txt(sp) and w<6 and x<4.0 and abs((y+h/2)-dcy)<0.30:
            d=abs(x-dcx)
            if best is None or d<best[0]: best=(d,sp)
    if best:
        a_ts=[t for t in best[1].getElementsByTagName("a:t") if t.firstChild]
        if a_ts:
            a_ts[0].firstChild.nodeValue=est
            for t in a_ts[1:]: t.firstChild.nodeValue=""
            for t in a_ts: color_run(t, strong)
def color_run(t, hexhex):
    """Pinta el color de la fuente del run que contiene el <a:t> t."""
    r=t.parentNode
    if r is None: return
    doc=t.ownerDocument; rpr=None
    for ch in r.childNodes:
        if ch.nodeType==1 and ch.tagName=="a:rPr": rpr=ch; break
    if rpr is None:
        rpr=doc.createElement("a:rPr"); r.insertBefore(rpr, r.firstChild)
    for sf in [c for c in rpr.childNodes if c.nodeType==1 and c.tagName=="a:solidFill"]:
        rpr.removeChild(sf)
    sf=doc.createElement("a:solidFill"); clr=doc.createElement("a:srgbClr")
    clr.setAttribute("val", hexhex.lstrip("#")); sf.appendChild(clr)
    rpr.insertBefore(sf, rpr.firstChild)

def clean_panel(files, slide_num):
    """Quita panel(es) y campos manuales de una corrida previa (idempotencia)."""
    sp=f"ppt/slides/slide{slide_num}.xml"
    if sp not in files: return
    doc=parseString(files[sp].decode("utf-8")); removed=False
    for tag in ("p:pic","p:sp"):
        for el in list(doc.getElementsByTagName(tag)):
            nv=el.getElementsByTagName("p:cNvPr")
            if nv and nv[0].getAttribute("name") in ("Panel Resumen","dato_manual"):
                el.parentNode.removeChild(el); removed=True
    if removed: files[sp]=doc.toxml().encode("utf-8")

def clear_body_shapes(files, slide_num, top_in=1.0, bottom_in=7.0):
    """Elimina el contenido ANTIGUO del cuerpo del resumen (panel/cards/tabla de la plantilla),
    conservando la cabecera (y<top_in), el pie (y>=bottom_in) y el logo. Se llama tras leer
    los datos y antes de insertar el panel nuevo, para que no quede 'versión antigua' debajo."""
    sp=f"ppt/slides/slide{slide_num}.xml"
    if sp not in files: return
    doc=parseString(files[sp].decode("utf-8")); tree=doc.getElementsByTagName("p:spTree")[0]
    removed=False
    for el in list(tree.childNodes):
        if el.nodeType!=1 or el.tagName not in ("p:sp","p:pic","p:graphicFrame"): continue
        offs=el.getElementsByTagName("a:off")
        if not offs or not offs[0].getAttribute("y").strip(): continue   # sin posición -> conservar
        y=int(offs[0].getAttribute("y"))/EMU
        nv=el.getElementsByTagName("p:cNvPr")
        name=nv[0].getAttribute("name") if nv else ""
        if name in ("Panel Resumen","dato_manual"): continue             # elementos nuevos -> conservar
        if top_in <= y < bottom_in:                                       # cuerpo viejo -> eliminar
            tree.removeChild(el); removed=True
    if removed: files[sp]=doc.toxml().encode("utf-8")

def redesign_summary(files, slide_num, avg_frac, rows, proy_total, card4=None):
    clean_panel(files, slide_num)   # idempotente: elimina panel previo si el deck ya fue generado
    """rows: [{name, avance(float), logro(float|None), proy(float)}]  avg_frac: fraccion (0-1)."""
    slide_xml=files[f"ppt/slides/slide{slide_num}.xml"].decode()
    ctx=read_context(slide_xml)
    staffing=read_table_staffing(slide_xml)
    for r in rows:                      # adjuntar Presupuesto y Conectados leidos de la plantilla
        st=staffing.get(r["name"].lower(), {})
        r["presupuesto"]=st.get("presupuesto") or "0"; r["conectados"]=st.get("conectados") or "0"
    region=body_region(files, slide_num)
    region_in=(region[2]/EMU, region[3]/EMU)
    n=len(rows); cumple=sum(1 for r in rows if r["logro"] is not None and r["logro"]>=0.999)
    # Secciones MF (banda superior guinda #6B150B: SBP, Scotia Tarjetas, BCP, Efectiva, SANNA)
    # llevan la cabecera de la tabla del mismo rojo; PlusMetas se queda en azul (NAVY).
    hdr_color = "#6B150B" if 'val="6B150B"' in slide_xml else NAVY
    png,edit_fields=render_panel(region_in, avg_frac, ctx, f"{cumple} / {n}", proy_total, rows, card4, hdr_color=hdr_color)
    clear_body_shapes(files, slide_num)   # elimina el panel/tabla ANTIGUO de la plantilla (queda solo el panel nuevo)
    insert_panel(files, slide_num, region, png)
    insert_text_fields(files, slide_num, region, edit_fields)
