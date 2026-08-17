# -*- coding: utf-8 -*-
"""
Motor de actualizacion del reporte BBVA CONVENIOS.
Estructura distinta a TLM: datos a nivel transaccion y 3 graficos por lamina
(Avances del ambito, Participacion por FAMILIA_CONVENIO, Asesores por Rango de Ventas).

Se actualizan las laminas 2 a 7 (2 = resumen, 3-7 = una por PLAZA).

SUPUESTOS (cambiables): ver notas al pie de load_data().
"""
import io, re, zipfile
from collections import Counter
import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import Patch
from matplotlib.lines import Line2D
from PIL import Image
from defusedxml.minidom import parseString
import summary_panel

# ===================== CONFIGURACION (unico lugar a tocar) =====================
# (clave_en_excel_PLAZA, texto_en_titulo_de_su_lamina, texto_en_fila_del_resumen)
PLAZAS = [
    ("LIMA",      "lima campo", "lima"),
    ("TELECAMPO", "telecampo",  "telecampo"),
    ("NORTE",     "norte",      "norte"),
    ("SUR",       "sur",        "sur"),
    ("ORIENTE",   "oriente",    "oriente"),
]
DIA_AVANCE = 11  # día por defecto; si el Excel trae una columna DIA, se usa ese valor
# familia -> color (dona). Familias no listadas usan colores de reserva.
FAMILY_COLORS = {
    "Fuerzas Armadas":     "#0060A0",
    "Gobierno Central":    "#2090C0",
    "Sector Salud":        "#50C0D0",
    "Sector Educación":    "#90C000",
    "Institución Pública": "#1F3B5B",
    "Institución Privada": "#7FB0D8",
    "Cajas y Cooperativas":"#B0C0E0",
}
FALLBACK_COLORS = ["#0060A0","#2090C0","#50C0D0","#90C000","#1F3B5B","#7FB0D8","#B0C0E0","#C0D000"]
# ==============================================================================

# estilo avances
BAR_AV, LINE_C = "#A0A0A0", "#90B0D0"     # barras Monto Avance / linea Cierre
LINE_PROJ      = "#B7C9DE"                # proyeccion (punteado, claro)
NAVY, REDLBL, RED = "#004060", "#A02020", "#F00000"
SLATE, TICKET  = "#405060", "#2E6CA6"
# estilo asesores
AS_BLUE, AS_GREY, AS_GREEN = "#90B0D0", "#808080", "#90C000"
DARK = "#203040"

MES_ABBR = {1:"ENE",2:"FEB",3:"MAR",4:"ABR",5:"MAY",6:"JUN",
            7:"JUL",8:"AGO",9:"SEP",10:"OCT",11:"NOV",12:"DIC"}
MES_FULL = {1:"Enero",2:"Febrero",3:"Marzo",4:"Abril",5:"Mayo",6:"Junio",
            7:"Julio",8:"Agosto",9:"Septiembre",10:"Octubre",11:"Noviembre",12:"Diciembre"}
EMU = 914400
PCT_RE  = re.compile(r"^\s*\d+([.,]\d+)?\s*%\s*$")
DATE_RE = re.compile(r"al\s+\d{1,2}\s+de\s+[A-Za-zÁÉÍÓÚáéíóú]+", re.IGNORECASE)
CIERRE_RE = re.compile(r"cierre\s+\w+\s+\d{4}", re.IGNORECASE)
TITLE_KW = {p[0]: p[1] for p in PLAZAS}
ROW_KW   = {p[0]: p[2] for p in PLAZAS}


# ---------- formato ----------
def kfmt(v):
    v = float(v)
    if v >= 1_000_000: return f"{v/1e6:.2f}M"
    if v >= 100_000:   return f"{v/1e3:.1f}K"
    if v >= 1000:      return f"{v/1e3:.2f}K"
    return f"{v:.0f}"
def fmt_pct(frac):
    v = float(frac)*100
    return f"{v:.0f}%" if v >= 100 else f"{v:.1f}%"
def fmt_mm(v):  return f"{v/1e6:.2f}MM"
def fmt_bar(v):
    if v == 0: return "0 und"
    return f"{v/1e6:.2f}MM" if v >= 1_000_000 else f"{v/1e3:.2f}k"
def fmt_ticket(v): return f"S/ {v/1e3:.2f}k"
def _norm(s): return str(s).strip().upper()


# ---------- datos ----------
def load_data(excel_bytes):
    xls = pd.ExcelFile(io.BytesIO(excel_bytes))
    df = None
    for sh in xls.sheet_names:
        t = xls.parse(sh)
        t = t.rename(columns={c: re.sub(r"\s+","",str(c).strip().upper()) for c in t.columns})
        cols = set(t.columns)
        # esquema NUEVO (crudo): UNIDAD/TERRITORIO/FAMILIA_CONVENIO/PERIODO/TIPO/MONTO
        # o esquema ANTIGUO: PLAZA/PERIODO/TIPO/MONTONETO
        if {"PERIODO","TIPO","MONTO","FAMILIA_CONVENIO"}.issubset(cols) or \
           {"PLAZA","PERIODO","TIPO","MONTONETO"}.issubset(cols):
            df = t; break
    if df is None:
        raise ValueError("No encontré la hoja de Convenios (esperaba TERRITORIO/FAMILIA_CONVENIO/PERIODO/TIPO/MONTO, "
                         "o el esquema antiguo PLAZA/PERIODO/TIPO/MONTONETO).")
    # normalizar nombres del esquema nuevo a los internos
    if "MONTONETO" not in df.columns and "MONTO" in df.columns:
        df = df.rename(columns={"MONTO":"MONTONETO"})
    if "DNI_ASESOR_VENTA" not in df.columns and "ID_EJECUTIVO" in df.columns:
        df = df.rename(columns={"ID_EJECUTIVO":"DNI_ASESOR_VENTA"})
    df["TIPO"]  = df["TIPO"].map(lambda x: _norm(x).replace(" ","_"))
    df = df[df["PERIODO"].notna()]; df["PERIODO"] = df["PERIODO"].astype(int)
    for c in ("MONTONETO","CANTIDAD"):
        if c in df.columns: df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)
    # agrupar automáticamente (banco+territorio -> plaza) si viene el crudo
    if "PLAZA" not in df.columns or not df["PLAZA"].notna().any():
        if "UNIDAD" in df.columns and "TERRITORIO" in df.columns:
            df = group_raw(df)
    df["PLAZA"] = df["PLAZA"].map(_norm)
    if "UNIDAD" in df.columns: df["UNIDAD"] = df["UNIDAD"].astype(str).str.strip().str.upper()
    dia = DIA_AVANCE
    if "DIA" in df.columns:
        vals = pd.to_numeric(df["DIA"], errors="coerce").dropna()
        if len(vals): dia = int(vals.mode().iloc[0])
    return df, int(df["PERIODO"].max()), dia


def plaza_series(df, plaza, current, bank=None):
    """series por periodo para el grafico de avances + KPIs de la plaza.
    Si plaza == 'GENERAL' se suman TODAS las plazas del banco."""
    if str(plaza).upper()=="GENERAL":
        d = df[(df.UNIDAD==bank)] if (bank and "UNIDAD" in df.columns) else df
    else:
        d = df[(df.PLAZA == plaza) & ((df.UNIDAD==bank) if (bank and "UNIDAD" in df.columns) else True)]
    periods = sorted(d.PERIODO.unique())
    def m(tipo, per):  # MontoNeto
        return d[(d.TIPO==tipo)&(d.PERIODO==per)]["MONTONETO"].sum()
    def q(tipo, per):  # Cantidad (#OP)
        return d[(d.TIPO==tipo)&(d.PERIODO==per)]["CANTIDAD"].sum()
    avance  = [m("AVANCE", p) for p in periods]
    ops     = [q("AVANCE", p) for p in periods]
    ticket  = [(a/o if o else 0) for a,o in zip(avance, ops)]
    cierre  = []
    for p in periods:
        cierre.append(m("PROYECCION", p) if p == current else m("CIERRE", p))
    cur_idx = periods.index(current) if current in periods else None
    kpi = {"avance": m("AVANCE",current), "meta_avance": m("META_AVANCE",current),
           "meta_mes": m("META_MES",current), "proyeccion": m("PROYECCION",current)}
    kpi["logro"] = (kpi["avance"]/kpi["meta_avance"] if kpi["meta_avance"] else np.nan)
    return {"periods":periods, "labels":[MES_ABBR[int(str(p)[4:6])] for p in periods],
            "avance":avance, "ops":ops, "ticket":ticket, "cierre":cierre,
            "cur_idx":cur_idx, "kpi":kpi}


def plaza_familia(df, plaza, current, bank=None):
    d = df[(df.PLAZA==plaza)&(df.PERIODO==current)&(df.TIPO=="CIERRE")]
    if bank and "UNIDAD" in df.columns: d=d[d.UNIDAD==bank]
    fam = d.groupby("FAMILIA_CONVENIO")["MONTONETO"].sum().sort_values(ascending=False)
    return fam[fam > 0]


def plaza_asesores(df, plaza, current, bank=None):
    d = df[(df.PLAZA==plaza)&(df.PERIODO==current)&(df.TIPO=="CIERRE")]
    if bank and "UNIDAD" in df.columns: d=d[d.UNIDAD==bank]
    ops = d.groupby("DNI_ASESOR_VENTA")["CANTIDAD"].sum()
    return {"1":int((ops==1).sum()), "2-3":int(((ops>=2)&(ops<=3)).sum()), ">=4":int((ops>=4).sum())}


# ---------- graficos ----------
def _fig(size_px, dpi=150):
    W,H = size_px
    fig, ax = plt.subplots(figsize=(W/dpi, H/dpi), dpi=dpi)
    fig.patch.set_alpha(0); ax.set_facecolor("none")
    return fig, ax

def chart_avances(s, plaza, size_px, out, dia):
    fig, ax = _fig(size_px)
    n = len(s["periods"]); xs = np.arange(n); cur = s["cur_idx"]
    if n == 0:                       # sin periodos -> lienzo vacio, nunca reventar
        ax.axis("off")
        ax.text(0.5, 0.5, "Sin información", transform=ax.transAxes,
                ha="center", va="center", color="#8A97A4", fontsize=16, fontweight="bold")
        fig.savefig(out, dpi=150, transparent=True); plt.close(fig)
        return
    cmax = max(s["cierre"]) or 1; amax = max(s["avance"]) or 1
    ax.set_ylim(0, cmax*1.52)
    # barras Monto Avance (banda baja, más corta para no chocar con la línea)
    barh = [a*((cmax*0.24)/amax) for a in s["avance"]]
    ax.bar(xs, barh, width=0.5, color=BAR_AV, zorder=3)
    for x,h,a,o in zip(xs, barh, s["avance"], s["ops"]):
        ax.text(x-0.07, h+cmax*0.02, fmt_bar(a), ha="center", va="bottom",
                color="#111", fontsize=15, fontweight="bold")
        ax.text(x+0.36, h+cmax*0.02, f"{int(o)}", ha="left", va="bottom",
                color=RED, fontsize=15, fontweight="bold")
    # linea Cierre + proyeccion
    if cur is not None and cur > 0:
        ax.plot(xs[:cur+1], s["cierre"][:cur+1], color=LINE_C, lw=4, zorder=4)
        ax.plot(xs[cur-1:cur+1], s["cierre"][cur-1:cur+1], color=LINE_PROJ, lw=4, ls=(0,(4,3)), zorder=5)
    else:
        ax.plot(xs, s["cierre"], color=LINE_C, lw=4, zorder=4)
    for i,(x,c) in enumerate(zip(xs, s["cierre"])):
        ax.text(x, c+cmax*0.06, fmt_mm(c), ha="center", va="bottom",
                color=REDLBL, fontsize=17, fontweight="bold")
    # ticket promedio (fila inferior, con mas separacion de los meses)
    for x,t in zip(xs, s["ticket"]):
        ax.text(x, -cmax*0.22, fmt_ticket(t), ha="center", va="top",
                color=TICKET, fontsize=15, fontweight="bold")
    # ejes
    ax.set_xticks(xs); ax.set_xticklabels(s["labels"], color="#111", fontsize=20, fontweight="bold")
    ax.tick_params(axis="x", length=0, pad=6); ax.set_yticks([])
    for sp in ("top","left","right"): ax.spines[sp].set_visible(False)
    ax.spines["bottom"].set_color("#B0B0B0")
    ax.set_xlim(-0.6, n-0.4)
    ax.set_title(f"Avances del ámbito {plaza}, al {dia} de cada mes",
                 color=NAVY, fontsize=25, fontweight="bold", pad=30)
    ax.text(0.5, -0.34, "Ticket Promedio (S/)", transform=ax.transAxes,
            ha="center", va="top", color=NAVY, fontsize=19, fontweight="bold")
    leg = [Patch(color=BAR_AV, label="Monto Avance"),
           Line2D([0],[0], color=LINE_C, lw=4, label="Cierre"),
           Line2D([0],[0], marker="s", color="w", markerfacecolor=RED, markersize=13, label="#OP"),
           Line2D([0],[0], color=LINE_PROJ, lw=4, ls=(0,(4,3)), label="Proyección")]
    ax.legend(handles=leg, loc="upper left", ncol=2, frameon=False, fontsize=15,
              handlelength=1.5, columnspacing=1.4, bbox_to_anchor=(0.0, 1.03))
    plt.subplots_adjust(left=0.02, right=0.98, top=0.83, bottom=0.28)
    fig.savefig(out, dpi=150, transparent=True); plt.close(fig)

def _donut_vacio(ax, size_px=None):
    """Dona gris de 'sin informacion' cuando no hay datos que graficar."""
    ax.pie([1.0], colors=["#E3E7EB"], startangle=90, counterclock=False,
           wedgeprops=dict(width=0.44, edgecolor="white", linewidth=3))
    ax.text(0, 0, "Sin\ninformación", ha="center", va="center",
            color="#8A97A4", fontsize=11, fontweight="bold")


def chart_familia(fam, mes_abbr, size_px, out):
    fig, ax = _fig(size_px)
    # --- saneamiento: quitar NaN/negativos y quedarnos solo con valores > 0 ---
    names, vals = [], []
    try:
        pares = list(zip(list(fam.index), list(fam.values)))
    except Exception:
        pares = []
    for nm, v in pares:
        try:
            v = float(v)
        except (TypeError, ValueError):
            continue
        if v != v or v <= 0:          # NaN o no positivo
            continue
        names.append(nm); vals.append(v)

    total = sum(vals)
    # matplotlib >= 3.11 lanza ValueError('All wedge sizes are zero') si la
    # lista esta vacia o suma cero -> dibujamos una dona vacia y salimos.
    if not vals or total <= 0:
        _donut_vacio(ax)
        ax.set_title("Participación por Familia",
                     color="#111", fontsize=15, fontweight="bold", loc="left", pad=10)
        ax.axis("equal")
        plt.subplots_adjust(left=0.02, right=0.49, top=0.85, bottom=0.05)
        fig.savefig(out, dpi=150, transparent=True); plt.close(fig)
        return

    colors = [FAMILY_COLORS.get(nm, FALLBACK_COLORS[i % len(FALLBACK_COLORS)]) for i,nm in enumerate(names)]
    wedges = ax.pie(vals, colors=colors, startangle=90, counterclock=False,
                    wedgeprops=dict(width=0.44, edgecolor="white", linewidth=3))[0]
    ax.set_title("Participación por Familia",
                 color="#111", fontsize=15, fontweight="bold", loc="left", pad=10)
    labels = []
    for nm,v in zip(names,vals):
        pct = v/total*100
        short = nm if len(nm)<=16 else nm[:15]+"."
        labels.append(f"{short} ({pct:.1f}%)")
    ax.legend(wedges, labels, loc="center left", bbox_to_anchor=(1.0, 0.5),
              frameon=False, fontsize=14, labelspacing=0.65, handlelength=1.2, handletextpad=0.45)
    plt.subplots_adjust(left=0.02, right=0.49, top=0.85, bottom=0.05)
    fig.savefig(out, dpi=150, transparent=True); plt.close(fig)

def chart_asesores(buckets, size_px, out):
    fig, ax = _fig(size_px)
    cats = ["1","2-3",">=4"]; vals = [buckets[c] for c in cats]
    cols = [AS_BLUE, AS_GREY, AS_GREEN]
    ax.bar(cats, vals, width=0.6, color=cols, zorder=3)
    ymax = max(vals) or 1
    ax.set_ylim(0, ymax*1.24)
    for i,v in enumerate(vals):
        ax.text(i, v+ymax*0.02, str(v), ha="center", va="bottom",
                color=DARK, fontsize=24, fontweight="bold")
    ax.set_title("Asesores por Rango de Ventas",
                 color="#25384a", fontsize=15, fontweight="bold", pad=12)
    ax.set_yticks([])
    for sp in ("top","left","right"): ax.spines[sp].set_visible(False)
    ax.spines["bottom"].set_color("#888")
    ax.tick_params(axis="x", length=0, labelsize=15)
    for lbl in ax.get_xticklabels(): lbl.set_color("#333")
    plt.subplots_adjust(left=0.05, right=0.96, top=0.84, bottom=0.12)
    fig.savefig(out, dpi=150, transparent=True); plt.close(fig)


# ---------- helpers XML (comunes) ----------
def _off(sp):
    o = sp.getElementsByTagName("a:off")
    if not o: return (None,None)
    try: return (int(o[0].getAttribute("x"))/EMU, int(o[0].getAttribute("y"))/EMU)
    except: return (None,None)
def _txt(sp): return " ".join(t.firstChild.nodeValue for t in sp.getElementsByTagName("a:t") if t.firstChild)
def _set(t,v):
    if t.firstChild: t.firstChild.nodeValue = v
def _fix_dates(doc, mes, year, dia):
    for t in doc.getElementsByTagName("a:t"):
        if not t.firstChild: continue
        s = t.firstChild.nodeValue
        if DATE_RE.search(s):     _set(t, DATE_RE.sub(f"al {dia} de {mes}", s))
        elif CIERRE_RE.search(s): _set(t, CIERRE_RE.sub(f"Cierre {mes} {year}", s))
def which(text_low, kw):
    hit,best = None,-1
    for k,v in kw.items():
        if v in text_low and len(v) > best: hit,best = k,len(v)
    return hit
def _title_plaza(doc):
    """Identifica la plaza SOLO por el titulo (evita choques con el pie de pagina)."""
    for sp in doc.getElementsByTagName("p:sp"):
        tt = _txt(sp).lower()
        if "convenios" in tt and "informaci" in tt:
            return which(tt, TITLE_KW)

def _title_bank_plaza(doc, df):
    """Devuelve (BANCO, PLAZA) del título de una lámina de convenios, según los datos."""
    for sp in doc.getElementsByTagName("p:sp"):
        tt = _txt(sp).lower()
        if "convenios" in tt and "informaci" in tt:
            bank=None
            for b in ("bbva","bcp","sbp"):
                if b in tt: bank=b.upper(); break
            if bank is None: return (None,None)
            if "general" in tt: return (bank, "GENERAL")
            if "UNIDAD" in df.columns:
                plazas=[str(p) for p in df[df.UNIDAD==bank].PLAZA.dropna().unique()]
            else:
                plazas=[str(p) for p in df.PLAZA.dropna().unique()]
            best=None
            for p in plazas:
                if p and p.lower() in tt and (best is None or len(p)>len(best)): best=p
            return (bank, best)
    return (None,None)
    return which(" ".join(_txt(sp).lower() for sp in doc.getElementsByTagName("p:sp")), TITLE_KW)
def _frame_aspect(slide_xml, rels_xml, image):
    rid2img = dict(re.findall(r'Id="([^"]+)"[^>]*Target="[^"]*media/([^"]+)"', rels_xml))
    doc = parseString(slide_xml)
    for pic in doc.getElementsByTagName("p:pic"):
        b = pic.getElementsByTagName("a:blip")
        if not b or rid2img.get(b[0].getAttribute("r:embed")) != image: continue
        for x in pic.getElementsByTagName("a:xfrm"):
            e = x.getElementsByTagName("a:ext")
            if e:
                cx,cy = int(e[0].getAttribute("cx")), int(e[0].getAttribute("cy"))
                if cy: return cx/cy
    return None

def edit_detail_text(doc, kpi):
    lg = kpi.get("logro", np.nan)
    for t in doc.getElementsByTagName("a:t"):
        if not t.firstChild: continue
        s = t.firstChild.nodeValue; low = s.strip().lower()
        if low.startswith("meta avance"):        _set(t, f"META AVANCE: {kfmt(kpi['meta_avance'])}")
        elif low.startswith("meta:") or low.startswith("meta "):
            if kpi["meta_mes"]:                   _set(t, f"META: {kfmt(kpi['meta_mes'])}")
        elif low.startswith("avance:"):           _set(t, f"AVANCE: {kfmt(kpi['avance'])}")
        elif PCT_RE.match(s) and pd.notna(lg):    _set(t, fmt_pct(lg)); summary_panel.color_run(t, summary_panel.pct_color(lg*100))
        elif ((re.match(r"^\s*EN\s+\w", s, re.I) and len(s.strip())<22) or s.strip().lower() in ("bajo la meta","cerca de la meta","meta alcanzada","bajo meta","cerca de meta","meta alcanzada")) and pd.notna(lg):
            _set(t, summary_panel.estado_text(lg*100)); summary_panel.color_run(t, summary_panel.pct_color(lg*100))
    if pd.notna(lg): summary_panel.recolor_pill(doc, lg*100)

def edit_summary(doc, kpis):
    shapes = [(sp,*_off(sp),_txt(sp)) for sp in doc.getElementsByTagName("p:sp")]
    colx = {}
    for sp,x,y,txt in shapes:
        if x is None: continue
        h = txt.strip().lower()
        if h == "% logro": colx["logro"] = x
        elif h == "avance": colx["avance"] = x
        elif h == "proyeccion": colx["proy"] = x
    rows = {}
    for sp,x,y,txt in shapes:
        if x is None or y is None or x>=3.5 or y<=3.5: continue
        p = which(txt.lower(), ROW_KW)
        if p: rows[p] = y
    def cell(col,p):
        if col not in colx or p not in rows: return None
        cx,cy = colx[col], rows[p]; best,bd = None,1e9
        for sp,x,y,txt in shapes:
            if x is None or y is None: continue
            if abs(x-cx)<0.85 and abs(y-cy)<0.18:
                for t in sp.getElementsByTagName("a:t"):
                    if t.firstChild and t.firstChild.nodeValue.strip() and abs(y-cy)<bd:
                        best,bd = t,abs(y-cy)
        return best
    logros=[]
    for p in rows:
        k = kpis.get(p); lg = k.get("logro",np.nan) if k else np.nan
        if pd.notna(lg): logros.append(lg)
        c=cell("logro",p);  c and pd.notna(lg) and _set(c, fmt_pct(lg))
        c=cell("avance",p); c and _set(c, kfmt(k["avance"]))
        c=cell("proy",p);   c and _set(c, kfmt(k["proyeccion"]))
    _gr=[{"avance":kpis[p]["avance"],
          "logro":(kpis[p]["logro"] if pd.notna(kpis[p].get("logro",np.nan)) else None),
          "meta_avance":kpis[p].get("meta_avance")} for p in rows if kpis.get(p)]
    avg = summary_panel.agg_frac(_gr)   # % agregado (avance/meta) = mismo criterio que el panel
    for sp,x,y,txt in shapes:
        if x is None or x>=2.0: continue
        for t in sp.getElementsByTagName("a:t"):
            if t.firstChild and PCT_RE.match(t.firstChild.nodeValue) and pd.notna(avg):
                _set(t, (f"{avg*100:.0f}%" if avg>=1 else f"{avg*100:.1f}%"))


# ---------- orquestador ----------
def update_presentation(template_bytes, excel_bytes):
    df, current, dia = load_data(excel_bytes)
    year, month = int(str(current)[:4]), int(str(current)[4:6])
    mes, mesab = MES_FULL[month], MES_ABBR[month]

    zin = zipfile.ZipFile(io.BytesIO(template_bytes))
    files = {n: zin.read(n) for n in zin.namelist()}; zin.close()

    refs = Counter(); slide_imgs = {}
    for n in list(files):
        mm = re.match(r"ppt/slides/_rels/(slide\d+)\.xml\.rels", n)
        if mm:
            imgs = re.findall(r"media/(image[0-9]+\.\w+)", files[n].decode("utf-8"))
            slide_imgs[mm.group(1)] = imgs
            for im in imgs: refs[im]+=1
    logo = refs.most_common(1)[0][0] if refs else None

    # precomputo KPIs de todas las plazas (para el resumen)
    kpis = {p[0]: plaza_series(df, p[0], current)["kpi"] for p in PLAZAS if p[0] in df.PLAZA.unique()}
    warnings = [f"'{p[0]}' no aparece en el Excel; se omite." for p in PLAZAS if p[0] not in df.PLAZA.unique()]

    slide_files = sorted([n for n in files if re.match(r"ppt/slides/slide\d+\.xml$", n)],
                         key=lambda x:int(re.findall(r"\d+", x.split("/")[-1])[0]))
    for sf in slide_files:
        sid = sf.split("/")[-1][:-4]
        xml = files[sf].decode("utf-8"); low = xml.lower()
        doc = parseString(xml)
        if "resumen general" in low:
            _fix_dates(doc, mes, year, dia)
            files[sf] = doc.toxml().encode("utf-8")
            disp = {"LIMA":"BBVA Conv. Lima","TELECAMPO":"BBVA Conv. Telecampo",
                    "NORTE":"BBVA Conv. Norte","SUR":"BBVA Conv. Sur","ORIENTE":"BBVA Conv. Oriente"}
            rows=[]; logros=[]; proy_total=0.0
            for key,_,_ in PLAZAS:
                if key not in kpis: continue
                k=kpis[key]; lg=k["logro"]
                rows.append({"name":disp.get(key,key.title()), "avance":k["avance"],
                             "logro":(float(lg) if pd.notna(lg) else None), "proy":k["proyeccion"],
                             "meta_avance":k["meta_avance"], "meta_mes":k["meta_mes"]})
                if pd.notna(lg): logros.append(lg)
                proy_total += k["proyeccion"]
            avg = summary_panel.agg_frac(rows)   # % agregado (avance/meta), coincide con GENERAL
            num = int(re.findall(r"\d+", sid)[0])
            summary_panel.redesign_summary(files, num, avg, rows, proy_total)
        elif "meta avance" in low:
            plaza = _title_plaza(doc)
            if plaza and plaza in df.PLAZA.unique():
                s = plaza_series(df, plaza, current)
                edit_detail_text(doc, s["kpi"]); _fix_dates(doc, mes, year, dia)
                files[sf] = doc.toxml().encode("utf-8")
                rels = files.get(f"ppt/slides/_rels/{sid}.xml.rels", b"").decode("utf-8")
                fam = plaza_familia(df, plaza, current)
                aso = plaza_asesores(df, plaza, current)
                for im in [i for i in slide_imgs.get(sid, []) if i != logo]:
                    w,h = Image.open(io.BytesIO(files[f"ppt/media/{im}"])).size
                    asp = _frame_aspect(xml, rels, im) or (w/h)
                    if w/h >= 1.85:                         # Avances (compuesto)
                        W=2400; size=(W, round(W/asp)); buf=io.BytesIO()
                        chart_avances(s, plaza, size, buf, dia); files[f"ppt/media/{im}"]=buf.getvalue()
                    elif h >= 600 or w >= 950:              # Dona participacion
                        W=1700; size=(W, round(W/asp)); buf=io.BytesIO()
                        chart_familia(fam, mesab, size, buf); files[f"ppt/media/{im}"]=buf.getvalue()
                    else:                                    # Asesores por rango
                        W=1500; size=(W, round(W/asp)); buf=io.BytesIO()
                        chart_asesores(aso, size, buf); files[f"ppt/media/{im}"]=buf.getvalue()
            else:
                _fix_dates(doc, mes, year, dia); files[sf]=doc.toxml().encode("utf-8")
        else:
            _fix_dates(doc, mes, year, dia); files[sf]=doc.toxml().encode("utf-8")

    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        for n,d in files.items(): z.writestr(n,d)
    out.seek(0)
    resumen = {"periodo":current, "mes":f"{mes} {year}",
               "plazas":{p:{"avance":kfmt(k["avance"]),
                            "meta":kfmt(k["meta_mes"]) if k["meta_mes"] else "-",
                            "logro":fmt_pct(k["logro"]) if pd.notna(k["logro"]) else "-",
                            "proyeccion":kfmt(k["proyeccion"]) if k["proyeccion"] else "-"}
                         for p,k in kpis.items()}}
    return out.getvalue(), resumen, warnings
# ===== Agrupación automática Convenios (banco + territorio -> plaza) =====
TERRITORIO_A_PLAZA = {
 "BBVA": {"APURIMAC":"Sur","AREQUIPA":"Sur","CAJAMARCA":"Norte","CALL":"TELECAMPO","CHICLAYO":"Norte",
          "HUANUCO":"Centro","HUARAZ":"Norte","ICA":"Sur","IQUITOS":"Oriente","LIMA":"LIMA",
          "PUCALLPA":"Oriente","PUNO":"Sur","TACNA":"Sur","TARAPOTO":"Oriente","TRUJILLO":"Norte"},
 "BCP":  {"AREQUIPA":"Sur/Oriente","CAJAMARCA":"Norte I","CHICLAYO":"Norte II","CHIMBOTE":"Norte I",
          "IQUITOS":"Sur/Oriente","LIMA":"LIMA","PIURA":"Norte II","PUNO":"Sur/Oriente",
          "TELECAMPO":"TELECAMPO","TRUJILLO":"Norte I"},
 "SBP":  {"CAJAMARCA":"Norte","CHICLAYO":"Norte","HUANCAYO":"COA","ICA":"SUR","IQUITOS":"COA",
          "LIMACAMPO":"LIMA","PUCALLPA":"COA","TELECAMPO":"TELECAMPO","TRUJILLO":"Norte"},
}
def group_raw(df):
    """Agrega la columna PLAZA a partir de (UNIDAD/banco, TERRITORIO) usando el mapeo. No pisa PLAZA existente."""
    import pandas as _pd
    if "PLAZA" in df.columns and df["PLAZA"].notna().any():
        return df
    def _p(r):
        b=str(r.get("UNIDAD","")).strip().upper(); t=str(r.get("TERRITORIO","")).strip().upper()
        return TERRITORIO_A_PLAZA.get(b,{}).get(t)
    df=df.copy(); df["PLAZA"]=df.apply(_p,axis=1)
    return df
# ========================================================================

