# -*- coding: utf-8 -*-
"""
Motor de actualizacion de reportes BBVA (TLM OUT + Digital u otras campanas).
Entrada: bytes de la plantilla .pptx + bytes del Excel semanal.
Salida:  bytes del .pptx actualizado (mismo diseno; solo cambian datos y graficos).

DISENO CONFIG-DRIVEN
--------------------
Para agregar/quitar/renombrar campanas solo se edita la lista PRODUCTS (abajo).
Todo lo demas se autodetecta:
  - la hoja del Excel con datos limpios (PRODUCTO/PERIODO/TIPO/MONTO),
  - el periodo actual (PERIODO mas alto),
  - las laminas resumen ("RESUMEN GENERAL"),
  - las laminas de detalle (las que tienen el bloque META/AVANCE),
  - el grafico de cada lamina (la imagen que NO es el logo compartido).
"""
import io, re, zipfile
from collections import Counter
import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import Patch
from PIL import Image
from defusedxml.minidom import parseString
import summary_panel

# ===================== CONFIGURACION (unico lugar a tocar) =====================
# (clave_en_excel, texto_titulo_lamina_detalle, texto_fila_resumen, titulo_grafico)
PRODUCTS = [
    ("TC OUT",     "tc telemarketing out", "tc telemarketing",    "FORMALIZADAS"),
    ("PLD OUT",    "pld out prestamos",    "pld out",             "FORMALIZADAS"),
    ("OPERA OUT",  "operaciones out",      "operaciones out",     "FORMALIZADAS"),
    ("PORTAFOLIO", "portafolio",           "portafolio",          "DESEMBOLSADO"),
    ("TC IN",      "tc hibrido",           "tc hibrido",          "FORMALIZADAS"),
    ("OPERA IN",   "operaciones digital",  "operaciones digital", "FORMALIZADAS"),
    ("PLD WEB",    "pld digital",          "pld digital",         "FORMALIZADAS"),
    ("TC START",   "tc respaldada",        "tc respaldada",       "FORMALIZADAS"),
]
DIA_AVANCE = 11  # "avance al dia 11"
# ==============================================================================

# estilo (muestreado de la plantilla)
TEAL, CYAN   = "#009090", "#00CCCC"
NAVY, NAVY_L = "#245490", "#7FA3CC"
SLATE, GRAY  = "#3C5460", "#787878"
MES_ABBR = {1:"ENE",2:"FEB",3:"MAR",4:"ABR",5:"MAY",6:"JUN",
            7:"JUL",8:"AGO",9:"SEP",10:"OCT",11:"NOV",12:"DIC"}
MES_FULL = {1:"Enero",2:"Febrero",3:"Marzo",4:"Abril",5:"Mayo",6:"Junio",
            7:"Julio",8:"Agosto",9:"Septiembre",10:"Octubre",11:"Noviembre",12:"Diciembre"}
EMU = 914400
PCT_RE = re.compile(r"^\s*\d+([.,]\d+)?\s*%\s*$")
DATE_RE = re.compile(r"al\s+\d{1,2}\s+de\s+[A-Za-zÁÉÍÓÚáéíóú]+", re.IGNORECASE)
CIERRE_RE = re.compile(r"cierre\s+\w+\s+\d{4}", re.IGNORECASE)

TITLE_KW = {p[0]: p[1] for p in PRODUCTS}
ROW_KW   = {p[0]: p[2] for p in PRODUCTS}
CHART_TT = {p[0]: p[3] for p in PRODUCTS}


def kfmt(v):
    v = float(v)
    if v >= 1_000_000: return f"{v/1e6:.2f}M"
    if v >= 100_000:   return f"{v/1e3:.1f}K"
    if v >= 1000:      return f"{v/1e3:.2f}K"
    return f"{v:.0f}"

def fmt_pct(frac):
    v=float(frac)*100
    return f"{v:.0f}%" if v>=100 else f"{v:.1f}%"

def _norm(s):
    return (str(s).strip().upper().replace("Á","A").replace("É","E").replace("Í","I")
            .replace("Ó","O").replace("Ú","U"))

# ---------------- Excel ----------------
def load_metrics(excel_bytes):
    xls = pd.ExcelFile(io.BytesIO(excel_bytes))
    need = {"PRODUCTO", "PERIODO", "TIPO", "MONTO"}
    df = None
    for sh in xls.sheet_names:                       # elegir la hoja con datos limpios
        tmp = xls.parse(sh)
        cols = {c: _norm(c).replace(" ", "") for c in tmp.columns}
        tmp = tmp.rename(columns=cols)
        if need.issubset(set(tmp.columns)):
            df = tmp; break
    if df is None:
        raise ValueError("No encontre una hoja con columnas PRODUCTO, PERIODO, TIPO, MONTO.")

    df["PRODUCTO"] = df["PRODUCTO"].map(_norm)
    df["TIPO"]     = df["TIPO"].map(lambda x: _norm(x).replace(" ", "_"))
    df = df[df["PERIODO"].notna()]
    df["PERIODO"]  = df["PERIODO"].astype(int)
    df["MONTO"]    = pd.to_numeric(df["MONTO"], errors="coerce")
    piv = df.pivot_table(index=["PRODUCTO","PERIODO"], columns="TIPO",
                         values="MONTO", aggfunc="first")
    current = int(df["PERIODO"].max())

    warnings, metrics = [], {}
    for key, *_ in PRODUCTS:
        if key not in piv.index.get_level_values(0):
            warnings.append(f"'{key}' no esta en el Excel; se omite.")
            continue
        row = piv.loc[(key, current)]
        av, ma = row.get("AVANCE", np.nan), row.get("META_AVANCE", np.nan)
        metrics[key] = {
            "avance": av, "meta_avance": ma,
            "meta_mes": row.get("META_MES", np.nan),
            "proyeccion": row.get("PROYECCION", np.nan),
            "logro": (av/ma if (pd.notna(av) and pd.notna(ma) and ma) else np.nan),
        }
        if pd.isna(ma) or ma == 0:
            warnings.append(f"{key}: META_AVANCE ausente/0 en {current}; %Logro no calculable.")
    return piv, current, metrics, warnings

# ---------------- grafico ----------------
def build_chart(piv, prod, title, current, size_px, out):
    sub = piv.loc[prod].sort_index()
    periods = list(sub.index)
    labels  = [MES_ABBR[int(str(p)[4:6])] for p in periods]
    avance  = [float(sub.loc[p, "AVANCE"]) for p in periods]
    cierre  = []
    for p in periods:
        proy = sub.loc[p].get("PROYECCION", np.nan)
        cierre.append(float(proy) if (p == current and pd.notna(proy)) else float(sub.loc[p, "CIERRE"]))
    n = len(periods); xs = np.arange(n)
    cur = periods.index(current) if current in periods else None
    W, H = size_px; dpi = 100
    fig, ax = plt.subplots(figsize=(W/dpi, H/dpi), dpi=dpi)
    fig.patch.set_alpha(0); ax.set_facecolor("none")
    cmax = max(cierre) or 1; amax = max(avance) or 1
    ax.set_ylim(0, cmax*1.34)
    barh = [a*((cmax*0.42)/amax) for a in avance]
    bcol = [CYAN if (cur is not None and i == cur) else TEAL for i in range(n)]
    ax.bar(xs, barh, width=0.42, color=bcol, zorder=3)
    for x, h, a in zip(xs, barh, avance):
        ax.text(x, h+cmax*0.02, kfmt(a), ha="center", va="bottom", color=SLATE, fontsize=40, fontweight="bold")
    if cur is not None and cur > 0:
        ax.plot(xs[:cur+1], cierre[:cur+1], color=NAVY, lw=6, solid_capstyle="round", zorder=4)
        ax.plot(xs[cur-1:cur+1], cierre[cur-1:cur+1], color=NAVY_L, lw=6, ls=(0,(4,3)), zorder=5)
        ax.plot(xs[:cur], cierre[:cur], "o", color=NAVY, ms=11, zorder=6)
        ax.plot([xs[cur]], [cierre[cur]], "o", color=NAVY_L, ms=11, zorder=6)
    else:
        ax.plot(xs, cierre, color=NAVY, lw=6, marker="o", ms=11, zorder=4)
    for i, (x, c) in enumerate(zip(xs, cierre)):
        col = NAVY_L if (cur is not None and i == cur) else NAVY
        ax.text(x, c+cmax*0.035, kfmt(c), ha="center", va="bottom", color=col, fontsize=42, fontweight="bold")
    ax.set_xticks(xs); ax.set_xticklabels(labels, color=GRAY, fontsize=42, fontweight="bold")
    ax.tick_params(axis="x", length=0, pad=14); ax.set_yticks([])
    for sp in ("top","left","right"): ax.spines[sp].set_visible(False)
    ax.spines["bottom"].set_color("#B0B0B0"); ax.spines["bottom"].set_linewidth(2)
    ax.set_xlim(-0.6, n-0.4)
    ax.text(-0.55, cmax*1.30, title, color=SLATE, fontsize=78, fontweight="bold", ha="left", va="top")
    ax.legend(handles=[Patch(color=SLATE, label="AVANCE"), Patch(color=NAVY, label="CIERRE")],
              loc="upper right", ncol=2, frameon=False, fontsize=46, handlelength=1.1,
              columnspacing=1.6, bbox_to_anchor=(1.0,0.99), labelcolor=[SLATE, NAVY])
    plt.subplots_adjust(left=0.01, right=0.99, top=0.99, bottom=0.09)
    fig.savefig(out, dpi=dpi, transparent=True); plt.close(fig)

# ---------------- helpers XML ----------------
def _txt(sp):
    return " ".join(t.firstChild.nodeValue for t in sp.getElementsByTagName("a:t") if t.firstChild)
def _off(sp):
    o = sp.getElementsByTagName("a:off")
    if not o: return (None, None)
    try:    return (int(o[0].getAttribute("x"))/EMU, int(o[0].getAttribute("y"))/EMU)
    except: return (None, None)
def _set(t, v):
    if t.firstChild: t.firstChild.nodeValue = v
def _fix_dates(doc, mes, year):
    for t in doc.getElementsByTagName("a:t"):
        if not t.firstChild: continue
        s = t.firstChild.nodeValue
        if DATE_RE.search(s):   _set(t, DATE_RE.sub(f"al {DIA_AVANCE} de {mes}", s))
        elif CIERRE_RE.search(s): _set(t, CIERRE_RE.sub(f"Cierre {mes} {year}", s))

def _strip_srcrect(doc, rels_xml, chart_image):
    """Quita el recorte (a:srcRect) del pic del grafico, para mostrarlo completo."""
    rid2img = dict(re.findall(r'Id="([^"]+)"[^>]*Target="[^"]*media/([^"]+)"', rels_xml))
    for pic in doc.getElementsByTagName("p:pic"):
        blip = pic.getElementsByTagName("a:blip")
        if not blip or rid2img.get(blip[0].getAttribute("r:embed")) != chart_image:
            continue
        for sr in pic.getElementsByTagName("a:srcRect"):
            sr.parentNode.removeChild(sr)

def _frame_aspect(slide_xml, rels_xml, chart_image):
    """Aspecto (cx/cy) del marco donde va el grafico, para no deformar la imagen."""
    rid2img = dict(re.findall(r'Id="([^"]+)"[^>]*Target="[^"]*media/([^"]+)"', rels_xml))
    doc = parseString(slide_xml)
    for pic in doc.getElementsByTagName("p:pic"):
        blip = pic.getElementsByTagName("a:blip")
        if not blip: continue
        rid = blip[0].getAttribute("r:embed")
        if rid2img.get(rid) != chart_image: continue
        for x in pic.getElementsByTagName("a:xfrm"):
            ext = x.getElementsByTagName("a:ext")
            if ext:
                cx, cy = int(ext[0].getAttribute("cx")), int(ext[0].getAttribute("cy"))
                if cy: return cx / cy
    return None

def which_product(text_low, kw_map):
    hit = None; best = -1
    for prod, kw in kw_map.items():
        if kw in text_low and len(kw) > best:
            hit, best = prod, len(kw)
    return hit

# ---------------- edicion de laminas ----------------
def edit_detail(doc, prod, m):
    logro = m.get("logro", np.nan)
    for sp in doc.getElementsByTagName("p:sp"):
        for t in sp.getElementsByTagName("a:t"):
            if not t.firstChild: continue
            s = t.firstChild.nodeValue; low = s.strip().lower()
            if low.startswith("meta avance"):        _set(t, f"META AVANCE: {kfmt(m['meta_avance'])}")
            elif low.startswith("meta:") or low.startswith("meta "):
                if pd.notna(m["meta_mes"]):           _set(t, f"META: {kfmt(m['meta_mes'])}")
            elif low.startswith("avance:"):           _set(t, f"AVANCE: {kfmt(m['avance'])}")
            elif PCT_RE.match(s) and pd.notna(logro):  _set(t, fmt_pct(logro)); summary_panel.color_run(t, summary_panel.pct_color(logro*100))
            elif ((re.match(r"^\s*EN\s+\w", s, re.I) and len(s.strip())<22) or s.strip().lower() in ("bajo la meta","cerca de la meta","meta alcanzada")) and pd.notna(logro):
                _set(t, summary_panel.estado_text(logro*100)); summary_panel.color_run(t, summary_panel.pct_color(logro*100))
    if pd.notna(logro): summary_panel.recolor_pill(doc, logro*100)

def edit_summary(doc, metrics):
    shapes = [(sp, *_off(sp), _txt(sp)) for sp in doc.getElementsByTagName("p:sp")]
    # x de cada columna a partir de los encabezados
    colx = {}
    for sp, x, y, txt in shapes:
        h = txt.strip().lower()
        if x is None: continue
        if h == "% logro":   colx["logro"] = x
        elif h == "avance":  colx["avance"] = x
        elif h == "proyeccion": colx["proy"] = x
    # filas: y de cada producto presente en esta lamina
    rows = {}
    for sp, x, y, txt in shapes:
        if x is None or y is None or x >= 3.5 or y <= 3.5: continue
        prod = which_product(txt.lower(), ROW_KW)
        if prod: rows[prod] = y
    def cell(colkey, prod):
        if colkey not in colx or prod not in rows: return None
        cx, cy = colx[colkey], rows[prod]; best, bd = None, 1e9
        for sp, x, y, txt in shapes:
            if x is None or y is None: continue
            if abs(x-cx) < 0.85 and abs(y-cy) < 0.18:
                for t in sp.getElementsByTagName("a:t"):
                    if t.firstChild and t.firstChild.nodeValue.strip() and abs(y-cy) < bd:
                        best, bd = t, abs(y-cy)
        return best
    logros = []
    for prod in rows:
        m = metrics.get(prod);  lg = m.get("logro", np.nan) if m else np.nan
        if pd.notna(lg): logros.append(lg)
        c = cell("logro", prod);  c and pd.notna(lg) and _set(c, fmt_pct(lg))
        c = cell("avance", prod); c and _set(c, kfmt(m["avance"]))
        c = cell("proy", prod);   c and pd.notna(m["proyeccion"]) and _set(c, kfmt(m["proyeccion"]))
    avg = float(np.mean(logros)) if logros else np.nan
    for sp, x, y, txt in shapes:            # avance promedio (% grande a la izquierda)
        if x is None or x >= 2.0: continue
        for t in sp.getElementsByTagName("a:t"):
            if t.firstChild and PCT_RE.match(t.firstChild.nodeValue) and pd.notna(avg):
                _set(t, f"{avg*100:.0f}%")

# ---------------- orquestador ----------------
def update_presentation(template_bytes, excel_bytes):
    piv, current, metrics, warnings = load_metrics(excel_bytes)
    year, month = int(str(current)[:4]), int(str(current)[4:6]); mes = MES_FULL[month]

    zin = zipfile.ZipFile(io.BytesIO(template_bytes))
    files = {n: zin.read(n) for n in zin.namelist()}; zin.close()

    # logo = imagen mas referenciada
    refs = Counter()
    slide_imgs = {}
    for n in files:
        mm = re.match(r"ppt/slides/_rels/(slide\d+)\.xml\.rels", n)
        if mm:
            imgs = re.findall(r"media/(image[0-9]+\.\w+)", files[n].decode("utf-8"))
            slide_imgs[mm.group(1)] = imgs
            for im in imgs: refs[im] += 1
    logo = refs.most_common(1)[0][0] if refs else None

    slide_files = sorted([n for n in files if re.match(r"ppt/slides/slide\d+\.xml$", n)],
                         key=lambda x: int(re.findall(r"\d+", x.split("/")[-1])[0]))

    for sf in slide_files:
        sid = sf.split("/")[-1][:-4]           # 'slideN'
        xml = files[sf].decode("utf-8")
        low_all = xml.lower()
        doc = parseString(xml)

        if "resumen general" in low_all:                       # ---- resumen ----
            _fix_dates(doc, mes, year); files[sf] = doc.toxml().encode("utf-8")
            # productos presentes en esta lamina (en orden), con su nombre visible
            found = []
            for sp in doc.getElementsByTagName("p:sp"):
                x, y = _off(sp); tt = _txt(sp)
                if x is None or y is None or x >= 3.5 or y <= 3.5: continue
                pr = which_product(tt.lower(), ROW_KW)
                if pr and pr not in [f[0] for f in found]:
                    found.append((pr, y, tt.strip()))
            found.sort(key=lambda f: f[1])
            rows = []; logros = []; proy_total = 0.0
            for pr, _, name in found:
                m = metrics.get(pr)
                if not m: continue
                lg = m.get("logro", np.nan)
                rows.append({"name": name, "key": pr, "avance": m["avance"],
                             "logro": (float(lg) if pd.notna(lg) else None),
                             "proy": (m["proyeccion"] if pd.notna(m["proyeccion"]) else 0)})
                if pd.notna(lg): logros.append(lg)
                if pd.notna(m["proyeccion"]): proy_total += m["proyeccion"]
            avg = float(np.mean(logros)) if logros else float("nan")
            num = int(re.findall(r"\d+", sid)[0])
            if rows:
                # 4a card TLM: campaña con menor % Logro (no se puede sumar proyecciones de campañas distintas)
                cand=[r for r in rows if r["logro"] is not None]
                card4=None
                if cand:
                    w=min(cand, key=lambda r: r["logro"]); p=w["logro"]*100
                    acc="#C0392B" if p<90 else "#E08A1E"
                    card4=("A REFORZAR", (f"{p:.0f}%" if p>=100 else f"{p:.1f}%"),
                           f"{w.get('key', w['name'])} · menor logro", acc, acc)
                summary_panel.redesign_summary(files, num, avg, rows, proy_total, card4)
        elif "meta avance" in low_all:                          # ---- detalle ----
            prod = which_product(low_all, TITLE_KW)
            if prod and prod in metrics:
                edit_detail(doc, prod, metrics[prod]); _fix_dates(doc, mes, year)
                # regenerar el grafico (imagen no-logo de esta lamina)
                charts = [im for im in slide_imgs.get(sid, []) if im != logo]
                rels = files.get(f"ppt/slides/_rels/{sid}.xml.rels", b"").decode("utf-8")
                if charts:
                    im = max(charts, key=lambda k: Image.open(io.BytesIO(files[f"ppt/media/{k}"])).size[0]
                                                    * Image.open(io.BytesIO(files[f"ppt/media/{k}"])).size[1])
                    px = Image.open(io.BytesIO(files[f"ppt/media/{im}"])).size
                    _strip_srcrect(doc, rels, im)   # mostrar el grafico completo (sin recorte)
                    files[sf] = doc.toxml().encode("utf-8")
                    aspect = _frame_aspect(files[sf].decode("utf-8"), rels, im) or (px[0]/px[1])
                    W = 4750; H = max(600, round(W / aspect))
                    buf = io.BytesIO(); build_chart(piv, prod, CHART_TT[prod], current, (W, H), buf)
                    files[f"ppt/media/{im}"] = buf.getvalue()
                    if px[0] / px[1] < 1.6:   # la imagen original era compuesta
                        warnings.append(f"{prod}: la imagen traia un panel extra 'KPIs' (CET/EFECTIVIDAD) "
                                        f"que no viene en el Excel; se regenero solo 'FORMALIZADAS' "
                                        f"con los datos disponibles (hasta el periodo actual).")
                else:
                    files[sf] = doc.toxml().encode("utf-8")
            else:
                _fix_dates(doc, mes, year); files[sf] = doc.toxml().encode("utf-8")
        else:                                                   # ---- portada/diagnostico ----
            _fix_dates(doc, mes, year); files[sf] = doc.toxml().encode("utf-8")

    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        for n, d in files.items(): z.writestr(n, d)
    out.seek(0)

    resumen = {"periodo": current, "mes": f"{mes} {year}",
               "productos": {p: {"avance": kfmt(m["avance"]),
                                 "meta": kfmt(m["meta_mes"]) if pd.notna(m["meta_mes"]) else "-",
                                 "logro": (f"{m['logro']*100:.1f}%" if pd.notna(m["logro"]) else "-"),
                                 "proyeccion": kfmt(m["proyeccion"]) if pd.notna(m["proyeccion"]) else "-"}
                             for p, m in metrics.items()}}
    return out.getvalue(), resumen, warnings
