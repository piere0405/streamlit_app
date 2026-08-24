# -*- coding: utf-8 -*-
"""Divide un PPTX consolidado en subconjuntos de láminas (convenios/call, o por campaña).
No re-genera nada: opera sobre el .pptx ya producido, conservando relaciones y medios."""
import io, re, zipfile

def _read(b):
    z=zipfile.ZipFile(io.BytesIO(b)); return {n:z.read(n) for n in z.namelist()}
def _write(files):
    o=io.BytesIO()
    with zipfile.ZipFile(o,"w",zipfile.ZIP_DEFLATED) as z:
        for n,d in files.items(): z.writestr(n,d)
    return o.getvalue()

def _slide_order(files):
    pres=files["ppt/presentation.xml"].decode("utf-8")
    rels=files["ppt/_rels/presentation.xml.rels"].decode("utf-8")
    rid2s={rid:t.split("/")[-1] for rid,t in re.findall(r'Id="([^"]+)"[^>]*Target="(slides/slide\d+\.xml)"',rels)}
    order=re.findall(r'<p:sldId[^>]*r:id="([^"]+)"',pres)
    return [rid2s[r] for r in order if r in rid2s]

def _text(files,name):
    xml=files[f"ppt/slides/{name}"].decode("utf-8")
    return " ".join(re.findall(r'<a:t>([^<]*)</a:t>',xml))

def _first_text(files,name):
    xml=files[f"ppt/slides/{name}"].decode("utf-8")
    for t in re.findall(r'<a:t>([^<]*)</a:t>',xml):
        if t.strip(): return t.strip()
    return name

def analyze(full_bytes):
    """Devuelve (seq, title_name, conv, call, sections)."""
    files=_read(full_bytes); seq=_slide_order(files)
    EXCL=("INFORMACI","RESUMEN","DIAGNOST","DIAGNÓST","COMITÉ SEMANAL","COMITE SEMANAL")
    title=None; conv=[]; call=[]; sections=[]; cur=None
    for n in seq:
        raw=_text(files,n); t=raw.upper()
        if ("COMITÉ SEMANAL" in t) or ("COMITE SEMANAL" in t):
            title=n; continue
        (conv if "CONVENIO" in t else call).append(n)
        if not any(x in t for x in EXCL):          # carátula -> nueva sección
            nm=_first_text(files,n)[:45]
            cur={"title":nm,"names":[n]}; sections.append(cur)
        else:
            if cur is None: cur={"title":"Sección","names":[]}; sections.append(cur)
            cur["names"].append(n)
    return seq, title, conv, call, sections

def subset(full_bytes, keep_names):
    """Nuevo pptx con SOLO las láminas keep_names (en su orden original)."""
    files=_read(full_bytes); keep=set(keep_names); seq=_slide_order(files)
    remove=[n for n in seq if n not in keep]
    pres=files["ppt/presentation.xml"].decode("utf-8")
    rels=files["ppt/_rels/presentation.xml.rels"].decode("utf-8")
    rid2s={rid:t.split("/")[-1] for rid,t in re.findall(r'Id="([^"]+)"[^>]*Target="(slides/slide\d+\.xml)"',rels)}
    s2rid={v:k for k,v in rid2s.items()}
    ct=files["[Content_Types].xml"].decode("utf-8")
    for n in remove:
        rid=s2rid.get(n)
        if rid:
            pres=re.sub(rf'<p:sldId[^>]*r:id="{re.escape(rid)}"\s*/>','',pres)
            rels=re.sub(rf'<Relationship[^>]*Id="{re.escape(rid)}"[^>]*/>','',rels)
        rp=f"ppt/slides/_rels/{n}.rels"; notes=None
        if rp in files:
            m=re.search(r'Target="\.\./(notesSlides/notesSlide\d+\.xml)"', files[rp].decode("utf-8","ignore"))
            if m: notes=f"ppt/{m.group(1)}"
        files.pop(f"ppt/slides/{n}",None); files.pop(rp,None)
        if notes:
            files.pop(notes,None); files.pop(f"ppt/notesSlides/_rels/{notes.split('/')[-1]}.rels",None)
            ct=re.sub(rf'<Override PartName="/{re.escape(notes)}"[^>]*/>','',ct)
        ct=re.sub(rf'<Override PartName="/ppt/slides/{re.escape(n)}"[^>]*/>','',ct)
    files["ppt/presentation.xml"]=pres.encode("utf-8")
    files["ppt/_rels/presentation.xml.rels"]=rels.encode("utf-8")
    # podar medios que ya no referencia ninguna relación (láminas restantes / layouts / masters / tema)
    used=set()
    for k,d in files.items():
        if k.endswith(".rels"):
            for tgt in re.findall(r'Target="([^"]*media/[^"]+)"', d.decode("utf-8","ignore")):
                used.add(tgt.split("media/")[-1])
    for k in list(files):
        if k.startswith("ppt/media/") and k.split("/")[-1] not in used:
            files.pop(k,None)
    files["[Content_Types].xml"]=ct.encode("utf-8")
    return _write(files)

def _safe(name):
    return re.sub(r'[^A-Za-z0-9]+','_', name).strip('_')[:40] or "seccion"

def split_two(full_bytes):
    """(convenios_bytes, call_bytes) — cada uno con la portada global + sus láminas."""
    seq,title,conv,call,_=analyze(full_bytes)
    head=[title] if title else []
    return (subset(full_bytes, head+conv), subset(full_bytes, head+call))

def split_sections(full_bytes):
    """Lista de (nombre_archivo, bytes) — una presentación por campaña/sección."""
    seq,title,conv,call,sections=analyze(full_bytes)
    out=[]; used=set()
    for i,sec in enumerate(sections,1):
        base=_safe(sec["title"]); fn=f"{i:02d}_{base}.pptx"
        if fn in used: fn=f"{i:02d}_{base}_{i}.pptx"
        used.add(fn)
        out.append((fn, subset(full_bytes, sec["names"])))
    return out

def zip_files(named_bytes):
    o=io.BytesIO()
    with zipfile.ZipFile(o,"w",zipfile.ZIP_DEFLATED) as z:
        for name,data in named_bytes: z.writestr(name,data)
    return o.getvalue()
