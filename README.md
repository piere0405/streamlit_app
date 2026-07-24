# Actualizador del reporte semanal — BBVA TLM OUT

App web sencilla: subes el Excel de la semana y descargas el PowerPoint actualizado.
El diseño, colores, tipografía y posición de cada elemento **no cambian**; solo se
reescriben los datos y se regeneran los 4 gráficos con el mismo estilo.

## Archivos
- `app.py` — interfaz web (Streamlit).
- `pptx_updater.py` — motor (Excel → PPTX). Reutilizable también sin la web.
- `plantilla.pptx` — plantilla base (diseño fijo). **No la borres.**
- `requirements.txt` — dependencias.

## Instalar (una sola vez)
Necesitas Python 3.10+ instalado.
```bash
pip install -r requirements.txt
```

## Ejecutar la app
```bash
streamlit run app.py
```
Se abre en el navegador. Sube el Excel → botón **Actualizar presentación** → **Descargar**.

## Sin interfaz (opcional, para lote/automatización)
```python
from pptx_updater import update_presentation
tpl = open("plantilla.pptx","rb").read()
xls = open("CALL.xlsx","rb").read()
pptx, resumen, warnings = update_presentation(tpl, xls)
open("call_info_actualizado.pptx","wb").write(pptx)
print(resumen, warnings)
```

## Agregar / quitar / renombrar campañas (lo ÚNICO que se edita)
Todo se maneja desde la lista `PRODUCTS` al inicio de `pptx_updater.py`. Cada línea es:
`(clave_en_excel, texto_del_título_de_su_lámina, texto_de_su_fila_en_el_resumen, título_del_gráfico)`.

```python
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
```
- **Renombrar**: cambia los textos de esa línea.  
- **Agregar**: añade una línea (y asegúrate de que la plantilla ya tenga esa lámina).  
- **Quitar**: borra la línea.  
El número/orden de láminas, qué hoja del Excel usar, cuál es el gráfico de cada lámina y cuáles son las láminas de resumen **se autodetectan**. No hay números de lámina en el código.

### Límite conocido
Las láminas con **gráfico compuesto** (el que incluye un panel inferior de KPIs CET/EFECTIVIDAD/CIERRE) no se pueden regenerar porque esos datos no vienen en el Excel: la app actualiza sus textos y **avisa** que la imagen del gráfico quedó igual.

## Reglas de negocio (implementadas)
- **Periodo actual** = el `PERIODO` más alto del Excel (avance al día 11 + proyección).
- **% Logro** = `AVANCE / META_AVANCE`.
- **Avance promedio** (slide 2) = promedio simple del %Logro de los 4 productos.
- **Gráficos**: barras = AVANCE, línea = CIERRE; el periodo actual usa PROYECCIÓN
  y se dibuja con el último tramo punteado (igual que la plantilla).
- **Formato**: `<1.000` entero · miles → `K` · millones → `M`.

## Qué NO toca (datos que no vienen del Excel)
Dotación, Conectados, Pend. Selección/Inducción, Presupuesto, Semana, y las columnas
*Presupuesto* y *Conectados* de la tabla de la slide 2. Los estados
("EN AVANCE / EN OBSERVACIÓN / EN SEGUIMIENTO") tampoco se modifican.

## Requisitos del Excel
Columnas: `PRODUCTO`, `PERIODO`, `TIPO`, `MONTO`.
- PRODUCTO ∈ {OPERA OUT, PLD OUT, TC OUT, PORTAFOLIO}
- TIPO ∈ {AVANCE, CIERRE, META_MES, META_AVANCE, PROYECCION}
- PERIODO en formato AAAAMM (p. ej. 202607).

## Publicar en la web (gratis, opcional)
Sube esta carpeta a un repositorio de GitHub y despliégala en **Streamlit Community Cloud**
(streamlit.io) apuntando a `app.py`. Tendrás una URL para abrir desde cualquier equipo.

## Dos reportes en la misma app
La app **detecta solo** el tipo de reporte por las columnas del Excel y usa la plantilla correcta:
- Excel con `PRODUCTO` → **TLM** (`plantilla_tlm.pptx`, motor `pptx_updater.py`).
- Excel con `PLAZA` / `FAMILIA_CONVENIO` → **Convenios** (`plantilla_convenios.pptx`, motor `convenios_updater.py`).

### Reporte Convenios (láminas 2–7)
- Datos a nivel transacción; el motor agrega por PLAZA/PERIODO.
- Cada lámina de plaza regenera **3 gráficos**: *Avances del ámbito* (barras Monto Avance + línea Cierre + #OP + Ticket Promedio), *Participación por FAMILIA_CONVENIO* (dona) y *Asesores por Rango de Ventas* (barras 1 / 2-3 / >=4).
- Config en `convenios_updater.py`: lista `PLAZAS` (clave, texto del título, texto de la fila del resumen) y `FAMILY_COLORS`.
- **Día del avance**: por defecto **11**. Si el Excel trae una columna `DIA` (por ejemplo 11 en todas las filas), se usa ese número tanto en el título de cada gráfico ("al N de cada mes") como en el título de las páginas ("al N de {mes}").
- Supuestos (cambiables en el motor): la dona y los asesores se calculan sobre `CIERRE` del **periodo actual**; #OP = `Cantidad` de `AVANCE`; Ticket = MontoNeto/Cantidad de `AVANCE`; la línea usa `PROYECCION` en el periodo actual.

## Rediseño de la lámina de resumen (nivel directorio)
El módulo `summary_panel.py` regenera el cuerpo de la lámina de resumen y lo coloca como panel de alta resolución (respeta cabecera, logo y pie nativos). Aplica a las 3 láminas de resumen (2 de TLM + 1 de Convenios).
- **Tabla con más presencia**: ocupa la mitad inferior a lo ancho, filas altas, sombreado alterno y barra de avance.
- **4 cards**: *Activos totales* y *Pend. ingreso* (se leen de la plantilla), *Cumplen meta* y *Proyección total* (se calculan del Excel y se actualizan solos).
- Paleta intacta. Para cambiar los KPIs de las cards o las columnas de la tabla, editar `summary_panel.py`.
