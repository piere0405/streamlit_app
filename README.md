# Actualizador de Reportes Semanales — PlusMetas · MF Asesoría y Consultoría

App en Streamlit que actualiza los PowerPoint semanales (Campañas/Telemarketing y Convenios)
a partir de los Excel de la semana, manteniendo diseño, colores y formato.

## Correr en local
```bash
pip install -r requirements.txt
streamlit run app_consolidado.py
```

## Desplegar en Streamlit Cloud
1. Sube esta carpeta a un repositorio de GitHub.
2. Entra a https://share.streamlit.io → **New app**.
3. Elige el repositorio y la rama (`main`).
4. **Main file path:** `app_consolidado.py`
5. Deploy. Streamlit instala solo lo de `requirements.txt`.

## Archivos
- `app_consolidado.py` — app principal (Comité de Campañas, 2 Excel → 1 PPT).
- `app.py` — app multi-reporte (Campañas / TLM / Convenios).
- `campanias_updater.py`, `convenios_updater.py`, `pptx_updater.py`, `summary_panel.py` — motores.
- `plantilla_campanias.pptx`, `plantilla_tlm.pptx`, `plantilla_convenios.pptx` — plantillas base.
- `logo_plusmetas.png`, `logo_mf.png` — logos de la cabecera.
