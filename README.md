# TD — Tablas Dinámicas por Fase (Streamlit)

App de Streamlit para analizar una o varias bases **Excel** (hoja *"Base de Datos"*) y generar **tablas tipo TD** por **Fase** (desde la columna **X**), con **segmentador por Mes** y **KPIs**.

## Cambios claves
- **Detección robusta de fases**: toma **todas** las columnas a partir de **X** que tengan texto/pocas categorías o contengan palabras clave (*fase, verificación, validación, malla, código, medidas*).
- **Exportación completa**: crea una hoja **TD_<Fase>** por **cada** fase detectada y una hoja consolidada **TD_FASES_TODAS**.
- Mapeo fijo por letras: **K** (Cantidad de procedimientos), **W** (Valor del servicio), **AH** (Estado de la facturación).

## Ejecutar localmente
```bash
python -m venv .venv
# Windows
.venv\Scripts\activate
# Linux/Mac
source .venv/bin/activate

pip install -r requirements.txt
streamlit run app_dashboard_td.py
```

## Desplegar en Streamlit Community Cloud
1. Sube este repo a GitHub.
2. En Streamlit Cloud → *New app*, selecciona repo y *main file* `app_dashboard_td.py`.
3. Deploy.

## Notas
- La lógica de **Facturado/No** se basa en texto del **Estado** y presencia de **Factura**. Ajusta reglas si tus estados difieren.
- Los nombres de hojas Excel se recortan a 31 chars (`TD_` + hasta 25).

---
Licencia: MIT
