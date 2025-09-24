# TD — Tablas Dinámicas por Fase (Streamlit)

App de Streamlit para analizar una o varias bases Excel (hoja **"Base de Datos"**) y generar **tablas tipo TD** por **Fase** (desde la columna **X**) con **segmentador por Mes** y **KPIs**.

## ✨ Características
- Carga **uno o varios** `.xlsx` y unifica.
- Mapeo fijo por letra de Excel: **K** (Cantidad de procedimientos), **W** (Valor del servicio), **AH** (Estado de la facturación); si no existen, cae a detección por nombre.
- Clasificación **Facturado / No Facturado** (reglas por texto del estado + existencia de número de factura).
- Segmentador por **Mes** con base **Mes Servicio** o **Mes Facturacion**.
- **Tablas por Estado** (por mes y totales).
- **TD por Fase** (filas = valores de la fase, columnas = Mes → Cant. Reg / Vlr. Servicio + Totales).
- **Export a Excel** con hojas: `Base_Filtrada`, `Estado_por_Mes`, `Estado_Total`, **`KPI_Mes`**, `TD_<Fase>` para **todas** las fases detectadas, y consolidado **`TD_FASES_TODAS`**.

## 🚀 Ejecutar localmente
```bash
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS / Linux
source .venv/bin/activate

pip install -r requirements.txt
streamlit run app_dashboard_td.py
```

## ☁️ Desplegar en Streamlit Community Cloud
1. Sube este repositorio a **GitHub**.
2. En https://share.streamlit.io/ → **New app**.
3. Selecciona el repo, rama y pon **`app_dashboard_td.py`** como *Main file*.
4. Deploy.

## 🧩 Notas
- La app asume hoja **"Base de Datos"** (si no existe, toma la primera).
- Las fases se detectan **desde la columna X** con detector robusto (texto/pocas categorías/palabras clave).
- Ajusta reglas de estados si tus textos son distintos (ver sección *Clasificación* en el código).
- Límite de nombre de hoja de Excel: 31 caracteres → se recorta a `TD_` + 25.

## 🛠 Estructura
```
.
├── app_dashboard_td.py
├── requirements.txt
├── .streamlit/
│   └── config.toml
├── .github/
│   └── workflows/
│       └── ci.yml
├── .gitignore
├── LICENSE
└── README.md
```
