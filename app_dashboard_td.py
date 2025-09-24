
# -*- coding: utf-8 -*-
import io, re
import numpy as np
import unicodedata
import pandas as pd
import streamlit as st

st.set_page_config(page_title="TD — Tablas Dinámicas por Fase", page_icon="📑", layout="wide")
st.title("📑 TD — Tablas Dinámicas por Fase (solo tablas)")

# ==============================
# Carga
# ==============================
st.sidebar.header("Entrada de datos")
uploaded_files = st.sidebar.file_uploader("Cargar 1 o varios .xlsx", type=["xlsx"], accept_multiple_files=True)
base_mes = st.sidebar.selectbox("Base de mes", ["Mes Servicio", "Mes Facturacion"], index=0)

def read_base(xlfile):
    try:
        xls = pd.ExcelFile(xlfile)
        target = "Base de Datos" if "Base de Datos" in xls.sheet_names else xls.sheet_names[0]
        df = pd.read_excel(xls, sheet_name=target)
        return df
    except Exception as e:
        st.error(f"Error leyendo: {getattr(xlfile,'name','archivo')}: {e}")
        return None

dfs = []
if uploaded_files:
    for f in uploaded_files:
        df = read_base(f)
        if df is not None:
            df["__archivo__"] = getattr(f, "name", "archivo.xlsx")
            dfs.append(df)
if not dfs:
    st.info("Cargue al menos un archivo .xlsx (hoja 'Base de Datos').")
    st.stop()

base = pd.concat(dfs, ignore_index=True, sort=False)
base = base.loc[:, ~base.columns.duplicated()]

# ==============================
# Helpers
# ==============================
def _norm(s): return str(s).strip().lower()

def _norm_noaccents(s: str) -> str:
    s = str(s)
    s = unicodedata.normalize('NFD', s)
    s = ''.join(ch for ch in s if unicodedata.category(ch) != 'Mn')
    return s.strip().lower()


def to_number(series: pd.Series) -> pd.Series:
    return pd.to_numeric(
        series.astype(str)
              .str.replace(r'[^0-9\-,\.]', '', regex=True)
              .str.replace(',', '', regex=False),
        errors='coerce'
    ).fillna(0.0)

def _col_letter_to_index(letter: str) -> int:
    letter = letter.upper()
    idx = 0
    for ch in letter:
        idx = idx * 26 + (ord(ch) - ord('A') + 1)
    return idx - 1

def _get_col_by_letter(df: pd.DataFrame, letter: str):
    i = _col_letter_to_index(letter)
    return df.columns[i] if 0 <= i < len(df.columns) else None

def _flatten_columns(df: pd.DataFrame) -> pd.DataFrame:
    if isinstance(df.columns, pd.MultiIndex):
        new_cols = []
        for tup in df.columns:
            if isinstance(tup, tuple):
                parts = [str(x) for x in tup if (x is not None and str(x) != "")]
                new_cols.append(" — ".join(parts))
            else:
                new_cols.append(str(tup))
        df = df.copy(); df.columns = new_cols
    return df

# ==============================
# Columnas principales (letras fijas + fallback por nombre)
# ==============================
# Usuario: K = Cantidad proc; W = Valor servicio; AH = Estado facturación
col_cant = _get_col_by_letter(base, "K")
col_val  = _get_col_by_letter(base, "W")
col_est  = _get_col_by_letter(base, "AH")

valor_col = col_val if col_val is not None else next((c for c in base.columns if ("valor" in _norm(c) and ("serv" in _norm(c) or _norm(c) in ["valor","valor total","valor unitario"]))), None)
base["_VALOR_"] = to_number(base[valor_col]) if valor_col else 0.0

cant_col = col_cant if col_cant is not None else next((c for c in base.columns if ("cantidad" in _norm(c) and "proced" in _norm(c))), None)
base["_CANT_PROC_"] = pd.to_numeric(base[cant_col], errors="coerce").fillna(0).astype(int) if cant_col else 1

estado_col = col_est
if estado_col is None:
    for c in base.columns:
        cn = _norm(c).replace(" ", "")
        if ("estado" in cn and "factur" in cn):
            estado_col = c; break
base["_ESTADO_"] = base[estado_col].astype(str).str.strip() if estado_col else "Sin estado"

# Facturado / No Facturado (reglas)
estado_lc = base["_ESTADO_"].str.lower()
no_fact = estado_lc.str.contains("no factur") | estado_lc.str.contains("sin factur") | estado_lc.str.contains("no aplica")
si_fact = estado_lc.str.contains("factur") & (~no_fact)
if "Factura" in base.columns:
    si_fact = si_fact | base["Factura"].notna()
base["_FACTURADO_"] = np.where(si_fact, "Facturado", "No Facturado")

# Mes base
if base_mes in base.columns:
    base["_MES_"] = base[base_mes].astype(str).str.strip()
elif "Mes Servicio" in base.columns:
    base["_MES_"] = base["Mes Servicio"].astype(str).str.strip()
else:
    base["_MES_"] = "Sin mes"

# Segmentador de Mes
meses = sorted(base["_MES_"].dropna().unique().tolist(), key=lambda x: x)
sel_meses = st.sidebar.multiselect("Mes (segmentador)", meses, default=meses)
base = base[base["_MES_"].isin(sel_meses)].copy()


# ==============================
# Fases a partir de columna X (inclusivo)
# ==============================
def detect_phase_cols(df: pd.DataFrame) -> list:
    i = _col_letter_to_index("X")
    cand = list(df.columns[i:]) if i < len(df.columns) else []
    phase_cols = [c for c in cand if isinstance(c, str) and (("fase" in c.lower()) or ("verific" in c.lower()))]
    # Si nada matchea, usa todas las columnas de X en adelante que no sean numéricas
    if not phase_cols:
        phase_cols = [c for c in cand if not pd.api.types.is_numeric_dtype(df[c])]
    # Unicos y en orden
    seen = set(); out = []
    for c in phase_cols:
        if c not in seen: seen.add(c); out.append(c)
    return out

phase_cols = detect_phase_cols(base)

# ==============================
# KPI
# ==============================
registros = int(base["_CANT_PROC_"].sum())
valor_total = float(base["_VALOR_"].sum())
valor_fact = float(base.loc[base["_FACTURADO_"]=="Facturado","_VALOR_"].sum())
valor_nofact = float(base.loc[base["_FACTURADO_"]=="No Facturado","_VALOR_"].sum())
msg = "Registros filtrados: {registros:,} | Valor: ${valor_total:,.0f} | Facturado: ${valor_fact:,.0f} | No Facturado: ${valor_nofact:,.0f}".format(
    registros=registros, valor_total=valor_total, valor_fact=valor_fact, valor_nofact=valor_nofact
)
st.success(msg)

# ==============================
# TD por Fase
# ==============================
def build_td_table(df: pd.DataFrame, phase_col: str, meses_order: list):
    tmp = df.rename(columns={phase_col: "__PHASE__", "_MES_": "__MES__"}).copy()
    grp = tmp.groupby(["__MES__", "__PHASE__"], dropna=False, as_index=False).agg(
        Cant_Reg=("_CANT_PROC_", "sum"),
        Vlr_Servicio=("_VALOR_", "sum")
    )
    grp = grp.rename(columns={"__PHASE__": "Fase", "__MES__": "_MES_"})
    fases = grp["Fase"].dropna().astype(str).unique().tolist()
    wide = pd.DataFrame({"Fase": fases}).set_index("Fase")
    for m in meses_order:
        sub = grp[grp["_MES_"]==m].set_index("Fase")
        wide[(m, "Cant. Reg")] = sub["Cant_Reg"]
        wide[(m, "Vlr. Servicio")] = sub["Vlr_Servicio"]
    totals = grp.groupby("Fase", as_index=True).agg(
        Cant_Reg=("Cant_Reg","sum"),
        Vlr_Servicio=("Vlr_Servicio","sum")
    )
    wide[("Total", "Cant. Reg")] = totals["Cant_Reg"]
    wide[("Total", "Vlr. Servicio")] = totals["Vlr_Servicio"]
    cols = []
    for m in meses_order:
        cols.extend([(m, "Cant. Reg"), (m, "Vlr. Servicio")])
    cols.extend([("Total","Cant. Reg"), ("Total","Vlr. Servicio")])
    wide = wide.reindex(columns=pd.MultiIndex.from_tuples(cols)).fillna(0)
    return wide.reset_index(), grp

meses_order = sorted(base["_MES_"].unique().tolist(), key=lambda x: x)

st.markdown("### 📑 Tablas dinámicas por Fase (formato TD)")
if phase_cols:
    for c in phase_cols:
        st.subheader(c)
        tbl, grp = build_td_table(base, c, meses_order)
        st.dataframe(tbl)
else:
    st.warning("No se detectaron columnas de Fase a partir de la columna X.")

# ==============================
# KPI_Mes + Fact/NoFact por Mes (Mes Servicio + Estado Facturación)
# ==============================
mes_serv_col = 'Mes Servicio' if 'Mes Servicio' in base.columns else '_MES_'
kpi_det = base.groupby([mes_serv_col, '_FACTURADO_'], dropna=False).agg(
    Cant_Serv=('_CANT_PROC_', 'sum'),
    Valor_Servicio=('_VALOR_', 'sum')
).reset_index().rename(columns={mes_serv_col:'Mes', '_FACTURADO_':'Fact_Status'})

# Pivot para tener columnas separadas por Facturado/No
qty = kpi_det.pivot(index='Mes', columns='Fact_Status', values='Cant_Serv').fillna(0)
val = kpi_det.pivot(index='Mes', columns='Fact_Status', values='Valor_Servicio').fillna(0)
for col in ['Facturado','No Facturado']:
    if col not in qty.columns: qty[col] = 0
    if col not in val.columns: val[col] = 0.0

kpi_mes = qty.copy()
kpi_mes = kpi_mes.rename(columns={'Facturado':'Cant_Facturado','No Facturado':'Cant_No_Facturado'})
kpi_mes['Valor_Facturado'] = val['Facturado']
kpi_mes['Valor_No_Facturado'] = val['No Facturado']
kpi_mes['Cant_Serv_Total'] = kpi_mes['Cant_Facturado'] + kpi_mes['Cant_No_Facturado']
kpi_mes['Valor_Total'] = kpi_mes['Valor_Facturado'] + kpi_mes['Valor_No_Facturado']
kpi_mes['Valor_Promedio_Servicio'] = kpi_mes.apply(lambda r: (r['Valor_Total']/r['Cant_Serv_Total']) if r['Cant_Serv_Total'] else 0.0, axis=1)
kpi_mes['%_Valor_Facturado'] = kpi_mes.apply(lambda r: (r['Valor_Facturado']/r['Valor_Total']) if r['Valor_Total'] else 0.0, axis=1)
kpi_mes['%_Valor_No_Facturado'] = kpi_mes.apply(lambda r: (r['Valor_No_Facturado']/r['Valor_Total']) if r['Valor_Total'] else 0.0, axis=1)
kpi_mes = kpi_mes.reset_index()

# Tabla explícita (detalle) por Mes Servicio y Estado para export
fact_nofact_mes = base.groupby([mes_serv_col, '_FACTURADO_'], dropna=False).agg(
    Cant_Serv=('_CANT_PROC_', 'sum'),
    Vlr_Servicio=('_VALOR_', 'sum')
).reset_index().rename(columns={mes_serv_col:'Mes','_FACTURADO_':'Estado_Fact'})

st.markdown('### 📈 KPI por Mes del Servicio (basado en Estado de Factura)')
st.dataframe(kpi_mes)



# ==============================
# Export
# ==============================

# ==============================
# Export limpio
# ==============================
out = io.BytesIO()
with pd.ExcelWriter(out, engine="xlsxwriter") as writer:
    # Base filtrada
    base.to_excel(writer, index=False, sheet_name="Base_Filtrada")
    # KPI
    kpi_mes.to_excel(writer, index=False, sheet_name="KPI_Mes")
    # Detalle Mes x Estado
    fact_nofact_mes.to_excel(writer, index=False, sheet_name="Fact_NoFact_por_Mes")
    # Tablas TD por cada fase + consolidado
    all_grps = []
    try:
        meses_order  # ensure exists
    except NameError:
        meses_order = sorted(base["_MES_"].unique().tolist())
    if 'phase_cols' in globals() and phase_cols:
        for c in phase_cols:
            tbl, grp = build_td_table(base, c, meses_order)
            tbl = _flatten_columns(tbl)
            sheet = re.sub(r'[^A-Za-z0-9]', '_', str(c))[:25] or "Fase"
            tbl.to_excel(writer, index=False, sheet_name=f"TD_{sheet}")
            grp2 = grp.copy()
            grp2.insert(0, "Columna_Fase", c)
            all_grps.append(grp2)
        if all_grps:
            td_all = pd.concat(all_grps, ignore_index=True)
            td_all.rename(columns={'_MES_':'Mes'}, inplace=True)
            td_all.to_excel(writer, index=False, sheet_name="TD_FASES_TODAS")

st.download_button(
    "Descargar Excel TD (fases)",
    data=out.getvalue(),
    file_name="td_tablas_fases.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
