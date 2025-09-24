
# -*- coding: utf-8 -*-
import io, re, os
import numpy as np
import pandas as pd
import streamlit as st

st.set_page_config(page_title="TD — Tablas Dinámicas por Fase", page_icon="📑", layout="wide")
st.title("📑 TD — Tablas Dinámicas por Fase (sin gráficos)")

# ------------------------------
# Entrada
# ------------------------------
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
base = base.loc[:, ~base.columns.duplicated()]  # columnas únicas

# ------------------------------
# Helpers
# ------------------------------
def _norm(s: str) -> str:
    return str(s).strip().lower()

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
    return idx - 1  # zero-based

def _get_col_by_letter(df: pd.DataFrame, letter: str):
    i = _col_letter_to_index(letter)
    return df.columns[i] if 0 <= i < len(df.columns) else None

def _flatten_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Aplana columnas MultiIndex en 'Mes — Métrica' para poder exportar a Excel con index=False."""
    if isinstance(df.columns, pd.MultiIndex):
        new_cols = []
        for tup in df.columns:
            if isinstance(tup, tuple):
                parts = [str(x) for x in tup if (x is not None and str(x) != "")]
                new_cols.append(" — ".join(parts))
            else:
                new_cols.append(str(tup))
        df = df.copy()
        df.columns = new_cols
    return df

# ------------------------------
# Mapeos de columnas por letra (prioridad) y por nombre (fallback)
# ------------------------------
# Usuario indicó: K = Cantidad de procedimientos, W = Valor del servicio, AH = Estado de la facturación
col_cant = _get_col_by_letter(base, "K")
col_val  = _get_col_by_letter(base, "W")
col_est  = _get_col_by_letter(base, "AH")

# Valor del servicio
if col_val is not None:
    valor_col = col_val
else:
    valor_candidates = [c for c in base.columns if ("valor" in _norm(c)) and ("serv" in _norm(c) or _norm(c) in ["valor","valor total","valor unitario"])]
    valor_col = valor_candidates[0] if valor_candidates else None
base["_VALOR_"] = to_number(base[valor_col]) if valor_col else 0.0

# Cantidad de procedimientos
if col_cant is not None:
    cant_col = col_cant
else:
    cant_candidates = [c for c in base.columns if ("cantidad" in _norm(c)) and ("proced" in _norm(c))]
    cant_col = cant_candidates[0] if cant_candidates else None
base["_CANT_PROC_"] = pd.to_numeric(base[cant_col], errors="coerce").fillna(0).astype(int) if cant_col else 1

# Estado de la facturación (robusto)
if col_est is not None:
    estado_col = col_est
else:
    estado_col = None
    for c in base.columns:
        cn = _norm(c).replace(" ", "")
        if (("estado" in cn) and ("factur" in cn)):
            estado_col = c
            break
base["_ESTADO_"] = base[estado_col].astype(str).str.strip() if estado_col else "Sin estado"

# Clasificación Facturado / No Facturado
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

# Segmentador Mes
meses = sorted(base["_MES_"].dropna().unique().tolist(), key=lambda x: x)
sel_meses = st.sidebar.multiselect("Mes (segmentador)", meses, default=meses)
base = base[base["_MES_"].isin(sel_meses)].copy()

# ------------------------------
# Detección de fases desde la columna X
# ------------------------------
def _cols_from_letter(df: pd.DataFrame, letter="X"):
    i = _col_letter_to_index(letter)
    return list(df.columns[i:]) if i < len(df.columns) else []

cols_X = _cols_from_letter(base, "X")
phase_cols = [c for c in cols_X if isinstance(c, str) and ("Fase" in c and "Verficacion" in c)]

# ------------------------------
# KPI (sin f-string para evitar conflictos)
# ------------------------------
registros = int(base["_CANT_PROC_"].sum())
valor_total = float(base["_VALOR_"].sum())
valor_fact = float(base.loc[base["_FACTURADO_"]=="Facturado","_VALOR_"].sum())
valor_nofact = float(base.loc[base["_FACTURADO_"]=="No Facturado","_VALOR_"].sum())
msg = "Registros filtrados: {registros:,} | Valor: ${valor_total:,.0f} | Facturado: ${valor_fact:,.0f} | No Facturado: ${valor_nofact:,.0f}".format(
    registros=registros, valor_total=valor_total, valor_fact=valor_fact, valor_nofact=valor_nofact
)
st.success(msg)

# ------------------------------
# Resúmenes por Estado (por Mes y total)
# ------------------------------
st.markdown("### 📊 Resumen por **Estado de la facturación**")
res_estado_mes = base.groupby(["_MES_", "_ESTADO_"], dropna=False).agg(
    Cant_Serv=("_CANT_PROC_", "sum"),
    Vlr_Servicio=("_VALOR_", "sum")
).reset_index().rename(columns={"_MES_":"Mes","_ESTADO_":"Estado"})
st.subheader("Por Mes y Estado")
st.dataframe(res_estado_mes)

st.subheader("Total por Estado (meses seleccionados)")
res_estado_total = base.groupby(["_ESTADO_"], dropna=False).agg(
    Cant_Serv=("_CANT_PROC_", "sum"),
    Vlr_Servicio=("_VALOR_", "sum")
).reset_index().rename(columns={"_ESTADO_":"Estado"})
st.dataframe(res_estado_total)

# ------------------------------
# Constructor de tablas formato TD por Fase (Mes → Cant. Reg / Vlr. Servicio)
# ------------------------------
def build_td_table(df: pd.DataFrame, phase_col: str, meses_order: list):
    grp = df.groupby(["_MES_", phase_col], dropna=False).agg(
        Cant_Reg=("_CANT_PROC_", "sum"),
        Vlr_Servicio=("_VALOR_", "sum")
    ).reset_index().rename(columns={phase_col: "Fase"})
    fases = grp["Fase"].dropna().astype(str).unique().tolist()
    wide = pd.DataFrame({"Fase": fases}).set_index("Fase")
    for m in meses_order:
        sub = grp[grp["_MES_"]==m].set_index("Fase")
        wide[(m, "Cant. Reg")] = sub["Cant_Reg"]
        wide[(m, "Vlr. Servicio")] = sub["Vlr_Servicio"]
    wide[("Total", "Cant. Reg")] = grp.groupby("Fase")["Cant_Reg"].sum()
    wide[("Total", "Vlr. Servicio")] = grp.groupby("Fase")["Vlr_Servicio"].sum()
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
    st.warning("No se detectaron columnas de fases a partir de la columna X. Si el nombre difiere, indícamelo.")

# ------------------------------
# Exportar Excel
# ------------------------------
out = io.BytesIO()
with pd.ExcelWriter(out, engine="xlsxwriter") as writer:
    base.to_excel(writer, index=False, sheet_name="Base_Filtrada")
    # Estados
    res_estado_mes.to_excel(writer, index=False, sheet_name="Estado_por_Mes")
    res_estado_total.to_excel(writer, index=False, sheet_name="Estado_Total")
    # Fases TD
    if phase_cols:
        for c in phase_cols:
            tbl, _ = build_td_table(base, c, meses_order)
            tbl = _flatten_columns(tbl)
            sheet = re.sub(r'[^A-Za-z0-9]', '_', c)[:25]
            tbl.to_excel(writer, index=False, sheet_name=f"TD_{sheet}")

st.download_button("Descargar Excel TD (fases)", data=out.getvalue(), file_name="td_tablas_fases.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
