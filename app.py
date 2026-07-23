from datetime import datetime
import glob
import os
import pandas as pd
import plotly.express as px
import streamlit as st

st.set_page_config(
    page_title="Dashboard de Ventas - Mystic",
    page_icon="📊",
    layout="wide",
)


@st.cache_data
def load_data():
  # Busca automáticamente cualquier archivo Excel (.xlsx) en la carpeta
  archivos_excel = glob.glob("*.xlsx")

  if not archivos_excel:
    raise FileNotFoundError(
        "No se encontró ningún archivo Excel en la carpeta. Coloca tu archivo"
        " aquí."
    )

  file_path = archivos_excel[0]

  df = pd.read_excel(file_path, sheet_name="report")
  df["Total Venta"] = (
      df["Informar la Cantidad Vendida"] * df["Informar precio de la venta"]
  )
  df["Fecha y hora de la encuesta"] = pd.to_datetime(
      df["Fecha y hora de la encuesta"]
  )
  df["Fecha"] = df["Fecha y hora de la encuesta"].dt.date
  return df, file_path


try:
  df, archivo_usado = load_data()
except Exception as e:
  st.error(f"Error al cargar el archivo Excel: {e}.")
  st.stop()

st.title("📈 Dashboard Interactivo de Ventas")
st.markdown(
    f"Visualización y análisis comercial en tiempo real (Fuente: *{archivo_usado}*)"
)

# --- BARRA LATERAL: FILTROS ---
st.sidebar.header("🔍 Filtros Globales")

metrica_seleccionada = st.sidebar.radio(
    "Ver gráficos en:", ("Dólares (US$)", "Unidades")
)

st.sidebar.markdown("---")

regional_options = ["Todas"] + sorted(df["Regional"].unique().tolist())
selected_regional = st.sidebar.selectbox("Seleccione Regional", regional_options)

supercat_options = ["Todas"] + sorted(df["Supercategoría"].unique().tolist())
selected_supercat = st.sidebar.selectbox(
    "Seleccione Supercategoría", supercat_options
)

modelo_options = ["Todos"] + sorted(
    df["Informar modelo "].dropna().unique().tolist()
)
selected_modelo = st.sidebar.selectbox("Seleccione Modelo", modelo_options)

# Filtro por PDV (Punto de Venta)
pdv_options = ["Todos"] + sorted(df["PDV"].dropna().unique().tolist())
selected_pdv = st.sidebar.selectbox("Seleccione PDV", pdv_options)

empleado_options = ["Todos"] + sorted(df["Empleado"].unique().tolist())
selected_empleado = st.sidebar.selectbox("Seleccione Empleado", empleado_options)

# Aplicar Filtros
df_filtered = df.copy()
if selected_regional != "Todas":
  df_filtered = df_filtered[df_filtered["Regional"] == selected_regional]
if selected_supercat != "Todas":
  df_filtered = df_filtered[df_filtered["Supercategoría"] == selected_supercat]
if selected_modelo != "Todos":
  df_filtered = df_filtered[df_filtered["Informar modelo "] == selected_modelo]
if selected_pdv != "Todos":
  df_filtered = df_filtered[df_filtered["PDV"] == selected_pdv]
if selected_empleado != "Todos":
  df_filtered = df_filtered[df_filtered["Empleado"] == selected_empleado]

total_revenue = df_filtered["Total Venta"].sum()
total_units = df_filtered["Informar la Cantidad Vendida"].sum()
total_transactions = len(df_filtered)
avg_ticket = total_revenue / total_transactions if total_transactions > 0 else 0

# --- TARJETAS DE MÉTRICAS (KPIs) ---
col1, col2, col3, col4 = st.columns(4)
with col1:
  st.metric(
      label="Venta Total (USD)",
      value=f"${total_revenue:,.2f}"
      .replace(",", "_")
      .replace(".", ",")
      .replace("_", "."),
  )
with col2:
  st.metric(label="Unidades Vendidas", value=f"{total_units:,}")
with col3:
  st.metric(label="Transacciones", value=f"{total_transactions:,}")
with col4:
  st.metric(
      label="Ticket Promedio",
      value=f"${avg_ticket:,.2f}"
      .replace(",", "_")
      .replace(".", ",")
      .replace("_", "."),
  )

st.markdown("---")

if metrica_seleccionada == "Dólares (US$)":
  columna_metrica = "Total Venta"
  etiqueta_eje = "Venta (USD)"
else:
  columna_metrica = "Informar la Cantidad Vendida"
  etiqueta_eje = "Cantidad (Unidades)"

# --- GRÁFICOS INTERACTIVOS (PLOTLY - ESCALA DE AZULES) ---
row1_col1, row1_col2 = st.columns(2)

with row1_col1:
  st.subheader(f"💡 Rendimiento por Supercategoría ({metrica_seleccionada})")
  if not df_filtered.empty:
    df_cat = (
        df_filtered.groupby("Supercategoría")[columna_metrica]
        .sum()
        .reset_index()
        .sort_values(by=columna_metrica, ascending=False)
    )

    fig_cat = px.bar(
        df_cat,
        x="Supercategoría",
        y=columna_metrica,
        text=columna_metrica,
        color=columna_metrica,
        color_continuous_scale="Blues",
        template="plotly_white",
    )

    if metrica_seleccionada == "Dólares (US$)":
      fig_cat.update_traces(texttemplate="$%{y:,.2f}", textposition="outside")
    else:
      fig_cat.update_traces(texttemplate="%{y:,.0f}", textposition="outside")

    fig_cat.update_layout(
        coloraxis_showscale=False,
        showlegend=False,
        xaxis_title="",
        yaxis_title=etiqueta_eje,
        xaxis={"categoryorder": "total descending"},
    )
    st.plotly_chart(fig_cat, use_container_width=True)
  else:
    st.info("No hay datos para los filtros seleccionados.")

with row1_col2:
  st.subheader(f"🗺️ Distribución por Regional ({metrica_seleccionada})")
  if not df_filtered.empty:
    fig_reg = px.pie(
        df_filtered.groupby("Regional")[columna_metrica].sum().reset_index(),
        names="Regional",
        values=columna_metrica,
        hole=0.4,
        color_discrete_sequence=px.colors.sequential.Blues_r,
        template="plotly_white",
    )
    fig_reg.update_traces(textposition="inside", textinfo="percent+label")
    st.plotly_chart(fig_reg, use_container_width=True)
  else:
    st.info("No hay datos para los filtros seleccionados.")

# Fila 2: Rendimiento por Modelo y Tendencia Diaria
row2_col1, row2_col2 = st.columns(2)

with row2_col1:
  st.subheader(f"🏷️ Rendimiento por Modelo ({metrica_seleccionada})")
  if not df_filtered.empty:
    df_mod = (
        df_filtered.groupby("Informar modelo ")[columna_metrica]
        .sum()
        .reset_index()
        .sort_values(by=columna_metrica, ascending=False)
    )

    fig_mod = px.bar(
        df_mod,
        x="Informar modelo ",
        y=columna_metrica,
        text=columna_metrica,
        color=columna_metrica,
        color_continuous_scale="Blues",
        template="plotly_white",
    )

    if metrica_seleccionada == "Dólares (US$)":
      fig_mod.update_traces(texttemplate="$%{y:,.2f}", textposition="outside")
    else:
      fig_mod.update_traces(texttemplate="%{y:,.0f}", textposition="outside")

    fig_mod.update_layout(
        coloraxis_showscale=False,
        showlegend=False,
        xaxis_title="Modelo",
        yaxis_title=etiqueta_eje,
        xaxis={"categoryorder": "total descending"},
    )
    st.plotly_chart(fig_mod, use_container_width=True)
  else:
    st.info("No hay datos para los filtros seleccionados.")

with row2_col2:
  st.subheader(f"📅 Tendencia Diaria ({metrica_seleccionada})")
  if not df_filtered.empty:
    df_daily = (
        df_filtered.groupby("Fecha")[columna_metrica].sum().reset_index()
    )
    fig_time = px.line(
        df_daily,
        x="Fecha",
        y=columna_metrica,
        markers=True,
        color_discrete_sequence=["#1f77b4"],
        template="plotly_white",
    )
    fig_time.update_layout(xaxis_title="Fecha", yaxis_title=etiqueta_eje)
    st.plotly_chart(fig_time, use_container_width=True)
  else:
    st.info("No hay datos para los filtros seleccionados.")

# Fila 3: Gráfico por Punto de Venta (PDV) y Comisión Acumulada
row3_col1, row3_col2 = st.columns(2)

with row3_col1:
  st.subheader(f"🏪 Rendimiento por Punto de Venta (PDV)")
  if not df_filtered.empty:
    df_pdv = (
        df_filtered.groupby("PDV")[columna_metrica]
        .sum()
        .reset_index()
        .sort_values(by=columna_metrica, ascending=True)
    )

    fig_pdv = px.bar(
        df_pdv,
        x=columna_metrica,
        y="PDV",
        orientation="h",
        text=columna_metrica,
        color=columna_metrica,
        color_continuous_scale="Blues",
        template="plotly_white",
    )

    if metrica_seleccionada == "Dólares (US$)":
      fig_pdv.update_traces(texttemplate="$%{x:,.2f}", textposition="outside")
    else:
      fig_pdv.update_traces(texttemplate="%{x:,.0f}", textposition="outside")

    fig_pdv.update_layout(
        coloraxis_showscale=False,
        xaxis_title=etiqueta_eje,
        yaxis_title="",
        height=500,
    )
    st.plotly_chart(fig_pdv, use_container_width=True)
  else:
    st.info("No hay datos para los filtros seleccionados.")

with row3_col2:
  st.subheader(
      f"🎯 Comisión Acumulada (Ventas >= $120) - por Empleado"
      f" ({metrica_seleccionada})"
  )
  df_comision = df_filtered[df_filtered["Informar precio de la venta"] >= 120]

  if not df_comision.empty:
    df_comm_emp = (
        df_comision.groupby("Empleado")[columna_metrica]
        .sum()
        .reset_index()
        .sort_values(by=columna_metrica, ascending=True)
    )

    fig_comm = px.bar(
        df_comm_emp,
        x=columna_metrica,
        y="Empleado",
        orientation="h",
        text=columna_metrica,
        color=columna_metrica,
        color_continuous_scale="Blues",
        template="plotly_white",
    )

    if metrica_seleccionada == "Dólares (US$)":
      fig_comm.update_traces(texttemplate="$%{x:,.2f}", textposition="outside")
    else:
      fig_comm.update_traces(texttemplate="%{x:,.0f}", textposition="outside")

    fig_comm.update_layout(
        coloraxis_showscale=False, xaxis_title=etiqueta_eje, yaxis_title=""
    )
    st.plotly_chart(fig_comm, use_container_width=True)
  else:
    st.info("No hay registros que cumplan con la condición de precio >= $120.")

# Fila 4: Top Empleados General
row4_col1 = st.container()
with row4_col1:
  st.subheader(f"🏆 Top Empleados - General ({metrica_seleccionada})")
  if not df_filtered.empty:
    df_emp = (
        df_filtered.groupby("Empleado")[columna_metrica]
        .sum()
        .reset_index()
        .sort_values(by=columna_metrica, ascending=True)
    )

    fig_emp = px.bar(
        df_emp,
        x=columna_metrica,
        y="Empleado",
        orientation="h",
        text=columna_metrica,
        color=columna_metrica,
        color_continuous_scale="Blues",
        template="plotly_white",
    )
    if metrica_seleccionada == "Dólares (US$)":
      fig_emp.update_traces(texttemplate="$%{x:,.2f}", textposition="outside")
    else:
      fig_emp.update_traces(texttemplate="%{x:,.0f}", textposition="outside")

    fig_emp.update_layout(
        coloraxis_showscale=False, xaxis_title=etiqueta_eje, yaxis_title=""
    )
    st.plotly_chart(fig_emp, use_container_width=True)
  else:
    st.info("No hay datos para los filtros seleccionados.")

st.markdown("---")

# --- TABLA DE DETALLES ---
st.subheader("📋 Detalle de Registros Filtrados")
st.dataframe(
    df_filtered[[
        "Fecha y hora de la encuesta",
        "Regional",
        "Ciudad",
        "PDV",
        "Empleado",
        "Supercategoría",
        "Informar modelo ",
        "Informar la Cantidad Vendida",
        "Informar precio de la venta",
        "Total Venta",
    ]],
    use_container_width=True,
)