from datetime import datetime
import glob
import os
import pandas as pd
import plotly.express as px
import streamlit as st
from PIL import Image

st.set_page_config(
    page_title="Dashboard de Ventas - Mystic",
    page_icon="📊",
    layout="wide",
)

# --- OCULTAR MENÚ DE NAVEGACIÓN AUTOMÁTICA EN SIDEBAR ---
st.markdown(
    """
    <style>
    div[data-testid="stSidebarNav"] {display: none;}
    </style>
    """,
    unsafe_allow_html=True,
)

# --- PALETA DE COLORES PERSONALIZADA ---
MYSTIC_BLUE_PALETTE = [
    "#182232",  # Azul marino muy oscuro / casi negro
    "#3f51b5",  # Azul índigo / vibrante
    "#7ea6e0",  # Azul celeste medio
    "#6b7280",  # Gris acero / pizarra oscuro
    "#9ca3af",  # Gris pizarra claro
    "#364f6b",  # Azul pizarra corporativo
    "#8fa4c7",  # Azul lavanda / grisáceo
    "#b9d2ee",  # Celeste pastel muy claro
]


@st.cache_data
def load_data():
  file_path = "Venta acumulada del mes.xlsx"
  if not os.path.exists(file_path):
    archivos_excel = glob.glob("*.xlsx")
    if not archivos_excel:
      raise FileNotFoundError(
          "No se encontró el archivo 'Venta acumulada del mes.xlsx' ni ningún"
          " otro Excel en la carpeta."
      )
    file_path = archivos_excel[0]

  try:
    df = pd.read_excel(file_path, sheet_name="report")
  except ValueError:
    xls = pd.ExcelFile(file_path)
    sheet_name = (
        "Ventas y Promotoras"
        if "Ventas y Promotoras" in xls.sheet_names
        else xls.sheet_names[0]
    )
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    st.warning(
        f"La hoja 'report' no se encontró. Se ha cargado la hoja '{sheet_name}'."
    )

  df.columns = df.columns.astype(str).str.strip()

  rename_dict = {
      "Plaza / Ciudad": "Regional",
      "Promotora": "Empleado",
      "Categoría Producto": "Supercategoría",
      "Tipo Canal": "Venta Directa o Indirecta",
      "Unidades Vendidas": "Informar la Cantidad Vendida",
      "Precio Unitario ($)": "Informar precio de la venta",
  }
  df = df.rename(columns=rename_dict)

  required_cols = [
      "Informar la Cantidad Vendida",
      "Informar precio de la venta",
      "Fecha y hora de la encuesta",
      "Venta Directa o Indirecta",
  ]
  missing_cols = [col for col in required_cols if col not in df.columns]
  if missing_cols:
    st.error(
        "El archivo Excel no contiene las columnas necesarias:"
        f" {', '.join(missing_cols)}. Columnas detectadas:"
        f" {list(df.columns)}"
    )
    st.stop()

  df["Venta Directa o Indirecta"] = (
      df["Venta Directa o Indirecta"].astype(str).str.strip().str.capitalize()
  )

  def limpiar_precio(val):
    if pd.isna(val):
      return 0.0
    val_str = (
        str(val).strip().replace("$", "").replace("USD", "").replace(" ", "")
    )
    if "," in val_str and "." in val_str:
      if val_str.rfind(",") > val_str.rfind("."):
        val_str = val_str.replace(".", "").replace(",", ".")
      else:
        val_str = val_str.replace(",", "")
    elif "," in val_str:
      val_str = val_str.replace(",", ".")
    try:
      return float(val_str)
    except ValueError:
      return 0.0

  df["Informar precio de la venta"] = df["Informar precio de la venta"].apply(
      limpiar_precio
  )

  def limpiar_unidades(val):
    if pd.isna(val):
      return 0.0
    val_str = str(val).strip().replace(" ", "").replace(",", ".")
    try:
      return float(val_str)
    except ValueError:
      return 0.0

  df["Informar la Cantidad Vendida"] = df[
      "Informar la Cantidad Vendida"
  ].apply(limpiar_unidades)

  if "Venta Total ($)" in df.columns:
    df["Total Venta"] = pd.to_numeric(df["Venta Total ($)"], errors="coerce")
    mask_nan = df["Total Venta"].isna()
    df.loc[mask_nan, "Total Venta"] = df.loc[
        mask_nan, "Informar precio de la venta"
    ]
  else:
    df["Total Venta"] = df["Informar precio de la venta"]

  df["Fecha y hora de la encuesta"] = pd.to_datetime(
      df["Fecha y hora de la encuesta"], errors="coerce"
  )
  df["Fecha"] = df["Fecha y hora de la encuesta"].dt.date
  df["Semana"] = df["Fecha y hora de la encuesta"].dt.to_period("W").astype(str)

  if "Empleado" not in df.columns:
    df["Empleado"] = "General"
  if "Regional" not in df.columns:
    df["Regional"] = "Nacional"
  if "Supercategoría" not in df.columns:
    df["Supercategoría"] = "General"
  else:
    df["Supercategoría"] = (
        df["Supercategoría"].astype(str).str.strip().str.capitalize()
    )
    df["Supercategoría"] = df["Supercategoría"].replace(
        {"Congeladores": "Congelador"}
    )
  if "Informar modelo" not in df.columns:
    df["Informar modelo"] = "General"
  if "PDV" not in df.columns:
    df["PDV"] = df["Regional"]

  return df, file_path


try:
  df, archivo_usado = load_data()
except Exception as e:
  st.error(f"Error al cargar el archivo Excel: {e}.")
  st.stop()

# --- EXTRACCIÓN AUTOMÁTICA DE FECHAS ---
fechas_validas = pd.to_datetime(df["Fecha y hora de la encuesta"]).dropna()
if not fechas_validas.empty:
  min_date = fechas_validas.min().strftime("%d/%m/%Y")
  max_date = fechas_validas.max().strftime("%d/%m/%Y")
  rango_fechas_str = f"(Del {min_date} al {max_date})"
else:
  rango_fechas_str = "(Periodo Actual)"

# --- CARGA DE LOGO ---
logo = None
for logo_file in ["Mystic.jpg", "image_0.png", "image_1.png"]:
  if os.path.exists(logo_file):
    try:
      logo = Image.open(logo_file)
      break
    except Exception:
      pass

# --- CABECERA ---
header_col1, header_col2 = st.columns([1.0, 11])
with header_col1:
  if logo:
    st.image(logo, width=130)
with header_col2:
  st.markdown(
      f"<h1 style='margin-bottom: 0px; margin-top: 5px; font-size: 32px;"
      f" color: #262730; line-height: 1.2;'>Dashboard interactivo de Sell-out"
      f" <span style='font-size: 18px; font-weight: normal; color: #555555;'>{rango_fechas_str}</span></h1>",
      unsafe_allow_html=True,
  )

# --- BARRA LATERAL: FILTROS ---
st.sidebar.header("Filtros Globales")

tipo_venta_options = ["Todos"] + sorted(
    df["Venta Directa o Indirecta"].dropna().unique().tolist()
)
selected_tipo_venta = st.sidebar.selectbox(
    "Seleccione Tipo de Venta", tipo_venta_options
)

st.sidebar.markdown("---")

empleado_options = ["Todos"] + sorted(df["Empleado"].dropna().unique().tolist())
selected_empleado = st.sidebar.selectbox(
    "Seleccione Empleado", empleado_options
)

regional_options = ["Todas"] + sorted(df["Regional"].dropna().unique().tolist())
selected_regional = st.sidebar.selectbox(
    "Seleccione Regional", regional_options
)

supercat_options = ["Todas"] + sorted(
    df["Supercategoría"].dropna().unique().tolist()
)
selected_supercat = st.sidebar.selectbox(
    "Seleccione Supercategoría", supercat_options
)

modelo_options = ["Todos"] + sorted(
    df["Informar modelo"].dropna().unique().tolist()
)
selected_modelo = st.sidebar.selectbox("Seleccione Modelo", modelo_options)

pdv_options = ["Todos"] + sorted(df["PDV"].dropna().unique().tolist())
selected_pdv = st.sidebar.selectbox("Seleccione PDV", pdv_options)

df_filtered_general = df.copy()

if selected_empleado != "Todos":
  df_filtered_general = df_filtered_general[
      df_filtered_general["Empleado"] == selected_empleado
  ]
if selected_regional != "Todas":
  df_filtered_general = df_filtered_general[
      df_filtered_general["Regional"] == selected_regional
  ]
if selected_supercat != "Todas":
  df_filtered_general = df_filtered_general[
      df_filtered_general["Supercategoría"] == selected_supercat
  ]
if selected_modelo != "Todos":
  df_filtered_general = df_filtered_general[
      df_filtered_general["Informar modelo"] == selected_modelo
  ]
if selected_pdv != "Todos":
  df_filtered_general = df_filtered_general[
      df_filtered_general["PDV"] == selected_pdv
  ]

total_revenue_gen = df_filtered_general["Total Venta"].sum()
total_units_gen = df_filtered_general["Informar la Cantidad Vendida"].sum()
total_transactions_gen = len(df_filtered_general)
avg_ticket_gen = (
    total_revenue_gen / total_transactions_gen
    if total_transactions_gen > 0
    else 0
)


def fmt_dolar(val):
  return f"${val:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


# --- TARJETAS DE MÉTRICAS (KPIs) GENERALES ---
col1, col2, col3, col4 = st.columns(4)
with col1:
  st.metric(label="Venta Total (USD)", value=fmt_dolar(total_revenue_gen))
with col2:
  st.metric(label="Unidades Vendidas", value=f"{total_units_gen:,.0f}")
with col3:
  st.metric(label="Transacciones", value=f"{total_transactions_gen:,}")
with col4:
  st.metric(label="Ticket Promedio", value=fmt_dolar(avg_ticket_gen))

# --- DESGLOSE DE VENTA DIRECTA E INDIRECTA ---
if (
    "Venta Directa o Indirecta" in df_filtered_general.columns
    and not df_filtered_general.empty
):
  st.markdown(
      "<h4 style='font-size: 18px; color: #262730; margin-bottom: 10px;'>Desglose"
      " por Tipo de Venta (Directa e Indirecta)</h4>",
      unsafe_allow_html=True,
  )
  df_tipo_resumen = (
      df_filtered_general.groupby("Venta Directa o Indirecta")
      .agg(
          Venta_USD=("Total Venta", "sum"),
          Unidades=("Informar la Cantidad Vendida", "sum"),
          Transacciones=("Total Venta", "count"),
      )
      .reset_index()
      .sort_values(by="Venta_USD", ascending=False)
  )

  df_tipo_resumen["Texto_Etiqueta"] = df_tipo_resumen.apply(
      lambda r: f"{fmt_dolar(r['Venta_USD'])}<br>({r['Unidades']:,.0f} u.)",
      axis=1,
  )

  lista_tipos_orden = df_tipo_resumen["Venta Directa o Indirecta"].tolist()

  fig_tipo = px.bar(
      df_tipo_resumen,
      x="Venta Directa o Indirecta",
      y="Venta_USD",
      text="Texto_Etiqueta",
      color="Venta Directa o Indirecta",
      color_discrete_sequence=MYSTIC_BLUE_PALETTE,
      template="plotly_white",
      category_orders={"Venta Directa o Indirecta": lista_tipos_orden},
  )
  fig_tipo.update_traces(textposition="outside")
  fig_tipo.update_layout(
      showlegend=False,
      xaxis_title="",
      yaxis_title="Venta (USD)",
      margin=dict(t=20, l=20, r=20, b=20),
      height=320,
  )
  st.plotly_chart(fig_tipo, use_container_width=True)

st.markdown("---")

df_filtered = df_filtered_general.copy()
if selected_tipo_venta != "Todos":
  df_filtered = df_filtered[
      df_filtered["Venta Directa o Indirecta"] == selected_tipo_venta
  ]

# --- HIGHLIGHTS ANALÍTICOS ---
if not df_filtered.empty:
  top_cat = (
      df_filtered.groupby("Supercategoría")["Total Venta"].sum().idxmax()
      if "Supercategoría" in df_filtered.columns and not df_filtered.empty
      else "N/A"
  )

  if "Informar modelo" in df_filtered.columns and not df_filtered.empty:
    top_models_series = (
        df_filtered.groupby("Informar modelo")["Informar la Cantidad Vendida"]
        .sum()
        .nlargest(3)
    )
    top_models_list = top_models_series.index.tolist()
    top_models_str = ", ".join(top_models_list) if top_models_list else "N/A"
  else:
    top_models_str = "N/A"

  top_reg = (
      df_filtered.groupby("Regional")["Total Venta"].sum().idxmax()
      if "Regional" in df_filtered.columns and not df_filtered.empty
      else "N/A"
  )
  top_pdv = (
      df_filtered.groupby("PDV")["Total Venta"].sum().idxmax()
      if "PDV" in df_filtered.columns and not df_filtered.empty
      else "N/A"
  )
  top_emp = (
      df_filtered.groupby("Empleado")["Total Venta"].sum().idxmax()
      if "Empleado" in df_filtered.columns and not df_filtered.empty
      else "N/A"
  )

  total_rev_filtered = df_filtered["Total Venta"].sum()
  if "Venta Directa o Indirecta" in df_filtered.columns:
    ventas_tipo = df_filtered.groupby("Venta Directa o Indirecta")[
        "Total Venta"
    ].sum()
    if not ventas_tipo.empty:
      dominant_tipo = ventas_tipo.idxmax()
      dominant_pct = (
          (ventas_tipo.max() / total_rev_filtered * 100)
          if total_rev_filtered > 0
          else 0
      )
    else:
      dominant_tipo, dominant_pct = "N/A", 0
  else:
    dominant_tipo, dominant_pct = "N/A", 0

  if selected_empleado != "Todos":
    st.info(
        f"**Highlights Individuales ({selected_empleado}):** Canal Principal:"
        f" {dominant_tipo} ({dominant_pct:.1f}%) | Categoría Destacada:"
        f" {top_cat} | Modelos Top: {top_models_str} | Regional: {top_reg} |"
        f" PDV Clave: {top_pdv}"
    )
  else:
    st.info(
        f"**Highlights Globales:** Canal Dominante: {dominant_tipo}"
        f" ({dominant_pct:.1f}%) | Categoría Principal: {top_cat} | Modelos"
        f" Top: {top_models_str} | Regional Principal: {top_reg} | PDV Clave:"
        f" {top_pdv} | Promotor Líder: {top_emp}"
    )

# --- GRÁFICOS ANALÍTICOS ---
st.markdown("### Análisis Comercial Consolidado en Dólares")

col_sec1, col_sec2 = st.columns(2)

with col_sec1:
  if not df_filtered.empty:
    df_cat = (
        df_filtered.groupby("Supercategoría")
        .agg(
            Total_Venta=("Total Venta", "sum"),
            Total_Unidades=("Informar la Cantidad Vendida", "sum"),
        )
        .reset_index()
        .sort_values(by="Total_Venta", ascending=False)
    )

    df_cat["Texto_Etiqueta"] = df_cat.apply(
        lambda r: f"{fmt_dolar(r['Total_Venta'])}<br>({r['Total_Unidades']:,.0f} u.)",
        axis=1,
    )

    total_dolares_cat = df_cat["Total_Venta"].sum()
    st.markdown(
        f"<h4 style='font-size: 16px; color: #262730; margin-bottom: 5px;'>Rendimiento"
        f" por Supercategoría (Total: {fmt_dolar(total_dolares_cat)})</h4>",
        unsafe_allow_html=True,
    )

    fig_cat = px.bar(
        df_cat,
        x="Supercategoría",
        y="Total_Venta",
        text="Texto_Etiqueta",
        color="Supercategoría",
        color_discrete_sequence=MYSTIC_BLUE_PALETTE,
        template="plotly_white",
    )
    fig_cat.update_traces(textposition="outside")
    fig_cat.update_layout(
        showlegend=False,
        xaxis_title="",
        yaxis_title="Venta (USD)",
        margin=dict(t=20, l=40, r=40),
    )
    st.plotly_chart(fig_cat, use_container_width=True)
  else:
    st.markdown(
        "<h4 style='font-size: 16px; color: #262730;'>Rendimiento por"
        " Supercategoría</h4>",
        unsafe_allow_html=True,
    )
    st.info("No hay datos para los filtros seleccionados.")

with col_sec2:
  if not df_filtered.empty:
    df_reg = (
        df_filtered.groupby("Regional")
        .agg(
            Total_Venta=("Total Venta", "sum"),
            Total_Unidades=("Informar la Cantidad Vendida", "sum"),
        )
        .reset_index()
    )

    total_dolares_reg = df_reg["Total_Venta"].sum()
    st.markdown(
        f"<h4 style='font-size: 16px; color: #262730; margin-bottom: 5px;'>Distribución"
        f" por Regional (Total: {fmt_dolar(total_dolares_reg)})</h4>",
        unsafe_allow_html=True,
    )

    fig_reg = px.pie(
        df_reg,
        names="Regional",
        values="Total_Venta",
        hole=0.4,
        color_discrete_sequence=MYSTIC_BLUE_PALETTE,
        template="plotly_white",
    )
    fig_reg.update_traces(textposition="inside", textinfo="percent+label")
    fig_reg.update_layout(margin=dict(t=20, l=40, r=40))
    st.plotly_chart(fig_reg, use_container_width=True)
  else:
    st.markdown(
        "<h4 style='font-size: 16px; color: #262730;'>Distribución por"
        " Regional</h4>",
        unsafe_allow_html=True,
    )
    st.info("No hay datos para los filtros seleccionados.")

st.markdown("---")

col_sec3, col_sec4 = st.columns(2)

with col_sec3:
  if not df_filtered.empty:
    df_mod = (
        df_filtered.groupby("Informar modelo")
        .agg(
            Total_Venta=("Total Venta", "sum"),
            Total_Unidades=("Informar la Cantidad Vendida", "sum"),
        )
        .reset_index()
        .sort_values(by="Total_Venta", ascending=False)
    )

    df_mod["Texto_Etiqueta"] = df_mod.apply(
        lambda r: f"{fmt_dolar(r['Total_Venta'])}<br>({r['Total_Unidades']:,.0f} u.)",
        axis=1,
    )

    total_dolares_mod = df_mod["Total_Venta"].sum()
    st.markdown(
        f"<h4 style='font-size: 16px; color: #262730; margin-bottom: 5px;'>Rendimiento"
        f" por Modelo (Total: {fmt_dolar(total_dolares_mod)})</h4>",
        unsafe_allow_html=True,
    )

    fig_mod = px.bar(
        df_mod,
        x="Informar modelo",
        y="Total_Venta",
        text="Texto_Etiqueta",
        color="Informar modelo",
        color_discrete_sequence=MYSTIC_BLUE_PALETTE,
        template="plotly_white",
    )
    fig_mod.update_traces(textposition="outside")
    fig_mod.update_layout(
        showlegend=False,
        xaxis_title="Modelo",
        yaxis_title="Venta (USD)",
        margin=dict(t=20, l=40, r=40),
    )
    st.plotly_chart(fig_mod, use_container_width=True)
  else:
    st.markdown(
        "<h4 style='font-size: 16px; color: #262730;'>Rendimiento por"
        " Modelo</h4>",
        unsafe_allow_html=True,
    )
    st.info("No hay datos para los filtros seleccionados.")

with col_sec4:
  if not df_filtered.empty:
    tipo_tendencia = st.radio(
        "Agrupación temporal:",
        ("Diaria", "Semanal"),
        horizontal=True,
        key="tendencia_radio",
    )
    columna_tiempo = "Fecha" if tipo_tendencia == "Diaria" else "Semana"

    total_dolares_time = df_filtered["Total Venta"].sum()
    st.markdown(
        f"<h4 style='font-size: 16px; color: #262730; margin-bottom: 5px;'>Tendencia"
        f" {tipo_tendencia} en Dólares (Total:"
        f" {fmt_dolar(total_dolares_time)})</h4>",
        unsafe_allow_html=True,
    )

    df_time = (
        df_filtered.groupby(columna_tiempo)["Total Venta"].sum().reset_index()
    )
    fig_time = px.line(
        df_time,
        x=columna_tiempo,
        y="Total Venta",
        markers=True,
        color_discrete_sequence=["#3f51b5"],
        template="plotly_white",
    )
    fig_time.update_layout(
        xaxis_title=tipo_tendencia,
        yaxis_title="Venta (USD)",
        margin=dict(t=20, l=40, r=40),
    )
    st.plotly_chart(fig_time, use_container_width=True)
  else:
    st.markdown(
        "<h4 style='font-size: 16px; color: #262730;'>Tendencia Temporal</h4>",
        unsafe_allow_html=True,
    )
    st.info("No hay datos para los filtros seleccionados.")

st.markdown("---")

col_sec5, col_sec6 = st.columns(2)

with col_sec5:
  if not df_filtered.empty:
    df_pdv = (
        df_filtered.groupby("PDV")
        .agg(
            Total_Venta=("Total Venta", "sum"),
            Total_Unidades=("Informar la Cantidad Vendida", "sum"),
        )
        .reset_index()
        .sort_values(by="Total_Venta", ascending=True)
    )

    df_pdv["Texto_Etiqueta"] = df_pdv.apply(
        lambda r: f"{fmt_dolar(r['Total_Venta'])} ({r['Total_Unidades']:,.0f} u.)",
        axis=1,
    )

    total_dolares_pdv = df_pdv["Total_Venta"].sum()
    st.markdown(
        f"<h4 style='font-size: 16px; color: #262730; margin-bottom: 5px;'>Rendimiento"
        f" por PDV (Total: {fmt_dolar(total_dolares_pdv)})</h4>",
        unsafe_allow_html=True,
    )

    fig_pdv = px.bar(
        df_pdv,
        x="Total_Venta",
        y="PDV",
        orientation="h",
        text="Texto_Etiqueta",
        color="PDV",
        color_discrete_sequence=MYSTIC_BLUE_PALETTE,
        template="plotly_white",
    )
    fig_pdv.update_traces(textposition="outside")
    fig_pdv.update_layout(
        showlegend=False,
        xaxis_title="Venta (USD)",
        yaxis_title="",
        height=500,
        margin=dict(t=20, l=120, r=100),
        yaxis={
            "categoryorder": "array",
            "categoryarray": df_pdv["PDV"].tolist(),
        },
    )
    st.plotly_chart(fig_pdv, use_container_width=True)
  else:
    st.markdown(
        "<h4 style='font-size: 16px; color: #262730;'>Rendimiento por PDV</h4>",
        unsafe_allow_html=True,
    )
    st.info("No hay datos para los filtros seleccionados.")

with col_sec6:
  if not df_filtered.empty:
    df_emp = (
        df_filtered.groupby("Empleado")
        .agg(
            Total_Venta=("Total Venta", "sum"),
            Total_Unidades=("Informar la Cantidad Vendida", "sum"),
        )
        .reset_index()
        .sort_values(by="Total_Venta", ascending=True)
    )

    df_emp["Texto_Etiqueta"] = df_emp.apply(
        lambda r: f"{fmt_dolar(r['Total_Venta'])} ({r['Total_Unidades']:,.0f} u.)",
        axis=1,
    )

    total_dolares_emp = df_emp["Total_Venta"].sum()
    st.markdown(
        f"<h4 style='font-size: 16px; color: #262730; margin-bottom: 5px;'>Rendimiento"
        f" por Empleado (Total: {fmt_dolar(total_dolares_emp)})</h4>",
        unsafe_allow_html=True,
    )

    fig_emp = px.bar(
        df_emp,
        x="Total_Venta",
        y="Empleado",
        orientation="h",
        text="Texto_Etiqueta",
        color="Empleado",
        color_discrete_sequence=MYSTIC_BLUE_PALETTE,
        template="plotly_white",
    )
    fig_emp.update_traces(textposition="outside")
    fig_emp.update_layout(
        showlegend=False,
        xaxis_title="Venta (USD)",
        yaxis_title="",
        height=500,
        margin=dict(t=20, l=150, r=100),
        yaxis={
            "categoryorder": "array",
            "categoryarray": df_emp["Empleado"].tolist(),
        },
    )
    st.plotly_chart(fig_emp, use_container_width=True)
  else:
    st.markdown(
        "<h4 style='font-size: 16px; color: #262730;'>Rendimiento por"
        " Empleado</h4>",
        unsafe_allow_html=True,
    )
    st.info("No hay datos para los filtros seleccionados.")

st.markdown("---")

# --- BOTÓN DE DESCARGA (EN LÍNEA PARA EVITAR ERRORES) ---
st.subheader("Exportar Base de Datos")


@st.cache_data
def convert_df_to_csv(df_to_conv):
  return df_to_conv.to_csv(index=False).encode("utf-8")


csv_data = convert_df_to_csv(df_filtered)
st.download_button(
    label="Descargar datos filtrados (CSV)",
    data=csv_data,
    file_name="base_datos_filtrada_mystic.csv",
    mime="text/csv",
)

st.markdown("---")

# --- MÓDULO DE NÓMINA PROTEGIDO ---
st.subheader(
    "Modulo de Nomina: Liquidacion de Bonos por Escalas (Ventas Directas y"
    " Precio >= $120)"
)

if "autenticado_nomina" not in st.session_state:
  st.session_state["autenticado_nomina"] = False

CLAVE_CORRECTA = "jorgito2026"

if not st.session_state["autenticado_nomina"]:
  password_ingresada = st.text_input(
      "Ingrese password de acceso",
      type="password",
      key="input_password_field",
  )
  if password_ingresada == CLAVE_CORRECTA:
    st.session_state["autenticado_nomina"] = True
    st.rerun()
  elif password_ingresada != "":
    st.error("Contraseña incorrecta. Por favor, verifique e intente de nuevo.")
else:
  st.success("¡Clave correcta! Módulo de nómina desbloqueado.")
  if st.button("Cerrar sesión de nómina"):
    st.session_state["autenticado_nomina"] = False
    st.rerun()

  # Parámetro base del monto máximo de bono ajustado
  BONO_BASE_PROMOTOR = 100.00
  BONO_BASE_OCCIDENTE = 100.00
  BONO_BASE_NACIONAL = 300.00

  # TABLAS VISUALES DE ESCALAS (DEFINIDAS EN LÍNEA CONTINUA)
  col_t1, col_t2, col_t3 = st.columns(3)

  with col_t1:
    st.markdown("**Tabla Tramos: Promotores**")
    df_tabla_prom = pd.DataFrame(
        {
            "Rango Unidades": [
                "< 30 u.",
                "30 a 49 u.",
                "50 a 69 u.",
                "70 a 89 u.",
                "90 a 109 u.",
                "≥ 110 u.",
            ],
            "% Bono": ["0%", "10%", "25%", "50%", "75%", "100%"],
            "Bono USD": [
                "$0.00",
                "$10.00",
                "$25.00",
                "$50.00",
                "$75.00",
                "$100.00",
            ],
        }
    )
    st.dataframe(df_tabla_prom, hide_index=True, use_container_width=True)

  with col_t2:
    st.markdown("**Tabla Tramos: Occidente**")
    df_tabla_occ = pd.DataFrame(
        {
            "Rango Unidades": [
                "< 90 u.",
                "90 a 119 u.",
                "120 a 149 u.",
                "150 a 189 u.",
                "190 a 229 u.",
                "≥ 230 u.",
            ],
            "% Bono": ["0%", "10%", "25%", "50%", "75%", "100%"],
            "Bono USD": [
                "$0.00",
                "$10.00",
                "$25.00",
                "$50.00",
                "$75.00",
                "$100.00",
            ],
        }
    )
    st.dataframe(df_tabla_occ, hide_index=True, use_container_width=True)

  with col_t3:
    st.markdown("**Tabla Tramos: Jefe Nacional**")
    df_tabla_nac = pd.DataFrame(
        {
            "Rango Unidades": [
                "< 380 u.",
                "380 a 489 u.",
                "490 a 549 u.",
                "550 a 579 u.",
                "580 a 599 u.",
                "≥ 600 u.",
            ],
            "% Bono": ["0%", "10%", "25%", "50%", "75%", "100%"],
            "Bono USD": [
                "$0.00",
                "$30.00",
                "$75.00",
                "$150.00",
                "$225.00",
                "$300.00",
            ],
        }
    )
    st.dataframe(df_tabla_nac, hide_index=True, use_container_width=True)

  st.markdown("---")

  df_nomina = df[
      df["Venta Directa o Indirecta"]
      .str.lower()
      .str.contains("directa")
      & (df["Informar precio de la venta"] >= 120)
  ].copy()

  if not df_nomina.empty:
    df_prom_resumen = (
        df_nomina.groupby("Empleado")
        .agg(
            Regional=("Regional", "first"),
            Unidades_Validas=("Informar la Cantidad Vendida", "sum"),
            Venta_Total_USD=("Total Venta", "sum"),
        )
        .reset_index()
    )

    resultados_nomina = []
    for _, row in df_prom_resumen.iterrows():
      nombre = row["Empleado"]
      regional = row["Regional"]
      unidades = row["Unidades_Validas"]
      v_usd = row["Venta_Total_USD"]

      if unidades >= 110:
        escala_txt, bono_pct = "≥ 110 unidades (100%)", 1.00
      elif unidades >= 90:
        escala_txt, bono_pct = "90 a 109 unidades (75%)", 0.75
      elif unidades >= 70:
        escala_txt, bono_pct = "70 a 89 unidades (50%)", 0.50
      elif unidades >= 50:
        escala_txt, bono_pct = "50 a 69 unidades (25%)", 0.25
      elif unidades >= 30:
        escala_txt, bono_pct = "30 a 49 unidades (10%)", 0.10
      else:
        escala_txt, bono_pct = "Sin Bono (< 30 u.)", 0.00

      monto_bono_usd = BONO_BASE_PROMOTOR * bono_pct

      resultados_nomina.append({
          "Empleado": nombre,
          "Región / Estado": regional,
          "Unidades_Validas": unidades,
          "Venta_Total_USD": round(v_usd, 2),
          "Escala_Alcanzada": escala_txt,
          "Porcentaje_Bono": f"{int(bono_pct * 100)}%",
          "Bono_USD": fmt_dolar(monto_bono_usd),
      })

    df_final_nomina = pd.DataFrame(resultados_nomina).sort_values(
        by="Unidades_Validas", ascending=False
    )

    st.markdown("### Tabla de Comisiones y Bonos por Promotor")
    st.dataframe(df_final_nomina, use_container_width=True)

    total_unidades_equipo = df_final_nomina["Unidades_Validas"].sum()

    st.markdown("---")

    # --- SECCIÓN 1: SUPERVISORA DE OCCIDENTE ---
    st.subheader("Liquidación: Supervisora de Occidente (Karen Chang)")
    st.markdown(
        "*(Táchira, Portuguesa, Barinas, Lara, Aragua - Basado en la suma de"
        " unidades válidas del equipo de Occidente)*"
    )

    regiones_occidente = [
        "táchira",
        "tachira",
        "portuguesa",
        "barinas",
        "lara",
        "aragua",
    ]
    df_occidente = df_nomina[
        df_nomina["Regional"].str.lower().isin(regiones_occidente)
    ]
    unidades_occidente = df_occidente["Informar la Cantidad Vendida"].sum()

    if unidades_occidente >= 230:
      escala_occ_txt, occ_pct = "≥ 230 unidades (100%)", 1.00
    elif unidades_occidente >= 190:
      escala_occ_txt, occ_pct = "190 a 229 unidades (75%)", 0.75
    elif unidades_occidente >= 150:
      escala_occ_txt, occ_pct = "150 a 189 unidades (50%)", 0.50
    elif unidades_occidente >= 120:
      escala_occ_txt, occ_pct = "120 a 149 unidades (25%)", 0.25
    elif unidades_occidente >= 90:
      escala_occ_txt, occ_pct = "90 a 119 unidades (10%)", 0.10
    else:
      escala_occ_txt, occ_pct = "Menor a 90 unidades (Sin Bono)", 0.00

    bono_occ_usd = BONO_BASE_OCCIDENTE * occ_pct

    col_o1, col_o2, col_o3, col_o4 = st.columns(4)
    with col_o1:
      st.metric(
          "Unidades Válidas Occidente", f"{unidades_occidente:,.0f} u."
      )
    with col_o2:
      st.metric("Escala Alcanzada", escala_occ_txt)
    with col_o3:
      st.metric("Porcentaje Aplicado", f"{int(occ_pct * 100)}%")
    with col_o4:
      st.metric("Bono en USD ($)", fmt_dolar(bono_occ_usd))

    st.markdown("---")

    # --- SECCIÓN 2: JEFE NACIONAL DE TM ---
    st.subheader("Liquidación: Jefe Nacional TM (Jorge Abraham)")
    st.markdown(
        "*(Dirección Nacional - Basado en la suma total de unidades válidas a"
        " nivel nacional)*"
    )

    unidades_nacionales = total_unidades_equipo

    if unidades_nacionales >= 600:
      escala_nac_txt, nac_pct = "≥ 600 unidades (100%)", 1.00
    elif unidades_nacionales >= 580:
      escala_nac_txt, nac_pct = "580 a 599 unidades (75%)", 0.75
    elif unidades_nacionales >= 550:
      escala_nac_txt, nac_pct = "550 a 579 unidades (50%)", 0.50
    elif unidades_nacionales >= 490:
      escala_nac_txt, nac_pct = "490 a 549 unidades (25%)", 0.25
    elif unidades_nacionales >= 380:
      escala_nac_txt, nac_pct = "380 a 489 unidades (10%)", 0.10
    else:
      escala_nac_txt, nac_pct = "Menor a 380 unidades (Sin Bono)", 0.00

    bono_nac_usd = BONO_BASE_NACIONAL * nac_pct

    col_n1, col_n2, col_n3, col_n4 = st.columns(4)
    with col_n1:
      st.metric("Total Unidades Nacionales", f"{unidades_nacionales:,.0f} u.")
    with col_n2:
      st.metric("Escala Alcanzada", escala_nac_txt)
    with col_n3:
      st.metric("Porcentaje Aplicado", f"{int(nac_pct * 100)}%")
    with col_n4:
      st.metric("Bono en USD ($)", fmt_dolar(bono_nac_usd))

  else:
    st.warning(
        "No se encontraron registros que cumplan con las condiciones de Venta"
        " Directa y Precio >= $120 en el conjunto de datos actual."
    )