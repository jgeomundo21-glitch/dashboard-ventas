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

@st.cache_data
def load_data():
    archivos_excel = glob.glob("*.xlsx")
    if not archivos_excel:
        raise FileNotFoundError(
            "No se encontró ningún archivo Excel en la carpeta. Coloca tu archivo aquí."
        )
    file_path = archivos_excel[0]
    try:
        df = pd.read_excel(file_path, sheet_name="report")
    except ValueError:
        df = pd.read_excel(file_path)
        st.warning(f"La hoja 'report' no se encontró en '{file_path}'. Se ha cargado la primera hoja disponible.")

    df.columns = df.columns.str.strip()

    required_cols = ["Informar la Cantidad Vendida", "Informar precio de la venta", "Fecha y hora de la encuesta", "Venta Directa o Indirecta"]
    if not all(col in df.columns for col in required_cols):
        st.error(f"El archivo Excel no contiene las columnas necesarias: {', '.join(required_cols)}")
        st.stop()

    # --- LIMPIEZA Y NORMALIZACIÓN PROFUNDA ---
    df["Venta Directa o Indirecta"] = df["Venta Directa o Indirecta"].astype(str).str.strip().str.capitalize()

    df["Informar precio de la venta"] = (
        df["Informar precio de la venta"]
        .astype(str)
        .str.replace("$", "", regex=False)
        .str.replace(" ", "", regex=False)
        .str.replace(",", ".", regex=False)
    )
    df["Informar precio de la venta"] = pd.to_numeric(df["Informar precio de la venta"], errors="coerce").fillna(0)

    df["Informar la Cantidad Vendida"] = (
        df["Informar la Cantidad Vendida"]
        .astype(str)
        .str.replace(" ", "", regex=False)
        .str.replace(",", ".", regex=False)
    )
    df["Informar la Cantidad Vendida"] = pd.to_numeric(df["Informar la Cantidad Vendida"], errors="coerce").fillna(0)
    # ------------------------------------------

    df["Total Venta"] = (
        df["Informar la Cantidad Vendida"] * df["Informar precio de la venta"]
    )
    df["Fecha y hora de la encuesta"] = pd.to_datetime(
        df["Fecha y hora de la encuesta"]
    )
    df["Fecha"] = df["Fecha y hora de la encuesta"].dt.date
    df["Semana"] = (
        df["Fecha y hora de la encuesta"]
        .dt.to_period("W")
        .astype(str)
    )
    return df, file_path

# --- CARGA DE LOGO ---
try:
    logo = Image.open("Mystic.jpg")
except FileNotFoundError:
    try:
        logo = Image.open("image_0.png")
    except FileNotFoundError:
        logo = None

# --- TÍTULO Y LOGO EN LA CABECERA ---
header_col1, header_col2 = st.columns([1.2, 10])
with header_col1:
    if logo:
        st.image(logo, width=150)
with header_col2:
    st.markdown(
        """
        <h1 style='margin-bottom: 0px; margin-top: 0px; font-size: 38px; color: #262730; line-height: 1.2;'>
            Dashboard interactivo de Sell-out <span style='font-size: 24px; font-weight: normal; color: #555555;'>(Del 1 al 31 de Julio de 2026)</span>
        </h1>
        """,
        unsafe_allow_html=True,
    )

try:
    df, archivo_usado = load_data()
except Exception as e:
    st.error(f"Error al cargar el archivo Excel: {e}.")
    st.stop()

# --- BARRA LATERAL: FILTROS ---
st.sidebar.header("🔍 Filtros Globales")

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
selected_regional = st.sidebar.selectbox("Seleccione Regional", regional_options)

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

# DataFrame general (no filtrado por Tipo de Venta para métricas globales)
df_filtered_general = df.copy()

if selected_empleado != "Todos":
    df_filtered_general = df_filtered_general[df_filtered_general["Empleado"] == selected_empleado]
if selected_regional != "Todas":
    df_filtered_general = df_filtered_general[df_filtered_general["Regional"] == selected_regional]
if selected_supercat != "Todas":
    df_filtered_general = df_filtered_general[df_filtered_general["Supercategoría"] == selected_supercat]
if selected_modelo != "Todos":
    df_filtered_general = df_filtered_general[df_filtered_general["Informar modelo"] == selected_modelo]
if selected_pdv != "Todos":
    df_filtered_general = df_filtered_general[df_filtered_general["PDV"] == selected_pdv]

total_revenue_gen = df_filtered_general["Total Venta"].sum()
total_units_gen = df_filtered_general["Informar la Cantidad Vendida"].sum()
total_transactions_gen = len(df_filtered_general)
avg_ticket_gen = total_revenue_gen / total_transactions_gen if total_transactions_gen > 0 else 0

# --- TARJETAS DE MÉTRICAS (KPIs) GENERALES ---
col1, col2, col3, col4 = st.columns(4)
with col1:
    st.metric(
        label="Venta Total (USD)",
        value=f"${total_revenue_gen:,.2f}"
        .replace(",", "_")
        .replace(".", ",")
        .replace("_", "."),
    )
with col2:
    st.metric(label="Unidades Vendidas", value=f"{total_units_gen:,.0f}")
with col3:
    st.metric(label="Transacciones", value=f"{total_transactions_gen:,}")
with col4:
    st.metric(
        label="Ticket Promedio",
        value=f"${avg_ticket_gen:,.2f}"
        .replace(",", "_")
        .replace(".", ",")
        .replace("_", "."),
    )

# --- DESGLOSE DE VENTA DIRECTA E INDIRECTA ---
if "Venta Directa o Indirecta" in df_filtered_general.columns and not df_filtered_general.empty:
    st.markdown("##### 📊 Desglose por Tipo de Venta (Directa e Indirecta)")
    df_tipo_resumen = df_filtered_general.groupby("Venta Directa o Indirecta").agg(
        Venta_USD=("Total Venta", "sum"),
        Unidades=("Informar la Cantidad Vendida", "sum"),
        Transacciones=("Total Venta", "count")
    ).reset_index()

    t_col1, t_col2 = st.columns(2)
    with t_col1:
        fig_tipo_usd = px.bar(
            df_tipo_resumen,
            x="Venta Directa o Indirecta",
            y="Venta_USD",
            text="Venta_USD",
            color="Venta Directa o Indirecta",
            color_discrete_sequence=px.colors.sequential.Blues,
            template="plotly_white",
        )
        fig_tipo_usd.update_traces(texttemplate="$%{y:,.2f}", textposition="outside")
        fig_tipo_usd.update_layout(
            showlegend=False,
            xaxis_title="",
            yaxis_title="Venta (USD)",
            margin=dict(t=20, l=20, r=20, b=20),
            height=250
        )
        st.plotly_chart(fig_tipo_usd, use_container_width=True)
    with t_col2:
        fig_tipo_un = px.bar(
            df_tipo_resumen,
            x="Venta Directa o Indirecta",
            y="Unidades",
            text="Unidades",
            color="Venta Directa o Indirecta",
            color_discrete_sequence=px.colors.sequential.Blues,
            template="plotly_white",
        )
        fig_tipo_un.update_traces(texttemplate="%{y:,.0f}", textposition="outside")
        fig_tipo_un.update_layout(
            showlegend=False,
            xaxis_title="",
            yaxis_title="Unidades",
            margin=dict(t=20, l=20, r=20, b=20),
            height=250
        )
        st.plotly_chart(fig_tipo_un, use_container_width=True)

st.markdown("---")

# Aplicar filtro de Tipo de Venta para df_filtered
df_filtered = df_filtered_general.copy()
if selected_tipo_venta != "Todos":
    df_filtered = df_filtered[
        df_filtered["Venta Directa o Indirecta"] == selected_tipo_venta
    ]

# --- RECUADRO DE HIGHLIGHTS ANALÍTICOS ---
if not df_filtered.empty:
    top_cat = df_filtered.groupby("Supercategoría")["Total Venta"].sum().idxmax() if "Supercategoría" in df_filtered.columns and not df_filtered.empty else "N/A"
    
    if "Informar modelo" in df_filtered.columns and not df_filtered.empty:
        top_models_series = df_filtered.groupby("Informar modelo")["Informar la Cantidad Vendida"].sum().nlargest(3)
        top_models_list = top_models_series.index.tolist()
        top_models_str = ", ".join(top_models_list) if top_models_list else "N/A"
    else:
        top_models_str = "N/A"

    top_reg = df_filtered.groupby("Regional")["Total Venta"].sum().idxmax() if "Regional" in df_filtered.columns and not df_filtered.empty else "N/A"
    top_pdv = df_filtered.groupby("PDV")["Total Venta"].sum().idxmax() if "PDV" in df_filtered.columns and not df_filtered.empty else "N/A"
    top_emp = df_filtered.groupby("Empleado")["Total Venta"].sum().idxmax() if "Empleado" in df_filtered.columns and not df_filtered.empty else "N/A"
    
    total_rev_filtered = df_filtered["Total Venta"].sum()
    if "Venta Directa o Indirecta" in df_filtered.columns:
        ventas_tipo = df_filtered.groupby("Venta Directa o Indirecta")["Total Venta"].sum()
        if not ventas_tipo.empty:
            dominant_tipo = ventas_tipo.idxmax()
            dominant_pct = (ventas_tipo.max() / total_rev_filtered * 100) if total_rev_filtered > 0 else 0
        else:
            dominant_tipo, dominant_pct = "N/A", 0
    else:
        dominant_tipo, dominant_pct = "N/A", 0

    if selected_empleado != "Todos":
        highlight_html = f"""
        <div style="background-color: #f0f2f6; padding: 20px; border-radius: 10px; border-left: 5px solid #1f77b4; margin-bottom: 25px;">
            <h4 style="margin-top: 0px; color: #262730; font-size: 18px;">💡 Highlights y Rendimiento Individual: {selected_empleado}</h4>
            <ul style="color: #31333F; font-size: 14px; margin-bottom: 0px; padding-left: 20px; line-height: 1.6;">
                <li><b>Canal de Venta Principal:</b> Concentra su mayor facturación en la modalidad de <b>{dominant_tipo}</b> (<b>{dominant_pct:.1f}%</b> de sus ventas).</li>
                <li><b>Supercategoría Destacada:</b> La categoría con mayor desempeño comercial para este promotor es <b>{top_cat}</b>.</li>
                <li><b>Modelos Top Ventas:</b> Sus tres modelos con mayor salida en unidades son <b>{top_models_str}</b>.</li>
                <li><b>Regional de Operación:</b> Su actividad comercial se consolida principalmente en la regional <b>{top_reg}</b>.</li>
                <li><b>Punto de Venta Clave:</b> El establecimiento con mejor facturación en su gestión es <b>{top_pdv}</b>.</li>
            </ul>
        </div>
        """
        st.markdown(highlight_html, unsafe_allow_html=True)
    else:
        highlight_html = f"""
        <div style="background-color: #f0f2f6; padding: 20px; border-radius: 10px; border-left: 5px solid #1f77b4; margin-bottom: 25px;">
            <h4 style="margin-top: 0px; color: #262730; font-size: 18px;">💡 Highlights y Drivers Clave de Rendimiento</h4>
            <ul style="color: #31333F; font-size: 14px; margin-bottom: 0px; padding-left: 20px; line-height: 1.6;">
                <li><b>Canal Dominante:</b> La modalidad de <b>{dominant_tipo}</b> lidera la facturación global representando el <b>{dominant_pct:.1f}%</b> de las ventas totales del periodo.</li>
                <li><b>Supercategoría Principal:</b> El principal motor de ingresos por categoría es <b>{top_cat}</b>, concentrando el mayor volumen de inversión y salida de producto en piso de venta.</li>
                <li><b>Modelos Top Ventas:</b> Los tres modelos con mayor salida en unidades son <b>{top_models_str}</b>.</li>
                <li><b>Distribución Geográfica:</b> La regional <b>{top_reg}</b> destaca como el núcleo de mayor tracción comercial y dinamismo en el territorio.</li>
                <li><b>Punto de Venta Clave:</b> El establecimiento <b>{top_pdv}</b> sobresale como el PDV con mejor desempeño en facturación general.</li>
                <li><b>Desempeño del Equipo:</b> El promotor/empleado con mayor aporte a la cuota general en el periodo es <b>{top_emp}</b>.</li>
            </ul>
        </div>
        """
        st.markdown(highlight_html, unsafe_allow_html=True)

st.markdown(
    """
    <style>
        div.stSubheader > p {
            font-size: 17px !important;
            font-weight: 600 !important;
            margin-bottom: -5px !important;
        }
    </style>
    """,
    unsafe_allow_html=True,
)

# --- GRÁFICOS INTERACTIVOS ---

# 1. SUPERCATEGORÍA
row1_col1, row1_col2 = st.columns(2)
with row1_col1:
    st.subheader("💡 Rendimiento por Supercategoría en Dólares (US$)")
    if not df_filtered.empty:
        df_cat_usd = (
            df_filtered.groupby("Supercategoría")["Total Venta"]
            .sum()
            .reset_index()
            .sort_values(by="Total Venta", ascending=False)
        )
        fig_cat_usd = px.bar(
            df_cat_usd,
            x="Supercategoría",
            y="Total Venta",
            text="Total Venta",
            color="Total Venta",
            color_continuous_scale="Blues",
            template="plotly_white",
        )
        fig_cat_usd.update_traces(texttemplate="$%{y:,.2f}", textposition="outside")
        fig_cat_usd.update_layout(
            coloraxis_showscale=False,
            showlegend=False,
            xaxis_title="",
            yaxis_title="Venta (USD)",
            xaxis={"categoryorder": "total descending"},
            margin=dict(t=40, l=40, r=40),
        )
        st.plotly_chart(fig_cat_usd, use_container_width=True)
    else:
        st.info("No hay datos para los filtros seleccionados.")

with row1_col2:
    if not df_filtered.empty:
        total_unidades_cat = df_filtered["Informar la Cantidad Vendida"].sum()
        st.subheader(f"💡 Rendimiento por Supercategoría en Unidades (Total: {total_unidades_cat:,.0f} u.)")
        df_cat_un = (
            df_filtered.groupby("Supercategoría")["Informar la Cantidad Vendida"]
            .sum()
            .reset_index()
            .sort_values(by="Informar la Cantidad Vendida", ascending=False)
        )
        fig_cat_un = px.bar(
            df_cat_un,
            x="Supercategoría",
            y="Informar la Cantidad Vendida",
            text="Informar la Cantidad Vendida",
            color="Informar la Cantidad Vendida",
            color_continuous_scale="Blues",
            template="plotly_white",
        )
        fig_cat_un.update_traces(texttemplate="%{y:,.0f}", textposition="outside")
        fig_cat_un.update_layout(
            coloraxis_showscale=False,
            showlegend=False,
            xaxis_title="",
            yaxis_title="Cantidad (Unidades)",
            xaxis={"categoryorder": "total descending"},
            margin=dict(t=40, l=40, r=40),
        )
        st.plotly_chart(fig_cat_un, use_container_width=True)
    else:
        st.subheader("💡 Rendimiento por Supercategoría en Unidades")
        st.info("No hay datos para los filtros seleccionados.")

st.markdown("---")

# 2. REGIONAL
row2_col1, row2_col2 = st.columns(2)
with row2_col1:
    st.subheader("🗺️ Distribución por Regional en Dólares (US$)")
    if not df_filtered.empty:
        fig_reg_usd = px.pie(
            df_filtered.groupby("Regional")["Total Venta"].sum().reset_index(),
            names="Regional",
            values="Total Venta",
            hole=0.4,
            color_discrete_sequence=px.colors.sequential.Blues_r,
            template="plotly_white",
        )
        fig_reg_usd.update_traces(textposition="inside", textinfo="percent+label")
        fig_reg_usd.update_layout(margin=dict(t=40, l=40, r=40))
        st.plotly_chart(fig_reg_usd, use_container_width=True)
    else:
        st.info("No hay datos para los filtros seleccionados.")

with row2_col2:
    if not df_filtered.empty:
        total_unidades_reg = df_filtered["Informar la Cantidad Vendida"].sum()
        st.subheader(f"🗺️ Distribución por Regional en Unidades (Total: {total_unidades_reg:,.0f} u.)")
        fig_reg_un = px.pie(
            df_filtered.groupby("Regional")["Informar la Cantidad Vendida"].sum().reset_index(),
            names="Regional",
            values="Informar la Cantidad Vendida",
            hole=0.4,
            color_discrete_sequence=px.colors.sequential.Blues_r,
            template="plotly_white",
        )
        fig_reg_un.update_traces(textposition="inside", textinfo="percent+label")
        fig_reg_un.update_layout(margin=dict(t=40, l=40, r=40))
        st.plotly_chart(fig_reg_un, use_container_width=True)
    else:
        st.subheader("🗺️ Distribución por Regional en Unidades")
        st.info("No hay datos para los filtros seleccionados.")

st.markdown("---")

# 3. MODELO
row3_col1, row3_col2 = st.columns(2)
with row3_col1:
    st.subheader("🏷️ Rendimiento por Modelo en Dólares (US$)")
    if not df_filtered.empty:
        df_mod_usd = (
            df_filtered.groupby("Informar modelo")["Total Venta"]
            .sum()
            .reset_index()
            .sort_values(by="Total Venta", ascending=False)
        )
        fig_mod_usd = px.bar(
            df_mod_usd,
            x="Informar modelo",
            y="Total Venta",
            text="Total Venta",
            color="Total Venta",
            color_continuous_scale="Blues",
            template="plotly_white",
        )
        fig_mod_usd.update_traces(texttemplate="$%{y:,.2f}", textposition="outside")
        fig_mod_usd.update_layout(
            coloraxis_showscale=False,
            showlegend=False,
            xaxis_title="Modelo",
            yaxis_title="Venta (USD)",
            xaxis={"categoryorder": "total descending"},
            margin=dict(t=40, l=40, r=40),
        )
        st.plotly_chart(fig_mod_usd, use_container_width=True)
    else:
        st.info("No hay datos para los filtros seleccionados.")

with row3_col2:
    if not df_filtered.empty:
        total_unidades_mod = df_filtered["Informar la Cantidad Vendida"].sum()
        st.subheader(f"🏷️ Rendimiento por Modelo en Unidades (Total: {total_unidades_mod:,.0f} u.)")
        df_mod_un = (
            df_filtered.groupby("Informar modelo")["Informar la Cantidad Vendida"]
            .sum()
            .reset_index()
            .sort_values(by="Informar la Cantidad Vendida", ascending=False)
        )
        fig_mod_un = px.bar(
            df_mod_un,
            x="Informar modelo",
            y="Informar la Cantidad Vendida",
            text="Informar la Cantidad Vendida",
            color="Informar la Cantidad Vendida",
            color_continuous_scale="Blues",
            template="plotly_white",
        )
        fig_mod_un.update_traces(texttemplate="%{y:,.0f}", textposition="outside")
        fig_mod_un.update_layout(
            coloraxis_showscale=False,
            showlegend=False,
            xaxis_title="Modelo",
            yaxis_title="Cantidad (Unidades)",
            xaxis={"categoryorder": "total descending"},
            margin=dict(t=40, l=40, r=40),
        )
        st.plotly_chart(fig_mod_un, use_container_width=True)
    else:
        st.subheader("🏷️ Rendimiento por Modelo en Unidades")
        st.info("No hay datos para los filtros seleccionados.")

st.markdown("---")

# 4. TENDENCIA TEMPORAL
tipo_tendencia = st.radio(
    "Agrupación temporal:", ("Diaria", "Semanal"), horizontal=True, key="tendencia_radio"
)

row4_col1, row4_col2 = st.columns(2)
with row4_col1:
    st.subheader(f"📅 Tendencia {tipo_tendencia} en Dólares (US$)")
    if not df_filtered.empty:
        columna_tiempo = "Fecha" if tipo_tendencia == "Diaria" else "Semana"
        df_time_usd = (
            df_filtered.groupby(columna_tiempo)["Total Venta"].sum().reset_index()
        )
        fig_time_usd = px.line(
            df_time_usd,
            x=columna_tiempo,
            y="Total Venta",
            markers=True,
            color_discrete_sequence=["#1f77b4"],
            template="plotly_white",
        )
        fig_time_usd.update_layout(xaxis_title=tipo_tendencia, yaxis_title="Venta (USD)", margin=dict(t=40, l=40, r=40))
        st.plotly_chart(fig_time_usd, use_container_width=True)
    else:
        st.info("No hay datos para los filtros seleccionados.")

with row4_col2:
    if not df_filtered.empty:
        total_unidades_time = df_filtered["Informar la Cantidad Vendida"].sum()
        st.subheader(f"📅 Tendencia {tipo_tendencia} en Unidades (Total: {total_unidades_time:,.0f} u.)")
        columna_tiempo = "Fecha" if tipo_tendencia == "Diaria" else "Semana"
        df_time_un = (
            df_filtered.groupby(columna_tiempo)["Informar la Cantidad Vendida"].sum().reset_index()
        )
        fig_time_un = px.line(
            df_time_un,
            x=columna_tiempo,
            y="Informar la Cantidad Vendida",
            markers=True,
            color_discrete_sequence=["#1f77b4"],
            template="plotly_white",
        )
        fig_time_un.update_layout(xaxis_title=tipo_tendencia, yaxis_title="Cantidad (Unidades)", margin=dict(t=40, l=40, r=40))
        st.plotly_chart(fig_time_un, use_container_width=True)
    else:
        st.subheader(f"📅 Tendencia {tipo_tendencia} en Unidades")
        st.info("No hay datos para los filtros seleccionados.")

st.markdown("---")

# 5. PUNTO DE VENTA (PDV)
row5_col1, row5_col2 = st.columns(2)
with row5_col1:
    if not df_filtered.empty:
        total_dolares_pdv = df_filtered["Total Venta"].sum()
        st.subheader(f"🏪 Rendimiento por PDV en Dólares (US$) (Total: ${total_dolares_pdv:,.2f})".replace(",", "_").replace(".", ",").replace("_", "."))
        df_pdv_usd = (
            df_filtered.groupby("PDV")["Total Venta"]
            .sum()
            .reset_index()
            .sort_values(by="Total Venta", ascending=True)
        )
        fig_pdv_usd = px.bar(
            df_pdv_usd,
            x="Total Venta",
            y="PDV",
            orientation="h",
            text="Total Venta",
            color="Total Venta",
            color_continuous_scale="Blues",
            template="plotly_white",
        )
        fig_pdv_usd.update_traces(texttemplate="$%{x:,.2f}", textposition="outside")
        fig_pdv_usd.update_layout(
            coloraxis_showscale=False,
            xaxis_title="Venta (USD)",
            yaxis_title="",
            height=500,
            margin=dict(t=40, l=120, r=80),
        )
        st.plotly_chart(fig_pdv_usd, use_container_width=True)
    else:
        st.subheader("🏪 Rendimiento por PDV en Dólares (US$)")
        st.info("No hay datos para los filtros seleccionados.")

with row5_col2:
    if not df_filtered.empty:
        total_unidades_pdv = df_filtered["Informar la Cantidad Vendida"].sum()
        st.subheader(f"🏪 Rendimiento por PDV en Unidades (Total: {total_unidades_pdv:,.0f} u.)")
        df_pdv_un = (
            df_filtered.groupby("PDV")["Informar la Cantidad Vendida"]
            .sum()
            .reset_index()
            .sort_values(by="Informar la Cantidad Vendida", ascending=True)
        )
        fig_pdv_un = px.bar(
            df_pdv_un,
            x="Informar la Cantidad Vendida",
            y="PDV",
            orientation="h",
            text="Informar la Cantidad Vendida",
            color="Informar la Cantidad Vendida",
            color_continuous_scale="Blues",
            template="plotly_white",
        )
        fig_pdv_un.update_traces(texttemplate="%{x:,.0f}", textposition="outside")
        fig_pdv_un.update_layout(
            coloraxis_showscale=False,
            xaxis_title="Cantidad (Unidades)",
            yaxis_title="",
            height=500,
            margin=dict(t=40, l=120, r=80),
        )
        st.plotly_chart(fig_pdv_un, use_container_width=True)
    else:
        st.subheader("🏪 Rendimiento por PDV en Unidades")
        st.info("No hay datos para los filtros seleccionados.")

st.markdown("---")

# 6. TOP EMPLEADOS - GENERAL
row6_col1, row6_col2 = st.columns(2)
with row6_col1:
    st.subheader("🏆 Top Empleados - General en Dólares (US$)")
    if not df_filtered.empty:
        df_emp_usd = (
            df_filtered.groupby("Empleado")["Total Venta"]
            .sum()
            .reset_index()
            .sort_values(by="Total Venta", ascending=True)
        )
        fig_emp_usd = px.bar(
            df_emp_usd,
            x="Total Venta",
            y="Empleado",
            orientation="h",
            text="Total Venta",
            color="Total Venta",
            color_continuous_scale="Blues",
            template="plotly_white",
        )
        fig_emp_usd.update_traces(texttemplate="$%{x:,.2f}", textposition="outside")
        fig_emp_usd.update_layout(
            coloraxis_showscale=False, xaxis_title="Venta (USD)", yaxis_title="", margin=dict(t=40, l=150, r=80)
        )
        st.plotly_chart(fig_emp_usd, use_container_width=True)
    else:
        st.info("No hay datos para los filtros seleccionados.")

with row6_col2:
    if not df_filtered.empty:
        total_unidades_emp = df_filtered["Informar la Cantidad Vendida"].sum()
        st.subheader(f"🏆 Top Empleados - General en Unidades (Total: {total_unidades_emp:,.0f} u.)")
        df_emp_un = (
            df_filtered.groupby("Empleado")["Informar la Cantidad Vendida"]
            .sum()
            .reset_index()
            .sort_values(by="Informar la Cantidad Vendida", ascending=True)
        )
        fig_emp_un = px.bar(
            df_emp_un,
            x="Informar la Cantidad Vendida",
            y="Empleado",
            orientation="h",
            text="Informar la Cantidad Vendida",
            color="Informar la Cantidad Vendida",
            color_continuous_scale="Blues",
            template="plotly_white",
        )
        fig_emp_un.update_traces(texttemplate="%{x:,.0f}", textposition="outside")
        fig_emp_un.update_layout(
            coloraxis_showscale=False, xaxis_title="Cantidad (Unidades)", yaxis_title="", margin=dict(t=40, l=150, r=80)
        )
        st.plotly_chart(fig_emp_un, use_container_width=True)
    else:
        st.subheader("🏆 Top Empleados - General en Unidades")
        st.info("No hay datos para los filtros seleccionados.")

st.markdown("---")

# --- BOTÓN DE DESCARGA DE LA BASE DE DATOS FILTRADA ---
st.subheader("📥 Exportar Base de Datos")
@st.cache_data
def convert_df_to_csv(df_to_conv):
    return df_to_conv.to_csv(index=False).encode('utf-8')

csv_data = convert_df_to_csv(df_filtered)
st.download_button(
    label="📥 Descargar datos filtrados (CSV)",
    data=csv_data,
    file_name="base_datos_filtrada_mystic.csv",
    mime="text/csv",
)

st.markdown("---")

# --- APARTADO DE UNIDADES PARA BONO POR CATEGORÍA (FILTRADO) ---
row_bono = st.container()
with row_bono:
    df_comision_temp = df_filtered.copy()
    df_comision_temp["Venta Limpia"] = df_comision_temp["Venta Directa o Indirecta"].astype(str).str.strip().str.lower()
    
    df_comision = df_comision_temp[
        df_comision_temp["Venta Limpia"].str.contains("directa", na=False) & 
        ~df_comision_temp["Venta Limpia"].str.contains("indirecta", na=False) &
        (df_comision_temp["Informar precio de la venta"] >= 120)
    ]

    if not df_comision.empty:
        total_unidades_bono = df_comision["Informar la Cantidad Vendida"].sum()
        st.subheader(f"🎯 Unidades para bono (Ventas Directas y Precio >= $120) - por Supercategoría (Total: {total_unidades_bono:,.0f} u.)")

        df_comm_cat = (
            df_comision.groupby("Supercategoría")["Informar la Cantidad Vendida"]
            .sum()
            .reset_index()
            .sort_values(by="Informar la Cantidad Vendida", ascending=True)
        )

        fig_comm = px.bar(
            df_comm_cat,
            x="Informar la Cantidad Vendida",
            y="Supercategoría",
            orientation="h",
            color="Informar la Cantidad Vendida",
            color_continuous_scale="Blues",
            template="plotly_white",
        )

        fig_comm.update_traces(texttemplate="%{x:,.0f}", textposition="outside")

        fig_comm.update_layout(
            coloraxis_showscale=False, xaxis_title="Cantidad (Unidades)", yaxis_title="", margin=dict(t=40, l=150, r=80)
        )
        st.plotly_chart(fig_comm, use_container_width=True)
    else:
        st.subheader("🎯 Unidades para bono (Ventas Directas y Precio de Venta >= $120) - por Supercategoría")
        st.info(
            "No hay registros de ventas directas con un Precio de Venta >= $120 para los filtros seleccionados."
        )

st.markdown("---")

# --- APARTADO DE NÓMINA PROTEGIDO POR CLAVE (USA 'df' COMPLETO) ---
st.subheader("💵 Módulo de Nómina: Liquidación de Bonos por Escalas (Ventas Directas y Precio >= $120)")

CLAVE_CORRECTA = "jorgito2026"

password_ingresada = st.text_input(
    "🔒 Ingrese password de acceso", 
    type="password"
)

if password_ingresada == CLAVE_CORRECTA:
    st.success("¡Clave correcta! Módulo de nómina desbloqueado.")
    
    st.markdown(
        "Este módulo procesa las unidades de aquellas ventas que son **Directas** y donde el **Precio de Venta >= $120**. "
        "Muestra el bono correspondiente para cada **Promotor** y calcula el acumulado global del equipo para liquidar el bono del **Supervisor**."
    )

    with st.expander("📊 Ver Escalas Oficiales de Bonificación (Referencia)"):
        col_esc1, col_esc2 = st.columns(2)
        with col_esc1:
            st.markdown("**Escala de Bonificación para Promotor**")
            df_ref_prom = pd.DataFrame([
                {"Meta de Unidades": "30 a 39 unidades", "% Bono": "10%", "Total Bono ($)": "$10.00"},
                {"Meta de Unidades": "40 a 49 unidades", "% Bono": "25%", "Total Bono ($)": "$25.00"},
                {"Meta de Unidades": "50 a 69 unidades", "% Bono": "50%", "Total Bono ($)": "$50.00"},
                {"Meta de Unidades": "70 a 79 unidades", "% Bono": "75%", "Total Bono ($)": "$75.00"},
                {"Meta de Unidades": "≥ 80 unidades", "% Bono": "100%", "Total Bono ($)": "$100.00"}
            ])
            st.table(df_ref_prom)
        with col_esc2:
            st.markdown("**Escala de Bonificación para Supervisor**")
            df_ref_sup = pd.DataFrame([
                {"Meta Total de Equipo": "90 a 149 unidades", "% Bono": "10%", "Total Bono ($)": "$60.00"},
                {"Meta Total de Equipo": "150 a 209 unidades", "% Bono": "25%", "Total Bono ($)": "$120.00"},
                {"Meta Total de Equipo": "210 a 269 unidades", "% Bono": "50%", "Total Bono ($)": "$180.00"},
                {"Meta Total de Equipo": "270 a 329 unidades", "% Bono": "75%", "Total Bono ($)": "$240.00"},
                {"Meta Total de Equipo": "≥ 330 unidades", "% Bono": "100%", "Total Bono ($)": "$300.00"}
            ])
            st.table(df_ref_sup)

    # --- PROCESAMIENTO DE NÓMINA (USA df COMPLETO) ---
    df_nomina_base = df.copy()
    df_nomina_base["Venta Limpia"] = df_nomina_base["Venta Directa o Indirecta"].astype(str).str.strip().str.lower()
    
    df_bonificable = df_nomina_base[
        df_nomina_base["Venta Limpia"].str.contains("directa", na=False) & 
        ~df_nomina_base["Venta Limpia"].str.contains("indirecta", na=False) &
        (df_nomina_base["Informar precio de la venta"] >= 120)
    ]

    if not df_bonificable.empty and "Empleado" in df_bonificable.columns:
        df_promotores = df_bonificable.groupby("Empleado")["Informar la Cantidad Vendida"].sum().reset_index()
        df_promotores.columns = ["Promotor", "Unidades Bonificables"]

        def calcular_bono_promotor(unidades):
            if unidades >= 80:
                return 100.0, "100%", "≥ 80 u."
            elif unidades >= 70:
                return 75.0, "75%", "70 - 79 u."
            elif unidades >= 50:
                return 50.0, "50%", "50 - 69 u."
            elif unidades >= 40:
                return 25.0, "25%", "40 - 49 u."
            elif unidades >= 30:
                return 10.0, "10%", "30 - 39 u."
            else:
                return 0.0, "0%", "< 30 u."

        resultados_nomina = []
        for _, row in df_promotores.iterrows():
            emp = row["Promotor"]
            units = row["Unidades Bonificables"]
            bono_val, bono_pct, escala = calcular_bono_promotor(units)
            resultados_nomina.append({
                "Promotor": emp,
                "Unidades Bonificables": units,
                "Escala Alcanzada": escala,
                "% Bono": bono_pct,
                "Bono Total ($)": bono_val
            })

        df_resumen_nomina = pd.DataFrame(resultados_nomina)
        df_resumen_nomina = df_resumen_nomina.sort_values(by="Unidades Bonificables", ascending=False).reset_index(drop=True)

        st.markdown("### 📋 Liquidación de Bonos por Promotor")
        st.dataframe(df_resumen_nomina, use_container_width=True)

        total_bonos_promotores = df_resumen_nomina["Bono Total ($)"].sum()
        total_unidades_equipo = df_promotores["Unidades Bonificables"].sum()

        def calcular_bono_supervisor(unidades_eq):
            if unidades_eq >= 330:
                return 300.0, "100%", "≥ 330 u."
            elif unidades_eq >= 270:
                return 240.0, "75%", "270 - 329 u."
            elif unidades_eq >= 210:
                return 180.0, "50%", "210 - 269 u."
            elif unidades_eq >= 150:
                return 120.0, "25%", "150 - 209 u."
            elif unidades_eq >= 90:
                return 60.0, "10%", "90 - 149 u."
            else:
                return 0.0, "0%", "< 90 u."

        sup_val, sup_pct, sup_escala = calcular_bono_supervisor(total_unidades_equipo)

        st.markdown("### 👔 Liquidación de Bono para Supervisor")
        col_sup1, col_sup2, col_sup3 = st.columns(3)
        with col_sup1:
            st.metric("Unidades Totales del Equipo", f"{total_unidades_equipo:,.0f} u.")
        with col_sup2:
            st.metric("Escala / % Supervisor", f"{sup_escala} ({sup_pct})")
        with col_sup3:
            st.metric("Bono Supervisor ($)", f"${sup_val:,.2f}")

        st.markdown("---")
        st.markdown("### 💰 Resumen Financiero Total de Nómina (Bonos)")
        gran_total_nomina = total_bonos_promotores + sup_val
        col_fin1, col_fin2, col_fin3 = st.columns(3)
        with col_fin1:
            st.metric("Total Bonos Promotores", f"${total_bonos_promotores:,.2f}")
        with col_fin2:
            st.metric("Total Bono Supervisor", f"${sup_val:,.2f}")
        with col_fin3:
            st.metric("Inversión Total Nómina Bonos", f"${gran_total_nomina:,.2f}")
    else:
        st.warning("No hay datos de ventas directas con precio >= $120 para calcular la nómina.")
else:
    if password_ingresada != "":
        st.error("Clave incorrecta. Intente de nuevo.")