import pandas as pd
import streamlit as st
from pathlib import Path
from io import BytesIO

# =====================
# 📁 CARGA DE DATOS
# =====================
ruta_carpeta = Path(__file__).parent
nombre_archivo = "df_materiales.xlsx"
hoja_materiales = "BD_Materiales"
ruta_completa = ruta_carpeta / nombre_archivo

df_materiales = pd.read_excel(ruta_completa, sheet_name=hoja_materiales)

# =====================
# 🧩 FUNCIONES DE FILTRO
# =====================
def filtrar_por_red(df, red):
    if red == "Profinet":
        return df[df["Red"].isin(["Ambos", "Profinet"])]
    elif red == "Profibus":
        return df[df["Red"].isin(["Ambos", "Profibus"])]
    return df.copy()

def filtrar_por_funcion(df, funcion):
    if funcion == "VisionControl":
        return df[df["VisControl"] == 1]
    elif funcion == "Vision":
        return df[df["Vis"] == 1]
    return df.copy()

def filtrar_por_vision_nocturna(df, vision_nocturna):
    df_filtrado = df.copy()

    if vision_nocturna.upper() == "SI":
        df_filtrado = df_filtrado[
            df_filtrado["Descripcion"].str.upper().str.strip() != "FOCO LED 10W 220V 6500K 1150LM"
        ]

    # Si es "NO" o cualquier otro valor, no se toca
    return df_filtrado


def filtrar_por_clima(df, clima):
    if clima == "Ambiente":
        return df[df["Clima"].isin(["Ambos", "Ambiente"])]
    elif clima == "Frio":
        return df[df["Clima"].isin(["Ambos", "Frio"])]
    return df.copy()

def filtrar_por_tablet(df, tablet):
    if tablet == "NO":
        return df[df["Descripcion"].str.upper() != "TABLET"]
    return df.copy()

def filtrar_por_pc_nuevo(df, pc_nuevo):
    if pc_nuevo == "NO":
        return df[df["Descripcion"].str.upper() != "CPC 1 PC"]
    return df.copy()

def tratar_por_distancia(df, distancia, red):
    df_filtrado = df.copy()

    # Actualizar cantidad de la manguera
    mask_manguera = df_filtrado["Descripcion"].str.upper() == "MANGUERA  3G1,0 NUM.-ML"
    df_filtrado.loc[mask_manguera, "Cantidad"] = distancia

    # Eliminar SWITCH si no aplica (solo se mantiene si Profibus y distancia > 150)
    if not (red == "Profibus" and distancia >= 150):
        df_filtrado = df_filtrado[df_filtrado["Descripcion"].str.upper() != "SWITCH ETHERNET GESTIONABLE"]

    return df_filtrado

def tratar_por_num_maquinas(df, num_maq):
    df_filtrado = df.copy()
    df_filtrado["Cantidad"] = pd.to_numeric(df_filtrado["Cantidad"], errors="coerce")
    df_filtrado.loc[df_filtrado["Unico"] == 0, "Cantidad"] *= int(num_maq)
    return df_filtrado

# =====================
# 🎛️ INTERFAZ STREAMLIT
# =====================
st.set_page_config(page_title="📋 Oferta Cámaras Webs", layout="centered")
st.title("📋 Oferta Cámaras Webs")

with st.sidebar:
    st.header("⚙️ Parámetros de configuración")

    red = st.selectbox("Red", ["Profinet", "Profibus"])
    funcion = st.selectbox("Función", ["Vision", "VisionControl"])
    vision_nocturna = st.selectbox("Visión Nocturna", ["SI", "NO"])
    clima = st.selectbox("Clima", ["Ambiente", "Frio"])
    tablet = st.selectbox("Tablet", ["SI", "NO"])
    pc_nuevo = st.selectbox("PC Nuevo", ["SI", "NO"])
    distancia = st.number_input("Distancia (m)", min_value=0, max_value=1000, value=100, step=10)
    num_maq = st.number_input("Número de máquinas", min_value=1, max_value=50, value=1, step=1)

# =====================
# 🧮 APLICAR FILTROS
# =====================
df_filtrado = filtrar_por_red(df_materiales, red)
df_filtrado = filtrar_por_funcion(df_filtrado, funcion)
df_filtrado = filtrar_por_vision_nocturna(df_filtrado, vision_nocturna)
df_filtrado = filtrar_por_clima(df_filtrado, clima)
df_filtrado = filtrar_por_tablet(df_filtrado, tablet)
df_filtrado = filtrar_por_pc_nuevo(df_filtrado, pc_nuevo)
df_filtrado = tratar_por_distancia(df_filtrado, distancia, red)
df_filtrado = tratar_por_num_maquinas(df_filtrado, num_maq)

df_interno = df_filtrado[df_filtrado["Categoria"] == "Interno"]
df_externo = df_filtrado[df_filtrado["Categoria"] == "Externo"]

# =====================
# 📊 MOSTRAR RESULTADOS
# =====================

st.markdown("### 🔧 Listado Interno")
st.dataframe(df_interno.reset_index(drop=True)[["Descripcion", "Cantidad"]].style.format({"Cantidad": "{:g}"}).hide(axis="index"), use_container_width=True)

st.markdown("### 📦 Listado Externo")
st.dataframe(df_externo.reset_index(drop=True)[["Descripcion", "Cantidad"]].style.format({"Cantidad": "{:g}"}).hide(axis="index"), use_container_width=True)

# =====================
# 💾 DESCARGA DE RESULTADOS
# =====================
output = BytesIO()
with pd.ExcelWriter(output, engine="openpyxl") as writer:
    df_filtrado.to_excel(writer, index=False, sheet_name="Listado")
output.seek(0)

st.download_button(
    label="💾 Descargar listado filtrado (Excel)",
    data=output,
    file_name="Listado_materiales_filtrado.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
