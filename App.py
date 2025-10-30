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

# Convertimos la columna PrecioUnit a numérico (por si viene con símbolo €)
df_materiales["PrecioUnit"] = (
    df_materiales["PrecioUnit"]
    .astype(str)
    .str.replace("€", "")
    .str.replace(",", ".")
    .astype(float)
)

# =====================
# 💵 CARGA DE TARIFAS
# =====================
hoja_tarifas = "BD_Tarifas"
df_tarifas = pd.read_excel(ruta_completa, sheet_name=hoja_tarifas)

# Limpiamos posibles espacios o formatos
df_tarifas.columns = df_tarifas.columns.str.strip()
df_tarifas["Técnico"] = df_tarifas["Técnico"].str.strip()
df_tarifas["Tipo"] = df_tarifas["Tipo"].str.strip()

# Convertimos la columna TARIFA a numérico (por si viene con símbolo €)
df_tarifas["TARIFA"] = (
    df_tarifas["TARIFA"]
    .astype(str)
    .str.replace("€", "")
    .str.replace(",", ".")
    .astype(float)
)

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
    if funcion == "Vision&Control":
        return df[df["VisControl"] == 1]
    elif funcion == "Vision":
        return df[df["Vis"] == 1]
    return df.copy()

def filtrar_por_iluminacion(df, iluminacion):
    df_filtrado = df.copy()

    if iluminacion.upper() == "NO":
        df_filtrado = df_filtrado[
            df_filtrado["Descripcion"].str.upper().str.strip() != "FOCO LED 10W 220V 6500K 1150LM"
        ]

    # Si es "SI" o cualquier otro valor, no se toca
    return df_filtrado

def filtrar_por_clima(df, clima):
    if clima == "Ambiente":
        return df[df["Clima"].isin(["Ambos", "Ambiente"])]
    elif clima == "Frio":
        return df[df["Clima"].isin(["Ambos", "Frio"])]
    return df.copy()

def filtrar_por_pc_nuevo(df, pc_nuevo):
    if pc_nuevo == "NO":
        return df[df["Descripcion"].str.upper() != "PC INDUSTRIAL CPC3G"]
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
    funcion = st.selectbox("Función", ["Vision", "Vision&Control"])
    iluminacion = st.selectbox("Iluminación", ["SI", "NO"])
    clima = st.selectbox("Clima", ["Ambiente", "Frio"])
    pc_nuevo = st.selectbox("PC Nuevo", ["SI", "NO"])
    distancia = st.number_input("Distancia Ft->Arm (m)", min_value=0, max_value=1000, value=100, step=10)
    num_maq = st.number_input("Número de máquinas", min_value=1, max_value=50, value=1, step=1)

    st.header("👷 Mano de obra")
    km_desplazamiento = st.number_input(
    "Kilómetros desde la delegación a la instalación (ida + vuelta)",
    min_value=0,
    max_value=9000,
    value=0,
    step=10,
    key="km_desplazamiento"
    )

    h_desplazamiento = st.number_input(
    "Horas de desplazamiento (ida + vuelta)",
    min_value=0,
    max_value=24,
    value=0,
    step=1,
    key="h_desplazamiento"
    )

    trabajo_frio = st.selectbox(
    "Trabajo en frío (<4ºC)",
    ["NO", "SI"],
    key="trabajo_frio"
    )

    estancia = st.selectbox(
    "Estancia del operario",
    ["VIAJA", "PERNOCTA"],
    key="estancia"
    )
# =====================
# 🧮 APLICAR FILTROS
# =====================
df_filtrado = filtrar_por_red(df_materiales, red)
df_filtrado = filtrar_por_funcion(df_filtrado, funcion)
df_filtrado = filtrar_por_iluminacion(df_filtrado, iluminacion)
df_filtrado = filtrar_por_clima(df_filtrado, clima)
df_filtrado = filtrar_por_pc_nuevo(df_filtrado, pc_nuevo)
df_filtrado = tratar_por_distancia(df_filtrado, distancia, red)
df_filtrado = tratar_por_num_maquinas(df_filtrado, num_maq)

df_interno = df_filtrado[df_filtrado["Categoria"] == "Interno"]
df_externo = df_filtrado[df_filtrado["Categoria"] == "Externo"]

# =====================
# 💰 CALCULAR PRECIOS
# =====================
df_filtrado["PrecioTotal"] = df_filtrado["Cantidad"] * df_filtrado["PrecioUnit"]

df_interno = df_filtrado[df_filtrado["Categoria"] == "Interno"]
df_externo = df_filtrado[df_filtrado["Categoria"] == "Externo"]


# =====================
# 📊 MOSTRAR RESULTADOS LISTADO DE MATERIALES
# =====================
st.markdown("## 1. Materiales")
st.markdown("### Material Interno")
st.dataframe(
    df_interno.reset_index(drop=True)[["Descripcion", "Cantidad", "PrecioUnit", "PrecioTotal"]],
    use_container_width=True,
    hide_index=True
)
st.markdown(f"**Subtotal Interno:** {df_interno['PrecioTotal'].sum():,.2f} €")

st.markdown("### Material Externo")
st.dataframe(
    df_externo.reset_index(drop=True)[["Descripcion", "Cantidad", "PrecioUnit", "PrecioTotal"]],
    use_container_width=True, hide_index=True
)
st.markdown(f"**Subtotal Externo:** {df_externo['PrecioTotal'].sum():,.2f} €")

st.markdown("---")
st.markdown(f"**Total Material:** {df_filtrado['PrecioTotal'].sum():,.2f} €")


# =====================
# 💾 DESCARGA DE RESULTADOS
# =====================

#output = BytesIO()
#with pd.ExcelWriter(output, engine="openpyxl") as writer:
#    df_filtrado.to_excel(writer, index=False, sheet_name="Listado")
#output.seek(0)

#st.download_button(
#    label="💾 Descargar listado filtrado (Excel)",
#    data=output,
#    file_name="Listado_materiales_filtrado.xlsx",
#    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
#)

# =====================
# 👷 3. MANO DE OBRA
# =====================

st.markdown("---")
st.markdown("## 2. Mano de obra")

# --- Cálculo de días laborales según red y función ---

def calcular_dias_tecnico_electromecanico(red, funcion, num_maq):
    if red == "Profibus":
        if funcion == "Vision":
            return 1 * num_maq
        elif funcion == "Vision&Control":
            return 1 * num_maq + 2
    elif red == "Profinet":
        if funcion == "Vision":
            return 0.5 * num_maq
        elif funcion == "Vision&Control":
            return 0.5 * num_maq + 2
    return 0

def calcular_dias_tecnico_control(red, funcion, num_maq):
    if red == "Profibus":
        if funcion == "Vision":
            return 0.5 * num_maq + 2
        elif funcion == "Vision&Control":
            return 0.5 * num_maq + 4
    elif red == "Profinet":
        if funcion == "Vision":
            return 0.25 * num_maq + 2
        elif funcion == "Vision&Control":
            return 0.25 * num_maq + 4
    return 0

# --- Aplicar lógica ---
dias_electromecanico = calcular_dias_tecnico_electromecanico(red, funcion, num_maq)
dias_control = calcular_dias_tecnico_control(red, funcion, num_maq)
num_electromecanico = 2
num_tec_control=1
horas_act =8

# --- Crear DataFrames para mostrar ---
df_mano_obra = pd.DataFrame({
    "Técnico": ["Técnico Electromecánico", "Técnico de Control"],
    "Nº Personas": [num_electromecanico, num_tec_control],
    "Días Laborales": [dias_electromecanico, dias_control]
})

# --- Mostrar resultados ---
st.markdown("### Técnico Electromecánico")
st.dataframe(df_mano_obra.reset_index(drop=True)[df_mano_obra["Técnico"] == "Técnico Electromecánico"][["Nº Personas", "Días Laborales"]],
             use_container_width=True,hide_index=True)

st.markdown("### Técnico de Control")
st.dataframe(df_mano_obra.reset_index(drop=True)[df_mano_obra["Técnico"] == "Técnico de Control"][["Nº Personas", "Días Laborales"]],
         use_container_width=True,hide_index=True)


# =====================
# 💰 CÁLCULO DE COSTES AGRUPADOS POR TIPO
# =====================

def obtener_tarifa(df, tecnico, palabra_clave):
    """Devuelve la tarifa y tipo correspondientes según el técnico y la descripción."""
    filtro = (
        (df["Técnico"].isin([tecnico, "Ambos"])) &
        (df["DESCRIPCION"].str.contains(palabra_clave, case=False, na=False))
    )
    try:
        fila = df.loc[filtro].iloc[0]
        return fila["TARIFA"], fila["Tipo"]
    except IndexError:
        return 0, None


# --- Obtener tarifas y tipos ---
tarifa_hora_electro, tipo_hora_electro = obtener_tarifa(df_tarifas, "Técnico Electromecánico", "HORA TRABAJO NORMAL")
tarifa_hora_ctrl, tipo_hora_ctrl = obtener_tarifa(df_tarifas, "Técnico de Control", "HORA TRABAJO NORMAL")
tarifa_hora_despl_electro, tipo_hora_despl_electro = obtener_tarifa(df_tarifas, "Técnico Electromecánico", "DESPLAZAMIENTO TECNICO")
tarifa_hora_despl_ctrl, tipo_hora_despl_ctrl = obtener_tarifa(df_tarifas, "Técnico de Control", "DESPLAZAMIENTO INGENIERO")

tarifa_media_dieta, tipo_media_dieta = obtener_tarifa(df_tarifas, "Ambos", "MEDIA DIETA")
tarifa_dieta_completa, tipo_dieta_completa = obtener_tarifa(df_tarifas, "Ambos", "DIETA COMPLETA")
tarifa_pernocta, tipo_pernocta = obtener_tarifa(df_tarifas, "Ambos", "PERNOCTACION")
tarifa_km, tipo_km = obtener_tarifa(df_tarifas, "Ambos", "KILOMETRO")
tarifa_frio, tipo_frio = obtener_tarifa(df_tarifas, "Ambos", "FRIO")

# --- Cálculo de cantidades ---
cant_horas_electro = dias_electromecanico * num_electromecanico * horas_act
cant_despl_electro = num_electromecanico * h_desplazamiento
cant_media_dieta_electro = dias_electromecanico * num_electromecanico if estancia == "VIAJA" else 0
cant_dieta_completa_electro = dias_electromecanico * num_electromecanico if estancia == "PERNOCTA" else 0
cant_pernocta_electro = dias_electromecanico * num_electromecanico if estancia == "PERNOCTA" else 0
cant_km_electro = km_desplazamiento
cant_frio_electro = dias_electromecanico * num_electromecanico * horas_act if trabajo_frio == "SI" else 0

cant_horas_ctrl = dias_control * num_tec_control * horas_act
cant_despl_ctrl = num_tec_control * h_desplazamiento
cant_media_dieta_ctrl = dias_control * num_tec_control if estancia == "VIAJA" else 0
cant_dieta_completa_ctrl = dias_control * num_tec_control if estancia == "PERNOCTA" else 0
cant_pernocta_ctrl = dias_control * num_tec_control if estancia == "PERNOCTA" else 0
cant_km_ctrl = km_desplazamiento
cant_frio_ctrl = dias_control * num_tec_control * horas_act if trabajo_frio == "SI" else 0

# --- Definición de estructura de costes con tipo ---
datos_coste = [
    # Tipo Normal
    {"Tipo": tipo_hora_electro, "Concepto": "Hora trabajo normal", "Electromecánico": cant_horas_electro * tarifa_hora_electro, "Control": cant_horas_ctrl * tarifa_hora_ctrl},
    {"Tipo": tipo_hora_despl_electro, "Concepto": "Hora desplazamiento", "Electromecánico": cant_despl_electro * tarifa_hora_despl_electro, "Control": cant_despl_ctrl * tarifa_hora_despl_ctrl},
    # Tipo Complementos
    {"Tipo": tipo_media_dieta, "Concepto": "Media dieta", "Electromecánico": cant_media_dieta_electro * tarifa_media_dieta, "Control": cant_media_dieta_ctrl * tarifa_media_dieta},
    {"Tipo": tipo_dieta_completa, "Concepto": "Dieta completa", "Electromecánico": cant_dieta_completa_electro * tarifa_dieta_completa, "Control": cant_dieta_completa_ctrl * tarifa_dieta_completa},
    {"Tipo": tipo_pernocta, "Concepto": "Pernocta", "Electromecánico": cant_pernocta_electro * tarifa_pernocta, "Control": cant_pernocta_ctrl * tarifa_pernocta},
    {"Tipo": tipo_km, "Concepto": "Kilómetros", "Electromecánico": cant_km_electro * tarifa_km, "Control": cant_km_ctrl * tarifa_km},
    {"Tipo": tipo_frio, "Concepto": "Trabajo en frío", "Electromecánico": cant_frio_electro * tarifa_frio, "Control": cant_frio_ctrl * tarifa_frio},
]

df_costes = pd.DataFrame(datos_coste)

# --- Subtotales por Tipo ---
df_subtotales = df_costes.groupby("Tipo", dropna=False)[["Electromecánico", "Control"]].sum().reset_index()
df_subtotales["Concepto"] = "Subtotal " + df_subtotales["Tipo"]
df_subtotales = df_subtotales[["Tipo", "Concepto", "Electromecánico", "Control"]]

# --- Total general ---
df_total = pd.DataFrame([{
    "Tipo": "",
    "Concepto": "TOTAL GENERAL",
    "Electromecánico": df_costes["Electromecánico"].sum(),
    "Control": df_costes["Control"].sum()
}])

# --- Unión final ---
df_final = pd.concat([df_costes, df_subtotales, df_total], ignore_index=True)

# --- Mostrar resultados ---
st.markdown("### Costes Mano de obra")
for col in ["Electromecánico", "Control"]:
    df_final[col] = pd.to_numeric(df_final[col], errors="coerce")
df_final[["Electromecánico", "Control"]] = df_final[["Electromecánico", "Control"]].fillna(0.0)
styler = df_final.style.format({
    "Electromecánico": "{:,.2f} €",
    "Control": "{:,.2f} €"
})
st.dataframe(styler, use_container_width=True,hide_index=True)


# =====================
# 🧠 4. HORAS DE INGENIERÍA
# =====================

# Definir tarifa e horas según tipo de instalación
tarifa_ingenieria = 190

if funcion == "Vision":
    horas_ingenieria = 40
elif funcion == "Vision&Control":
    horas_ingenieria = 64
else:
    horas_ingenieria = 0

# Calcular coste total
coste_ingenieria = horas_ingenieria * tarifa_ingenieria

# Mostrar resultados
st.markdown("---")
st.markdown("## 3. Horas de ingeniería")

df_ingenieria = pd.DataFrame({
    "Concepto": ["Horas de ingeniería"],
    "Horas": [horas_ingenieria],
    "Tarifa (€/h)": [tarifa_ingenieria],
    "Coste Total (€)": [coste_ingenieria]
})

st.dataframe(
    df_ingenieria.style.format({
        "Tarifa (€/h)": "{:,.2f} €",
        "Coste Total (€)": "{:,.2f} €"
    }),
    use_container_width=True,hide_index=True
)



# =====================
# 📊 5. RESUMEN DE OFERTA
# =====================
st.markdown("---")
st.header("📊 4. Resumen de oferta")

# --- Datos generales ---
st.subheader("Datos generales")
col1, col2 = st.columns(2)
with col1:
    fecha_decision = st.date_input("📅 Fecha de decisión")
    poblacion = st.text_input("🏙️ Población", "")
with col2:
    cliente = st.text_input("👤 Cliente", "")
    delegacion = st.text_input("🏢 Delegación", "")

# --- Sección de descuentos ---
st.subheader("Descuentos aplicables (%)")

col1, col2, col3 = st.columns(3)
with col1:
    descuento_interno = st.number_input("Descuento Materiales Internos (%)", min_value=0.0, max_value=100.0, value=10.0, step=0.5)
with col2:
    descuento_externo = st.number_input("Descuento Materiales Externos (%)", min_value=0.0, max_value=100.0, value=55.0, step=0.5)
with col3:
    descuento_mano_obra = st.number_input("Descuento Mano de Obra (%)", min_value=0.0, max_value=100.0, value=30.0, step=0.5)

descuento_ingenieria = st.number_input("Descuento Horas de Ingeniería (%)", min_value=0.0, max_value=100.0, value=30.0, step=0.5)

# --- Calcular totales ---
total_interno = df_interno["PrecioTotal"].sum()
total_externo = df_externo["PrecioTotal"].sum()

total_mano_obra = df_final.loc[df_final["Concepto"] == "TOTAL GENERAL", ["Electromecánico", "Control"]].sum().sum()
total_ingenieria = coste_ingenieria

# Aplicar descuentos
neto_interno = total_interno * (1 - descuento_interno / 100)
neto_externo = total_externo * (1 - descuento_externo / 100)
neto_mano_obra = total_mano_obra * (1 - descuento_mano_obra / 100)
neto_ingenieria = total_ingenieria * (1 - descuento_ingenieria / 100)

# Total general
total_tarifa = total_interno + total_externo + total_mano_obra + total_ingenieria
total_descuento = ((total_tarifa - (neto_interno + neto_externo + neto_mano_obra + neto_ingenieria)) / total_tarifa) * 100
total_neto = neto_interno + neto_externo + neto_mano_obra + neto_ingenieria

# --- Crear tabla resumen ---
df_resumen = pd.DataFrame({
    "Concepto": [
        "1. Materiales Internos",
        "2. Materiales Externos",
        "3. Mano de Obra, Control y Puesta en Marcha",
        "4. Ingeniería"
    ],
    "Tarifa (€)": [
        total_interno,
        total_externo,
        total_mano_obra,
        total_ingenieria
    ],
    "Descuento (%)": [
        descuento_interno,
        descuento_externo,
        descuento_mano_obra,
        descuento_ingenieria
    ],
    "Neto Cliente (€)": [
        neto_interno,
        neto_externo,
        neto_mano_obra,
        neto_ingenieria
    ]
})

# --- Mostrar resumen ---
st.markdown("### 🧾 Resumen de importes")
st.dataframe(
    df_resumen.style.format({
        "Tarifa (€)": "{:,.2f} €",
        "Descuento (%)": "{:.2f} %",
        "Neto Cliente (€)": "{:,.2f} €"
    }),
    use_container_width=True,hide_index=True
)

# --- Totales ---
st.markdown("### 💵 Totales generales")
col1, col2, col3 = st.columns([1, 1, 1])
col1.metric("Total Tarifa", f"{total_tarifa:,.2f} €")
col2.metric("Descuento medio", f"{total_descuento:,.2f} %")
col3.metric("Total Neto Cliente", f"{total_neto:,.2f} €")


# =====================
# 💾 EXPORTAR A EXCEL
# =====================

output = BytesIO()
with pd.ExcelWriter(output, engine="openpyxl") as writer:
    # Hoja 1: materiales filtrados
    df_filtrado.to_excel(writer, index=False, sheet_name="Listado")

    # Hoja 2: costes mano de obra
    df_final.to_excel(writer, index=False, sheet_name="Costes Mano de Obra")

    # Hoja 3: resumen de oferta
    df_resumen.to_excel(writer, index=False, sheet_name="Resumen Oferta")

    # (Opcional) incluir una hoja con los parámetros seleccionados
    parametros = pd.DataFrame({
        "Campo": ["Fecha decisión", "Población", "Cliente", "Delegación"],
        "Valor": [fecha_decision, poblacion, cliente, delegacion]
    })
    parametros.to_excel(writer, index=False, sheet_name="Datos Generales")

output.seek(0)

st.download_button(
    label="💾 Descargar oferta completa (Excel)",
    data=output,
    file_name=f"Oferta_Camaras_{delegacion or 'General'}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
