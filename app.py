
from pathlib import Path
from io import BytesIO
import re
import unicodedata

import pandas as pd
import plotly.express as px
import streamlit as st


# ==========================================================
# 1. CONFIGURACIÓN
# ==========================================================

st.set_page_config(
    page_title="Proyección Académica | Saber Pro y TyT",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

CARPETA = Path(__file__).resolve().parent
ARCHIVO_BASE = CARPETA / "PROYECCIÓN.xlsx"

CATEGORIAS = [
    "Candidato",
    "Próximo",
    "Fuera de proyección",
    "Sin información"
]

COLORES = {
    "Candidato": "#248653",
    "Próximo": "#D99120",
    "Fuera de proyección": "#64748B",
    "Sin información": "#A8B2C1"
}


# ==========================================================
# 2. ESTILO VISUAL
# ==========================================================

st.markdown("""
<style>
    .stApp {
        background-color: #F4F7FB;
    }

    .block-container {
        max-width: 1550px;
        padding-top: 1.5rem;
        padding-bottom: 2rem;
    }

    h1 {
        color: #17365D;
        font-weight: 800;
    }

    h2, h3 {
        color: #24476B;
    }

    div[data-testid="stMetric"] {
        background: #FFFFFF;
        border: 1px solid #E1E8F0;
        padding: 17px 19px;
        border-radius: 14px;
        box-shadow: 0 3px 12px rgba(23,54,93,0.05);
    }

    div[data-testid="stMetricLabel"] {
        color: #64748B;
        font-size: 0.88rem;
    }

    div[data-testid="stMetricValue"] {
        color: #17365D;
        font-weight: 750;
    }

    div[data-testid="stTabs"] button {
        font-weight: 650;
    }

    section[data-testid="stSidebar"] {
        background-color: #EAF0F7;
    }
</style>
""", unsafe_allow_html=True)


# ==========================================================
# 3. FUNCIONES DE LIMPIEZA
# ==========================================================

def normalizar_nombre(nombre):
    texto = str(nombre).strip().upper()
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(
        c for c in texto if not unicodedata.combining(c)
    )
    texto = re.sub(r"[^A-Z0-9]+", "_", texto)
    return texto.strip("_")


def normalizar_texto(valor):
    if pd.isna(valor):
        return ""

    texto = str(valor).strip().upper()
    texto = unicodedata.normalize("NFKD", texto)

    return "".join(
        c for c in texto if not unicodedata.combining(c)
    )


def convertir_numero(serie):

    def convertir(valor):
        if pd.isna(valor):
            return float("nan")

        texto = str(valor).strip()

        if texto in ("", "nan", "None", "-"):
            return float("nan")

        texto = texto.replace("%", "").replace(" ", "")

        if "," in texto and "." in texto:
            if texto.rfind(",") > texto.rfind("."):
                texto = texto.replace(".", "").replace(",", ".")
            else:
                texto = texto.replace(",", "")
        elif "," in texto:
            texto = texto.replace(",", ".")

        try:
            return float(texto)
        except ValueError:
            return float("nan")

    return serie.map(convertir)


def clasificar_avance(porcentaje):

    if pd.isna(porcentaje):
        return "Sin información"

    if porcentaje >= 75:
        return "Candidato"

    if porcentaje >= 60:
        return "Próximo"

    return "Fuera de proyección"


def clasificar_nivel(nivel):

    nivel = normalizar_texto(nivel)

    if "TECNICO PROFESIONAL" in nivel or "TECNOLOG" in nivel:
        return "Saber TyT"

    if "PREGRADO" in nivel or "LICENCIATURA" in nivel:
        return "Saber Pro"

    return "Por validar"


# ==========================================================
# 4. PREPARACIÓN DE LA BASE
# ==========================================================

def preparar_base(base):

    df = base.copy()

    df.columns = [
        normalizar_nombre(c) for c in df.columns
    ]

    equivalencias = {
        "ID": [
            "ID", "CODIGO_ESTUDIANTE",
            "ID_ESTUDIANTE", "CODIGO"
        ],
        "TIPO": ["TIPO", "TIPO_SEDE"],
        "SEDE": [
            "SEDE", "CENTRO_UNIVERSITARIO",
            "CENTRO", "CU"
        ],
        "FACULTAD": ["FACULTAD", "COD_FACULTAD"],
        "COD": [
            "COD", "CODIGO_PROGRAMA", "COD_PROGRAMA"
        ],
        "PROGRAMA": ["PROGRAMA", "NOMBRE_PROGRAMA"],
        "NIVEL": ["NIVEL", "NIVEL_ACADEMICO"],
        "JORNADA": ["JORNADA"],
        "TOTAL": [
            "TOTAL", "TOTAL_CREDITOS", "CREDITOS_TOTALES"
        ],
        "APROBADOS": [
            "APROBADOS", "CREDITOS_APROBADOS"
        ],
        "SISTEMA": ["SISTEMA", "SEMESTRE_SISTEMA"],
        "2026": ["2026", "SEMESTRE_2026"],
        "2027": ["2027", "SEMESTRE_2027"],
        "PORCENTAJE": [
            "PORCENTAJE", "AVANCE",
            "PORCENTAJE_AVANCE", "PORCENTAJE_CREDITOS"
        ]
    }

    renombrar = {}

    for destino, alternativas in equivalencias.items():
        for alternativa in alternativas:
            nombre = normalizar_nombre(alternativa)

            if nombre in df.columns:
                if destino not in df.columns or nombre == destino:
                    renombrar[nombre] = destino
                    break

    df = df.rename(columns=renombrar)

    requeridas = [
        "ID", "TIPO", "SEDE", "FACULTAD", "COD",
        "PROGRAMA", "NIVEL", "APROBADOS", "TOTAL",
        "2026", "2027"
    ]

    faltantes = [
        columna for columna in requeridas
        if columna not in df.columns
    ]

    if faltantes:
        raise ValueError(
            "Faltan columnas requeridas: "
            + ", ".join(faltantes)
            + ". Verifica los encabezados y la fila inicial "
              "de PROYECCIÓN.xlsx."
        )

    # Quitar únicamente filas completamente vacías.
    df = df.dropna(how="all").copy()

    columnas_texto = [
        "ID", "TIPO", "SEDE", "FACULTAD",
        "COD", "PROGRAMA", "NIVEL", "JORNADA"
    ]

    for col in columnas_texto:
        if col in df.columns:
            df[col] = df[col].astype("string").str.strip()

    # Excluir filas que no tienen identificador.
    df = df[
        df["ID"].notna() & (df["ID"].str.strip() != "")
    ].copy()

    # Convertir datos numéricos.
    for col in ["TOTAL", "APROBADOS", "2026", "2027"]:
        df[col] = convertir_numero(df[col])

    # Normalizar campos categóricos.
    for col in [
        "TIPO", "SEDE", "FACULTAD",
        "COD", "PROGRAMA", "NIVEL"
    ]:
        df[col] = df[col].fillna("Sin dato").astype(str).str.strip()
        df.loc[df[col] == "", col] = "Sin dato"

    # Calcular el avance a partir de los créditos.
    calculado = (
        df["APROBADOS"]
        / df["TOTAL"].where(df["TOTAL"] > 0)
        * 100
    )

    # Usar el porcentaje de la fuente cuando no pueda calcularse.
    if "PORCENTAJE" in df.columns:
        original = convertir_numero(df["PORCENTAJE"])
        df["PORCENTAJE_ANALISIS"] = calculado.fillna(original)
    else:
        df["PORCENTAJE_ANALISIS"] = calculado

    # Los porcentajes fuera de 0–100 no se consideran válidos.
    df.loc[
        ~df["PORCENTAJE_ANALISIS"].between(0, 100),
        "PORCENTAJE_ANALISIS"
    ] = float("nan")

    # Clasificación única y consistente.
    df["CATEGORIA"] = df["PORCENTAJE_ANALISIS"].apply(
        clasificar_avance
    )

    df["EN_PROYECCION"] = df["CATEGORIA"].isin(
        ["Candidato", "Próximo"]
    )

    df["PRUEBA_ORIENTATIVA"] = df["NIVEL"].apply(
        clasificar_nivel
    )

    return df


# ==========================================================
# 5. FUNCIONES DE RESUMEN
# ==========================================================

def resumen_por_grupo(df, columnas):

    filas = []

    for clave, grupo in df.groupby(columnas, dropna=False):

        if not isinstance(clave, tuple):
            clave = (clave,)

        fila = dict(zip(columnas, clave))

        fila["Estudiantes únicos"] = grupo["ID"].nunique()

        # Contar IDs únicos por categoría.
        for categoria in CATEGORIAS:
            fila[categoria] = grupo.loc[
                grupo["CATEGORIA"] == categoria, "ID"
            ].nunique()

        fila["En proyección"] = (
            fila["Candidato"] + fila["Próximo"]
        )

        # Promedio sobre una fila por ID en cada grupo.
        fila["Avance promedio (%)"] = (
            grupo.drop_duplicates("ID")["PORCENTAJE_ANALISIS"].mean()
        )

        filas.append(fila)

    return pd.DataFrame(filas)


def formato_tabla(df):

    resultado = df.copy()

    if "Avance promedio (%)" in resultado.columns:
        resultado["Avance promedio (%)"] = (
            resultado["Avance promedio (%)"].round(1)
        )

    return resultado


def exportar_excel(hojas):

    archivo = BytesIO()

    with pd.ExcelWriter(archivo, engine="openpyxl") as writer:
        for nombre, datos in hojas.items():
            datos.to_excel(
                writer,
                sheet_name=nombre[:31],
                index=False
            )

    archivo.seek(0)
    return archivo


def grafico_barras_categoria(df, eje_x, titulo):

    fig = px.bar(
        df,
        x=eje_x,
        y=["Candidato", "Próximo"],
        barmode="stack",
        color_discrete_map=COLORES,
        template="plotly_white",
        title=titulo,
        labels={
            "value": "Estudiantes únicos",
            "variable": "Categoría"
        }
    )

    fig.update_layout(
        legend_title_text="Clasificación",
        margin=dict(l=15, r=15, t=60, b=20)
    )

    return fig


# ==========================================================
# 6. ENCABEZADO
# ==========================================================

st.title("Dashboard de Proyección Académica")

st.caption(
    "Seguimiento de población estudiantil y proyección "
    "para Saber Pro y Saber TyT"
)

st.markdown(
    """
    **Criterios de seguimiento**

    🟢 Candidatos: ≥75 % ·
    🟠 Próximos: ≥60 % y <75 % ·
    ⚪ Fuera de proyección: <60 %
    """
)


# ==========================================================
# 7. CARGA DEL ARCHIVO EXCEL
# ==========================================================

st.sidebar.header("Fuente de datos")

archivo_cargado = st.sidebar.file_uploader(
    "Cargar archivo Excel",
    type=["xlsx", "xls"]
)

if archivo_cargado is not None:
    bytes_excel = archivo_cargado.getvalue()
    nombre_archivo = archivo_cargado.name

elif ARCHIVO_BASE.exists():
    bytes_excel = ARCHIVO_BASE.read_bytes()
    nombre_archivo = ARCHIVO_BASE.name

else:
    st.info(
        "No se encontró PROYECCIÓN.xlsx junto a app.py. "
        "Carga el archivo desde el panel izquierdo."
    )
    st.stop()

try:
    libro = pd.ExcelFile(BytesIO(bytes_excel))

    hoja = st.sidebar.selectbox(
        "Hoja de Excel",
        options=libro.sheet_names
    )

    base_original = pd.read_excel(
        BytesIO(bytes_excel),
        sheet_name=hoja,
        dtype=str
    )

    df_completo = preparar_base(base_original)

except Exception as e:
    st.error(f"No se pudo preparar la base de datos: {e}")
    st.stop()

st.sidebar.success("Archivo cargado correctamente")
st.sidebar.caption(f"Fuente: {nombre_archivo}")


# ==========================================================
# 8. FILTROS
# ==========================================================

st.sidebar.divider()
st.sidebar.header("Filtros del análisis")


def opciones_columna(df, columna):
    return sorted(
        df[columna].dropna().astype(str).unique().tolist()
    )


def seleccionar_filtro(df, columna, etiqueta):

    opciones = opciones_columna(df, columna)

    return st.sidebar.multiselect(
        etiqueta,
        options=opciones,
        default=opciones,
        key=f"filtro_{columna}"
    )


sedes = seleccionar_filtro(
    df_completo, "SEDE", "Sede / CU"
)

facultades = seleccionar_filtro(
    df_completo, "FACULTAD", "Facultad"
)

niveles = seleccionar_filtro(
    df_completo, "NIVEL", "Nivel académico"
)

programas = seleccionar_filtro(
    df_completo, "PROGRAMA", "Programa"
)

pruebas = seleccionar_filtro(
    df_completo, "PRUEBA_ORIENTATIVA", "Prueba orientativa"
)

categorias = st.sidebar.multiselect(
    "Categoría de avance",
    options=CATEGORIAS,
    default=CATEGORIAS
)

df = df_completo[
    df_completo["SEDE"].isin(sedes)
    & df_completo["FACULTAD"].isin(facultades)
    & df_completo["NIVEL"].isin(niveles)
    & df_completo["PROGRAMA"].isin(programas)
    & df_completo["PRUEBA_ORIENTATIVA"].isin(pruebas)
    & df_completo["CATEGORIA"].isin(categorias)
].copy()

if df.empty:
    st.warning("No hay registros para los filtros seleccionados.")
    st.stop()


# ==========================================================
# 9. INDICADORES GENERALES
# ==========================================================

# Los indicadores cuentan estudiantes únicos.
total = df["ID"].nunique()

candidatos = df.loc[
    df["CATEGORIA"] == "Candidato", "ID"
].nunique()

proximos = df.loc[
    df["CATEGORIA"] == "Próximo", "ID"
].nunique()

fuera = df.loc[
    df["CATEGORIA"] == "Fuera de proyección", "ID"
].nunique()

sin_info = df.loc[
    df["CATEGORIA"] == "Sin información", "ID"
].nunique()

en_proyeccion = candidatos + proximos

avance_promedio = (
    df.drop_duplicates("ID")["PORCENTAJE_ANALISIS"].mean()
)

st.subheader("Resumen general")

k1, k2, k3 = st.columns(3)

k1.metric(
    "Estudiantes únicos",
    f"{total:,}"
)

k2.metric(
    "Candidatos · ≥75 %",
    f"{candidatos:,}"
)

k3.metric(
    "Próximos · 60 % a <75 %",
    f"{proximos:,}"
)

k4, k5, k6 = st.columns(3)

k4.metric(
    "En proyección inmediata",
    f"{en_proyeccion:,}"
)

k5.metric(
    "Fuera de proyección · <60 %",
    f"{fuera:,}"
)

k6.metric(
    "Avance académico promedio",
    f"{avance_promedio:.1f} %"
    if pd.notna(avance_promedio) else "N/D"
)

st.caption(
    f"Estudiantes sin información de avance: {sin_info:,}. "
    "En proyección = candidatos + próximos."
)


# ==========================================================
# 10. PESTAÑAS DEL DASHBOARD
# ==========================================================

tab_general, tab_sede, tab_facultad, tab_programa, tab_anios, tab_datos = st.tabs([
    "Resumen general",
    "Por sede / CU",
    "Por facultad",
    "Por programa",
    "Proyección 2026–2027",
    "Datos y exportación"
])


# ==========================================================
# PESTAÑA 1. RESUMEN GENERAL
# ==========================================================

with tab_general:

    izquierda, derecha = st.columns(2)

    with izquierda:

        st.subheader("Distribución por avance")

        distribucion = (
            df.groupby("CATEGORIA")["ID"]
            .nunique()
            .reindex(CATEGORIAS, fill_value=0)
            .rename_axis("Categoría")
            .reset_index(name="Estudiantes")
        )

        fig = px.pie(
            distribucion,
            names="Categoría",
            values="Estudiantes",
            hole=0.55,
            color="Categoría",
            color_discrete_map=COLORES,
            template="plotly_white"
        )

        fig.update_traces(
            textposition="inside",
            textinfo="percent+label"
        )

        fig.update_layout(
            showlegend=False,
            margin=dict(l=10, r=10, t=25, b=10)
        )

        st.plotly_chart(fig, use_container_width=True)

    with derecha:

        st.subheader("Población por tipo de prueba")

        prueba_resumen = resumen_por_grupo(
            df, ["PRUEBA_ORIENTATIVA"]
        )

        prueba_resumen = prueba_resumen.rename(
            columns={"PRUEBA_ORIENTATIVA": "Prueba"}
        )

        columnas_grafico = [
            "Prueba",
            "Candidato",
            "Próximo",
            "Fuera de proyección",
            "Sin información"
        ]

        datos_grafico = prueba_resumen[
            columnas_grafico
        ].melt(
            id_vars="Prueba",
            var_name="Categoría",
            value_name="Estudiantes"
        )

        fig = px.bar(
            datos_grafico,
            x="Prueba",
            y="Estudiantes",
            color="Categoría",
            barmode="stack",
            color_discrete_map=COLORES,
            template="plotly_white"
        )

        fig.update_layout(
            legend_title_text="Clasificación",
            margin=dict(l=10, r=10, t=25, b=10)
        )

        st.plotly_chart(fig, use_container_width=True)

    st.subheader("Resumen de clasificación")

    resumen_clasificacion = pd.DataFrame({
        "Categoría": CATEGORIAS,
        "Estudiantes únicos": [
            df.loc[df["CATEGORIA"] == c, "ID"].nunique()
            for c in CATEGORIAS
        ]
    })

    st.dataframe(
        resumen_clasificacion,
        use_container_width=True,
        hide_index=True
    )


# ==========================================================
# PESTAÑA 2. ANÁLISIS POR SEDE
# ==========================================================

with tab_sede:

    st.subheader("Comparativo por sede / CU")

    resumen_sede = formato_tabla(
        resumen_por_grupo(df, ["SEDE"])
    )

    resumen_sede = resumen_sede.sort_values(
        "En proyección", ascending=False
    )

    fig = grafico_barras_categoria(
        resumen_sede,
        "SEDE",
        "Candidatos y próximos por sede"
    )

    st.plotly_chart(fig, use_container_width=True)

    st.dataframe(
        resumen_sede,
        use_container_width=True,
        hide_index=True
    )


# ==========================================================
# PESTAÑA 3. ANÁLISIS POR FACULTAD
# ==========================================================

with tab_facultad:

    st.subheader("Población por facultad")

    resumen_facultad = formato_tabla(
        resumen_por_grupo(df, ["FACULTAD"])
    )

    resumen_facultad = resumen_facultad.sort_values(
        "En proyección", ascending=False
    )

    fig = px.bar(
        resumen_facultad,
        x="FACULTAD",
        y="En proyección",
        color="FACULTAD",
        template="plotly_white",
        title="Estudiantes en proyección inmediata por facultad",
        labels={
            "En proyección": "Candidatos + próximos",
            "FACULTAD": "Facultad"
        }
    )

    fig.update_layout(
        showlegend=False,
        margin=dict(l=15, r=15, t=60, b=20)
    )

    st.plotly_chart(fig, use_container_width=True)

    st.dataframe(
        resumen_facultad,
        use_container_width=True,
        hide_index=True
    )


# ==========================================================
# PESTAÑA 4. DETALLE POR PROGRAMA
# ==========================================================

with tab_programa:

    st.subheader("Detalle por facultad, código y programa")

    resumen_programa = formato_tabla(
        resumen_por_grupo(
            df,
            ["FACULTAD", "COD", "PROGRAMA", "NIVEL"]
        )
    )

    resumen_programa = resumen_programa.sort_values(
        ["FACULTAD", "En proyección"],
        ascending=[True, False]
    )

    st.dataframe(
        resumen_programa,
        use_container_width=True,
        hide_index=True
    )

    st.subheader("Programas con mayor población en proyección")

    top_programas = resumen_programa.sort_values(
        "En proyección", ascending=False
    ).head(20).copy()

    top_programas["Etiqueta"] = (
        top_programas["COD"].astype(str)
        + " | "
        + top_programas["PROGRAMA"].astype(str)
    )

    fig = px.bar(
        top_programas,
        x="En proyección",
        y="Etiqueta",
        color="NIVEL",
        orientation="h",
        template="plotly_white",
        hover_data=[
            "FACULTAD",
            "Estudiantes únicos",
            "Candidato",
            "Próximo"
        ],
        title="Top 20 programas"
    )

    fig.update_layout(
        yaxis={"categoryorder": "total ascending"},
        xaxis_title="Candidatos + próximos",
        yaxis_title="Programa",
        margin=dict(l=15, r=15, t=60, b=20)
    )

    st.plotly_chart(fig, use_container_width=True)

    st.download_button(
        "Descargar resumen por programa",
        data=exportar_excel({
            "Por programa": resumen_programa
        }),
        file_name="proyeccion_por_programa.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


# ==========================================================
# PESTAÑA 5. PROYECCIÓN 2026 Y 2027
# ==========================================================

with tab_anios:

    st.subheader("Análisis por año de proyección")

    anio = st.radio(
        "Selecciona el año",
        ["2026", "2027"],
        horizontal=True,
        format_func=lambda x: (
            "📘 Año 2026" if x == "2026"
            else "📗 Año 2027"
        )
    )

    st.markdown(f"### Proyección académica para {anio}")

    df_anio = df[df[anio].notna()].copy()

    if df_anio.empty:

        st.info(
            f"No hay semestres registrados para {anio}."
        )

    else:

        df_anio[anio] = df_anio[anio].astype(int)

        total_anio = df_anio["ID"].nunique()

        candidatos_anio = df_anio.loc[
            df_anio["CATEGORIA"] == "Candidato", "ID"
        ].nunique()

        proximos_anio = df_anio.loc[
            df_anio["CATEGORIA"] == "Próximo", "ID"
        ].nunique()

        a1, a2, a3 = st.columns(3)

        a1.metric(
            f"Estudiantes con semestre {anio}",
            f"{total_anio:,}"
        )

        a2.metric(
            f"Candidatos {anio}",
            f"{candidatos_anio:,}"
        )

        a3.metric(
            f"Próximos {anio}",
            f"{proximos_anio:,}"
        )

        st.caption(
            "La categoría de avance se basa en los créditos de la fuente. "
            "Seleccionar 2027 no incrementa automáticamente el porcentaje "
            "de avance académico."
        )

        resumen_anio = resumen_por_grupo(
            df_anio, [anio]
        ).sort_values(anio)

        fig = grafico_barras_categoria(
            resumen_anio,
            anio,
            f"Candidatos y próximos por semestre — {anio}"
        )

        st.plotly_chart(fig, use_container_width=True)

        st.subheader(f"Detalle de semestres — {anio}")

        st.dataframe(
            formato_tabla(resumen_anio),
            use_container_width=True,
            hide_index=True
        )

        st.download_button(
            f"Descargar proyección {anio}",
            data=exportar_excel({
                f"Proyeccion {anio}": resumen_anio
            }),
            file_name=f"proyeccion_{anio}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )


# ==========================================================
# PESTAÑA 6. DATOS Y EXPORTACIÓN
# ==========================================================

with tab_datos:

    st.subheader("Base académica filtrada")

    st.caption(
        "Los campos personales se excluyen de las tablas y "
        "exportaciones para reducir la exposición de información sensible."
    )

    columnas_privadas = [
        "NOMBRE", "DOCUMENTO", "CORREO",
        "EMAIL", "TELEFONO", "CELULAR"
    ]

    columnas_visibles = [
        c for c in df.columns
        if c not in columnas_privadas
    ]

    base_visible = df[columnas_visibles].copy()

    st.dataframe(
        base_visible,
        use_container_width=True,
        hide_index=True
    )

    st.download_button(
        "Descargar base filtrada en Excel",
        data=exportar_excel({
            "Base filtrada": base_visible
        }),
        file_name="proyeccion_base_filtrada.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    st.subheader("Resumen consolidado")

    resumen_consolidado = formato_tabla(
        resumen_por_grupo(
            df,
            [
                "SEDE", "FACULTAD", "COD",
                "PROGRAMA", "NIVEL"
            ]
        )
    )

    st.dataframe(
        resumen_consolidado,
        use_container_width=True,
        hide_index=True
    )

    st.download_button(
        "Descargar resumen consolidado",
        data=exportar_excel({
            "Resumen": resumen_consolidado,
            "Por sede": formato_tabla(
                resumen_por_grupo(df, ["SEDE"])
            ),
            "Por facultad": formato_tabla(
                resumen_por_grupo(df, ["FACULTAD"])
            ),
            "Por programa": formato_tabla(
                resumen_por_grupo(
                    df,
                    ["FACULTAD", "COD", "PROGRAMA", "NIVEL"]
                )
            )
        }),
        file_name="resumen_proyeccion_academica.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


# ==========================================================
# 11. NOTA FINAL
# ==========================================================

st.divider()

st.caption(
    "La clasificación entre Saber Pro y Saber TyT es orientativa "
    "según el nivel académico. La elegibilidad debe verificarse con "
    "los criterios vigentes del ICFES. Los umbrales de avance son "
    "criterios de seguimiento definidos para este análisis."
)
