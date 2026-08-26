import streamlit as st
from textwrap import dedent
import pandas as pd
import matplotlib.pyplot as plt
from textblob import TextBlob
import plotly.express as px
import plotly.graph_objects as go

# ==========================================================
# CONFIGURACIÓN DE LA PÁGINA
# ==========================================================

st.set_page_config(
    page_title="Reporte de Soluciones",
    page_icon="📊",
    layout="wide"
)


# ==========================================================
# USUARIOS / ÁREAS
# ==========================================================

USUARIOS = {
    "CONTABILIDAD": "X",
    "ADMINISTRACION": "X",
    "OR": "X",
    "SSOMA": "X"
}


# ==========================================================
# INICIALIZAR SESIÓN
# ==========================================================

if "logueado" not in st.session_state:
    st.session_state.logueado = False

if "area" not in st.session_state:
    st.session_state.area = None


# ==========================================================
# ESTILOS
# ==========================================================

st.markdown("""
<style>
    
/* ==========================================================
   FONDO
========================================================== */

.stApp {
    background: linear-gradient(
        135deg,
        #EEF4FF 0%,
        #F7F3FF 50%,
        #ECFBFF 100%
    );
}


/* ==========================================================
   OCULTAR MENÚ
========================================================== */

#MainMenu {
    visibility: hidden;
}

footer {
    visibility: hidden;
}


/* ==========================================================
   BOTÓN
========================================================== */

.stButton > button {
    width: 100%;
    height: 48px;
    border-radius: 10px;
    border: none;
    background: linear-gradient(
        90deg,
        #5B5FEF,
        #7C5CFC
    );
    color: white;
    font-size: 15px;
    font-weight: 600;
}

.stButton > button:hover {
    background: linear-gradient(
        90deg,
        #4F46E5,
        #6D28D9
    );
    color: white;
}


/* ==========================================================
   INPUTS
========================================================== */

.stSelectbox > div > div,
.stTextInput > div > div {
    border-radius: 10px;
}


/* ==========================================================
   LABELS
========================================================== */

.stSelectbox label,
.stTextInput label {
    font-weight: 500;
    color: #374151;
}


/* ==========================================================
   ALERTAS
========================================================== */

.stAlert {
    border-radius: 10px;
}

</style>
""", unsafe_allow_html=True)


# ==========================================================
# PANTALLA DE LOGIN
# ==========================================================

def pantalla_login():

     # ======================================================
    # CABECERA
    # ======================================================

    st.markdown("""
    <div style="width:100%; text-align:center; margin-top:15px; margin-bottom:20px;">

    <div style="font-size:42px; margin-bottom:6px;">
        🤖
    </div>

    <div style="font-size:32px; font-weight:700; color:#252B48; margin-bottom:8px;">
        ¡Bienvenido a nuestro portal!
    </div>

    <div style="font-size:14px; font-weight:700; letter-spacing:1px; color:#5B5FEF; margin-bottom:8px;">
        DIGITAL INNOVATION - TIC
    </div>

    <div style="font-size:14px; color:#6B7280; line-height:1.4;">
        Ingresa para consultar nuestros reportes de soluciones digitales
    </div>

    </div>
    """, unsafe_allow_html=True)


    # ======================================================
    # ESPACIO
    # ======================================================

    st.write("")


    # ======================================================
    # CAMPO ÁREA
    # ======================================================

    area = st.selectbox(
        "Área",
        options=list(USUARIOS.keys()),
        index=None,
        placeholder="Selecciona tu área"
    )


    # ======================================================
    # CONTRASEÑA
    # ======================================================

    contraseña = st.text_input(
        "Contraseña",
        type="password",
        placeholder="Ingresa tu contraseña"
    )


    # ======================================================
    # ESPACIO
    # ======================================================

    st.write("")


    # ======================================================
    # BOTÓN
    # ======================================================

    ingresar = st.button(
        "Ingresar al portal →",
        use_container_width=True
    )


    # ======================================================
    # VALIDACIÓN
    # ======================================================

    if ingresar:

        if area is None:

            st.warning(
                "Selecciona tu área para continuar."
            )

        elif contraseña == USUARIOS[area]:

            st.session_state.logueado = True
            st.session_state.area = area

            st.rerun()

        else:

            st.error(
                "La contraseña ingresada es incorrecta."
            )


    # ======================================================
    # PIE
    # ======================================================

    st.write("")

    st.markdown(
        """
        <div style="
            width: 100%;
            text-align: center;
            color: #9CA3AF;
            font-size: 12px;
            margin-top: 10px;
        ">
            Portal de Reportes · Digital Innovation - TIC
        </div>
        """,
        unsafe_allow_html=True
    )
    

# CONFIGURACIÓN DE NAVEGACIÓN
# ==========================================================

if "pantalla" not in st.session_state:
    st.session_state.pantalla = "inicio"


# ==========================================================
# PANTALLA PRINCIPAL
# ==========================================================

def pantalla_principal():

    # ======================================================
    # CERRAR SESIÓN
    # ======================================================
    
    if st.button("< Cerrar sesión"):
    
        st.session_state.logueado = False
        st.session_state.area = None
    
        st.rerun()

    # ======================================================
    # TÍTULO
    # ======================================================

    st.title("Status de Soluciones")

    st.markdown(
        f"""
        <p style="
            color:#6B7280;
            font-size:15px;
            margin-top:-10px;
            margin-bottom:25px;
        ">
            Estado de las automatizaciones del área
            <strong>{st.session_state.area}</strong>
        </p>
        """,
        unsafe_allow_html=True
    )


    # ======================================================
    # LEER EXCEL
    # ======================================================

    try:

        archivo = "STATUS_ROBOTS.xlsx"

        tabla = pd.read_excel(
            archivo,
            sheet_name="REPORTE AGO"
        )

    except Exception as e:

        st.error(
            f"No se pudo cargar la hoja REPORTE AGO: {e}"
        )

        return


    # ======================================================
    # FILTRAR POR ÁREA DEL USUARIO
    # ======================================================

    area_usuario = st.session_state.area

    tabla_area = tabla[
        tabla["AREA"]
        .astype(str)
        .str.strip()
        .str.upper()
        ==
        area_usuario.strip().upper()
    ].copy()


    # ======================================================
    # VERIFICAR SI EXISTEN DATOS
    # ======================================================

    if tabla_area.empty:

        st.info(
            f"No se encontraron automatizaciones "
            f"registradas para el área {area_usuario}."
        )

        return


    # ======================================================
    # LIMPIAR COLUMNAS DE EJECUCIONES
    # ======================================================

    columnas_ejecuciones = [
        "DETENIDOS",
        "EXITOSOS",
        "ERROR"
    ]

    for columna in columnas_ejecuciones:

        if columna in tabla_area.columns:

            tabla_area[columna] = pd.to_numeric(
                tabla_area[columna],
                errors="coerce"
            ).fillna(0)


    # ======================================================
    # INDICADORES
    # ======================================================

    # ------------------------------------------------------
    # 1. NÚMERO DE ROBOTS
    # ------------------------------------------------------

    numero_robots = tabla_area["ROBOT"].nunique()


    # ------------------------------------------------------
    # 2. EJECUCIONES EXITOSAS
    # ------------------------------------------------------

    ejecuciones_exitosas = int(
        tabla_area["EXITOSOS"].sum()
    )


    # ------------------------------------------------------
    # 3. EJECUCIONES DETENIDAS
    # ------------------------------------------------------

    ejecuciones_detenidas = int(
        tabla_area["DETENIDOS"].sum()
    )


    # ------------------------------------------------------
    # 4. EJECUCIONES FALLIDAS
    # ------------------------------------------------------

    ejecuciones_fallidas = int(
        tabla_area["ERROR"].sum()
    )


    # ======================================================
    # SUBTÍTULO
    # ======================================================

    st.subheader("Automatizaciones")


    # ======================================================
    # ESTILO DE INDICADORES
    # ======================================================

    st.html("""
    <style>

    .metric-container {
        display: flex;
        gap: 20px;
        width: 100%;
        margin-bottom: 30px;
    }

    .metric-card {
        flex: 1;
        background: white;
        border: 1px solid #E0E0E0;
        border-radius: 18px;
        padding: 24px 28px;
        min-height: 150px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
    }

    .metric-title {
        font-size: 18px;
        color: #555555;
        margin-bottom: 10px;
    }

    .metric-value {
        font-size: 42px;
        font-weight: 600;
        color: #111111;
        line-height: 1.1;
    }

    .metric-subtitle {
        font-size: 16px;
        color: #666666;
        margin-top: 10px;
    }

    </style>
    """)


    # ======================================================
    # MOSTRAR LOS 4 INDICADORES
    # ======================================================

    st.html(f"""
    <div class="metric-container">

        <!-- INDICADOR 1 -->
        <div class="metric-card">

            <div class="metric-title">
                N.° de robots
            </div>

            <div class="metric-value">
                {numero_robots}
            </div>

            <div class="metric-subtitle">
                robots registrados en el área
            </div>

        </div>


        <!-- INDICADOR 2 -->
        <div class="metric-card">

            <div class="metric-title">
                Ejecuciones exitosas
            </div>

            <div class="metric-value">
                {ejecuciones_exitosas}
            </div>

            <div class="metric-subtitle">
                ejecuciones completadas
            </div>

        </div>


        <!-- INDICADOR 3 -->
        <div class="metric-card">

            <div class="metric-title">
                Ejecuciones detenidas
            </div>

            <div class="metric-value">
                {ejecuciones_detenidas}
            </div>

            <div class="metric-subtitle">
                ejecuciones detenidas
            </div>

        </div>


        <!-- INDICADOR 4 -->
        <div class="metric-card">

            <div class="metric-title">
                Ejecuciones fallidas
            </div>

            <div class="metric-value">
                {ejecuciones_fallidas}
            </div>

            <div class="metric-subtitle">
                ejecuciones con error
            </div>

        </div>

    </div>
    """)


    # ======================================================
    # GRÁFICA DE EJECUCIONES POR ROBOT
    # ======================================================

    st.subheader("Ejecuciones por robot")


    # ======================================================
    # AGRUPAR ROBOTS
    # ======================================================

    grafica = (
        tabla_area
        .groupby("ROBOT")[
            [
                "EXITOSOS",
                "DETENIDOS",
                "ERROR"
            ]
        ]
        .sum()
        .reset_index()
    )


    # ======================================================
    # TOTAL DE EJECUCIONES POR ROBOT
    # ======================================================

    grafica["TOTAL"] = (
        grafica["EXITOSOS"]
        + grafica["DETENIDOS"]
        + grafica["ERROR"]
    )


    # ======================================================
    # CREAR GRÁFICA
    # ======================================================

    fig = go.Figure()


    # ======================================================
    # EXITOSAS
    # ======================================================

    fig.add_trace(
        go.Bar(
            x=grafica["ROBOT"],
            y=grafica["EXITOSOS"],
            name="Exitosas",
            marker_color="#0875D1",

            hovertemplate=(
                "<b>Exitosas</b><br>"
                "Ejecuciones: %{y}"
                "<extra></extra>"
            )
        )
    )


    # ======================================================
    # DETENIDAS
    # ======================================================

    fig.add_trace(
        go.Bar(
            x=grafica["ROBOT"],
            y=grafica["DETENIDOS"],
            name="Detenidas",
            marker_color="#75B9E8",

            hovertemplate=(
                "<b>Detenidas</b><br>"
                "Ejecuciones: %{y}"
                "<extra></extra>"
            )
        )
    )


    # ======================================================
    # FALLIDAS
    # ======================================================

    fig.add_trace(
        go.Bar(
            x=grafica["ROBOT"],
            y=grafica["ERROR"],
            name="Fallidas",
            marker_color="#FF2B32",

            hovertemplate=(
                "<b>Fallidas</b><br>"
                "Ejecuciones: %{y}"
                "<extra></extra>"
            )
        )
    )


    # ======================================================
    # TOTAL ENCIMA DE CADA BARRA
    # ======================================================

    fig.add_trace(
        go.Scatter(

            x=grafica["ROBOT"],

            y=grafica["TOTAL"],

            mode="text",

            text=(
                grafica["TOTAL"]
                .astype(int)
                .astype(str)
            ),

            textposition="top center",

            textfont=dict(
                size=14,
                color="#374151"
            ),

            hoverinfo="skip",

            showlegend=False
        )
    )


    # ======================================================
    # CONFIGURACIÓN DE LA GRÁFICA
    # ======================================================

    fig.update_layout(

        barmode="stack",

        height=450,

        margin=dict(
            l=20,
            r=20,
            t=30,
            b=120
        ),

        xaxis=dict(
            title=None,
            tickangle=-35
        ),

        yaxis=dict(
            title="N.° de ejecuciones",
            rangemode="tozero"
        ),

        legend=dict(
            title="Resultado",
            orientation="h",
            yanchor="top",
            y=-0.25,
            xanchor="center",
            x=0.5
        ),

        hovermode="closest"
    )


    # ======================================================
    # MOSTRAR GRÁFICA
    # ======================================================

    st.plotly_chart(
        fig,
        use_container_width=True
    )


    # ======================================================
    # TABLA
    # ======================================================

    st.subheader("Detalle de ejecuciones")


    # ------------------------------------------------------
    # COLUMNAS QUE SE MOSTRARÁN
    # ------------------------------------------------------

    columnas = [
        "ROBOT",
        "FECHA",
        "DETENIDOS",
        "EXITOSOS",
        "ERROR"
    ]

    columnas_disponibles = [
        col
        for col in columnas
        if col in tabla_area.columns
    ]

    tabla_mostrar = tabla_area[
        columnas_disponibles
    ].copy()


    # ------------------------------------------------------
    # ASEGURAR QUE LAS EJECUCIONES SEAN ENTEROS
    # ------------------------------------------------------

    columnas_ejecuciones = [
        "DETENIDOS",
        "EXITOSOS",
        "ERROR"
    ]

    for columna in columnas_ejecuciones:

        if columna in tabla_mostrar.columns:

            tabla_mostrar[columna] = pd.to_numeric(
                tabla_mostrar[columna],
                errors="coerce"
            ).fillna(0).astype(int)


    # ------------------------------------------------------
    # FORMATO DE FECHA
    # ------------------------------------------------------

    if "FECHA" in tabla_mostrar.columns:

        tabla_mostrar["FECHA"] = pd.to_datetime(
            tabla_mostrar["FECHA"],
            errors="coerce"
        ).dt.strftime("%d/%m/%Y")


    # ======================================================
    # ESTILO DE LA TABLA
    # ======================================================

    def colorear_detenidos(valor):

        if valor > 0:
            return (
                "color: #F59E0B; "
                "font-weight: 600;"
            )

        return "color: #9CA3AF;"


    def colorear_exitosos(valor):

        if valor > 0:
            return (
                "color: #16A34A; "
                "font-weight: 600;"
            )

        return "color: #9CA3AF;"


    def colorear_fallidos(valor):

        if valor > 0:
            return (
                "color: #DC2626; "
                "font-weight: 600;"
            )

        return "color: #9CA3AF;"


    # ======================================================
    # APLICAR ESTILOS
    # ======================================================

    tabla_estilizada = (
        tabla_mostrar
        .style
        .map(
            colorear_detenidos,
            subset=["DETENIDOS"]
        )
        .map(
            colorear_exitosos,
            subset=["EXITOSOS"]
        )
        .map(
            colorear_fallidos,
            subset=["ERROR"]
        )
    )


    # ======================================================
    # MOSTRAR TABLA
    # ======================================================

    st.dataframe(
        tabla_estilizada,
        use_container_width=True,
        hide_index=True,
        height=500
    )
    

        
# ==========================================================
# CONTROL DE PANTALLAS
# ==========================================================

if st.session_state.logueado:
    pantalla_principal()
else:
    pantalla_login()