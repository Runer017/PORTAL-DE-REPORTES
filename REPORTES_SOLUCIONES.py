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
    "ADMINISTRACIÓN": "X",
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

    import pandas as pd

    # ======================================================
    # CARGAR EXCEL
    # ======================================================

    archivo_excel = "STATUS_ROBOTS.xlsx"

    try:

        df_robots = pd.read_excel(
            archivo_excel,
            sheet_name="REPORTE AGO"
        )

    except Exception as e:

        st.error(
            f"No se pudo cargar la información de robots: {e}"
        )
        return


    # ======================================================
    # LIMPIAR NOMBRES DE COLUMNAS
    # ======================================================

    df_robots.columns = (
        df_robots.columns
        .str.strip()
        .str.upper()
    )


    # ======================================================
    # ÁREA DEL USUARIO
    # ======================================================

    area_usuario = st.session_state.area.upper().strip()


    # ======================================================
    # FILTRAR POR ÁREA
    # ======================================================

    df_area = df_robots[
        df_robots["AREA"]
        .astype(str)
        .str.upper()
        .str.strip()
        == area_usuario
    ].copy()


    # ======================================================
    # TÍTULO
    # ======================================================

    st.title("Status de Soluciones")


    st.markdown(
        f"""
        <div style="
            color:#6B7280;
            font-size:15px;
            margin-top:-10px;
            margin-bottom:25px;
        ">
            Estado de las soluciones digitales del área
            <strong>{st.session_state.area}</strong>
        </div>
        """,
        unsafe_allow_html=True
    )


    # ======================================================
    # SUBTÍTULO
    # ======================================================

    st.subheader("Automatizaciones")


    # ======================================================
    # PREPARAR DATOS
    # ======================================================

    if df_area.empty:

        st.info(
            f"No se encontraron robots registrados para el área "
            f"{st.session_state.area}."
        )

    else:

        # --------------------------------------------------
        # REEMPLAZAR VACÍOS POR 0
        # --------------------------------------------------

        columnas_estado = [
            "DETENIDOS",
            "EXITOSOS",
            "ERROR"
        ]

        for columna in columnas_estado:

            if columna in df_area.columns:

                df_area[columna] = (
                    pd.to_numeric(
                        df_area[columna],
                        errors="coerce"
                    )
                    .fillna(0)
                    .astype(int)
                )


        # --------------------------------------------------
        # OBTENER FECHA MÁS RECIENTE
        # --------------------------------------------------

        df_area["FECHA"] = pd.to_datetime(
            df_area["FECHA"],
            errors="coerce"
        )

        fecha_maxima = df_area["FECHA"].max()


        # --------------------------------------------------
        # MOSTRAR SOLO EL ÚLTIMO REPORTE
        # --------------------------------------------------

        df_actual = df_area[
            df_area["FECHA"] == fecha_maxima
        ].copy()


        # ==================================================
        # CALCULAR ESTADO
        # ==================================================

        def obtener_estado(row):

            if row["ERROR"] > 0:

                return "🔴 ERROR"

            elif row["DETENIDOS"] > 0:

                return "🟡 DETENIDO"

            elif row["EXITOSOS"] > 0:

                return "🟢 OK"

            else:

                return "⚪ SIN DATOS"


        df_actual["ESTADO"] = df_actual.apply(
            obtener_estado,
            axis=1
        )


        # ==================================================
        # INDICADORES
        # ==================================================

        total_ok = (
            df_actual["ESTADO"] == "🟢 OK"
        ).sum()

        total_error = (
            df_actual["ESTADO"] == "🔴 ERROR"
        ).sum()

        total_detenidos = (
            df_actual["ESTADO"] == "🟡 DETENIDO"
        ).sum()


        col1, col2, col3 = st.columns(3)


        with col1:

            st.metric(
                "🟢 Exitosos",
                total_ok
            )


        with col2:

            st.metric(
                "🔴 Con error",
                total_error
            )


        with col3:

            st.metric(
                "🟡 Detenidos",
                total_detenidos
            )


        st.write("")


        # ==================================================
        # TABLA DE AUTOMATIZACIONES
        # ==================================================

        tabla = df_actual[
            [
                "ROBOT",
                "AREA",
                "FECHA",
                "DETENIDOS",
                "EXITOSOS",
                "ERROR",
                "ESTADO"
            ]
        ].copy()


        # --------------------------------------------------
        # FORMATO FECHA
        # --------------------------------------------------

        tabla["FECHA"] = tabla["FECHA"].dt.strftime(
            "%d/%m/%Y"
        )


        # ==================================================
        # MOSTRAR TABLA
        # ==================================================

        st.dataframe(
            tabla,
            use_container_width=True,
            hide_index=True,
            column_config={

                "ROBOT": st.column_config.TextColumn(
                    "Robot"
                ),

                "AREA": st.column_config.TextColumn(
                    "Área"
                ),

                "FECHA": st.column_config.TextColumn(
                    "Último reporte"
                ),

                "DETENIDOS": st.column_config.NumberColumn(
                    "Detenidos"
                ),

                "EXITOSOS": st.column_config.NumberColumn(
                    "Exitosos"
                ),

                "ERROR": st.column_config.NumberColumn(
                    "Errores"
                ),

                "ESTADO": st.column_config.TextColumn(
                    "Estado"
                )
            }
        )


    # ======================================================
    # CERRAR SESIÓN
    # ======================================================

    st.write("")

    if st.button("Cerrar sesión"):

        st.session_state.logueado = False
        st.session_state.area = None

        st.rerun()