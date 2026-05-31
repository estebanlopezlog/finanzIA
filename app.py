# ============================================================
# FinanzIA - Gestor Inteligente de Finanzas Personales
# Proyecto Final - Prompt Engineering para Programadores
# Diplomatura en Data Science - CoderHouse
# ============================================================

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import google.generativeai as genai
from datetime import datetime
import io

# ============================================================
# CONFIGURACIÓN DE LA PÁGINA
# ============================================================
st.set_page_config(
    page_title="FinanzIA - Finanzas Personales",
    page_icon="💰",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ============================================================
# ESTILOS CSS PERSONALIZADOS (Header, Footer, paleta de colores)
# ============================================================
st.markdown("""
<style>
    /* Header principal */
    .main-header {
        background: linear-gradient(135deg, #0f2027, #203a43, #2c5364);
        padding: 25px 30px;
        border-radius: 12px;
        text-align: center;
        margin-bottom: 25px;
        box-shadow: 0 4px 15px rgba(0,0,0,0.3);
    }
    .main-header h1 { color: #f0c040; margin: 0; font-size: 2.2rem; }
    .main-header p  { color: #b0c4de; margin: 6px 0 0 0; font-size: 1rem; }

    /* Tarjetas de métricas */
    .metric-box {
        background: #1e2d40;
        border-radius: 10px;
        padding: 16px;
        text-align: center;
        color: white;
    }

    /* Footer */
    .footer {
        text-align: center;
        color: #888;
        padding: 18px;
        border-top: 1px solid #333;
        margin-top: 40px;
        font-size: 0.85rem;
    }

    /* Tabla de datos más compacta */
    .dataframe { font-size: 0.85rem; }

    /* Chips de categoría */
    .chip {
        display: inline-block;
        background: #2c5364;
        color: #f0c040;
        border-radius: 20px;
        padding: 3px 12px;
        margin: 3px;
        font-size: 0.82rem;
    }

    /* Barra de límite de gasto */
    .limite-ok      { color: #4CAF50; font-weight: bold; }
    .limite-warning { color: #FF9800; font-weight: bold; }
    .limite-danger  { color: #f44336; font-weight: bold; }

    /* Tarjeta de alerta de límite */
    .alerta-warning {
        background: #3a2800;
        border-left: 4px solid #FF9800;
        border-radius: 8px;
        padding: 12px 16px;
        margin: 6px 0;
    }
    .alerta-danger {
        background: #3a0000;
        border-left: 4px solid #f44336;
        border-radius: 8px;
        padding: 12px 16px;
        margin: 6px 0;
    }
</style>
""", unsafe_allow_html=True)


# ============================================================
# FUNCIONES DE CARGA Y PROCESAMIENTO DE DATOS
# ============================================================

def cargar_gastos(xl: pd.ExcelFile) -> pd.DataFrame:
    """
    Carga la hoja 'Gastos' del Excel y normaliza su estructura.
    El archivo tiene un título en la fila 0 ('Registro de compras')
    y los nombres de columna reales en la fila 1.
    """
    raw = xl.parse("Gastos", header=None)
    # La fila 1 contiene los nombres de columnas reales (fila 0 es el título)
    raw.columns = raw.iloc[1].tolist()
    raw = raw.iloc[2:].reset_index(drop=True)

    # Renombrar columnas clave al formato estándar interno
    rename_map = {}
    for col in raw.columns:
        col_lower = str(col).lower().strip()
        if "fecha" in col_lower and "año" not in col_lower and "mes" not in col_lower:
            rename_map[col] = "Fecha"
        elif col_lower == "año_mes" or col_lower == "año mes":
            rename_map[col] = "Anio_Mes"
        elif col_lower == "mes":
            rename_map[col] = "Mes"
        elif col_lower == "year" or col_lower == "año":
            rename_map[col] = "Anio"
        elif "concepto" in col_lower:
            rename_map[col] = "Concepto"
        elif "monto" in col_lower or "importe" in col_lower or "amount" in col_lower:
            rename_map[col] = "Monto"
        elif "medio" in col_lower and "pago" in col_lower:
            rename_map[col] = "Medio_Pago"
        elif "moneda" in col_lower or "currency" in col_lower:
            rename_map[col] = "Moneda"
        elif "observ" in col_lower or "nota" in col_lower:
            rename_map[col] = "Observaciones"

    raw.rename(columns=rename_map, inplace=True)

    # Convertir tipos
    if "Fecha" in raw.columns:
        raw["Fecha"] = pd.to_datetime(raw["Fecha"], errors="coerce")
    if "Monto" in raw.columns:
        raw["Monto"] = pd.to_numeric(raw["Monto"], errors="coerce")
    if "Anio" in raw.columns:
        raw["Anio"] = pd.to_numeric(raw["Anio"], errors="coerce")

    # Filtrar solo filas con monto válido
    raw = raw.dropna(subset=["Monto"])
    return raw


def cargar_sueldos(xl: pd.ExcelFile) -> pd.DataFrame:
    """Carga la hoja 'Sueldos' del Excel y normaliza su estructura."""
    df = xl.parse("Sueldos")

    # Renombrar columnas clave
    rename_map = {}
    for col in df.columns:
        col_lower = str(col).lower().strip()
        if "fecha" in col_lower:
            rename_map[col] = "Fecha"
        elif col_lower == "mes":
            rename_map[col] = "Mes"
        elif "año" in col_lower or "year" in col_lower or col_lower == "año":
            rename_map[col] = "Anio"
        elif "sueldo" in col_lower and "pesos" in col_lower:
            rename_map[col] = "Sueldo_Pesos"
        elif "sueldo" in col_lower and "u$s" in col_lower:
            rename_map[col] = "Sueldo_USD"
        elif col_lower == "dolar":
            rename_map[col] = "Dolar"
        elif col_lower == "acumulado":
            rename_map[col] = "Acumulado"

    df.rename(columns=rename_map, inplace=True)

    if "Fecha" in df.columns:
        df["Fecha"] = pd.to_datetime(df["Fecha"], errors="coerce")
    if "Sueldo_Pesos" in df.columns:
        df["Sueldo_Pesos"] = pd.to_numeric(df["Sueldo_Pesos"], errors="coerce")

    return df.dropna(subset=["Sueldo_Pesos"])


# ============================================================
# FUNCIONES DE CÁLCULO Y ANÁLISIS
# ============================================================

def calcular_resumen(df_gastos: pd.DataFrame, df_sueldos: pd.DataFrame) -> dict:
    """Calcula las métricas principales del resumen financiero."""
    total_ingresos = df_sueldos["Sueldo_Pesos"].sum() if "Sueldo_Pesos" in df_sueldos.columns else 0
    total_gastos   = df_gastos["Monto"].sum() if "Monto" in df_gastos.columns else 0
    balance        = total_ingresos - total_gastos
    tasa_ahorro    = (balance / total_ingresos * 100) if total_ingresos > 0 else 0

    return {
        "total_ingresos": total_ingresos,
        "total_gastos":   total_gastos,
        "balance":        balance,
        "tasa_ahorro":    tasa_ahorro,
    }


def gastos_por_familia(df_gastos: pd.DataFrame, familias: dict) -> dict:
    """
    Agrupa el total de gastos por familia según la asignación del Maestro.
    Conceptos sin familia van a 'Sin clasificar'.
    """
    resultado = {}
    conceptos_asignados = set()

    for familia, conceptos in familias.items():
        if not conceptos:
            continue
        mask = df_gastos["Concepto"].isin(conceptos)
        total = df_gastos.loc[mask, "Monto"].sum()
        if total > 0:
            resultado[familia] = total
            conceptos_asignados.update(conceptos)

    # Sin clasificar
    sin_clasificar = df_gastos.loc[~df_gastos["Concepto"].isin(conceptos_asignados), "Monto"].sum()
    if sin_clasificar > 0:
        resultado["Sin clasificar"] = sin_clasificar

    return resultado


def generar_contexto_ia(df_gastos: pd.DataFrame, df_sueldos: pd.DataFrame, familias: dict) -> str:
    """
    Construye el contexto financiero resumido que se enviará a la IA.
    Solo se envía información estadística, no datos crudos, para minimizar tokens.
    """
    resumen = calcular_resumen(df_gastos, df_sueldos)

    # Gastos por concepto (top 8)
    top_conceptos = (
        df_gastos.groupby("Concepto")["Monto"]
        .sum()
        .sort_values(ascending=False)
        .head(8)
    )

    # Gastos por mes
    if "Mes" in df_gastos.columns:
        por_mes = df_gastos.groupby("Mes")["Monto"].sum().to_dict()
    else:
        por_mes = {}

    # Sueldos por mes
    sueldos_por_mes = {}
    if "Mes" in df_sueldos.columns:
        sueldos_por_mes = df_sueldos.set_index("Mes")["Sueldo_Pesos"].to_dict()

    # Familias con montos
    fam_montos = gastos_por_familia(df_gastos, familias)

    ctx = f"""
=== RESUMEN FINANCIERO ===
- Total ingresos (sueldos): ${resumen['total_ingresos']:,.0f} ARS
- Total gastos: ${resumen['total_gastos']:,.0f} ARS
- Balance: ${resumen['balance']:,.0f} ARS
- Tasa de ahorro: {resumen['tasa_ahorro']:.1f}%

=== GASTOS POR CATEGORÍA (familias) ===
{chr(10).join(f"- {k}: ${v:,.0f}" for k, v in fam_montos.items())}

=== TOP CONCEPTOS DE GASTO ===
{top_conceptos.to_string()}

=== GASTOS POR MES ===
{chr(10).join(f"- {mes}: ${monto:,.0f}" for mes, monto in por_mes.items())}

=== INGRESOS POR MES ===
{chr(10).join(f"- {mes}: ${sueldo:,.0f}" for mes, sueldo in sueldos_por_mes.items())}
"""
    return ctx.strip()


# ============================================================
# FUNCIÓN DE LLAMADA A LA IA (prompt dirigido con salida estructurada)
# ============================================================

def consultar_ia(api_key: str, contexto: str, pregunta: str = None) -> str:
    """
    Llama a Gemini 1.5 Flash con un system prompt de asesor financiero.
    Parámetros:
        api_key  : Google Gemini API key
        contexto : resumen financiero del usuario
        pregunta : consulta puntual (None = análisis general)
    """

    # Configurar Gemini con la API key del usuario
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel(
        model_name="gemini-2.0-flash",
        system_instruction="""Eres FinanzIA, un asesor experto en finanzas personales en Argentina.
Tu misión es ayudar al usuario a entender su situación financiera y mejorar sus hábitos de gasto.

REGLAS:
1. Respondé siempre en español, de forma clara y empática.
2. Usá emojis estratégicamente para hacer la lectura más amigable.
3. Estructurá tu respuesta con secciones en markdown (## para títulos, - para listas).
4. Basate ÚNICAMENTE en los datos financieros provistos; no inventes cifras.
5. Sé concreto: citá números reales del contexto cuando sea relevante.
6. Limitá tu respuesta a 400 palabras máximo para ser conciso.
7. Si hay algo preocupante, mencionalo con tacto pero con claridad.""",
    )

    # Construir el mensaje
    if pregunta:
        user_msg = f"Mis datos financieros:\n{contexto}\n\nMi pregunta: {pregunta}"
    else:
        user_msg = f"Analizá mi situación financiera y dame recomendaciones accionables:\n{contexto}"

    response = model.generate_content(user_msg)
    return response.text


# ============================================================
# FUNCIONES DE LÍMITES Y ALERTAS
# ============================================================

def gasto_total_familia(familia: str, df_gastos: pd.DataFrame | None,
                         gastos_manuales: list, familias: dict) -> float:
    """
    Suma todos los gastos (Excel + manuales) asociados a una familia.
    """
    conceptos = familias.get(familia, [])
    total = 0.0

    # Gastos del Excel
    if df_gastos is not None and conceptos:
        mask  = df_gastos["Concepto"].isin(conceptos)
        total += float(df_gastos.loc[mask, "Monto"].sum())

    # Gastos registrados manualmente
    for g in gastos_manuales:
        if g["familia"] == familia:
            total += g["monto"]

    return total


def estado_limite(gastado: float, limite: float) -> dict:
    """
    Devuelve el estado de consumo respecto al límite fijado.
    Retorna un dict con:
        pct        : porcentaje consumido (0-100+)
        restante   : $ que faltan para llegar al límite
        estado     : 'ok' | 'warning' | 'danger'
        mensaje    : texto legible para el usuario
    """
    if limite <= 0:
        return {"pct": 0, "restante": 0, "estado": "ok", "mensaje": ""}

    pct      = (gastado / limite) * 100
    restante = limite - gastado

    if pct >= 100:
        return {
            "pct": pct, "restante": restante, "estado": "danger",
            "mensaje": f"⛔ ¡Límite ALCANZADO! Gastaste ${gastado:,.0f} de ${limite:,.0f} ({pct:.0f}%)"
        }
    elif pct >= 70:          # faltan <= 30 % del límite
        return {
            "pct": pct, "restante": restante, "estado": "warning",
            "mensaje": f"⚠️ Te faltan ${restante:,.0f} para llegar al límite (${limite:,.0f}). Llevás ${gastado:,.0f} ({pct:.0f}%)"
        }
    else:
        return {
            "pct": pct, "restante": restante, "estado": "ok",
            "mensaje": f"✅ Vas bien: ${gastado:,.0f} de ${limite:,.0f} ({pct:.0f}%)"
        }


def verificar_y_mostrar_alertas(familia: str, df_gastos, gastos_manuales: list,
                                 familias: dict, limites: dict,
                                 usar_toast: bool = True):
    """
    Evalúa el estado de una familia y muestra alertas emergentes (toast)
    y/o mensajes inline según corresponda.
    Retorna el dict de estado para uso posterior.
    """
    limite  = limites.get(familia, 0)
    if limite <= 0:
        return None   # Sin límite definido, nada que alertar

    gastado = gasto_total_familia(familia, df_gastos, gastos_manuales, familias)
    est     = estado_limite(gastado, limite)

    if usar_toast:
        if est["estado"] == "danger":
            st.toast(est["mensaje"], icon="⛔")
        elif est["estado"] == "warning":
            st.toast(est["mensaje"], icon="⚠️")

    return est


# ============================================================
# INICIALIZACIÓN DE SESSION STATE
# ============================================================

def init_state():
    """Inicializa todas las variables de estado de la sesión."""
    defaults = {
        "df_gastos":      None,
        "df_sueldos":     None,
        "familias":       {
            "Alimentación":   ["Comida", "Super", "Mercado"],
            "Transporte":     ["Nafta", "Combustible", "Peaje"],
            "Servicios":      ["Servicios", "Luz", "Gas", "Internet", "Agua"],
            "Entretenimiento": ["Salida", "Viaje", "Vacaciones"],
            "Salud":          ["Farmacia", "Médico", "Prepaga"],
            "Otros":          [],
        },
        "analisis_cache": None,   # Cachea el análisis general para no repetir llamadas
        "chat_historial": [],
        # Límites de gasto por familia (en ARS). 0 = sin límite definido.
        "limites": {},
        # Gastos registrados manualmente en la app (además del Excel)
        "gastos_manuales": [],
    }
    for key, val in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = val


init_state()


# ============================================================
# HEADER
# ============================================================
st.markdown("""
<div class="main-header">
    <h1>💰 FinanzIA</h1>
    <p>Tu asistente inteligente de finanzas personales · Diplomatura Data Science · CoderHouse</p>
</div>
""", unsafe_allow_html=True)


# ============================================================
# SIDEBAR: Configuración y carga de datos
# ============================================================
with st.sidebar:
    st.title("⚙️ Configuración")

    # --- API Key ---
    api_key = st.text_input(
        "🔑 Gemini API Key",
        type="password",
        help="Tu clave de Google Gemini (aistudio.google.com). No se almacena.",
    )

    st.divider()

    # --- Carga de Excel ---
    st.subheader("📂 Cargar Datos")
    archivo = st.file_uploader(
        "Subí tu archivo Excel (Finanzas.xlsx)",
        type=["xlsx", "xls"],
        help="Debe tener hojas 'Gastos' y 'Sueldos'",
    )

    if archivo:
        try:
            xl = pd.ExcelFile(archivo)
            hojas = xl.sheet_names
            st.success(f"✅ Archivo cargado · {len(hojas)} hojas detectadas")

            # Verificar hojas esperadas
            tiene_gastos  = "Gastos"  in hojas
            tiene_sueldos = "Sueldos" in hojas

            if tiene_gastos and tiene_sueldos:
                if st.button("⚡ Procesar Datos", type="primary", use_container_width=True):
                    with st.spinner("Procesando..."):
                        st.session_state.df_gastos  = cargar_gastos(xl)
                        st.session_state.df_sueldos = cargar_sueldos(xl)
                        st.session_state.analisis_cache = None  # Reset caché IA
                    st.success("¡Datos listos!")
                    st.rerun()
            else:
                st.warning(
                    "El archivo debe tener hojas llamadas exactamente "
                    "'Gastos' y 'Sueldos'."
                )
        except Exception as e:
            st.error(f"Error al leer el archivo: {e}")

    # Estado de datos cargados
    st.divider()
    if st.session_state.df_gastos is not None:
        st.success(f"📊 Gastos: {len(st.session_state.df_gastos)} registros")
    if st.session_state.df_sueldos is not None:
        st.success(f"💼 Sueldos: {len(st.session_state.df_sueldos)} registros")

    st.divider()
    st.caption("FinanzIA v1.0 | Proyecto Final Prompt Engineering")


# ============================================================
# TABS PRINCIPALES
# ============================================================
tab_dash, tab_maestro, tab_ia, tab_info = st.tabs([
    "📊 Dashboard",
    "📋 Maestro de Gastos",
    "🤖 Análisis IA",
    "ℹ️ Cómo Funciona",
])


# ============================================================
# TAB 1 – DASHBOARD
# ============================================================
with tab_dash:
    st.header("📊 Dashboard Financiero")

    if st.session_state.df_gastos is None:
        # Estado vacío: instrucciones
        st.info("👈 Cargá tu **Finanzas.xlsx** desde el panel lateral para ver el dashboard.")

        st.subheader("Formato esperado del Excel")
        col_a, col_b = st.columns(2)
        with col_a:
            st.caption("📋 Hoja **Gastos**")
            st.dataframe(
                pd.DataFrame({
                    "Fecha": ["2026-01-01", "2026-01-05"],
                    "Concepto": ["Comida", "Nafta"],
                    "Monto": [9200, 35002],
                    "Medio de Pago": ["Débito", "Efectivo"],
                }),
                hide_index=True,
            )
        with col_b:
            st.caption("💼 Hoja **Sueldos**")
            st.dataframe(
                pd.DataFrame({
                    "Fecha": ["2026-01-01", "2026-02-01"],
                    "Mes": ["enero", "febrero"],
                    "Sueldo en Pesos": [3323469, 3620015],
                }),
                hide_index=True,
            )
    else:
        df_g = st.session_state.df_gastos
        df_s = st.session_state.df_sueldos
        resumen = calcular_resumen(df_g, df_s)

        # --- Métricas principales ---
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("💚 Total Ingresos",   f"${resumen['total_ingresos']:,.0f}")
        c2.metric("🔴 Total Gastos",     f"${resumen['total_gastos']:,.0f}")
        c3.metric(
            "💙 Balance",
            f"${resumen['balance']:,.0f}",
            delta=f"{resumen['tasa_ahorro']:.1f}% ahorro",
            delta_color="normal" if resumen["balance"] >= 0 else "inverse",
        )
        c4.metric("📈 Tasa de Ahorro",   f"{resumen['tasa_ahorro']:.1f}%")

        st.divider()

        # --- Fila de gráficos ---
        col1, col2 = st.columns(2)

        # Gráfico 1: Ingresos vs Gastos por mes
        with col1:
            meses_orden = ["ENERO","FEBRERO","MARZO","ABRIL","MAYO","JUNIO",
                           "JULIO","AGOSTO","SEPTIEMBRE","OCTUBRE","NOVIEMBRE","DICIEMBRE"]

            if "Mes" in df_g.columns and "Mes" in df_s.columns:
                gastos_mes  = df_g.groupby("Mes")["Monto"].sum().reset_index()
                gastos_mes.columns = ["Mes", "Gastos"]
                gastos_mes["Mes"] = gastos_mes["Mes"].str.upper()

                sueldos_mes = df_s.copy()
                sueldos_mes["Mes_upper"] = sueldos_mes["Mes"].str.upper()
                merged = gastos_mes.merge(
                    sueldos_mes[["Mes_upper","Sueldo_Pesos"]].rename(columns={"Mes_upper":"Mes"}),
                    on="Mes", how="outer"
                ).fillna(0)
                # Ordenar por mes calendario
                merged["orden"] = merged["Mes"].map(
                    {m: i for i, m in enumerate(meses_orden)}
                ).fillna(99)
                merged = merged.sort_values("orden")

                fig_mes = go.Figure()
                fig_mes.add_trace(go.Bar(
                    name="Ingresos", x=merged["Mes"], y=merged["Sueldo_Pesos"],
                    marker_color="#4CAF50"
                ))
                fig_mes.add_trace(go.Bar(
                    name="Gastos", x=merged["Mes"], y=merged["Gastos"],
                    marker_color="#f44336"
                ))
                fig_mes.update_layout(
                    title="Ingresos vs Gastos por Mes",
                    barmode="group", height=370,
                    legend=dict(orientation="h", y=-0.2)
                )
                st.plotly_chart(fig_mes, use_container_width=True)

        # Gráfico 2: Torta por concepto
        with col2:
            conceptos_total = df_g.groupby("Concepto")["Monto"].sum().reset_index()
            fig_pie = px.pie(
                conceptos_total,
                names="Concepto", values="Monto",
                title="Distribución de Gastos por Concepto",
                color_discrete_sequence=px.colors.qualitative.Bold,
            )
            fig_pie.update_layout(height=370)
            st.plotly_chart(fig_pie, use_container_width=True)

        st.divider()

        # Gráfico 3: Evolución temporal de gastos
        if "Fecha" in df_g.columns:
            df_evol = df_g.dropna(subset=["Fecha"]).copy()
            df_evol["FechaMes"] = df_evol["Fecha"].dt.to_period("M").astype(str)
            evol = df_evol.groupby("FechaMes")["Monto"].sum().reset_index()
            fig_line = px.line(
                evol, x="FechaMes", y="Monto",
                title="📈 Evolución de Gastos Mensuales",
                markers=True,
                color_discrete_sequence=["#f0c040"],
            )
            fig_line.update_layout(height=280, xaxis_title="Mes", yaxis_title="$ ARS")
            st.plotly_chart(fig_line, use_container_width=True)

        st.divider()

        # --- Panel de alertas de límites (si hay límites configurados) ---
        limites_activos = {
            fam: lim for fam, lim in st.session_state.limites.items() if lim > 0
        }
        if limites_activos:
            alertas_warn   = []
            alertas_danger = []

            for fam, lim in limites_activos.items():
                gastado = gasto_total_familia(
                    fam, df_g,
                    st.session_state.gastos_manuales,
                    st.session_state.familias,
                )
                est = estado_limite(gastado, lim)
                if est["estado"] == "danger":
                    alertas_danger.append((fam, est))
                elif est["estado"] == "warning":
                    alertas_warn.append((fam, est))

            if alertas_danger or alertas_warn:
                st.subheader("🔔 Alertas de Límites")
                for fam, est in alertas_danger:
                    st.markdown(
                        f'<div class="alerta-danger">'
                        f"<b>{fam}:</b> {est['mensaje']}</div>",
                        unsafe_allow_html=True,
                    )
                for fam, est in alertas_warn:
                    st.markdown(
                        f'<div class="alerta-warning">'
                        f"<b>{fam}:</b> {est['mensaje']}</div>",
                        unsafe_allow_html=True,
                    )
                st.divider()

        # Tablas de datos
        with st.expander("🔍 Ver tabla completa de gastos"):
            st.dataframe(df_g, hide_index=True, use_container_width=True)
        with st.expander("🔍 Ver tabla de sueldos"):
            st.dataframe(df_s, hide_index=True, use_container_width=True)


# ============================================================
# TAB 2 – MAESTRO DE GASTOS
# ============================================================
with tab_maestro:
    st.header("📋 Maestro de Gastos")
    st.caption(
        "Organizá tus gastos por familias, fijá límites de consumo y registrá "
        "nuevos gastos para recibir alertas en tiempo real."
    )

    # ── Sub-tabs internos ──────────────────────────────────────
    sub_familias, sub_registrar, sub_limites = st.tabs([
        "📁 Familias & Conceptos",
        "➕ Registrar Gasto",
        "🎯 Límites de Consumo",
    ])

    # ── SUB-TAB A: Familias & Conceptos ───────────────────────
    with sub_familias:
        col_izq, col_der = st.columns([1, 2])

        with col_izq:
            # --- Crear nueva familia ---
            st.subheader("Nueva Familia")
            nueva_familia = st.text_input(
                "Nombre",
                placeholder="Ej: Streaming",
                key="inp_nueva_familia",
            )
            if st.button("Crear Familia", type="primary", use_container_width=True):
                nombre = nueva_familia.strip()
                if nombre and nombre not in st.session_state.familias:
                    st.session_state.familias[nombre] = []
                    st.session_state.limites[nombre] = 0
                    st.success(f"✅ Familia '{nombre}' creada")
                    st.rerun()
                elif nombre in st.session_state.familias:
                    st.warning("Esa familia ya existe.")
                else:
                    st.warning("Ingresá un nombre.")

            st.divider()

            # --- Agregar concepto a familia ---
            st.subheader("Agregar Conceptos")

            # Todos los conceptos ya asignados (en cualquier familia)
            conceptos_ya_asignados = {
                c for items in st.session_state.familias.values() for c in items
            }

            familia_destino = st.selectbox(
                "Familia de destino",
                list(st.session_state.familias.keys()),
                key="sel_familia_destino",
            )

            if st.session_state.df_gastos is not None:
                # Conceptos del Excel aún sin asignar a ninguna familia
                todos_conceptos = sorted(
                    st.session_state.df_gastos["Concepto"].dropna().unique().tolist()
                )
                sin_asignar = [c for c in todos_conceptos if c not in conceptos_ya_asignados]

                if sin_asignar:
                    st.caption(f"{len(sin_asignar)} conceptos del Excel sin asignar:")
                    # Multiselect: elegís varios de una vez
                    seleccionados = st.multiselect(
                        "Seleccioná uno o más",
                        sin_asignar,
                        key="multi_conceptos",
                        placeholder="Buscá o elegí...",
                    )
                    if st.button("Asignar seleccionados", type="primary",
                                 use_container_width=True, disabled=not seleccionados):
                        for c in seleccionados:
                            if c not in st.session_state.familias[familia_destino]:
                                st.session_state.familias[familia_destino].append(c)
                        st.success(f"✅ {len(seleccionados)} concepto(s) → '{familia_destino}'")
                        st.rerun()
                else:
                    st.success("✅ Todos los conceptos del Excel ya están asignados.")

            st.divider()

            # Escritura manual (para conceptos que no están en el Excel)
            st.caption("Agregar concepto personalizado:")
            concepto_manual = st.text_input(
                "Nombre", placeholder="Ej: Netflix", key="inp_concepto_manual"
            )
            if st.button("Agregar manual", use_container_width=True):
                c = concepto_manual.strip()
                if not c:
                    st.warning("Escribí un nombre.")
                elif c in conceptos_ya_asignados:
                    # Buscar en qué familia está
                    fam_actual = next(
                        (f for f, items in st.session_state.familias.items() if c in items), "?"
                    )
                    st.warning(f"'{c}' ya está en la familia **{fam_actual}**. "
                               f"Quitalo de ahí primero si querés moverlo.")
                else:
                    st.session_state.familias[familia_destino].append(c)
                    st.success(f"✅ '{c}' → '{familia_destino}'")
                    st.rerun()

        with col_der:
            st.subheader("Familias Actuales")

            for familia, conceptos in st.session_state.familias.items():
                n = len(conceptos)
                gastado = gasto_total_familia(
                    familia,
                    st.session_state.df_gastos,
                    st.session_state.gastos_manuales,
                    st.session_state.familias,
                )
                limite = st.session_state.limites.get(familia, 0)
                est    = estado_limite(gastado, limite) if limite > 0 else None

                # Etiqueta del expander con estado de límite
                if est:
                    icono = {"ok": "✅", "warning": "⚠️", "danger": "⛔"}[est["estado"]]
                    label = f"📂 {familia}  ({n} conceptos · ${gastado:,.0f} / ${limite:,.0f}) {icono}"
                elif gastado > 0:
                    label = f"📂 {familia}  ({n} conceptos · ${gastado:,.0f} ARS)"
                else:
                    label = f"📂 {familia}  ({n} conceptos)"

                with st.expander(label, expanded=bool(est and est["estado"] != "ok")):
                    # Barra de progreso si hay límite
                    if est:
                        pct_clamp = min(est["pct"] / 100, 1.0)
                        st.progress(pct_clamp)
                        if est["estado"] == "danger":
                            st.markdown(
                                f'<div class="alerta-danger">{est["mensaje"]}</div>',
                                unsafe_allow_html=True,
                            )
                        elif est["estado"] == "warning":
                            st.markdown(
                                f'<div class="alerta-warning">{est["mensaje"]}</div>',
                                unsafe_allow_html=True,
                            )

                    if conceptos:
                        chips_html = "".join(
                            f'<span class="chip">{c}</span>' for c in conceptos
                        )
                        st.markdown(chips_html, unsafe_allow_html=True)
                        st.write("")

                        concepto_a_eliminar = st.selectbox(
                            "Eliminar concepto:",
                            ["(seleccionar)"] + conceptos,
                            key=f"del_concepto_{familia}",
                        )
                        if concepto_a_eliminar != "(seleccionar)":
                            if st.button(
                                f"🗑️ Quitar '{concepto_a_eliminar}'",
                                key=f"btn_del_{familia}_{concepto_a_eliminar}",
                            ):
                                st.session_state.familias[familia].remove(concepto_a_eliminar)
                                st.rerun()
                    else:
                        st.caption("Sin conceptos asignados aún.")

                    # Eliminar familia vacía (no las predeterminadas)
                    predeterminadas = {"Alimentación","Transporte","Servicios",
                                       "Entretenimiento","Salud","Otros"}
                    if n == 0 and familia not in predeterminadas:
                        if st.button(f"🗑️ Eliminar familia", key=f"btn_del_fam_{familia}"):
                            del st.session_state.familias[familia]
                            if familia in st.session_state.limites:
                                del st.session_state.limites[familia]
                            st.rerun()

    # ── SUB-TAB B: Registrar Gasto ─────────────────────────────
    with sub_registrar:
        st.subheader("➕ Registrar nuevo gasto")
        st.caption(
            "Registrá un gasto manual (fuera del Excel). "
            "Si la familia tiene un límite configurado, verás la alerta al instante."
        )

        col_form, col_historial = st.columns([1, 1])

        with col_form:
            with st.form("form_nuevo_gasto", clear_on_submit=True):
                fecha_gasto    = st.date_input("Fecha", value=datetime.today())
                familia_gasto  = st.selectbox(
                    "Familia", list(st.session_state.familias.keys())
                )
                concepto_gasto = st.text_input(
                    "Concepto", placeholder="Ej: Supermercado Día"
                )
                monto_gasto    = st.number_input(
                    "Monto ($)", min_value=0.0, step=100.0, format="%.0f"
                )
                medio_pago     = st.selectbox(
                    "Medio de pago",
                    ["Débito", "Crédito", "Efectivo", "Transferencia", "Otro"],
                )
                registrar_btn  = st.form_submit_button(
                    "💾 Registrar Gasto", type="primary", use_container_width=True
                )

            if registrar_btn:
                if not concepto_gasto.strip():
                    st.warning("Ingresá un concepto para el gasto.")
                elif monto_gasto <= 0:
                    st.warning("El monto debe ser mayor a cero.")
                else:
                    # Guardar en session state
                    nuevo = {
                        "fecha":    str(fecha_gasto),
                        "familia":  familia_gasto,
                        "concepto": concepto_gasto.strip(),
                        "monto":    float(monto_gasto),
                        "medio":    medio_pago,
                    }
                    st.session_state.gastos_manuales.append(nuevo)

                    # ── EVALUAR ALERTAS INMEDIATAMENTE ──────────
                    est = verificar_y_mostrar_alertas(
                        familia_gasto,
                        st.session_state.df_gastos,
                        st.session_state.gastos_manuales,
                        st.session_state.familias,
                        st.session_state.limites,
                        usar_toast=True,   # dispara el toast emergente
                    )

                    # Mensaje de confirmación + alerta inline
                    gastado_total = gasto_total_familia(
                        familia_gasto,
                        st.session_state.df_gastos,
                        st.session_state.gastos_manuales,
                        st.session_state.familias,
                    )
                    st.success(
                        f"✅ Gasto registrado: {concepto_gasto} · ${monto_gasto:,.0f}"
                    )

                    if est:
                        if est["estado"] == "danger":
                            st.markdown(
                                f'<div class="alerta-danger">{est["mensaje"]}</div>',
                                unsafe_allow_html=True,
                            )
                        elif est["estado"] == "warning":
                            st.markdown(
                                f'<div class="alerta-warning">{est["mensaje"]}</div>',
                                unsafe_allow_html=True,
                            )
                    else:
                        limite = st.session_state.limites.get(familia_gasto, 0)
                        if limite > 0:
                            restante = limite - gastado_total
                            st.info(f"Te quedan ${restante:,.0f} disponibles en '{familia_gasto}'.")

        with col_historial:
            st.subheader("Historial de gastos registrados")
            if st.session_state.gastos_manuales:
                df_manual = pd.DataFrame(st.session_state.gastos_manuales)
                st.dataframe(df_manual, hide_index=True, use_container_width=True)
                total_manual = df_manual["monto"].sum()
                st.metric("Total registrado manualmente", f"${total_manual:,.0f}")

                if st.button("🗑️ Limpiar historial manual"):
                    st.session_state.gastos_manuales = []
                    st.rerun()
            else:
                st.info("Todavía no registraste gastos manualmente.")

    # ── SUB-TAB C: Límites de Consumo ─────────────────────────
    with sub_limites:
        st.subheader("🎯 Configurar límites por familia")
        st.caption(
            "Fijá un límite mensual (en $) para cada familia de gastos. "
            "Recibirás una alerta cuando falte el 30% o menos para alcanzarlo, "
            "y otra alerta al superarlo."
        )

        st.info(
            "**¿Cómo funcionan las alertas?**\n\n"
            "- 🟠 **Alerta amarilla** — Llevás el 70% o más del límite "
            "(te falta ≤ 30% del tope)\n"
            "- 🔴 **Alerta roja** — Alcanzaste o superaste el límite\n\n"
            "Las alertas aparecen como notificaciones emergentes (toast) al "
            "registrar un gasto, y también en forma inline en la tarjeta de cada familia."
        )

        st.divider()

        # Formulario de límites por familia
        familias_lista = list(st.session_state.familias.keys())

        # Mostramos en grilla de 2 columnas
        pares = [familias_lista[i:i+2] for i in range(0, len(familias_lista), 2)]

        for par in pares:
            cols = st.columns(2)
            for col, fam in zip(cols, par):
                with col:
                    gastado = gasto_total_familia(
                        fam,
                        st.session_state.df_gastos,
                        st.session_state.gastos_manuales,
                        st.session_state.familias,
                    )
                    limite_actual = st.session_state.limites.get(fam, 0)
                    est = estado_limite(gastado, limite_actual) if limite_actual > 0 else None

                    with st.container(border=True):
                        st.markdown(f"**📂 {fam}**")

                        nuevo_limite = st.number_input(
                            f"Límite mensual ($)",
                            min_value=0.0,
                            value=float(limite_actual),
                            step=1000.0,
                            format="%.0f",
                            key=f"limite_{fam}",
                            help="0 = sin límite",
                        )

                        # Guardar automáticamente al cambiar
                        if nuevo_limite != limite_actual:
                            st.session_state.limites[fam] = float(nuevo_limite)
                            st.rerun()

                        # Estado actual
                        if nuevo_limite > 0:
                            est_nuevo = estado_limite(gastado, nuevo_limite)
                            pct_clamp = min(est_nuevo["pct"] / 100, 1.0)
                            st.progress(pct_clamp)

                            color_class = {
                                "ok":      "limite-ok",
                                "warning": "limite-warning",
                                "danger":  "limite-danger",
                            }[est_nuevo["estado"]]
                            st.markdown(
                                f'<span class="{color_class}">'
                                f"${gastado:,.0f} / ${nuevo_limite:,.0f} "
                                f"({est_nuevo['pct']:.0f}%)</span>",
                                unsafe_allow_html=True,
                            )
                        else:
                            st.caption(f"Gastado: ${gastado:,.0f} · Sin límite")


# ============================================================
# TAB 3 – ANÁLISIS IA
# ============================================================
with tab_ia:
    st.header("🤖 Análisis con Inteligencia Artificial")

    datos_listos = st.session_state.df_gastos is not None
    key_lista    = bool(api_key)

    if not key_lista:
        st.warning("⚠️ Ingresá tu **Gemini API Key** en el panel lateral para usar esta función.")
    elif not datos_listos:
        st.info("📂 Primero cargá tu archivo **Finanzas.xlsx** desde el panel lateral.")
    else:
        contexto = generar_contexto_ia(
            st.session_state.df_gastos,
            st.session_state.df_sueldos,
            st.session_state.familias,
        )

        # --- Análisis General ---
        st.subheader("📊 Análisis Financiero General")
        st.caption(
            "Se genera un análisis completo de tu situación financiera. "
            "Se cachea para no repetir la llamada a la API."
        )

        col_btn, col_reset = st.columns([2, 1])
        with col_btn:
            if st.session_state.analisis_cache is None:
                if st.button("🔍 Generar Análisis General", type="primary", use_container_width=True):
                    with st.spinner("FinanzIA está analizando tus finanzas..."):
                        try:
                            st.session_state.analisis_cache = consultar_ia(api_key, contexto)
                        except Exception as e:
                            st.error(f"Error al conectar con Gemini: {e}")
        with col_reset:
            if st.session_state.analisis_cache:
                if st.button("🔄 Regenerar", use_container_width=True):
                    st.session_state.analisis_cache = None
                    st.rerun()

        if st.session_state.analisis_cache:
            with st.container(border=True):
                st.markdown(st.session_state.analisis_cache)

        st.divider()

        # --- Chat financiero ---
        st.subheader("💬 Consultá a FinanzIA")
        st.caption(
            "Hacé preguntas específicas sobre tus finanzas. "
            "Cada mensaje genera una nueva llamada a la API."
        )

        # Mostrar historial de chat
        for msg in st.session_state.chat_historial:
            with st.chat_message(msg["role"]):
                st.markdown(msg["content"])

        # Sugerencias de preguntas
        with st.expander("💡 Ejemplos de preguntas"):
            st.markdown("""
            - ¿En qué concepto gasto más?
            - ¿Cuál fue mi mes con mayor gasto?
            - ¿Cómo puedo reducir mis gastos en alimentación?
            - ¿Mi tasa de ahorro es saludable?
            - ¿Cómo evolucionaron mis ingresos?
            """)

        # Input de chat
        pregunta = st.chat_input("Preguntale algo a FinanzIA sobre tus finanzas...")
        if pregunta:
            st.session_state.chat_historial.append({"role": "user", "content": pregunta})
            with st.chat_message("user"):
                st.markdown(pregunta)

            with st.chat_message("assistant"):
                with st.spinner("Analizando..."):
                    try:
                        respuesta = consultar_ia(api_key, contexto, pregunta)
                        st.markdown(respuesta)
                        st.session_state.chat_historial.append(
                            {"role": "assistant", "content": respuesta}
                        )
                    except Exception as e:
                        st.error(f"Error: {e}")

        if st.session_state.chat_historial:
            if st.button("🗑️ Limpiar conversación"):
                st.session_state.chat_historial = []
                st.rerun()


# ============================================================
# TAB 4 – CÓMO FUNCIONA
# ============================================================
with tab_info:
    st.header("ℹ️ ¿Cómo funciona FinanzIA?")

    st.markdown("""
    **FinanzIA** es una aplicación web de finanzas personales potenciada por Inteligencia Artificial.
    Cargás tu propio historial de gastos e ingresos en Excel, organizás los conceptos por familias
    y obtenés análisis y recomendaciones personalizadas gracias a Gemini 1.5 Flash de Google.
    """)

    st.divider()

    col1, col2 = st.columns(2)

    with col1:
        st.subheader("🚀 Pasos para usar la app")
        st.markdown("""
        **Paso 1 — Cargá tu Excel** 📂
        Desde el panel lateral subí tu archivo `Finanzas.xlsx`.
        Debe tener dos hojas: `Gastos` y `Sueldos`.

        **Paso 2 — Procesá los datos** ⚡
        Hacé clic en "Procesar Datos". La app normaliza
        automáticamente las columnas y los tipos de datos.

        **Paso 3 — Explorá el Dashboard** 📊
        Visualizá ingresos vs gastos por mes, la distribución
        por concepto y la evolución temporal.

        **Paso 4 — Organizá el Maestro de Gastos** 🗂️
        Creá familias (ej. Streaming, Transporte) y asigná
        tus conceptos de gasto a cada una para un análisis
        más preciso.

        **Paso 5 — Analizá con IA** 🤖
        Ingresá tu Gemini API Key y generá un análisis
        automático o hacé preguntas en el chat financiero.
        """)

    with col2:
        st.subheader("🧠 Cómo funciona la IA")
        st.markdown("""
        La IA usa el modelo **Gemini 2.0 Flash** de Google con
        un *prompt de sistema* que define el rol de asesor
        financiero personal.

        **Prompt inicial:**
        > *"Eres FinanzIA, un asesor experto en finanzas personales
        en Argentina. Tu misión es ayudar al usuario a entender
        su situación financiera y mejorar sus hábitos de gasto..."*

        **¿Qué datos se envían a la IA?**
        Solo un resumen estadístico (totales, promedios, top conceptos),
        **nunca datos crudos ni información personal** más allá de cifras.

        **Optimización de costos:**
        - Modelo: gpt-3.5-turbo (~$0.001 por análisis)
        - El análisis general se **cachea** en la sesión para
          no repetir llamadas innecesarias.
        - El chat solo llama a la API cuando el usuario envía
          un mensaje (no hay loops ni llamadas automáticas).
        - Máximo 450 tokens por respuesta.
        """)

        st.divider()

        st.subheader("🔒 Privacidad")
        st.markdown("""
        - Tu **API Key nunca se almacena** en ningún servidor.
        - Los datos del Excel se procesan **localmente** en
          la sesión de Streamlit y no se persisten entre sesiones.
        - Solo se envían **resúmenes numéricos** a la API de Gemini.
        """)

    st.divider()

    st.subheader("📁 Formato del Excel requerido")
    col_a, col_b = st.columns(2)
    with col_a:
        st.caption("Hoja **Gastos**")
        st.dataframe(pd.DataFrame({
            "Fecha":        ["2026-01-01", "2026-01-05"],
            "Año_Mes":      ["2026-01-01", "2026-01-01"],
            "Mes":          ["ENERO",      "ENERO"],
            "Year":         [2026,         2026],
            "Concepto":     ["Comida",     "Nafta"],
            "Monto":        [9200,         35002],
            "Medio de Pago":["Débito",     "Efectivo"],
        }), hide_index=True)
    with col_b:
        st.caption("Hoja **Sueldos**")
        st.dataframe(pd.DataFrame({
            "Fecha":          ["2026-01-01", "2026-02-01"],
            "Mes":            ["enero",      "febrero"],
            "Año":            [2026,         2026],
            "Sueldo en Pesos":[3323469,      3620015],
            "Sueldo en U$S":  [2245,         2462],
        }), hide_index=True)


# ============================================================
# FOOTER
# ============================================================
st.markdown("""
<div class="footer">
    💰 <strong>FinanzIA</strong> &nbsp;·&nbsp;
    Proyecto Final — Prompt Engineering para Programadores &nbsp;·&nbsp;
    Diplomatura en Data Science — CoderHouse 2026
</div>
""", unsafe_allow_html=True)
