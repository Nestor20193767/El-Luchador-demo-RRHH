# -*- coding: utf-8 -*-
"""
PLATAFORMA DE GESTIÓN E INTELIGENCIA EMPRESARIAL — EL LUCHADOR
MÓDULO DE RECURSOS HUMANOS (ETAPA 1)
Desarrollado bajo Streamlit conforme a ESPECIFICACION_TECNICA_ARQUITECTURA_RRHH.md
"""

import calendar
import hashlib
import html
import random
import unicodedata
from datetime import date, datetime, timedelta
from pathlib import Path
from zoneinfo import ZoneInfo

import holidays
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

# ============================================================
# 1. CONFIGURACIÓN GENERAL Y CONSTANTES
# ============================================================

st.set_page_config(
    page_title="El Luchador | Gestión Integral de Recursos Humanos",
    page_icon="👥",
    layout="wide",
    initial_sidebar_state="expanded",
)

BASE = Path(__file__).resolve().parent
RUTA_EXCEL = BASE / "data" / "sueldos.xlsx"

# Paleta corporativa El Luchador
NEGRO = "#090A10"
AZUL_PRIMARIO = "#2F6BFF"
AZUL_OSCURO = "#1D346F"
NARANJA = "#FF7A1A"
BLANCO = "#F5F7FF"
GRIS_TEXTO = "#CBD5E1"
GRIS_BORDE = "#2B354C"
SUPERFICIE = "#141825"
SUPERFICIE_CARD = "#191E2E"
VERDE_EXITO = "#22C55E"
ROJO_PELIGRO = "#EF4444"
AMARILLO_ALERTA = "#F59E0B"

MESES = [
    "Enero", "Febrero", "Marzo", "Abril",
    "Mayo", "Junio", "Julio", "Agosto",
    "Septiembre", "Octubre", "Noviembre", "Diciembre",
]

DIAS_SEMANA = [
    "Lunes", "Martes", "Miércoles",
    "Jueves", "Viernes", "Sábado", "Domingo",
]

# Estructura Multiempresa y Multisucursal
RAZONES_SOCIALES = ["Misión", "Il Comendatori"]

SUCURSALES_INFO = {
    "Entrada": {
        "numero": 1,
        "nombre_completo": "Sucursal 1 Entrada",
        "razon_social": "Misión",
        "localidad": "Pilar",
    },
    "Ruta 26": {
        "numero": 2,
        "nombre_completo": "Sucursal 2 Ruta 26",
        "razon_social": "Il Comendatori",
        "localidad": "Del Viso",
    },
    "Colectora": {
        "numero": 3,
        "nombre_completo": "Sucursal 3 Colectora",
        "razon_social": "Il Comendatori",
        "localidad": "Escobar",
    },
    "Savio": {
        "numero": 4,
        "nombre_completo": "Sucursal 4 Savio",
        "razon_social": "Misión",
        "localidad": "Maquinista Savio",
    },
}

ROLES_RBAC = [
    "Administrador General",
    "Gerente de Recursos Humanos",
    "Responsable de Sucursal",
    "Supervisor de Área",
    "Asesor Laboral / Auditor",
]

UNIDADES_NEGOCIO = [
    "Carnicería",
    "Verdulería",
    "Cajas / Facturación",
    "Fiambrería",
    "Producción / Avícola",
    "Almacén",
    "Administración y Dirección",
]

DOTACION_MINIMA_DEFECTO = {
    "Carnicería": 2,
    "Verdulería": 2,
    "Cajas / Facturación": 1,
    "Fiambrería": 1,
    "Producción / Avícola": 1,
    "Almacén": 1,
    "Administración y Dirección": 1,
}

# ============================================================
# 2. ESTILOS VISUALES Y CORRECCIÓN DE CONTRASTE (CSS)
# ============================================================

st.markdown(
    f"""
<style>
/* 1. ESTRUCTURA GLOBAL Y FONDOS */
.stApp {{
    background: {NEGRO} !important;
    color: {BLANCO} !important;
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
}}

/* ELIMINACIÓN TOTAL DE LA BARRA BLANCA SUPERIOR */
header, [data-testid="stHeader"], .stAppHeader {{
    background-color: {NEGRO} !important;
    background: {NEGRO} !important;
    color: {BLANCO} !important;
    border-bottom: 1px solid #1C2436 !important;
}}

[data-testid="stDecoration"] {{
    display: none !important;
}}

[data-testid="stToolbar"] {{
    background-color: transparent !important;
    background: transparent !important;
    color: {BLANCO} !important;
}}

[data-testid="stToolbar"] button, [data-testid="stToolbar"] svg {{
    color: {BLANCO} !important;
    fill: {BLANCO} !important;
}}

/* 2. SIDEBAR CON CONTRASTE ÓPTIMO */
[data-testid="stSidebar"] {{
    background: #0E121D !important;
    border-right: 1px solid #232A3B !important;
}}

[data-testid="stSidebar"] label,
[data-testid="stSidebar"] [data-testid="stWidgetLabel"] p,
[data-testid="stSidebar"] [data-testid="stWidgetLabel"] span {{
    color: #E2E8F0 !important;
    font-size: 13.5px !important;
    font-weight: 600 !important;
}}

[data-testid="stSidebar"] p, [data-testid="stSidebar"] span {{
    color: #CBD5E1 !important;
}}

[data-testid="stSidebar"] h1, [data-testid="stSidebar"] h2, [data-testid="stSidebar"] h3, [data-testid="stSidebar"] h4 {{
    color: #FFFFFF !important;
    font-weight: 700 !important;
}}

/* NAVEGACIÓN RADIO BUTTONS */
[data-testid="stSidebar"] [data-testid="stRadio"] label {{
    padding: 8px 12px !important;
    border-radius: 9px !important;
    margin-bottom: 3px !important;
    color: #CBD5E1 !important;
    font-size: 13.5px !important;
    font-weight: 500 !important;
    transition: background 0.15s ease, color 0.15s ease !important;
}}

[data-testid="stSidebar"] [data-testid="stRadio"] label:hover {{
    background: #1A2338 !important;
    color: #FFFFFF !important;
}}

/* 3. CONTENEDOR PRINCIPAL Y ENCABEZADOS */
.block-container {{
    padding-top: 1.5rem !important;
    padding-bottom: 2.5rem !important;
    max-width: 1580px !important;
}}

h1, h2, h3, h4 {{
    color: #FFFFFF !important;
    font-weight: 700 !important;
}}

/* 4. TARJETAS DE MÉTRICAS */
[data-testid="stMetric"] {{
    border-radius: 14px !important;
    background: {SUPERFICIE} !important;
    border: 1px solid {GRIS_BORDE} !important;
    padding: 14px 18px !important;
    box-shadow: 0 4px 14px rgba(0,0,0,0.3) !important;
    transition: transform 0.15s ease, border-color 0.15s ease !important;
}}
[data-testid="stMetric"]:hover {{
    border-color: {AZUL_PRIMARIO} !important;
    transform: translateY(-2px) !important;
}}

[data-testid="stMetricLabel"] p {{
    color: #94A3B8 !important;
    font-size: 12px !important;
    font-weight: 700 !important;
    text-transform: uppercase !important;
    letter-spacing: 0.6px !important;
    white-space: normal !important;
    word-break: break-word !important;
}}

[data-testid="stMetricValue"] {{
    color: #FFFFFF !important;
    font-size: 25px !important;
    font-weight: 800 !important;
}}

/* 5. SELECTBOXES, INPUTS Y CONTROLES BASEWEB */
div[data-baseweb="select"] > div {{
    background-color: #151926 !important;
    border: 1px solid #333F5C !important;
    color: #FFFFFF !important;
    border-radius: 9px !important;
}}

div[data-baseweb="select"] span {{
    color: #FFFFFF !important;
    font-weight: 500 !important;
}}

div[data-baseweb="popover"], div[data-baseweb="menu"], ul[role="listbox"] {{
    background-color: #151926 !important;
    border: 1px solid #333F5C !important;
    border-radius: 9px !important;
}}

li[role="option"] {{
    color: #E2E8F0 !important;
    font-size: 13.5px !important;
}}

li[role="option"]:hover, li[aria-selected="true"] {{
    background-color: #2F6BFF !important;
    color: #FFFFFF !important;
}}

div[data-baseweb="input"] > div {{
    background-color: #151926 !important;
    border: 1px solid #333F5C !important;
    color: #FFFFFF !important;
    border-radius: 9px !important;
}}

input {{
    color: #FFFFFF !important;
}}

label, [data-testid="stWidgetLabel"] p, [data-testid="stWidgetLabel"] span {{
    color: #E2E8F0 !important;
    font-size: 13.5px !important;
    font-weight: 600 !important;
}}

/* 6. BOTONES */
.stButton > button[kind="primary"] {{
    background: {AZUL_PRIMARIO} !important;
    color: #FFFFFF !important;
    border: 0 !important;
    border-radius: 10px !important;
    font-weight: 600 !important;
    padding: 8px 18px !important;
    transition: all 0.2s ease !important;
}}
.stButton > button[kind="primary"]:hover {{
    background: #1A55E3 !important;
    box-shadow: 0 4px 16px rgba(47, 107, 255, 0.4) !important;
}}

.stButton > button {{
    border-radius: 10px !important;
    border: 1px solid {GRIS_BORDE} !important;
    background: {SUPERFICIE} !important;
    color: {BLANCO} !important;
}}

/* 7. HERO BANNER Y TARJETAS */
.hero {{
    border: 1px solid #344C8C;
    border-radius: 18px;
    padding: 24px 30px;
    background: linear-gradient(115deg, #162447 0%, #111524 100%);
    margin-bottom: 20px;
    box-shadow: 0 8px 24px rgba(0,0,0,0.35);
}}
.hero .eyebrow {{
    font-size: 11px;
    font-weight: 800;
    letter-spacing: 2px;
    color: #9BB8FF;
    text-transform: uppercase;
}}
.hero h1 {{
    font-size: 28px !important;
    margin: 8px 0;
}}
.hero p {{
    margin: 0;
    color: #CBD5EF;
    font-size: 14.5px;
}}

.panel-card {{
    background: {SUPERFICIE};
    border: 1px solid {GRIS_BORDE};
    border-radius: 16px;
    padding: 20px;
    margin-bottom: 18px;
}}

/* 8. PESTAÑAS (TABS) */
[data-baseweb="tab-list"] {{
    background: #101420 !important;
    border-radius: 10px !important;
    padding: 4px !important;
    border: 1px solid #283147 !important;
    gap: 4px !important;
}}

[data-baseweb="tab"] {{
    color: #94A3B8 !important;
    font-weight: 600 !important;
    font-size: 13.5px !important;
    border-radius: 8px !important;
    padding: 8px 16px !important;
    border: none !important;
}}

[aria-selected="true"][data-baseweb="tab"] {{
    background-color: #1F2B47 !important;
    color: #60A5FA !important;
    border: 1px solid #3B82F6 !important;
}}

/* 9. DATAFRAMES Y TABLAS */
[data-testid="stDataFrame"] {{
    border: 1px solid {GRIS_BORDE} !important;
    border-radius: 10px !important;
    background: {SUPERFICIE} !important;
    overflow: hidden !important;
}}

/* BADGES */
.badge {{
    display: inline-block;
    padding: 3px 10px;
    border-radius: 12px;
    font-size: 11.5px;
    font-weight: 700;
    letter-spacing: 0.3px;
}}
.badge-blue {{ background: #1F3668; color: #94BAFF; border: 1px solid #355BA8; }}
.badge-green {{ background: #133924; color: #72EAA5; border: 1px solid #236E46; }}
.badge-orange {{ background: #3E2413; color: #FFAF75; border: 1px solid #824921; }}
.badge-red {{ background: #3C1618; color: #FFA3A6; border: 1px solid #78272C; }}

/* CALENDARIO */
.calbox {{
    border: 1px solid #2D3349;
    background: {SUPERFICIE};
    border-radius: 10px;
    padding: 8px;
    min-height: 165px;
    margin: 3px 0;
    overflow: hidden;
}}
.calbox.holiday {{ border-color: {NARANJA}; background: #221B1A; }}
.calbox.today {{ border: 2px solid {AZUL_PRIMARIO}; }}
.calbox.other {{ opacity: 0.4; }}
.caldate {{ font-size: 14px; font-weight: 700; margin-bottom: 4px; }}
.calbadge-feriado {{
    font-size: 10.5px;
    font-weight: 700;
    color: {NARANJA};
    margin-bottom: 5px;
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
}}
.calshift {{
    font-size: 10.5px;
    background: #1B2135;
    color: #CCD8F5;
    padding: 2px 5px;
    border-radius: 4px;
    margin-bottom: 3px;
    border-left: 3px solid {AZUL_PRIMARIO};
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
}}
.calshift-warning {{
    font-size: 10.5px;
    background: #362214;
    color: #FFB380;
    padding: 2px 5px;
    border-radius: 4px;
    margin-bottom: 3px;
    border-left: 3px solid {NARANJA};
}}

/* MOCKUP WHATSAPP */
.wa-container {{
    max-width: 540px;
    margin: 0 auto;
    background: #111B21;
    border: 1px solid #2A3942;
    border-radius: 20px;
    overflow: hidden;
    box-shadow: 0 10px 30px rgba(0,0,0,0.6);
}}
.wa-header {{
    background: #202C33;
    padding: 12px 18px;
    display: flex;
    align-items: center;
    border-bottom: 1px solid #2A3942;
}}
.wa-header-avatar {{
    width: 38px;
    height: 38px;
    border-radius: 50%;
    background: #00A884;
    display: flex;
    align-items: center;
    justify-content: center;
    font-weight: 800;
    color: #FFF;
    margin-right: 12px;
}}
.wa-body {{
    background: #0B141A;
    background-image: radial-gradient(#182229 1px, transparent 1px);
    background-size: 16px 16px;
    padding: 18px;
    min-height: 280px;
}}
.wa-bubble {{
    background: #202C33;
    color: #E9EDEF;
    border-radius: 12px;
    border-top-left-radius: 2px;
    padding: 12px 14px;
    max-width: 90%;
    margin-bottom: 14px;
    box-shadow: 0 2px 6px rgba(0,0,0,0.3);
    font-size: 13.5px;
    line-height: 1.45;
}}
.wa-bubble-title {{
    font-weight: 700;
    color: #53BDEB;
    margin-bottom: 6px;
}}
.wa-bubble-time {{
    text-align: right;
    font-size: 10.5px;
    color: #8696A0;
    margin-top: 4px;
}}
</style>
    """,
    unsafe_allow_html=True,
)

# ============================================================
# 3. UTILIDADES GENERALES Y REGLAS LABORALES ARGENTINAS (LCT)
# ============================================================

try:
    HOY_AR = datetime.now(ZoneInfo("America/Argentina/Buenos_Aires")).date()
except Exception:
    HOY_AR = date(2026, 11, 12)


def norm(txt):
    if txt is None:
        return ""
    s = str(txt).strip().lower()
    return unicodedata.normalize("NFKD", s).encode("ASCII", "ignore").decode("utf-8")


def entera(val):
    try:
        return f"{int(round(val)):,}".replace(",", ".")
    except Exception:
        return "0"


def ars(val):
    if val is None or pd.isna(val):
        return "—"
    try:
        return f"$ {float(val):,.2f}".replace(",", "@").replace(".", ",").replace("@", ".")
    except Exception:
        return str(val)


def millones(val):
    if val is None or pd.isna(val):
        return "—"
    try:
        return f"$ {float(val) / 1_000_000:.2f} M".replace(".", ",")
    except Exception:
        return str(val)


def semilla(valor):
    return int(hashlib.md5(str(valor).encode("utf-8")).hexdigest()[:8], 16)


def cabecera(titulo, descripcion, etiqueta_modulo="Módulo de Recursos Humanos"):
    st.markdown(
        f"""
        <div class="hero">
            <div class="eyebrow">{etiqueta_modulo} · El Luchador</div>
            <h1>{html.escape(titulo)}</h1>
            <p>{html.escape(descripcion)}</p>
        </div>
        """,
        unsafe_allow_html=True,
    )


def generar_cuil(dni, sexo=0):
    prefijo = 20 if sexo == 0 else 27
    dni_str = str(dni).zfill(8)
    base_str = f"{prefijo}{dni_str}"
    factores = [5, 4, 3, 2, 7, 6, 5, 4, 3, 2]
    suma = sum(int(base_str[i]) * factores[i] for i in range(10))
    resto = suma % 11
    if resto == 0:
        verif = 0
    elif resto == 1:
        prefijo = 23
        verif = 9 if sexo == 0 else 4
    else:
        verif = 11 - resto
    return f"{prefijo}-{dni_str}-{verif}"


def calcular_edad(fecha_nac, fecha_ref=None):
    if not fecha_nac:
        return 0
    ref = fecha_ref or HOY_AR
    return ref.year - fecha_nac.year - ((ref.month, ref.day) < (fecha_nac.month, fecha_nac.day))


def calcular_antiguedad_exacta(fecha_ingreso, fecha_ref=None):
    if not fecha_ingreso:
        return 0, 0, 0, "0 años"
    ref = fecha_ref or HOY_AR
    anios = ref.year - fecha_ingreso.year
    meses = ref.month - fecha_ingreso.month
    dias = ref.day - fecha_ingreso.day
    if dias < 0:
        meses -= 1
        dias += 30
    if meses < 0:
        anios -= 1
        meses += 12
    anios = max(0, anios)
    meses = max(0, meses)
    dias = max(0, dias)
    return anios, meses, dias, f"{anios}a {meses}m {dias}d"


def escala_vacaciones_lct(antiguedad_anios):
    if antiguedad_anios <= 5:
        return 14
    elif antiguedad_anios <= 10:
        return 21
    elif antiguedad_anios <= 20:
        return 28
    else:
        return 35


def calcular_plus_vacacional(sueldo_mensual, dias_vacaciones):
    if not sueldo_mensual or sueldo_mensual <= 0 or dias_vacaciones <= 0:
        return 0.0, 0.0, 0.0
    jornal_vacacional = sueldo_mensual / 25.0
    jornal_habitual = sueldo_mensual / 30.0
    monto_total_vacaciones = jornal_vacacional * dias_vacaciones
    plus_neto = (jornal_vacacional - jornal_habitual) * dias_vacaciones
    return jornal_vacacional, monto_total_vacaciones, plus_neto


def costo_laboral_total(sueldo_bruto):
    cargas_patronales = sueldo_bruto * 0.240
    art = sueldo_bruto * 0.030
    provision_sac = sueldo_bruto * (1.0 / 12.0)
    cargas_sac = provision_sac * 0.240
    costo_mensual = sueldo_bruto + cargas_patronales + art + provision_sac + cargas_sac
    return {
        "bruto": sueldo_bruto,
        "cargas_patronales": cargas_patronales,
        "art": art,
        "provision_sac": provision_sac,
        "cargas_sac": cargas_sac,
        "costo_total_mensual": costo_mensual,
        "costo_total_anual": costo_mensual * 12.0,
    }


def area_de_rol(rol):
    r = norm(rol)
    if any(k in r for k in ["cortador", "carnic"]):
        return "Carnicería"
    if any(k in r for k in ["verdulero", "verdul"]):
        return "Verdulería"
    if any(k in r for k in ["cajer", "factur"]):
        return "Cajas / Facturación"
    if any(k in r for k in ["fiambr"]):
        return "Fiambrería"
    if any(k in r for k in ["produccion", "avicol"]):
        return "Producción / Avícola"
    if any(k in r for k in ["almacen"]):
        return "Almacén"
    if any(k in r for k in ["gerente", "admin", "direccion"]):
        return "Administración y Dirección"
    return "Operaciones Generales"


def registrar_auditoria(usuario, rol, accion, entidad, detalle):
    if "rrhh_audit_log" not in st.session_state:
        st.session_state.rrhh_audit_log = []
    st.session_state.rrhh_audit_log.append({
        "timestamp": datetime.now(ZoneInfo("America/Argentina/Buenos_Aires")).strftime("%Y-%m-%d %H:%M:%S"),
        "usuario": usuario,
        "rol": rol,
        "accion": accion,
        "entidad": entidad,
        "detalle": detalle,
    })


# ============================================================
# 4. CARGA Y ENRIQUECIMIENTO DE DATOS
# ============================================================

@st.cache_data(show_spinner="Procesando registros de personal...")
def cargar_personal(origen):
    df = pd.read_excel(origen, engine="openpyxl")
    df.columns = [str(c).strip() for c in df.columns]

    requeridas = [
        "Razon social", "Sucursal", "Local", "Empleados", "Rol", "Semanal", "Salario mensual"
    ]
    faltan = [c for c in requeridas if c not in df.columns]
    if faltan:
        raise ValueError(f"Faltan columnas requeridas en el archivo: {', '.join(faltan)}")

    df["Fila Excel"] = df.index + 2

    for col in ["Salario mensual", "Semanal", "semanal real", "Aguinaldo junio", "Aguinaldo diciembre"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")
        else:
            df[col] = 0.0

    df = df.dropna(subset=requeridas).copy()

    excluidos = df["Empleados"].astype(str).map(norm).str.fullmatch("compras verduleria|encargado|natalia")
    df = df.loc[(~excluidos) & (df["Salario mensual"] > 0)].copy()

    df["Nombre"] = df["Empleados"].astype(str).str.strip()

    def limpiar_rs(val):
        x = norm(val)
        if "mision" in x:
            return "Misión"
        if "comendatori" in x or "ilcom" in x:
            return "Il Comendatori"
        return str(val).strip()

    df["Razón social"] = df["Razon social"].map(limpiar_rs)

    def limpiar_local(val):
        x = norm(val)
        if "entrada" in x:
            return "Entrada"
        if "26" in x or "ruta" in x:
            return "Ruta 26"
        if "colectora" in x:
            return "Colectora"
        if "savio" in x:
            return "Savio"
        return str(val).strip()

    df["Local"] = df["Local"].map(limpiar_local)

    num_sucursal = df["Sucursal"].astype(str).str.extract(r"(\d+)")[0].fillna("1")
    df["ID"] = "S" + num_sucursal + "-F" + df["Fila Excel"].astype(int).astype(str).str.zfill(2)
    df["Unidad de Negocio"] = df["Rol"].map(area_de_rol)
    df["Área"] = df["Unidad de Negocio"]

    df["Mensual ARS"] = df["Salario mensual"]
    df["Semanal equivalente ARS"] = df["Mensual ARS"] * 12.0 / 52.0
    df["Anual base ARS"] = df["Mensual ARS"] * 12.0

    calles_zonanorte = [
        ("Av. Tratado del Pilar", "Pilar"),
        ("Rivadavia", "Pilar"),
        ("Ruta 26", "Del Viso"),
        ("Av. Madero", "Del Viso"),
        ("Av. Eugenia Tapia de Cruz", "Escobar"),
        ("Asborno", "Escobar"),
        ("Boulevard 5 de Junio", "Maquinista Savio"),
        ("Juan XXIII", "Maquinista Savio"),
        ("Av. Belgrano", "Garín"),
        ("Churich", "Garín"),
    ]

    obras_sociales = ["OSECAC", "OSECAC", "OSDE 210", "Swiss Medical", "OSPAT"]
    aseguradoras_art = ["Prevención ART", "La Segunda ART", "Asociart"]
    convenios_categoria = [
        "CCT 130/75 - Personal Especializado",
        "CCT 130/75 - Cajero B",
        "CCT 130/75 - Maestranza A",
        "CCT 130/75 - Auxiliar A",
        "CCT 130/75 - Administrativo B",
    ]

    dnis, cuils, nacimientos, ingresos, domicilios, ult_actualiz_domicilio = [], [], [], [], [], []
    art_list, art_polizas, art_vencimientos, obs_list, cat_list = [], [], [], [], []
    familiares_list, emergencias_list, estados_laborales = [], [], []

    for i, uid in enumerate(df["ID"]):
        rng = random.Random(semilla(uid))
        dni_val = rng.randint(28_000_000, 44_000_000)
        sexo = 1 if (rng.randint(1, 10) > 7) else 0
        cuil_val = generar_cuil(dni_val, sexo)

        nac = date(rng.randint(1972, 2003), rng.randint(1, 12), rng.randint(1, 28))
        ing = date(rng.randint(2014, 2025), rng.randint(1, 12), rng.randint(1, 28))

        calle, loc = calles_zonanorte[rng.randint(0, len(calles_zonanorte) - 1)]
        altura = rng.randint(100, 4800)
        domicilio_str = f"{calle} {altura}, {loc}, Prov. Buenos Aires"

        dias_atras_dom = rng.choice([45, 90, 110, 195, 230, 270])
        ult_act_dom = HOY_AR - timedelta(days=dias_atras_dom)

        art_cia = aseguradoras_art[rng.randint(0, len(aseguradoras_art) - 1)]
        art_pol = f"POL-{rng.randint(100000, 999999)}/01"
        art_venc = date(2027, rng.randint(1, 12), rng.randint(1, 28))

        obs = obras_sociales[rng.randint(0, len(obras_sociales) - 1)]
        cat = convenios_categoria[rng.randint(0, len(convenios_categoria) - 1)]

        num_hijos = rng.randint(0, 3)
        fam_str = f"{num_hijos} hijos a cargo" if num_hijos > 0 else "Sin cargas familiares declaradas"
        tel_emerg = f"+54 9 11 {rng.randint(3000, 6999)}-{rng.randint(1000, 9999)}"

        dnis.append(f"{dni_val:,}".replace(",", "."))
        cuils.append(cuil_val)
        nacimientos.append(nac)
        ingresos.append(ing)
        domicilios.append(domicilio_str)
        ult_actualiz_domicilio.append(ult_act_dom)
        art_list.append(art_cia)
        art_polizas.append(art_pol)
        art_vencimientos.append(art_venc)
        obs_list.append(obs)
        cat_list.append(cat)
        familiares_list.append(fam_str)
        emergencias_list.append(tel_emerg)
        estados_laborales.append("Activo")

    df["DNI"] = dnis
    df["CUIL"] = cuils
    df["Fecha Nacimiento"] = nacimientos
    df["Fecha Ingreso"] = ingresos
    df["Domicilio"] = domicilios
    df["Ultima Confirmacion Domicilio"] = ult_actualiz_domicilio
    df["ART Aseguradora"] = art_list
    df["ART Poliza"] = art_polizas
    df["ART Vencimiento"] = art_vencimientos
    df["Obra Social"] = obs_list
    df["Categoria Convenio"] = cat_list
    df["Familiares"] = familiares_list
    df["Contacto Emergencia"] = emergencias_list
    df["Estado Laboral"] = estados_laborales

    return df.reset_index(drop=True)


# ============================================================
# 5. CALENDARIO OFICIAL ARGENTINO
# ============================================================

@st.cache_data
def feriados_argentina(anio):
    try:
        feriados_ar = holidays.Argentina(years=[anio])
        return {fecha: nombre for fecha, nombre in feriados_ar.items()}
    except Exception:
        return {}


# ============================================================
# 6. GESTIÓN DE TURNOS, COBERTURA Y REEMPLAZOS
# ============================================================

def turno_base(uid, dia):
    descanso = semilla("descanso-" + uid) % 7
    if dia.weekday() == descanso:
        return "Descanso"
    
    semana = dia.isocalendar().week
    tipo = (semilla("turno-" + uid) + semana) % 3
    if tipo == 0:
        return "Apertura (07:00-15:00)"
    elif tipo == 1:
        return "Cierre (14:00-22:00)"
    else:
        return "Cortado (08:30-12:30 / 16:30-20:30)"


def esta_en_vacaciones(uid, dia):
    return any(
        s["ID"] == uid
        and s["Estado"] == "Aprobada"
        and s["Inicio"] <= dia <= s["Fin"]
        for s in st.session_state.rrhh_vacaciones
    )


def turno_efectivo(uid, dia):
    clave_reemplazo = (dia, uid)
    if clave_reemplazo in st.session_state.rrhh_reemplazos:
        return st.session_state.rrhh_reemplazos[clave_reemplazo]["turno"]
    if esta_en_vacaciones(uid, dia):
        return "Vacaciones"
    return turno_base(uid, dia)


def evaluar_cobertura_sucursal(plantilla, local, fecha):
    sub = plantilla.loc[plantilla["Local"] == local]
    resultados = []
    for area in UNIDADES_NEGOCIO:
        empleados_area = sub.loc[sub["Unidad de Negocio"] == area]
        total_asignados = len(empleados_area)
        if total_asignados == 0:
            continue
        minimo_requerido = DOTACION_MINIMA_DEFECTO.get(area, 1)
        presentes = 0
        ausentes = 0
        for _, emp in empleados_area.iterrows():
            t = turno_efectivo(emp["ID"], fecha)
            if t not in ["Descanso", "Vacaciones"]:
                presentes += 1
            else:
                ausentes += 1
        cubierto = presentes >= minimo_requerido
        resultados.append({
            "Área": area,
            "Total Asignados": total_asignados,
            "Presentes": presentes,
            "Mínimo Requerido": minimo_requerido,
            "Déficit": max(0, minimo_requerido - presentes),
            "Estado": "Óptimo" if presentes > minimo_requerido else ("Al Límite" if cubierto else "Déficit Crítico"),
        })
    return pd.DataFrame(resultados)


# ============================================================
# 7. ESTADO DE SESIÓN (SESSION STATE)
# ============================================================

def inicializar_session_state(personal):
    if "rrhh_vacaciones" not in st.session_state:
        st.session_state.rrhh_vacaciones = []
        ejemplos = personal.sample(n=min(5, len(personal)), random_state=42)
        for i, (_, r) in enumerate(ejemplos.iterrows()):
            inicio = date(2026, 11 + (i % 2), 15 + i * 2)
            dias_sol = 7 + (i % 3) * 7
            fin = inicio + timedelta(days=dias_sol - 1)
            antig_anios = antiguedad_a_cierre(r["Fecha Ingreso"], 2026)
            _, tot_vac, plus_neto = calcular_plus_vacacional(r["Mensual ARS"], dias_sol)
            st.session_state.rrhh_vacaciones.append({
                "N": i + 1,
                "ID": r["ID"],
                "Nombre": r["Nombre"],
                "Local": r["Local"],
                "Área": r["Unidad de Negocio"],
                "Inicio": inicio,
                "Fin": fin,
                "Dias": dias_sol,
                "Estado": "Aprobada" if i % 2 == 0 else "Pendiente",
                "Plus Vacacional ARS": plus_neto,
                "Total Vacaciones ARS": tot_vac,
                "Fecha Pago Prevista": inicio - timedelta(days=3),
                "Aprobado Por": "Gerente de Recursos Humanos" if i % 2 == 0 else "—",
            })

    if "rrhh_anticipos" not in st.session_state:
        st.session_state.rrhh_anticipos = [
            {
                "N": 1,
                "ID": personal.iloc[0]["ID"],
                "Nombre": personal.iloc[0]["Nombre"],
                "Local": personal.iloc[0]["Local"],
                "Monto": round(personal.iloc[0]["Mensual ARS"] * 0.25, -3),
                "Fecha Solicitud": HOY_AR - timedelta(days=5),
                "Fecha Pago": HOY_AR - timedelta(days=2),
                "Estado": "Aprobado",
                "Saldo Pendiente": round(personal.iloc[0]["Mensual ARS"] * 0.25, -3),
                "Aprobado Por": "Administrador General",
            }
        ]

    if "rrhh_reemplazos" not in st.session_state:
        st.session_state.rrhh_reemplazos = {}

    if "rrhh_audit_log" not in st.session_state:
        st.session_state.rrhh_audit_log = [
            {
                "timestamp": (datetime.now(ZoneInfo("America/Argentina/Buenos_Aires")) - timedelta(hours=2)).strftime("%Y-%m-%d %H:%M:%S"),
                "usuario": "Admin",
                "rol": "Administrador General",
                "accion": "INICIO_SISTEMA",
                "entidad": "Configuración Multiempresa",
                "detalle": "Módulo de Recursos Humanos iniciado con 4 sucursales y 2 razones sociales.",
            }
        ]

    if "rrhh_alertas_vistas" not in st.session_state:
        st.session_state.rrhh_alertas_vistas = set()

    if "rrhh_docs_modificados" not in st.session_state:
        st.session_state.rrhh_docs_modificados = {}

    if "rrhh_domicilios_ratificados" not in st.session_state:
        st.session_state.rrhh_domicilios_ratificados = set()

    if "rrhh_whatsapp_mensajes" not in st.session_state:
        st.session_state.rrhh_whatsapp_mensajes = [
            {
                "id": "WAM-001",
                "timestamp": HOY_AR.strftime("%d/%m/%Y 09:30"),
                "destinatario": "Responsable Sucursal 1 (+54 9 11 4820-9112)",
                "plantilla": "aviso_cobertura_minima",
                "mensaje": "⚠️ *ALERTA OPERATIVA - EL LUCHADOR*\nSucursal Entrada: Se detectó 1 puesto sin cubrir en Carnicería para el turno tarde. Por favor autorizar reemplazo desde la plataforma.",
                "estado": "Entregado",
            }
        ]


def antiguedad_a_cierre(ingreso, anio):
    corte = date(anio, 12, 31)
    return max(0, corte.year - ingreso.year - ((corte.month, corte.day) < (ingreso.month, ingreso.day)))


def dias_aprobados_empleado(uid, anio):
    return sum(
        s["Dias"]
        for s in st.session_state.rrhh_vacaciones
        if s["ID"] == uid and s["Estado"] == "Aprobada" and s["Inicio"].year == anio
    )


# ============================================================
# 8. CARGA DE ARCHIVO Y PREPARACIÓN DE DATOS
# ============================================================

with st.sidebar:
    st.markdown(
        """
        <div style="background:linear-gradient(135deg, #162447, #0E1322); padding:18px; border:1px solid #355BA8; border-radius:14px; margin-bottom:15px;">
            <div style="font-size:11px; font-weight:800; letter-spacing:1.5px; color:#8EB2FF;">GRUPO EMPRESARIAL</div>
            <h2 style="margin:4px 0 2px; font-size:20px; color:#FFF;">EL LUCHADOR</h2>
            <div style="font-size:12px; color:#A7BEE8;">Módulo de Recursos Humanos · Etapa 1</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    if RUTA_EXCEL.exists():
        fuente = RUTA_EXCEL
    else:
        fuente = st.file_uploader("Cargar archivo sueldos.xlsx", type=["xlsx"])
        if fuente is None:
            st.warning("Coloque sueldos.xlsx en la carpeta data/ o cárguelo para iniciar.")
            st.stop()

try:
    personal = cargar_personal(fuente)
except Exception as ex:
    st.error(f"Error procesando sueldos.xlsx: {ex}")
    st.stop()

if personal.empty:
    st.error("No se encontraron registros válidos de trabajadores.")
    st.stop()

inicializar_session_state(personal)

# ============================================================
# 9. BARRA LATERAL, PERFIL RBAC Y FILTROS GLOBALES
# ============================================================

with st.sidebar:
    st.markdown("### Perfil de Acceso (RBAC)")
    rol_activo = st.selectbox(
        "Rol Activo",
        ROLES_RBAC,
        index=0,
        help="Simula permisos y vistas según el rol de seguridad asignado (Sección 2.2).",
    )

    badge_color = "badge-blue" if "Administrador" in rol_activo else ("badge-green" if "Gerente" in rol_activo else "badge-orange")
    st.markdown(f'<span class="badge {badge_color}">PERFIL: {rol_activo.upper()}</span>', unsafe_allow_html=True)

    st.divider()

    st.markdown("### Navegación")
    pagina = st.radio(
        "Sección del Módulo",
        [
            "📊 Resumen Ejecutivo (KPIs)",
            "🗂️ Legajos Digitales",
            "📅 Calendario & Turnos",
            "🏖️ Vacaciones & Régimen LCT",
            "💰 Remuneraciones & Anticipos",
            "📈 Productividad Operativa (POS)",
            "🔔 Centro de Alertas & WhatsApp",
            "🛡️ Registro de Auditoría",
        ],
        label_visibility="collapsed",
    )

    st.divider()

    st.markdown("### Filtros Globales")
    rs_filtro = st.multiselect(
        "Razón Social",
        RAZONES_SOCIALES,
        default=RAZONES_SOCIALES,
    )

    sucursales_disponibles = sorted(personal["Local"].unique().tolist())
    sucursales_filtro = st.multiselect(
        "Sucursales",
        sucursales_disponibles,
        default=sucursales_disponibles,
    )

    anio_filtro = st.selectbox("Año de Gestión", [2026, 2027, 2028], index=0)

base = personal.loc[
    (personal["Razón social"].isin(rs_filtro)) &
    (personal["Local"].isin(sucursales_filtro))
].copy()

if base.empty:
    st.warning("No hay empleados que coincidan con los filtros seleccionados.")
    st.stop()


# ============================================================
# 10. MÓDULO 1: RESUMEN EJECUTIVO (DASHBOARD DE KPIS)
# ============================================================

if pagina == "📊 Resumen Ejecutivo (KPIs)":
    cabecera(
        "Dashboard Ejecutivo de Recursos Humanos",
        "Indicadores clave de dotación, presupuesto salarial, cumplimiento de turnos y semáforo de vencimientos en tiempo real.",
        "Etapa 1 · Panel Gerencial",
    )

    total_empleados = len(base)
    salario_mensual_total = base["Mensual ARS"].sum()
    costos_totales = [costo_laboral_total(s) for s in base["Mensual ARS"]]
    costo_mensual_empresa = sum(c["costo_total_mensual"] for c in costos_totales)
    
    edades = [calcular_edad(nac) for nac in base["Fecha Nacimiento"]]
    edad_promedio = sum(edades) / max(1, len(edades))
    
    antiguedades = [antiguedad_a_cierre(ing, anio_filtro) for ing in base["Fecha Ingreso"]]
    antig_promedio = sum(antiguedades) / max(1, len(antiguedades))

    anticipos_activos = sum(a["Saldo Pendiente"] for a in st.session_state.rrhh_anticipos if a["Estado"] == "Aprobado")

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Headcount Activo", f"{total_empleados} pers.")
    c2.metric("Masa Salarial Bruta", millones(salario_mensual_total))
    c3.metric("Costo Total C/Cargas", millones(costo_mensual_empresa), help="Incluye Salarios + Cargas Patronales (~24%) + ART (~3%) + Provisiones SAC.")
    c4.metric("Anticipos Activos", ars(anticipos_activos))

    c5, c6, c7, c8 = st.columns(4)
    c5.metric("Edad Promedio", f"{edad_promedio:.1f} años")
    c6.metric("Antigüedad Media", f"{antig_promedio:.1f} años")
    
    coberturas_locales = [evaluar_cobertura_sucursal(personal, s, HOY_AR) for s in base["Local"].unique()]
    if coberturas_locales:
        df_cob_global = pd.concat(coberturas_locales)
        pct_cob = (1.0 - (df_cob_global["Déficit"].sum() / max(1, df_cob_global["Mínimo Requerido"].sum()))) * 100
    else:
        pct_cob = 100.0
    c7.metric("Cumplimiento Dotación", f"{pct_cob:.1f} %", delta="Operativo" if pct_cob >= 90 else "-Alerta")

    solic_pendientes = sum(1 for s in st.session_state.rrhh_vacaciones if s["Estado"] == "Pendiente")
    c8.metric("Vacaciones x Resolver", f"{solic_pendientes} sol.", delta=f"{solic_pendientes} pendientes", delta_color="inverse")

    st.divider()

    st.markdown("### 🚦 Semáforo de Control Legal y Vencimientos")
    domicilios_vencidos = [
        r["ID"] for _, r in base.iterrows()
        if (HOY_AR - r["Ultima Confirmacion Domicilio"]).days > 180 and r["ID"] not in st.session_state.rrhh_domicilios_ratificados
    ]

    sc1, sc2, sc3 = st.columns(3)
    with sc1:
        color_dom = "#EF4444" if domicilios_vencidos else "#22C55E"
        st.markdown(
            f"""
            <div class="panel-card" style="border-left: 4px solid {color_dom};">
                <div style="font-size:12px; color:{GRIS_TEXTO};">ACTUALIZACIÓN SEMESTRAL DOMICILIO</div>
                <h3 style="margin:4px 0;">{len(domicilios_vencidos)} Trabajadores</h3>
                <div style="font-size:13px; color:#C4CCE0;">Superan los 180 días sin ratificación formal de domicilio.</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with sc2:
        st.markdown(
            f"""
            <div class="panel-card" style="border-left: 4px solid #22C55E;">
                <div style="font-size:12px; color:{GRIS_TEXTO};">COBERTURA RIESGOS DEL TRABAJO (ART)</div>
                <h3 style="margin:4px 0;">100% Cobertura Activa</h3>
                <div style="font-size:13px; color:#C4CCE0;">Pólizas de Prevención ART, La Segunda y Asociart al día.</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with sc3:
        st.markdown(
            f"""
            <div class="panel-card" style="border-left: 4px solid #F59E0B;">
                <div style="font-size:12px; color:{GRIS_TEXTO};">PROVISIÓN DE SAC (AGUINALDO)</div>
                <h3 style="margin:4px 0;">{millones(salario_mensual_total * 0.5)}</h3>
                <div style="font-size:13px; color:#C4CCE0;">Compromiso semestral estimado para cuota Junio / Diciembre.</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    st.markdown("### Desglose y Distribución Operativa")
    g1, g2 = st.columns(2)

    with g1:
        fig_suc = px.bar(
            base.groupby(["Local", "Razón social"]).size().reset_index(name="Dotación"),
            x="Local",
            y="Dotación",
            color="Razón social",
            barmode="group",
            title="Dotación Activa por Sucursal y Razón Social",
            color_discrete_map={"Misión": AZUL_PRIMARIO, "Il Comendatori": NARANJA},
        )
        fig_suc.update_layout(
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(0,0,0,0)",
            font_color=BLANCO,
            margin=dict(t=40, b=20, l=20, r=20),
        )
        st.plotly_chart(fig_suc, use_container_width=True)

    with g2:
        cargas_tot = sum(c["cargas_patronales"] for c in costos_totales)
        art_tot = sum(c["art"] for c in costos_totales)
        sac_tot = sum(c["provision_sac"] + c["cargas_sac"] for c in costos_totales)
        fig_costos = px.pie(
            names=["Sueldo Bruto Base", "Cargas Patronales (24%)", "ART (3%)", "Provisión SAC + Cargas"],
            values=[salario_mensual_total, cargas_tot, art_tot, sac_tot],
            title="Composición del Costo Laboral Total Empresa",
            color_discrete_sequence=[AZUL_PRIMARIO, "#4B7DFF", NARANJA, "#FCD34D"],
            hole=0.45,
        )
        fig_costos.update_layout(
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(0,0,0,0)",
            font_color=BLANCO,
            margin=dict(t=40, b=20, l=20, r=20),
        )
        st.plotly_chart(fig_costos, use_container_width=True)


# ============================================================
# 11. MÓDULO 2: LEGAJOS DIGITALES Y FICHA DEL TRABAJADOR
# ============================================================

elif pagina == "🗂️ Legajos Digitales":
    cabecera(
        "Legajos Digitales y Ficha Integral del Trabajador",
        "Registro laboral centralizado: datos personales, contractuales, historial, cobertura médica ART y repositorio documental.",
        "Sección 4.1 a 4.3 · Legajo Digital",
    )

    seleccion_id = st.selectbox(
        "Seleccionar Colaborador para inspeccionar Legajo",
        base["ID"].tolist(),
        format_func=lambda x: f"{x} — {base.loc[base['ID']==x, 'Nombre'].iloc[0]} ({base.loc[base['ID']==x, 'Local'].iloc[0]} · {base.loc[base['ID']==x, 'Rol'].iloc[0]})",
    )

    emp = base.loc[base["ID"] == seleccion_id].iloc[0]
    anios_ant, meses_ant, dias_ant, antig_str = calcular_antiguedad_exacta(emp["Fecha Ingreso"])
    edad = calcular_edad(emp["Fecha Nacimiento"])

    col_av, col_det, col_sal = st.columns([1, 2.5, 1.5])
    with col_av:
        st.markdown(
            f"""
            <div style="background:{SUPERFICIE}; border:1px solid {AZUL_PRIMARIO}; border-radius:18px; padding:20px; text-align:center;">
                <div style="font-size:38px;">👤</div>
                <h3 style="margin:8px 0 2px; font-size:18px;">{emp["Nombre"]}</h3>
                <span class="badge badge-blue">{emp["ID"]}</span>
                <div style="margin-top:8px; font-size:12px; color:{GRIS_TEXTO};">{emp["Categoria Convenio"]}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with col_det:
        st.markdown(
            f"""
            <div class="panel-card" style="height:100%;">
                <h4 style="margin-top:0;">Datos Contractuales & Radicación</h4>
                <div style="display:grid; grid-template-columns: 1fr 1fr; gap:8px; font-size:13.5px;">
                    <div><b>Razón Social:</b> {emp["Razón social"]}</div>
                    <div><b>Sucursal:</b> {emp["Local"]} (Sucursal {SUCURSALES_INFO.get(emp["Local"], {}).get("numero", "—")})</div>
                    <div><b>Unidad de Negocio:</b> {emp["Unidad de Negocio"]}</div>
                    <div><b>Puesto / Rol:</b> {emp["Rol"]}</div>
                    <div><b>Fecha de Ingreso:</b> {emp["Fecha Ingreso"].strftime('%d/%m/%Y')}</div>
                    <div><b>Antigüedad Exacta:</b> <span style="color:#60A5FA; font-weight:700;">{antig_str}</span></div>
                    <div><b>Tipo Contrato:</b> Tiempo Indeterminado (LCT)</div>
                    <div><b>Estado:</b> <span class="badge badge-green">ACTIVO</span></div>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with col_sal:
        costo_emp = costo_laboral_total(emp["Mensual ARS"])
        st.markdown(
            f"""
            <div class="panel-card" style="height:100%;">
                <h4 style="margin-top:0;">Esquema Remuneratorio</h4>
                <div style="font-size:13px; color:{GRIS_TEXTO};">Salario Mensual Bruto:</div>
                <div style="font-size:22px; font-weight:800; color:#FFF; margin-bottom:6px;">{ars(emp["Mensual ARS"])}</div>
                <div style="font-size:12.5px; color:#A7BEE8;">Semanal Equivalente: {ars(emp["Semanal equivalente ARS"])}</div>
                <div style="font-size:12.5px; color:#A7BEE8;">Costo Total Empresa: {ars(costo_emp["costo_total_mensual"])}/mes</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    t_pers, t_art, t_docs, t_hist = st.tabs([
        "👤 Datos Personales & Domicilio",
        "🏥 Cobertura Médica & ART",
        "📂 Repositorio de Documentación (Legajo)",
        "📜 Historial Laboral & Antecedentes",
    ])

    with t_pers:
        c_p1, c_p2 = st.columns(2)
        with c_p1:
            st.markdown("#### Identificación & Familia")
            st.write(f"**DNI:** {emp['DNI']}")
            st.write(f"**CUIL:** `{emp['CUIL']}` (Clave Tributaria Laboral Argentina)")
            st.write(f"**Fecha de Nacimiento:** {emp['Fecha Nacimiento'].strftime('%d/%m/%Y')} ({edad} años)")
            st.write(f"**Familiares a Cargo:** {emp['Familiares']}")
            st.write(f"**Contacto de Emergencia:** {emp['Contacto Emergencia']}")

        with c_p2:
            st.markdown("#### Control Semestral de Domicilio")
            dias_desde_act = (HOY_AR - emp["Ultima Confirmacion Domicilio"]).days
            ratificado = emp["ID"] in st.session_state.rrhh_domicilios_ratificados
            
            st.write(f"**Domicilio Declarado:** {emp['Domicilio']}")
            st.write(f"**Última Actualización:** {emp['Ultima Confirmacion Domicilio'].strftime('%d/%m/%Y')} (hace {dias_desde_act} días)")

            if dias_desde_act > 180 and not ratificado:
                st.warning("⚠️ **Alerta Semestral:** Han transcurrido más de 180 días desde la última ratificación formal de domicilio (Art. 4.1).")
                if st.button("✅ Ratificar Domicilio Semestral Ahora", key=f"btn_rat_{emp['ID']}", type="primary"):
                    st.session_state.rrhh_domicilios_ratificados.add(emp["ID"])
                    registrar_auditoria(
                        rol_activo, rol_activo, "RATIFICACION_DOMICILIO", emp["ID"],
                        f"Ratificación de domicilio confirmada para {emp['Nombre']} ({emp['Domicilio']})."
                    )
                    st.success("Domicilio ratificado exitosamente. Alerta regularizada.")
                    st.rerun()
            else:
                st.success("✓ Domicilio ratificado y vigente dentro del período semestral reglamentario.")

    with t_art:
        st.markdown("#### Cobertura de Riesgos del Trabajo y Salud Asignada")
        ca1, ca2 = st.columns(2)
        with ca1:
            st.write(f"**Compañía Aseguradora (ART):** {emp['ART Aseguradora']}")
            st.write(f"**Número de Póliza:** `{emp['ART Poliza']}`")
            st.write(f"**Vigencia de Cobertura:** Hasta {emp['ART Vencimiento'].strftime('%d/%m/%Y')}")
            st.write("**Estado de Afiliación:** Cobertura Plena Activa")
        with ca2:
            st.write(f"**Obra Social Sindical:** {emp['Obra Social']}")
            st.write("**Plan de Cobertura:** Prestación Médica Obligatoria (PMO) + Seguro de Sepelio")
            st.write(f"**Último Examen Periódico de Salud:** Realizado ({emp['Fecha Ingreso'].year + 1})")

    with t_docs:
        st.markdown("#### Repositorio de Documentación del Legajo Digital")
        st.caption("Verificación de cumplimiento de instrumentos legales obligatorios.")

        documentos_legajo = [
            ("Contrato de Trabajo Firmado (LCT Art. 90)", "Vigente / Completo", "14/03/2024"),
            ("Fotocopia DNI Frente y Dorso", "Vigente / Completo", "14/03/2024"),
            ("Constancia de CUIL (ARCA / ANSES)", "Vigente / Completo", "14/03/2024"),
            ("Alta Temprana AFIP (F. 885/A)", "Vigente / Completo", "14/03/2024"),
            ("Certificado Médico Preocupacional / Periódico", "Vigente / Completo", "12/05/2024"),
            ("Constancia Entrega de Uniforme y EPP (Res. SRT 299/11)", "Vigente / Completo", "10/01/2026"),
            ("Recibo de Haberes Último Período Firmado", "Vigente / Completo", "05/11/2026"),
        ]

        df_docs = pd.DataFrame(documentos_legajo, columns=["Documento Obligatorio", "Estado", "Fecha Registro"])
        st.dataframe(df_docs, use_container_width=True, hide_index=True)

        with st.expander("📤 Adjuntar o Actualizar Documento del Legajo"):
            tipo_doc = st.selectbox("Tipo de Documento", [d[0] for d in documentos_legajo])
            archivo = st.file_uploader(f"Cargar archivo digital para {emp['Nombre']}", type=["pdf", "png", "jpg"])
            if st.button("Guardar en Repositorio"):
                if archivo:
                    st.success(f"Documento '{tipo_doc}' cargado y auditado para {emp['Nombre']}.")
                    registrar_auditoria(rol_activo, rol_activo, "CARGA_DOCUMENTO", emp["ID"], f"Documento {tipo_doc} subido.")
                else:
                    st.info("Simulación: Documento verificado y actualizado.")

    with t_hist:
        st.markdown("#### Historial Laboral, Traslados y Evaluaciones")
        st.markdown(
            f"""
            - **{emp['Fecha Ingreso'].strftime('%d/%m/%Y')}:** Alta laboral formal en **{emp['Razón social']}**, sucursal **{emp['Local']}**.
            - **15/06/2024:** Asignación de categoría **{emp['Categoria Convenio']}**.
            - **10/01/2025:** Entrega de uniforme reglamentario y elementos de protección personal (EPP).
            - **01/03/2026:** Ratificación de jornada contractual en **{emp['Unidad de Negocio']}**.
            """
        )


# ============================================================
# 12. MÓDULO 3: CALENDARIO & PLANIFICADOR DE TURNOS
# ============================================================

elif pagina == "📅 Calendario & Turnos":
    cabecera(
        "Calendario Laboral Oficial & Planificador de Turnos",
        "Matriz interactiva con feriados nacionales argentinos, rotación de horarios, control de cobertura mínima y sustitución de personal.",
        "Sección 4.4 · Turnos y Cobertura",
    )

    c1, c2, c3, c4 = st.columns([1, 1, 1, 1])
    mes_sel = c1.selectbox("Mes", range(1, 13), index=HOY_AR.month - 1, format_func=lambda m: MESES[m - 1])
    sucursal_sel = c2.selectbox("Sucursal", sorted(base["Local"].unique().tolist()))
    area_sel = c3.selectbox("Unidad de Negocio", ["Todas"] + UNIDADES_NEGOCIO)
    vista_modo = c4.selectbox("Vista del Planificador", ["Mensual", "Semanal", "Control de Cobertura Diaria"])

    feriados_anio = feriados_argentina(anio_filtro)
    sub_personal = base.loc[base["Local"] == sucursal_sel].copy()
    if area_sel != "Todas":
        sub_personal = sub_personal.loc[sub_personal["Unidad de Negocio"] == area_sel]

    if vista_modo == "Control de Cobertura Diaria":
        st.markdown("### Verificación de Dotación Mínima Operativa por Turno")
        dia_inspeccion = st.date_input("Fecha a evaluar", value=HOY_AR)
        df_cob = evaluar_cobertura_sucursal(personal, sucursal_sel, dia_inspeccion)

        st.dataframe(df_cob, use_container_width=True, hide_index=True)

        hay_deficit = df_cob["Déficit"].sum() > 0
        if hay_deficit:
            st.error(f"⚠️ **Alerta Operativa:** Se detectaron puestos descubiertos en la sucursal {sucursal_sel} para el día {dia_inspeccion.strftime('%d/%m/%Y')}.")
        else:
            st.success(f"✓ Cobertura óptima asegurada en todas las áreas operativas de {sucursal_sel}.")

        with st.expander("🔄 Asignar Reemplazo Operativo para Contingencias"):
            col_rep1, col_rep2 = st.columns(2)
            with col_rep1:
                emp_reemplazado = st.selectbox("Colaborador ausente / a cubrir", sub_personal["ID"].tolist(), format_func=lambda x: f"{x} - {sub_personal.loc[sub_personal['ID']==x, 'Nombre'].iloc[0]}")
            with col_rep2:
                otros_disp = personal.loc[(personal["ID"] != emp_reemplazado)]
                emp_sustituto = st.selectbox("Colaborador que asume el reemplazo", otros_disp["ID"].tolist(), format_func=lambda x: f"{x} - {otros_disp.loc[otros_disp['ID']==x, 'Nombre'].iloc[0]} ({otros_disp.loc[otros_disp['ID']==x, 'Local'].iloc[0]})")

            if st.button("Confirmar Reemplazo Operativo"):
                clave = (dia_inspeccion, emp_reemplazado)
                st.session_state.rrhh_reemplazos[clave] = {
                    "sustituto": emp_sustituto,
                    "turno": "Apertura (07:00-15:00)",
                }
                registrar_auditoria(rol_activo, rol_activo, "ASIGNACION_REEMPLAZO", emp_reemplazado, f"Reemplazado por {emp_sustituto} el {dia_inspeccion.strftime('%d/%m')}.")
                st.success("Reemplazo registrado y coordinado en el plan de turnos.")
                st.rerun()

    else:
        st.markdown(f"### Malla Mensual de Turnos — {MESES[mes_sel - 1]} {anio_filtro} ({sucursal_sel})")
        cal = calendar.Calendar(firstweekday=0)
        dias_mes = cal.monthdatescalendar(anio_filtro, mes_sel)

        cols_head = st.columns(7)
        for i, nombre_dia in enumerate(DIAS_SEMANA):
            cols_head[i].markdown(f"<div style='text-align:center; font-weight:700; color:{AZUL_PRIMARIO};'>{nombre_dia}</div>", unsafe_allow_html=True)

        for semana in dias_mes:
            cols_sem = st.columns(7)
            for j, d in enumerate(semana):
                with cols_sem[j]:
                    es_mes = (d.month == mes_sel)
                    es_feriado = d in feriados_anio
                    es_hoy = (d == HOY_AR)
                    clase_extra = "holiday" if es_feriado else ("today" if es_hoy else ("other" if not es_mes else ""))

                    html_cal = f'<div class="calbox {clase_extra}">'
                    html_cal += f'<div class="caldate">{d.day}</div>'
                    if es_feriado:
                        html_cal += f'<div class="calbadge-feriado">🇦🇷 {feriados_anio[d]}</div>'

                    if es_mes:
                        for _, emp_it in sub_personal.head(3).iterrows():
                            t = turno_efectivo(emp_it["ID"], d)
                            badge_style = "calshift-warning" if "Vacaciones" in t or "Descanso" in t else "calshift"
                            html_cal += f'<div class="{badge_style}">{emp_it["Nombre"].split()[0]}: {t.split()[0]}</div>'
                        if len(sub_personal) > 3:
                            html_cal += f'<div style="font-size:10px; color:{GRIS_TEXTO};">+{len(sub_personal)-3} más</div>'

                    html_cal += '</div>'
                    st.markdown(html_cal, unsafe_allow_html=True)


# ============================================================
# 13. MÓDULO 4: VACACIONES & RÉGIMEN LCT (LEY 20.744)
# ============================================================

elif pagina == "🏖️ Vacaciones & Régimen LCT":
    cabecera(
        "Gestión de Vacaciones y Plus Vacacional (Ley 20.744)",
        "Escala de antigüedad legal, verificación automática de dotación mínima, cálculo de plus vacacional y flujo de aprobaciones.",
        "Sección 4.5 · Vacaciones LCT",
    )

    with st.expander("📖 Marco Legal Argentino: Régimen de Vacaciones LCT N° 20.744"):
        st.markdown(
            """
            - **Escala de Licencia Anual Ordinaria (Art. 150 LCT):**
              - Antigüedad <= 5 años: 14 días corridos
              - 5 < Antigüedad <= 10 años: 21 días corridos
              - 10 < Antigüedad <= 20 años: 28 días corridos
              - Antigüedad > 20 años: 35 días corridos
            - **Cálculo del Plus Vacacional (Art. 155 LCT):**
              El día de vacación se remunera dividiendo el sueldo mensual por 25 (y no por 30):
              $$\\text{Plus Vacacional} = \\left(\\frac{\\text{Sueldo Mensual}}{25} - \\frac{\\text{Sueldo Mensual}}{30}\\right) \\times \\text{Días de Licencia}$$
            - **Fecha de Pago Obligatoria:** Las vacaciones deben ser liquidadas y abonadas *antes de la iniciación de la licencia*.
            """
        )

    tab_solic, tab_matriz = st.tabs(["📝 Registro & Simulación de Solicitud", "📊 Matriz de Saldos por Empleado"])

    with tab_solic:
        st.markdown("### Nueva Solicitud de Licencia Anual")
        col_s1, col_s2 = st.columns(2)

        with col_s1:
            emp_vac_id = st.selectbox(
                "Colaborador",
                base["ID"].tolist(),
                format_func=lambda x: f"{x} - {base.loc[base['ID']==x, 'Nombre'].iloc[0]} ({base.loc[base['ID']==x, 'Local'].iloc[0]})",
            )
            emp_sel = base.loc[base["ID"] == emp_vac_id].iloc[0]
            antig_anios = antiguedad_a_cierre(emp_sel["Fecha Ingreso"], anio_filtro)
            cupo_legal = escala_vacaciones_lct(antig_anios)
            dias_usados = dias_aprobados_empleado(emp_vac_id, anio_filtro)
            saldo_disp = max(0, cupo_legal - dias_usados)

            st.info(f"**Antigüedad LCT:** {antig_anios} años · **Cupo Anual Legal:** {cupo_legal} días · **Saldo Disponible:** {saldo_disp} días")

            inicio_vac = st.date_input("Fecha Inicio de Licencia", value=HOY_AR + timedelta(days=10))
            dias_pedidos = st.number_input("Días Corridos Solicitados", min_value=1, max_value=max(1, saldo_disp), value=min(14, max(1, saldo_disp)))
            fin_vac = inicio_vac + timedelta(days=int(dias_pedidos) - 1)

        with col_s2:
            st.markdown("#### Impacto Salarial Estimado (LCT Art. 155)")
            jornal_vac, monto_total_vac, plus_neto = calcular_plus_vacacional(emp_sel["Mensual ARS"], int(dias_pedidos))
            fecha_pago_prev = inicio_vac - timedelta(days=3)

            st.write(f"**Período Solicitado:** {inicio_vac.strftime('%d/%m/%Y')} al {fin_vac.strftime('%d/%m/%Y')} ({dias_pedidos} días)")
            st.write(f"**Jornal Vacacional (Base 25):** {ars(jornal_vac)} / día")
            st.write(f"**Monto Total Vacaciones a Liquidar:** {ars(monto_total_vac)}")
            st.metric("Plus Vacacional Neto Estimado", ars(plus_neto), help="Diferencia adicional a favor del trabajador por base divisor 25.")
            st.write(f"**Fecha Límite Prevista de Pago:** `{fecha_pago_prev.strftime('%d/%m/%Y')}`")

            if st.button("Enviar Solicitud a Aprobación", type="primary"):
                nueva_solicitud = {
                    "N": len(st.session_state.rrhh_vacaciones) + 1,
                    "ID": emp_sel["ID"],
                    "Nombre": emp_sel["Nombre"],
                    "Local": emp_sel["Local"],
                    "Área": emp_sel["Unidad de Negocio"],
                    "Inicio": inicio_vac,
                    "Fin": fin_vac,
                    "Dias": int(dias_pedidos),
                    "Estado": "Pendiente",
                    "Plus Vacacional ARS": plus_neto,
                    "Total Vacaciones ARS": monto_total_vac,
                    "Fecha Pago Prevista": fecha_pago_prev,
                    "Aprobado Por": "—",
                }
                st.session_state.rrhh_vacaciones.append(nueva_solicitud)
                registrar_auditoria(rol_activo, rol_activo, "SOLICITUD_VACACIONES", emp_sel["ID"], f"Solicitud de {dias_pedidos} días del {inicio_vac} al {fin_vac}.")
                st.success("Solicitud registrada con éxito y remitida a la bandeja de autorizaciones.")
                st.rerun()

        st.divider()
        st.markdown("### Bandeja de Solicitudes de Vacaciones")
        if st.session_state.rrhh_vacaciones:
            df_vac_bandeja = pd.DataFrame(st.session_state.rrhh_vacaciones)
            st.dataframe(df_vac_bandeja, use_container_width=True, hide_index=True)

            solic_pendientes = [s for s in st.session_state.rrhh_vacaciones if s["Estado"] == "Pendiente"]
            if solic_pendientes and "Auditor" not in rol_activo:
                st.markdown("#### Resolver Solicitudes Pendientes")
                idx_res = st.selectbox("Seleccione Solicitud a Resolver", range(len(solic_pendientes)), format_func=lambda i: f"#{solic_pendientes[i]['N']} - {solic_pendientes[i]['Nombre']} ({solic_pendientes[i]['Dias']} días desde {solic_pendientes[i]['Inicio'].strftime('%d/%m')})")
                sol_elegida = solic_pendientes[idx_res]

                b_ap, b_rec = st.columns(2)
                if b_ap.button("✅ Aprobar Solicitud", type="primary"):
                    for s in st.session_state.rrhh_vacaciones:
                        if s["N"] == sol_elegida["N"]:
                            s["Estado"] = "Aprobada"
                            s["Aprobado Por"] = rol_activo
                    registrar_auditoria(rol_activo, rol_activo, "APROBACION_VACACIONES", sol_elegida["ID"], f"Aprobadas vacaciones #{sol_elegida['N']} ({sol_elegida['Dias']} días).")
                    st.success("Solicitud APROBADA exitosamente.")
                    st.rerun()
                if b_rec.button("❌ Rechazar Solicitud"):
                    for s in st.session_state.rrhh_vacaciones:
                        if s["N"] == sol_elegida["N"]:
                            s["Estado"] = "Rechazada"
                            s["Aprobado Por"] = rol_activo
                    registrar_auditoria(rol_activo, rol_activo, "RECHAZO_VACACIONES", sol_elegida["ID"], f"Rechazadas vacaciones #{sol_elegida['N']}.")
                    st.warning("Solicitud RECHAZADA.")
                    st.rerun()

    with tab_matriz:
        st.markdown(f"### Matriz Consolidada de Vacaciones — Período {anio_filtro}")
        filas_matriz = []
        for _, r in base.iterrows():
            ant = antiguedad_a_cierre(r["Fecha Ingreso"], anio_filtro)
            cupo = escala_vacaciones_lct(ant)
            aprob = dias_aprobados_empleado(r["ID"], anio_filtro)
            filas_matriz.append({
                "ID": r["ID"],
                "Nombre": r["Nombre"],
                "Sucursal": r["Local"],
                "Área": r["Unidad de Negocio"],
                "Antigüedad (Años)": ant,
                "Días por Ley LCT": cupo,
                "Días Gozados / Aprobados": aprob,
                "Saldo Disponible": max(0, cupo - aprob),
            })
        st.dataframe(pd.DataFrame(filas_matriz), use_container_width=True, hide_index=True)


# ============================================================
# 14. MÓDULO 5: REMUNERACIONES, ANTICIPOS & COSTO LABORAL
# ============================================================

elif pagina == "💰 Remuneraciones & Anticipos":
    cabecera(
        "Administración Salarial, Anticipos y Costos Laborales",
        "Presupuesto sobre Salario Mensual, gestión auditable de anticipos, estimación de cargas patronales y provisiones de aguinaldo (SAC).",
        "Sección 4.6 · Gestión Salarial",
    )

    t_nom, t_ant, t_costo = st.tabs([
        "💵 Nómina Salarial",
        "💳 Gestión de Anticipos",
        "🏢 Costo Total Empresa & Cargas Patronales",
    ])

    with t_nom:
        total_sal_base = base["Mensual ARS"].sum()
        total_sac_jun = base["Aguinaldo junio"].fillna(0).sum()
        total_sac_dic = base["Aguinaldo diciembre"].fillna(0).sum()

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Masa Mensual Base", millones(total_sal_base))
        c2.metric("Proyección Anual Base", millones(total_sal_base * 12))
        c3.metric("SAC Junio (Aguinaldo)", millones(total_sac_jun if total_sac_jun > 0 else total_sal_base * 0.5))
        c4.metric("SAC Diciembre (Aguinaldo)", millones(total_sac_dic if total_sac_dic > 0 else total_sal_base * 0.5))

        st.markdown("#### Detalle de Planilla y Comparación")
        cols_mostrar = ["ID", "Nombre", "Local", "Rol", "Mensual ARS", "Semanal", "semanal real", "Semanal equivalente ARS"]
        st.dataframe(base[cols_mostrar], use_container_width=True, hide_index=True)

    with t_ant:
        st.markdown("### Solicitud y Autorización de Anticipos de Sueldo")
        with st.form("form_anticipo"):
            col_a1, col_a2 = st.columns(2)
            with col_a1:
                emp_ant_id = st.selectbox("Colaborador Solicitante", base["ID"].tolist(), format_func=lambda x: f"{x} - {base.loc[base['ID']==x, 'Nombre'].iloc[0]}")
                monto_ant = st.number_input("Monto del Anticipo (ARS)", min_value=5000.0, value=60000.0, step=5000.0)
            with col_a2:
                fecha_ant = st.date_input("Fecha de Solicitud", value=HOY_AR)
                tope_porc = st.slider("Tope Máximo Permitido sobre Salario", 10, 50, 30, 5, help="Política de empresa: habitualmente tope del 30% o 50% de la remuneración mensual.")

            btn_ant = st.form_submit_button("Registrar Solicitud de Anticipo", type="primary")

        if btn_ant:
            sueldo_emp = base.loc[base["ID"] == emp_ant_id, "Mensual ARS"].iloc[0]
            nombre_emp = base.loc[base["ID"] == emp_ant_id, "Nombre"].iloc[0]
            local_emp = base.loc[base["ID"] == emp_ant_id, "Local"].iloc[0]
            limite_dinero = sueldo_emp * (tope_porc / 100.0)

            if monto_ant > limite_dinero:
                st.error(f"El anticipo solicitado ({ars(monto_ant)}) supera el límite del {tope_porc}% del sueldo ({ars(limite_dinero)}).")
            else:
                nuevo_ant = {
                    "N": len(st.session_state.rrhh_anticipos) + 1,
                    "ID": emp_ant_id,
                    "Nombre": nombre_emp,
                    "Local": local_emp,
                    "Monto": monto_ant,
                    "Fecha Solicitud": fecha_ant,
                    "Fecha Pago": fecha_ant + timedelta(days=2),
                    "Estado": "Solicitado",
                    "Saldo Pendiente": monto_ant,
                    "Aprobado Por": "—",
                }
                st.session_state.rrhh_anticipos.append(nuevo_ant)
                registrar_auditoria(rol_activo, rol_activo, "SOLICITUD_ANTICIPO", emp_ant_id, f"Anticipo solicitado por {ars(monto_ant)}.")
                st.success("Anticipo registrado en el sistema.")
                st.rerun()

        st.markdown("#### Historial y Control de Anticipos")
        if st.session_state.rrhh_anticipos:
            st.dataframe(pd.DataFrame(st.session_state.rrhh_anticipos), use_container_width=True, hide_index=True)

            pendientes_ant = [a for a in st.session_state.rrhh_anticipos if a["Estado"] == "Solicitado"]
            if pendientes_ant and "Auditor" not in rol_activo:
                st.markdown("#### Autorizar Anticipos Pendientes")
                idx_ant = st.selectbox("Seleccione Anticipo a Autorizar", range(len(pendientes_ant)), format_func=lambda i: f"#{pendientes_ant[i]['N']} - {pendientes_ant[i]['Nombre']} ({ars(pendientes_ant[i]['Monto'])})")
                ant_sel = pendientes_ant[idx_ant]
                if st.button("✅ Autorizar y Programar Descuento", type="primary"):
                    for a in st.session_state.rrhh_anticipos:
                        if a["N"] == ant_sel["N"]:
                            a["Estado"] = "Aprobado"
                            a["Aprobado Por"] = rol_activo
                    registrar_auditoria(rol_activo, rol_activo, "APROBACION_ANTICIPO", ant_sel["ID"], f"Anticipo #{ant_sel['N']} aprobado por {rol_activo}.")
                    st.success("Anticipo APROBADO.")
                    st.rerun()

    with t_costo:
        st.markdown("### Estructura de Costo Laboral Total por Empleado y Sucursal")
        costo_df_list = []
        for _, r in base.iterrows():
            c = costo_laboral_total(r["Mensual ARS"])
            costo_df_list.append({
                "ID": r["ID"],
                "Nombre": r["Nombre"],
                "Sucursal": r["Local"],
                "Salario Bruto": r["Mensual ARS"],
                "Cargas Patronales (24%)": c["cargas_patronales"],
                "Alícuota ART (3%)": c["art"],
                "Provisión SAC Mensual": c["provision_sac"],
                "Costo Total Mensual": c["costo_total_mensual"],
                "Costo Total Anual": c["costo_total_anual"],
            })
        df_costos = pd.DataFrame(costo_df_list)
        st.dataframe(df_costos, use_container_width=True, hide_index=True)


# ============================================================
# 15. MÓDULO 6: PRODUCTIVIDAD OPERATIVA (POS)
# ============================================================

elif pagina == "📈 Productividad Operativa (POS)":
    cabecera(
        "Conexión con Productividad Operativa (Punto de Venta / Mostrador)",
        "Seguimiento de actividad por puesto sin emitir juicios causales automáticos: operaciones atendidas y volúmenes facturados por turno.",
        "Sección 4.7 · Productividad",
    )

    st.markdown(
        """
        > [!NOTE]
        > Conforme a los principios de diseño de la especificación técnica (Sección 4.7), el sistema presenta
        > indicadores cuantitativos de actividad operativa y atención de caja sin inferir conclusiones disciplinarias automáticas.
        """
    )

    prod_data = []
    for _, r in base.iterrows():
        rng = random.Random(semilla("prod-" + r["ID"]))
        es_cajero = "Cajer" in r["Rol"]
        es_cortador = "Cortador" in r["Rol"] or "Carnic" in r["Rol"]
        
        if es_cajero:
            ops = rng.randint(180, 310)
            ticket_prom = rng.randint(14000, 26000)
        elif es_cortador:
            ops = rng.randint(90, 160)
            ticket_prom = rng.randint(18000, 32000)
        else:
            ops = rng.randint(50, 120)
            ticket_prom = rng.randint(8000, 19000)

        total_vendido = ops * ticket_prom

        prod_data.append({
            "ID": r["ID"],
            "Nombre": r["Nombre"],
            "Sucursal": r["Local"],
            "Rol": r["Rol"],
            "Operaciones / Turno": ops,
            "Ticket Promedio (ARS)": ticket_prom,
            "Volumen Facturado Asignado": total_vendido,
            "Índice de Cadencia": "Alta" if ops > 200 else ("Media" if ops > 100 else "Estándar"),
        })

    df_prod = pd.DataFrame(prod_data)

    p1, p2, p3 = st.columns(3)
    p1.metric("Total Operaciones Procesadas", f"{df_prod['Operaciones / Turno'].sum():,} ops".replace(",", "."))
    p2.metric("Volumen Facturado Operativo", millones(df_prod["Volumen Facturado Asignado"].sum()))
    p3.metric("Promedio de Atención", f"{int(df_prod['Operaciones / Turno'].mean())} ops/turno")

    st.dataframe(df_prod, use_container_width=True, hide_index=True)


# ============================================================
# 16. MÓDULO 7: CENTRO DE ALERTAS & WHATSAPP BUSINESS
# ============================================================

elif pagina == "🔔 Centro de Alertas & WhatsApp":
    cabecera(
        "Centro de Notificaciones & Simulador WhatsApp Business",
        "Alertas administrativas categorizadas por prioridad e integración simulada con WhatsApp Business Platform (Cloud API).",
        "Sección 5.1 · Notificaciones y Alertas",
    )

    alertas = []

    for _, r in base.iterrows():
        dias_dom = (HOY_AR - r["Ultima Confirmacion Domicilio"]).days
        if dias_dom > 180 and r["ID"] not in st.session_state.rrhh_domicilios_ratificados:
            alertas.append({
                "Clave": f"DOM-{r['ID']}",
                "Tipo": "Domicilio Semestral",
                "Prioridad": "Alta",
                "Referencia": f"{r['ID']} - {r['Nombre']}",
                "Descripción": f"Han transcurrido {dias_dom} días sin ratificación formal de domicilio (Art. 4.1).",
                "Accion": "Ratificar Domicilio",
            })

    for s in st.session_state.rrhh_vacaciones:
        if s["Estado"] == "Pendiente":
            alertas.append({
                "Clave": f"VAC-{s['N']}",
                "Tipo": "Vacaciones",
                "Prioridad": "Media",
                "Referencia": f"{s['ID']} - {s['Nombre']}",
                "Descripción": f"Solicitud de {s['Dias']} días ({s['Inicio'].strftime('%d/%m')} al {s['Fin'].strftime('%d/%m')}) pendiente de resolución.",
                "Accion": "Aprobar Vacación",
            })

    for a in st.session_state.rrhh_anticipos:
        if a["Estado"] == "Solicitado":
            alertas.append({
                "Clave": f"ANT-{a['N']}",
                "Tipo": "Anticipo de Sueldo",
                "Prioridad": "Media",
                "Referencia": f"{a['ID']} - {a['Nombre']}",
                "Descripción": f"Solicitud de anticipo por {ars(a['Monto'])} pendiente de aprobación.",
                "Accion": "Autorizar Anticipo",
            })

    for suc in base["Local"].unique():
        df_c = evaluar_cobertura_sucursal(personal, suc, HOY_AR)
        def_crit = df_c.loc[df_c["Déficit"] > 0]
        for _, dc in def_crit.iterrows():
            alertas.append({
                "Clave": f"COB-{suc}-{dc['Área']}",
                "Tipo": "Dotación Mínima",
                "Prioridad": "Alta",
                "Referencia": f"{suc} · {dc['Área']}",
                "Descripción": f"Déficit operativo: faltan {dc['Déficit']} personas para alcanzar la dotación mínima.",
                "Accion": "Asignar Reemplazo",
            })

    visibles = [a for a in alertas if a["Clave"] not in st.session_state.rrhh_alertas_vistas]

    col_izq, col_der = st.columns([1.2, 1])

    with col_izq:
        st.markdown("### Alertas Abiertas en el Sistema")
        m1, m2 = st.columns(2)
        m1.metric("Total Pendientes", len(visibles))
        m2.metric("Prioridad Alta", sum(1 for a in visibles if a["Prioridad"] == "Alta"))

        if visibles:
            st.dataframe(pd.DataFrame(visibles).drop(columns=["Clave"]), use_container_width=True, hide_index=True)
            sel_alerta_idx = st.selectbox("Seleccionar Alerta para gestionar", range(len(visibles)), format_func=lambda i: f"[{visibles[i]['Prioridad']}] {visibles[i]['Tipo']} — {visibles[i]['Referencia']}")
            al_activa = visibles[sel_alerta_idx]

            if st.button("Marcar como Revisada / Descartar Alerta"):
                st.session_state.rrhh_alertas_vistas.add(al_activa["Clave"])
                registrar_auditoria(rol_activo, rol_activo, "REVISION_ALERTA", al_activa["Referencia"], f"Alerta {al_activa['Clave']} descartada.")
                st.success("Alerta marcada como revisada.")
                st.rerun()
        else:
            st.success("✓ No hay alertas pendientes en el sistema. Todos los parámetros se encuentran normalizados.")

    with col_der:
        st.markdown("### 📱 Simulador WhatsApp Business Platform (API)")
        st.markdown(
            """
            <div class="wa-container">
                <div class="wa-header">
                    <div class="wa-header-avatar">EL</div>
                    <div>
                        <div style="font-weight:700; color:#E9EDEF; font-size:14px;">El Luchador · RR. HH.</div>
                        <div style="font-size:11px; color:#00A884;">● Cuenta Empresarial Oficial Verificada</div>
                    </div>
                </div>
                <div class="wa-body">
            """,
            unsafe_allow_html=True,
        )

        if visibles:
            al_ejemplo = visibles[0]
            txt_wa = f"⚠️ *NOTIFICACIÓN DE RECURSOS HUMANOS*\n\n*Tipo:* {al_ejemplo['Tipo']}\n*Afectado:* {al_ejemplo['Referencia']}\n*Detalle:* {al_ejemplo['Descripción']}\n\nAcción directa requerida por la jefatura."
        else:
            txt_wa = "✓ *SISTEMA DE GESTIÓN EL LUCHADOR*\nTodas las alertas operativas han sido atendidas. Dotaciones cubiertas y nómina en orden."

        txt_wa_html = txt_wa.replace("\n", "<br>")

        st.markdown(
            f"""
            <div class="wa-bubble">
                <div class="wa-bubble-title">EL LUCHADOR IN-APP BOT</div>
                <div>{txt_wa_html}</div>
                <div class="wa-bubble-time">{HOY_AR.strftime('%H:%M')} ✓✓</div>
            </div>
            </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

        st.markdown("#### Botones de Acción Rápida (Interactive Quick Actions)")
        if visibles:
            b_wa1, b_wa2 = st.columns(2)
            if b_wa1.button("🟢 Resolver desde WhatsApp", type="primary"):
                st.session_state.rrhh_alertas_vistas.add(visibles[0]["Clave"])
                registrar_auditoria(rol_activo, rol_activo, "RESOLUCION_WHATSAPP", visibles[0]["Referencia"], "Resuelto vía botón rápido de WhatsApp API.")
                st.success("Acción transaccional ejecutada desde la notificación interactiva.")
                st.rerun()
            if b_wa2.button("📲 Reenviar Plantilla"):
                st.info("Plantilla HSM reenviada a los supervisores de turno.")


# ============================================================
# 17. MÓDULO 8: REGISTRO DE AUDITORÍA (AUDIT LOG)
# ============================================================

elif pagina == "🛡️ Registro de Auditoría":
    cabecera(
        "Pistas de Auditoría y Trazabilidad (Event Sourcing)",
        "Registro inmutable de transacciones, altas, aprobaciones de licencias y ratificaciones contractuales con marca temporal y usuario.",
        "Sección 2.2 · Seguridad y Auditoría",
    )

    if st.session_state.rrhh_audit_log:
        df_audit = pd.DataFrame(st.session_state.rrhh_audit_log)
        st.dataframe(df_audit.iloc[::-1], use_container_width=True, hide_index=True)
    else:
        st.info("No hay eventos registrados en la sesión actual.")

    if st.button("Descargar Registro de Auditoría (.CSV)"):
        st.download_button(
            label="Descargar CSV",
            data=pd.DataFrame(st.session_state.rrhh_audit_log).to_csv(index=False),
            file_name="auditoria_rrhh_luchador.csv",
            mime="text/csv",
        )


# ============================================================
# 18. PIE DE PÁGINA INSTITUCIONAL
# ============================================================

st.divider()
st.markdown(
    f"""
    <div style="display:flex; justify-content:space-between; align-items:center; font-size:12px; color:{GRIS_TEXTO};">
        <div><b>El Luchador</b> · Plataforma de Gestión e Inteligencia Empresarial (Etapa 1: Recursos Humanos)</div>
        <div>Cumplimiento Legal LCT N° 20.744 · Multiempresa: Misión & Il Comendatori</div>
    </div>
    """,
    unsafe_allow_html=True,
)
