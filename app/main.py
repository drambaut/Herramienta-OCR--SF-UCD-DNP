"""
Herramienta OCR-SF-UCD
Extracción automática de planillas de seguridad social colombianas.
Desarrollado para la Subdirección Financiera - Unidad de Costos y Dependencias (DNP).
"""

import os
import io
import json
import base64
import time
import zipfile
import tempfile
import pandas as pd
import fitz  
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
from azure.ai.documentintelligence import DocumentIntelligenceClient
from azure.core.credentials import AzureKeyCredential
from openai import AzureOpenAI
from dotenv import load_dotenv

load_dotenv()

# ─────────────────────────────────────────
# CONFIGURACIÓN
# ─────────────────────────────────────────
ENDPOINT                 = os.getenv("AZURE_FORM_RECOGNIZER_ENDPOINT")
API_KEY                  = os.getenv("AZURE_FORM_RECOGNIZER_KEY")
AZURE_OPENAI_ENDPOINT    = os.getenv("AZURE_OPENAI_ENDPOINT")
AZURE_OPENAI_API_KEY     = os.getenv("AZURE_OPENAI_API_KEY")
AZURE_OPENAI_API_VERSION = os.getenv("AZURE_OPENAI_API_VERSION", "2024-05-01-preview")
AZURE_OPENAI_DEPLOYMENT  = os.getenv("AZURE_OPENAI_DEPLOYMENT_NAME")

# ─────────────────────────────────────────
# ESTILOS GLOBALES
# ─────────────────────────────────────────
CUSTOM_CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;500;600&family=IBM+Plex+Sans:wght@300;400;600;700&display=swap');

/* ── Reset global ── */
html, body, [class*="css"] {
    font-family: 'IBM Plex Sans', sans-serif;
    background-color: #070c18;
    color: #c9d4e8;
}

/* ── Fondo app ── */
.stApp {
    background: linear-gradient(160deg, #070c18 0%, #0d1525 50%, #070c18 100%);
    min-height: 100vh;
}

/* ── Header principal ── */
.ocr-header {
    padding: 2.5rem 0 1.5rem 0;
    border-bottom: 1px solid #1a2744;
    margin-bottom: 2rem;
}
.ocr-header h1 {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 1.6rem;
    font-weight: 600;
    color: #ffffff;
    letter-spacing: -0.02em;
    margin: 0;
}
.ocr-header .subtitle {
    font-size: 0.82rem;
    color: #4a6090;
    letter-spacing: 0.12em;
    text-transform: uppercase;
    margin-top: 0.4rem;
}
.ocr-badge {
    display: inline-block;
    background: #0f2044;
    border: 1px solid #1e3a6e;
    color: #3b82f6;
    font-family: 'IBM Plex Mono', monospace;
    font-size: 0.68rem;
    padding: 0.2rem 0.7rem;
    border-radius: 2px;
    letter-spacing: 0.1em;
    margin-bottom: 0.8rem;
}

/* ── Upload zone ── */
.upload-zone {
    border: 1px dashed #1e3a6e;
    border-radius: 8px;
    padding: 2.5rem;
    text-align: center;
    background: #0a1022;
    transition: border-color 0.2s;
}
.upload-zone:hover { border-color: #3b82f6; }

/* ── Tarjetas KPI ── */
.kpi-card {
    background: linear-gradient(135deg, #0d1830 0%, #111e38 100%);
    border: 1px solid #1a2d55;
    border-radius: 10px;
    padding: 1.6rem 1.4rem;
    position: relative;
    overflow: hidden;
}
.kpi-card::before {
    content: '';
    position: absolute;
    top: 0; left: 0; right: 0;
    height: 2px;
    background: linear-gradient(90deg, #1e3a6e, #3b82f6, #1e3a6e);
}
.kpi-value {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 2.4rem;
    font-weight: 600;
    color: #3b82f6;
    line-height: 1;
    margin-bottom: 0.5rem;
}
.kpi-label {
    font-size: 0.68rem;
    letter-spacing: 0.14em;
    text-transform: uppercase;
    color: #4a6090;
    font-weight: 600;
}
.kpi-sub {
    font-size: 0.75rem;
    color: #2a4070;
    margin-top: 0.3rem;
    font-family: 'IBM Plex Mono', monospace;
}

/* ── Barra de progreso animada ── */
.progress-wrapper {
    background: #0a1022;
    border: 1px solid #1a2744;
    border-radius: 8px;
    padding: 1.8rem 2rem;
    margin: 1.5rem 0;
}
.progress-title {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 0.78rem;
    color: #3b82f6;
    letter-spacing: 0.1em;
    text-transform: uppercase;
    margin-bottom: 1rem;
}
.progress-track {
    background: #0d1525;
    border-radius: 4px;
    height: 6px;
    overflow: hidden;
    position: relative;
}
.progress-fill {
    height: 100%;
    border-radius: 4px;
    background: linear-gradient(90deg, #1d4ed8, #3b82f6, #60a5fa);
    background-size: 200% 100%;
    animation: shimmer 1.4s ease infinite;
    transition: width 0.6s ease;
}
@keyframes shimmer {
    0%   { background-position: 200% center; }
    100% { background-position: -200% center; }
}
.progress-pulse {
    display: flex;
    align-items: center;
    gap: 0.6rem;
    margin-top: 0.9rem;
}
.pulse-dot {
    width: 7px;
    height: 7px;
    background: #3b82f6;
    border-radius: 50%;
    animation: pulse 1.2s ease-in-out infinite;
    flex-shrink: 0;
}
@keyframes pulse {
    0%, 100% { opacity: 1; transform: scale(1); }
    50%       { opacity: 0.3; transform: scale(0.6); }
}
.pulse-dot:nth-child(2) { animation-delay: 0.2s; }
.pulse-dot:nth-child(3) { animation-delay: 0.4s; }
.progress-msg {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 0.75rem;
    color: #4a6090;
}
.progress-counter {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 0.72rem;
    color: #2a4070;
    margin-top: 0.5rem;
    text-align: right;
}

/* ── Sección de gráficas ── */
.section-title {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 0.72rem;
    letter-spacing: 0.14em;
    text-transform: uppercase;
    color: #2a4880;
    border-bottom: 1px solid #111e38;
    padding-bottom: 0.6rem;
    margin: 2rem 0 1.2rem 0;
}

/* ── Botón de descarga ── */
.stDownloadButton > button {
    background: linear-gradient(135deg, #1d4ed8, #2563eb) !important;
    color: white !important;
    border: none !important;
    border-radius: 6px !important;
    font-family: 'IBM Plex Mono', monospace !important;
    font-size: 0.78rem !important;
    letter-spacing: 0.08em !important;
    padding: 0.6rem 1.4rem !important;
    transition: all 0.2s !important;
    box-shadow: 0 0 20px rgba(59,130,246,0.15) !important;
}
.stDownloadButton > button:hover {
    background: linear-gradient(135deg, #2563eb, #3b82f6) !important;
    box-shadow: 0 0 30px rgba(59,130,246,0.3) !important;
    transform: translateY(-1px) !important;
}

/* ── File uploader ── */
[data-testid="stFileUploader"] {
    background: #0a1022;
    border: 1px dashed #1e3a6e;
    border-radius: 8px;
    padding: 1rem;
}
[data-testid="stFileUploader"]:hover {
    border-color: #3b82f6;
}

/* ── Alertas/info ── */
.stAlert {
    background: #0a1022 !important;
    border-color: #1a2744 !important;
    color: #c9d4e8 !important;
}

/* ── Tabla de datos ── */
.stDataFrame {
    border: 1px solid #1a2744;
    border-radius: 6px;
    overflow: hidden;
}

/* ── Sidebar ── */
[data-testid="stSidebar"] {
    background: #070c18;
    border-right: 1px solid #111e38;
}
[data-testid="stSidebar"] .stMarkdown {
    color: #4a6090;
}

/* ── Scrollbar ── */
::-webkit-scrollbar { width: 4px; }
::-webkit-scrollbar-track { background: #070c18; }
::-webkit-scrollbar-thumb { background: #1a2744; border-radius: 2px; }

/* ── Toast / success ── */
.success-banner {
    background: linear-gradient(135deg, #052e16, #064e3b);
    border: 1px solid #065f46;
    border-radius: 8px;
    padding: 1rem 1.4rem;
    color: #6ee7b7;
    font-family: 'IBM Plex Mono', monospace;
    font-size: 0.8rem;
    margin: 1rem 0;
}

/* ── Ocultar elementos Streamlit por defecto ── */
#MainMenu, footer, header { visibility: hidden; }
.block-container { padding-top: 1rem; max-width: 1100px; }
</style>
"""

# ─────────────────────────────────────────
# CLIENTES AZURE
# ─────────────────────────────────────────
@st.cache_resource
def get_clients():
    cliente_di = DocumentIntelligenceClient(ENDPOINT, AzureKeyCredential(API_KEY))
    cliente_oai = AzureOpenAI(
        azure_endpoint=AZURE_OPENAI_ENDPOINT,
        api_key=AZURE_OPENAI_API_KEY,
        api_version=AZURE_OPENAI_API_VERSION,
    )
    return cliente_di, cliente_oai

# ─────────────────────────────────────────
# LÓGICA OCR (misma del script batch)
# ─────────────────────────────────────────
def pdf_primera_pagina_a_base64(ruta_pdf: str, dpi: int = 150) -> str:
    doc  = fitz.open(ruta_pdf)
    pag  = doc[0]
    pix  = pag.get_pixmap(dpi=dpi)
    data = pix.tobytes("png")
    doc.close()
    return base64.standard_b64encode(data).decode("utf-8")


def extraer_texto_con_azure(ruta_pdf: str, cliente_di) -> tuple[str, int]:
    with open(ruta_pdf, "rb") as f:
        tarea     = cliente_di.begin_analyze_document(
            "prebuilt-layout", body=f, content_type="application/octet-stream"
        )
        resultado = tarea.result()

    bloques = []
    for pagina in resultado.pages:
        for linea in pagina.lines:
            bloques.append(linea.content)

    if resultado.tables:
        bloques.append("\n--- TABLAS DETECTADAS ---")
        for tabla in resultado.tables:
            for celda in tabla.cells:
                bloques.append(f"[Fila {celda.row_index}, Col {celda.column_index}]: {celda.content}")

    return "\n".join(bloques), len(resultado.pages)


SYSTEM_PROMPT = """
Eres un experto en planillas de seguridad social colombianas (PILA).
Recibes la imagen de la primera página y el texto extraído del documento.

Los 6 operadores de información más comunes en Colombia son:
  MiPlanilla (Compensar), Aportes en Línea, ARUS, SOI, Asopago, Simple.

Identifica el operador por el logo visible en la imagen O por el nombre en el texto.
Responde ÚNICAMENTE con un JSON válido, sin texto adicional, sin markdown.
"""

USER_PROMPT_TEMPLATE = """
Extrae exactamente estos campos del documento y devuelve solo el JSON:

{{
  "operador": "nombre del operador de información",
  "tipo_identificacion": "CC o NIT",
  "identificacion_aportante": "número de CC o NIT (solo dígitos, sin puntos ni guión)",
  "nombre_aportante": "nombre o razón social de quien paga",
  "fecha_pago": "fecha de pago en formato YYYY-MM-DD",
  "periodo": "periodo reportado, ej: 2026-01",
  "total_pagado": <número entero sin separadores de miles, ej: 1139500>,
  "moneda": "COP, USD u otra divisa detectada"
}}

Si un campo no puede determinarse con certeza, usa null.

TEXTO EXTRAÍDO DEL DOCUMENTO:
{texto}
"""

def extraer_campos_con_openai(texto: str, imagen_base64: str, cliente_oai) -> dict:
    prompt_usuario = USER_PROMPT_TEMPLATE.format(texto=texto[:5000])

    respuesta = cliente_oai.chat.completions.create(
        model    = AZURE_OPENAI_DEPLOYMENT,
        messages = [
            {"role": "system", "content": SYSTEM_PROMPT},
            {
                "role"   : "user",
                "content": [
                    {
                        "type"     : "image_url",
                        "image_url": {
                            "url"   : f"data:image/png;base64,{imagen_base64}",
                            "detail": "high",
                        },
                    },
                    {"type": "text", "text": prompt_usuario},
                ],
            },
        ],
        max_tokens      = 600,
        temperature     = 0,
        response_format = {"type": "json_object"},
    )

    raw = respuesta.choices[0].message.content.strip()
    raw = raw.replace("```json", "").replace("```", "").strip()

    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        return {k: None for k in [
            "operador","tipo_identificacion","identificacion_aportante",
            "nombre_aportante","fecha_pago","periodo","total_pagado","moneda"
        ]} | {"_error_json": raw[:200]}


def validar_datos(datos: dict) -> str:
    alertas = []
    id_num  = str(datos.get("identificacion_aportante") or "")
    if id_num and not id_num.isdigit():
        alertas.append("Identificación no numérica")
    if id_num and len(id_num) < 6:
        alertas.append("Identificación muy corta")

    total = datos.get("total_pagado")
    if total is not None:
        if total <= 0:
            alertas.append("Total cero o negativo")
        elif total > 5_000_000:
            alertas.append("Monto > 5M — Revisión Nivel 2")

    moneda = datos.get("moneda")
    if moneda and moneda.upper() != "COP":
        alertas.append(f"Moneda inusual: {moneda}")

    fecha_str = datos.get("fecha_pago")
    if fecha_str:
        try:
            fecha = datetime.strptime(fecha_str, "%Y-%m-%d")
            delta = (datetime.now() - fecha).days
            if delta > 365:
                alertas.append("Fecha > 12 meses")
            elif delta < 0:
                alertas.append("Fecha futura")
        except ValueError:
            alertas.append("Formato de fecha inválido")

    if not datos.get("operador"):
        alertas.append("Operador no identificado")
    if datos.get("_error_json"):
        alertas.append("Error parseo LLM — revisar")

    return ", ".join(alertas) if alertas else "Validado"


# ─────────────────────────────────────────
# COMPONENTES UI
# ─────────────────────────────────────────
def render_progress(current: int, total: int, filename: str, mensaje: str, slot):
    pct = int((current / total) * 100) if total > 0 else 0
    slot.markdown(f"""
    <div class="progress-wrapper">
        <div class="progress-title">⬡ Procesando archivos</div>
        <div class="progress-track">
            <div class="progress-fill" style="width:{pct}%"></div>
        </div>
        <div class="progress-pulse">
            <div class="pulse-dot"></div>
            <div class="pulse-dot"></div>
            <div class="pulse-dot"></div>
            <span class="progress-msg">{mensaje}</span>
        </div>
        <div class="progress-counter">{current} / {total} archivos &nbsp;·&nbsp; {pct}% completado</div>
        <div style="font-family:'IBM Plex Mono',monospace;font-size:0.65rem;color:#1e3a6e;margin-top:0.4rem;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;">
            ▶ {filename}
        </div>
    </div>
    """, unsafe_allow_html=True)


def kpi_card(value: str, label: str, sub: str = "") -> str:
    return f"""
    <div class="kpi-card">
        <div class="kpi-value">{value}</div>
        <div class="kpi-label">{label}</div>
        {"<div class='kpi-sub'>" + sub + "</div>" if sub else ""}
    </div>
    """


MENSAJES_PROGRESO = [
    "Enviando a Azure Document Intelligence…",
    "Extrayendo tablas y estructuras…",
    "Analizando imagen con GPT-4o…",
    "Identificando operador de información…",
    "Leyendo IBC, aportes y fechas…",
    "Aplicando validaciones de negocio…",
    "Consolidando en el registro…",
    "Verificando consistencia de datos…",
    "Cruzando campos con reglas DNP…",
    "Guardando resultado…",
]


# ─────────────────────────────────────────
# GRÁFICAS
# ─────────────────────────────────────────
PLOTLY_LAYOUT = dict(
    paper_bgcolor="rgba(0,0,0,0)",
    plot_bgcolor="rgba(0,0,0,0)",
    font=dict(family="IBM Plex Mono, monospace", color="#6b8cc4", size=11),
    margin=dict(l=10, r=10, t=40, b=10),
)

COLOR_PALETTE = [
    "#3b82f6","#1d4ed8","#60a5fa","#93c5fd","#2563eb",
    "#1e40af","#bfdbfe","#0ea5e9","#38bdf8","#7dd3fc",
]

def grafica_barras_proveedor(df: pd.DataFrame):
    conteo = df["operador"].fillna("Sin identificar").value_counts().reset_index()
    conteo.columns = ["Operador", "Cantidad"]
    fig = px.bar(
        conteo, x="Operador", y="Cantidad",
        color="Operador", color_discrete_sequence=COLOR_PALETTE,
        title="Planillas por Operador",
    )
    fig.update_traces(marker_line_width=0, opacity=0.9)
    fig.update_layout(**PLOTLY_LAYOUT,
        showlegend=False,
        title_font=dict(size=13, color="#3b82f6"),
        xaxis=dict(gridcolor="#0d1525", linecolor="#1a2744"),
        yaxis=dict(gridcolor="#0d1525", linecolor="#1a2744"),
    )
    return fig


def grafica_dona_validacion(df: pd.DataFrame):
    conteo = df["estado_validacion"].fillna("Sin datos").apply(
        lambda x: "Validado" if x == "Validado" else "Con alertas"
    ).value_counts().reset_index()
    conteo.columns = ["Estado", "Cantidad"]

    fig = go.Figure(go.Pie(
        labels=conteo["Estado"],
        values=conteo["Cantidad"],
        hole=0.62,
        marker=dict(colors=["#3b82f6","#1e3a6e"], line=dict(color="#070c18", width=2)),
        textfont=dict(family="IBM Plex Mono, monospace", size=11),
    ))
    fig.update_layout(**PLOTLY_LAYOUT,
        title=dict(text="Estado de Validación", font=dict(size=13, color="#3b82f6")),
        legend=dict(font=dict(color="#6b8cc4")),
        annotations=[dict(
            text=f"{conteo['Cantidad'].sum()}<br>total",
            x=0.5, y=0.5, font_size=13,
            font=dict(family="IBM Plex Mono", color="#3b82f6"),
            showarrow=False,
        )],
    )
    return fig


def grafica_pagos_proveedor(df: pd.DataFrame):
    df2 = df.dropna(subset=["total_pagado","operador"]).copy()
    df2["operador"] = df2["operador"].fillna("Sin identificar")
    agg = df2.groupby("operador")["total_pagado"].sum().reset_index()
    agg.columns = ["Operador", "Total_COP"]
    agg = agg.sort_values("Total_COP", ascending=True)
    agg["Total_M"] = (agg["Total_COP"] / 1_000_000).round(2)

    fig = px.bar(
        agg, x="Total_M", y="Operador", orientation="h",
        color="Total_M", color_continuous_scale=["#1e3a6e","#3b82f6","#93c5fd"],
        title="Pagos por Operador (millones COP)",
        labels={"Total_M": "Millones COP"},
    )
    fig.update_traces(marker_line_width=0)
    fig.update_layout(**PLOTLY_LAYOUT,
        coloraxis_showscale=False,
        showlegend=False,
        title_font=dict(size=13, color="#3b82f6"),
        xaxis=dict(gridcolor="#0d1525", linecolor="#1a2744"),
        yaxis=dict(gridcolor="rgba(0,0,0,0)", linecolor="#1a2744"),
    )
    return fig


# ─────────────────────────────────────────
# APP PRINCIPAL
# ─────────────────────────────────────────
def main():
    st.set_page_config(
        page_title="Herramienta OCR-SF-UCD",
        page_icon="⬡",
        layout="wide",
        initial_sidebar_state="collapsed",
    )
    st.markdown(CUSTOM_CSS, unsafe_allow_html=True)

    # ── Header ──────────────────────────
    st.markdown("""
    <div class="ocr-header">
        <div class="ocr-badge">DNP · SF · UCD · 2026</div>
        <h1>⬡ Herramienta OCR-SF-UCD</h1>
        <div class="subtitle">Extracción automática de planillas de seguridad social · Seguridad Social -- Colombia</div>
    </div>
    """, unsafe_allow_html=True)

    # ── Validación de credenciales ───────
    credenciales_ok = all([ENDPOINT, API_KEY, AZURE_OPENAI_ENDPOINT,
                           AZURE_OPENAI_API_KEY, AZURE_OPENAI_DEPLOYMENT])
    if not credenciales_ok:
        st.error(
            "⚠ Variables de entorno incompletas. "
            "Revisa AZURE_FORM_RECOGNIZER_ENDPOINT, AZURE_FORM_RECOGNIZER_KEY, "
            "AZURE_OPENAI_ENDPOINT, AZURE_OPENAI_API_KEY y AZURE_OPENAI_DEPLOYMENT_NAME."
        )
        st.stop()

    cliente_di, cliente_oai = get_clients()

    # ── Upload ───────────────────────────
    st.markdown('<div class="section-title">① Carga de archivos</div>', unsafe_allow_html=True)
    uploaded_zip = st.file_uploader(
        "Arrastra o selecciona un archivo .zip con las planillas PDF",
        type=["zip"],
        help="El ZIP debe contener únicamente archivos PDF de planillas de seguridad social.",
    )

    if not uploaded_zip:
        st.markdown("""
        <div style="text-align:center;padding:3rem 0;color:#1e3a6e;font-family:'IBM Plex Mono',monospace;font-size:0.8rem;">
            Esperando archivo ZIP…<br>
            <span style="font-size:0.65rem;letter-spacing:0.1em;color:#111e38;">
            [ Soporta planillas de MiPlanilla · Aportes en Línea · ARUS · SOI · Asopago · Simple ]
            </span>
        </div>
        """, unsafe_allow_html=True)
        return

    # ── Procesar ZIP ─────────────────────
    with tempfile.TemporaryDirectory() as tmpdir:
        # Extraer ZIPcon
        with zipfile.ZipFile(io.BytesIO(uploaded_zip.read()), "r") as zf:
            zf.extractall(tmpdir)

        # Listar PDFs (incluye subdirectorios)
        archivos_pdf = []
        for root, _, files in os.walk(tmpdir):
            for f in files:
                if f.lower().endswith(".pdf"):
                    archivos_pdf.append(os.path.join(root, f))

        total_en_zip = len(archivos_pdf)

        if total_en_zip == 0:
            st.warning("El ZIP no contiene archivos PDF.")
            return

        st.markdown(f"""
        <div style="font-family:'IBM Plex Mono',monospace;font-size:0.75rem;color:#3b82f6;
                    padding:0.6rem 1rem;background:#0a1022;border:1px solid #1a2744;
                    border-radius:6px;margin-bottom:1rem;">
            ✓ ZIP detectado &nbsp;·&nbsp; <strong>{total_en_zip}</strong> archivos PDF encontrados
        </div>
        """, unsafe_allow_html=True)

        iniciar = st.button("▶ Iniciar análisis", use_container_width=False)
        if not iniciar:
            return

        # ── Procesamiento ─────────────────
        st.markdown('<div class="section-title">② Procesamiento</div>', unsafe_allow_html=True)
        progress_slot = st.empty()
        lista_final   = []

        for idx, ruta_pdf in enumerate(archivos_pdf):
            nombre = os.path.basename(ruta_pdf)
            msg_idx = idx % len(MENSAJES_PROGRESO)

            render_progress(idx, total_en_zip, nombre, MENSAJES_PROGRESO[msg_idx], progress_slot)

            try:
                texto, paginas = extraer_texto_con_azure(ruta_pdf, cliente_di)

                # Actualizar mensaje al paso de OpenAI
                render_progress(idx, total_en_zip, nombre, "Analizando imagen con GPT-4o…", progress_slot)

                imagen_b64 = pdf_primera_pagina_a_base64(ruta_pdf)
                datos      = extraer_campos_con_openai(texto, imagen_b64, cliente_oai)

                datos["estado_validacion"] = validar_datos(datos)
                datos["archivo"]           = nombre
                datos["paginas_planilla"]     = paginas
                lista_final.append(datos)

            except Exception as e:
                lista_final.append({
                    "archivo"          : nombre,
                    "estado_validacion": f"ERROR: {str(e)[:100]}",
                })

            time.sleep(0.3)

        # Barra al 100%
        render_progress(total_en_zip, total_en_zip, "Proceso completado", "✓ Análisis finalizado", progress_slot)
        time.sleep(0.8)
        progress_slot.empty()

        # ── Construir DataFrame ───────────
        columnas = [
            "archivo","operador","tipo_identificacion","identificacion_aportante",
            "nombre_aportante","fecha_pago","periodo","total_pagado","moneda",
            "estado_validacion","paginas_planilla",
        ]
        df = pd.DataFrame(lista_final)
        for col in columnas:
            if col not in df.columns:
                df[col] = None
        df = df[columnas]

        # ── Banner de éxito ───────────────
        planillas_ok  = len(df)
        total_cop     = pd.to_numeric(df["total_pagado"], errors="coerce").sum()
        total_millones = round(total_cop / 1_000_000, 1)

        st.markdown(f"""
        <div class="success-banner">
            ✓ Análisis completado exitosamente &nbsp;·&nbsp;
            {planillas_ok} planillas procesadas &nbsp;·&nbsp;
            Total: ${total_millones:,.1f}M COP
        </div>
        """, unsafe_allow_html=True)

        # ── KPIs ─────────────────────────
        st.markdown('<div class="section-title">③ Indicadores clave</div>', unsafe_allow_html=True)
        col1, col2, col3 = st.columns(3)

        with col1:
            st.markdown(kpi_card(
                str(total_en_zip),
                "Archivos en el ZIP",
                sub="PDFs detectados"
            ), unsafe_allow_html=True)

        with col2:
            st.markdown(kpi_card(
                str(planillas_ok),
                "Planillas registradas",
                sub=f"en el Excel consolidado"
            ), unsafe_allow_html=True)

        with col3:
            st.markdown(kpi_card(
                f"${total_millones:,.1f}M",
                "Total pagos (aprox.)",
                sub="Suma de aportes · COP"
            ), unsafe_allow_html=True)

        # ── Gráficas ─────────────────────
        st.markdown('<div class="section-title">④ Análisis visual</div>', unsafe_allow_html=True)

        col_g1, col_g2 = st.columns([1.4, 1])
        with col_g1:
            st.plotly_chart(grafica_barras_proveedor(df), use_container_width=True)
        with col_g2:
            st.plotly_chart(grafica_dona_validacion(df), use_container_width=True)

        st.plotly_chart(grafica_pagos_proveedor(df), use_container_width=True)

        # ── Vista previa tabla ────────────
        st.markdown('<div class="section-title">⑤ Datos consolidados</div>', unsafe_allow_html=True)
        st.dataframe(
            df.head(50),
            use_container_width=True,
            hide_index=True,
        )
        if len(df) > 50:
            st.caption(f"Mostrando 50 de {len(df)} filas. Descarga el Excel para ver todos.")

        # ── Descarga Excel ────────────────
        st.markdown('<div class="section-title">⑥ Exportar</div>', unsafe_allow_html=True)
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Planillas SS")
            ws = writer.sheets["Planillas SS"]
            for col_cells in ws.columns:
                max_len = max(
                    (len(str(c.value)) for c in col_cells if c.value is not None),
                    default=10,
                )
                ws.column_dimensions[col_cells[0].column_letter].width = min(max_len + 4, 45)
        buffer.seek(0)

        ts = datetime.now().strftime("%Y%m%d_%H%M")
        st.download_button(
            label="⬇ Descargar Excel consolidado",
            data=buffer,
            file_name=f"Consolidado_Financiero_DNP_{ts}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        st.markdown("""
        <div style="font-family:'IBM Plex Mono',monospace;font-size:0.62rem;
                    color:#1a2744;text-align:center;margin-top:3rem;padding-top:1rem;
                    border-top:1px solid #0d1525;">
            Herramienta OCR-SF-UCD · DNP Colombia · Subdirección Financiera
        </div>
        """, unsafe_allow_html=True)


if __name__ == "__main__":
    main()
