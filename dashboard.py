# =============================================================================
# DASHBOARD.PY — INTERFAZ INTERACTIVA EN STREAMLIT
# Portfolio Analyzer · Versión 1.3
# =============================================================================

import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
from datetime import datetime, timedelta
from io import BytesIO

from config import (
    APP_TITULO, APP_SUBTITULO, PERIODOS_DISPONIBLES,
    DIAS_ANIO_DEFAULT, RF_DEFAULT, EXCEL_EXPORT_NOMBRE
)
from datos import combinar_activos, calcular_rendimientos, cargar_excel
from calculos import (
    precios_base100, retorno_nominal, retorno_ea,
    volatilidad_diaria, volatilidad_anual, sharpe_ratio,
    analisis_percentiles, rango_percentil_actual,
    matriz_correlacion, covarianza_diaria, covarianza_anual,
    montecarlo_iteraciones, tabla_activos, tabla_portafolio,
    information_ratio
)
from optimizacion import (
    optimizar_todos, frontera_eficiente,
    portafolio_combinado_sharpe, portafolio_combinado_ir,
    portafolio_consenso
)

st.set_page_config(
    page_title="Portfolio Analyzer",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ─────────────────────────────────────────────────────────────────────────────
# HELPERS INTERNOS
# ─────────────────────────────────────────────────────────────────────────────

def _ir_portafolio(pesos: dict, df_ir: pd.DataFrame) -> float:
    """IR ponderado: suma(peso_activo * IR_promedio_activo)."""
    ir_prom = df_ir.mean()
    total, w_sum = 0.0, 0.0
    for activo, peso in pesos.items():
        if activo in ir_prom.index:
            total  += peso * ir_prom[activo]
            w_sum  += peso
    return total / w_sum if w_sum > 0 else 0.0


def _safe_str(text) -> str:
    """Sanitiza texto a Latin-1 para fpdf2."""
    replacements = {'—': '-', '–': '-', '’': "'", '“': '"', '”': '"',
                    'é': 'e', 'á': 'a', 'í': 'i', 'ó': 'o', 'ú': 'u', 'ñ': 'n',
                    'É': 'E', 'Á': 'A', 'Ó': 'O', 'Ú': 'U', 'Ñ': 'N', '…': '...'}
    t = str(text) if text is not None else ""
    for ch, rep in replacements.items():
        t = t.replace(ch, rep)
    return t.encode('latin-1', errors='ignore').decode('latin-1')


def generar_pdf_portafolio(
    cols_activos, benchmark, rf,
    df_opt, port_combinado, criterio_combinado,
    port_cons_ret, port_cons_vol,
    df_ir, df_sharpe, df_ea, df_vol_a
) -> bytes:
    """Genera investing memo PDF del analisis de portafolio con fpdf2."""
    try:
        from fpdf import FPDF
    except ImportError:
        return b""

    C_DARK  = (30, 58, 95)
    C_ALT   = (240, 244, 250)

    class _PDF(FPDF):
        def header(self):
            self.set_fill_color(*C_DARK)
            self.rect(0, 0, 210, 20, 'F')
            self.set_font('Helvetica', 'B', 10)
            self.set_text_color(255, 255, 255)
            self.set_xy(10, 5)
            self.cell(0, 10, 'PORTFOLIO ANALYZER - INVESTING MEMO')
            self.set_text_color(0, 0, 0)

        def footer(self):
            self.set_y(-15)
            self.set_font('Helvetica', 'I', 7)
            self.set_text_color(120, 120, 120)
            self.cell(0, 8,
                f'Pagina {self.page_no()} | Generado: {datetime.now().strftime("%Y-%m-%d %H:%M")} | '
                'Solo fines educativos. No constituye asesoria de inversion.',
                align='C')
            self.set_text_color(0, 0, 0)

        def cell(self, w=0, h=0, txt='', border=0, ln=False, align='', fill=False, link=''):
            super().cell(w, h, _safe_str(txt), border=border, ln=ln, align=align, fill=fill, link=link)

    def section(pdf, title):
        pdf.ln(4)
        pdf.set_fill_color(*C_DARK)
        pdf.set_font('Helvetica', 'B', 10)
        pdf.set_text_color(255, 255, 255)
        pdf.cell(0, 8, f'  {title}', fill=True, ln=True)
        pdf.set_text_color(0, 0, 0)
        pdf.ln(2)

    def kv(pdf, label, value, alt=False):
        pdf.set_fill_color(*(C_ALT if alt else (255, 255, 255)))
        pdf.set_font('Helvetica', 'B', 8)
        pdf.cell(70, 6.5, f'  {label}', border='LTB', fill=True)
        pdf.set_font('Helvetica', '', 8)
        pdf.cell(110, 6.5, f'  {value}', border='RTB', fill=True, ln=True)

    pdf = _PDF()
    pdf.set_margins(15, 25, 15)
    pdf.set_auto_page_break(auto=True, margin=20)

    # ── Portada ────────────────────────────────────────────────────────────────
    pdf.add_page()
    pdf.set_fill_color(*C_DARK)
    pdf.rect(0, 50, 210, 80, 'F')
    pdf.set_font('Helvetica', 'B', 22)
    pdf.set_text_color(255, 255, 255)
    pdf.set_xy(15, 65)
    pdf.cell(0, 14, 'INVESTING MEMO', ln=True)
    pdf.set_font('Helvetica', 'B', 14)
    pdf.set_x(15)
    pdf.cell(0, 10, 'Analisis de Portafolio - Optimizacion Markowitz', ln=True)
    pdf.set_font('Helvetica', '', 11)
    pdf.set_x(15)
    pdf.cell(0, 8, f'Generado el {datetime.now().strftime("%d/%m/%Y %H:%M")}', ln=True)
    pdf.set_x(15)
    pdf.cell(0, 8, f'Activos: {", ".join(cols_activos)}  |  Benchmark: {benchmark}  |  RF: {rf:.2%}', ln=True)
    pdf.set_text_color(0, 0, 0)

    # ── Resumen ────────────────────────────────────────────────────────────────
    pdf.set_xy(15, 145)
    section(pdf, '1. RESUMEN EJECUTIVO')
    kv(pdf, 'Activos analizados', str(len(cols_activos)), alt=True)
    kv(pdf, 'Benchmark',          benchmark)
    kv(pdf, 'Tasa libre de riesgo', f'{rf:.2%}', alt=True)
    kv(pdf, 'Metodo combinacion',  criterio_combinado)
    kv(pdf, 'Portafolio combinado - Retorno',    f'{port_combinado["retorno"]:.2%}', alt=True)
    kv(pdf, 'Portafolio combinado - Volatilidad',f'{port_combinado["volatilidad"]:.2%}')
    kv(pdf, 'Portafolio combinado - Sharpe',     f'{port_combinado["sharpe"]:.3f}', alt=True)

    # ── 9 Portafolios ─────────────────────────────────────────────────────────
    pdf.add_page()
    section(pdf, '2. LOS 9 PORTAFOLIOS OPTIMOS (3 Metodos x 3 Objetivos)')

    if not df_opt.empty:
        # Cabecera
        pdf.set_fill_color(*C_DARK)
        pdf.set_font('Helvetica', 'B', 7)
        pdf.set_text_color(255, 255, 255)
        hdrs  = ['Metodo', 'Objetivo', 'Retorno', 'Volatilidad', 'Sharpe', 'IR Pond.']
        wdths = [28, 32, 22, 25, 20, 22]
        for h, w in zip(hdrs, wdths):
            pdf.cell(w, 7, h, border=1, fill=True, align='C')
        pdf.ln()
        pdf.set_text_color(0, 0, 0)

        for i, (idx, row) in enumerate(df_opt.iterrows()):
            pesos_row = {a: row[a] for a in cols_activos if a in row.index}
            ir_val = _ir_portafolio(pesos_row, df_ir)
            alt = (i % 2 == 0)
            pdf.set_fill_color(*(C_ALT if alt else (255, 255, 255)))
            pdf.set_font('Helvetica', '', 7)
            pdf.cell(28, 6.5, str(idx[0]), border=1, fill=True)
            pdf.cell(32, 6.5, str(idx[1]), border=1, fill=True)
            pdf.cell(22, 6.5, f'{row["Retorno"]:.2%}',     border=1, fill=True, align='C')
            pdf.cell(25, 6.5, f'{row["Volatilidad"]:.2%}', border=1, fill=True, align='C')
            pdf.cell(20, 6.5, f'{row["Sharpe"]:.3f}',      border=1, fill=True, align='C')
            pdf.cell(22, 6.5, f'{ir_val:.3f}',             border=1, fill=True, align='C')
            pdf.ln()

    # ── Portafolio Combinado ───────────────────────────────────────────────────
    section(pdf, '3. PORTAFOLIO COMBINADO')
    kv(pdf, 'Criterio ponderacion', criterio_combinado, alt=True)
    kv(pdf, 'Retorno esperado',    f'{port_combinado["retorno"]:.2%}')
    kv(pdf, 'Volatilidad',         f'{port_combinado["volatilidad"]:.2%}', alt=True)
    kv(pdf, 'Sharpe Ratio',        f'{port_combinado["sharpe"]:.3f}')
    kv(pdf, 'N portafolios usados',str(port_combinado['n_portafolios_usados']), alt=True)

    pdf.ln(3)
    pdf.set_font('Helvetica', 'B', 8)
    pdf.cell(0, 6, '  Asignacion de activos:', ln=True)
    for i, (activo, peso) in enumerate(sorted(port_combinado['pesos'].items(), key=lambda x: -x[1])):
        kv(pdf, activo, f'{peso:.2%}', alt=(i % 2 == 0))

    # ── Portafolios Consenso ───────────────────────────────────────────────────
    pdf.add_page()
    section(pdf, '4. PORTAFOLIOS CONSENSO PONDERADOS')

    for label, pc in [
        ('Consenso 60% Max Retorno',       port_cons_ret),
        ('Consenso 60% Min Volatilidad',   port_cons_vol),
    ]:
        pdf.ln(2)
        pdf.set_font('Helvetica', 'B', 9)
        pdf.cell(0, 7, f'  {label}', ln=True)
        kv(pdf, 'Objetivo principal',  f'{pc["objetivo_principal"]} ({pc["peso_principal"]:.0%})', alt=True)
        kv(pdf, 'Retorno esperado',    f'{pc["retorno"]:.2%}')
        kv(pdf, 'Volatilidad',         f'{pc["volatilidad"]:.2%}', alt=True)
        kv(pdf, 'Sharpe Ratio',        f'{pc["sharpe"]:.3f}')
        kv(pdf, 'IR Ponderado',        f'{_ir_portafolio(pc["pesos"], df_ir):.3f}', alt=True)
        pdf.ln(2)
        pdf.set_font('Helvetica', 'B', 7.5)
        pdf.set_fill_color(*C_DARK)
        pdf.set_text_color(255, 255, 255)
        for h, w in zip(['Activo', 'Peso'], [40, 30]):
            pdf.cell(w, 6.5, h, border=1, fill=True, align='C')
        pdf.ln()
        pdf.set_text_color(0, 0, 0)
        for j, (activo, peso) in enumerate(sorted(pc['pesos'].items(), key=lambda x: -x[1])):
            pdf.set_fill_color(*(C_ALT if j % 2 == 0 else (255, 255, 255)))
            pdf.set_font('Helvetica', '', 7.5)
            pdf.cell(40, 6, activo, border=1, fill=True)
            pdf.cell(30, 6, f'{peso:.2%}', border=1, fill=True, align='C')
            pdf.ln()
        pdf.ln(4)

    # ── Metricas por activo ───────────────────────────────────────────────────
    section(pdf, '5. METRICAS POR ACTIVO (ultimo periodo disponible)')
    pdf.set_fill_color(*C_DARK)
    pdf.set_font('Helvetica', 'B', 7)
    pdf.set_text_color(255, 255, 255)
    for h, w in zip(['Activo', 'Retorno EA', 'Volatilidad', 'Sharpe', 'IR'], [35, 28, 28, 22, 22]):
        pdf.cell(w, 7, h, border=1, fill=True, align='C')
    pdf.ln()
    pdf.set_text_color(0, 0, 0)

    ret_last = df_ea.iloc[-1] if not df_ea.empty else pd.Series()
    vol_last = df_vol_a.iloc[-1] if not df_vol_a.empty else pd.Series()
    shr_last = df_sharpe.iloc[-1] if not df_sharpe.empty else pd.Series()
    ir_last  = df_ir.iloc[-1] if not df_ir.empty else pd.Series()

    for i, activo in enumerate(cols_activos):
        alt = (i % 2 == 0)
        pdf.set_fill_color(*(C_ALT if alt else (255, 255, 255)))
        pdf.set_font('Helvetica', '', 7)
        def gv(series, key):
            return f'{series[key]:.2%}' if key in series.index and not pd.isna(series[key]) else 'N/A'
        def gf(series, key, fmt='.3f'):
            return f'{series[key]:{fmt}}' if key in series.index and not pd.isna(series[key]) else 'N/A'
        pdf.cell(35, 6.5, activo,               border=1, fill=True)
        pdf.cell(28, 6.5, gv(ret_last, activo), border=1, fill=True, align='C')
        pdf.cell(28, 6.5, gv(vol_last, activo), border=1, fill=True, align='C')
        pdf.cell(22, 6.5, gf(shr_last, activo), border=1, fill=True, align='C')
        pdf.cell(22, 6.5, gf(ir_last,  activo), border=1, fill=True, align='C')
        pdf.ln()

    # ── Disclaimer ────────────────────────────────────────────────────────────
    pdf.add_page()
    section(pdf, 'DISCLAIMER Y METODOLOGIA')
    pdf.set_font('Helvetica', '', 8)
    pdf.set_text_color(80, 80, 80)
    disclaimer = _safe_str(
        "Este documento es de caracter educativo e informativo unicamente. "
        "No constituye asesoria financiera ni recomendacion de inversion. "
        "Los modelos de optimizacion (Markowitz, CAPM, Montecarlo) implican supuestos "
        "sobre distribucion de retornos que pueden no verificarse en el futuro.\n\n"
        "Metodologia: WACC via CAPM, optimizacion via scipy SLSQP, "
        "Information Ratio = (Retorno portafolio - Retorno benchmark) / Tracking Error. "
        "Portafolios consenso: ponderacion por objetivo segun fraccion definida por el usuario."
    )
    pdf.multi_cell(0, 5, disclaimer)

    return bytes(pdf.output())

st.title(APP_TITULO)
st.caption(APP_SUBTITULO)

# -----------------------------------------------------------------------------
# SIDEBAR
# -----------------------------------------------------------------------------

st.sidebar.header("⚙️ Configuración")

st.sidebar.subheader("Fuente de datos")
fuente = st.sidebar.radio(
    "Selecciona la fuente",
    options=["Yahoo Finance", "Cargar Excel"],
    horizontal=True
)

st.sidebar.subheader("Activos")

if fuente == "Yahoo Finance":
    tickers_yahoo_input = st.sidebar.text_input(
        "Tickers Yahoo Finance (último = benchmark)",
        value="AAPL, NVDA, KO, ^GSPC",
        help="Separa con comas. El último ticker será el benchmark."
    )
    tickers_fics_input = st.sidebar.text_input(
        "FICs Colombia (palabras clave)",
        value="",
        help="Opcional. Separa con comas."
    )
    archivo_excel_input = None
else:
    archivo_excel_input = st.sidebar.file_uploader(
        "Sube tu archivo Excel",
        type=["xlsx"],
        help="Primera columna = fechas. Última columna = benchmark."
    )
    tickers_yahoo_input = ""
    tickers_fics_input = ""

st.sidebar.subheader("Período de análisis")

fecha_especifica = st.sidebar.date_input(
    "Fecha inicio específica (opcional)",
    value=None
)

periodos_seleccionados = st.sidebar.multiselect(
    "Períodos a analizar",
    options=list(PERIODOS_DISPONIBLES.keys()),
    default=["1 año", "2 años", "3 años"]
)

st.sidebar.subheader("Parámetros")

dias_anio = st.sidebar.number_input(
    "Días hábiles por año",
    min_value=252,
    max_value=365,
    value=DIAS_ANIO_DEFAULT,
    step=1
)

rf_pct = st.sidebar.number_input(
    "Tasa libre de riesgo RF (%)",
    min_value=0.0,
    max_value=20.0,
    value=RF_DEFAULT,
    step=0.1,
    format="%.2f"
)
rf = rf_pct / 100

st.sidebar.subheader("Criterio portafolio combinado")
criterio_combinado = st.sidebar.radio(
    "Ponderar por:",
    options=[
        "Sharpe Ratio",
        "Information Ratio",
        "Max Retorno (60% peso)",
        "Min Volatilidad (60% peso)",
    ],
    help=(
        "Sharpe / IR: pondera los 9 portafolios por la métrica elegida.\n"
        "Max Retorno: asigna 60% a los 3 portafolios de máximo retorno, 40% al resto.\n"
        "Min Volatilidad: asigna 60% a los 3 portafolios de mínima volatilidad, 40% al resto."
    )
)

ejecutar = st.sidebar.button("▶ Ejecutar análisis", type="primary", use_container_width=True)

# -----------------------------------------------------------------------------
# LÓGICA PRINCIPAL
# -----------------------------------------------------------------------------

if ejecutar:

    tickers_yahoo = [t.strip().upper() for t in tickers_yahoo_input.split(',') if t.strip()]
    tickers_fics = [t.strip() for t in tickers_fics_input.split(',') if t.strip()]

    if fuente == "Yahoo Finance":
        if len(tickers_yahoo) < 2 and not tickers_fics:
            st.error("Necesitas al menos 2 tickers en Yahoo Finance (activos + benchmark).")
            st.stop()
        if len(tickers_yahoo) >= 2:
            benchmark = tickers_yahoo[-1]
            activos_yahoo = tickers_yahoo[:-1]
        else:
            benchmark = ""
            activos_yahoo = []
    else:
        benchmark = ""
        activos_yahoo = []

    if fecha_especifica:
        fecha_inicio = fecha_especifica.strftime('%Y-%m-%d')
    elif "Máximo disponible" in periodos_seleccionados:
        fecha_inicio = "2000-01-01"
    else:
        años_max = max(
            [v for v in [PERIODOS_DISPONIBLES.get(p) for p in periodos_seleccionados] if v is not None],
            default=3
        )
        fecha_inicio = (datetime.today() - timedelta(days=int(años_max * 365) + 30)).strftime('%Y-%m-%d')

    periodos = {k: v for k, v in PERIODOS_DISPONIBLES.items() if k in periodos_seleccionados}
    if not periodos:
        st.error("Selecciona al menos un período de análisis.")
        st.stop()

    with st.spinner("Cargando datos..."):
        try:
            if fuente == "Yahoo Finance":
                df_precios, benchmark_nombre = combinar_activos(
                    tickers_yahoo, tickers_fics, benchmark, fecha_inicio
                )
                if not benchmark:
                    benchmark = df_precios.columns[-1]
                    benchmark_nombre = benchmark
            else:
                if archivo_excel_input is None:
                    st.error("Sube un archivo Excel para continuar.")
                    st.stop()
                df_precios = cargar_excel(archivo_excel_input)
                benchmark = df_precios.columns[-1]
                benchmark_nombre = benchmark
        except Exception as e:
            st.error(f"Error cargando datos: {e}")
            st.stop()

    cols_activos = [c for c in df_precios.columns if c != benchmark]
    n_activos = len(cols_activos)
    pesos = {c: 1/n_activos for c in cols_activos}

    with st.spinner("Calculando métricas..."):
        df_rend = calcular_rendimientos(df_precios, pesos)
        df_base100 = precios_base100(df_precios)
        df_nom = retorno_nominal(df_precios, periodos)
        df_ea = retorno_ea(df_nom, periodos, dias_anio, total_dias=len(df_precios))
        df_vol_d = volatilidad_diaria(df_rend, periodos)
        df_vol_a = volatilidad_anual(df_vol_d, dias_anio)
        df_sharpe = sharpe_ratio(df_ea, df_vol_a, rf)
        df_ir = information_ratio(df_rend, benchmark, periodos, dias_anio)
        df_percentiles = analisis_percentiles(df_rend)
        df_rango = rango_percentil_actual(df_rend)
        df_corr = matriz_correlacion(df_rend)
        df_cov_d = covarianza_diaria(df_rend)
        df_cov_a = covarianza_anual(df_cov_d, dias_anio)
        retornos_mc = montecarlo_iteraciones(df_rend, dias_anio=dias_anio)
        tablas_act = tabla_activos(df_rend, df_cov_a, pesos, rf, benchmark, retornos_mc, dias_anio)
        tablas_port = tabla_portafolio(tablas_act, pesos, cols_activos)

    with st.spinner("Optimizando portafolios..."):
        df_opt = optimizar_todos(tablas_act, df_cov_a, rf)
        port_cons_ret = portafolio_consenso(df_opt, cols_activos, 'Max Retorno', 0.60)
        port_cons_vol = portafolio_consenso(df_opt, cols_activos, 'Min Volatilidad', 0.60)
        if criterio_combinado == "Sharpe Ratio":
            port_combinado = portafolio_combinado_sharpe(df_opt, cols_activos)
        elif criterio_combinado == "Information Ratio":
            port_combinado = portafolio_combinado_ir(df_opt, df_ir, cols_activos)
        elif criterio_combinado == "Max Retorno (60% peso)":
            port_combinado = port_cons_ret
        else:  # Min Volatilidad (60% peso)
            port_combinado = port_cons_vol

    # -------------------------------------------------------------------------
    # SECCIÓN 1: RESUMEN
    # -------------------------------------------------------------------------
    st.header("📈 Resumen del Portafolio")
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Activos analizados", len(cols_activos))
    col2.metric("Benchmark", benchmark)
    col3.metric("Fechas disponibles", len(df_precios))
    col4.metric("RF utilizada", f"{rf_pct:.2f}%")

    # -------------------------------------------------------------------------
    # SECCIÓN 2: PORTAFOLIOS ÓPTIMOS
    # -------------------------------------------------------------------------
    st.header("🎯 Portafolios Óptimos")
    st.caption("9 portafolios: 3 métodos × 3 objetivos. Pesos en % por activo.")

    if not df_opt.empty:
        cols_pesos = [c for c in df_opt.columns if c in cols_activos]
        cols_metricas = ['Retorno', 'Volatilidad', 'Sharpe']
        df_display = df_opt[cols_pesos + cols_metricas].copy()
        for col in cols_pesos:
            df_display[col] = df_display[col].map(lambda x: f"{x:.1%}")
        for col in cols_metricas:
            df_display[col] = df_display[col].map(lambda x: f"{x:.2%}" if col != 'Sharpe' else f"{x:.3f}")
        st.dataframe(df_display, use_container_width=True)

    # -------------------------------------------------------------------------
    # SECCIÓN 3: PORTAFOLIO COMBINADO
    # -------------------------------------------------------------------------
    st.header(f"🔀 Portafolio Combinado (Ponderado por {criterio_combinado})")
    st.caption("Combinación de los portafolios óptimos con métrica positiva.")

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Retorno", f"{port_combinado['retorno']:.2%}")
    col2.metric("Volatilidad", f"{port_combinado['volatilidad']:.2%}")
    col3.metric("Sharpe", f"{port_combinado['sharpe']:.3f}")
    col4.metric("Portafolios usados", port_combinado['n_portafolios_usados'])

    df_pesos_comb = pd.DataFrame(
        port_combinado['pesos'].items(),
        columns=['Activo', 'Peso']
    )
    df_pesos_comb['Peso'] = df_pesos_comb['Peso'].map(lambda x: f"{x:.1%}")
    st.dataframe(df_pesos_comb, use_container_width=True)

    # -------------------------------------------------------------------------
    # SECCIÓN 4: ANÁLISIS UNIFICADO DE PORTAFOLIOS
    # -------------------------------------------------------------------------
    st.header("🏆 Análisis Unificado de Portafolios")
    st.caption("Comparación de los 9 portafolios optimizados + 2 portafolios consenso ponderados.")

    # Tabla comparativa: 9 + 2 portafolios
    if not df_opt.empty:
        filas_unif = []
        for idx, row in df_opt.iterrows():
            pesos_row = {a: row[a] for a in cols_activos if a in row.index}
            filas_unif.append({
                'Tipo':        'Optimizado',
                'Metodo':      idx[0],
                'Objetivo':    idx[1],
                'Retorno':     row['Retorno'],
                'Volatilidad': row['Volatilidad'],
                'Sharpe':      row['Sharpe'],
                'IR Pond.':    _ir_portafolio(pesos_row, df_ir),
            })

        for label, pc in [
            ('Consenso 60% Max Retorno',     port_cons_ret),
            ('Consenso 60% Min Volatilidad', port_cons_vol),
        ]:
            filas_unif.append({
                'Tipo':        'Consenso',
                'Metodo':      label,
                'Objetivo':    f'{pc["objetivo_principal"]} ({pc["peso_principal"]:.0%})',
                'Retorno':     pc['retorno'],
                'Volatilidad': pc['volatilidad'],
                'Sharpe':      pc['sharpe'],
                'IR Pond.':    _ir_portafolio(pc['pesos'], df_ir),
            })

        df_unif = pd.DataFrame(filas_unif)
        df_unif_display = df_unif.copy()
        for col in ['Retorno', 'Volatilidad']:
            df_unif_display[col] = df_unif_display[col].map(lambda x: f"{x:.2%}")
        df_unif_display['Sharpe']    = df_unif_display['Sharpe'].map(lambda x: f"{x:.3f}")
        df_unif_display['IR Pond.']  = df_unif_display['IR Pond.'].map(lambda x: f"{x:.3f}")
        st.dataframe(df_unif_display, use_container_width=True, hide_index=True)

        # Scatter retorno vs volatilidad — todos los portafolios
        fig_unif = go.Figure()
        colores_tipo = {'Optimizado': '#4C8BF5', 'Consenso': '#FFD600'}
        simbolos     = {'Optimizado': 'circle',  'Consenso': 'star'}
        for tipo in ['Optimizado', 'Consenso']:
            sub = df_unif[df_unif['Tipo'] == tipo]
            fig_unif.add_trace(go.Scatter(
                x=sub['Volatilidad'],
                y=sub['Retorno'],
                mode='markers+text',
                name=tipo,
                text=sub['Objetivo'],
                textposition='top center',
                textfont=dict(size=8),
                marker=dict(size=12 if tipo == 'Consenso' else 9,
                            color=colores_tipo[tipo],
                            symbol=simbolos[tipo],
                            line=dict(color='white', width=1))
            ))
        fig_unif.update_layout(
            title='Retorno vs Volatilidad — 9 Portafolios + 2 Consenso',
            xaxis_title='Volatilidad Anual',
            yaxis_title='Retorno Esperado',
            xaxis=dict(tickformat='.1%'),
            yaxis=dict(tickformat='.1%'),
            height=420,
            template='plotly_dark',
        )
        st.plotly_chart(fig_unif, use_container_width=True)

    # Portafolios consenso — detalle
    st.subheader("Portafolios Consenso Ponderados")
    col_c1, col_c2 = st.columns(2)
    for col_c, pc, label in [
        (col_c1, port_cons_ret, "60% Max Retorno"),
        (col_c2, port_cons_vol, "60% Min Volatilidad"),
    ]:
        with col_c:
            st.markdown(f"**{label}**")
            st.metric("Retorno",     f'{pc["retorno"]:.2%}')
            st.metric("Volatilidad", f'{pc["volatilidad"]:.2%}')
            st.metric("Sharpe",      f'{pc["sharpe"]:.3f}')
            st.metric("IR Pond.",    f'{_ir_portafolio(pc["pesos"], df_ir):.3f}')
            df_cons = pd.DataFrame(
                sorted(pc['pesos'].items(), key=lambda x: -x[1]),
                columns=['Activo', 'Peso']
            )
            df_cons['Peso'] = df_cons['Peso'].map(lambda x: f"{x:.1%}")
            st.dataframe(df_cons, use_container_width=True, hide_index=True)

    # -------------------------------------------------------------------------
    # SECCIÓN 5: ANÁLISIS DETALLADO
    # -------------------------------------------------------------------------
    st.header("📊 Análisis Detallado")

    with st.expander("Retorno Nominal por Período"):
        st.dataframe(df_nom.style.format("{:.2%}"), use_container_width=True)

    with st.expander("Retorno Efectivo Anual (EA)"):
        st.dataframe(df_ea.style.format("{:.2%}"), use_container_width=True)

    with st.expander("Volatilidad Diaria"):
        st.dataframe(df_vol_d.style.format("{:.4%}"), use_container_width=True)

    with st.expander("Volatilidad Anual"):
        st.dataframe(df_vol_a.style.format("{:.2%}"), use_container_width=True)

    with st.expander("Sharpe Ratio"):
        st.dataframe(df_sharpe.style.format("{:.3f}"), use_container_width=True)

    with st.expander("Information Ratio por Período"):
        st.caption("IR > 1.0: excelente | 0.5-1.0: buena | 0-0.5: mediocre | < 0: destruye valor vs benchmark")
        st.dataframe(df_ir.style.format("{:.3f}"), use_container_width=True)

    with st.expander("Análisis de Percentiles (rendimientos diarios)"):
        st.dataframe(df_percentiles.style.format("{:.4%}"), use_container_width=True)

    with st.expander("Rango Percentil Actual"):
        st.dataframe(df_rango.to_frame().T.style.format("{:.2%}"), use_container_width=True)

    with st.expander("Matriz de Correlación"):
        fig_corr = px.imshow(
            df_corr, text_auto=".2f",
            color_continuous_scale='RdBu_r',
            zmin=-1, zmax=1,
            title="Correlación de Rendimientos"
        )
        st.plotly_chart(fig_corr, use_container_width=True)

    with st.expander("Covarianza Diaria"):
        st.dataframe(df_cov_d.style.format("{:.8f}"), use_container_width=True)

    with st.expander("Covarianza Anual"):
        st.dataframe(df_cov_a.style.format("{:.6f}"), use_container_width=True)

    with st.expander("Tabla por Activo — Markowitz"):
        st.dataframe(tablas_act['markowitz'].style.format("{:.4f}"), use_container_width=True)

    with st.expander("Tabla por Activo — CAPM"):
        st.dataframe(tablas_act['capm'].style.format("{:.4f}"), use_container_width=True)

    with st.expander("Tabla por Activo — Montecarlo"):
        st.dataframe(tablas_act['montecarlo'].style.format("{:.4f}"), use_container_width=True)

    with st.expander("Verificación del Portafolio"):
        col1, col2, col3 = st.columns(3)
        for i, (metodo, df_p) in enumerate(tablas_port.items()):
            [col1, col2, col3][i].subheader(metodo.capitalize())
            [col1, col2, col3][i].dataframe(df_p.style.format("{:.4f}"))

    # -------------------------------------------------------------------------
    # SECCIÓN 5: FRONTERA EFICIENTE
    # -------------------------------------------------------------------------
    st.header("📉 Frontera Eficiente")

    fig = go.Figure()
    colores = {'markowitz': 'blue', 'capm': 'green', 'montecarlo': 'red'}

    for metodo in ['markowitz', 'capm', 'montecarlo']:
        if metodo not in tablas_act:
            continue
        retornos_metodo = tablas_act[metodo].loc['Retorno']
        cov_activos = df_cov_a.loc[cols_activos, cols_activos]
        df_frontera = frontera_eficiente(retornos_metodo, cov_activos, rf)
        if not df_frontera.empty:
            fig.add_trace(go.Scatter(
                x=df_frontera['Volatilidad'],
                y=df_frontera['Retorno'],
                mode='lines',
                name=f'Frontera {metodo.capitalize()}',
                line=dict(color=colores[metodo], width=2)
            ))

    if not df_opt.empty:
        fig.add_trace(go.Scatter(
            x=df_opt['Volatilidad'],
            y=df_opt['Retorno'],
            mode='markers+text',
            name='Portafolios Óptimos',
            text=df_opt.index.get_level_values('Objetivo'),
            textposition='top center',
            marker=dict(size=10, color='black', symbol='star')
        ))

    fig.add_trace(go.Scatter(
        x=[port_combinado['volatilidad']],
        y=[port_combinado['retorno']],
        mode='markers+text',
        name=f'Portafolio Combinado ({criterio_combinado})',
        text=['Combinado'],
        textposition='top center',
        marker=dict(size=14, color='gold', symbol='diamond', line=dict(color='black', width=2))
    ))

    fig.update_layout(
        title=f"Frontera Eficiente — 3 Métodos + Portafolio Combinado ({criterio_combinado})",
        xaxis_title="Volatilidad Anual",
        yaxis_title="Retorno Esperado",
        xaxis=dict(tickformat='.1%'),
        yaxis=dict(tickformat='.1%'),
        height=500,
        legend=dict(orientation='h', yanchor='bottom', y=1.02)
    )
    st.plotly_chart(fig, use_container_width=True)

    with st.expander("Ver evolución de precios (Base 100)"):
        fig2 = px.line(
            df_base100,
            title="Evolución de Precios — Base 100",
            labels={'value': 'Precio Normalizado', 'fecha': 'Fecha'}
        )
        st.plotly_chart(fig2, use_container_width=True)

    # -------------------------------------------------------------------------
    # SECCIÓN 6: EXPORTAR
    # -------------------------------------------------------------------------
    st.header("💾 Exportar")
    col_exp1, col_exp2 = st.columns(2)

    # ── Excel ─────────────────────────────────────────────────────────────────
    def generar_excel():
        buffer = BytesIO()
        try:
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                df_precios.to_excel(writer,    sheet_name='Precios')
                df_rend.to_excel(writer,       sheet_name='Rendimientos')
                df_base100.to_excel(writer,    sheet_name='Precios Base 100')
                df_nom.to_excel(writer,        sheet_name='Retorno Nominal')
                df_ea.to_excel(writer,         sheet_name='Retorno EA')
                df_vol_d.to_excel(writer,      sheet_name='Volatilidad Diaria')
                df_vol_a.to_excel(writer,      sheet_name='Volatilidad Anual')
                df_sharpe.to_excel(writer,     sheet_name='Sharpe Ratio')
                df_ir.to_excel(writer,         sheet_name='Information Ratio')
                df_percentiles.to_excel(writer,sheet_name='Percentiles')
                df_rango.to_frame().to_excel(writer, sheet_name='Rango Percentil')
                df_corr.to_excel(writer,       sheet_name='Correlacion')
                df_cov_d.to_excel(writer,      sheet_name='Covarianza Diaria')
                df_cov_a.to_excel(writer,      sheet_name='Covarianza Anual')
                for metodo, df_t in tablas_act.items():
                    sname = f'Tabla {metodo[:5].title()}'[:31]
                    df_t.to_excel(writer, sheet_name=sname)
                if not df_opt.empty:
                    df_opt.to_excel(writer, sheet_name='Portafolios Optimos')
                # Portafolio combinado
                df_pc = pd.DataFrame([port_combinado['pesos']]).T
                df_pc.columns = ['Peso']
                for col in ['retorno', 'volatilidad', 'sharpe']:
                    df_pc.loc[col.capitalize(), 'Peso'] = port_combinado[col]
                df_pc.to_excel(writer, sheet_name='Port Combinado')
                # Portafolios consenso
                for label_sheet, pc in [
                    ('Consenso MaxRetorno', port_cons_ret),
                    ('Consenso MinVol',     port_cons_vol),
                ]:
                    df_cs = pd.DataFrame(list(pc['pesos'].items()), columns=['Activo', 'Peso'])
                    metr = pd.DataFrame([{'Activo': 'Retorno',    'Peso': pc['retorno']},
                                         {'Activo': 'Volatilidad','Peso': pc['volatilidad']},
                                         {'Activo': 'Sharpe',     'Peso': pc['sharpe']},
                                         {'Activo': 'IR Pond.',   'Peso': _ir_portafolio(pc['pesos'], df_ir)}])
                    pd.concat([df_cs, metr], ignore_index=True).to_excel(
                        writer, sheet_name=label_sheet[:31], index=False)
                # Analisis unificado
                if not df_opt.empty:
                    filas_xl = []
                    for idx, row in df_opt.iterrows():
                        pw = {a: row[a] for a in cols_activos if a in row.index}
                        filas_xl.append({'Tipo':'Optimizado','Metodo':idx[0],'Objetivo':idx[1],
                                         'Retorno':row['Retorno'],'Volatilidad':row['Volatilidad'],
                                         'Sharpe':row['Sharpe'],'IR Pond.':_ir_portafolio(pw, df_ir)})
                    for lbl, pc in [('Consenso 60% Max Retorno', port_cons_ret),
                                    ('Consenso 60% Min Vol',     port_cons_vol)]:
                        filas_xl.append({'Tipo':'Consenso','Metodo':lbl,'Objetivo':pc['objetivo_principal'],
                                         'Retorno':pc['retorno'],'Volatilidad':pc['volatilidad'],
                                         'Sharpe':pc['sharpe'],'IR Pond.':_ir_portafolio(pc['pesos'], df_ir)})
                    pd.DataFrame(filas_xl).to_excel(writer, sheet_name='Analisis Unificado', index=False)
        except Exception as e:
            st.error(f"Error generando Excel: {e}")
        buffer.seek(0)
        return buffer

    with col_exp1:
        st.markdown("**📊 Excel — Modelo Completo**")
        st.caption("19 hojas: precios, retornos, riesgo, correlacion, portafolios, consenso.")
        st.download_button(
            label="📥 Descargar Excel",
            data=generar_excel(),
            file_name=EXCEL_EXPORT_NOMBRE,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

    # ── PDF ───────────────────────────────────────────────────────────────────
    with col_exp2:
        st.markdown("**📄 PDF — Investing Memo**")
        st.caption("Portafolios, consenso, metricas Sharpe e IR, asignacion de activos.")
        try:
            _pdf_bytes = generar_pdf_portafolio(
                cols_activos, benchmark, rf,
                df_opt, port_combinado, criterio_combinado,
                port_cons_ret, port_cons_vol,
                df_ir, df_sharpe, df_ea, df_vol_a
            )
        except Exception as _e:
            _pdf_bytes = b""
            st.error(f"Error generando PDF: {_e}")
        if _pdf_bytes:
            st.download_button(
                label="📥 Descargar Investing Memo PDF",
                data=_pdf_bytes,
                file_name=f"portfolio_memo_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf",
                mime="application/pdf",
                use_container_width=True,
                key="dl_pdf_pa"
            )
        else:
            st.warning("fpdf2 no instalado. Ejecuta: pip install fpdf2")

else:
    st.info("👈 Configura los parámetros en el panel izquierdo y presiona **Ejecutar análisis**.")
    st.markdown("""
    ### ¿Qué hace este sistema?
    1. **Descarga** precios de Yahoo Finance o carga un Excel propio
    2. **Calcula** 14 tablas de análisis financiero
    3. **Optimiza** 9 portafolios: 3 métodos × 3 objetivos
    4. **Combina** los portafolios ponderado por Sharpe o IR
    5. **Mide** Information Ratio vs benchmark
    6. **Visualiza** la frontera eficiente interactiva
    7. **Exporta** todos los resultados a Excel
    """)