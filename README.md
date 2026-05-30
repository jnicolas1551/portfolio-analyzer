# 📊 Portfolio Analyzer — Optimizacion Markowitz

Dashboard interactivo de analisis estadistico y optimizacion de portafolios de inversion.
Soporta activos del S&P 500, BVC Colombia y FICs colombianos.

[![Streamlit App](https://static.streamlit.io/badges/streamlit_badge_black_white.svg)](https://portfolio-analyzer-jnrc.streamlit.app)

---

## Funcionalidades

| Modulo | Descripcion |
|---|---|
| **Retornos** | Retorno nominal y efectivo anual (EA), base 100, periodos ajustables |
| **Riesgo** | Volatilidad diaria y anual, Value at Risk, analisis de percentiles |
| **Metricas** | Sharpe Ratio, Information Ratio vs benchmark |
| **Correlacion** | Matriz de correlacion y covarianza entre activos |
| **Monte Carlo** | Simulacion de escenarios de portafolio |
| **Markowitz** | 9 portafolios optimos: 3 metodos x 3 objetivos |
| **Frontera eficiente** | Visualizacion interactiva de la frontera de Markowitz |
| **Exportacion** | Excel con todas las tablas y graficas |

### Metodos de optimizacion (3 x 3)

| Metodo \ Objetivo | Max Retorno | Max Sharpe | Min Volatilidad |
|---|---|---|---|
| **Historico** | ✓ | ✓ | ✓ |
| **CAPM** | ✓ | ✓ | ✓ |
| **Monte Carlo** | ✓ | ✓ | ✓ |

---

## Instalacion local

```bash
git clone https://github.com/jnicolas1551/portfolio-analyzer.git
cd portfolio-analyzer
pip install -r requirements.txt
streamlit run dashboard.py
```

---

## Estructura del proyecto

```
portfolio-analyzer/
├── dashboard.py          # Entrada principal Streamlit
├── config.py             # Constantes: periodos, URLs, limites API
├── datos.py              # Descarga datos: Yahoo Finance + FICs Colombia
├── calculos.py           # Motor de calculo: 14 funciones financieras
├── optimizacion.py       # Optimizacion Markowitz: 9 portafolios optimos
├── descarga_acciones.py  # Script auxiliar descarga precios historicos
├── requirements.txt
└── .streamlit/
    └── config.toml       # Tema oscuro
```

---

## Fuentes de datos

| Fuente | Tipo de activo |
|---|---|
| Yahoo Finance (yfinance) | Acciones S&P 500, indices, ETFs |
| datos.gov.co API | FICs (Fondos de Inversion Colectiva) Colombia |
| Carga manual Excel | Cualquier serie de precios historicos |

---

## Metodologia

### Retornos
- **Nominal:** `(P_t / P_{t-1}) - 1`
- **Efectivo anual:** `(1 + r_diario)^252 - 1`
- **Base 100:** precio normalizado desde fecha de inicio

### Optimizacion Markowitz
- **Restricciones:** pesos suman 100%, pesos minimo/maximo configurables
- **Funcion objetivo:** maximizar Sharpe = `(E[r] - Rf) / sigma`
- **Frontera eficiente:** 500 portafolios simulados + 9 puntos optimos

### Sharpe e Information Ratio
```
Sharpe = (Retorno_portafolio - Rf) / Volatilidad_portafolio
IR     = (Retorno_portafolio - Retorno_benchmark) / Tracking_Error
```

---

## Mercados soportados

| Region | Ejemplos |
|---|---|
| USA (S&P 500) | AAPL, MSFT, ^GSPC, SPY, QQQ |
| Colombia (BVC) | ECOPETROL.CL, PFBCOLOM.CL, ISA.CL |
| FICs Colombia | Fondo Fiducolombia, Fondo Skandia, etc. |

---

## Disclaimer

Herramienta educativa. No constituye asesoria financiera ni recomendacion de inversion.

---

*Desarrollado con Python + Streamlit + yfinance + scipy + plotly*
