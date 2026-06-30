"""
==============================================================================
  DASHBOARD DE GESTIÓN DE PORTAFOLIO - VALORACIÓN
  Versión: 1.1  |  Entorno: Corporativo Windows
  Fuente de datos: J:\\VALORACION\\ALEJANDRO APONTE\\POSICION PROPIA\\Data Dashboard.xlsx
==============================================================================
  Librerías requeridas (instalar con pip si no están disponibles):
      pip install dash dash-bootstrap-components plotly pandas openpyxl
==============================================================================
"""

import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots
import dash
from dash import dcc, html, dash_table, Input, Output, callback
import dash_bootstrap_components as dbc
from datetime import date, datetime
import warnings
warnings.filterwarnings("ignore")

# ─────────────────────────────────────────────────────────────────────────────
#  CONFIGURACIÓN  –  Ajusta la ruta según tu entorno
# ─────────────────────────────────────────────────────────────────────────────
RUTA_ARCHIVO = r"J:\VALORACION\ALEJANDRO APONTE\POSICION PROPIA\Data Dashboard.xlsx"

# Colores corporativos
COLORS = {
    "bg_dark":    "#0F1923",
    "bg_card":    "#1A2535",
    "bg_card2":   "#1E2D42",
    "accent":     "#00C4FF",
    "accent2":    "#00E5A0",
    "accent3":    "#FF6B6B",
    "accent4":    "#FFB347",
    "text":       "#E8EDF2",
    "text_muted": "#7A8FA6F2",
    "border":     "#263548",
    "positive":   "#00E5A0",
    "negative":   "#FF6B6B",
    "neutral":    "#00C4FF",
}

# ─────────────────────────────────────────────────────────────────────────────
#  CARGA Y TRANSFORMACIÓN DE DATOS
# ─────────────────────────────────────────────────────────────────────────────
def cargar_datos(ruta: str) -> pd.DataFrame:
    df = pd.read_excel(ruta)
    df.columns = [c.strip() for c in df.columns]

    # Extraer fecha desde Source.Name  (formato YYYYMMDD)
    df["Fecha_Val"] = pd.to_datetime(
        df["Source.Name"].str.extract(r"(\d{8})")[0], format="%Y%m%d"
    )

    # Limpiar strings
    for col in ["Especie", "ISIN/Nemotécnic", "Moned", "Est"]:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()

    # Limpiar columna PORTAFOLIO (columna AF = índice 31)
    # Detectar automáticamente: busca por nombre o usa posición
    col_port = None
    for c in df.columns:
        if c.strip().upper() in ["PORTAFOLIO", "PORTFOLIO", "PORT"]:
            col_port = c
            break
    if col_port is None and len(df.columns) > 31:
        col_port = df.columns[31]
    if col_port is not None:
        df["PORTAFOLIO"] = df[col_port].astype(str).str.strip()
    else:
        df["PORTAFOLIO"] = "ÚNICO"

    # Columnas numéricas clave
    numericas = [
        "Vlr Nominal", "Vlr Mer. Ant", "Vlr Mer. Hoy",
        "Causación Mer", "Causación TIR", "TIR.Mercado",
        "Precio", "Dias", "Adeudados", "Mnd Val",
        "Causación Moneda", "Causación Tasa",
    ]
    for col in numericas:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    # Variación porcentual de precio respecto al día anterior
    df["Variacion_Abs"] = df["Vlr Mer. Hoy"] - df["Vlr Mer. Ant"]
    df["Variacion_Pct"] = (df["Variacion_Abs"] / df["Vlr Mer. Ant"].replace(0, pd.NA)) * 100

    # Duración en años
    df["Duracion_Anos"] = df["Dias"] / 365

    # Año de vencimiento — F.Vcto ya viene como datetime en el archivo
    df["Año_Vcto"] = pd.to_datetime(df["F.Vcto"], errors="coerce").dt.year

    # Causación Total — columna nativa del archivo (ya no se calcula)
    if "Causación Total" in df.columns:
        df["Causacion_Total"] = pd.to_numeric(df["Causación Total"], errors="coerce").fillna(0)
    else:
        df["Causacion_Total"] = df["Causación Mer"].fillna(0) + df["Causación TIR"].fillna(0)

    return df

df_global = cargar_datos(RUTA_ARCHIVO)
fechas_disponibles = sorted(df_global["Fecha_Val"].dt.date.unique())
portafolios_disponibles = sorted(df_global["PORTAFOLIO"].unique())

# ─────────────────────────────────────────────────────────────────────────────
#  FUNCIONES AUXILIARES
# ─────────────────────────────────────────────────────────────────────────────
def fmt_cop(valor: float) -> str:
    """Muestra el valor completo sin abreviaciones."""
    if pd.isna(valor):
        return "—"
    return f"${valor:,.0f}"

def color_var(valor: float) -> str:
    if pd.isna(valor) or valor == 0:
        return COLORS["neutral"]
    return COLORS["positive"] if valor > 0 else COLORS["negative"]

def filtrar_df(fecha_ini: date, fecha_fin: date) -> pd.DataFrame:
    mask = (df_global["Fecha_Val"].dt.date >= fecha_ini) & \
           (df_global["Fecha_Val"].dt.date <= fecha_fin)
    return df_global[mask].copy()

def calcular_rentabilidad_portafolio(df: pd.DataFrame, portafolio: str):
    """
    Lógica Excel:
      VPN_Inicial  = Vlr Mer. Hoy del último día del mes anterior
      VPN_Final    = Vlr Mer. Hoy del último día del rango filtrado
      TIR Acumulada= (VPN_Final / VPN_Inicial) - 1
      TIR Diaria   = Causación último día / VPN día anterior dentro del rango
    """
    df_p = df[df["PORTAFOLIO"] == portafolio].copy()
    if df_p.empty:
        return 0.0, 0.0

    causacion_diaria = df_p.groupby("Fecha_Val")["Causación Mer"].sum()
    vlr_diario       = df_p.groupby("Fecha_Val")["Vlr Mer. Hoy"].sum()

    # VPN Final: último día del rango filtrado
    vpn_final = vlr_diario.iloc[-1]

    # VPN Inicial: último día del mes anterior buscado en df_global
    fecha_fin_periodo  = vlr_diario.index.max()
    ultimo_dia_mes_ant = (fecha_fin_periodo.to_period("M") - 1).to_timestamp("M")

    df_mes_ant = df_global[
        df_global["Fecha_Val"].dt.date == ultimo_dia_mes_ant.date()
    ]
    if not df_mes_ant.empty:
        vpn_inicial = df_mes_ant[
            df_mes_ant["PORTAFOLIO"] == portafolio
        ]["Vlr Mer. Hoy"].sum()
    else:
        vpn_inicial = vlr_diario.iloc[0]  # fallback: primer día del rango

    if vpn_inicial == 0:
        return 0.0, 0.0

    # TIR Acumulada MTD = (VPN_Final - VPN_Inicial) / VPN_Inicial
    ret_acumulada = (vpn_final - vpn_inicial) / vpn_inicial

    # TIR Diaria = causación último día / valor mercado penúltimo día del rango
    if len(vlr_diario) > 1:
        base_diaria = vlr_diario.iloc[-2]
    else:
        base_diaria = vpn_inicial  # solo hay un día: usar VPN mes anterior

    ret_diaria = (
        causacion_diaria.iloc[-1] / base_diaria
        if base_diaria != 0 else 0.0
    )

    return float(ret_diaria), float(ret_acumulada)

# ─────────────────────────────────────────────────────────────────────────────
#  COMPONENTES DE LAYOUT
# ─────────────────────────────────────────────────────────────────────────────
def kpi_card(titulo, valor, subtitulo="", color_val=None, icono=""):
    col = color_val or COLORS["text"]
    return html.Div([
        html.Div([
            html.Span(icono, style={"fontSize": "22px", "marginRight": "8px"}),
            html.Span(titulo, style={"fontSize": "11px", "color": COLORS["text_muted"],
                                      "textTransform": "uppercase", "letterSpacing": "1px"}),
        ], style={"display": "flex", "alignItems": "center", "marginBottom": "8px"}),
        html.Div(valor, style={"fontSize": "24px", "fontWeight": "700",
                                "color": col, "lineHeight": "1"}),
        html.Div(subtitulo, style={"fontSize": "11px", "color": COLORS["text_muted"],
                                    "marginTop": "4px"}),
    ], style={
        "background": COLORS["bg_card"],
        "border": f"1px solid {COLORS['border']}",
        "borderTop": f"3px solid {col}",
        "borderRadius": "8px",
        "padding": "16px 18px",
        "height": "100%",
    })


def kpi_card_portafolio(titulo, valor, subtitulo="", color_val=None):
    """Tarjeta KPI compacta para la fila de portafolios — una sola fila."""
    col = color_val or COLORS["text"]
    return html.Div([
        html.Div(titulo, style={
            "fontSize": "9px", "color": COLORS["text_muted"],
            "textTransform": "uppercase", "letterSpacing": "0.6px",
            "marginBottom": "5px", "fontWeight": "600",
            "whiteSpace": "normal",
            "lineHeight": "1.3",
            "wordBreak": "break-word",
        }),
        html.Div(valor, style={
            "fontSize": "16px", "fontWeight": "700",
            "color": col, "lineHeight": "1",
        }),
        html.Div(subtitulo, style={
            "fontSize": "9px", "color": COLORS["text_muted"],
            "marginTop": "3px",
        }),
    ], style={
        "background": COLORS["bg_card2"],
        "border": f"1px solid {COLORS['border']}",
        "borderLeft": f"3px solid {col}",
        "borderRadius": "8px",
        "padding": "8px",
        "height": "100%",
        "minWidth": "0",
    })

# ─────────────────────────────────────────────────────────────────────────────
#  APP DASH
# ─────────────────────────────────────────────────────────────────────────────
app = dash.Dash(
    __name__,
    external_stylesheets=[dbc.themes.CYBORG],
    title="Dashboard Portafolio | Valoración",
    meta_tags=[{"name": "viewport", "content": "width=device-width, initial-scale=1"}],
    suppress_callback_exceptions=True,
)

# ── ESTILOS SIDEBAR ───────────────────────────────────────────────────────────
SIDEBAR_W = "220px"

def nav_item(label, icon, page_id, active=False):
    return html.Div(
        id=f"nav-{page_id}",
        n_clicks=0,
        children=[
            html.Span(icon, style={"fontSize": "18px", "marginRight": "10px"}),
            html.Span(label, style={"fontSize": "13px", "fontWeight": "600",
                                    "letterSpacing": "0.5px"}),
        ],
        style={
            "display": "flex", "alignItems": "center",
            "padding": "11px 16px", "borderRadius": "8px",
            "cursor": "pointer", "marginBottom": "4px",
            "background": COLORS["accent"] + "22" if active else "transparent",
            "borderLeft": f"3px solid {COLORS['accent']}" if active
                          else f"3px solid transparent",
            "color": COLORS["accent"] if active else COLORS["text_muted"],
            "transition": "all 0.2s",
        },
    )

# ── LAYOUT PRINCIPAL ──────────────────────────────────────────────────────────
app.layout = html.Div(
    style={"background": COLORS["bg_dark"], "minHeight": "100vh",
           "fontFamily": "'Segoe UI', Arial, sans-serif", "color": COLORS["text"],
           "display": "flex", "flexDirection": "column"},
    children=[

        # ── HEADER ──────────────────────────────────────────────────────────
        html.Div([
            html.Div([
                html.Span("📶", style={"fontSize": "26px", "color": COLORS["accent"],
                                      "marginRight": "12px"}),
                html.Div([
                    html.H1("OWN PORTFOLIO RETURNS",
                            style={"margin": 0, "fontSize": "19px", "fontWeight": "700",
                                   "color": COLORS["text"], "letterSpacing": "2px"}),
                    html.Div("Desempeño y valorización del portafolio",
                             style={"fontSize": "9px", "color": COLORS["text_muted"],
                                    "letterSpacing": "1px"}),
                ]),
            ], style={"display": "flex", "alignItems": "center"}),

            html.Div([
                html.Div([
                    html.Label("RANGO DE FECHAS", style={
                        "fontSize": "9px", "color": COLORS["text_muted"],
                        "letterSpacing": "1px", "marginBottom": "4px", "display": "block",
                    }),
                    dcc.DatePickerRange(
                        id="date-range",
                        min_date_allowed=min(fechas_disponibles),
                        max_date_allowed=max(fechas_disponibles),
                        start_date=min(fechas_disponibles),
                        end_date=max(fechas_disponibles),
                        display_format="DD/MM/YYYY",
                        style={"fontSize": "12px", "color": "#1f1f1f"},
                    ),
                ]),
                html.Div([
                    html.Label("PORTAFOLIO", style={
                        "fontSize": "9px", "color": COLORS["text_muted"],
                        "letterSpacing": "1px", "marginBottom": "4px", "display": "block",
                    }),
                    dcc.Dropdown(
                        id="especie-filter",
                        options=[{"label": "Todos", "value": "Todas"}] +
                                [{"label": p, "value": p}
                                 for p in sorted(df_global["PORTAFOLIO"].unique())],
                        value="Todas", clearable=False,
                        style={"width": "210px", "fontSize": "12px",
                               "background": COLORS["bg_card2"], "color": "#1f1f1f"},
                    ),
                ], style={"marginLeft": "18px"}),
            ], style={"display": "flex", "alignItems": "flex-end"}),
        ], style={
            "background": COLORS["bg_card"],
            "borderBottom": f"1px solid {COLORS['border']}",
            "padding": "13px 24px",
            "display": "flex", "justifyContent": "space-between",
            "alignItems": "center", "flexWrap": "wrap", "gap": "10px",
            "position": "sticky", "top": 0, "zIndex": 999,
        }),

        # ── CUERPO: SIDEBAR + CONTENIDO ──────────────────────────────────────
        html.Div(style={"display": "flex", "flex": "1"}, children=[

            # ── SIDEBAR ───────────────────────────────────────────────────────
            html.Div([
                html.Div("NAVEGACIÓN", style={
                    "fontSize": "9px", "color": COLORS["text_muted"],
                    "letterSpacing": "2px", "padding": "18px 16px 10px",
                }),
                nav_item("Inicio",                "🏠︎", "inicio",   active=True),
                nav_item("Análisis de Portafolio","🔍︎", "analisis", active=False),

                # Separador
                html.Div(style={
                    "height": "1px", "background": COLORS["border"],
                    "margin": "16px 12px",
                }),
                html.Div(
                    f"Datos al {max(fechas_disponibles).strftime('%d/%m/%Y')}",
                    style={"fontSize": "9px", "color": COLORS["text_muted"],
                           "padding": "0 16px", "lineHeight": "1.6"},
                ),
            ], style={
                "width": SIDEBAR_W, "minWidth": SIDEBAR_W,
                "background": COLORS["bg_card"],
                "borderRight": f"1px solid {COLORS['border']}",
                "minHeight": "calc(100vh - 60px)",
                "padding": "4px 8px",
                "position": "sticky", "top": "60px",
                "alignSelf": "flex-start",
            }),

            # ── ÁREA DE CONTENIDO ─────────────────────────────────────────────
            html.Div(style={"flex": "1", "padding": "20px 24px",
                            "overflowX": "hidden"}, children=[

                # Store para página activa
                dcc.Store(id="pagina-activa", data="inicio"),

                # ════════════════════════════════════════════════════════════
                # PÁGINA: INICIO
                # ════════════════════════════════════════════════════════════
                html.Div(id="pagina-inicio", children=[

                    # ── FILA KPIs GENERALES ───────────────────────────────────
                    dbc.Row(id="kpi-row", className="g-3 mb-3"),

                    # ── SEPARADOR + TÍTULO FILA PORTAFOLIOS ──────────────────
                    html.Div([
                        html.Div(style={
                            "height": "1px", "background": COLORS["border"],
                            "flex": "1",
                        }),
                        html.Span("RENTABILIDAD POR PORTAFOLIO", style={
                            "fontSize": "9px", "color": COLORS["text_muted"],
                            "letterSpacing": "2px", "padding": "0 12px",
                            "whiteSpace": "nowrap",
                        }),
                        html.Div(style={
                            "height": "1px", "background": COLORS["border"],
                            "flex": "1",
                        }),
                    ], style={
                        "display": "flex", "alignItems": "center",
                        "marginBottom": "10px",
                    }),

                    # ── FILA KPIs POR PORTAFOLIO ─────────────────────────────
                    dbc.Row(id="kpi-portafolios-row", className="g-2 mb-3",
                            style={"flexWrap": "nowrap", "display": "flex",
                                   "width": "100%"}),

                    dbc.Row([
                        dbc.Col(html.Div([
                            html.Div("CAUSACIÓN DIARIA DE MERCADO",
                                     style={"fontSize": "11px", "color": COLORS["text_muted"],
                                            "letterSpacing": "1px", "marginBottom": "12px"}),
                            dcc.Graph(id="chart-pnl", config={"displaylogo": False},
                                      style={"height": "320px"}),
                        ], style={"background": COLORS["bg_card"],
                                  "border": f"1px solid {COLORS['border']}",
                                  "borderRadius": "8px", "padding": "16px"}), md=7),

                        dbc.Col(html.Div([
                            html.Div("COMPOSICIÓN POR ESPECIE",
                                     style={"fontSize": "11px", "color": COLORS["text_muted"],
                                            "letterSpacing": "1px", "marginBottom": "12px"}),
                            dcc.Graph(id="chart-composicion", config={"displaylogo": False},
                                      style={"height": "320px"}),
                        ], style={"background": COLORS["bg_card"],
                                  "border": f"1px solid {COLORS['border']}",
                                  "borderRadius": "8px", "padding": "16px"}), md=5),
                    ], className="g-3 mb-3"),

                ]),  # fin pagina-inicio

                # ════════════════════════════════════════════════════════════
                # PÁGINA: ANÁLISIS DE PORTAFOLIO
                # ════════════════════════════════════════════════════════════
                html.Div(id="pagina-analisis", style={"display": "none"}, children=[

                    # Fila: Tabla dist vcto + Gráfica causación por plazo
                    dbc.Row([
                        dbc.Col(html.Div([
                            html.Div([
                                html.Span("DISTRIBUCIÓN POR AÑO DE VENCIMIENTO",
                                          style={"fontSize": "11px",
                                                 "color": COLORS["text_muted"],
                                                 "letterSpacing": "1px"}),
                                dcc.Dropdown(
                                    id="dist-especie-filter",
                                    options=[{"label": "Todas", "value": "Todas"}] +
                                            [{"label": e, "value": e}
                                             for e in sorted(df_global["Especie"].unique())],
                                    value="Todas", clearable=False,
                                    style={"width": "185px", "fontSize": "11px",
                                           "color": "#1f1f1f"},
                                ),
                            ], style={"display": "flex", "justifyContent": "space-between",
                                      "alignItems": "center", "marginBottom": "10px"}),
                            # Tabla con altura fija + scroll
                            html.Div(
                                id="tabla-dist-vcto",
                                style={"height": "340px", "overflowY": "auto"},
                            ),
                        ], style={"background": COLORS["bg_card"],
                                  "border": f"1px solid {COLORS['border']}",
                                  "borderRadius": "8px", "padding": "16px"}), md=6),

                        dbc.Col(html.Div([
                            html.Div("CAUSACIÓN ACUMULADA POR PLAZO DE VENCIMIENTO",
                                     style={"fontSize": "11px", "color": COLORS["text_muted"],
                                            "letterSpacing": "1px", "marginBottom": "12px"}),
                            dcc.Graph(id="chart-duracion", config={"displaylogo": False},
                                      style={"height": "340px"}),
                        ], style={"background": COLORS["bg_card"],
                                  "border": f"1px solid {COLORS['border']}",
                                  "borderRadius": "8px", "padding": "16px"}), md=6),
                    ], className="g-3 mb-3"),

                    # Fila: Tabla detalle posiciones (ancho completo)
                    dbc.Row([
                        dbc.Col(html.Div([
                            html.Div([
                                html.Span("DETALLE DE POSICIONES",
                                          style={"fontSize": "11px",
                                                 "color": COLORS["text_muted"],
                                                 "letterSpacing": "1px"}),
                                html.Span(id="tabla-subtitle",
                                          style={"fontSize": "11px",
                                                 "color": COLORS["accent"],
                                                 "marginLeft": "12px"}),
                            ], style={"marginBottom": "12px"}),
                            html.Div(id="tabla-detalle"),
                        ], style={"background": COLORS["bg_card"],
                                  "border": f"1px solid {COLORS['border']}",
                                  "borderRadius": "8px", "padding": "16px"}), md=12),
                    ], className="g-3 mb-3"),

                ]),  # fin pagina-analisis

                # Footer
                html.Div(
                    f"Dashboard · Fuente: {RUTA_ARCHIVO}",
                    style={"fontSize": "9px", "color": COLORS["text_muted"],
                           "textAlign": "center", "padding": "10px 0"},
                ),
            ]),
        ]),
    ]
)

# ─────────────────────────────────────────────────────────────────────────────
#  CALLBACK NAVEGACIÓN SIDEBAR
# ─────────────────────────────────────────────────────────────────────────────
@callback(
    Output("pagina-inicio",   "style"),
    Output("pagina-analisis", "style"),
    Output("nav-inicio",   "style"),
    Output("nav-analisis", "style"),
    Input("nav-inicio",   "n_clicks"),
    Input("nav-analisis", "n_clicks"),
)
def navegar(n_inicio, n_analisis):
    from dash import ctx
    pagina = ctx.triggered_id or "nav-inicio"

    estilo_visible = {"display": "block"}
    estilo_oculto  = {"display": "none"}

    def estilo_nav(activo):
        return {
            "display": "flex", "alignItems": "center",
            "padding": "11px 16px", "borderRadius": "8px",
            "cursor": "pointer", "marginBottom": "4px",
            "background": COLORS["accent"] + "22" if activo else "transparent",
            "borderLeft": f"3px solid {COLORS['accent']}" if activo
                          else f"3px solid transparent",
            "color": COLORS["accent"] if activo else COLORS["text_muted"],
            "transition": "all 0.2s",
        }

    if pagina == "nav-analisis":
        return estilo_oculto, estilo_visible, estilo_nav(False), estilo_nav(True)
    return estilo_visible, estilo_oculto, estilo_nav(True), estilo_nav(False)

# ─────────────────────────────────────────────────────────────────────────────
#  CALLBACK PRINCIPAL — DATOS
# ─────────────────────────────────────────────────────────────────────────────
@callback(
    Output("kpi-row",             "children"),
    Output("kpi-portafolios-row", "children"),   # ← NUEVA FILA PORTAFOLIOS
    Output("chart-composicion",   "figure"),
    Output("chart-pnl",           "figure"),
    Output("chart-duracion",      "figure"),
    Output("tabla-dist-vcto",     "children"),
    Output("tabla-detalle",       "children"),
    Output("tabla-subtitle",      "children"),
    Input("date-range",          "start_date"),
    Input("date-range",          "end_date"),
    Input("especie-filter",      "value"),
    Input("dist-especie-filter", "value"),
)
def actualizar_dashboard(start_date, end_date, especie, dist_especie):
    # ── Filtrar ──────────────────────────────────────────────────────────────
    fecha_ini = pd.to_datetime(start_date).date()
    fecha_fin = pd.to_datetime(end_date).date()

    # df completo por fechas (sin filtro de portafolio) — para la fila de rentabilidad
    df_todos_ports = filtrar_df(fecha_ini, fecha_fin)

    # df filtrado (con portafolio si aplica) — para KPIs, gráficos y tablas
    df = df_todos_ports.copy()
    if especie != "Todas":
        df = df[df["PORTAFOLIO"] == especie]

    if df.empty:
        vacio = go.Figure()
        vacio.update_layout(paper_bgcolor=COLORS["bg_card"], plot_bgcolor=COLORS["bg_card"])
        sin_datos = html.Div("Sin datos", style={"color": COLORS["text_muted"]})
        return [], [], vacio, vacio, vacio, sin_datos, sin_datos, ""

    # Última fecha disponible en el rango
    ultima_fecha = df["Fecha_Val"].max()
    df_hoy = df[df["Fecha_Val"] == ultima_fecha]

    # VPN día anterior
    fechas_ordenadas = sorted(df["Fecha_Val"].unique())

    vpn_inicial = 0
    vpn_final = df_hoy["Vlr Mer. Hoy"].sum()

    if len(fechas_ordenadas) > 1:
        fecha_anterior = fechas_ordenadas[-2]
        df_anterior = df[df["Fecha_Val"] == fecha_anterior]
        vpn_inicial = df_anterior["Vlr Mer. Hoy"].sum()

    # ── KPIs generales ───────────────────────────────────────────────────────
    vlr_total = df_hoy["Vlr Mer. Hoy"].sum()
    vlr_ant   = df_hoy["Vlr Mer. Ant"].sum()
    variacion = vlr_total - vlr_ant
    var_pct   = (variacion / vlr_ant * 100) if vlr_ant != 0 else 0
    causacion = df_hoy["Causación Mer"].sum()
    nominal   = df_hoy["Vlr Nominal"].sum()
    tir_pond  = (df_hoy["TIR.Mercado"] * df_hoy["Vlr Mer. Hoy"]).sum() / vlr_total \
                if vlr_total != 0 else 0
    n_titulos = df_hoy["ISIN/Nemotécnic"].nunique()
    n_inver   = df_hoy["Inver"].nunique()

    causacion_acumulada = df["Causación Mer"].sum()

    # ── Rentabilidad con lógica MTD (igual que tarjetas de portafolio) ────────
    fecha_fin_kpi      = df["Fecha_Val"].max()
    ultimo_dia_mes_kpi = (fecha_fin_kpi.to_period("M") - 1).to_timestamp("M")
    df_base_kpi        = df_global[
        df_global["Fecha_Val"].dt.date == ultimo_dia_mes_kpi.date()
    ]
    if especie != "Todas":
        df_base_kpi = df_base_kpi[df_base_kpi["PORTAFOLIO"] == especie]

    vpn_base_kpi = df_base_kpi["Vlr Mer. Hoy"].sum() if not df_base_kpi.empty else vlr_total

    rentabilidad_acumulada = (
        (vlr_total - vpn_base_kpi) / vpn_base_kpi if vpn_base_kpi != 0 else 0
    )

    # Rentabilidad diaria: causación hoy / vlr mercado ayer
    if len(fechas_ordenadas) > 1:
        rentabilidad_diaria = causacion / vpn_inicial if vpn_inicial != 0 else 0
    else:
        rentabilidad_diaria = causacion / vpn_base_kpi if vpn_base_kpi != 0 else 0

    kpis = dbc.Row([
        dbc.Col(
            kpi_card(
                "Valor de Mercado Total",
                fmt_cop(vlr_total),
                f"Variación: {fmt_cop(variacion)} ({var_pct:+.2f}%)",
                color_var(variacion),
                "💼"
            ),
            md=3
        ),
        dbc.Col(
            kpi_card(
                "Causación Acumulada",
                fmt_cop(causacion_acumulada),
                f"{fecha_ini.strftime('%d/%m/%Y')} → {fecha_fin.strftime('%d/%m/%Y')}",
                color_var(causacion_acumulada),
                "📈"
            ),
            md=3
        ),
        dbc.Col(
            kpi_card(
                "Causación Mercado (Día)",
                fmt_cop(causacion),
                f"P&L del día: {ultima_fecha.strftime('%d/%m/%Y')}",
                color_var(causacion),
                "📅"
            ),
            md=2
        ),
        dbc.Col(
            kpi_card(
                "RENTABILIDAD ACUMULADA",
                f"{rentabilidad_acumulada:.2%}",
                f"Diaria: {rentabilidad_diaria:.2%}",
                color_var(rentabilidad_diaria),
                "📊"
            ),
            md=2
        ),
        dbc.Col(
            kpi_card(
                "Títulos / Inversores",
                f"{n_titulos} / {n_inver}",
                "Diversificación del portafolio",
                COLORS["accent2"],
                "🔢"
            ),
            md=2
        ),
    ], className="g-3 mb-3")

    # ── KPIs POR PORTAFOLIO (nueva fila) ─────────────────────────────────────
    # Si hay filtro de portafolio activo, mostrar solo ese; si no, todos los del rango
    if especie != "Todas":
        portafolios_en_rango = [especie]
    else:
        portafolios_en_rango = sorted(df["PORTAFOLIO"].unique())

    # Paleta de colores de acento para diferenciar portafolios
    colores_port = [
        COLORS["accent"],   # azul
        COLORS["accent2"],  # verde
        COLORS["accent4"],  # amarillo
        COLORS["accent3"],  # rojo
        "#B57BFF",          # morado
        "#FF8C42",          # naranja
        "#4ECDC4",          # turquesa
        "#F7B731",          # dorado
    ]

    # Calcular ancho de columna para que quepan todos en una sola fila
    n_ports = len(portafolios_en_rango)
    if n_ports == 0:
        md_col = 2
    elif n_ports <= 6:
        md_col = 12 // n_ports
    else:
        md_col = max(1, 12 // n_ports)

    cols_portafolio = []
    for i, port in enumerate(portafolios_en_rango):
        ret_d, ret_a = calcular_rentabilidad_portafolio(df_todos_ports, port)
        color_port   = color_var(ret_a)

        cols_portafolio.append(
            dbc.Col(
                kpi_card_portafolio(
                    port,
                    f"{ret_a:.2%}",
                    f"Diaria: {ret_d:.2%}",
                    color_val=color_port,
                ),
                style={"flex": "1", "minWidth": "0", "padding": "0 4px"},
            )
        )

    kpis_portafolios = cols_portafolio  # lista de dbc.Col — dbc.Row ya está en el layout

    # ── Estilos base para gráficos ────────────────────────────────────────────
    layout_base = dict(
        paper_bgcolor=COLORS["bg_card"],
        plot_bgcolor=COLORS["bg_card"],
        font=dict(family="Segoe UI", color=COLORS["text"], size=11),
        margin=dict(l=12, r=12, t=12, b=12),
        legend=dict(bgcolor="rgba(0,0,0,0)", borderwidth=0),
        xaxis=dict(gridcolor=COLORS["border"], zerolinecolor=COLORS["border"],
                   tickfont=dict(size=10)),
        yaxis=dict(gridcolor=COLORS["border"], zerolinecolor=COLORS["border"],
                   tickfont=dict(size=10)),
    )

    colores_esp = [COLORS["accent"], COLORS["accent2"], COLORS["accent4"]]

    # Detectar columna CLASIFICACION una sola vez (usada en CHART 1 y CHART 2)
    col_clas = None
    for c in df.columns:
        if c.strip().upper() in ["CLASIFICACION", "CLASIFICACIÓN", "CLASSIFICATION"]:
            col_clas = c
            break

    # ── CHART 1: Composición (donut) ─────────────────────────────────────────
    if col_clas:
        comp = df_hoy.groupby(col_clas)["Vlr Mer. Hoy"].sum().reset_index()
        comp = comp.rename(columns={col_clas: "CLASIFICACION"})
        comp_labels = comp["CLASIFICACION"]
    else:
        comp = df_hoy.groupby("Especie")["Vlr Mer. Hoy"].sum().reset_index()
        comp_labels = comp["Especie"]

    fig_comp = go.Figure(go.Pie(
        labels=comp_labels, values=comp["Vlr Mer. Hoy"],
        hole=0.55,
        marker=dict(colors=colores_esp, line=dict(width=2)),
        textinfo="label+percent",
        textfont=dict(size=10, color=COLORS["text"]),
        hovertemplate="<b>%{label}</b><br>$%{value:,.0f}<br>%{percent}<extra></extra>",
    ))
    total_str = fmt_cop(vlr_total)
    fig_comp.update_layout(
        **{k: v for k, v in layout_base.items() if k not in ["xaxis", "yaxis"]},
        annotations=[dict(text=f"<b>{total_str}</b>", x=0.5, y=0.5,
                          font=dict(size=12, color=COLORS["text"]), showarrow=False)],
    )

    # ── CHART 2: P&L diario ───────────────────────────────────────────────────
    if col_clas:
        pnl = df.groupby(["Fecha_Val", col_clas])["Causación Mer"].sum().reset_index()
        pnl = pnl.rename(columns={col_clas: "CLASIFICACION"})
        grupo_col = "CLASIFICACION"
    else:
        # Fallback a Especie si no existe la columna
        pnl = df.groupby(["Fecha_Val", "Especie"])["Causación Mer"].sum().reset_index()
        grupo_col = "Especie"

    pnl_total = df.groupby("Fecha_Val")["Causación Mer"].sum().reset_index()
    fig_pnl = go.Figure()
    for i, grp in enumerate(pnl[grupo_col].unique()):
        sub = pnl[pnl[grupo_col] == grp]
        fig_pnl.add_trace(go.Bar(
            x=sub["Fecha_Val"], y=sub["Causación Mer"], name=grp,
            marker_color=colores_esp[i % len(colores_esp)],
            opacity=0.85,
            hovertemplate="<b>%{x|%d/%m/%Y}</b><br>P&L: $%{y:,.0f}<extra></extra>"
        ))
    fig_pnl.add_trace(go.Scatter(
        x=pnl_total["Fecha_Val"], y=pnl_total["Causación Mer"],
        name="Total", mode="lines+markers",
        line=dict(color=COLORS["accent4"], width=2, dash="dot"),
        marker=dict(size=8, color=COLORS["accent4"]),
    ))
    fig_pnl.update_layout(**layout_base, barmode="group")
    fig_pnl.update_yaxes(tickformat="$,.0f")

    # ── CHART 3: Causación acumulada por plazo ────────────────────────────────
    df_dur = df[df["Especie"] == dist_especie] if dist_especie != "Todas" else df
    dur_df = df_dur.copy()
    dur_df["Bucket"] = pd.cut(
        dur_df["Dias"],
        bins=[0, 365, 730, 1825, 3650, 7300, 20000],
        labels=["< 1 año", "1-2 años", "2-5 años", "5-10 años", "10-20 años", "> 20 años"],
    )
    dur_agg = (
        dur_df.groupby(["Bucket", "Especie"], observed=True)["Causacion_Total"]
        .sum()
        .reset_index()
    )
    fig_dur = go.Figure()
    for i, esp in enumerate(sorted(dur_agg["Especie"].unique())):
        sub = dur_agg[dur_agg["Especie"] == esp]
        fig_dur.add_trace(go.Bar(
            x=sub["Bucket"].astype(str),
            y=sub["Causacion_Total"],
            name=esp,
            marker_color=colores_esp[i % len(colores_esp)],
            hovertemplate="<b>%{x}</b><br>Causación: $%{y:,.2f}<extra></extra>",
            text=sub["Causacion_Total"].apply(lambda v: f"{v:,.0f}"),
            textposition="outside",
            textfont=dict(size=9),
        ))
    fig_dur.update_layout(**layout_base, barmode="stack")
    fig_dur.update_yaxes(tickformat="$,.0f", title_text="Causación Acumulada")
    fig_dur.update_xaxes(title_text="Plazo al Vencimiento")

    # ── TABLA DISTRIBUCIÓN POR AÑO DE VENCIMIENTO ────────────────────────────
    df_dist = df[df["Especie"] == dist_especie] if dist_especie != "Todas" else df
    dist = df_dist.groupby("Año_Vcto").agg(
        Vlr_Mer_Hoy=("Vlr Mer. Hoy",    "sum"),
        Causacion_Total=("Causacion_Total", "sum"),
    ).reset_index().sort_values("Año_Vcto")

    total_causacion = dist["Causacion_Total"].sum()
    dist["Participacion"] = (dist["Causacion_Total"] / total_causacion * 100) if total_causacion else 0

    fila_total = pd.DataFrame([{
        "Año_Vcto":        "TOTAL",
        "Vlr_Mer_Hoy":     0,
        "Causacion_Total": dist["Causacion_Total"].sum(),
        "Participacion":   100.0,
    }])
    dist_display = pd.concat([dist, fila_total], ignore_index=True)

    dist_display["Año_Vcto"]       = dist_display["Año_Vcto"].astype(str)
    dist_display["Participacion"]  = dist_display["Participacion"].apply(
        lambda x: f"{x:.2f}%" if pd.notna(x) else "—"
    )
    dist_display["Vlr_Mer_Hoy"]    = dist_display["Vlr_Mer_Hoy"].apply(fmt_cop)
    dist_display["Causacion_Total"]= dist_display["Causacion_Total"].apply(
        lambda x: f"{x:,.2f}" if pd.notna(x) else "—"
    )

    dist_display.columns = [
        "Año Vcto.", "Val. Mercado Hoy", "Causación Total", "% Aporte Causación"
    ]
    dist_display = dist_display[["Año Vcto.", "% Aporte Causación", "Causación Total"]]

    idx_total = len(dist_display) - 1

    tabla_dist = dash_table.DataTable(
        data=dist_display.to_dict("records"),
        columns=[{"name": c, "id": c} for c in dist_display.columns],
        style_table={"overflowX": "auto"},
        style_header={
            "backgroundColor": COLORS["bg_card2"],
            "color": COLORS["accent"],
            "fontWeight": "700",
            "fontSize": "11px",
            "border": f"1px solid {COLORS['border']}",
            "letterSpacing": "0.5px",
            "textAlign": "center",
        },
        style_cell={
            "backgroundColor": COLORS["bg_card"],
            "color": COLORS["text"],
            "fontSize": "12px",
            "border": f"1px solid {COLORS['border']}",
            "padding": "8px 14px",
            "textAlign": "right",
        },
        style_cell_conditional=[
            {"if": {"column_id": "Año Vcto."},
             "textAlign": "center", "fontWeight": "600", "width": "90px"},
            {"if": {"column_id": "% Aporte Causación"},
             "textAlign": "center"},
        ],
        style_data_conditional=[
            {"if": {"row_index": idx_total},
             "backgroundColor": COLORS["bg_card2"],
             "fontWeight": "700",
             "borderTop": f"2px solid {COLORS['accent']}",
             "color": COLORS["accent"]},
            {"if": {"filter_query": "{Causación Total} contains '-'"},
             "color": COLORS["negative"]},
        ],
        sort_action="native",
        sort_mode="single",
        page_action="none",
    )

    # ── TABLA DETALLE ─────────────────────────────────────────────────────────
    tabla_df = df_hoy[[
        "ISIN/Nemotécnic", "Especie", "Vlr Nominal", "Precio",
        "Vlr Mer. Hoy", "Variacion_Abs", "Variacion_Pct",
        "Causación Mer", "TIR.Mercado", "Dias",
    ]].copy()

    tabla_df.columns = [
        "ISIN", "Especie", "Val. Nominal", "Precio",
        "Val. Mercado Hoy", "Variación $", "Variación %",
        "Causación Mer.", "TIR Mercado", "Días",
    ]

    for col in ["Val. Nominal", "Val. Mercado Hoy", "Variación $", "Causación Mer."]:
        tabla_df[col] = tabla_df[col].apply(fmt_cop)
    tabla_df["Precio"]      = tabla_df["Precio"].apply(lambda x: f"{x:.3f}" if pd.notna(x) else "—")
    tabla_df["TIR Mercado"] = tabla_df["TIR Mercado"].apply(lambda x: f"{x:.4f}%" if pd.notna(x) else "—")
    tabla_df["Variación %"] = tabla_df["Variación %"].apply(lambda x: f"{x:+.2f}%" if pd.notna(x) else "—")
    tabla_df["Días"]        = tabla_df["Días"].apply(lambda x: f"{int(x):,}" if pd.notna(x) else "—")

    tabla = dash_table.DataTable(
        data=tabla_df.to_dict("records"),
        columns=[{"name": c, "id": c} for c in tabla_df.columns],
        style_table={"overflowX": "auto"},
        style_header={
            "backgroundColor": COLORS["bg_card2"],
            "color": COLORS["accent"],
            "fontWeight": "700",
            "fontSize": "11px",
            "border": f"1px solid {COLORS['border']}",
            "letterSpacing": "0.5px",
        },
        style_cell={
            "backgroundColor": COLORS["bg_card"],
            "color": COLORS["text"],
            "fontSize": "12px",
            "border": f"1px solid {COLORS['border']}",
            "padding": "8px 12px",
            "textAlign": "right",
        },
        style_cell_conditional=[
            {"if": {"column_id": "ISIN"}, "textAlign": "left", "fontWeight": "600"},
            {"if": {"column_id": "Especie"}, "textAlign": "left"},
        ],
        style_data_conditional=[
            {"if": {"filter_query": "{Variación %} contains '+'"},
             "color": COLORS["positive"]},
            {"if": {"filter_query": "{Variación %} contains '-'"},
             "color": COLORS["negative"]},
        ],
        sort_action="native",
        filter_action="native",
        page_size=15,
        page_action="native",
    )

    subtitle = (
        f"Mostrando datos al {ultima_fecha.strftime('%d/%m/%Y')} "
        f"| {len(tabla_df)} posiciones"
    )

    return (kpis, kpis_portafolios, fig_comp, fig_pnl, fig_dur,
            tabla_dist, tabla, subtitle)

# ─────────────────────────────────────────────────────────────────────────────
#  PUNTO DE ENTRADA
# ─────────────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    print("=" * 65)
    print("  DASHBOARD DE PORTAFOLIO - VALORACIÓN")
    print("=" * 65)
    print(f"  Fuente de datos : {RUTA_ARCHIVO}")
    print(f"  Fechas cargadas : {min(fechas_disponibles)} → {max(fechas_disponibles)}")
    print(f"  Registros       : {len(df_global):,}")
    print(f"  Portafolios     : {', '.join(portafolios_disponibles)}")
    print("=" * 65)
    print("  Abre tu navegador en:  http://127.0.0.1:8050")
    print("  Presiona Ctrl+C para detener el servidor")
    print("=" * 65)
    app.run(debug=False, port=8050, host="127.0.0.1")
