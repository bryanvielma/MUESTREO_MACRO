import os
import re
import base64
import hashlib
import numpy as np
import pandas as pd
from io import BytesIO
from datetime import datetime
import dash
from dash import dcc, html, Input, Output, State
import dash_bootstrap_components as dbc
import xlsxwriter
import warnings

warnings.filterwarnings("ignore", category=UserWarning)

# =============================================================================
# CONFIGURACIÓN DE RUTAS (RENDER SAFE)
# =============================================================================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

OUTPUT_DIR = os.path.join(BASE_DIR, "output")
IMAGES_DIR = os.path.join(BASE_DIR, "images")

ARCHIVO_EXCEL = os.path.join(OUTPUT_DIR, "Muestreos_Activos.xlsx")
IMAGEN_RECORTADA = os.path.join(IMAGES_DIR, "imagen_recortada.jpg")

# =============================================================================
# VALIDACIÓN DE ARCHIVOS
# =============================================================================
if not os.path.exists(ARCHIVO_EXCEL):
    raise Exception(f"❌ No se encontró {ARCHIVO_EXCEL} en el repositorio")

# =============================================================================
# CARGA DE DATOS
# =============================================================================
muestreos_hoy = pd.read_excel(ARCHIVO_EXCEL, sheet_name="Hoy")
muestreos_proximos = pd.read_excel(ARCHIVO_EXCEL, sheet_name="Proximos")

for df in [muestreos_hoy, muestreos_proximos]:
    for col in ['Macetas actuales', 'Alveolos', 'Bandeja']:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce')

    if 'fecha_activadora' in df.columns:
        df['fecha_activadora'] = pd.to_datetime(df['fecha_activadora'], errors='coerce')

# =============================================================================
# FUNCIÓN PARA EXTRAER CANTIDAD DESDE I-M-C
# =============================================================================
def extraer_cantidad_desde_imc(imc_str):
    if not isinstance(imc_str, str):
        return 0
    numeros = re.findall(r'-C(\d+)', imc_str)
    return sum(int(num) for num in numeros) if numeros else 0

# =============================================================================
# APP DASH
# =============================================================================
app = dash.Dash(__name__, external_stylesheets=[dbc.themes.FLATLY])
server = app.server  # 🔴 CLAVE PARA RENDER

app.title = "Cálculo de Tamaño de Muestras"

app.layout = dbc.Container([
    html.H1("Cálculo de Tamaño de Muestras", className="text-center my-4"),
    dbc.Row([
        dbc.Col([
            dcc.Dropdown(
                id="dropdown-lote-hoy",
                options=[
                    {"label": f"{row['Código']} - {row.get('Variedad', '')}", "value": row["Código"]}
                    for _, row in muestreos_hoy.iterrows()
                ],
                placeholder="Seleccione un lote con muestreo hoy",
                className="mb-3"
            ),
            dcc.Dropdown(
                id="dropdown-lote",
                options=[
                    {"label": f"{row['Código']} - {row.get('Variedad', '')}", "value": row["Código"]}
                    for _, row in muestreos_proximos.iterrows()
                ],
                placeholder="Seleccione un lote próximo",
                className="mb-3"
            ),
            dbc.Button("Calcular", id="btn-calcular", color="primary", className="w-100"),
            html.A("Descargar Excel", id="btn-descargar", href="", download="", className="btn btn-success mt-3 w-100"),
            dbc.Card([
                dbc.CardHeader("Resultados"),
                dbc.CardBody(html.Div(id="resultados"))
            ], className="mt-4"),
        ], width=12)
    ])
], fluid=True)

# =============================================================================
# CALLBACK
# =============================================================================
@app.callback(
    [Output("resultados", "children"),
     Output("btn-descargar", "href"),
     Output("btn-descargar", "download")],
    Input("btn-calcular", "n_clicks"),
    [State("dropdown-lote-hoy", "value"),
     State("dropdown-lote", "value")]
)
def calcular_muestra_y_generar_excel(n_clicks, codigo_hoy, codigo):

    codigo = codigo_hoy if codigo_hoy else codigo

    if not n_clicks or not codigo:
        return "Seleccione un lote y presione 'Calcular'.", "", ""

    df_origen = muestreos_hoy if codigo_hoy else muestreos_proximos

    try:
        lote = df_origen[df_origen["Código"] == codigo].iloc[0]

        # -------- DATOS BASE --------
        imc_raw = lote.get("I-M-C", "") or ""
        cantidad = extraer_cantidad_desde_imc(str(imc_raw))

        if cantidad == 0:
            cantidad = int(lote.get("Macetas actuales", 0) or 0)

        bandejas_val = int(lote.get("Bandeja", 0) or 0)

        macetero_raw = lote.get("Macetero", "") or ""
        litros = 0.0
        match = re.search(r'(\d+(?:[.,]\d+)?)\s*L', str(macetero_raw))
        if match:
            litros = float(match.group(1).replace(",", "."))

        # -------- MUESTREO --------
        hileras = 20

        def calcular_tamano(c):
            for limite, muestra in [
                (8, 2), (15, 3), (25, 5), (50, 8), (90, 13),
                (150, 20), (280, 40), (500, 60), (1200, 80),
                (3200, 140), (10000, 200), (35000, 320),
                (150000, 500), (500000, 800)
            ]:
                if c <= limite:
                    return muestra
            return 1260

        muestra_tamano = calcular_tamano(cantidad)

        rows_count = max(cantidad // hileras, 1)
        rows_needed = max(muestra_tamano // hileras, 1)

        seed = int(hashlib.sha256(str(codigo).encode()).hexdigest(), 16) % (10**6)
        np.random.seed(seed)

        chosen_rows = sorted(np.random.choice(range(1, rows_count+1), size=rows_needed, replace=False))

        rows = []
        muestra_num = 1
        for r in chosen_rows:
            for i in range(hileras):
                rows.append({
                    "Muestra": muestra_num,
                    "Número Planta": (r-1)*hileras + i + 1,
                    "Fila": r
                })
                muestra_num += 1

        tabla_df = pd.DataFrame(rows)

        # -------- EXCEL --------
        output = BytesIO()

        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            tabla_df.to_excel(writer, index=False)

        output.seek(0)
        excel_data = base64.b64encode(output.read()).decode()

        href = f"data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{excel_data}"

        download_name = f"Muestreo_{codigo}.xlsx"

        # -------- RESULTADOS --------
        resultados = [
            html.P(f"Código: {codigo}"),
            html.P(f"Cantidad: {cantidad}"),
            html.P(f"Muestra: {muestra_tamano}"),
            html.P(f"Bandeja: {bandejas_val}"),
            html.P(f"Volumen: {litros:.2f} L"),
        ]

        return resultados, href, download_name

    except Exception as e:
        return f"Error: {str(e)}", "", ""

# =============================================================================
# RUN LOCAL
# =============================================================================
if __name__ == "__main__":
    app.run(debug=True)