import os
import re
import base64
import hashlib
import numpy as np
import pandas as pd
from io import BytesIO
from datetime import datetime
import dash
from dash import dcc, html, Input, Output, State, dash_table
import dash_bootstrap_components as dbc
import plotly.express as px
import xlsxwriter
import warnings

warnings.filterwarnings("ignore")

# =============================================================================
# CONFIGURACIÓN DE RUTAS
# =============================================================================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.path.join(BASE_DIR, "output")
IMAGES_DIR = os.path.join(BASE_DIR, "images")
ARCHIVO_EXCEL = os.path.join(OUTPUT_DIR, "Muestreos_Activos.xlsx")
IMAGEN_RECORTADA = os.path.join(IMAGES_DIR, "imagen_recortada.jpg")

# =============================================================================
# CARGA Y FILTRADO DE DATOS
# =============================================================================
if not os.path.exists(ARCHIVO_EXCEL):
    raise FileNotFoundError(f"No se encontró {ARCHIVO_EXCEL}. Asegúrate de subir el archivo.")

muestreos_hoy_raw = pd.read_excel(ARCHIVO_EXCEL, sheet_name="Hoy")
muestreos_proximos_raw = pd.read_excel(ARCHIVO_EXCEL, sheet_name="Proximos")

# Función para identificar lotes con "MN" (Vivero los Viñedos - Perú)
def es_lote_peru(row):
    imc = row.get("I-M-C", "")
    return isinstance(imc, str) and "MN" in imc.upper()

# Filtrar por fecha actual (solo para hoy)
hoy_date = datetime.now().date()
if 'fecha_activadora' in muestreos_hoy_raw.columns:
    muestreos_hoy_raw['fecha_activadora'] = pd.to_datetime(muestreos_hoy_raw['fecha_activadora'], errors='coerce')
    mascara_fecha = muestreos_hoy_raw['fecha_activadora'].dt.date == hoy_date
    muestreos_hoy_raw = muestreos_hoy_raw[mascara_fecha].copy()

# IDs excluidos (con MN) de hoy
ids_excluidos_hoy = muestreos_hoy_raw[muestreos_hoy_raw.apply(es_lote_peru, axis=1)]["ID"].tolist()
ids_excluidos = sorted(set(ids_excluidos_hoy))

# Filtrar lotes que NO contienen "MN" (los que sí se muestrean)
def filtrar_sin_mn(df):
    if 'I-M-C' not in df.columns:
        return df
    mask = df.apply(lambda row: not es_lote_peru(row), axis=1)
    return df[mask].copy()

muestreos_hoy = filtrar_sin_mn(muestreos_hoy_raw)
muestreos_proximos = filtrar_sin_mn(muestreos_proximos_raw)

# Convertir columnas numéricas
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
    patron = r'-C(\d+)'
    numeros = re.findall(patron, imc_str)
    if not numeros:
        return 0
    return sum(int(num) for num in numeros)

# =============================================================================
# FUNCIÓN PARA GENERAR DATOS DE UN LOTE (similar a antes, pero reutilizable)
# =============================================================================
def generar_datos_lote(lote):
    imc_raw = lote.get("I-M-C", "")
    if pd.isna(imc_raw):
        imc_raw = ""

    cantidad = extraer_cantidad_desde_imc(str(imc_raw))
    if cantidad == 0:
        cantidad = lote.get("Macetas actuales", 0)
        if pd.isna(cantidad):
            cantidad = 0
        else:
            cantidad = int(cantidad)
    else:
        cantidad = int(cantidad)

    imc_val = imc_raw
    bandejas_val = lote.get("Bandeja", 0)
    if pd.isna(bandejas_val):
        bandejas_val = 0
    macetero_raw = lote.get("Macetero", "")
    if pd.isna(macetero_raw):
        macetero_raw = ""
    litros = 0.0
    vol_match = re.search(r'(\d+(?:[.,]\d+)?)\s*L', str(macetero_raw))
    if vol_match:
        litros = float(vol_match.group(1).replace(",", "."))

    hileras = 20

    def calcular_tamano(c):
        for limite, muestra in [(8, 2), (15, 3), (25, 5), (50, 8), (90, 13),
                                (150, 20), (280, 40), (500, 60), (1200, 80),
                                (3200, 140), (10000, 200), (35000, 320),
                                (150000, 500), (500000, 800)]:
            if c <= limite:
                return muestra
        return 1260

    muestra_tamano = calcular_tamano(cantidad)
    rows_count = max(cantidad // hileras, 1)
    full_rows = muestra_tamano // hileras
    remainder = muestra_tamano % hileras
    rows_needed = full_rows + (1 if remainder > 0 else 0)

    seed = int(hashlib.sha256(str(lote["Código"]).encode()).hexdigest(), 16) % (10**6)
    np.random.seed(seed)
    spacing = max(rows_count // rows_needed, 1)
    primera_fila = np.random.randint(1, spacing + 1)
    chosen_rows = sorted(set(min(primera_fila + i * spacing, rows_count) for i in range(rows_needed)))

    rows = []
    muestra_num = 1
    for i, r_idx in enumerate(chosen_rows):
        start_plant = (r_idx - 1) * hileras + 1
        if (i == len(chosen_rows) - 1) and remainder > 0:
            for off_p in range(remainder):
                plant_num = start_plant + off_p
                rows.append({"Muestra": muestra_num, "Número Planta": plant_num, "Fila": r_idx})
                muestra_num += 1
        else:
            for off_p in range(hileras):
                plant_num = start_plant + off_p
                rows.append({"Muestra": muestra_num, "Número Planta": plant_num, "Fila": r_idx})
                muestra_num += 1

    if not rows:
        raise ValueError("No se generaron datos de muestra.")

    tabla_df = pd.DataFrame(rows)
    return {
        "tabla_df": tabla_df,
        "lote": lote,
        "cantidad": cantidad,
        "imc_val": imc_val,
        "bandejas_val": bandejas_val,
        "litros": litros,
        "muestra_tamano": muestra_tamano
    }

# =============================================================================
# FUNCIÓN PARA ESCRIBIR UNA HOJA EN EL WORKBOOK (reutilizable)
# =============================================================================
def escribir_hoja(workbook, datos, nombre_hoja):
    worksheet = workbook.add_worksheet(nombre_hoja[:31])
    fmt_bold = workbook.add_format({'bold': True, 'border': 1, 'align': 'center', 'valign': 'vcenter'})
    fmt_norm = workbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter', 'font_size': 10})
    fmt_center = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'font_size': 10, 'bold': True})
    fmt_small = workbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter', 'font_size': 9})

    start_row = 4
    lote = datos["lote"]
    tabla_df = datos["tabla_df"]
    cantidad = datos["cantidad"]
    imc_val = datos["imc_val"]
    bandejas_val = datos["bandejas_val"]
    litros = datos["litros"]

    # Información del lote
    info = {
        "ID": lote.get("ID", "N/A"),
        "Fecha Inicial": lote.get("Fecha", "").strftime('%d-%m-%Y') if pd.notnull(lote.get("Fecha")) else "N/A",
        "Especie": lote.get("Especie", "N/A"),
        "Variedad": lote.get("Variedad", "N/A"),
        "Muestreo": lote.get("muestreo_activador", "N/A"),
        "Fecha Muestreo": lote["fecha_activadora"].strftime('%d-%m-%Y') if pd.notnull(lote.get("fecha_activadora")) else "N/A",
        "Alveolos": cantidad,
        "Muestra": len(tabla_df)
    }
    for col, (key, val) in enumerate(info.items()):
        worksheet.write(start_row, col, key, fmt_bold)
        worksheet.write(start_row+1, col, val, fmt_norm)
    worksheet.set_row(start_row, 15)
    worksheet.set_row(start_row+1, 12)
    start_row += 2

    # Bandeja, Volumen, Código
    worksheet.write(start_row, 0, "Bandeja", fmt_bold)
    worksheet.write(start_row, 1, "Vol. Sustrato (L)", fmt_bold)
    worksheet.write(start_row, 2, "Código", fmt_bold)
    worksheet.write(start_row+1, 0, bandejas_val, fmt_norm)
    worksheet.write(start_row+1, 1, f"{litros:.2f}", fmt_norm)
    worksheet.write(start_row+1, 2, lote["Código"], fmt_norm)
    worksheet.set_row(start_row, 12)
    worksheet.set_row(start_row+1, 12)
    start_row += 2

    # I-M-C
    worksheet.merge_range(f'D{start_row-1}:H{start_row-1}', 'INVERNADERO - MESÓN - CANTIDAD (I-M-C)', fmt_bold)
    worksheet.merge_range(f'D{start_row}:H{start_row}', str(imc_val), fmt_norm)
    worksheet.set_row(start_row-1, 12)
    worksheet.set_row(start_row, 12)
    start_row += 2

    # Tabla de resumen por fila
    worksheet.set_row(8, 5)
    header_row = 9
    filas_resumen = tabla_df.groupby("Fila").size().reset_index(name="Cantidad de Repeticiones")
    cols_vacias = ["Sobrevivencia", "Ejes ≥ 2", "Ocup sustrato ≥ 80%", "Altura ≥ 12 cm", "Talla Comercial", "% Col"]
    all_cols = ["Fila", "Máximo"] + cols_vacias
    for c, name in enumerate(all_cols):
        worksheet.write(header_row, c, name, fmt_bold)
    worksheet.set_row(header_row, 12)
    data_row = header_row + 1
    for idx, row in filas_resumen.iterrows():
        worksheet.write(data_row+idx, 0, row["Fila"], fmt_norm)
        worksheet.write(data_row+idx, 1, row["Cantidad de Repeticiones"], fmt_norm)
        for c in range(2, len(all_cols)):
            worksheet.write(data_row+idx, c, "", fmt_norm)
        worksheet.set_row(data_row+idx, 11)
    last_data = data_row + len(filas_resumen) - 1

    # Responsable, fecha, firma
    blank = last_data + 1
    worksheet.set_row(blank, 5)
    resp_row = blank + 1
    worksheet.merge_range(f'A{resp_row+1}:D{resp_row+1}', 'Responsable: _________________________________________________', fmt_center)
    worksheet.merge_range(f'F{resp_row+1}:H{resp_row+1}', 'Fecha: ______ /______ /_________', fmt_center)
    worksheet.write(f'E{resp_row+1}', 'Firma: ___________', fmt_center)
    worksheet.set_row(resp_row, 12)

    # Porcentajes
    blank2 = resp_row + 1
    worksheet.set_row(blank2, 5)
    pct_row = blank2 + 1
    worksheet.merge_range(f'A{pct_row+1}:B{pct_row+1}', '% PLANTAS PLANTABLES', fmt_small)
    worksheet.write(f'C{pct_row+1}', '', fmt_small)
    worksheet.merge_range(f'F{pct_row+1}:G{pct_row+1}', '% TALLA COMERCIAL', fmt_small)
    worksheet.write(f'H{pct_row+1}', '', fmt_small)
    worksheet.set_row(pct_row, 10)

    # Encabezado superior (imagen, título)
    worksheet.merge_range('A1:B2', '', fmt_bold)
    if os.path.exists(IMAGEN_RECORTADA):
        worksheet.insert_image('A1', IMAGEN_RECORTADA, {
            'x_scale': 30/45, 'y_scale': 22/65, 'x_offset': 12, 'y_offset': 5, 'positioning': 1
        })
    titulo = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'font_name': 'Comic Sans MS', 'font_size': 11, 'text_wrap': True, 'border': 1})
    worksheet.merge_range('C1:E2', 'Sociedad de Investigación, Desarrollo y Servicios de Biotecnología Aplicada Ltda.', titulo)
    cell_fmt = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'font_name': 'Arial', 'font_size': 10, 'text_wrap': True, 'border': 1})
    worksheet.write('F1', 'RAC-XXX', cell_fmt)
    worksheet.write('G1', 'POE XXX', cell_fmt)
    worksheet.write('F2', 'Edición 00', cell_fmt)
    worksheet.write('G2', 'Pág. 1 de 1', cell_fmt)
    titulo2 = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'font_name': 'Arial', 'font_size': 11, 'text_wrap': True, 'border': 1})
    worksheet.merge_range('A3:E3', 'REGISTRO PARA EL CONTROL DE LA TALLA COMERCIAL EN MACRO', titulo2)
    worksheet.write('F3', 'Vigente: 01/ 01/2025', cell_fmt)
    worksheet.write('G3', 'Folio:', cell_fmt)

    # Anchos de columna
    for col, w in [('A',8),('B',13),('C',23),('D',12),('E',18),('F',18),('G',13),('H',9)]:
        worksheet.set_column(f'{col}:{col}', w)
    worksheet.set_row(3, 5)

# =============================================================================
# APP DASH
# =============================================================================
app = dash.Dash(__name__, external_stylesheets=[dbc.themes.FLATLY])
server = app.server
app.title = "MACRO - Muestreo y Supervivencia"

app.layout = dbc.Container([
    html.H1("Sistema de Gestión de Muestreos", className="text-center my-4"),
    dcc.Tabs(id="tabs", value="tab-muestra", children=[
        dcc.Tab(label="📊 Cálculo de Tamaño de Muestra", value="tab-muestra"),
        dcc.Tab(label="📈 Análisis de Supervivencia", value="tab-supervivencia"),
    ]),
    html.Div(id="tab-content", className="mt-3")
], fluid=True)

# =============================================================================
# CONTENIDO DE LA PESTAÑA 1 (MULTIPLE)
# =============================================================================
@app.callback(
    Output("tab-content", "children"),
    Input("tabs", "value")
)
def render_tab(tab):
    if tab == "tab-muestra":
        # Mensaje de exclusión
        mensaje_exclusion = None
        if ids_excluidos:
            ids_texto = ", ".join(str(id_) for id_ in ids_excluidos)
            mensaje_exclusion = dbc.Alert(
                [html.I(className="fas fa-info-circle me-2"), 
                 f"⚠️ Los siguientes IDs corresponden a lotes en Vivero los Viñedos (PERÚ)  y no se incluyen en los muestreos de hoy: {ids_texto}"],
                color="warning",
                dismissable=True,
                className="mt-2"
            )
        
        # Mostrar información de lotes disponibles
        info_lotes = html.Div()
        if not muestreos_hoy.empty:
            info_lotes = dbc.Card(
                dbc.CardBody([
                    html.H5(f"Lotes a muestrear hoy ({len(muestreos_hoy)}):", className="card-title"),
                    html.Ul([html.Li(f"{row['Código']} - {row.get('Variedad', '')} (ID: {row['ID']})") for _, row in muestreos_hoy.iterrows()])
                ]),
                className="mb-3"
            )
        else:
            info_lotes = dbc.Alert("No hay lotes programados para hoy (sin MN).", color="info")
        
        return dbc.Row([
            dbc.Col([
                mensaje_exclusion if mensaje_exclusion else html.Div(),
                info_lotes,
                dbc.Button("Generar Excel de muestreo hoy", id="btn-generar-multiple", color="primary", className="w-100 mb-3"),
                html.A("Descargar Excel", id="btn-descargar-multiple", href="", download="", className="btn btn-success w-100"),
                dbc.Card([
                    dbc.CardHeader("Resultado de la generación"),
                    dbc.CardBody(html.Div(id="resultado-multiple"))
                ], className="mt-4"),
            ], width=12)
        ])
    else:
        # Pestaña de supervivencia (sin cambios)
        return dbc.Container([
            dbc.Row([
                dbc.Col([
                    dcc.Upload(
                        id='upload-data',
                        children=html.Div([
                            'Arrastra y suelta o ',
                            html.A('Selecciona un archivo Excel')
                        ]),
                        style={
                            'width': '100%', 'height': '60px', 'lineHeight': '60px',
                            'borderWidth': '1px', 'borderStyle': 'dashed', 'borderRadius': '5px',
                            'textAlign': 'center', 'margin': '10px'
                        },
                        multiple=False
                    ),
                    html.Div(id='output-alertas', style={'marginTop': '20px'}),
                    html.Div(id='output-data-upload', style={'marginTop': '20px'}),
                ], width=12)
            ]),
            dbc.Row([
                dbc.Col(dcc.Graph(id="grafico-supervivencia", config={'displayModeBar': False}), width=4, style={'height': '450px'}),
                dbc.Col(dcc.Graph(id="grafico-talla-comercial", config={'displayModeBar': False}), width=4, style={'height': '450px'}),
                dbc.Col(dcc.Graph(id="grafico-ejes", config={'displayModeBar': False}), width=4, style={'height': '450px'}),
            ], className="mt-3", style={'marginBottom': '20px'}),
            dbc.Row([
                dbc.Col(dcc.Graph(id="grafico-ocupacion", config={'displayModeBar': False}), width=4, style={'height': '450px'}),
                dbc.Col(dcc.Graph(id="grafico-altura", config={'displayModeBar': False}), width=4, style={'height': '450px'}),
                dbc.Col(dcc.Graph(id="grafico-porcentaje-col", config={'displayModeBar': False}), width=4, style={'height': '450px'}),
            ], className="mt-3", style={'marginBottom': '30px'}),
        ], fluid=True)

# =============================================================================
# CALLBACK PARA GENERAR EXCEL MÚLTIPLE
# =============================================================================
@app.callback(
    [Output("btn-descargar-multiple", "href"),
     Output("btn-descargar-multiple", "download"),
     Output("resultado-multiple", "children")],
    Input("btn-generar-multiple", "n_clicks"),
    prevent_initial_call=True
)
def generar_excel_multiple(n_clicks):
    if not n_clicks:
        return "", "", "Presione el botón para generar el Excel."
    
    if muestreos_hoy.empty:
        return "", "", "No hay lotes para muestrear hoy (todos contienen MN o no hay datos)."
    
    fecha_str = datetime.now().strftime("%d-%m-%Y")
    nombre_excel = f"MUESTREOS_MACRO_{fecha_str}.xlsx"
    
    output = BytesIO()
    workbook = xlsxwriter.Workbook(output)
    
    lotes_procesados = []
    errores = []
    
    for _, lote in muestreos_hoy.iterrows():
        try:
            datos = generar_datos_lote(lote)
            # Nombre de la hoja: CÓDIGO_DÍAS
            codigo = str(lote["Código"])
            muestreo_activador = lote.get("muestreo_activador", "")
            if pd.notna(muestreo_activador):
                dias_str = str(muestreo_activador).strip()
                dias_clean = re.sub(r'\s+', '_', dias_str)
                nombre_hoja_base = f"{codigo}_{dias_clean}"
            else:
                nombre_hoja_base = codigo
            nombre_hoja = re.sub(r'[\\/*?:\[\]]', '_', nombre_hoja_base)[:31]
            
            escribir_hoja(workbook, datos, nombre_hoja)
            lotes_procesados.append(codigo)
        except Exception as e:
            errores.append(f"{lote.get('Código')}: {str(e)}")
    
    workbook.close()
    output.seek(0)
    
    excel_data = base64.b64encode(output.read()).decode("utf-8")
    href = f"data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{excel_data}"
    
    # Mensaje de resultado
    resultado = html.Div([
        html.P(f"✅ Excel generado correctamente con {len(lotes_procesados)} hoja(s)."),
        html.P(f"Lotes incluidos: {', '.join(lotes_procesados)}"),
    ])
    if errores:
        resultado.children.append(html.P(f"❌ Errores en: {', '.join(errores)}", style={"color": "red"}))
    
    return href, nombre_excel, resultado

# =============================================================================
# CALLBACK PARA SUPERVIVENCIA (sin cambios, solo copiado)
# =============================================================================
@app.callback(
    [Output('output-alertas', 'children'),
     Output('output-data-upload', 'children'),
     Output('grafico-supervivencia', 'figure'),
     Output('grafico-talla-comercial', 'figure'),
     Output('grafico-ejes', 'figure'),
     Output('grafico-ocupacion', 'figure'),
     Output('grafico-altura', 'figure'),
     Output('grafico-porcentaje-col', 'figure')],
    [Input('upload-data', 'contents')],
    [State('upload-data', 'filename')]
)
def procesar_archivo(contents, filename):
    if contents is None:
        empty_fig = {}
        return html.Div(["Por favor, carga un archivo Excel."]), None, empty_fig, empty_fig, empty_fig, empty_fig, empty_fig, empty_fig

    content_type, content_string = contents.split(',')
    decoded = base64.b64decode(content_string)

    try:
        df_raw = pd.read_excel(BytesIO(decoded), header=None)
        
        # Buscar la fila donde aparece "Fila" (encabezado de la tabla)
        header_row_idx = None
        for i in range(len(df_raw)):
            if df_raw.iloc[i, 0] == 'Fila':
                header_row_idx = i
                break
        
        if header_row_idx is None:
            return html.Div(["No se encontró la fila de encabezado 'Fila' en el archivo."]), None, {}, {}, {}, {}, {}, {}
        
        df = pd.read_excel(BytesIO(decoded), header=header_row_idx)
        
        # Limpiar: eliminar filas donde 'Fila' no sea numérico
        df['Fila_temp'] = df['Fila'].astype(str).str.strip()
        mask_fila_valida = df['Fila_temp'].str.match(r'^\d+(\.\d+)?$', na=False)
        df = df[mask_fila_valida].copy()
        df.drop(columns=['Fila_temp'], inplace=True)
        
        if df.empty:
            return html.Div(["No se encontraron filas de datos numéricos en la tabla."]), None, {}, {}, {}, {}, {}, {}
        
        columnas_numericas = ['Máximo', 'Sobrevivencia', 'Talla Comercial', 'Ejes ≥ 2',
                              'Ocup sustrato ≥ 80%', 'Altura ≥ 12 cm']
        for col in columnas_numericas:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce')
            else:
                df[col] = 0
        
        if '% Col' in df.columns:
            df['% Col'] = pd.to_numeric(df['% Col'], errors='coerce')
            columnas_numericas.append('% Col')
        else:
            df['% Col'] = 0
        
        df[columnas_numericas] = df[columnas_numericas].fillna(0)
        df['Fila'] = pd.to_numeric(df['Fila'], errors='coerce').fillna(0).astype(int).astype(str)
        
        if 'Máximo' not in df.columns:
            return html.Div(["Columna 'Máximo' no encontrada."]), None, {}, {}, {}, {}, {}, {}
        
        total_maximo = df['Máximo'].sum()
        if total_maximo == 0:
            return html.Div(["El total de 'Máximo' es cero, no se puede calcular porcentajes."]), None, {}, {}, {}, {}, {}, {}
        
        total_sobrevivencia = df['Sobrevivencia'].sum()
        tasa_supervivencia = (total_sobrevivencia / total_maximo) * 100
        total_talla_comercial = df['Talla Comercial'].sum()
        tasa_talla_comercial = (total_talla_comercial / total_maximo) * 100
        total_ejes = df['Ejes ≥ 2'].sum()
        tasa_ejes = (total_ejes / total_maximo) * 100
        total_ocupacion = df['Ocup sustrato ≥ 80%'].sum()
        tasa_ocupacion = (total_ocupacion / total_maximo) * 100
        total_altura = df['Altura ≥ 12 cm'].sum()
        tasa_altura = (total_altura / total_maximo) * 100
        
        if '% Col' in df.columns and df['% Col'].sum() > 0:
            total_porcentaje_col = df['% Col'].sum()
            tasa_porcentaje_col = (total_porcentaje_col / total_maximo) * 100
        else:
            tasa_porcentaje_col = 0
        
        condiciones = (
            (df['Sobrevivencia'] > df['Máximo']) |
            (df['Talla Comercial'] > df['Máximo']) |
            (df['Ejes ≥ 2'] > df['Máximo']) |
            (df['Ocup sustrato ≥ 80%'] > df['Máximo']) |
            (df['Altura ≥ 12 cm'] > df['Máximo'])
        )
        if '% Col' in df.columns and '% Col' in df:
            condiciones = condiciones | (df['% Col'] > df['Máximo'])
        
        filas_alerta = df[condiciones]
        
        alerta = html.Div([
            html.H5("⚠️ Alarmas detectadas:", style={"color": "red"}),
            html.P(f"Se encontraron {len(filas_alerta)} filas con valores fuera de rango."),
            dash_table.DataTable(
                data=filas_alerta.to_dict('records'),
                columns=[{'name': i, 'id': i} for i in filas_alerta.columns],
                style_table={'overflowX': 'auto', 'maxWidth': '100%'},
                style_cell={'textAlign': 'center', 'padding': '5px', 'fontSize': '12px'},
                style_header={'backgroundColor': 'lightgrey', 'fontWeight': 'bold'},
                page_size=10
            )
        ]) if not filas_alerta.empty else html.Div([
            html.H5("✅ No se detectaron alarmas.", style={"color": "green"})
        ])
        
        tabla = dash_table.DataTable(
            data=df.to_dict('records'),
            columns=[{'name': i, 'id': i} for i in df.columns],
            style_table={'overflowX': 'auto', 'maxWidth': '100%'},
            style_cell={'textAlign': 'center', 'padding': '5px', 'fontSize': '12px'},
            style_header={'backgroundColor': 'lightgrey', 'fontWeight': 'bold'},
            page_size=10
        )
        
        filas_unicas = df['Fila'].tolist()
        
        def crear_grafico(col_y, titulo, color, label_y):
            if col_y not in df.columns:
                return px.bar(title=f"{titulo} - Columna no encontrada")
            fig = px.bar(
                df, x='Fila', y=col_y,
                title=titulo,
                labels={'Fila': 'Fila', col_y: label_y},
                color_discrete_sequence=[color]
            )
            fig.update_traces(text=df[col_y], textposition='outside')
            fig.update_layout(
                xaxis=dict(tickmode='array', tickvals=filas_unicas, ticktext=filas_unicas, tickangle=-45),
                xaxis_title="Fila", yaxis_title=label_y,
                plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
                font=dict(size=10), margin=dict(t=60, b=80, l=50, r=50),
                height=400
            )
            return fig
        
        fig_supervivencia = crear_grafico('Sobrevivencia', f'Supervivencia: {tasa_supervivencia:.2f}%', '#1f77b4', 'Plantas Vivas')
        fig_talla_comercial = crear_grafico('Talla Comercial', f'Talla Comercial: {tasa_talla_comercial:.2f}%', '#ff7f0e', 'Plantas en Talla Comercial')
        fig_ejes = crear_grafico('Ejes ≥ 2', f'Ejes ≥ 2: {tasa_ejes:.2f}%', '#2ca02c', 'Plantas con Ejes ≥ 2')
        fig_ocupacion = crear_grafico('Ocup sustrato ≥ 80%', f'Ocupación Sustrato ≥ 80%: {tasa_ocupacion:.2f}%', '#d62728', 'Plantas con Ocupación ≥ 80%')
        fig_altura = crear_grafico('Altura ≥ 12 cm', f'Altura ≥ 12 cm: {tasa_altura:.2f}%', '#9467bd', 'Plantas con Altura ≥ 12 cm')
        
        if '% Col' in df.columns and df['% Col'].sum() > 0:
            fig_porcentaje_col = crear_grafico('% Col', f'% Col: {tasa_porcentaje_col:.2f}%', '#8c564b', 'Plantas con % Col')
        else:
            fig_porcentaje_col = px.bar(title="% Col no disponible en el archivo")
        
        # Leer metadatos (fecha y lote) desde posiciones fijas del Excel original
        metadata_df = pd.read_excel(BytesIO(decoded), header=None)
        fecha_muestreo = metadata_df.iloc[5, 5] if metadata_df.shape[0] > 5 and metadata_df.shape[1] > 5 else "No disponible"
        lote = metadata_df.iloc[7, 2] if metadata_df.shape[0] > 7 and metadata_df.shape[1] > 2 else "No disponible"
        
        try:
            if isinstance(fecha_muestreo, str):
                fecha_muestreo = pd.to_datetime(fecha_muestreo, format="%d-%m-%Y", errors="raise")
            elif isinstance(fecha_muestreo, (int, float)):
                fecha_muestreo = pd.to_datetime("1899-12-30") + pd.to_timedelta(int(fecha_muestreo), unit="D")
            fecha_muestreo = fecha_muestreo.strftime('%d-%m-%Y')
        except Exception:
            fecha_muestreo = "Formato de fecha inválido"
        
        resumen = dbc.Container([
            dbc.Card(
                dbc.CardBody([
                    html.H5(f"Archivo cargado: {filename}", className="text-center text-primary mb-4"),
                    html.P(f"Lote maceta: {lote}", className="text-center mb-2"),
                    html.P(f"Fecha Muestreo: {fecha_muestreo}", className="text-center mb-2"),
                    html.P(f"N° macetas muestreo: {int(total_maximo):,}".replace(",", "."), className="text-center mb-2"),
                    html.P(f"% plantas vivas: {tasa_supervivencia:.2f}%".replace('.', ','), className="text-center mb-2"),
                    html.P(f"% plantas comerciales: {tasa_talla_comercial:.2f}%".replace('.', ','), className="text-center mb-2"),
                ]),
                className="shadow-sm bg-light p-4 mx-auto",
                style={"maxWidth": "500px"}
            ),
            html.Div([
                html.H5("Tabla de Datos", className="text-center text-primary mt-4"),
                tabla
            ], style={'overflowX': 'auto'})
        ])
        
        return alerta, resumen, fig_supervivencia, fig_talla_comercial, fig_ejes, fig_ocupacion, fig_altura, fig_porcentaje_col
        
    except Exception as e:
        empty_fig = {}
        return html.Div([f"Error al procesar el archivo: {str(e)}"]), None, empty_fig, empty_fig, empty_fig, empty_fig, empty_fig, empty_fig

# =============================================================================
# EJECUCIÓN
# =============================================================================
if __name__ == "__main__":
    app.run(host='0.0.0.0', port=8050, debug=True)s