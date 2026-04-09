import os
import re
import base64
import pandas as pd
import numpy as np
from io import BytesIO
import dash
from dash import dcc, html, Input, Output, State, dash_table
import dash_bootstrap_components as dbc
import plotly.express as px
import warnings

warnings.filterwarnings("ignore", category=UserWarning)

# =============================================================================
# APP DASH
# =============================================================================
app = dash.Dash(__name__, external_stylesheets=[dbc.themes.FLATLY])
server = app.server  # 🔴 CLAVE PARA RENDER

# 🔴 Aumentar límite de carga (50MB)
app.server.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024

app.title = "Análisis de Supervivencia - MACRO"

app.layout = dbc.Container([
    html.H1("Análisis Automático de Supervivencia con Alarmas", className="text-center my-4"),
    
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
        dbc.Col(dcc.Graph(id="grafico-supervivencia"), width=4),
        dbc.Col(dcc.Graph(id="grafico-talla-comercial"), width=4),
        dbc.Col(dcc.Graph(id="grafico-ejes"), width=4),
    ]),

    dbc.Row([
        dbc.Col(dcc.Graph(id="grafico-ocupacion"), width=4),
        dbc.Col(dcc.Graph(id="grafico-altura"), width=4),
        dbc.Col(dcc.Graph(id="grafico-porcentaje-col"), width=4),
    ]),
], fluid=True)

# =============================================================================
# CALLBACK
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
        # 🔥 usar un solo buffer
        excel_file = BytesIO(decoded)

        df_raw = pd.read_excel(excel_file, header=None)
        excel_file.seek(0)

        # Buscar encabezado
        header_row_idx = None
        for i in range(len(df_raw)):
            if df_raw.iloc[i, 0] == 'Fila':
                header_row_idx = i
                break

        if header_row_idx is None:
            return html.Div(["No se encontró la fila 'Fila'."]), None, {}, {}, {}, {}, {}, {}

        df = pd.read_excel(excel_file, header=header_row_idx)
        excel_file.seek(0)

        # Limpiar filas
        df['Fila_temp'] = df['Fila'].astype(str).str.strip()
        df = df[df['Fila_temp'].str.match(r'^\d+(\.\d+)?$', na=False)].copy()
        df.drop(columns=['Fila_temp'], inplace=True)

        if df.empty:
            return html.Div(["No hay datos válidos."]), None, {}, {}, {}, {}, {}, {}

        # Columnas
        columnas_numericas = [
            'Máximo', 'Sobrevivencia', 'Talla Comercial',
            'Ejes ≥ 2', 'Ocup sustrato ≥ 80%', 'Altura ≥ 12 cm'
        ]

        for col in columnas_numericas:
            df[col] = pd.to_numeric(df.get(col, 0), errors='coerce')

        # % Col
        if '% Col' in df.columns:
            df['% Col'] = pd.to_numeric(df['% Col'], errors='coerce')
        else:
            df['% Col'] = 0

        df = df.fillna(0)

        df['Fila'] = pd.to_numeric(df['Fila'], errors='coerce').fillna(0).astype(int).astype(str)

        total_maximo = df['Máximo'].sum()
        if total_maximo == 0:
            return html.Div(["Máximo = 0."]), None, {}, {}, {}, {}, {}, {}

        # Tasas
        tasa = lambda col: (df[col].sum() / total_maximo) * 100

        tasa_supervivencia = tasa('Sobrevivencia')
        tasa_talla = tasa('Talla Comercial')
        tasa_ejes = tasa('Ejes ≥ 2')
        tasa_ocup = tasa('Ocup sustrato ≥ 80%')
        tasa_altura = tasa('Altura ≥ 12 cm')
        tasa_col = tasa('% Col') if df['% Col'].sum() > 0 else 0

        # Alarmas
        condiciones = (
            (df['Sobrevivencia'] > df['Máximo']) |
            (df['Talla Comercial'] > df['Máximo']) |
            (df['Ejes ≥ 2'] > df['Máximo']) |
            (df['Ocup sustrato ≥ 80%'] > df['Máximo']) |
            (df['Altura ≥ 12 cm'] > df['Máximo']) |
            (df['% Col'] > df['Máximo'])
        )

        filas_alerta = df[condiciones]

        alerta = html.Div([
            html.H5("⚠️ Alarmas detectadas", style={"color": "red"}),
            html.P(f"{len(filas_alerta)} filas con error."),
            dash_table.DataTable(data=filas_alerta.to_dict('records'),
                                 columns=[{'name': i, 'id': i} for i in df.columns])
        ]) if not filas_alerta.empty else html.H5("✅ Sin alarmas", style={"color": "green"})

        # Gráficos
        def graf(col, titulo):
            return px.bar(df, x='Fila', y=col, title=f"{titulo}: {tasa(col):.2f}%")

        figs = [
            graf('Sobrevivencia', 'Supervivencia'),
            graf('Talla Comercial', 'Talla Comercial'),
            graf('Ejes ≥ 2', 'Ejes'),
            graf('Ocup sustrato ≥ 80%', 'Ocupación'),
            graf('Altura ≥ 12 cm', 'Altura'),
            graf('% Col', '% Col')
        ]

        resumen = html.Div([
            html.H5(f"Archivo: {filename}"),
            html.P(f"Total: {int(total_maximo)}"),
            html.P(f"Supervivencia: {tasa_supervivencia:.2f}%")
        ])

        return alerta, resumen, *figs

    except Exception as e:
        return html.Div([str(e)]), None, {}, {}, {}, {}, {}, {}

# =============================================================================
# RUN LOCAL
# =============================================================================
if __name__ == "__main__":
    app.run(debug=True)