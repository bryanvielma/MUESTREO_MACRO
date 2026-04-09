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

warnings.filterwarnings("ignore")

# =============================================================================
# APP DASH (SOLO ANÁLISIS DE SUPERVIVENCIA)
# =============================================================================
app = dash.Dash(__name__, external_stylesheets=[dbc.themes.FLATLY])
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
# CALLBACK PARA PROCESAR EL ARCHIVO EXCEL SUBIDO
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
        # Leer todo el archivo sin asumir header
        df_raw = pd.read_excel(BytesIO(decoded), header=None)
        
        # Buscar la fila donde aparece "Fila" (encabezado de la tabla)
        header_row_idx = None
        for i in range(len(df_raw)):
            if df_raw.iloc[i, 0] == 'Fila':
                header_row_idx = i
                break
        
        if header_row_idx is None:
            return html.Div(["No se encontró la fila de encabezado 'Fila' en el archivo."]), None, {}, {}, {}, {}, {}, {}
        
        # Leer los datos a partir de la fila siguiente al encabezado
        df = pd.read_excel(BytesIO(decoded), header=header_row_idx)
        
        # Limpiar: eliminar filas donde 'Fila' no sea numérico (texto como "Responsable", etc.)
        df['Fila_temp'] = df['Fila'].astype(str).str.strip()
        mask_fila_valida = df['Fila_temp'].str.match(r'^\d+(\.\d+)?$', na=False)
        df = df[mask_fila_valida].copy()
        df.drop(columns=['Fila_temp'], inplace=True)
        
        # Si después del filtro no hay filas, error
        if df.empty:
            return html.Div(["No se encontraron filas de datos numéricos en la tabla."]), None, {}, {}, {}, {}, {}, {}
        
        # Convertir columnas numéricas
        columnas_numericas = ['Máximo', 'Sobrevivencia', 'Talla Comercial', 'Ejes ≥ 2',
                              'Ocup sustrato ≥ 80%', 'Altura ≥ 12 cm']
        for col in columnas_numericas:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce')
            else:
                df[col] = 0
        
        # % Col puede no existir
        if '% Col' in df.columns:
            df['% Col'] = pd.to_numeric(df['% Col'], errors='coerce')
            columnas_numericas.append('% Col')
        else:
            df['% Col'] = 0
        
        df[columnas_numericas] = df[columnas_numericas].fillna(0)
        
        # Asegurar que 'Fila' sea entero y luego string para el eje X
        df['Fila'] = pd.to_numeric(df['Fila'], errors='coerce').fillna(0).astype(int).astype(str)
        
        # Verificar que existe 'Máximo' y total > 0
        if 'Máximo' not in df.columns:
            return html.Div(["Columna 'Máximo' no encontrada."]), None, {}, {}, {}, {}, {}, {}
        
        total_maximo = df['Máximo'].sum()
        if total_maximo == 0:
            return html.Div(["El total de 'Máximo' es cero, no se puede calcular porcentajes."]), None, {}, {}, {}, {}, {}, {}
        
        # Cálculo de tasas
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
        
        # Alarmas
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
# EJECUCIÓN DE LA APP
# =============================================================================
if __name__ == "__main__":
    app.run(host='127.0.0.1', port=8050, debug=True)