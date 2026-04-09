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

# =============================================================================
# CONFIGURACIÓN DE RUTAS
# =============================================================================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.path.join(BASE_DIR, "output")
IMAGES_DIR = os.path.join(BASE_DIR, "images")
ARCHIVO_EXCEL = os.path.join(OUTPUT_DIR, "Muestreos_Activos.xlsx")
IMAGEN_RECORTADA = os.path.join(IMAGES_DIR, "imagen_recortada.jpg")

# =============================================================================
# CARGA DE DATOS
# =============================================================================
if not os.path.exists(ARCHIVO_EXCEL):
    raise FileNotFoundError(f"No se encontró {ARCHIVO_EXCEL}. Ejecuta primero el script principal.")

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
    patron = r'-C(\d+)'
    numeros = re.findall(patron, imc_str)
    if not numeros:
        return 0
    return sum(int(num) for num in numeros)

# =============================================================================
# APP DASH
# =============================================================================
app = dash.Dash(__name__, external_stylesheets=[dbc.themes.FLATLY])
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
        ], width=12)  # Ahora ocupa todo el ancho
    ])
], fluid=True)

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

    if codigo_hoy:
        df_origen = muestreos_hoy
    else:
        df_origen = muestreos_proximos

    try:
        lote = df_origen[df_origen["Código"] == codigo].iloc[0]

        # Extraer cantidad desde I-M-C
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

        # Bandejas y volumen
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

        seed = int(hashlib.sha256(str(codigo).encode()).hexdigest(), 16) % (10**6)
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
            return "Error: No se generaron datos de muestra.", "", ""

        tabla_df = pd.DataFrame(rows)

        # --- Generar archivo Excel ---
        output = BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            workbook = writer.book
            formato_negrita_bordes = workbook.add_format({'bold': True, 'border': 1, 'align': 'center', 'valign': 'vcenter'})
            formato_bordes = workbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter', 'font_size': 10})
            formato_centrado = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'font_size': 10, 'bold': True})
            formato_bordes_pequeno = workbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter', 'font_size': 9})

            worksheet = workbook.add_worksheet("Resumen_Muestreo")
            start_row = 4

            info_previa = {
                "ID": lote.get("ID", "N/A"),
                "Fecha Inicial": lote.get("Fecha", "").strftime('%d-%m-%Y') if pd.notnull(lote.get("Fecha")) else "N/A",
                "Especie": lote.get("Especie", "N/A"),
                "Variedad": lote.get("Variedad", "N/A"),
                "Muestreo": lote.get("muestreo_activador", "N/A"),
                "Fecha Muestreo": lote["fecha_activadora"].strftime('%d-%m-%Y') if pd.notnull(lote.get("fecha_activadora")) else "N/A",
                "Alveolos": cantidad,
                "Muestra": len(tabla_df)
            }
            for col_num, (key, value) in enumerate(info_previa.items()):
                worksheet.write(start_row, col_num, key, formato_negrita_bordes)
                worksheet.write(start_row + 1, col_num, value, formato_bordes)
            worksheet.set_row(start_row, 15)
            worksheet.set_row(start_row + 1, 12)
            start_row += 2

            worksheet.write(start_row, 0, "Bandeja", formato_negrita_bordes)
            worksheet.write(start_row, 1, "Vol. Sustrato (L)", formato_negrita_bordes)
            worksheet.write(start_row, 2, "Código", formato_negrita_bordes)
            worksheet.write(start_row + 1, 0, bandejas_val, formato_bordes)
            worksheet.write(start_row + 1, 1, f"{litros:.2f}", formato_bordes)
            worksheet.write(start_row + 1, 2, lote["Código"], formato_bordes)
            worksheet.set_row(start_row, 12)
            worksheet.set_row(start_row + 1, 12)
            start_row += 2

            worksheet.merge_range(f'D{start_row-1}:H{start_row-1}', 'INVERNADERO - MESÓN - CANTIDAD (I-M-C)', formato_negrita_bordes)
            worksheet.merge_range(f'D{start_row}:H{start_row}', str(imc_val), formato_bordes)
            worksheet.set_row(start_row - 1, 12)
            worksheet.set_row(start_row, 12)
            start_row += 2

            # Tabla de repeticiones
            worksheet.set_row(8, 5)
            header_row = 9
            filas_resumen = tabla_df.groupby("Fila").size().reset_index(name="Cantidad de Repeticiones")
            columnas_vacias = ["Sobrevivencia", "Ejes ≥ 2", "Ocup sustrato ≥ 80%", "Altura ≥ 12 cm", "Talla Comercial", "% Col"]
            todas_columnas = ["Fila", "Máximo"] + columnas_vacias
            for col_num, col_name in enumerate(todas_columnas):
                worksheet.write(header_row, col_num, col_name, formato_negrita_bordes)
            worksheet.set_row(header_row, 12)
            data_start_row = header_row + 1
            for idx, row in filas_resumen.iterrows():
                worksheet.write(data_start_row + idx, 0, row["Fila"], formato_bordes)
                worksheet.write(data_start_row + idx, 1, row["Cantidad de Repeticiones"], formato_bordes)
                for col_num in range(2, len(todas_columnas)):
                    worksheet.write(data_start_row + idx, col_num, "", formato_bordes)
                worksheet.set_row(data_start_row + idx, 11)
            last_data_row = data_start_row + len(filas_resumen) - 1

            blank_row_after_table = last_data_row + 1
            worksheet.set_row(blank_row_after_table, 5)
            responsable_row = blank_row_after_table + 1
            worksheet.merge_range(f'A{responsable_row+1}:D{responsable_row+1}', 'Responsable: _________________________________________________', formato_centrado)
            worksheet.merge_range(f'F{responsable_row+1}:H{responsable_row+1}', 'Fecha: ______ /______ /_________', formato_centrado)
            worksheet.write(f'E{responsable_row+1}', 'Firma: ___________', formato_centrado)
            worksheet.set_row(responsable_row, 12)

            blank_row_before_percent = responsable_row + 1
            worksheet.set_row(blank_row_before_percent, 5)
            percent_row = blank_row_before_percent + 1
            worksheet.merge_range(f'A{percent_row+1}:B{percent_row+1}', '% PLANTAS PLANTABLES', formato_bordes_pequeno)
            worksheet.write(f'C{percent_row+1}', '', formato_bordes_pequeno)
            worksheet.merge_range(f'F{percent_row+1}:G{percent_row+1}', '% TALLA COMERCIAL', formato_bordes_pequeno)
            worksheet.write(f'H{percent_row+1}', '', formato_bordes_pequeno)
            worksheet.set_row(percent_row, 10)

            # Encabezado con logo
            worksheet.merge_range('A1:B2', '', formato_negrita_bordes)
            cell_width = 30
            cell_height = 22
            image_scale_x = cell_width / 45
            image_scale_y = cell_height / 65
            if os.path.exists(IMAGEN_RECORTADA):
                worksheet.insert_image('A1', IMAGEN_RECORTADA, {
                    'x_scale': image_scale_x,
                    'y_scale': image_scale_y,
                    'x_offset': 12,
                    'y_offset': 5,
                    'positioning': 1
                })
            merge_format_titulo = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'font_name': 'Comic Sans MS', 'font_size': 11, 'text_wrap': True, 'border': 1})
            worksheet.merge_range('C1:E2', 'Sociedad de Investigación, Desarrollo y Servicios de Biotecnología Aplicada Ltda.', merge_format_titulo)
            cell_format = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'font_name': 'Arial', 'font_size': 10, 'text_wrap': True, 'border': 1})
            worksheet.write('F1', 'RAC-XXX', cell_format)
            worksheet.write('G1', 'POE XXX', cell_format)
            worksheet.write('F2', 'Edición 00', cell_format)
            worksheet.write('G2', 'Pág. 1 de 1', cell_format)
            cell_format_11 = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'font_name': 'Arial', 'font_size': 11, 'text_wrap': True, 'border': 1})
            worksheet.merge_range('A3:E3', 'REGISTRO PARA EL CONTROL DE LA TALLA COMERCIAL EN MACRO', cell_format_11)
            worksheet.write('F3', 'Vigente: 01/ 01/2025', cell_format)
            worksheet.write('G3', 'Folio:', cell_format)
            for col, width in [('A', 8), ('B', 13), ('C', 23), ('D', 12), ('E', 18), ('F', 18), ('G', 13), ('H', 9)]:
                worksheet.set_column(f'{col}:{col}', width)
            worksheet.set_row(3, 5)

        output.seek(0)
        excel_data = base64.b64encode(output.read()).decode("utf-8")
        href = f"data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{excel_data}"
        download_name = f'MACRO_{lote["Código"]}_{lote["muestreo_activador"]}_{lote["fecha_activadora"].strftime("%d-%m-%Y")}.xlsx'

        # --- Resultados en texto ---
        resultados = [
            html.P(f"ID del Lote: {lote.get('ID', 'N/A')}"),
            html.P(f"Fecha Inicial: {lote.get('Fecha', '').strftime('%d-%m-%Y') if pd.notnull(lote.get('Fecha')) else 'N/A'}"),
            html.P(f"Especie: {lote.get('Especie', 'N/A')}"),
            html.P(f"Variedad: {lote.get('Variedad', 'N/A')}"),
            html.P(f"Código: {lote.get('Código', 'N/A')}"),
            html.P(f"Estado: {lote.get('Estado', 'N/A')}"),
            html.P(f"Reagrupado: {lote.get('Reagrupado', 'N/A')}"),
            html.P(f"LMC Dominante: {lote.get('LMC Dominante', 'N/A')}"),
            html.P(f"Muestreo: {lote.get('muestreo_activador', 'N/A')}"),
            html.P(f"Fecha Muestreo: {lote['fecha_activadora'].strftime('%d-%m-%Y') if pd.notnull(lote.get('fecha_activadora')) else 'N/A'}"),
            html.P(f"Alveolos: {cantidad}"),
            html.P(f"Muestra: {muestra_tamano}"),
            html.P(f"Bandeja: {bandejas_val}"),
            html.P(f"Vol. Sustrato: {litros:.2f} L"),
            html.P(f"I-M-C: {imc_val}")
        ]

        return resultados, href, download_name

    except Exception as e:
        import traceback
        traceback.print_exc()
        return f"Error: {str(e)}", "", ""

if __name__ == "__main__":
    app.run(host='127.0.0.1', port=8050, debug=True)