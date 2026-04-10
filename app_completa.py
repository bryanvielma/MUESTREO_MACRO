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
    raise FileNotFoundError(f"No se encontró {ARCHIVO_EXCEL}.")

muestreos_hoy_raw      = pd.read_excel(ARCHIVO_EXCEL, sheet_name="Hoy")
muestreos_proximos_raw = pd.read_excel(ARCHIVO_EXCEL, sheet_name="Proximos")

def es_lote_peru(row):
    imc = row.get("I-M-C", "")
    return isinstance(imc, str) and "MN" in imc.upper()

hoy_date = datetime.now().date()
if 'fecha_activadora' in muestreos_hoy_raw.columns:
    muestreos_hoy_raw['fecha_activadora'] = pd.to_datetime(
        muestreos_hoy_raw['fecha_activadora'], errors='coerce')
    muestreos_hoy_raw = muestreos_hoy_raw[
        muestreos_hoy_raw['fecha_activadora'].dt.date == hoy_date].copy()

ids_excluidos = sorted(set(
    muestreos_hoy_raw[muestreos_hoy_raw.apply(es_lote_peru, axis=1)]["ID"].tolist()
))

def filtrar_sin_mn(df):
    if 'I-M-C' not in df.columns:
        return df
    return df[df.apply(lambda r: not es_lote_peru(r), axis=1)].copy()

muestreos_hoy      = filtrar_sin_mn(muestreos_hoy_raw)
muestreos_proximos = filtrar_sin_mn(muestreos_proximos_raw)

for df in [muestreos_hoy, muestreos_proximos]:
    for col in ['Macetas actuales', 'Alveolos', 'Bandeja']:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce')
    if 'fecha_activadora' in df.columns:
        df['fecha_activadora'] = pd.to_datetime(df['fecha_activadora'], errors='coerce')

# =============================================================================
# FUNCIONES AUXILIARES
# =============================================================================
def extraer_cantidad_desde_imc(imc_str):
    if not isinstance(imc_str, str):
        return 0
    numeros = re.findall(r'-C(\d+)', imc_str)
    return sum(int(n) for n in numeros) if numeros else 0

def generar_datos_lote(lote):
    imc_raw = lote.get("I-M-C", "")
    if pd.isna(imc_raw):
        imc_raw = ""
    cantidad = extraer_cantidad_desde_imc(str(imc_raw))
    if cantidad == 0:
        cantidad = lote.get("Macetas actuales", 0)
        cantidad = int(cantidad) if not pd.isna(cantidad) else 0
    else:
        cantidad = int(cantidad)
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
        rangos = [(8,2),(15,3),(25,5),(50,8),(90,13),(150,20),(280,40),
                  (500,60),(1200,80),(3200,140),(10000,200),(35000,320),
                  (150000,500),(500000,800)]
        for limite, muestra in rangos:
            if c <= limite:
                return muestra
        return 1260

    muestra_tamano = calcular_tamano(cantidad)
    rows_count     = max(cantidad // hileras, 1)
    full_rows      = muestra_tamano // hileras
    remainder      = muestra_tamano % hileras
    rows_needed    = full_rows + (1 if remainder > 0 else 0)

    seed = int(hashlib.sha256(str(lote["Código"]).encode()).hexdigest(), 16) % (10**6)
    np.random.seed(seed)
    spacing      = max(rows_count // rows_needed, 1)
    primera_fila = np.random.randint(1, spacing + 1)
    chosen_rows  = sorted(set(
        min(primera_fila + i * spacing, rows_count) for i in range(rows_needed)
    ))

    rows = []
    muestra_num = 1
    for i, r_idx in enumerate(chosen_rows):
        start_plant = (r_idx - 1) * hileras + 1
        if (i == len(chosen_rows) - 1) and remainder > 0:
            for off_p in range(remainder):
                rows.append({"Muestra": muestra_num,
                              "Número Planta": start_plant + off_p,
                              "Fila": r_idx})
                muestra_num += 1
        else:
            for off_p in range(hileras):
                rows.append({"Muestra": muestra_num,
                              "Número Planta": start_plant + off_p,
                              "Fila": r_idx})
                muestra_num += 1

    if not rows:
        raise ValueError("No se generaron datos de muestra.")

    return {
        "tabla_df": pd.DataFrame(rows), "lote": lote,
        "cantidad": cantidad, "imc_val": imc_raw,
        "bandejas_val": bandejas_val, "litros": litros,
        "muestra_tamano": muestra_tamano
    }

def escribir_hoja(workbook, datos, nombre_hoja):
    worksheet    = workbook.add_worksheet(nombre_hoja[:31])
    fmt_bold     = workbook.add_format({'bold':True,'border':1,'align':'center','valign':'vcenter'})
    fmt_norm     = workbook.add_format({'border':1,'align':'center','valign':'vcenter','font_size':10})
    fmt_center   = workbook.add_format({'align':'center','valign':'vcenter','font_size':10,'bold':True})
    fmt_small    = workbook.add_format({'border':1,'align':'center','valign':'vcenter','font_size':9})
    start_row    = 4
    lote         = datos["lote"]
    tabla_df     = datos["tabla_df"]
    cantidad     = datos["cantidad"]
    imc_val      = datos["imc_val"]
    bandejas_val = datos["bandejas_val"]
    litros       = datos["litros"]

    info = {
        "ID": lote.get("ID","N/A"),
        "Fecha Inicial": lote.get("Fecha","").strftime('%d-%m-%Y') if pd.notnull(lote.get("Fecha")) else "N/A",
        "Especie": lote.get("Especie","N/A"),
        "Variedad": lote.get("Variedad","N/A"),
        "Muestreo": lote.get("muestreo_activador","N/A"),
        "Fecha Muestreo": lote["fecha_activadora"].strftime('%d-%m-%Y') if pd.notnull(lote.get("fecha_activadora")) else "N/A",
        "Alveolos": cantidad,
        "Muestra": len(tabla_df)
    }
    for col, (key, val) in enumerate(info.items()):
        worksheet.write(start_row, col, key, fmt_bold)
        worksheet.write(start_row+1, col, val, fmt_norm)
    worksheet.set_row(start_row, 15); worksheet.set_row(start_row+1, 12)
    start_row += 2

    worksheet.write(start_row,0,"Bandeja",fmt_bold)
    worksheet.write(start_row,1,"Vol. Sustrato (L)",fmt_bold)
    worksheet.write(start_row,2,"Código",fmt_bold)
    worksheet.write(start_row+1,0,bandejas_val,fmt_norm)
    worksheet.write(start_row+1,1,f"{litros:.2f}",fmt_norm)
    worksheet.write(start_row+1,2,lote["Código"],fmt_norm)
    worksheet.set_row(start_row,12); worksheet.set_row(start_row+1,12)
    start_row += 2

    worksheet.merge_range(f'D{start_row-1}:H{start_row-1}',
                          'INVERNADERO - MESÓN - CANTIDAD (I-M-C)', fmt_bold)
    worksheet.merge_range(f'D{start_row}:H{start_row}', str(imc_val), fmt_norm)
    worksheet.set_row(start_row-1,12); worksheet.set_row(start_row,12)
    start_row += 2

    worksheet.set_row(8,5)
    header_row    = 9
    filas_resumen = tabla_df.groupby("Fila").size().reset_index(name="Cantidad de Repeticiones")
    cols_vacias   = ["Sobrevivencia","Ejes ≥ 2","Ocup sustrato ≥ 80%",
                     "Altura ≥ 12 cm","Talla Comercial","% Col"]
    all_cols = ["Fila","Máximo"] + cols_vacias
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

    blank = last_data + 1; worksheet.set_row(blank, 5)
    resp_row = blank + 1
    worksheet.merge_range(f'A{resp_row+1}:D{resp_row+1}',
                          'Responsable: _________________________________________________', fmt_center)
    worksheet.merge_range(f'F{resp_row+1}:H{resp_row+1}',
                          'Fecha: ______ /______ /_________', fmt_center)
    worksheet.write(f'E{resp_row+1}','Firma: ___________', fmt_center)
    worksheet.set_row(resp_row, 12)

    blank2 = resp_row + 1; worksheet.set_row(blank2, 5)
    pct_row = blank2 + 1
    worksheet.merge_range(f'A{pct_row+1}:B{pct_row+1}','% PLANTAS PLANTABLES', fmt_small)
    worksheet.write(f'C{pct_row+1}','', fmt_small)
    worksheet.merge_range(f'F{pct_row+1}:G{pct_row+1}','% TALLA COMERCIAL', fmt_small)
    worksheet.write(f'H{pct_row+1}','', fmt_small)
    worksheet.set_row(pct_row, 10)

    worksheet.merge_range('A1:B2','', fmt_bold)
    if os.path.exists(IMAGEN_RECORTADA):
        worksheet.insert_image('A1', IMAGEN_RECORTADA,
            {'x_scale':30/45,'y_scale':22/65,'x_offset':12,'y_offset':5,'positioning':1})
    titulo = workbook.add_format({'align':'center','valign':'vcenter',
                                   'font_name':'Comic Sans MS','font_size':11,
                                   'text_wrap':True,'border':1})
    worksheet.merge_range('C1:E2',
        'Sociedad de Investigación, Desarrollo y Servicios de Biotecnología Aplicada Ltda.',titulo)
    cell_fmt = workbook.add_format({'align':'center','valign':'vcenter',
                                     'font_name':'Arial','font_size':10,'text_wrap':True,'border':1})
    worksheet.write('F1','RAC-XXX',cell_fmt); worksheet.write('G1','POE XXX',cell_fmt)
    worksheet.write('F2','Edición 00',cell_fmt); worksheet.write('G2','Pág. 1 de 1',cell_fmt)
    titulo2 = workbook.add_format({'align':'center','valign':'vcenter',
                                    'font_name':'Arial','font_size':11,'text_wrap':True,'border':1})
    worksheet.merge_range('A3:E3',
        'REGISTRO PARA EL CONTROL DE LA TALLA COMERCIAL EN MACRO', titulo2)
    worksheet.write('F3','Vigente: 01/ 01/2025',cell_fmt)
    worksheet.write('G3','Folio:',cell_fmt)
    for col, w in [('A',8),('B',13),('C',23),('D',12),('E',18),('F',18),('G',13),('H',9)]:
        worksheet.set_column(f'{col}:{col}', w)
    worksheet.set_row(3, 5)

# =============================================================================
# BIOTECH LABORATORY CSS - Modern, clean, no cream colors
# =============================================================================
CUSTOM_CSS = """
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&family=JetBrains+Mono:wght@400;500&display=swap');

:root {
  --bg-page:         #F4F7FB;
  --surface:         #FFFFFF;
  --primary-dark:    #0B3B2F;
  --primary:         #1A6D4C;
  --primary-light:   #4A9E7A;
  --primary-bglight: #E8F3EF;
  --accent-blue:     #2C7DA0;
  --accent-blue-bg:  #EAF4F8;
  --accent-teal:     #20B2AA;
  --gray-50:         #F8FAFC;
  --gray-100:        #F1F5F9;
  --gray-200:        #E2E8F0;
  --gray-300:        #CBD5E1;
  --gray-600:        #475569;
  --gray-700:        #334155;
  --gray-800:        #1E293B;
  --success:         #2E7D32;
  --warning:         #ED6C02;
  --error:           #D32F2F;
  --info:            #0288D1;
  --border-light:    #E2E8F0;
  --shadow-sm:       0 1px 2px rgba(0,0,0,0.04), 0 1px 2px rgba(0,0,0,0.02);
  --shadow-md:       0 4px 6px -1px rgba(0,0,0,0.05), 0 2px 4px -1px rgba(0,0,0,0.03);
  --shadow-lg:       0 10px 15px -3px rgba(0,0,0,0.05), 0 4px 6px -2px rgba(0,0,0,0.025);
  --radius-sm:       6px;
  --radius-md:       12px;
  --radius-lg:       18px;
}

* {
  box-sizing: border-box;
}

html, body {
  background: var(--bg-page) !important;
  color: var(--gray-700) !important;
  font-family: 'Inter', sans-serif !important;
  min-height: 100vh;
}

body::before {
  display: none;
}

.container-fluid {
  position: relative;
  z-index: 1;
}

/* HEADER */
.app-header {
  padding: 32px 0 0;
  display: flex;
  align-items: center;
  gap: 18px;
}

.header-icon {
  width: 50px;
  height: 50px;
  background: linear-gradient(135deg, var(--primary-dark), var(--primary));
  border-radius: var(--radius-md);
  display: flex;
  align-items: center;
  justify-content: center;
  font-size: 1.6rem;
  flex-shrink: 0;
  box-shadow: var(--shadow-md);
}

.header-text h1 {
  font-family: 'Inter', sans-serif !important;
  font-weight: 600 !important;
  font-size: 1.7rem !important;
  color: var(--gray-800) !important;
  letter-spacing: -0.01em;
  margin: 0 !important;
}

.header-text .tagline {
  font-size: 0.7rem;
  font-weight: 500;
  color: var(--gray-600);
  letter-spacing: 0.2em;
  text-transform: uppercase;
  margin-top: 5px;
}

.header-divider {
  margin-top: 18px;
  height: 2px;
  border-radius: 2px;
  background: linear-gradient(90deg, var(--primary) 0%, var(--accent-blue) 70%, transparent 100%);
  opacity: 0.6;
}

/* TABS */
.dash-tabs {
  border-bottom: 2px solid var(--border-light) !important;
  background: transparent !important;
  margin-top: 22px !important;
}

.dash-tab {
  font-family: 'Inter', sans-serif !important;
  font-weight: 500 !important;
  font-size: 0.85rem !important;
  color: var(--gray-600) !important;
  background: transparent !important;
  border: none !important;
  border-bottom: 2px solid transparent !important;
  padding: 10px 24px !important;
  margin-bottom: -2px !important;
  transition: all 0.2s ease !important;
}

.dash-tab:hover {
  color: var(--primary) !important;
  border-bottom-color: var(--primary-light) !important;
  background: rgba(26,109,76,0.04) !important;
}

.dash-tab--selected {
  color: var(--primary) !important;
  font-weight: 600 !important;
  border-bottom: 2px solid var(--primary) !important;
  background: transparent !important;
}

/* CARDS */
.card {
  background: var(--surface) !important;
  border: 1px solid var(--border-light) !important;
  border-radius: var(--radius-md) !important;
  box-shadow: var(--shadow-sm) !important;
  transition: box-shadow 0.2s ease, transform 0.1s ease !important;
}

.card:hover {
  box-shadow: var(--shadow-md) !important;
}

.card-header {
  background: var(--surface) !important;
  border-bottom: 1px solid var(--border-light) !important;
  padding: 14px 20px !important;
  font-family: 'Inter', sans-serif !important;
  font-weight: 600 !important;
  font-size: 0.75rem !important;
  letter-spacing: 0.05em;
  text-transform: uppercase !important;
  color: var(--primary) !important;
}

.card-body {
  background: var(--surface) !important;
  color: var(--gray-700) !important;
}

.border-forest { border-left: 3px solid var(--primary) !important; }
.border-amber  { border-left: 3px solid var(--accent-teal) !important; }
.border-sage   { border-left: 3px solid var(--primary-light) !important; }
.border-terra  { border-left: 3px solid var(--accent-blue) !important; }

/* KPI BAR */
.kpi-bar {
  display: flex;
  background: var(--surface);
  border: 1px solid var(--border-light);
  border-radius: var(--radius-md);
  box-shadow: var(--shadow-sm);
  margin-bottom: 20px;
  overflow: hidden;
}

.kpi-cell {
  flex: 1;
  padding: 16px 12px;
  border-right: 1px solid var(--border-light);
  text-align: center;
}

.kpi-cell:last-child {
  border-right: none;
}

.kpi-val {
  font-family: 'Inter', sans-serif;
  font-weight: 700;
  font-size: 1.6rem;
  color: var(--gray-800);
  line-height: 1.2;
}

.kpi-val.amber { color: var(--accent-teal); }
.kpi-val.sage  { color: var(--primary-light); }
.kpi-val.sky   { color: var(--accent-blue); }
.kpi-val.sm    { font-size: 1rem; }

.kpi-lbl {
  font-size: 0.6rem;
  font-weight: 600;
  letter-spacing: 0.1em;
  text-transform: uppercase;
  color: var(--gray-600);
  margin-top: 6px;
}

/* LOTE LIST */
.lote-list {
  list-style: none;
  padding: 0;
  margin: 0;
}

.lote-list li {
  display: flex;
  align-items: center;
  gap: 12px;
  padding: 8px 12px;
  border-radius: var(--radius-sm);
  font-size: 0.85rem;
  transition: background 0.15s;
}

.lote-list li:hover {
  background: var(--primary-bglight);
}

.lote-badge {
  display: inline-block;
  background: var(--primary-bglight);
  color: var(--primary-dark);
  font-size: 0.65rem;
  font-weight: 600;
  letter-spacing: 0.05em;
  padding: 2px 8px;
  border-radius: 30px;
  font-family: 'JetBrains Mono', monospace;
  flex-shrink: 0;
}

.lote-code {
  font-family: 'JetBrains Mono', monospace;
  font-size: 0.8rem;
  font-weight: 500;
  color: var(--gray-800);
}

.lote-var {
  font-size: 0.75rem;
  color: var(--gray-600);
  font-style: italic;
}

/* BUTTONS */
.btn-primary {
  background: var(--primary) !important;
  border: none !important;
  border-radius: var(--radius-sm) !important;
  color: white !important;
  font-family: 'Inter', sans-serif !important;
  font-weight: 600 !important;
  font-size: 0.8rem !important;
  letter-spacing: 0.03em !important;
  padding: 10px 22px !important;
  box-shadow: 0 1px 2px rgba(0,0,0,0.05) !important;
  transition: background 0.2s, box-shadow 0.2s !important;
}

.btn-primary:hover {
  background: var(--primary-dark) !important;
  box-shadow: 0 2px 4px rgba(0,0,0,0.1) !important;
}

.btn-success {
  background: transparent !important;
  border: 1.5px solid var(--primary) !important;
  border-radius: var(--radius-sm) !important;
  color: var(--primary) !important;
  font-family: 'Inter', sans-serif !important;
  font-weight: 600 !important;
  font-size: 0.8rem !important;
  letter-spacing: 0.03em !important;
  padding: 9px 22px !important;
  transition: all 0.2s !important;
}

.btn-success:hover {
  background: var(--primary) !important;
  color: white !important;
  box-shadow: 0 2px 8px rgba(26,109,76,0.2) !important;
}

/* ALERTS */
.alert {
  border-radius: var(--radius-sm) !important;
  font-size: 0.8rem !important;
  font-family: 'Inter', sans-serif !important;
  border: none !important;
  padding: 12px 16px !important;
}

.alert-warning {
  background: #FFF8E7 !important;
  color: #B45309 !important;
  border-left: 3px solid var(--warning) !important;
}

.alert-info {
  background: var(--accent-blue-bg) !important;
  color: var(--accent-blue) !important;
  border-left: 3px solid var(--accent-blue) !important;
}

.alert-success {
  background: var(--primary-bglight) !important;
  color: var(--primary-dark) !important;
  border-left: 3px solid var(--primary) !important;
}

.alert-danger {
  background: #FEF2F2 !important;
  color: var(--error) !important;
  border-left: 3px solid var(--error) !important;
}

/* UPLOAD ZONE */
.upload-zone {
  border: 2px dashed var(--gray-300) !important;
  border-radius: var(--radius-md) !important;
  background: var(--gray-50) !important;
  color: var(--gray-600) !important;
  font-size: 0.85rem !important;
  transition: all 0.2s ease !important;
  cursor: pointer;
}

.upload-zone:hover {
  border-color: var(--primary) !important;
  background: var(--primary-bglight) !important;
  color: var(--primary) !important;
}

/* DROPDOWN */
.Select-control {
  background: var(--surface) !important;
  border: 1px solid var(--gray-300) !important;
  border-radius: var(--radius-sm) !important;
}

.Select-menu-outer {
  background: var(--surface) !important;
  border-radius: var(--radius-sm) !important;
  box-shadow: var(--shadow-md) !important;
}

.Select-option {
  background: var(--surface) !important;
  color: var(--gray-700) !important;
}

.Select-option:hover,
.VirtualizedSelectFocusedOption {
  background: var(--primary-bglight) !important;
  color: var(--primary-dark) !important;
}

.Select-value-label {
  color: var(--gray-800) !important;
  font-weight: 500 !important;
}

.Select-placeholder {
  color: var(--gray-600) !important;
}

/* DATA TABLE */
.dash-header {
  background: var(--primary-bglight) !important;
  color: var(--primary-dark) !important;
  font-family: 'Inter', sans-serif !important;
  font-size: 0.7rem !important;
  font-weight: 600 !important;
  letter-spacing: 0.05em !important;
  text-transform: uppercase !important;
}

.dash-cell {
  background: var(--surface) !important;
  color: var(--gray-700) !important;
  font-family: 'JetBrains Mono', monospace !important;
  font-size: 0.75rem !important;
  border-color: var(--border-light) !important;
}

.dash-cell:hover {
  background: var(--gray-100) !important;
}

/* SECTION LABEL */
.sec-lbl {
  font-size: 0.65rem;
  font-weight: 600;
  letter-spacing: 0.15em;
  text-transform: uppercase;
  color: var(--gray-600);
  margin-bottom: 12px;
  display: flex;
  align-items: center;
  gap: 10px;
}

.sec-lbl::before {
  content: '';
  width: 14px;
  height: 2px;
  background: var(--primary-light);
  border-radius: 2px;
  flex-shrink: 0;
}

.sec-lbl::after {
  content: '';
  flex: 1;
  height: 1px;
  background: var(--border-light);
}

/* LOG BOX */
.log-box {
  background: var(--gray-50);
  border: 1px solid var(--border-light);
  border-radius: var(--radius-sm);
  padding: 14px 18px;
  font-family: 'JetBrains Mono', monospace;
  font-size: 0.75rem;
  color: var(--gray-600);
  min-height: 62px;
}

.log-ok {
  color: var(--primary);
}

.log-warn {
  color: var(--warning);
}

.log-error {
  color: var(--error);
}

/* STATUS DOT */
.dot {
  display: inline-block;
  width: 8px;
  height: 8px;
  border-radius: 50%;
  background: var(--primary);
  margin-right: 8px;
  vertical-align: middle;
  animation: pulse 2s ease-in-out infinite;
}

@keyframes pulse {
  0%, 100% { opacity: 1; transform: scale(1); }
  50% { opacity: 0.5; transform: scale(0.95); }
}

/* CHART WRAPPER */
.chart-wrap {
  background: var(--surface);
  border: 1px solid var(--border-light);
  border-radius: var(--radius-md);
  box-shadow: var(--shadow-sm);
  overflow: hidden;
  transition: box-shadow 0.2s;
}

.chart-wrap:hover {
  box-shadow: var(--shadow-md);
}

/* SCROLLBAR */
::-webkit-scrollbar {
  width: 5px;
  height: 5px;
}

::-webkit-scrollbar-track {
  background: var(--gray-100);
}

::-webkit-scrollbar-thumb {
  background: var(--gray-300);
  border-radius: 4px;
}

::-webkit-scrollbar-thumb:hover {
  background: var(--gray-400);
}
"""

# ── PLOTLY TEMPLATE ──────────────────────────────────────────────────────────
PLOT_LAYOUT = dict(
    plot_bgcolor="rgba(248,250,252,0.8)",
    paper_bgcolor="rgba(0,0,0,0)",
    font=dict(family="Inter, sans-serif", color="#475569", size=10),
    title_font=dict(family="Inter, sans-serif", color="#0B3B2F", size=12),
    xaxis=dict(gridcolor="rgba(203,213,225,0.4)", linecolor="rgba(203,213,225,0.6)",
               tickcolor="rgba(203,213,225,0.6)", zerolinecolor="rgba(203,213,225,0.3)"),
    yaxis=dict(gridcolor="rgba(203,213,225,0.4)", linecolor="rgba(203,213,225,0.6)",
               tickcolor="rgba(203,213,225,0.6)", zerolinecolor="rgba(203,213,225,0.3)"),
    margin=dict(t=50, b=70, l=50, r=20),
    height=350,
    legend=dict(bgcolor="rgba(0,0,0,0)", font=dict(color="#475569")),
)

# Modern biotech palette for charts
CHART_COLORS = ["#1A6D4C", "#2C7DA0", "#20B2AA", "#4A9E7A", "#0B3B2F", "#0288D1"]

# =============================================================================
# APP
# =============================================================================
app = dash.Dash(
    __name__,
    external_stylesheets=[
        dbc.themes.BOOTSTRAP,
        "https://use.fontawesome.com/releases/v5.15.4/css/all.css",
    ]
)

app.index_string = (
    '<!DOCTYPE html><html><head>{%metas%}'
    '<title>MACRO · Gestión de Muestreos</title>{%favicon%}{%css%}'
    '<style>' + CUSTOM_CSS + '</style>'
    '</head><body>{%app_entry%}'
    '<footer>{%config%}{%scripts%}{%renderer%}</footer>'
    '</body></html>'
)

server    = app.server
app.title = "MACRO · Gestión de Muestreos"

# ── LAYOUT ──────────────────────────────────────────────────────────────────
app.layout = dbc.Container([

    html.Div([
        html.Div([
            html.Div("🧬", className="header-icon"),
            html.Div([
                html.H1("Sistema de Gestión de Muestreos"),
                html.P("Biotecnología Aplicada · Control de Calidad · Muestreo de Plantas",
                       className="tagline"),
            ], className="header-text"),
        ], className="app-header"),
        html.Div(className="header-divider"),
    ]),

    dcc.Tabs(id="tabs", value="tab-muestra", children=[
        dcc.Tab(label="📊  Cálculo de Muestra",       value="tab-muestra"),
        dcc.Tab(label="📈  Análisis de Supervivencia", value="tab-supervivencia"),
    ]),

    html.Div(id="tab-content", style={"paddingTop":"24px","paddingBottom":"48px"})

], fluid=True, style={"maxWidth":"1360px","margin":"0 auto","padding":"0 28px"})


# =============================================================================
# TAB ROUTING
# =============================================================================
@app.callback(Output("tab-content","children"), Input("tabs","value"))
def render_tab(tab):

    if tab == "tab-muestra":

        excl = None
        if ids_excluidos:
            ids_txt = ", ".join(str(i) for i in ids_excluidos)
            excl = dbc.Alert([
                html.I(className="fas fa-info-circle me-2"),
                f"IDs excluidos (Vivero Perú · MN): {ids_txt}"
            ], color="warning", dismissable=True, style={"marginBottom":"16px"})

        n_lotes   = len(muestreos_hoy)
        n_plantas = int(muestreos_hoy['Macetas actuales'].sum()) \
                    if 'Macetas actuales' in muestreos_hoy.columns else 0

        kpi = html.Div([
            html.Div([
                html.Div(str(n_lotes), className="kpi-val"),
                html.Div("Lotes activos hoy", className="kpi-lbl"),
            ], className="kpi-cell"),
            html.Div([
                html.Div(f"{n_plantas:,}".replace(",","."), className="kpi-val amber"),
                html.Div("Plantas totales", className="kpi-lbl"),
            ], className="kpi-cell"),
            html.Div([
                html.Div(datetime.now().strftime("%d · %m · %Y"),
                         className="kpi-val sage sm"),
                html.Div("Fecha de muestreo", className="kpi-lbl"),
            ], className="kpi-cell"),
            html.Div([
                html.Div([html.Span(className="dot"), "Activo"],
                         className="kpi-val sm",
                         style={"display":"flex","alignItems":"center","justifyContent":"center"}),
                html.Div("Estado del sistema", className="kpi-lbl"),
            ], className="kpi-cell"),
        ], className="kpi-bar")

        if not muestreos_hoy.empty:
            items = [
                html.Li([
                    html.Span(f"ID {row['ID']}", className="lote-badge"),
                    html.Span(str(row['Código']), className="lote-code"),
                    html.Span(f"— {row.get('Variedad','')}", className="lote-var"),
                ]) for _, row in muestreos_hoy.iterrows()
            ]
            lotes_ui = dbc.Card([
                dbc.CardHeader([
                    html.I(className="fas fa-seedling me-2"),
                    f"Lotes programados hoy  ·  {n_lotes} registros"
                ]),
                dbc.CardBody(
                    html.Ul(items, className="lote-list"),
                    style={"maxHeight":"290px","overflowY":"auto","padding":"10px 14px"}
                )
            ], className="border-forest", style={"marginBottom":"20px"})
        else:
            lotes_ui = dbc.Alert("Sin lotes activos para hoy.", color="info",
                                  style={"marginBottom":"16px"})

        acc_row = dbc.Row([
            dbc.Col(dbc.Button([html.I(className="fas fa-file-excel me-2"),
                                "Generar Excel de muestreo"],
                               id="btn-generar-multiple", color="primary",
                               className="w-100"), md=5),
            dbc.Col(html.A([html.I(className="fas fa-download me-2"), "Descargar Excel"],
                           id="btn-descargar-multiple", href="", download="",
                           className="btn btn-success w-100"), md=5),
        ], className="g-3 mb-4")

        log_card = dbc.Card([
            dbc.CardHeader([html.I(className="fas fa-terminal me-2"), "Registro de operación"]),
            dbc.CardBody(
                html.Div("Presiona 'Generar Excel' para iniciar el procesamiento de lotes.",
                         id="resultado-multiple", className="log-box"),
            )
        ], className="border-sage")

        return html.Div([excl or html.Div(), kpi, lotes_ui, acc_row, log_card])

    else:
        return dbc.Container([
            dbc.Row([
                dbc.Col([
                    html.Div("Cargar archivo de resultados", className="sec-lbl"),
                    dcc.Upload(
                        id='upload-data',
                        children=html.Div([
                            html.I(className="fas fa-file-upload me-2",
                                   style={"color":"var(--primary)"}),
                            "Arrastra el Excel aquí o ",
                            html.Span("haz clic para explorar",
                                      style={"color":"var(--primary)","fontWeight":"600",
                                             "textDecoration":"underline"}),
                        ]),
                        style={'width':'100%','height':'66px','lineHeight':'66px',
                               'textAlign':'center','cursor':'pointer'},
                        className="upload-zone", multiple=False
                    ),
                ], md=7),
                dbc.Col([
                    html.Div("Hoja de análisis", className="sec-lbl"),
                    html.Div(id='selector-hoja-wrapper'),
                    dcc.Dropdown(id='selector-hoja',
                                 placeholder="Seleccione una hoja…",
                                 style={"borderRadius":"6px","fontSize":"0.84rem"}),
                ], md=5,
                   style={"display":"flex","flexDirection":"column","justifyContent":"flex-end"}),
            ], className="g-4 mb-4"),

            html.Div(id='output-alertas',    style={"marginBottom":"16px"}),
            html.Div(id='output-data-upload',style={"marginBottom":"24px"}),

            html.Div("Indicadores de calidad por fila", className="sec-lbl",
                     style={"marginBottom":"14px"}),

            dbc.Row([
                dbc.Col(html.Div(dcc.Graph(id="grafico-supervivencia",
                                           config={'displayModeBar':False}),
                                 className="chart-wrap"), md=4),
                dbc.Col(html.Div(dcc.Graph(id="grafico-talla-comercial",
                                           config={'displayModeBar':False}),
                                 className="chart-wrap"), md=4),
                dbc.Col(html.Div(dcc.Graph(id="grafico-ejes",
                                           config={'displayModeBar':False}),
                                 className="chart-wrap"), md=4),
            ], className="g-3 mb-3"),

            dbc.Row([
                dbc.Col(html.Div(dcc.Graph(id="grafico-ocupacion",
                                           config={'displayModeBar':False}),
                                 className="chart-wrap"), md=4),
                dbc.Col(html.Div(dcc.Graph(id="grafico-altura",
                                           config={'displayModeBar':False}),
                                 className="chart-wrap"), md=4),
                dbc.Col(html.Div(dcc.Graph(id="grafico-porcentaje-col",
                                           config={'displayModeBar':False}),
                                 className="chart-wrap"), md=4),
            ], className="g-3 mb-4"),
        ], fluid=True)


# =============================================================================
# CALLBACK GENERAR EXCEL
# =============================================================================
@app.callback(
    [Output("btn-descargar-multiple","href"),
     Output("btn-descargar-multiple","download"),
     Output("resultado-multiple","children")],
    Input("btn-generar-multiple","n_clicks"),
    prevent_initial_call=True
)
def generar_excel_multiple(n_clicks):
    if not n_clicks:
        return "", "", "En espera…"
    if muestreos_hoy.empty:
        return "", "", html.Span("Sin lotes disponibles para hoy.", className="log-warn")

    fecha_str    = datetime.now().strftime("%d-%m-%Y")
    nombre_excel = f"MUESTREOS_MACRO_{fecha_str}.xlsx"
    output       = BytesIO()
    workbook     = xlsxwriter.Workbook(output)
    lotes_ok = []; errores = []

    for _, lote in muestreos_hoy.iterrows():
        try:
            datos  = generar_datos_lote(lote)
            codigo = str(lote["Código"])
            m_act  = lote.get("muestreo_activador","")
            base   = f"{codigo}_{re.sub(r'\\s+','_',str(m_act).strip())}" if pd.notna(m_act) else codigo
            nombre = re.sub(r'[\\/*?:\[\]]','_', base)[:31]
            escribir_hoja(workbook, datos, nombre)
            lotes_ok.append(codigo)
        except Exception as e:
            errores.append(f"{lote.get('Código')}: {e}")

    workbook.close(); output.seek(0)
    href = ("data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,"
            + base64.b64encode(output.read()).decode())

    lines = [
        html.Div(f"✔  Proceso completado · {len(lotes_ok)} hoja(s) generadas",
                 className="log-ok", style={"marginBottom":"5px","fontWeight":"500"}),
        html.Div(f"Lotes procesados: {', '.join(lotes_ok)}",
                 style={"color":"var(--gray-600)"}),
    ]
    if errores:
        lines.append(html.Div(f"✗  Errores: {', '.join(errores)}",
                               className="log-error", style={"marginTop":"6px"}))
    return href, nombre_excel, html.Div(lines, className="log-box")


# =============================================================================
# CALLBACK SUPERVIVENCIA (MODIFICADO PARA INCLUIR EL ID)
# =============================================================================
@app.callback(
    [Output('selector-hoja-wrapper','children'),
     Output('selector-hoja','options'),
     Output('output-alertas','children'),
     Output('output-data-upload','children'),
     Output('grafico-supervivencia','figure'),
     Output('grafico-talla-comercial','figure'),
     Output('grafico-ejes','figure'),
     Output('grafico-ocupacion','figure'),
     Output('grafico-altura','figure'),
     Output('grafico-porcentaje-col','figure')],
    [Input('upload-data','contents'), Input('selector-hoja','value')],
    [State('upload-data','filename')]
)
def procesar_archivo_con_hoja(contents, hoja_seleccionada, filename):
    ef = {"layout": {**PLOT_LAYOUT,
                     "title":{"text":"Sin datos · carga un archivo",
                               "font":{"color":"#94A3B8"}}}}

    if contents is None:
        msg = html.Div(
            [html.I(className="fas fa-leaf me-2", style={"color":"var(--primary-light)"}),
             "Carga un archivo Excel para comenzar el análisis."],
            style={"textAlign":"center","padding":"28px","color":"var(--gray-500)",
                   "fontSize":"0.85rem","background":"var(--gray-50)",
                   "border":"1px dashed var(--gray-300)","borderRadius":"var(--radius-md)"}
        )
        return "", [], msg, None, ef, ef, ef, ef, ef, ef

    _, cs = contents.split(',')
    decoded = base64.b64decode(cs)

    try:
        hojas = pd.ExcelFile(BytesIO(decoded)).sheet_names
    except Exception as e:
        return "", [], dbc.Alert(f"Error: {e}", color="danger"), None, ef,ef,ef,ef,ef,ef

    opciones = [{'label': h, 'value': h} for h in hojas]

    if hoja_seleccionada is None or hoja_seleccionada not in hojas:
        return html.Div(), opciones, None, None, ef,ef,ef,ef,ef,ef

    try:
        df_raw = pd.read_excel(BytesIO(decoded), sheet_name=hoja_seleccionada, header=None)
    except Exception as e:
        return html.Div(), opciones, dbc.Alert(f"Error: {e}", color="danger"), None, ef,ef,ef,ef,ef,ef

    hri = next((i for i in range(len(df_raw)) if df_raw.iloc[i,0] == 'Fila'), None)
    if hri is None:
        return html.Div(), opciones, \
               dbc.Alert("No se encontró el encabezado 'Fila'.", color="warning"), \
               None, ef,ef,ef,ef,ef,ef

    df = pd.read_excel(BytesIO(decoded), sheet_name=hoja_seleccionada, header=hri)
    df['_t'] = df['Fila'].astype(str).str.strip()
    df = df[df['_t'].str.match(r'^\d+(\.\d+)?$', na=False)].drop(columns=['_t']).copy()

    if df.empty:
        return html.Div(), opciones, \
               dbc.Alert("Sin filas numéricas válidas.", color="warning"), \
               None, ef,ef,ef,ef,ef,ef

    for c in ['Máximo','Sobrevivencia','Talla Comercial','Ejes ≥ 2',
              'Ocup sustrato ≥ 80%','Altura ≥ 12 cm','% Col']:
        df[c] = pd.to_numeric(df.get(c, 0), errors='coerce').fillna(0)

    df['Fila'] = pd.to_numeric(df['Fila'], errors='coerce').fillna(0).astype(int).astype(str)

    tot = df['Máximo'].sum()
    if tot == 0:
        return html.Div(), opciones, \
               dbc.Alert("Total Máximo = 0, no se puede calcular.", color="danger"), \
               None, ef,ef,ef,ef,ef,ef

    def pct(c): return (df[c].sum() / tot) * 100

    ts = pct('Sobrevivencia'); tc = pct('Talla Comercial'); tejs = pct('Ejes ≥ 2')
    tocu = pct('Ocup sustrato ≥ 80%'); talt = pct('Altura ≥ 12 cm')
    tcol = pct('% Col') if df['% Col'].sum() > 0 else 0

    # Alarmas
    cond = ((df['Sobrevivencia']>df['Máximo'])|(df['Talla Comercial']>df['Máximo'])|
            (df['Ejes ≥ 2']>df['Máximo'])|(df['Ocup sustrato ≥ 80%']>df['Máximo'])|
            (df['Altura ≥ 12 cm']>df['Máximo'])|(df['% Col']>df['Máximo']))
    fa = df[cond]

    alerta_ui = (
        html.Div([
            dbc.Alert([html.I(className="fas fa-exclamation-triangle me-2"),
                       f"{len(fa)} fila(s) con valores fuera de rango"],
                      color="warning", style={"marginBottom":"10px"}),
            dash_table.DataTable(
                data=fa.to_dict('records'),
                columns=[{'name':i,'id':i} for i in fa.columns],
                style_table={'overflowX':'auto'},
                style_cell={'textAlign':'center','padding':'6px','fontSize':'12px',
                            'fontFamily':'JetBrains Mono, monospace',
                            'backgroundColor':'#fff','color':'#1E293B',
                            'border':'1px solid #E2E8F0'},
                style_header={'backgroundColor':'#E8F3EF','fontWeight':'600',
                              'color':'#0B3B2F','border':'1px solid #CBD5E1',
                              'fontSize':'0.7rem','letterSpacing':'0.05em'},
                page_size=8)
        ]) if not fa.empty else
        dbc.Alert([html.I(className="fas fa-check-circle me-2"),
                   "Sin alarmas — todos los valores dentro del rango esperado."],
                  color="success")
    )

    # ========== METADATOS: se agrega la lectura del ID ==========
    try:
        meta = pd.read_excel(BytesIO(decoded), sheet_name=hoja_seleccionada, header=None)
        # ID del lote (fila 5, columna A)
        id_lote = meta.iloc[5,0] if meta.shape[0]>5 and meta.shape[1]>0 else "—"
        # Fecha muestreo (fila 5, columna F)
        fm = meta.iloc[5,5] if meta.shape[0]>5 and meta.shape[1]>5 else "—"
        # Código de lote (fila 7, columna C)
        lm = meta.iloc[7,2] if meta.shape[0]>7 and meta.shape[1]>2 else "—"
        if isinstance(fm,(int,float)):
            fm = (pd.to_datetime("1899-12-30")+pd.to_timedelta(int(fm),"D")).strftime('%d-%m-%Y')
        elif hasattr(fm,'strftime'):
            fm = fm.strftime('%d-%m-%Y')
        else:
            fm = str(fm)
    except Exception:
        id_lote = "—"
        fm = "—"
        lm = "—"

    # ========== KPI BAR con 6 celdas (agregamos ID) ==========
    kpi = html.Div([
        html.Div([
            html.Div(f"{ts:.1f}%".replace('.',','), className="kpi-val"),
            html.Div("Supervivencia", className="kpi-lbl")
        ], className="kpi-cell"),
        html.Div([
            html.Div(f"{tc:.1f}%".replace('.',','), className="kpi-val amber"),
            html.Div("Talla comercial", className="kpi-lbl")
        ], className="kpi-cell"),
        html.Div([
            html.Div(f"{int(tot):,}".replace(",","."), className="kpi-val sage sm"),
            html.Div("Macetas muestreadas", className="kpi-lbl")
        ], className="kpi-cell"),
        # Nueva celda para el ID
        html.Div([
            html.Div(str(id_lote), className="kpi-val sm"),
            html.Div("ID", className="kpi-lbl")
        ], className="kpi-cell"),
        html.Div([
            html.Div(str(lm), className="kpi-val sm"),
            html.Div("Lote", className="kpi-lbl")
        ], className="kpi-cell"),
        html.Div([
            html.Div(str(fm), className="kpi-val sky sm"),
            html.Div("Fecha muestreo", className="kpi-lbl")
        ], className="kpi-cell"),
    ], className="kpi-bar")

    tabla = html.Div([
        html.Div("Datos registrados por fila", className="sec-lbl",
                 style={"marginTop":"20px","marginBottom":"10px"}),
        dash_table.DataTable(
            data=df.to_dict('records'),
            columns=[{'name':i,'id':i} for i in df.columns],
            style_table={'overflowX':'auto'},
            style_cell={'textAlign':'center','padding':'7px','fontSize':'12px',
                        'fontFamily':'JetBrains Mono, monospace',
                        'backgroundColor':'#fff','color':'#1E293B',
                        'border':'1px solid #E2E8F0'},
            style_header={'backgroundColor':'#E8F3EF','fontWeight':'600',
                          'color':'#0B3B2F','border':'1px solid #CBD5E1',
                          'fontSize':'0.7rem','letterSpacing':'0.05em',
                          'textTransform':'uppercase'},
            page_size=10)
    ])

    filas_u = df['Fila'].tolist()

    def graf(col, titulo, color, ylabel):
        if col not in df.columns:
            return {**ef}
        fig = px.bar(df, x='Fila', y=col,
                     labels={'Fila':'Fila', col: ylabel},
                     color_discrete_sequence=[color])
        fig.update_traces(text=df[col], textposition='outside',
                          marker_color=color,
                          marker_line_color="rgba(255,255,255,0.7)",
                          marker_line_width=0.8, opacity=0.9)
        lay = dict(PLOT_LAYOUT)
        lay["title"] = {"text": titulo,
                        "font":{"family":"Inter, sans-serif","color":"#0B3B2F","size":12}}
        lay["xaxis"] = dict(PLOT_LAYOUT["xaxis"],
                            tickmode='array', tickvals=filas_u,
                            ticktext=filas_u, tickangle=-45)
        lay["yaxis"] = dict(PLOT_LAYOUT["yaxis"], title=ylabel)
        fig.update_layout(**lay)
        return fig

    C = CHART_COLORS
    return (html.Div(), opciones, alerta_ui, html.Div([kpi, tabla]),
            graf('Sobrevivencia',       f'Supervivencia — {ts:.1f}%',   C[0], 'Plantas vivas'),
            graf('Talla Comercial',     f'Talla Comercial — {tc:.1f}%', C[1], 'En talla comercial'),
            graf('Ejes ≥ 2',            f'Ejes ≥ 2 — {tejs:.1f}%',     C[2], 'Con ejes ≥ 2'),
            graf('Ocup sustrato ≥ 80%', f'Ocup. Sustrato — {tocu:.1f}%',C[3], 'Ocup ≥ 80%'),
            graf('Altura ≥ 12 cm',      f'Altura ≥ 12 cm — {talt:.1f}%',C[4], 'Alt ≥ 12 cm'),
            graf('% Col', f'% Col — {tcol:.1f}%', C[5], '% Col')
            if df['% Col'].sum() > 0 else
            {**ef, "layout":{**PLOT_LAYOUT,"title":{"text":"% Col — sin datos"}}})


# =============================================================================
if __name__ == "__main__":
    app.run(host='0.0.0.0', port=8050, debug=True)