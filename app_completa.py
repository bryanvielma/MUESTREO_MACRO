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

def es_lote_peru(row):
    imc = row.get("I-M-C", "")
    return isinstance(imc, str) and "MN" in imc.upper()

hoy_date = datetime.now().date()
if 'fecha_activadora' in muestreos_hoy_raw.columns:
    muestreos_hoy_raw['fecha_activadora'] = pd.to_datetime(muestreos_hoy_raw['fecha_activadora'], errors='coerce')
    mascara_fecha = muestreos_hoy_raw['fecha_activadora'].dt.date == hoy_date
    muestreos_hoy_raw = muestreos_hoy_raw[mascara_fecha].copy()

ids_excluidos_hoy = muestreos_hoy_raw[muestreos_hoy_raw.apply(es_lote_peru, axis=1)]["ID"].tolist()
ids_excluidos = sorted(set(ids_excluidos_hoy))

def filtrar_sin_mn(df):
    if 'I-M-C' not in df.columns:
        return df
    mask = df.apply(lambda row: not es_lote_peru(row), axis=1)
    return df[mask].copy()

muestreos_hoy = filtrar_sin_mn(muestreos_hoy_raw)
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
        rangos = [(8,2),(15,3),(25,5),(50,8),(90,13),(150,20),(280,40),
                  (500,60),(1200,80),(3200,140),(10000,200),(35000,320),
                  (150000,500),(500000,800)]
        for limite, muestra in rangos:
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
                rows.append({"Muestra": muestra_num, "Número Planta": start_plant + off_p, "Fila": r_idx})
                muestra_num += 1
        else:
            for off_p in range(hileras):
                rows.append({"Muestra": muestra_num, "Número Planta": start_plant + off_p, "Fila": r_idx})
                muestra_num += 1

    if not rows:
        raise ValueError("No se generaron datos de muestra.")

    return {
        "tabla_df": pd.DataFrame(rows),
        "lote": lote,
        "cantidad": cantidad,
        "imc_val": imc_val,
        "bandejas_val": bandejas_val,
        "litros": litros,
        "muestra_tamano": muestra_tamano
    }

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

    worksheet.write(start_row, 0, "Bandeja", fmt_bold)
    worksheet.write(start_row, 1, "Vol. Sustrato (L)", fmt_bold)
    worksheet.write(start_row, 2, "Código", fmt_bold)
    worksheet.write(start_row+1, 0, bandejas_val, fmt_norm)
    worksheet.write(start_row+1, 1, f"{litros:.2f}", fmt_norm)
    worksheet.write(start_row+1, 2, lote["Código"], fmt_norm)
    worksheet.set_row(start_row, 12)
    worksheet.set_row(start_row+1, 12)
    start_row += 2

    worksheet.merge_range(f'D{start_row-1}:H{start_row-1}', 'INVERNADERO - MESÓN - CANTIDAD (I-M-C)', fmt_bold)
    worksheet.merge_range(f'D{start_row}:H{start_row}', str(imc_val), fmt_norm)
    worksheet.set_row(start_row-1, 12)
    worksheet.set_row(start_row, 12)
    start_row += 2

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

    blank = last_data + 1
    worksheet.set_row(blank, 5)
    resp_row = blank + 1
    worksheet.merge_range(f'A{resp_row+1}:D{resp_row+1}', 'Responsable: _________________________________________________', fmt_center)
    worksheet.merge_range(f'F{resp_row+1}:H{resp_row+1}', 'Fecha: ______ /______ /_________', fmt_center)
    worksheet.write(f'E{resp_row+1}', 'Firma: ___________', fmt_center)
    worksheet.set_row(resp_row, 12)

    blank2 = resp_row + 1
    worksheet.set_row(blank2, 5)
    pct_row = blank2 + 1
    worksheet.merge_range(f'A{pct_row+1}:B{pct_row+1}', '% PLANTAS PLANTABLES', fmt_small)
    worksheet.write(f'C{pct_row+1}', '', fmt_small)
    worksheet.merge_range(f'F{pct_row+1}:G{pct_row+1}', '% TALLA COMERCIAL', fmt_small)
    worksheet.write(f'H{pct_row+1}', '', fmt_small)
    worksheet.set_row(pct_row, 10)

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

    for col, w in [('A',8),('B',13),('C',23),('D',12),('E',18),('F',18),('G',13),('H',9)]:
        worksheet.set_column(f'{col}:{col}', w)
    worksheet.set_row(3, 5)

# =============================================================================
# CUSTOM CSS — DARK TECH / BIOPUNK AESTHETIC
# =============================================================================
CUSTOM_CSS = """
@import url('https://fonts.googleapis.com/css2?family=Share+Tech+Mono&family=Rajdhani:wght@300;400;500;600;700&family=Exo+2:wght@200;300;400;600&display=swap');

:root {
  --bg-void:        #060a0f;
  --bg-deep:        #0a1018;
  --bg-panel:       #0d1520;
  --bg-card:        #111d2b;
  --bg-card-hover:  #152130;
  --accent-cyan:    #00e5ff;
  --accent-green:   #39ff8a;
  --accent-lime:    #b0ff00;
  --accent-amber:   #ffaa00;
  --accent-red:     #ff3a5c;
  --text-primary:   #e0f0ff;
  --text-secondary: #6fa8c8;
  --text-dim:       #3a6280;
  --border-glow:    rgba(0, 229, 255, 0.25);
  --border-subtle:  rgba(0, 229, 255, 0.08);
  --glow-cyan:      0 0 20px rgba(0,229,255,0.35), 0 0 60px rgba(0,229,255,0.08);
  --glow-green:     0 0 20px rgba(57,255,138,0.35), 0 0 60px rgba(57,255,138,0.08);
}

*, *::before, *::after { box-sizing: border-box; }

html, body, #react-entry-point, ._dash-loading, .dash-renderer {
  background: var(--bg-void) !important;
  color: var(--text-primary) !important;
  font-family: 'Exo 2', sans-serif !important;
  min-height: 100vh;
}

/* ── ANIMATED GRID BACKGROUND ── */
body::before {
  content: '';
  position: fixed;
  inset: 0;
  background-image:
    linear-gradient(rgba(0,229,255,0.03) 1px, transparent 1px),
    linear-gradient(90deg, rgba(0,229,255,0.03) 1px, transparent 1px);
  background-size: 40px 40px;
  pointer-events: none;
  z-index: 0;
}

body::after {
  content: '';
  position: fixed;
  inset: 0;
  background: radial-gradient(ellipse 80% 60% at 50% 0%, rgba(0,229,255,0.04) 0%, transparent 70%);
  pointer-events: none;
  z-index: 0;
}

.container-fluid { position: relative; z-index: 1; }

/* ── HEADER ── */
.app-header {
  padding: 32px 0 20px;
  text-align: center;
  position: relative;
}

.app-header h1 {
  font-family: 'Rajdhani', sans-serif !important;
  font-size: 2.4rem !important;
  font-weight: 700 !important;
  letter-spacing: 0.18em !important;
  text-transform: uppercase;
  color: var(--accent-cyan) !important;
  text-shadow: var(--glow-cyan);
  margin: 0 !important;
}

.app-header .subtitle {
  font-family: 'Share Tech Mono', monospace;
  font-size: 0.7rem;
  color: var(--text-dim);
  letter-spacing: 0.3em;
  text-transform: uppercase;
  margin-top: 6px;
}

.header-line {
  width: 100%;
  height: 1px;
  background: linear-gradient(90deg, transparent, var(--accent-cyan), var(--accent-green), transparent);
  margin-top: 20px;
  opacity: 0.6;
}

/* ── TABS ── */
.custom-tabs .tab-container {
  display: flex;
  gap: 4px;
  padding: 12px 0;
  border-bottom: 1px solid var(--border-glow);
  margin-bottom: 24px;
}

/* Override Dash default tab styles */
.dash-tabs-container { background: transparent !important; }
.dash-tabs { border-bottom: 1px solid var(--border-glow) !important; background: transparent !important; }

.dash-tab {
  font-family: 'Rajdhani', sans-serif !important;
  font-weight: 600 !important;
  font-size: 0.85rem !important;
  letter-spacing: 0.12em !important;
  text-transform: uppercase !important;
  color: var(--text-secondary) !important;
  background: var(--bg-card) !important;
  border: 1px solid var(--border-subtle) !important;
  border-bottom: none !important;
  padding: 12px 28px !important;
  transition: all 0.25s ease !important;
  clip-path: polygon(8px 0%, 100% 0%, 100% 100%, 0% 100%);
}

.dash-tab:hover {
  color: var(--accent-cyan) !important;
  background: var(--bg-card-hover) !important;
  border-color: var(--border-glow) !important;
}

.dash-tab--selected {
  color: var(--accent-cyan) !important;
  background: var(--bg-panel) !important;
  border-color: var(--accent-cyan) !important;
  border-bottom: 2px solid var(--accent-cyan) !important;
  box-shadow: var(--glow-cyan) !important;
}

/* ── CARDS ── */
.tech-card {
  background: var(--bg-card) !important;
  border: 1px solid var(--border-subtle) !important;
  border-radius: 2px !important;
  position: relative;
  overflow: hidden;
  transition: border-color 0.3s ease, box-shadow 0.3s ease;
}

.tech-card::before {
  content: '';
  position: absolute;
  top: 0; left: 0; right: 0;
  height: 2px;
  background: linear-gradient(90deg, var(--accent-cyan), var(--accent-green));
}

.tech-card:hover {
  border-color: var(--border-glow) !important;
  box-shadow: 0 4px 32px rgba(0,229,255,0.1) !important;
}

.card { background: var(--bg-card) !important; border: 1px solid var(--border-subtle) !important; border-radius: 2px !important; }
.card-header {
  background: var(--bg-panel) !important;
  border-bottom: 1px solid var(--border-glow) !important;
  font-family: 'Rajdhani', sans-serif !important;
  font-weight: 600 !important;
  font-size: 0.8rem !important;
  letter-spacing: 0.15em !important;
  text-transform: uppercase !important;
  color: var(--accent-cyan) !important;
  padding: 10px 16px !important;
}
.card-body { background: var(--bg-card) !important; color: var(--text-primary) !important; }

/* ── BUTTONS ── */
.btn-primary {
  background: transparent !important;
  border: 1px solid var(--accent-cyan) !important;
  color: var(--accent-cyan) !important;
  font-family: 'Rajdhani', sans-serif !important;
  font-weight: 600 !important;
  font-size: 0.8rem !important;
  letter-spacing: 0.2em !important;
  text-transform: uppercase !important;
  padding: 10px 24px !important;
  border-radius: 1px !important;
  clip-path: polygon(6px 0%, 100% 0%, calc(100% - 6px) 100%, 0% 100%);
  transition: all 0.2s ease !important;
  position: relative;
  overflow: hidden;
}

.btn-primary::before {
  content: '';
  position: absolute;
  inset: 0;
  background: var(--accent-cyan);
  transform: scaleX(0);
  transform-origin: left;
  transition: transform 0.2s ease;
  z-index: -1;
}

.btn-primary:hover {
  color: var(--bg-void) !important;
  background: var(--accent-cyan) !important;
  box-shadow: var(--glow-cyan) !important;
}

.btn-success {
  background: transparent !important;
  border: 1px solid var(--accent-green) !important;
  color: var(--accent-green) !important;
  font-family: 'Rajdhani', sans-serif !important;
  font-weight: 600 !important;
  font-size: 0.8rem !important;
  letter-spacing: 0.2em !important;
  text-transform: uppercase !important;
  padding: 10px 24px !important;
  border-radius: 1px !important;
  clip-path: polygon(6px 0%, 100% 0%, calc(100% - 6px) 100%, 0% 100%);
  transition: all 0.2s ease !important;
}

.btn-success:hover {
  color: var(--bg-void) !important;
  background: var(--accent-green) !important;
  box-shadow: var(--glow-green) !important;
}

/* ── ALERTS ── */
.alert {
  border-radius: 1px !important;
  border-left: 3px solid var(--accent-amber) !important;
  background: rgba(255,170,0,0.06) !important;
  color: var(--accent-amber) !important;
  border-top: none !important; border-right: none !important; border-bottom: none !important;
  font-family: 'Share Tech Mono', monospace !important;
  font-size: 0.78rem !important;
}

.alert-info {
  border-left-color: var(--accent-cyan) !important;
  background: rgba(0,229,255,0.05) !important;
  color: var(--accent-cyan) !important;
}

.alert-success {
  border-left-color: var(--accent-green) !important;
  background: rgba(57,255,138,0.05) !important;
  color: var(--accent-green) !important;
}

/* ── LOTE LIST ── */
.lote-list {
  list-style: none;
  padding: 0;
  margin: 0;
}

.lote-list li {
  padding: 8px 14px;
  margin-bottom: 4px;
  background: var(--bg-panel);
  border-left: 2px solid var(--accent-cyan);
  font-family: 'Share Tech Mono', monospace;
  font-size: 0.78rem;
  color: var(--text-secondary);
  transition: all 0.2s;
}

.lote-list li:hover {
  background: var(--bg-card-hover);
  color: var(--accent-cyan);
  border-left-color: var(--accent-green);
}

/* ── UPLOAD ZONE ── */
.upload-zone {
  border: 1px dashed rgba(0,229,255,0.3) !important;
  border-radius: 2px !important;
  background: var(--bg-panel) !important;
  color: var(--text-secondary) !important;
  font-family: 'Exo 2', sans-serif !important;
  font-size: 0.85rem !important;
  transition: all 0.3s ease !important;
  cursor: pointer;
  position: relative;
}

.upload-zone::before {
  content: '⬆ ';
  color: var(--accent-cyan);
}

.upload-zone:hover {
  border-color: var(--accent-cyan) !important;
  background: rgba(0,229,255,0.04) !important;
  color: var(--accent-cyan) !important;
  box-shadow: inset 0 0 20px rgba(0,229,255,0.05) !important;
}

/* ── DROPDOWN ── */
.Select-control, .Select-menu-outer {
  background: var(--bg-card) !important;
  border: 1px solid var(--border-glow) !important;
  border-radius: 1px !important;
  color: var(--text-primary) !important;
}

.Select-value-label { color: var(--accent-cyan) !important; font-family: 'Share Tech Mono', monospace !important; }
.Select-option { background: var(--bg-panel) !important; color: var(--text-secondary) !important; }
.Select-option:hover { background: var(--bg-card-hover) !important; color: var(--accent-cyan) !important; }
.Select-placeholder { color: var(--text-dim) !important; }

.VirtualizedSelectOption { background: var(--bg-panel) !important; color: var(--text-secondary) !important; }
.VirtualizedSelectFocusedOption { background: var(--bg-card-hover) !important; color: var(--accent-cyan) !important; }

/* ── DATA TABLE ── */
.dash-spreadsheet-container, .dash-table-container { background: transparent !important; }

.dash-spreadsheet { background: var(--bg-card) !important; }

.dash-header {
  background: var(--bg-panel) !important;
  color: var(--accent-cyan) !important;
  font-family: 'Share Tech Mono', monospace !important;
  font-size: 0.72rem !important;
  letter-spacing: 0.1em !important;
  border-bottom: 1px solid var(--border-glow) !important;
}

.dash-cell {
  background: var(--bg-card) !important;
  color: var(--text-secondary) !important;
  font-family: 'Share Tech Mono', monospace !important;
  font-size: 0.75rem !important;
  border-color: var(--border-subtle) !important;
}

.dash-cell:hover { background: var(--bg-card-hover) !important; color: var(--text-primary) !important; }

/* ── GRAPHS ── */
.js-plotly-plot .plotly { background: transparent !important; }

/* ── STAT BADGE ── */
.stat-badge {
  display: inline-flex;
  align-items: center;
  gap: 8px;
  padding: 6px 14px;
  background: var(--bg-panel);
  border: 1px solid var(--border-glow);
  border-radius: 1px;
  font-family: 'Share Tech Mono', monospace;
  font-size: 0.75rem;
  color: var(--text-secondary);
  margin: 4px;
}

.stat-badge span { color: var(--accent-cyan); font-weight: bold; font-size: 0.9rem; }

/* ── SECTION LABEL ── */
.section-label {
  font-family: 'Share Tech Mono', monospace;
  font-size: 0.65rem;
  letter-spacing: 0.3em;
  text-transform: uppercase;
  color: var(--text-dim);
  margin-bottom: 10px;
  display: flex;
  align-items: center;
  gap: 10px;
}

.section-label::after {
  content: '';
  flex: 1;
  height: 1px;
  background: var(--border-subtle);
}

/* ── SCROLLBAR ── */
::-webkit-scrollbar { width: 6px; height: 6px; }
::-webkit-scrollbar-track { background: var(--bg-void); }
::-webkit-scrollbar-thumb { background: var(--accent-cyan); opacity: 0.3; border-radius: 0; }
::-webkit-scrollbar-thumb:hover { opacity: 0.8; }

/* ── SUMMARY CARD ── */
.summary-stat {
  text-align: center;
  padding: 16px;
  border-right: 1px solid var(--border-subtle);
}
.summary-stat:last-child { border-right: none; }
.summary-stat .val {
  font-family: 'Rajdhani', sans-serif;
  font-size: 1.8rem;
  font-weight: 700;
  color: var(--accent-cyan);
  text-shadow: var(--glow-cyan);
  line-height: 1;
}
.summary-stat .lbl {
  font-family: 'Share Tech Mono', monospace;
  font-size: 0.62rem;
  letter-spacing: 0.15em;
  color: var(--text-dim);
  text-transform: uppercase;
  margin-top: 4px;
}

/* ── TICKER / PULSE ── */
@keyframes pulse-border {
  0%, 100% { opacity: 0.4; }
  50% { opacity: 1; }
}

.pulse { animation: pulse-border 2s ease-in-out infinite; }

@keyframes scan-line {
  0% { transform: translateY(-100%); }
  100% { transform: translateY(100vh); }
}

.header-scan {
  position: fixed;
  top: 0; left: 0; right: 0;
  height: 2px;
  background: linear-gradient(90deg, transparent, var(--accent-cyan), transparent);
  opacity: 0.2;
  animation: scan-line 8s linear infinite;
  pointer-events: none;
  z-index: 9999;
}
"""

# =============================================================================
# APP DASH — REDESIGNED
# =============================================================================
app = dash.Dash(
    __name__,
    external_stylesheets=[
        dbc.themes.CYBORG,
        "https://use.fontawesome.com/releases/v5.15.4/css/all.css"
    ]
)

# Inject custom CSS
app.index_string = '''
<!DOCTYPE html>
<html>
    <head>
        {%metas%}
        <title>MACRO — Sistema de Muestreos</title>
        {%favicon%}
        {%css%}
        <style>
''' + CUSTOM_CSS + '''
        </style>
    </head>
    <body>
        <div class="header-scan"></div>
        {%app_entry%}
        <footer>
            {%config%}
            {%scripts%}
            {%renderer%}
        </footer>
    </body>
</html>
'''

server = app.server
app.title = "MACRO — Sistema de Muestreos"

# ── PLOTLY TEMPLATE ─────────────────────────────────────────────────────────
PLOT_LAYOUT = dict(
    plot_bgcolor="rgba(13,21,32,0.8)",
    paper_bgcolor="rgba(0,0,0,0)",
    font=dict(family="Share Tech Mono, monospace", color="#6fa8c8", size=10),
    title_font=dict(family="Rajdhani, sans-serif", color="#00e5ff", size=13),
    xaxis=dict(
        gridcolor="rgba(0,229,255,0.06)",
        linecolor="rgba(0,229,255,0.2)",
        tickcolor="rgba(0,229,255,0.2)",
        zerolinecolor="rgba(0,229,255,0.1)",
    ),
    yaxis=dict(
        gridcolor="rgba(0,229,255,0.06)",
        linecolor="rgba(0,229,255,0.2)",
        tickcolor="rgba(0,229,255,0.2)",
        zerolinecolor="rgba(0,229,255,0.1)",
    ),
    margin=dict(t=55, b=75, l=50, r=20),
    height=380,
    legend=dict(bgcolor="rgba(0,0,0,0)", font=dict(color="#6fa8c8")),
)

CHART_COLORS = ["#00e5ff", "#39ff8a", "#b0ff00", "#ffaa00", "#ff3a5c", "#c47aff"]

# ── LAYOUT ──────────────────────────────────────────────────────────────────
app.layout = dbc.Container([

    # Header
    html.Div([
        html.Div([
            html.H1("MACRO // Sistema de Gestión de Muestreos"),
            html.P("BIOTECNOLOGÍA APLICADA · CONTROL DE CALIDAD · ANÁLISIS EN TIEMPO REAL",
                   className="subtitle"),
        ], className="app-header"),
        html.Div(className="header-line"),
    ]),

    # Tabs
    dcc.Tabs(
        id="tabs",
        value="tab-muestra",
        className="custom-tabs",
        children=[
            dcc.Tab(label="⬡  CÁLCULO DE MUESTRA", value="tab-muestra"),
            dcc.Tab(label="◈  ANÁLISIS DE SUPERVIVENCIA", value="tab-supervivencia"),
        ]
    ),

    html.Div(id="tab-content", style={"paddingTop": "8px"})

], fluid=True, style={"maxWidth": "1400px", "margin": "0 auto", "padding": "0 24px 48px"})


# =============================================================================
# TAB ROUTING
# =============================================================================
@app.callback(Output("tab-content", "children"), Input("tabs", "value"))
def render_tab(tab):
    if tab == "tab-muestra":
        # ── Warning exclusion alert
        alerta_excl = None
        if ids_excluidos:
            ids_texto = ", ".join(str(i) for i in ids_excluidos)
            alerta_excl = html.Div([
                html.Div([
                    html.Span("⚠", style={"fontSize": "1rem", "marginRight": "10px", "color": "#ffaa00"}),
                    html.Span(f"IDs excluidos (Vivero Perú · MN): {ids_texto}",
                              style={"fontFamily": "Share Tech Mono, monospace", "fontSize": "0.75rem"})
                ], className="alert")
            ], style={"marginBottom": "16px"})

        # ── Lote count badge
        n_lotes = len(muestreos_hoy)
        badge_row = html.Div([
            html.Div([
                html.Div(str(n_lotes), className="val"),
                html.Div("LOTES HOY", className="lbl"),
            ], className="summary-stat"),
            html.Div([
                html.Div(
                    str(int(muestreos_hoy['Macetas actuales'].sum())) if 'Macetas actuales' in muestreos_hoy.columns else "—",
                    className="val"
                ),
                html.Div("PLANTAS TOTALES", className="lbl"),
            ], className="summary-stat"),
            html.Div([
                html.Div(datetime.now().strftime("%d/%m/%Y"), className="val",
                         style={"fontSize": "1.1rem", "letterSpacing": "0.05em"}),
                html.Div("FECHA ACTIVA", className="lbl"),
            ], className="summary-stat"),
        ], style={
            "display": "flex",
            "background": "var(--bg-card)",
            "border": "1px solid var(--border-subtle)",
            "marginBottom": "20px",
        })

        # ── Lote list card
        if not muestreos_hoy.empty:
            lote_items = [
                html.Li([
                    html.Span(f"[{row['ID']}]",
                              style={"color": "var(--accent-green)", "marginRight": "10px", "fontSize": "0.72rem"}),
                    html.Span(f"{row['Código']}",
                              style={"color": "var(--accent-cyan)", "marginRight": "6px"}),
                    html.Span(f"— {row.get('Variedad', '')}",
                              style={"color": "var(--text-dim)", "fontSize": "0.75rem"}),
                ])
                for _, row in muestreos_hoy.iterrows()
            ]
            lote_card = dbc.Card([
                dbc.CardHeader([
                    html.I(className="fas fa-leaf me-2"),
                    f"LOTES PROGRAMADOS HOY  ·  {n_lotes} ACTIVOS"
                ]),
                dbc.CardBody(
                    html.Ul(lote_items, className="lote-list"),
                    style={"maxHeight": "280px", "overflowY": "auto", "padding": "12px 16px"}
                )
            ], className="tech-card", style={"marginBottom": "20px"})
        else:
            lote_card = html.Div(
                "⊘  Sin lotes activos para hoy",
                className="alert alert-info",
                style={"marginBottom": "16px"}
            )

        # ── Action buttons
        btn_row = dbc.Row([
            dbc.Col(
                dbc.Button([
                    html.I(className="fas fa-file-excel me-2"),
                    "GENERAR EXCEL DE MUESTREO"
                ], id="btn-generar-multiple", color="primary", className="w-100"),
                width=6
            ),
            dbc.Col(
                html.A([
                    html.I(className="fas fa-download me-2"),
                    "DESCARGAR EXCEL"
                ], id="btn-descargar-multiple", href="", download="",
                   className="btn btn-success w-100"),
                width=6
            ),
        ], className="g-3 mb-3")

        # ── Result card
        result_card = dbc.Card([
            dbc.CardHeader([html.I(className="fas fa-terminal me-2"), "LOG DE GENERACIÓN"]),
            dbc.CardBody(
                html.Div(
                    html.Span("// En espera de ejecución...",
                              style={"fontFamily": "Share Tech Mono, monospace",
                                     "fontSize": "0.78rem",
                                     "color": "var(--text-dim)"}),
                    id="resultado-multiple"
                )
            )
        ], className="tech-card")

        return html.Div([
            alerta_excl or html.Div(),
            badge_row,
            lote_card,
            btn_row,
            result_card,
        ])

    else:
        # ── TAB 2: Supervivencia
        return dbc.Container([
            # Upload + sheet selector row
            dbc.Row([
                dbc.Col([
                    html.Div("// CARGAR ARCHIVO DE RESULTADOS", className="section-label"),
                    dcc.Upload(
                        id='upload-data',
                        children=html.Div([
                            "Arrastra el archivo Excel o ",
                            html.Span("haz clic para seleccionar",
                                      style={"color": "var(--accent-cyan)", "textDecoration": "underline"})
                        ]),
                        style={
                            'width': '100%', 'height': '64px', 'lineHeight': '64px',
                            'textAlign': 'center', 'cursor': 'pointer',
                        },
                        className="upload-zone",
                        multiple=False
                    ),
                ], md=7),
                dbc.Col([
                    html.Div("// SELECCIONAR HOJA", className="section-label"),
                    html.Div(id='selector-hoja-wrapper'),
                    dcc.Dropdown(
                        id='selector-hoja',
                        placeholder="— seleccione hoja —",
                        style={
                            "background": "var(--bg-card)",
                            "border": "1px solid var(--border-glow)",
                            "borderRadius": "1px",
                            "color": "var(--text-primary)",
                            "fontFamily": "Share Tech Mono, monospace",
                            "fontSize": "0.8rem",
                        }
                    ),
                ], md=5),
            ], className="g-4 mb-3"),

            # Alerts
            html.Div(id='output-alertas', style={"marginBottom": "16px"}),

            # Summary / metadata card
            html.Div(id='output-data-upload', style={"marginBottom": "20px"}),

            # Charts row 1
            html.Div("// INDICADORES DE CALIDAD", className="section-label", style={"marginBottom": "12px"}),
            dbc.Row([
                dbc.Col(dcc.Graph(id="grafico-supervivencia",
                                  config={'displayModeBar': False}), md=4),
                dbc.Col(dcc.Graph(id="grafico-talla-comercial",
                                  config={'displayModeBar': False}), md=4),
                dbc.Col(dcc.Graph(id="grafico-ejes",
                                  config={'displayModeBar': False}), md=4),
            ], className="g-3 mb-2"),

            # Charts row 2
            dbc.Row([
                dbc.Col(dcc.Graph(id="grafico-ocupacion",
                                  config={'displayModeBar': False}), md=4),
                dbc.Col(dcc.Graph(id="grafico-altura",
                                  config={'displayModeBar': False}), md=4),
                dbc.Col(dcc.Graph(id="grafico-porcentaje-col",
                                  config={'displayModeBar': False}), md=4),
            ], className="g-3 mb-4"),

        ], fluid=True)


# =============================================================================
# CALLBACK GENERAR EXCEL (TAB 1)
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
        return "", "", "// En espera..."
    if muestreos_hoy.empty:
        return "", "", html.Span("⊘  Sin lotes disponibles para muestreo hoy.",
                                  style={"color": "var(--accent-amber)",
                                         "fontFamily": "Share Tech Mono, monospace",
                                         "fontSize": "0.8rem"})

    fecha_str = datetime.now().strftime("%d-%m-%Y")
    nombre_excel = f"MUESTREOS_MACRO_{fecha_str}.xlsx"
    output = BytesIO()
    workbook = xlsxwriter.Workbook(output)

    lotes_procesados = []
    errores = []
    for _, lote in muestreos_hoy.iterrows():
        try:
            datos = generar_datos_lote(lote)
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

    lines = [
        html.Div(f"✓ PROCESO COMPLETADO  ·  {len(lotes_procesados)} HOJA(S) GENERADAS",
                 style={"color": "var(--accent-green)", "marginBottom": "8px",
                        "fontFamily": "Share Tech Mono, monospace", "fontSize": "0.78rem"}),
        html.Div(f"LOTES: {', '.join(lotes_procesados)}",
                 style={"color": "var(--text-secondary)", "fontFamily": "Share Tech Mono, monospace",
                        "fontSize": "0.72rem"}),
    ]
    if errores:
        lines.append(
            html.Div(f"✗ ERRORES: {', '.join(errores)}",
                     style={"color": "var(--accent-red)", "marginTop": "8px",
                            "fontFamily": "Share Tech Mono, monospace", "fontSize": "0.72rem"})
        )
    return href, nombre_excel, html.Div(lines)


# =============================================================================
# CALLBACK SUPERVIVENCIA (TAB 2)
# =============================================================================
@app.callback(
    [Output('selector-hoja-wrapper', 'children'),
     Output('selector-hoja', 'options'),
     Output('output-alertas', 'children'),
     Output('output-data-upload', 'children'),
     Output('grafico-supervivencia', 'figure'),
     Output('grafico-talla-comercial', 'figure'),
     Output('grafico-ejes', 'figure'),
     Output('grafico-ocupacion', 'figure'),
     Output('grafico-altura', 'figure'),
     Output('grafico-porcentaje-col', 'figure')],
    [Input('upload-data', 'contents'),
     Input('selector-hoja', 'value')],
    [State('upload-data', 'filename')]
)
def procesar_archivo_con_hoja(contents, hoja_seleccionada, filename):
    empty_fig = {"layout": {**PLOT_LAYOUT, "title": {"text": "// Sin datos"}}}

    if contents is None:
        msg = html.Div(
            "⊘  Carga un archivo Excel para comenzar el análisis.",
            style={"fontFamily": "Share Tech Mono, monospace",
                   "color": "var(--text-dim)", "fontSize": "0.8rem",
                   "padding": "20px", "textAlign": "center",
                   "border": "1px dashed var(--border-subtle)",
                   "borderRadius": "2px"}
        )
        return "", [], msg, None, empty_fig, empty_fig, empty_fig, empty_fig, empty_fig, empty_fig

    content_type, content_string = contents.split(',')
    decoded = base64.b64decode(content_string)

    try:
        excel_file = pd.ExcelFile(BytesIO(decoded))
        hojas = excel_file.sheet_names
    except Exception as e:
        return "", [], html.Div(f"Error al leer archivo: {e}"), None, \
               empty_fig, empty_fig, empty_fig, empty_fig, empty_fig, empty_fig

    opciones = [{'label': h, 'value': h} for h in hojas]

    if hoja_seleccionada is None:
        return html.Div(), opciones, None, None, \
               empty_fig, empty_fig, empty_fig, empty_fig, empty_fig, empty_fig

    if hoja_seleccionada not in hojas:
        return html.Div(), opciones, \
               html.Div("Hoja no válida.", style={"color": "var(--accent-red)"}), None, \
               empty_fig, empty_fig, empty_fig, empty_fig, empty_fig, empty_fig

    try:
        df_raw = pd.read_excel(BytesIO(decoded), sheet_name=hoja_seleccionada, header=None)
    except Exception as e:
        return html.Div(), opciones, html.Div(f"Error: {e}"), None, \
               empty_fig, empty_fig, empty_fig, empty_fig, empty_fig, empty_fig

    header_row_idx = None
    for i in range(len(df_raw)):
        if df_raw.iloc[i, 0] == 'Fila':
            header_row_idx = i
            break

    if header_row_idx is None:
        return html.Div(), opciones, \
               html.Div("No se encontró el encabezado 'Fila'.",
                        style={"color": "var(--accent-amber)",
                               "fontFamily": "Share Tech Mono, monospace",
                               "fontSize": "0.78rem"}), None, \
               empty_fig, empty_fig, empty_fig, empty_fig, empty_fig, empty_fig

    df = pd.read_excel(BytesIO(decoded), sheet_name=hoja_seleccionada, header=header_row_idx)
    df['Fila_temp'] = df['Fila'].astype(str).str.strip()
    mask_fila_valida = df['Fila_temp'].str.match(r'^\d+(\.\d+)?$', na=False)
    df = df[mask_fila_valida].copy()
    df.drop(columns=['Fila_temp'], inplace=True)

    if df.empty:
        return html.Div(), opciones, \
               html.Div("Sin filas numéricas válidas.",
                        style={"color": "var(--accent-amber)"}), None, \
               empty_fig, empty_fig, empty_fig, empty_fig, empty_fig, empty_fig

    columnas_numericas = ['Máximo', 'Sobrevivencia', 'Talla Comercial', 'Ejes ≥ 2',
                          'Ocup sustrato ≥ 80%', 'Altura ≥ 12 cm']
    for col in columnas_numericas:
        df[col] = pd.to_numeric(df.get(col, 0), errors='coerce')

    df['% Col'] = pd.to_numeric(df.get('% Col', 0), errors='coerce')
    columnas_numericas.append('% Col')
    df[columnas_numericas] = df[columnas_numericas].fillna(0)
    df['Fila'] = pd.to_numeric(df['Fila'], errors='coerce').fillna(0).astype(int).astype(str)

    total_maximo = df['Máximo'].sum()
    if total_maximo == 0:
        return html.Div(), opciones, \
               html.Div("Total 'Máximo' = 0, sin datos para calcular.",
                        style={"color": "var(--accent-red)"}), None, \
               empty_fig, empty_fig, empty_fig, empty_fig, empty_fig, empty_fig

    # ── KPIs
    total_sobrevivencia = df['Sobrevivencia'].sum()
    tasa_sup   = (total_sobrevivencia / total_maximo) * 100
    tasa_tc    = (df['Talla Comercial'].sum() / total_maximo) * 100
    tasa_ejes  = (df['Ejes ≥ 2'].sum() / total_maximo) * 100
    tasa_ocup  = (df['Ocup sustrato ≥ 80%'].sum() / total_maximo) * 100
    tasa_alt   = (df['Altura ≥ 12 cm'].sum() / total_maximo) * 100
    tasa_col   = (df['% Col'].sum() / total_maximo) * 100 if df['% Col'].sum() > 0 else 0

    # ── Alarmas
    condiciones = (
        (df['Sobrevivencia'] > df['Máximo']) |
        (df['Talla Comercial'] > df['Máximo']) |
        (df['Ejes ≥ 2'] > df['Máximo']) |
        (df['Ocup sustrato ≥ 80%'] > df['Máximo']) |
        (df['Altura ≥ 12 cm'] > df['Máximo']) |
        (df['% Col'] > df['Máximo'])
    )
    filas_alerta = df[condiciones]

    if not filas_alerta.empty:
        alerta_ui = html.Div([
            html.Div(f"⚠  {len(filas_alerta)} FILAS CON VALORES FUERA DE RANGO",
                     style={"fontFamily": "Share Tech Mono, monospace",
                            "color": "var(--accent-amber)",
                            "fontSize": "0.78rem",
                            "marginBottom": "10px",
                            "borderLeft": "2px solid var(--accent-amber)",
                            "paddingLeft": "10px"}),
            dash_table.DataTable(
                data=filas_alerta.to_dict('records'),
                columns=[{'name': i, 'id': i} for i in filas_alerta.columns],
                style_table={'overflowX': 'auto', 'background': 'var(--bg-card)'},
                style_cell={
                    'textAlign': 'center', 'padding': '5px',
                    'fontSize': '11px', 'fontFamily': 'Share Tech Mono, monospace',
                    'backgroundColor': 'var(--bg-panel)', 'color': 'var(--text-secondary)',
                    'border': '1px solid var(--border-subtle)'
                },
                style_header={
                    'backgroundColor': 'var(--bg-void)', 'fontWeight': 'bold',
                    'color': 'var(--accent-cyan)', 'border': '1px solid var(--border-glow)'
                },
                page_size=8
            )
        ])
    else:
        alerta_ui = html.Div(
            "✓  Sin alarmas detectadas — todos los valores dentro del rango esperado.",
            style={"fontFamily": "Share Tech Mono, monospace",
                   "color": "var(--accent-green)",
                   "fontSize": "0.75rem",
                   "borderLeft": "2px solid var(--accent-green)",
                   "paddingLeft": "10px",
                   "padding": "8px 12px",
                   "background": "rgba(57,255,138,0.04)"}
        )

    # ── Metadata
    try:
        meta_df = pd.read_excel(BytesIO(decoded), sheet_name=hoja_seleccionada, header=None)
        fecha_m = meta_df.iloc[5, 5] if meta_df.shape[0] > 5 and meta_df.shape[1] > 5 else "—"
        lote_m  = meta_df.iloc[7, 2] if meta_df.shape[0] > 7 and meta_df.shape[1] > 2 else "—"
        if isinstance(fecha_m, (int, float)):
            fecha_m = (pd.to_datetime("1899-12-30") + pd.to_timedelta(int(fecha_m), unit="D")).strftime('%d-%m-%Y')
        elif hasattr(fecha_m, 'strftime'):
            fecha_m = fecha_m.strftime('%d-%m-%Y')
        else:
            fecha_m = str(fecha_m)
    except Exception:
        fecha_m = "—"
        lote_m  = "—"

    kpi_cards = html.Div([
        html.Div([
            html.Div(f"{tasa_sup:.1f}%".replace('.', ','), className="val"),
            html.Div("SUPERVIVENCIA", className="lbl"),
        ], className="summary-stat"),
        html.Div([
            html.Div(f"{tasa_tc:.1f}%".replace('.', ','), className="val",
                     style={"color": "var(--accent-green)", "textShadow": "var(--glow-green)"}),
            html.Div("TALLA COMERCIAL", className="lbl"),
        ], className="summary-stat"),
        html.Div([
            html.Div(f"{int(total_maximo):,}".replace(",", "."), className="val",
                     style={"fontSize": "1.4rem"}),
            html.Div("MACETAS MUESTREADAS", className="lbl"),
        ], className="summary-stat"),
        html.Div([
            html.Div(str(lote_m), className="val",
                     style={"fontSize": "1rem", "letterSpacing": "0.05em"}),
            html.Div("LOTE", className="lbl"),
        ], className="summary-stat"),
        html.Div([
            html.Div(str(fecha_m), className="val",
                     style={"fontSize": "1rem", "letterSpacing": "0.05em"}),
            html.Div("FECHA MUESTREO", className="lbl"),
        ], className="summary-stat"),
    ], style={
        "display": "flex",
        "flexWrap": "wrap",
        "background": "var(--bg-card)",
        "border": "1px solid var(--border-subtle)",
        "marginBottom": "20px",
    })

    # ── Data table
    tabla_ui = html.Div([
        html.Div("// DATOS CRUDOS", className="section-label", style={"marginTop": "16px"}),
        dash_table.DataTable(
            data=df.to_dict('records'),
            columns=[{'name': i, 'id': i} for i in df.columns],
            style_table={'overflowX': 'auto'},
            style_cell={
                'textAlign': 'center', 'padding': '5px',
                'fontSize': '11px', 'fontFamily': 'Share Tech Mono, monospace',
                'backgroundColor': 'var(--bg-panel)', 'color': 'var(--text-secondary)',
                'border': '1px solid var(--border-subtle)'
            },
            style_header={
                'backgroundColor': 'var(--bg-void)', 'fontWeight': 'bold',
                'color': 'var(--accent-cyan)', 'border': '1px solid var(--border-glow)',
                'letterSpacing': '0.08em', 'fontSize': '0.7rem'
            },
            page_size=10
        )
    ])

    resumen_ui = html.Div([kpi_cards, tabla_ui])

    filas_unicas = df['Fila'].tolist()

    def crear_grafico(col_y, titulo, color_hex, label_y):
        if col_y not in df.columns:
            return {"layout": {**PLOT_LAYOUT, "title": {"text": f"// {titulo} — no disponible"}}}
        fig = px.bar(df, x='Fila', y=col_y,
                     labels={'Fila': 'Fila', col_y: label_y},
                     color_discrete_sequence=[color_hex])
        fig.update_traces(
            text=df[col_y], textposition='outside',
            marker_line_color=color_hex,
            marker_line_width=0.5,
            marker_color=color_hex,
            opacity=0.85,
        )
        layout = dict(PLOT_LAYOUT)
        layout["title"] = {"text": titulo, "font": {"family": "Rajdhani, sans-serif",
                                                     "color": color_hex, "size": 13}}
        layout["xaxis"] = dict(PLOT_LAYOUT["xaxis"],
                                tickmode='array', tickvals=filas_unicas,
                                ticktext=filas_unicas, tickangle=-45)
        layout["yaxis"] = dict(PLOT_LAYOUT["yaxis"], title=label_y)
        fig.update_layout(**layout)
        return fig

    c = CHART_COLORS
    fig_sup  = crear_grafico('Sobrevivencia',        f'Supervivencia  {tasa_sup:.1f}%',  c[0], 'Plantas vivas')
    fig_tc   = crear_grafico('Talla Comercial',      f'Talla Comercial  {tasa_tc:.1f}%', c[1], 'Talla comercial')
    fig_ej   = crear_grafico('Ejes ≥ 2',             f'Ejes ≥ 2  {tasa_ejes:.1f}%',      c[2], 'Con ejes ≥ 2')
    fig_oc   = crear_grafico('Ocup sustrato ≥ 80%',  f'Ocup. Sustrato  {tasa_ocup:.1f}%',c[3], 'Ocup ≥ 80%')
    fig_alt  = crear_grafico('Altura ≥ 12 cm',       f'Altura ≥ 12 cm  {tasa_alt:.1f}%', c[4], 'Alt ≥ 12 cm')

    if df['% Col'].sum() > 0:
        fig_col = crear_grafico('% Col', f'% Col  {tasa_col:.1f}%', c[5], '% Col')
    else:
        fig_col = {"layout": {**PLOT_LAYOUT, "title": {"text": "// % Col no disponible"}}}

    return html.Div(), opciones, alerta_ui, resumen_ui, fig_sup, fig_tc, fig_ej, fig_oc, fig_alt, fig_col


# =============================================================================
# EJECUCIÓN
# =============================================================================
if __name__ == "__main__":
    app.run(host='0.0.0.0', port=8050, debug=True)