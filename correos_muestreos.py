import os
import re
import hashlib
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
import numpy as np
import pandas as pd
import xlsxwriter
from datetime import datetime
import warnings

warnings.filterwarnings("ignore")

# =============================================================================
# CONFIGURACIÓN DE RUTAS
# =============================================================================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.path.join(BASE_DIR, "output")
IMAGES_DIR = os.path.join(BASE_DIR, "images")
MUESTREO_DIR = os.path.join(BASE_DIR, "muestreo")
os.makedirs(MUESTREO_DIR, exist_ok=True)

ARCHIVO_EXCEL = os.path.join(OUTPUT_DIR, "Muestreos_Activos.xlsx")
IMAGEN_RECORTADA = os.path.join(IMAGES_DIR, "imagen_recortada.jpg")

# =============================================================================
# CONFIGURACIÓN DE CORREO
# =============================================================================
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587
CORREO_REMITENTE = "bryan.vielma@synergiabio.com"
# Lista de destinatarios (se envía a todos en el campo 'Para')
DESTINATARIOS = [
    "karel.rodriguez@synergiabio.com",
    "francisco.lincir@synergiabio.com",
    "milena.gonzalez@synergiabio.com",
    "briseys.calvo@synergiabio.com",
    "elizabeth.garcia@synergiabio.com",
    "eurekadatanalytics@gmail.com",
]
CORREO_DESTINATARIO = ", ".join(DESTINATARIOS)
CONTRASENA = "zxvv ijfx dojz acqm"

# =============================================================================
# CARGA DE DATOS
# =============================================================================
if not os.path.exists(ARCHIVO_EXCEL):
    raise FileNotFoundError(f"No se encontró {ARCHIVO_EXCEL}.")

muestreos_hoy = pd.read_excel(ARCHIVO_EXCEL, sheet_name="Hoy")
for col in ['Macetas actuales', 'Alveolos', 'Bandeja']:
    if col in muestreos_hoy.columns:
        muestreos_hoy[col] = pd.to_numeric(muestreos_hoy[col], errors='coerce')
if 'fecha_activadora' in muestreos_hoy.columns:
    muestreos_hoy['fecha_activadora'] = pd.to_datetime(muestreos_hoy['fecha_activadora'], errors='coerce')

# =============================================================================
# FUNCIONES AUXILIARES
# =============================================================================
def extraer_cantidad_desde_imc(imc_str):
    if not isinstance(imc_str, str):
        return 0
    numeros = re.findall(r'-C(\d+)', imc_str)
    return sum(int(n) for n in numeros) if numeros else 0

def calcular_tamano_muestra(cantidad):
    rangos = [(8,2),(15,3),(25,5),(50,8),(90,13),(150,20),(280,40),
              (500,60),(1200,80),(3200,140),(10000,200),(35000,320),
              (150000,500),(500000,800)]
    for limite, muestra in rangos:
        if cantidad <= limite:
            return muestra
    return 1260

def generar_datos_lote(lote):
    """Prepara todos los datos necesarios para una hoja."""
    imc_raw = lote.get("I-M-C", "")
    if pd.isna(imc_raw):
        imc_raw = ""

    cantidad = extraer_cantidad_desde_imc(str(imc_raw))
    if cantidad == 0:
        cantidad = lote.get("Macetas actuales", 0)
        cantidad = int(cantidad) if not pd.isna(cantidad) else 0

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
    muestra_tamano = calcular_tamano_muestra(cantidad)
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
    """Escribe una hoja con el formato idéntico al de la app Dash."""
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
# GENERAR EXCEL CON MÚLTIPLES HOJAS (SOLO LOTES SIN MN)
# =============================================================================
def generar_excel_multiple(lotes_a_procesar):
    """
    Genera un archivo Excel con una hoja por cada lote en la lista.
    lotes_a_procesar: lista de Series (filas del DataFrame) que NO contienen 'MN'
    """
    if not lotes_a_procesar:
        print("⚠️ No hay lotes para procesar (todos tienen MN o ninguno con muestreo).")
        return None

    hoy = datetime.now().strftime("%d-%m-%Y")
    nombre_excel = f"MUESTREO_MACRO_{hoy}.xlsx"
    ruta_excel = os.path.join(MUESTREO_DIR, nombre_excel)

    lotes_datos = []
    for lote in lotes_a_procesar:
        try:
            datos_lote = generar_datos_lote(lote)
            lotes_datos.append(datos_lote)
        except Exception as e:
            print(f"❌ Error en lote {lote.get('Código')}: {e}")

    if not lotes_datos:
        print("⚠️ No se pudo generar datos para ningún lote.")
        return None

    workbook = xlsxwriter.Workbook(ruta_excel)
    for datos in lotes_datos:
        nombre = str(datos["lote"]["Código"])[:31]
        nombre = re.sub(r'[\\/*?:\[\]]', '_', nombre)
        escribir_hoja(workbook, datos, nombre)
    workbook.close()
    print(f"✅ Excel generado: {ruta_excel}")
    return ruta_excel

# =============================================================================
# ENVÍO DE CORREO
# =============================================================================
def enviar_correo(ruta_archivo, nombre_archivo, asunto, cuerpo):
    if not os.path.exists(ruta_archivo):
        return False

    msg = MIMEMultipart()
    msg['From'] = CORREO_REMITENTE
    msg['To'] = CORREO_DESTINATARIO
    msg['Subject'] = asunto
    msg.attach(MIMEText(cuerpo, 'plain'))

    with open(ruta_archivo, 'rb') as f:
        part = MIMEBase('application', 'vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        part.set_payload(f.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f'attachment; filename="{nombre_archivo}"')
        msg.attach(part)

    try:
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(CORREO_REMITENTE, CONTRASENA)
            server.send_message(msg)
        print(f"📧 Correo enviado a {CORREO_DESTINATARIO}")
        return True
    except Exception as e:
        print(f"❌ Error al enviar: {e}")
        return False

# =============================================================================
# MAIN
# =============================================================================
if __name__ == "__main__":
    fecha_hoy = datetime.now().strftime("%d-%m-%Y")

    # Clasificar lotes según presencia de "MN" en I-M-C
    lotes_sin_mn = []   # SynergiaBio Chile (se incluyen en Excel)
    lotes_con_mn = []   # Vivero los viñedos (se excluyen)

    for _, lote in muestreos_hoy.iterrows():
        imc_raw = lote.get("I-M-C", "")
        if isinstance(imc_raw, str) and "MN" in imc_raw.upper():
            lotes_con_mn.append(lote)
        else:
            lotes_sin_mn.append(lote)

    # Generar Excel solo con los lotes sin MN
    ruta_excel = generar_excel_multiple(lotes_sin_mn)

    if ruta_excel:
        asunto = f"MUESTREO - MACRO {fecha_hoy}"

        ids_todos = [str(lote["ID"]) for lote in lotes_sin_mn + lotes_con_mn]
        ids_synergia = [str(lote["ID"]) for lote in lotes_sin_mn]
        ids_vivero = [str(lote["ID"]) for lote in lotes_con_mn]

        cuerpo = f"Adjunto el archivo con los muestreos del día {fecha_hoy}.\n\n"
        cuerpo += f"ID lotes con muestreo hoy: {', '.join(ids_todos)} ({len(ids_todos)} en total)\n"
        cuerpo += f"ID Lotes SynergiaBio Chile: {', '.join(ids_synergia)}\n"
        cuerpo += f"ID Lotes vivero los viñedos: {', '.join(ids_vivero)}\n"
        cuerpo += "Lotes que están en vivero los viñedos NO se incluyen en el muestreo."

        enviar_correo(ruta_excel, os.path.basename(ruta_excel), asunto, cuerpo)
    else:
        print("No se generó Excel. No se envía correo.")