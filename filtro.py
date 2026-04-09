"""
Script robusto para extraer muestreos activos desde BioTecnos.
- Usa Selenium estándar + webdriver-manager (misma configuración que lotes).
- Driver persistente con manejo de errores de conexión.
Genera output/Muestreos_Activos.xlsx con hojas: Hoy, Proximos, SemanaActual.
"""

import os
import re
import time
import logging
import traceback
from datetime import datetime, timedelta
import pandas as pd
import numpy as np
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import TimeoutException, NoSuchElementException, WebDriverException
from webdriver_manager.chrome import ChromeDriverManager

# ================= CONFIGURACIÓN =================
HOY = pd.Timestamp.now().normalize()
INI = HOY - timedelta(days=7)
FIN = HOY + timedelta(days=30)

USERNAME = "bryan"
PASSWORD = "Synergia333"
LOGIN_URL = "https://sisbiotecnos.cl/login"

HEADLESS = True  # Cambiar a False si quieres ver el navegador
MAX_RETRIES = 3
RETRY_DELAY = 5

# Rutas
DATA_DIR = "data"
OUTPUT_DIR = "output"
os.makedirs(DATA_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

ARCHIVO_BIOTECNOS = os.path.join(DATA_DIR, "BioTecnos.xlsx")

# Configurar logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(os.path.join(OUTPUT_DIR, "muestreos.log"), encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# ================= FUNCIONES DE DRIVER (estilo lotes) =================
def crear_driver():
    """Crea un driver de Chrome con la misma configuración que funcionó en lotes."""
    for intento in range(1, MAX_RETRIES + 1):
        try:
            options = webdriver.ChromeOptions()
            if HEADLESS:
                options.add_argument('--headless=new')
            # Opciones esenciales para evitar problemas de conexión
            options.add_argument('--no-sandbox')
            options.add_argument('--disable-dev-shm-usage')
            options.add_argument('--disable-gpu')
            options.add_argument('--window-size=1920,1080')
            options.add_argument('--disable-blink-features=AutomationControlled')
            options.add_argument('--log-level=3')
            # Evitar que Chrome cierre conexiones inesperadamente
            options.add_argument('--disable-features=VizDisplayCompositor')
            options.add_experimental_option('excludeSwitches', ['enable-logging'])
            
            service = Service(ChromeDriverManager().install())
            driver = webdriver.Chrome(service=service, options=options)
            driver.set_page_load_timeout(30)
            # Pequeña pausa para estabilizar la conexión
            time.sleep(1)
            logger.info("✅ Driver creado exitosamente.")
            return driver
        except Exception as e:
            logger.error(f"Intento {intento} falló al crear driver: {e}")
            if intento < MAX_RETRIES:
                time.sleep(RETRY_DELAY)
            else:
                raise

def iniciar_sesion(driver):
    """Inicia sesión en BioTecnos con manejo de errores de conexión."""
    for intento in range(1, MAX_RETRIES + 1):
        try:
            logger.info("🔐 Iniciando sesión...")
            driver.get(LOGIN_URL)
            WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.NAME, "login")))
            driver.find_element(By.NAME, "login").send_keys(USERNAME)
            driver.find_element(By.NAME, "pass").send_keys(PASSWORD)
            driver.find_element(By.XPATH, "//button[@type='submit']").click()
            WebDriverWait(driver, 15).until(EC.url_changes(LOGIN_URL))
            logger.info("✅ Sesión iniciada correctamente")
            time.sleep(2)
            return
        except WebDriverException as e:
            logger.error(f"Intento {intento} - Error de conexión: {e}")
            # Si el driver murió, lo recreamos
            if "invalid session id" in str(e) or "connection refused" in str(e).lower():
                logger.info("🔄 El driver perdió la conexión. Recreando...")
                driver.quit()
                driver = crear_driver()  # Nota: esto no afecta al driver externo, pero podemos manejarlo
                # Para que el cambio persista, necesitamos retornar el nuevo driver
                # Como es complejo, mejor lanzamos excepción para que main lo maneje
                raise Exception("Conexión perdida, reiniciar driver")
            if intento == MAX_RETRIES:
                raise
            time.sleep(RETRY_DELAY)
        except Exception as e:
            logger.error(f"Intento {intento} de login falló: {e}")
            if intento == MAX_RETRIES:
                raise
            time.sleep(RETRY_DELAY)

# ================= FUNCIONES AUXILIARES DE NEGOCIO =================
# (mantener igual que antes, sin cambios)
def sumar_dias_habil(fecha, dias):
    if pd.isnull(fecha):
        return None
    nueva = fecha + timedelta(days=dias)
    if nueva.weekday() == 5:
        return nueva - timedelta(days=1)
    elif nueva.weekday() == 6:
        return nueva + timedelta(days=1)
    return nueva

def ajustar_a_habil(fecha):
    if pd.isnull(fecha):
        return None
    if fecha.weekday() == 5:
        return fecha - timedelta(days=1)
    if fecha.weekday() == 6:
        return fecha + timedelta(days=1)
    return fecha

def limpiar_macetero(texto):
    if pd.isna(texto) or not isinstance(texto, str):
        return None
    match = re.search(r'(\d+)\s*mL', texto)
    if match:
        litros = int(match.group(1)) / 1000
        return f"{litros:.5f}".rstrip('0').rstrip('.') + " L"
    return texto

def extraer_cantidad_bandejas(texto):
    if pd.isna(texto) or not isinstance(texto, str):
        return None
    match_bandeja = re.search(r'Bandeja de (\d+)', texto)
    if match_bandeja:
        return int(match_bandeja.group(1))
    if re.search(r'(\d+)\s*mL', texto):
        return 1
    return None

def extraer_alveolos(texto):
    if pd.isna(texto):
        return None
    match = re.search(r'C(\d+(?:\.\d+)?)', texto)
    if match:
        return int(match.group(1).replace('.', ''))
    return None

def limpiar_macetas_miles(valor):
    try:
        if pd.isna(valor):
            return None
        valor_str = str(valor).replace(".", "").replace(",", "")
        return int(valor_str)
    except:
        return None

def expandir_por_bandeja(df_muestreos):
    if 'Bandejas' not in df_muestreos.columns:
        return df_muestreos
    df_exp = df_muestreos.copy()
    df_exp['Bandejas'] = df_exp['Bandejas'].astype(str)
    mask = df_exp['Bandejas'].str.contains(',')
    if mask.any():
        df_exp = pd.concat([
            df_exp[~mask],
            df_exp[mask].assign(Bandejas=df_exp[mask]['Bandejas'].str.split(',')).explode('Bandejas')
        ], ignore_index=True)
        df_exp['Bandejas'] = df_exp['Bandejas'].str.strip()
    df_exp['ID'] = df_exp['ID'].astype(str) + "_B" + df_exp['Bandejas'].astype(str)
    return df_exp

# ================= WEB SCRAPING CON DRIVER PERSISTENTE =================
def extraer_tabla_2_para_ids(driver, ids):
    tablas = []
    estados = []
    total = len(ids)
    for idx, id_lote in enumerate(ids, 1):
        estado = "Entregado"
        try:
            num_id = int(str(id_lote).split('_')[0])
            url = f"https://sisbiotecnos.cl/ex_vitro/lote_maceta/ver/{num_id}"
            driver.get(url)
            # Verificar redirección a login
            if "login" in driver.current_url:
                logger.warning(f"ID {id_lote}: sesión expirada, reintentando login...")
                iniciar_sesion(driver)
                driver.get(url)
                if "login" in driver.current_url:
                    estado = "Error (login fallido)"
                else:
                    estado = "Activo"
            WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.TAG_NAME, "table")))
            tablas_html = driver.find_elements(By.TAG_NAME, "table")
            if len(tablas_html) >= 2:
                tabla = tablas_html[1]
                filas = tabla.find_elements(By.TAG_NAME, "tr")
                if len(filas) >= 2:
                    data = []
                    for fila in filas:
                        celdas = fila.find_elements(By.TAG_NAME, "td")
                        if celdas:
                            data.append([celda.text.strip() for celda in celdas])
                    data = [fila for fila in data if any(fila)]
                    columnas = ["Tipo Contenedor", "Sector", "Invernadero", "Mesón",
                                "Macetas", "% Macetas", "Plantas plantables", "Plantas comerciales"]
                    cuerpo = [fila for fila in data if len(fila) == len(columnas)]
                    if cuerpo:
                        df_tabla = pd.DataFrame(cuerpo, columns=columnas)
                        df_tabla.insert(0, "ID", id_lote)
                        tablas.append(df_tabla)
                        estado = "Activo"
                    else:
                        estado = "Sin datos en tabla"
                else:
                    estado = "Tabla vacía"
            else:
                estado = "No hay segunda tabla"
        except Exception as e:
            logger.error(f"Error en ID {id_lote} ({idx}/{total}): {e}")
            estado = f"Error: {str(e)[:50]}"
        estados.append({"ID": id_lote, "Estado del Lote": estado})
        time.sleep(0.5)
        if idx % 10 == 0:
            logger.info(f"   Procesados {idx}/{total} IDs...")
    df_tabla2 = pd.concat(tablas, ignore_index=True) if tablas else pd.DataFrame()
    df_estados = pd.DataFrame(estados)
    return df_tabla2, df_estados

def procesar_df_tabla2(df_tabla2):
    if df_tabla2.empty:
        return pd.DataFrame()
    df = df_tabla2.copy()
    df["Macetero"] = df["Tipo Contenedor"].apply(limpiar_macetero)
    df["Invernadero"] = df["Invernadero"].apply(lambda x: x.split()[-1] if isinstance(x, str) else x)
    df["Bandejas"] = df["Tipo Contenedor"].apply(extraer_cantidad_bandejas)
    df["Mesón"] = df["Mesón"].astype(str).str.replace("M", "", regex=False).str.zfill(2)
    df["Macetas"] = df["Macetas"].apply(limpiar_macetas_miles)
    df["I-M-C"] = "I" + df["Invernadero"].astype(str) + "-M" + df["Mesón"] + "-C" + df["Macetas"].astype(str)
    df["I-M-C-V-B"] = df["I-M-C"] + "-V" + df["Macetero"].astype(str) + "-B" + df["Bandejas"].astype(str)
    df["Alveolos"] = df["I-M-C"].apply(extraer_alveolos)
    df_resumen = df.groupby(["ID", "Bandejas", "Macetero"]).agg({
        "Macetas": "sum",
        "Alveolos": "sum",
        "I-M-C": lambda x: ", ".join(dict.fromkeys(x.dropna())),
        "I-M-C-V-B": lambda x: ", ".join(dict.fromkeys(x.dropna()))
    }).reset_index()
    df_resumen["Alveolos"] = pd.to_numeric(df_resumen["Alveolos"], errors="coerce")
    df_resumen["Macetas actuales"] = pd.to_numeric(df_resumen["Macetas"], errors="coerce")
    df_resumen["Combinación"] = df_resumen["Bandejas"].astype(str) + " | " + df_resumen["Macetero"].astype(str)
    df_resumen.rename(columns={"Bandejas": "Bandeja"}, inplace=True)
    return df_resumen

# ================= PROCESAMIENTO PRINCIPAL =================
def main():
    if not os.path.exists(ARCHIVO_BIOTECNOS):
        raise FileNotFoundError(f"No se encontró '{ARCHIVO_BIOTECNOS}'. Ejecuta primero el script de lotes.")
    logger.info(f"📂 Leyendo archivo: {ARCHIVO_BIOTECNOS}")
    df = pd.read_excel(ARCHIVO_BIOTECNOS, header=0)
    if "Fecha" not in df.columns:
        raise ValueError("La columna 'Fecha' no existe en el archivo.")
    df["Fecha"] = pd.to_datetime(df["Fecha"], errors='coerce', dayfirst=True)
    for d in [30, 60, 120, 180]:
        df[f'{d} Días'] = df["Fecha"].apply(lambda x: sumar_dias_habil(x, d))
    def determinar_muestreo(row):
        for col in ['30 Días', '60 Días', '120 Días', '180 Días']:
            if pd.notnull(row[col]) and (INI <= row[col] <= FIN):
                return col, row[col]
        return None, None
    df[['muestreo_activador', 'fecha_activadora']] = df.apply(lambda r: pd.Series(determinar_muestreo(r)), axis=1)
    activos = df[df['fecha_activadora'].notna()].copy()
    if activos.empty:
        logger.warning("⚠️ No hay muestreos activos en la ventana de tiempo.")
        return
    muestreos_hoy = activos[activos['fecha_activadora'] == HOY].copy()
    muestreos_proximos = activos[activos['fecha_activadora'] != HOY].copy()
    muestreos_hoy = expandir_por_bandeja(muestreos_hoy)
    muestreos_proximos = expandir_por_bandeja(muestreos_proximos)
    ids_hoy = muestreos_hoy["ID"].dropna().unique()
    ids_proximos = muestreos_proximos["ID"].dropna().unique()
    ids_total = pd.unique(np.concatenate([ids_hoy, ids_proximos]))
    logger.info(f"📋 Total de IDs únicos a procesar: {len(ids_total)}")
    
    logger.info("🚀 Iniciando driver...")
    driver = None
    try:
        driver = crear_driver()
        iniciar_sesion(driver)
        logger.info("🔍 Extrayendo datos de producción para cada ID...")
        df_tabla2, df_estados = extraer_tabla_2_para_ids(driver, ids_total)
        df_resumen = procesar_df_tabla2(df_tabla2)
    except Exception as e:
        logger.error(f"❌ Error fatal durante el scraping: {e}")
        logger.error(traceback.format_exc())
        if driver:
            try:
                driver.save_screenshot(os.path.join(OUTPUT_DIR, "error_scraping_muestreos.png"))
                logger.info("📸 Captura guardada como 'error_scraping_muestreos.png'")
            except:
                pass
        raise
    finally:
        if driver:
            driver.quit()
            logger.info("🚪 Driver cerrado.")
    
    # Merge y procesamiento posterior (igual que antes)
    for df_temp in [muestreos_hoy, muestreos_proximos]:
        for col in ['Macetas act.', 'Macetas actuales', 'I-M-C', 'Bandeja', 'Alveolos', 'Combinación']:
            if col in df_temp.columns:
                df_temp.drop(columns=[col], inplace=True)
    muestreos_proximos = muestreos_proximos.merge(df_estados, on="ID", how="left")
    muestreos_proximos = muestreos_proximos.merge(df_resumen, on="ID", how="left")
    muestreos_proximos = muestreos_proximos[muestreos_proximos["Estado del Lote"] == "Activo"]
    muestreos_hoy = muestreos_hoy.merge(df_estados, on="ID", how="left")
    muestreos_hoy = muestreos_hoy.merge(df_resumen, on="ID", how="left")
    muestreos_hoy = muestreos_hoy[muestreos_hoy["Estado del Lote"] == "Activo"]
    muestreos_proximos["fecha_activadora"] = muestreos_proximos["fecha_activadora"].apply(ajustar_a_habil)
    muestreos_hoy["fecha_activadora"] = muestreos_hoy["fecha_activadora"].apply(ajustar_a_habil)
    
    def formatear_codigo_excel(row):
        codigo_base = str(row["Código"])
        bandeja = str(row.get("Bandeja", ""))
        macetero = str(row.get("Macetero", "")).strip().replace(" ", "").replace(",", ".")
        try:
            bandeja = str(int(float(bandeja)))
        except:
            bandeja = bandeja.strip()
        if not bandeja.endswith("B"):
            bandeja += "B"
        if not macetero.endswith("L"):
            macetero = macetero.replace("L", "") + "L"
        return f"{codigo_base}-{bandeja}-{macetero}"
    for df_m in [muestreos_hoy, muestreos_proximos]:
        if all(col in df_m.columns for col in ['Código', 'Bandeja', 'Macetero']):
            df_m['Código'] = df_m.apply(formatear_codigo_excel, axis=1)
        else:
            logger.warning("⚠️ No se pudo reconstruir Código: faltan columnas")
    for df_m in [muestreos_hoy, muestreos_proximos]:
        if 'Macetas actuales' in df_m.columns:
            df_m['Macetas actuales'] = pd.to_numeric(df_m['Macetas actuales'], errors='coerce')
    
    muestreos_totales = pd.concat([muestreos_hoy, muestreos_proximos], ignore_index=True)
    hoy_ts = pd.Timestamp.now().normalize()
    inicio_semana = hoy_ts - timedelta(days=hoy_ts.weekday())
    fin_semana = inicio_semana + timedelta(days=6)
    muestreos_semana = muestreos_totales[muestreos_totales['fecha_activadora'].between(inicio_semana, fin_semana)].copy()
    if not muestreos_semana.empty:
        dias_es = ['Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes', 'Sábado', 'Domingo']
        muestreos_semana['Día de la semana'] = muestreos_semana['fecha_activadora'].dt.weekday.map(lambda x: dias_es[x])
        muestreos_semana = muestreos_semana.sort_values('fecha_activadora')
        cols = ['Día de la semana'] + [c for c in muestreos_semana.columns if c != 'Día de la semana']
        muestreos_semana = muestreos_semana[cols]
    else:
        muestreos_semana = pd.DataFrame(columns=['Día de la semana'] + muestreos_totales.columns.tolist())
    
    salida_path = os.path.join(OUTPUT_DIR, "Muestreos_Activos.xlsx")
    with pd.ExcelWriter(salida_path) as writer:
        muestreos_hoy.to_excel(writer, sheet_name="Hoy", index=False)
        muestreos_proximos.to_excel(writer, sheet_name="Proximos", index=False)
        muestreos_semana.to_excel(writer, sheet_name="SemanaActual", index=False)
    logger.info(f"\n✅ Reporte generado exitosamente.")
    logger.info(f"   📅 HOY: {len(muestreos_hoy)} filas")
    logger.info(f"   📅 PRÓXIMOS: {len(muestreos_proximos)} filas")
    logger.info(f"   📅 SEMANA ACTUAL: {len(muestreos_semana)} filas")
    logger.info(f"   💾 Archivo: {salida_path}")

if __name__ == "__main__":
    main()