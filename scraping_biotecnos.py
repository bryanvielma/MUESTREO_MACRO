"""
Script robusto para extraer lotes en maceta desde BioTecnos.
Rango de fechas: hoy y 8 meses atrás.
Guarda el resultado en r'C:\SYNERGIABIO\APP_MACRO\data\BioTecnos.xlsx'
VERSIÓN FINAL: sin backups, con reintentos y logging.
"""

import os
import time
import logging
import traceback
from datetime import datetime
from dateutil.relativedelta import relativedelta
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from webdriver_manager.chrome import ChromeDriverManager

# ================= CONFIGURACIÓN =================
USERNAME = "bryan"
PASSWORD = "Synergia333"
LOGIN_URL = "https://sisbiotecnos.cl/login"
BASE_URL = "https://sisbiotecnos.cl/ex_vitro/lote_maceta/index"
ESTADO = 2  # 2 = En inventario
HEADLESS = True  # Cambiar a False para ver el navegador
MAX_RETRIES = 3
RETRY_DELAY = 5

# Rutas
RUTA_DESTINO = r"C:\SYNERGIABIO\APP_MACRO\data"
os.makedirs(RUTA_DESTINO, exist_ok=True)
RUTA_COMPLETA = os.path.join(RUTA_DESTINO, "BioTecnos.xlsx")

# Columnas esperadas
NOMBRES_COLUMNAS = [
    'ID', 'Código', 'Fecha', 'Especie', 'Variedad',
    'LMC Dominante', 'Macetas actuales', 'Reagrupado', 'Estado'
]

# Configurar logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(os.path.join(RUTA_DESTINO, "scraping.log"), encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# ================= FUNCIONES =================
def calcular_fechas_dinamicas(meses_atras=8):
    hoy = datetime.now().date()
    fecha_inicio = hoy - relativedelta(months=meses_atras)
    return fecha_inicio.strftime("%Y-%m-%d"), hoy.strftime("%Y-%m-%d")

def crear_driver():
    for intento in range(1, MAX_RETRIES + 1):
        try:
            options = webdriver.ChromeOptions()
            if HEADLESS:
                options.add_argument('--headless=new')
            options.add_argument('--no-sandbox')
            options.add_argument('--disable-dev-shm-usage')
            options.add_argument('--disable-gpu')
            options.add_argument('--window-size=1920,1080')
            options.add_argument('--disable-blink-features=AutomationControlled')
            options.add_argument('--log-level=3')
            service = Service(ChromeDriverManager().install())
            driver = webdriver.Chrome(service=service, options=options)
            driver.set_page_load_timeout(30)
            logger.info("✅ Driver creado exitosamente.")
            return driver
        except Exception as e:
            logger.error(f"Intento {intento} falló: {e}")
            if intento < MAX_RETRIES:
                time.sleep(RETRY_DELAY)
            else:
                raise

def iniciar_sesion(driver):
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
        except Exception as e:
            logger.error(f"Intento {intento} de login falló: {e}")
            if intento == MAX_RETRIES:
                raise
            time.sleep(RETRY_DELAY)

def aplicar_filtros(driver):
    try:
        btn = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Aplicar filtros')]"))
        )
        logger.info("🔘 Haciendo clic en 'Aplicar filtros'...")
        btn.click()
        time.sleep(2)
        WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.ID, "tabla_lotes_maceta")))
    except TimeoutException:
        logger.warning("⚠️ Botón 'Aplicar filtros' no encontrado.")

def extraer_filas_tabla(driver):
    try:
        WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.ID, "tabla_lotes_maceta")))
        tabla = driver.find_element(By.ID, "tabla_lotes_maceta")
        filas = tabla.find_elements(By.CSS_SELECTOR, "tbody tr")
        datos = []
        for fila in filas:
            celdas = fila.find_elements(By.TAG_NAME, "td")
            if not celdas:
                continue
            fila_texto = [celda.text.strip().replace("Ver", "").strip() for celda in celdas]
            datos.append(fila_texto[:9] if len(fila_texto) >= 9 else fila_texto)
        return datos
    except Exception as e:
        logger.error(f"Error extrayendo filas: {e}")
        return []

def extraer_todas_paginas(driver, fecha_inicio, fecha_fin):
    url = f"{BASE_URL}?fecha_inicio={fecha_inicio}&fecha_termino={fecha_fin}&especie=*&variedad=*&estado={ESTADO}"
    logger.info(f"🌐 Navegando a: {url}")
    driver.get(url)
    time.sleep(3)

    aplicar_filtros(driver)

    todas_filas = []
    pagina = 1

    while True:
        logger.info(f"📄 Procesando página {pagina}...")
        datos_pagina = extraer_filas_tabla(driver)

        if datos_pagina:
            df_pagina = pd.DataFrame(datos_pagina)
            todas_filas.append(df_pagina)
            logger.info(f"   ✅ {len(datos_pagina)} filas extraídas")
        else:
            logger.warning(f"   ⚠️ Sin filas en página {pagina}")

        # Paginación simple
        try:
            siguiente = driver.find_element(By.XPATH, "//a[contains(text(), 'Siguiente')]")
            li_padre = siguiente.find_element(By.XPATH, "..")
            if "disabled" in li_padre.get_attribute("class"):
                logger.info("📌 Última página alcanzada.")
                break
            siguiente.click()
            time.sleep(2)
            pagina += 1
        except NoSuchElementException:
            logger.info("📌 No hay botón 'Siguiente'. Fin de paginación.")
            break
        except Exception as e:
            logger.error(f"Error en paginación: {e}")
            break

    if todas_filas:
        df_final = pd.concat(todas_filas, ignore_index=True)
        df_final.columns = NOMBRES_COLUMNAS[:df_final.shape[1]]
        logger.info(f"\n🎉 Total de registros extraídos: {len(df_final)}")
        return df_final
    else:
        return pd.DataFrame()

# ================= EJECUCIÓN =================
if __name__ == "__main__":
    fecha_inicio, fecha_fin = calcular_fechas_dinamicas(8)
    logger.info(f"📅 Rango: {fecha_inicio} → {fecha_fin} (8 meses atrás)")

    driver = None
    try:
        driver = crear_driver()
        iniciar_sesion(driver)
        df = extraer_todas_paginas(driver, fecha_inicio, fecha_fin)

        if not df.empty:
            logger.info("\n📊 Vista previa:")
            logger.info(f"\n{df.head().to_string()}")
            df.to_excel(RUTA_COMPLETA, index=False)
            logger.info(f"💾 Datos guardados en '{RUTA_COMPLETA}'")
        else:
            logger.warning("⚠️ No se extrajeron datos.")
    except Exception as e:
        logger.error(f"❌ Error: {e}")
        logger.error(traceback.format_exc())
        if driver:
            try:
                driver.save_screenshot(os.path.join(RUTA_DESTINO, "error.png"))
                logger.info("📸 Captura guardada como error.png")
            except:
                pass
    finally:
        if driver:
            driver.quit()
            logger.info("🚪 Driver cerrado.")