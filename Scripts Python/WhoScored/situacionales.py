"""
situacionales.py — Scraper de estadísticas situacionales de equipo (WhoScored)
===============================================================================
Extrae la pestaña "Situacionales" de estadísticas de equipo de WhoScored,
que desglosa las métricas según la situación de juego:
  - Juego abierto, Contragolpe, Balón parado, Penaltis

Para cada situación se recorren los tres filtros de localía (General, Local,
Visitante), generando columnas con sufijos _gen, _loc y _vis.

Los datos se acumulan en _datos_temp.json junto con el resto de scrapers.

Uso: invocado automáticamente por who.py
"""

import re
import json
import time
import random
from datetime import datetime
from pathlib import Path

from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# ─── Configuración ────────────────────────────────────────────────────────────
BASE_URL = "https://es.whoscored.com/regions/206/tournaments/4/seasons/10803/stages/24622/teamstatistics/espa%C3%B1a-laliga-2025-2026"

# ─── Driver ───────────────────────────────────────────────────────────────────
def init_driver():
    options = Options()
    options.add_argument("--headless")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    driver = webdriver.Chrome()
    driver.implicitly_wait(5)
    return driver

# ══════════════════════════════════════════════════════════════════
# STORE TEMPORAL EN JSON
# ══════════════════════════════════════════════════════════════════
STORE_PATH = Path(__file__).parent / "_datos_temp.json"

def _load():
    """Carga el store temporal desde disco, o devuelve estructura vacía si no existe."""
    if STORE_PATH.exists():
        with open(STORE_PATH, 'r', encoding='utf-8') as f:
            return json.load(f)
    return {"equipos": {}, "jugadores": {}}

def _save(store):
    """Persiste el store temporal en disco."""
    with open(STORE_PATH, 'w', encoding='utf-8') as f:
        json.dump(store, f, ensure_ascii=False, indent=2)

def actualiza_equipos(datos):
    """Añade o actualiza las métricas de una lista de equipos en el store temporal."""
    store = _load()
    for dato in datos:
        equipo = dato.get('equipo')
        if not equipo:
            continue
        if equipo not in store['equipos']:
            store['equipos'][equipo] = {}
        store['equipos'][equipo].update({k: v for k, v in dato.items() if k != 'equipo'})
    _save(store)


# ══════════════════════════════════════════════════════════════════
# CONFIGURACIÓN
# ══════════════════════════════════════════════════════════════════

DELAY_CORTO     = (1.5, 2.5)
DELAY_SCROLL    = 1
PAUSA_REINTENTO = (8, 15)


# ══════════════════════════════════════════════════════════════════
# HELPERS
# click_boton(): función genérica para pulsar cualquier botón de la
# página esperando a que sea visible y haciendo scroll hasta él.
# ══════════════════════════════════════════════════════════════════

def limpiar_prefijo_equipo(nombre):
    match = re.search(r'[^0-9\.\s]', nombre)
    return nombre[match.start():].strip() if match else nombre.strip()


def rand_sleep(rango=DELAY_CORTO):
    time.sleep(random.uniform(*rango))


def click_boton(driver, xpath, timeout=10, presence=False):
    """Espera un botón, hace scroll hasta él y lo pulsa con JS. Devuelve True si OK."""
    wait      = WebDriverWait(driver, timeout)
    condition = (EC.presence_of_element_located if presence
                 else EC.visibility_of_element_located)
    try:
        boton = wait.until(condition((By.XPATH, xpath)))
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", boton)
        time.sleep(DELAY_SCROLL)
        driver.execute_script("arguments[0].click();", boton)
        rand_sleep()
        return True
    except Exception as e:
        print(f"  Error al pulsar botón [{xpath}]: {e}")
        return False


def extraer_tabla(driver):
    """
    Extrae todas las filas de la tabla situacional activa en el div 'stage-situation-stats'.

    Returns:
        Lista de listas de strings, una por equipo.
    """
    soup = BeautifulSoup(driver.page_source, 'lxml')
    div  = soup.find('div', id='stage-situation-stats')
    if div is None:
        return []
    filas = []
    for tbody in div.find_all("tbody"):
        for fila in tbody.find_all("tr"):
            celdas = [td.text.strip() for td in fila.find_all("td")]
            if celdas:
                filas.append(celdas)
    return filas


def actualiza_db(datos):
    """Wrapper que llama a actualiza_equipos para guardar los datos en el store."""
    actualiza_equipos(datos)


def capturar_y_guardar(driver, columnas, label):
    """
    Extrae la tabla activa, la mapea a los campos indicados y la guarda en el store.

    Args:
        driver:   WebDriver con la tabla ya cargada.
        columnas: Dict {nombre_campo: índice_celda} que define qué columnas capturar.
        label:    Etiqueta descriptiva para el log de progreso.
    """
    filas = extraer_tabla(driver)
    if not filas:
        print(f"  AVISO: sin datos — {label}")
        return
    datos = []
    for fila in filas:
        # fila[1] = nombre del equipo (fila[0] es el número de ranking)
        registro = {'equipo': limpiar_prefijo_equipo(fila[1])}
        for nombre_campo, idx in columnas.items():
            registro[nombre_campo] = fila[idx]
        datos.append(registro)
    actualiza_db(datos)
    print(label.ljust(60) + " : OK")


# ══════════════════════════════════════════════════════════════════
# XPATHS REUTILIZABLES
# Centralizados aquí para no repetirlos dentro del bucle de ejecución.
# Los filtros de localía y de dirección (A favor / En contra) están
# dentro de divs específicos para cada sección, de ahí los XPaths largos.
# ══════════════════════════════════════════════════════════════════
XPATH = {
    # Menú principal
    'tipos_gol':   "//div[@id='stage-situation-stats']//a[normalize-space()='Tipos de Gol']",
    'tipos_pase':  "//div[@id='stage-situation-stats']//a[normalize-space()='Tipos de Pase']",
    'tarjetas':    "//div[@id='stage-situation-stats']//a[normalize-space()='Situación Tarjetas']",

    # Filtros de localía
    'gol_general':       "//div[@id='stage-goals-filter-field']//a[normalize-space()='General']",
    'gol_local':         "//div[@id='stage-goals-filter-field']//a[normalize-space()='Local']",
    'gol_visitante':     "//div[@id='stage-goals-filter-field']//a[normalize-space()='Visitante']",
    'pase_general':      "//div[@id='stage-passes-filter-field']//a[normalize-space()='General']",
    'pase_local':        "//div[@id='stage-passes-filter-field']//a[normalize-space()='Local']",
    'pase_visitante':    "//div[@id='stage-passes-filter-field']//a[normalize-space()='Visitante']",
    'tarjeta_general':   "//div[@id='stage-cards-filter-field']//a[normalize-space()='General']",
    'tarjeta_local':     "//div[@id='stage-cards-filter-field']//a[normalize-space()='Local']",
    'tarjeta_visitante': "//div[@id='stage-cards-filter-field']//a[normalize-space()='Visitante']",

    # Filtros a favor / en contra
    'gol_a_favor':    "//div[@id='stage-goals-filter-against']//a[normalize-space()='A favor']",
    'gol_en_contra':  "//div[@id='stage-goals-filter-against']//a[normalize-space()='En contra']",
    'pase_a_favor':   "//div[@id='stage-passes-filter-against']//a[normalize-space()='A favor']",
    'pase_en_contra': "//div[@id='stage-passes-filter-against']//a[normalize-space()='En contra']",
}


# ══════════════════════════════════════════════════════════════════
# MAPEO DE COLUMNAS POR MENÚ
# Índices de celda para cada campo a capturar en cada sección.
# Los índices son absolutos respecto a fila[0] = nombre del equipo.
# ══════════════════════════════════════════════════════════════════
COLUMNAS_GOL = {
    'gol_juegoAbierto': 2,
    'gol_contraataque': 3,
    'gol_balonParado':  4,
    'gol_penalty':      5,
    'gol_propia':       6,
}

COLUMNAS_PASE = {
    'pase_centros': 2,
    'pase_alHueco': 3,
    'pase_largo':   4,
    'pase_corto':   5,
}

COLUMNAS_TARJETA = {
    'tarjetas_faltas':        2,
    'tarjetas_antideportivo': 3,
    'tarjetas_simulacion':    4,
    'tarjetas_otro':          5,
}


# ══════════════════════════════════════════════════════════════════
# EJECUCIÓN PRINCIPAL
# Tipos de Gol y Tipos de Pase tienen doble bucle: localía × dirección
# (A favor / En contra). Tarjetas solo tiene bucle de localía.
# ══════════════════════════════════════════════════════════════════

def ejecutar_bloque_situacionales():
    print(datetime.now())

    driver = init_driver()
    driver.get(BASE_URL)
    rand_sleep((5, 8))

    # ── TIPOS DE GOL y TIPOS DE PASE (con bucle de localía y dirección) ───────
    for menu_nombre, xpath_menu, localidad_xpaths, dir_xpaths, columnas_base in [
        (
            'TiposGol',  XPATH['tipos_gol'],
            [('gen', XPATH['gol_general']),   ('loc', XPATH['gol_local']),   ('vis', XPATH['gol_visitante'])],
            [('fav', XPATH['gol_a_favor']),   ('con', XPATH['gol_en_contra'])],
            COLUMNAS_GOL,
        ),
        (
            'TiposPase', XPATH['tipos_pase'],
            [('gen', XPATH['pase_general']),  ('loc', XPATH['pase_local']),  ('vis', XPATH['pase_visitante'])],
            [('fav', XPATH['pase_a_favor']),  ('con', XPATH['pase_en_contra'])],
            COLUMNAS_PASE,
        ),
    ]:
        # Pulsar menú
        click_boton(driver, xpath_menu, presence=True)

        # Desde opción 1 a 3 (general, local, visitante)
        for suf_loc, xpath_loc in localidad_xpaths:
            click_boton(driver, xpath_loc, presence=True)

            # Desde botón 1 a 2 (a favor, en contra)
            for suf_dir, xpath_dir in dir_xpaths:
                click_boton(driver, xpath_dir, presence=True)

                columnas = {f"{campo}_{suf_dir}_{suf_loc}": idx
                            for campo, idx in columnas_base.items()}
                label = f"Equipos.Situacionales.{menu_nombre}.{suf_dir}.{suf_loc}."
                capturar_y_guardar(driver, columnas, label)

    # ── TARJETAS (sin bucle de dirección) ─────────────────────────────────────
    click_boton(driver, XPATH['tarjetas'], presence=True)

    # Desde opción 1 a 3 (general, local, visitante)
    for suf_loc, xpath_loc in [
        ('gen', XPATH['tarjeta_general']),
        ('loc', XPATH['tarjeta_local']),
        ('vis', XPATH['tarjeta_visitante']),
    ]:
        click_boton(driver, xpath_loc, presence=True)

        columnas = {f"{campo}_{suf_loc}": idx
                    for campo, idx in COLUMNAS_TARJETA.items()}
        label = f"Equipos.Situacionales.Tarjetas.{suf_loc}."
        capturar_y_guardar(driver, columnas, label)

    driver.quit()
    print(datetime.now())

if __name__ == "__main__":
    ejecutar_bloque_situacionales()
