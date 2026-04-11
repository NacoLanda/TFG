"""
posicionales 24-25.py — Scraper estadísticas posicionales · LaLiga 2024-25
===========================================================================
Versión de la temporada pasada del script who/posicionales.py.
Ver who/posicionales.py (temporada actual) para documentación completa.

Uso: invocado automáticamente por who 24-25.py
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
BASE_URL = "https://es.whoscored.com/regions/206/tournaments/4/seasons/10317/stages/23401/teamstatistics/espa%C3%B1a-laliga-2024-2025"

# ─── Driver ───────────────────────────────────────────────────────────────────
def init_driver():
    options = Options()
    options.add_argument("--headless")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    driver = webdriver.Chrome()
    driver.implicitly_wait(5)
    return driver

# ─── Store temporal ───────────────────────────────────────────────────────────
STORE_PATH = Path(__file__).parent / "_datos_temp_24-25.json"

def _load():
    if STORE_PATH.exists():
        with open(STORE_PATH, 'r', encoding='utf-8') as f:
            return json.load(f)
    return {"equipos": {}, "jugadores": {}}

def _save(store):
    with open(STORE_PATH, 'w', encoding='utf-8') as f:
        json.dump(store, f, ensure_ascii=False, indent=2)

def actualiza_equipos(datos):
    store = _load()
    for dato in datos:
        equipo = dato.get('equipo')
        if not equipo:
            continue
        if equipo not in store['equipos']:
            store['equipos'][equipo] = {}
        store['equipos'][equipo].update({k: v for k, v in dato.items() if k != 'equipo'})
    _save(store)


# ─────────────────────────────────────────────
# CONFIGURACIÓN
# ─────────────────────────────────────────────
DELAY_CORTO     = (1.5, 2.5)
DELAY_SCROLL    = 1
PAUSA_REINTENTO = (8, 15)


# ─────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────

def limpiar_prefijo_equipo(nombre):
    """Elimina números, puntos y espacios antes del nombre del equipo."""
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


def extraer_tabla(driver, div_id):
    """Extrae todas las filas del div indicado."""
    soup = BeautifulSoup(driver.page_source, 'lxml')
    div  = soup.find('div', id=div_id)
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
    actualiza_equipos(datos)


def ejecutar_tarea(driver, tarea):
    """Pulsa los botones, extrae la tabla y guarda en BD. Devuelve True si OK."""
    for xpath in tarea['clics']:
        if not click_boton(driver, xpath, presence=True):
            return False

    filas = extraer_tabla(driver, tarea['div_id'])
    if not filas:
        return False

    datos = []
    for fila in filas:
        registro = {'equipo': limpiar_prefijo_equipo(fila[1])}
        for nombre_campo, idx in tarea['columnas'].items():
            registro[nombre_campo] = fila[idx]
        datos.append(registro)

    actualiza_db(datos)
    return True


# ─────────────────────────────────────────────
# XPATHS REUTILIZABLES
# ─────────────────────────────────────────────
XPATH = {
    # Menú principal
    'lados_ataque':    "//div[@id='stage-pitch-stats']//a[normalize-space()='Lados de Ataque']",
    'dir_tiro':        "//div[@id='stage-pitch-stats']//a[normalize-space()='Direcciones de Tiro']",
    'zonas_tiro':      "//div[@id='stage-pitch-stats']//a[normalize-space()='Zonas de Tiro']",
    'zonas_accion':    "//div[@id='stage-pitch-stats']//a[normalize-space()='Zonas de Acción']",

    # Filtros campo (Lados de Ataque)
    'lados_general':   "//div[@id='stage-touch-channels-filter-field']//a[normalize-space()='General']",
    'lados_local':     "//div[@id='stage-touch-channels-filter-field']//a[normalize-space()='Local']",
    'lados_visitante': "//div[@id='stage-touch-channels-filter-field']//a[normalize-space()='Visitante']",

    # Filtros campo (Direcciones de Tiro)
    'dir_general':     "//div[@id='stage-attempt-directions-filter-field']//a[normalize-space()='General']",
    'dir_local':       "//div[@id='stage-attempt-directions-filter-field']//a[normalize-space()='Local']",
    'dir_visitante':   "//div[@id='stage-attempt-directions-filter-field']//a[normalize-space()='Visitante']",

    # Filtros a favor/en contra (Direcciones de Tiro)
    'dir_a_favor':     "//div[@id='stage-attempt-directions-filter-against']//a[normalize-space()='A favor']",
    'dir_en_contra':   "//div[@id='stage-attempt-directions-filter-against']//a[normalize-space()='En contra']",

    # Filtros campo (Zonas de Tiro)
    'ztiro_general':   "//div[@id='stage-attempt-zones-filter-field']//a[normalize-space()='General']",
    'ztiro_local':     "//div[@id='stage-attempt-zones-filter-field']//a[normalize-space()='Local']",
    'ztiro_visitante': "//div[@id='stage-attempt-zones-filter-field']//a[normalize-space()='Visitante']",

    # Filtros a favor/en contra (Zonas de Tiro)
    'ztiro_a_favor':   "//div[@id='stage-attempt-zones-filter-against']//a[normalize-space()='A favor']",
    'ztiro_en_contra': "//div[@id='stage-attempt-zones-filter-against']//a[normalize-space()='En contra']",

    # Filtros campo (Zonas de Acción)
    'zaccion_general':   "//div[@id='stage-touch-zones-filter']//a[normalize-space()='General']",
    'zaccion_local':     "//div[@id='stage-touch-zones-filter']//a[normalize-space()='Local']",
    'zaccion_visitante': "//div[@id='stage-touch-zones-filter']//a[normalize-space()='Visitante']",
}


# ─────────────────────────────────────────────
# DEFINICIÓN DE TAREAS
# ─────────────────────────────────────────────

TAREAS = [

    # ── LADOS DE ATAQUE (sin A Favor/En Contra) ───────────────────────────────

    {   'label':   'Equipos.Posicionales.Lados de Ataque.General.',
        'div_id':  'stage-touch-channels',
        'clics':   [XPATH['lados_ataque']],
        'columnas': {
            'ladosAtaque_izquierda_gen': 2,
            'ladosAtaque_centro_gen':    3,
            'ladosAtaque_derecha_gen':   4,
        }
    },
    {   'label':   'Equipos.Posicionales.Lados de Ataque.Local.',
        'div_id':  'stage-touch-channels',
        'clics':   [XPATH['lados_ataque'], XPATH['lados_local']],
        'columnas': {
            'ladosAtaque_izquierda_loc': 2,
            'ladosAtaque_centro_loc':    3,
            'ladosAtaque_derecha_loc':   4,
        }
    },
    {   'label':   'Equipos.Posicionales.Lados de Ataque.Visitante.',
        'div_id':  'stage-touch-channels',
        'clics':   [XPATH['lados_ataque'], XPATH['lados_visitante']],
        'columnas': {
            'ladosAtaque_izquierda_vis': 2,
            'ladosAtaque_centro_vis':    3,
            'ladosAtaque_derecha_vis':   4,
        }
    },

    # ── DIRECCIONES DE TIRO ───────────────────────────────────────────────────

    {   'label':   'Equipos.Posicionales.Direcciones de Tiro.A Favor.General.',
        'div_id':  'stage-attempt-directions',
        'clics':   [XPATH['dir_tiro'], XPATH['dir_a_favor']],
        'columnas': {
            'direccionesTiro_izquierda_fav_gen': 2,
            'direccionesTiro_centro_fav_gen':    3,
            'direccionesTiro_derecha_fav_gen':   4,
        }
    },
    {   'label':   'Equipos.Posicionales.Direcciones de Tiro.En Contra.General.',
        'div_id':  'stage-attempt-directions',
        'clics':   [XPATH['dir_tiro'], XPATH['dir_en_contra']],
        'columnas': {
            'direccionesTiro_izquierda_con_gen': 2,
            'direccionesTiro_centro_con_gen':    3,
            'direccionesTiro_derecha_con_gen':   4,
        }
    },
    {   'label':   'Equipos.Posicionales.Direcciones de Tiro.A Favor.Local.',
        'div_id':  'stage-attempt-directions',
        'clics':   [XPATH['dir_tiro'], XPATH['dir_local'], XPATH['dir_a_favor']],
        'columnas': {
            'direccionesTiro_izquierda_fav_loc': 2,
            'direccionesTiro_centro_fav_loc':    3,
            'direccionesTiro_derecha_fav_loc':   4,
        }
    },
    {   'label':   'Equipos.Posicionales.Direcciones de Tiro.En Contra.Local.',
        'div_id':  'stage-attempt-directions',
        'clics':   [XPATH['dir_tiro'], XPATH['dir_local'], XPATH['dir_en_contra']],
        'columnas': {
            'direccionesTiro_izquierda_con_loc': 2,
            'direccionesTiro_centro_con_loc':    3,
            'direccionesTiro_derecha_con_loc':   4,
        }
    },
    {   'label':   'Equipos.Posicionales.Direcciones de Tiro.A Favor.Visitante.',
        'div_id':  'stage-attempt-directions',
        'clics':   [XPATH['dir_tiro'], XPATH['dir_visitante'], XPATH['dir_a_favor']],
        'columnas': {
            'direccionesTiro_izquierda_fav_vis': 2,
            'direccionesTiro_centro_fav_vis':    3,
            'direccionesTiro_derecha_fav_vis':   4,
        }
    },
    {   'label':   'Equipos.Posicionales.Direcciones de Tiro.En Contra.Visitante.',
        'div_id':  'stage-attempt-directions',
        'clics':   [XPATH['dir_tiro'], XPATH['dir_visitante'], XPATH['dir_en_contra']],
        'columnas': {
            'direccionesTiro_izquierda_con_vis': 2,
            'direccionesTiro_centro_con_vis':    3,
            'direccionesTiro_derecha_con_vis':   4,
        }
    },

    # ── ZONAS DE TIRO ─────────────────────────────────────────────────────────

    {   'label':   'Equipos.Posicionales.Zonas de Tiro.A Favor.General.',
        'div_id':  'stage-attempt-zones',
        'clics':   [XPATH['zonas_tiro'], XPATH['ztiro_a_favor']],
        'columnas': {
            'zonasTiro_areaPequeña_fav_gen': 2,
            'zonasTiro_areaPenalty_fav_gen': 3,
            'zonasTiro_areaFuera_fav_gen':   4,
        }
    },
    {   'label':   'Equipos.Posicionales.Zonas de Tiro.En Contra.General.',
        'div_id':  'stage-attempt-zones',
        'clics':   [XPATH['zonas_tiro'], XPATH['ztiro_en_contra']],
        'columnas': {
            'zonasTiro_areaPequeña_con_gen': 2,
            'zonasTiro_areaPenalty_con_gen': 3,
            'zonasTiro_areaFuera_con_gen':   4,
        }
    },
    {   'label':   'Equipos.Posicionales.Zonas de Tiro.A Favor.Local.',
        'div_id':  'stage-attempt-zones',
        'clics':   [XPATH['zonas_tiro'], XPATH['ztiro_local'], XPATH['ztiro_a_favor']],
        'columnas': {
            'zonasTiro_areaPequeña_fav_loc': 2,
            'zonasTiro_areaPenalty_fav_loc': 3,
            'zonasTiro_areaFuera_fav_loc':   4,
        }
    },
    {   'label':   'Equipos.Posicionales.Zonas de Tiro.En Contra.Local.',
        'div_id':  'stage-attempt-zones',
        'clics':   [XPATH['zonas_tiro'], XPATH['ztiro_local'], XPATH['ztiro_en_contra']],
        'columnas': {
            'zonasTiro_areaPequeña_con_loc': 2,
            'zonasTiro_areaPenalty_con_loc': 3,
            'zonasTiro_areaFuera_con_loc':   4,
        }
    },
    {   'label':   'Equipos.Posicionales.Zonas de Tiro.A Favor.Visitante.',
        'div_id':  'stage-attempt-zones',
        'clics':   [XPATH['zonas_tiro'], XPATH['ztiro_visitante'], XPATH['ztiro_a_favor']],
        'columnas': {
            'zonasTiro_areaPequeña_fav_vis': 2,
            'zonasTiro_areaPenalty_fav_vis': 3,
            'zonasTiro_areaFuera_fav_vis':   4,
        }
    },
    {   'label':   'Equipos.Posicionales.Zonas de Tiro.En Contra.Visitante.',
        'div_id':  'stage-attempt-zones',
        'clics':   [XPATH['zonas_tiro'], XPATH['ztiro_visitante'], XPATH['ztiro_en_contra']],
        'columnas': {
            'zonasTiro_areaPequeña_con_vis': 2,
            'zonasTiro_areaPenalty_con_vis': 3,
            'zonasTiro_areaFuera_con_vis':   4,
        }
    },

    # ── ZONAS DE ACCIÓN (sin A Favor/En Contra) ───────────────────────────────

    {   'label':   'Equipos.Posicionales.Zonas de Accion.General.',
        'div_id':  'stage-touch-zones',
        'clics':   [XPATH['zonas_accion']],
        'columnas': {
            'zonas_accion_defensa_gen':     2,
            'zonas_accion_mediocampo_gen':  3,
            'zonas_accion_ataque_gen':      4,
        }
    },
    {   'label':   'Equipos.Posicionales.Zonas de Accion.Local.',
        'div_id':  'stage-touch-zones',
        'clics':   [XPATH['zonas_accion'], XPATH['zaccion_local']],
        'columnas': {
            'zonas_accion_defensa_loc':     2,
            'zonas_accion_mediocampo_loc':  3,
            'zonas_accion_ataque_loc':      4,
        }
    },
    {   'label':   'Equipos.Posicionales.Zonas de Accion.Visitante.',
        'div_id':  'stage-touch-zones',
        'clics':   [XPATH['zonas_accion'], XPATH['zaccion_visitante']],
        'columnas': {
            'zonas_accion_defensa_vis':     2,
            'zonas_accion_mediocampo_vis':  3,
            'zonas_accion_ataque_vis':      4,
        }
    },
]


# ─────────────────────────────────────────────
# EJECUCIÓN — un único driver para todo
# ─────────────────────────────────────────────

def ejecutar_bloque_posicionales():
    print(datetime.now())

    driver = init_driver()
    driver.get(BASE_URL)
    rand_sleep((10, 15))

    for tarea in TAREAS:
        ok = ejecutar_tarea(driver, tarea)
        if ok:
            print(tarea['label'].ljust(60) + " : OK")
        else:
            print(f"  Fallo en {tarea['label']}, recargando y reintentando...")
            time.sleep(random.uniform(*PAUSA_REINTENTO))
            driver.get(BASE_URL)
            rand_sleep((10, 15))
            ok2 = ejecutar_tarea(driver, tarea)
            if ok2:
                print(tarea['label'].ljust(60) + " : OK (reintento)")
            else:
                print(tarea['label'].ljust(60) + " : FALLIDO")

    driver.quit()
    print(datetime.now())

if __name__ == "__main__":
    ejecutar_bloque_posicionales()
