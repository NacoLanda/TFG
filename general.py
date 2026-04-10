"""
general.py — Scraper de estadísticas generales de equipo (WhoScored)
=====================================================================
Extrae las secciones Resumen, Defensivo, Ofensivo y xG de la página de
estadísticas de equipos de WhoScored para LaLiga 2025-26.

Para cada sección se recorren tres filtros de localía: General, Local y Visitante,
generando columnas con sufijos _gen, _loc y _vis respectivamente.

La sección xG se trata de forma especial porque tiene un filtro adicional
"A Favor / En Contra" que requiere reiniciar el subfiltro entre ambos estados.

Los datos se acumulan en _datos_temp.json (compartido con el resto de scrapers).

Métricas extraídas (con sufijos _gen / _loc / _vis):
  Resumen  : rating, tiros_pp, aciertoPasePct, aereos
  Defensivo: tiros_contra, entradas_pp, intercep_pp, faltas_pp, fueraJuego_pp
  Ofensivo : tirosAP_pp, regates_pp, faltasFavor_pp
  xG       : xG_fav/con, goles_fav/con, xGDif_fav/con, tiros_fav/con, xGTiros_fav/con

Uso: invocado automáticamente por who.py
"""

import re
import json
from pathlib import Path

from bs4 import BeautifulSoup
import time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# ─── Configuración ────────────────────────────────────────────────────────────
BASE_URL = "https://es.whoscored.com/regions/206/tournaments/4/seasons/10803/stages/24622/teamstatistics/espa%C3%B1a-laliga-2025-2026"

# ─── Driver ───────────────────────────────────────────────────────────────────
def init_driver():
    """Crea y devuelve un WebDriver de Chrome en modo headless."""
    options = Options()
    options.add_argument("--headless")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    driver = webdriver.Chrome()
    driver.implicitly_wait(5)
    return driver

# ─── Store temporal ───────────────────────────────────────────────────────────
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

# ─── Funciones de extracción ────────────────────────────────────────────────── 

def limpiar_prefijo_equipo(nombre):
    """
    Elimina números, puntos y espacios ANTES del nombre del equipo.
    
    Args:
        nombre (str): "1. REAL MADRID" → "REAL MADRID"
                     "12. REAL MADRID" → "REAL MADRID"
    
    Returns:
        str: Nombre limpio
    """
    match = re.search(r'[^0-9\.\s]', nombre)
    if match:
        return nombre[match.start():].strip()
    return nombre.strip()

def actualiza_db(datos):
    actualiza_equipos(datos)


# ─── Configuración de secciones ──────────────────────────────────────────────
# Cada entrada define el id de la tabla HTML, el XPath del botón de sección
# y el mapeo columna_nombre → índice de celda en la fila de la tabla.
CONFIG_SECCIONES = {
    "Resumen": {
        "id_tabla": "statistics-team-table-summary",
        "xpath_sec": None, 
        "mapeo": {"rating": 7, "tiros_pp": 2, "aciertoPasePct": 5, "aereos": 6}
    },
    "Defensivo": {
        "id_tabla": "statistics-team-table-defensive",
        "xpath_sec": "//a[normalize-space()='Defensivo']",
        "mapeo": {"tiros_contra": 1, "entradas_pp": 2, "intercep_pp": 3, "faltas_pp": 4, "fueraJuego_pp": 5}
    },
    "Ofensivo": {
        "id_tabla": "statistics-team-table-offensive",
        "xpath_sec": "//a[normalize-space()='Ofensivo']",
        "mapeo": {"tirosAP_pp": 2, "regates_pp": 3, "faltasFavor_pp": 4}
    },
    "xG": {
        "id_tabla": "statistics-team-table-xg",
        "xpath_sec": "//a[normalize-space()='xG']",
        "mapeo": {"xG": 1, "goles": 2, "xGDif": 3, "tiros": 4, "xGTiros": 5}
    }
}

# ─── Navegación y extracción ─────────────────────────────────────────────────

def extraer_tabla(driver, seccion, sufijo, selector_sub=None, es_contra=False):
    """
    Extrae los datos de la tabla de la sección activa en el navegador.

    Si se indica selector_sub, hace clic en el sub-filtro (Local/Visitante)
    antes de leer el HTML. La tabla se busca por su id definido en CONFIG_SECCIONES.

    Args:
        driver:       WebDriver activo con la página ya cargada.
        seccion:      Clave de CONFIG_SECCIONES ('Resumen', 'Defensivo', etc.).
        sufijo:       Sufijo a añadir al nombre de cada columna ('gen', 'loc', 'vis').
        selector_sub: Selector CSS del botón de sub-filtro (Local/Visitante) o None.
        es_contra:    True cuando se extraen datos xG "En Contra" (cambia el nombre de columna).
    """
    conf = CONFIG_SECCIONES[seccion]
    wait = WebDriverWait(driver, 10)

    if selector_sub:
        btn_sub = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, selector_sub)))
        driver.execute_script("arguments[0].click();", btn_sub)
        time.sleep(3)

    soup = BeautifulSoup(driver.page_source, 'lxml')
    div = soup.find('div', id=conf["id_tabla"])
    if not div:
        return []

    datos_finales = []
    for tr in div.select("tbody tr"):
        tds = [td.text.strip() for td in tr.find_all("td")]
        if len(tds) > 1:
            item = {'equipo': limpiar_prefijo_equipo(tds[0])}
            sub_tipo = "con" if es_contra else "fav"
            for campo, idx in conf["mapeo"].items():
                nombre_col = f"{campo}_{sub_tipo}_{sufijo}" if seccion == "xG" else f"{campo}_{sufijo}"
                item[nombre_col] = tds[idx]
            datos_finales.append(item)
    return datos_finales

# ─── Ejecución principal ─────────────────────────────────────────────────────

def ejecutar_bloque_general():
    """
    Abre el navegador, carga la página y extrae todas las secciones y sub-filtros.

    Bloque 1: Resumen, Defensivo, Ofensivo → 3 secciones × 3 localías = 9 extracciones.
    Bloque 2: xG A Favor y En Contra → 2 estados × 3 localías = 6 extracciones.
    """
    driver = init_driver()
    try:
        driver.get(BASE_URL)
        wait = WebDriverWait(driver, 10)

        # BLOQUE 1: Resumen, Defensivo, Ofensivo — una sola sesión, navegando entre pestañas
        for sec in ["Resumen", "Defensivo", "Ofensivo"]:
            conf = CONFIG_SECCIONES[sec]
            tab_id = conf["id_tabla"].replace("statistics-team-table-", "")

            # Click en la pestaña (el click resetea el subfiltro a General)
            if conf["xpath_sec"]:
                btn_sec = wait.until(EC.element_to_be_clickable((By.XPATH, conf["xpath_sec"])))
                driver.execute_script("arguments[0].click();", btn_sec)
                time.sleep(2)

            for sub_nombre, sufijo, selector_sub in [
                ("General",   "gen", None),
                ("Local",     "loc", f"#stage-team-stats-{tab_id} a[data-value='Home']"),
                ("Visitante", "vis", f"#stage-team-stats-{tab_id} a[data-value='Away']"),
            ]:
                datos = extraer_tabla(driver, sec, sufijo, selector_sub)
                if datos:
                    actualiza_db(datos)
                    print(f"OK: {sec} -> {sub_nombre}")
                else:
                    print(f"AVISO: {sec} -> {sub_nombre}: sin datos")

        # BLOQUE 2: xG — click en pestaña una vez, luego alternar A Favor / En Contra
        conf_xg = CONFIG_SECCIONES["xG"]
        btn_xg = wait.until(EC.element_to_be_clickable((By.XPATH, conf_xg["xpath_sec"])))
        driver.execute_script("arguments[0].click();", btn_xg)
        time.sleep(2)

        for es_con in [False, True]:
            txt = "En Contra" if es_con else "A Favor"

            if es_con:
                # Pulsar "General" primero para resetear el filtro de localía
                btn_general = wait.until(EC.element_to_be_clickable(
                    (By.XPATH, "//div[@id='stage-team-stats-xg']//a[normalize-space()='General']")
                ))
                driver.execute_script("arguments[0].click();", btn_general)
                time.sleep(2)
                # Después pulsar "En Contra"
                btn_con = wait.until(EC.element_to_be_clickable(
                    (By.CSS_SELECTOR, "#stage-team-stats-xg a[data-value='true']")
                ))
                driver.execute_script("arguments[0].click();", btn_con)
                time.sleep(2)

            for sub_nombre, sufijo, selector_sub in [
                ("General",   "gen", None),
                ("Local",     "loc", "#stage-team-stats-xg a[data-value='Home']"),
                ("Visitante", "vis", "#stage-team-stats-xg a[data-value='Away']"),
            ]:
                datos = extraer_tabla(driver, "xG", sufijo, selector_sub, es_contra=es_con)
                if datos:
                    actualiza_db(datos)
                    print(f"OK: xG {txt} -> {sub_nombre}")
                else:
                    print(f"AVISO: xG {txt} -> {sub_nombre}: sin datos")

    except Exception as e:
        print(f"ERROR general: {e}")
    finally:
        driver.quit()

if __name__ == "__main__":
    ejecutar_bloque_general()