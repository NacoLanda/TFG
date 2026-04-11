"""
general 24-25.py — Scraper estadísticas generales de equipo · LaLiga 2024-25
=============================================================================
Versión de la temporada pasada del script who/general.py.
Ver who/general.py (temporada actual) para documentación completa.

Uso: invocado automáticamente por who 24-25.py
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

#################################################################################### FUNCIONES

def limpiar_prefijo_equipo(nombre):
    match = re.search(r'[^0-9\.\s]', nombre)
    if match:
        return nombre[match.start():].strip()
    return nombre.strip()

def actualiza_db(datos):
    actualiza_equipos(datos)


#################################################################################### CONFIGURACIÓN
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

#################################################################################### NAVEGACIÓN Y EXTRACCIÓN

def extraer_tabla(driver, seccion, sufijo, selector_sub=None, es_contra=False):
    """Extrae datos de la tabla activa sin abrir nueva sesión."""
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

#################################################################################### EJECUCIÓN

def ejecutar_bloque_general():
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
