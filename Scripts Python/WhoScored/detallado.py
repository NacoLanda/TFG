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
from selenium.webdriver.support.ui import Select

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

# ─── Store temporal ───────────────────────────────────────────────────────────
STORE_PATH = Path(__file__).parent / "_datos_temp.json"

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


def extraer_datos_tabla(driver, sufijo, categoria):
    """Extrae la información de la tabla y formatea el diccionario de salida."""
    soup = BeautifulSoup(driver.page_source, 'lxml')
    div = soup.find('div', id='statistics-team-table-detailed')

    if not div:
        return []

    # 1. Definimos las columnas por categoría
    config_columnas = {
        'Entradas': [
            f'entrada_exito_{sufijo}',
            f'entrada_fallo_{sufijo}'
        ],
        'Intercepciones': [
            f'intercepciones_{sufijo}'
        ],
        'Faltas': [
            f'faltas_recibidas_{sufijo}',
            f'faltas_cometidas_{sufijo}'
        ],
        'Despejes': [
            f'despejes_{sufijo}'
        ],
        'Bloqueos': [
            f'bloqueados_tiros_{sufijo}',
            f'bloqueados_centros_{sufijo}',
            f'bloqueados_pases_{sufijo}'
        ],
        'Paradas': [
            f'paradas_total_{sufijo}',
            f'paradas_pequeña_{sufijo}',
            f'paradas_area_{sufijo}',
            f'paradas_fuera_{sufijo}'
        ],
        'Tiros': [
            f'tiros_fueraArea_{sufijo}',
            f'tiros_areaPequeña_{sufijo}',
            f'tiros_area_{sufijo}'
        ],
        'Goles': [
            f'goles_areaPequeña_{sufijo}',
            f'goles_area_{sufijo}',
            f'goles_fueraArea_{sufijo}',
        ],
        'Regates': [
            f'regates_fallidos_{sufijo}',
            f'regates_exitosos_{sufijo}'
        ],
        'Posesión perdida': [
            f'perdida_toqueFallido_{sufijo}',
            f'perdida_desposeido_{sufijo}'
        ],
        'Aéreo': [
            f'balonesAereos_ganados_{sufijo}',
            f'balonesAereos_perdidos_{sufijo}'
        ],
        'Pases': [
            f'pases_total_{sufijo}',
            f'pases_largosPrecisos_{sufijo}',
            f'pases_largosImprecisos_{sufijo}',
            f'pases_cortosPrecisos_{sufijo}',
            f'pases_cortosImprecisos_{sufijo}'
        ],
        'Pases clave': [
            f'pasesClave_largo_{sufijo}',
            f'pasesClave_corto_{sufijo}'
        ],
        'Asistencias': [
            f'asistencias_centro_{sufijo}',
            f'asistencias_corner_{sufijo}',
            f'asistencias_alHueco_{sufijo}',
            f'asistencias_tiroLibre_{sufijo}',
            f'asistencias_banda_{sufijo}',
            f'asistencias_otro_{sufijo}'
        ],
        'Tarjetas': [
            f'tarjetas_amarilla_{sufijo}',
            f'tarjetas_roja_{sufijo}'
        ],
    }

    # Obtenemos la lista de nombres de columnas para la categoría actual
    columnas_db = config_columnas.get(categoria, [])

    resultados = []
    for fila in div.select("tbody tr"):
        celdas = [td.get_text(strip=True) for td in fila.find_all("td")]

        if celdas:
            # El primer elemento (celdas[0]) siempre es el nombre del equipo
            item = {'equipo': limpiar_prefijo_equipo(celdas[0])}

            # MAPEO DINÁMICO: emparejamos nombres de columnas con celdas
            for i, nombre_columna in enumerate(columnas_db):
                # i+1 porque la celda 0 es el nombre del equipo, empezamos en la 1
                # Si hay columnas intercaladas hay que hacer el desplazamiento en i:
                if (categoria == 'Goles' or categoria == 'Tiros' or categoria == 'Aéreo' or categoria == 'Pases clave'):
                    i = i + 1 #Saltamos la columna 'Total'

                if (i + 1) < len(celdas):
                    item[nombre_columna] = celdas[i+1]

            resultados.append(item)

    return resultados

#################################################################################### EJECUCIÓN

def ejecutar_bloque_detallado_completo():
    categorias = [
        'Entradas', 'Intercepciones', 'Faltas', 'Despejes', 'Bloqueos',
        'Paradas', 'Tiros', 'Goles', 'Regates', 'Posesión perdida',
        'Aéreo', 'Pases', 'Pases clave', 'Asistencias', 'Tarjetas'
    ]

    driver = init_driver()
    try:
        driver.get(BASE_URL)
        wait = WebDriverWait(driver, 10)

        # Click en pestaña 'Detallado' una sola vez
        detallado_xpath = "//a[normalize-space()='Detallado' and @href='#stage-team-stats-detailed']"
        detallado_link = wait.until(EC.element_to_be_clickable((By.XPATH, detallado_xpath)))
        driver.execute_script("arguments[0].click();", detallado_link)
        wait.until(EC.presence_of_element_located((By.ID, "category")))

        botones_localidad = [
            ('General',   'gen'),
            ('Local',     'loc'),
            ('Visitante', 'vis'),
        ]

        for categoria in categorias:
            # Elegir opción del desplegable
            select_element = wait.until(EC.presence_of_element_located((By.ID, "category")))
            Select(select_element).select_by_visible_text(categoria)
            time.sleep(2)

            # Pulsar botón 1 (General), 2 (Local), 3 (Visitante) y capturar datos
            for texto_boton, sufijo in botones_localidad:
                try:
                    boton = wait.until(EC.element_to_be_clickable(
                        (By.XPATH, f"//div[@id='stage-team-stats-detailed']//a[normalize-space()='{texto_boton}']")
                    ))
                    driver.execute_script("arguments[0].click();", boton)
                    time.sleep(2)
                except Exception as e:
                    print(f"ERROR al pulsar '{texto_boton}' en '{categoria}': {e}")
                    continue

                datos = extraer_datos_tabla(driver, sufijo, categoria)
                if datos:
                    actualiza_db(datos)
                    print(f"Equipos.Detallado.{categoria.capitalize()}.{texto_boton}.".ljust(60) + " : OK")
                else:
                    print(f"AVISO: Detallado.{categoria}.{texto_boton}: sin datos")

    except Exception as e:
        print(f"ERROR general: {e}")
    finally:
        driver.quit()

if __name__ == "__main__":
    ejecutar_bloque_detallado_completo()
