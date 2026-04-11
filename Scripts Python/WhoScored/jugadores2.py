"""
jugadores2.py — Scraper de estadísticas detalladas de jugadores (WhoScored)
============================================================================
Extrae el menú "Detallado" de estadísticas de jugadores de WhoScored, que
desglosa tiros, goles, pases y bloqueos por zona del campo, situación de juego
y parte del cuerpo empleada.

El menú Detallado requiere seleccionar categoría, subcategoría y tipo de
acumulación mediante desplegables, a diferencia del resto de menús que son
simples pestañas. Por eso este scraper tiene su propia lógica de navegación
independiente de jugadores.py.

Para cada una de las 16 tareas definidas en TAREAS se recorren los tres
filtros de localía (General, Local, Visitante), paginando hasta el final.

Los datos se acumulan en _datos_temp.json junto con el resto de scrapers.

Flujo:
  1. Abre Chrome, navega a WhoScored y pulsa la pestaña "Detallado".
  2. Para cada tarea: selecciona categoría, subcategoría y tipo de acumulación.
  3. Para cada localía: pulsa el botón correspondiente y pagina hasta el final.
  4. Acumula los datos en un dict {jugador → {campo: valor}}.
  5. Convierte el dict a DataFrame y lo guarda en el store JSON.

Uso: invocado automáticamente por who.py
"""

import json
import unicodedata
import time
import random
from pathlib import Path
import pandas as pd
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

# ══════════════════════════════════════════════════════════════════
# STORE TEMPORAL EN JSON
# Mismo mecanismo de persistencia que jugadores.py.
# Los datos de este scraper se añaden sobre los ya existentes:
# no reemplaza, sino que enriquece el registro de cada jugador
# con las columnas detalladas extraídas aquí.
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

def _normalizar(nombre):
    """Convierte un nombre a minúsculas sin tildes para comparaciones robustas."""
    nfkd = unicodedata.normalize('NFKD', nombre.strip().lower())
    return ''.join(c for c in nfkd if not unicodedata.combining(c))

def _buscar_jugador(store, nombre):
    """Devuelve la clave exacta del jugador en el store si existe (con o sin tildes)."""
    clave_norm = _normalizar(nombre)
    for key in store['jugadores']:
        if _normalizar(key) == clave_norm:
            return key
    return None

def actualiza_jugadores(datos):
    """
    Añade o actualiza las columnas detalladas de los jugadores en el store temporal.

    Igual que en jugadores.py, usa normalización de nombres para evitar duplicados.
    Actualiza sobre los datos existentes (no sobreescribe los campos previos).

    Args:
        datos: DataFrame o lista de dicts con clave 'Jugador' o 'jugador'.
    """
    store = _load()
    if hasattr(datos, 'to_dict'):
        records = datos.to_dict(orient='records')
    else:
        records = datos
    for dato in records:
        jugador = (dato.get('Jugador') or dato.get('jugador') or '').strip()
        if not jugador:
            continue
        clave = _buscar_jugador(store, jugador) or jugador
        if clave not in store['jugadores']:
            store['jugadores'][clave] = {}
        store['jugadores'][clave].update(
            {k: v for k, v in dato.items() if k not in ('Jugador', 'jugador')}
        )
    _save(store)


# ══════════════════════════════════════════════════════════════════
# CONFIGURACIÓN
# TAREAS: lista de diccionarios que define cada combinación a extraer.
# Cada tarea especifica qué categoría/subcategoría seleccionar en
# los desplegables, el tipo de acumulación (por partido o total),
# el prefijo de columna y los índices de las celdas a capturar.
# ══════════════════════════════════════════════════════════════════
URL = (
    "https://es.whoscored.com/regions/206/tournaments/4/seasons/10803/"
    "stages/24622/playerstatistics/espa%C3%B1a-laliga-2025-2026"
)

ACTION_DELAY = (1.5, 2.5)
PAGE_DELAY   = (2.5, 4.0)
MAX_PAGES    = 70

OPCIONES = {"General": "Overall", "Local": "Home", "Visitante": "Away"}
SUFIJOS  = {"General": "general", "Local": "local", "Visitante": "visitante"}

# ─────────────────────────────────────────────
# DEFINICIÓN DE TAREAS
# Cada tarea describe qué seleccionar y qué columnas capturar
# ─────────────────────────────────────────────
# Formato: (category, subcategory_or_None, accum_type, prefijo, [col_indices_1based])
# col_indices: posiciones de las columnas a capturar en la tabla (1=primera tras Jugador)
# accum_type: '0'=Por partido, '2'=Total, None=no cambiar

TAREAS = [
    # ── Bloqueos (sin subcategory, por partido) ──────────────────────────
    {
        "label":    "Bloqueos",
        "category": "blocks",
        "subcat":   None,
        "accum":    "0",
        "prefijo":  "bloq2",
        "cols":     {"TirosParados": 1, "CentrBloq": 2, "PasesBloq": 3},
    },
    # ── Paradas (sin subcategory, por partido) ───────────────────────────
    {
        "label":    "Paradas",
        "category": "saves",
        "subcat":   None,
        "accum":    "0",
        "prefijo":  "par",
        "cols":     {"AreaPeq": 2, "AreaPenalti": 3, "FueraArea": 4},
    },
    # ── Tiros - Zonas ────────────────────────────────────────────────────
    {
        "label":    "Tiros-Zonas",
        "category": "shots",
        "subcat":   "zones",
        "accum":    "0",
        "prefijo":  "tir_z",
        "cols":     {"FueraArea": 2, "AreaPeq": 3, "AreaPenalti": 4},
    },
    # ── Tiros - Situaciones ──────────────────────────────────────────────
    {
        "label":    "Tiros-Situaciones",
        "category": "shots",
        "subcat":   "situations",
        "accum":    "0",
        "prefijo":  "tir_s",
        "cols":     {"JuegoAb": 2, "Contra": 3, "BParado": 4, "Penaltis": 5},
    },
    # ── Tiros - Precisión ────────────────────────────────────────────────
    {
        "label":    "Tiros-Precision",
        "category": "shots",
        "subcat":   "accuracy",
        "accum":    "0",
        "prefijo":  "tir_p",
        "cols":     {"FueraPorteria": 2, "AlPoste": 3, "APorteria": 4, "Bloqueado": 5},
    },
    # ── Tiros - Partes del cuerpo ────────────────────────────────────────
    {
        "label":    "Tiros-Cuerpo",
        "category": "shots",
        "subcat":   "bodyparts",
        "accum":    "0",
        "prefijo":  "tir_c",
        "cols":     {"PieDer": 2, "PieIzq": 3, "Cabeza": 4, "Otro": 5},
    },
    # ── Goles - Zonas ────────────────────────────────────────────────────
    {
        "label":    "Goles-Zonas",
        "category": "goals",
        "subcat":   "zones",
        "accum":    "2",
        "prefijo":  "gol_z",
        "cols":     {"AreaPeq": 2, "AreaPenalti": 3, "FueraArea": 4},
    },
    # ── Goles - Situaciones ──────────────────────────────────────────────
    {
        "label":    "Goles-Situaciones",
        "category": "goals",
        "subcat":   "situations",
        "accum":    "2",
        "prefijo":  "gol_s",
        "cols":     {"JuegoAb": 2, "Contra": 3, "BParado": 4, "Penaltis": 5, "Normal": 7},
    },
    # ── Goles - Partes del cuerpo ────────────────────────────────────────
    {
        "label":    "Goles-Cuerpo",
        "category": "goals",
        "subcat":   "bodyparts",
        "accum":    "2",
        "prefijo":  "gol_c",
        "cols":     {"PieDer": 2, "PieIzq": 3, "Cabeza": 4, "Otro": 5},
    },
    # ── Regates (por partido) ────────────────────────────────────────────
    {
        "label":    "Regates",
        "category": "dribbles",
        "subcat":   None,
        "accum":    "0",
        "prefijo":  "rgt",
        "cols":     {"Exitoso": 2, "NoExitoso": 1},   # col 1=Exitoso, col 2=NoExitoso
    },
    # ── Aéreo (por partido) ──────────────────────────────────────────────
    {
        "label":    "Aereo",
        "category": "aerial",
        "subcat":   None,
        "accum":    "0",
        "prefijo":  "aer",
        "cols":     {"Perdidos": 3},    # col 1=Ganados, col 2=Perdidos
    },
    # ── Pases - Longitud (por partido) ──────────────────────────────────
    {
        "label":    "Pases-Longitud",
        "category": "passes",
        "subcat":   "length",
        "accum":    "0",
        "prefijo":  "pas_l",
        "cols":     {"BLPrec": 2, "BLImp": 3, "PCortoPre": 4, "PCortoImp": 5},
    },
    # ── Pases - Tipo (por partido) ───────────────────────────────────────
    {
        "label":    "Pases-Tipo",
        "category": "passes",
        "subcat":   "type",
        "accum":    "0",
        "prefijo":  "pas_t",
        "cols":     {"CentrPrec": 1, "CentrImp": 2, "CrnPrec": 3,
                     "CrnImp": 4, "TirLibPrec": 5, "TirLibImp": 6},
    },
    # ── Pases Clave - Longitud (por partido) ────────────────────────────
    {
        "label":    "PClave-Longitud",
        "category": "key-passes",
        "subcat":   "length",
        "accum":    "0",
        "prefijo":  "pc_l",
        "cols":     {"Largo": 2, "Corto": 3},
    },
    # ── Pases Clave - Tipo (por partido) ────────────────────────────────
    {
        "label":    "PClave-Tipo",
        "category": "key-passes",
        "subcat":   "type",
        "accum":    "0",
        "prefijo":  "pc_t",
        "cols":     {"Centro": 1, "Corner": 2, "PHueco": 3,
                     "TiroLibr": 4, "SBanda": 5, "Otro": 6},
    },
    # ── Asistencias - Total ──────────────────────────────────────────────
    {
        "label":    "Asistencias",
        "category": "assists",
        "subcat":   None,
        "accum":    "2",
        "prefijo":  "ast",
        "cols":     {"Centro": 1, "Corner": 2, "PHueco": 3,
                     "TiroLibr": 4, "SBanda": 5, "Otro": 6},
    },
]


# ══════════════════════════════════════════════════════════════════
# DRIVER Y HELPERS DE SCRAPING
# ══════════════════════════════════════════════════════════════════

def init_driver():
    """Crea y devuelve un WebDriver de Chrome sin modo headless."""
    driver = webdriver.Chrome()
    driver.implicitly_wait(5)
    return driver


def rand_sleep(lo, hi):
    """Espera un tiempo aleatorio entre lo y hi segundos."""
    time.sleep(random.uniform(lo, hi))


def remove_overlays(driver):
    """Elimina via JS los elementos flotantes con z-index extremadamente alto."""
    driver.execute_script("""
        document.querySelectorAll('div').forEach(el => {
            const z = parseInt(window.getComputedStyle(el).zIndex);
            if (z > 1000000) el.remove();
        });
    """)


# ══════════════════════════════════════════════════════════════════
# NAVEGACIÓN — MENÚ DETALLADO
# La vista Detallada de jugadores tiene una estructura distinta a los
# menús estándar: usa desplegables <select> para categoría, subcategoría
# y tipo de acumulación, más botones de id='overall'/'home'/'away'.
# Tras cada selección hay que pulsar el botón "Buscar" para recargar
# la tabla con los nuevos filtros aplicados.
# ══════════════════════════════════════════════════════════════════

def ir_a_detallado(driver):
    """Pulsa la pestaña 'Detallado' del menú de jugadores."""
    a = WebDriverWait(driver, 15).until(
        EC.presence_of_element_located((By.CSS_SELECTOR,
            "#stage-top-player-stats-options a[href='#stage-top-player-stats-detailed']"))
    )
    driver.execute_script("arguments[0].click();", a)
    rand_sleep(*ACTION_DELAY)
    # Esperar a que aparezca el select de categoría
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "category"))
    )


def seleccionar_todos_jugadores(driver):
    """Activa 'Todos los jugadores' si está disponible en la vista detallada."""
    sel = "a.option[data-backbone-model-attribute='isMinApp'][data-value='false']"
    for el in driver.find_elements(By.CSS_SELECTOR, sel):
        if el.is_displayed():
            try:
                el.click()
            except Exception:
                driver.execute_script("arguments[0].click();", el)
            rand_sleep(*ACTION_DELAY)
            print("  -> 'Todos los jugadores' activado.")
            return
    print("  -> 'Todos los jugadores' ya activo o no encontrado.")


def select_categoria(driver, value):
    sel = Select(WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "category"))
    ))
    sel.select_by_value(value)
    rand_sleep(1, 1.5)
    # Pulsar Buscar para recargar la tabla
    try:
        remove_overlays(driver)
        buscar = WebDriverWait(driver, 8).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "button.search-button"))
        )
        driver.execute_script("arguments[0].click();", buscar)
        rand_sleep(*ACTION_DELAY)
    except Exception:
        pass


def select_subcategoria(driver, value):
    try:
        sel = Select(WebDriverWait(driver, 8).until(
            EC.presence_of_element_located((By.ID, "subcategory"))
        ))
        sel.select_by_value(value)
        rand_sleep(*ACTION_DELAY)
    except Exception as e:
        print(f"    [WARN] subcategory '{value}' no encontrado: {e}")


def select_acumulacion(driver, value):
    try:
        sel = Select(WebDriverWait(driver, 8).until(
            EC.presence_of_element_located((By.ID, "statsAccumulationType"))
        ))
        sel.select_by_value(value)
        rand_sleep(*ACTION_DELAY)
    except Exception as e:
        print(f"    [WARN] statsAccumulationType '{value}' no encontrado: {e}")


# Mapa de data_value a id del botón en la vista Detallada
CAMPO_ID = {"Overall": "overall", "Home": "home", "Away": "away"}

def select_opcion_campo(driver, data_value):
    """Selecciona General/Local/Visitante en la vista Detallada.
    Los botones usan id='overall'/'home'/'away' y data-option-value.
    Tras el click hay que pulsar el botón Buscar para recargar la tabla."""
    btn_id = CAMPO_ID.get(data_value)
    if not btn_id:
        print(f"    [WARN] data_value desconocido: {data_value}")
        return False
    try:
        el = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, btn_id))
        )
        driver.execute_script("arguments[0].click();", el)
        rand_sleep(1, 1.5)
        # Pulsar "Buscar" para que recargue la tabla con el nuevo filtro
        remove_overlays(driver)
        buscar = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "button.search-button"))
        )
        driver.execute_script("arguments[0].click();", buscar)
        rand_sleep(*ACTION_DELAY)
        return True
    except Exception as e:
        print(f"    [WARN] select_opcion_campo '{data_value}': {e}")
        return False


# ══════════════════════════════════════════════════════════════════
# EXTRACCIÓN DE DATOS
# extraer_pagina(): lee la tabla activa y devuelve una lista de diccionarios
# con los campos definidos en tarea["cols"] para cada jugador.
# Los índices de columna en tarea["cols"] son 1-based respecto a
# la primera celda de datos (td[4] en el DOM real).
# ══════════════════════════════════════════════════════════════════

def get_tabla_detallada(driver):
    """Devuelve la tabla activa del panel detallado."""
    soup = BeautifulSoup(driver.page_source, "lxml")
    div  = soup.find("div", {"id": "statistics-table-detailed"})
    if div is None:
        return None
    tables = div.find_all("table")
    tables_con_datos = [t for t in tables
                        if t.find("tbody") and t.find("tbody").find("tr")]
    return tables_con_datos[-1] if tables_con_datos else None


def extraer_nombre_jugador(td_ghost):
    """
    Extrae solo el nombre del jugador desde td[1] (grid-ghost-cell).
    El nombre está en <a class="player-link">, sin equipo ni posición.
    """
    a_player = td_ghost.find("a", class_="player-link")
    return a_player.get_text(strip=True) if a_player else ""


def extraer_pagina(driver, tarea, opcion_text):
    """
    Extrae los datos de la página actual para una tarea dada.
    Devuelve lista de dicts {Jugador: ..., col_prefijo_sufijo: valor, ...}
    td[0] = celda visible (ranking + nombre + equipo + edad + posición)
    td[1] = grid-ghost-cell (misma info, usada para extraer nombre limpio)
    td[2..N] = datos estadísticos
    """
    tabla = get_tabla_detallada(driver)
    if tabla is None:
        return []

    sufijo  = SUFIJOS[opcion_text]
    prefijo = tarea["prefijo"]
    cols    = tarea["cols"]   # {nombre_campo: idx_columna_1based}

    filas = []
    tbody = tabla.find("tbody")
    for fila in tbody.find_all("tr"):
        celdas = fila.find_all("td")
        # Necesitamos al menos td[1] (ghost) + las columnas de datos
        if len(celdas) < 2:
            continue

        # td[1] es la ghost-cell con nombre y equipo separados
        jugador = extraer_nombre_jugador(celdas[1])
        if not jugador:
            continue

        # Estructura de celdas:
        # td[0]=ranking, td[1]=ghost, td[2]=Jgdos, td[3]=Mins, td[4..N]=datos
        # Los índices en tarea["cols"] son 1-based respecto a td[4]
        # es decir idx=1 -> celdas[4], idx=2 -> celdas[5], etc.
        registro = {"Jugador": jugador}
        for nombre_col, idx in cols.items():
            campo = f"{prefijo}_{nombre_col}_{sufijo}"
            celda_real = idx + 3   # offset: td[0]+td[1]+td[2]+td[3] = 4 celdas previas
            if celda_real < len(celdas):
                registro[campo] = celdas[celda_real].get_text(strip=True)
            else:
                registro[campo] = None

        filas.append(registro)

    return filas


def has_next_detailed(driver):
    """Compara currentPage con totalPages para saber si hay más páginas."""
    try:
        current = int(driver.find_element(By.CSS_SELECTOR,
            "#statistics-paging-detailed #currentPage").get_attribute("value") or 0)
        total   = int(driver.find_element(By.CSS_SELECTOR,
            "#statistics-paging-detailed #totalPages").get_attribute("value") or 0)
        return current < total
    except Exception:
        return False


def click_next_detailed(driver):
    """Pulsa el botón siguiente via JS (no es visible para Selenium)."""
    try:
        btn = driver.find_element(By.CSS_SELECTOR,
            "#statistics-paging-detailed a.option#next")
        driver.execute_script("arguments[0].click();", btn)
        rand_sleep(*PAGE_DELAY)
        return True
    except Exception:
        return False


def ir_a_primera_pagina(driver):
    """Vuelve a la primera página pulsando el botón 'inicial' via JS.
    Si ya estamos en página 1 no hace nada."""
    try:
        cp = driver.find_element(By.CSS_SELECTOR,
            "#statistics-paging-detailed #currentPage").get_attribute("value")
        if str(cp) == "1":
            return
        btn = driver.find_element(By.CSS_SELECTOR,
            "#statistics-paging-detailed a.option#first")
        driver.execute_script("arguments[0].click();", btn)
        rand_sleep(*ACTION_DELAY)
    except Exception:
        pass


# ══════════════════════════════════════════════════════════════════
# FLUJO PRINCIPAL
# ══════════════════════════════════════════════════════════════════

def scrape_detallado():
    """
    Ejecuta el scraping completo del menú Detallado de jugadores de WhoScored.

    Para cada tarea en TAREAS: selecciona los filtros correspondientes y
    pagina por General, Local y Visitante, acumulando los datos en all_data.
    Al terminar devuelve un DataFrame con una fila por jugador.

    Returns:
        DataFrame con los datos de todas las tareas y localías scrapeados.
    """
    driver = init_driver()
    # all_data: dict {jugador -> {campo: valor}} — acumulador entre tareas y páginas
    all_data = {}

    try:
        print(f"\nAbriendo {URL}\n")
        driver.get(URL)
        rand_sleep(5, 8)
        remove_overlays(driver)

        print("Navegando a Detallado...")
        ir_a_detallado(driver)
        seleccionar_todos_jugadores(driver)
        print("En menú Detallado.\n")

        for tarea in TAREAS:
            label   = tarea["label"]
            print(f"\n{'='*52}")
            print(f"  TAREA: {label}")
            print(f"{'='*52}")

            # Seleccionar categoría
            select_categoria(driver, tarea["category"])

            # Seleccionar subcategoría si aplica
            if tarea["subcat"]:
                select_subcategoria(driver, tarea["subcat"])

            # Seleccionar tipo de acumulación
            if tarea["accum"]:
                select_acumulacion(driver, tarea["accum"])

            rand_sleep(1, 2)

            for opcion_text, data_value in OPCIONES.items():
                print(f"\n  [ {opcion_text} ]")

                # Volver a aplicar categoría/subcat/acum tras cambio de opción
                select_categoria(driver, tarea["category"])
                if tarea["subcat"]:
                    select_subcategoria(driver, tarea["subcat"])
                if tarea["accum"]:
                    select_acumulacion(driver, tarea["accum"])

                ok = select_opcion_campo(driver, data_value)
                if not ok:
                    print(f"    SKIP opción '{opcion_text}'")
                    continue

                # Volver siempre a la página 1 antes de empezar a paginar
                ir_a_primera_pagina(driver)
                rand_sleep(*ACTION_DELAY)

                contador = 1
                while contador <= MAX_PAGES:
                    filas = extraer_pagina(driver, tarea, opcion_text)

                    for f in filas:
                        jugador = f.pop("Jugador")
                        if jugador not in all_data:
                            all_data[jugador] = {"Jugador": jugador}
                        all_data[jugador].update(f)

                    print(f"    Pág {contador:2d}: {len(filas)} jugadores")

                    if not has_next_detailed(driver):
                        print(f"    -> Última página.")
                        break

                    if not click_next_detailed(driver):
                        print(f"    -> No hay siguiente.")
                        break

                    contador += 1

    except KeyboardInterrupt:
        print("\nInterrumpido por el usuario.")
    finally:
        driver.quit()

    df = pd.DataFrame(list(all_data.values()))
    print(f"\nTotal jugadores scrapeados: {len(df)}")
    return df


# ══════════════════════════════════════════════════════════════════
# GUARDAR RESULTADOS
# Persiste el DataFrame en el store JSON compartido.
# ══════════════════════════════════════════════════════════════════

def save_results(df):
    """Guarda el DataFrame en el store JSON. No sobreescribe datos existentes de otros scrapers."""
    if df.empty:
        print("\nNo se obtuvieron datos.")
        return
    print(f"\nTotal jugadores : {len(df)}")
    actualiza_jugadores(df)
    print(f"Store  -> {len(df)} jugadores actualizados.")


# ─────────────────────────────────────────────
# ENTRY POINT
# ─────────────────────────────────────────────

def ejecutar_bloque_jugadores2():
    df = scrape_detallado()
    save_results(df)

if __name__ == "__main__":
    from datetime import datetime
    inicio = datetime.now()
    print(f"\n⏱  Inicio: {inicio.strftime('%H:%M:%S')}")
    ejecutar_bloque_jugadores2()
    fin = datetime.now()
    print(f"\n⏱  Fin:    {fin.strftime('%H:%M:%S')}")
    print(f"⏱  Duración: {str(fin - inicio).split('.')[0]}")
