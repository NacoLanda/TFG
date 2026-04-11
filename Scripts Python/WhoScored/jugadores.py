"""
jugadores.py — Scraper de estadísticas individuales de jugadores (WhoScored)
=============================================================================
Extrae las estadísticas individuales de jugadores de LaLiga 2025-26 desde
WhoScored, navegando por los cinco menús disponibles: Resumen, Defensivo,
Ofensivo, Distribución y xG.

Para cada menú se recorren tres filtros de localía (General, Local, Visitante),
paginando hasta agotar todos los jugadores disponibles (máx. MAX_PAGES páginas
de ~10 jugadores por página).

El resultado se consolida: las ~15 combinaciones menú × localía se reducen a
una única fila por jugador, con columnas con sufijos _general, _local, _visitante.

Los datos se acumulan en _datos_temp.json (compartido con el resto de scrapers).

Flujo:
  1. Abre Chrome y carga la URL de estadísticas de jugadores de WhoScored.
  2. Para cada menú, activa "Todos los jugadores" y recorre General/Local/Visitante.
  3. Por cada combinación, pagina hasta el final extrayendo las filas de la tabla.
  4. Consolida todas las filas en un DataFrame de una fila por jugador.
  5. Guarda el resultado en _datos_temp.json mediante actualiza_jugadores().

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
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

# ══════════════════════════════════════════════════════════════════
# STORE TEMPORAL EN JSON
# Los sub-scrapers de who.py comparten datos a través de _datos_temp.json.
# Cada scraper lee el store, añade sus datos y lo vuelve a guardar.
# La normalización de nombres evita duplicados por tildes o mayúsculas.
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

def reset_jugadores():
    """Borra todos los datos de jugadores del store temporal."""
    store = _load()
    store['jugadores'] = {}
    _save(store)

def actualiza_jugadores(datos):
    """
    Añade o actualiza los datos de una lista de jugadores en el store temporal.

    Acepta tanto un DataFrame como una lista de dicts. Usa normalización de
    nombres para evitar duplicados entre pasadas (ej. con/sin tilde).

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
        # Buscar la clave existente (para no crear duplicados con distintas tildes)
        clave = _buscar_jugador(store, jugador) or jugador
        if clave not in store['jugadores']:
            store['jugadores'][clave] = {}
        store['jugadores'][clave].update(
            {k: v for k, v in dato.items() if k not in ('Jugador', 'jugador')}
        )
    _save(store)


# ══════════════════════════════════════════════════════════════════
# CONFIGURACIÓN DE SCRAPING
# URL de la página de estadísticas de jugadores en WhoScored.
# MENUS: cinco secciones de estadísticas disponibles en la página.
# OPCIONES: los tres filtros de localía disponibles.
# TABLE_ID: id del elemento <table> donde se muestran los resultados.
# MAX_PAGES: límite de seguridad para evitar bucles infinitos.
# ══════════════════════════════════════════════════════════════════
URL = (
    "https://es.whoscored.com/regions/206/tournaments/4/seasons/10803/"
    "stages/24622/playerstatistics/espa%C3%B1a-laliga-2025-2026"
)

MENUS = {
    "Resumen":      "#stage-top-player-stats-summary",
    "Defensivo":    "#stage-top-player-stats-defensive",
    "Ofensivo":     "#stage-top-player-stats-offensive",
    "Distribucion": "#stage-top-player-stats-passing",
    "xG":           "#stage-top-player-stats-xg",
}

OPCIONES = {
    "General":   "Overall",
    "Local":     "Home",
    "Visitante": "Away",
}

TABLE_ID     = "top-player-stats-summary-grid"
ACTION_DELAY = (1.5, 2.5)
PAGE_DELAY   = (2.5, 4.0)
MAX_PAGES    = 70


# ══════════════════════════════════════════════════════════════════
# ESTRUCTURA DE COLUMNAS Y CONSOLIDACIÓN
# El scraping genera ~15 filas por jugador (5 menús × 3 localías).
# COLS_BASE: columnas que aparecen en todos los menús.
# COLS_MENU: columnas específicas de cada menú.
# RENAME_MAP: mapeo nombre_pantalla → nombre_campo para el store JSON.
# La consolidación agrupa todo en una sola fila por jugador.
# ══════════════════════════════════════════════════════════════════
COLS_BASE = ["Jgdos", "Mins", "Rating"]

COLS_MENU = {
    "Resumen":      ["Goles", "Asist", "Amar", "Roja", "TpP", "AP%", "Aéreos", "JdelP"],
    "Defensivo":    ["Entrad", "Interc", "Falt", "FJuegoG", "Despe", "Rgts", "Bloq", "Propia"],
    "Ofensivo":     ["PClave", "FaltF", "FJuego", "Despo", "PdasB"],
    "Distribucion": ["PromeP", "Centr", "BLargos", "PHueco"],
    "xG":           ["xG", "xGDif", "xG/90", "Tiros", "xG/Tiros"],
}

# Mapeo exacto columna Excel -> nombre campo (sin sufijo)
RENAME_MAP = {
    "Jgdos":    "jgdos",
    "Suplente": "suplente",  
    "Mins":     "mins",
    "Rating":   "rating",
    "Goles":    "goles",
    "Asist":    "asist",
    "Amar":     "amar",
    "Roja":     "roja",
    "TpP":      "tpp",
    "AP%":      "ap_pct",
    "Aéreos":   "aereos",
    "JdelP":    "jdelp",
    "Entrad":   "entrad",
    "Interc":   "interc",
    "Falt":     "falt",
    "FJuegoG":  "fjuego_g",
    "Despe":    "despe",
    "Rgts":     "rgts",
    "Bloq":     "bloq",
    "Propia":   "propia",
    "PClave":   "pclave",
    "FaltF":    "faltf",
    "FJuego":   "fjuego_c",
    "Despo":    "despo",
    "PdasB":    "pdasb",
    "PromeP":   "prome_p",
    "Centr":    "centr",
    "BLargos":  "blargos",
    "PHueco":   "phueco",
    "xG":       "xg",
    "xGDif":    "xg_dif",
    "xG/90":    "xg_90",
    "Tiros":    "tiros",
    "xG/Tiros": "xg_tiros",
}


def consolidar(df_raw):
    """
    Reduce el DataFrame en bruto (5 menús × 3 localías × N páginas) a una sola fila por jugador.

    Estrategia: para cada combinación menú+localía, agrupa por Jugador tomando el
    primer valor no-nulo (elimina duplicados de paginación), añade el sufijo de
    localía a cada columna y hace un merge externo por Jugador con el acumulado.

    Args:
        df_raw: DataFrame con columnas Menu, Opcion, Jugador y las métricas scrapeadas.

    Returns:
        DataFrame con una fila por jugador y columnas con sufijos _general/_local/_visitante.
    """
    sufijos = {"General": "general", "Local": "local", "Visitante": "visitante"}
    resultado = None

    for opcion_texto, sufijo in sufijos.items():
        bloque = None
        for menu, cols_menu in COLS_MENU.items():
            mask = (df_raw["Menu"] == menu) & (df_raw["Opcion"] == opcion_texto)
            if not mask.any():
                continue

            cols_disp = [c for c in COLS_BASE + cols_menu if c in df_raw.columns]

            # Suplente solo existe en Resumen — añadirlo únicamente en ese menú
            if menu == "Resumen" and "Suplente" in df_raw.columns:
                cols_disp = cols_disp + ["Suplente"]


            # Agrupar por Jugador tomando primer valor no-nulo → elimina duplicados
            subset = (
                df_raw[mask][["Jugador"] + cols_disp]
                .groupby("Jugador", as_index=False)
                .first()
            )

            rename_map = {
                col: f"{RENAME_MAP[col]}_{sufijo}"
                for col in cols_disp if col in RENAME_MAP
            }
            subset = subset.rename(columns=rename_map)

            if bloque is None:
                bloque = subset
            else:
                cols_ya = [c for c in subset.columns
                           if c != "Jugador" and c in bloque.columns]
                subset = subset.drop(columns=cols_ya, errors="ignore")
                bloque = bloque.merge(subset, on="Jugador", how="outer")

        if bloque is None:
            continue
        if resultado is None:
            resultado = bloque
        else:
            resultado = resultado.merge(bloque, on="Jugador", how="outer")

    return resultado


# ══════════════════════════════════════════════════════════════════
# DRIVER Y HELPERS DE SCRAPING
# remove_overlays: elimina divs con z-index muy alto (banners, GDPR)
# que bloquean los clics de Selenium.
# rand_sleep: esperas aleatorias para imitar comportamiento humano
# y evitar que WhoScored detecte el scraper.
# ══════════════════════════════════════════════════════════════════

def init_driver():
    """Crea y devuelve un WebDriver de Chrome sin modo headless (WhoScored lo detecta)."""
    driver = webdriver.Chrome()
    driver.implicitly_wait(5)
    return driver


def rand_sleep(lo, hi):
    """Espera un tiempo aleatorio entre lo y hi segundos."""
    time.sleep(random.uniform(lo, hi))


def remove_overlays(driver):
    """Elimina via JS los elementos flotantes con z-index extremadamente alto (banners/popups)."""
    driver.execute_script("""
        document.querySelectorAll('div').forEach(el => {
            const z = parseInt(window.getComputedStyle(el).zIndex);
            if (z > 1000000) el.remove();
        });
    """)


def select_menu(driver, menu_name):
    """
    Hace clic en la pestaña del menú indicado (Resumen, Defensivo, etc.).

    Elimina overlays antes de intentar el clic para evitar que un banner
    intercepte la interacción. Usa JS como fallback si el clic directo falla.

    Returns:
        True si el clic fue exitoso, False en caso contrario.
    """
    remove_overlays(driver)
    href     = MENUS[menu_name]
    selector = f"#stage-top-player-stats-options a[href='{href}']"
    try:
        a = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, selector))
        )
        try:
            a.click()
        except Exception:
            driver.execute_script("arguments[0].click();", a)
        rand_sleep(*ACTION_DELAY)
        return True
    except Exception as e:
        print(f"    Error seleccionando menú '{menu_name}': {e}")
        return False


def select_opcion(driver, data_value):
    """
    Selecciona el filtro de localía (Overall/Home/Away) buscando el botón visible.

    WhoScored puede tener varios botones con el mismo data-value; solo se pulsa
    el que está actualmente visible en pantalla.

    Returns:
        True si se encontró y pulsó el botón, False si no hay ninguno visible.
    """
    sel = (f"a.option[data-value='{data_value}']"
           f"[data-backbone-model-attribute='field']")
    for el in driver.find_elements(By.CSS_SELECTOR, sel):
        if el.is_displayed():
            try:
                el.click()
            except Exception:
                driver.execute_script("arguments[0].click();", el)
            rand_sleep(*ACTION_DELAY)
            return True
    return False


def click_siguiente(driver):
    """Pulsa el botón siguiente via JS (el último en el DOM = el activo)."""
    try:
        btns = driver.find_elements(By.CSS_SELECTOR, "a.option#next")
        if not btns:
            return False
        driver.execute_script("arguments[0].click();", btns[-1])
        rand_sleep(*PAGE_DELAY)
        return True
    except Exception:
        return False


# Mapa de menú → id del div contenedor
MENU_DIV = {
    "Resumen":      "stage-top-player-stats-summary",
    "Defensivo":    "stage-top-player-stats-defensive",
    "Ofensivo":     "stage-top-player-stats-offensive",
    "Distribucion": "stage-top-player-stats-passing",
    "xG":           "stage-top-player-stats-xg",
}


def get_active_table(driver, menu_name=None):
    """
    Localiza la tabla de datos activa en el DOM del navegador.

    Primero intenta buscarla dentro del div específico del menú activo
    (más preciso). Si no se pasa menu_name o no se encuentra, cae al
    fallback de buscar por TABLE_ID en todo el DOM.

    Args:
        driver:    WebDriver con la página cargada.
        menu_name: Nombre del menú activo (clave de MENUS) o None.

    Returns:
        Elemento BeautifulSoup de la tabla, o None si no se encuentra.
    """
    soup = BeautifulSoup(driver.page_source, "lxml")

    if menu_name and menu_name in MENU_DIV:
        div = soup.find("div", {"id": MENU_DIV[menu_name]})
        if div:
            tables = div.find_all("table")
            tables_con_datos = [t for t in tables
                                if t.find("tbody") and
                                t.find("tbody").find("tr") and
                                t.find("tbody").find("td")]
            if tables_con_datos:
                return tables_con_datos[-1]

    # Fallback: última tabla del DOM con TABLE_ID
    tables = soup.find_all("table", {"id": TABLE_ID})
    tables_con_datos = [t for t in tables
                        if t.find("tbody") and
                        t.find("tbody").find("tr") and
                        t.find("tbody").find("td")]
    return tables_con_datos[-1] if tables_con_datos else None


def count_rows(driver):
    table = get_active_table(driver)
    if table is None:
        return 0
    tbody = table.find("tbody")
    return len([r for r in tbody.find_all("tr") if r.find("td")])


def extraer_nombre_jugador(td_ghost):
    """
    Extrae solo el nombre del jugador desde td[1] (grid-ghost-cell).
    El nombre está en <a class="player-link">, sin equipo ni posición.
    """
    a_player = td_ghost.find("a", class_="player-link")
    return a_player.get_text(strip=True) if a_player else ""



def parsear_jgdos(texto: str) -> tuple[str, str]:
    """Separa el formato '23(2)' en titular y suplente.
    
    Args:
        texto: Valor crudo de la celda Jgdos, ej: '23(2)' o '15'
    
    Returns:
        Tupla (titular, suplente). Si no hay suplentes, suplente = '0'.
    
    Ejemplos:
        '23(2)' → ('23', '2')
        '15'    → ('15', '0')
        '-'     → ('0', '0')
    """
    import re
    texto = texto.strip()
    if not texto or texto == '-':
        return ('0', '0')
    match = re.match(r'^(\d+)\((\d+)\)$', texto)
    if match:
        return (match.group(1), match.group(2))
    return (texto, '0')



def limpiar_celda(texto):
    """Elimina contenido entre paréntesis como '22(2)' → '22'."""
    import re
    
    return re.sub(r'\(\d+\)', '', texto).strip()


def extract_table(driver, menu_name, opcion):
    """
    Extrae todas las filas de la tabla activa para un menú y opción dados.

    Estructura de celdas en cada fila de WhoScored:
      td[0] = ranking + info (celda visible compleja)
      td[1] = ghost-cell con nombre del jugador limpio (se usa para extraer nombre)
      td[2] = Jgdos (formato '23(2)' — se descompone en titular y suplente)
      td[3..N] = métricas numéricas

    Args:
        driver:    WebDriver con la tabla cargada.
        menu_name: Nombre del menú activo (para localizar el div correcto).
        opcion:    Texto de la opción activa ('General', 'Local', 'Visitante').

    Returns:
        Lista de dicts con una entrada por jugador, más columnas Menu y Opcion.
    """
    table = get_active_table(driver, menu_name)
    if table is None:
        print(f"    Tabla no encontrada para {menu_name}/{opcion}")
        return []

    headers = []
    thead   = table.find("thead")
    if thead:
        first_row = thead.find("tr")
        if first_row:
            headers = [th.get_text(strip=True) for th in first_row.find_all("th")]
            # Eliminar cabeceras vacías al final (artefacto del HTML de WhoScored)
            while headers and not headers[-1]:
                headers.pop()

    rows_data = []
    tbody = table.find("tbody")
    for row in tbody.find_all("tr"):
        cells = row.find_all("td")
        if len(cells) < 3:
            continue

        # td[1] es la ghost-cell con el enlace limpio del jugador
        nombre = extraer_nombre_jugador(cells[1])
        if not nombre:
            continue

        # td[2] tiene el formato '23(2)': 23 partidos titular, 2 como suplente
        titular, suplente = parsear_jgdos(cells[2].get_text(strip=True))

        # Reconstruir la fila en el mismo orden que las cabeceras
        values = (
            [cells[0].get_text(strip=True)]
            + [nombre]
            + [titular]
            + [limpiar_celda(c.get_text(strip=True)) for c in cells[3:]]
        )

        if headers:
            n      = len(headers)
            padded = (values + [""] * n)[:n]
            record = dict(zip(headers, padded))
            record["Jugador"]  = nombre
            record["Suplente"] = suplente
        else:
            record = {f"col_{i}": v for i, v in enumerate(values)}

        record["Menu"]   = menu_name
        record["Opcion"] = opcion
        rows_data.append(record)

    return rows_data



def select_todos_jugadores(driver):
    """Pulsa el botón 'Todos los jugadores' para desactivar el filtro
    de mínimo de partidos y obtener todos los jugadores."""
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
    print("  -> Botón 'Todos los jugadores' no encontrado (puede que ya esté activo).")


# ══════════════════════════════════════════════════════════════════
# FLUJO PRINCIPAL
# Orquesta la navegación completa: menú × localía × páginas.
# Cuando hay menos de 10 jugadores en una página se asume que es
# la última (WhoScored no muestra un indicador de página final explícito).
# ══════════════════════════════════════════════════════════════════

def scrape_all():
    """
    Ejecuta el scraping completo de estadísticas de jugadores de WhoScored.

    Navega por los 5 menús × 3 localías, pagina hasta el final de cada
    combinación y consolida todos los registros en un DataFrame de una fila
    por jugador. Si hay un CAPTCHA, pausa y espera input del usuario.

    Returns:
        DataFrame consolidado con una fila por jugador y columnas con sufijos.
    """
    driver      = init_driver()
    all_records = []

    try:
        print(f"\nAbriendo {URL}\n")
        driver.get(URL)
        rand_sleep(5, 8)
        remove_overlays(driver)

        print("Esperando tabla principal...")
        try:
            WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.ID, TABLE_ID))
            )
            print("Tabla detectada.\n")
        except TimeoutException:
            print("No se detectó la tabla. Si hay captcha, resuélvelo y pulsa Enter...")
            input()

        for menu_name in MENUS:
            print(f"\n{'='*52}")
            print(f"  MENÚ: {menu_name}")
            print(f"{'='*52}")

            # 1. Seleccionar menú
            if not select_menu(driver, menu_name):
                print(f"  SKIP menú '{menu_name}'")
                continue

            # 2. Todos los jugadores — una sola vez por menú
            select_todos_jugadores(driver)

            # 3. Iterar opciones General/Local/Visitante
            for opcion_text, data_value in OPCIONES.items():
                print(f"\n  [ {opcion_text} ]")

                if not select_opcion(driver, data_value):
                    print(f"  SKIP opción '{opcion_text}'")
                    continue

                contador = 1
                while contador <= MAX_PAGES:
                    records = extract_table(driver, menu_name, opcion_text)
                    n = len(records)
                    for r in records:
                        r["Pagina"] = contador
                    all_records.extend(records)

                    print(f"    Pág {contador:2d}: {n} jugadores")

                    # Última página: menos de 10 jugadores
                    if n < 10:
                        print(f"    -> Última página.")
                        break

                    if not click_siguiente(driver):
                        print(f"    -> No hay siguiente.")
                        break

                    contador += 1

    except KeyboardInterrupt:
        print("\nInterrumpido por el usuario.")
    finally:
        driver.quit()

    return consolidar(pd.DataFrame(all_records))


# ══════════════════════════════════════════════════════════════════
# GUARDAR RESULTADOS
# ══════════════════════════════════════════════════════════════════

def save_results(df):
    """Guarda el DataFrame en el store JSON. Resetea primero para evitar datos obsoletos."""
    if df.empty:
        print("\nNo se obtuvieron datos.")
        return

    print(f"\nTotal jugadores : {len(df)}")
    reset_jugadores()
    actualiza_jugadores(df)
    print(f"Store  -> {len(df)} jugadores guardados.")


# ─────────────────────────────────────────────
# ENTRY POINT
# ─────────────────────────────────────────────

if __name__ == "__main__":
    from datetime import datetime
    inicio = datetime.now()
    print(f"\n⏱  Inicio: {inicio.strftime('%H:%M:%S')}")

    #Aqui ejecuta la parte de jugadores.py
    df = scrape_all()
    save_results(df)

    fin = datetime.now()
    print(f"\n⏱  Fin:    {fin.strftime('%H:%M:%S')}")
    print(f"⏱  Duración: {str(fin - inicio).split('.')[0]}")



