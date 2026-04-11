"""
WhoScored LaLiga 2024-2025 - Player Statistics Scraper
=======================================================
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

def _normalizar(nombre):
    nfkd = unicodedata.normalize('NFKD', nombre.strip().lower())
    return ''.join(c for c in nfkd if not unicodedata.combining(c))

def _buscar_jugador(store, nombre):
    clave_norm = _normalizar(nombre)
    for key in store['jugadores']:
        if _normalizar(key) == clave_norm:
            return key
    return None

def reset_jugadores():
    store = _load()
    store['jugadores'] = {}
    _save(store)

def actualiza_jugadores(datos):
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


# ─────────────────────────────────────────────
# CONFIGURACIÓN SCRAPING
# ─────────────────────────────────────────────
URL = (
    "https://es.whoscored.com/regions/206/tournaments/4/seasons/10317/"
    "stages/23401/playerstatistics/espa%C3%B1a-laliga-2024-2025"
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


# ─────────────────────────────────────────────
# CONSOLIDACIÓN: de 15 filas por jugador a 1
# ─────────────────────────────────────────────
COLS_BASE = ["Jgdos", "Mins", "Rating"]

COLS_MENU = {
    "Resumen":      ["Goles", "Asist", "Amar", "Roja", "TpP", "AP%", "Aéreos", "JdelP"],
    "Defensivo":    ["Entrad", "Interc", "Falt", "FJuegoG", "Despe", "Rgts", "Bloq", "Propia"],
    "Ofensivo":     ["PClave", "FaltF", "FJuego", "Despo", "PdasB"],
    "Distribucion": ["PromeP", "Centr", "BLargos", "PHueco"],
    "xG":           ["xG", "xGDif", "xG/90", "Tiros", "xG/Tiros"],
}

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
    sufijos = {"General": "general", "Local": "local", "Visitante": "visitante"}
    resultado = None

    for opcion_texto, sufijo in sufijos.items():
        bloque = None
        for menu, cols_menu in COLS_MENU.items():
            mask = (df_raw["Menu"] == menu) & (df_raw["Opcion"] == opcion_texto)
            if not mask.any():
                continue

            cols_disp = [c for c in COLS_BASE + cols_menu if c in df_raw.columns]

            if menu == "Resumen" and "Suplente" in df_raw.columns:
                cols_disp = cols_disp + ["Suplente"]

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


# ─────────────────────────────────────────────
# DRIVER
# ─────────────────────────────────────────────

def init_driver():
    driver = webdriver.Chrome()
    driver.implicitly_wait(5)
    return driver


# ─────────────────────────────────────────────
# HELPERS SCRAPING
# ─────────────────────────────────────────────

def rand_sleep(lo, hi):
    time.sleep(random.uniform(lo, hi))


def remove_overlays(driver):
    driver.execute_script("""
        document.querySelectorAll('div').forEach(el => {
            const z = parseInt(window.getComputedStyle(el).zIndex);
            if (z > 1000000) el.remove();
        });
    """)


def select_menu(driver, menu_name):
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


MENU_DIV = {
    "Resumen":      "stage-top-player-stats-summary",
    "Defensivo":    "stage-top-player-stats-defensive",
    "Ofensivo":     "stage-top-player-stats-offensive",
    "Distribucion": "stage-top-player-stats-passing",
    "xG":           "stage-top-player-stats-xg",
}


def get_active_table(driver, menu_name=None):
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
    a_player = td_ghost.find("a", class_="player-link")
    return a_player.get_text(strip=True) if a_player else ""


def parsear_jgdos(texto: str) -> tuple[str, str]:
    import re
    texto = texto.strip()
    if not texto or texto == '-':
        return ('0', '0')
    match = re.match(r'^(\d+)\((\d+)\)$', texto)
    if match:
        return (match.group(1), match.group(2))
    return (texto, '0')


def limpiar_celda(texto):
    import re
    return re.sub(r'\(\d+\)', '', texto).strip()


def extract_table(driver, menu_name, opcion):
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
            while headers and not headers[-1]:
                headers.pop()

    rows_data = []
    tbody = table.find("tbody")
    for row in tbody.find_all("tr"):
        cells = row.find_all("td")
        if len(cells) < 3:
            continue

        nombre = extraer_nombre_jugador(cells[1])
        if not nombre:
            continue

        titular, suplente = parsear_jgdos(cells[2].get_text(strip=True))

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


# ─────────────────────────────────────────────
# FLUJO PRINCIPAL
# ─────────────────────────────────────────────

def scrape_all():
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

            if not select_menu(driver, menu_name):
                print(f"  SKIP menú '{menu_name}'")
                continue

            select_todos_jugadores(driver)

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


def insertar_en_store(df):
    reset_jugadores()
    actualiza_jugadores(df)
    print(f"Store  -> {len(df)} jugadores guardados.")


def save_results(df):
    if df.empty:
        print("\nNo se obtuvieron datos.")
        return
    print(f"\nTotal jugadores : {len(df)}")
    insertar_en_store(df)


if __name__ == "__main__":
    from datetime import datetime
    inicio = datetime.now()
    print(f"\n⏱  Inicio: {inicio.strftime('%H:%M:%S')}")

    df = scrape_all()
    save_results(df)

    fin = datetime.now()
    print(f"\n⏱  Fin:    {fin.strftime('%H:%M:%S')}")
    print(f"⏱  Duración: {str(fin - inicio).split('.')[0]}")
