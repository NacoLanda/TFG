"""
who.py — Orquestador del scraping de WhoScored
================================================
Ejecuta en secuencia los seis sub-scrapers de WhoScored, cada uno de los
cuales extrae una sección de estadísticas y la almacena en un JSON temporal
(_datos_temp.json). Al finalizar todos, convierte el JSON a Excel.

Sub-scrapers ejecutados en orden:
  1. general.py      → Resumen, Defensivo, Ofensivo, xG (por equipo)
  2. detallado.py    → Pases, Regates, Bloqueos, Zonas de tiro (por equipo)
  3. situacionales.py → Stats desglosadas por situación de juego (por equipo)
  4. posicionales.py  → Stats desglosadas por zona del campo (por equipo)
  5. jugadores.py    → Estadísticas individuales de jugadores (página a página)
  6. jugadores2.py   → Estadísticas detalladas individuales (tiros, pases, goles)

Salida: Datos/Datos WhoScored.xlsx  (2 hojas: Equipos y Jugadores)

Uso: python3 who/who.py
     (o desde Actualización.py, que lo llama automáticamente)
"""

import sys
import subprocess
import json
from datetime import datetime
from pathlib import Path
import pandas as pd

# ══════════════════════════════════════════════════════════════════
# STORE TEMPORAL EN JSON
# Los sub-scrapers comparten datos a través de _datos_temp.json,
# que se crea al inicio de who.py y se borra al generar el Excel final.
# Cada sub-scraper añade sus datos sobre lo que ya existe en el JSON.
# ══════════════════════════════════════════════════════════════════
STORE_PATH = Path(__file__).parent / "_datos_temp.json"
EXCEL_PATH = Path("/Users/NacoLG/Documentos/UFV 4/TFG/Ingeniería del Dato/Datos") / "Datos WhoScored.xlsx"

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

def reset():
    """Reinicia el store temporal (borra todos los datos acumulados)."""
    _save({"equipos": {}, "jugadores": {}})

# ══════════════════════════════════════════════════════════════════
# NORMALIZACIÓN DE NOMBRES
# WhoScored usa algunos nombres distintos al resto del proyecto.
# NOMBRE_EQUIPOS: corrige los nombres antes de escribir el Excel.
# ORDEN_EQUIPOS: define el orden de filas en la hoja Equipos.
# ══════════════════════════════════════════════════════════════════
NOMBRE_EQUIPOS = {
    "Deportivo Alaves":  "Alavés",
    "Atletico Madrid":   "Atlético Madrid",
    "Real Oviedo":       "Oviedo",
}

ORDEN_EQUIPOS = [
    "Alavés", "Athletic Club", "Atlético Madrid", "Barcelona",
    "Celta Vigo", "Elche", "Espanyol", "Getafe", "Girona", "Levante",
    "Mallorca", "Osasuna", "Oviedo", "Rayo Vallecano", "Real Betis",
    "Real Madrid", "Real Sociedad", "Sevilla", "Valencia", "Villarreal",
]

def convertir_numericos(df, col_id):
    """Convierte a numérico todas las columnas excepto la de identificación."""
    for col in df.columns:
        if col == col_id:
            continue
        df[col] = pd.to_numeric(df[col], errors='coerce')
    return df

def to_excel():
    """
    Consolida el JSON temporal en un Excel con dos hojas: Equipos y Jugadores.

    Los equipos se ordenan según ORDEN_EQUIPOS. Los valores NaN se
    rellenan con 0 para que Excel los trate como números, no como vacíos.
    Al finalizar, borra el JSON temporal para dejar el directorio limpio.
    """
    store = _load()
    df_equipos = pd.DataFrame([
        {'equipo': NOMBRE_EQUIPOS.get(nombre, nombre), **cols}
        for nombre, cols in store['equipos'].items()
    ])
    # Reordenar según ORDEN_EQUIPOS; equipos no listados van al final
    orden_map = {nombre: i for i, nombre in enumerate(ORDEN_EQUIPOS)}
    df_equipos['_orden'] = df_equipos['equipo'].map(lambda x: orden_map.get(x, len(ORDEN_EQUIPOS)))
    df_equipos = df_equipos.sort_values('_orden').drop(columns='_orden').reset_index(drop=True)
    df_jugadores = pd.DataFrame([
        {'jugador': nombre, **cols}
        for nombre, cols in store['jugadores'].items()
    ])

    # Convertir strings numéricos ("2.1") a float para que Excel los trate como números
    # Las celdas sin dato quedan como 0
    df_equipos  = convertir_numericos(df_equipos,  'equipo')
    df_jugadores = convertir_numericos(df_jugadores, 'jugador')

    cols_num_eq  = [c for c in df_equipos.columns  if c != 'equipo']
    cols_num_jug = [c for c in df_jugadores.columns if c != 'jugador']
    df_equipos[cols_num_eq]   = df_equipos[cols_num_eq].fillna(0)
    df_jugadores[cols_num_jug] = df_jugadores[cols_num_jug].fillna(0)

    EXCEL_PATH.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(EXCEL_PATH, engine='openpyxl') as writer:
        df_equipos.to_excel(writer, sheet_name='Equipos', index=False)
        df_jugadores.to_excel(writer, sheet_name='Jugadores', index=False)
    print(f"Excel guardado en: {EXCEL_PATH}")
    STORE_PATH.unlink(missing_ok=True)

# ══════════════════════════════════════════════════════════════════
# EJECUCIÓN PRINCIPAL
# ══════════════════════════════════════════════════════════════════
SCRIPTS = [
    "general.py",
    "detallado.py",
    "situacionales.py",
    "posicionales.py",
    "jugadores.py",
    "jugadores2.py",
]

def ejecutar_script(script):
    """Ejecuta un sub-scraper como subproceso, mostrando timestamps de inicio y fin."""
    ruta = Path(__file__).parent / script
    print(f"\n{'─'*60}")
    print(f"  Iniciando: {script}  ({datetime.now().strftime('%H:%M:%S')})")
    print(f"{'─'*60}")
    subprocess.run([sys.executable, str(ruta)], check=True)
    print(f"  Finalizado: {script}  ({datetime.now().strftime('%H:%M:%S')})")


if __name__ == "__main__":
    inicio = datetime.now()
    print(f"\n{'='*60}")
    print(f"  INICIO: {inicio.strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"{'='*60}")

    reset()

    for script in SCRIPTS:
        ejecutar_script(script)

    print(f"\n{'─'*60}")
    print(f"  Generando Excel...")
    to_excel()

    fin = datetime.now()
    horas, resto = divmod(int((fin - inicio).total_seconds()), 3600)
    minutos, segundos = divmod(resto, 60)

    print(f"\n{'='*60}")
    print(f"  FIN:          {fin.strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"  Tiempo total: {horas:02d}h {minutos:02d}m {segundos:02d}s")
    print(f"{'='*60}\n")
