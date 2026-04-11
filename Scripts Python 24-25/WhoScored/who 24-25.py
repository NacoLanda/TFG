"""
who 24-25.py — Orquestador del scraping de WhoScored (temporada 2024-25)
=========================================================================
Versión de la temporada pasada del script who/who.py.
Ejecuta los sub-scrapers de WhoScored para LaLiga 2024-25 y genera
Datos WhoScored 24-25.xlsx en Datos/Temporada Pasada/.

Ver who/who.py (temporada actual) para documentación completa.

Uso: python3 "Temporada Pasada/who 24-25/who 24-25.py"
"""

import sys
import subprocess
import json
from datetime import datetime
from pathlib import Path
import pandas as pd

# ─── Store temporal ───────────────────────────────────────────────────────────
STORE_PATH = Path(__file__).parent / "_datos_temp_24-25.json"
EXCEL_PATH = Path("/Users/NacoLG/Documentos/UFV 4/TFG/Ingeniería del Dato/Datos") / "Temporada Pasada" / "Datos WhoScored 24-25.xlsx"

def _load():
    if STORE_PATH.exists():
        with open(STORE_PATH, 'r', encoding='utf-8') as f:
            return json.load(f)
    return {"equipos": {}, "jugadores": {}}

def _save(store):
    with open(STORE_PATH, 'w', encoding='utf-8') as f:
        json.dump(store, f, ensure_ascii=False, indent=2)

def reset():
    _save({"equipos": {}, "jugadores": {}})

NOMBRE_EQUIPOS = {
    "Deportivo Alaves":  "Alavés",
    "Atletico Madrid":   "Atlético Madrid",
    "Las Palmas":        "Las Palmas",
    "Leganes":           "Leganés",
    "Real Valladolid":   "Valladolid",
}

ORDEN_EQUIPOS = [
    "Alavés", "Athletic Club", "Atlético Madrid", "Barcelona",
    "Celta Vigo", "Espanyol", "Getafe", "Girona",
    "Las Palmas", "Leganés", "Mallorca", "Osasuna",
    "Rayo Vallecano", "Real Betis", "Real Madrid", "Real Sociedad",
    "Sevilla", "Valencia", "Valladolid", "Villarreal",
]

def convertir_numericos(df, col_id):
    """Convierte a numérico todas las columnas salvo la de identificación."""
    for col in df.columns:
        if col == col_id:
            continue
        df[col] = pd.to_numeric(df[col], errors='coerce')
    return df

def to_excel():
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

SCRIPTS = [
    "general 24-25.py",
    "detallado 24-25.py",
    "situacionales 24-25.py",
    "posicionales 24-25.py",
    "jugadores 24-25.py",
    "jugadores2 24-25.py",
]

def ejecutar_script(script):
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
