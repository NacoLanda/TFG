"""
Actualización.py
================
Orquestador principal: ejecuta en orden todos los scripts de recopilación
y enriquecimiento de datos de LaLiga 2025-26.

Cada script se lanza como subproceso independiente, mostrando timestamps
de inicio y fin para facilitar el diagnóstico de tiempos.

Orden de ejecución:
  1. Código Partidos.py     → genera Partidos.xlsx sin lluvia (base para Lluvias)
  2. Código FBref.py        → estadísticas de equipo desde HTML de FBref
  3. who/who.py             → estadísticas de WhoScored (equipos + jugadores)
  4. Código Jugadores FBref.py → estadísticas individuales desde HTML de FBref
  5. Código Jugadores Unificados.py → cruza FBref con WhoScored por jugador
  6. Código Lesiones y Sanciones.py → estado actual de bajas y sanciones
  7. Código Árbitros.py     → datos de árbitros de tres fuentes web
  8. Código Lluvias.py      → descarga lluvia histórica y la cruza con partidos
  9. Código Partidos.py     → segunda pasada: enriquece Partidos.xlsx con lluvia

NOTA: "Código Partidos.py" se ejecuta dos veces intencionadamente.
La primera pasada genera el fichero que necesita "Código Lluvias.py".
La segunda pasada añade la columna de lluvia una vez que esos datos existen.

Entrada:  Varios archivos HTML en Descargas FBref/ + conexión a internet
Salida:   Varios archivos .xlsx en Datos/

Uso: python3 "Actualización.py"
"""

import sys
import subprocess
from datetime import datetime
from pathlib import Path

# ══════════════════════════════════════════════════════════════════
# CONFIGURACIÓN — lista ordenada de scripts a ejecutar
# ══════════════════════════════════════════════════════════════════
BASE   = Path(__file__).parent
WHO    = BASE / "who" / "who.py"

SCRIPTS = [
    BASE / "Código Partidos.py",
    BASE / "Código FBref.py",
    WHO,
    BASE / "Código Jugadores Unificados.py",
    BASE / "Código Lesiones y Sanciones.py",
    BASE / "Código Árbitros.py",
    BASE / "Código Lluvias.py",
    BASE / "Código Partidos.py",   # segunda pasada: incorpora datos de lluvia
]


def ejecutar(ruta):
    """Ejecuta un script como subproceso, mostrando timestamps de inicio y fin."""
    nombre = ruta.name
    print(f"\n{'─' * 60}")
    print(f"  Iniciando : {nombre}  ({datetime.now().strftime('%H:%M:%S')})")
    print(f"{'─' * 60}")
    subprocess.run([sys.executable, str(ruta)], check=True)
    print(f"  Finalizado: {nombre}  ({datetime.now().strftime('%H:%M:%S')})")

if __name__ == "__main__":
    inicio = datetime.now()
    print(f"\n{'=' * 60}")
    print(f"  ACTUALIZACIÓN — {inicio.strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"{'=' * 60}")

    for script in SCRIPTS:
        ejecutar(script)

    fin = datetime.now()
    horas, resto = divmod(int((fin - inicio).total_seconds()), 3600)
    minutos, segundos = divmod(resto, 60)

    print(f"\n{'=' * 60}")
    print(f"  FIN:          {fin.strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"  Tiempo total: {horas:02d}h {minutos:02d}m {segundos:02d}s")
    print(f"{'=' * 60}\n")
