#!/usr/bin/env python3
"""
demo_video.py — Montaje del vídeo demo de la aplicación Futbolytics
=====================================================================
Genera un vídeo MP4 a partir de:
  - 10 capturas de pantalla de la app (en orden numérico)
  - Un vídeo de demostración (Video.mp4), con 2 segundos de pausa al inicio
  - Una narración en audio (Audio demo.mp3)
  - Subtítulos sincronizados automáticamente con la voz mediante Whisper

El texto de los subtítulos está definido en SUBTITLE_PARAGRAPHS y coincide
exactamente con lo que dice la narración. Si Whisper no está instalado,
los tiempos se distribuyen de forma proporcional al número de palabras.

Uso: python3 demo_video.py
Salida: presentacion_app.mp4
"""
import sys, subprocess, re
from pathlib import Path

ASSETS = Path("/Users/NacoLG/Documentos/UFV 4/TFG/Análisis de Negocio/Capturas Demo")
TMP    = Path("/Users/NacoLG/Documentos/UFV 4/TFG/Análisis de Negocio/_demo_tmp")
OUT    = Path("/Users/NacoLG/Documentos/UFV 4/TFG/Análisis de Negocio/presentacion_app.mp4")
AUDIO  = Path("/Users/NacoLG/Documentos/UFV 4/TFG/Análisis de Negocio/Audio demo.mp3")

W, H = 1920, 1080
FPS  = 24

IMAGES = [
    (ASSETS / "1.png",  11),
    (ASSETS / "2.png",   5),
    (ASSETS / "3.png",   5),
    (ASSETS / "4.png",   5),
    (ASSETS / "5.png",   4),
    (ASSETS / "6.png",   2),
    (ASSETS / "7.png",   5),
    (ASSETS / "8.png",   6),
    (ASSETS / "9.png",  15),
    (ASSETS / "10.png", 15),
]
VIDEO = ASSETS / "Video.mp4"

# Texto exacto que aparecerá en los subtítulos (en el mismo orden que la voz)
SUBTITLE_PARAGRAPHS = [
    "Esta es una herramienta de análisis prescriptivo para LaLiga. Genera informes tácticos personalizados integrando más de cien variables estadísticas en tiempo real.",
    "En el panel izquierdo seleccionamos nuestro equipo. Elegimos el Real Madrid. A continuación fijamos al rival: el Fútbol Club Barcelona. Seleccionamos opcionalmente al árbitro designado: José Luis Munuera Montero.",
    "Activamos la opción de lluvia prevista en caso de haberla.",
    "Los parámetros quedan registrados.",
    "Configuramos el once titular posición a posición. Si no se toca nada, saldrá en cada posición el jugador más probable.",
    "La pantalla principal ya muestra la tabla de LaLiga ordenada por el Índice de Éxito: una métrica propia que integra potencial ofensivo, solidez defensiva y rating de rendimiento. El Barcelona encabeza con 0.97, el Real Madrid le sigue con 0.858.",
    "El mapa ofensivo-defensivo sitúa a cada equipo en uno de cuatro cuadrantes. Real Madrid y Barcelona comparten el cuadrante élite: alto ataque, alta defensa. El enfrentamiento se anticipa muy equilibrado. Al pulsar Generar informe, el sistema procesa todos los datos. Lo primero que determina es si existe un favorito estadístico claro o si el partido está equilibrado, comparando el Índice de Éxito de ambos equipos y calculando el diferencial.",
    "El Bloque 1 muestra el Resumen General. Primero muestra el resultado del modelo predictivo que estima que habrá un empate, con una ligera ventaja de 0.19 goles para el Barcelona. Después analiza brevemente los indicadores principales de nuestro equipo, en este caso el Real Madrid, y compara las métricas más importantes con las del rival. A continuación hay un gráfico de red que compara el perfil táctico de ambos equipos mediante 6 métricas principales. Por último, compara la forma reciente de ambos equipos, indicando resultados y métricas básicas de los últimos 5 partidos, además del resultado del anterior encuentro entre estos dos equipos.",
    "El Bloque 2 analiza la Estrategia Ofensiva: se compara nuestro potencial ofensivo frente las vulnerabilidades defensivas del rival, incluyendo nuestro xG por partido, los disparos a puerta, y sus XG y disparos concedidos. También se analiza a su portero principal y la distribución goleadora por zonas de nuestro equipo. Por último, se muestran las recomendaciones tácticas ofensivas, respaldadas con datos reales de rendimiento.",
    "El Bloque 3 examina la Estrategia Defensiva: analiza la solidez defensiva propia frente al potencial ofensivo del rival, comparando igual que antes las métricas clave de cada uno. Se analiza a nuestro portero principal y se muestran las recomendaciones tácticas defensivas, con los datos que las sustentan. Esta parte se divide entre cómo ataca el rival y recomendaciones defensivas. Por último, se muestran las duplas peligrosas del rival con el objetivo de centrar la defensa en evitar conexiones entre ellos.",
    "El Bloque 4 evalúa a los Jugadores Clave: presenta la alineación titular propia con el índice de éxito individual de cada jugador y lo mismo para el rival con su alineación más probable. Se muestran los jugadores más influyentes de ambos equipos y tarjetas con el máximo goleador y el máximo asistente del equipo, indicando el peso que tienen sobre la producción total.",
    "El Bloque 5 contextualiza el partido: se analiza el historial estadístico del árbitro designado (en el caso de haberse seleccionado) con la media de sus métricas más importantes, comparadas con la media de LaLiga. Se incorporan los datos climatológicos de lluvia acumulada y días de precipitación de ambas ciudades durante la temporada. Este bloque muestra datos que se deben saber de cara a cada partido, pero se ha concluido en el análisis estadístico que son dos factores que no tienen correlación alguna con el resultado final.",
    "Por último, el Resumen Ejecutivo sintetiza los hallazgos del informe en un conjunto de recomendaciones tácticas concretas listas para ser trasladadas al cuerpo técnico antes del partido.",
]


# ── Dependencias ──────────────────────────────────────────────────────────────
def _ensure_deps():
    """Comprueba que las librerías necesarias están instaladas e instala las que falten."""
    import importlib.util
    pkgs = [
        ("moviepy", "moviepy==1.0.3"),
        ("numpy",   "numpy"),
        ("scipy",   "scipy"),
        ("PIL",     "Pillow"),
    ]
    miss = [pip for mod, pip in pkgs if not importlib.util.find_spec(mod)]
    if miss:
        print(f"[dep] Instalando: {miss}")
        subprocess.check_call([sys.executable, "-m", "pip", "install"] + miss + ["-q"])

_ensure_deps()

import numpy as np
from scipy.io import wavfile
from PIL import Image, ImageDraw, ImageFont
if not hasattr(Image, "ANTIALIAS"):
    Image.ANTIALIAS = Image.LANCZOS
from moviepy.editor import (ImageClip, VideoFileClip, AudioFileClip,
                             CompositeVideoClip, concatenate_videoclips)


# ── Fragmentación del texto en líneas cortas ──────────────────────────────────
def build_subtitle_chunks(max_chars: int = 70) -> list[str]:
    """
    Divide SUBTITLE_PARAGRAPHS en fragmentos cortos aptos para una sola línea de subtítulo.

    Primero divide cada párrafo en frases por puntuación (.!?:) y luego,
    si alguna frase sigue siendo más larga que max_chars, la parte por palabras.

    Returns:
        Lista de strings, cada uno de max_chars caracteres o menos.
    """
    chunks = []
    for para in SUBTITLE_PARAGRAPHS:
        sentences = re.split(r'(?<=[.!?:])\s+', para.strip())
        for sent in sentences:
            sent = sent.strip()
            if not sent:
                continue
            if len(sent) <= max_chars:
                chunks.append(sent)
            else:
                words, current = sent.split(), ""
                for word in words:
                    candidate = (current + " " + word).strip()
                    if len(candidate) <= max_chars:
                        current = candidate
                    else:
                        if current:
                            chunks.append(current)
                        current = word
                if current:
                    chunks.append(current)
    return [c for c in chunks if c.strip()]


# ── Sincronización con Whisper: tiempos reales, texto de SUBTITLE_PARAGRAPHS ──
def _whisper_timings(audio_path: Path, max_chars: int = 70) -> list:
    """
    Usa Whisper (IA de reconocimiento de voz) para obtener el instante exacto en que
    se pronuncia cada palabra del audio, y asigna esos tiempos a los chunks de texto
    de SUBTITLE_PARAGRAPHS mediante una proporción de índice de palabras.

    De este modo el subtítulo muestra el texto exacto definido en el código
    pero cambia en el momento preciso en que la voz lo va diciendo.

    Returns:
        Lista de tuplas (inicio_s, fin_s, texto) listas para crear los overlays.
    """
    import whisper
    import imageio_ffmpeg

    print("[whisper] Cargando modelo 'base'…")
    model = whisper.load_model("base")
    print("[whisper] Transcribiendo audio…")

    # Convertir MP3 → WAV 16 kHz mono con el ffmpeg empaquetado
    ffmpeg_bin = imageio_ffmpeg.get_ffmpeg_exe()
    wav_16k = TMP / "voice_16k.wav"
    TMP.mkdir(exist_ok=True)
    subprocess.check_call(
        [ffmpeg_bin, "-y", "-i", str(audio_path),
         "-ar", "16000", "-ac", "1", "-f", "wav", str(wav_16k)],
        stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
    )
    _, raw = wavfile.read(str(wav_16k))
    audio_arr = raw.astype("float32") / 32768.0

    result = model.transcribe(audio_arr, language="es", word_timestamps=True, verbose=False)

    # Extraer timestamps de cada palabra detectada
    whisper_words = []
    for seg in result.get("segments", []):
        for w in seg.get("words", []):
            if w["word"].strip():
                whisper_words.append((w["start"], w["end"]))

    if not whisper_words:
        print("[whisper] No se detectaron palabras — usando distribución proporcional")
        return _proportional_timings(audio_path, max_chars)

    # Nuestros chunks de texto (exactamente los de SUBTITLE_PARAGRAPHS)
    chunks = build_subtitle_chunks(max_chars)
    chunk_wc = [len(c.split()) for c in chunks]
    total_our = sum(chunk_wc)
    total_w   = len(whisper_words)

    # Mapear cada chunk al rango de palabras de Whisper por proporción de índice
    timings, our_idx = [], 0
    for chunk, wc in zip(chunks, chunk_wc):
        i0 = min(round(our_idx * total_w / total_our), total_w - 1)
        i1 = min(round((our_idx + wc) * total_w / total_our) - 1, total_w - 1)
        i1 = max(i1, i0)
        timings.append((whisper_words[i0][0], whisper_words[i1][1], chunk))
        our_idx += wc

    print(f"[whisper] {len(timings)} chunks · {total_w} palabras detectadas")
    return timings


# ── Fallback proporcional (sin Whisper) ───────────────────────────────────────
def _proportional_timings(audio_path: Path, max_chars: int = 70) -> list:
    """
    Alternativa a _whisper_timings cuando Whisper no está instalado.
    Distribuye el tiempo total del audio entre los chunks de texto
    proporcionalmente al número de palabras de cada uno.

    Returns:
        Lista de tuplas (inicio_s, fin_s, texto).
    """
    chunks = build_subtitle_chunks(max_chars)
    dur    = AudioFileClip(str(audio_path)).duration
    counts = [len(c.split()) for c in chunks]
    total  = sum(counts)
    result, t = [], 0.0
    for text, wc in zip(chunks, counts):
        d = (wc / total) * dur
        result.append((t, t + d, text))
        t += d
    return result


def compute_subtitle_timings(audio_path: Path, max_chars: int = 70) -> list:
    """
    Punto de entrada para calcular los tiempos de los subtítulos.
    Usa Whisper si está instalado; si no, cae al método proporcional.
    """
    try:
        import whisper  # noqa: F401
        return _whisper_timings(audio_path, max_chars)
    except ImportError:
        print("[subs] openai-whisper no instalado → distribución proporcional")
        return _proportional_timings(audio_path, max_chars)


# ── Overlay de subtítulo: una sola línea centrada ─────────────────────────────
def make_subtitle_overlay(text: str, duration: float) -> ImageClip:
    """
    Genera un fotograma RGBA transparente con una barra de subtítulo centrada
    en la parte inferior de la pantalla.

    Dibuja un rectángulo semitransparente oscuro y encima el texto en blanco
    con sombra negra para garantizar legibilidad sobre cualquier fondo.

    Args:
        text:     Texto a mostrar en el subtítulo.
        duration: Duración en segundos que debe estar visible este subtítulo.

    Returns:
        ImageClip de moviepy listo para superponer sobre el vídeo principal.
    """
    fs   = 38
    font = None
    for fp in [
        "/System/Library/Fonts/Supplemental/Arial Bold.ttf",
        "/System/Library/Fonts/Supplemental/Arial.ttf",
        "/System/Library/Fonts/Helvetica.ttc",
    ]:
        if Path(fp).exists():
            try: font = ImageFont.truetype(fp, fs); break
            except Exception: pass
    if font is None:
        font = ImageFont.load_default()

    lh = fs + 12
    y0 = H - lh - 36

    img  = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    draw.rectangle([(0, y0 - 10), (W, y0 + lh + 10)], fill=(5, 8, 20, 175))
    bb = draw.textbbox((0, 0), text, font=font)
    tw = bb[2] - bb[0]
    x  = (W - tw) // 2
    draw.text((x + 1, y0 + 1), text, font=font, fill=(0, 0, 0, 210))
    draw.text((x,     y0),     text, font=font, fill=(255, 255, 255, 255))

    return (ImageClip(np.array(img), ismask=False)
            .set_duration(duration)
            .set_fps(FPS))


# ── Main ──────────────────────────────────────────────────────────────────────
def main():
    """
    Orquesta el montaje completo del vídeo demo en estos pasos:
      1. Concatena las capturas de pantalla con sus duraciones definidas en IMAGES.
      2. Añade 2 segundos de pausa en el primer fotograma de Video.mp4 y lo incorpora.
      3. Carga la narración de voz (Audio demo.mp3).
      4. Si la voz es más larga que el vídeo, congela el último fotograma para cubrirla.
      5. Calcula los tiempos de los subtítulos (Whisper o proporcional) y los superpone.
      6. Mezcla el vídeo con el audio de voz y exporta el resultado como MP4.
    """
    TMP.mkdir(exist_ok=True)

    # 1. Vídeo base
    clips = []
    for img_path, dur in IMAGES:
        c = (ImageClip(str(img_path))
             .resize((W, H))
             .set_duration(dur)
             .set_fps(FPS))
        clips.append(c)
        print(f"[img]  {img_path.name}  →  {dur} s")

    vid = (VideoFileClip(str(VIDEO))
           .without_audio()
           .resize((W, H)))
    freeze = ImageClip(vid.get_frame(0)).set_duration(2).set_fps(FPS)
    clips.append(freeze)
    clips.append(vid)
    print(f"[vid]  {VIDEO.name}  →  {vid.duration:.1f} s")

    video = concatenate_videoclips(clips, method="compose")
    total = video.duration
    print(f"[info] Duración vídeo base: {total:.1f} s")

    # 2. Voz
    voice = AudioFileClip(str(AUDIO))
    print(f"[voz]  {AUDIO.name}  →  {voice.duration:.1f} s")

    # 3. Si la voz supera el vídeo, congelar el último frame
    if voice.duration > total:
        extra      = voice.duration - total
        last_frame = video.get_frame(total - 0.05)
        freeze     = ImageClip(last_frame).set_duration(extra).set_fps(FPS)
        video      = concatenate_videoclips([video, freeze], method="compose")
        total      = video.duration
        print(f"[info] Vídeo extendido a {total:.1f} s")

    # 4. Subtítulos sincronizados (texto exacto de SUBTITLE_PARAGRAPHS)
    timings  = compute_subtitle_timings(AUDIO)
    overlays = []
    for start, end, text in timings:
        ov = make_subtitle_overlay(text, max(end - start, 0.1)).set_start(start)
        overlays.append(ov)
    print(f"[subs] {len(overlays)} segmentos")

    video = CompositeVideoClip([video] + overlays)

    # 5. Solo voz, sin música
    video = video.set_audio(voice)

    # 6. Exportar
    print(f"[out]  Exportando {OUT.name}…")
    video.write_videofile(
        str(OUT), fps=FPS, codec="libx264", audio_codec="aac",
        preset="slow", ffmpeg_params=["-crf", "18"], logger=None,
    )
    print(f"\n✅  {OUT}")


if __name__ == "__main__":
    main()
