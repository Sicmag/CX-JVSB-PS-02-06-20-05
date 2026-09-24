"""
EscribIA - Versión web (Streamlit)
Convierte fotos de apuntes manuscritos en documentos digitales con IA.

Proveedores de IA:
  - Groq (principal)   → rápida y estable
  - Gemini (respaldo)  → cuando Groq falla

Salidas:
  - Solo lo anotado (transcripción literal)
  - Documento completo de estudio (expandido con IA)
  - Presentación de PowerPoint (con imágenes automáticas de Unsplash)
  - Resumen corto
"""

import base64
import io
from datetime import datetime

import requests
import streamlit as st
from PIL import Image
from docx import Document
from pptx import Presentation
from pptx.util import Inches, Pt


# ============ LECTURA DE CLAVES ============

def _leer_clave(nombre):
    """Intenta leer una clave desde Secrets o config.py sin crashear."""
    try:
        return st.secrets[nombre]
    except Exception:
        pass
    try:
        import config
        return getattr(config, nombre)
    except Exception:
        return None


GEMINI_API_KEY = _leer_clave("GEMINI_API_KEY")
GROQ_API_KEY = _leer_clave("GROQ_API_KEY")
UNSPLASH_ACCESS_KEY = _leer_clave("UNSPLASH_ACCESS_KEY")

if not GEMINI_API_KEY and not GROQ_API_KEY:
    st.error(
        "⚠️ No hay ninguna API key de IA configurada. "
        "Agrega GEMINI_API_KEY o GROQ_API_KEY en los Secrets de Streamlit."
    )
    st.stop()


# ============ CONFIGURACIÓN ============

MODELO_GEMINI = "gemini-3.8-flash"
MODELO_GROQ = "qwen/qwen3.8-27b"

URL_GEMINI = (
    "https://generativelanguage.googleapis.com/v1beta/models/"
    f"{MODELO_GEMINI}:generateContent"
)
URL_GROQ = "https://api.groq.com/openai/v1/chat/completions"


# --- Prompts por tipo de salida ---

PROMPT_WORD = """Eres un asistente que digitaliza apuntes manuscritos en español.

Transcribe con fidelidad, corrige ortografía, organiza en secciones con
títulos claros, usa guiones para listas. Si hay partes ilegibles márcalas
como [ilegible]. Devuelve SOLO el texto, sin markdown con asteriscos."""


PROMPT_COMPLETO = """Eres un asistente experto en crear documentos de estudio
a partir de apuntes manuscritos.

Recibirás la foto de un apunte. Genera un DOCUMENTO COMPLETO Y BIEN
ESTRUCTURADO con esta estructura exacta:

TITULO: <título principal del tema, máximo 10 palabras>

INTRODUCCION:
<2-3 líneas explicando de qué trata el tema>

DESARROLLO:
SECCION: <nombre de la primera sección>
<contenido desarrollado, 2-4 párrafos>

SECCION: <nombre de la segunda sección>
<contenido desarrollado>

SECCION: <y así con las secciones necesarias, entre 3 y 6>

CONCEPTOS CLAVE:
- <concepto 1: definición breve>
- <concepto 2>
- <entre 5 y 8 conceptos>

CONCLUSION:
<síntesis de 2-3 líneas sobre lo más importante del tema>

PREGUNTAS DE REPASO:
1. <pregunta 1>
2. <pregunta 2>
3. <pregunta 3>

Reglas:
- Lenguaje claro, formal y didáctico.
- Corrige errores ortográficos sin cambiar el sentido.
- NO inventes información que no esté en el apunte original.
- Si algo no se entiende, márcalo como [ilegible].
- Devuelve SOLO el texto, sin markdown con asteriscos.
- Usa exactamente las etiquetas TITULO:, INTRODUCCION:, DESARROLLO:,
  SECCION:, CONCEPTOS CLAVE:, CONCLUSION:, PREGUNTAS DE REPASO:."""


PROMPT_PPT = """Convierte este apunte manuscrito en diapositivas con explicaciones.

Formato EXACTO:
TITULO: <título corto, máximo 6 palabras>
EXPLICACION: <1-2 frases explicando el tema, sin inventar información>
- punto 1
- punto 2
- punto 3
---
TITULO: <otro título>
EXPLICACION: <explicación>
- punto 1
---
Entre 4 y 8 diapositivas, 3-5 puntos cada una. Puntos de máximo 12 palabras.
Sin asteriscos ni markdown."""


PROMPT_RESUMEN = """Crea un resumen corto (máx 300 palabras) del apunte.
Empieza con un párrafo introductorio, luego viñetas con conceptos clave,
y termina con 3 preguntas de repaso. Sin markdown."""


# ============ MOTOR DE IA ============

def _groq(imagen_b64, mime, prompt):
    """Llama a Groq (proveedor principal)."""
    if not GROQ_API_KEY:
        raise RuntimeError("Groq no disponible.")
    headers = {"Authorization": f"Bearer {GROQ_API_KEY}"}
    payload = {
        "model": MODELO_GROQ,
        "messages": [{
            "role": "user",
            "content": [
                {"type": "text", "text": prompt},
                {"type": "image_url",
                 "image_url": {"url": f"data:{mime};base64,{imagen_b64}"}},
            ],
        }],
        "temperature": 0.2,
    }
    r = requests.post(URL_GROQ, headers=headers, json=payload, timeout=120)
    if r.status_code != 200:
        raise RuntimeError(f"Groq {r.status_code}: {r.text[:200]}")
    return r.json()["choices"][0]["message"]["content"]


def _gemini(imagen_b64, mime, prompt):
    """Llama a Gemini (proveedor de respaldo)."""
    if not GEMINI_API_KEY:
        raise RuntimeError("Gemini no disponible.")
    url = f"{URL_GEMINI}?key={GEMINI_API_KEY}"
    payload = {
        "contents": [{
            "parts": [
                {"text": prompt},
                {"inline_data": {"mime_type": mime, "data": imagen_b64}},
            ]
        }]
    }
    r = requests.post(url, json=payload, timeout=120)
    if r.status_code != 200:
        raise RuntimeError(f"Gemini {r.status_code}: {r.text[:200]}")
    return r.json()["candidates"][0]["content"]["parts"][0]["text"]


def procesar(imagen_b64, mime, prompt):
    """Intenta con Groq primero. Si falla, cae a Gemini."""
    errores = []

    if GROQ_API_KEY:
        try:
            return _groq(imagen_b64, mime, prompt), "Groq"
        except Exception as e:
            errores.append(f"Groq: {e}")

    if GEMINI_API_KEY:
        try:
            return _gemini(imagen_b64, mime, prompt), "Gemini"
        except Exception as e:
            errores.append(f"Gemini: {e}")

    raise RuntimeError(" | ".join(errores) or "Sin proveedores disponibles.")


# ============ BÚSQUEDA DE IMÁGENES ============

def buscar_imagen_unsplash(query):
    """Busca una imagen en Unsplash y devuelve la URL."""
    if not UNSPLASH_ACCESS_KEY:
        return None
    try:
        url = "https://api.unsplash.com/search/photos"
        params = {
            "query": query,
            "per_page": 1,
            "orientation": "landscape",
            "client_id": UNSPLASH_ACCESS_KEY,
        }
        r = requests.get(url, params=params, timeout=15)
        if r.status_code == 200:
            data = r.json()
            if data.get("results"):
                return data["results"][0]["urls"]["regular"]
    except Exception:
        pass
    return None


# ============ GENERACIÓN DE DOCUMENTOS ============

def docx_bytes(texto, titulo="Documento generado por EscribIA"):
    """Genera un Word con formato, detectando etiquetas de sección."""
    doc = Document()

    if "TITULO:" not in texto.upper():
        doc.add_heading(titulo, level=1)

    for linea in texto.split("\n"):
        l = linea.strip()
        if not l:
            continue
        up = l.upper()

        if up.startswith("TITULO:"):
            doc.add_heading(l.split(":", 1)[1].strip(), level=0)
        elif up.startswith("SECCION:"):
            doc.add_heading(l.split(":", 1)[1].strip(), level=2)
        elif up.startswith("INTRODUCCION") or up.startswith("INTRODUCCIÓN"):
            doc.add_heading("Introducción", level=1)
        elif up.startswith("DESARROLLO"):
            doc.add_heading("Desarrollo", level=1)
        elif up.startswith("CONCEPTOS CLAVE"):
            doc.add_heading("Conceptos clave", level=1)
        elif up.startswith("CONCLUSION") or up.startswith("CONCLUSIÓN"):
            doc.add_heading("Conclusión", level=1)
        elif up.startswith("PREGUNTAS DE REPASO"):
            doc.add_heading("Preguntas de repaso", level=1)
        elif l.startswith("-"):
            doc.add_paragraph(l[1:].strip(), style="List Bullet")
        elif l[:2].strip().rstrip(".").isdigit():
            doc.add_paragraph(l, style="List Number")
        else:
            doc.add_paragraph(l)

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


def parsear_slides(texto):
    """Convierte el texto estructurado en lista de diapositivas."""
    slides = []
    for bloque in texto.split("---"):
        lineas = [l.strip() for l in bloque.strip().split("\n") if l.strip()]
        if not lineas:
            continue
        titulo, explicacion, puntos = "", "", []
        for linea in lineas:
            mayus = linea.upper()
            if mayus.startswith("TITULO:"):
                titulo = linea.split(":", 1)[1].strip()
            elif mayus.startswith("EXPLICACION:"):
                explicacion = linea.split(":", 1)[1].strip()
            elif linea.startswith("-"):
                puntos.append(linea[1:].strip())
            elif not titulo:
                titulo = linea
            elif not explicacion:
                explicacion = linea
            else:
                puntos.append(linea)
        if titulo or puntos:
            slides.append((titulo or "Sin título", explicacion, puntos))
    return slides


def pptx_bytes(texto, con_imagenes=True):
    """Genera una presentación con título, explicación, puntos e imágenes."""
    slides = parsear_slides(texto)

    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    # --- Portada ---
    portada = prs.slides.add_slide(prs.slide_layouts[0])
    portada.shapes.title.text = "Apuntes digitalizados"
    portada.placeholders[1].text = (
        f"EscribIA - {datetime.now().strftime('%d/%m/%Y')}"
    )

    # --- Una diapositiva por sección ---
    for titulo, explicacion, puntos in slides:
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        slide.shapes.title.text = titulo

        # --- Imagen automática de Unsplash ---
        tiene_imagen = False
        if con_imagenes and UNSPLASH_ACCESS_KEY and titulo:
            imagen_url = buscar_imagen_unsplash(titulo)
            if imagen_url:
                try:
                    img_data = requests.get(imagen_url, timeout=15).content
                    slide.shapes.add_picture(
                        io.BytesIO(img_data),
                        left=Inches(7.0), top=Inches(1.8),
                        width=Inches(5.8),
                    )
                    tiene_imagen = True
                except Exception:
                    pass

        # Ancho del texto: reducido si hay imagen, completo si no
        ancho_texto = Inches(6.2) if tiene_imagen else Inches(12.1)

        # Caja de explicación
        if explicacion:
            caja_expl = slide.shapes.add_textbox(
                Inches(0.6), Inches(1.6), ancho_texto, Inches(1.5)
            )
            tf = caja_expl.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.text = explicacion
            p.font.size = Pt(16)
            p.font.italic = True

        # Caja de puntos
        if puntos:
            top_puntos = Inches(3.4) if explicacion else Inches(2.0)
            caja_puntos = slide.shapes.add_textbox(
                Inches(0.6), top_puntos, ancho_texto, Inches(4.0)
            )
            tf2 = caja_puntos.text_frame
            tf2.word_wrap = True
            for i, punto in enumerate(puntos):
                if i == 0:
                    par = tf2.paragraphs[0]
                else:
                    par = tf2.add_paragraph()
                par.text = f"•  {punto}"
                par.font.size = Pt(18)

    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()


# ============ INTERFAZ ============

st.set_page_config(page_title="EscribIA", page_icon="📝", layout="centered")

st.title("📝 EscribIA")
st.caption("Apuntes escritos a mano, convertidos en documentos digitales con IA")

st.divider()

formato = st.radio(
    "¿Qué quieres generar?",
    [
        "📝 Solo lo anotado (transcripción literal)",
        "📚 Documento completo de estudio (con IA)",
        "📊 Presentación de PowerPoint",
        "📄 Resumen corto",
    ],
    captions=[
        "Respeta el texto tal cual lo escribiste. Solo corrige ortografía y organiza en párrafos.",
        "La IA amplía y estructura: introducción, secciones desarrolladas, conceptos clave, conclusión y preguntas.",
        "Convierte el apunte en diapositivas con imágenes automáticas de Unsplash.",
        "Resumen de máximo 300 palabras con viñetas de conceptos y preguntas de repaso.",
    ],
)

foto = st.file_uploader(
    "Sube o toma la foto del apunte",
    type=["jpg", "jpeg", "png", "webp"],
)

if foto:
    img = Image.open(foto)
    st.image(img, caption="Apunte cargado", use_container_width=True)

    if st.button("✨ Procesar con IA", type="primary", use_container_width=True):
        # Preparar imagen
        img.thumbnail((1600, 1600))
        buf = io.BytesIO()
        img.convert("RGB").save(buf, "JPEG", quality=85)
        b64 = base64.b64encode(buf.getvalue()).decode()
        mime = "image/jpeg"

        # Elegir prompt según formato
        if "Documento completo" in formato:
            prompt = PROMPT_COMPLETO
        elif "Solo lo anotado" in formato:
            prompt = PROMPT_WORD
        elif "PowerPoint" in formato:
            prompt = PROMPT_PPT
        else:
            prompt = PROMPT_RESUMEN

        with st.spinner("Procesando con inteligencia artificial..."):
            try:
                texto, proveedor = procesar(b64, mime, prompt)
            except Exception as e:
                st.error(f"Error: {e}")
                st.stop()

        st.success(f"✅ Procesado con {proveedor}")
        st.subheader("Contenido generado")
        st.text_area("Texto", texto, height=300, label_visibility="collapsed")

        # Generar archivo según formato
        if "Documento completo" in formato:
            datos = docx_bytes(texto, "Documento de estudio - EscribIA")
            nombre = f"escribia_documento_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
            mime_dl = ("application/vnd.openxmlformats-officedocument."
                       "wordprocessingml.document")

        elif "Solo lo anotado" in formato:
            datos = docx_bytes(texto, "Apunte transcrito - EscribIA")
            nombre = f"escribia_apunte_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
            mime_dl = ("application/vnd.openxmlformats-officedocument."
                       "wordprocessingml.document")

        elif "PowerPoint" in formato:
            if UNSPLASH_ACCESS_KEY:
                with st.spinner("Buscando imágenes para las diapositivas..."):
                    datos = pptx_bytes(texto, con_imagenes=True)
            else:
                st.info(
                    "ℹ️ Para insertar imágenes automáticas, agrega "
                    "UNSPLASH_ACCESS_KEY en los Secrets de Streamlit."
                )
                datos = pptx_bytes(texto, con_imagenes=False)
            nombre = f"escribia_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pptx"
            mime_dl = ("application/vnd.openxmlformats-officedocument."
                       "presentationml.presentation")

        else:
            datos = docx_bytes(texto, "Resumen de estudio - EscribIA")
            nombre = f"escribia_resumen_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
            mime_dl = ("application/vnd.openxmlformats-officedocument."
                       "wordprocessingml.document")

        st.download_button(
            "⬇️ Descargar archivo",
            data=datos,
            file_name=nombre,
            mime=mime_dl,
            type="primary",
            use_container_width=True,
          )
