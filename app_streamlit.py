"""
EscribIA - Versión web (Streamlit)
Convierte fotos de apuntes manuscritos en documentos digitales con IA.

Proveedores:
  - Gemini (principal) → documentos largos y detallados
  - Groq (respaldo)    → cuando Gemini falla o está saturado

Salidas:
  - Documento de Word (transcripción organizada)
  - Documento completo de estudio (introducción, desarrollo, conceptos...)
  - Presentación de PowerPoint
  - Resumen corto con preguntas de repaso
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

# Verificar que al menos una exista
if not GEMINI_API_KEY and not GROQ_API_KEY:
    st.error(
        "⚠️ No hay ninguna API key configurada. "
        "Agrega GEMINI_API_KEY o GROQ_API_KEY en los Secrets de Streamlit."
    )
    st.stop()

# Si la clave de Gemini no tiene el formato correcto (empieza con AIzaSy),
# la ignoramos para evitar crasheos con claves inválidas.
if GEMINI_API_KEY and not GEMINI_API_KEY.startswith("AIzaSy"):
    st.warning(
        "⚠️ La clave de GEMINI_API_KEY no tiene el formato correcto. "
        "Se usará Groq mientras se corrige."
    )
    GEMINI_API_KEY = None


# ============ CONFIGURACIÓN ============

MODELO_GEMINI = "gemini-2.5-flash"
MODELO_GROQ = "meta-llama/llama-4-scout-17b-16e-instruct"

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
TITULO: <título corto, máximo 8 palabras>
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

def _gemini(imagen_b64, mime, prompt):
    """Llama a Gemini (proveedor principal)."""
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


def _groq(imagen_b64, mime, prompt):
    """Llama a Groq (proveedor de respaldo)."""
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


def procesar(imagen_b64, mime, prompt):
    """Intenta con Gemini primero. Si falla, cae a Groq."""
    errores = []

    if GEMINI_API_KEY:
        try:
            return _gemini(imagen_b64, mime, prompt), "Gemini"
        except Exception as e:
            errores.append(f"Gemini: {e}")

    if GROQ_API_KEY:
        try:
            return _groq(imagen_b64, mime, prompt), "Groq"
        except Exception as e:
            errores.append(f"Groq: {e}")

    raise RuntimeError(" | ".join(errores) or "Sin proveedores disponibles.")


# ============ GENERACIÓN DE DOCUMENTOS ============

def docx_bytes(texto, titulo="Documento generado por EscribIA"):
    """Genera un Word con formato, detectando etiquetas de sección."""
    doc = Document()

    # Si el texto no trae TITULO:, agregamos un encabezado por defecto
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


def pptx_bytes(texto):
    """Genera una presentación con título, explicación y puntos por diapositiva."""
    slides = parsear_slides(texto)

    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    portada = prs.slides.add_slide(prs.slide_layouts[0])
    portada.shapes.title.text = "Apuntes digitalizados"
    portada.placeholders[1].text = (
        f"EscribIA - {datetime.now().strftime('%d/%m/%Y')}"
    )

    for titulo, explicacion, puntos in slides:
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        slide.shapes.title.text = titulo

        if explicacion:
            caja = slide.shapes.add_textbox(
                Inches(0.6), Inches(1.6), Inches(12.1), Inches(1.1)
            )
            caja.text_frame.word_wrap = True
            p = caja.text_frame.paragraphs[0]
            p.text = explicacion
            p.font.size = Pt(16)
            p.font.italic = True

        if puntos:
            top = Inches(3.0) if explicacion else Inches(2.0)
            caja2 = slide.shapes.add_textbox(
                Inches(0.6), top, Inches(12.1), Inches(4.2)
            )
            caja2.text_frame.word_wrap = True
            for i, punto in enumerate(puntos):
                if i == 0:
                    par = caja2.text_frame.paragraphs[0]
                else:
                    par = caja2.text_frame.add_paragraph()
                par.text = f"•  {punto}"
                par.font.size = Pt(20)

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
        "📄 Documento de Word",
        "📚 Documento completo de estudio",
        "📊 Presentación de PowerPoint",
        "📝 Resumen corto",
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
        elif "Word" in formato:
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
        elif "Word" in formato:
            datos = docx_bytes(texto, "Documento generado por EscribIA")
            nombre = f"escribia_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
            mime_dl = ("application/vnd.openxmlformats-officedocument."
                       "wordprocessingml.document")
        elif "PowerPoint" in formato:
            datos = pptx_bytes(texto)
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
