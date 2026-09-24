"""
EscribIA - Versión web unificada con autenticación y base de datos.

Incluye:
  - Login y registro con Supabase (usando st-login-form)
  - 4 formatos: Solo lo anotado, Documento completo, PowerPoint, Resumen
  - Contador de usos mensuales (plan freemium, límite 7)
  - Doble proveedor de IA: Groq (principal) + Gemini (respaldo)
  - Imágenes automáticas de Unsplash para PowerPoint
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
from supabase import create_client
from st_login_form import login_form


# ============ LECTURA DE CLAVES ============

def _leer_clave(nombre):
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
SUPABASE_URL = _leer_clave("SUPABASE_URL")
SUPABASE_KEY = _leer_clave("SUPABASE_KEY")

if not GEMINI_API_KEY and not GROQ_API_KEY:
    st.error("⚠️ No hay ninguna API key de IA configurada.")
    st.stop()

if not SUPABASE_URL or not SUPABASE_KEY:
    st.error("⚠️ Falta configurar SUPABASE_URL o SUPABASE_KEY.")
    st.stop()

LIMITE_GRATUITO = 7


# ============ CONEXIÓN A SUPABASE ============

try:
    supabase_client = create_client(SUPABASE_URL, SUPABASE_KEY)
except Exception as e:
    st.error(f"Error conectando a la base de datos: {e}")
    st.stop()


# ============ CONFIGURACIÓN DE IA ============

MODELO_GEMINI = "gemini-3.8-flash"
MODELO_GROQ = "qwen/qwen3.8-27b"

URL_GEMINI = (
    "https://generativelanguage.googleapis.com/v1beta/models/"
    f"{MODELO_GEMINI}:generateContent"
)
URL_GROQ = "https://api.groq.com/openai/v1/chat/completions"


# --- Prompts ---

PROMPT_WORD = """Eres un asistente que digitaliza apuntes manuscritos en español.

Transcribe con fidelidad, corrige ortografía, organiza en secciones con
títulos claros, usa guiones para listas. Si hay partes ilegibles márcalas
como [ilegible]. Devuelve SOLO el texto, sin markdown con asteriscos.
Sé conciso: no agregues contenido extra."""

PROMPT_COMPLETO = """Crea un documento de estudio breve y organizado a partir del apunte.

Estructura exacta:

TITULO: <título del tema, máximo 8 palabras>

INTRODUCCION:
<2 líneas sobre de qué trata>

DESARROLLO:
SECCION: <nombre de la sección 1>
<1 párrafo de 3-4 líneas>

SECCION: <nombre de la sección 2>
<1 párrafo de 3-4 líneas>

SECCION: <nombre de la sección 3>
<1 párrafo de 3-4 líneas>

CONCEPTOS CLAVE:
- <concepto 1: definición corta>
- <concepto 2>
- <concepto 3>
- <concepto 4>

CONCLUSION:
<1-2 líneas>

PREGUNTAS DE REPASO:
1. <pregunta 1>
2. <pregunta 2>
3. <pregunta 3>

Reglas:
- Sé conciso. Total máximo: 500 palabras.
- No inventes información que no esté en el apunte.
- Sin asteriscos ni markdown."""

PROMPT_PPT = """Convierte este apunte en diapositivas.

Formato EXACTO:
TITULO: <título corto, máximo 6 palabras>
EXPLICACION: <1 frase explicando el tema>
- punto 1
- punto 2
- punto 3
---
TITULO: <otro título>
EXPLICACION: <explicación>
- punto 1
---

Reglas:
- Entre 4 y 6 diapositivas.
- 3 puntos por diapositiva, máximo 10 palabras cada uno.
- Sin asteriscos ni markdown."""

PROMPT_RESUMEN = """Resume este apunte de forma MUY breve y directa.

Formato exacto:
RESUMEN:
<un párrafo de 3-4 líneas>

CONCEPTOS:
- <concepto 1>
- <concepto 2>
- <concepto 3>
- <concepto 4>

PREGUNTAS:
1. <pregunta 1>
2. <pregunta 2>
3. <pregunta 3>

Reglas:
- Sé conciso. Total máximo: 200 palabras.
- No inventes información.
- Sin asteriscos ni markdown.
- Empieza directo con "RESUMEN:"."""


# ============ MOTOR DE IA ============

def _groq(imagen_b64, mime, prompt):
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
        "max_tokens": 900,
    }
    r = requests.post(URL_GROQ, headers=headers, json=payload, timeout=120)
    if r.status_code != 200:
        raise RuntimeError(f"Groq {r.status_code}: {r.text[:200]}")
    return r.json()["choices"][0]["message"]["content"]


def _gemini(imagen_b64, mime, prompt):
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


# ============ UNSPLASH ============

def buscar_imagen_unsplash(query):
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


# ============ GENERADORES DE DOCUMENTOS ============

def docx_bytes(texto, titulo="Documento generado por EscribIA"):
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
        elif up.startswith("CONCEPTOS CLAVE") or up.startswith("CONCEPTOS:"):
            doc.add_heading("Conceptos clave", level=1)
        elif up.startswith("CONCLUSION") or up.startswith("CONCLUSIÓN"):
            doc.add_heading("Conclusión", level=1)
        elif up.startswith("PREGUNTAS DE REPASO") or up.startswith("PREGUNTAS:"):
            doc.add_heading("Preguntas de repaso", level=1)
        elif up.startswith("RESUMEN:"):
            doc.add_heading("Resumen", level=1)
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

        ancho_texto = Inches(6.2) if tiene_imagen else Inches(12.1)

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


# ============ FUNCIONES DE BASE DE DATOS ============

def obtener_usos_mes(user_id, mes_actual):
    try:
        response = supabase_client.table("usage_logs") \
            .select("usage_count") \
            .eq("user_id", user_id) \
            .eq("month_year", mes_actual) \
            .execute()
        if response.data:
            return response.data[0]["usage_count"]
        return 0
    except Exception:
        return 0


def incrementar_uso(user_id, mes_actual, usos_actuales):
    try:
        if usos_actuales == 0:
            supabase_client.table("usage_logs").insert({
                "user_id": user_id,
                "month_year": mes_actual,
                "usage_count": 1,
            }).execute()
        else:
            supabase_client.table("usage_logs") \
                .update({"usage_count": usos_actuales + 1}) \
                .eq("user_id", user_id) \
                .eq("month_year", mes_actual) \
                .execute()
    except Exception as e:
        st.warning(f"No se pudo registrar el uso: {e}")


def guardar_conversion(user_id, source_type, texto):
    try:
        supabase_client.table("conversions").insert({
            "user_id": user_id,
            "source_type": source_type,
            "generated_text": texto[:5000],
        }).execute()
    except Exception:
        pass


# ============ CONFIGURACIÓN DE PÁGINA ============

st.set_page_config(page_title="EscribIA", page_icon="📝", layout="centered")


# ============ AUTENTICACIÓN ============

if "authenticated" not in st.session_state:
    st.session_state["authenticated"] = False

if not st.session_state.get("authenticated", False):
    st.title("📝 EscribIA")
    st.caption("Apuntes escritos a mano, convertidos en documentos digitales con IA")
    st.divider()

    # st-login-form lee SUPABASE_URL y SUPABASE_KEY de los Secrets
    # automáticamente. No hay que pasarle el cliente.
    login_form(
        title="Inicia sesión o crea tu cuenta",
        user_tablename="profiles",
    )

    if st.session_state.get("authenticated", False):
        st.rerun()
    st.stop()


# ============ USUARIO AUTENTICADO ============

user_id = st.session_state.get("user_id")
username = st.session_state.get("username", "Usuario")

with st.sidebar:
    st.markdown(f"### 👤 {username}")
    mes_actual = datetime.now().strftime("%Y-%m")
    usos = obtener_usos_mes(user_id, mes_actual)
    restantes = max(0, LIMITE_GRATUITO - usos)
    st.caption(f"Plan gratuito: {restantes} usos restantes este mes")
    st.progress(min(usos / LIMITE_GRATUITO, 1.0))
    st.divider()
    if st.button("🚪 Cerrar sesión", use_container_width=True):
        st.session_state["authenticated"] = False
        st.session_state.pop("user_id", None)
        st.session_state.pop("username", None)
        st.rerun()


# ============ INTERFAZ PRINCIPAL ============

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
        "Respeta el texto tal cual lo escribiste.",
        "La IA amplía y estructura en secciones.",
        "Diapositivas con imágenes automáticas.",
        "Resumen breve con conceptos y preguntas.",
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
        # --- Verificar límite de uso ---
        mes_actual = datetime.now().strftime("%Y-%m")
        usos_actuales = obtener_usos_mes(user_id, mes_actual)

        if usos_actuales >= LIMITE_GRATUITO:
            st.error(
                f"🚫 Has alcanzado tu límite de {LIMITE_GRATUITO} usos "
                f"gratuitos este mes. Actualiza a un plan de pago para seguir."
            )
            st.stop()

        # --- Preparar imagen ---
        img.thumbnail((1600, 1600))
        buf = io.BytesIO()
        img.convert("RGB").save(buf, "JPEG", quality=85)
        b64 = base64.b64encode(buf.getvalue()).decode()
        mime = "image/jpeg"

        # --- Elegir prompt ---
        if "Documento completo" in formato:
            prompt = PROMPT_COMPLETO
            source_type = "full_doc"
        elif "Solo lo anotado" in formato:
            prompt = PROMPT_WORD
            source_type = "word"
        elif "PowerPoint" in formato:
            prompt = PROMPT_PPT
            source_type = "ppt"
        else:
            prompt = PROMPT_RESUMEN
            source_type = "summary"

        # --- Procesar ---
        with st.spinner("Procesando con inteligencia artificial..."):
            try:
                texto, proveedor = procesar(b64, mime, prompt)
            except Exception as e:
                st.error(f"Error: {e}")
                st.stop()

        # --- Registrar uso y guardar en BD ---
        incrementar_uso(user_id, mes_actual, usos_actuales)
        guardar_conversion(user_id, source_type, texto)

        st.success(f"✅ Procesado con {proveedor}")
        st.subheader("Contenido generado")
        st.text_area("Texto", texto, height=300, label_visibility="collapsed")

        # --- Generar archivo ---
        if "Documento completo" in formato:
            datos = docx_bytes(texto, "Documento de estudio - EscribIA")
            nombre = f"escribia_documento_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
            mime_dl = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        elif "Solo lo anotado" in formato:
            datos = docx_bytes(texto, "Apunte transcrito - EscribIA")
            nombre = f"escribia_apunte_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
            mime_dl = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        elif "PowerPoint" in formato:
            if UNSPLASH_ACCESS_KEY:
                with st.spinner("Buscando imágenes..."):
                    datos = pptx_bytes(texto, con_imagenes=True)
            else:
                datos = pptx_bytes(texto, con_imagenes=False)
            nombre = f"escribia_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pptx"
            mime_dl = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
        else:
            datos = docx_bytes(texto, "Resumen de estudio - EscribIA")
            nombre = f"escribia_resumen_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
            mime_dl = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"

        st.download_button(
            "⬇️ Descargar archivo",
            data=datos,
            file_name=nombre,
            mime=mime_dl,
            type="primary",
            use_container_width=True,
              )
