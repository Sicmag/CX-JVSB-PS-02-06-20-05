import base64
import io
from datetime import datetime, timezone

import requests
import streamlit as st
from PIL import Image
from docx import Document
from pptx import Presentation
from pptx.util import Inches, Pt
from supabase import create_client
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill


# ============ CLAVES ============

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

if not GROQ_API_KEY and not GEMINI_API_KEY:
    st.error("Falta configurar claves de IA.")
    st.stop()
if not SUPABASE_URL or not SUPABASE_KEY:
    st.error("Falta configurar Supabase.")
    st.stop()

LIMITE_GRATUITO = 7


# ============ SUPABASE ============

def get_supabase():
    return create_client(SUPABASE_URL, SUPABASE_KEY)


supabase_client = get_supabase()


def restaurar_sesion():
    if st.session_state.get("access_token") and st.session_state.get("refresh_token"):
        try:
            supabase_client.auth.set_session(
                st.session_state["access_token"],
                st.session_state["refresh_token"],
            )
        except Exception:
            pass


restaurar_sesion()


# ============ CONFIG IA ============

MODELO_GEMINI = "gemini-3.8-flash"
MODELO_GROQ = "qwen/qwen3.8-27b"

URL_GEMINI = f"https://generativelanguage.googleapis.com/v1beta/models/{MODELO_GEMINI}:generateContent"
URL_GROQ = "https://api.groq.com/openai/v1/chat/completions"


PROMPT_WORD = """Transcribe este apunte manuscrito en español con fidelidad.
Corrige ortografía, organiza en párrafos. Devuelve SOLO el texto, sin markdown."""

PROMPT_COMPLETO = """Crea un documento de estudio del apunte.

TITULO: <título>
INTRODUCCION:
<2 líneas>
DESARROLLO:
SECCION: <nombre>
<párrafo>
SECCION: <nombre>
<párrafo>
CONCEPTOS CLAVE:
- <concepto>
- <concepto>
CONCLUSION:
<1-2 líneas>
PREGUNTAS DE REPASO:
1. <pregunta>
2. <pregunta>

Máximo 500 palabras. Sin asteriscos."""

PROMPT_PPT = """Convierte el apunte en diapositivas.

TITULO: <título corto>
EXPLICACION: <1 frase>
- punto 1
- punto 2
- punto 3
---
TITULO: <otro título>
EXPLICACION: <explicación>
- punto 1
---

Entre 4 y 6 diapositivas. Sin asteriscos."""

PROMPT_RESUMEN = """Resume el apunte muy breve.

RESUMEN:
<párrafo de 3 líneas>
CONCEPTOS:
- <concepto>
- <concepto>
PREGUNTAS:
1. <pregunta>
2. <pregunta>

Máximo 200 palabras. Sin asteriscos."""

PROMPT_ECUACIONES = """Analiza este apunte manuscrito. Si contiene ecuaciones,
sistemas de ecuaciones o problemas matemáticos, resuélvelos paso a paso.

Devuelve SOLO esto, con este formato exacto:

TITULO: <título del tema>

ECUACIONES ORIGINALES:
<escribe las ecuaciones tal como aparecen en el apunte>

METODO:
<nombre del método: sustitución, igualación, reducción, determinantes, etc.>

PASOS:
1. <paso 1 detallado>
2. <paso 2 detallado>
3. <paso 3 detallado>
4. <los pasos necesarios>

SOLUCION:
<valores finales de las variables>

VERIFICACION:
<sustituye los valores en las ecuaciones originales y comprueba>

Si el apunte NO contiene ecuaciones, devuelve solo: SIN_ECUACIONES

Reglas:
- Sé claro y didáctico.
- No inventes datos que no estén en el apunte.
- Sin asteriscos ni markdown."""

PROMPT_EXCEL = """Analiza este apunte manuscrito. Extrae TODOS los datos que
estén en formato tabular: tablas, listas con columnas, datos numéricos
organizados, inventarios, calificaciones, presupuestos, etc.

Devuelve SOLO los datos en formato CSV, con la primera fila como encabezados.

Ejemplo:
Producto,Cantidad,Precio
Manzanas,10,2500
Peras,5,3000

Reglas:
- Separador: coma
- Sin comillas alrededor de los valores
- Sin líneas vacías
- Si un dato no está claro, déjalo vacío
- Si NO hay datos tabulares, devuelve solo: SIN_DATOS
- NO agregues explicaciones ni texto adicional, solo el CSV"""


# ============ IA ============

def _groq(b64, mime, prompt):
    headers = {"Authorization": f"Bearer {GROQ_API_KEY}"}
    payload = {
        "model": MODELO_GROQ,
        "messages": [{"role": "user", "content": [
            {"type": "text", "text": prompt},
            {"type": "image_url", "image_url": {"url": f"data:{mime};base64,{b64}"}},
        ]}],
        "temperature": 0.2,
        "max_tokens": 900,
    }
    r = requests.post(URL_GROQ, headers=headers, json=payload, timeout=120)
    if r.status_code != 200:
        raise RuntimeError(f"Groq {r.status_code}")
    return r.json()["choices"][0]["message"]["content"]


def _gemini(b64, mime, prompt):
    url = f"{URL_GEMINI}?key={GEMINI_API_KEY}"
    payload = {"contents": [{"parts": [
        {"text": prompt},
        {"inline_data": {"mime_type": mime, "data": b64}},
    ]}]}
    r = requests.post(url, json=payload, timeout=120)
    if r.status_code != 200:
        raise RuntimeError(f"Gemini {r.status_code}")
    return r.json()["candidates"][0]["content"]["parts"][0]["text"]


def procesar(b64, mime, prompt):
    errores = []
    if GROQ_API_KEY:
        try:
            return _groq(b64, mime, prompt), "Groq"
        except Exception as e:
            errores.append(f"Groq: {e}")
    if GEMINI_API_KEY:
        try:
            return _gemini(b64, mime, prompt), "Gemini"
        except Exception as e:
            errores.append(f"Gemini: {e}")
    raise RuntimeError(" | ".join(errores))


# ============ UNSPLASH ============

def buscar_imagen(query):
    if not UNSPLASH_ACCESS_KEY:
        return None
    try:
        r = requests.get(
            "https://api.unsplash.com/search/photos",
            params={"query": query, "per_page": 1, "orientation": "landscape",
                    "client_id": UNSPLASH_ACCESS_KEY},
            timeout=15,
        )
        if r.status_code == 200:
            res = r.json().get("results", [])
            if res:
                return res[0]["urls"]["regular"]
    except Exception:
        pass
    return None


# ============ DOCX ============

def docx_bytes(texto, titulo="Documento EscribIA"):
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
        elif up.startswith("CONCEPTOS"):
            doc.add_heading("Conceptos clave", level=1)
        elif up.startswith("CONCLUSION") or up.startswith("CONCLUSIÓN"):
            doc.add_heading("Conclusión", level=1)
        elif up.startswith("PREGUNTAS"):
            doc.add_heading("Preguntas", level=1)
        elif up.startswith("RESUMEN:"):
            doc.add_heading("Resumen", level=1)
        elif up.startswith("ECUACIONES ORIGINALES"):
            doc.add_heading("Ecuaciones originales", level=1)
        elif up.startswith("METODO") or up.startswith("MÉTODO"):
            doc.add_heading("Método", level=1)
        elif up.startswith("PASOS:"):
            doc.add_heading("Pasos", level=1)
        elif up.startswith("SOLUCION") or up.startswith("SOLUCIÓN"):
            doc.add_heading("Solución", level=1)
        elif up.startswith("VERIFICACION") or up.startswith("VERIFICACIÓN"):
            doc.add_heading("Verificación", level=1)
        elif l.startswith("-"):
            doc.add_paragraph(l[1:].strip(), style="List Bullet")
        elif l[:2].strip().rstrip(".").isdigit():
            doc.add_paragraph(l, style="List Number")
        else:
            doc.add_paragraph(l)
    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


# ============ EXCEL ============

def excel_bytes(csv_texto):
    import csv
    from io import StringIO

    if "SIN_DATOS" in csv_texto.upper():
        raise ValueError("No se detectaron datos tabulares en el apunte.")

    # Limpiar posibles bloques de código markdown
    limpio = csv_texto.strip()
    if limpio.startswith("```"):
        lineas = limpio.split("\n")
        lineas = [l for l in lineas if not l.strip().startswith("```")]
        limpio = "\n".join(lineas)

    wb = Workbook()
    ws = wb.active
    ws.title = "Datos"

    reader = csv.reader(StringIO(limpio))
    filas_validas = 0
    for fila in reader:
        if fila and any(c.strip() for c in fila):
            ws.append(fila)
            filas_validas += 1

    if filas_validas == 0:
        raise ValueError("El CSV no contiene filas de datos válidas.")

    # Encabezados con estilo
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E78")

    # Ancho automático de columnas
    for col in ws.columns:
        max_len = 0
        letra = col[0].column_letter
        for cell in col:
            try:
                if cell.value is not None:
                    max_len = max(max_len, len(str(cell.value)))
            except Exception:
                pass
        ws.column_dimensions[letra].width = min(max_len + 4, 40)

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ============ PPTX ============

def parsear_slides(texto):
    slides = []
    for bloque in texto.split("---"):
        lineas = [l.strip() for l in bloque.strip().split("\n") if l.strip()]
        if not lineas:
            continue
        titulo, explicacion, puntos = "", "", []
        for linea in lineas:
            up = linea.upper()
            if up.startswith("TITULO:"):
                titulo = linea.split(":", 1)[1].strip()
            elif up.startswith("EXPLICACION:"):
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
    slides = parsear_slides(texto)
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    portada = prs.slides.add_slide(prs.slide_layouts[0])
    portada.shapes.title.text = "Apuntes digitalizados"
    portada.placeholders[1].text = f"EscribIA - {datetime.now().strftime('%d/%m/%Y')}"

    for titulo, explicacion, puntos in slides:
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        slide.shapes.title.text = titulo

        tiene_img = False
        if UNSPLASH_ACCESS_KEY and titulo:
            url = buscar_imagen(titulo)
            if url:
                try:
                    data = requests.get(url, timeout=15).content
                    slide.shapes.add_picture(io.BytesIO(data), left=Inches(7.0),
                                              top=Inches(1.8), width=Inches(5.8))
                    tiene_img = True
                except Exception:
                    pass

        ancho = Inches(6.2) if tiene_img else Inches(12.1)

        if explicacion:
            caja = slide.shapes.add_textbox(Inches(0.6), Inches(1.6), ancho, Inches(1.5))
            caja.text_frame.word_wrap = True
            p = caja.text_frame.paragraphs[0]
            p.text = explicacion
            p.font.size = Pt(16)
            p.font.italic = True

        if puntos:
            top = Inches(3.4) if explicacion else Inches(2.0)
            caja2 = slide.shapes.add_textbox(Inches(0.6), top, ancho, Inches(4.0))
            caja2.text_frame.word_wrap = True
            for i, pt in enumerate(puntos):
                if i == 0:
                    par = caja2.text_frame.paragraphs[0]
                else:
                    par = caja2.text_frame.add_paragraph()
                par.text = f"•  {pt}"
                par.font.size = Pt(18)

    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()


# ============ BD ============

def obtener_usos(uid, mes):
    restaurar_sesion()
    try:
        r = supabase_client.table("usage_logs").select("usage_count") \
            .eq("user_id", uid).eq("month_year", mes).execute()
        return r.data[0]["usage_count"] if r.data else 0
    except Exception:
        return 0


def sumar_uso(uid, mes, actual):
    restaurar_sesion()
    try:
        if actual == 0:
            supabase_client.table("usage_logs").insert({
                "user_id": uid, "month_year": mes, "usage_count": 1
            }).execute()
        else:
            supabase_client.table("usage_logs").update(
                {"usage_count": actual + 1}
            ).eq("user_id", uid).eq("month_year", mes).execute()
    except Exception:
        pass


def guardar_conv(uid, tipo, texto):
    restaurar_sesion()
    try:
        supabase_client.table("conversions").insert({
            "user_id": uid, "source_type": tipo, "generated_text": texto[:5000]
        }).execute()
    except Exception:
        pass


def obtener_perfil(uid):
    restaurar_sesion()
    try:
        r = supabase_client.table("profiles") \
            .select("premium_until, premium_plan") \
            .eq("id", uid).execute()
        return r.data[0] if r.data else {}
    except Exception:
        return {}


def es_premium(perfil):
    if not perfil or not perfil.get("premium_until"):
        return False
    try:
        fecha = datetime.fromisoformat(perfil["premium_until"].replace("Z", "+00:00"))
        return fecha > datetime.now(timezone.utc)
    except Exception:
        return False


def activar_codigo(code):
    restaurar_sesion()
    try:
        resp = supabase_client.rpc("activar_codigo", {"p_code": code}).execute()
        return resp.data
    except Exception as e:
        return {"success": False, "message": str(e)}


# ============ ESTADO ============

st.set_page_config(page_title="EscribIA", page_icon="📝", layout="centered")

if "user" not in st.session_state:
    st.session_state["user"] = None
if "usos_cache" not in st.session_state:
    st.session_state["usos_cache"] = None
if "perfil_cache" not in st.session_state:
    st.session_state["perfil_cache"] = None


# ============ LOGIN ============

if st.session_state["user"] is None:
    st.title("📝 EscribIA")
    st.caption("Apuntes escritos a mano, convertidos en documentos digitales con IA")
    st.divider()

    tab1, tab2 = st.tabs(["🔐 Iniciar sesión", "✨ Crear cuenta"])

    with tab1:
        with st.form("login"):
            email = st.text_input("Correo")
            pwd = st.text_input("Contraseña", type="password")
            ok = st.form_submit_button("Iniciar sesión", use_container_width=True)
        if ok:
            try:
                resp = supabase_client.auth.sign_in_with_password(
                    {"email": email, "password": pwd}
                )
                if resp.user:
                    st.session_state["user"] = {
                        "id": resp.user.id,
                        "email": resp.user.email,
                    }
                    st.session_state["access_token"] = resp.session.access_token
                    st.session_state["refresh_token"] = resp.session.refresh_token
                    st.session_state["usos_cache"] = None
                    st.session_state["perfil_cache"] = None
                    st.rerun()
            except Exception as e:
                st.error(f"Error: {e}")

    with tab2:
        st.caption("Mínimo 8 caracteres con mayúscula, minúscula, número y un símbolo.")
        with st.form("registro"):
            email_r = st.text_input("Correo", key="r_email")
            pwd_r = st.text_input("Contraseña", type="password", key="r_pwd")
            pwd_r2 = st.text_input("Repite contraseña", type="password", key="r_pwd2")
            ok_r = st.form_submit_button("Crear cuenta", use_container_width=True)
        if ok_r:
            if pwd_r != pwd_r2:
                st.error("Las contraseñas no coinciden.")
            elif len(pwd_r) < 8:
                st.error("Mínimo 8 caracteres.")
            else:
                try:
                    resp = supabase_client.auth.sign_up(
                        {"email": email_r, "password": pwd_r}
                    )
                    if resp.user:
                        st.success("✅ Cuenta creada. Revisa tu correo para confirmar.")
                except Exception as e:
                    st.error(f"Error: {e}")
    st.stop()


# ============ APP PRINCIPAL ============

uid = st.session_state["user"]["id"]
email = st.session_state["user"]["email"]
nombre = email.split("@")[0] if email else "Usuario"
mes_actual = datetime.now().strftime("%Y-%m")

if st.session_state["perfil_cache"] is None:
    st.session_state["perfil_cache"] = obtener_perfil(uid)
if st.session_state["usos_cache"] is None:
    st.session_state["usos_cache"] = obtener_usos(uid, mes_actual)

perfil = st.session_state["perfil_cache"]
premium_activo = es_premium(perfil)
usos = st.session_state["usos_cache"]


with st.sidebar:
    st.markdown(f"### 👤 {nombre}")
    st.caption(email)

    if premium_activo:
        try:
            fecha = datetime.fromisoformat(perfil["premium_until"].replace("Z", "+00:00"))
            st.success(f"⭐ **Premium activo**\n\nVence el {fecha.strftime('%d/%m/%Y')}")
        except Exception:
            st.success("⭐ **Premium activo**")
        st.caption("Uso ilimitado este mes")
    else:
        restantes = max(0, LIMITE_GRATUITO - usos)
        st.caption(f"Plan gratuito: {restantes}/{LIMITE_GRATUITO} usos")
        st.progress(min(usos / LIMITE_GRATUITO, 1.0))

    st.divider()

    if not premium_activo:
        with st.expander("⭐ Activar Premium"):
            st.caption("Ingresa tu código de activación")
            codigo_input = st.text_input("Código", key="codigo_premium",
                                          placeholder="ESCRIBIA-XXXX-XXXX")
            if st.button("Activar", use_container_width=True, type="primary"):
                if codigo_input:
                    with st.spinner("Validando..."):
                        resultado = activar_codigo(codigo_input)
                    if resultado and resultado.get("success"):
                        st.success(resultado.get("message", "¡Premium activado!"))
                        st.session_state["perfil_cache"] = None
                        st.session_state["usos_cache"] = None
                        st.balloons()
                        st.rerun()
                    else:
                        msg = resultado.get("message", "Código inválido") if resultado else "Código inválido"
                        st.error(msg)
                else:
                    st.warning("Escribe un código.")

    st.divider()
    if st.button("🚪 Cerrar sesión", use_container_width=True):
        try:
            supabase_client.auth.sign_out()
        except Exception:
            pass
        for k in ["user", "access_token", "refresh_token", "usos_cache", "perfil_cache"]:
            st.session_state.pop(k, None)
        st.rerun()


st.title("📝 EscribIA")
st.caption("Apuntes escritos a mano, convertidos en documentos digitales con IA")

if premium_activo:
    st.info("⭐ Tienes **Premium activo**. Uso ilimitado.")

st.divider()

formato = st.radio(
    "¿Qué quieres generar?",
    [
        "📝 Solo lo anotado",
        "📚 Documento completo de estudio",
        
