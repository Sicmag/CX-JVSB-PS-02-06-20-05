"""
EscribIA Admin - Generador de códigos premium
Panel privado para generar, ver y exportar códigos de activación.
Permite definir duración en días (1-365) y horas (0-23).
"""

import random
import string
from datetime import datetime

import pandas as pd
import streamlit as st
from supabase import create_client


# ============ CONFIGURACIÓN DE PÁGINA ============

st.set_page_config(page_title="EscribIA Admin", page_icon="🔑", layout="wide")


# ============ CSS ============

st.markdown("""
<style>
    .main-header {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 2rem;
        border-radius: 16px;
        color: white;
        margin-bottom: 2rem;
    }
    .main-header h1 { color: white; margin: 0; font-size: 2.5rem; }
    .main-header p { color: rgba(255,255,255,0.9); margin: 0.5rem 0 0 0; }
    .metric-card {
        background: #f9fafb;
        border-radius: 12px;
        padding: 1.5rem;
        box-shadow: 0 2px 12px rgba(0,0,0,0.08);
        border-left: 4px solid #667eea;
    }
    .metric-value { font-size: 2.2rem; font-weight: 700; color: #1a1a2e; margin: 0; }
    .metric-label { color: #6b7280; font-size: 0.85rem; text-transform: uppercase; letter-spacing: 0.5px; }
    .success-banner {
        background: linear-gradient(90deg, #10b981, #34d399);
        color: white;
        padding: 1rem 1.5rem;
        border-radius: 10px;
        margin: 1rem 0;
        font-weight: 600;
    }
    .preview-box {
        background: #eef2ff;
        border: 2px solid #667eea;
        border-radius: 10px;
        padding: 1rem 1.25rem;
        color: #3730a3;
        font-weight: 600;
        margin: 0.5rem 0 1rem 0;
    }
    .stButton > button { border-radius: 8px; font-weight: 600; }
</style>
""", unsafe_allow_html=True)


# ============ LECTURA DE SECRETS ============

def _leer(nombre, default=""):
    try:
        return st.secrets[nombre]
    except Exception:
        return default


ADMIN_PASSWORD = _leer("ADMIN_PASSWORD")
SUPABASE_URL = _leer("SUPABASE_URL")
SUPABASE_SERVICE_KEY = _leer("SUPABASE_SERVICE_KEY")

if not SUPABASE_URL or not SUPABASE_SERVICE_KEY:
    st.error("⚠️ Faltan SUPABASE_URL o SUPABASE_SERVICE_KEY en los Secrets.")
    st.stop()

if not ADMIN_PASSWORD:
    st.error("⚠️ Falta configurar ADMIN_PASSWORD en los Secrets.")
    st.stop()


# ============ LOGIN ADMIN ============

if "admin_ok" not in st.session_state:
    st.session_state["admin_ok"] = False

if not st.session_state["admin_ok"]:
    st.markdown("""
    <div class="main-header">
        <h1>🔑 EscribIA Admin</h1>
        <p>Panel privado de generación de códigos premium</p>
    </div>
    """, unsafe_allow_html=True)

    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        with st.form("admin_login"):
            pwd = st.text_input("Contraseña de administrador", type="password")
            ok = st.form_submit_button("Entrar", use_container_width=True, type="primary")
        if ok:
            if pwd == ADMIN_PASSWORD:
                st.session_state["admin_ok"] = True
                st.rerun()
            else:
                st.error("Contraseña incorrecta")
    st.stop()


# ============ CLIENTE SUPABASE ============

@st.cache_resource
def get_admin_client():
    return create_client(SUPABASE_URL, SUPABASE_SERVICE_KEY)


try:
    admin_client = get_admin_client()
except Exception as e:
    st.error(f"Error conectando a Supabase: {e}")
    st.stop()


# ============ HELPERS ============

def generar_codigo():
    chars = string.ascii_uppercase + string.digits
    p1 = "".join(random.choices(chars, k=4))
    p2 = "".join(random.choices(chars, k=4))
    return f"ESCRIBIA-{p1}-{p2}"


def nombre_plan(dias, horas):
    """Genera el identificador del plan según días y horas."""
    if horas > 0:
        return f"premium_{dias}d_{horas}h"
    return f"premium_{dias}d"


def texto_duracion(dias, horas):
    """Texto legible: '2 días y 5 horas'."""
    partes = []
    if dias > 0:
        partes.append(f"{dias} día{'s' if dias != 1 else ''}")
    if horas > 0:
        partes.append(f"{horas} hora{'s' if horas != 1 else ''}")
    return " y ".join(partes) if partes else "0"


def generar_lote(cantidad, dias, horas):
    codigos = []
    intentos = 0
    while len(codigos) < cantidad and intentos < cantidad * 3:
        codigos.append(generar_codigo())
        intentos += 1

    plan = nombre_plan(dias, horas)
    data = [{
        "code": c,
        "plan_type": plan,
        "duration_days": dias,
        "duration_hours": horas,
    } for c in codigos]

    try:
        resp = admin_client.table("activation_codes").insert(data).execute()
        return resp.data or [], None
    except Exception as e:
        return [], str(e)


def obtener_codigos():
    try:
        resp = admin_client.table("activation_codes").select("*").execute()
        return resp.data or []
    except Exception:
        return []


# ============ HEADER ============

st.markdown("""
<div class="main-header">
    <h1>🔑 EscribIA Admin</h1>
    <p>Genera y gestiona los códigos de activación premium</p>
</div>
""", unsafe_allow_html=True)


# ============ ESTADÍSTICAS ============

codigos = obtener_codigos()
total = len(codigos)
usados = sum(1 for c in codigos if c.get("used"))
disponibles = total - usados
porcentaje = (usados / total * 100) if total > 0 else 0

col1, col2, col3, col4 = st.columns(4)
with col1:
    st.markdown(f"""
    <div class="metric-card">
        <p class="metric-label">Total generados</p>
        <p class="metric-value">{total}</p>
    </div>
    """, unsafe_allow_html=True)
with col2:
    st.markdown(f"""
    <div class="metric-card" style="border-left-color:#10b981;">
        <p class="metric-label">Disponibles</p>
        <p class="metric-value" style="color:#10b981;">{disponibles}</p>
    </div>
    """, unsafe_allow_html=True)
with col3:
    st.markdown(f"""
    <div class="metric-card" style="border-left-color:#ef4444;">
        <p class="metric-label">Usados</p>
        <p class="metric-value" style="color:#ef4444;">{usados}</p>
    </div>
    """, unsafe_allow_html=True)
with col4:
    st.markdown(f"""
    <div class="metric-card" style="border-left-color:#f59e0b;">
        <p class="metric-label">Tasa de uso</p>
        <p class="metric-value" style="color:#f59e0b;">{porcentaje:.1f}%</p>
    </div>
    """, unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)


# ============ GENERADOR ============

st.subheader("🎯 Generar nuevo lote de códigos")
st.caption("Configura la duración exacta del premium: días y/o horas.")

col_izq, col_der = st.columns([2, 1])

with col_izq:
    with st.form("generar"):
        c1, c2, c3 = st.columns(3)
        with c1:
            cantidad = st.number_input(
                "Cantidad de códigos",
                min_value=1, max_value=200, value=5, step=1,
            )
        with c2:
            dias = st.number_input(
                "Días (1 - 365)",
                min_value=1, max_value=365, value=30, step=1,
            )
        with c3:
            horas = st.number_input(
                "Horas (0 - 23)",
                min_value=0, max_value=23, value=0, step=1,
                help="0 = solo días. Suma horas extra al vencimiento.",
            )

        st.markdown(
            f'<div class="preview-box">⏱️ Duración por código: {texto_duracion(dias, horas)}</div>',
            unsafe_allow_html=True,
        )

        generar = st.form_submit_button(
            "⚡ Generar códigos",
            use_container_width=True,
            type="primary",
        )

    if generar:
        with st.spinner(f"Generando {cantidad} códigos..."):
            nuevos, error = generar_lote(cantidad, int(dias), int(horas))
        if error:
            st.error(f"Error: {error}")
        elif nuevos:
            st.markdown(
                f'<div class="success-banner">✅ {len(nuevos)} códigos generados ({texto_duracion(dias, horas)})</div>',
                unsafe_allow_html=True,
            )
            st.markdown("**Primeros códigos generados:**")
            for c in nuevos[:5]:
                st.code(c["code"], language=None)
            if len(nuevos) > 5:
                st.caption(f"... y {len(nuevos)-5} más")

            df_nuevos = pd.DataFrame(nuevos)[
                ["code", "plan_type", "duration_days", "duration_hours"]
            ]
            csv = df_nuevos.to_csv(index=False).encode("utf-8")
            st.download_button(
                "📥 Descargar lote como CSV",
                data=csv,
                file_name=f"codigos_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                mime="text/csv",
                use_container_width=True,
            )

with col_der:
    st.markdown("""
    <div style="background:#f9fafb;padding:1rem;border-radius:10px;">
    <p style="font-weight:700;margin:0 0 0.5rem 0;">💡 Ejemplos de uso</p>
    <p style="font-size:0.85rem;color:#4b5563;margin:0 0 0.5rem 0;">
    • <b>1 día, 0 horas</b> → prueba rápida<br>
    • <b>7 días, 0 horas</b> → semana<br>
    • <b>30 días, 0 horas</b> → mes<br>
    • <b>365 días, 0 horas</b> → año completo<br>
    • <b>0? no, mínimo 1 día</b> — usa horas si quieres menos<br>
    • <b>1 día, 12 horas</b> → día y medio
    </p>
    <p style="font-weight:700;margin:0.75rem 0 0.5rem 0;">💰 Precios sugeridos</p>
    <p style="font-size:0.85rem;color:#4b5563;margin:0;">
    • 1 hora → $1.000 COP<br>
    • 1 día → $3.900 COP<br>
    • 30 días → $19.900 COP<br>
    • 365 días → $149.000 COP
    </p>
    </div>
    """, unsafe_allow_html=True)


st.divider()


# ============ LISTADO ============

st.subheader("📋 Códigos existentes")

if not codigos:
    st.info("Aún no has generado ningún código.")
else:
    filtro = st.radio(
        "Filtrar:", ["Todos", "Solo disponibles", "Solo usados"], horizontal=True
    )

    lista = codigos
    if filtro == "Solo disponibles":
        lista = [c for c in codigos if not c.get("used")]
    elif filtro == "Solo usados":
        lista = [c for c in codigos if c.get("used")]

    lista = sorted(lista, key=lambda x: x.get("created_at") or "", reverse=True)

    def duracion_legible(c):
        d = c.get("duration_days", 0) or 0
        h = c.get("duration_hours", 0) or 0
        return texto_duracion(d, h)

    df = pd.DataFrame([{
        "Código": c["code"],
        "Duración": duracion_legible(c),
        "Estado": "✅ Usado" if c.get("used") else "🟢 Disponible",
        "Creado": (c.get("created_at") or "")[:10],
        "Usado el": (c.get("used_at") or "")[:10] if c.get("used") else "—",
    } for c in lista])

    st.dataframe(df, use_container_width=True, hide_index=True)

    csv_all = df.to_csv(index=False).encode("utf-8")
    st.download_button(
        "📥 Exportar TODOS a CSV",
        data=csv_all,
        file_name=f"codigos_completo_{datetime.now().strftime('%Y%m%d')}.csv",
        mime="text/csv",
    )


st.divider()
if st.button("🚪 Cerrar sesión admin"):
    st.session_state["admin_ok"] = False
    st.rerun()
