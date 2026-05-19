import streamlit as st

from pages.visitas       import page_agregar_seller, page_edicion_visitas, page_eliminacion_visitas, page_visit_types_skills
from pages.rutas         import page_flotas, page_iniciar_cerrar_rutas, page_zonas
from pages.gps           import page_analisis_gps, page_validacion_gps
from pages.configuracion import page_cambiar_rol, page_configurar_addons, page_copiar_configs, page_crear_webhook, page_desbloqueo, page_permisos_usuario, page_reenviar_webhooks
from pages.tms           import page_tms_document_types, page_tms_transportistas

APP_VERSION = "2.0.0"

st.set_page_config(page_title="SimpliRoute Tools", page_icon="🚀", layout="wide")

# ── NAVEGACIÓN ────────────────────────────────────────────────────────────────
MENU_GROUPS = [
    ("📋 Visitas", [
        "🧑‍💼 Agregar Seller a Visitas",
        "✏️ Edición de Visitas",
        "🗑️ Eliminación Masiva de Visitas",
        "🏷️ Tipos de Visita y Skills",
    ]),
    ("🚗 Rutas & Flotas", [
        "🚛 Flotas",
        "🚦 Iniciar / Cerrar Rutas",
        "🗺️ Zonas",
    ]),
    ("📡 GPS", [
        "📍 Análisis de Recorrido GPS",
        "📡 Validación de GPS",
    ]),
    ("⚙️ Configuración", [
        "👤 Cambiar Rol de Usuario",
        "⚙️ Configurar Addons",
        "⚙️ Copiar Configs",
        "🔔 Crear Webhook",
        "🔓 Desbloqueo de Contraseña",
        "🔍 Permisos de Usuario",
        "🔁 Reenviar Webhooks",
    ]),
    ("📦 TMS", [
        "📄 Tipos de Documento",
        "🚚 Transportistas",
    ]),
]

if "current_page" not in st.session_state:
    st.session_state["current_page"] = MENU_GROUPS[0][1][0]

selected = st.session_state["current_page"]


def nav_item(label):
    is_active = st.session_state["current_page"] == label
    if st.button(label, key=f"nav_{label}", use_container_width=True,
                 type="primary" if is_active else "secondary"):
        for key in list(st.session_state.keys()):
            if key != "current_page":
                del st.session_state[key]
        st.session_state["current_page"] = label
        st.rerun()


with st.sidebar:
    st.markdown("## 🚀 SimpliRoute Tools")
    st.markdown("---")
    for group_name, items in MENU_GROUPS:
        is_group_active = selected in items
        with st.expander(group_name, expanded=is_group_active):
            for item in items:
                nav_item(item)
    st.markdown("---")
    st.link_button("🔧 Tools Julio", "https://simpliroute-tools.streamlit.app/", use_container_width=True)
    st.markdown("---")
    st.caption(f"SimpliRoute Internal Tools v{APP_VERSION}")


# ── ROUTER ────────────────────────────────────────────────────────────────────
PAGES = {
    "🧑‍💼 Agregar Seller a Visitas":     page_agregar_seller,
    "✏️ Edición de Visitas":            page_edicion_visitas,
    "🗑️ Eliminación Masiva de Visitas": page_eliminacion_visitas,
    "🏷️ Tipos de Visita y Skills":      page_visit_types_skills,
    "🚛 Flotas":                        page_flotas,
    "🚦 Iniciar / Cerrar Rutas":        page_iniciar_cerrar_rutas,
    "🗺️ Zonas":                         page_zonas,
    "📍 Análisis de Recorrido GPS":     page_analisis_gps,
    "📡 Validación de GPS":             page_validacion_gps,
    "👤 Cambiar Rol de Usuario":         page_cambiar_rol,
    "⚙️ Configurar Addons":             page_configurar_addons,
    "⚙️ Copiar Configs":                page_copiar_configs,
    "🔔 Crear Webhook":                 page_crear_webhook,
    "🔓 Desbloqueo de Contraseña":      page_desbloqueo,
    "🔍 Permisos de Usuario":           page_permisos_usuario,
    "🔁 Reenviar Webhooks":             page_reenviar_webhooks,
    "📄 Tipos de Documento":            page_tms_document_types,
    "🚚 Transportistas":                page_tms_transportistas,
}

page_fn = PAGES.get(selected)
if page_fn:
    page_fn()
