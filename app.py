import streamlit as st
import openpyxl
import io
import re
import requests

st.set_page_config(
    page_title="SimpliRoute Tools",
    page_icon="🚀",
    layout="wide"
)

# --- Sidebar ---
with st.sidebar:
    st.markdown("## 🚀 SimpliRoute Tools")
    st.markdown("---")
    st.markdown("### 🛠️ Herramientas")
    menu = st.radio(
        label="",
        options=[
            "🗺️ Cargar Zonas",
            "👤 Cambiar Rol de Usuario",
        ],
        label_visibility="collapsed"
    )
    st.markdown("---")
    st.caption("SimpliRoute Internal Tools v1.0")


# ─────────────────────────────────────────
# FEATURE 1: CARGAR ZONAS
# ─────────────────────────────────────────
def decode_file(raw_bytes):
    for enc in ["utf-8", "latin-1", "cp1252"]:
        try:
            return raw_bytes.decode(enc, errors="ignore")
        except Exception:
            continue
    return raw_bytes.decode("utf-8", errors="ignore")

def strip_rtf_codes(text):
    text = re.sub(r"\\'[0-9a-fA-F]{2}", "", text)
    text = re.sub(r"\\[a-zA-Z]+\d*\s?", " ", text)
    text = re.sub(r"[{}]", "", text)
    text = re.sub(r" {2,}", " ", text)
    return text

def parse_polygons(text):
    polygons = []
    placemark_blocks = re.findall(r'<Placemark>(.*?)</Placemark>', text, re.DOTALL)
    for block in placemark_blocks:
        if '<Polygon>' not in block and '<Polygon ' not in block:
            continue
        name_match = re.search(r'<name>\s*(.*?)\s*</name>', block)
        name = name_match.group(1).strip() if name_match else "Sin nombre"
        coords_match = re.search(r'<coordinates>\s*(.*?)\s*</coordinates>', block, re.DOTALL)
        if not coords_match:
            continue
        raw_coords = coords_match.group(1).strip()
        points = []
        for token_str in raw_coords.split():
            parts = token_str.strip().split(",")
            if len(parts) < 2:
                continue
            try:
                a, b = float(parts[0]), float(parts[1])
            except ValueError:
                continue
            lng, lat = a, b
            if abs(lat) > 90:
                lng, lat = lat, lng
            if not (-90 <= lat <= 90 and -180 <= lng <= 180):
                continue
            points.append((lat, lng))
        if points:
            polygons.append({"name": name, "coords": points})
    return polygons

def coords_to_str(coords):
    parts = [f"{{'lat':'{lat:.6f}','lng':'{lng:.6f}'}}" for lat, lng in coords]
    return "[" + ",".join(parts) + "]"

def generate_excel(polygons):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Hoja 1"
    ws.append(["is_name", "is_coordinates"])
    for p in polygons:
        ws.append([p["name"], coords_to_str(p["coords"])])
    ws.column_dimensions["A"].width = 25
    ws.column_dimensions["B"].width = 80
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

def upload_zone(name, coordinates, auth_token):
    url = "http://api.simpliroute.com/v1/zones/"
    headers = {"authorization": f"Token {auth_token}", "content-type": "application/json"}
    payload = {"name": name, "coordinates": coordinates, "vehicles": []}
    try:
        r = requests.post(url, headers=headers, json=payload, timeout=15)
        return r.status_code, r.json()
    except Exception as e:
        return None, str(e)

def page_cargar_zonas():
    st.title("🗺️ Cargar Zonas")
    st.markdown("Sube tu archivo KML o RTF, ingresa tu token y carga las zonas directo a SimpliRoute.")

    token = st.text_input("🔑 Token de SimpliRoute", type="password", placeholder="Ingresa tu token aquí")
    uploaded = st.file_uploader("📂 Sube tu archivo KML o RTF", type=["kml", "rtf", "txt"])

    if uploaded:
        content = uploaded.read()
        with st.spinner("Procesando archivo..."):
            try:
                text = decode_file(content)
                if text.strip().startswith("{\\rtf") or "\\rtf" in text[:100]:
                    text = strip_rtf_codes(text)
                polygons = parse_polygons(text)
                error = None
            except Exception as e:
                polygons = []
                error = str(e)

        if error:
            st.error(f"Error: {error}")
        elif polygons:
            st.success(f"✅ {len(polygons)} polígono(s) encontrado(s)")
            st.subheader("Preview")
            for p in polygons:
                with st.expander(f"📍 {p['name']} — {len(p['coords'])} puntos"):
                    st.code(coords_to_str(p["coords"][:3]) + ",...]", language=None)

            excel_buf = generate_excel(polygons)
            st.download_button(
                label="⬇️ Descargar Excel",
                data=excel_buf,
                file_name="ZONES.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            st.divider()
            st.subheader("📡 Cargar zonas a SimpliRoute")
            if not token:
                st.warning("⚠️ Ingresa tu token de SimpliRoute arriba para poder cargar las zonas.")
            else:
                if st.button("🚀 Cargar zonas a SimpliRoute", type="primary"):
                    results = []
                    progress = st.progress(0)
                    status_box = st.empty()
                    for i, p in enumerate(polygons):
                        status_box.info(f"Cargando: **{p['name']}** ({i+1}/{len(polygons)})")
                        code, resp = upload_zone(p["name"], coords_to_str(p["coords"]), token)
                        results.append({"name": p["name"], "code": code, "resp": resp})
                        progress.progress((i + 1) / len(polygons))
                    status_box.empty()
                    progress.empty()

                    ok = sum(1 for r in results if r["code"] in [200, 201])
                    st.markdown(f"✅ **{ok} zonas cargadas** | ❌ **{len(results)-ok} con error**")
                    for r in results:
                        if r["code"] in [200, 201]:
                            st.success(f"✅ {r['name']} — Zona creada correctamente")
                        elif r["code"] == 400:
                            st.warning(f"⚠️ {r['name']} — Ya existe una zona con este nombre")
                        elif r["code"] == 401:
                            st.error(f"❌ {r['name']} — Token inválido o sin permisos")
                        elif r["code"] is None:
                            st.error(f"❌ {r['name']} — Sin conexión: {r['resp']}")
                        else:
                            st.error(f"❌ {r['name']} — Error {r['code']}: {r['resp']}")
        else:
            st.warning("No se encontraron polígonos en el archivo.")


# ─────────────────────────────────────────
# FEATURE 2: CAMBIAR ROL DE USUARIO
# ─────────────────────────────────────────
ROLES = {
    "Administrador": "is_admin",
    "Conductor": "is_driver",
    "Co-conductor": "is_codriver",
    "Router Jr": "is_router_jr",
    "Monitor": "is_monitor",
    "Coordinador": "is_coordinator",
    "Router": "is_router",
    "Staff": "is_staff",
    "Seller Viewer": "is_seller_viewer",
    "Seller": "is_seller",
}

ALL_ROLE_KEYS = list(ROLES.values()) + ["is_owner"]

def get_current_role(user):
    for label, key in ROLES.items():
        if user.get(key):
            return label, key
    if user.get("is_owner"):
        return "Owner", "is_owner"
    return "Sin rol", None

def get_user(user_id, auth_token):
    url = f"http://api.simpliroute.com/v1/accounts/users/{user_id}/"
    headers = {"Authorization": f"Token {auth_token}", "Content-Type": "application/json"}
    try:
        r = requests.get(url, headers=headers, timeout=15)
        return r.status_code, r.json()
    except Exception as e:
        return None, str(e)

def update_user_role(user_id, new_role_key, auth_token, user_data={}):
    url = f"http://api.simpliroute.com/v1/accounts/users/{user_id}/"
    headers = {"Authorization": f"Token {auth_token}", "Content-Type": "application/json"}
    # Solo enviar los campos de rol, todos en False excepto el nuevo
    payload = {key: (key == new_role_key) for key in ALL_ROLE_KEYS}
    payload["username"] = user_data.get("username", "")
    payload["name"] = user_data.get("name", "")
    try:
        r = requests.put(url, headers=headers, json=payload, timeout=15)
        return r.status_code, r.json()
    except Exception as e:
        return None, str(e)

def page_cambiar_rol():
    st.title("👤 Cambiar Rol de Usuario")
    st.markdown("Consulta un usuario por ID y cambia su rol en la plataforma.")

    col1, col2 = st.columns(2)
    with col1:
        token = st.text_input("🔑 Token de SimpliRoute", type="password", placeholder="Ingresa tu token aquí")
    with col2:
        user_id = st.text_input("🆔 ID del Usuario", placeholder="Ej: 524614")

    if st.button("🔍 Consultar usuario", type="primary", disabled=not (token and user_id)):
        with st.spinner("Consultando usuario..."):
            code, resp = get_user(user_id.strip(), token.strip())

        if code == 200:
            st.session_state["user_data"] = resp
            st.session_state["user_token"] = token.strip()
        elif code == 401:
            st.error("❌ Token inválido o sin permisos.")
            st.session_state.pop("user_data", None)
        elif code == 404:
            st.error("❌ Usuario no encontrado. Verifica el ID.")
            st.session_state.pop("user_data", None)
        elif code is None:
            st.error(f"❌ Sin conexión: {resp}")
            st.session_state.pop("user_data", None)
        else:
            st.error(f"❌ Error {code}: {resp}")
            st.session_state.pop("user_data", None)


    if "user_data" in st.session_state:
        user = st.session_state["user_data"]
        current_label, current_key = get_current_role(user)

        st.divider()
        st.subheader("Información del usuario")
        col1, col2, col3 = st.columns(3)
        col1.markdown(f"**Nombre**<br>{user.get('name', '—')}", unsafe_allow_html=True)
        col2.markdown(f"**Username**<br>{user.get('username', '—')}", unsafe_allow_html=True)
        col3.markdown(f"**Rol actual**<br>{current_label}", unsafe_allow_html=True)
        st.caption(f"📧 {user.get('email', '—')} · Estado: {user.get('status', '—')}")

        st.divider()

        if current_key == "is_owner":
            st.warning("⚠️ Este usuario es **Owner** de la cuenta. Su rol no puede ser modificado.")
        else:
            st.subheader("Cambiar rol")
            # Opciones de rol excluyendo el actual y owner
            opciones = [label for label, key in ROLES.items() if key != current_key]
            nuevo_rol_label = st.selectbox("Selecciona el nuevo rol:", opciones)
            nuevo_rol_key = ROLES[nuevo_rol_label]

            if st.button(f"💾 Cambiar rol a {nuevo_rol_label}", type="primary"):
                with st.spinner("Actualizando rol..."):
                    code, resp = update_user_role(
                        user["id"],
                        nuevo_rol_key,
                        st.session_state["user_token"],
                        user
                    )
                if code == 200:
                    st.success(f"✅ Rol actualizado correctamente a **{nuevo_rol_label}**")
                    st.session_state.pop("user_data", None)
                elif code == 401:
                    st.error("❌ Token inválido o sin permisos.")
                elif code == 400:
                    st.error(f"❌ Error de validación: {resp}")
                elif code is None:
                    st.error(f"❌ Sin conexión: {resp}")
                else:
                    st.error(f"❌ Error {code}: {resp}")


# ─────────────────────────────────────────
# ROUTER
# ─────────────────────────────────────────
if menu == "🗺️ Cargar Zonas":
    page_cargar_zonas()
elif menu == "👤 Cambiar Rol de Usuario":
    page_cambiar_rol()