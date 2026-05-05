import streamlit as st
import openpyxl
import io
import re
import requests
import json
from datetime import date

st.set_page_config(page_title="SimpliRoute Tools", page_icon="🚀", layout="wide")

# ── SIDEBAR ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## 🚀 SimpliRoute Tools")
    st.markdown("---")

    if "active_section" not in st.session_state:
        st.session_state["active_section"] = "core"
    if "menu_core" not in st.session_state:
        st.session_state["menu_core"] = "🧑‍💼 Agregar Seller a Visitas"
    if "menu_tms" not in st.session_state:
        st.session_state["menu_tms"] = None

    def on_core_change():
        st.session_state["active_section"] = "core"
        st.session_state["menu_tms"] = None

    def on_tms_change():
        st.session_state["active_section"] = "tms"
        st.session_state["menu_core"] = None

    st.markdown("### 🛠️ SimpliRoute Core")
    menu = st.radio(
        label="Navegación Core",
        options=[
            "🧑‍💼 Agregar Seller a Visitas",
            "🚛 Asignación de Flotas",
            "🗺️ Cargar Zonas",
            "👤 Cambiar Rol de Usuario",
            "🔔 Crear Webhook",
            "🔓 Desbloqueo de Contraseña",
            "🔁 Reenviar Webhooks",
            "🏷️ Tipos de Visita y Skills",
        ],
        label_visibility="collapsed",
        key="menu_core",
        index=None if st.session_state.get("menu_core") is None else [
            "🧑‍💼 Agregar Seller a Visitas",
            "🚛 Asignación de Flotas",
            "🗺️ Cargar Zonas",
            "👤 Cambiar Rol de Usuario",
            "🔔 Crear Webhook",
            "🔓 Desbloqueo de Contraseña",
            "🔁 Reenviar Webhooks",
            "🏷️ Tipos de Visita y Skills",
        ].index(st.session_state["menu_core"]) if st.session_state.get("menu_core") in [
            "🧑‍💼 Agregar Seller a Visitas",
            "🚛 Asignación de Flotas",
            "🗺️ Cargar Zonas",
            "👤 Cambiar Rol de Usuario",
            "🔔 Crear Webhook",
            "🔓 Desbloqueo de Contraseña",
            "🔁 Reenviar Webhooks",
            "🏷️ Tipos de Visita y Skills",
        ] else None,
        on_change=on_core_change
    )

    st.markdown("---")
    st.markdown("### 📦 TMS")
    menu_tms = st.radio(
        label="Navegación TMS",
        options=[
            "📄 Tipos de Documento",
            "🚚 Transportistas",
        ],
        label_visibility="collapsed",
        key="menu_tms",
        index=None if st.session_state.get("menu_tms") is None else [
            "📄 Tipos de Documento",
            "🚚 Transportistas",
        ].index(st.session_state["menu_tms"]) if st.session_state.get("menu_tms") in [
            "📄 Tipos de Documento",
            "🚚 Transportistas",
        ] else None,
        on_change=on_tms_change
    )

    st.markdown("---")
    st.caption("SimpliRoute Internal Tools v1.0")

active_section = st.session_state["active_section"]
menu = st.session_state.get("menu_core")
menu_tms = st.session_state.get("menu_tms")


# ── HELPERS COMPARTIDOS ───────────────────────────────────────────────────────
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

def generate_excel_zones(polygons):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Hoja 1"
    ws.append(["is_name", "is_coordinates"])
    for p in polygons:
        ws.append([p["name"], coords_to_str(p["coords"])])
    ws.column_dimensions["A"].width = 25
    ws.column_dimensions["B"].width = 80
    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    return buf

def read_excel_column(file, columns):
    wb = openpyxl.load_workbook(file)
    ws = wb.active
    headers = [str(cell.value).strip() if cell.value else "" for cell in ws[1]]
    rows = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if all(v is None for v in row):
            continue
        entry = {}
        for col in columns:
            idx = headers.index(col) if col in headers else -1
            val = row[idx] if 0 <= idx < len(row) else None
            entry[col] = str(val).strip() if val is not None else ""
        if any(entry[c] for c in columns):
            rows.append(entry)
    return rows

def show_results(results, name_key):
    ok = sum(1 for r in results if r["code"] in [200, 201])
    st.markdown(f"✅ **{ok} creados correctamente** | ❌ **{len(results)-ok} con error**")
    for r in results:
        if r["code"] in [200, 201]:
            st.success(f"✅ {r[name_key]} — Creado correctamente")
        elif r["code"] == 400:
            st.warning(f"⚠️ {r[name_key]} — Ya existe o datos inválidos: {r['resp']}")
        elif r["code"] == 401:
            st.error(f"❌ {r[name_key]} — Token inválido o sin permisos")
        elif r["code"] is None:
            st.error(f"❌ {r[name_key]} — Sin conexión: {r['resp']}")
        else:
            st.error(f"❌ {r[name_key]} — Error {r['code']}: {r['resp']}")

def get_users_list(token):
    url = "http://api.simpliroute.com/v1/accounts/users/"
    headers = {"Authorization": f"Token {token}", "accept": "application/json"}
    try:
        r = requests.get(url, headers=headers, timeout=300)
        return r.status_code, r.json()
    except Exception as e:
        return None, str(e)

def put_user_full(token, user, new_email):
    url = f"http://api.simpliroute.com/v1/accounts/users/{user['id']}/"
    headers = {"Authorization": f"Token {token}", "Content-Type": "application/json"}
    payload = {
        "id": user["id"],
        "username": user.get("username", ""),
        "name": user.get("name", ""),
        "phone": user.get("phone", ""),
        "email": new_email,
        "is_owner": user.get("is_owner", False),
        "is_admin": user.get("is_admin", False),
        "is_driver": user.get("is_driver", False),
        "is_codriver": user.get("is_codriver", False),
        "is_router_jr": user.get("is_router_jr", False),
        "is_monitor": user.get("is_monitor", False),
        "is_coordinator": user.get("is_coordinator", False),
        "is_router": user.get("is_router", False),
        "is_staff": user.get("is_staff", False),
        "is_seller_viewer": user.get("is_seller_viewer", False),
        "is_seller": user.get("is_seller", False),
        "blocked": user.get("blocked", False),
        "status": user.get("status", "active"),
    }
    try:
        r = requests.put(url, headers=headers, json=payload, timeout=30)
        return r.status_code, r.json()
    except Exception as e:
        return None, str(e)


# ── FEATURE 1: ASIGNACIÓN DE FLOTAS ──────────────────────────────────────────
def make_template_flotas():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Flotas"
    ws.append(["Nombre de flota", "Vehículos", "Usuarios"])
    ws.append(["Flota Norte", "MC4327,QC4380,MC7959", "juan.perez,ana.lopez"])
    ws.append(["Flota Sur", "DC8761,MA7532", "carlos.gomez,maria.ruiz"])
    ws.column_dimensions["A"].width = 25
    ws.column_dimensions["B"].width = 40
    ws.column_dimensions["C"].width = 40
    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    return buf

def get_fleets(token):
    url = "http://api.simpliroute.com/v1/fleets/"
    headers = {"Authorization": f"Token {token}", "accept": "application/json"}
    try:
        r = requests.get(url, headers=headers, timeout=300)
        return r.status_code, r.json()
    except Exception as e:
        return None, str(e)

def get_vehicles(token):
    url = "http://api.simpliroute.com/v1/routes/vehicles/"
    headers = {"Authorization": f"Token {token}", "Content-Type": "application/json"}
    try:
        r = requests.get(url, headers=headers, timeout=300)
        return r.status_code, r.json()
    except Exception as e:
        return None, str(e)

def update_fleet(fleet_id, fleet_name, vehicle_ids, user_ids, token):
    url = f"http://api.simpliroute.com/v1/fleets/{fleet_id}/"
    headers = {"Authorization": f"Token {token}", "Content-Type": "application/json"}
    payload = {"id": fleet_id, "name": fleet_name, "vehicles": vehicle_ids, "users": user_ids}
    try:
        r = requests.put(url, headers=headers, json=payload, timeout=300)
        return r.status_code, r.json()
    except Exception as e:
        return None, str(e)

def page_asignacion_flotas():
    st.title("🚛 Asignación de Flotas")
    st.markdown("Sube un Excel con la asignación de vehículos y usuarios por flota para actualizarlas masivamente.")
    st.info("ℹ️ Cada fila del Excel representa una edición independiente de flota. Si una misma flota aparece en dos filas, la segunda edición sobreescribirá a la primera.")
    token = st.text_input("🔑 Token de SimpliRoute", type="password", placeholder="Ingresa tu token aquí", key="token_flotas")
    st.download_button("📥 Descargar plantilla Flotas", data=make_template_flotas(),
                       file_name="plantilla_flotas.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    fleet_file = st.file_uploader("📂 Sube tu Excel de Flotas", type=["xlsx"], key="upload_flotas")
    if fleet_file and token:
        if st.button("🚀 Procesar y actualizar flotas", type="primary"):
            try:
                rows = read_excel_column(fleet_file, ["Nombre de flota", "Vehículos", "Usuarios"])
                if not rows:
                    st.error("❌ No se encontraron filas en el Excel.")
                    return
            except Exception as e:
                st.error(f"❌ Error leyendo el Excel: {e}"); return
            with st.spinner("⏳ Consultando flotas, vehículos y usuarios..."):
                code_f, fleets = get_fleets(token)
                code_v, vehicles = get_vehicles(token)
                code_u, users = get_users_list(token)
            if code_f != 200: st.error(f"❌ Error consultando flotas: {code_f}"); return
            if code_v != 200: st.error(f"❌ Error consultando vehículos: {code_v}"); return
            if code_u != 200: st.error(f"❌ Error consultando usuarios: {code_u}"); return
            fleet_map = {f["name"].strip().lower(): f for f in fleets}
            vehicle_map = {v["name"].strip().lower(): v["id"] for v in vehicles if v.get("name")}
            user_map = {}
            for u in users:
                if u.get("username"): user_map[u["username"].strip().lower()] = u["id"]
                if u.get("email"): user_map[u["email"].strip().lower()] = u["id"]
            st.success(f"✅ {len(fleets)} flotas · {len(vehicles)} vehículos · {len(users)} usuarios")
            st.divider()
            for row in rows:
                fleet_name = row["Nombre de flota"].strip()
                if fleet_name.lower() not in fleet_map:
                    st.error(f"❌ **{fleet_name}** — Flota no encontrada"); continue
                fleet = fleet_map[fleet_name.lower()]
                vehicle_ids, vehicle_errors = [], []
                for v in [x.strip() for x in row["Vehículos"].split(",") if x.strip()]:
                    vid = vehicle_map.get(v.lower())
                    (vehicle_ids if vid else vehicle_errors).append(vid if vid else v)
                user_ids, user_errors = [], []
                for u in [x.strip() for x in row["Usuarios"].split(",") if x.strip()]:
                    uid = user_map.get(u.lower())
                    (user_ids if uid else user_errors).append(uid if uid else u)
                if vehicle_errors: st.warning(f"⚠️ **{fleet_name}** — Vehículos no encontrados: {', '.join(str(x) for x in vehicle_errors)}")
                if user_errors: st.warning(f"⚠️ **{fleet_name}** — Usuarios no encontrados: {', '.join(str(x) for x in user_errors)}")
                code, resp = update_fleet(fleet["id"], fleet_name, vehicle_ids, user_ids, token)
                if code == 200: st.success(f"✅ **{fleet_name}** — Actualizada ({len(vehicle_ids)} vehículos · {len(user_ids)} usuarios)")
                elif code == 400: st.error(f"❌ **{fleet_name}** — Error: {resp}")
                elif code == 401: st.error(f"❌ **{fleet_name}** — Token inválido")
                else: st.error(f"❌ **{fleet_name}** — Error {code}: {resp}")
    elif fleet_file and not token:
        st.warning("⚠️ Ingresa tu token arriba para continuar.")


# ── FEATURE 2: CARGAR ZONAS ───────────────────────────────────────────────────
def upload_zone(name, coordinates, token):
    url = "http://api.simpliroute.com/v1/zones/"
    headers = {"authorization": f"Token {token}", "content-type": "application/json"}
    try:
        r = requests.post(url, headers=headers, json={"name": name, "coordinates": coordinates, "vehicles": []}, timeout=15)
        return r.status_code, r.json()
    except Exception as e:
        return None, str(e)

def page_cargar_zonas():
    st.title("🗺️ Cargar Zonas")
    st.markdown("Sube tu archivo KML o RTF, ingresa tu token y carga las zonas directo a SimpliRoute.")
    token = st.text_input("🔑 Token de SimpliRoute", type="password", placeholder="Ingresa tu token aquí", key="token_zonas")
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
                polygons = []; error = str(e)
        if error: st.error(f"Error: {error}")
        elif polygons:
            st.success(f"✅ {len(polygons)} polígono(s) encontrado(s)")
            for p in polygons:
                with st.expander(f"📍 {p['name']} — {len(p['coords'])} puntos"):
                    st.code(coords_to_str(p["coords"][:3]) + ",...]", language=None)
            st.download_button("⬇️ Descargar Excel", data=generate_excel_zones(polygons),
                               file_name="ZONES.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            st.divider()
            st.subheader("📡 Cargar zonas a SimpliRoute")
            if not token: st.warning("⚠️ Ingresa tu token arriba para cargar las zonas.")
            else:
                if st.button("🚀 Cargar zonas", type="primary"):
                    results = []; prog = st.progress(0); status = st.empty()
                    for i, p in enumerate(polygons):
                        status.info(f"Cargando: **{p['name']}** ({i+1}/{len(polygons)})")
                        code, resp = upload_zone(p["name"], coords_to_str(p["coords"]), token)
                        results.append({"name": p["name"], "code": code, "resp": resp})
                        prog.progress((i+1)/len(polygons))
                    status.empty(); prog.empty()
                    for r in results:
                        if r["code"] in [200, 201]: st.success(f"✅ {r['name']} — Zona creada")
                        elif r["code"] == 400: st.warning(f"⚠️ {r['name']} — Ya existe")
                        elif r["code"] == 401: st.error(f"❌ {r['name']} — Token inválido")
                        else: st.error(f"❌ {r['name']} — Error {r['code']}")
        else: st.warning("No se encontraron polígonos en el archivo.")


# ── FEATURE 3: CAMBIAR ROL DE USUARIO ────────────────────────────────────────
ROLES = {
    "Administrador": "is_admin", "Conductor": "is_driver", "Co-conductor": "is_codriver",
    "Router Jr": "is_router_jr", "Monitor": "is_monitor", "Coordinador": "is_coordinator",
    "Router": "is_router", "Staff": "is_staff", "Seller Viewer": "is_seller_viewer", "Seller": "is_seller",
}
ALL_ROLE_KEYS = list(ROLES.values()) + ["is_owner"]

def get_current_role(user):
    for label, key in ROLES.items():
        if user.get(key): return label, key
    if user.get("is_owner"): return "Owner", "is_owner"
    return "Sin rol", None

def get_user(user_id, token):
    url = f"http://api.simpliroute.com/v1/accounts/users/{user_id}/"
    headers = {"Authorization": f"Token {token}", "Content-Type": "application/json"}
    try:
        r = requests.get(url, headers=headers, timeout=15)
        return r.status_code, r.json()
    except Exception as e:
        return None, str(e)

def update_user_role(user_id, new_role_key, token, user_data={}):
    url = f"http://api.simpliroute.com/v1/accounts/users/{user_id}/"
    headers = {"Authorization": f"Token {token}", "Content-Type": "application/json"}
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
    col1, col2 = st.columns(2)
    with col1:
        token = st.text_input("🔑 Token de SimpliRoute", type="password", placeholder="Token", key="token_rol")
    with col2:
        user_id = st.text_input("🆔 ID del Usuario", placeholder="Ej: 524614")
    if st.button("🔍 Consultar usuario", type="primary", disabled=not (token and user_id)):
        with st.spinner("Consultando..."):
            code, resp = get_user(user_id.strip(), token.strip())
        if code == 200:
            st.session_state["user_data"] = resp
            st.session_state["user_token"] = token.strip()
        elif code == 401: st.error("❌ Token inválido.")
        elif code == 404: st.error("❌ Usuario no encontrado.")
        else: st.error(f"❌ Error {code}: {resp}")
        st.session_state.pop("user_data", None) if code != 200 else None
    if "user_data" in st.session_state:
        user = st.session_state["user_data"]
        current_label, current_key = get_current_role(user)
        st.divider()
        col1, col2, col3 = st.columns(3)
        col1.markdown(f"**Nombre**<br>{user.get('name','—')}", unsafe_allow_html=True)
        col2.markdown(f"**Username**<br>{user.get('username','—')}", unsafe_allow_html=True)
        col3.markdown(f"**Rol actual**<br>{current_label}", unsafe_allow_html=True)
        st.caption(f"📧 {user.get('email','—')} · Estado: {user.get('status','—')}")
        st.divider()
        if current_key == "is_owner":
            st.warning("⚠️ Este usuario es **Owner**. Su rol no puede modificarse.")
        else:
            opciones = [l for l, k in ROLES.items() if k != current_key]
            nuevo_label = st.selectbox("Nuevo rol:", opciones)
            nuevo_key = ROLES[nuevo_label]
            if st.button(f"💾 Cambiar a {nuevo_label}", type="primary"):
                with st.spinner("Actualizando..."):
                    code, resp = update_user_role(user["id"], nuevo_key, st.session_state["user_token"], user)
                if code == 200:
                    st.success(f"✅ Rol actualizado a **{nuevo_label}**")
                    _, fresh = get_user(user["id"], st.session_state["user_token"])
                    st.session_state["user_data"] = fresh if isinstance(fresh, dict) and "id" in fresh else None
                    if st.session_state["user_data"] is None: st.session_state.pop("user_data", None)
                    st.rerun()
                else: st.error(f"❌ Error {code}: {resp}")


# ── FEATURE 4: CREAR WEBHOOK ──────────────────────────────────────────────────
WEBHOOK_EVENTS = {
    "Plan creado": "plan_created", "Plan editado": "plan_edited",
    "Ruta creada": "route_created", "Ruta editada": "route_edited",
    "Ruta iniciada": "route_started", "Ruta finalizada": "route_finished",
    "En camino": "on_its_way", "Checkout": "visit_checkout",
    "Checkout detallado": "visit_checkout_detailed",
}

def page_crear_webhook():
    st.title("🔔 Crear Webhook")
    token = st.text_input("🔑 Token de SimpliRoute", type="password", placeholder="Ingresa tu token aquí", key="token_webhook")
    url = st.text_input("🌐 URL de destino", placeholder="https://tu-servidor.com/webhook")
    evento_label = st.selectbox("📋 Tipo de evento", options=list(WEBHOOK_EVENTS.keys()))
    evento_key = WEBHOOK_EVENTS[evento_label]
    st.markdown("**🔧 Headers**")
    custom_header = st.checkbox("Usar header personalizado")
    if custom_header:
        header_json_str = st.text_area("Header personalizado (JSON)",
                                       value='{\n    "Content-Type": "application/json",\n    "X-Custom-Header": "valor"\n}',
                                       height=150, label_visibility="collapsed")
        try:
            headers_payload = json.loads(header_json_str)
        except Exception:
            st.error("❌ JSON inválido."); headers_payload = None
    else:
        headers_payload = {"Content-Type": "application/json"}
        st.code('{"Content-Type": "application/json"}', language="json")
    st.divider()
    if st.button("🚀 Crear Webhook", type="primary", disabled=not (token and url)):
        if not url.startswith("http"): st.error("❌ La URL debe comenzar con http:// o https://"); return
        if custom_header and headers_payload is None: st.error("❌ Corrige el JSON."); return
        with st.spinner("Creando webhook..."):
            api_url = "http://api.simpliroute.com/v1/addons/webhooks/"
            headers = {"Authorization": f"Token {token}", "Content-Type": "application/json"}
            try:
                r = requests.post(api_url, headers=headers, json={"webhook": evento_key, "url": url, "headers": headers_payload}, timeout=30)
                code, resp = r.status_code, r.json()
            except Exception as e:
                code, resp = None, str(e)
        if code in [200, 201]: st.success(f"✅ Webhook **{evento_label}** creado"); st.json(resp)
        elif code == 400: st.error(f"❌ Error: {resp}")
        elif code == 401: st.error("❌ Token inválido.")
        else: st.error(f"❌ Error {code}: {resp}")


# ── FEATURE 5: DESBLOQUEO DE CONTRASEÑA ──────────────────────────────────────
SUPPORT_AGENTS = {
    "Brandon Vargas": "brandon.vargas@simpliroute.com",
    "Carlos Junior": "carlos.celestino@simpliroute.com",
    "David Martinez": "david.martinez@simpliroute.com",
    "Itzel Meza": "itzel.meza@simpliroute.com",
    "Jorge Cruz": "jorge.cruz@simpliroute.com",
    "Julio Mares": "julio.mares@simpliroute.com",
    "Roger Camacho": "roger.camacho@simpliroute.com",
    "Silmary Guedez": "silmary.guedez@simpliroute.com",
}

def page_desbloqueo():
    st.title("🔓 Desbloqueo de Contraseña")
    st.markdown("Consulta usuarios bloqueados y envía el link de desbloqueo al correo del agente.")

    token = st.text_input("🔑 Token de SimpliRoute", type="password",
                          placeholder="Ingresa tu token aquí", key="token_unlock")

    if st.button("🔍 Consultar usuarios bloqueados", type="primary", disabled=not token):
        with st.spinner("Consultando usuarios..."):
            code, resp = get_users_list(token)
        if code == 200:
            blocked = [u for u in resp if u.get("blocked") is True]
            if not blocked:
                st.info("✅ No hay usuarios bloqueados en esta cuenta.")
                st.session_state.pop("blocked_users", None)
                st.session_state.pop("unlock_token", None)
            else:
                st.session_state["blocked_users"] = blocked
                st.session_state["unlock_token"] = token
        elif code == 401:
            st.error("❌ Token inválido o sin permisos.")
        elif code is None:
            st.error(f"❌ Sin conexión: {resp}")
        else:
            st.error(f"❌ Error {code}: {resp}")

    # Botón de restauración pendiente
    if "pending_restore" in st.session_state:
        pr = st.session_state["pending_restore"]
        st.divider()
        st.subheader("📧 Restauración pendiente")
        st.info(f"Usuario: **{pr['user'].get('username')}** · Email original: **{pr['original_email'] or 'Sin email'}**")
        st.markdown("Una vez que el usuario haya reseteado su contraseña exitosamente, presiona el botón para restaurar su email original.")
        if st.button("↩️ Confirmar y restaurar email original", type="primary", key="btn_restore"):
            with st.spinner("Restaurando email original..."):
                code, resp = put_user_full(pr["token"], pr["user"], pr["original_email"])
            if code in [200, 201]:
                st.success("✅ Email original restaurado correctamente.")
                st.session_state.pop("pending_restore", None)
            else:
                st.error(f"❌ Error al restaurar: {code} — {resp}")

    if "blocked_users" in st.session_state:
        blocked = st.session_state["blocked_users"]
        st.divider()
        st.success(f"⚠️ {len(blocked)} usuario(s) bloqueado(s) encontrado(s)")
        st.subheader("Usuario bloqueado")

        def user_label(u):
            name = u.get("name") or u.get("username") or str(u.get("id"))
            username = u.get("username", "")
            return f"{name} — {username}" if username != name else name

        opciones = {user_label(u): u for u in blocked}
        selected_label = st.selectbox("Selecciona el usuario bloqueado:", list(opciones.keys()))
        selected_user = opciones[selected_label]
        original_email = selected_user.get("email", "")
        st.caption(f"📧 Email actual del usuario: **{original_email or 'Sin email'}**")

        st.divider()
        st.subheader("Agente que recibirá el link")
        agent_name = st.selectbox("Selecciona el agente:", list(SUPPORT_AGENTS.keys()))
        recovery_email = SUPPORT_AGENTS[agent_name]
        st.caption(f"📨 El link se enviará a: **{recovery_email}**")

        st.divider()
        if st.button("🔓 Enviar link de desbloqueo", type="primary"):
            tok = st.session_state["unlock_token"]
            prog = st.progress(0); status = st.empty()

            status.info("1/2 — Asignando email temporal al usuario...")
            code1, resp1 = put_user_full(tok, selected_user, recovery_email)
            prog.progress(50)

            if code1 not in [200, 201]:
                st.error(f"❌ Error al actualizar email: {code1} — {resp1}")
                status.empty(); prog.empty()
            else:
                status.info("2/2 — Enviando link de desbloqueo...")
                try:
                    r = requests.post(
                        "https://api.simpliroute.com/v2/auth/unlock/",
                        headers={"Content-Type": "application/json;charset=UTF-8", "authorization": "null"},
                        json={"username": selected_user.get("username")},
                        timeout=30
                    )
                    code2 = r.status_code
                    resp2 = r.json() if r.content else {}
                except Exception as e:
                    code2, resp2 = None, str(e)

                prog.progress(100); status.empty(); prog.empty()

                if code2 in [200, 201]:
                    st.success(f"✅ Link enviado a **{recovery_email}** ({agent_name})")
                    st.warning("⚠️ Una vez que el usuario haya reseteado su contraseña, presiona el botón de restauración que aparecerá arriba.")
                    st.session_state["pending_restore"] = {
                        "user": selected_user,
                        "original_email": original_email,
                        "token": tok
                    }
                    st.session_state.pop("blocked_users", None)
                    st.rerun()
                else:
                    put_user_full(tok, selected_user, original_email)
                    st.error(f"❌ No se pudo enviar el link. Email restaurado. ({code2}: {resp2})")


# ── FEATURE 6: REENVIAR WEBHOOKS ─────────────────────────────────────────────
def page_reenviar_webhooks():
    st.title("🔁 Reenviar Webhooks")
    st.info("🔑 Usa el token de tu cuenta de SimpliRoute.")
    token = st.text_input("🔑 Token de SimpliRoute", type="password", placeholder="Ingresa tu token aquí", key="token_resend")
    account_id = st.text_input("🏢 ID de la cuenta", placeholder="Ej: 30610")
    st.markdown("**📋 IDs de visita**")
    st.caption("Pega los IDs de visita uno por línea:\n```\n799841373\n808472905\n```")
    visit_ids_raw = st.text_area("IDs de visita", placeholder="799841373\n808472905", height=200, label_visibility="collapsed")
    today = date.today().strftime("%Y-%m-%d")
    st.caption(f"📅 Fecha utilizada: **{today}**")
    st.divider()
    if st.button("🚀 Reenviar Webhooks", type="primary", disabled=not (token and account_id and visit_ids_raw)):
        visit_ids = []; errors = []
        for line in visit_ids_raw.strip().splitlines():
            line = line.strip()
            if not line: continue
            try: visit_ids.append(int(line))
            except ValueError: errors.append(line)
        if errors: st.error(f"❌ IDs inválidos: {', '.join(errors)}"); return
        if not visit_ids: st.error("❌ No se encontraron IDs válidos."); return
        with st.spinner(f"Reenviando para {len(visit_ids)} visitas..."):
            try:
                r = requests.post("https://api.simpliroute.com/v1/mobile/send-webhooks",
                                  headers={"Authorization": f"Token {token}", "Content-Type": "application/json"},
                                  json={"account_ids": [int(account_id)], "planned_date": today, "visit_ids": visit_ids},
                                  timeout=60)
                code, resp = r.status_code, r.json() if r.content else {}
            except Exception as e:
                code, resp = None, str(e)
        if code in [200, 201]: st.success(f"✅ Webhooks reenviados para **{len(visit_ids)} visitas**")
        elif code == 400: st.error(f"❌ Error: {resp}")
        elif code == 401: st.error("❌ Token inválido.")
        elif code == 404: st.error("❌ Cuenta no encontrada.")
        else: st.error(f"❌ Error {code}: {resp}")


# ── FEATURE 7: TIPOS DE VISITA Y SKILLS ──────────────────────────────────────
def make_template_skills():
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Skills"
    ws.append(["skill"]); ws.append(["Manejo de carga pesada"]); ws.append(["Refrigeración"])
    ws.column_dimensions["A"].width = 35
    buf = io.BytesIO(); wb.save(buf); buf.seek(0); return buf

def make_template_visit_types():
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Tipos de Visita"
    ws.append(["label", "key"]); ws.append(["Entrega express", "entrega_express"]); ws.append(["Retiro de mercancía", "retiro_mercancia"])
    ws.column_dimensions["A"].width = 30; ws.column_dimensions["B"].width = 30
    buf = io.BytesIO(); wb.save(buf); buf.seek(0); return buf

def page_visit_types_skills():
    st.title("🏷️ Tipos de Visita y Skills")
    token = st.text_input("🔑 Token de SimpliRoute", type="password", placeholder="Ingresa tu token aquí", key="token_vts")

    st.markdown("---")
    st.subheader("🔧 Skills")
    st.download_button("📥 Descargar plantilla Skills", data=make_template_skills(),
                       file_name="plantilla_skills.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    skills_file = st.file_uploader("📂 Sube tu Excel de Skills", type=["xlsx"], key="upload_skills")
    if skills_file:
        try:
            rows = read_excel_column(skills_file, ["skill"])
            st.success(f"✅ {len(rows)} skill(s)")
            with st.expander("Ver lista"):
                for r in rows: st.markdown(f"- {r['skill']}")
        except Exception as e:
            rows = []; st.error(f"Error: {e}")
        if rows and token:
            if st.button("🚀 Crear Skills", type="primary", key="btn_skills"):
                results = []; prog = st.progress(0); status = st.empty()
                for i, row in enumerate(rows):
                    status.info(f"Creando: **{row['skill']}** ({i+1}/{len(rows)})")
                    try:
                        r = requests.post("http://api.simpliroute.com/v1/routes/skills/",
                                          headers={"Authorization": f"Token {token}", "Content-Type": "application/json"},
                                          json={"skill": row["skill"]}, timeout=15)
                        code, resp = r.status_code, r.json()
                    except Exception as e:
                        code, resp = None, str(e)
                    results.append({"skill": row["skill"], "code": code, "resp": resp})
                    prog.progress((i+1)/len(rows))
                status.empty(); prog.empty(); show_results(results, "skill")

    st.markdown("---")
    st.subheader("📋 Tipos de Visita")
    st.download_button("📥 Descargar plantilla Tipos de Visita", data=make_template_visit_types(),
                       file_name="plantilla_tipos_visita.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    vt_file = st.file_uploader("📂 Sube tu Excel de Tipos de Visita", type=["xlsx"], key="upload_vt")
    if vt_file:
        try:
            rows_vt = read_excel_column(vt_file, ["label", "key"])
            st.success(f"✅ {len(rows_vt)} tipo(s)")
            with st.expander("Ver lista"):
                for r in rows_vt: st.markdown(f"- **{r['label']}** → `{r['key']}`")
        except Exception as e:
            rows_vt = []; st.error(f"Error: {e}")
        if rows_vt and token:
            if st.button("🚀 Crear Tipos de Visita", type="primary", key="btn_vt"):
                results_vt = []; prog = st.progress(0); status = st.empty()
                for i, row in enumerate(rows_vt):
                    status.info(f"Creando: **{row['label']}** ({i+1}/{len(rows_vt)})")
                    try:
                        r = requests.post("http://api.simpliroute.com/v1/accounts/visit-types/",
                                          headers={"Authorization": f"Token {token}", "Content-Type": "application/json"},
                                          json=[{"label": row["label"], "key": row["key"]}], timeout=15)
                        code, resp = r.status_code, r.json()
                    except Exception as e:
                        code, resp = None, str(e)
                    results_vt.append({"label": row["label"], "code": code, "resp": resp})
                    prog.progress((i+1)/len(rows_vt))
                status.empty(); prog.empty(); show_results(results_vt, "label")


# ── FEATURE 8: AGREGAR SELLER A VISITAS ──────────────────────────────────────
def get_sellers(token):
    url = "http://api.simpliroute.com/v1/sellers/"
    headers = {"Authorization": f"Token {token}", "Content-Type": "application/json"}
    try:
        r = requests.get(url, headers=headers, timeout=60)
        return r.status_code, r.json()
    except Exception as e:
        return None, str(e)

def page_agregar_seller():
    st.title("🧑‍💼 Agregar Seller a Visitas")
    token = st.text_input("🔑 Token de SimpliRoute", type="password", placeholder="Ingresa tu token aquí", key="token_seller")
    if st.button("🔍 Consultar Sellers", type="primary", disabled=not token):
        with st.spinner("Consultando sellers..."):
            code, resp = get_sellers(token)
        if code == 200:
            if not resp:
                st.warning("⚠️ No se encontraron sellers.")
                st.session_state.pop("sellers", None)
            else:
                st.session_state["sellers"] = resp
                st.session_state["seller_token"] = token
                st.success(f"✅ {len(resp)} seller(s) encontrado(s)")
        elif code == 401: st.error("❌ Token inválido.")
        else: st.error(f"❌ Error {code}: {resp}")

    if "sellers" in st.session_state:
        sellers = st.session_state["sellers"]
        st.divider()

        def seller_label(s):
            name = s.get("name") or s.get("username") or s.get("email") or str(s.get("uuid", ""))
            email = s.get("email", "")
            return f"{name} — {email}" if email and email != name else name

        opciones = {seller_label(s): s for s in sellers}
        selected = opciones[st.selectbox("Seller:", list(opciones.keys()))]
        seller_uuid = selected.get("uuid") or selected.get("id")

        st.divider()
        st.caption("Pega los IDs de visita uno por línea:\n```\n799841373\n808472905\n```")
        visit_ids_raw = st.text_area("IDs de visita", placeholder="799841373\n808472905", height=200, label_visibility="collapsed")

        if st.button("💾 Asignar Seller", type="primary", disabled=not visit_ids_raw):
            visit_ids = []; errors = []
            for line in visit_ids_raw.strip().splitlines():
                line = line.strip()
                if not line: continue
                try: visit_ids.append(int(line))
                except ValueError: errors.append(line)
            if errors: st.error(f"❌ IDs inválidos: {', '.join(errors)}"); return
            with st.spinner(f"Asignando a {len(visit_ids)} visitas..."):
                try:
                    r = requests.patch("http://api.simpliroute.com/v1/routes/visits/",
                                       headers={"Authorization": f"Token {st.session_state['seller_token']}", "Content-Type": "application/json"},
                                       json=[{"id": vid, "seller": seller_uuid} for vid in visit_ids], timeout=120)
                    code = r.status_code
                except Exception as e:
                    code = None
            if code in [200, 201]: st.success(f"✅ Seller asignado a **{len(visit_ids)} visitas**")
            elif code == 401: st.error("❌ Token inválido.")
            else: st.error(f"❌ Error {code}")


# ── TMS FEATURE 1: TIPOS DE DOCUMENTO ────────────────────────────────────────
COUNTRY_OPTIONS = ["PE", "CL", "MX", "CO", "AR", "EC", "BO", "UY", "PY", "VE"]
ENTITY_TYPES = ["provider", "customer", "driver", "vehicle"]

def make_template_doc_types():
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Tipos de Documento"
    ws.append(["name", "entity_type"])
    ws.append(["DNI", "driver"])
    ws.append(["RUC", "provider"])
    ws.column_dimensions["A"].width = 25
    ws.column_dimensions["B"].width = 25
    buf = io.BytesIO(); wb.save(buf); buf.seek(0); return buf

def page_tms_document_types():
    st.title("📄 Tipos de Documento TMS")
    st.markdown("Crea tipos de documento de manera masiva en TMS.")

    col1, col2, col3 = st.columns(3)
    with col1:
        token = st.text_input("🔑 Token", type="password", placeholder="Token SimpliRoute", key="token_doctype")
    with col2:
        account_id = st.text_input("🏢 Account ID", placeholder="Ej: 9695", key="account_doctype")
    with col3:
        country = st.selectbox("🌎 País", COUNTRY_OPTIONS, key="country_doctype")

    st.download_button("📥 Descargar plantilla", data=make_template_doc_types(),
                       file_name="plantilla_tipos_documento.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    doc_file = st.file_uploader("📂 Sube tu Excel", type=["xlsx"], key="upload_doctype")

    if doc_file:
        try:
            rows = read_excel_column(doc_file, ["name", "entity_type"])
            st.success(f"✅ {len(rows)} tipo(s) encontrado(s)")
            with st.expander("Ver lista"):
                for r in rows:
                    st.markdown(f"- **{r['name']}** → `{r['entity_type']}`")
        except Exception as e:
            rows = []; st.error(f"Error: {e}")

        if rows:
            if not (token and account_id):
                st.warning("⚠️ Completa el token y account ID.")
            else:
                if st.button("🚀 Crear Tipos de Documento", type="primary", key="btn_doctype"):
                    results = []; prog = st.progress(0); status = st.empty()
                    for i, row in enumerate(rows):
                        status.info(f"Creando: **{row['name']}** ({i+1}/{len(rows)})")
                        try:
                            r = requests.post(
                                "https://api.simpliroute.com/tms/api/v1/document-types/",
                                headers={
                                    "Authorization": f"Token {token}",
                                    "Related-Account": account_id,
                                    "Content-Type": "application/json"
                                },
                                json={"name": row["name"], "country": country,
                                      "entity_type": row["entity_type"], "status": "active"},
                                timeout=30
                            )
                            code, resp = r.status_code, r.json()
                        except Exception as e:
                            code, resp = None, str(e)
                        results.append({"name": row["name"], "code": code, "resp": resp})
                        prog.progress((i+1)/len(rows))
                    status.empty(); prog.empty()
                    show_results(results, "name")


# ── TMS FEATURE 2: TRANSPORTISTAS ────────────────────────────────────────────
def make_template_transportistas():
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Transportistas"
    ws.append(["trade_name", "legal_name", "tax_id_number"])
    ws.append(["Transportes Sur S.A.", "Transportes Sur Sociedad Anónima", "20123456789"])
    ws.append(["LogiRápido", "LogiRápido E.I.R.L.", "20987654321"])
    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 35
    ws.column_dimensions["C"].width = 20
    buf = io.BytesIO(); wb.save(buf); buf.seek(0); return buf

def page_tms_transportistas():
    st.title("🚚 Transportistas TMS")
    st.markdown("Crea transportistas de manera masiva en TMS.")

    col1, col2 = st.columns(2)
    with col1:
        token = st.text_input("🔑 Token", type="password", placeholder="Token SimpliRoute", key="token_transp")
    with col2:
        account_id = st.text_input("🏢 Account ID", placeholder="Ej: 82761", key="account_transp")

    st.download_button("📥 Descargar plantilla", data=make_template_transportistas(),
                       file_name="plantilla_transportistas.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    transp_file = st.file_uploader("📂 Sube tu Excel", type=["xlsx"], key="upload_transp")

    if transp_file:
        try:
            rows = read_excel_column(transp_file, ["trade_name", "legal_name", "tax_id_number"])
            st.success(f"✅ {len(rows)} transportista(s) encontrado(s)")
            with st.expander("Ver lista"):
                for r in rows:
                    st.markdown(f"- **{r['trade_name']}** | {r['legal_name']} | `{r['tax_id_number']}`")
        except Exception as e:
            rows = []; st.error(f"Error: {e}")

        if rows:
            if not (token and account_id):
                st.warning("⚠️ Completa el token y account ID.")
            else:
                if st.button("🚀 Crear Transportistas", type="primary", key="btn_transp"):
                    results = []; prog = st.progress(0); status = st.empty()
                    for i, row in enumerate(rows):
                        status.info(f"Creando: **{row['trade_name']}** ({i+1}/{len(rows)})")
                        try:
                            r = requests.post(
                                "https://api.simpliroute.com/tms/api/v1/transportation-providers/",
                                headers={
                                    "Authorization": f"Token {token}",
                                    "Content-Type": "application/json"
                                },
                                json={
                                    "trade_name": row["trade_name"],
                                    "legal_name": row["legal_name"],
                                    "tax_id_number": row["tax_id_number"],
                                    "account_id": int(account_id)
                                },
                                timeout=30
                            )
                            code, resp = r.status_code, r.json()
                        except Exception as e:
                            code, resp = None, str(e)
                        results.append({"trade_name": row["trade_name"], "code": code, "resp": resp})
                        prog.progress((i+1)/len(rows))
                    status.empty(); prog.empty()
                    show_results(results, "trade_name")


# ── ROUTER ────────────────────────────────────────────────────────────────────
if menu == "🚛 Asignación de Flotas":
    page_asignacion_flotas()
elif menu == "🗺️ Cargar Zonas":
    page_cargar_zonas()
elif menu == "👤 Cambiar Rol de Usuario":
    page_cambiar_rol()
elif menu == "🔔 Crear Webhook":
    page_crear_webhook()
elif menu == "🔓 Desbloqueo de Contraseña":
    page_desbloqueo()
elif menu == "🔁 Reenviar Webhooks":
    page_reenviar_webhooks()
elif menu == "🏷️ Tipos de Visita y Skills":
    page_visit_types_skills()
elif menu == "🧑‍💼 Agregar Seller a Visitas":
    page_agregar_seller()