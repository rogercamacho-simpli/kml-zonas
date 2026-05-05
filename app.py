import streamlit as st
import openpyxl
import io
import re
import requests

st.set_page_config(page_title="SimpliRoute Tools", page_icon="🚀", layout="wide")

# ── SIDEBAR ──────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## 🚀 SimpliRoute Tools")
    st.markdown("---")
    st.markdown("### 🛠️ Herramientas")
    menu = st.radio(
        label="Navegación",
        options=[
            "🚛 Asignación de Flotas",
            "🗺️ Cargar Zonas",
            "👤 Cambiar Rol de Usuario",
            "🔔 Crear Webhook",
            "🔁 Reenviar Webhooks",
            "🏷️ Tipos de Visita y Skills",
            "🧑‍💼 Agregar Seller a Visitas",
        ],
        label_visibility="collapsed"
    )
    st.markdown("---")
    st.caption("SimpliRoute Internal Tools v1.0")


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


# ── FEATURE 1: CARGAR ZONAS ───────────────────────────────────────────────────
def upload_zone(name, coordinates, auth_token):
    url = "http://api.simpliroute.com/v1/zones/"
    headers = {"authorization": f"Token {auth_token}", "content-type": "application/json"}
    try:
        r = requests.post(url, headers=headers, json={"name": name, "coordinates": coordinates, "vehicles": []}, timeout=15)
        return r.status_code, r.json()
    except Exception as e:
        return None, str(e)

def page_cargar_zonas():
    st.title("🗺️ Cargar Zonas")
    st.markdown("Sube tu archivo KML o RTF, ingresa tu token y carga las zonas directo a SimpliRoute.")
    st.info("ℹ️ Cada fila del Excel representa una edición independiente de flota. Si una misma flota aparece en dos filas, la segunda edición sobreescribirá a la primera.")
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
            for p in polygons:
                with st.expander(f"📍 {p['name']} — {len(p['coords'])} puntos"):
                    st.code(coords_to_str(p["coords"][:3]) + ",...]", language=None)
            st.download_button("⬇️ Descargar Excel", data=generate_excel(polygons),
                               file_name="ZONES.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            st.divider()
            st.subheader("📡 Cargar zonas a SimpliRoute")
            if not token:
                st.warning("⚠️ Ingresa tu token arriba para cargar las zonas.")
            else:
                if st.button("🚀 Cargar zonas a SimpliRoute", type="primary"):
                    results = []
                    prog = st.progress(0)
                    status = st.empty()
                    for i, p in enumerate(polygons):
                        status.info(f"Cargando: **{p['name']}** ({i+1}/{len(polygons)})")
                        code, resp = upload_zone(p["name"], coords_to_str(p["coords"]), token)
                        results.append({"name": p["name"], "code": code, "resp": resp})
                        prog.progress((i+1)/len(polygons))
                    status.empty(); prog.empty()
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


# ── FEATURE 2: CAMBIAR ROL DE USUARIO ────────────────────────────────────────
ROLES = {
    "Administrador": "is_admin", "Conductor": "is_driver", "Co-conductor": "is_codriver",
    "Router Jr": "is_router_jr", "Monitor": "is_monitor", "Coordinador": "is_coordinator",
    "Router": "is_router", "Staff": "is_staff", "Seller Viewer": "is_seller_viewer", "Seller": "is_seller",
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
            opciones = [label for label, key in ROLES.items() if key != current_key]
            nuevo_rol_label = st.selectbox("Selecciona el nuevo rol:", opciones)
            nuevo_rol_key = ROLES[nuevo_rol_label]
            if st.button(f"💾 Cambiar rol a {nuevo_rol_label}", type="primary"):
                with st.spinner("Actualizando rol..."):
                    code, resp = update_user_role(user["id"], nuevo_rol_key, st.session_state["user_token"], user)
                if code == 200:
                    st.success(f"✅ Rol actualizado correctamente a **{nuevo_rol_label}**")
                    _, fresh = get_user(user["id"], st.session_state["user_token"])
                    if isinstance(fresh, dict) and "id" in fresh:
                        st.session_state["user_data"] = fresh
                    else:
                        st.session_state.pop("user_data", None)
                    st.rerun()
                elif code == 401:
                    st.error("❌ Token inválido o sin permisos.")
                elif code == 400:
                    st.error(f"❌ Error de validación: {resp}")
                elif code is None:
                    st.error(f"❌ Sin conexión: {resp}")
                else:
                    st.error(f"❌ Error {code}: {resp}")


# ── FEATURE 3: TIPOS DE VISITA Y SKILLS ──────────────────────────────────────
def make_template_skills():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Skills"
    ws.append(["skill"])
    ws.append(["Manejo de carga pesada"])
    ws.append(["Refrigeración"])
    ws.column_dimensions["A"].width = 35
    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    return buf

def make_template_visit_types():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Tipos de Visita"
    ws.append(["label", "key"])
    ws.append(["Entrega express", "entrega_express"])
    ws.append(["Retiro de mercancía", "retiro_mercancia"])
    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 30
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

def create_skill(skill, auth_token):
    url = "http://api.simpliroute.com/v1/routes/skills/"
    headers = {"Authorization": f"Token {auth_token}", "Content-Type": "application/json"}
    try:
        r = requests.post(url, headers=headers, json={"skill": skill}, timeout=15)
        return r.status_code, r.json()
    except Exception as e:
        return None, str(e)

def create_visit_type(label, key, auth_token):
    url = "http://api.simpliroute.com/v1/accounts/visit-types/"
    headers = {"Authorization": f"Token {auth_token}", "Content-Type": "application/json"}
    try:
        r = requests.post(url, headers=headers, json=[{"label": label, "key": key}], timeout=15)
        return r.status_code, r.json()
    except Exception as e:
        return None, str(e)

def page_visit_types_skills():
    st.title("🏷️ Tipos de Visita y Skills")
    st.markdown("Crea tipos de visita y skills de manera masiva desde un archivo Excel.")
    token = st.text_input("🔑 Token de SimpliRoute", type="password", placeholder="Ingresa tu token aquí", key="token_vts")

    # ── Skills
    st.markdown("---")
    st.subheader("🔧 Skills")
    st.markdown("Sube un Excel con una columna `skill` para crear habilidades en masa.")
    st.download_button("📥 Descargar plantilla Skills", data=make_template_skills(),
                       file_name="plantilla_skills.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    skills_file = st.file_uploader("📂 Sube tu Excel de Skills", type=["xlsx"], key="upload_skills")
    if skills_file:
        try:
            rows = read_excel_column(skills_file, ["skill"])
            st.success(f"✅ {len(rows)} skill(s) encontrado(s)")
            with st.expander("Ver lista"):
                for r in rows:
                    st.markdown(f"- {r['skill']}")
        except Exception as e:
            rows = []
            st.error(f"Error leyendo el archivo: {e}")
        if rows:
            if not token:
                st.warning("⚠️ Ingresa tu token arriba para crear los skills.")
            else:
                if st.button("🚀 Crear Skills", type="primary", key="btn_skills"):
                    results = []
                    prog = st.progress(0); status = st.empty()
                    for i, row in enumerate(rows):
                        status.info(f"Creando skill: **{row['skill']}** ({i+1}/{len(rows)})")
                        code, resp = create_skill(row["skill"], token)
                        results.append({"skill": row["skill"], "code": code, "resp": resp})
                        prog.progress((i+1)/len(rows))
                    status.empty(); prog.empty()
                    show_results(results, "skill")

    # ── Tipos de Visita
    st.markdown("---")
    st.subheader("📋 Tipos de Visita")
    st.markdown("Sube un Excel con columnas `label` y `key` para crear tipos de visita en masa.")
    st.download_button("📥 Descargar plantilla Tipos de Visita", data=make_template_visit_types(),
                       file_name="plantilla_tipos_visita.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    vt_file = st.file_uploader("📂 Sube tu Excel de Tipos de Visita", type=["xlsx"], key="upload_vt")
    if vt_file:
        try:
            rows_vt = read_excel_column(vt_file, ["label", "key"])
            st.success(f"✅ {len(rows_vt)} tipo(s) de visita encontrado(s)")
            with st.expander("Ver lista"):
                for r in rows_vt:
                    st.markdown(f"- **{r['label']}** → `{r['key']}`")
        except Exception as e:
            rows_vt = []
            st.error(f"Error leyendo el archivo: {e}")
        if rows_vt:
            if not token:
                st.warning("⚠️ Ingresa tu token arriba para crear los tipos de visita.")
            else:
                if st.button("🚀 Crear Tipos de Visita", type="primary", key="btn_vt"):
                    results_vt = []
                    prog = st.progress(0); status = st.empty()
                    for i, row in enumerate(rows_vt):
                        status.info(f"Creando: **{row['label']}** ({i+1}/{len(rows_vt)})")
                        code, resp = create_visit_type(row["label"], row["key"], token)
                        results_vt.append({"label": row["label"], "code": code, "resp": resp})
                        prog.progress((i+1)/len(rows_vt))
                    status.empty(); prog.empty()
                    show_results(results_vt, "label")


# ── FEATURE 4: ASIGNACIÓN DE FLOTAS ──────────────────────────────────────────
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

def get_fleets(auth_token):
    url = "http://api.simpliroute.com/v1/fleets/"
    headers = {"Authorization": f"Token {auth_token}", "accept": "application/json"}
    try:
        r = requests.get(url, headers=headers, timeout=300)
        return r.status_code, r.json()
    except Exception as e:
        return None, str(e)

def get_vehicles(auth_token):
    url = "http://api.simpliroute.com/v1/routes/vehicles/"
    headers = {"Authorization": f"Token {auth_token}", "Content-Type": "application/json"}
    try:
        r = requests.get(url, headers=headers, timeout=300)
        return r.status_code, r.json()
    except Exception as e:
        return None, str(e)

def get_users_list(auth_token):
    url = "http://api.simpliroute.com/v1/accounts/users/"
    headers = {"Authorization": f"Token {auth_token}", "accept": "application/json"}
    try:
        r = requests.get(url, headers=headers, timeout=300)
        return r.status_code, r.json()
    except Exception as e:
        return None, str(e)

def update_fleet(fleet_id, fleet_name, vehicle_ids, user_ids, auth_token):
    url = f"http://api.simpliroute.com/v1/fleets/{fleet_id}/"
    headers = {"Authorization": f"Token {auth_token}", "Content-Type": "application/json"}
    payload = {
        "id": fleet_id,
        "name": fleet_name,
        "vehicles": vehicle_ids,
        "users": user_ids
    }
    try:
        r = requests.put(url, headers=headers, json=payload, timeout=300)
        return r.status_code, r.json()
    except Exception as e:
        return None, str(e)

def page_asignacion_flotas():
    st.title("🚛 Asignación de Flotas")
    st.markdown("Sube un Excel con la asignación de vehículos y usuarios por flota para actualizarlas masivamente.")

    token = st.text_input("🔑 Token de SimpliRoute", type="password",
                          placeholder="Ingresa tu token aquí", key="token_flotas")

    st.download_button(
        label="📥 Descargar plantilla Flotas",
        data=make_template_flotas(),
        file_name="plantilla_flotas.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    fleet_file = st.file_uploader("📂 Sube tu Excel de Flotas", type=["xlsx"], key="upload_flotas")

    if fleet_file and token:
        if st.button("🚀 Procesar y actualizar flotas", type="primary", key="btn_flotas"):

            # Leer Excel
            try:
                rows = read_excel_column(fleet_file, ["Nombre de flota", "Vehículos", "Usuarios"])
                if not rows:
                    st.error("❌ No se encontraron filas en el Excel.")
                    return
            except Exception as e:
                st.error(f"❌ Error leyendo el Excel: {e}")
                return

            # Consultar flotas, vehículos y usuarios
            with st.spinner("⏳ Consultando flotas, vehículos y usuarios (puede tardar varios minutos)..."):
                code_f, fleets = get_fleets(token)
                code_v, vehicles = get_vehicles(token)
                code_u, users = get_users_list(token)

            if code_f != 200:
                st.error(f"❌ Error consultando flotas: {code_f} — {fleets}")
                return
            if code_v != 200:
                st.error(f"❌ Error consultando vehículos: {code_v} — {vehicles}")
                return
            if code_u != 200:
                st.error(f"❌ Error consultando usuarios: {code_u} — {users}")
                return

            # Crear mapas nombre → id
            fleet_map = {f["name"].strip().lower(): f for f in fleets}
            vehicle_map = {v["name"].strip().lower(): v["id"] for v in vehicles if v.get("name")}
            # Usuarios: mapear por username y email
            user_map = {}
            for u in users:
                if u.get("username"):
                    user_map[u["username"].strip().lower()] = u["id"]
                if u.get("email"):
                    user_map[u["email"].strip().lower()] = u["id"]

            st.success(f"✅ Consultados: {len(fleets)} flotas · {len(vehicles)} vehículos · {len(users)} usuarios")
            st.divider()
            st.subheader("Resultados")

            for row in rows:
                fleet_name = row["Nombre de flota"].strip()
                fleet_key = fleet_name.lower()

                # Buscar flota
                if fleet_key not in fleet_map:
                    st.error(f"❌ **{fleet_name}** — Flota no encontrada en la cuenta")
                    continue

                fleet = fleet_map[fleet_key]
                fleet_id = fleet["id"]

                # Resolver vehículos
                vehicle_ids = []
                vehicle_errors = []
                for v_name in [v.strip() for v in row["Vehículos"].split(",") if v.strip()]:
                    vid = vehicle_map.get(v_name.lower())
                    if vid:
                        vehicle_ids.append(vid)
                    else:
                        vehicle_errors.append(v_name)

                # Resolver usuarios
                user_ids = []
                user_errors = []
                for u_name in [u.strip() for u in row["Usuarios"].split(",") if u.strip()]:
                    uid = user_map.get(u_name.lower())
                    if uid:
                        user_ids.append(uid)
                    else:
                        user_errors.append(u_name)

                # Advertencias de no encontrados
                if vehicle_errors:
                    st.warning(f"⚠️ **{fleet_name}** — Vehículos no encontrados: {', '.join(vehicle_errors)}")
                if user_errors:
                    st.warning(f"⚠️ **{fleet_name}** — Usuarios no encontrados: {', '.join(user_errors)}")

                # Actualizar flota
                code, resp = update_fleet(fleet_id, fleet_name, vehicle_ids, user_ids, token)
                if code == 200:
                    st.success(f"✅ **{fleet_name}** — Actualizada correctamente ({len(vehicle_ids)} vehículos · {len(user_ids)} usuarios)")
                elif code == 400:
                    st.error(f"❌ **{fleet_name}** — Error de validación: {resp}")
                elif code == 401:
                    st.error(f"❌ **{fleet_name}** — Token inválido o sin permisos")
                elif code is None:
                    st.error(f"❌ **{fleet_name}** — Sin conexión: {resp}")
                else:
                    st.error(f"❌ **{fleet_name}** — Error {code}: {resp}")

    elif fleet_file and not token:
        st.warning("⚠️ Ingresa tu token arriba para continuar.")


# ── FEATURE 5: CREAR WEBHOOK ─────────────────────────────────────────────────
WEBHOOK_EVENTS = {
    "Plan creado": "plan_created",
    "Plan editado": "plan_edited",
    "Ruta creada": "route_created",
    "Ruta editada": "route_edited",
    "Ruta iniciada": "route_started",
    "Ruta finalizada": "route_finished",
    "En camino": "on_its_way",
    "Checkout": "visit_checkout",
    "Checkout detallado": "visit_checkout_detailed",
}

def create_webhook(token, webhook_key, url, headers_payload):
    api_url = "http://api.simpliroute.com/v1/addons/webhooks/"
    headers = {"Authorization": f"Token {token}", "Content-Type": "application/json"}
    payload = {"webhook": webhook_key, "url": url, "headers": headers_payload}
    try:
        r = requests.post(api_url, headers=headers, json=payload, timeout=30)
        return r.status_code, r.json()
    except Exception as e:
        return None, str(e)

def page_crear_webhook():
    st.title("🔔 Crear Webhook")
    st.markdown("Configura un webhook para recibir notificaciones de eventos de SimpliRoute.")

    token = st.text_input("🔑 Token de SimpliRoute", type="password",
                          placeholder="Ingresa tu token aquí", key="token_webhook")

    url = st.text_input("🌐 URL de destino", placeholder="https://tu-servidor.com/webhook")

    evento_label = st.selectbox("📋 Tipo de evento", options=list(WEBHOOK_EVENTS.keys()))
    evento_key = WEBHOOK_EVENTS[evento_label]

    st.markdown("**🔧 Headers**")
    custom_header = st.checkbox("Usar header personalizado")

    if custom_header:
        st.caption("Pega aquí el JSON del header personalizado. El Content-Type se incluye por defecto.")
        header_json_str = st.text_area(
            "Header personalizado (JSON)",
            value='{\n    "Content-Type": "application/json",\n    "X-Custom-Header": "valor"\n}',
            height=150,
            label_visibility="collapsed"
        )
        try:
            import json
            headers_payload = json.loads(header_json_str)
        except Exception:
            st.error("❌ El JSON del header no es válido. Verifica el formato.")
            headers_payload = None
    else:
        headers_payload = {"Content-Type": "application/json"}
        st.code('{"Content-Type": "application/json"}', language="json")

    st.divider()

    if st.button("🚀 Crear Webhook", type="primary", disabled=not (token and url)):
        if not url.startswith("http"):
            st.error("❌ La URL debe comenzar con http:// o https://")
            return
        if custom_header and headers_payload is None:
            st.error("❌ Corrige el JSON del header antes de continuar.")
            return
        with st.spinner("Creando webhook..."):
            code, resp = create_webhook(token, evento_key, url, headers_payload)
        if code in [200, 201]:
            st.success(f"✅ Webhook **{evento_label}** creado correctamente")
            st.json(resp)
        elif code == 400:
            st.error(f"❌ Error de validación: {resp}")
        elif code == 401:
            st.error("❌ Token inválido o sin permisos.")
        elif code is None:
            st.error(f"❌ Sin conexión: {resp}")
        else:
            st.error(f"❌ Error {code}: {resp}")


# ── FEATURE 6: REENVIAR WEBHOOKS ─────────────────────────────────────────────
def resend_webhooks(token, account_id, planned_date, visit_ids):
    url = "https://api.simpliroute.com/v1/mobile/send-webhooks"
    headers = {"Authorization": f"Token {token}", "Content-Type": "application/json"}
    payload = {
        "account_ids": [int(account_id)],
        "planned_date": planned_date,
        "visit_ids": visit_ids
    }
    try:
        r = requests.post(url, headers=headers, json=payload, timeout=60)
        return r.status_code, r.json() if r.content else {}
    except Exception as e:
        return None, str(e)

def page_reenviar_webhooks():
    from datetime import date
    st.title("🔁 Reenviar Webhooks")
    st.markdown("Reenvía webhooks de visitas específicas a tu sistema.")

    st.info("🔑 Usa el token de tu cuenta de SimpliRoute.")
    token = st.text_input("🔑 Token de SimpliRoute", type="password",
                          placeholder="Ingresa tu token aquí", key="token_resend")

    account_id = st.text_input("🏢 ID de la cuenta", placeholder="Ej: 30610")

    st.markdown("**📋 IDs de visita**")
    st.caption("Pega los IDs de visita uno por línea, por ejemplo:\n```\n799841373\n808472905\n819900123\n```")
    visit_ids_raw = st.text_area(
        "IDs de visita",
        placeholder="799841373\n808472905\n819900123",
        height=200,
        label_visibility="collapsed"
    )

    # Fecha de hoy seteada por backend
    today = date.today().strftime("%Y-%m-%d")
    st.caption(f"📅 La fecha utilizada será la de hoy: **{today}**")

    st.divider()

    if st.button("🚀 Reenviar Webhooks", type="primary", disabled=not (token and account_id and visit_ids_raw)):
        # Parsear IDs
        visit_ids = []
        errors = []
        for line in visit_ids_raw.strip().splitlines():
            line = line.strip()
            if not line:
                continue
            try:
                visit_ids.append(int(line))
            except ValueError:
                errors.append(line)

        if errors:
            st.error(f"❌ Los siguientes valores no son IDs válidos: {', '.join(errors)}")
            return
        if not visit_ids:
            st.error("❌ No se encontraron IDs de visita válidos.")
            return

        with st.spinner(f"Reenviando webhooks para {len(visit_ids)} visitas..."):
            code, resp = resend_webhooks(token, account_id.strip(), today, visit_ids)

        if code in [200, 201]:
            st.success(f"✅ Webhooks reenviados correctamente para **{len(visit_ids)} visitas**")
            if resp:
                st.json(resp)
        elif code == 400:
            st.error(f"❌ Error de validación: {resp}")
        elif code == 401:
            st.error("❌ Token inválido o sin permisos.")
        elif code == 404:
            st.error("❌ Cuenta no encontrada. Verifica el ID.")
        elif code is None:
            st.error(f"❌ Sin conexión: {resp}")
        else:
            st.error(f"❌ Error {code}: {resp}")


# ── FEATURE 7: AGREGAR SELLER A VISITAS ──────────────────────────────────────
def get_sellers(token):
    url = "http://api.simpliroute.com/v1/sellers/"
    headers = {"Authorization": f"Token {token}", "Content-Type": "application/json"}
    try:
        r = requests.get(url, headers=headers, timeout=60)
        return r.status_code, r.json()
    except Exception as e:
        return None, str(e)

def patch_visits_seller(token, visit_ids, seller_uuid):
    url = "http://api.simpliroute.com/v1/routes/visits/"
    headers = {"Authorization": f"Token {token}", "Content-Type": "application/json"}
    payload = [{"id": vid, "seller": seller_uuid} for vid in visit_ids]
    try:
        r = requests.patch(url, headers=headers, json=payload, timeout=120)
        return r.status_code, r.json() if r.content else {}
    except Exception as e:
        return None, str(e)

def page_agregar_seller():
    st.title("🧑‍💼 Agregar Seller a Visitas")
    st.markdown("Consulta los sellers de la cuenta, selecciona uno y asígnalo a una lista de visitas.")

    token = st.text_input("🔑 Token de SimpliRoute", type="password",
                          placeholder="Ingresa tu token aquí", key="token_seller")

    if st.button("🔍 Consultar Sellers", type="primary", disabled=not token):
        with st.spinner("Consultando sellers..."):
            code, resp = get_sellers(token)
        if code == 200:
            if not resp:
                st.warning("⚠️ No se encontraron sellers en esta cuenta.")
                st.session_state.pop("sellers", None)
            else:
                st.session_state["sellers"] = resp
                st.session_state["seller_token"] = token
                st.success(f"✅ {len(resp)} seller(s) encontrado(s)")
        elif code == 401:
            st.error("❌ Token inválido o sin permisos.")
            st.session_state.pop("sellers", None)
        elif code is None:
            st.error(f"❌ Sin conexión: {resp}")
            st.session_state.pop("sellers", None)
        else:
            st.error(f"❌ Error {code}: {resp}")
            st.session_state.pop("sellers", None)

    if "sellers" in st.session_state:
        sellers = st.session_state["sellers"]

        st.divider()
        st.subheader("Selecciona un Seller")

        # Construir opciones — usar nombre o email o uuid como fallback
        def seller_label(s):
            name = s.get("name") or s.get("username") or s.get("email") or s.get("uuid", "Sin nombre")
            email = s.get("email", "")
            return f"{name} — {email}" if email and email != name else name

        opciones = {seller_label(s): s for s in sellers}
        selected_label = st.selectbox("Seller:", list(opciones.keys()))
        selected_seller = opciones[selected_label]
        seller_uuid = selected_seller.get("uuid") or selected_seller.get("id")

        st.divider()
        st.subheader("IDs de Visita")
        st.caption("Pega los IDs de visita uno por línea, por ejemplo:\n```\n799841373\n808472905\n819900123\n```")
        visit_ids_raw = st.text_area(
            "IDs de visita",
            placeholder="799841373\n808472905\n819900123",
            height=200,
            label_visibility="collapsed"
        )

        if st.button("💾 Asignar Seller a Visitas", type="primary", disabled=not visit_ids_raw):
            # Parsear IDs
            visit_ids = []
            errors = []
            for line in visit_ids_raw.strip().splitlines():
                line = line.strip()
                if not line:
                    continue
                try:
                    visit_ids.append(int(line))
                except ValueError:
                    errors.append(line)

            if errors:
                st.error(f"❌ Los siguientes valores no son IDs válidos: {', '.join(errors)}")
            elif not visit_ids:
                st.error("❌ No se encontraron IDs de visita válidos.")
            else:
                with st.spinner(f"Asignando seller a {len(visit_ids)} visitas..."):
                    code, resp = patch_visits_seller(
                        st.session_state["seller_token"], visit_ids, seller_uuid
                    )
                if code in [200, 201]:
                    st.success(f"✅ Seller **{selected_label}** asignado correctamente a **{len(visit_ids)} visitas**")
                elif code == 400:
                    st.error(f"❌ Error de validación: {resp}")
                elif code == 401:
                    st.error("❌ Token inválido o sin permisos.")
                elif code is None:
                    st.error(f"❌ Sin conexión: {resp}")
                else:
                    st.error(f"❌ Error {code}: {resp}")


# ── ROUTER ────────────────────────────────────────────────────────────────────
if menu == "🗺️ Cargar Zonas":
    page_cargar_zonas()
elif menu == "👤 Cambiar Rol de Usuario":
    page_cambiar_rol()
elif menu == "🏷️ Tipos de Visita y Skills":
    page_visit_types_skills()
elif menu == "🚛 Asignación de Flotas":
    page_asignacion_flotas()
elif menu == "🔔 Crear Webhook":
    page_crear_webhook()
elif menu == "🧑‍💼 Agregar Seller a Visitas":
    page_agregar_seller()