import streamlit as st
import openpyxl
import io
import re
import requests
import json
from datetime import date

st.set_page_config(page_title="SimpliRoute Tools", page_icon="🚀", layout="wide")

# ── NAVEGACIÓN ────────────────────────────────────────────────────────────────
CORE_OPTIONS = [
    "🧑‍💼 Agregar Seller a Visitas",
    "🚛 Asignación de Flotas",
    "🗺️ Cargar Zonas",
    "👤 Cambiar Rol de Usuario",
    "🔔 Crear Webhook",
    "🔓 Desbloqueo de Contraseña",
    "🚦 Iniciar / Cerrar Rutas",
    "🔁 Reenviar Webhooks",
    "🏷️ Tipos de Visita y Skills",
    "📡 Validación de GPS",
]
TMS_OPTIONS = [
    "📄 Tipos de Documento",
    "🚚 Transportistas",
]

if "current_page" not in st.session_state:
    st.session_state["current_page"] = CORE_OPTIONS[0]

def nav_item(label):
    is_active = st.session_state["current_page"] == label
    if st.sidebar.button(label, key=f"nav_{label}", use_container_width=True,
                         type="primary" if is_active else "secondary"):
        st.session_state["current_page"] = label
        st.rerun()

with st.sidebar:
    st.markdown("## 🚀 SimpliRoute Tools")
    st.markdown("---")
    st.markdown("### 🛠️ SimpliRoute Core")
    for item in CORE_OPTIONS:
        nav_item(item)
    st.markdown("---")
    st.markdown("### 📦 TMS")
    for item in TMS_OPTIONS:
        nav_item(item)
    st.markdown("---")
    st.caption("SimpliRoute Internal Tools v1.0")

selected = st.session_state["current_page"]


# ── HELPERS ───────────────────────────────────────────────────────────────────
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
    for block in re.findall(r'<Placemark>(.*?)</Placemark>', text, re.DOTALL):
        if '<Polygon>' not in block and '<Polygon ' not in block:
            continue
        name_match = re.search(r'<name>\s*(.*?)\s*</name>', block)
        name = name_match.group(1).strip() if name_match else "Sin nombre"
        coords_match = re.search(r'<coordinates>\s*(.*?)\s*</coordinates>', block, re.DOTALL)
        if not coords_match:
            continue
        points = []
        for token_str in coords_match.group(1).strip().split():
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
            if -90 <= lat <= 90 and -180 <= lng <= 180:
                points.append((lat, lng))
        if points:
            polygons.append({"name": name, "coords": points})
    return polygons

def coords_to_str(coords):
    return "[" + ",".join(f"{{'lat':'{lat:.6f}','lng':'{lng:.6f}'}}" for lat, lng in coords) + "]"

def generate_excel_zones(polygons):
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Hoja 1"
    ws.append(["is_name", "is_coordinates"])
    for p in polygons:
        ws.append([p["name"], coords_to_str(p["coords"])])
    ws.column_dimensions["A"].width = 25; ws.column_dimensions["B"].width = 80
    buf = io.BytesIO(); wb.save(buf); buf.seek(0); return buf

def read_excel_column(file, columns):
    wb = openpyxl.load_workbook(file); ws = wb.active
    headers = [str(c.value).strip() if c.value else "" for c in ws[1]]
    rows = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if all(v is None for v in row):
            continue
        entry = {col: str(row[headers.index(col)] if col in headers and headers.index(col) < len(row) else "").strip() for col in columns}
        if any(entry[c] for c in columns):
            rows.append(entry)
    return rows

def show_results(results, name_key):
    ok = sum(1 for r in results if r["code"] in [200, 201])
    st.markdown(f"✅ **{ok} creados** | ❌ **{len(results)-ok} con error**")
    for r in results:
        if r["code"] in [200, 201]: st.success(f"✅ {r[name_key]} — OK")
        elif r["code"] == 400: st.warning(f"⚠️ {r[name_key]} — Ya existe o datos inválidos")
        elif r["code"] == 401: st.error(f"❌ {r[name_key]} — Token inválido")
        elif r["code"] is None: st.error(f"❌ {r[name_key]} — Sin conexión")
        else: st.error(f"❌ {r[name_key]} — Error {r['code']}")

def get_users_list(token):
    try:
        r = requests.get("http://api.simpliroute.com/v1/accounts/users/",
                         headers={"Authorization": f"Token {token}", "accept": "application/json"}, timeout=300)
        return r.status_code, r.json()
    except Exception as e:
        return None, str(e)

def put_user_full(token, user, new_email):
    try:
        r = requests.put(f"http://api.simpliroute.com/v1/accounts/users/{user['id']}/",
                         headers={"Authorization": f"Token {token}", "Content-Type": "application/json"},
                         json={"id": user["id"], "username": user.get("username",""), "name": user.get("name",""),
                               "phone": user.get("phone",""), "email": new_email,
                               "is_owner": user.get("is_owner",False), "is_admin": user.get("is_admin",False),
                               "is_driver": user.get("is_driver",False), "is_codriver": user.get("is_codriver",False),
                               "is_router_jr": user.get("is_router_jr",False), "is_monitor": user.get("is_monitor",False),
                               "is_coordinator": user.get("is_coordinator",False), "is_router": user.get("is_router",False),
                               "is_staff": user.get("is_staff",False), "is_seller_viewer": user.get("is_seller_viewer",False),
                               "is_seller": user.get("is_seller",False), "blocked": user.get("blocked",False),
                               "status": user.get("status","active")}, timeout=30)
        return r.status_code, r.json()
    except Exception as e:
        return None, str(e)


# ── FEATURE: AGREGAR SELLER A VISITAS ────────────────────────────────────────
def page_agregar_seller():
    st.title("🧑‍💼 Agregar Seller a Visitas")
    token = st.text_input("🔑 Token de SimpliRoute", type="password", key="token_seller")
    if st.button("🔍 Consultar Sellers", type="primary", disabled=not token):
        with st.spinner("Consultando..."):
            try:
                r = requests.get("http://api.simpliroute.com/v1/sellers/",
                                 headers={"Authorization": f"Token {token}", "Content-Type": "application/json"}, timeout=60)
                code, resp = r.status_code, r.json()
            except Exception as e:
                code, resp = None, str(e)
        if code == 200:
            st.session_state["sellers"] = resp; st.session_state["seller_token"] = token
            st.success(f"✅ {len(resp)} seller(s) encontrado(s)")
        else:
            st.error(f"❌ Error {code}: {resp}")
            st.session_state.pop("sellers", None)

    if "sellers" in st.session_state:
        sellers = st.session_state["sellers"]
        st.divider()
        def slabel(s):
            name = s.get("name") or s.get("username") or s.get("email") or str(s.get("uuid",""))
            email = s.get("email","")
            return f"{name} — {email}" if email and email != name else name
        opciones = {slabel(s): s for s in sellers}
        selected_s = opciones[st.selectbox("Seller:", list(opciones.keys()))]
        seller_uuid = selected_s.get("uuid") or selected_s.get("id")
        st.divider()
        st.caption("Pega los IDs de visita uno por línea:\n```\n799841373\n808472905\n```")
        visit_ids_raw = st.text_area("IDs", placeholder="799841373\n808472905", height=200, label_visibility="collapsed")
        if st.button("💾 Asignar Seller", type="primary", disabled=not visit_ids_raw):
            ids, errs = [], []
            for l in visit_ids_raw.strip().splitlines():
                l = l.strip()
                try: ids.append(int(l))
                except: errs.append(l)
            if errs: st.error(f"❌ IDs inválidos: {', '.join(errs)}"); return
            with st.spinner(f"Asignando a {len(ids)} visitas..."):
                try:
                    r = requests.patch("http://api.simpliroute.com/v1/routes/visits/",
                                       headers={"Authorization": f"Token {st.session_state['seller_token']}", "Content-Type": "application/json"},
                                       json=[{"id": vid, "seller": seller_uuid} for vid in ids], timeout=120)
                    code = r.status_code
                except Exception as e:
                    code = None
            if code in [200,201]: st.success(f"✅ Seller asignado a {len(ids)} visitas")
            else: st.error(f"❌ Error {code}")


# ── FEATURE: ASIGNACIÓN DE FLOTAS ────────────────────────────────────────────
def page_asignacion_flotas():
    st.title("🚛 Asignación de Flotas")
    st.info("ℹ️ Cada fila representa una edición independiente. Si la misma flota aparece dos veces, prevalece la última.")
    token = st.text_input("🔑 Token de SimpliRoute", type="password", key="token_flotas")
    def make_tpl():
        wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Flotas"
        ws.append(["Nombre de flota","Vehículos","Usuarios"])
        ws.append(["Flota Norte","MC4327,QC4380","juan.perez,ana.lopez"])
        ws.column_dimensions["A"].width=25; ws.column_dimensions["B"].width=40; ws.column_dimensions["C"].width=40
        buf=io.BytesIO(); wb.save(buf); buf.seek(0); return buf
    st.download_button("📥 Descargar plantilla", data=make_tpl(), file_name="plantilla_flotas.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    fleet_file = st.file_uploader("📂 Sube tu Excel", type=["xlsx"], key="upload_flotas")
    if fleet_file and token:
        if st.button("🚀 Procesar flotas", type="primary"):
            try:
                rows = read_excel_column(fleet_file, ["Nombre de flota","Vehículos","Usuarios"])
            except Exception as e:
                st.error(f"❌ Error leyendo Excel: {e}"); return
            with st.spinner("⏳ Consultando flotas, vehículos y usuarios..."):
                try:
                    rf = requests.get("http://api.simpliroute.com/v1/fleets/", headers={"Authorization":f"Token {token}","accept":"application/json"}, timeout=300)
                    rv = requests.get("http://api.simpliroute.com/v1/routes/vehicles/", headers={"Authorization":f"Token {token}"}, timeout=300)
                    ru = get_users_list(token)
                    fleets, vehicles = rf.json(), rv.json()
                    code_u, users = ru
                except Exception as e:
                    st.error(f"❌ Error consultando API: {e}"); return
            fleet_map = {f["name"].strip().lower(): f for f in fleets}
            vehicle_map = {v["name"].strip().lower(): v["id"] for v in vehicles if v.get("name")}
            user_map = {}
            for u in (users if isinstance(users, list) else []):
                if u.get("username"): user_map[u["username"].strip().lower()] = u["id"]
                if u.get("email"): user_map[u["email"].strip().lower()] = u["id"]
            st.success(f"✅ {len(fleets)} flotas · {len(vehicles)} vehículos · {len(user_map)} usuarios")
            st.divider()
            for row in rows:
                fname = row["Nombre de flota"].strip()
                if fname.lower() not in fleet_map:
                    st.error(f"❌ **{fname}** — Flota no encontrada"); continue
                fleet = fleet_map[fname.lower()]
                vids, verrs = [], []
                for v in [x.strip() for x in row["Vehículos"].split(",") if x.strip()]:
                    vid = vehicle_map.get(v.lower())
                    (vids if vid else verrs).append(vid if vid else v)
                uids, uerrs = [], []
                for u in [x.strip() for x in row["Usuarios"].split(",") if x.strip()]:
                    uid = user_map.get(u.lower())
                    (uids if uid else uerrs).append(uid if uid else u)
                if verrs: st.warning(f"⚠️ **{fname}** — Vehículos no encontrados: {', '.join(str(x) for x in verrs)}")
                if uerrs: st.warning(f"⚠️ **{fname}** — Usuarios no encontrados: {', '.join(str(x) for x in uerrs)}")
                try:
                    r = requests.put(f"http://api.simpliroute.com/v1/fleets/{fleet['id']}/",
                                     headers={"Authorization":f"Token {token}","Content-Type":"application/json"},
                                     json={"id":fleet["id"],"name":fname,"vehicles":vids,"users":uids}, timeout=300)
                    code = r.status_code
                except Exception as e:
                    code = None
                if code == 200: st.success(f"✅ **{fname}** — Actualizada ({len(vids)} vehículos · {len(uids)} usuarios)")
                else: st.error(f"❌ **{fname}** — Error {code}")
    elif fleet_file and not token:
        st.warning("⚠️ Ingresa tu token.")


# ── FEATURE: CARGAR ZONAS ─────────────────────────────────────────────────────
def page_cargar_zonas():
    st.title("🗺️ Cargar Zonas")
    token = st.text_input("🔑 Token de SimpliRoute", type="password", key="token_zonas")
    uploaded = st.file_uploader("📂 Sube tu archivo KML o RTF", type=["kml","rtf","txt"])
    if uploaded:
        text = decode_file(uploaded.read())
        if text.strip().startswith("{\\rtf") or "\\rtf" in text[:100]:
            text = strip_rtf_codes(text)
        polygons = parse_polygons(text)
        if not polygons:
            st.warning("No se encontraron polígonos."); return
        st.success(f"✅ {len(polygons)} polígono(s)")
        for p in polygons:
            with st.expander(f"📍 {p['name']} — {len(p['coords'])} puntos"):
                st.code(coords_to_str(p["coords"][:3]) + ",...]")
        st.download_button("⬇️ Descargar Excel", data=generate_excel_zones(polygons), file_name="ZONES.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        st.divider()
        if not token:
            st.warning("⚠️ Ingresa tu token para cargar zonas.")
        elif st.button("🚀 Cargar zonas", type="primary"):
            prog = st.progress(0); status = st.empty()
            for i, p in enumerate(polygons):
                status.info(f"Cargando: **{p['name']}** ({i+1}/{len(polygons)})")
                try:
                    r = requests.post("http://api.simpliroute.com/v1/zones/",
                                      headers={"authorization":f"Token {token}","content-type":"application/json"},
                                      json={"name":p["name"],"coordinates":coords_to_str(p["coords"]),"vehicles":[]}, timeout=15)
                    code = r.status_code
                except: code = None
                if code in [200,201]: st.success(f"✅ {p['name']}")
                elif code == 400: st.warning(f"⚠️ {p['name']} — Ya existe")
                elif code == 401: st.error(f"❌ {p['name']} — Token inválido")
                else: st.error(f"❌ {p['name']} — Error {code}")
                prog.progress((i+1)/len(polygons))
            status.empty(); prog.empty()


# ── FEATURE: CAMBIAR ROL ──────────────────────────────────────────────────────
ROLES = {"Administrador":"is_admin","Conductor":"is_driver","Co-conductor":"is_codriver",
         "Router Jr":"is_router_jr","Monitor":"is_monitor","Coordinador":"is_coordinator",
         "Router":"is_router","Staff":"is_staff","Seller Viewer":"is_seller_viewer","Seller":"is_seller"}
ALL_ROLE_KEYS = list(ROLES.values()) + ["is_owner"]

def get_current_role(user):
    for label, key in ROLES.items():
        if user.get(key): return label, key
    if user.get("is_owner"): return "Owner", "is_owner"
    return "Sin rol", None

def page_cambiar_rol():
    st.title("👤 Cambiar Rol de Usuario")
    col1, col2 = st.columns(2)
    with col1: token = st.text_input("🔑 Token", type="password", key="token_rol")
    with col2: user_id = st.text_input("🆔 ID del Usuario", placeholder="Ej: 524614")
    if st.button("🔍 Consultar", type="primary", disabled=not (token and user_id)):
        try:
            r = requests.get(f"http://api.simpliroute.com/v1/accounts/users/{user_id.strip()}/",
                             headers={"Authorization":f"Token {token}","Content-Type":"application/json"}, timeout=15)
            code, resp = r.status_code, r.json()
        except Exception as e:
            code, resp = None, str(e)
        if code == 200:
            st.session_state["user_data"] = resp; st.session_state["user_token"] = token
        else:
            st.error(f"❌ Error {code}: {resp}")
            st.session_state.pop("user_data", None)
    if "user_data" in st.session_state:
        user = st.session_state["user_data"]
        cl, ck = get_current_role(user)
        st.divider()
        c1,c2,c3 = st.columns(3)
        c1.markdown(f"**Nombre**<br>{user.get('name','—')}", unsafe_allow_html=True)
        c2.markdown(f"**Username**<br>{user.get('username','—')}", unsafe_allow_html=True)
        c3.markdown(f"**Rol actual**<br>{cl}", unsafe_allow_html=True)
        st.caption(f"📧 {user.get('email','—')} · {user.get('status','—')}")
        st.divider()
        if ck == "is_owner":
            st.warning("⚠️ Este usuario es Owner. No se puede modificar.")
        else:
            nuevo_label = st.selectbox("Nuevo rol:", [l for l, k in ROLES.items() if k != ck])
            nuevo_key = ROLES[nuevo_label]
            if st.button(f"💾 Cambiar a {nuevo_label}", type="primary"):
                payload = {k: (k == nuevo_key) for k in ALL_ROLE_KEYS}
                payload["username"] = user.get("username",""); payload["name"] = user.get("name","")
                try:
                    r = requests.put(f"http://api.simpliroute.com/v1/accounts/users/{user['id']}/",
                                     headers={"Authorization":f"Token {st.session_state['user_token']}","Content-Type":"application/json"},
                                     json=payload, timeout=15)
                    code = r.status_code
                except: code = None
                if code == 200:
                    st.success(f"✅ Rol actualizado a **{nuevo_label}**")
                    try:
                        r2 = requests.get(f"http://api.simpliroute.com/v1/accounts/users/{user['id']}/",
                                          headers={"Authorization":f"Token {st.session_state['user_token']}"}, timeout=15)
                        if r2.status_code == 200: st.session_state["user_data"] = r2.json()
                    except: pass
                    st.rerun()
                else: st.error(f"❌ Error {code}")


# ── FEATURE: CREAR WEBHOOK ────────────────────────────────────────────────────
WEBHOOK_EVENTS = {"Plan creado":"plan_created","Plan editado":"plan_edited","Ruta creada":"route_created",
                  "Ruta editada":"route_edited","Ruta iniciada":"route_started","Ruta finalizada":"route_finished",
                  "En camino":"on_its_way","Checkout":"visit_checkout","Checkout detallado":"visit_checkout_detailed"}

def page_crear_webhook():
    st.title("🔔 Crear Webhook")
    token = st.text_input("🔑 Token de SimpliRoute", type="password", key="token_webhook")
    url = st.text_input("🌐 URL de destino", placeholder="https://tu-servidor.com/webhook")
    evento_key = WEBHOOK_EVENTS[st.selectbox("📋 Tipo de evento", list(WEBHOOK_EVENTS.keys()))]
    custom = st.checkbox("Usar header personalizado")
    if custom:
        raw = st.text_area("Header JSON", value='{\n    "Content-Type": "application/json",\n    "X-Custom": "valor"\n}',
                           height=150, label_visibility="collapsed")
        try: headers_payload = json.loads(raw)
        except: st.error("❌ JSON inválido"); headers_payload = None
    else:
        headers_payload = {"Content-Type": "application/json"}
        st.code('{"Content-Type": "application/json"}', language="json")
    st.divider()
    if st.button("🚀 Crear Webhook", type="primary", disabled=not (token and url)):
        if not url.startswith("http"): st.error("❌ URL inválida"); return
        if custom and not headers_payload: st.error("❌ Corrige el JSON"); return
        try:
            r = requests.post("http://api.simpliroute.com/v1/addons/webhooks/",
                              headers={"Authorization":f"Token {token}","Content-Type":"application/json"},
                              json={"webhook":evento_key,"url":url,"headers":headers_payload}, timeout=30)
            code, resp = r.status_code, r.json()
        except Exception as e:
            code, resp = None, str(e)
        if code in [200,201]: st.success("✅ Webhook creado"); st.json(resp)
        elif code == 401: st.error("❌ Token inválido")
        else: st.error(f"❌ Error {code}: {resp}")


# ── FEATURE: DESBLOQUEO DE CONTRASEÑA ────────────────────────────────────────
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
    token = st.text_input("🔑 Token de SimpliRoute", type="password", key="token_unlock")

    if st.button("🔍 Consultar usuarios bloqueados", type="primary", disabled=not token):
        with st.spinner("Consultando..."):
            code, resp = get_users_list(token)
        if code == 200:
            blocked = [u for u in resp if u.get("blocked") is True]
            if not blocked:
                st.info("✅ No hay usuarios bloqueados.")
                st.session_state.pop("blocked_users", None)
            else:
                st.session_state["blocked_users"] = blocked
                st.session_state["unlock_token"] = token
        else:
            st.error(f"❌ Error {code}")

    if "pending_restore" in st.session_state:
        pr = st.session_state["pending_restore"]
        st.divider()
        st.subheader("📧 Restauración pendiente")
        st.info(f"Usuario: **{pr['user'].get('username')}** · Email original: **{pr['original_email'] or 'Sin email'}**")
        st.markdown("Una vez que el usuario haya reseteado su contraseña, presiona el botón para restaurar su email original.")
        if st.button("↩️ Confirmar y restaurar email original", type="primary", key="btn_restore"):
            code, _ = put_user_full(pr["token"], pr["user"], pr["original_email"])
            if code in [200,201]:
                st.success("✅ Email restaurado."); st.session_state.pop("pending_restore", None)
            else:
                st.error(f"❌ Error {code}")

    if "blocked_users" in st.session_state:
        blocked = st.session_state["blocked_users"]
        st.divider()
        st.success(f"⚠️ {len(blocked)} usuario(s) bloqueado(s)")

        def ulabel(u):
            name = u.get("name") or u.get("username") or str(u.get("id"))
            un = u.get("username","")
            return f"{name} — {un}" if un != name else name

        selected_user = {ulabel(u): u for u in blocked}[st.selectbox("Usuario bloqueado:", [ulabel(u) for u in blocked])]
        original_email = selected_user.get("email","")
        st.caption(f"📧 Email actual: **{original_email or 'Sin email'}**")
        st.divider()
        agent_name = st.selectbox("Agente que recibirá el link:", list(SUPPORT_AGENTS.keys()))
        recovery_email = SUPPORT_AGENTS[agent_name]
        st.caption(f"📨 Se enviará a: **{recovery_email}**")
        st.divider()
        if st.button("🔓 Enviar link de desbloqueo", type="primary"):
            tok = st.session_state["unlock_token"]
            prog = st.progress(0); status = st.empty()
            status.info("1/2 — Asignando email temporal...")
            code1, resp1 = put_user_full(tok, selected_user, recovery_email)
            prog.progress(50)
            if code1 not in [200,201]:
                st.error(f"❌ Error al actualizar email: {code1}"); status.empty(); prog.empty()
            else:
                status.info("2/2 — Enviando link...")
                try:
                    r = requests.post("https://api.simpliroute.com/v2/auth/unlock/",
                                      headers={"Content-Type":"application/json;charset=UTF-8","authorization":"null"},
                                      json={"username": selected_user.get("username")}, timeout=30)
                    code2 = r.status_code
                except Exception as e:
                    code2 = None
                prog.progress(100); status.empty(); prog.empty()
                if code2 in [200,201]:
                    st.success(f"✅ Link enviado a **{recovery_email}** ({agent_name})")
                    st.warning("⚠️ Cuando el usuario resetee su contraseña, usa el botón de restauración de arriba.")
                    st.session_state["pending_restore"] = {"user": selected_user, "original_email": original_email, "token": tok}
                    st.session_state.pop("blocked_users", None)
                    st.rerun()
                else:
                    put_user_full(tok, selected_user, original_email)
                    st.error(f"❌ No se pudo enviar el link. Email restaurado. ({code2})")


# ── FEATURE: REENVIAR WEBHOOKS ────────────────────────────────────────────────
def page_reenviar_webhooks():
    st.title("🔁 Reenviar Webhooks")
    st.info("🔑 Usa el token de tu cuenta de SimpliRoute.")
    token = st.text_input("🔑 Token", type="password", key="token_resend")
    account_id = st.text_input("🏢 ID de la cuenta", placeholder="Ej: 30610")
    st.caption("Pega los IDs de visita uno por línea:\n```\n799841373\n808472905\n```")
    visit_ids_raw = st.text_area("IDs", placeholder="799841373\n808472905", height=200, label_visibility="collapsed")
    today = date.today().strftime("%Y-%m-%d")
    st.caption(f"📅 Fecha: **{today}**")
    st.divider()
    if st.button("🚀 Reenviar", type="primary", disabled=not (token and account_id and visit_ids_raw)):
        ids, errs = [], []
        for l in visit_ids_raw.strip().splitlines():
            l = l.strip()
            try: ids.append(int(l))
            except: errs.append(l)
        if errs: st.error(f"❌ IDs inválidos: {', '.join(errs)}"); return
        with st.spinner(f"Reenviando {len(ids)} visitas..."):
            try:
                r = requests.post("https://api.simpliroute.com/v1/mobile/send-webhooks",
                                  headers={"Authorization":f"Token {token}","Content-Type":"application/json"},
                                  json={"account_ids":[int(account_id)],"planned_date":today,"visit_ids":ids}, timeout=60)
                code = r.status_code
            except: code = None
        if code in [200,201]: st.success(f"✅ Webhooks reenviados para {len(ids)} visitas")
        elif code == 401: st.error("❌ Token inválido")
        elif code == 404: st.error("❌ Cuenta no encontrada")
        else: st.error(f"❌ Error {code}")


# ── FEATURE: TIPOS DE VISITA Y SKILLS ────────────────────────────────────────
def page_visit_types_skills():
    st.title("🏷️ Tipos de Visita y Skills")
    token = st.text_input("🔑 Token de SimpliRoute", type="password", key="token_vts")

    st.markdown("---"); st.subheader("🔧 Skills")
    def tpl_skills():
        wb=openpyxl.Workbook(); ws=wb.active; ws.title="Skills"
        ws.append(["skill"]); ws.append(["Manejo de carga pesada"]); ws.append(["Refrigeración"])
        ws.column_dimensions["A"].width=35; buf=io.BytesIO(); wb.save(buf); buf.seek(0); return buf
    st.download_button("📥 Plantilla Skills", data=tpl_skills(), file_name="plantilla_skills.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    sf = st.file_uploader("📂 Excel Skills", type=["xlsx"], key="upload_skills")
    if sf:
        try:
            rows = read_excel_column(sf, ["skill"])
            st.success(f"✅ {len(rows)} skill(s)")
            with st.expander("Ver"): [st.markdown(f"- {r['skill']}") for r in rows]
        except Exception as e:
            rows = []; st.error(str(e))
        if rows and token and st.button("🚀 Crear Skills", type="primary", key="btn_skills"):
            results=[]; prog=st.progress(0); status=st.empty()
            for i, row in enumerate(rows):
                status.info(f"Creando: **{row['skill']}** ({i+1}/{len(rows)})")
                try:
                    r=requests.post("http://api.simpliroute.com/v1/routes/skills/",
                                    headers={"Authorization":f"Token {token}","Content-Type":"application/json"},
                                    json={"skill":row["skill"]}, timeout=15)
                    code,resp=r.status_code,r.json()
                except Exception as e: code,resp=None,str(e)
                results.append({"skill":row["skill"],"code":code,"resp":resp}); prog.progress((i+1)/len(rows))
            status.empty(); prog.empty(); show_results(results, "skill")

    st.markdown("---"); st.subheader("📋 Tipos de Visita")
    def tpl_vt():
        wb=openpyxl.Workbook(); ws=wb.active; ws.title="Tipos de Visita"
        ws.append(["label","key"]); ws.append(["Entrega express","entrega_express"]); ws.append(["Retiro","retiro"])
        ws.column_dimensions["A"].width=30; ws.column_dimensions["B"].width=30
        buf=io.BytesIO(); wb.save(buf); buf.seek(0); return buf
    st.download_button("📥 Plantilla Tipos de Visita", data=tpl_vt(), file_name="plantilla_tipos_visita.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    vf = st.file_uploader("📂 Excel Tipos de Visita", type=["xlsx"], key="upload_vt")
    if vf:
        try:
            rows_vt = read_excel_column(vf, ["label","key"])
            st.success(f"✅ {len(rows_vt)} tipo(s)")
            with st.expander("Ver"): [st.markdown(f"- **{r['label']}** → `{r['key']}`") for r in rows_vt]
        except Exception as e:
            rows_vt=[]; st.error(str(e))
        if rows_vt and token and st.button("🚀 Crear Tipos de Visita", type="primary", key="btn_vt"):
            results=[]; prog=st.progress(0); status=st.empty()
            for i, row in enumerate(rows_vt):
                status.info(f"Creando: **{row['label']}** ({i+1}/{len(rows_vt)})")
                try:
                    r=requests.post("http://api.simpliroute.com/v1/accounts/visit-types/",
                                    headers={"Authorization":f"Token {token}","Content-Type":"application/json"},
                                    json=[{"label":row["label"],"key":row["key"]}], timeout=15)
                    code,resp=r.status_code,r.json()
                except Exception as e: code,resp=None,str(e)
                results.append({"label":row["label"],"code":code,"resp":resp}); prog.progress((i+1)/len(rows_vt))
            status.empty(); prog.empty(); show_results(results, "label")


# ── FEATURE: INICIAR / CERRAR RUTAS ──────────────────────────────────────────
def page_iniciar_cerrar_rutas():
    st.title("🚦 Iniciar / Cerrar Rutas")
    st.markdown("Registra el evento de inicio o cierre para una lista de rutas.")

    token = st.text_input("🔑 Token de SimpliRoute", type="password", key="token_routes")

    evento = st.selectbox("📋 Tipo de evento", ["Iniciar ruta", "Finalizar ruta"])
    event_type = "ROUTE_STARTED" if evento == "Iniciar ruta" else "ROUTE_FINISHED"

    selected_date = st.date_input("📅 Fecha", value=date.today())
    hora = "12:00:00.000Z" if event_type == "ROUTE_STARTED" else "22:00:00.000Z"
    date_time = f"{selected_date.strftime('%Y-%m-%d')}T{hora}"
    st.caption(f"📅 Fecha y hora que se usará: **{date_time}**")

    st.divider()
    st.caption("Pega los IDs de ruta uno por línea, por ejemplo:\n```\n637f11a2-a1a6-4609-8c23-83e8c76dccbf\nb9055313-a698-4af9-a676-fe63ca175ace\n```")
    route_ids_raw = st.text_area("IDs de ruta", placeholder="637f11a2-a1a6-4609-8c23-83e8c76dccbf\nb9055313-...",
                                 height=200, label_visibility="collapsed")

    if st.button(f"🚀 {evento}", type="primary", disabled=not (token and route_ids_raw)):
        route_ids = [l.strip() for l in route_ids_raw.strip().splitlines() if l.strip()]
        if not route_ids:
            st.error("❌ No se encontraron IDs válidos."); return

        prog = st.progress(0); status = st.empty()

        for i, route_id in enumerate(route_ids):
            status.info(f"Procesando: **{route_id}** ({i+1}/{len(route_ids)})")

            # GET info de la ruta
            try:
                r = requests.get(f"http://api.simpliroute.com/v1/routes/routes/{route_id}/",
                                 headers={"Authorization": f"Token {token}", "Content-Type": "application/json"},
                                 timeout=30)
                code_get, route_data = r.status_code, r.json()
            except Exception as e:
                st.error(f"❌ **{route_id}** — Error consultando ruta: {e}")
                prog.progress((i+1)/len(route_ids)); continue

            if code_get != 200:
                st.error(f"❌ **{route_id}** — Error {code_get} al consultar ruta")
                prog.progress((i+1)/len(route_ids)); continue

            # Seleccionar lat/lng según el tipo de evento
            if event_type == "ROUTE_STARTED":
                lat = route_data.get("location_start_latitude")
                lng = route_data.get("location_start_longitude")
            else:
                lat = route_data.get("location_end_latitude")
                lng = route_data.get("location_end_longitude")

            if not lat or not lng:
                st.warning(f"⚠️ **{route_id}** — Sin coordenadas disponibles para este evento")
                prog.progress((i+1)/len(route_ids)); continue

            # POST evento
            try:
                r2 = requests.post("http://api.simpliroute.com/v1/events/register/",
                                   headers={"Authorization": f"Token {token}", "Content-Type": "application/json"},
                                   json={"date_time": date_time, "latitude": float(lat),
                                         "longitude": float(lng), "route_id": route_id, "type": event_type},
                                   timeout=30)
                code_post = r2.status_code
            except Exception as e:
                st.error(f"❌ **{route_id}** — Error registrando evento: {e}")
                prog.progress((i+1)/len(route_ids)); continue

            if code_post in [200, 201]:
                st.success(f"✅ **{route_id}** — {evento} registrado correctamente")
            elif code_post == 401:
                st.error(f"❌ **{route_id}** — Token inválido")
            else:
                st.error(f"❌ **{route_id}** — Error {code_post}")

            prog.progress((i+1)/len(route_ids))

        status.empty(); prog.empty()



def page_tms_document_types():
    st.title("📄 Tipos de Documento TMS")
    COUNTRY_OPTIONS = ["PE","CL","MX","CO","AR","EC","BO","UY","PY","VE"]
    c1,c2,c3 = st.columns(3)
    with c1: token = st.text_input("🔑 Token", type="password", key="token_doctype")
    with c2: account_id = st.text_input("🏢 Account ID", placeholder="Ej: 9695")
    with c3: country = st.selectbox("🌎 País", COUNTRY_OPTIONS)
    def tpl():
        wb=openpyxl.Workbook(); ws=wb.active; ws.title="Tipos de Documento"
        ws.append(["name","entity_type"]); ws.append(["DNI","driver"]); ws.append(["RUC","provider"])
        ws.column_dimensions["A"].width=25; ws.column_dimensions["B"].width=25
        buf=io.BytesIO(); wb.save(buf); buf.seek(0); return buf
    st.download_button("📥 Plantilla", data=tpl(), file_name="plantilla_tipos_documento.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    doc_file = st.file_uploader("📂 Excel", type=["xlsx"], key="upload_doctype")
    if doc_file:
        try:
            rows = read_excel_column(doc_file, ["name","entity_type"])
            st.success(f"✅ {len(rows)} tipo(s)")
            with st.expander("Ver"): [st.markdown(f"- **{r['name']}** → `{r['entity_type']}`") for r in rows]
        except Exception as e:
            rows=[]; st.error(str(e))
        if rows and token and account_id and st.button("🚀 Crear", type="primary", key="btn_doctype"):
            results=[]; prog=st.progress(0); status=st.empty()
            for i, row in enumerate(rows):
                status.info(f"Creando: **{row['name']}** ({i+1}/{len(rows)})")
                try:
                    r=requests.post("https://api.simpliroute.com/tms/api/v1/document-types/",
                                    headers={"Authorization":f"Token {token}","Related-Account":account_id,"Content-Type":"application/json"},
                                    json={"name":row["name"],"country":country,"entity_type":row["entity_type"],"status":"active"}, timeout=30)
                    code,resp=r.status_code,r.json()
                except Exception as e: code,resp=None,str(e)
                results.append({"name":row["name"],"code":code,"resp":resp}); prog.progress((i+1)/len(rows))
            status.empty(); prog.empty(); show_results(results, "name")


# ── TMS: TRANSPORTISTAS ───────────────────────────────────────────────────────
def page_tms_transportistas():
    st.title("🚚 Transportistas TMS")
    c1,c2 = st.columns(2)
    with c1: token = st.text_input("🔑 Token", type="password", key="token_transp")
    with c2: account_id = st.text_input("🏢 Account ID", placeholder="Ej: 82761")
    def tpl():
        wb=openpyxl.Workbook(); ws=wb.active; ws.title="Transportistas"
        ws.append(["trade_name","legal_name","tax_id_number"])
        ws.append(["Transportes Sur S.A.","Transportes Sur Sociedad Anónima","20123456789"])
        ws.column_dimensions["A"].width=30; ws.column_dimensions["B"].width=35; ws.column_dimensions["C"].width=20
        buf=io.BytesIO(); wb.save(buf); buf.seek(0); return buf
    st.download_button("📥 Plantilla", data=tpl(), file_name="plantilla_transportistas.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    tf = st.file_uploader("📂 Excel", type=["xlsx"], key="upload_transp")
    if tf:
        try:
            rows = read_excel_column(tf, ["trade_name","legal_name","tax_id_number"])
            st.success(f"✅ {len(rows)} transportista(s)")
            with st.expander("Ver"): [st.markdown(f"- **{r['trade_name']}** | {r['legal_name']} | `{r['tax_id_number']}`") for r in rows]
        except Exception as e:
            rows=[]; st.error(str(e))
        if rows and token and account_id and st.button("🚀 Crear", type="primary", key="btn_transp"):
            results=[]; prog=st.progress(0); status=st.empty()
            for i, row in enumerate(rows):
                status.info(f"Creando: **{row['trade_name']}** ({i+1}/{len(rows)})")
                try:
                    r=requests.post("https://api.simpliroute.com/tms/api/v1/transportation-providers/",
                                    headers={"Authorization":f"Token {token}","Content-Type":"application/json"},
                                    json={"trade_name":row["trade_name"],"legal_name":row["legal_name"],
                                          "tax_id_number":row["tax_id_number"],"account_id":int(account_id)}, timeout=30)
                    code,resp=r.status_code,r.json()
                except Exception as e: code,resp=None,str(e)
                results.append({"trade_name":row["trade_name"],"code":code,"resp":resp}); prog.progress((i+1)/len(rows))
            status.empty(); prog.empty(); show_results(results, "trade_name")


# ── ROUTER ────────────────────────────────────────────────────────────────────
if selected == "🧑‍💼 Agregar Seller a Visitas": page_agregar_seller()
elif selected == "🚛 Asignación de Flotas": page_asignacion_flotas()
elif selected == "🗺️ Cargar Zonas": page_cargar_zonas()
elif selected == "👤 Cambiar Rol de Usuario": page_cambiar_rol()
elif selected == "🔔 Crear Webhook": page_crear_webhook()
elif selected == "🔓 Desbloqueo de Contraseña": page_desbloqueo()
elif selected == "🚦 Iniciar / Cerrar Rutas": page_iniciar_cerrar_rutas()
elif selected == "🔁 Reenviar Webhooks": page_reenviar_webhooks()
elif selected == "🏷️ Tipos de Visita y Skills": page_visit_types_skills()
elif selected == "📡 Validación de GPS": page_validacion_gps()
elif selected == "📄 Tipos de Documento": page_tms_document_types()
elif selected == "🚚 Transportistas": page_tms_transportistas()