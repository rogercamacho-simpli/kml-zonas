import streamlit as st
import json
import time
import requests
from datetime import date

from utils.api import sr_get, sr_post, sr_put, get_users_list, put_user_full, API_GW

APP_VERSION = "2.0.0"

# ── CONSTANTES ────────────────────────────────────────────────────────────────
MASTER_ADDONS = [
    {"key": "visit_signature",            "title": "Firma de Visita",              "description": "",                             "logo": ""},
    {"key": "path_finder",                "title": "path_finder",                  "description": "path_finder",                  "logo": ""},
    {"key": "ada",                        "title": "Ada",                          "description": "ada",                          "logo": "ada"},
    {"key": "live_tracking",              "title": "Live Tracking",                "description": "Live Tracking",                "logo": ""},
    {"key": "rest_time",                  "title": "Tiempo de Descanso",           "description": "rest_time",                    "logo": ""},
    {"key": "mobile_route_edition",       "title": "Edición de Ruta Móvil",        "description": "Addon mobile_route_edition",   "logo": "-"},
    {"key": "mobile_features",            "title": "Funciones Móviles",            "description": "Addon mobile_features",        "logo": "-"},
    {"key": "seller_management",          "title": "Gestión de Sellers",           "description": "seller_management",            "logo": ""},
    {"key": "territory_planner",          "title": "force field",                  "description": "force field",                  "logo": ""},
    {"key": "account_configs_edition",    "title": "Edición de Configuración",     "description": "account_configs_edition",      "logo": ""},
    {"key": "management_comments",        "title": "Comentarios",                  "description": "comentarios",                  "logo": ""},
    {"key": "accounting",                 "title": "accounting",                   "description": "accounting",                   "logo": ""},
    {"key": "simpli_chat",                "title": "Simpli Chat",                  "description": "SImpli Chat",                  "logo": ""},
    {"key": "zones",                      "title": "Zonas",                        "description": "Addon zones",                  "logo": ""},
    {"key": "delivery_survey",            "title": "Delivery Survey",              "description": "delivery_survey",              "logo": ""},
    {"key": "tms",                        "title": "TMS",                          "description": "tms",                          "logo": ""},
    {"key": "visit_card_additional_data", "title": "Visit Card Additional Data",   "description": "visit_card_additional_data",   "logo": "-"},
    {"key": "roles",                      "title": "Roles Personalizados",         "description": "roles",                        "logo": ""},
]

PERMISOS_AMIGABLES = {
    "reports.edit":                               "📊 Reportes",
    "watchtower.tracking.edit":                   "🗺️ Watch Tower - Tracking",
    "watchtower.visits.edit":                     "📍 Watch Tower - Visitas",
    "settings.communications.live_tracking.edit": "📡 Comunicaciones",
    "settings.communications.wsp_pro.edit":       "💬 WhatsApp Pro",
    "clients.edit":                               "👥 Clientes",
    "chat.edit":                                  "💬 Chat",
    "ada.edit":                                   "🤖 Ada",
    "forcefield.edit":                            "🛡️ ForceField",
    "settings.vehicles.edit":                     "🚗 Vehículos",
    "pathfinder.plans.edit":                      "📋 Mis planes",
    "pathfinder.router.edit":                     "🔀 Ruteador",
    "pathfinder.assigner.edit":                   "📌 Asignador",
    "settings.users.edit":                        "👤 Usuarios",
    "settings.vehicles.fleets.edit":              "🚛 Flotas",
    "settings.zones.edit":                        "🗺️ Zonas",
    "settings.optimizers.edit":                   "⚙️ Creación de rutas",
    "settings.mobile.edit":                       "📱 App móvil",
    "settings.security.edit":                     "🔒 Seguridad",
    "settings.customization.edit":                "🎨 Personalización",
}

SYSTEM_ROLE_FLAGS = {
    "is_owner":         "Owner",
    "is_admin":         "Administrador",
    "is_router":        "Router",
    "is_router_jr":     "Router Jr",
    "is_coordinator":   "Coordinador",
    "is_monitor":       "Monitor",
    "is_seller":        "Seller",
    "is_seller_viewer": "Seller Viewer",
    "is_driver":        "Conductor",
    "is_codriver":      "Co-conductor",
}

ROLES = {
    "Administrador": "is_admin", "Conductor": "is_driver", "Co-conductor": "is_codriver",
    "Router Jr": "is_router_jr", "Monitor": "is_monitor", "Coordinador": "is_coordinator",
    "Router": "is_router", "Staff": "is_staff", "Seller Viewer": "is_seller_viewer", "Seller": "is_seller",
}
ALL_ROLE_KEYS = list(ROLES.values()) + ["is_owner"]

SUPPORT_AGENTS = {
    "Brandon Vargas":  "brandon.vargas@simpliroute.com",
    "Carlos Junior":   "carlos.celestino@simpliroute.com",
    "David Martinez":  "david.martinez@simpliroute.com",
    "Itzel Meza":      "itzel.meza@simpliroute.com",
    "Jorge Cruz":      "jorge.cruz@simpliroute.com",
    "Julio Mares":     "julio.mares@simpliroute.com",
    "Roger Camacho":   "roger.camacho@simpliroute.com",
    "Silmary Guedez":  "silmary.guedez@simpliroute.com",
}

WEBHOOK_EVENTS = {
    "Plan creado": "plan_created", "Plan editado": "plan_edited",
    "Ruta creada": "route_created", "Ruta editada": "route_edited",
    "Ruta iniciada": "route_started", "Ruta finalizada": "route_finished",
    "En camino": "on_its_way", "Checkout": "visit_checkout",
    "Checkout detallado": "visit_checkout_detailed",
}


# ── PÁGINAS ───────────────────────────────────────────────────────────────────
def page_cambiar_rol():
    st.title("👤 Cambiar Rol de Usuario")
    col1, col2 = st.columns(2)
    with col1: token   = st.text_input("🔑 Token", type="password", key="token_rol")
    with col2: user_id = st.text_input("🆔 ID del Usuario", placeholder="Ej: 524614")

    if st.button("🔍 Consultar", type="primary", disabled=not (token and user_id)):
        code, resp = sr_get(f"/v1/accounts/users/{user_id.strip()}/", token)
        if code == 200:
            st.session_state["user_data"]  = resp
            st.session_state["user_token"] = token
        else:
            st.error(f"❌ Error {code}")
            st.session_state.pop("user_data", None)

    if "user_data" in st.session_state:
        user = st.session_state["user_data"]
        def get_current_role(u):
            for label, key in ROLES.items():
                if u.get(key): return label, key
            if u.get("is_owner"): return "Owner", "is_owner"
            return "Sin rol", None
        cl, ck = get_current_role(user)
        st.divider()
        c1, c2, c3 = st.columns(3)
        c1.markdown(f"**Nombre**<br>{user.get('name','—')}",      unsafe_allow_html=True)
        c2.markdown(f"**Username**<br>{user.get('username','—')}", unsafe_allow_html=True)
        c3.markdown(f"**Rol actual**<br>{cl}",                     unsafe_allow_html=True)
        st.caption(f"📧 {user.get('email','—')} · {user.get('status','—')}")
        st.divider()
        if ck == "is_owner":
            st.warning("⚠️ Este usuario es Owner. No se puede modificar.")
        else:
            nuevo_label = st.selectbox("Nuevo rol:", [l for l, k in ROLES.items() if k != ck])
            nuevo_key   = ROLES[nuevo_label]
            if st.button(f"💾 Cambiar a {nuevo_label}", type="primary"):
                payload = {k: (k == nuevo_key) for k in ALL_ROLE_KEYS}
                payload["username"] = user.get("username", "")
                payload["name"]     = user.get("name", "")
                code, _ = sr_put(f"/v1/accounts/users/{user['id']}/", st.session_state["user_token"], payload)
                if code == 200:
                    st.success(f"✅ Rol actualizado a **{nuevo_label}**")
                    code2, resp2 = sr_get(f"/v1/accounts/users/{user['id']}/", st.session_state["user_token"])
                    if code2 == 200: st.session_state["user_data"] = resp2
                    st.rerun()
                else:
                    st.error(f"❌ Error {code}")


def page_configurar_addons():
    st.title("⚙️ Configurar Addons")
    token = st.text_input("🔑 Token de SimpliRoute", type="password", key="token_addons")

    if st.button("🔍 Consultar Addons", type="primary", disabled=not token):
        with st.spinner("Consultando addons..."):
            code, resp = sr_get("/v1/addons/addons/", token)
        if code == 200:
            acc_id = resp[0]["account_id"] if resp else "—"
            st.session_state["addons_data"]    = resp
            st.session_state["addons_token"]   = token
            st.session_state["addons_account"] = acc_id
        elif code == 401: st.error("❌ Token inválido")
        else:
            st.error(f"❌ Error {code}")
            st.session_state.pop("addons_data", None)

    if "addons_data" in st.session_state:
        addons        = st.session_state["addons_data"]
        tok           = st.session_state["addons_token"]
        acc_id        = st.session_state["addons_account"]
        existing_keys = {a["key"] for a in addons}

        st.info(f"🏢 Cuenta: **{acc_id}** · {len(addons)} addon(s) configurados")
        st.divider()
        st.subheader("📋 Addons de la cuenta")

        cols = st.columns(3)
        for idx, addon in enumerate(addons):
            master_entry = next((m for m in MASTER_ADDONS if m["key"] == addon["key"]), None)
            label  = master_entry["title"] if master_entry else (addon.get("title") or addon.get("key"))
            estado = addon.get("enable", False)
            color_borde  = "#1D9E75" if estado else "#E24B4A"
            status_color = "#0F6E56" if estado else "#A32D2D"
            badge_bg     = "#E1F5EE" if estado else "#FCEBEB"
            estado_txt   = "✅ Activo"    if estado else "🔴 Inactivo"
            btn_label    = "🔴 Desactivar" if estado else "🟢 Activar"
            with cols[idx % 3]:
                st.markdown(
                    f"""<div style="border:1.5px solid {color_borde};border-radius:10px;padding:14px 16px 8px 16px;margin-bottom:12px;background:#ffffff;">
                        <div style="font-size:14px;font-weight:600;color:#1a1a1a;margin-bottom:4px;">{label}</div>
                        <div style="font-size:11px;font-family:monospace;background:{badge_bg};color:{status_color};display:inline-block;padding:1px 7px;border-radius:4px;margin-bottom:8px;">{addon['key']}</div>
                        <div style="font-size:13px;color:{status_color};font-weight:500;">{estado_txt}</div>
                    </div>""", unsafe_allow_html=True)
                if st.button(btn_label, key=f"toggle_{addon['id']}", use_container_width=True):
                    new_val = not estado
                    payload = [{"id": addon["id"], "account_id": addon["account_id"],
                                "title": addon.get("title",""), "key": addon["key"],
                                "description": addon.get("description",""), "logo": addon.get("logo",""),
                                "value": new_val}]
                    code, _ = sr_put("/v1/addons/addons/", tok, payload)
                    if code in [200, 201]:
                        for a in st.session_state["addons_data"]:
                            if a["id"] == addon["id"]: a["enable"] = new_val
                        st.rerun()
                    elif code == 401: st.error("❌ Token inválido")
                    else:             st.error(f"❌ Error {code}")

        missing = [m for m in MASTER_ADDONS if m["key"] not in existing_keys]
        if missing:
            st.divider()
            st.subheader("➕ Addons no configurados")
            st.caption("Estos addons no están en la cuenta. Puedes agregarlos aquí.")
            staff_tok_input = st.text_input("🔐 Token Staff *(requerido para agregar addons)*",
                                            type="password", key="staff_token_addons",
                                            placeholder="Token de usuario Staff")
            cols2 = st.columns(3)
            for idx2, master in enumerate(missing):
                with cols2[idx2 % 3]:
                    st.markdown(
                        f"""<div style="border:1.5px dashed #ccc;border-radius:10px;padding:14px 16px 8px 16px;margin-bottom:12px;background:#fafafa;">
                            <div style="font-size:14px;font-weight:600;color:#555;margin-bottom:4px;">{master['title']}</div>
                            <div style="font-size:11px;font-family:monospace;background:#f0f0f0;color:#888;display:inline-block;padding:1px 7px;border-radius:4px;margin-bottom:8px;">{master['key']}</div>
                            <div style="font-size:12px;color:#aaa;">No configurado</div>
                        </div>""", unsafe_allow_html=True)
                    if st.button("➕ Agregar", key=f"add_{master['key']}", use_container_width=True):
                        use_tok = staff_tok_input.strip() if staff_tok_input and staff_tok_input.strip() else tok
                        payload = {"account_id": int(acc_id), "title": master["title"], "key": master["key"],
                                   "description": master["description"], "logo": master["logo"], "value": True}
                        code, post_resp = sr_post("/v1/addons/addons/", use_tok, payload)
                        if code in [200, 201]:
                            st.success(f"✅ **{master['title']}** agregado")
                            code2, data2 = sr_get("/v1/addons/addons/", tok)
                            if code2 == 200: st.session_state["addons_data"] = data2
                            st.rerun()
                        elif code == 403: st.error("❌ Error 403 — Solo un usuario Staff puede agregar addons.")
                        elif code == 401: st.error("❌ Token inválido")
                        else:             st.error(f"❌ Error {code}: {post_resp}")
        else:
            st.divider()
            st.info("✅ La cuenta tiene todos los addons del catálogo configurados.")


def page_copiar_configs():
    st.title("⚙️ Copiar Configs")
    st.markdown("Copia las configuraciones de una cuenta origen a una cuenta destino.")
    col1, col2 = st.columns(2)
    with col1: token_origen  = st.text_input("🔑 Token cuenta origen",  type="password", key="token_configs_origen")
    with col2: token_destino = st.text_input("🔑 Token cuenta destino", type="password", key="token_configs_destino")

    if st.button("🔍 Consultar configs", type="primary", disabled=not token_origen, key="btn_consultar_configs"):
        code, data = sr_get("/v1/accounts/configs/", token_origen)
        if code == 401:   st.error("❌ Token inválido")
        elif code != 200: st.error(f"❌ Error {code}")
        else:             st.session_state["configs_origen"] = data

    if "configs_origen" in st.session_state:
        configs = st.session_state["configs_origen"]
        KEYS_EXCLUIDAS = {"routing_sources_provider"}
        configs_activas = [c for c in configs if c["value"] not in [False, None, "", "false"]
                           and c["key"] not in KEYS_EXCLUIDAS]
        st.success(f"✅ {len(configs_activas)} configuración(es) activas")
        st.divider()
        col_sel, col_desel = st.columns(2)
        with col_sel:
            if st.button("☑️ Seleccionar todas", key="sel_todas_configs", use_container_width=True):
                for c in configs_activas: st.session_state[f"cfg_check_{c['key']}"] = True
        with col_desel:
            if st.button("🔲 Deseleccionar todas", key="desel_todas_configs", use_container_width=True):
                for c in configs_activas: st.session_state[f"cfg_check_{c['key']}"] = False

        st.markdown("**Selecciona las configs a copiar:**")
        seleccionadas = []
        cols = st.columns(2)
        for i, c in enumerate(configs_activas):
            with cols[i % 2]:
                if st.checkbox(f"`{c['key']}` → **{c['value']}**", key=f"cfg_check_{c['key']}"):
                    seleccionadas.append(c)

        st.divider()
        if seleccionadas: st.info(f"**{len(seleccionadas)} config(s) seleccionadas**")

        if st.button("🚀 Copiar configs", type="primary",
                     disabled=not (token_destino and seleccionadas), key="btn_copiar_configs"):
            code_d, data_d = sr_get("/v1/accounts/configs/", token_destino)
            if code_d == 401: st.error("❌ Token destino inválido"); return
            if code_d != 200: st.error(f"❌ Error consultando destino: {code_d}"); return
            keys_destino = {c["key"] for c in data_d}

            prog = st.progress(0); status = st.empty()
            ok_count = skip_count = err_count = 0
            start_time = time.time()
            for i, config in enumerate(seleccionadas):
                elapsed = time.time() - start_time
                rate    = (i + 1) / elapsed if elapsed > 0 else 1
                eta     = (len(seleccionadas) - (i + 1)) / rate
                status.info(f"Procesando: **{config['key']}** ({i+1}/{len(seleccionadas)}) — ETA: {int(eta)}s")
                if config["key"] in keys_destino:
                    st.warning(f"⏭️ **{config['key']}** — Ya existe, omitida")
                    skip_count += 1; prog.progress((i+1)/len(seleccionadas)); continue
                payload = {"key": config["key"], "value": config["value"],
                           "string_value": "false", "is_public": config["is_public"]}
                code, _ = sr_post("/v1/accounts/configs/", token_destino, payload)
                if code in [200, 201]:  ok_count += 1
                elif code == 401: st.error("❌ Token destino inválido."); status.empty(); prog.empty(); return
                else: st.error(f"❌ **{config['key']}** — Error {code}"); err_count += 1
                prog.progress((i+1)/len(seleccionadas))
            status.empty(); prog.empty()
            st.divider()
            st.success(f"✅ **{ok_count} copiadas** · ⏭️ **{skip_count} omitidas** · ❌ **{err_count} con error**")


def page_crear_webhook():
    st.title("🔔 Crear Webhook")
    token      = st.text_input("🔑 Token de SimpliRoute", type="password", key="token_webhook")
    url        = st.text_input("🌐 URL de destino", placeholder="https://tu-servidor.com/webhook")
    evento_key = WEBHOOK_EVENTS[st.selectbox("📋 Tipo de evento", list(WEBHOOK_EVENTS.keys()))]
    custom     = st.checkbox("Usar header personalizado")
    if custom:
        raw = st.text_area("Header JSON",
                           value='{\n    "Content-Type": "application/json",\n    "X-Custom": "valor"\n}',
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
        code, resp = sr_post("/v1/addons/webhooks/", token,
                             {"webhook": evento_key, "url": url, "headers": headers_payload})
        if code in [200, 201]: st.success("✅ Webhook creado"); st.json(resp)
        elif code == 401:      st.error("❌ Token inválido")
        else:                  st.error(f"❌ Error {code}: {resp}")


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
                st.session_state["unlock_token"]  = token
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
            if code in [200, 201]:
                st.success("✅ Email restaurado.")
                st.session_state.pop("pending_restore", None)
            else:
                st.error(f"❌ Error {code}")

    if "blocked_users" in st.session_state:
        blocked = st.session_state["blocked_users"]
        st.divider()
        st.success(f"⚠️ {len(blocked)} usuario(s) bloqueado(s)")
        def ulabel(u):
            name = u.get("name") or u.get("username") or str(u.get("id"))
            un   = u.get("username", "")
            return f"{name} — {un}" if un != name else name
        selected_user  = {ulabel(u): u for u in blocked}[st.selectbox("Usuario bloqueado:", [ulabel(u) for u in blocked])]
        original_email = selected_user.get("email", "")
        st.caption(f"📧 Email actual: **{original_email or 'Sin email'}**")
        st.divider()
        agent_name     = st.selectbox("Agente que recibirá el link:", list(SUPPORT_AGENTS.keys()))
        recovery_email = SUPPORT_AGENTS[agent_name]
        st.caption(f"📨 Se enviará a: **{recovery_email}**")
        st.divider()
        if st.button("🔓 Enviar link de desbloqueo", type="primary"):
            tok  = st.session_state["unlock_token"]
            prog = st.progress(0); status = st.empty()
            status.info("1/2 — Asignando email temporal...")
            code1, _ = put_user_full(tok, selected_user, recovery_email)
            prog.progress(50)
            if code1 not in [200, 201]:
                st.error(f"❌ Error al actualizar email: {code1}"); status.empty(); prog.empty()
            else:
                status.info("2/2 — Enviando link...")
                try:
                    r = requests.post("https://api.simpliroute.com/v2/auth/unlock/",
                                      headers={"Content-Type": "application/json;charset=UTF-8", "authorization": "null"},
                                      json={"username": selected_user.get("username")}, timeout=30)
                    code2 = r.status_code
                except: code2 = None
                prog.progress(100); status.empty(); prog.empty()
                if code2 in [200, 201]:
                    st.success(f"✅ Link enviado a **{recovery_email}** ({agent_name})")
                    st.warning("⚠️ Cuando el usuario resetee su contraseña, usa el botón de restauración de arriba.")
                    st.session_state["pending_restore"] = {"user": selected_user, "original_email": original_email, "token": tok}
                    st.session_state.pop("blocked_users", None)
                    st.rerun()
                else:
                    put_user_full(tok, selected_user, original_email)
                    st.error(f"❌ No se pudo enviar el link. Email restaurado. ({code2})")


def page_permisos_usuario():
    st.title("🔍 Permisos de Usuario")
    st.markdown("Consulta el rol y permisos de un usuario, incluyendo roles personalizados.")
    col1, col2 = st.columns(2)
    with col1: token   = st.text_input("🔑 Token", type="password", key="token_permisos")
    with col2: user_id = st.text_input("🆔 User ID", placeholder="Ej: 503734", key="user_permisos")

    if not st.button("🔍 Consultar", type="primary", disabled=not (token and user_id)):
        return

    code_me, me_data = sr_get("/v1/accounts/me", token)
    if code_me == 401: st.error("❌ Token inválido"); return
    if code_me != 200: st.error(f"❌ Error obteniendo cuenta: {code_me}"); return
    account_id = me_data.get("account", {}).get("id")
    if not account_id: st.error("❌ No se pudo obtener el Account ID"); return

    code_u, user = sr_get(f"/v1/accounts/users/{user_id.strip()}/", token)
    if code_u == 401: st.error("❌ Token inválido"); return
    if code_u != 200: st.error(f"❌ Error {code_u}"); return

    st.divider()
    c1, c2, c3 = st.columns(3)
    c1.markdown(f"**Nombre**<br>{user.get('name','—')}",      unsafe_allow_html=True)
    c2.markdown(f"**Username**<br>{user.get('username','—')}", unsafe_allow_html=True)
    estado = "🔴 Bloqueado" if user.get("blocked") else "🟢 Activo"
    c3.markdown(f"**Estado**<br>{estado}", unsafe_allow_html=True)
    st.caption(f"📧 {user.get('email') or 'Sin email'} · ID: {user.get('id')} · Cuenta: {account_id}")

    roles = []
    code_r, roles_data = sr_get(f"/v1/accounts/{account_id}/roles/", token, base=API_GW)
    if code_r == 200: roles = roles_data

    custom_role_name = user.get("custom_role_name")
    st.divider()
    if custom_role_name:
        role_data = next((r for r in roles if r.get("name") == custom_role_name), None)
        st.markdown(f"**🏷️ Rol:** `{custom_role_name}` — *personalizado*")
    else:
        flag_key   = next((f for f in SYSTEM_ROLE_FLAGS if user.get(f)), None)
        role_label = SYSTEM_ROLE_FLAGS.get(flag_key, "Sin rol")
        role_data  = next((r for r in roles if r.get("legacy_role") == flag_key and r.get("is_system_role")), None)
        st.markdown(f"**🏷️ Rol:** `{role_label}` — *sistema*")

    if role_data:
        perm_keys = [p["key"] for p in role_data.get("permissions", [])]
        tiene     = [PERMISOS_AMIGABLES[k] for k in PERMISOS_AMIGABLES if k in perm_keys]
        no_tiene  = [PERMISOS_AMIGABLES[k] for k in PERMISOS_AMIGABLES if k not in perm_keys]
        st.markdown(f"**✅ Tiene acceso a ({len(tiene)}):**")
        cols = st.columns(2)
        for i, label in enumerate(tiene): cols[i % 2].markdown(f"• {label}")
        if no_tiene:
            with st.expander(f"🚫 Sin acceso ({len(no_tiene)})"):
                cols2 = st.columns(2)
                for i, label in enumerate(no_tiene): cols2[i % 2].markdown(f"• {label}")
    else:
        st.warning("⚠️ No se encontraron los permisos detallados para este rol.")


def page_reenviar_webhooks():
    st.title("🔁 Reenviar Webhooks")
    st.info("🔑 Usa el token de tu cuenta de SimpliRoute.")
    token      = st.text_input("🔑 Token", type="password", key="token_resend")
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
            code, _ = sr_post("/v1/mobile/send-webhooks", token,
                              {"account_ids": [int(account_id)], "planned_date": today, "visit_ids": ids},
                              base="https://api.simpliroute.com", timeout=60)
        if code in [200, 201]: st.success(f"✅ Webhooks reenviados para {len(ids)} visitas")
        elif code == 401:      st.error("❌ Token inválido")
        elif code == 404:      st.error("❌ Cuenta no encontrada")
        else:                  st.error(f"❌ Error {code}")
