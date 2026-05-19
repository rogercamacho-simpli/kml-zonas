import streamlit as st
import openpyxl
import io
from datetime import date

from utils.helpers import read_excel_column, coords_to_str, generate_excel_zones, decode_file, strip_rtf_codes, parse_polygons
from utils.api import sr_get, sr_post, sr_put, sr_delete, get_users_list


def page_flotas():
    st.title("🚛 Flotas")
    tab1, tab2 = st.tabs(["📋 Asignar Flotas", "🗑️ Eliminar Flotas"])

    with tab1:
        st.info("ℹ️ Cada fila representa una edición independiente. Si la misma flota aparece dos veces, prevalece la última.")
        token = st.text_input("🔑 Token de SimpliRoute", type="password", key="token_flotas")

        def make_tpl():
            wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Flotas"
            ws.append(["Nombre de flota", "Vehículos", "Usuarios"])
            ws.append(["Flota Norte", "MC4327,QC4380", "juan.perez,ana.lopez"])
            ws.column_dimensions["A"].width = 25
            ws.column_dimensions["B"].width = 40
            ws.column_dimensions["C"].width = 40
            buf = io.BytesIO(); wb.save(buf); buf.seek(0); return buf

        st.download_button("📥 Descargar plantilla", data=make_tpl(), file_name="plantilla_flotas.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        fleet_file = st.file_uploader("📂 Sube tu Excel", type=["xlsx"], key="upload_flotas")

        if fleet_file and token:
            if st.button("🚀 Procesar flotas", type="primary"):
                try:
                    rows = read_excel_column(fleet_file, ["Nombre de flota", "Vehículos", "Usuarios"])
                except Exception as e:
                    st.error(f"❌ Error leyendo Excel: {e}"); return
                with st.spinner("⏳ Consultando flotas, vehículos y usuarios..."):
                    code_f, fleets   = sr_get("/v1/fleets/", token, timeout=300)
                    code_v, vehicles = sr_get("/v1/routes/vehicles/", token, timeout=300)
                    code_u, users    = get_users_list(token)

                fleet_map   = {f["name"].strip().lower(): f for f in (fleets if isinstance(fleets, list) else [])}
                vehicle_map = {v["name"].strip().lower(): v["id"] for v in (vehicles if isinstance(vehicles, list) else []) if v.get("name")}
                user_map = {}
                for u in (users if isinstance(users, list) else []):
                    if u.get("username"): user_map[u["username"].strip().lower()] = u["id"]
                    if u.get("email"):    user_map[u["email"].strip().lower()]    = u["id"]

                st.success(f"✅ {len(fleet_map)} flotas · {len(vehicle_map)} vehículos · {len(user_map)} usuarios")
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
                    code, _ = sr_put(f"/v1/fleets/{fleet['id']}/", token,
                                     {"id": fleet["id"], "name": fname, "vehicles": vids, "users": uids}, timeout=300)
                    if code == 200:
                        st.success(f"✅ **{fname}** — Actualizada ({len(vids)} vehículos · {len(uids)} usuarios)")
                    else:
                        st.error(f"❌ **{fname}** — Error {code}")
        elif fleet_file and not token:
            st.warning("⚠️ Ingresa tu token.")

    with tab2:
        st.error("⚠️ **ADVERTENCIA:** La eliminación de flotas es permanente y no se puede deshacer.")
        token_del = st.text_input("🔑 Token de SimpliRoute", type="password", key="token_flotas_del")
        st.caption("Pega los IDs de flota uno por línea:\n```\n47524\n47525\n```")
        fleet_ids_raw = st.text_area("IDs de flota", placeholder="47524\n47525", height=180,
                                     label_visibility="collapsed", key="fleet_ids_del")
        confirm_del = st.checkbox("✅ Confirmo que quiero eliminar estas flotas de forma permanente", key="confirm_flotas_del")
        if st.button("🗑️ Eliminar Flotas", type="primary", disabled=not (token_del and fleet_ids_raw and confirm_del)):
            fleet_ids = []
            for l in fleet_ids_raw.strip().splitlines():
                l = l.strip()
                if not l: continue
                try: fleet_ids.append(int(l))
                except: st.warning(f"⚠️ ID inválido ignorado: {l}")
            if not fleet_ids: st.error("❌ No se encontraron IDs válidos."); return
            prog = st.progress(0); status = st.empty()
            ok_count = 0; err_count = 0
            for i, fid in enumerate(fleet_ids):
                status.info(f"Eliminando flota **{fid}** ({i+1}/{len(fleet_ids)})...")
                code, _ = sr_delete(f"/v1/fleets/{fid}", token_del, timeout=30)
                if code in [200, 204]: st.success(f"✅ Flota **{fid}** eliminada"); ok_count += 1
                elif code == 401: st.error("❌ Token inválido."); status.empty(); prog.empty(); return
                elif code == 404: st.warning(f"⚠️ **{fid}** — No encontrada")
                else: st.error(f"❌ **{fid}** — Error {code}"); err_count += 1
                prog.progress((i+1)/len(fleet_ids))
            status.empty(); prog.empty()
            st.divider()
            if err_count == 0: st.success(f"✅ Completado — **{ok_count} flotas eliminadas**")
            else: st.warning(f"⚠️ Completado con errores — **{ok_count} eliminadas**, {err_count} con error")


def page_iniciar_cerrar_rutas():
    st.title("🚦 Iniciar / Cerrar Rutas")
    token = st.text_input("🔑 Token de SimpliRoute", type="password", key="token_routes")
    evento = st.selectbox("📋 Tipo de evento", ["Iniciar ruta", "Finalizar ruta"])
    event_type = "ROUTE_STARTED" if evento == "Iniciar ruta" else "ROUTE_FINISHED"
    selected_date = st.date_input("📅 Fecha", value=date.today(), key="route_date")
    hora = "12:00:00.000Z" if event_type == "ROUTE_STARTED" else "22:00:00.000Z"
    date_time = f"{selected_date.strftime('%Y-%m-%d')}T{hora}"
    st.caption(f"📅 Fecha y hora que se usará: **{date_time}**")
    st.divider()
    st.caption("Pega los IDs de ruta uno por línea:\n```\n637f11a2-a1a6-4609-8c23-83e8c76dccbf\n```")
    route_ids_raw = st.text_area("IDs de ruta", placeholder="637f11a2-a1a6-4609-8c23-83e8c76dccbf",
                                 height=200, label_visibility="collapsed")

    if st.button(f"🚀 {evento}", type="primary", disabled=not (token and route_ids_raw)):
        route_ids = [l.strip() for l in route_ids_raw.strip().splitlines() if l.strip()]
        if not route_ids: st.error("❌ No se encontraron IDs."); return
        prog = st.progress(0); status = st.empty()
        for i, route_id in enumerate(route_ids):
            status.info(f"Procesando: **{route_id}** ({i+1}/{len(route_ids)})")
            code_get, route_data = sr_get(f"/v1/routes/routes/{route_id}/", token)
            if code_get != 200:
                st.error(f"❌ **{route_id}** — Error {code_get}"); prog.progress((i+1)/len(route_ids)); continue
            lat = route_data.get("location_start_latitude") if event_type == "ROUTE_STARTED" else route_data.get("location_end_latitude")
            lng = route_data.get("location_start_longitude") if event_type == "ROUTE_STARTED" else route_data.get("location_end_longitude")
            if not lat or not lng:
                st.warning(f"⚠️ **{route_id}** — Sin coordenadas"); prog.progress((i+1)/len(route_ids)); continue
            code_post, _ = sr_post("/v1/events/register/", token,
                                   {"date_time": date_time, "latitude": float(lat), "longitude": float(lng),
                                    "route_id": route_id, "type": event_type})
            if code_post in [200, 201]: st.success(f"✅ **{route_id}** — {evento} registrado")
            elif code_post == 401: st.error(f"❌ **{route_id}** — Token inválido")
            else: st.error(f"❌ **{route_id}** — Error {code_post}")
            prog.progress((i+1)/len(route_ids))
        status.empty(); prog.empty()


def page_zonas():
    st.title("🗺️ Zonas")
    tab1, tab2, tab3 = st.tabs(["📂 Cargar Zonas", "🗑️ Eliminar Zonas", "📋 Copiar Zonas"])

    with tab1:
        token = st.text_input("🔑 Token de SimpliRoute", type="password", key="token_zonas")
        uploaded = st.file_uploader("📂 Sube tu archivo KML o RTF", type=["kml", "rtf", "txt"])
        if uploaded:
            text = decode_file(uploaded.read())
            if text.strip().startswith("{\\rtf") or "\\rtf" in text[:100]:
                text = strip_rtf_codes(text)
            polygons = parse_polygons(text)
            if not polygons: st.warning("No se encontraron polígonos."); return
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
                    code, _ = sr_post("/v1/zones/", token,
                                      {"name": p["name"], "coordinates": coords_to_str(p["coords"]), "vehicles": []})
                    if code in [200, 201]: st.success(f"✅ {p['name']}")
                    elif code == 400: st.warning(f"⚠️ {p['name']} — Ya existe")
                    elif code == 401: st.error(f"❌ {p['name']} — Token inválido")
                    else: st.error(f"❌ {p['name']} — Error {code}")
                    prog.progress((i+1)/len(polygons))
                status.empty(); prog.empty()

    with tab3:
        st.markdown("Consulta las zonas de una cuenta origen y cópialas a una cuenta destino.")
        col1, col2 = st.columns(2)
        with col1: token_origen  = st.text_input("🔑 Token cuenta origen",  type="password", key="token_zonas_origen")
        with col2: token_destino = st.text_input("🔑 Token cuenta destino", type="password", key="token_zonas_destino")

        if st.button("🔍 Consultar zonas", type="primary", disabled=not token_origen, key="btn_consultar_zonas"):
            code, data = sr_get("/v1/zones/", token_origen)
            if code == 401: st.error("❌ Token inválido")
            elif code != 200: st.error(f"❌ Error {code}")
            else: st.session_state["zonas_origen"] = data

        if "zonas_origen" in st.session_state:
            zonas = st.session_state["zonas_origen"]
            st.success(f"✅ {len(zonas)} zona(s) encontradas")
            st.divider()
            st.markdown("**Selecciona las zonas a copiar:**")
            col_sel, col_desel = st.columns(2)
            with col_sel:
                if st.button("☑️ Seleccionar todas", key="sel_todas", use_container_width=True):
                    for zona in zonas: st.session_state[f"zona_check_{zona['id']}"] = True
            with col_desel:
                if st.button("🔲 Deseleccionar todas", key="desel_todas", use_container_width=True):
                    for zona in zonas: st.session_state[f"zona_check_{zona['id']}"] = False

            seleccionadas = []
            cols = st.columns(2)
            for i, zona in enumerate(zonas):
                with cols[i % 2]:
                    if st.checkbox(f"📍 {zona['name']}", key=f"zona_check_{zona['id']}"):
                        seleccionadas.append(zona)

            st.divider()
            if seleccionadas: st.info(f"**{len(seleccionadas)} zona(s) seleccionadas** para copiar")

            if st.button("🚀 Copiar zonas", type="primary",
                         disabled=not (token_destino and seleccionadas), key="btn_copiar_zonas"):
                prog = st.progress(0); status = st.empty()
                ok_count = 0; err_count = 0
                for i, zona in enumerate(seleccionadas):
                    status.info(f"Copiando: **{zona['name']}** ({i+1}/{len(seleccionadas)})")
                    coords = zona.get("coordinates", [])
                    coords_str = "[" + ",".join(f"{{'lat':'{p['lat']}','lng':'{p['lng']}'}}" for p in coords) + "]"
                    code, _ = sr_post("/v1/zones/", token_destino,
                                      {"name": zona["name"], "coordinates": coords_str, "vehicles": []})
                    if code in [200, 201]: st.success(f"✅ {zona['name']}"); ok_count += 1
                    elif code == 400: st.warning(f"⚠️ {zona['name']} — Ya existe")
                    elif code == 401: st.error("❌ Token destino inválido."); status.empty(); prog.empty(); return
                    else: st.error(f"❌ {zona['name']} — Error {code}"); err_count += 1
                    prog.progress((i+1)/len(seleccionadas))
                status.empty(); prog.empty()
                st.divider()
                if err_count == 0: st.success(f"✅ Completado — **{ok_count} zonas copiadas**")
                else: st.warning(f"⚠️ Completado — **{ok_count} copiadas**, {err_count} con error")

    with tab2:
        st.error("⚠️ **ADVERTENCIA:** La eliminación de zonas es permanente y no se puede deshacer.")
        token_del = st.text_input("🔑 Token de SimpliRoute", type="password", key="token_zonas_del")
        st.caption("Pega los IDs de zona uno por línea:\n```\n12345\n12346\n```")
        zone_ids_raw = st.text_area("IDs de zona", placeholder="12345\n12346", height=180,
                                    label_visibility="collapsed", key="zone_ids_del")
        confirm_del = st.checkbox("✅ Confirmo que quiero eliminar estas zonas de forma permanente", key="confirm_zonas_del")
        if st.button("🗑️ Eliminar Zonas", type="primary", disabled=not (token_del and zone_ids_raw and confirm_del)):
            zone_ids = []
            for l in zone_ids_raw.strip().splitlines():
                l = l.strip()
                if not l: continue
                try: zone_ids.append(int(l))
                except: st.warning(f"⚠️ ID inválido ignorado: {l}")
            if not zone_ids: st.error("❌ No se encontraron IDs válidos."); return
            prog = st.progress(0); status = st.empty()
            ok_count = 0; err_count = 0
            for i, zid in enumerate(zone_ids):
                status.info(f"Eliminando zona **{zid}** ({i+1}/{len(zone_ids)})...")
                code, _ = sr_delete(f"/v1/zones/{zid}", token_del)
                if code in [200, 204]: st.success(f"✅ Zona **{zid}** eliminada"); ok_count += 1
                elif code == 401: st.error("❌ Token inválido."); status.empty(); prog.empty(); return
                elif code == 404: st.warning(f"⚠️ **{zid}** — No encontrada")
                else: st.error(f"❌ **{zid}** — Error {code}"); err_count += 1
                prog.progress((i+1)/len(zone_ids))
            status.empty(); prog.empty()
            st.divider()
            if err_count == 0: st.success(f"✅ Completado — **{ok_count} zonas eliminadas**")
            else: st.warning(f"⚠️ Completado con errores — **{ok_count} eliminadas**, {err_count} con error")
