import streamlit as st
import openpyxl
import io
import re
import requests

st.set_page_config(page_title="KML → SimpliRoute Zonas", page_icon="🗺️")
st.title("🗺️ KML → SimpliRoute Zonas")
st.markdown("Sube tu archivo KML o RTF, ingresa tu token y carga las zonas directo a SimpliRoute.")

# --- Token input ---
token = st.text_input("🔑 Token de SimpliRoute", type="password", placeholder="Ingresa tu token aquí")

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
    headers = {
        "authorization": f"Token {auth_token}",
        "content-type": "application/json"
    }
    payload = {
        "name": name,
        "coordinates": coordinates,
        "vehicles": []
    }
    try:
        response = requests.post(url, headers=headers, json=payload, timeout=15)
        return response.status_code, response.json()
    except Exception as e:
        return None, str(e)

# --- UI ---
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

        # Preview
        st.subheader("Preview")
        for p in polygons:
            with st.expander(f"📍 {p['name']} — {len(p['coords'])} puntos"):
                st.code(coords_to_str(p["coords"][:3]) + ",...]", language=None)

        # Descargar Excel
        excel_buf = generate_excel(polygons)
        st.download_button(
            label="⬇️ Descargar Excel",
            data=excel_buf,
            file_name="ZONES.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.divider()

        # Cargar a SimpliRoute
        st.subheader("📡 Cargar zonas a SimpliRoute")
        if not token:
            st.warning("⚠️ Ingresa tu token de SimpliRoute arriba para poder cargar las zonas.")
        else:
            if st.button("🚀 Cargar zonas a SimpliRoute", type="primary"):
                results = []
                progress = st.progress(0)
                status_container = st.empty()

                for i, p in enumerate(polygons):
                    status_container.info(f"Cargando: **{p['name']}** ({i+1}/{len(polygons)})")
                    code, resp = upload_zone(p["name"], coords_to_str(p["coords"]), token)
                    results.append({"name": p["name"], "code": code, "resp": resp})
                    progress.progress((i + 1) / len(polygons))

                status_container.empty()
                progress.empty()

                st.subheader("Resultados")
                ok = sum(1 for r in results if r["code"] in [200, 201])
                fail = len(results) - ok
                st.markdown(f"✅ **{ok} zonas cargadas correctamente** | ❌ **{fail} con error**")

                for r in results:
                    if r["code"] in [200, 201]:
                        st.success(f"✅ {r['name']} — Zona creada correctamente")
                    else:
                        st.error(f"❌ {r['name']} — Error {r['code']}: {r['resp']}")
    else:
        st.warning("No se encontraron polígonos en el archivo.")