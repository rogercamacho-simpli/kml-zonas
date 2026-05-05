import streamlit as st
import openpyxl
import io
import re

st.set_page_config(page_title="KML → Excel", page_icon="🗺️")
st.title("🗺️ KML → Excel Zonas")
st.markdown("Sube tu archivo KML o RTF y descarga el Excel listo para importar.")

def decode_file(raw_bytes):
    for enc in ["utf-8", "latin-1", "cp1252"]:
        try:
            return raw_bytes.decode(enc, errors="ignore")
        except Exception:
            continue
    return raw_bytes.decode("utf-8", errors="ignore")

def strip_rtf_codes(text):
    """Limpia códigos RTF básicos del texto."""
    # Eliminar secuencias de escape RTF como \'xx
    text = re.sub(r"\\'[0-9a-fA-F]{2}", "", text)
    # Eliminar comandos RTF como \palabra o \palabra123
    text = re.sub(r"\\[a-zA-Z]+\d*\s?", " ", text)
    # Eliminar llaves RTF
    text = re.sub(r"[{}]", "", text)
    # Limpiar espacios múltiples
    text = re.sub(r" {2,}", " ", text)
    return text

def parse_polygons(text):
    """Extrae polígonos directamente con regex sin parsear XML."""
    polygons = []

    # Encontrar todos los bloques Placemark
    placemark_blocks = re.findall(
        r'<Placemark>(.*?)</Placemark>',
        text, re.DOTALL
    )

    for block in placemark_blocks:
        # Verificar que tenga Polygon (no Point)
        if '<Polygon>' not in block and '<Polygon ' not in block:
            continue

        # Extraer nombre
        name_match = re.search(r'<name>\s*(.*?)\s*</name>', block)
        name = name_match.group(1).strip() if name_match else "Sin nombre"

        # Extraer coordenadas
        coords_match = re.search(r'<coordinates>\s*(.*?)\s*</coordinates>', block, re.DOTALL)
        if not coords_match:
            continue

        raw_coords = coords_match.group(1).strip()
        points = []

        for token in raw_coords.split():
            token = token.strip()
            parts = token.split(",")
            if len(parts) < 2:
                continue
            try:
                a, b = float(parts[0]), float(parts[1])
            except ValueError:
                continue

            # KML orden: lng, lat, alt
            lng, lat = a, b
            # Si lat > 90 están invertidos
            if abs(lat) > 90:
                lng, lat = lat, lng
            # Validación
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

# --- UI ---
uploaded = st.file_uploader("📂 Sube tu archivo KML o RTF", type=["kml", "rtf", "txt"])

if uploaded:
    content = uploaded.read()
    with st.spinner("Procesando archivo..."):
        try:
            text = decode_file(content)
            # Si es RTF, limpiar códigos
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
    else:
        st.warning("No se encontraron polígonos en el archivo.")