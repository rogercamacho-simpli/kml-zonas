import streamlit as st
import openpyxl
import io

from utils.helpers import read_excel_column, show_results
from utils.api import sr_post


def page_tms_document_types():
    st.title("📄 Tipos de Documento TMS")
    COUNTRY_OPTIONS = ["PE", "CL", "MX", "CO", "AR", "EC", "BO", "UY", "PY", "VE"]
    c1, c2, c3 = st.columns(3)
    with c1: token      = st.text_input("🔑 Token",      type="password", key="token_doctype")
    with c2: account_id = st.text_input("🏢 Account ID", placeholder="Ej: 9695")
    with c3: country    = st.selectbox("🌎 País", COUNTRY_OPTIONS)

    def tpl():
        wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Tipos de Documento"
        ws.append(["name", "entity_type"])
        ws.append(["DNI",  "driver"])
        ws.append(["RUC",  "provider"])
        ws.column_dimensions["A"].width = 25
        ws.column_dimensions["B"].width = 25
        buf = io.BytesIO(); wb.save(buf); buf.seek(0); return buf

    st.download_button("📥 Plantilla", data=tpl(), file_name="plantilla_tipos_documento.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    doc_file = st.file_uploader("📂 Excel", type=["xlsx"], key="upload_doctype")

    if doc_file:
        try:
            rows = read_excel_column(doc_file, ["name", "entity_type"])
            st.success(f"✅ {len(rows)} tipo(s)")
            with st.expander("Ver"):
                for r in rows: st.markdown(f"- **{r['name']}** → `{r['entity_type']}`")
        except Exception as e:
            rows = []; st.error(str(e))
        if rows and token and account_id and st.button("🚀 Crear", type="primary", key="btn_doctype"):
            results = []; prog = st.progress(0); status = st.empty()
            for i, row in enumerate(rows):
                status.info(f"Creando: **{row['name']}** ({i+1}/{len(rows)})")
                code, resp = sr_post(
                    "/tms/api/v1/document-types/", token,
                    {"name": row["name"], "country": country,
                     "entity_type": row["entity_type"], "status": "active"},
                    base="https://api.simpliroute.com",
                    extra_headers={"Related-Account": account_id},
                    timeout=30,
                )
                results.append({"name": row["name"], "code": code, "resp": resp})
                prog.progress((i+1) / len(rows))
            status.empty(); prog.empty()
            show_results(results, "name")


def page_tms_transportistas():
    st.title("🚚 Transportistas TMS")
    c1, c2 = st.columns(2)
    with c1: token      = st.text_input("🔑 Token",      type="password", key="token_transp")
    with c2: account_id = st.text_input("🏢 Account ID", placeholder="Ej: 82761")

    def tpl():
        wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Transportistas"
        ws.append(["trade_name", "legal_name", "tax_id_number"])
        ws.append(["Transportes Sur S.A.", "Transportes Sur Sociedad Anónima", "20123456789"])
        ws.column_dimensions["A"].width = 30
        ws.column_dimensions["B"].width = 35
        ws.column_dimensions["C"].width = 20
        buf = io.BytesIO(); wb.save(buf); buf.seek(0); return buf

    st.download_button("📥 Plantilla", data=tpl(), file_name="plantilla_transportistas.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    tf = st.file_uploader("📂 Excel", type=["xlsx"], key="upload_transp")

    if tf:
        try:
            rows = read_excel_column(tf, ["trade_name", "legal_name", "tax_id_number"])
            st.success(f"✅ {len(rows)} transportista(s)")
            with st.expander("Ver"):
                for r in rows:
                    st.markdown(f"- **{r['trade_name']}** | {r['legal_name']} | `{r['tax_id_number']}`")
        except Exception as e:
            rows = []; st.error(str(e))
        if rows and token and account_id and st.button("🚀 Crear", type="primary", key="btn_transp"):
            results = []; prog = st.progress(0); status = st.empty()
            for i, row in enumerate(rows):
                status.info(f"Creando: **{row['trade_name']}** ({i+1}/{len(rows)})")
                code, resp = sr_post(
                    "/tms/api/v1/transportation-providers/", token,
                    {"trade_name": row["trade_name"], "legal_name": row["legal_name"],
                     "tax_id_number": row["tax_id_number"], "account_id": int(account_id)},
                    base="https://api.simpliroute.com",
                    timeout=30,
                )
                results.append({"trade_name": row["trade_name"], "code": code, "resp": resp})
                prog.progress((i+1) / len(rows))
            status.empty(); prog.empty()
            show_results(results, "trade_name")
