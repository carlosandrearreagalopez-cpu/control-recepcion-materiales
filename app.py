from datetime import datetime
import os
from fpdf import FPDF
import openpyxl
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Control de Recepción y Calidad", layout="wide"
)

st.title("📦 Sistema de Trazabilidad y Control de Recepción")

ARCHIVO_PENDIENTES = "registros_pendientes.csv"


def cargar_pendientes():
  if os.path.exists(ARCHIVO_PENDIENTES):
    return pd.read_csv(ARCHIVO_PENDIENTES)
  else:
    return pd.DataFrame(columns=[
        "ID",
        "Planta",
        "TipoMaterial",
        "NombreMaterial",
        "Proveedor",
        "LoteProveedor",
        "Responsable",
        "Fecha",
        "Estado",
        "FirmaCalidad",
        "RutaCarpeta",
    ])


def guardar_pendientes(df):
  df.to_csv(ARCHIVO_PENDIENTES, index=False)


st.sidebar.header("⚙️ Configuración del Sistema")
archivo_catalogo = st.sidebar.file_uploader(
    "Actualizar Catálogo (proveedores_materiales.xlsx)", type=["xlsx"]
)

if archivo_catalogo is not None:
  df_db = pd.read_excel(archivo_catalogo)
  st.sidebar.success("¡Catálogo actualizado!")
else:
  if os.path.exists("proveedores_materiales.xlsx"):
    df_db = pd.read_excel("proveedores_materiales.xlsx")
  else:
    df_db = None
    st.sidebar.warning("⚠️ Sube tu Excel de proveedores para habilitar el sistema.")

tab1, tab2, tab3 = st.tabs([
    "📝 1. Ingreso de Operador",
    "✍️ 2. Panel de Aprobación y Firma",
    "🔍 3. Búsqueda y Trazabilidad",
])

with tab1:
  st.header("Formulario de Recepción de Materiales")
  st.write(
      "Complete los datos correspondientes al checklist y adjunte el"
      " certificado."
  )

  if df_db is not None:
    with st.form("form_operador"):
      col1, col2 = st.columns(2)

      with col1:
        plantas = df_db["Planta"].dropna().unique()
        planta_sel = st.selectbox("Planta", plantas)

        tipos_material = (
            df_db[df_db["Planta"] == planta_sel]["TipoMaterial"]
            .dropna()
            .unique()
        )
        tipo_mat_sel = st.selectbox("Tipo de Material", tipos_material)

      with col2:
        materiales = (
            df_db[
                (df_db["Planta"] == planta_sel)
                & (df_db["TipoMaterial"] == tipo_mat_sel)
            ]["NombreMaterial"]
            .dropna()
            .unique()
        )
        material_sel = st.selectbox("Nombre del Material", materiales)

        proveedores_disponibles = (
            df_db[
                (df_db["Planta"] == planta_sel)
                & (df_db["TipoMaterial"] == tipo_mat_sel)
                & (df_db["NombreMaterial"] == material_sel)
            ]["Proveedor"]
            .dropna()
            .unique()
        )
        proveedor_sel = st.selectbox("Proveedor", proveedores_disponibles)

      st.markdown("---")
      col3, col4 = st.columns(2)
      with col3:
        lote_proveedor = st.text_input("Lote del Proveedor")
      with col4:
        responsable = st.text_input("Nombre del Operador Responsable")

      st.markdown("---")
      st.subheader("📎 Certificado de Calidad (Imagen o PDF)")
      certificado_file = st.file_uploader(
          "Subir archivo de certificado", type=["jpg", "jpeg", "png", "pdf"]
      )

      enviar_revision = st.form_submit_button(
          "Enviar a Revisión de Calidad 🚀"
      )

    if enviar_revision:
      if not lote_proveedor or not responsable:
        st.error("Por favor, completa el lote y el nombre del operador.")
      elif certificado_file is None:
        st.error("Es obligatorio adjuntar el certificado de calidad.")
      else:
        os.makedirs("temp_certificados", exist_ok=True)
        ruta_cert_temp = os.path.join(
            "temp_certificados", certificado_file.name
        )
        with open(ruta_cert_temp, "wb") as f:
          f.write(certificado_file.getbuffer())

        df_p = cargar_pendientes()
        nuevo_id = len(df_p) + 1
        nuevo_registro = {
            "ID": nuevo_id,
            "Planta": planta_sel,
            "TipoMaterial": tipo_mat_sel,
            "NombreMaterial": material_sel,
            "Proveedor": proveedor_sel,
            "LoteProveedor": lote_proveedor,
            "Responsable": responsable,
            "Fecha": datetime.now().strftime("%Y-%m-%d %H:%M"),
            "Estado": "Pendiente",
            "FirmaCalidad": "Sin firmar",
            "RutaCarpeta": ruta_cert_temp,
        }
        df_p = pd.concat(
            [df_p, pd.DataFrame([nuevo_registro])], ignore_index=True
        )
        guardar_pendientes(df_p)

        st.success(
            "¡Registro enviado con éxito al panel de calidad para su"
            " revisión y firma!"
        )
  else:
    st.info("Sube el catálogo de proveedores en la barra lateral para comenzar.")

with tab2:
  st.header("✍️ Panel de Control, Revisión y Firma de Calidad")
  st.write(
      "Revise los ingresos pendientes, ingrese su firma/nombre y apruebe para"
      " generar la estructura final."
  )

  df_p = cargar_pendientes()
  pendientes = df_p[df_p["Estado"] == "Pendiente"]

  if len(pendientes) == 0:
    st.info("No hay registros pendientes de aprobación.")
  else:
    for index, row in pendientes.iterrows():
      with st.expander(
          f"📥 Recepción #{row['ID']} - {row['NombreMaterial']} (Lote:"
          f" {row['LoteProveedor']})"
      ):
        st.write(f"**Planta:** {row['Planta']}")
        st.write(f"**Tipo de Material:** {row['TipoMaterial']}")
        st.write(f"**Proveedor:** {row['Proveedor']}")
        st.write(f"**Operador:** {row['Responsable']}")
        st.write(f"**Fecha:** {row['Fecha']}")

        if os.path.exists(row["RutaCarpeta"]):
          if row["RutaCarpeta"].endswith((".jpg", ".jpeg", ".png")):
            st.image(row["RutaCarpeta"], caption="Certificado Adjunto", width=300)
          else:
            st.write(
                "📄 Archivo de certificado PDF adjunto listo para procesar."
            )

        with st.form(f"form_firma_{row['ID']}"):
          nombre_supervisor = st.text_input(
              "Firma Digital (Escriba su Nombre Completo como Supervisor de"
              " Calidad)"
          )
          aprobar = st.form_submit_button("Aprobar, Firmar y Generar Carpetas 🛡️")

        if aprobar:
          if not nombre_supervisor:
            st.error("Debe ingresar su nombre para firmar el registro.")
          else:
            with st.spinner("Generando estructura jerárquica y archivos..."):
              carpeta_base = "./Registros_Oficiales"
              subcarpeta_lote = (
                  f"Lote_{row['LoteProveedor']}_{row['NombreMaterial']}".replace(
                      " ", "_"
                  )
              )

              ruta_destino = os.path.join(
                  carpeta_base,
                  str(row["Planta"]),
                  str(row["TipoMaterial"]),
                  str(row["NombreMaterial"]),
                  subcarpeta_lote,
              )
              os.makedirs(ruta_destino, exist_ok=True)

              nombre_cert_oficial = f"Certificado_{row['LoteProveedor']}.pdf"
              ruta_cert_final = os.path.join(
                  ruta_destino, nombre_cert_oficial
              )
              if os.path.exists(row["RutaCarpeta"]):
                if row["RutaCarpeta"].endswith((".jpg", ".jpeg", ".png")):
                  pdf = FPDF()
                  pdf.add_page()
                  pdf.set_font("Arial", "B", 14)
                  pdf.cell(
                      0,
                      10,
                      f"Certificado de Calidad - {row['NombreMaterial']}",
                      0,
                      1,
                      "C",
                  )
                  pdf.set_font("Arial", "", 11)
                  pdf.cell(
                      0,
                      10,
                      f"Planta: {row['Planta']} | Lote: {row['LoteProveedor']}",
                      0,
                      1,
                      "C",
                  )
                  pdf.cell(
                      0,
                      10,
                      f"Aprobado y Firmado por: {nombre_supervisor}",
                      0,
                      1,
                      "C",
                  )
                  pdf.ln(5)
                  pdf.image(row["RutaCarpeta"], x=15, y=45, w=180)
                  pdf.output(ruta_cert_final)
                else:
                  os.rename(row["RutaCarpeta"], ruta_cert_final)

              prefijo_archivo = (
                  f"Checklist_{row['NombreMaterial']}_{row['LoteProveedor']}".replace(
                      " ", "_"
                  )
              )
              ruta_plantilla = "plantilla_checklist.xlsx"

              if os.path.exists(ruta_plantilla):
                wb = openpyxl.load_workbook(ruta_plantilla)
                ws = wb.active
                ruta_excel_final = os.path.join(
                    ruta_destino, f"{prefijo_archivo}.xlsx"
                )
                wb.save(ruta_excel_final)

              df_p.loc[df_p["ID"] == row["ID"], "Estado"] = "Aprobado"
              df_p.loc[df_p["ID"] == row["ID"], "FirmaCalidad"] = (
                  nombre_supervisor
              )
              df_p.loc[df_p["ID"] == row["ID"], "RutaCarpeta"] = ruta_destino
              guardar_pendientes(df_p)

              st.success(
                  f"¡Registro aprobado y guardado en la ruta: {ruta_destino}!"
              )
              st.rerun()

with tab3:
  st.header("🔍 Buscador y Consulta de Registros Oficiales")
  st.write(
      "Busca rápidamente por lote, material o proveedor dentro de la"
      " estructura."
  )

  df_p = cargar_pendientes()
  aprobados = df_p[df_p["Estado"] == "Aprobado"]

  if len(aprobados) == 0:
    st.info("Aún no hay registros aprobados y archivados.")
  else:
    busqueda = st.text_input(
        "Buscar por Lote, Material o Proveedor", placeholder="Ej. Lote 12345..."
    )

    if busqueda:
      resultado = aprobados[
          aprobados["LoteProveedor"]
          .astype(str)
          .str.contains(busqueda, case=False, na=False)
          | aprobados["NombreMaterial"]
          .astype(str)
          .str.contains(busqueda, case=False, na=False)
          | aprobados["Proveedor"]
          .astype(str)
          .str.contains(busqueda, case=False, na=False)
      ]
    else:
      resultado = aprobados

    st.dataframe(resultado[[
        "Planta",
        "TipoMaterial",
        "NombreMaterial",
        "Proveedor",
        "LoteProveedor",
        "Fecha",
        "FirmaCalidad",
    ]])
