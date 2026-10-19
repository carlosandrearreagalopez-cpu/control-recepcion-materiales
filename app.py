from datetime import datetime
import os
from fpdf import FPDF
import openpyxl
import pandas as pd
import streamlit as st

# Configuración de diseño de página
st.set_page_config(
    page_title="Control de Recepción y Calidad",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
    <style>
    .main { background-color: #f8f9fa; }
    .stTabs [data-baseweb="tab-list"] { gap: 10px; }
    .stTabs [data-baseweb="tab"] {
        height: 50px;
        white-space: pre-wrap;
        background-color: #ffffff;
        border-radius: 5px;
        font-weight: 600;
        padding: 10px 20px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
    }
    .stTabs [aria-selected="true"] {
        background-color: #0066cc !important;
        color: #ffffff !important;
    }
    </style>
""", unsafe_allow_html=True)

st.title("📦 Sistema de Trazabilidad y Control de Calidad")
st.markdown("---")

ARCHIVO_PENDIENTES = "registros_pendientes.csv"
ARCHIVO_CATALOGO = "Maestro_Materiales_Proveedores (2).xlsx"


def cargar_pendientes():
  if os.path.exists(ARCHIVO_PENDIENTES):
    df = pd.read_csv(ARCHIVO_PENDIENTES)
    columnas_texto = [
        "Planta",
        "TipoMaterial",
        "NombreMaterial",
        "Proveedor",
        "LoteProveedor",
        "FacturaGuia",
        "Transporte",
        "Placa",
        "Unidad",
        "Temperatura",
        "FechaProduccion",
        "FechaVencimiento",
        "Criticidad",
        "Responsable",
        "Fecha",
        "Estado",
        "FirmaCalidad",
        "RutaCarpeta",
        "RutaExcel",
        "RutaPdfCert",
    ]
    for col in columnas_texto:
      if col not in df.columns:
        df[col] = ""
      else:
        df[col] = df[col].astype(str)
    return df
  else:
    return pd.DataFrame(columns=[
        "ID",
        "Planta",
        "TipoMaterial",
        "NombreMaterial",
        "Proveedor",
        "LoteProveedor",
        "FacturaGuia",
        "Transporte",
        "Placa",
        "Cantidad",
        "Unidad",
        "Temperatura",
        "FechaProduccion",
        "FechaVencimiento",
        "Criticidad",
        "Responsable",
        "Fecha",
        "Estado",
        "FirmaCalidad",
        "RutaCarpeta",
        "RutaExcel",
        "RutaPdfCert",
    ])


def guardar_pendientes(df):
  df.to_csv(ARCHIVO_PENDIENTES, index=False)


# Inicializar carrito temporal en sesión
if "items_ingreso_actual" not in st.session_state:
  st.session_state.items_ingreso_actual = []


# Cargar catálogo maestro
df_db = None
if os.path.exists(ARCHIVO_CATALOGO):
  try:
    df_db = pd.read_excel(ARCHIVO_CATALOGO)
    df_db.columns = df_db.columns.str.strip()
  except Exception as e:
    st.error(f"Error al leer el catálogo automático: {e}")

st.sidebar.header("⚙️ Panel de Administración")
archivo_catalogo_nuevo = st.sidebar.file_uploader(
    "Actualizar Catálogo Excel", type=["xlsx"]
)
if archivo_catalogo_nuevo is not None:
  df_db = pd.read_excel(archivo_catalogo_nuevo)
  df_db.columns = df_db.columns.str.strip()
  st.sidebar.success("¡Catálogo actualizado correctamente!")

tab1, tab2, tab3 = st.tabs([
    "📝 1. Ingreso de Operador",
    "✍️ 2. Panel de Calidad y Firma",
    "🔍 3. Búsqueda y Trazabilidad",
])

# ------------------------------------------
# PESTAÑA 1: INGRESO DE OPERADOR
# ------------------------------------------
with tab1:
  st.header("Formulario de Recepción de Materiales")
  st.markdown(
      "Complete los datos de transporte, factura, lote y seleccione los"
      " productos adjuntando su certificado."
  )

  if df_db is not None:
    col_planta = "Planta"
    col_tipo = "Tipo de material"
    col_material = "Material"
    col_proveedor = "Proveedor"
    col_criticidad = (
        "Criticidad" if "Criticidad" in df_db.columns else None
    )

    plantas = df_db[col_planta].dropna().unique()
    planta_sel = st.selectbox("Planta", plantas)

    df_planta = df_db[df_db[col_planta] == planta_sel]

    col1, col2 = st.columns(2)
    with col1:
      tipos_material = df_planta[col_tipo].dropna().unique()
      tipo_mat_sel = st.selectbox("Tipo de Material", tipos_material)
      df_tipo = df_planta[df_planta[col_tipo] == tipo_mat_sel]

    with col2:
      proveedores_disponibles = df_tipo[col_proveedor].dropna().unique()
      proveedor_sel = st.selectbox("Proveedor", proveedores_disponibles)
      df_prov = df_tipo[df_tipo[col_proveedor] == proveedor_sel]

    materiales_disponibles = df_prov[col_material].dropna().unique()
    material_sel = st.selectbox(
        "Nombre del Material (Producto)", materiales_disponibles
    )

    df_material = df_prov[df_prov[col_material] == material_sel]
    criticidad_val = "N/A"
    if col_criticidad and not df_material.empty:
      criticidad_val = str(df_material.iloc[0][col_criticidad])

    st.info(f"🏷️ **Criticidad Asociada al Material:** {criticidad_val}")

    material_lower = str(material_sel).lower()
    tipo_lower = str(tipo_mat_sel).lower()
    requiere_temperatura = any(
        kw in material_lower or kw in tipo_lower
        for kw in ["pulpa", "concentrado", "leche", "refrigerado", "congelado"]
    )

    with st.form("form_item_recepcion"):
      st.subheader("2. Datos de Transporte, Factura y Lote")
      t1, t2, t3 = st.columns(3)
      with t1:
        factura_guia = st.text_input(
            "No. de Factura y/o Guía", placeholder="Ej. FAC-12345"
        )
      with t2:
        nombre_transporte = st.text_input(
            "Nombre del Transporte", placeholder="Ej. Transportes S.A."
        )
      with t3:
        placa_transporte = st.text_input(
            "Placa del Transporte", placeholder="Ej. C-123ABC"
        )

      c3, c4, c5 = st.columns(3)
      with c3:
        lote_proveedor = st.text_input(
            "Lote del Proveedor", placeholder="Ej. LOTE-98765"
        )
      with c4:
        cantidad = st.number_input(
            "Cantidad Recibida", min_value=0.0, format="%.2f"
        )
      with c5:
        unidad = st.selectbox(
            "Unidad de Medida",
            ["Kg", "Litros", "Unidades", "Toneladas", "Galones", "Cajas"],
        )

      c6, c7 = st.columns(2)
      with c6:
        fecha_produccion = st.date_input("Fecha de Producción")
      with c7:
        fecha_vencimiento = st.date_input("Fecha de Vencimiento / Caducidad")

      temperatura = "N/A"
      if requiere_temperatura:
        st.markdown("---")
        st.subheader("❄️ Control de Temperatura (Requerido)")
        temperatura = st.number_input(
            "Temperatura de Recepción (°C)", value=4.0, format="%.1f"
        )

      st.markdown("---")
      st.subheader("📎 Certificado de Calidad del Producto")
      certificado_file = st.file_uploader(
          "Subir Foto o PDF del Certificado", type=["jpg", "jpeg", "png", "pdf"]
      )

      agregar_item = st.form_submit_button(
          "➕ Agregar este producto a la lista de ingreso"
      )

    if agregar_item:
      if not lote_proveedor or not factura_guia:
        st.warning(
            "⚠️ Debes ingresar el número de factura/guía y el número de lote."
        )
      elif certificado_file is None:
        st.warning(
            "⚠️ Es obligatorio adjuntar la foto o PDF del certificado para este"
            " producto."
        )
      else:
        os.makedirs("temp_certificados", exist_ok=True)
        ruta_cert_temp = os.path.join(
            "temp_certificados", certificado_file.name
        )
        with open(ruta_cert_temp, "wb") as f:
          f.write(certificado_file.getbuffer())

        nuevo_item = {
            "Planta": planta_sel,
            "TipoMaterial": tipo_mat_sel,
            "Proveedor": proveedor_sel,
            "NombreMaterial": material_sel,
            "FacturaGuia": factura_guia,
            "Transporte": nombre_transporte,
            "Placa": placa_transporte,
            "Lote": lote_proveedor,
            "Cantidad": cantidad,
            "Unidad": unidad,
            "Temperatura": temperatura,
            "Produccion": str(fecha_produccion),
            "Vencimiento": str(fecha_vencimiento),
            "Criticidad": criticidad_val,
            "RutaTemp": ruta_cert_temp,
        }
        st.session_state.items_ingreso_actual.append(nuevo_item)
        st.success(
            f"✅ ¡Producto '{material_sel}' agregado con su certificado a la"
            " lista!"
        )

    if len(st.session_state.items_ingreso_actual) > 0:
      st.markdown("---")
      st.subheader("📋 Resumen de Productos Agregados en este Ingreso")
      df_carrito = pd.DataFrame(st.session_state.items_ingreso_actual)
      st.dataframe(
          df_carrito[[
              "Planta",
              "Proveedor",
              "NombreMaterial",
              "FacturaGuia",
              "Lote",
              "Cantidad",
              "Unidad",
          ]],
          use_container_width=True,
      )

      if st.button("🗑️ Limpiar lista actual"):
        st.session_state.items_ingreso_actual = []
        st.rerun()

      st.markdown("---")
      st.subheader("3. Finalizar y Enviar a Revisión de Calidad")
      with st.form("form_finalizar_envio"):
        responsable = st.text_input(
            "Operador Responsable", placeholder="Tu Nombre Completo"
        )
        enviar_todo = st.form_submit_button(
            "🚀 Enviar Todos los Productos a Calidad", type="primary"
        )

      if enviar_todo:
        if not responsable:
          st.error("⚠️ Por favor ingresa el nombre del operador responsable.")
        else:
          df_p = cargar_pendientes()
          nuevo_id = int(df_p["ID"].max()) + 1 if not df_p.empty and not pd.isna(df_p["ID"].max()) else 1

          for item in st.session_state.items_ingreso_actual:
            nuevo_registro = {
                "ID": nuevo_id,
                "Planta": str(item["Planta"]),
                "TipoMaterial": str(item["TipoMaterial"]),
                "NombreMaterial": str(item["NombreMaterial"]),
                "Proveedor": str(item["Proveedor"]),
                "FacturaGuia": str(item["FacturaGuia"]),
                "Transporte": str(item["Transporte"]),
                "Placa": str(item["Placa"]),
                "LoteProveedor": str(item["Lote"]),
                "Cantidad": item["Cantidad"],
                "Unidad": str(item["Unidad"]),
                "Temperatura": str(item["Temperatura"]),
                "FechaProduccion": str(item["Produccion"]),
                "FechaVencimiento": str(item["Vencimiento"]),
                "Criticidad": str(item["Criticidad"]),
                "Responsable": str(responsable),
                "Fecha": datetime.now().strftime("%Y-%m-%d %H:%M"),
                "Estado": "Pendiente",
                "FirmaCalidad": "Sin firmar",
                "RutaCarpeta": "",
                "RutaExcel": "",
                "RutaPdfCert": str(item["RutaTemp"]),
            }
            df_p = pd.concat(
                [df_p, pd.DataFrame([nuevo_registro])], ignore_index=True
            )
            nuevo_id += 1

          guardar_pendientes(df_p)
          st.session_state.items_ingreso_actual = []
          st.success(
              "🎉 ¡Todos los productos con sus certificados fueron enviados"
              " con éxito al panel de calidad!"
          )
  else:
    st.warning(
        "⚠️ No se detectó automáticamente el catálogo en el repositorio."
        " Asegúrate de haber subido el archivo"
        " `Maestro_Materiales_Proveedores (2).xlsx` a GitHub."
    )

# ------------------------------------------
# PESTAÑA 2: PANEL DE CALIDAD Y FIRMA
# ------------------------------------------
with tab2:
  st.header("✍️ Panel de Control y Firma de Calidad")
  st.markdown(
      "Espacio exclusivo para supervisores. Ingrese su contraseña, revise los"
      " datos y apruebe para rellenar exactamente el checklist oficial y"
      " generar el certificado PDF."
  )

  clave_ingresada = st.text_input(
      "Contraseña de Supervisor de Calidad", type="password", key="pass_calidad"
  )
  CLAVE_SUPERVISOR = "1234"

  if clave_ingresada == CLAVE_SUPERVISOR:
    st.success("🔓 Acceso de supervisor autorizado.")
    df_p = cargar_pendientes()
    pendientes = (
        df_p[df_p["Estado"] == "Pendiente"] if not df_p.empty else pd.DataFrame()
    )

    if len(pendientes) == 0:
      st.info(
          "No hay registros pendientes de aprobación en este momento."
          " (Puedes ingresar nuevos registros en la Pestaña 1)."
      )
    else:
      for index, row in pendientes.iterrows():
        with st.expander(
            f"📥 Recepción ID #{row['ID']} | Material: {row['NombreMaterial']}"
            f" | Lote: {row['LoteProveedor']}"
        ):
          c1, c2 = st.columns(2)
          with c1:
            st.write(f"**Planta:** {row['Planta']}")
            st.write(f"**Proveedor:** {row['Proveedor']}")
            st.write(f"**Factura / Guía:** {row['FacturaGuia']}")
            st.write(f"**Transporte:** {row['Transporte']} (Placa: {row['Placa']})")
          with c2:
            st.write(f"**Cantidad:** {row['Cantidad']} {row['Unidad']}")
            st.write(f"**Prod. / Venc.:** {row['FechaProduccion']} al {row['FechaVencimiento']}")
            st.write(f"**Operador:** {row['Responsable']}")

          if row["RutaPdfCert"] and os.path.exists(str(row["RutaPdfCert"])):
            if str(row["RutaPdfCert"]).endswith((".jpg", ".jpeg", ".png")):
              st.image(
                  row["RutaPdfCert"], caption="Certificado Adjunto", width=350
              )
            else:
              st.write("📄 Certificado adjunto disponible.")

          with st.form(f"form_firma_{row['ID']}"):
            nombre_supervisor = st.text_input(
                "Nombre del Supervisor (Firma Digital)", key=f"sup_{row['ID']}"
            )
            aprobar = st.form_submit_button(
                "Aprobar, Generar Certificado PDF e Inyectar en Checklist 🛡️",
                type="primary",
            )

          if aprobar:
            if not nombre_supervisor:
              st.error(
                  "Debe ingresar su nombre para autorizar y firmar el registro."
              )
            else:
              with st.spinner(
                  "Generando certificado oficial PDF y completando Excel..."
              ):
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

                # 1. GENERAR CERTIFICADO PDF OFICIAL IDÉNTICO
                nombre_cert_oficial = (
                    f"Certificado_Firmado_{row['LoteProveedor']}.pdf"
                )
                ruta_cert_final = os.path.join(
                    ruta_destino, nombre_cert_oficial
                )

                pdf = FPDF()
                pdf.add_page()
                pdf.set_font("Arial", "B", 16)
                pdf.cell(
                    0,
                    10,
                    f"Certificado de Calidad - {row['NombreMaterial']}",
                    0,
                    1,
                    "C",
                )
                pdf.set_font("Arial", "", 12)
                pdf.cell(
                    0,
                    10,
                    f"Planta: {row['Planta']} | Proveedor: {row['Proveedor']}",
                    0,
                    1,
                    "C",
                )
                pdf.cell(
                    0,
                    10,
                    f"Lote: {row['LoteProveedor']} | Factura/Guía: {row['FacturaGuia']}",
                    0,
                    1,
                    "C",
                )
                pdf.cell(
                    0,
                    10,
                    (
                        f"Aprobado por: {nombre_supervisor} | Fecha:"
                        f" {row['Fecha']}"
                    ),
                    0,
                    1,
                    "C",
                )
                pdf.ln(10)

                if row["RutaPdfCert"] and os.path.exists(
                    str(row["RutaPdfCert"])
                ):
                  if str(row["RutaPdfCert"]).endswith((".jpg", ".jpeg", ".png")):
                    pdf.image(row["RutaPdfCert"], x=15, y=65, w=180)

                pdf.output(ruta_cert_final)

                # 2. INYECTAR DATOS EN TU CHECKLIST OFICIAL DE EXCEL
                prefijo_archivo = (
                    f"Checklist_{row['NombreMaterial']}_Lote_{row['LoteProveedor']}".replace(
                        " ", "_"
                    )
                )
                
                archivos_repo = os.listdir('.')
                plantillas_disponibles = [
                    f for f in archivos_repo 
                    if f.endswith('.xlsx') and ("R ICC" in f or "Check" in f or "Checklist" in f) and f != "Maestro_Materiales_Proveedores (2).xlsx"
                ]
                
                ruta_plantilla = plantillas_disponibles[0] if plantillas_disponibles else "R ICC 14-1 Check List Recepción Materiales.xlsx"
                ruta_excel_final = os.path.join(
                    ruta_destino, f"{prefijo_archivo}.xlsx"
                )

                if os.path.exists(ruta_plantilla):
                  wb = openpyxl.load_workbook(ruta_plantilla)
                  ws = wb["Recepción"] if "Recepción" in wb.sheetnames else wb.active
                  
                  # Celdas exactas según tu formato oficial del PDF
                  ws['C5'] = row['Fecha']                   # Fecha de recepción
                  ws['C16'] = row['Proveedor']              # Proveedor
                  ws['C17'] = row['NombreMaterial']         # Material
                  ws['C18'] = float(row['Cantidad']) if pd.notna(row['Cantidad']) else 0.0 # Cantidad
                  ws['C19'] = row['Unidad']                 # Unidad de medida
                  ws['C20'] = row['LoteProveedor']          # Lote
                  ws['C21'] = row['FechaProduccion']        # Fecha de Producción
                  ws['C22'] = row['FechaVencimiento']       # Fecha de Vencimiento
                  
                  if str(row['Temperatura']) != "N/A":
                    ws['D39'] = float(row['Temperatura'])   # Temperatura del material
                  
                  ws['C51'] = row['Responsable']            # Persona que realiza la recepción
                  ws['C53'] = nombre_supervisor             # Firma de Calidad

                  wb.save(ruta_excel_final)
                else:
                  df_resp = pd.DataFrame([row])
                  df_resp.to_excel(ruta_excel_final, index=False)

                # Actualizar estado en CSV temporal
                df_p.loc[df_p["ID"] == row["ID"], "Estado"] = "Aprobado"
                df_p.loc[df_p["ID"] == row["ID"], "FirmaCalidad"] = str(
                    nombre_supervisor
                )
                df_p.loc[df_p["ID"] == row["ID"], "RutaCarpeta"] = str(
                    ruta_destino
                )
                df_p.loc[df_p["ID"] == row["ID"], "RutaExcel"] = str(
                    ruta_excel_final
                )
                df_p.loc[df_p["ID"] == row["ID"], "RutaPdfCert"] = str(
                    ruta_cert_final
                )
                guardar_pendientes(df_p)

                st.success(
                    "🎉 ¡Certificado PDF y Checklist de Excel oficial generados"
                    f" con éxito! Guardado en: `{ruta_destino}`"
                )
                st.rerun()
  elif clave_ingresada == "":
    st.info("🔒 Por favor, ingrese la contraseña de supervisor (1234) para ver los registros pendientes.")
  else:
    st.error("❌ Contraseña incorrecta. Intente nuevamente.")

# ------------------------------------------
# PESTAÑA 3: BÚSQUEDA Y TRAZABILIDAD
# ------------------------------------------
with tab3:
  st.header("🔍 Buscador y Consulta de Registros Oficiales")
  st.markdown(
      "Consulte el historial de recepciones aprobadas, descargue su checklist"
      " oficial de Excel rellenado y los certificados firmados."
  )

  df_p = cargar_pendientes()
  aprobados = (
      df_p[df_p["Estado"] == "Aprobado"] if not df_p.empty else pd.DataFrame()
  )

  if len(aprobados) == 0:
    st.info("Aún no hay registros aprobados y archivados en el sistema.")
  else:
    busqueda = st.text_input(
        "Filtro de búsqueda rápida",
        placeholder="Escriba el Lote, Material o Proveedor...",
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

    st.dataframe(
        resultado[[
            "Planta",
            "TipoMaterial",
            "NombreMaterial",
            "Proveedor",
            "FacturaGuia",
            "LoteProveedor",
            "Cantidad",
            "Unidad",
            "FechaProduccion",
            "FechaVencimiento",
            "FirmaCalidad",
        ]],
        use_container_width=True,
    )

    st.markdown("---")
    st.subheader("📥 Descarga de Archivos Oficiales por Registro")

    for idx, row in resultado.iterrows():
      with st.expander(
          f"📁 Material: {row['NombreMaterial']} | Lote: {row['LoteProveedor']}"
          f" | Aprobado por: {row['FirmaCalidad']}"
      ):
        col_d1, col_d2 = st.columns(2)

        with col_d1:
          if "RutaExcel" in row and pd.notna(row["RutaExcel"]) and str(row["RutaExcel"]).strip() != "" and os.path.exists(str(row["RutaExcel"])):
            with open(row["RutaExcel"], "rb") as f:
              st.download_button(
                  label="📊 Descargar Checklist Excel Oficial",
                  data=f,
                  file_name=os.path.basename(str(row["RutaExcel"])),
                  mime=(
                      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                  ),
                  key=f"dl_excel_{row['ID']}",
              )
          else:
            st.info("Archivo Excel en proceso o no disponible.")

        with col_d2:
          if "RutaPdfCert" in row and pd.notna(row["RutaPdfCert"]) and str(row["RutaPdfCert"]).strip() != "" and os.path.exists(str(row["RutaPdfCert"])):
            with open(row["RutaPdfCert"], "rb") as f:
              st.download_button(
                  label="📄 Descargar Certificado Firmado (PDF)",
                  data=f,
                  file_name=os.path.basename(str(row["RutaPdfCert"])),
                  mime="application/pdf",
                  key=f"dl_pdf_{row['ID']}",
              )
          else:
            st.info("Certificado PDF en proceso o no disponible.")
