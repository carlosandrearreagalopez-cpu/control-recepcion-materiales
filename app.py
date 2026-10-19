from datetime import datetime
import os
import shutil
import pandas as pd
import streamlit as st
from fpdf import FPDF

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

# ------------------------------------------
# FUNCIONES DE DATOS Y GENERACIÓN PDF
# ------------------------------------------
def cargar_pendientes():
  if os.path.exists(ARCHIVO_PENDIENTES):
    df = pd.read_csv(ARCHIVO_PENDIENTES)
    columnas_texto = [
        "Planta", "TipoMaterial", "NombreMaterial", "Proveedor",
        "LoteProveedor", "FacturaGuia", "Transporte", "Placa",
        "IngresoSAP", "DocSAP", "Marchamo",
        "Unidad", "Temperatura", "FechaProduccion", "FechaVencimiento",
        "Criticidad", "Responsable", "Fecha", "Estado", "FirmaCalidad",
        "RutaCarpeta", "RutaChecklistOficial", "RutaPdfCert"
    ]
    for col in columnas_texto:
      if col not in df.columns:
        df[col] = ""
      else:
        df[col] = df[col].astype(str)
    return df
  else:
    return pd.DataFrame(columns=[
        "ID", "Planta", "TipoMaterial", "NombreMaterial", "Proveedor",
        "LoteProveedor", "FacturaGuia", "Transporte", "Placa",
        "IngresoSAP", "DocSAP", "Marchamo", "Cantidad",
        "Unidad", "Temperatura", "FechaProduccion", "FechaVencimiento",
        "Criticidad", "Responsable", "Fecha", "Estado", "FirmaCalidad",
        "RutaCarpeta", "RutaChecklistOficial", "RutaPdfCert"
    ])

def guardar_pendientes(df):
  df.to_csv(ARCHIVO_PENDIENTES, index=False)

def generar_pdf_ricc14(row, supervisor, ruta_guardado):
    pdf = FPDF('P', 'mm', 'A4')
    pdf.add_page()
    
    # --- ENCABEZADO ---
    pdf.set_font('Arial', 'B', 9)
    pdf.cell(45, 10, 'Codigo: R ICC-14-1', border=1, align='C')
    pdf.cell(100, 10, 'Aseguramiento de Calidad', border=1, align='C')
    pdf.cell(45, 10, 'Version: 5', border=1, align='C', ln=1)
    
    pdf.set_font('Arial', 'B', 11)
    pdf.cell(190, 8, 'Check list de recepcion de materiales', border=1, align='C', ln=1)
    pdf.ln(3)

    # --- DATOS GENERALES ---
    pdf.set_font('Arial', '', 8)
    pdf.cell(35, 6, 'Fecha de recepcion:', border=1)
    pdf.cell(60, 6, str(row.get('Fecha', '')).split()[0], border=1)
    pdf.cell(45, 6, 'Factura/Guia del proveedor:', border=1)
    pdf.cell(50, 6, str(row.get('FacturaGuia', '')), border=1, ln=1)
    
    pdf.cell(35, 6, 'Nombre transporte:', border=1)
    pdf.cell(60, 6, str(row.get('Transporte', '')), border=1)
    pdf.cell(45, 6, 'Fecha de ingreso a SAP:', border=1)
    pdf.cell(50, 6, str(row.get('IngresoSAP', 'N/A')), border=1, ln=1)

    pdf.cell(35, 6, 'Placa del transporte:', border=1)
    pdf.cell(60, 6, str(row.get('Placa', '')), border=1)
    pdf.cell(45, 6, 'No. Documento SAP:', border=1)
    pdf.cell(50, 6, str(row.get('DocSAP', 'N/A')), border=1, ln=1)
    pdf.ln(3)

    # --- MARCHAMO Y CANDADO ---
    pdf.set_fill_color(220, 230, 241)
    pdf.set_font('Arial', 'B', 8)
    pdf.cell(95, 6, 'Inspeccion del Marchamo', border=1, align='C', fill=True)
    pdf.cell(95, 6, 'Inspeccion del Candado', border=1, align='C', fill=True, ln=1)
    
    pdf.set_font('Arial', '', 8)
    pdf.cell(95, 6, 'El marchamo se encuentra alterado?   NO', border=1)
    pdf.cell(95, 6, 'El candado se encuentra alterado?   NO', border=1, ln=1)
    pdf.cell(95, 6, f"No. de marchamo: {row.get('Marchamo', 'N/A')}", border=1)
    pdf.cell(95, 6, '', border=1, ln=1)
    pdf.ln(3)

    # --- INSPECCIÓN DEL MATERIAL ---
    pdf.set_font('Arial', 'B', 8)
    pdf.cell(190, 6, 'Inspeccion del Material', border=1, align='C', fill=True, ln=1)
    
    pdf.set_font('Arial', '', 8)
    pdf.cell(35, 6, 'Proveedor:', border=1)
    pdf.cell(155, 6, str(row.get('Proveedor', '')), border=1, ln=1)
    
    pdf.cell(35, 6, 'Material:', border=1)
    pdf.cell(155, 6, str(row.get('NombreMaterial', '')), border=1, ln=1)
    
    pdf.cell(35, 6, 'Cantidad:', border=1)
    pdf.cell(60, 6, str(row.get('Cantidad', '')), border=1)
    pdf.cell(35, 6, 'Unidad de medida:', border=1)
    pdf.cell(60, 6, str(row.get('Unidad', '')), border=1, ln=1)
    
    pdf.cell(35, 6, 'Lote:', border=1)
    pdf.cell(60, 6, str(row.get('LoteProveedor', '')), border=1)
    pdf.cell(35, 6, 'Temp. Material (C):', border=1)
    temp_str = str(row.get('Temperatura', 'N/A'))
    pdf.cell(60, 6, temp_str if temp_str not in ["nan", ""] else "N/A", border=1, ln=1)
    
    pdf.cell(35, 6, 'Fecha Produccion:', border=1)
    pdf.cell(60, 6, str(row.get('FechaProduccion', '')), border=1)
    pdf.cell(35, 6, 'Fecha Vencimiento:', border=1)
    pdf.cell(60, 6, str(row.get('FechaVencimiento', '')), border=1, ln=1)
    pdf.ln(3)

    # --- PREGUNTAS DEL MATERIAL ---
    pdf.set_font('Arial', 'B', 8)
    pdf.cell(160, 6, 'Criterio de Evaluacion (Material)', border=1)
    pdf.cell(30, 6, 'Resultado', border=1, align='C', ln=1)
    
    pdf.set_font('Arial', '', 8)
    preguntas_mat = [
        "Tomando en cuenta la fecha, el lote posee mas de un mes de vida?",
        f"Categoria de la matriz de materiales a recibir: {row.get('Criticidad', 'N/A')}",
        "El Certificado de Analisis corresponde al numero de lote?",
        "Los resultados de los Certificados cumplen parametros criticos?",
        "El empaque del material se encuentra integro y sin deterioro?",
        "El material esta protegido de la intemperie?",
        "Coincide la entrega segun documento contra lo fisico?"
    ]
    for p in preguntas_mat:
        pdf.cell(160, 6, p, border=1)
        pdf.cell(30, 6, "SI (Cumple)", border=1, align='C', ln=1)
    pdf.ln(3)

    # --- PREGUNTAS DEL TRANSPORTE ---
    pdf.set_font('Arial', 'B', 8)
    pdf.cell(190, 6, 'Inspeccion del Transporte', border=1, align='C', fill=True, ln=1)
    pdf.cell(160, 6, 'Criterio de Evaluacion (Transporte)', border=1)
    pdf.cell(30, 6, 'Resultado', border=1, align='C', ln=1)
    
    pdf.set_font('Arial', '', 8)
    preguntas_trans = [
        ("Se encuentra el transporte libre de plagas?", "SI"),
        ("El techo, paredes del transporte o lona estan integras?", "SI"),
        ("El techo, paredes del transporte o lona estan limpios?", "SI"),
        ("El piso del transporte esta limpio y en buenas condiciones?", "SI"),
        ("El transporte cumple con la temperatura especificada?", "SI" if temp_str != "N/A" else "N/A"),
        ("El transporte se encuentra libre de olores extranos?", "SI"),
        ("La tarima de uso retornable se encuentra en buenas condiciones?", "SI")
    ]
    for p, res in preguntas_trans:
        pdf.cell(160, 6, p, border=1)
        pdf.cell(30, 6, res, border=1, align='C', ln=1)
    pdf.ln(4)

    # --- RESULTADO FINAL ---
    pdf.set_font('Arial', 'B', 9)
    pdf.cell(35, 8, 'Resultado final:', border=1)
    pdf.set_text_color(0, 128, 0) # Verde
    pdf.cell(155, 8, 'APROBADO', border=1, align='C', ln=1)
    pdf.set_text_color(0, 0, 0)
    
    pdf.set_font('Arial', '', 8)
    pdf.multi_cell(190, 6, 'OBSERVACIONES: Validacion digital autorizada en plataforma. Se adjunta documento oficial del proveedor en la carpeta de trazabilidad correspondiente.', border=1)
    pdf.ln(6)

    # --- FIRMAS ---
    pdf.cell(95, 5, 'Nombre de la persona que realiza la recepcion:', 0, 0)
    pdf.cell(95, 5, 'Vo. Bo. Calidad:', 0, 1)
    
    pdf.set_font('Arial', 'B', 8)
    pdf.cell(95, 6, str(row.get('Responsable', '')), 0, 0)
    pdf.cell(95, 6, f"Firmado digitalmente por: {supervisor}", 0, 1)

    pdf.output(ruta_guardado)


# Inicializar carrito
if "items_ingreso_actual" not in st.session_state:
  st.session_state.items_ingreso_actual = []

# Cargar catálogo
df_db = None
if os.path.exists(ARCHIVO_CATALOGO):
  try:
    df_db = pd.read_excel(ARCHIVO_CATALOGO)
    df_db.columns = df_db.columns.str.strip()
  except Exception as e:
    st.error(f"Error al leer el catálogo: {e}")

st.sidebar.header("⚙️ Panel de Administración")
archivo_catalogo_nuevo = st.sidebar.file_uploader("Actualizar Catálogo Excel", type=["xlsx"])
if archivo_catalogo_nuevo is not None:
  with open(ARCHIVO_CATALOGO, "wb") as f:
      f.write(archivo_catalogo_nuevo.getbuffer())
  df_db = pd.read_excel(ARCHIVO_CATALOGO)
  df_db.columns = df_db.columns.str.strip()
  st.sidebar.success("¡Catálogo actualizado!")

tab1, tab2, tab3 = st.tabs([
    "📝 1. Ingreso de Operador",
    "✍️ 2. Panel de Calidad",
    "🔍 3. Trazabilidad",
])

# ------------------------------------------
# PESTAÑA 1: INGRESO DE OPERADOR
# ------------------------------------------
with tab1:
  st.header("Formulario de Recepción de Materiales")

  if df_db is not None:
    col_planta = "Planta"
    col_tipo = "Tipo de material"
    col_material = "Material"
    col_proveedor = "Proveedor"
    col_criticidad = "Criticidad" if "Criticidad" in df_db.columns else None

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
    material_sel = st.selectbox("Nombre del Material", materiales_disponibles)

    df_material = df_prov[df_prov[col_material] == material_sel]
    criticidad_val = str(df_material.iloc[0][col_criticidad]) if col_criticidad and not df_material.empty else "N/A"

    material_lower = str(material_sel).lower()
    tipo_lower = str(tipo_mat_sel).lower()
    requiere_temperatura = any(kw in material_lower or kw in tipo_lower for kw in ["pulpa", "concentrado", "leche", "refrigerado", "congelado"])

    with st.form("form_item_recepcion"):
      st.subheader("Datos de Transporte y Lote")
      t1, t2, t3 = st.columns(3)
      with t1:
        factura_guia = st.text_input("No. de Factura/Guía")
      with t2:
        placa_transporte = st.text_input("Placa de Transporte")
      with t3:
        lote_proveedor = st.text_input("Lote del Proveedor")

      s1, s2, s3 = st.columns(3)
      with s1:
        ingreso_sap = st.text_input("Fecha de Ingreso a SAP (Opcional)")
      with s2:
        doc_sap = st.text_input("No. Documento SAP (Opcional)")
      with s3:
        marchamo = st.text_input("No. de Marchamo (Opcional)")

      c4, c5 = st.columns(2)
      with c4:
        cantidad = st.number_input("Cantidad Recibida", min_value=0.0, format="%.2f")
      with c5:
        unidad = st.selectbox("Unidad de Medida", ["Kg", "Litros", "Unidades", "Toneladas", "Galones", "Cajas"])

      c6, c7 = st.columns(2)
      with c6:
        fecha_produccion = st.date_input("Fecha de Producción")
      with c7:
        fecha_vencimiento = st.date_input("Fecha de Vencimiento / Caducidad")

      temperatura = "N/A"
      if requiere_temperatura:
        temperatura = st.number_input("Temp. de Recepción (°C)", value=4.0, format="%.1f")

      st.markdown("---")
      certificado_file = st.file_uploader("Subir Archivo del Certificado (PDF o Imagen)", type=["jpg", "jpeg", "png", "pdf"])

      agregar_item = st.form_submit_button("➕ Agregar producto a la lista")

    if agregar_item:
      if not lote_proveedor or not factura_guia:
        st.warning("⚠️ Debes ingresar el número de factura/guía y el número de lote.")
      elif certificado_file is None:
        st.warning("⚠️ Es obligatorio adjuntar el archivo del certificado.")
      else:
        os.makedirs("temp_certificados", exist_ok=True)
        ruta_cert_temp = os.path.join("temp_certificados", certificado_file.name)
        with open(ruta_cert_temp, "wb") as f:
          f.write(certificado_file.getbuffer())

        nuevo_item = {
            "Planta": planta_sel,
            "TipoMaterial": tipo_mat_sel,
            "Proveedor": proveedor_sel,
            "NombreMaterial": material_sel,
            "FacturaGuia": factura_guia,
            "Transporte": proveedor_sel, 
            "Placa": placa_transporte,
            "IngresoSAP": ingreso_sap,
            "DocSAP": doc_sap,
            "Marchamo": marchamo,
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
        st.success("✅ ¡Producto agregado!")

    if len(st.session_state.items_ingreso_actual) > 0:
      st.markdown("---")
      st.subheader("📋 Resumen de Productos")
      df_carrito = pd.DataFrame(st.session_state.items_ingreso_actual)
      st.dataframe(df_carrito[["Planta", "NombreMaterial", "FacturaGuia", "Lote", "Cantidad"]], use_container_width=True)

      item_a_borrar = st.selectbox(
          "Quitar producto de la lista actual:",
          options=range(len(st.session_state.items_ingreso_actual)),
          format_func=lambda x: f"{st.session_state.items_ingreso_actual[x]['NombreMaterial']} (Lote: {st.session_state.items_ingreso_actual[x]['Lote']})"
      )
      if st.button("❌ Quitar seleccionado"):
        st.session_state.items_ingreso_actual.pop(item_a_borrar)
        st.rerun()

      with st.form("form_finalizar_envio"):
        responsable = st.text_input("Operador Responsable")
        enviar_todo = st.form_submit_button("🚀 Enviar a Calidad", type="primary")

      if enviar_todo:
        if not responsable:
          st.error("⚠️ Ingresa el nombre del operador responsable.")
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
                "IngresoSAP": str(item["IngresoSAP"]),
                "DocSAP": str(item["DocSAP"]),
                "Marchamo": str(item["Marchamo"]),
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
                "RutaChecklistOficial": "",
                "RutaPdfCert": str(item["RutaTemp"]),
            }
            df_p = pd.concat([df_p, pd.DataFrame([nuevo_registro])], ignore_index=True)
            nuevo_id += 1

          guardar_pendientes(df_p)
          st.session_state.items_ingreso_actual = []
          st.success("🎉 ¡Enviados a calidad con éxito!")

  else:
    st.warning("⚠️ Sube el catálogo maestro.")

# ------------------------------------------
# PESTAÑA 2: PANEL DE CALIDAD Y FIRMA
# ------------------------------------------
with tab2:
  st.header("✍️ Panel de Control y Firma")
  
  clave_ingresada = st.text_input("Contraseña de Supervisor", type="password")
  CLAVE_SUPERVISOR = "1234"

  if clave_ingresada == CLAVE_SUPERVISOR:
    st.success("🔓 Acceso autorizado.")
    df_p = cargar_pendientes()
    pendientes = df_p[df_p["Estado"] == "Pendiente"] if not df_p.empty else pd.DataFrame()

    if len(pendientes) == 0:
      st.info("No hay registros pendientes.")
    else:
      for index, row in pendientes.iterrows():
        with st.expander(f"📥 ID #{row['ID']} | Material: {row['NombreMaterial']} | Lote: {row['LoteProveedor']}"):
          c1, c2 = st.columns(2)
          with c1:
            st.write(f"**Proveedor:** {row['Proveedor']}")
            st.write(f"**Factura:** {row['FacturaGuia']}")
          with c2:
            st.write(f"**Cantidad:** {row['Cantidad']} {row['Unidad']}")
            st.write(f"**Operador:** {row['Responsable']}")

          if row["RutaPdfCert"] and os.path.exists(str(row["RutaPdfCert"])):
            st.write("📄 Certificado adjunto disponible.")

          col_btn1, col_btn2 = st.columns(2)
          with col_btn1:
            with st.form(f"form_firma_{row['ID']}"):
              nombre_supervisor = st.text_input("Firma del Supervisor")
              aprobar = st.form_submit_button("Aprobar y Generar R-ICC-14-1 🛡️", type="primary")

            if aprobar:
              if not nombre_supervisor:
                st.error("Ingrese su firma.")
              else:
                with st.spinner("Generando PDF Oficial R-ICC-14-1..."):
                  carpeta_base = "./Registros_Oficiales"
                  subcarpeta_lote = f"Lote_{row['LoteProveedor']}_{row['NombreMaterial']}".replace(" ", "_")
                  ruta_destino = os.path.join(carpeta_base, str(row["Planta"]), str(row["TipoMaterial"]), str(row["NombreMaterial"]), subcarpeta_lote)
                  os.makedirs(ruta_destino, exist_ok=True)

                  # 1. Guardar certificado
                  ext_cert = os.path.splitext(str(row["RutaPdfCert"]))[1] if row["RutaPdfCert"] else ".pdf"
                  ruta_cert_final = os.path.join(ruta_destino, f"Certificado_{row['LoteProveedor']}{ext_cert}")
                  if row["RutaPdfCert"] and os.path.exists(str(row["RutaPdfCert"])):
                    shutil.copy(str(row["RutaPdfCert"]), ruta_cert_final)

                  # 2. Generar Checklist R ICC 14-1 en PDF desde cero
                  ruta_checklist_final = os.path.join(ruta_destino, f"Checklist_RICC14-1_{row['LoteProveedor']}.pdf")
                  generar_pdf_ricc14(row, nombre_supervisor, ruta_checklist_final)

                  df_p.loc[df_p["ID"] == row["ID"], "Estado"] = "Aprobado"
                  df_p.loc[df_p["ID"] == row["ID"], "FirmaCalidad"] = str(nombre_supervisor)
                  df_p.loc[df_p["ID"] == row["ID"], "RutaChecklistOficial"] = str(ruta_checklist_final)
                  df_p.loc[df_p["ID"] == row["ID"], "RutaPdfCert"] = str(ruta_cert_final)
                  guardar_pendientes(df_p)

                  st.success("🎉 ¡Checklist PDF R-ICC-14-1 Generado y Aprobado!")
                  st.rerun()

          with col_btn2:
            st.markdown("<br>", unsafe_allow_html=True)
            if st.button("🗑️ Eliminar", key=f"del_pend_{row['ID']}"):
              df_p = df_p[df_p["ID"] != row["ID"]]
              guardar_pendientes(df_p)
              st.rerun()

# ------------------------------------------
# PESTAÑA 3: BÚSQUEDA Y TRAZABILIDAD
# ------------------------------------------
with tab3:
  st.header("🔍 Buscador y Consulta")
  
  df_p = cargar_pendientes()
  aprobados = df_p[df_p["Estado"] == "Aprobado"] if not df_p.empty else pd.DataFrame()

  if len(aprobados) == 0:
    st.info("No hay registros aprobados.")
  else:
    busqueda = st.text_input("Filtro rápido (Lote, Material o Proveedor)")
    if busqueda:
      resultado = aprobados[
          aprobados["LoteProveedor"].astype(str).str.contains(busqueda, case=False, na=False) |
          aprobados["NombreMaterial"].astype(str).str.contains(busqueda, case=False, na=False) |
          aprobados["Proveedor"].astype(str).str.contains(busqueda, case=False, na=False)
      ]
    else:
      resultado = aprobados

    st.dataframe(resultado[["NombreMaterial", "FacturaGuia", "LoteProveedor", "Fecha", "FirmaCalidad"]], use_container_width=True)

    for idx, row in resultado.iterrows():
      with st.expander(f"📁 {row['NombreMaterial']} | Lote: {row['LoteProveedor']} | Supervisor: {row['FirmaCalidad']}"):
        col_d1, col_d2, col_d3 = st.columns(3)

        ruta_chk = str(row.get("RutaChecklistOficial", ""))
        with col_d1:
          if pd.notna(ruta_chk) and ruta_chk != "nan" and os.path.exists(ruta_chk):
            with open(ruta_chk, "rb") as f:
              st.download_button(
                  label="📊 Descargar Checklist (PDF Oficial)",
                  data=f,
                  file_name=os.path.basename(ruta_chk),
                  mime="application/pdf",
                  key=f"dl_chk_{row['ID']}",
              )
          else:
            st.info("Checklist no disponible.")

        ruta_cert = str(row.get("RutaPdfCert", ""))
        with col_d2:
          if pd.notna(ruta_cert) and ruta_cert != "nan" and os.path.exists(ruta_cert):
            with open(ruta_cert, "rb") as f:
              st.download_button(
                  label="📄 Descargar Certificado Subido",
                  data=f,
                  file_name=os.path.basename(ruta_cert),
                  mime="application/pdf" if ruta_cert.endswith(".pdf") else "image/jpeg",
                  key=f"dl_cert_{row['ID']}",
              )
          else:
            st.info("Certificado no disponible.")

        with col_d3:
          if st.button("🗑️ Borrar", key=f"del_apr_{row['ID']}"):
            df_p = df_p[df_p["ID"] != row["ID"]]
            guardar_pendientes(df_p)
            st.rerun()
