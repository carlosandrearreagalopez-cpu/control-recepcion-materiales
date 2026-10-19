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

def cargar_pendientes():
  if os.path.exists(ARCHIVO_PENDIENTES):
    df = pd.read_csv(ARCHIVO_PENDIENTES)
    columnas_texto = [
        "Planta", "TipoMaterial", "NombreMaterial", "Proveedor",
        "LoteProveedor", "FacturaGuia", "Transporte", "Placa",
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
        "LoteProveedor", "FacturaGuia", "Transporte", "Placa", "Cantidad",
        "Unidad", "Temperatura", "FechaProduccion", "FechaVencimiento",
        "Criticidad", "Responsable", "Fecha", "Estado", "FirmaCalidad",
        "RutaCarpeta", "RutaChecklistOficial", "RutaPdfCert"
    ])

def guardar_pendientes(df):
  df.to_csv(ARCHIVO_PENDIENTES, index=False)

def generar_pdf_checklist(row, supervisor, ruta_guardado):
    # Crear PDF en formato Horizontal (Landscape)
    pdf = FPDF(orientation='L', unit='mm', format='A4')
    pdf.add_page()
    
    # --- ENCABEZADO ---
    pdf.set_font('Arial', 'B', 12)
    pdf.cell(277, 8, 'Check list de recepción de materiales', border=1, align='C', ln=1)
    pdf.set_font('Arial', '', 9)
    pdf.cell(92, 6, 'Código: R ICC-14-1', border=1, align='C')
    pdf.cell(93, 6, 'Aseguramiento de Calidad', border=1, align='C')
    pdf.cell(92, 6, 'Versión: 5', border=1, align='C', ln=1)
    pdf.ln(3)

    # --- 1. DATOS GENERALES ---
    pdf.set_fill_color(220, 230, 241)
    pdf.set_font('Arial', 'B', 9)
    pdf.cell(277, 6, '1. DATOS GENERALES Y TRANSPORTE', border=1, align='L', fill=True, ln=1)
    
    pdf.set_font('Arial', 'B', 8)
    pdf.cell(40, 6, 'Fecha de recepción:', border=1)
    pdf.set_font('Arial', '', 8)
    pdf.cell(50, 6, str(row['Fecha']).split()[0], border=1)
    
    pdf.set_font('Arial', 'B', 8)
    pdf.cell(45, 6, 'No. Factura/Guía:', border=1)
    pdf.set_font('Arial', '', 8)
    pdf.cell(45, 6, str(row['FacturaGuia']), border=1)
    
    pdf.set_font('Arial', 'B', 8)
    pdf.cell(45, 6, 'Planta:', border=1)
    pdf.set_font('Arial', '', 8)
    pdf.cell(52, 6, str(row['Planta']), border=1, ln=1)

    pdf.set_font('Arial', 'B', 8)
    pdf.cell(40, 6, 'Nombre Transporte:', border=1)
    pdf.set_font('Arial', '', 8)
    pdf.cell(50, 6, str(row.get('Transporte', row['Proveedor'])), border=1)
    
    pdf.set_font('Arial', 'B', 8)
    pdf.cell(45, 6, 'Placa del transporte:', border=1)
    pdf.set_font('Arial', '', 8)
    pdf.cell(45, 6, str(row.get('Placa', '')), border=1)
    
    pdf.set_font('Arial', 'B', 8)
    pdf.cell(45, 6, 'No. de Marchamo:', border=1)
    pdf.set_font('Arial', '', 8)
    pdf.cell(52, 6, 'N/A', border=1, ln=1)

    # --- 2. DATOS DEL MATERIAL ---
    pdf.set_font('Arial', 'B', 9)
    pdf.cell(277, 6, '2. DATOS DEL MATERIAL', border=1, align='L', fill=True, ln=1)
    
    pdf.set_font('Arial', 'B', 8)
    pdf.cell(30, 6, 'Proveedor:', border=1)
    pdf.set_font('Arial', '', 8)
    pdf.cell(60, 6, str(row['Proveedor']), border=1)
    
    pdf.set_font('Arial', 'B', 8)
    pdf.cell(30, 6, 'Material:', border=1)
    pdf.set_font('Arial', '', 8)
    pdf.cell(60, 6, str(row['NombreMaterial']), border=1)
    
    pdf.set_font('Arial', 'B', 8)
    pdf.cell(30, 6, 'Cantidad/U:', border=1)
    pdf.set_font('Arial', '', 8)
    pdf.cell(67, 6, f"{row['Cantidad']} {row['Unidad']}", border=1, ln=1)

    pdf.set_font('Arial', 'B', 8)
    pdf.cell(30, 6, 'Lote:', border=1)
    pdf.set_font('Arial', '', 8)
    pdf.cell(60, 6, str(row['LoteProveedor']), border=1)
    
    pdf.set_font('Arial', 'B', 8)
    pdf.cell(30, 6, 'F. Producción:', border=1)
    pdf.set_font('Arial', '', 8)
    pdf.cell(60, 6, str(row['FechaProduccion']), border=1)
    
    pdf.set_font('Arial', 'B', 8)
    pdf.cell(30, 6, 'F. Vencimiento:', border=1)
    pdf.set_font('Arial', '', 8)
    pdf.cell(67, 6, str(row['FechaVencimiento']), border=1, ln=1)

    pdf.set_font('Arial', 'B', 8)
    pdf.cell(30, 6, 'Temp. (°C):', border=1)
    pdf.set_font('Arial', '', 8)
    temp_val = str(row['Temperatura']) if str(row['Temperatura']) not in ["N/A", "nan", ""] else "N/A"
    pdf.cell(60, 6, temp_val, border=1)
    pdf.set_font('Arial', 'B', 8)
    pdf.cell(30, 6, 'Criticidad:', border=1)
    pdf.set_font('Arial', '', 8)
    pdf.cell(157, 6, str(row['Criticidad']), border=1, ln=1)

    pdf.ln(3)

    # --- 3. PUNTOS DE INSPECCIÓN (CHECKLIST EXACTO) ---
    pdf.set_font('Arial', 'B', 8)
    pdf.cell(217, 6, '3. INSPECCIÓN DE MATERIAL Y TRANSPORTE', border=1, align='C', fill=True)
    pdf.cell(20, 6, 'SI', border=1, align='C', fill=True)
    pdf.cell(20, 6, 'NO', border=1, align='C', fill=True)
    pdf.cell(20, 6, 'N/A', border=1, align='C', fill=True, ln=1)
    
    preguntas = [
        ("1. ¿El marchamo se encuentra alterado?", "NO"),
        ("2. ¿El candado se encuentra alterado?", "NO"),
        ("3. ¿El (los) lote(s) posee(n) más de un mes de vida?", "SI"),
        ("4. ¿El Certificado de Análisis corresponde al número de lote?", "SI"),
        ("5. ¿Los resultados del Certificado cumplen con los parámetros?", "SI"),
        ("6. ¿El empaque del material se encuentra íntegro y sin deterioro?", "SI"),
        ("7. ¿El material está protegido de la intemperie?", "SI"),
        ("8. ¿Coincide la entrega según documento contra lo físico?", "SI"),
        ("9. ¿Se encuentra el transporte libre de plagas?", "SI"),
        ("10. ¿El techo, paredes o lona se encuentran íntegras y sin agujeros?", "SI"),
        ("11. ¿El techo, paredes o lona se encuentran limpios?", "SI"),
        ("12. ¿El piso del transporte está limpio y en buenas condiciones?", "SI"),
        ("13. ¿El transporte cumple con la temperatura especificada?", "SI" if temp_val != "N/A" else "N/A"),
        ("14. ¿El transporte se encuentra libre de olores extraños?", "NO"),
        ("15. ¿La tarima retornable se encuentra en buenas condiciones?", "SI")
    ]
    
    pdf.set_font('Arial', '', 8)
    for preg, resp_ideal in preguntas:
        pdf.cell(217, 5, preg, border=1)
        pdf.cell(20, 5, 'X' if resp_ideal == 'SI' else '', border=1, align='C')
        pdf.cell(20, 5, 'X' if resp_ideal == 'NO' else '', border=1, align='C')
        pdf.cell(20, 5, 'X' if resp_ideal == 'N/A' else '', border=1, align='C', ln=1)
        
    pdf.ln(3)
    
    # --- RESULTADO Y FIRMAS ---
    pdf.set_font('Arial', 'B', 9)
    pdf.cell(40, 8, 'Resultado Final:', border=1)
    pdf.set_text_color(0, 128, 0)
    pdf.cell(237, 8, 'APROBADO (Validado Digitalmente)', border=1, align='C', ln=1)
    pdf.set_text_color(0, 0, 0)
    
    pdf.ln(6)
    pdf.cell(138, 5, 'Nombre de la persona que realiza la recepción:', 0, 0, 'L')
    pdf.cell(139, 5, 'Vo. Bo. Calidad:', 0, 1, 'L')
    
    pdf.set_font('Arial', '', 9)
    pdf.cell(138, 5, str(row['Responsable']), 0, 0, 'L')
    pdf.cell(139, 5, str(supervisor), 0, 1, 'L')
    
    pdf.output(ruta_guardado)


# Inicializar variables
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
archivo_catalogo_nuevo = st.sidebar.file_uploader("Actualizar Catálogo Excel", type=["xlsx"])
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
  st.markdown("Complete los datos de factura, placa, lote y agregue los productos.")

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
    material_sel = st.selectbox("Nombre del Material (Producto)", materiales_disponibles)

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
      st.subheader("2. Datos de Factura, Transporte y Lote")
      t1, t2 = st.columns(2)
      with t1:
        factura_guia = st.text_input("No. de Factura y/o Guía", placeholder="Ej. FAC-12345")
      with t2:
        placa_transporte = st.text_input("Placa del Transporte", placeholder="Ej. C-123ABC")

      c3, c4, c5 = st.columns(3)
      with c3:
        lote_proveedor = st.text_input("Lote del Proveedor", placeholder="Ej. LOTE-98765")
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
        st.markdown("---")
        st.subheader("❄️ Control de Temperatura (Requerido)")
        temperatura = st.number_input("Temperatura de Recepción (°C)", value=4.0, format="%.1f")

      st.markdown("---")
      st.subheader("📎 Certificado de Calidad del Producto")
      certificado_file = st.file_uploader(
          "Subir Archivo del Certificado (PDF o Imagen)", type=["jpg", "jpeg", "png", "pdf"]
      )

      agregar_item = st.form_submit_button("➕ Agregar este producto a la lista de ingreso")

    if agregar_item:
      if not lote_proveedor or not factura_guia:
        st.warning("⚠️ Debes ingresar el número de factura/guía y el número de lote.")
      elif certificado_file is None:
        st.warning("⚠️ Es obligatorio adjuntar el archivo del certificado para este producto.")
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
            "Transporte": proveedor_sel, # Transporte es igual al proveedor
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
        st.success(f"✅ ¡Producto '{material_sel}' agregado con su certificado a la lista!")

    if len(st.session_state.items_ingreso_actual) > 0:
      st.markdown("---")
      st.subheader("📋 Resumen de Productos Agregados en este Ingreso")
      df_carrito = pd.DataFrame(st.session_state.items_ingreso_actual)
      st.dataframe(df_carrito[["Planta", "Proveedor", "NombreMaterial", "FacturaGuia", "Lote", "Cantidad", "Unidad"]], use_container_width=True)

      item_a_borrar = st.selectbox(
          "Seleccione un producto para quitar de la lista actual",
          options=range(len(st.session_state.items_ingreso_actual)),
          format_func=lambda x: f"{st.session_state.items_ingreso_actual[x]['NombreMaterial']} (Lote: {st.session_state.items_ingreso_actual[x]['Lote']})"
      )
      if st.button("❌ Quitar producto seleccionado"):
        st.session_state.items_ingreso_actual.pop(item_a_borrar)
        st.rerun()

      if st.button("🗑️ Limpiar toda la lista"):
        st.session_state.items_ingreso_actual = []
        st.rerun()

      st.markdown("---")
      st.subheader("3. Finalizar y Enviar a Revisión de Calidad")
      with st.form("form_finalizar_envio"):
        responsable = st.text_input("Operador Responsable", placeholder="Tu Nombre Completo")
        enviar_todo = st.form_submit_button("🚀 Enviar Todos los Productos a Calidad", type="primary")

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
                "RutaChecklistOficial": "",
                "RutaPdfCert": str(item["RutaTemp"]),
            }
            df_p = pd.concat([df_p, pd.DataFrame([nuevo_registro])], ignore_index=True)
            nuevo_id += 1

          guardar_pendientes(df_p)
          st.session_state.items_ingreso_actual = []
          st.success("🎉 ¡Todos los productos con sus certificados fueron enviados con éxito al panel de calidad!")

  else:
    st.warning("⚠️ No se detectó automáticamente el catálogo en el repositorio.")

# ------------------------------------------
# PESTAÑA 2: PANEL DE CALIDAD Y FIRMA
# ------------------------------------------
with tab2:
  st.header("✍️ Panel de Control y Firma de Calidad")
  st.markdown("Ingrese su contraseña, revise los datos, apruebe para generar el checklist final en PDF o elimine registros incorrectos.")

  clave_ingresada = st.text_input("Contraseña de Supervisor de Calidad", type="password", key="pass_calidad")
  CLAVE_SUPERVISOR = "1234"

  if clave_ingresada == CLAVE_SUPERVISOR:
    st.success("🔓 Acceso de supervisor autorizado.")
    df_p = cargar_pendientes()
    pendientes = df_p[df_p["Estado"] == "Pendiente"] if not df_p.empty else pd.DataFrame()

    if len(pendientes) == 0:
      st.info("No hay registros pendientes de aprobación en este momento.")
    else:
      for index, row in pendientes.iterrows():
        with st.expander(f"📥 Recepción ID #{row['ID']} | Material: {row['NombreMaterial']} | Lote: {row['LoteProveedor']}"):
          c1, c2 = st.columns(2)
          with c1:
            st.write(f"**Planta:** {row['Planta']}")
            st.write(f"**Proveedor (Transporte):** {row['Proveedor']}")
            st.write(f"**Factura / Guía:** {row['FacturaGuia']}")
            st.write(f"**Placa de Transporte:** {row['Placa']}")
          with c2:
            st.write(f"**Cantidad:** {row['Cantidad']} {row['Unidad']}")
            st.write(f"**Prod. / Venc.:** {row['FechaProduccion']} al {row['FechaVencimiento']}")
            st.write(f"**Operador:** {row['Responsable']}")

          if row["RutaPdfCert"] and os.path.exists(str(row["RutaPdfCert"])):
            if str(row["RutaPdfCert"]).endswith((".jpg", ".jpeg", ".png")):
              st.image(row["RutaPdfCert"], caption="Certificado Adjunto", width=350)
            else:
              st.write("📄 Certificado PDF adjunto disponible.")

          col_btn1, col_btn2 = st.columns(2)
          
          with col_btn1:
            with st.form(f"form_firma_{row['ID']}"):
              nombre_supervisor = st.text_input("Nombre del Supervisor (Firma Digital)", key=f"sup_{row['ID']}")
              aprobar = st.form_submit_button("Aprobar y Generar Checklist Horizontal PDF 🛡️", type="primary")

            if aprobar:
              if not nombre_supervisor:
                st.error("Debe ingresar su nombre para firmar.")
              else:
                with st.spinner("Creando el formato oficial horizontal..."):
                  carpeta_base = "./Registros_Oficiales"
                  subcarpeta_lote = f"Lote_{row['LoteProveedor']}_{row['NombreMaterial']}".replace(" ", "_")
                  ruta_destino = os.path.join(carpeta_base, str(row["Planta"]), str(row["TipoMaterial"]), str(row["NombreMaterial"]), subcarpeta_lote)
                  os.makedirs(ruta_destino, exist_ok=True)

                  # 1. Guardar certificado original del proveedor
                  ext_cert = os.path.splitext(str(row["RutaPdfCert"]))[1] if row["RutaPdfCert"] else ".pdf"
                  nombre_cert_oficial = f"Certificado_Original_{row['LoteProveedor']}{ext_cert}"
                  ruta_cert_final = os.path.join(ruta_destino, nombre_cert_oficial)

                  if row["RutaPdfCert"] and os.path.exists(str(row["RutaPdfCert"])):
                    shutil.copy(str(row["RutaPdfCert"]), ruta_cert_final)

                  # 2. Generar el nuevo Checklist completo desde cero en PDF Horizontal
                  nombre_checklist_pdf = f"Checklist_Generado_{row['LoteProveedor']}.pdf"
                  ruta_checklist_final = os.path.join(ruta_destino, nombre_checklist_pdf)
                  
                  generar_pdf_checklist(row, nombre_supervisor, ruta_checklist_final)

                  df_p.loc[df_p["ID"] == row["ID"], "Estado"] = "Aprobado"
                  df_p.loc[df_p["ID"] == row["ID"], "FirmaCalidad"] = str(nombre_supervisor)
                  df_p.loc[df_p["ID"] == row["ID"], "RutaCarpeta"] = str(ruta_destino)
                  df_p.loc[df_p["ID"] == row["ID"], "RutaChecklistOficial"] = str(ruta_checklist_final)
                  df_p.loc[df_p["ID"] == row["ID"], "RutaPdfCert"] = str(ruta_cert_final)
                  guardar_pendientes(df_p)

                  st.success("🎉 ¡Aprobado! El Checklist completo ha sido generado en PDF.")
                  st.rerun()

          with col_btn2:
            st.markdown("<br>", unsafe_allow_html=True)
            if st.button("🗑️ Eliminar Registro Pendiente", key=f"del_pend_{row['ID']}"):
              df_p = df_p[df_p["ID"] != row["ID"]]
              guardar_pendientes(df_p)
              st.warning("⚠️ Registro pendiente eliminado.")
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
  st.markdown("Consulte el historial de recepciones aprobadas, descargue los checklists generados o elimine registros.")

  df_p = cargar_pendientes()
  aprobados = df_p[df_p["Estado"] == "Aprobado"] if not df_p.empty else pd.DataFrame()

  if len(aprobados) == 0:
    st.info("Aún no hay registros aprobados y archivados en el sistema.")
  else:
    busqueda = st.text_input("Filtro de búsqueda rápida", placeholder="Escriba el Lote, Material o Proveedor...")

    if busqueda:
      resultado = aprobados[
          aprobados["LoteProveedor"].astype(str).str.contains(busqueda, case=False, na=False) |
          aprobados["NombreMaterial"].astype(str).str.contains(busqueda, case=False, na=False) |
          aprobados["Proveedor"].astype(str).str.contains(busqueda, case=False, na=False)
      ]
    else:
      resultado = aprobados

    st.dataframe(resultado[["Planta", "NombreMaterial", "Proveedor", "FacturaGuia", "Placa", "LoteProveedor", "Cantidad", "Unidad", "FechaProduccion", "FirmaCalidad"]], use_container_width=True)

    st.markdown("---")
    st.subheader("📥 Descarga de Archivos Oficiales y Gestión")

    for idx, row in resultado.iterrows():
      with st.expander(f"📁 Material: {row['NombreMaterial']} | Lote: {row['LoteProveedor']} | Aprobado por: {row['FirmaCalidad']}"):
        col_d1, col_d2, col_d3 = st.columns(3)

        with col_d1:
          if "RutaChecklistOficial" in row and pd.notna(row["RutaChecklistOficial"]) and str(row["RutaChecklistOficial"]).strip() != "" and os.path.exists(str(row["RutaChecklistOficial"])):
            with open(row["RutaChecklistOficial"], "rb") as f:
              st.download_button(
                  label="📊 Descargar Checklist Oficial (PDF)",
                  data=f,
                  file_name=os.path.basename(str(row["RutaChecklistOficial"])),
                  mime="application/pdf",
                  key=f"dl_chk_{row['ID']}",
              )
          else:
            st.info("Checklist no disponible.")

        with col_d2:
          if "RutaPdfCert" in row and pd.notna(row["RutaPdfCert"]) and str(row["RutaPdfCert"]).strip() != "" and os.path.exists(str(row["RutaPdfCert"])):
            with open(row["RutaPdfCert"], "rb") as f:
              st.download_button(
                  label="📄 Descargar Certificado del Proveedor",
                  data=f,
                  file_name=os.path.basename(str(row["RutaPdfCert"])),
                  mime="application/pdf" if str(row["RutaPdfCert"]).endswith(".pdf") else "image/jpeg",
                  key=f"dl_pdf_{row['ID']}",
              )
          else:
            st.info("Certificado no disponible.")

        with col_d3:
          if st.button("🗑️ Borrar del Historial", key=f"del_apr_{row['ID']}"):
            df_p = df_p[df_p["ID"] != row["ID"]]
            guardar_pendientes(df_p)
            st.warning("⚠️ Registro eliminado del historial oficial.")
            st.rerun()
