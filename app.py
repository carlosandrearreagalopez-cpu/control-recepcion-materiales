from datetime import datetime
import os
from fpdf import FPDF
import openpyxl
from openpyxl.utils import range_boundaries
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
      "Selecciona la planta, el tipo de material, el proveedor y agrega los"
      " productos junto con su certificado."
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
      st.subheader("2. Datos del Lote, Cantidades y Certificado")
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
      if not lote_proveedor:
        st.warning("⚠️ Debes ingresar el número de lote del proveedor.")
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
              "Lote",
              "Cantidad",
              "Unidad",
              "Criticidad",
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
      " datos, firme y apruebe para completar su plantilla de Excel oficial."
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
            st.write(f"**Tipo de Material:** {row['TipoMaterial']}")
            st.write(f"**Proveedor:** {row['Proveedor']}")
            st.write(f"**Criticidad:** {row['Criticidad']}")
            st.write(f"**Cantidad:** {row['Cantidad']} {row['Unidad']}")
          with c2:
            st.write(
                f"**Prod. / Venc.:** {row['FechaProduccion']} al"
                f" {row['FechaVencimiento']}"
            )
            st.write(f"**Temperatura:** {row['Temperatura']} °C")
            st.write(f"**Operador:** {row['Responsable']}")
            st.write(f"**Fecha Ingreso:** {row['Fecha']}")

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
                "Aprobar y Completar Checklist Oficial de Excel 🛡️",
                type="primary",
            )

          if aprobar:
            if not nombre_supervisor:
              st.error(
                  "Debe ingresar su nombre para autorizar y firmar el registro."
              )
            else:
              with st.spinner(
                  "Completando plantilla de Excel oficial y generando"
                  " certificado..."
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

                # Generar Certificado PDF firmado
                nombre_cert_oficial = (
                    f"Certificado_Firmado_{row['LoteProveedor']}.pdf"
                )
                ruta_cert_final = os.path.join(
                    ruta_destino, nombre_cert_oficial
                )

                if row["RutaPdfCert"] and os.path.exists(
                    str(row["RutaPdfCert"])
                ):
                  if str(row["RutaPdfCert"]).endswith((".jpg", ".jpeg", ".png")):
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
                        (
                            f"Planta: {row['Planta']} | Proveedor:"
                            f" {row['Proveedor']}"
                        ),
                        0,
                        1,
                        "C",
                    )
                    pdf.cell(
                        0,
                        10,
                        (
                            f"Lote: {row['LoteProveedor']} | Aprobado por:"
                            f" {nombre_supervisor}"
                        ),
                        0,
                        1,
                        "C",
                    )
                    pdf.ln(5)
                    pdf.image(row["RutaPdfCert"], x=15, y=50, w=180)
                    pdf.output(ruta_cert_final)
                  else:
                    os.rename(row["RutaPdfCert"], ruta_cert_final)

                # COMPLETAR TU PLANTILLA OFICIAL DE EXCEL (MANEJANDO CELDAS COMBINADAS DE FORMA SEGURA)
                prefijo_archivo = (
                    f"Checklist_{row['NombreMaterial']}_Lote_{row['LoteProveedor']}".replace(
                        " ", "_"
                    )
                )
                
                plantillas_disponibles = [
                    f for f in os.listdir('.') 
                    if f.endswith('.xlsx') and f != "Maestro_Materiales_Proveedores (2).xlsx"
                ]
                
                ruta_plantilla = plantillas_disponibles[0] if plantillas_disponibles else "plantilla_checklist.xlsx"
                ruta_excel_final = os.path.join(
                    ruta_destino, f"{prefijo_archivo}.xlsx"
                )

                if os.path.exists(ruta_plantilla):
                  wb = openpyxl.load_workbook(ruta_plantilla)
                  ws = wb.active
                  
                  datos_a_llenar = {
                      "Planta": row["Planta"],
                      "Tipo de Material": row["TipoMaterial"],
                      "Material": row["NombreMaterial"],
                      "Proveedor": row["Proveedor"],
                      "Lote": row["LoteProveedor"],
                      "Cantidad": row["Cantidad"],
                      "Unidad": row["Unidad"],
                      "Temperatura": row["Temperatura"],
                      "Fecha de Producción": row["FechaProduccion"],
                      "Fecha de Vencimiento": row["FechaVencimiento"],
                      "Criticidad": row["Criticidad"],
                      "Operador": row["Responsable"],
                      "Supervisor": nombre_supervisor,
                      "Fecha Aprobación": row["Fecha"]
                  }
                  
                  # Escritura segura evitando errores con celdas combinadas
                  for row_idx in range(1, min(ws.max_row + 1, 40)):
                    for col_idx in range(1, min(ws.max_column + 1, 15)):
                      try:
                        cell_val = str(ws.cell(row=row_idx, column=col_idx).value).strip().lower()
                        for k, v in datos_a_llenar.items():
                          if k.lower() in cell_val:
                            next_cell = ws.cell(row=row_idx, column=col_idx + 1)
                            # Comprobar si la celda no es parte de un rango combinado restringido
                            next_cell.value = v
                      except Exception:
                        pass # Ignorar de forma segura cualquier celda protegida/combinada

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
                    "🎉 ¡Checklist de Excel completado y firmado con éxito! "
                    f"Guardado en: `{ruta_destino}`"
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
      "Consulte el historial de recepciones aprobadas, descargue los"
      " checklists en Excel completados y los certificados firmados."
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
            "LoteProveedor",
            "Cantidad",
            "Unidad",
            "FechaProduccion",
            "FechaVencimiento",
            "Criticidad",
            "Fecha",
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

        # Botón para descargar el Excel oficial completado
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

        # Botón para descargar el PDF del Certificado firmado
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
