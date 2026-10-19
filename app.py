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
    return pd.read_csv(ARCHIVO_PENDIENTES)
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
    ])


def guardar_pendientes(df):
  df.to_csv(ARCHIVO_PENDIENTES, index=False)


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
      "Complete la información correspondiente al checklist de ingreso y"
      " adjunte el certificado oficial."
  )

  if df_db is not None:
    col_planta = "Planta"
    col_tipo = "Tipo de material"
    col_material = "Material"
    col_proveedor = "Proveedor"
    col_criticidad = (
        "Criticidad" if "Criticidad" in df_db.columns else None
    )

    with st.form("form_operador", clear_on_submit=False):
      st.subheader("1. Jerarquía y Selección de Material")

      # Nivel 1: Selección de Planta
      plantas = df_db[col_planta].dropna().unique()
      planta_sel = st.selectbox("Planta", plantas)

      # Filtrar estrictamente por la Planta seleccionada
      df_planta = df_db[df_db[col_planta] == planta_sel]

      col1, col2 = st.columns(2)

      with col1:
        # Nivel 2: Tipo de Material exclusivo de la planta seleccionada
        tipos_material = df_planta[col_tipo].dropna().unique()
        tipo_mat_sel = st.selectbox("Tipo de Material", tipos_material)

        # Filtrar por Planta y Tipo de Material
        df_tipo = df_planta[df_planta[col_tipo] == tipo_mat_sel]

      with col2:
        # Nivel 3: Nombre del Material exclusivo de la planta y tipo seleccionados
        materiales = df_tipo[col_material].dropna().unique()
        material_sel = st.selectbox("Nombre del Material", materiales)

        # Filtrar por Planta, Tipo y Material seleccionado para obtener sus proveedores
        df_material = df_tipo[df_tipo[col_material] == material_sel]
        proveedores_disponibles = df_material[col_proveedor].dropna().unique()
        proveedor_sel = st.selectbox("Proveedor", proveedores_disponibles)

      # Obtener criticidad automática del material seleccionado
      criticidad_val = "N/A"
      if col_criticidad and not df_material.empty:
        criticidad_val = str(df_material.iloc[0][col_criticidad])

      st.info(f"🏷️ **Criticidad Asociada al Material:** {criticidad_val}")

      st.markdown("---")
      st.subheader("2. Datos de Lote, Cantidad y Fechas")
      col3, col4, col5 = st.columns(3)
      with col3:
        lote_proveedor = st.text_input(
            "Lote del Proveedor", placeholder="Ej. LOTE-98765"
        )
      with col4:
        cantidad = st.number_input(
            "Cantidad Recibida", min_value=0.0, format="%.2f"
        )
      with col5:
        unidad = st.selectbox(
            "Unidad de Medida",
            ["Kg", "Litros", "Unidades", "Toneladas", "Galones", "Cajas"],
        )

      col6, col7 = st.columns(2)
      with col6:
        fecha_produccion = st.date_input("Fecha de Producción")
      with col7:
        fecha_vencimiento = st.date_input("Fecha de Vencimiento / Caducidad")

      # Control de Temperatura condicional si es materia prima, concentrado o pulpa
      temperatura = "N/A"
      if (
          "materia prima" in str(tipo_mat_sel).lower()
          or "concentrado" in str(material_sel).lower()
          or "pulpa" in str(material_sel).lower()
      ):
        st.markdown("---")
        st.subheader("❄️ Control de Temperatura (Requerido para Materia Prima)")
        temperatura = st.number_input(
            "Temperatura de Recepción (°C)",
            value=4.0,
            format="%.1f",
            help=(
                "Ingrese la temperatura medida al recibir el concentrado o"
                " pulpa."
            ),
        )

      st.markdown("---")
      st.subheader("3. Responsable y Documentación")
      col8, col9 = st.columns(2)
      with col8:
        responsable = st.text_input(
            "Operador Responsable", placeholder="Tu Nombre"
        )
      with col9:
        certificado_file = st.file_uploader(
            "Adjuntar Certificado de Calidad (Imagen o PDF)",
            type=["jpg", "jpeg", "png", "pdf"],
        )

      st.markdown("")
      enviar_revision = st.form_submit_button(
          "Enviar a Revisión de Calidad 🚀", type="primary"
      )

    if enviar_revision:
      if not lote_proveedor or not responsable:
        st.error(
            "⚠️ Por favor, completa el número de lote y el nombre del"
            " operador."
        )
      elif certificado_file is None:
        st.error(
            "⚠️ Es obligatorio adjuntar el certificado de calidad en formato"
            " imagen o PDF."
        )
      else:
        os.makedirs("temp_certificados", exist_ok=True)
        ruta_cert_temp = os.path.join(
            "temp_certificados", certificado_file.name
        )
        with open(ruta_cert_temp, "wb") as f:
          f.write(certificado_file.getbuffer())

        df_p = cargar_pendientes()
        nuevo_id = len(df_p) + 1 if not df_p.empty else 1
        nuevo_registro = {
            "ID": nuevo_id,
            "Planta": planta_sel,
            "TipoMaterial": tipo_mat_sel,
            "NombreMaterial": material_sel,
            "Proveedor": proveedor_sel,
            "LoteProveedor": lote_proveedor,
            "Cantidad": cantidad,
            "Unidad": unidad,
            "Temperatura": temperatura,
            "FechaProduccion": str(fecha_produccion),
            "FechaVencimiento": str(fecha_vencimiento),
            "Criticidad": criticidad_val,
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
            "¡Registro enviado exitosamente al panel de control de calidad"
            " para su validación y firma!"
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
      "Espacio exclusivo para supervisores. Ingrese su contraseña de acceso,"
      " revise los datos, certificado, firme y apruebe."
  )

  clave_ingresada = st.text_input(
      "Contraseña de Supervisor de Calidad", type="password"
  )
  CLAVE_SUPERVISOR = "1234"

  if clave_ingresada == CLAVE_SUPERVISOR:
    st.success("🔓 Acceso de supervisor autorizado.")
    df_p = cargar_pendientes()
    pendientes = (
        df_p[df_p["Estado"] == "Pendiente"] if not df_p.empty else pd.DataFrame()
    )

    if len(pendientes) == 0:
      st.info("No hay registros pendientes de aprobación en este momento.")
    else:
      for index, row in pendientes.iterrows():
        with st.expander(
            f"📥 Recepción #{row['ID']} | Material: {row['NombreMaterial']} |"
            f" Lote: {row['LoteProveedor']}"
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

          if os.path.exists(str(row["RutaCarpeta"])):
            if row["RutaCarpeta"].endswith((".jpg", ".jpeg", ".png")):
              st.image(
                  row["RutaCarpeta"], caption="Certificado Adjunto", width=350
              )
            else:
              st.write(
                  "📄 Archivo PDF del certificado adjunto listo para procesar."
              )

          with st.form(f"form_firma_{row['ID']}"):
            nombre_supervisor = st.text_input(
                "Nombre del Supervisor (Firma Digital)"
            )
            aprobar = st.form_submit_button(
                "Aprobar, Firmar y Generar Estructura Oficial 🛡️",
                type="primary",
            )

          if aprobar:
            if not nombre_supervisor:
              st.error(
                  "Debe ingresar su nombre para autorizar y firmar el registro."
              )
            else:
              with st.spinner("Generando carpetas y archivos oficiales..."):
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

                if os.path.exists(str(row["RutaCarpeta"])):
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
                    f"¡Registro aprobado con éxito! Guardado en la estructura:"
                    f" `{ruta_destino}`"
                )
                st.rerun()
  elif clave_ingresada == "":
    st.info("🔒 Por favor, ingrese la contraseña de supervisor para ver los registros pendientes.")
  else:
    st.error("❌ Contraseña incorrecta. Intente nuevamente.")

# ------------------------------------------
# PESTAÑA 3: BÚSQUEDA Y TRAZABILIDAD
# ------------------------------------------
with tab3:
  st.header("🔍 Buscador y Consulta de Registros Oficiales")
  st.markdown(
      "Consulte el historial de recepciones aprobadas filtrando por lote,"
      " material o proveedor."
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
