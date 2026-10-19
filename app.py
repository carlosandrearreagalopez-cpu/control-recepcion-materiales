from datetime import datetime
import os
from fpdf import FPDF
import openpyxl
import pandas as pd
import streamlit as st

# Configuración de diseño de página limpio y profesional
st.set_page_config(
    page_title="Control de Recepción y Calidad",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Estilos visuales modernos
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
        "Responsable",
        "Fecha",
        "Estado",
        "FirmaCalidad",
        "RutaCarpeta",
    ])


def guardar_pendientes(df):
  df.to_csv(ARCHIVO_PENDIENTES, index=False)


# CARGAR CATÁLOGO AUTOMÁTICAMENTE DESDE EL REPOSITORIO
df_db = None
if os.path.exists(ARCHIVO_CATALOGO):
  try:
    df_db = pd.read_excel(ARCHIVO_CATALOGO)
  except Exception as e:
    st.error(f"Error al leer el catálogo automático: {e}")

# Opción en barra lateral por si deseas actualizarlo manualmente después
st.sidebar.header("⚙️ Panel de Administración")
archivo_catalogo_nuevo = st.sidebar.file_uploader(
    "Actualizar Catálogo Excel", type=["xlsx"]
)
if archivo_catalogo_nuevo is not None:
  df_db = pd.read_excel(archivo_catalogo_nuevo)
  st.sidebar.success("¡Catálogo actualizado correctamente!")

# PESTAÑAS PRINCIPALES
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
    with st.form("form_operador", clear_on_submit=False):
      st.subheader("1. Selección de Material")
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
      st.subheader("2. Datos de Lote y Responsable")
      col3, col4 = st.columns(2)
      with col3:
        lote_proveedor = st.text_input(
            "Lote del Proveedor", placeholder="Ej. LOTE-98765"
        )
      with col4:
        responsable = st.text_input(
            "Operador Responsable", placeholder="Tu Nombre"
        )

      st.markdown("---")
      st.subheader("3. Documentación")
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
            "¡Registro enviado exitosamente al panel de control de calidad para"
            " su validación y firma!"
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
      " revise el certificado, firme y apruebe para generar la jerarquía"
      " oficial de carpetas."
  )

  # Sistema de Clave de Seguridad para el Panel de Calidad
  clave_ingresada = st.text_input(
      "Contraseña de Supervisor de Calidad", type="password"
  )
  CLAVE_SUPERVISOR = (
      "1234"  # Puedes cambiar la contraseña aquí cuando gustes
  )

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
          with c2:
            st.write(f"**Operador:** {row['Responsable']}")
            st.write(f"**Fecha:** {row['Fecha']}")

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
                # Jerarquía exacta: Planta -> TipoMaterial -> NombreMaterial -> Lote_Nombre
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
            "Fecha",
            "FirmaCalidad",
        ]],
        use_container_width=True,
    )
