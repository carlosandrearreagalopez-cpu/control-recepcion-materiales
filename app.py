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
# FUNCIONES DE DATOS
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

# ------------------------------------------
# GENERADOR DEL PDF HORIZONTAL (R ICC-14-1)
# ------------------------------------------
def dibujar_casilla(pdf, texto, x, y, marcado=False):
    pdf.set_xy(x, y)
    pdf.set_font('Arial', '', 8)
    pdf.cell(10, 5, texto, 0, 0, 'R')
    pdf.rect(x + 11, y + 1, 3, 3, style='D' if not marcado else 'DF')

def generar_pdf_ricc14(row, supervisor, ruta_guardado):
    # Formato A4 Horizontal (297 x 210 mm). Márgenes de 10mm -> Ancho útil = 277mm
    pdf = FPDF(orientation='L', unit='mm', format='A4')
    pdf.add_page()
    pdf.set_margins(10, 10, 10)
    pdf.set_auto_page_break(auto=True, margin=10)
    pdf.set_fill_color(220, 230, 241) # Color gris claro para encabezados
    
    # ---------------- ENCABEZADO ----------------
    pdf.set_font('Arial', 'B', 8)
    # Fila 1
    pdf.cell(30, 5, 'Código:', 'LT', 0, 'R')
    pdf.set_font('Arial', '', 8)
    pdf.cell(40, 5, 'R ICC - 14-1', 'TR', 0, 'C')
    pdf.set_font('Arial', 'B', 8)
    pdf.cell(167, 5, 'Aseguramiento de Calidad', 'LTR', 0, 'C')
    pdf.set_font('Arial', '', 8)
    pdf.cell(40, 5, 'Vigente a partir de:', 'LTR', 1, 'C')
    # Fila 2
    pdf.set_font('Arial', 'B', 8)
    pdf.cell(30, 5, 'Versión:', 'LB', 0, 'R')
    pdf.set_font('Arial', '', 8)
    pdf.cell(40, 5, '5', 'BR', 0, 'C')
    pdf.set_font('Arial', 'B', 10)
    pdf.cell(167, 5, 'Check list de recepción de materiales', 'LBR', 0, 'C')
    pdf.set_font('Arial', '', 8)
    pdf.cell(40, 5, 'feb-23', 'LBR', 1, 'C')
    pdf.ln(3)

    # ---------------- DATOS GENERALES ----------------
    pdf.set_font('Arial', 'B', 8)
    
    pdf.cell(40, 6, 'Fecha de recepción:', 1, 0, 'R')
    pdf.set_font('Arial', '', 8)
    pdf.cell(70, 6, str(row.get('Fecha', '')).split()[0], 1, 0, 'C')
    pdf.cell(20, 6, '', 0, 0) # Espacio
    pdf.set_font('Arial', 'B', 8)
    pdf.cell(60, 6, 'Número de factura y/o guía del proveedor:', 1, 0, 'R')
    pdf.set_font('Arial', '', 8)
    pdf.cell(70, 6, str(row.get('FacturaGuia', '')), 1, 1, 'C')

    pdf.set_font('Arial', 'B', 8)
    pdf.cell(40, 6, 'Nombre del transporte:', 1, 0, 'R')
    pdf.set_font('Arial', '', 8)
    pdf.cell(70, 6, str(row.get('Transporte', '')), 1, 0, 'C')
    pdf.cell(20, 6, '', 0, 0)
    pdf.set_font('Arial', 'B', 8)
    pdf.cell(60, 6, 'Fecha de Ingreso a SAP:', 1, 0, 'R')
    pdf.set_font('Arial', '', 8)
    pdf.cell(70, 6, str(row.get('IngresoSAP', '')), 1, 1, 'C')

    pdf.set_font('Arial', 'B', 8)
    pdf.cell(40, 6, 'Placa del transporte:', 1, 0, 'R')
    pdf.set_font('Arial', '', 8)
    pdf.cell(70, 6, str(row.get('Placa', '')), 1, 0, 'C')
    pdf.cell(20, 6, '', 0, 0)
    pdf.set_font('Arial', 'B', 8)
    pdf.cell(60, 6, 'Número de documento de ingreso a SAP:', 1, 0, 'R')
    pdf.set_font('Arial', '', 8)
    pdf.cell(70, 6, str(row.get('DocSAP', '')), 1, 1, 'C')
    pdf.ln(3)

    # ---------------- MARCHAMO Y CANDADO ----------------
    pdf.set_font('Arial', 'B', 8)
    pdf.cell(100, 5, 'Inspección del Marchamo', 0, 0, 'L', fill=True)
    pdf.cell(5, 5, '', 0, 0)
    pdf.cell(100, 5, 'Inspección del Candado', 0, 0, 'L', fill=True)
    pdf.cell(5, 5, '', 0, 0)
    pdf.cell(20, 5, 'Valor esperado', 0, 0, 'C')
    pdf.cell(47, 5, 'Acción en caso de desviación', 0, 1, 'C')

    y_start = pdf.get_y()
    pdf.rect(10, y_start, 100, 10)
    pdf.rect(115, y_start, 100, 10)
    pdf.rect(220, y_start, 20, 10)
    pdf.rect(240, y_start, 47, 10)

    # Q1 & Q2
    pdf.set_xy(10, y_start + 1)
    pdf.cell(5, 8, '1', 0, 0, 'C')
    pdf.cell(50, 4, '¿El marchamo se encuentra alterado?', 0, 0)
    # Checkboxes Marchamo (Marcamos NO por defecto ideal)
    dibujar_casilla(pdf, "SI", 65, y_start + 1)
    dibujar_casilla(pdf, "NO", 80, y_start + 1, marcado=True)
    dibujar_casilla(pdf, "N/A", 95, y_start + 1)
    
    pdf.set_xy(15, y_start + 5)
    pdf.cell(25, 4, 'No. de marchamo:', 0, 0)
    pdf.cell(50, 4, str(row.get('Marchamo', '')), 'B', 0, 'C')

    pdf.set_xy(115, y_start + 2)
    pdf.cell(5, 6, '2', 0, 0, 'C')
    pdf.cell(50, 6, '¿El candado se encuentra alterado?', 0, 0)
    dibujar_casilla(pdf, "SI", 170, y_start + 2)
    dibujar_casilla(pdf, "NO", 185, y_start + 2, marcado=True)
    dibujar_casilla(pdf, "N/A", 200, y_start + 2)

    pdf.set_xy(220, y_start + 2)
    pdf.set_font('Arial', 'B', 8)
    pdf.cell(20, 6, 'NO', 0, 0, 'C')
    pdf.set_font('Arial', '', 7)
    pdf.cell(47, 6, 'Retener el lote y emitir reclamo.', 0, 1, 'C')
    pdf.ln(3)

    # ---------------- INSPECCIÓN DEL MATERIAL ----------------
    pdf.set_font('Arial', 'B', 8)
    pdf.cell(210, 5, 'Inspección del Material', 0, 0, 'L', fill=True)
    pdf.cell(20, 5, 'Valor esperado', 0, 0, 'C', fill=True)
    pdf.cell(47, 5, 'Acción en caso de desviación', 0, 1, 'C', fill=True)

    # Filas de datos de material
    datos_mat = [
        ("Proveedor:", str(row.get('Proveedor', ''))),
        ("Material:", str(row.get('NombreMaterial', ''))),
        ("Cantidad:", str(row.get('Cantidad', ''))),
        ("Unidad de medida:", str(row.get('Unidad', ''))),
        ("Lote:", str(row.get('LoteProveedor', ''))),
        ("Fecha de Producción:", str(row.get('FechaProduccion', ''))),
        ("Fecha de Vencimiento:", str(row.get('FechaVencimiento', '')))
    ]

    for etiqueta, valor in datos_mat:
        pdf.set_font('Arial', 'B', 8)
        pdf.cell(40, 5, etiqueta, 1, 0, 'R')
        pdf.set_font('Arial', '', 8)
        pdf.cell(170, 5, valor, 1, 0, 'C')
        # Celdas grises derechas
        pdf.set_fill_color(200, 200, 200)
        pdf.cell(20, 5, '------', 1, 0, 'C', fill=True)
        pdf.cell(47, 5, '------', 1, 1, 'C', fill=True)

    pdf.set_fill_color(220, 230, 241) # Reset fill
    
    # Preguntas Material
    preguntas_mat = [
        ("3", "Tomando en cuenta la fecha del día de hoy y la respuesta anterior, ¿El (los) lote(s) posee(n) más de un mes de vida?", ["SI", "NO", "N/A"], "SI", "Retener el lote y emitir reclamo"),
        ("4", "¿A qué categoría de la matriz de materiales pertenece el insumo a recibir?", ["I", "II", "III"], "------", "------"),
        ("5", "¿El Certificado de Análisis físico/digital (si aplica), corresponde al número de lote?", ["SI", "NO", "N/A"], "SI", "Rechazar el lote y emitir reclamo"),
        ("6", "¿Los resultados de los Certificados de Análisis cumple(n) con los parámetros críticos de la Especificación del Material a Recibir?", ["SI", "NO", "N/A"], "SI", "Rechazar el lote y emitir reclamo"),
        ("7", "¿El empaque del material se encuentra íntegro y sin deterioro?", ["SI", "NO", "N/A"], "SI", "Rechazar el lote y emitir reclamo"),
        ("8", "¿El material está protegido de la intemperie?", ["SI", "NO", "N/A"], "SI", "Rechazar el lote y emitir reclamo"),
        ("9", "¿Coincide la entrega según documento, contra lo físico en cantidad y forma?", ["SI", "NO", "N/A"], "SI", "Rechazar el lote y emitir reclamo")
    ]

    criticidad = str(row.get('Criticidad', 'I'))

    for i, (num, preg, opts, val_esp, accion) in enumerate(preguntas_mat):
        y_curr = pdf.get_y()
        pdf.set_font('Arial', 'B', 8)
        pdf.cell(5, 6, num, 1, 0, 'C')
        pdf.set_font('Arial', '', 7)
        # Adapt text width
        pdf.cell(100, 6, preg[:85] + ("..." if len(preg)>85 else ""), 1, 0, 'L')
        
        # Bloque para la caja (simulando 1 proveedor)
        pdf.cell(105, 6, '', 1, 0)
        
        # Dibujar casillas
        if num == "4":
            dibujar_casilla(pdf, opts[0], 120, y_curr+1, marcado=(criticidad=="I"))
            dibujar_casilla(pdf, opts[1], 135, y_curr+1, marcado=(criticidad=="II"))
            dibujar_casilla(pdf, opts[2], 150, y_curr+1, marcado=(criticidad=="III"))
        else:
            dibujar_casilla(pdf, opts[0], 120, y_curr+1, marcado=True)
            dibujar_casilla(pdf, opts[1], 135, y_curr+1)
            dibujar_casilla(pdf, opts[2], 150, y_curr+1)

        pdf.set_font('Arial', 'B', 8)
        if val_esp == "------":
            pdf.set_fill_color(200, 200, 200)
            pdf.cell(20, 6, val_esp, 1, 0, 'C', fill=True)
            pdf.cell(47, 6, accion, 1, 1, 'C', fill=True)
            pdf.set_fill_color(220, 230, 241)
        else:
            pdf.cell(20, 6, val_esp, 1, 0, 'C')
            pdf.set_font('Arial', '', 7)
            pdf.cell(47, 6, accion, 1, 1, 'C')

    pdf.ln(2)

    # ---------------- INSPECCIÓN DEL TRANSPORTE ----------------
    pdf.set_font('Arial', 'B', 8)
    pdf.cell(210, 5, 'Inspección del Transporte', 0, 0, 'L', fill=True)
    pdf.cell(67, 5, '', 0, 1)

    preguntas_trans = [
        ("10", "¿Se encuentra el transporte libre de plagas?", "SI", "Rechazar el lote y emitir reclamo"),
        ("11", "¿El techo, las paredes del transporte o lona (cuando aplique) se encuentran íntegras y sin agujeros?", "SI", "Rechazar el lote y emitir reclamo"),
        ("12", "¿El techo, las paredes del transporte o lona (cuando aplique) se encuentran limpios?", "SI", "Rechazar el lote y emitir reclamo"),
        ("13", "¿El piso del transporte está limpio y en buenas condiciones?", "SI", "Emitir reclamo")
    ]

    for num, preg, val_esp, accion in preguntas_trans:
        y_curr = pdf.get_y()
        pdf.set_font('Arial', 'B', 8)
        pdf.cell(5, 5, num, 1, 0, 'C')
        pdf.set_font('Arial', '', 7)
        pdf.cell(100, 5, preg, 1, 0, 'L')
        pdf.cell(105, 5, '', 1, 0)
        
        dibujar_casilla(pdf, "SI", 120, y_curr+0.5, marcado=True)
        dibujar_casilla(pdf, "NO", 135, y_curr+0.5)

        pdf.set_font('Arial', 'B', 8)
        pdf.cell(20, 5, val_esp, 1, 0, 'C')
        pdf.set_font('Arial', '', 7)
        pdf.cell(47, 5, accion, 1, 1, 'C')

    pdf.ln(5)
    
    # ---------------- RESULTADO, OBSERVACIONES Y FIRMAS ----------------
    pdf.set_font('Arial', 'B', 8)
    pdf.cell(277, 4, '** Si el marchamo N/A por ser proveedores locales, basta con la inspección del material y del transporte.', 0, 1, 'L')
    pdf.cell(277, 4, '** Para los casos que se realiza la revisión en el transporte y se determina que éste no cumple y que los insumos se ven DIRECTAMENTE AFECTADOS, SE PROCEDE A RECHAZARLOS.', 0, 1, 'L')
    
    pdf.ln(4)
    pdf.set_font('Arial', 'B', 9)
    pdf.cell(70, 6, 'Nombre de la persona que realiza la recepción:', 0, 0)
    pdf.set_font('Arial', '', 9)
    pdf.cell(80, 6, str(row.get('Responsable', '')), 'B', 0, 'C')
    pdf.cell(10, 6, '', 0, 0)
    pdf.set_font('Arial', 'B', 9)
    pdf.cell(25, 6, 'Vo. Bo. Calidad:', 0, 0)
    pdf.set_font('Arial', '', 9)
    pdf.cell(80, 6, supervisor, 'B', 1, 'C')

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
        ingreso_sap = st.date_input("Fecha de Ingreso a SAP")
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
            "IngresoSAP": str(ingreso_sap),
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
              aprobar = st.form_submit_button("Aprobar y Generar R-ICC-14-1 Horizontal 🛡️", type="primary")

            if aprobar:
              if not nombre_supervisor:
                st.error("Ingrese su firma.")
              else:
                with st.spinner("Generando PDF Oficial R-ICC-14-1 (Landscape)..."):
                  carpeta_base = "./Registros_Oficiales"
                  subcarpeta_lote = f"Lote_{row['LoteProveedor']}_{row['NombreMaterial']}".replace(" ", "_")
                  ruta_destino = os.path.join(carpeta_base, str(row["Planta"]), str(row["TipoMaterial"]), str(row["NombreMaterial"]), subcarpeta_lote)
                  os.makedirs(ruta_destino, exist_ok=True)

                  # 1. Guardar certificado
                  ext_cert = os.path.splitext(str(row["RutaPdfCert"]))[1] if row["RutaPdfCert"] else ".pdf"
                  ruta_cert_final = os.path.join(ruta_destino, f"Certificado_{row['LoteProveedor']}{ext_cert}")
                  if row["RutaPdfCert"] and os.path.exists(str(row["RutaPdfCert"])):
                    shutil.copy(str(row["RutaPdfCert"]), ruta_cert_final)

                  # 2. Generar Checklist R ICC 14-1 Horizontal en PDF desde cero
                  ruta_checklist_final = os.path.join(ruta_destino, f"Checklist_RICC14-1_{row['LoteProveedor']}.pdf")
                  generar_pdf_ricc14(row, nombre_supervisor, ruta_checklist_final)

                  df_p.loc[df_p["ID"] == row["ID"], "Estado"] = "Aprobado"
                  df_p.loc[df_p["ID"] == row["ID"], "FirmaCalidad"] = str(nombre_supervisor)
                  df_p.loc[df_p["ID"] == row["ID"], "RutaChecklistOficial"] = str(ruta_checklist_final)
                  df_p.loc[df_p["ID"] == row["ID"], "RutaPdfCert"] = str(ruta_cert_final)
                  guardar_pendientes(df_p)

                  st.success("🎉 ¡Checklist PDF Horizontal R-ICC-14-1 Generado!")
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
                  label="📊 Descargar Checklist (PDF Horizontal)",
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
