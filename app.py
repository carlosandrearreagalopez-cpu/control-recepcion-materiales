from datetime import datetime
import os
import shutil
import pandas as pd
import streamlit as st
from fpdf import FPDF
import io

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
# FUNCIONES BASE Y DE DATOS
# ------------------------------------------
def cargar_pendientes():
  if os.path.exists(ARCHIVO_PENDIENTES):
    df = pd.read_csv(ARCHIVO_PENDIENTES)
    columnas_texto = [
        "Planta", "TipoMaterial", "NombreMaterial", "Proveedor",
        "LoteProveedor", "FacturaGuia", "Transporte", "Placa",
        "IngresoSAP", "DocSAP", "EstadoMarchamo", "Marchamo", "EstadoCandado",
        "Cantidad", "Unidad", "Temperatura", "FechaProduccion", "FechaVencimiento",
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
        "IngresoSAP", "DocSAP", "EstadoMarchamo", "Marchamo", "EstadoCandado", "Cantidad",
        "Unidad", "Temperatura", "FechaProduccion", "FechaVencimiento",
        "Criticidad", "Responsable", "Fecha", "Estado", "FirmaCalidad",
        "RutaCarpeta", "RutaChecklistOficial", "RutaPdfCert"
    ])

def guardar_pendientes(df):
  df.to_csv(ARCHIVO_PENDIENTES, index=False)


# ------------------------------------------
# MOTOR DE GENERACIÓN PDF (FORMATO EXACTO R ICC-14-1)
# ------------------------------------------
def dibujar_casilla(pdf, x, y, label, checked=False):
    pdf.set_xy(x, y)
    pdf.set_font('Arial', '', 7)
    pdf.cell(7, 4, label, 0, 0, 'R')
    pdf.rect(x + 8, y + 0.5, 3, 3)
    if checked:
        pdf.set_xy(x + 8, y + 0.5)
        pdf.set_font('Arial', 'B', 8)
        pdf.cell(3, 3, 'X', 0, 0, 'C')

def generar_pdf_ricc14(row, supervisor, ruta_guardado):
    pdf = FPDF(orientation='L', unit='mm', format='A4')
    pdf.add_page()
    pdf.set_auto_page_break(auto=False)
    
    # ---------------- ENCABEZADO ----------------
    pdf.set_font('Arial', 'B', 7)
    pdf.cell(40, 4, 'Código:', 'LT', 0, 'R')
    pdf.set_font('Arial', '', 7)
    pdf.cell(40, 4, 'R ICC - 14-1', 'TR', 0, 'C')
    pdf.set_font('Arial', 'B', 7)
    pdf.cell(157, 4, 'Aseguramiento de Calidad', 'LTR', 0, 'C')
    pdf.set_font('Arial', '', 7)
    pdf.cell(40, 4, 'Vigente a partir de:', 'LTR', 1, 'C')
    
    pdf.set_font('Arial', 'B', 7)
    pdf.cell(40, 4, 'Versión:', 'LB', 0, 'R')
    pdf.set_font('Arial', '', 7)
    pdf.cell(40, 4, '5', 'BR', 0, 'C')
    pdf.set_font('Arial', 'B', 9)
    pdf.cell(157, 4, 'Check list de recepción de materiales', 'LBR', 0, 'C')
    pdf.set_font('Arial', '', 7)
    pdf.cell(40, 4, 'feb-23', 'LBR', 1, 'C')
    pdf.ln(2)

    # ---------------- DATOS GENERALES ----------------
    h_row = 4.5
    pdf.set_font('Arial', 'B', 7)
    pdf.cell(40, h_row, 'Fecha de recepción:', 1, 0, 'R')
    pdf.set_font('Arial', '', 7)
    pdf.cell(70, h_row, str(row.get('Fecha', '')).split()[0], 1, 0, 'C')
    pdf.cell(27, h_row, '', 0, 0)
    pdf.set_font('Arial', 'B', 7)
    pdf.cell(60, h_row, 'Número de factura y/o guía del proveedor:', 1, 0, 'R')
    pdf.set_font('Arial', '', 7)
    pdf.cell(80, h_row, str(row.get('FacturaGuia', '')), 1, 1, 'C')

    pdf.set_font('Arial', 'B', 7)
    pdf.cell(40, h_row, 'Nombre del transporte:', 1, 0, 'R')
    pdf.set_font('Arial', '', 7)
    pdf.cell(70, h_row, str(row.get('Transporte', '')), 1, 0, 'C')
    pdf.cell(27, h_row, '', 0, 0)
    pdf.set_font('Arial', 'B', 7)
    pdf.cell(60, h_row, 'Fecha de Ingreso a SAP:', 1, 0, 'R')
    pdf.set_font('Arial', '', 7)
    pdf.cell(80, h_row, str(row.get('IngresoSAP', '')), 1, 1, 'C')

    pdf.set_font('Arial', 'B', 7)
    pdf.cell(40, h_row, 'Placa del transporte:', 1, 0, 'R')
    pdf.set_font('Arial', '', 7)
    pdf.cell(70, h_row, str(row.get('Placa', '')), 1, 0, 'C')
    pdf.cell(27, h_row, '', 0, 0)
    pdf.set_font('Arial', 'B', 7)
    pdf.cell(60, h_row, 'Número de documento de ingreso a SAP:', 1, 0, 'R')
    pdf.set_font('Arial', '', 7)
    pdf.cell(80, h_row, str(row.get('DocSAP', '')), 1, 1, 'C')
    pdf.ln(2)

    # ---------------- MARCHAMO Y CANDADO ----------------
    pdf.set_font('Arial', 'B', 7)
    pdf.set_fill_color(220, 230, 241)
    pdf.cell(100, h_row, 'Inspección del Marchamo', 0, 0, 'L', fill=True)
    pdf.cell(5, h_row, '', 0, 0)
    pdf.cell(100, h_row, 'Inspección del Candado', 0, 0, 'L', fill=True)
    pdf.cell(5, h_row, '', 0, 0)
    pdf.cell(25, h_row, 'Valor esperado', 0, 0, 'C')
    pdf.cell(42, h_row, 'Acción en caso de desviación', 0, 1, 'C')

    y_start = pdf.get_y()
    pdf.rect(10, y_start, 100, 9)
    pdf.rect(115, y_start, 100, 9)
    pdf.rect(220, y_start, 25, 9)
    pdf.rect(245, y_start, 42, 9)

    pdf.set_xy(10, y_start + 1)
    pdf.cell(5, 4, '1', 0, 0, 'C')
    pdf.cell(45, 4, '¿El marchamo se encuentra alterado?', 0, 0)
    est_mar = str(row.get('EstadoMarchamo', 'N/A'))
    dibujar_casilla(pdf, 60, y_start + 1, "SI", checked=(est_mar=="SI"))
    dibujar_casilla(pdf, 75, y_start + 1, "NO", checked=(est_mar=="NO"))
    dibujar_casilla(pdf, 90, y_start + 1, "N/A", checked=(est_mar=="N/A"))
    
    pdf.set_xy(15, y_start + 5)
    pdf.cell(25, 4, 'No. de marchamo:', 0, 0)
    pdf.cell(55, 4, str(row.get('Marchamo', '')), 'B', 0, 'C')

    pdf.set_xy(115, y_start + 2.5)
    pdf.cell(5, 4, '2', 0, 0, 'C')
    pdf.cell(45, 4, '¿El candado se encuentra alterado?', 0, 0)
    est_can = str(row.get('EstadoCandado', 'N/A'))
    dibujar_casilla(pdf, 165, y_start + 2.5, "SI", checked=(est_can=="SI"))
    dibujar_casilla(pdf, 180, y_start + 2.5, "NO", checked=(est_can=="NO"))
    dibujar_casilla(pdf, 195, y_start + 2.5, "N/A", checked=(est_can=="N/A"))

    pdf.set_xy(220, y_start + 2.5)
    pdf.set_font('Arial', 'B', 8)
    pdf.cell(25, 4, 'NO', 0, 0, 'C')
    pdf.set_font('Arial', '', 6)
    pdf.cell(42, 4, 'Retener el lote y emitir reclamo.', 0, 1, 'C')
    pdf.ln(2.5)

    # ---------------- INSPECCIÓN DEL MATERIAL ----------------
    pdf.set_font('Arial', 'B', 7)
    pdf.cell(210, h_row, 'Inspección del Material', 0, 0, 'L', fill=True)
    pdf.cell(25, h_row, 'Valor esperado', 0, 0, 'C', fill=True)
    pdf.cell(42, h_row, 'Acción en caso de desviación', 0, 1, 'C', fill=True)

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
        pdf.set_font('Arial', 'B', 7)
        pdf.cell(40, h_row, etiqueta, 1, 0, 'R')
        pdf.set_font('Arial', '', 7)
        pdf.cell(170, h_row, valor, 1, 0, 'C')
        pdf.set_fill_color(200, 200, 200)
        pdf.cell(25, h_row, '------', 1, 0, 'C', fill=True)
        pdf.cell(42, h_row, '------', 1, 1, 'C', fill=True)

    pdf.set_fill_color(220, 230, 241)
    
    preguntas_mat = [
        ("3", "Tomando en cuenta la fecha de hoy y respuesta anterior, ¿El (los) lote(s) posee(n) más de un mes de vida?", ["SI", "NO", "N/A"], "SI", "Retener el lote y emitir reclamo"),
        ("4", "¿A qué categoría de la matriz de materiales pertenece el insumo a recibir?", ["I", "II", "III"], "------", "------"),
        ("5", "¿El Certificado de Análisis (si aplica), corresponde al número de lote?", ["SI", "NO", "N/A"], "SI", "Rechazar el lote y emitir reclamo"),
        ("6", "¿Los resultados de los Certificados de Análisis cumple(n) con los parámetros críticos de la Especificación?", ["SI", "NO", "N/A"], "SI", "Rechazar el lote y emitir reclamo"),
        ("7", "¿El empaque del material se encuentra íntegro y sin deterioro?", ["SI", "NO", "N/A"], "SI", "Rechazar el lote y emitir reclamo"),
        ("8", "¿El material está protegido de la intemperie?", ["SI", "NO", "N/A"], "SI", "Rechazar el lote y emitir reclamo"),
        ("9", "¿Coincide la entrega según documento, contra lo físico en cantidad y forma?", ["SI", "NO", "N/A"], "SI", "Rechazar el lote y emitir reclamo")
    ]

    criticidad = str(row.get('Criticidad', 'I'))

    for num, preg, opts, val_esp, accion in preguntas_mat:
        y_curr = pdf.get_y()
        pdf.set_font('Arial', 'B', 7)
        pdf.cell(6, h_row, num, 1, 0, 'C')
        pdf.set_font('Arial', '', 6.5)
        pdf.cell(134, h_row, preg, 1, 0, 'L')
        pdf.cell(70, h_row, '', 1, 0)
        
        if num == "4":
            dibujar_casilla(pdf, 150, y_curr+(h_row-4)/2, opts[0], marcado=(criticidad=="I"))
            dibujar_casilla(pdf, 170, y_curr+(h_row-4)/2, opts[1], marcado=(criticidad=="II"))
            dibujar_casilla(pdf, 190, y_curr+(h_row-4)/2, opts[2], marcado=(criticidad=="III"))
        else:
            dibujar_casilla(pdf, 150, y_curr+(h_row-4)/2, opts[0], marcado=True) # SI premarcado ideal
            dibujar_casilla(pdf, 170, y_curr+(h_row-4)/2, opts[1])
            dibujar_casilla(pdf, 190, y_curr+(h_row-4)/2, opts[2])

        pdf.set_font('Arial', 'B', 7)
        if val_esp == "------":
            pdf.set_fill_color(200, 200, 200)
            pdf.cell(25, h_row, val_esp, 1, 0, 'C', fill=True)
            pdf.cell(42, h_row, accion, 1, 1, 'C', fill=True)
            pdf.set_fill_color(220, 230, 241)
        else:
            pdf.cell(25, h_row, val_esp, 1, 0, 'C')
            pdf.set_font('Arial', '', 6.5)
            pdf.cell(42, h_row, accion, 1, 1, 'C')

    pdf.ln(1)

    # ---------------- INSPECCIÓN DEL TRANSPORTE ----------------
    pdf.set_font('Arial', 'B', 7)
    pdf.cell(210, h_row, 'Inspección del Transporte', 0, 0, 'L', fill=True)
    pdf.cell(67, h_row, '', 0, 1)

    preguntas_trans = [
        ("10", "¿Se encuentra el transporte libre de plagas?", "SI", "Rechazar el lote y emitir reclamo"),
        ("11", "¿El techo, las paredes del transporte o lona se encuentran íntegras y sin agujeros?", "SI", "Rechazar el lote y emitir reclamo"),
        ("12", "¿El techo, las paredes del transporte o lona se encuentran limpios?", "SI", "Rechazar el lote y emitir reclamo"),
        ("13", "¿El piso del transporte está limpio y en buenas condiciones?", "SI", "Emitir reclamo"),
        ("14", "¿El transporte cumple con la temperatura especificada en la matriz?", "SI", "Rechazar el lote y emitir reclamo"),
        ("15", "¿El transporte se encuentra libre de olores extraños?", "SI", "Rechazar el lote y emitir reclamo")
    ]

    for num, preg, val_esp, accion in preguntas_trans:
        y_curr = pdf.get_y()
        pdf.set_font('Arial', 'B', 7)
        pdf.cell(6, h_row, num, 1, 0, 'C')
        pdf.set_font('Arial', '', 6.5)
        pdf.cell(134, h_row, preg, 1, 0, 'L')
        pdf.cell(70, h_row, '', 1, 0)
        
        dibujar_casilla(pdf, 150, y_curr+(h_row-4)/2, "SI", marcado=True)
        dibujar_casilla(pdf, 170, y_curr+(h_row-4)/2, "NO")
        if num == "14":
             dibujar_casilla(pdf, 190, y_curr+(h_row-4)/2, "N/A", marcado=(str(row.get('Temperatura', ''))=="N/A"))

        pdf.set_font('Arial', 'B', 7)
        pdf.cell(25, h_row, val_esp, 1, 0, 'C')
        pdf.set_font('Arial', '', 6.5)
        pdf.cell(42, h_row, accion, 1, 1, 'C')
        
        # Campo de temperatura extra bajo la 14
        if num == "14":
            temp_val = str(row.get('Temperatura', 'N/A'))
            pdf.cell(6, h_row, '', 'LR', 0)
            pdf.cell(204, h_row, f"Temperatura (°C): {temp_val}", 'LR', 0, 'L')
            pdf.cell(67, h_row, '', 'LR', 1)

    pdf.ln(2)
    
    # ---------------- RESULTADO, OBSERVACIONES Y FIRMAS ----------------
    y_curr = pdf.get_y()
    pdf.rect(10, y_curr, 277, 8)
    pdf.set_xy(10, y_curr+2)
    pdf.set_font('Arial', 'B', 9)
    pdf.cell(50, 4, 'Resultado final:', 0, 0, 'R')
    dibujar_casilla(pdf, 70, y_curr+2, "", checked=True)
    pdf.set_xy(82, y_curr+2)
    pdf.cell(20, 4, 'Aprobado', 0, 0, 'L')
    dibujar_casilla(pdf, 110, y_curr+2, "", checked=False)
    pdf.set_xy(122, y_curr+2)
    pdf.cell(30, 4, 'Desaprobado', 0, 1, 'L')
    
    pdf.ln(4)
    pdf.set_font('Arial', 'B', 6)
    pdf.cell(277, 3, 'OBSERVACIONES:', 0, 1, 'L')
    pdf.cell(277, 3, '** Si el marchamo N/A por ser proveedores locales, basta con la inspección del material y del transporte.', 0, 1, 'L')
    pdf.cell(277, 3, '** Para los casos que se realiza la revisión en el transporte y determina que no cumple y los insumos se ven AFECTADOS, PROCEDE A RECHAZARLOS Y EMITIR RECLAMO.', 0, 1, 'L')
    
    pdf.ln(5)
    pdf.set_font('Arial', 'B', 8)
    pdf.cell(65, 5, 'Nombre de la persona que realiza la recepción:', 0, 0)
    pdf.set_font('Arial', '', 8)
    pdf.cell(80, 5, str(row.get('Responsable', '')), 'B', 0, 'C')
    pdf.cell(10, 5, '', 0, 0)
    pdf.set_font('Arial', 'B', 8)
    pdf.cell(25, 5, 'Vo. Bo. Calidad:', 0, 0)
    pdf.set_font('Arial', '', 8)
    pdf.cell(80, 5, supervisor, 'B', 1, 'C')

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

st.sidebar.header("⚙️ Configuración")
archivo_catalogo_nuevo = st.sidebar.file_uploader("Actualizar Catálogo Excel", type=["xlsx"])
if archivo_catalogo_nuevo is not None:
  with open(ARCHIVO_CATALOGO, "wb") as f:
      f.write(archivo_catalogo_nuevo.getbuffer())
  df_db = pd.read_excel(ARCHIVO_CATALOGO)
  df_db.columns = df_db.columns.str.strip()
  st.sidebar.success("¡Catálogo actualizado!")

tab1, tab2, tab3, tab4 = st.tabs([
    "📝 1. Ingreso de Operador",
    "✍️ 2. Panel de Calidad",
    "🔍 3. Trazabilidad",
    "📊 4. Control Excel (Gerencial)"
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
      st.subheader("Datos de Transporte, SAP y Lote")
      t1, t2, t3 = st.columns(3)
      with t1:
        factura_guia = st.text_input("No. de Factura/Guía", placeholder="Ej. 122")
      with t2:
        placa_transporte = st.text_input("Placa de Transporte")
      with t3:
        lote_proveedor = st.text_input("Lote del Proveedor")

      s1, s2, s3 = st.columns(3)
      with s1:
        ingreso_sap = st.date_input("Fecha de Ingreso a SAP")
      with s2:
        doc_sap = st.text_input("No. Documento SAP")
      with s3:
        marchamo = st.text_input("No. de Marchamo")

      m1, m2 = st.columns(2)
      with m1:
        estado_marchamo = st.selectbox("¿El marchamo se encuentra alterado?", ["NO", "SI", "N/A"])
      with m2:
        estado_candado = st.selectbox("¿El candado se encuentra alterado?", ["NO", "SI", "N/A"])

      c4, c5 = st.columns(2)
      with c4:
        cantidad = st.number_input("Cantidad Recibida", min_value=0.0, format="%.2f")
      with c5:
        unidad = st.selectbox("Unidad de Medida", ["Kg", "Litros", "Unidades", "Toneladas", "Galones", "Cajas"])

      c6, c7 = st.columns(2)
      with c6:
        fecha_produccion = st.date_input("Fecha de Producción")
      with c7:
        fecha_vencimiento = st.date_input("Fecha de Vencimiento")

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
            "EstadoMarchamo": estado_marchamo,
            "Marchamo": marchamo,
            "EstadoCandado": estado_candado,
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
                "EstadoMarchamo": str(item["EstadoMarchamo"]),
                "Marchamo": str(item["Marchamo"]),
                "EstadoCandado": str(item["EstadoCandado"]),
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

          col_btn1, col_btn2 = st.columns(2)
          with col_btn1:
            with st.form(f"form_firma_{row['ID']}"):
              nombre_supervisor = st.text_input("Firma del Supervisor")
              aprobar = st.form_submit_button("Aprobar y Generar Checklist PDF 🛡️", type="primary")

            if aprobar:
              if not nombre_supervisor:
                st.error("Ingrese su firma.")
              else:
                with st.spinner("Generando Formato Oficial R ICC-14-1..."):
                  carpeta_base = "./Registros_Oficiales"
                  subcarpeta_lote = f"Lote_{row['LoteProveedor']}"
                  # Estructura escalonada de carpetas automatizada
                  ruta_destino = os.path.join(carpeta_base, str(row["Planta"]), str(row["TipoMaterial"]), str(row["NombreMaterial"]), subcarpeta_lote)
                  os.makedirs(ruta_destino, exist_ok=True)

                  ext_cert = os.path.splitext(str(row["RutaPdfCert"]))[1] if row["RutaPdfCert"] else ".pdf"
                  ruta_cert_final = os.path.join(ruta_destino, f"Certificado_Original_{row['LoteProveedor']}{ext_cert}")
                  if row["RutaPdfCert"] and os.path.exists(str(row["RutaPdfCert"])):
                    shutil.copy(str(row["RutaPdfCert"]), ruta_cert_final)

                  # Generación nativa del PDF horizontal R ICC-14-1
                  ruta_checklist_final = os.path.join(ruta_destino, f"Checklist_RICC14-1_{row['LoteProveedor']}.pdf")
                  generar_pdf_ricc14(row, nombre_supervisor, ruta_checklist_final)

                  df_p.loc[df_p["ID"] == row["ID"], "Estado"] = "Aprobado"
                  df_p.loc[df_p["ID"] == row["ID"], "FirmaCalidad"] = str(nombre_supervisor)
                  df_p.loc[df_p["ID"] == row["ID"], "RutaCarpeta"] = str(ruta_destino)
                  df_p.loc[df_p["ID"] == row["ID"], "RutaChecklistOficial"] = str(ruta_checklist_final)
                  df_p.loc[df_p["ID"] == row["ID"], "RutaPdfCert"] = str(ruta_cert_final)
                  guardar_pendientes(df_p)

                  st.success("🎉 ¡Checklist PDF Generado, organizado en carpetas y Aprobado!")
                  st.rerun()

          with col_btn2:
            st.markdown("<br>", unsafe_allow_html=True)
            if st.button("🗑️ Eliminar Registro", key=f"del_pend_{row['ID']}"):
              df_p = df_p[df_p["ID"] != row["ID"]]
              guardar_pendientes(df_p)
              st.rerun()

# ------------------------------------------
# PESTAÑA 3: BÚSQUEDA Y TRAZABILIDAD
# ------------------------------------------
with tab3:
  st.header("🔍 Buscador y Trazabilidad")
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

    st.dataframe(resultado[["Planta", "NombreMaterial", "FacturaGuia", "LoteProveedor", "Fecha", "FirmaCalidad"]], use_container_width=True)

    for idx, row in resultado.iterrows():
      with st.expander(f"📁 {row['NombreMaterial']} | Lote: {row['LoteProveedor']} | Supervisor: {row['FirmaCalidad']}"):
        col_d1, col_d2, col_d3 = st.columns(3)
        ruta_chk = str(row.get("RutaChecklistOficial", ""))
        with col_d1:
          if pd.notna(ruta_chk) and ruta_chk != "nan" and os.path.exists(ruta_chk):
            with open(ruta_chk, "rb") as f:
              st.download_button("📊 Descargar Checklist", data=f, file_name=os.path.basename(ruta_chk), mime="application/pdf", key=f"dl_chk_{row['ID']}")
        
        ruta_cert = str(row.get("RutaPdfCert", ""))
        with col_d2:
          if pd.notna(ruta_cert) and ruta_cert != "nan" and os.path.exists(ruta_cert):
            with open(ruta_cert, "rb") as f:
              st.download_button("📄 Descargar Certificado", data=f, file_name=os.path.basename(ruta_cert), mime="application/pdf" if ruta_cert.endswith(".pdf") else "image/jpeg", key=f"dl_cert_{row['ID']}")
        
        with col_d3:
          if st.button("🗑️ Borrar", key=f"del_apr_{row['ID']}"):
            df_p = df_p[df_p["ID"] != row["ID"]]
            guardar_pendientes(df_p)
            st.rerun()

# ------------------------------------------
# PESTAÑA 4: CONTROL EXCEL (GERENCIAL)
# ------------------------------------------
with tab4:
  st.header("📊 Control de Ingresos a Detalle")
  st.markdown("Consolide todos los ingresos aprobados en una sola base de datos de Excel para análisis gerencial.")
  
  df_p = cargar_pendientes()
  aprobados_excel = df_p[df_p["Estado"] == "Aprobado"].copy() if not df_p.empty else pd.DataFrame()

  if len(aprobados_excel) == 0:
    st.info("No hay información suficiente para generar el reporte Excel.")
  else:
    # Columnas seleccionadas para el reporte gerencial
    columnas_reporte = [
        "Planta", "TipoMaterial", "NombreMaterial", "Proveedor", "LoteProveedor", 
        "FacturaGuia", "Placa", "IngresoSAP", "DocSAP", "Cantidad", "Unidad", 
        "Temperatura", "Criticidad", "Fecha", "Responsable", "FirmaCalidad"
    ]
    df_reporte = aprobados_excel[columnas_reporte].copy()
    st.dataframe(df_reporte, use_container_width=True)

    # Convertir DataFrame a Excel en memoria
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_reporte.to_excel(writer, index=False, sheet_name='Historial de Ingresos')
    excel_data = output.getvalue()

    st.markdown("<br>", unsafe_allow_html=True)
    st.download_button(
        label="📥 Descargar Base de Datos Completa (Excel)",
        data=excel_data,
        file_name=f"Control_Ingresos_LIF_{datetime.now().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary"
    )
