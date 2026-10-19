from datetime import datetime
import os
import shutil
import pandas as pd
import numpy as np
from PIL import Image
import streamlit as st
from fpdf import FPDF
import io

# Librerías adicionales
try:
    from streamlit_drawable_canvas import st_canvas
except ImportError:
    st_canvas = None
    st.error("⚠️ Falta 'streamlit-drawable-canvas' en requirements.txt")

try:
    from pypdf import PdfWriter, PdfReader
except ImportError:
    st.error("⚠️ Falta 'pypdf' en requirements.txt para unir los PDFs.")

# Configuración de diseño de página
st.set_page_config(page_title="Control de Recepción y Calidad", page_icon="📦", layout="wide", initial_sidebar_state="expanded")

st.markdown("""
    <style>
    .main { background-color: #f8f9fa; }
    .stTabs [data-baseweb="tab-list"] { gap: 10px; }
    .stTabs [data-baseweb="tab"] { height: 50px; white-space: pre-wrap; background-color: #ffffff; border-radius: 5px; font-weight: 600; padding: 10px 20px; box-shadow: 0 2px 4px rgba(0,0,0,0.05); }
    .stTabs [aria-selected="true"] { background-color: #0066cc !important; color: #ffffff !important; }
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
        "IngresoSAP", "DocSAP", "EstadoMarchamo", "Marchamo", "EstadoCandado",
        "Cantidad", "Unidad", "TempTransporte", "TempMaterial", "FechaProduccion", "FechaVencimiento",
        "Criticidad", "Responsable", "Fecha", "Estado", "FirmaCalidad",
        "RutaCarpeta", "RutaChecklistOficial", "RutaPdfCert", "RutaFirmaOp", "RutaFirmaSup"
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
        "Unidad", "TempTransporte", "TempMaterial", "FechaProduccion", "FechaVencimiento",
        "Criticidad", "Responsable", "Fecha", "Estado", "FirmaCalidad",
        "RutaCarpeta", "RutaChecklistOficial", "RutaPdfCert", "RutaFirmaOp", "RutaFirmaSup"
    ])

def guardar_pendientes(df):
  df.to_csv(ARCHIVO_PENDIENTES, index=False)

# ------------------------------------------
# GENERADOR PDF (DISEÑO MEJORADO)
# ------------------------------------------
def dibujar_casilla(pdf, x, y, label, checked=False):
    pdf.set_xy(x, y)
    pdf.set_font('Arial', '', 8)
    pdf.cell(8, 4, label, 0, 0, 'R')
    pdf.rect(x + 9, y + 0.5, 3, 3)
    if checked:
        pdf.set_xy(x + 9, y + 0.5)
        pdf.set_font('Arial', 'B', 8)
        pdf.cell(3, 3, 'X', 0, 0, 'C')

def generar_pdf_ricc14(row, supervisor, ruta_guardado_temp):
    pdf = FPDF(orientation='L', unit='mm', format='A4')
    pdf.add_page()
    pdf.set_auto_page_break(auto=False)
    pdf.set_fill_color(220, 230, 241) 
    
    # --- ENCABEZADO ---
    pdf.set_font('Arial', 'B', 8)
    pdf.cell(40, 5, 'Código:', 'LT', 0, 'R')
    pdf.set_font('Arial', '', 8)
    pdf.cell(40, 5, 'R ICC - 14-1', 'TR', 0, 'C')
    pdf.set_font('Arial', 'B', 8)
    pdf.cell(157, 5, 'Aseguramiento de Calidad', 'LTR', 0, 'C')
    pdf.set_font('Arial', '', 8)
    pdf.cell(40, 5, 'Vigente a partir de:', 'LTR', 1, 'C')
    
    pdf.set_font('Arial', 'B', 8)
    pdf.cell(40, 5, 'Versión:', 'LB', 0, 'R')
    pdf.set_font('Arial', '', 8)
    pdf.cell(40, 5, '5', 'BR', 0, 'C')
    pdf.set_font('Arial', 'B', 10)
    pdf.cell(157, 5, 'Check list de recepción de materiales', 'LBR', 0, 'C')
    pdf.set_font('Arial', '', 8)
    pdf.cell(40, 5, 'feb-23', 'LBR', 1, 'C')
    pdf.ln(3)

    # --- DATOS GENERALES ---
    h_row = 5
    pdf.set_font('Arial', 'B', 8)
    pdf.cell(40, h_row, 'Fecha de recepción:', 1, 0, 'R')
    pdf.set_font('Arial', '', 8)
    pdf.cell(70, h_row, str(row.get('Fecha', '')).split()[0], 1, 0, 'C')
    pdf.cell(27, h_row, '', 0, 0)
    pdf.set_font('Arial', 'B', 8)
    pdf.cell(60, h_row, 'Número de factura y/o guía:', 1, 0, 'R')
    pdf.set_font('Arial', '', 8)
    pdf.cell(80, h_row, str(row.get('FacturaGuia', '')), 1, 1, 'C')

    pdf.set_font('Arial', 'B', 8)
    pdf.cell(40, h_row, 'Nombre del transporte:', 1, 0, 'R')
    pdf.set_font('Arial', '', 8)
    pdf.cell(70, h_row, str(row.get('Transporte', '')), 1, 0, 'C')
    pdf.cell(27, h_row, '', 0, 0)
    pdf.set_font('Arial', 'B', 8)
    pdf.cell(60, h_row, 'Fecha de Ingreso a SAP:', 1, 0, 'R')
    pdf.set_font('Arial', '', 8)
    pdf.cell(80, h_row, str(row.get('IngresoSAP', '')), 1, 1, 'C')

    pdf.set_font('Arial', 'B', 8)
    pdf.cell(40, h_row, 'Placa del transporte:', 1, 0, 'R')
    pdf.set_font('Arial', '', 8)
    pdf.cell(70, h_row, str(row.get('Placa', '')), 1, 0, 'C')
    pdf.cell(27, h_row, '', 0, 0)
    pdf.set_font('Arial', 'B', 8)
    pdf.cell(60, h_row, 'Número de documento SAP:', 1, 0, 'R')
    pdf.set_font('Arial', '', 8)
    pdf.cell(80, h_row, str(row.get('DocSAP', '')), 1, 1, 'C')
    pdf.ln(3)

    # --- MARCHAMO Y CANDADO ---
    pdf.set_font('Arial', 'B', 8)
    pdf.cell(100, h_row, 'Inspección del Marchamo', 0, 0, 'L', fill=True)
    pdf.cell(5, h_row, '', 0, 0)
    pdf.cell(100, h_row, 'Inspección del Candado', 0, 0, 'L', fill=True)
    pdf.cell(5, h_row, '', 0, 0)
    pdf.cell(25, h_row, 'Valor esperado', 0, 0, 'C')
    pdf.cell(42, h_row, 'Acción en caso desviación', 0, 1, 'C')

    y_start = pdf.get_y()
    pdf.rect(10, y_start, 100, 10)
    pdf.rect(115, y_start, 100, 10)
    pdf.rect(220, y_start, 25, 10)
    pdf.rect(245, y_start, 42, 10)

    pdf.set_xy(10, y_start + 1.5)
    pdf.cell(5, 4, '1', 0, 0, 'C')
    pdf.cell(45, 4, '¿El marchamo se encuentra alterado?', 0, 0)
    est_mar = str(row.get('EstadoMarchamo', 'N/A'))
    dibujar_casilla(pdf, 60, y_start + 1.5, "SI", checked=(est_mar=="SI"))
    dibujar_casilla(pdf, 75, y_start + 1.5, "NO", checked=(est_mar=="NO"))
    dibujar_casilla(pdf, 90, y_start + 1.5, "N/A", checked=(est_mar=="N/A"))
    
    pdf.set_xy(15, y_start + 5.5)
    pdf.cell(25, 4, 'No. de marchamo:', 0, 0)
    pdf.cell(55, 4, str(row.get('Marchamo', '')), 'B', 0, 'C')

    pdf.set_xy(115, y_start + 3)
    pdf.cell(5, 4, '2', 0, 0, 'C')
    pdf.cell(45, 4, '¿El candado se encuentra alterado?', 0, 0)
    est_can = str(row.get('EstadoCandado', 'N/A'))
    dibujar_casilla(pdf, 165, y_start + 3, "SI", checked=(est_can=="SI"))
    dibujar_casilla(pdf, 180, y_start + 3, "NO", checked=(est_can=="NO"))
    dibujar_casilla(pdf, 195, y_start + 3, "N/A", checked=(est_can=="N/A"))

    pdf.set_xy(220, y_start + 3)
    pdf.set_font('Arial', 'B', 8)
    pdf.cell(25, 4, 'NO', 0, 0, 'C')
    pdf.set_font('Arial', '', 7)
    pdf.cell(42, 4, 'Retener lote y emitir reclamo.', 0, 1, 'C')
    pdf.ln(3.5)

    # --- MATERIAL ---
    pdf.set_font('Arial', 'B', 8)
    pdf.cell(210, h_row, 'Inspección del Material', 0, 0, 'L', fill=True)
    pdf.cell(25, h_row, 'Valor esperado', 0, 0, 'C', fill=True)
    pdf.cell(42, h_row, 'Acción en caso desviación', 0, 1, 'C', fill=True)

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
        pdf.cell(40, h_row, etiqueta, 1, 0, 'R')
        pdf.set_font('Arial', '', 8)
        pdf.cell(170, h_row, valor, 1, 0, 'C')
        pdf.set_fill_color(200, 200, 200)
        pdf.cell(25, h_row, '------', 1, 0, 'C', fill=True)
        pdf.cell(42, h_row, '------', 1, 1, 'C', fill=True)

    pdf.set_fill_color(220, 230, 241)
    preguntas_mat = [
        ("3", "¿El (los) lote(s) posee(n) más de un mes de vida?", ["SI", "NO", "N/A"], "SI", "Retener el lote y emitir reclamo"),
        ("4", "¿A qué categoría de la matriz de materiales pertenece?", ["I", "II", "III"], "------", "------"),
        ("5", "¿El Certificado de Análisis corresponde al lote?", ["SI", "NO", "N/A"], "SI", "Rechazar lote y emitir reclamo"),
        ("6", "¿Los resultados del Certificado cumplen parámetros críticos?", ["SI", "NO", "N/A"], "SI", "Rechazar lote y emitir reclamo"),
        ("7", "¿El empaque se encuentra íntegro y sin deterioro?", ["SI", "NO", "N/A"], "SI", "Rechazar lote y emitir reclamo"),
        ("8", "¿El material está protegido de la intemperie?", ["SI", "NO", "N/A"], "SI", "Rechazar lote y emitir reclamo"),
        ("9", "¿Coincide la entrega según documento contra lo físico?", ["SI", "NO", "N/A"], "SI", "Rechazar lote y emitir reclamo")
    ]

    criticidad = str(row.get('Criticidad', 'I'))

    for num, preg, opts, val_esp, accion in preguntas_mat:
        y_curr = pdf.get_y()
        pdf.set_font('Arial', 'B', 8)
        pdf.cell(6, h_row, num, 1, 0, 'C')
        pdf.set_font('Arial', '', 7.5)
        pdf.cell(120, h_row, preg, 1, 0, 'L')
        pdf.cell(84, h_row, '', 1, 0) # Espacio para checkboxes
        
        x_base = 135
        if num == "4":
            dibujar_casilla(pdf, x_base, y_curr+(h_row-4)/2, opts[0], checked=(criticidad=="I"))
            dibujar_casilla(pdf, x_base+25, y_curr+(h_row-4)/2, opts[1], checked=(criticidad=="II"))
            dibujar_casilla(pdf, x_base+50, y_curr+(h_row-4)/2, opts[2], checked=(criticidad=="III"))
        else:
            dibujar_casilla(pdf, x_base, y_curr+(h_row-4)/2, opts[0], checked=True)
            dibujar_casilla(pdf, x_base+25, y_curr+(h_row-4)/2, opts[1])
            dibujar_casilla(pdf, x_base+50, y_curr+(h_row-4)/2, opts[2])

        pdf.set_font('Arial', 'B', 8)
        if val_esp == "------":
            pdf.set_fill_color(200, 200, 200)
            pdf.cell(25, h_row, val_esp, 1, 0, 'C', fill=True)
            pdf.cell(42, h_row, accion, 1, 1, 'C', fill=True)
            pdf.set_fill_color(220, 230, 241)
        else:
            pdf.cell(25, h_row, val_esp, 1, 0, 'C')
            pdf.set_font('Arial', '', 7)
            pdf.cell(42, h_row, accion, 1, 1, 'C')

    pdf.ln(1)

    # --- TRANSPORTE ---
    pdf.set_font('Arial', 'B', 8)
    pdf.cell(210, h_row, 'Inspección del Transporte', 0, 0, 'L', fill=True)
    pdf.cell(67, h_row, '', 0, 1)

    preguntas_trans = [
        ("10", "¿Se encuentra el transporte libre de plagas?", "SI", "Rechazar lote y emitir reclamo"),
        ("11", "¿El techo, paredes o lona se encuentran íntegras?", "SI", "Rechazar lote y emitir reclamo"),
        ("12", "¿El techo, paredes o lona se encuentran limpios?", "SI", "Rechazar lote y emitir reclamo"),
        ("13", "¿El piso del transporte está limpio y buenas condiciones?", "SI", "Emitir reclamo")
    ]

    for num, preg, val_esp, accion in preguntas_trans:
        y_curr = pdf.get_y()
        pdf.set_font('Arial', 'B', 8)
        pdf.cell(6, h_row, num, 1, 0, 'C')
        pdf.set_font('Arial', '', 7.5)
        pdf.cell(120, h_row, preg, 1, 0, 'L')
        pdf.cell(84, h_row, '', 1, 0)
        
        dibujar_casilla(pdf, 135, y_curr+(h_row-4)/2, "SI", checked=True)
        dibujar_casilla(pdf, 160, y_curr+(h_row-4)/2, "NO")

        pdf.set_font('Arial', 'B', 8)
        pdf.cell(25, h_row, val_esp, 1, 0, 'C')
        pdf.set_font('Arial', '', 7)
        pdf.cell(42, h_row, accion, 1, 1, 'C')
        
    # --- TEMPERATURAS (Transporte y Material) ---
    temp_trans = str(row.get('TempTransporte', 'N/A'))
    temp_mat = str(row.get('TempMaterial', 'N/A'))
    
    y_curr = pdf.get_y()
    pdf.set_font('Arial', 'B', 8)
    pdf.cell(6, h_row, '14', 1, 0, 'C')
    pdf.set_font('Arial', '', 7.5)
    pdf.cell(120, h_row, '¿El transporte cumple con la temperatura especificada?', 1, 0, 'L')
    pdf.cell(84, h_row, '', 1, 0)
    dibujar_casilla(pdf, 135, y_curr+(h_row-4)/2, "SI", checked=(temp_trans != "N/A" or temp_mat != "N/A"))
    dibujar_casilla(pdf, 160, y_curr+(h_row-4)/2, "NO")
    dibujar_casilla(pdf, 185, y_curr+(h_row-4)/2, "N/A", checked=(temp_trans == "N/A" and temp_mat == "N/A"))
    
    pdf.set_font('Arial', 'B', 8)
    pdf.cell(25, h_row, "SI", 1, 0, 'C')
    pdf.set_font('Arial', '', 7)
    pdf.cell(42, h_row, "Rechazar lote", 1, 1, 'C')

    # Fila de recuadros de temperatura
    pdf.cell(6, h_row+2, '', 'LR', 0)
    pdf.set_font('Arial', 'B', 8)
    pdf.cell(102, h_row+2, f" Temperatura Transporte (°C):   {temp_trans}", 'L', 0, 'L')
    pdf.cell(102, h_row+2, f" Temperatura del Material (°C):   {temp_mat}", 'R', 0, 'L')
    pdf.cell(67, h_row+2, '', 'R', 1)

    # Pregunta 15
    y_curr = pdf.get_y()
    pdf.set_font('Arial', 'B', 8)
    pdf.cell(6, h_row, '15', 1, 0, 'C')
    pdf.set_font('Arial', '', 7.5)
    pdf.cell(120, h_row, '¿El transporte se encuentra libre de olores extraños?', 1, 0, 'L')
    pdf.cell(84, h_row, '', 1, 0)
    dibujar_casilla(pdf, 135, y_curr+(h_row-4)/2, "SI", checked=True)
    dibujar_casilla(pdf, 160, y_curr+(h_row-4)/2, "NO")
    pdf.set_font('Arial', 'B', 8)
    pdf.cell(25, h_row, "SI", 1, 0, 'C')
    pdf.set_font('Arial', '', 7)
    pdf.cell(42, h_row, "Rechazar lote", 1, 1, 'C')

    pdf.ln(3)
    
    # --- RESULTADO, OBSERVACIONES Y FIRMAS ---
    y_curr = pdf.get_y()
    pdf.rect(10, y_curr, 277, 10)
    pdf.set_xy(10, y_curr+3)
    pdf.set_font('Arial', 'B', 10)
    pdf.cell(60, 4, 'Resultado final:', 0, 0, 'R')
    dibujar_casilla(pdf, 80, y_curr+3, "", checked=True)
    pdf.set_xy(92, y_curr+3)
    pdf.cell(30, 4, 'Aprobado', 0, 0, 'L')
    dibujar_casilla(pdf, 130, y_curr+3, "", checked=False)
    pdf.set_xy(142, y_curr+3)
    pdf.cell(30, 4, 'Desaprobado', 0, 1, 'L')
    
    pdf.ln(5)
    pdf.set_font('Arial', 'B', 7)
    pdf.cell(277, 3, 'OBSERVACIONES:', 0, 1, 'L')
    pdf.cell(277, 3, '** Si el marchamo N/A por ser proveedores locales, basta con la inspección del material y del transporte.', 0, 1, 'L')
    pdf.cell(277, 3, '** Para los casos que se realiza la revisión en el transporte y determina que no cumple, PROCEDE A RECHAZARLOS.', 0, 1, 'L')
    
    pdf.ln(12) 
    y_firmas = pdf.get_y()
    
    # Inserción de imágenes de firma convertidas a RGB para evitar crashes
    ruta_firma_op = str(row.get('RutaFirmaOp', ''))
    if ruta_firma_op and os.path.exists(ruta_firma_op):
        try:
            pdf.image(ruta_firma_op, x=50, y=y_firmas-12, w=40, h=15)
        except: pass
        
    ruta_firma_sup = str(row.get('RutaFirmaSup', ''))
    if ruta_firma_sup and os.path.exists(ruta_firma_sup):
        try:
            pdf.image(ruta_firma_sup, x=180, y=y_firmas-12, w=40, h=15)
        except: pass

    pdf.set_font('Arial', 'B', 9)
    pdf.cell(80, 5, 'Nombre de la persona que realiza la recepción:', 0, 0)
    pdf.set_font('Arial', '', 9)
    pdf.cell(60, 5, str(row.get('Responsable', '')), 'T', 0, 'C')
    pdf.cell(10, 5, '', 0, 0)
    pdf.set_font('Arial', 'B', 9)
    pdf.cell(30, 5, 'Vo. Bo. Calidad:', 0, 0)
    pdf.set_font('Arial', '', 9)
    pdf.cell(60, 5, supervisor, 'T', 1, 'C')

    pdf.output(ruta_guardado_temp)

# --- FUSIÓN DE PDFS ---
def generar_documento_unificado(ruta_checklist, ruta_certificado, ruta_final):
    # Si el certificado es un PDF, los unimos. Si es imagen, la ignoramos para el merge pero sigue guardada en carpeta
    try:
        from pypdf import PdfWriter, PdfReader
        merger = PdfWriter()
        
        # 1. Agregar el Checklist Oficial
        merger.append(ruta_checklist)
        
        # 2. Agregar el Certificado si es PDF
        if ruta_certificado and os.path.exists(ruta_certificado) and ruta_certificado.lower().endswith('.pdf'):
            merger.append(ruta_certificado)
            
        merger.write(ruta_final)
        merger.close()
    except Exception as e:
        # Fallback de seguridad: si no se puede unir, el documento final será solo el checklist
        shutil.copy(ruta_checklist, ruta_final)


# Inicializar
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
archivo_catalogo_nuevo = st.sidebar.file_uploader("Actualizar Catálogo", type=["xlsx"])
if archivo_catalogo_nuevo is not None:
  with open(ARCHIVO_CATALOGO, "wb") as f:
      f.write(archivo_catalogo_nuevo.getbuffer())
  df_db = pd.read_excel(ARCHIVO_CATALOGO)
  df_db.columns = df_db.columns.str.strip()
  st.sidebar.success("¡Catálogo actualizado!")

tab1, tab2, tab3, tab4 = st.tabs([
    "📝 1. Ingreso de Operador", "✍️ 2. Panel de Calidad", "🔍 3. Trazabilidad", "📊 4. Control Excel"
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

    requiere_temperatura = any(kw in str(material_sel).lower() or kw in str(tipo_mat_sel).lower() for kw in ["pulpa", "concentrado", "leche", "refrigerado", "congelado"])

    with st.form("form_item_recepcion"):
      st.subheader("Datos de Transporte, SAP y Lote")
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
        fecha_vencimiento = st.date_input("Fecha de Vencimiento / Caducidad")

      temp_trans = "N/A"
      temp_mat = "N/A"
      if requiere_temperatura:
        st.markdown("---")
        st.subheader("❄️ Control de Temperaturas (Requerido)")
        tmp1, tmp2 = st.columns(2)
        with tmp1:
            temp_trans = st.number_input("Temp. Transporte (°C)", value=4.0, format="%.1f")
        with tmp2:
            temp_mat = st.number_input("Temp. del Material (°C)", value=4.0, format="%.1f")

      st.markdown("---")
      certificado_file = st.file_uploader("Subir Archivo del Certificado (PDF Recomendado)", type=["jpg", "jpeg", "png", "pdf"])

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
            "Planta": planta_sel, "TipoMaterial": tipo_mat_sel, "Proveedor": proveedor_sel,
            "NombreMaterial": material_sel, "FacturaGuia": factura_guia, "Transporte": proveedor_sel, 
            "Placa": placa_transporte, "IngresoSAP": str(ingreso_sap), "DocSAP": doc_sap,
            "EstadoMarchamo": estado_marchamo, "Marchamo": marchamo, "EstadoCandado": estado_candado,
            "Lote": lote_proveedor, "Cantidad": cantidad, "Unidad": unidad, 
            "TempTransporte": temp_trans, "TempMaterial": temp_mat,
            "Produccion": str(fecha_produccion), "Vencimiento": str(fecha_vencimiento),
            "Criticidad": criticidad_val, "RutaTemp": ruta_cert_temp,
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

      st.markdown("---")
      st.subheader("3. Finalizar y Enviar a Revisión")
      responsable = st.text_input("Operador Responsable")
      
      st.write("✍️ **Dibuja tu Firma de Operador:**")
      canvas_op = None
      if st_canvas:
          canvas_op = st_canvas(stroke_width=2, stroke_color="#000", background_color="#eee", height=100, width=400, drawing_mode="freedraw", key="firma_op")

      if st.button("🚀 Enviar a Calidad", type="primary"):
        if not responsable:
          st.error("⚠️ Ingresa el nombre del operador responsable.")
        else:
          ruta_firma_op = ""
          if canvas_op is not None:
              try:
                  if canvas_op.image_data is not None and np.any(canvas_op.image_data):
                      os.makedirs("temp_firmas", exist_ok=True)
                      ruta_firma_op = f"temp_firmas/firma_op_{datetime.now().strftime('%Y%m%d%H%M%S')}.png"
                      # Convertir a RGB para evitar problemas de compatibilidad en PDF
                      img = Image.fromarray(canvas_op.image_data.astype('uint8'), 'RGBA').convert('RGB')
                      img.save(ruta_firma_op)
              except Exception:
                  pass

          df_p = cargar_pendientes()
          nuevo_id = int(df_p["ID"].max()) + 1 if not df_p.empty and not pd.isna(df_p["ID"].max()) else 1

          for item in st.session_state.items_ingreso_actual:
            nuevo_registro = {
                "ID": nuevo_id, "Planta": str(item["Planta"]), "TipoMaterial": str(item["TipoMaterial"]),
                "NombreMaterial": str(item["NombreMaterial"]), "Proveedor": str(item["Proveedor"]),
                "FacturaGuia": str(item["FacturaGuia"]), "Transporte": str(item["Transporte"]),
                "Placa": str(item["Placa"]), "IngresoSAP": str(item["IngresoSAP"]), "DocSAP": str(item["DocSAP"]),
                "EstadoMarchamo": str(item["EstadoMarchamo"]), "Marchamo": str(item["Marchamo"]),
                "EstadoCandado": str(item["EstadoCandado"]), "LoteProveedor": str(item["Lote"]),
                "Cantidad": item["Cantidad"], "Unidad": str(item["Unidad"]), 
                "TempTransporte": str(item["TempTransporte"]), "TempMaterial": str(item["TempMaterial"]),
                "FechaProduccion": str(item["Produccion"]), "FechaVencimiento": str(item["Vencimiento"]),
                "Criticidad": str(item["Criticidad"]), "Responsable": str(responsable),
                "Fecha": datetime.now().strftime("%Y-%m-%d %H:%M"), "Estado": "Pendiente", "FirmaCalidad": "Sin firmar",
                "RutaCarpeta": "", "RutaChecklistOficial": "", "RutaPdfCert": str(item["RutaTemp"]),
                "RutaFirmaOp": ruta_firma_op, "RutaFirmaSup": ""
            }
            df_p = pd.concat([df_p, pd.DataFrame([nuevo_registro])], ignore_index=True)
            nuevo_id += 1

          guardar_pendientes(df_p)
          st.session_state.items_ingreso_actual = []
          st.success("🎉 ¡Enviados a calidad con éxito!")
          st.rerun()

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
          st.write(f"**Operador Responsable:** {row['Responsable']}")
          if row["RutaPdfCert"] and os.path.exists(str(row["RutaPdfCert"])):
            st.write("📄 Certificado adjunto.")

          nombre_supervisor = st.text_input("Nombre del Supervisor", key=f"sup_nom_{row['ID']}")
          st.write("✍️ **Dibuja tu Firma de Supervisor:**")
          
          canvas_sup = None
          if st_canvas:
              canvas_sup = st_canvas(stroke_width=2, stroke_color="#000", background_color="#eee", height=100, width=400, drawing_mode="freedraw", key=f"firma_sup_{row['ID']}")

          col_btn1, col_btn2 = st.columns(2)
          with col_btn1:
            if st.button("Aprobar y Generar Checklist + Certificado 🛡️", key=f"btn_apr_{row['ID']}", type="primary"):
              if not nombre_supervisor:
                st.error("Ingrese su nombre para firmar.")
              else:
                with st.spinner("Generando y unificando el PDF de Auditoría..."):
                  carpeta_base = "./Registros_Oficiales"
                  subcarpeta_lote = f"Lote_{row['LoteProveedor']}_{row['NombreMaterial']}".replace(" ", "_")
                  ruta_destino = os.path.join(carpeta_base, str(row["Planta"]), str(row["TipoMaterial"]), str(row["NombreMaterial"]), subcarpeta_lote)
                  os.makedirs(ruta_destino, exist_ok=True)

                  ruta_firma_sup = ""
                  if canvas_sup is not None:
                      try:
                          if canvas_sup.image_data is not None and np.any(canvas_sup.image_data):
                              os.makedirs("temp_firmas", exist_ok=True)
                              ruta_firma_sup = f"temp_firmas/firma_sup_{row['ID']}_{datetime.now().strftime('%Y%m%d%H%M%S')}.png"
                              img = Image.fromarray(canvas_sup.image_data.astype('uint8'), 'RGBA').convert('RGB')
                              img.save(ruta_firma_sup)
                      except Exception:
                          pass
                  
                  row['RutaFirmaSup'] = ruta_firma_sup

                  ext_cert = os.path.splitext(str(row["RutaPdfCert"]))[1] if row["RutaPdfCert"] else ".pdf"
                  ruta_cert_final = os.path.join(ruta_destino, f"Certificado_Original_{row['LoteProveedor']}{ext_cert}")
                  if row["RutaPdfCert"] and os.path.exists(str(row["RutaPdfCert"])):
                    shutil.copy(str(row["RutaPdfCert"]), ruta_cert_final)

                  # 1. Generar Checklist Base (PDF Temporal)
                  ruta_checklist_temp = os.path.join(ruta_destino, f"TEMP_Checklist_{row['LoteProveedor']}.pdf")
                  generar_pdf_ricc14(row, nombre_supervisor, ruta_checklist_temp)

                  # 2. Unificar Documentos (Checklist + Certificado) en un solo archivo PDF
                  ruta_unificada = os.path.join(ruta_destino, f"Registro_Auditoria_Completo_{row['LoteProveedor']}.pdf")
                  generar_documento_unificado(ruta_checklist_temp, ruta_cert_final, ruta_unificada)
                  
                  # Opcional: Eliminar temporal
                  if os.path.exists(ruta_checklist_temp) and os.path.exists(ruta_unificada):
                      os.remove(ruta_checklist_temp)

                  df_p.loc[df_p["ID"] == row["ID"], "Estado"] = "Aprobado"
                  df_p.loc[df_p["ID"] == row["ID"], "FirmaCalidad"] = str(nombre_supervisor)
                  df_p.loc[df_p["ID"] == row["ID"], "RutaCarpeta"] = str(ruta_destino)
                  df_p.loc[df_p["ID"] == row["ID"], "RutaChecklistOficial"] = str(ruta_unificada)
                  df_p.loc[df_p["ID"] == row["ID"], "RutaPdfCert"] = str(ruta_cert_final)
                  df_p.loc[df_p["ID"] == row["ID"], "RutaFirmaSup"] = str(ruta_firma_sup)
                  guardar_pendientes(df_p)

                  st.success("🎉 ¡Aprobado! Checklist y Certificado unificados en un solo PDF.")
                  st.rerun()

          with col_btn2:
            if st.button("🗑️ Eliminar Registro", key=f"del_pend_{row['ID']}"):
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

    st.dataframe(resultado[["Planta", "NombreMaterial", "FacturaGuia", "LoteProveedor", "Fecha", "FirmaCalidad"]], use_container_width=True)

    for idx, row in resultado.iterrows():
      with st.expander(f"📁 {row['NombreMaterial']} | Lote: {row['LoteProveedor']} | Supervisor: {row['FirmaCalidad']}"):
        col_d1, col_d2, col_d3 = st.columns(3)
        ruta_chk = str(row.get("RutaChecklistOficial", ""))
        with col_d1:
          if pd.notna(ruta_chk) and ruta_chk != "nan" and os.path.exists(ruta_chk):
            with open(ruta_chk, "rb") as f:
              st.download_button("📊 Descargar PDF Unificado (Checklist + Certificado)", data=f, file_name=os.path.basename(ruta_chk), mime="application/pdf", key=f"dl_chk_{row['ID']}")
        
        ruta_cert = str(row.get("RutaPdfCert", ""))
        with col_d2:
          if pd.notna(ruta_cert) and ruta_cert != "nan" and os.path.exists(ruta_cert):
            with open(ruta_cert, "rb") as f:
              st.download_button("📄 Descargar Certificado Original", data=f, file_name=os.path.basename(ruta_cert), mime="application/pdf" if ruta_cert.endswith(".pdf") else "image/jpeg", key=f"dl_cert_{row['ID']}")
        
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
  df_p = cargar_pendientes()
  aprobados_excel = df_p[df_p["Estado"] == "Aprobado"].copy() if not df_p.empty else pd.DataFrame()

  if len(aprobados_excel) == 0:
    st.info("No hay información suficiente.")
  else:
    columnas_reporte = [
        "Planta", "TipoMaterial", "NombreMaterial", "Proveedor", "LoteProveedor", 
        "FacturaGuia", "Placa", "IngresoSAP", "DocSAP", "EstadoMarchamo", "Marchamo",
        "Cantidad", "Unidad", "TempTransporte", "TempMaterial", "Criticidad", "Fecha", "Responsable", "FirmaCalidad"
    ]
    df_reporte = aprobados_excel[columnas_reporte].copy()
    st.dataframe(df_reporte, use_container_width=True)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_reporte.to_excel(writer, index=False, sheet_name='Historial')
    st.download_button("📥 Descargar Base Completa (Excel)", data=output.getvalue(), file_name=f"Control_LIF_{datetime.now().strftime('%Y%m%d')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")
