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
    from pypdf import PdfWriter
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
# FUNCIONES DE DATOS Y FIRMAS
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

def guardar_firma(canvas_result, prefix):
    if canvas_result is not None:
        try:
            data = canvas_result.image_data
            if data is not None and np.any(data):
                os.makedirs("temp_firmas", exist_ok=True)
                ruta_firma = f"temp_firmas/{prefix}_{datetime.now().strftime('%Y%m%d%H%M%S')}.png"
                
                rgba_image = Image.fromarray(data.astype('uint8'), 'RGBA')
                background = Image.new('RGBA', rgba_image.size, (255, 255, 255, 255))
                alpha_composite = Image.alpha_composite(background, rgba_image)
                rgb_image = alpha_composite.convert('RGB')
                
                rgb_image.save(ruta_firma, "PNG")
                return ruta_firma
        except Exception:
            pass
    return ""

def generar_documento_unificado(ruta_checklist, ruta_certificado, ruta_final):
    try:
        from pypdf import PdfWriter
        merger = PdfWriter()
        merger.append(ruta_checklist)
        
        if ruta_certificado and os.path.exists(ruta_certificado):
            if ruta_certificado.lower().endswith(('.png', '.jpg', '.jpeg')):
                img = Image.open(ruta_certificado).convert('RGB')
                ruta_temp_pdf = ruta_certificado + "_temp.pdf"
                img.save(ruta_temp_pdf, "PDF", resolution=100.0)
                merger.append(ruta_temp_pdf)
                merger.write(ruta_final)
                merger.close()
                os.remove(ruta_temp_pdf)
            elif ruta_certificado.lower().endswith('.pdf'):
                merger.append(ruta_certificado)
                merger.write(ruta_final)
                merger.close()
        else:
            merger.write(ruta_final)
            merger.close()
    except Exception:
        shutil.copy(ruta_checklist, ruta_final)

# ------------------------------------------
# MOTOR DE GENERACIÓN PDF (DISEÑO MILIMÉTRICO)
# ------------------------------------------
def dibujar_opciones_cabecera(pdf, x_start, y, h_row, opts, idx_checked):
    """Distribución limpia y exacta para los bloques superiores de Marchamo y Candado."""
    curr_x = x_start
    for i, opt in enumerate(opts):
        pdf.set_font('Arial', '', 7)
        pdf.set_xy(curr_x, y)
        pdf.cell(7, h_row, opt, 0, 0, 'R')
        box_x = curr_x + 8.5
        box_y = y + (h_row - 3) / 2
        pdf.rect(box_x, box_y, 3, 3)
        if i == idx_checked:
            pdf.set_xy(box_x, box_y)
            pdf.set_font('Arial', 'B', 7)
            pdf.cell(3, 3, 'X', 0, 0, 'C')
        curr_x += 15

def dibujar_opciones_bloque(pdf, x_start, y, h_row, opts, idx_checked):
    """Función centrada y uniforme para las preguntas principales de la tabla."""
    n = len(opts)
    ancho_total = 55.0
    espacio_item = ancho_total / n
    for i, opt in enumerate(opts):
        sub_x = x_start + (i * espacio_item)
        pdf.set_font('Arial', '', 7)
        w_texto = pdf.get_string_width(opt)
        text_x = sub_x + (espacio_item - (w_texto + 4.5)) / 2
        pdf.set_xy(text_x, y)
        pdf.cell(w_texto, h_row, opt, 0, 0, 'L')
        
        box_x = text_x + w_texto + 1.5
        box_y = y + (h_row - 3) / 2
        pdf.rect(box_x, box_y, 3, 3)
        
        if i == idx_checked:
            pdf.set_xy(box_x, box_y)
            pdf.set_font('Arial', 'B', 7)
            pdf.cell(3, 3, 'X', 0, 0, 'C')

def generar_pdf_ricc14(row, supervisor, ruta_guardado_temp):
    pdf = FPDF(orientation='L', unit='mm', format='A4')
    pdf.add_page()
    pdf.set_auto_page_break(auto=False)
    pdf.set_fill_color(220, 230, 241) 
    
    h_row = 4.5
    
    # --- ENCABEZADO ---
    pdf.set_font('Arial', 'B', 7)
    pdf.cell(35, h_row, 'Código:', 'LT', 0, 'R')
    pdf.set_font('Arial', '', 7)
    pdf.cell(35, h_row, 'R ICC - 14-1', 'TR', 0, 'C')
    pdf.set_font('Arial', 'B', 7)
    pdf.cell(167, h_row, 'Aseguramiento de Calidad', 'LTR', 0, 'C')
    pdf.set_font('Arial', '', 7)
    pdf.cell(40, h_row, 'Vigente a partir de:', 'LTR', 1, 'C')
    
    pdf.set_font('Arial', 'B', 7)
    pdf.cell(35, h_row, 'Versión:', 'LB', 0, 'R')
    pdf.set_font('Arial', '', 7)
    pdf.cell(35, h_row, '5', 'BR', 0, 'C')
    pdf.set_font('Arial', 'B', 9)
    pdf.cell(167, h_row, 'Check list de recepción de materiales', 'LBR', 0, 'C')
    pdf.set_font('Arial', '', 7)
    pdf.cell(40, h_row, 'feb-23', 'LBR', 1, 'C')
    pdf.ln(1)

    # --- DATOS GENERALES ---
    pdf.set_font('Arial', 'B', 7)
    pdf.cell(35, h_row, 'Fecha de recepción:', 1, 0, 'R')
    pdf.set_font('Arial', '', 7)
    pdf.cell(60, h_row, str(row.get('Fecha', '')).split()[0], 1, 0, 'C')
    pdf.cell(15, h_row, '', 0, 0)
    pdf.set_font('Arial', 'B', 7)
    pdf.cell(67, h_row, 'Número de factura y/o guía:', 1, 0, 'R')
    pdf.set_font('Arial', '', 7)
    pdf.cell(100, h_row, str(row.get('FacturaGuia', '')), 1, 1, 'C')

    pdf.set_font('Arial', 'B', 7)
    pdf.cell(35, h_row, 'Nombre del transporte:', 1, 0, 'R')
    pdf.set_font('Arial', '', 7)
    pdf.cell(60, h_row, str(row.get('Transporte', '')), 1, 0, 'C')
    pdf.cell(15, h_row, '', 0, 0)
    pdf.set_font('Arial', 'B', 7)
    pdf.cell(67, h_row, 'Fecha de Ingreso a SAP:', 1, 0, 'R')
    pdf.set_font('Arial', '', 7)
    pdf.cell(100, h_row, str(row.get('IngresoSAP', '')), 1, 1, 'C')

    pdf.set_font('Arial', 'B', 7)
    pdf.cell(35, h_row, 'Placa del transporte:', 1, 0, 'R')
    pdf.set_font('Arial', '', 7)
    pdf.cell(60, h_row, str(row.get('Placa', '')), 1, 0, 'C')
    pdf.cell(15, h_row, '', 0, 0)
    pdf.set_font('Arial', 'B', 7)
    pdf.cell(67, h_row, 'Número de documento SAP:', 1, 0, 'R')
    pdf.set_font('Arial', '', 7)
    pdf.cell(100, h_row, str(row.get('DocSAP', '')), 1, 1, 'C')
    pdf.ln(2)

    # --- MARCHAMO Y CANDADO ---
    pdf.set_font('Arial', 'B', 7)
    pdf.cell(100, h_row, 'Inspección del Marchamo', 0, 0, 'L', fill=True)
    pdf.cell(5, h_row, '', 0, 0)
    pdf.cell(85, h_row, 'Inspección del Candado', 0, 0, 'L', fill=True)
    pdf.cell(30, h_row, 'Valor esperado', 0, 0, 'C')
    pdf.cell(57, h_row, 'Acción en caso desviación', 0, 1, 'C')

    y_start = pdf.get_y()
    pdf.rect(10, y_start, 100, h_row*2)
    pdf.rect(115, y_start, 85, h_row*2)
    pdf.rect(200, y_start, 30, h_row*2)
    pdf.rect(230, y_start, 57, h_row*2)

    # Marchamo (Pregunta 1)
    pdf.set_xy(10, y_start)
    pdf.cell(5, h_row, '1', 0, 0, 'C')
    pdf.cell(42, h_row, '¿El marchamo se encuentra alterado?', 0, 0)
    
    est_mar = str(row.get('EstadoMarchamo', 'N/A'))
    idx_m = 0 if est_mar == 'SI' else (1 if est_mar == 'NO' else 2)
    dibujar_opciones_cabecera(pdf, 54, y_start, h_row, ['SI', 'NO', 'N/A'], idx_m)

    pdf.set_xy(15, y_start + h_row)
    pdf.cell(25, h_row, 'No. de marchamo:', 0, 0)
    pdf.cell(55, h_row, str(row.get('Marchamo', '')), 'B', 0, 'C')

    # Candado (Pregunta 2) - Ajustado x_start a 155 para que N/A entre perfectamente
    pdf.set_xy(115, y_start + h_row/2)
    pdf.cell(5, h_row, '2', 0, 0, 'C')
    pdf.cell(42, h_row, '¿El candado se encuentra alterado?', 0, 0)
    est_can = str(row.get('EstadoCandado', 'N/A'))
    idx_c = 0 if est_can == 'SI' else (1 if est_can == 'NO' else 2)
    dibujar_opciones_cabecera(pdf, 155, y_start + h_row/2, h_row, ['SI', 'NO', 'N/A'], idx_c)

    pdf.set_xy(200, y_start + h_row/2)
    pdf.set_font('Arial', 'B', 7)
    pdf.cell(30, h_row, 'NO', 0, 0, 'C')
    pdf.set_font('Arial', '', 7)
    pdf.cell(57, h_row, 'Retener lote y emitir reclamo.', 0, 1, 'C')
    pdf.ln(h_row/2 + 2)

    # --- MATERIAL ---
    pdf.set_font('Arial', 'B', 7)
    pdf.cell(190, h_row, 'Inspección del Material', 0, 0, 'L', fill=True)
    pdf.cell(30, h_row, 'Valor esperado', 0, 0, 'C', fill=True)
    pdf.cell(57, h_row, 'Acción en caso desviación', 0, 1, 'C', fill=True)

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
        pdf.cell(35, h_row, etiqueta, 1, 0, 'R')
        pdf.set_font('Arial', '', 7)
        pdf.cell(155, h_row, valor, 1, 0, 'C')
        pdf.set_fill_color(200, 200, 200)
        pdf.cell(30, h_row, '------', 1, 0, 'C', fill=True)
        pdf.cell(57, h_row, '------', 1, 1, 'C', fill=True)

    pdf.set_fill_color(220, 230, 241)
    preguntas_mat = [
        ("3", "¿El (los) lote(s) posee(n) más de un mes de vida?", ["SI", "NO", "N/A"], "SI", "Retener lote y emitir reclamo"),
        ("4", "¿A qué categoría de la matriz pertenece?", ["I", "II", "III"], "------", "------"),
        ("5", "¿El Certificado de Análisis corresponde al lote?", ["SI", "NO", "N/A"], "SI", "Rechazar lote y emitir reclamo"),
        ("6", "¿Los resultados del Certificado cumplen parámetros críticos?", ["SI", "NO", "N/A"], "SI", "Rechazar lote y emitir reclamo"),
        ("7", "¿El empaque se encuentra íntegro y sin deterioro?", ["SI", "NO", "N/A"], "SI", "Rechazar lote y emitir reclamo"),
        ("8", "¿El material está protegido de la intemperie?", ["SI", "NO", "N/A"], "SI", "Rechazar lote y emitir reclamo"),
        ("9", "¿Coincide la entrega según documento contra lo físico?", ["SI", "NO", "N/A"], "SI", "Rechazar lote y emitir reclamo")
    ]

    criticidad = str(row.get('Criticidad', 'I'))

    for num, preg, opts, val_esp, accion in preguntas_mat:
        y_curr = pdf.get_y()
        pdf.set_xy(10, y_curr)
        pdf.set_font('Arial', 'B', 7)
        pdf.cell(5, h_row, num, 1, 0, 'C')
        pdf.set_font('Arial', '', 7)
        pdf.cell(130, h_row, preg, 1, 0, 'L')
        pdf.cell(55, h_row, '', 1, 0)
        
        idx_chk = 0 
        if num == "4":
            idx_chk = 0 if criticidad=="I" else (1 if criticidad=="II" else 2)
            
        dibujar_opciones_bloque(pdf, 145, y_curr, h_row, opts, idx_chk)

        pdf.set_xy(200, y_curr)
        pdf.set_font('Arial', 'B', 7)
        if val_esp == "------":
            pdf.set_fill_color(200, 200, 200)
            pdf.cell(30, h_row, val_esp, 1, 0, 'C', fill=True)
            pdf.cell(57, h_row, accion, 1, 1, 'C', fill=True)
            pdf.set_fill_color(220, 230, 241)
        else:
            pdf.cell(30, h_row, val_esp, 1, 0, 'C')
            pdf.set_font('Arial', '', 7)
            pdf.cell(57, h_row, accion, 1, 1, 'C')

    pdf.ln(1)

    # --- TRANSPORTE ---
    pdf.set_font('Arial', 'B', 7)
    pdf.cell(190, h_row, 'Inspección del Transporte', 0, 0, 'L', fill=True)
    pdf.cell(87, h_row, '', 0, 1)

    preguntas_trans = [
        ("10", "¿Se encuentra el transporte libre de plagas?", ["SI", "NO"], "SI", "Rechazar lote y emitir reclamo"),
        ("11", "¿El techo, paredes o lona se encuentran íntegras?", ["SI", "NO"], "SI", "Rechazar lote y emitir reclamo"),
        ("12", "¿El techo, paredes o lona se encuentran limpios?", ["SI", "NO"], "SI", "Rechazar lote y emitir reclamo"),
        ("13", "¿El piso del transporte está limpio y en buenas condiciones?", ["SI", "NO"], "SI", "Emitir reclamo")
    ]

    for num, preg, opts, val_esp, accion in preguntas_trans:
        y_curr = pdf.get_y()
        pdf.set_xy(10, y_curr)
        pdf.set_font('Arial', 'B', 7)
        pdf.cell(5, h_row, num, 1, 0, 'C')
        pdf.set_font('Arial', '', 7)
        pdf.cell(130, h_row, preg, 1, 0, 'L')
        pdf.cell(55, h_row, '', 1, 0)
        
        dibujar_opciones_bloque(pdf, 145, y_curr, h_row, opts, 0)

        pdf.set_xy(200, y_curr)
        pdf.set_font('Arial', 'B', 7)
        pdf.cell(30, h_row, val_esp, 1, 0, 'C')
        pdf.set_font('Arial', '', 7)
        pdf.cell(57, h_row, accion, 1, 1, 'C')
        
    # --- TEMPERATURAS ---
    temp_trans = str(row.get('TempTransporte', 'N/A'))
    temp_mat = str(row.get('TempMaterial', 'N/A'))
    
    y_curr = pdf.get_y()
    pdf.set_xy(10, y_curr)
    pdf.set_font('Arial', 'B', 7)
    pdf.cell(5, h_row, '14', 'LTR', 0, 'C')
    pdf.set_font('Arial', '', 7)
    pdf.cell(130, h_row, '¿El transporte cumple con la temperatura especificada?', 'LTR', 0, 'L')
    pdf.cell(55, h_row, '', 'LTR', 0)
    
    idx_temp = 2 if (temp_trans == "N/A" and temp_mat == "N/A") else 0
    dibujar_opciones_bloque(pdf, 145, y_curr, h_row, ["SI", "NO", "N/A"], idx_temp)
    
    pdf.set_xy(200, y_curr)
    pdf.set_font('Arial', 'B', 7)
    pdf.cell(30, h_row, "SI", 'LTR', 0, 'C')
    pdf.set_font('Arial', '', 7)
    pdf.cell(57, h_row, "Rechazar lote", 'LTR', 1, 'C')

    # Fila combinada de temperaturas
    y_curr = pdf.get_y()
    pdf.set_xy(10, y_curr)
    pdf.cell(5, h_row, '', 'LBR', 0)
    pdf.set_font('Arial', 'B', 7)
    pdf.cell(130, h_row, f"Temperatura Transporte (°C): {temp_trans}           Temperatura del Material (°C): {temp_mat}", 'BR', 0, 'C')
    pdf.cell(55, h_row, '', 'BR', 0)
    pdf.cell(30, h_row, '', 'BR', 0)
    pdf.cell(57, h_row, '', 'BR', 1)

    # Pregunta 15
    y_curr = pdf.get_y()
    pdf.set_xy(10, y_curr)
    pdf.set_font('Arial', 'B', 7)
    pdf.cell(5, h_row, '15', 1, 0, 'C')
    pdf.set_font('Arial', '', 7)
    pdf.cell(130, h_row, '¿El transporte se encuentra libre de olores extraños?', 1, 0, 'L')
    pdf.cell(55, h_row, '', 1, 0)
    
    dibujar_opciones_bloque(pdf, 145, y_curr, h_row, ["SI", "NO"], 0) 
    
    pdf.set_xy(200, y_curr)
    pdf.set_font('Arial', 'B', 7)
    pdf.cell(30, h_row, "SI", 1, 0, 'C')
    pdf.set_font('Arial', '', 7)
    pdf.cell(57, h_row, "Rechazar lote", 1, 1, 'C')

    pdf.ln(2)
    
    # --- RESULTADO Y OBSERVACIONES ---
    y_curr = pdf.get_y()
    pdf.rect(10, y_curr, 277, 8)
    pdf.set_xy(10, y_curr+2)
    pdf.set_font('Arial', 'B', 9)
    pdf.cell(70, 4, 'Resultado final:', 0, 0, 'R')
    dibujar_opciones_bloque(pdf, 85, y_curr+2, 4, ["Aprobado", "Desaprobado"], 0)
    
    pdf.ln(6)
    pdf.set_font('Arial', 'B', 7)
    pdf.cell(277, 3, 'OBSERVACIONES:', 0, 1, 'L')
    pdf.cell(277, 3, '** Si el marchamo N/A por ser proveedores locales, basta con la inspección del material y del transporte.', 0, 1, 'L')
    pdf.cell(277, 3, '** Para los casos que se realiza la revisión en el transporte y determina que no cumple, PROCEDE A RECHAZARLOS.', 0, 1, 'L')
    
    # --- ESPACIO DE FIRMAS ---
    pdf.ln(12) 
    y_linea1 = pdf.get_y()

    # 1. Primera línea (Línea para el nombre)
    pdf.line(55, y_linea1, 135, y_linea1)
    pdf.line(175, y_linea1, 255, y_linea1)

    pdf.set_xy(55, y_linea1 - 4)
    pdf.set_font('Arial', '', 8)
    pdf.cell(80, 4, str(row.get('Responsable', '')), 0, 0, 'C')
    
    pdf.set_xy(175, y_linea1 - 4)
    pdf.cell(80, 4, supervisor, 0, 0, 'C')

    # Ajustado: Etiqueta movida a x=5 para que no se corte y se lea completa
    pdf.set_xy(5, y_linea1 - 1)
    pdf.set_font('Arial', 'B', 7)
    pdf.cell(48, 4, 'Nombre de la persona que realiza la recepción:', 0, 0, 'R')
    
    pdf.set_xy(145, y_linea1 - 1)
    pdf.set_font('Arial', 'B', 8)
    pdf.cell(28, 4, 'Vo. Bo. Calidad:', 0, 0, 'R')

    # 2. Segunda línea (Línea para la Firma)
    y_linea2 = y_linea1 + 10
    pdf.line(55, y_linea2, 135, y_linea2)
    pdf.line(175, y_linea2, 255, y_linea2)

    pdf.set_xy(35, y_linea2 - 1)
    pdf.cell(18, 4, 'Firma:', 0, 0, 'R')
    
    pdf.set_xy(155, y_linea2 - 1)
    pdf.cell(18, 4, 'Firma:', 0, 0, 'R')

    ruta_firma_op = str(row.get('RutaFirmaOp', ''))
    if ruta_firma_op and os.path.exists(ruta_firma_op):
        try:
            pdf.image(ruta_firma_op, x=70, y=y_linea2 - 12, w=50, h=11)
        except Exception: pass
        
    ruta_firma_sup = str(row.get('RutaFirmaSup', ''))
    if ruta_firma_sup and os.path.exists(ruta_firma_sup):
        try:
            pdf.image(ruta_firma_sup, x=190, y=y_linea2 - 12, w=50, h=11)
        except Exception: pass

    pdf.output(ruta_guardado_temp)


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
          canvas_op = st_canvas(stroke_width=2, stroke_color="#000", background_color="#ffffff", height=100, width=400, drawing_mode="freedraw", key="firma_op")

      if st.button("🚀 Enviar a Calidad", type="primary"):
        if not responsable:
          st.error("⚠️ Ingresa el nombre del operador responsable.")
        else:
          ruta_firma_op = guardar_firma(canvas_op, "firma_op")

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
          st.write(f"**Operador Responsable:** {row['Responsable']}")
          if row["RutaPdfCert"] and os.path.exists(str(row["RutaPdfCert"])):
            st.write("📄 Certificado adjunto.")

          nombre_supervisor = st.text_input("Nombre del Supervisor", key=f"sup_nom_{row['ID']}")
          st.write("✍️ **Dibuja tu Firma de Supervisor:**")
          
          canvas_sup = None
          if st_canvas:
              canvas_sup = st_canvas(stroke_width=2, stroke_color="#000", background_color="#ffffff", height=100, width=400, drawing_mode="freedraw", key=f"firma_sup_{row['ID']}")

          col_btn1, col_btn2 = st.columns(2)
          with col_btn1:
            if st.button("Aprobar y Generar Documento Maestro 🛡️", key=f"btn_apr_{row['ID']}", type="primary"):
              if not nombre_supervisor:
                st.error("Ingrese su nombre para firmar.")
              else:
                with st.spinner("Generando y unificando el PDF de Auditoría..."):
                  carpeta_base = "./Registros_Oficiales"
                  subcarpeta_lote = f"Lote_{row['LoteProveedor']}_{row['NombreMaterial']}".replace(" ", "_")
                  ruta_destino = os.path.join(carpeta_base, str(row["Planta"]), str(row["TipoMaterial"]), str(row["NombreMaterial"]), subcarpeta_lote)
                  os.makedirs(ruta_destino, exist_ok=True)

                  ruta_firma_sup = guardar_firma(canvas_sup, f"firma_sup_{row['ID']}")
                  row['RutaFirmaSup'] = ruta_firma_sup

                  ext_cert = os.path.splitext(str(row["RutaPdfCert"]))[1] if row["RutaPdfCert"] else ".pdf"
                  ruta_cert_final = os.path.join(ruta_destino, f"Certificado_{row['LoteProveedor']}{ext_cert}")
                  if row["RutaPdfCert"] and os.path.exists(str(row["RutaPdfCert"])):
                    shutil.copy(str(row["RutaPdfCert"]), ruta_cert_final)

                  # 1. Generar Checklist
                  ruta_checklist_temp = os.path.join(ruta_destino, f"TEMP_Checklist_{row['LoteProveedor']}.pdf")
                  generar_pdf_ricc14(row, nombre_supervisor, ruta_checklist_temp)

                  # 2. Unificar
                  ruta_unificada = os.path.join(ruta_destino, f"Registro_Auditoria_{row['LoteProveedor']}.pdf")
                  generar_documento_unificado(ruta_checklist_temp, ruta_cert_final, ruta_unificada)
                  
                  if os.path.exists(ruta_checklist_temp) and os.path.exists(ruta_unificada):
                      os.remove(ruta_checklist_temp)

                  df_p.loc[df_p["ID"] == row["ID"], "Estado"] = "Aprobado"
                  df_p.loc[df_p["ID"] == row["ID"], "FirmaCalidad"] = str(nombre_supervisor)
                  df_p.loc[df_p["ID"] == row["ID"], "RutaCarpeta"] = str(ruta_destino)
                  df_p.loc[df_p["ID"] == row["ID"], "RutaChecklistOficial"] = str(ruta_unificada)
                  df_p.loc[df_p["ID"] == row["ID"], "RutaPdfCert"] = str(ruta_cert_final)
                  df_p.loc[df_p["ID"] == row["ID"], "RutaFirmaSup"] = str(ruta_firma_sup)
                  guardar_pendientes(df_p)

                  st.success("🎉 ¡Aprobado! Documento maestro listo y carpetas creadas.")
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
    col_f1, col_f2 = st.columns(2)
    with col_f1:
      busqueda = st.text_input("Filtro rápido (Lote, Material o Proveedor)")
    with col_f2:
      fecha_filtro = st.date_input("Filtrar por Fecha de Ingreso a SAP", value=None)

    resultado = aprobados.copy()
    if busqueda:
      resultado = resultado[
          resultado["LoteProveedor"].astype(str).str.contains(busqueda, case=False, na=False) |
          resultado["NombreMaterial"].astype(str).str.contains(busqueda, case=False, na=False) |
          resultado["Proveedor"].astype(str).str.contains(busqueda, case=False, na=False)
      ]
    if fecha_filtro:
      resultado = resultado[resultado["IngresoSAP"] == str(fecha_filtro)]

    st.dataframe(resultado[["Planta", "NombreMaterial", "FacturaGuia", "LoteProveedor", "IngresoSAP", "FirmaCalidad"]], use_container_width=True)

    for idx, row in resultado.iterrows():
      with st.expander(f"📁 {row['NombreMaterial']} | Lote: {row['LoteProveedor']} | Fecha SAP: {row['IngresoSAP']}"):
        col_d1, col_d2, col_d3 = st.columns(3)
        ruta_chk = str(row.get("RutaChecklistOficial", ""))
        with col_d1:
          if pd.notna(ruta_chk) and ruta_chk != "nan" and os.path.exists(ruta_chk):
            with open(ruta_chk, "rb") as f:
              st.download_button("📊 Descargar PDF Maestro Unificado", data=f, file_name=os.path.basename(ruta_chk), mime="application/pdf", key=f"dl_chk_{row['ID']}")
        
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
