from datetime import datetime
import os
import shutil
import pandas as pd
import numpy as np
from PIL import Image
import streamlit as st
from fpdf import FPDF
import io
import smtplib
import sqlite3
import zipfile
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

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
st.set_page_config(page_title="Control de Recepción y Calidad - LIF", page_icon="📦", layout="wide", initial_sidebar_state="expanded")

st.markdown("""
    <style>
    .main { background-color: #f8f9fa; }
    .stTabs [data-baseweb="tab-list"] { gap: 10px; }
    .stTabs [data-baseweb="tab"] { height: 50px; white-space: pre-wrap; background-color: #ffffff; border-radius: 5px; font-weight: 600; padding: 10px 20px; box-shadow: 0 2px 4px rgba(0,0,0,0.05); }
    .stTabs [aria-selected="true"] { background-color: #0066cc !important; color: #ffffff !important; }
    </style>
""", unsafe_allow_html=True)

DB_NAME = "registros_lif.db"
ARCHIVO_CATALOGO = "Maestro_Materiales_Proveedores (2).xlsx"
CONTRASENA_SISTEMA = "Lif1012"
CORREO_NOTIFICACION = "carreaga@lifbrands.co"

# Inicializar Base de Datos SQLite (Persistencia Definitiva)
def inicializar_db():
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS registros (
            ID INTEGER PRIMARY KEY AUTOINCREMENT,
            Planta TEXT,
            TipoMaterial TEXT,
            NombreMaterial TEXT,
            Proveedor TEXT,
            LoteProveedor TEXT,
            FacturaGuia TEXT,
            Transporte TEXT,
            Placa TEXT,
            IngresoSAP TEXT,
            DocSAP TEXT,
            EstadoMarchamo TEXT,
            Marchamo TEXT,
            EstadoCandado TEXT,
            EstadoTarima TEXT,
            Cantidad REAL,
            Unidad TEXT,
            TempTransporte TEXT,
            TempMaterial TEXT,
            FechaProduccion TEXT,
            FechaVencimiento TEXT,
            Criticidad TEXT,
            Responsable TEXT,
            Fecha TEXT,
            Estado TEXT,
            FirmaCalidad TEXT,
            RutaCarpeta TEXT,
            RutaChecklistOficial TEXT,
            RutaPdfCert TEXT,
            RutaFirmaOp TEXT,
            RutaFirmaSup TEXT
        )
    ''')
    conn.commit()
    conn.close()

inicializar_db()

# ------------------------------------------
# INTERFAZ DE AUTENTICACIÓN INICIAL (LIF)
# ------------------------------------------
if "autenticado" not in st.session_state:
    st.session_state.autenticado = False

if not st.session_state.autenticado:
    st.markdown("""
        <div style='text-align: center; padding: 40px 0;'>
            <h1 style='color: #0066cc; font-size: 3.5rem; margin-bottom: 0; font-weight: 800;'>LIF</h1>
            <p style='color: #555; font-size: 1.2rem; font-weight: 600;'>Sistema de Trazabilidad y Control de Calidad</p>
        </div>
    """, unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns([1, 1.5, 1])
    with col2:
        with st.form("form_login_lif"):
            st.markdown("### 🔐 Acceso Restringido")
            password_ingresada = st.text_input("Contraseña del Sistema", type="password", placeholder="Ingrese la contraseña")
            submit_login = st.form_submit_button("Ingresar al Sistema", use_container_width=True)
            
            if submit_login:
                if password_ingresada == CONTRASENA_SISTEMA:
                    st.session_state.autenticado = True
                    st.success("¡Acceso concedido!")
                    st.rerun()
                else:
                    st.error("❌ Contraseña incorrecta. Intente nuevamente.")
    
    st.stop()

st.title("📦 Sistema de Trazabilidad y Control de Calidad - LIF")
st.markdown("---")

# ------------------------------------------
# FUNCIONES DE DATOS CON SQLITE
# ------------------------------------------
def enviar_notificacion_outlook(asunto, cuerpo_mensaje):
    remitente = CORREO_NOTIFICACION
    password = "tu_contraseña_o_app_password"
    destinatario = CORREO_NOTIFICACION
    try:
        mensaje = MIMEMultipart()
        mensaje["From"] = remitente
        mensaje["To"] = destinatario
        mensaje["Subject"] = asunto
        mensaje.attach(MIMEText(cuerpo_mensaje, "plain"))
        
        servidor = smtplib.SMTP("smtp.office365.com", 587)
        servidor.starttls()
        servidor.login(remitente, password)
        servidor.sendmail(remitente, destinatario, mensaje.as_string())
        servidor.quit()
        return True
    except Exception as e:
        print(f"Error al enviar correo: {e}")
        return False

def cargar_pendientes():
    conn = sqlite3.connect(DB_NAME)
    df = pd.read_sql_query("SELECT * FROM registros ORDER BY IngresoSAP DESC", conn)
    conn.close()
    
    columnas_texto = [
        "Planta", "TipoMaterial", "NombreMaterial", "Proveedor",
        "LoteProveedor", "FacturaGuia", "Transporte", "Placa",
        "IngresoSAP", "DocSAP", "EstadoMarchamo", "Marchamo", "EstadoCandado",
        "Cantidad", "Unidad", "TempTransporte", "TempMaterial", "FechaProduccion", "FechaVencimiento",
        "Criticidad", "Responsable", "Fecha", "Estado", "FirmaCalidad",
        "RutaCarpeta", "RutaChecklistOficial", "RutaPdfCert", "RutaFirmaOp", "RutaFirmaSup", "EstadoTarima"
    ]
    for col in columnas_texto:
        if col not in df.columns:
            df[col] = ""
        else:
            df[col] = df[col].astype(str)
            
    return df

def guardar_registro_db(nuevo_registro):
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute('''
        INSERT INTO registros (
            Planta, TipoMaterial, NombreMaterial, Proveedor, LoteProveedor,
            FacturaGuia, Transporte, Placa, IngresoSAP, DocSAP, EstadoMarchamo,
            Marchamo, EstadoCandado, EstadoTarima, Cantidad, Unidad, TempTransporte,
            TempMaterial, FechaProduccion, FechaVencimiento, Criticidad, Responsable,
            Fecha, Estado, FirmaCalidad, RutaCarpeta, RutaChecklistOficial,
            RutaPdfCert, RutaFirmaOp, RutaFirmaSup
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    ''', (
        nuevo_registro.get("Planta"), nuevo_registro.get("TipoMaterial"), nuevo_registro.get("NombreMaterial"),
        nuevo_registro.get("Proveedor"), nuevo_registro.get("LoteProveedor"), nuevo_registro.get("FacturaGuia"),
        nuevo_registro.get("Transporte"), nuevo_registro.get("Placa"), nuevo_registro.get("IngresoSAP"),
        nuevo_registro.get("DocSAP"), nuevo_registro.get("EstadoMarchamo"), nuevo_registro.get("Marchamo"),
        nuevo_registro.get("EstadoCandado"), nuevo_registro.get("EstadoTarima"), nuevo_registro.get("Cantidad"),
        nuevo_registro.get("Unidad"), nuevo_registro.get("TempTransporte"), nuevo_registro.get("TempMaterial"),
        nuevo_registro.get("FechaProduccion"), nuevo_registro.get("FechaVencimiento"), nuevo_registro.get("Criticidad"),
        nuevo_registro.get("Responsable"), nuevo_registro.get("Fecha"), nuevo_registro.get("Estado"),
        nuevo_registro.get("FirmaCalidad"), nuevo_registro.get("RutaCarpeta"), nuevo_registro.get("RutaChecklistOficial"),
        nuevo_registro.get("RutaPdfCert"), nuevo_registro.get("RutaFirmaOp"), nuevo_registro.get("RutaFirmaSup")
    ))
    conn.commit()
    conn.close()

def actualizar_registro_db(id_reg, campos_actualizados):
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    for campo, valor in campos_actualizados.items():
        cursor.execute(f"UPDATE registros SET {campo} = ? WHERE ID = ?", (valor, id_reg))
    conn.commit()
    conn.close()

def eliminar_registro_db(id_reg):
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute("DELETE FROM registros WHERE ID = ?", (id_reg,))
    conn.commit()
    conn.close()

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

def dibujar_opciones_cabecera(pdf, x_start, y, h_row, opts, idx_checked, step=12.0):
    curr_x = x_start
    for i, opt in enumerate(opts):
        pdf.set_font('Arial', '', 7)
        pdf.set_xy(curr_x, y)
        pdf.cell(7, h_row, opt, 0, 0, 'R')
        box_x = curr_x + 8.0
        box_y = y + (h_row - 3) / 2
        pdf.rect(box_x, box_y, 3, 3)
        if i == idx_checked:
            pdf.set_xy(box_x, box_y)
            pdf.set_font('Arial', 'B', 7)
            pdf.cell(3, 3, 'X', 0, 0, 'C')
        curr_x += step

def dibujar_opciones_bloque(pdf, x_start, y, h_row, opts, idx_checked):
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

    pdf.set_font('Arial', 'B', 7)
    pdf.cell(35, h_row, 'Fecha de recepción:', 1, 0, 'R')
    pdf.set_font('Arial', '', 7)
    fecha_ingreso_sap = str(row.get('IngresoSAP', ''))
    pdf.cell(60, h_row, fecha_ingreso_sap if fecha_ingreso_sap != 'nan' else str(row.get('Fecha', '')).split()[0], 1, 0, 'C')
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
    pdf.cell(100, h_row, fecha_ingreso_sap, 1, 1, 'C')

    pdf.set_font('Arial', 'B', 7)
    pdf.cell(35, h_row, 'Placa del transporte:', 1, 0, 'R')
    pdf.set_font('Arial', '', 7)
    pdf.cell(60, h_row, str(row.get('Placa', '')), 1, 0, 'C')
    pdf.cell(15, h_row, '', 0, 0)
    pdf.set_font('Arial', 'B', 7)
    pdf.cell(67, h_row, 'Número de documento SAP:', 1, 0, 'R')
    pdf.set_font('Arial', '', 7)
    val_docsap = str(row.get('DocSAP', ''))
    pdf.cell(100, h_row, val_docsap if val_docsap != 'nan' else '', 1, 1, 'C')
    pdf.ln(2)

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

    pdf.set_xy(10, y_start)
    pdf.cell(5, h_row, '1', 0, 0, 'C')
    pdf.cell(54, h_row, '¿El marchamo se encuentra alterado?', 0, 0)
    est_mar = str(row.get('EstadoMarchamo', 'N/A'))
    idx_m = 0 if est_mar == 'SI' else (1 if est_mar == 'NO' else 2)
    dibujar_opciones_cabecera(pdf, 69, y_start, h_row, ['SI', 'NO', 'N/A'], idx_m, step=12.0)

    pdf.set_xy(15, y_start + h_row)
    pdf.cell(25, h_row, 'No. de marchamo:', 0, 0)
    val_marchamo = str(row.get('Marchamo', ''))
    pdf.cell(55, h_row, val_marchamo if val_marchamo != 'nan' else '', 'B', 0, 'C')

    pdf.set_xy(115, y_start + h_row/2)
    pdf.cell(5, h_row, '2', 0, 0, 'C')
    pdf.cell(43, h_row, '¿El candado se encuentra alterado?', 0, 0)
    est_can = str(row.get('EstadoCandado', 'N/A'))
    idx_c = 0 if est_can == 'SI' else (1 if est_can == 'NO' else 2)
    dibujar_opciones_cabecera(pdf, 163, y_start + h_row/2, h_row, ['SI', 'NO', 'N/A'], idx_c, step=12.0)

    pdf.set_xy(200, y_start + h_row/2)
    pdf.set_font('Arial', 'B', 7)
    pdf.cell(30, h_row, 'NO', 0, 0, 'C')
    pdf.set_font('Arial', '', 7)
    pdf.cell(57, h_row, 'Retener lote y emitir reclamo.', 0, 1, 'C')
    pdf.ln(h_row/2 + 2)

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
        
    temp_trans = str(row.get('TempTransporte', 'N/A'))
    temp_mat = str(row.get('TempMaterial', 'N/A'))
    
    y_curr = pdf.get_y()
    pdf.set_xy(10, y_curr)
    pdf.set_font('Arial', 'B', 7)
    pdf.cell(5, h_row, '14', 'LTR', 0, 'C')
    pdf.set_font('Arial', '', 7)
    pdf.cell(130, h_row, 'Si el material es perecedero, ¿el transporte cumple con la temperatura especificada?', 'LTR', 0, 'L')
    pdf.cell(55, h_row, '', 'LTR', 0)
    idx_temp = 2 if (temp_trans == "N/A" or temp_trans == "nan" or not temp_trans) else 0
    dibujar_opciones_bloque(pdf, 145, y_curr, h_row, ["SI", "NO", "N/A"], idx_temp)
    
    pdf.set_xy(200, y_curr)
    pdf.set_font('Arial', 'B', 7)
    pdf.cell(30, h_row, "SI", 'LTR', 0, 'C')
    pdf.set_font('Arial', '', 7)
    pdf.cell(57, h_row, "Rechazar lote y emitir reclamo", 'LTR', 1, 'C')

    y_curr = pdf.get_y()
    pdf.set_xy(10, y_curr)
    pdf.cell(5, h_row, '', 'LBR', 0)
    pdf.set_font('Arial', 'B', 7)
    txt_temps = f"Temperatura Transporte (°C): {temp_trans if temp_trans!='nan' else 'N/A'}          Temperatura del Material (°C): {temp_mat if temp_mat!='nan' else 'N/A'}"
    if idx_temp == 2:
        txt_temps = "Temperatura Transporte (°C): N/A          Temperatura del Material (°C): N/A"
    pdf.cell(130, h_row, txt_temps, 'BR', 0, 'C')
    pdf.cell(55, h_row, '', 'BR', 0)
    pdf.cell(30, h_row, '', 'BR', 0)
    pdf.cell(57, h_row, '', 'BR', 1)

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
    pdf.cell(57, h_row, "Retener lote y emitir reclamo", 1, 1, 'C')

    est_tarima = str(row.get('EstadoTarima', 'SI'))
    idx_tarima = 0 if est_tarima == 'SI' else (1 if est_tarima == 'NO' else 2)

    y_curr = pdf.get_y()
    pdf.set_xy(10, y_curr)
    pdf.set_font('Arial', 'B', 7)
    pdf.cell(5, h_row, '16', 1, 0, 'C')
    pdf.set_font('Arial', '', 7)
    pdf.cell(130, h_row, '¿La tarima de uso retornable se encuentra en buenas condiciones y estado inocuo?', 1, 0, 'L')
    pdf.cell(55, h_row, '', 1, 0)
    dibujar_opciones_bloque(pdf, 145, y_curr, h_row, ["SI", "NO", "N/A"], idx_tarima)
    
    pdf.set_xy(200, y_curr)
    pdf.set_font('Arial', 'B', 7)
    pdf.cell(30, h_row, "SI", 1, 0, 'C')
    pdf.set_font('Arial', '', 7)
    pdf.cell(57, h_row, "Rechazar lote y emitir reclamo", 1, 1, 'C')

    pdf.ln(2)
    
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
    
    pdf.ln(12) 
    y_linea1 = pdf.get_y()

    pdf.line(55, y_linea1, 135, y_linea1)
    pdf.line(175, y_linea1, 255, y_linea1)

    pdf.set_xy(55, y_linea1 - 4)
    pdf.set_font('Arial', '', 8)
    pdf.cell(80, 4, str(row.get('Responsable', '')), 0, 0, 'C')
    pdf.set_xy(175, y_linea1 - 4)
    pdf.cell(80, 4, supervisor, 0, 0, 'C')

    pdf.set_xy(10, y_linea1 - 0.5)
    pdf.set_font('Arial', 'B', 6.5)
    pdf.cell(43, 4, 'Nombre de la persona que realiza la recepción:', 0, 0, 'R')
    pdf.set_xy(145, y_linea1 - 0.5)
    pdf.set_font('Arial', 'B', 8)
    pdf.cell(28, 4, 'Vo. Bo. Calidad:', 0, 0, 'R')

    y_linea2 = y_linea1 + 10
    pdf.line(55, y_linea2, 135, y_linea2)
    pdf.line(175, y_linea2, 255, y_linea2)

    pdf.set_xy(35, y_linea2 - 1)
    pdf.set_font('Arial', 'B', 8)
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


if "items_ingreso_actual" not in st.session_state:
  st.session_state.items_ingreso_actual = []

df_db = None
if os.path.exists(ARCHIVO_CATALOGO):
  try:
    df_db = pd.read_excel(ARCHIVO_CATALOGO)
    df_db.columns = df_db.columns.str.strip()
  except Exception as e:
    st.error(f"Error al leer el catálogo: {e}")

st.sidebar.header("⚙️ Configuración y Respaldo")
archivo_catalogo_nuevo = st.sidebar.file_uploader("Actualizar Catálogo", type=["xlsx"])
if archivo_catalogo_nuevo is not None:
  with open(ARCHIVO_CATALOGO, "wb") as f:
      f.write(archivo_catalogo_nuevo.getbuffer())
  df_db = pd.read_excel(ARCHIVO_CATALOGO)
  df_db.columns = df_db.columns.str.strip()
  st.sidebar.success("¡Catálogo actualizado!")

tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📝 1. Ingreso de Operador", "✍️ 2. Panel de Calidad", "🔍 3. Trazabilidad", "📁 4. Explorador de Archivos", "📊 5. Control Excel"
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

    palabras_clave_empaque = ["envase", "etiqueta", "tapa", "botella", "caja", "bolsa", "film", "saco", "precinto", "carton"]
    es_empaque = any(kw in str(material_sel).lower() or kw in str(tipo_mat_sel).lower() for kw in palabras_clave_empaque)

    es_aceite = "aceite" in str(material_sel).lower() or "aceite" in str(tipo_mat_sel).lower()
    
    if es_empaque or es_aceite:
        requiere_temperatura = False
    else:
        palabras_clave_temp = ["pulpa", "concentrado", "leche", "refrigerado", "congelado", "fresco", "crema", "mantequilla", "queso", "carne", "pollo", "embutido"]
        requiere_temperatura = any(kw in str(material_sel).lower() or kw in str(tipo_mat_sel).lower() for kw in palabras_clave_temp)

    mat_lower = str(material_sel).lower()
    exime_coa = "coco" in mat_lower or "naranja valencia hn" in mat_lower or "mora" in mat_lower

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
        ingreso_sap = st.date_input("Fecha de Ingreso")
      with s2:
        doc_sap = st.text_input("No. Documento SAP", disabled=True, value="")
      with s3:
        marchamo = st.text_input("No. de Marchamo", disabled=True, value="")

      m1, m2 = st.columns(2)
      with m1:
        estado_marchamo = st.selectbox("¿El marchamo se encuentra alterado?", ["NO", "SI", "N/A"])
      with m2:
        estado_candado = st.selectbox("¿El candado se encuentra alterado?", ["NO", "SI", "N/A"])

      estado_tarima = st.selectbox("¿La tarima de uso retornable se encuentra en buenas condiciones y estado inocuo?", ["SI", "NO", "N/A"])

      c4, c5 = st.columns(2)
      with c4:
        cantidad = st.number_input("Cantidad Recibida", min_value=0.0, format="%.2f")
      with c5:
        unidad = st.selectbox("Unidad de Medida", ["Kg", "Litros", "Sacos", "Unidades", "Toneladas", "Galones", "Cajas"])

      c6, c7 = st.columns(2)
      with c6:
        fecha_produccion = st.date_input("Fecha de Producción")
      with c7:
        fecha_vencimiento = st.date_input("Fecha de Vencimiento / Caducidad")

      temp_trans = "N/A"
      temp_mat = "N/A"
      if requiere_temperatura:
        st.markdown("---")
        st.subheader("❄️ Control de Temperaturas (Requerido por tipo de material)")
        tmp1, tmp2 = st.columns(2)
        with tmp1:
            temp_trans = st.text_input("Temp. Transporte (°C)", value="4.0", placeholder="Ej. 4.0 o -18.5")
        with tmp2:
            temp_mat = st.text_input("Temp. del Material (°C)", value="4.0", placeholder="Ej. 4.0 o -18.5")
      else:
        st.info("ℹ️ Este material no requiere control de temperatura (se registrará automáticamente como N/A).")

      st.markdown("---")
      if exime_coa:
          st.info("ℹ️ Este material (Coco / Naranja Valencia HN / Mora) no requiere adjuntar COA obligatoriamente.")
          certificado_file = st.file_uploader("Subir Archivo del Certificado / COA (Opcional)", type=["jpg", "jpeg", "png", "pdf"])
      else:
          certificado_file = st.file_uploader("Subir Archivo del Certificado / COA (Obligatorio)", type=["jpg", "jpeg", "png", "pdf"])

      agregar_item = st.form_submit_button("➕ Agregar producto a la lista")

    if agregar_item:
      if not lote_proveedor or not factura_guia:
        st.warning("⚠️ Debes ingresar el número de factura/guía y el número de lote.")
      elif certificado_file is None and not exime_coa:
        st.warning("⚠️ Es obligatorio adjuntar el archivo del certificado para este material.")
      else:
        ruta_cert_temp = ""
        nombre_cert_str = "Sin certificado"
        
        if certificado_file is not None:
            os.makedirs("temp_certificados", exist_ok=True)
            nombre_archivo_unico = f"{datetime.now().strftime('%Y%m%d_%H%M%S_%f')}_{certificado_file.name}"
            ruta_cert_temp = os.path.join("temp_certificados", nombre_archivo_unico)
            with open(ruta_cert_temp, "wb") as f:
              f.write(certificado_file.getbuffer())
            nombre_cert_str = certificado_file.name

        nuevo_item = {
            "Planta": planta_sel, "TipoMaterial": tipo_mat_sel, "Proveedor": proveedor_sel,
            "NombreMaterial": material_sel, "FacturaGuia": factura_guia, "Transporte": proveedor_sel, 
            "Placa": placa_transporte, "IngresoSAP": str(ingreso_sap), "DocSAP": doc_sap,
            "EstadoMarchamo": estado_marchamo, "Marchamo": marchamo, "EstadoCandado": estado_candado,
            "EstadoTarima": estado_tarima,
            "Lote": lote_proveedor, "Cantidad": cantidad, "Unidad": unidad, 
            "TempTransporte": str(temp_trans), "TempMaterial": str(temp_mat),
            "Produccion": str(fecha_produccion), "Vencimiento": str(fecha_vencimiento),
            "Criticidad": criticidad_val, "RutaTemp": ruta_cert_temp,
            "NombreCertificado": nombre_cert_str
        }
        st.session_state.items_ingreso_actual.append(nuevo_item)
        st.success("✅ ¡Producto agregado correctamente!")

    if len(st.session_state.items_ingreso_actual) > 0:
      st.markdown("---")
      st.subheader("📋 Resumen de Productos Agregados y Edición")
      for idx, item in enumerate(st.session_state.items_ingreso_actual):
          st.markdown(f"""
          <div style='background-color: #eef2f7; padding: 10px 15px; border-radius: 6px; margin-bottom: 8px; border-left: 4px solid #0066cc;'>
              <b>#{idx+1} - {item['NombreMaterial']}</b> | Proveedor: {item['Proveedor']} | Lote: {item['Lote']} | Cantidad: {item['Cantidad']} {item['Unidad']}<br>
              <span style='color: #0d6efd; font-size: 0.9rem;'>✅ Certificado: <b>{item['NombreCertificado']}</b></span>
          </div>
          """, unsafe_allow_html=True)

      with st.expander("✏️ Editar un producto agregado"):
          idx_editar = st.selectbox(
              "Selecciona el producto a editar:",
              options=range(len(st.session_state.items_ingreso_actual)),
              format_func=lambda x: f"#{x+1} - {st.session_state.items_ingreso_actual[x]['NombreMaterial']} (Lote: {st.session_state.items_ingreso_actual[x]['Lote']})"
          )
          item_actual = st.session_state.items_ingreso_actual[idx_editar]
          with st.form(f"form_editar_item_{idx_editar}"):
              nuevo_lote = st.text_input("Lote del Proveedor", value=item_actual['Lote'])
              nueva_factura = st.text_input("No. de Factura/Guía", value=item_actual['FacturaGuia'])
              nueva_cantidad = st.number_input("Cantidad Recibida", min_value=0.0, value=float(item_actual['Cantidad']), format="%.2f")
              nuevo_estado_tarima = st.selectbox("¿La tarima de uso retornable se encuentra en buenas condiciones y estado inocuo?", ["SI", "NO", "N/A"], index=0 if item_actual.get('EstadoTarima','SI')=='SI' else (1 if item_actual.get('EstadoTarima')=='NO' else 2))
              
              temp_trans_edit = item_actual['TempTransporte']
              temp_mat_edit = item_actual['TempMaterial']
              if temp_trans_edit != "N/A":
                  e_t1, e_t2 = st.columns(2)
                  with e_t1:
                      temp_trans_edit = st.text_input("Temp. Transporte (°C)", value=str(temp_trans_edit))
                  with e_t2:
                      temp_mat_edit = st.text_input("Temp. del Material (°C)", value=str(temp_mat_edit))

              guardar_cambios = st.form_submit_button("💾 Guardar Cambios en este Producto")
              if guardar_cambios:
                  st.session_state.items_ingreso_actual[idx_editar]['Lote'] = nuevo_lote
                  st.session_state.items_ingreso_actual[idx_editar]['FacturaGuia'] = nueva_factura
                  st.session_state.items_ingreso_actual[idx_editar]['Cantidad'] = nueva_cantidad
                  st.session_state.items_ingreso_actual[idx_editar]['EstadoTarima'] = nuevo_estado_tarima
                  if temp_trans_edit != "N/A":
                      st.session_state.items_ingreso_actual[idx_editar]['TempTransporte'] = temp_trans_edit
                      st.session_state.items_ingreso_actual[idx_editar]['TempMaterial'] = temp_mat_edit
                  st.success("¡Producto actualizado correctamente!")
                  st.rerun()

      col_del1, col_del2 = st.columns(2)
      with col_del1:
          item_a_borrar = st.selectbox(
              "Quitar producto de la lista actual:",
              options=range(len(st.session_state.items_ingreso_actual)),
              format_func=lambda x: f"#{x+1} - {st.session_state.items_ingreso_actual[x]['NombreMaterial']}"
          )
          if st.button("❌ Quitar seleccionado"):
            st.session_state.items_ingreso_actual.pop(item_a_borrar)
            st.rerun()

      st.markdown("---")
      st.subheader("3. Finalizar y Enviar a Revisión")
      responsable = st.selectbox("Operador Responsable", ["Carlos López", "Carlos Gallardo", "Carlos Tema", "Oliver", "Willy", "Ronald"])
      
      st.write("✍️ **Dibuja tu Firma de Operador:**")
      canvas_op = None
      if st_canvas:
          canvas_op = st_canvas(stroke_width=2, stroke_color="#000", background_color="#ffffff", height=100, width=400, drawing_mode="freedraw", key="firma_op")

      if st.button("🚀 Enviar a Calidad", type="primary"):
        ruta_firma_op = guardar_firma(canvas_op, "firma_op")

        for item in st.session_state.items_ingreso_actual:
          nuevo_registro = {
              "Planta": str(item["Planta"]), "TipoMaterial": str(item["TipoMaterial"]),
              "NombreMaterial": str(item["NombreMaterial"]), "Proveedor": str(item["Proveedor"]),
              "FacturaGuia": str(item["FacturaGuia"]), "Transporte": str(item["Transporte"]),
              "Placa": str(item["Placa"]), "IngresoSAP": str(item["IngresoSAP"]), "DocSAP": str(item["DocSAP"]),
              "EstadoMarchamo": str(item["EstadoMarchamo"]), "Marchamo": str(item["Marchamo"]),
              "EstadoCandado": str(item["EstadoCandado"]), "EstadoTarima": str(item.get("EstadoTarima", "SI")), "LoteProveedor": str(item["Lote"]),
              "Cantidad": item["Cantidad"], "Unidad": str(item["Unidad"]), 
              "TempTransporte": str(item["TempTransporte"]), "TempMaterial": str(item["TempMaterial"]),
              "FechaProduccion": str(item["Produccion"]), "FechaVencimiento": str(item["Vencimiento"]),
              "Criticidad": str(item["Criticidad"]), "Responsable": str(responsable),
              "Fecha": str(item["IngresoSAP"]), "Estado": "Pendiente", "FirmaCalidad": "Sin firmar",
              "RutaCarpeta": "", "RutaChecklistOficial": "", "RutaPdfCert": str(item["RutaTemp"]),
              "RutaFirmaOp": ruta_firma_op, "RutaFirmaSup": ""
          }
          guardar_registro_db(nuevo_registro)
        
        asunto_mail = "🚨 Nuevo Material Pendiente de Revisión - Sistema LIF"
        cuerpo_mail = f"Se ha enviado un nuevo lote de material para revisión en el sistema LIF.\n\nOperador responsable: {responsable}\nFecha de Ingreso: {item['IngresoSAP']}\n\nPor favor ingrese al sistema para su verificación y aprobación."
        enviar_notificacion_outlook(asunto_mail, cuerpo_mail)

        st.session_state.items_ingreso_actual = []
        st.success("🎉 ¡Enviados a calidad con éxito y guardados en la Base de Datos SQLite!")
        st.rerun()

  else:
    st.warning("⚠️ Sube el catálogo maestro.")

# ------------------------------------------
# PESTAÑA 2: PANEL DE CALIDAD Y FIRMA
# ------------------------------------------
with tab2:
  st.header("✍️ Panel de Control y Firma")
  
  df_p = cargar_pendientes()
  pendientes = df_p[df_p["Estado"] == "Pendiente"] if not df_p.empty else pd.DataFrame()

  if len(pendientes) == 0:
    st.info("No hay registros pendientes.")
  else:
    for index, row in pendientes.iterrows():
      with st.expander(f"📥 ID #{row['ID']} | Planta: {row.get('Planta', 'N/A')} | Material: {row['NombreMaterial']} | Lote: {row['LoteProveedor']} | Ingreso SAP: {row['IngresoSAP']}"):
        st.write(f"**Operador Responsable:** {row['Responsable']}")
        if row["RutaPdfCert"] and os.path.exists(str(row["RutaPdfCert"])):
          st.write("📄 Certificado adjunto.")

        nombre_supervisor = st.text_input("Nombre del Supervisor", value="Carlos López", key=f"sup_nom_{row['ID']}")
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
              with st.spinner("Generando y unificando documento maestro..."):
                carpeta_base = "./Registros_Oficiales"
                
                planta_safe = str(row.get("Planta", "Planta")).replace("/", "_").replace(chr(92), "_")
                tipo_mat_safe = str(row.get("TipoMaterial", "Tipo")).replace("/", "_").replace(chr(92), "_")
                nombre_mat_safe = str(row.get("NombreMaterial", "Material")).replace("/", "_").replace(chr(92), "_")
                lote_safe = str(row.get("LoteProveedor", "Lote")).replace("/", "_").replace(chr(92), "_")
                fecha_safe = str(row.get("IngresoSAP", "")).replace("/", "-")
                
                nombre_base_archivo = f"{nombre_mat_safe}_{lote_safe}_{fecha_safe}"
                
                subcarpeta_lote = f"Lote_{lote_safe}_{nombre_mat_safe}"
                ruta_destino = os.path.join(carpeta_base, planta_safe, tipo_mat_safe, nombre_mat_safe, subcarpeta_lote)
                os.makedirs(ruta_destino, exist_ok=True)

                ruta_firma_sup = guardar_firma(canvas_sup, f"firma_sup_{row['ID']}")

                ruta_cert_origen = str(row.get("RutaPdfCert", ""))
                ext_cert = os.path.splitext(ruta_cert_origen)[1] if ruta_cert_origen and ruta_cert_origen != "nan" else ".pdf"
                ruta_cert_final = os.path.join(ruta_destino, f"Certificado_{nombre_base_archivo}{ext_cert}")
                
                try:
                    if ruta_cert_origen and ruta_cert_origen != "nan" and os.path.exists(ruta_cert_origen):
                        shutil.copy(ruta_cert_origen, ruta_cert_final)
                    else:
                        ruta_cert_final = ""
                except Exception:
                    ruta_cert_final = ""

                ruta_checklist_temp = os.path.join(ruta_destino, f"TEMP_{nombre_base_archivo}.pdf")
                generar_pdf_ricc14(row, nombre_supervisor, ruta_checklist_temp)

                ruta_unificada = os.path.join(ruta_destino, f"{nombre_base_archivo}.pdf")
                if ruta_cert_final and os.path.exists(ruta_cert_final):
                    generar_documento_unificado(ruta_checklist_temp, ruta_cert_final, ruta_unificada)
                else:
                    generar_documento_unificado(ruta_checklist_temp, "", ruta_unificada)
                
                if os.path.exists(ruta_checklist_temp):
                    os.remove(ruta_checklist_temp)

                actualizar_registro_db(row["ID"], {
                    "Estado": "Aprobado",
                    "FirmaCalidad": str(nombre_supervisor),
                    "RutaCarpeta": str(ruta_destino),
                    "RutaChecklistOficial": str(ruta_unificada),
                    "RutaPdfCert": str(ruta_cert_final),
                    "RutaFirmaSup": str(ruta_firma_sup)
                })

                st.success("🎉 ¡Aprobado! Documento maestro guardado exitosamente en carpetas locales.")
                st.rerun()

        with col_btn2:
          if st.button("🗑️ Eliminar Registro", key=f"del_pend_{row['ID']}"):
            eliminar_registro_db(row["ID"])
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
      busqueda = st.text_input("Filtro rápido (Planta, Lote, Material o Proveedor)")
    with col_f2:
      fecha_filtro = st.date_input("Filtrar por Fecha de Ingreso", value=None)

    resultado = aprobados.copy()
    if busqueda:
      resultado = resultado[
          resultado["Planta"].astype(str).str.contains(busqueda, case=False, na=False) |
          resultado["LoteProveedor"].astype(str).str.contains(busqueda, case=False, na=False) |
          resultado["NombreMaterial"].astype(str).str.contains(busqueda, case=False, na=False) |
          resultado["Proveedor"].astype(str).str.contains(busqueda, case=False, na=False)
      ]
    if fecha_filtro:
      resultado = resultado[resultado["IngresoSAP"].str.startswith(str(fecha_filtro))]

    st.dataframe(resultado[["Planta", "NombreMaterial", "FacturaGuia", "LoteProveedor", "IngresoSAP", "FirmaCalidad"]], use_container_width=True)

    for idx, row in resultado.iterrows():
      planta_val = row.get('Planta', 'N/A')
      fecha_rec = str(row.get('IngresoSAP', 'N/A'))
      with st.expander(f"📁 Planta: {planta_val} | Material: {row['NombreMaterial']} | Lote: {row['LoteProveedor']} | Ingreso SAP: {fecha_rec}"):
        
        nombre_mat_descarga = str(row.get('NombreMaterial', 'Material')).replace("/", "_")
        lote_descarga = str(row.get('LoteProveedor', 'Lote')).replace("/", "_")
        fecha_descarga = str(row.get('IngresoSAP', 'Fecha')).replace("/", "-")
        nombre_descarga_base = f"{nombre_mat_descarga}_{lote_descarga}_{fecha_descarga}"

        col_d1, col_d2, col_d3 = st.columns(3)
        ruta_chk = str(row.get("RutaChecklistOficial", ""))
        with col_d1:
          if pd.notna(ruta_chk) and ruta_chk != "nan" and os.path.exists(ruta_chk):
            with open(ruta_chk, "rb") as f:
              st.download_button("📊 Descargar PDF Maestro Unificado", data=f, file_name=f"{nombre_descarga_base}.pdf", mime="application/pdf", key=f"dl_chk_{row['ID']}")
        
        ruta_cert = str(row.get("RutaPdfCert", ""))
        with col_d2:
          if pd.notna(ruta_cert) and ruta_cert != "nan" and os.path.exists(ruta_cert):
            ext_cert_dl = os.path.splitext(ruta_cert)[1]
            with open(ruta_cert, "rb") as f:
              st.download_button("📄 Descargar Certificado Original", data=f, file_name=f"Certificado_{nombre_descarga_base}{ext_cert_dl}", mime="application/pdf" if ruta_cert.endswith(".pdf") else "image/jpeg", key=f"dl_cert_{row['ID']}")
        
        with col_d3:
          if st.button("🗑️ Borrar", key=f"del_apr_{row['ID']}"):
            eliminar_registro_db(row["ID"])
            st.rerun()

# ------------------------------------------
# PESTAÑA 4: EXPLORADOR DE ARCHIVOS
# ------------------------------------------
with tab4:
  st.header("📁 Explorador de Archivos Corporativo")
  st.markdown("Navegue por la estructura de carpetas: **Planta ➡️ Tipo de Material ➡️ Proveedor ➡️ Material ➡️ Lotes**.")
  
  df_p = cargar_pendientes()
  aprobados_exp = df_p[df_p["Estado"] == "Aprobado"] if not df_p.empty else pd.DataFrame()

  if len(aprobados_exp) == 0:
    st.info("No hay registros aprobados para mostrar en el explorador.")
  else:
    busqueda_exp = st.text_input("🔍 Buscar dentro del Explorador (Material, Lote o Proveedor)", placeholder="Escriba para filtrar carpetas...")
    
    df_filtrado_exp = aprobados_exp.copy()
    if busqueda_exp:
      df_filtrado_exp = df_filtrado_exp[
          df_filtrado_exp["Planta"].astype(str).str.contains(busqueda_exp, case=False, na=False) |
          df_filtrado_exp["LoteProveedor"].astype(str).str.contains(busqueda_exp, case=False, na=False) |
          df_filtrado_exp["NombreMaterial"].astype(str).str.contains(busqueda_exp, case=False, na=False) |
          df_filtrado_exp["Proveedor"].astype(str).str.contains(busqueda_exp, case=False, na=False)
      ]

    if len(df_filtrado_exp) == 0:
      st.warning("⚠️ No se encontraron coincidencias con el término de búsqueda en el explorador.")
    else:
      plantas_disponibles = df_filtrado_exp["Planta"].dropna().unique()
      
      for planta in plantas_disponibles:
        df_planta_exp = df_filtrado_exp[df_filtrado_exp["Planta"] == planta]
        
        with st.expander(f"📁 Planta: {planta} ({len(df_planta_exp)} registros)"):
          tipos_material_exp = df_planta_exp["TipoMaterial"].dropna().unique()
          
          for tipo_mat in tipos_material_exp:
            df_tipo_exp = df_planta_exp[df_planta_exp["TipoMaterial"] == tipo_mat]
            
            with st.expander(f"📂 Tipo de Material: {tipo_mat} ({len(df_tipo_exp)} registros)"):
              proveedores_exp = df_tipo_exp["Proveedor"].dropna().unique()
              
              for proveedor in proveedores_exp:
                df_prov_exp = df_tipo_exp[df_tipo_exp["Proveedor"] == proveedor]
                
                with st.expander(f"🏢 Proveedor: {proveedor} ({len(df_prov_exp)} registros)"):
                  materiales_exp = df_prov_exp["NombreMaterial"].dropna().unique()
                  
                  for material in materiales_exp:
                    df_mat_exp = df_prov_exp[df_prov_exp["NombreMaterial"] == material]
                    
                    with st.expander(f"📦 Material: {material} ({len(df_mat_exp)} lotes)"):
                      for idx_exp, row_exp in df_mat_exp.iterrows():
                        lote_exp = row_exp.get("LoteProveedor", "S/Lote")
                        fecha_exp = str(row_exp.get("IngresoSAP", "S/Fecha"))
                        
                        st.markdown(f"""
                        <div style='background-color: #f1f3f5; padding: 8px 12px; border-radius: 5px; margin-bottom: 6px; border-left: 3px solid #0066cc;'>
                            <b>Lote:</b> {lote_exp} | <b>Ingreso SAP:</b> {fecha_exp}
                        </div>
                        """, unsafe_allow_html=True)
                        
                        nombre_mat_d = str(row_exp.get('NombreMaterial', 'Material')).replace("/", "_")
                        lote_d = str(row_exp.get('LoteProveedor', 'Lote')).replace("/", "_")
                        fecha_d = str(row_exp.get('IngresoSAP', 'Fecha')).replace("/", "-")
                        base_d = f"{nombre_mat_d}_{lote_d}_{fecha_d}"
                        
                        col_ex1, col_ex2 = st.columns(2)
                        ruta_chk_exp = str(row_exp.get("RutaChecklistOficial", ""))
                        with col_ex1:
                          if pd.notna(ruta_chk_exp) and ruta_chk_exp != "nan" and os.path.exists(ruta_chk_exp):
                            with open(ruta_chk_exp, "rb") as f_chk:
                              st.download_button("📊 Descargar PDF Maestro Unificado", data=f_chk, file_name=f"{base_d}.pdf", mime="application/pdf", key=f"exp_chk_{row_exp['ID']}")
                        
                        ruta_cert_exp = str(row_exp.get("RutaPdfCert", ""))
                        with col_ex2:
                          if pd.notna(ruta_cert_exp) and ruta_cert_exp != "nan" and os.path.exists(ruta_cert_exp):
                            ext_c = os.path.splitext(ruta_cert_exp)[1]
                            with open(ruta_cert_exp, "rb") as f_cer:
                              st.download_button("📄 Descargar Certificado Original", data=f_cer, file_name=f"Certificado_{base_d}{ext_c}", mime="application/pdf" if ruta_cert_exp.endswith(".pdf") else "image/jpeg", key=f"exp_cert_{row_exp['ID']}")
                        st.markdown("---")

# ------------------------------------------
# PESTAÑA 5: CONTROL EXCEL Y DESCARGA MASIVA
# ------------------------------------------
with tab5:
  st.header("📊 Control de Ingresos y Descaarga Masiva")
  df_p = cargar_pendientes()
  aprobados_excel = df_p[df_p["Estado"] == "Aprobado"].copy() if not df_p.empty else pd.DataFrame()

  if len(aprobados_excel) == 0:
    st.info("No hay información suficiente para generar reportes.")
  else:
    columnas_reporte = [
        "Planta", "TipoMaterial", "NombreMaterial", "Proveedor", "LoteProveedor", 
        "FacturaGuia", "Placa", "IngresoSAP", "DocSAP", "EstadoMarchamo", "Marchamo",
        "EstadoTarima", "Cantidad", "Unidad", "TempTransporte", "TempMaterial", "Criticidad", "Responsable", "FirmaCalidad"
    ]
    df_reporte = aprobados_excel[columnas_reporte].copy()
    st.dataframe(df_reporte, use_container_width=True)

    col_dl1, col_dl2 = st.columns(2)
    with col_dl1:
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_reporte.to_excel(writer, index=False, sheet_name='Historial')
        st.download_button("📥 Descargar Base Completa (Excel)", data=output.getvalue(), file_name=f"Control_LIF_{datetime.now().strftime('%Y%m%d')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")

    with col_dl2:
        if st.button("📦 Descargar Todos los Certificados y PDFs (ZIP)", type="secondary"):
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                if os.path.exists("./Registros_Oficiales"):
                    for root, dirs, files in os.walk("./Registros_Oficiales"):
                        for file in files:
                            ruta_completa = os.path.join(root, file)
                            ruta_relativa = os.path.relpath(ruta_completa, "./Registros_Oficiales")
                            zip_file.write(ruta_completa, ruta_relativa)
            
            zip_buffer.seek(0)
            st.download_button(
                label="⬇️ Click aquí para descargar el ZIP",
                data=zip_buffer,
                file_name=f"Certificados_Y_Documentos_LIF_{datetime.now().strftime('%Y%m%d_%H%M')}.zip",
                mime="application/zip"
            )
