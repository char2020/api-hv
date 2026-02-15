from flask import Flask, request, send_file, jsonify, redirect
from flask_cors import CORS
from docx import Document
from docx.shared import RGBColor, Pt, Inches
from docx.oxml import parse_xml
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
import os
import re
import io
import base64
import requests
import time
from datetime import datetime
try:
    from google.oauth2.credentials import Credentials
    from google_auth_oauthlib.flow import InstalledAppFlow, Flow
    from googleapiclient.discovery import build
    from googleapiclient.http import MediaIoBaseUpload
    from googleapiclient.errors import HttpError
    GOOGLE_DRIVE_AVAILABLE = True
except ImportError:
    GOOGLE_DRIVE_AVAILABLE = False
    Flow = None
    Credentials = None
    InstalledAppFlow = None
    build = None
    MediaIoBaseUpload = None
    HttpError = Exception

app = Flask(__name__)
# CORS con restricciones de seguridad - solo permitir orígenes específicos
allowed_origins = os.getenv('ALLOWED_ORIGINS', 'https://generador-hojas-vida.web.app,https://generador-hojas-vida.firebaseapp.com').split(',')
CORS(app, origins=allowed_origins, methods=['GET', 'POST'], allow_headers=['Content-Type'])

# Configuración de APIs de iLovePDF
# API Principal: Usada en la API de cursos-certificados (GitHub) - ~250 conversiones
# API de Respaldo: Para cuando se agoten los créditos de la principal - ~250 conversiones
# Total disponible: ~500 conversiones
ILOVEPDF_APIS = [
    {
        'name': 'primary',
        # API Principal - Credenciales de la API de cursos-certificados
        # TODO: Reemplazar con las credenciales reales de la API principal
        'public_key': os.getenv('ILOVEPDF_PRIMARY_PUBLIC_KEY', None),  # Agregar aquí la public_key principal
        'secret_key': os.getenv('ILOVEPDF_PRIMARY_SECRET_KEY', None)   # Agregar aquí la secret_key principal
    },
    {
        'name': 'backup',
        # API de Respaldo - Se usa automáticamente cuando la principal se queda sin créditos
        # NOTA: Estas credenciales deben moverse a variables de entorno en producción
        'public_key': os.getenv('ILOVEPDF_BACKUP_PUBLIC_KEY', 'project_public_e8de4c9dde8d3130930dc8f9620f9fd0_4gcUq34631a35630e89502c9cb2229d123ff4'),
        'secret_key': os.getenv('ILOVEPDF_BACKUP_SECRET_KEY', 'secret_key_5f1ab1bb9dc866aadc8a05671e460491_zNqoaf28f8b33e1755f025940359d1d4a70a3')
    }
]

# INSTRUCCIONES PARA CONFIGURAR LA API PRINCIPAL:
# 1. Busca las credenciales de iLovePDF en el código de la API de cursos-certificados (GitHub)
# 2. Reemplaza los valores None arriba con las credenciales reales, O
# 3. Configura variables de entorno en Render:
#    - ILOVEPDF_PRIMARY_PUBLIC_KEY = "tu_public_key_aqui"
#    - ILOVEPDF_PRIMARY_SECRET_KEY = "tu_secret_key_aqui"

# Variable para rastrear qué API está activa
current_api_index = 0

# Meses en español
MESES = {
    1: 'enero', 2: 'febrero', 3: 'marzo', 4: 'abril',
    5: 'mayo', 6: 'junio', 7: 'julio', 8: 'agosto',
    9: 'septiembre', 10: 'octubre', 11: 'noviembre', 12: 'diciembre'
}

def formatear_fecha(fecha_str: str) -> str:
    """Convierte una fecha a formato '20 de noviembre de 1990'
    Acepta formatos: YYYY-MM-DD, DD/MM/YYYY, DD-MM-YYYY
    """
    if not fecha_str:
        return ''
    
    fecha_str = fecha_str.strip()
    
    try:
        # Intentar formato YYYY-MM-DD (estándar de date picker)
        if fecha_str.count('-') == 2:
            partes = fecha_str.split('-')
            if len(partes[0]) == 4:  # YYYY-MM-DD
                año, mes, dia = partes
                año, mes, dia = int(año), int(mes), int(dia)
                if 1 <= mes <= 12 and 1 <= dia <= 31:
                    return f"{dia} de {MESES[mes]} de {año}"
        
        # Intentar formato DD/MM/YYYY o DD-MM-YYYY
        if fecha_str.count('/') == 2 or fecha_str.count('-') == 2:
            separador = '/' if '/' in fecha_str else '-'
            partes = fecha_str.split(separador)
            if len(partes) == 3:
                # Determinar si es DD/MM/YYYY o MM/DD/YYYY
                # Asumimos DD/MM/YYYY si el primer número es <= 31
                if int(partes[0]) <= 31:
                    dia, mes, año = int(partes[0]), int(partes[1]), int(partes[2])
                else:
                    mes, dia, año = int(partes[0]), int(partes[1]), int(partes[2])
                
                if 1 <= mes <= 12 and 1 <= dia <= 31:
                    return f"{dia} de {MESES[mes]} de {año}"
    except (ValueError, IndexError, KeyError):
        pass
    
    # Si no se puede parsear, devolver la fecha original
    return fecha_str

@app.route('/', methods=['GET'])
def root():
    """Endpoint raíz para verificar que el servidor está funcionando"""
    return jsonify({
        "status": "ok", 
        "message": "API de Generación de Hojas de Vida funcionando",
        "endpoints": {
            "/health": "GET - Verificar estado del servidor",
            "/generate-word": "POST - Generar documento Word (Hoja de Vida)",
            "/generate-cuenta-cobro": "POST - Generar cuenta de cobro desde template",
            "/convert-word-to-pdf": "POST - Convertir Word a PDF usando iLovePDF"
        }
    })

@app.route('/health', methods=['GET'])
def health():
    """Endpoint de salud para verificar que el servidor está funcionando"""
    return jsonify({"status": "ok", "message": "API funcionando correctamente"})

def convert_word_to_pdf_with_ilovepdf(word_file_bytes, filename='document.docx'):
    """
    Convierte un archivo Word a PDF usando la API de iLovePDF con fallback automático
    si se acaban los créditos.
    """
    global current_api_index
    
    max_retries = len(ILOVEPDF_APIS)
    
    for attempt in range(max_retries):
        api_config = ILOVEPDF_APIS[current_api_index]
        
        # Si no hay credenciales configuradas, saltar esta API
        if not api_config['public_key'] or not api_config['secret_key']:
            current_api_index = (current_api_index + 1) % len(ILOVEPDF_APIS)
            continue
        
        try:
            # Paso 1: Autenticarse y obtener token
            auth_url = 'https://api.ilovepdf.com/v1/auth'
            auth_response = requests.post(auth_url, json={
                'public_key': api_config['public_key']
            })
            
            # Verificar respuesta de autenticación
            if auth_response.status_code != 200:
                error_text = auth_response.text.lower()
                error_json = {}
                try:
                    error_json = auth_response.json()
                except:
                    pass
                
                # Detectar errores de créditos
                if any(ind in error_text for ind in ['credits', 'quota', 'limit', 'exceeded', 'insufficient', 'balance']):
                    raise Exception(f"CREDITS_EXHAUSTED: {auth_response.text}")
                
                raise Exception(f"Error de autenticación ({auth_response.status_code}): {auth_response.text}")
            
            auth_data = auth_response.json()
            token = auth_data.get('token')
            
            if not token:
                raise Exception("No se recibió token de autenticación")
            
            # Paso 2: Iniciar tarea de conversión
            start_url = 'https://api.ilovepdf.com/v1/start/officepdf'
            headers = {'Authorization': f'Bearer {token}'}
            start_response = requests.get(start_url, headers=headers)
            
            if start_response.status_code != 200:
                error_text = start_response.text.lower()
                if any(ind in error_text for ind in ['credits', 'quota', 'limit', 'exceeded', 'insufficient', 'balance']):
                    raise Exception(f"CREDITS_EXHAUSTED: {start_response.text}")
                raise Exception(f"Error al iniciar tarea ({start_response.status_code}): {start_response.text}")
            
            task_data = start_response.json()
            server = task_data.get('server')
            task = task_data.get('task')
            
            if not server or not task:
                raise Exception("No se recibieron datos de servidor o tarea")
            
            # Paso 3: Subir archivo Word
            upload_url = f'https://{server}/v1/upload'
            files = {'file': (filename, word_file_bytes, 'application/vnd.openxmlformats-officedocument.wordprocessingml.document')}
            upload_response = requests.post(upload_url, files=files, headers=headers)
            
            if upload_response.status_code != 200:
                error_text = upload_response.text.lower()
                if any(ind in error_text for ind in ['credits', 'quota', 'limit', 'exceeded', 'insufficient', 'balance']):
                    raise Exception(f"CREDITS_EXHAUSTED: {upload_response.text}")
                raise Exception(f"Error al subir archivo ({upload_response.status_code}): {upload_response.text}")
            
            upload_data = upload_response.json()
            server_filename = upload_data.get('server_filename')
            
            if not server_filename:
                raise Exception("No se recibió nombre de archivo del servidor")
            
            # Paso 4: Procesar conversión
            process_url = f'https://{server}/v1/process'
            process_data = {
                'task': task,
                'tool': 'officepdf',
                'files': [{'server_filename': server_filename, 'filename': filename}]
            }
            process_response = requests.post(process_url, json=process_data, headers=headers)
            
            if process_response.status_code != 200:
                error_text = process_response.text.lower()
                if any(ind in error_text for ind in ['credits', 'quota', 'limit', 'exceeded', 'insufficient', 'balance']):
                    raise Exception(f"CREDITS_EXHAUSTED: {process_response.text}")
                raise Exception(f"Error al procesar ({process_response.status_code}): {process_response.text}")
            
            # Esperar un momento para que el procesamiento termine
            time.sleep(1)
            
            # Paso 5: Descargar PDF resultante
            download_url = f'https://{server}/v1/download/{task}'
            download_response = requests.get(download_url, headers=headers)
            
            if download_response.status_code != 200:
                error_text = download_response.text.lower()
                if any(ind in error_text for ind in ['credits', 'quota', 'limit', 'exceeded', 'insufficient', 'balance']):
                    raise Exception(f"CREDITS_EXHAUSTED: {download_response.text}")
                raise Exception(f"Error al descargar PDF ({download_response.status_code}): {download_response.text}")
            
            # Si llegamos aquí, la conversión fue exitosa
            print(f"✅ Conversión exitosa usando API {api_config['name']}")
            return download_response.content
            
        except Exception as e:
            error_message = str(e)
            is_credits_error = 'CREDITS_EXHAUSTED' in error_message or any(
                ind in error_message.lower() for ind in ['credits', 'quota', 'limit', 'exceeded', 'insufficient', 'balance', '401', '403']
            )
            
            if is_credits_error:
                # Cambiar a la siguiente API
                current_api_index = (current_api_index + 1) % len(ILOVEPDF_APIS)
                print(f"⚠️ Créditos agotados en API {api_config['name']}. Cambiando a API de respaldo...")
                
                # Si no hay más APIs, lanzar error
                if attempt == max_retries - 1:
                    raise Exception(f"Todas las APIs de iLovePDF han agotado sus créditos. Último error: {error_message}")
                
                # Continuar con el siguiente intento
                continue
            else:
                # Error diferente, relanzar
                raise e
    
    raise Exception("No se pudo convertir el archivo después de intentar todas las APIs disponibles")

@app.route('/convert-word-to-pdf', methods=['POST'])
def convert_word_to_pdf():
    """
    Convierte un documento Word a PDF usando iLovePDF con fallback automático
    """
    try:
        # Verificar si se envió un archivo
        if 'file' not in request.files:
            return jsonify({"error": "No se proporcionó ningún archivo"}), 400
        
        file = request.files['file']
        
        if file.filename == '':
            return jsonify({"error": "Nombre de archivo vacío"}), 400
        
        # Leer el archivo
        word_file_bytes = file.read()
        
        # Convertir a PDF
        pdf_bytes = convert_word_to_pdf_with_ilovepdf(word_file_bytes, file.filename)
        
        # Preparar respuesta
        output = io.BytesIO(pdf_bytes)
        output.seek(0)
        
        # Nombre del archivo PDF
        pdf_filename = file.filename.replace('.docx', '.pdf').replace('.doc', '.pdf')
        if not pdf_filename.endswith('.pdf'):
            pdf_filename += '.pdf'
        
        return send_file(
            output,
            mimetype='application/pdf',
            as_attachment=True,
            download_name=pdf_filename
        )
        
    except Exception as e:
        import traceback
        return jsonify({
            "error": str(e),
            "traceback": traceback.format_exc()
        }), 500

@app.route('/generate-word', methods=['POST'])
def generate_word():
    """Genera un documento Word desde cero con todos los datos recibidos"""
    try:
        data = request.json
        
        # Obtener datos básicos
        nombre = data.get('fullName', '').strip()
        cedula = data.get('idNumber', '').strip()
        fecha_raw = data.get('birthDate', '').strip()
        fecha = formatear_fecha(fecha_raw)  # Convertir a formato "20 de noviembre de 1990"
        telefono = data.get('phone', '').strip()
        direccion = data.get('address', '').strip()
        ciudad = data.get('place', '').strip()
        estado_civil = data.get('estadoCivil', '').strip().upper()  # Convertir a mayúsculas
        correo = data.get('email', '').strip()
        exp = data.get('idIssuePlace', '').strip()
        texto_perfil = data.get('profile', '').strip()
        
        # Obtener referencias
        referencias_familiares = data.get('referenciasFamiliares', [])
        referencias_personales = data.get('referenciasPersonales', [])
        
        # Obtener experiencias laborales
        experiencias = data.get('experiencias', [])
        
        # Obtener formaciones académicas
        formaciones = data.get('formaciones', [])
        high_school = data.get('highSchool', '').strip()
        institution = data.get('institution', '').strip()
        
        # Crear un nuevo documento desde cero
        doc = Document()
        
        # Configurar encabezado con fondo azul
        section = doc.sections[0]
        header = section.header
        
        # Limpiar párrafos existentes del encabezado
        for para in header.paragraphs:
            para.clear()
        
        # Crear nuevo párrafo para el encabezado
        header_para = header.paragraphs[0] if header.paragraphs else header.add_paragraph()
        header_para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        
        # Agregar fondo azul al encabezado usando XML
        header_xml = header_para._element
        pPr = header_xml.get_or_add_pPr()
        
        # Crear elemento de sombreado (fondo azul #5B9BD5)
        shd = parse_xml(r'<w:shd xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:fill="5B9BD5" w:val="clear"/>')
        pPr.append(shd)
        
        # Agregar espaciado superior e inferior para que el fondo se vea como una barra
        spacing = parse_xml(r'<w:spacing xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:before="240" w:after="240"/>')
        pPr.append(spacing)
        
        # Agregar indentación derecha para que el texto no toque el borde
        ind = parse_xml(r'<w:ind xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:right="360"/>')
        pPr.append(ind)
        
        # Agregar el nombre en blanco, negrita
        header_run = header_para.add_run(nombre.upper())
        header_run.font.name = "Calibri"
        header_run.font.size = Pt(11)
        header_run.font.color.rgb = RGBColor(255, 255, 255)  # Blanco
        header_run.bold = True
        
        # Agregar espacio
        doc.add_paragraph()
        
        # Nombre principal (Cambria 18, color #4472C4, mayúsculas, negrita)
        p_nombre = doc.add_paragraph()
        run_nombre = p_nombre.add_run(nombre.upper())
        run_nombre.font.name = "Cambria"
        run_nombre.font.size = Pt(18)
        run_nombre.font.color.rgb = RGBColor(0x44, 0x72, 0xC4)
        run_nombre.bold = True
        p_nombre.alignment = WD_ALIGN_PARAGRAPH.LEFT
        doc.add_paragraph()
        
        # Información personal - etiquetas en negrita azul, valores en negro
        p_cedula = doc.add_paragraph()
        run_cedula_label = p_cedula.add_run("Número de cédula: ")
        run_cedula_label.font.color.rgb = RGBColor(0x44, 0x72, 0xC4)
        run_cedula_label.bold = True
        run_cedula_valor = p_cedula.add_run(cedula)
        run_cedula_valor.font.color.rgb = RGBColor(0, 0, 0)
        
        p_fecha = doc.add_paragraph()
        run_fecha_label = p_fecha.add_run("Fecha de nacimiento: ")
        run_fecha_label.font.color.rgb = RGBColor(0x44, 0x72, 0xC4)
        run_fecha_label.bold = True
        run_fecha_valor = p_fecha.add_run(fecha)
        run_fecha_valor.font.color.rgb = RGBColor(0, 0, 0)
        
        p_tel = doc.add_paragraph()
        run_tel_label = p_tel.add_run("Teléfono móvil: ")
        run_tel_label.font.color.rgb = RGBColor(0x44, 0x72, 0xC4)
        run_tel_label.bold = True
        run_tel_valor = p_tel.add_run(telefono)
        run_tel_valor.font.color.rgb = RGBColor(0, 0, 0)
        
        p_dir = doc.add_paragraph()
        run_dir_label = p_dir.add_run("Dirección: ")
        run_dir_label.font.color.rgb = RGBColor(0x44, 0x72, 0xC4)
        run_dir_label.bold = True
        run_dir_valor = p_dir.add_run(direccion)
        run_dir_valor.font.color.rgb = RGBColor(0, 0, 0)
        
        p_ciu = doc.add_paragraph()
        run_ciu_label = p_ciu.add_run("Ciudad: ")
        run_ciu_label.font.color.rgb = RGBColor(0x44, 0x72, 0xC4)
        run_ciu_label.bold = True
        run_ciu_valor = p_ciu.add_run(ciudad)
        run_ciu_valor.font.color.rgb = RGBColor(0, 0, 0)
        
        p_est = doc.add_paragraph()
        run_est_label = p_est.add_run("Estado civil: ")
        run_est_label.font.color.rgb = RGBColor(0x44, 0x72, 0xC4)
        run_est_label.bold = True
        run_est_valor = p_est.add_run(estado_civil)
        run_est_valor.font.color.rgb = RGBColor(0, 0, 0)
        
        if correo:
            p_corr = doc.add_paragraph()
            run_corr_label = p_corr.add_run("Correo: ")
            run_corr_label.font.color.rgb = RGBColor(0x44, 0x72, 0xC4)
            run_corr_label.bold = True
            run_corr_valor = p_corr.add_run(correo)
            run_corr_valor.font.color.rgb = RGBColor(0, 0, 0)
        
        # Perfil Profesional - título en azul, negrita, mayúsculas
        if texto_perfil:
            p_perfil_titulo = doc.add_paragraph()
            p_perfil_titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run_perfil_titulo = p_perfil_titulo.add_run("PERFIL PROFESIONAL")
            run_perfil_titulo.bold = True
            run_perfil_titulo.font.size = Pt(12)
            run_perfil_titulo.font.color.rgb = RGBColor(0x44, 0x72, 0xC4)
            doc.add_paragraph()
            
            p_perfil_texto = doc.add_paragraph()
            p_perfil_texto.add_run(texto_perfil)
        
        # Si NO hay experiencia laboral, agregar formación académica en la hoja 1
        if not experiencias:
            # Solo agregar formación académica si hay datos
            if high_school or institution or formaciones:
                doc.add_paragraph()
                doc.add_paragraph()
                
                # Formación Académica - título en azul, negrita, mayúsculas (hoja 1)
                p_formacion_titulo = doc.add_paragraph()
                p_formacion_titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run_formacion_titulo = p_formacion_titulo.add_run("FORMACIÓN ACADÉMICA")
                run_formacion_titulo.bold = True
                run_formacion_titulo.font.size = Pt(12)
                run_formacion_titulo.font.color.rgb = RGBColor(0x44, 0x72, 0xC4)
                doc.add_paragraph()
                
                # Formación académica sin tabla, solo texto alineado
                if high_school or institution:
                    p_sec = doc.add_paragraph()
                    run_sec_label = p_sec.add_run("BACHILLER: ")
                    run_sec_label.font.color.rgb = RGBColor(0x44, 0x72, 0xC4)
                    run_sec_label.bold = True
                    run_sec_valor = p_sec.add_run(high_school)
                    run_sec_valor.font.color.rgb = RGBColor(0, 0, 0)
                    
                    p_inst = doc.add_paragraph()
                    run_inst_label = p_inst.add_run("INSTITUCION: ")
                    run_inst_label.font.color.rgb = RGBColor(0x44, 0x72, 0xC4)
                    run_inst_label.bold = True
                    run_inst_valor = p_inst.add_run(institution)
                    run_inst_valor.font.color.rgb = RGBColor(0, 0, 0)
                
                # Formación técnica/universitaria (puede haber múltiples)
                # Solo agregar si NO es "Bachiller" (ya está arriba)
                for form in formaciones:
                    tipo_form = form.get('tipo', '').strip().upper()
                    nombre_form = form.get('nombre', '').strip()
                    
                    # Filtrar: no mostrar si es "BACHILLER" (ya está en la sección de secundaria)
                    if tipo_form and tipo_form != 'BACHILLER' and nombre_form:
                        doc.add_paragraph()
                        p_tec = doc.add_paragraph()
                        run_tec_label = p_tec.add_run(tipo_form)
                        run_tec_label.font.color.rgb = RGBColor(0x44, 0x72, 0xC4)
                        run_tec_label.bold = True
                        # El valor en la misma línea con dos puntos
                        run_tec_colon = p_tec.add_run(": ")
                        run_tec_colon.font.color.rgb = RGBColor(0x44, 0x72, 0xC4)
                        run_tec_colon.bold = True
                        run_tec_valor = p_tec.add_run(nombre_form)
                        run_tec_valor.font.color.rgb = RGBColor(0, 0, 0)
                
                # Salto de página después de formación académica (inicio de hoja 2 para referencias)
                p_break1 = doc.add_paragraph()
                run_break1 = p_break1.add_run()
                run_break1.add_break(WD_BREAK.PAGE)
        else:
            # Si hay experiencia laboral, salto de página después del perfil profesional (inicio de hoja 2)
            p_break1 = doc.add_paragraph()
            run_break1 = p_break1.add_run()
            run_break1.add_break(WD_BREAK.PAGE)
            
            # Formación Académica - título en azul, negrita, mayúsculas (hoja 2)
            # Solo agregar si hay datos
            if high_school or institution or formaciones:
                p_formacion_titulo = doc.add_paragraph()
                p_formacion_titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run_formacion_titulo = p_formacion_titulo.add_run("FORMACIÓN ACADÉMICA")
                run_formacion_titulo.bold = True
                run_formacion_titulo.font.size = Pt(12)
                run_formacion_titulo.font.color.rgb = RGBColor(0x44, 0x72, 0xC4)
                doc.add_paragraph()
                
                # Formación académica sin tabla, solo texto alineado
                if high_school or institution:
                    p_sec = doc.add_paragraph()
                    run_sec_label = p_sec.add_run("BACHILLER: ")
                    run_sec_label.font.color.rgb = RGBColor(0x44, 0x72, 0xC4)
                    run_sec_label.bold = True
                    run_sec_valor = p_sec.add_run(high_school)
                    run_sec_valor.font.color.rgb = RGBColor(0, 0, 0)
                    
                    p_inst = doc.add_paragraph()
                    run_inst_label = p_inst.add_run("INSTITUCION: ")
                    run_inst_label.font.color.rgb = RGBColor(0x44, 0x72, 0xC4)
                    run_inst_label.bold = True
                    run_inst_valor = p_inst.add_run(institution)
                    run_inst_valor.font.color.rgb = RGBColor(0, 0, 0)
                
                # Formación técnica/universitaria (puede haber múltiples)
                # Solo agregar si NO es "Bachiller" (ya está arriba)
                for form in formaciones:
                    tipo_form = form.get('tipo', '').strip().upper()
                    nombre_form = form.get('nombre', '').strip()
                    
                    # Filtrar: no mostrar si es "BACHILLER" (ya está en la sección de secundaria)
                    if tipo_form and tipo_form != 'BACHILLER' and nombre_form:
                        doc.add_paragraph()
                        p_tec = doc.add_paragraph()
                        run_tec_label = p_tec.add_run(tipo_form)
                        run_tec_label.font.color.rgb = RGBColor(0x44, 0x72, 0xC4)
                        run_tec_label.bold = True
                        # El valor en la misma línea con dos puntos
                        run_tec_colon = p_tec.add_run(": ")
                        run_tec_colon.font.color.rgb = RGBColor(0x44, 0x72, 0xC4)
                        run_tec_colon.bold = True
                        run_tec_valor = p_tec.add_run(nombre_form)
                        run_tec_valor.font.color.rgb = RGBColor(0, 0, 0)
                
                doc.add_paragraph()
                doc.add_paragraph()
        
        # Experiencia Laboral - título en azul, negrita, mayúsculas, centrado (hoja 2, solo si hay experiencia)
        if experiencias:
            p_exp_titulo = doc.add_paragraph()
            p_exp_titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run_exp_titulo = p_exp_titulo.add_run("EXPERIENCIA LABORAL")
            run_exp_titulo.bold = True
            run_exp_titulo.font.size = Pt(12)
            run_exp_titulo.font.color.rgb = RGBColor(0x44, 0x72, 0xC4)
            doc.add_paragraph()
            
            for experiencia in experiencias:
                # Las experiencias pueden venir con 'empresa' o 'local', 'cargo' o 'cargo', 'tiempo' o 'fechaInicio/fechaFin'
                empresa = experiencia.get('empresa', experiencia.get('local', '')).strip()
                cargo = experiencia.get('cargo', '').strip()
                tiempo = experiencia.get('tiempo', '')
                
                # Si no viene tiempo, construirlo desde fechaInicio y fechaFin
                if not tiempo:
                    fecha_inicio = experiencia.get('fechaInicio', '').strip()
                    fecha_fin = experiencia.get('fechaFin', '').strip()
                    if fecha_inicio and fecha_fin:
                        tiempo = f"Desde {fecha_inicio} hasta {fecha_fin}"
                
                if empresa and cargo:
                    p_estab = doc.add_paragraph()
                    run_estab_label = p_estab.add_run("ESTABLECIMIENTO: ")
                    run_estab_label.font.color.rgb = RGBColor(0x44, 0x72, 0xC4)
                    run_estab_label.bold = True
                    run_estab_valor = p_estab.add_run(empresa)
                    run_estab_valor.font.color.rgb = RGBColor(0, 0, 0)
                    
                    p_cargo = doc.add_paragraph()
                    run_cargo_label = p_cargo.add_run("CARGO: ")
                    run_cargo_label.font.color.rgb = RGBColor(0x44, 0x72, 0xC4)
                    run_cargo_label.bold = True
                    run_cargo_valor = p_cargo.add_run(cargo)
                    run_cargo_valor.font.color.rgb = RGBColor(0, 0, 0)
                    
                    if tiempo:
                        p_periodo = doc.add_paragraph()
                        run_periodo_label = p_periodo.add_run("PERIODO LABORAL: ")
                        run_periodo_label.font.color.rgb = RGBColor(0x44, 0x72, 0xC4)
                        run_periodo_label.bold = True
                        run_periodo_valor = p_periodo.add_run(tiempo)
                        run_periodo_valor.font.color.rgb = RGBColor(0, 0, 0)
                    
                    doc.add_paragraph()
                    doc.add_paragraph()
            
            # Si hay experiencia, salto de página para referencias (hoja 3)
            p_break2 = doc.add_paragraph()
            run_break2 = p_break2.add_run()
            run_break2.add_break(WD_BREAK.PAGE)
        
        # Referencias Familiares - título en azul, negrita, mayúsculas, centrado
        # Solo agregar si hay referencias familiares
        if referencias_familiares:
            p_ref_fam_titulo = doc.add_paragraph()
            p_ref_fam_titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run_ref_fam_titulo = p_ref_fam_titulo.add_run("REFERENCIAS FAMILIARES")
            run_ref_fam_titulo.bold = True
            run_ref_fam_titulo.font.size = Pt(12)
            run_ref_fam_titulo.font.color.rgb = RGBColor(0x44, 0x72, 0xC4)
            doc.add_paragraph()
            
            for ref in referencias_familiares:
                nombre_ref = ref.get('nombre', '').strip()
                telefono_ref = ref.get('telefono', ref.get('celular', '')).strip()
                
                if nombre_ref:
                    p_ref_fam_nombre = doc.add_paragraph()
                    run_ref_fam_nombre = p_ref_fam_nombre.add_run(nombre_ref)
                    run_ref_fam_nombre.font.color.rgb = RGBColor(0x44, 0x72, 0xC4)
                    run_ref_fam_nombre.italic = True
                    
                    if telefono_ref:
                        p_ref_fam_tel = doc.add_paragraph()
                        run_ref_fam_tel_label = p_ref_fam_tel.add_run("Teléfono: ")
                        run_ref_fam_tel_label.font.color.rgb = RGBColor(0x44, 0x72, 0xC4)
                        run_ref_fam_tel_label.bold = True
                        run_ref_fam_tel_valor = p_ref_fam_tel.add_run(telefono_ref)
                        run_ref_fam_tel_valor.font.color.rgb = RGBColor(0, 0, 0)
                        run_ref_fam_tel_valor.bold = True
                    
                    doc.add_paragraph()
        
        # Referencias Personales - título en azul, negrita, mayúsculas, centrado
        # Solo agregar si hay referencias personales
        if referencias_personales:
            p_ref_per_titulo = doc.add_paragraph()
            p_ref_per_titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run_ref_per_titulo = p_ref_per_titulo.add_run("REFERENCIAS PERSONALES")
            run_ref_per_titulo.bold = True
            run_ref_per_titulo.font.size = Pt(12)
            run_ref_per_titulo.font.color.rgb = RGBColor(0x44, 0x72, 0xC4)
            doc.add_paragraph()
            
            for ref in referencias_personales:
                nombre_ref = ref.get('nombre', '').strip()
                telefono_ref = ref.get('telefono', ref.get('celular', '')).strip()
                
                if nombre_ref:
                    p_ref_per_nombre = doc.add_paragraph()
                    run_ref_per_nombre = p_ref_per_nombre.add_run(nombre_ref)
                    run_ref_per_nombre.font.color.rgb = RGBColor(0x44, 0x72, 0xC4)
                    run_ref_per_nombre.italic = True
                    
                    if telefono_ref:
                        p_ref_per_tel = doc.add_paragraph()
                        run_ref_per_tel_label = p_ref_per_tel.add_run("Teléfono: ")
                        run_ref_per_tel_label.font.color.rgb = RGBColor(0x44, 0x72, 0xC4)
                        run_ref_per_tel_label.bold = True
                        run_ref_per_tel_valor = p_ref_per_tel.add_run(telefono_ref)
                        run_ref_per_tel_valor.font.color.rgb = RGBColor(0, 0, 0)
                        run_ref_per_tel_valor.bold = True
                    
                    doc.add_paragraph()
        
        # Espacios finales antes del pie de página
        doc.add_paragraph()
        doc.add_paragraph()
        doc.add_paragraph()
        
        # Pie de página con nombre en azul, negrita, mayúsculas
        p_final = doc.add_paragraph()
        run_final = p_final.add_run(nombre.upper())
        run_final.font.color.rgb = RGBColor(0x44, 0x72, 0xC4)
        run_final.bold = True
        
        p_cedula_final = doc.add_paragraph()
        run_cedula_final = p_cedula_final.add_run(f"C.C. {cedula} de {exp}")
        run_cedula_final.font.color.rgb = RGBColor(0, 0, 0)
        
        # Guardar en memoria
        output = io.BytesIO()
        doc.save(output)
        output.seek(0)
        
        # Nombre del archivo
        nombre_archivo = nombre.replace(' ', '_') if nombre else 'Hoja_de_Vida'
        filename = f"HV_{nombre_archivo}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
        
        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            as_attachment=True,
            download_name=filename
        )
        
    except Exception as e:
        import traceback
        return jsonify({"error": str(e), "traceback": traceback.format_exc()}), 500

def reemplazar_texto_en_documento(doc, reemplazos):
    """
    Reemplaza texto en un documento Word manteniendo el formato.
    Busca en párrafos y tablas. Busca placeholders de forma case-insensitive.
    Mejora: Busca en todos los runs de texto para encontrar variables divididas.
    """
    import re
    
    def reemplazar_en_parrafo(paragraph, reemplazos_dict):
        """Reemplaza texto en un párrafo manteniendo formato, especialmente negrilla"""
        # Obtener todo el texto del párrafo
        texto_original = paragraph.text
        if not texto_original:
            return
        
        # Trabajar directamente con los runs para conservar formato individual
        if not paragraph.runs:
            return
        
        # Buscar y reemplazar en cada run, conservando formato
        sorted_reemplazos = sorted(reemplazos_dict.items(), key=lambda x: len(x[0]), reverse=True)
        
        for placeholder, valor in sorted_reemplazos:
            if not placeholder:
                continue
            
            # Si el placeholder tiene formato {{VARIABLE}}, buscar solo en mayúsculas y case-sensitive
            if placeholder.startswith('{{') and placeholder.endswith('}}'):
                pattern = re.escape(placeholder)
                case_sensitive = True
            # Para dia1 y dia2, buscar sin word boundaries para capturar variaciones
            elif 'dia' in placeholder.lower() and len(placeholder) <= 5:
                pattern = re.escape(placeholder)
                case_sensitive = False
            elif placeholder.upper() in ['4 TURNOS', 'ADICIONALES', 'SUELDO FIJO MENSUAL', 'SUELDO PROPORCIONAL', 
                                         'BONO SEGURIDAD', 'AUXILIO DE TRANSPORTE', 'TURNOS', 'DESCANSOS']:
                continue
            elif 'SUELDO FIJO' in placeholder.upper() or 'SUELDO PROPORCIONAL' in placeholder.upper():
                continue
            elif 'BONO SEGURIDAD' in placeholder.upper() or 'AUXILIO' in placeholder.upper():
                continue
            else:
                pattern = re.escape(placeholder)
                case_sensitive = False
            
            # Buscar el placeholder en todos los runs
            texto_completo = ''.join([run.text for run in paragraph.runs])
            flags = 0 if case_sensitive else re.IGNORECASE
            matches = list(re.finditer(pattern, texto_completo, flags))
            
            if matches:
                # Reemplazar desde el final hacia el inicio para mantener índices
                for match in reversed(matches):
                    start, end = match.span()
                    valor_reemplazo = str(valor)
                    
                    # Para dia1 y dia2, verificar si necesita espacios alrededor
                    if 'dia' in placeholder.lower() and len(placeholder) <= 5:
                        tiene_espacio_antes = start > 0 and texto_completo[start-1].isspace()
                        tiene_espacio_despues = end < len(texto_completo) and texto_completo[end].isspace()
                        
                        if valor_reemplazo.startswith(' ') and tiene_espacio_antes:
                            valor_reemplazo = valor_reemplazo.lstrip()
                        if valor_reemplazo.endswith(' ') and tiene_espacio_despues:
                            valor_reemplazo = valor_reemplazo.rstrip()
                        
                        if not tiene_espacio_antes and not valor_reemplazo.startswith(' '):
                            if start > 0 and texto_completo[start-1].isalnum():
                                valor_reemplazo = ' ' + valor_reemplazo
                        if not tiene_espacio_despues and not valor_reemplazo.endswith(' '):
                            if end < len(texto_completo) and texto_completo[end].isalnum():
                                valor_reemplazo = valor_reemplazo + ' '
                    
                    # Encontrar en qué run(s) está el placeholder y conservar formato (especialmente negrilla)
                    current_pos = 0
                    formato_aplicar = None
                    placeholder_encontrado = False
                    
                    for run in paragraph.runs:
                        run_start = current_pos
                        run_end = current_pos + len(run.text)
                        
                        # Si el placeholder está completamente en este run
                        if run_start <= start < run_end and run_start <= end <= run_end:
                            # Conservar formato del run (incluyendo negrilla)
                            formato_aplicar = {
                                'font_name': run.font.name if run.font.name else None,
                                'font_size': run.font.size if run.font.size else None,
                                'bold': run.bold if run.bold is not None else False,
                                'italic': run.italic if run.italic is not None else False,
                                'color': run.font.color.rgb if run.font.color and run.font.color.rgb else None
                            }
                            
                            # Reemplazar en el run - el formato se conserva automáticamente
                            run_start_in_run = start - run_start
                            run_end_in_run = end - run_start
                            nuevo_texto = run.text[:run_start_in_run] + valor_reemplazo + run.text[run_end_in_run:]
                            run.text = nuevo_texto
                            placeholder_encontrado = True
                            break
                        
                        # Si el placeholder comienza en este run (puede estar dividido)
                        elif run_start <= start < run_end and formato_aplicar is None:
                            # Usar el formato del run donde comienza el placeholder (conservar negrilla)
                            formato_aplicar = {
                                'font_name': run.font.name if run.font.name else None,
                                'font_size': run.font.size if run.font.size else None,
                                'bold': run.bold if run.bold is not None else False,
                                'italic': run.italic if run.italic is not None else False,
                                'color': run.font.color.rgb if run.font.color and run.font.color.rgb else None
                            }
                        
                        current_pos = run_end
                    
                    # Si el placeholder estaba dividido entre múltiples runs, reconstruir conservando formato
                    if not placeholder_encontrado and formato_aplicar:
                        # Reconstruir texto completo
                        texto_completo_nuevo = ''.join([run.text for run in paragraph.runs])
                        if placeholder in texto_completo_nuevo:
                            # Reemplazar en el texto completo
                            texto_completo_nuevo = texto_completo_nuevo.replace(placeholder, valor_reemplazo, 1)
                            
                            # Limpiar y reconstruir con el formato del run donde estaba el placeholder
                            paragraph.clear()
                            nuevo_run = paragraph.add_run(texto_completo_nuevo)
                            
                            # Aplicar formato conservado (incluyendo negrilla)
                            if formato_aplicar:
                                if formato_aplicar['font_name']:
                                    nuevo_run.font.name = formato_aplicar['font_name']
                                if formato_aplicar['font_size']:
                                    nuevo_run.font.size = formato_aplicar['font_size']
                                if formato_aplicar['color']:
                                    nuevo_run.font.color.rgb = formato_aplicar['color']
                                nuevo_run.bold = formato_aplicar['bold']
                                nuevo_run.italic = formato_aplicar['italic']
    
    def reemplazar_en_runs(runs, reemplazos_dict):
        """Reemplaza texto en runs individuales conservando formato, especialmente negrilla"""
        texto_completo = ''.join([run.text for run in runs])
        if not texto_completo:
            return False
        
        texto_nuevo = texto_completo
        cambios_realizados = False
        formato_aplicar = None
        
        # Ordenar por longitud descendente
        sorted_reemplazos = sorted(reemplazos_dict.items(), key=lambda x: len(x[0]), reverse=True)
        
        for placeholder, valor in sorted_reemplazos:
            if not placeholder:
                continue
            
            # Si el placeholder tiene formato {{VARIABLE}}, buscar solo en mayúsculas y case-sensitive
            if placeholder.startswith('{{') and placeholder.endswith('}}'):
                pattern = re.escape(placeholder)
                matches = list(re.finditer(pattern, texto_nuevo))  # Case-sensitive
            else:
                pattern = re.escape(placeholder)
                matches = list(re.finditer(pattern, texto_nuevo, re.IGNORECASE))
            
            if matches:
                cambios_realizados = True
                # Encontrar el formato del run donde está el placeholder (conservar negrilla)
                for match in reversed(matches):
                    start, end = match.span()
                    
                    # Encontrar en qué run está el placeholder para conservar su formato
                    current_pos = 0
                    for run in runs:
                        run_start = current_pos
                        run_end = current_pos + len(run.text)
                        
                        if run_start <= start < run_end and formato_aplicar is None:
                            # Conservar formato del run donde está el placeholder (incluyendo negrilla)
                            formato_aplicar = {
                                'font_name': run.font.name if run.font.name else None,
                                'font_size': run.font.size if run.font.size else None,
                                'bold': run.bold if run.bold is not None else False,
                                'italic': run.italic if run.italic is not None else False,
                                'color': run.font.color.rgb if run.font.color and run.font.color.rgb else None
                            }
                            break
                        
                        current_pos = run_end
                    
                    texto_nuevo = texto_nuevo[:start] + str(valor) + texto_nuevo[end:]
        
        if cambios_realizados and texto_nuevo != texto_completo:
            # Limpiar todos los runs y crear uno nuevo con el texto reemplazado
            for run in runs:
                run.text = ''
            if runs:
                runs[0].text = texto_nuevo
                # Aplicar formato conservado (incluyendo negrilla)
                if formato_aplicar:
                    if formato_aplicar['font_name']:
                        runs[0].font.name = formato_aplicar['font_name']
                    if formato_aplicar['font_size']:
                        runs[0].font.size = formato_aplicar['font_size']
                    if formato_aplicar['color']:
                        runs[0].font.color.rgb = formato_aplicar['color']
                    runs[0].bold = formato_aplicar['bold']
                    runs[0].italic = formato_aplicar['italic']
            return True
        
        return False
    
    # Reemplazar en párrafos
    for paragraph in doc.paragraphs:
        # Primero intentar reemplazo en el párrafo completo
        reemplazar_en_parrafo(paragraph, reemplazos)
        
        # También intentar reemplazo en runs individuales (por si la variable está dividida)
        if paragraph.runs and len(paragraph.runs) > 1:
            reemplazar_en_runs(paragraph.runs, reemplazos)
    
    # Reemplazar en tablas
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    reemplazar_en_parrafo(paragraph, reemplazos)
                    # También en runs individuales
                    if paragraph.runs and len(paragraph.runs) > 1:
                        reemplazar_en_runs(paragraph.runs, reemplazos)

def formatear_monto(monto, incluir_signo=True):
    """Formatea un monto como moneda colombiana"""
    if not monto:
        return ''
    try:
        monto_num = float(monto)
        # Formatear sin el símbolo $ para evitar duplicaciones en el template
        monto_formateado = f"{monto_num:,.0f}".replace(',', '.')
        if incluir_signo:
            return f"${monto_formateado}"
        return monto_formateado
    except:
        return str(monto)

def sanitize_input(value, max_length=500):
    """Sanitiza un input removiendo caracteres peligrosos y limitando longitud"""
    if not value:
        return ''
    # Remover caracteres de control y limitar longitud
    sanitized = ''.join(c for c in str(value) if c.isprintable() or c in ['\n', '\r', '\t'])[:max_length]
    return sanitized.strip()

def validate_numeric(value, min_val=None, max_val=None, default=0):
    """Valida y convierte un valor numérico"""
    try:
        num = float(str(value).replace(',', '.').replace('.', '', str(value).count('.') - 1))
        if min_val is not None and num < min_val:
            return default
        if max_val is not None and num > max_val:
            return default
        return num
    except (ValueError, AttributeError):
        return default

@app.route('/generate-cuenta-cobro', methods=['POST'])
def generate_cuenta_cobro():
    """Genera una cuenta de cobro usando el template Word"""
    try:
        if not request.json:
            return jsonify({'error': 'No se recibieron datos'}), 400
        
        data = request.json
        
        # Validar y sanitizar datos del formulario
        nombre = sanitize_input(data.get('nombre', ''), max_length=200)
        if not nombre:
            return jsonify({'error': 'El nombre es obligatorio'}), 400
        
        cedula = sanitize_input(data.get('cedula', ''), max_length=50)
        if not cedula:
            return jsonify({'error': 'La cédula es obligatoria'}), 400
        
        telefono = sanitize_input(data.get('phone', '') or data.get('telefono', '') or data.get('phoneNumber', ''), max_length=20)
        # Remover print de debug con datos sensibles en producción
        # print(f"📞 Teléfono recibido: '{telefono}'")  # Debug - removido por seguridad
        mes = sanitize_input(data.get('mes', ''), max_length=50)
        año = sanitize_input(data.get('año', ''), max_length=10)
        
        # Validar valores numéricos
        mes_completo = bool(data.get('mesCompleto', True))
        dia_inicio = str(int(validate_numeric(data.get('diaInicio', '1'), min_val=1, max_val=31, default=1)))
        dia_fin = str(int(validate_numeric(data.get('diaFin', '30'), min_val=1, max_val=31, default=30)))
        
        # Calcular días trabajados
        dias_num = 30  # Valor por defecto
        if not mes_completo:
            try:
                dia_inicio_num = int(dia_inicio)
                dia_fin_num = int(dia_fin)
                if dia_fin_num >= dia_inicio_num:
                    dias_num = (dia_fin_num - dia_inicio_num) + 1
                else:
                    dias_num = 30
            except (ValueError, TypeError):
                dias_num = 30
        else:
            # Si es mes completo, usar el valor del campo o calcular desde el mes
            try:
                dias_trabajados_input = data.get('diasTrabajados', '')
                if dias_trabajados_input:
                    dias_num = int(validate_numeric(dias_trabajados_input, min_val=1, max_val=31, default=30))
                else:
                    dias_num = 30
            except:
                dias_num = 30
        
        # Obtener el número de días del mes seleccionado
        try:
            mes_num = int(mes) if mes.isdigit() else 0
            año_num = int(año) if año.isdigit() else datetime.now().year
            if 1 <= mes_num <= 12:
                from calendar import monthrange
                dias_del_mes = monthrange(año_num, mes_num)[1]
            else:
                dias_del_mes = 30
        except:
            dias_del_mes = 30
        
        # Limitar días trabajados al máximo de días del mes
        if dias_num > dias_del_mes:
            dias_num = dias_del_mes
        if dias_num < 1:
            dias_num = 30
        
        # Parsear sueldo fijo correctamente (soporta formatos: "2000000", "2.000.000", "2,000,000")
        sueldo_fijo_num = 0
        try:
            sueldo_fijo_raw = str(data.get('sueldoFijo', '0')).strip()
            if sueldo_fijo_raw:
                # Remover puntos (separadores de miles) y reemplazar coma por punto (decimal)
                sueldo_fijo_limpio = sueldo_fijo_raw.replace('.', '').replace(',', '.')
                sueldo_fijo_num = float(sueldo_fijo_limpio)
                if sueldo_fijo_num < 0:
                    sueldo_fijo_num = 0
                if sueldo_fijo_num > 10000000:
                    sueldo_fijo_num = 10000000
        except (ValueError, TypeError):
            sueldo_fijo_num = 0
        
        # Calcular sueldo proporcional según días trabajados
        sueldo_proporcional = 0
        if sueldo_fijo_num > 0 and dias_del_mes > 0:
            valor_por_dia = sueldo_fijo_num / dias_del_mes
            sueldo_proporcional = round(valor_por_dia * dias_num)
        
        # Validar y sanitizar valores monetarios y otros campos
        turnos_descansos = str(int(validate_numeric(data.get('turnosDescansos', '0'), min_val=0, max_val=100, default=0)))
        paciente = sanitize_input(data.get('paciente', '') or data.get('patientName', ''), max_length=200)
        cuenta_bancaria = sanitize_input(data.get('cuentaBancaria', '') or data.get('bankAccount', ''), max_length=50)
        banco = sanitize_input(data.get('banco', ''), max_length=100).upper() or 'Bancolombia'
        tipo_cuenta_cobro = sanitize_input(data.get('tipoCuentaCobro', '12h'), max_length=10)
        if tipo_cuenta_cobro not in ['12h', '8h']:
            tipo_cuenta_cobro = '12h'
        tiene_auxilio_transporte = bool(data.get('tieneAuxilioTransporte', False))
        
        # Parsear bono de seguridad CORRECTAMENTE
        bono_seguridad_num = 0
        try:
            bono_raw = data.get('bonoSeguridad', '0')
            if bono_raw:
                # Convertir a string y limpiar
                bono_str = str(bono_raw).strip()
                # Remover puntos (separadores de miles) y reemplazar coma por punto (decimal)
                bono_limpio = bono_str.replace('.', '').replace(',', '.')
                bono_seguridad_num = float(bono_limpio)
                # Validar rango
                if bono_seguridad_num < 0:
                    bono_seguridad_num = 0
                if bono_seguridad_num > 10000000:
                    bono_seguridad_num = 10000000
        except (ValueError, TypeError, AttributeError):
            bono_seguridad_num = 0
        
        # Parsear auxilio de transporte
        auxilio_transporte_num = 0
        if tiene_auxilio_transporte:
            try:
                auxilio_raw = data.get('auxilioTransporte', '0')
                if auxilio_raw:
                    auxilio_str = str(auxilio_raw).strip()
                    auxilio_limpio = auxilio_str.replace('.', '').replace(',', '.')
                    auxilio_transporte_num = float(auxilio_limpio)
                    if auxilio_transporte_num < 0:
                        auxilio_transporte_num = 0
                    if auxilio_transporte_num > 10000000:
                        auxilio_transporte_num = 10000000
            except (ValueError, TypeError, AttributeError):
                auxilio_transporte_num = 0
        
        # Calcular adicionales (turnos * 60000)
        turnos_num = int(turnos_descansos) if turnos_descansos.isdigit() else 0
        valor_por_turno = 60000
        adicionales_valor = turnos_num * valor_por_turno
        
        # El total se calculará después de formatear los valores
        
        # Formatear fecha (mes en texto)
        fecha_texto = ''
        if mes and año:
            mes_num = int(mes) if mes.isdigit() else 0
            if 1 <= mes_num <= 12:
                fecha_texto = f"{MESES[mes_num].upper()} DE {año}"
        
        # Cargar template
        # Seleccionar template según tipo de cuenta de cobro
        if tipo_cuenta_cobro == '8h':
            template_path = os.path.join(os.path.dirname(__file__), 'templates', 'cobro_8h.docx')
        else:
            template_path = os.path.join(os.path.dirname(__file__), 'templates', 'cobro_ 2026.docx')
        if not os.path.exists(template_path):
            return jsonify({"error": f"Template no encontrado en: {template_path}"}), 404
        
        doc = Document(template_path)
        
        # Preparar reemplazos usando los placeholders exactos del template
        # Buscar todas las variaciones posibles de las variables
        reemplazos = {}
        
        # ============================================
        # FORMATEO DE VALORES - CÓDIGO NUEVO DESDE CERO
        # ============================================
        
        # Variable sf1: Sueldo proporcional según días trabajados
        # Ejemplo: sueldo fijo 2.000.000, enero 31 días = 2.000.000, febrero 28 días = 2.000.000
        sf1_valor = sueldo_proporcional
        sf1_formateado = formatear_monto(sf1_valor, incluir_signo=False)
        
        # Variable bs1: Bono de seguridad
        # Ejemplo: 200.000
        bs1_valor = bono_seguridad_num
        bs1_formateado = formatear_monto(bs1_valor, incluir_signo=False) if bs1_valor > 0 else ''
        
        # Variable ad1: Adicionales (turnos de descansos)
        # Ejemplo: 4 turnos * 60.000 = 240.000
        ad1_valor = adicionales_valor
        ad1_formateado = formatear_monto(ad1_valor, incluir_signo=False) if ad1_valor > 0 else ''
        
        # Variable ax1: Auxilio de transporte
        ax1_valor = auxilio_transporte_num
        ax1_formateado = formatear_monto(ax1_valor, incluir_signo=False) if ax1_valor > 0 else ''
        
        # Calcular TOTAL: sf1 + bs1 + ad1 + ax1
        total_calculado = sf1_valor + bs1_valor + ad1_valor + ax1_valor
        total_formateado = formatear_monto(total_calculado, incluir_signo=False)
        
        # Solo reemplazar variables con formato {{VARIABLE}} (llaves dobles y mayúsculas)
        # Nombre
        reemplazos['{{Name1}}'] = nombre.upper()
        
        # Cédula
        reemplazos['{{Cedu1}}'] = cedula
        
        # Teléfono - asegurar que siempre se reemplace, incluso si está vacío
        telefono_valor = telefono if telefono else ''
        reemplazos['{{Num1}}'] = telefono_valor
        
        # Banco
        reemplazos['{{banco1}}'] = banco
        
        # Número de cuenta bancaria
        reemplazos['{{nbanco1}}'] = cuenta_bancaria
        
        # Cuenta bancaria
        reemplazos['{{cuenta1}}'] = cuenta_bancaria
        
        # Mes y año
        reemplazos['{{mes1}}'] = fecha_texto
        
        # Total/Valor
        reemplazos['{{valor1}}'] = total_formateado
        reemplazos['{{total1}}'] = total_formateado
        
        # Paciente
        paciente_valor = paciente.upper() if paciente else ''
        reemplazos['{{paciente1}}'] = paciente_valor
        
        # Variable sf1: Sueldo proporcional (NO incluye bono)
        reemplazos['{{sf1}}'] = sf1_formateado
        
        # Días trabajados - mantener texto descriptivo y variable con llaves
        dias_texto = f"{dias_num} DÍAS" if dias_num < 30 else 'MES COMPLETO'
        reemplazos['MES COMPLETO'] = dias_texto
        reemplazos['30 DÍAS'] = dias_texto
        reemplazos['30 DIAS'] = dias_texto
        reemplazos['{{dias1}}'] = str(dias_num)
        
        # Variable dia1 y dia2 - dia1 es el día de inicio, dia2 es el último día del mes
        # Obtener diaInicio de los datos
        dia_inicio = data.get('diaInicio', '1').strip() if data.get('diaInicio') else '1'
        # dia2 debe ser el último día del mes (ej: enero=31, febrero=28)
        dia_fin = str(dias_del_mes)
        
        # dia1 siempre es el día de inicio - múltiples variaciones para asegurar el reemplazo
        # Agregar espacios alrededor para evitar que quede pegado
        reemplazos['dia1'] = f' {dia_inicio} '  # Agregar espacios alrededor
        reemplazos['DIA1'] = f' {dia_inicio} '
        reemplazos['{dia1}'] = dia_inicio
        reemplazos['{{dia1}}'] = dia_inicio
        reemplazos['[dia1]'] = dia_inicio
        reemplazos['<<dia1>>'] = dia_inicio
        reemplazos[' dia1 '] = f' {dia_inicio} '  # Mantener espacios
        reemplazos[' DIA1 '] = f' {dia_inicio} '
        reemplazos['dia1 '] = f' {dia_inicio} '  # Mantener espacio al final
        reemplazos[' dia1'] = f' {dia_inicio} '  # Mantener espacio al inicio
        reemplazos['diaInicio'] = dia_inicio
        reemplazos['dia_inicio'] = dia_inicio
        
        # dia2 siempre es el último día del mes - múltiples variaciones para asegurar el reemplazo
        # Agregar espacios alrededor para evitar que quede pegado
        reemplazos['dia2'] = f' {dia_fin} '  # Último día del mes (ej: enero=31, febrero=28)
        reemplazos['DIA2'] = f' {dia_fin} '
        reemplazos['{dia2}'] = dia_fin
        reemplazos['{{dia2}}'] = dia_fin
        reemplazos['[dia2]'] = dia_fin
        reemplazos['<<dia2>>'] = dia_fin
        reemplazos[' dia2 '] = f' {dia_fin} '  # Mantener espacios
        reemplazos[' DIA2 '] = f' {dia_fin} '
        reemplazos['dia2 '] = f' {dia_fin} '  # Mantener espacio al final
        reemplazos[' dia2'] = f' {dia_fin} '  # Mantener espacio al inicio
        reemplazos['diaFin'] = dia_fin
        reemplazos['dia_fin'] = dia_fin
        
        # Patrones comunes con "al" para reemplazar correctamente
        reemplazos[f'dia1 al dia2'] = f'{dia_inicio} al {dia_fin}'
        reemplazos[f'DIA1 AL DIA2'] = f'{dia_inicio} al {dia_fin}'
        reemplazos[f'dia1al dia2'] = f'{dia_inicio} al {dia_fin}'
        reemplazos[f'dia1al dia2'] = f'{dia_inicio} al {dia_fin}'
        reemplazos[f'dia1al dia2'] = f'{dia_inicio} al {dia_fin}'
        
        # Variable bs1/sb1: Bono de seguridad (200.000)
        # Variable principal: bs1
        reemplazos['{{bs1}}'] = bs1_formateado
        # Variable alternativa: sb1 (usada en el template)
        reemplazos['{{sb1}}'] = bs1_formateado
        
        # Variable ad1: Adicionales
        reemplazos['{{ad1}}'] = ad1_formateado
        # NO reemplazar "ADICIONALES" - debe mantenerse como texto descriptivo
        
        # Variable ax1: Auxilio de transporte
        reemplazos['{{ax1}}'] = ax1_formateado
        
        # Limpiar duplicaciones de texto comunes ANTES de reemplazar
        # Duplicaciones de año - múltiples variaciones (ordenar por longitud descendente)
        # Primero los más largos para evitar reemplazos parciales
        reemplazos['DE ' + año + ' DEL ' + año] = f'DE {año}'
        reemplazos['DEL ' + año + ' DE ' + año] = f'DEL {año}'
        reemplazos['DE ' + año + ' DE ' + año] = f'DE {año}'
        reemplazos['DEL ' + año + ' DEL ' + año] = f'DEL {año}'
        # También valores hardcodeados comunes
        reemplazos['DE 2026 DEL 2026'] = f'DE {año}'
        reemplazos['DEL 2026 DE 2026'] = f'DEL {año}'
        reemplazos['DE 2026 DE 2026'] = f'DE {año}'
        reemplazos['DEL 2026 DEL 2026'] = f'DEL {año}'
        
        # Log de reemplazos para debug - especialmente dia1 y dia2
        print(f"🔍 Reemplazos a realizar: {len(reemplazos)} variables")
        print(f"📅 dia1 (día inicio): '{dia_inicio}'")
        print(f"📅 dia2 (día fin): '{dia_fin}'")
        # Debug removido por seguridad - comentado para producción
        # for key, value in sorted(reemplazos.items()):
        #     if value and ('dia' in key.lower() or 'DIA' in key):
        #         print(f"  - {key} -> {value}")
        
        # Reemplazar texto en el documento
        reemplazar_texto_en_documento(doc, reemplazos)
        print("✅ Reemplazos completados en el documento")
        
        # Verificar si dia1 y dia2 fueron reemplazados correctamente
        texto_completo = ' '.join([para.text for para in doc.paragraphs])
        if 'dia1' in texto_completo.lower() or 'dia2' in texto_completo.lower():
            # Debug removido por seguridad - solo en desarrollo
            # print(f"⚠️ ADVERTENCIA: Todavía hay 'dia1' o 'dia2' sin reemplazar en el documento")
            # print(f"   Texto encontrado: {texto_completo[texto_completo.lower().find('dia'):texto_completo.lower().find('dia')+50]}")
            pass
        
        # Limpiar duplicaciones después del reemplazo
        # Buscar y limpiar patrones comunes de duplicación
        import re
        for paragraph in doc.paragraphs:
            texto = paragraph.text
            # Limpiar duplicaciones de año - múltiples patrones (aplicar varias veces para asegurar)
            # Patrón 1: DE 2026 DE 2026 -> DE 2026
            texto = re.sub(r'DE (\d{4}) DE \1', r'DE \1', texto)
            # Patrón 2: DEL 2026 DEL 2026 -> DEL 2026
            texto = re.sub(r'DEL (\d{4}) DEL \1', r'DEL \1', texto)
            # Patrón 3: DE 2026 DEL 2026 -> DE 2026 (el más común)
            texto = re.sub(r'DE (\d{4}) DEL \1', r'DE \1', texto)
            # Patrón 4: DEL 2026 DE 2026 -> DEL 2026
            texto = re.sub(r'DEL (\d{4}) DE \1', r'DEL \1', texto)
            # Aplicar nuevamente para casos anidados
            texto = re.sub(r'DE (\d{4}) DEL \1', r'DE \1', texto)
            texto = re.sub(r'DEL (\d{4}) DE \1', r'DEL \1', texto)
            # Limpiar múltiples símbolos $ seguidos
            texto = re.sub(r'\$\$+', '$', texto)
            texto = re.sub(r'\$ \$+', '$', texto)
            # Limpiar espacios múltiples
            texto = re.sub(r'  +', ' ', texto)
            
            if texto != paragraph.text:
                # Guardar formato
                formato_original = None
                if paragraph.runs:
                    primer_run = paragraph.runs[0]
                    formato_original = {
                        'font_name': primer_run.font.name if primer_run.font.name else None,
                        'font_size': primer_run.font.size if primer_run.font.size else None,
                        'bold': primer_run.bold if primer_run.bold is not None else False,
                        'italic': primer_run.italic if primer_run.italic is not None else False,
                        'color': primer_run.font.color.rgb if primer_run.font.color and primer_run.font.color.rgb else None
                    }
                
                paragraph.clear()
                nuevo_run = paragraph.add_run(texto)
                
                if formato_original:
                    if formato_original['font_name']:
                        nuevo_run.font.name = formato_original['font_name']
                    if formato_original['font_size']:
                        nuevo_run.font.size = formato_original['font_size']
                    if formato_original['color']:
                        nuevo_run.font.color.rgb = formato_original['color']
                    nuevo_run.bold = formato_original['bold']
                    nuevo_run.italic = formato_original['italic']
        
        # También limpiar en tablas
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        texto = paragraph.text
                        # Limpiar duplicaciones de año - múltiples patrones (aplicar varias veces)
                        # Patrón 1: DE 2026 DE 2026 -> DE 2026
                        texto = re.sub(r'DE (\d{4}) DE \1', r'DE \1', texto)
                        # Patrón 2: DEL 2026 DEL 2026 -> DEL 2026
                        texto = re.sub(r'DEL (\d{4}) DEL \1', r'DEL \1', texto)
                        # Patrón 3: DE 2026 DEL 2026 -> DE 2026 (el más común)
                        texto = re.sub(r'DE (\d{4}) DEL \1', r'DE \1', texto)
                        # Patrón 4: DEL 2026 DE 2026 -> DEL 2026
                        texto = re.sub(r'DEL (\d{4}) DE \1', r'DEL \1', texto)
                        # Aplicar nuevamente para casos anidados
                        texto = re.sub(r'DE (\d{4}) DEL \1', r'DE \1', texto)
                        texto = re.sub(r'DEL (\d{4}) DE \1', r'DEL \1', texto)
                        texto = re.sub(r'\$\$+', '$', texto)
                        texto = re.sub(r'\$ \$+', '$', texto)
                        texto = re.sub(r'  +', ' ', texto)
                        
                        if texto != paragraph.text:
                            formato_original = None
                            if paragraph.runs:
                                primer_run = paragraph.runs[0]
                                formato_original = {
                                    'font_name': primer_run.font.name if primer_run.font.name else None,
                                    'font_size': primer_run.font.size if primer_run.font.size else None,
                                    'bold': primer_run.bold if primer_run.bold is not None else False,
                                    'italic': primer_run.italic if primer_run.italic is not None else False,
                                    'color': primer_run.font.color.rgb if primer_run.font.color and primer_run.font.color.rgb else None
                                }
                            
                            paragraph.clear()
                            nuevo_run = paragraph.add_run(texto)
                            
                            if formato_original:
                                if formato_original['font_name']:
                                    nuevo_run.font.name = formato_original['font_name']
                                if formato_original['font_size']:
                                    nuevo_run.font.size = formato_original['font_size']
                                if formato_original['color']:
                                    nuevo_run.font.color.rgb = formato_original['color']
                                nuevo_run.bold = formato_original['bold']
                                nuevo_run.italic = formato_original['italic']
        
        print("✅ Limpieza de duplicaciones completada")
        
        # Procesar tablas: eliminar fila de ADICIONALES si no hay turnos, agregar AUXILIO DE TRANSPORTE si está seleccionado
        for table in doc.tables:
            filas_a_eliminar = []
            indice_adicionales = -1
            indice_ultima_fila_datos = -1
            
            # Buscar la fila de ADICIONALES y la última fila de datos (antes del TOTAL)
            for idx, row in enumerate(table.rows):
                row_text = ' '.join([cell.text.strip() for cell in row.cells]).upper()
                
                # Buscar fila de ADICIONALES
                if 'ADICIONALES' in row_text and 'SUELDO FIJO' not in row_text and 'BONO' not in row_text:
                    indice_adicionales = idx
                
                # Buscar última fila de datos (antes del TOTAL)
                if 'TOTAL' not in row_text and 'SUELDO FIJO' in row_text or 'BONO' in row_text or 'ADICIONALES' in row_text:
                    indice_ultima_fila_datos = idx
            
            # Eliminar fila de ADICIONALES si no hay turnos
            if turnos_num == 0 and indice_adicionales >= 0:
                filas_a_eliminar.append(indice_adicionales)
                print(f"🗑️ Eliminando fila de ADICIONALES (índice {indice_adicionales}) - no hay turnos")
            
            # Eliminar filas en orden inverso para mantener los índices correctos
            for idx in sorted(filas_a_eliminar, reverse=True):
                if idx < len(table.rows):
                    table._element.remove(table.rows[idx]._element)
            
            # Agregar fila de AUXILIO DE TRANSPORTE si está seleccionado
            if tiene_auxilio_transporte and auxilio_transporte_num > 0:
                # Buscar la fila del TOTAL para insertar antes de ella
                indice_total = -1
                for idx, row in enumerate(table.rows):
                    row_text = ' '.join([cell.text.strip() for cell in row.cells]).upper()
                    if 'TOTAL' in row_text:
                        indice_total = idx
                        break
                
                # Si no se encuentra TOTAL, usar el final de la tabla
                if indice_total < 0:
                    indice_total = len(table.rows)
                
                # Crear nueva fila al final primero
                nueva_fila = table.add_row()
                
                # Llenar las celdas de la nueva fila
                if len(nueva_fila.cells) >= 4:
                    # Columna 1: Descripción
                    nueva_fila.cells[0].text = 'AUXILIO DE TRANSPORTE'
                    # Columna 2: Cantidad
                    nueva_fila.cells[1].text = 'MES COMPLETO'
                    # Columna 3: Valor
                    nueva_fila.cells[2].text = ax1_formateado
                    # Columna 4: Paciente (vacía)
                    nueva_fila.cells[3].text = ''
                
                # Mover la fila a la posición correcta (antes del TOTAL)
                if indice_total < len(table.rows) - 1:
                    nueva_fila_element = nueva_fila._element
                    tbl = table._element
                    # Remover de la posición actual
                    tbl.remove(nueva_fila_element)
                    # Insertar antes del TOTAL
                    fila_total = table.rows[indice_total]._element
                    fila_total.addprevious(nueva_fila_element)
                
                print(f"✅ Agregada fila de AUXILIO DE TRANSPORTE con valor {ax1_formateado}")
        
        # Guardar en memoria
        output = io.BytesIO()
        doc.save(output)
        output.seek(0)
        
        # Nombre del archivo
        nombre_archivo = nombre.replace(' ', '_') if nombre else 'Cuenta_Cobro'
        filename = f"Cuenta_Cobro_{nombre_archivo}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
        
        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            as_attachment=True,
            download_name=filename
        )
        
    except Exception as e:
        import traceback
        return jsonify({"error": str(e), "traceback": traceback.format_exc()}), 500

@app.route('/generate-contrato-arrendamiento', methods=['POST'])
def generate_contrato_arrendamiento():
    """Genera un contrato de arrendamiento de predio rural usando el template Word"""
    try:
        if not request.json:
            return jsonify({'error': 'No se recibieron datos'}), 400
        
        data = request.json
        
        # Validar y sanitizar datos del formulario
        nombre_arrendador = sanitize_input(data.get('nombreArrendador', ''), max_length=200)
        if not nombre_arrendador:
            return jsonify({'error': 'El nombre del arrendador es obligatorio'}), 400
        
        cedula_arrendador = sanitize_input(data.get('cedulaArrendador', ''), max_length=50)
        if not cedula_arrendador:
            return jsonify({'error': 'La cédula del arrendador es obligatoria'}), 400
        
        ciudad_expedicion_arrendador = sanitize_input(data.get('ciudadExpedicionArrendador', ''), max_length=100)
        nombre_arrendatario = sanitize_input(data.get('nombreArrendatario', ''), max_length=200)
        cedula_arrendatario = sanitize_input(data.get('cedulaArrendatario', ''), max_length=50)
        ciudad_expedicion_arrendatario = sanitize_input(data.get('ciudadExpedicionArrendatario', ''), max_length=100)
        nombre_predio = sanitize_input(data.get('nombrePredio', ''), max_length=200)
        nombre_vereda = sanitize_input(data.get('nombreVereda', ''), max_length=200)
        municipio = sanitize_input(data.get('municipio', ''), max_length=100)
        departamento = sanitize_input(data.get('departamento', ''), max_length=100)
        direccion_referencia = sanitize_input(data.get('direccionReferencia', ''), max_length=500)
        hectareas_arrendadas = sanitize_input(data.get('hectareasArrendadas', ''), max_length=50)
        hectareas_totales = sanitize_input(data.get('hectareasTotales', ''), max_length=50)
        valor_canon = sanitize_input(data.get('valorCanon', ''), max_length=50)
        duracion_contrato_anios = sanitize_input(data.get('duracionContratoAnios', ''), max_length=10)
        fecha_inicio_contrato = sanitize_input(data.get('fechaInicioContrato', ''), max_length=50)
        ciudad_firma_contrato = sanitize_input(data.get('ciudadFirmaContrato', ''), max_length=100)
        dia_firma = sanitize_input(data.get('diaFirma', ''), max_length=10)
        mes_firma = sanitize_input(data.get('mesFirma', ''), max_length=10)
        anio_firma = sanitize_input(data.get('anioFirma', ''), max_length=10)
        
        # Obtener hectáreas en texto si viene del formulario
        hectareas_arrendadas_texto = data.get('hectareasArrendadasTexto', '')
        if not hectareas_arrendadas_texto and hectareas_arrendadas:
            try:
                hectareas_num = float(hectareas_arrendadas.replace(',', '.'))
                # Convertir a texto simple
                hectareas_arrendadas_texto = str(hectareas_num)
            except:
                hectareas_arrendadas_texto = hectareas_arrendadas
        
        # Obtener nombre del mes
        mes_nombre = data.get('mesFirmaNombre', '')
        if not mes_nombre and mes_firma:
            try:
                mes_num = int(mes_firma)
                if 1 <= mes_num <= 12:
                    mes_nombre = MESES.get(mes_num, mes_firma)
            except:
                mes_nombre = mes_firma
        
        # Cargar template
        base_dir = os.path.dirname(__file__)
        template_path = os.path.join(base_dir, 'templates', 'contrato.docx')
        
        # Debug: Listar archivos en templates si no existe
        if not os.path.exists(template_path):
            templates_dir = os.path.join(base_dir, 'templates')
            available_files = []
            if os.path.exists(templates_dir):
                available_files = os.listdir(templates_dir)
            error_msg = f"Template no encontrado en: {template_path}\n"
            error_msg += f"Directorio base: {base_dir}\n"
            error_msg += f"Directorio templates: {templates_dir}\n"
            error_msg += f"Archivos disponibles en templates: {', '.join(available_files) if available_files else 'Ninguno'}"
            return jsonify({"error": error_msg}), 404
        
        doc = Document(template_path)
        
        # Preparar reemplazos con todas las variaciones posibles
        reemplazos = {}
        
        # Solo reemplazar variables con formato {{VARIABLE}} (llaves dobles y mayúsculas)
        # Arrendador
        reemplazos['{{NOMBRE_ARRENDADOR}}'] = nombre_arrendador.upper()
        reemplazos['{{CEDULA_ARRENDADOR}}'] = cedula_arrendador
        reemplazos['{{CIUDAD_EXPEDICION_ARRENDADOR}}'] = ciudad_expedicion_arrendador.upper()
        
        # Arrendatario
        reemplazos['{{NOMBRE_ARRENDATARIO}}'] = nombre_arrendatario.upper()
        reemplazos['{{CEDULA_ARRENDATARIO}}'] = cedula_arrendatario
        reemplazos['{{CIUDAD_EXPEDICION_ARRENDATARIO}}'] = ciudad_expedicion_arrendatario.upper()
        
        # Predio
        reemplazos['{{NOMBRE_PREDIO}}'] = nombre_predio.upper()
        reemplazos['{{NOMBRE_VEREDA}}'] = nombre_vereda.upper()
        reemplazos['{{MUNICIPIO}}'] = municipio.upper()
        reemplazos['{{DEPARTAMENTO}}'] = departamento.upper()
        reemplazos['{{DIRECCION_REFERENCIA}}'] = direccion_referencia.upper()
        
        # Hectáreas
        reemplazos['{{HECTAREAS_ARRENDADAS}}'] = hectareas_arrendadas
        reemplazos['{{HECTAREAS_ARRENDADAS_TEXTO}}'] = hectareas_arrendadas_texto.upper()
        reemplazos['{{HECTAREAS_TOTALES}}'] = hectareas_totales
        
        # Valor del canon
        reemplazos['{{VALOR_CANON}}'] = valor_canon
        
        # Duración y fecha inicio
        reemplazos['{{DURACION_CONTRATO_ANIOS}}'] = duracion_contrato_anios
        
        # Formatear fecha de inicio
        fecha_inicio_formateada = ''
        if fecha_inicio_contrato:
            try:
                fecha_obj = datetime.strptime(fecha_inicio_contrato, '%Y-%m-%d')
                fecha_inicio_formateada = formatear_fecha(fecha_inicio_contrato)
            except:
                fecha_inicio_formateada = fecha_inicio_contrato
        
        reemplazos['{{FECHA_INICIO_CONTRATO}}'] = fecha_inicio_formateada
        
        # Firma
        reemplazos['{{CIUDAD_FIRMA_CONTRATO}}'] = ciudad_firma_contrato.upper()
        reemplazos['{{DIA_FIRMA}}'] = dia_firma
        reemplazos['{{MES_FIRMA}}'] = mes_nombre.upper()
        reemplazos['{{ANIO_FIRMA}}'] = anio_firma
        
        # Reemplazar texto en el documento
        reemplazar_texto_en_documento(doc, reemplazos)
        
        # Limpiar duplicaciones después del reemplazo
        # Buscar y limpiar patrones comunes de duplicación
        import re
        for paragraph in doc.paragraphs:
            texto = paragraph.text
            # Limpiar duplicaciones específicas primero
            # "CONVENCIÓN de CONVENCIÓN" -> "CONVENCIÓN"
            texto = re.sub(r'\b(CONVENCIÓN)\s+de\s+\1\b', r'\1', texto, flags=re.IGNORECASE)
            # "NORTE DE SANTANDER de NORTE DE SANTANDER" -> "NORTE DE SANTANDER"
            texto = re.sub(r'\b(NORTE DE SANTANDER)\s+de\s+\1\b', r'\1', texto, flags=re.IGNORECASE)
            # Limpiar duplicaciones generales: "TEXTO de TEXTO" -> "TEXTO"
            # Aplicar múltiples veces para casos anidados
            for _ in range(3):  # Aplicar hasta 3 veces para casos complejos
                # Patrón que captura cualquier texto seguido de " de " y el mismo texto
                texto_anterior = texto
                # Mejorar el patrón para capturar mejor textos con espacios
                texto = re.sub(r'\b([A-ZÁÉÍÓÚÑ][A-ZÁÉÍÓÚÑ\s]{2,}?)\s+de\s+\1\b', r'\1', texto, flags=re.IGNORECASE)
                if texto == texto_anterior:
                    break  # No hay más cambios
            # Limpiar duplicaciones de año - múltiples patrones
            texto = re.sub(r'DE (\d{4}) DE \1', r'DE \1', texto)
            texto = re.sub(r'DEL (\d{4}) DEL \1', r'DEL \1', texto)
            texto = re.sub(r'DE (\d{4}) DEL \1', r'DE \1', texto)
            texto = re.sub(r'DEL (\d{4}) DE \1', r'DEL \1', texto)
            # Aplicar nuevamente para casos anidados
            texto = re.sub(r'DE (\d{4}) DEL \1', r'DE \1', texto)
            texto = re.sub(r'DEL (\d{4}) DE \1', r'DEL \1', texto)
            # Limpiar espacios múltiples
            texto = re.sub(r'  +', ' ', texto)
            
            if texto != paragraph.text:
                # Limpiar el párrafo y reconstruirlo
                paragraph.clear()
                paragraph.add_run(texto)
        
        # Guardar en memoria
        output = io.BytesIO()
        doc.save(output)
        output.seek(0)
        
        # Nombre del archivo
        nombre_archivo = nombre_arrendador.replace(' ', '_') if nombre_arrendador else 'Contrato_Arrendamiento'
        filename = f"Contrato_Arrendamiento_{nombre_archivo}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
        
        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            as_attachment=True,
            download_name=filename
        )
        
    except Exception as e:
        import traceback
        return jsonify({"error": str(e), "traceback": traceback.format_exc()}), 500

# Google Drive API Configuration
SCOPES = ['https://www.googleapis.com/auth/drive.file']
GOOGLE_DRIVE_FOLDER_ID = os.getenv('GOOGLE_DRIVE_FOLDER_ID', '')  # ID de la carpeta raíz en Google Drive

def get_google_drive_service():
    """Obtiene el servicio de Google Drive usando credenciales"""
    try:
        # Intentar cargar credenciales desde variable de entorno (token JSON)
        creds_json = os.getenv('GOOGLE_DRIVE_CREDENTIALS', None)
        if creds_json:
            import json
            creds_dict = json.loads(creds_json)
            creds = Credentials.from_authorized_user_info(creds_dict, SCOPES)
        else:
            # Intentar cargar desde archivo token.json
            token_path = os.path.join(os.path.dirname(__file__), 'token.json')
            if os.path.exists(token_path):
                creds = Credentials.from_authorized_user_file(token_path, SCOPES)
            else:
                # Si no hay credenciales, retornar None (se manejará en el endpoint)
                return None
        
        # Si las credenciales están expiradas, intentar refrescar
        if not creds.valid:
            if creds.expired and creds.refresh_token:
                creds.refresh(requests.Request())
        
        service = build('drive', 'v3', credentials=creds)
        return service
    except Exception as e:
        import traceback
        print(f'Error obteniendo servicio de Google Drive: {e}')
        print(traceback.format_exc())
        return None

def create_or_get_folder(service, folder_name, parent_folder_id=None):
    """Crea una carpeta en Google Drive o retorna su ID si ya existe"""
    try:
        # Buscar si la carpeta ya existe
        query = f"name='{folder_name}' and mimeType='application/vnd.google-apps.folder' and trashed=false"
        if parent_folder_id:
            query += f" and '{parent_folder_id}' in parents"
        else:
            query += " and 'root' in parents"
        
        results = service.files().list(q=query, spaces='drive', fields='files(id, name)').execute()
        items = results.get('files', [])
        
        if items:
            return items[0]['id']
        
        # Si no existe, crear la carpeta
        file_metadata = {
            'name': folder_name,
            'mimeType': 'application/vnd.google-apps.folder'
        }
        if parent_folder_id:
            file_metadata['parents'] = [parent_folder_id]
        
        folder = service.files().create(body=file_metadata, fields='id').execute()
        return folder.get('id')
    except Exception as e:
        print(f'Error creando/obteniendo carpeta: {e}')
        return None

def upload_file_to_drive(service, file_data, file_name, folder_id):
    """Sube un archivo a Google Drive en la carpeta especificada"""
    try:
        file_metadata = {
            'name': file_name,
            'parents': [folder_id]
        }
        
        # Convertir base64 a bytes si es necesario
        if isinstance(file_data, str):
            if file_data.startswith('data:'):
                # Es un data URL, extraer el base64 y el mimetype
                header, encoded = file_data.split(',', 1)
                file_bytes = base64.b64decode(encoded)
                # Intentar extraer mimetype del header si está disponible
                mimetype = 'application/pdf'  # Por defecto
                if 'mimetype=' in header or ':' in header:
                    # Buscar mimetype en el header (ej: data:application/pdf;base64)
                    if ':' in header:
                        mime_part = header.split(':')[1].split(';')[0]
                        if mime_part:
                            mimetype = mime_part
            else:
                # Es base64 directo
                file_bytes = base64.b64decode(file_data)
                mimetype = 'application/pdf'  # Por defecto
        else:
            file_bytes = file_data
            mimetype = 'application/pdf'  # Por defecto
        
        # Detectar mimetype por extensión si no se detectó
        if mimetype == 'application/pdf':
            file_ext = os.path.splitext(file_name)[1].lower()
            mimetype_map = {
                '.pdf': 'application/pdf',
                '.jpg': 'image/jpeg',
                '.jpeg': 'image/jpeg',
                '.png': 'image/png',
                '.doc': 'application/msword',
                '.docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                '.xls': 'application/vnd.ms-excel',
                '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            }
            mimetype = mimetype_map.get(file_ext, 'application/pdf')
        
        media = MediaIoBaseUpload(io.BytesIO(file_bytes), mimetype=mimetype, resumable=True)
        
        file = service.files().create(
            body=file_metadata,
            media_body=media,
            fields='id, webViewLink'
        ).execute()
        
        return {
            'file_id': file.get('id'),
            'web_link': file.get('webViewLink'),
            'name': file_name
        }
    except Exception as e:
        import traceback
        print(f'Error subiendo archivo {file_name} a Google Drive: {e}')
        print(f'Traceback: {traceback.format_exc()}')
        return None

@app.route('/upload-attachments-to-drive', methods=['POST'])
def upload_attachments_to_drive():
    """Sube anexos a Google Drive organizados por cliente"""
    if not GOOGLE_DRIVE_AVAILABLE:
        return jsonify({'error': 'Google Drive API no está disponible. Instala las dependencias: pip install google-api-python-client google-auth-httplib2 google-auth-oauthlib'}), 503
    
    try:
        if not request.json:
            return jsonify({'error': 'No se recibieron datos'}), 400
        
        data = request.json
        
        # Obtener datos del cliente
        client_name = data.get('clientName', '').strip()
        client_id = data.get('clientId', '').strip()  # Cédula o ID
        attachments = data.get('attachments', {})  # Diccionario de anexos
        
        if not client_name or not client_id:
            return jsonify({'error': 'Se requiere nombre del cliente e ID (cédula)'}), 400
        
        if not attachments:
            return jsonify({'error': 'No se proporcionaron anexos para subir'}), 400
        
        # Obtener servicio de Google Drive
        service = get_google_drive_service()
        if not service:
            return jsonify({'error': 'No se pudo conectar con Google Drive. Verifica las credenciales.'}), 500
        
        # Crear nombre de carpeta: nombre_cliente_cedula
        folder_name = f"{client_name}_{client_id}".replace(' ', '_').replace('/', '_')
        
        # Crear o obtener carpeta del cliente
        parent_folder_id = GOOGLE_DRIVE_FOLDER_ID if GOOGLE_DRIVE_FOLDER_ID else None
        client_folder_id = create_or_get_folder(service, folder_name, parent_folder_id)
        
        if not client_folder_id:
            return jsonify({'error': 'No se pudo crear la carpeta del cliente en Google Drive'}), 500
        
        # Mapeo de IDs de anexos a nombres descriptivos
        attachment_names = {
            'cedula': 'Cedula',
            'actaBachiller': 'Acta_Bachiller',
            'diplomaBachiller': 'Diploma_Bachiller',
            'actaOtro': 'Acta_Otro_Estudio',
            'diplomaOtro': 'Diploma_Otro_Estudio',
            'cursos': 'Cursos',
            'antecedentes': 'Antecedentes',
            'rut': 'RUT',
            'pension': 'Certificado_Pension',
            'eps': 'Certificado_EPS',
            'vacunas': 'Vacunas',
            'arl': 'Certificado_ARL',
            'otros': 'Otros_Documentos'
        }
        
        uploaded_files = []
        errors = []
        
        # Subir cada anexo
        for attachment_key, attachment_data in attachments.items():
            try:
                if not attachment_data or not attachment_data.get('dataUrl'):
                    print(f'⚠️ Saltando {attachment_key}: no tiene dataUrl')
                    continue
                
                # Obtener nombre descriptivo del anexo
                attachment_name = attachment_names.get(attachment_key, attachment_key)
                
                # Obtener extensión del archivo original
                original_name = attachment_data.get('name', 'documento')
                file_extension = os.path.splitext(original_name)[1] or '.pdf'
                
                # Crear nombre del archivo: nombre_anexo.extensión
                file_name = f"{attachment_name}{file_extension}"
                
                print(f'📤 Subiendo archivo: {file_name} (tipo: {attachment_key})')
                
                # Subir archivo
                result = upload_file_to_drive(
                    service,
                    attachment_data['dataUrl'],
                    file_name,
                    client_folder_id
                )
                
                if result:
                    print(f'✅ Archivo subido exitosamente: {file_name} (ID: {result["file_id"]})')
                    uploaded_files.append({
                        'key': attachment_key,
                        'name': file_name,
                        'file_id': result['file_id'],
                        'web_link': result.get('web_link', '')
                    })
                else:
                    error_msg = f"Error subiendo {attachment_key} ({file_name})"
                    print(f'❌ {error_msg}')
                    errors.append(error_msg)
            except Exception as e:
                import traceback
                error_msg = f"Error procesando {attachment_key}: {str(e)}"
                print(f'❌ {error_msg}')
                print(f'Traceback: {traceback.format_exc()}')
                errors.append(error_msg)
        
        # Enlace directo a la carpeta en Drive para que el admin pueda ver los archivos
        drive_folder_link = f'https://drive.google.com/drive/folders/{client_folder_id}'
        return jsonify({
            'success': True,
            'folder_name': folder_name,
            'folder_id': client_folder_id,
            'drive_folder_link': drive_folder_link,
            'uploaded_files': uploaded_files,
            'errors': errors,
            'message': f'Se subieron {len(uploaded_files)} archivo(s) a Google Drive'
        }), 200
        
    except Exception as e:
        import traceback
        return jsonify({
            'error': str(e),
            'traceback': traceback.format_exc()
        }), 500


# Mapeo de keys de anexos a prefijos de nombre en Drive (reutilizable)
ATTACHMENT_KEY_TO_DRIVE_PREFIX = {
    'cedula': 'Cedula',
    'actaBachiller': 'Acta_Bachiller',
    'diplomaBachiller': 'Diploma_Bachiller',
    'actaOtro': 'Acta_Otro_Estudio',
    'diplomaOtro': 'Diploma_Otro_Estudio',
    'cursos': 'Cursos',
    'antecedentes': 'Antecedentes',
    'rut': 'RUT',
    'pension': 'Certificado_Pension',
    'eps': 'Certificado_EPS',
    'vacunas': 'Vacunas',
    'arl': 'Certificado_ARL',
    'otros': 'Otros_Documentos',
    'contrato': 'Contrato',  # ARL
}


@app.route('/delete-attachment-from-drive', methods=['POST'])
def delete_attachment_from_drive():
    """Elimina un anexo de Google Drive por folder_id y attachment_key"""
    if not GOOGLE_DRIVE_AVAILABLE:
        return jsonify({'error': 'Google Drive API no está disponible'}), 503
    
    try:
        if not request.json:
            return jsonify({'error': 'No se recibieron datos'}), 400
        
        data = request.json
        folder_id = data.get('folderId', '').strip()
        attachment_key = data.get('attachmentKey', '').strip()
        
        if not folder_id or not attachment_key:
            return jsonify({'error': 'Se requiere folderId y attachmentKey'}), 400
        
        # Obtener prefijo del nombre del archivo en Drive
        file_prefix = ATTACHMENT_KEY_TO_DRIVE_PREFIX.get(attachment_key, attachment_key)
        
        service = get_google_drive_service()
        if not service:
            return jsonify({'error': 'No se pudo conectar con Google Drive'}), 500
        
        # Listar archivos en la carpeta
        results = service.files().list(
            q=f"'{folder_id}' in parents and trashed=false",
            fields="files(id, name)",
            pageSize=100
        ).execute()
        
        files = results.get('files', [])
        file_to_delete = None
        for f in files:
            name = f.get('name', '')
            if name.startswith(file_prefix) or file_prefix in name:
                file_to_delete = f
                break
        
        if not file_to_delete:
            return jsonify({
                'success': True,
                'message': 'Archivo no encontrado en Drive (puede que ya se eliminó)'
            }), 200
        
        # Eliminar archivo de Drive
        service.files().delete(fileId=file_to_delete['id']).execute()
        print(f"✅ Archivo eliminado de Drive: {file_to_delete.get('name')} (ID: {file_to_delete['id']})")
        
        return jsonify({
            'success': True,
            'message': f"Archivo '{file_to_delete.get('name')}' eliminado de Google Drive"
        }), 200
        
    except Exception as e:
        import traceback
        print(f"❌ Error eliminando anexo de Drive: {e}")
        print(traceback.format_exc())
        return jsonify({'error': str(e), 'traceback': traceback.format_exc()}), 500


@app.route('/test-upload-drive', methods=['GET', 'POST'])
def test_upload_drive():
    """
    Prueba de subida a Google Drive: genera un Word de prueba (cuenta de cobro)
    y lo sube a Drive. Sirve para verificar que credenciales y carpeta funcionan.
    """
    if not GOOGLE_DRIVE_AVAILABLE:
        return jsonify({
            'success': False,
            'error': 'Google Drive API no está disponible',
            'detail': 'Instala: pip install google-api-python-client google-auth-httplib2 google-auth-oauthlib'
        }), 503

    try:
        # 1. Obtener servicio
        service = get_google_drive_service()
        if not service:
            creds_env = 'GOOGLE_DRIVE_CREDENTIALS' if os.getenv('GOOGLE_DRIVE_CREDENTIALS') else None
            token_path = os.path.join(os.path.dirname(__file__), 'token.json')
            return jsonify({
                'success': False,
                'error': 'No se pudo conectar con Google Drive',
                'detail': 'Verifica: GOOGLE_DRIVE_CREDENTIALS (env) o token.json en la carpeta api. Ninguno encontrado o inválido.'
            }), 500

        # 2. Crear un Word de prueba (tipo cuenta de cobro)
        doc = Document()
        doc.add_paragraph('PRUEBA DE SUBIDA A GOOGLE DRIVE', style='Heading 1')
        doc.add_paragraph('')
        doc.add_paragraph('Documento generado para verificar que la subida a Drive funciona.')
        doc.add_paragraph('Fecha: ' + datetime.now().strftime('%Y-%m-%d %H:%M'))
        doc.add_paragraph('')
        doc.add_paragraph('— Cuenta de cobro de prueba —')
        buffer = io.BytesIO()
        doc.save(buffer)
        buffer.seek(0)
        file_bytes = buffer.read()

        # 3. Carpeta destino: la configurada o una de pruebas
        parent_folder_id = GOOGLE_DRIVE_FOLDER_ID.strip() if GOOGLE_DRIVE_FOLDER_ID else None
        folder_name = 'Pruebas_Upload'
        folder_id = create_or_get_folder(service, folder_name, parent_folder_id)

        if not folder_id:
            return jsonify({
                'success': False,
                'error': 'No se pudo crear o acceder a la carpeta en Drive',
                'detail': 'Revisa GOOGLE_DRIVE_FOLDER_ID o permisos de la cuenta.'
            }), 500

        # 4. Subir archivo
        file_name = f'Prueba_Cuenta_Cobro_{datetime.now().strftime("%Y%m%d_%H%M%S")}.docx'
        result = upload_file_to_drive(service, file_bytes, file_name, folder_id)

        if not result:
            return jsonify({
                'success': False,
                'error': 'La subida del archivo falló',
                'detail': 'Revisa los logs del servidor para el traceback.'
            }), 500

        return jsonify({
            'success': True,
            'message': 'Archivo de prueba subido correctamente a Google Drive',
            'file_name': result['name'],
            'file_id': result['file_id'],
            'web_view_link': result.get('web_link', ''),
            'folder_name': folder_name,
            'folder_id': folder_id
        }), 200

    except Exception as e:
        import traceback
        return jsonify({
            'success': False,
            'error': str(e),
            'traceback': traceback.format_exc()
        }), 500


# --- OAuth para cliente Web: obtener token desde el navegador (una sola vez) ---
@app.route('/get-drive-token', methods=['GET'])
def get_drive_token():
    """
    Inicia el flujo OAuth con tu cliente Web. Redirige a Google para que inicies sesión
    y autorices; luego Google redirige a /oauth2callback y se guarda el token en token.json.
    Usa tu Web client (ID y secreto) en GOOGLE_DRIVE_CLIENT_ID y GOOGLE_DRIVE_CLIENT_SECRET.
    En la consola de Google añade como URI de redirección: http://localhost:5000/oauth2callback
    (y si tienes API en producción: https://tu-dominio.com/oauth2callback).
    """
    if not GOOGLE_DRIVE_AVAILABLE or Flow is None:
        return jsonify({'error': 'Google Auth no disponible. Instala google-auth-oauthlib.'}), 503

    client_id = os.getenv('GOOGLE_DRIVE_CLIENT_ID', '').strip()
    client_secret = os.getenv('GOOGLE_DRIVE_CLIENT_SECRET', '').strip()
    if not client_id or not client_secret:
        return (
            '<p>Faltan variables de entorno. En el servidor configura:</p>'
            '<ul><li><b>GOOGLE_DRIVE_CLIENT_ID</b> = tu ID de cliente (Web)</li>'
            '<li><b>GOOGLE_DRIVE_CLIENT_SECRET</b> = tu secreto del cliente (Web)</li></ul>'
            '<p>En Google Cloud Console, en tu cliente Web, añade en "URIs de redirección":<br>'
            '<code>http://localhost:5000/oauth2callback</code></p>',
            400,
            {'Content-Type': 'text/html; charset=utf-8'}
        )

    # Redirect URI debe coincidir con lo configurado en la consola (cliente Web)
    use_https = request.headers.get('X-Forwarded-Proto') == 'https' or request.is_secure
    base_url = request.host_url.rstrip('/')
    redirect_uri = f'{base_url}/oauth2callback'

    client_config = {
        'web': {
            'client_id': client_id,
            'client_secret': client_secret,
            'redirect_uris': [redirect_uri],
            'auth_uri': 'https://accounts.google.com/o/oauth2/auth',
            'token_uri': 'https://oauth2.googleapis.com/token',
        }
    }
    try:
        flow = Flow.from_client_config(client_config, SCOPES, redirect_uri=redirect_uri)
        authorization_url, _ = flow.authorization_url(access_type='offline', prompt='consent')
        return redirect(authorization_url)
    except Exception as e:
        import traceback
        return f'<pre>Error iniciando OAuth: {e}\n{traceback.format_exc()}</pre>', 500


@app.route('/oauth2callback', methods=['GET'])
def oauth2callback():
    """
    Google redirige aquí tras autorizar. Intercambiamos el código por el token
    y lo guardamos en token.json (en la carpeta api).
    """
    if not GOOGLE_DRIVE_AVAILABLE or Flow is None:
        return '<p>Google Auth no disponible.</p>', 503

    code = request.args.get('code')
    if not code:
        error = request.args.get('error', 'Falta código')
        return f'<p>Error: {error}. No se recibió código de autorización.</p>', 400

    client_id = os.getenv('GOOGLE_DRIVE_CLIENT_ID', '').strip()
    client_secret = os.getenv('GOOGLE_DRIVE_CLIENT_SECRET', '').strip()
    if not client_id or not client_secret:
        return '<p>Faltan GOOGLE_DRIVE_CLIENT_ID o GOOGLE_DRIVE_CLIENT_SECRET.</p>', 500

    base_url = request.host_url.rstrip('/')
    redirect_uri = f'{base_url}/oauth2callback'
    client_config = {
        'web': {
            'client_id': client_id,
            'client_secret': client_secret,
            'redirect_uris': [redirect_uri],
            'auth_uri': 'https://accounts.google.com/o/oauth2/auth',
            'token_uri': 'https://oauth2.googleapis.com/token',
        }
    }
    try:
        flow = Flow.from_client_config(client_config, SCOPES, redirect_uri=redirect_uri)
        flow.fetch_token(code=code)
        creds = flow.credentials
        token_path = os.path.join(os.path.dirname(__file__), 'token.json')
        token_data = {
            'token': creds.token,
            'refresh_token': getattr(creds, 'refresh_token', None) or '',
            'token_uri': creds.token_uri,
            'client_id': creds.client_id,
            'client_secret': creds.client_secret,
            'scopes': list(creds.scopes) if creds.scopes else SCOPES,
        }
        with open(token_path, 'w', encoding='utf-8') as f:
            import json
            json.dump(token_data, f, indent=2)
        return (
            '<h2>Token guardado</h2>'
            '<p>El token de Google Drive se guardó en <code>token.json</code>.</p>'
            '<p>Ya puedes cerrar esta pestaña y probar la subida con <code>/test-upload-drive</code>.</p>',
            200,
            {'Content-Type': 'text/html; charset=utf-8'}
        )
    except Exception as e:
        import traceback
        return f'<pre>Error guardando token: {e}\n{traceback.format_exc()}</pre>', 500


if __name__ == '__main__':
    # Crear directorio de templates si no existe
    os.makedirs(os.path.join(os.path.dirname(__file__), 'templates'), exist_ok=True)
    app.run(debug=True, host='0.0.0.0', port=5000)
