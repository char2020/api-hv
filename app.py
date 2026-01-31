from flask import Flask, request, send_file, jsonify
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
        estado_civil = data.get('estadoCivil', '').strip()
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
        """Reemplaza texto en un párrafo manteniendo formato"""
        # Obtener todo el texto del párrafo
        texto_original = paragraph.text
        if not texto_original:
            return
        
        texto_nuevo = texto_original
        cambios_realizados = False
        
        # Buscar y reemplazar (case-insensitive) - ordenar por longitud descendente para evitar reemplazos parciales
        sorted_reemplazos = sorted(reemplazos_dict.items(), key=lambda x: len(x[0]), reverse=True)
        
        for placeholder, valor in sorted_reemplazos:
            if not placeholder:  # Saltar placeholders vacíos
                continue
            
            # Buscar placeholder de forma case-insensitive
            # Para dia1 y dia2, buscar sin word boundaries para capturar variaciones
            if 'dia' in placeholder.lower() and len(placeholder) <= 5:
                # Buscar dia1 o dia2 incluso si está en medio de texto (ej: "AL dia2")
                pattern = re.escape(placeholder)
            elif placeholder == '4 TURNOS':
                # NO reemplazar "4 TURNOS" - esto afectaría "4 TURNOS DOMICILIARIOS" en el texto descriptivo
                continue
            elif placeholder == 'ADICIONALES':
                # NO reemplazar "ADICIONALES" - debe mantenerse como texto descriptivo
                continue
            else:
                pattern = re.escape(placeholder)
            
            matches = list(re.finditer(pattern, texto_nuevo, re.IGNORECASE))
            
            if matches:
                cambios_realizados = True
                # Reemplazar desde el final hacia el inicio para mantener índices
                for match in reversed(matches):
                    start, end = match.span()
                    valor_reemplazo = str(valor)
                    
                    # Para dia1 y dia2, verificar si necesita espacios alrededor
                    if 'dia' in placeholder.lower() and len(placeholder) <= 5:
                        # Verificar si hay espacios alrededor en el texto original
                        tiene_espacio_antes = start > 0 and texto_nuevo[start-1].isspace()
                        tiene_espacio_despues = end < len(texto_nuevo) and texto_nuevo[end].isspace()
                        
                        # Si el valor ya tiene espacios pero el placeholder no los tenía, ajustar
                        if valor_reemplazo.startswith(' ') and tiene_espacio_antes:
                            valor_reemplazo = valor_reemplazo.lstrip()
                        if valor_reemplazo.endswith(' ') and tiene_espacio_despues:
                            valor_reemplazo = valor_reemplazo.rstrip()
                        
                        # Si no hay espacios alrededor y el valor tampoco los tiene, agregarlos
                        if not tiene_espacio_antes and not valor_reemplazo.startswith(' '):
                            # Verificar si el carácter anterior es una letra (necesita espacio)
                            if start > 0 and texto_nuevo[start-1].isalnum():
                                valor_reemplazo = ' ' + valor_reemplazo
                        if not tiene_espacio_despues and not valor_reemplazo.endswith(' '):
                            # Verificar si el carácter siguiente es una letra (necesita espacio)
                            if end < len(texto_nuevo) and texto_nuevo[end].isalnum():
                                valor_reemplazo = valor_reemplazo + ' '
                    
                    texto_nuevo = texto_nuevo[:start] + valor_reemplazo + texto_nuevo[end:]
                    # Log para debug de dia1 y dia2 (removido por seguridad)
                    # if 'dia' in placeholder.lower():
                    #     print(f"  ✅ Reemplazado '{placeholder}' por '{valor_reemplazo}' en: ...{texto_original[max(0, start-20):min(len(texto_original), start+20)]}...")
        
        if cambios_realizados and texto_nuevo != texto_original:
            # Guardar formato del primer run si existe
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
            
            # Limpiar y reemplazar
            paragraph.clear()
            nuevo_run = paragraph.add_run(texto_nuevo)
            
            # Restaurar formato si existe
            if formato_original:
                if formato_original['font_name']:
                    nuevo_run.font.name = formato_original['font_name']
                if formato_original['font_size']:
                    nuevo_run.font.size = formato_original['font_size']
                if formato_original['color']:
                    nuevo_run.font.color.rgb = formato_original['color']
                nuevo_run.bold = formato_original['bold']
                nuevo_run.italic = formato_original['italic']
    
    def reemplazar_en_runs(runs, reemplazos_dict):
        """Reemplaza texto en runs individuales - útil para variables divididas en múltiples runs"""
        texto_completo = ''.join([run.text for run in runs])
        if not texto_completo:
            return False
        
        texto_nuevo = texto_completo
        cambios_realizados = False
        
        # Ordenar por longitud descendente
        sorted_reemplazos = sorted(reemplazos_dict.items(), key=lambda x: len(x[0]), reverse=True)
        
        for placeholder, valor in sorted_reemplazos:
            if not placeholder:
                continue
            
            pattern = re.escape(placeholder)
            matches = list(re.finditer(pattern, texto_nuevo, re.IGNORECASE))
            
            if matches:
                cambios_realizados = True
                for match in reversed(matches):
                    start, end = match.span()
                    texto_nuevo = texto_nuevo[:start] + str(valor) + texto_nuevo[end:]
        
        if cambios_realizados and texto_nuevo != texto_completo:
            # Limpiar todos los runs y crear uno nuevo con el texto reemplazado
            for run in runs:
                run.text = ''
            if runs:
                runs[0].text = texto_nuevo
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
        bono_seguridad = str(validate_numeric(data.get('bonoSeguridad', '0'), min_val=0, max_val=10000000))
        turnos_descansos = str(int(validate_numeric(data.get('turnosDescansos', '0'), min_val=0, max_val=100, default=0)))
        paciente = sanitize_input(data.get('paciente', '') or data.get('patientName', ''), max_length=200)
        cuenta_bancaria = sanitize_input(data.get('cuentaBancaria', '') or data.get('bankAccount', ''), max_length=50)
        banco = sanitize_input(data.get('banco', ''), max_length=100).upper() or 'Bancolombia'
        tipo_cuenta_cobro = sanitize_input(data.get('tipoCuentaCobro', '12h'), max_length=10)
        if tipo_cuenta_cobro not in ['12h', '8h']:
            tipo_cuenta_cobro = '12h'
        tiene_auxilio_transporte = bool(data.get('tieneAuxilioTransporte', False))
        auxilio_transporte = str(validate_numeric(data.get('auxilioTransporte', '0'), min_val=0, max_val=10000000))
        
        # Parsear bono de seguridad
        bono_seguridad_num = 0
        try:
            if bono_seguridad:
                bono_limpio = str(bono_seguridad).replace('.', '').replace(',', '.')
                bono_seguridad_num = float(bono_limpio)
                if bono_seguridad_num < 0:
                    bono_seguridad_num = 0
        except (ValueError, TypeError):
            bono_seguridad_num = 0
        
        # Parsear auxilio de transporte
        auxilio_transporte_num = 0
        if tiene_auxilio_transporte and auxilio_transporte:
            try:
                auxilio_limpio = str(auxilio_transporte).replace('.', '').replace(',', '.')
                auxilio_transporte_num = float(auxilio_limpio)
                if auxilio_transporte_num < 0:
                    auxilio_transporte_num = 0
            except (ValueError, TypeError):
                auxilio_transporte_num = 0
        
        # Calcular adicionales (turnos * 60000)
        turnos_num = int(turnos_descansos) if turnos_descansos.isdigit() else 0
        valor_por_turno = 60000
        adicionales_valor = turnos_num * valor_por_turno
        
        # Calcular total
        total = sueldo_proporcional + bono_seguridad_num + adicionales_valor + auxilio_transporte_num
        
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
        
        # Formatear bono seguridad (sin $ para evitar duplicaciones)
        bono_seguridad_formateado = formatear_monto(bono_seguridad_num, incluir_signo=False) if bono_seguridad_num > 0 else ''
        
        # Formatear sueldo fijo mensual (no el proporcional) - este es el valor base
        sueldo_fijo_formateado = formatear_monto(sueldo_fijo_num, incluir_signo=False)
        
        # Formatear sueldo proporcional (sin $ para evitar duplicaciones)
        sueldo_proporcional_formateado = formatear_monto(sueldo_proporcional, incluir_signo=False)
        
        # Formatear adicionales (sin $ para evitar duplicaciones)
        adicionales_formateado = formatear_monto(adicionales_valor, incluir_signo=False) if turnos_num > 0 else ''
        
        # Formatear total (sin $ para evitar duplicaciones)
        total_formateado = formatear_monto(total, incluir_signo=False)
        
        # Nombre - múltiples variaciones
        reemplazos['Name1'] = nombre.upper()
        reemplazos['NAME1'] = nombre.upper()
        reemplazos['name1'] = nombre.upper()
        reemplazos['{Name1}'] = nombre.upper()
        reemplazos['{{Name1}}'] = nombre.upper()
        reemplazos['[Name1]'] = nombre.upper()
        reemplazos['<<Name1>>'] = nombre.upper()
        
        # Cédula - múltiples variaciones
        reemplazos['Cedu1'] = cedula
        reemplazos['CEDU1'] = cedula
        reemplazos['cedu1'] = cedula
        reemplazos['{Cedu1}'] = cedula
        reemplazos['{{Cedu1}}'] = cedula
        reemplazos['[Cedu1]'] = cedula
        reemplazos['<<Cedu1>>'] = cedula
        
        # Teléfono - múltiples variaciones (variable Num1 en el Word)
        # Asegurar que siempre se reemplace, incluso si está vacío
        telefono_valor = telefono if telefono else ''
        reemplazos['Num1'] = telefono_valor
        reemplazos['NUM1'] = telefono_valor
        reemplazos['num1'] = telefono_valor
        reemplazos['{Num1}'] = telefono_valor
        reemplazos['{{Num1}}'] = telefono_valor
        reemplazos['[Num1]'] = telefono_valor
        reemplazos['<<Num1>>'] = telefono_valor
        # También variaciones con espacios
        reemplazos['Num 1'] = telefono_valor
        reemplazos['NUM 1'] = telefono_valor
        reemplazos['{Num 1}'] = telefono_valor
        
        # Banco - múltiples variaciones (variable banco1 en el Word para NOMBRE del banco)
        reemplazos['banco1'] = banco
        reemplazos['BANCO1'] = banco
        reemplazos['{banco1}'] = banco
        reemplazos['{{banco1}}'] = banco
        reemplazos['[banco1]'] = banco
        reemplazos['<<banco1>>'] = banco
        
        # Número de cuenta bancaria - variable nbanco1 en el Word
        reemplazos['nbanco1'] = cuenta_bancaria
        reemplazos['NBANCO1'] = cuenta_bancaria
        reemplazos['NBanco1'] = cuenta_bancaria
        reemplazos['{nbanco1}'] = cuenta_bancaria
        reemplazos['{{nbanco1}}'] = cuenta_bancaria
        reemplazos['[nbanco1]'] = cuenta_bancaria
        reemplazos['<<nbanco1>>'] = cuenta_bancaria
        reemplazos['n banco1'] = cuenta_bancaria
        reemplazos['N BANCO1'] = cuenta_bancaria
        
        # Cuenta bancaria - múltiples variaciones (para mantener compatibilidad)
        reemplazos['cuenta1'] = cuenta_bancaria
        reemplazos['CUENTA1'] = cuenta_bancaria
        reemplazos['{cuenta1}'] = cuenta_bancaria
        
        # Mes y año - múltiples variaciones
        reemplazos['mes1'] = fecha_texto
        reemplazos['MES1'] = fecha_texto
        reemplazos['{mes1}'] = fecha_texto
        reemplazos['{{mes1}}'] = fecha_texto
        reemplazos['[mes1]'] = fecha_texto
        reemplazos['<<mes1>>'] = fecha_texto
        
        # Total/Valor - múltiples variaciones (sin $ adicional para evitar duplicaciones)
        reemplazos['valor1'] = total_formateado
        reemplazos['VALOR1'] = total_formateado
        reemplazos['{valor1}'] = total_formateado
        reemplazos['{{valor1}}'] = total_formateado
        reemplazos['[valor1]'] = total_formateado
        reemplazos['<<valor1>>'] = total_formateado
        reemplazos['total1'] = total_formateado
        reemplazos['TOTAL1'] = total_formateado
        reemplazos['{total1}'] = total_formateado
        # Limpiar duplicaciones de $ en el total
        reemplazos['$ $2.440.000'] = total_formateado  # Limpiar múltiples $
        reemplazos['$$2.440.000'] = total_formateado  # Limpiar múltiples $
        reemplazos['$2.440.000'] = total_formateado  # Sin $ adicional
        reemplazos['2.440.000'] = total_formateado
        # Reemplazar "23.000.000" que aparece incorrectamente en el template
        reemplazos['23.000.000'] = total_formateado
        reemplazos['$23.000.000'] = total_formateado  # Sin $ adicional
        reemplazos['$$23.000.000'] = total_formateado  # Limpiar múltiples $
        reemplazos['$ 23.000.000'] = total_formateado  # Sin $ adicional
        
        # Paciente - múltiples variaciones
        paciente_valor = paciente.upper() if paciente else ''
        reemplazos['paciente1'] = paciente_valor
        reemplazos['PACIENTE1'] = paciente_valor
        reemplazos['{paciente1}'] = paciente_valor
        reemplazos['{{paciente1}}'] = paciente_valor
        reemplazos['[paciente1]'] = paciente_valor
        reemplazos['<<paciente1>>'] = paciente_valor
        
        # Sueldo fijo mensual - múltiples variaciones (este es el valor base, no el proporcional)
        reemplazos['sueldoFijoMensual'] = sueldo_fijo_formateado
        reemplazos['SUELDO FIJO MENSUAL'] = sueldo_fijo_formateado
        reemplazos['sueldoFijo'] = sueldo_fijo_formateado
        reemplazos['sueldoFijo1'] = sueldo_fijo_formateado
        reemplazos['SUELDOFIJO1'] = sueldo_fijo_formateado
        reemplazos['{sueldoFijo}'] = sueldo_fijo_formateado
        reemplazos['{{sueldoFijo}}'] = sueldo_fijo_formateado
        
        # Sueldo proporcional - múltiples variaciones y valores de ejemplo
        # Reemplazar valores con $ y sin $ para evitar duplicaciones (solo para el proporcional)
        reemplazos['sueldo1'] = sueldo_proporcional_formateado
        reemplazos['SUELDO1'] = sueldo_proporcional_formateado
        reemplazos['{sueldo1}'] = sueldo_proporcional_formateado
        reemplazos['sueldoProporcional'] = sueldo_proporcional_formateado
        reemplazos['SUELDO PROPORCIONAL'] = sueldo_proporcional_formateado
        
        # Días trabajados
        dias_texto = f"{dias_num} DÍAS" if dias_num < 30 else 'MES COMPLETO'
        reemplazos['MES COMPLETO'] = dias_texto
        reemplazos['30 DÍAS'] = dias_texto
        reemplazos['30 DIAS'] = dias_texto
        reemplazos['dias1'] = str(dias_num)
        reemplazos['DIAS1'] = str(dias_num)
        reemplazos['{dias1}'] = str(dias_num)
        reemplazos['diasTrabajados'] = str(dias_num)
        
        # Variable dia1 y dia2 - dia1 es el día de inicio, dia2 es el día final
        # Esto aplica para ambos tipos (12h y 8h)
        # Obtener diaInicio y diaFin de los datos (ya se obtuvieron antes, pero asegurarse de tenerlos)
        dia_inicio = data.get('diaInicio', '1').strip() if data.get('diaInicio') else '1'
        dia_fin = data.get('diaFin', '30').strip() if data.get('diaFin') else str(dias_num)
        
        # Debug: verificar valores
        # Debug removido por seguridad - solo en desarrollo
        # print(f"🔍 DEBUG dia1/dia2: diaInicio='{dia_inicio}', diaFin='{dia_fin}'")
        # print(f"🔍 DEBUG data keys: {list(data.keys())}")
        # if 'diaInicio' in data:
        #     print(f"🔍 DEBUG diaInicio raw: '{data.get('diaInicio')}'")
        # if 'diaFin' in data:
        #     print(f"🔍 DEBUG diaFin raw: '{data.get('diaFin')}'")
        
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
        
        # dia2 siempre es el día final - múltiples variaciones para asegurar el reemplazo
        # Agregar espacios alrededor para evitar que quede pegado
        reemplazos['dia2'] = f' {dia_fin} '  # Agregar espacios alrededor
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
        
        # Bono seguridad - múltiples variaciones (sin $ adicional para evitar duplicaciones)
        # Solo reemplazar variables específicas, NO valores fijos del template como "200.000" o "BONO SEGURIDAD"
        if bono_seguridad_formateado:
            reemplazos['bono1'] = bono_seguridad_formateado
            reemplazos['BONO1'] = bono_seguridad_formateado
            reemplazos['{bono1}'] = bono_seguridad_formateado
            reemplazos['{{bono1}}'] = bono_seguridad_formateado
            reemplazos['[bono1]'] = bono_seguridad_formateado
            reemplazos['<<bono1>>'] = bono_seguridad_formateado
            reemplazos['bonoSeguridad'] = bono_seguridad_formateado
            # NO reemplazar "BONO SEGURIDAD" ni "200.000" - son valores fijos del template
        
        # Adicionales - Solo si hay turnos de descansos
        if turnos_num > 0:
            adicionales_texto_turnos = f"{turnos_num} TURNOS"
            # Limpiar duplicaciones de "4 TURNOS" -> solo el número
            reemplazos['4 4 TURNOS'] = adicionales_texto_turnos  # Limpiar duplicación
            # NO reemplazar "4 TURNOS" directamente - esto afectaría "4 TURNOS DOMICILIARIOS" en el texto descriptivo
            # NO reemplazar "TURNOS" solo - esto afectaría el texto descriptivo "TURNOS DOMICILIARIOS"
            # Solo reemplazar variaciones específicas con placeholders
            reemplazos['{turnos}'] = adicionales_texto_turnos
            reemplazos['{{turnos}}'] = adicionales_texto_turnos
            reemplazos['[turnos]'] = adicionales_texto_turnos
            reemplazos['<<turnos>>'] = adicionales_texto_turnos
            # Valores sin $ adicional para evitar duplicaciones
            reemplazos['240.000'] = adicionales_formateado
            reemplazos['$240.000'] = adicionales_formateado  # Sin $ adicional
            reemplazos['$$240.000'] = adicionales_formateado  # Limpiar múltiples $
            reemplazos['$$$240.000'] = adicionales_formateado  # Limpiar múltiples $
            reemplazos['$ 240.000'] = adicionales_formateado  # Sin $ adicional
            reemplazos['$ $ 240.000'] = adicionales_formateado  # Limpiar múltiples $
            reemplazos['adicionales1'] = adicionales_formateado
            reemplazos['ADICIONALES1'] = adicionales_formateado
            reemplazos['{adicionales1}'] = adicionales_formateado
            # NO reemplazar "ADICIONALES" - debe mantenerse como texto descriptivo en la tabla
            # NO reemplazar "DESCANSOS" - es el texto de descripción que debe mantenerse
            # Solo reemplazar el valor monetario, no el texto descriptivo
        else:
            # Si no hay turnos, dejar vacío los valores pero NO eliminar "DESCANSOS", "TURNOS" ni "ADICIONALES"
            reemplazos['4 4 TURNOS'] = ''  # Limpiar duplicación
            # NO reemplazar "4 TURNOS" - esto afectaría "4 TURNOS DOMICILIARIOS"
            reemplazos['240.000'] = ''
            reemplazos['$240.000'] = ''
            reemplazos['$$240.000'] = ''
            reemplazos['$$$240.000'] = ''
            reemplazos['$ 240.000'] = ''
            # NO reemplazar "ADICIONALES" - debe mantenerse como texto descriptivo
            # NO reemplazar "DESCANSOS" ni "TURNOS" - deben mantenerse como texto descriptivo
            # Solo limpiar placeholders específicos
            reemplazos['{turnos}'] = ''
            reemplazos['{{turnos}}'] = ''
            reemplazos['[turnos]'] = ''
            reemplazos['<<turnos>>'] = ''
        
        # Auxilio de transporte
        if tiene_auxilio_transporte and auxilio_transporte_num > 0:
            auxilio_formateado = formatear_monto(auxilio_transporte_num, incluir_signo=False)
            reemplazos['auxilioTransporte'] = auxilio_formateado
            reemplazos['AUXILIO TRANSPORTE'] = auxilio_formateado
            reemplazos['auxilio1'] = auxilio_formateado
            reemplazos['AUXILIO1'] = auxilio_formateado
            reemplazos['{auxilio1}'] = auxilio_formateado
            reemplazos['{{auxilio1}}'] = auxilio_formateado
            reemplazos['[auxilio1]'] = auxilio_formateado
            reemplazos['<<auxilio1>>'] = auxilio_formateado
        else:
            reemplazos['auxilioTransporte'] = ''
            reemplazos['AUXILIO TRANSPORTE'] = ''
            reemplazos['auxilio1'] = ''
            reemplazos['AUXILIO1'] = ''
        
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
        
        # Si no hay turnos de descansos, limpiar solo los valores pero mantener "DESCANSOS" como texto
        if turnos_num == 0:
            # Buscar tablas y limpiar solo los valores (columna VALOR) pero mantener "DESCANSOS" en la descripción
            for table in doc.tables:
                for row in table.rows:
                    row_text = ' '.join([cell.text for cell in row.cells])
                    # Si la fila contiene "DESCANSOS" pero no tiene valores de sueldo o bono
                    if 'DESCANSOS' in row_text.upper() and 'SUELDO FIJO' not in row_text.upper() and 'BONO' not in row_text.upper():
                        # Limpiar solo las celdas de cantidad y valor, pero mantener "DESCANSOS" en la primera columna
                        for idx, cell in enumerate(row.cells):
                            # Si no es la primera columna (donde está "DESCANSOS"), limpiar
                            if idx > 0:
                                for paragraph in cell.paragraphs:
                                    paragraph.clear()
                                    paragraph.add_run('')
        
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

if __name__ == '__main__':
    # Crear directorio de templates si no existe
    os.makedirs(os.path.join(os.path.dirname(__file__), 'templates'), exist_ok=True)
    app.run(debug=True, host='0.0.0.0', port=5000)
