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
CORS(app)  # Permitir CORS para llamadas desde React

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
        'public_key': 'project_public_e8de4c9dde8d3130930dc8f9620f9fd0_4gcUq34631a35630e89502c9cb2229d123ff4',
        'secret_key': 'secret_key_5f1ab1bb9dc866aadc8a05671e460491_zNqoaf28f8b33e1755f025940359d1d4a70a3'
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
    """
    def reemplazar_en_parrafo(paragraph, reemplazos_dict):
        """Reemplaza texto en un párrafo manteniendo formato"""
        texto_original = paragraph.text
        texto_nuevo = texto_original
        
        # Buscar y reemplazar (case-insensitive)
        for placeholder, valor in reemplazos_dict.items():
            # Buscar placeholder de forma case-insensitive
            import re
            pattern = re.escape(placeholder)
            matches = re.finditer(pattern, texto_original, re.IGNORECASE)
            for match in reversed(list(matches)):  # Reversed para mantener índices
                start, end = match.span()
                texto_nuevo = texto_nuevo[:start] + str(valor) + texto_nuevo[end:]
        
        if texto_nuevo != texto_original:
            # Guardar formato del primer run si existe
            formato_original = None
            if paragraph.runs:
                primer_run = paragraph.runs[0]
                formato_original = {
                    'font_name': primer_run.font.name,
                    'font_size': primer_run.font.size,
                    'bold': primer_run.bold,
                    'italic': primer_run.italic,
                    'color': primer_run.font.color.rgb if primer_run.font.color.rgb else None
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
    
    # Reemplazar en párrafos
    for paragraph in doc.paragraphs:
        reemplazar_en_parrafo(paragraph, reemplazos)
    
    # Reemplazar en tablas
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    reemplazar_en_parrafo(paragraph, reemplazos)

def formatear_monto(monto):
    """Formatea un monto como moneda colombiana"""
    if not monto:
        return ''
    try:
        monto_num = float(monto)
        return f"${monto_num:,.0f}".replace(',', '.')
    except:
        return str(monto)

@app.route('/generate-cuenta-cobro', methods=['POST'])
def generate_cuenta_cobro():
    """Genera una cuenta de cobro usando el template Word"""
    try:
        data = request.json
        
        # Obtener datos del formulario
        nombre = data.get('nombre', '').strip()
        cedula = data.get('cedula', '').strip()
        mes = data.get('mes', '').strip()
        año = data.get('año', '').strip()
        sueldo_fijo = data.get('sueldoFijo', '').strip()
        dias_trabajados = data.get('diasTrabajados', '30').strip()
        bono_seguridad = data.get('bonoSeguridad', '').strip()
        turnos_descansos = data.get('turnosDescansos', '0').strip()
        paciente = data.get('paciente', '').strip()
        cuenta_bancaria = data.get('cuentaBancaria', '').strip()
        
        # Calcular sueldo proporcional según días trabajados
        sueldo_proporcional = 0
        dias_num = int(dias_trabajados) if dias_trabajados.isdigit() else 30
        if dias_num > 30:
            dias_num = 30
        if dias_num < 1:
            dias_num = 30
        
        try:
            if sueldo_fijo:
                sueldo_fijo_num = float(sueldo_fijo.replace('.', '').replace(',', '.'))
                # Calcular: (sueldo fijo / 30) * días trabajados
                sueldo_proporcional = (sueldo_fijo_num / 30) * dias_num
        except:
            pass
        
        # Calcular adicionales (turnos * 70000)
        turnos_num = int(turnos_descansos) if turnos_descansos.isdigit() else 0
        valor_por_turno = 70000
        adicionales_valor = turnos_num * valor_por_turno
        
        # Calcular total
        total = 0
        try:
            # Agregar sueldo proporcional
            total += sueldo_proporcional
            
            # Agregar bono seguridad
            if bono_seguridad:
                total += float(bono_seguridad.replace('.', '').replace(',', '.'))
            
            # Agregar adicionales si hay turnos de descansos
            if turnos_num > 0:
                total += adicionales_valor
        except:
            pass
        
        # Formatear fecha (mes en texto)
        fecha_texto = ''
        if mes and año:
            mes_num = int(mes) if mes.isdigit() else 0
            if 1 <= mes_num <= 12:
                fecha_texto = f"{MESES[mes_num].upper()} DE {año}"
        
        # Cargar template
        template_path = os.path.join(os.path.dirname(__file__), 'templates', 'cobro_ 2026.docx')
        if not os.path.exists(template_path):
            return jsonify({"error": f"Template no encontrado en: {template_path}"}), 404
        
        doc = Document(template_path)
        
        # Preparar reemplazos usando los placeholders exactos del template
        reemplazos = {}
        
        # Name1 - Nombre
        reemplazos['Name1'] = nombre.upper()
        
        # banco1 - Cuenta bancaria
        reemplazos['banco1'] = cuenta_bancaria
        
        # Cedu1 - Cédula
        reemplazos['Cedu1'] = cedula
        
        # mes1 - Mes y año
        reemplazos['mes1'] = fecha_texto
        
        # valor1 - Total
        reemplazos['valor1'] = formatear_monto(total)
        
        # paciente1 - Nombre del paciente (si existe)
        if paciente:
            reemplazos['paciente1'] = paciente.upper()
        else:
            reemplazos['paciente1'] = ''  # Dejar vacío si no hay paciente
        
        # Reemplazar sueldo fijo mensual con sueldo proporcional en la tabla
        # El template tiene "SUELDO FIJO MENSUAL" y "MES COMPLETO" o el valor
        reemplazos['MES COMPLETO'] = f"{dias_num} DÍAS" if dias_num < 30 else 'MES COMPLETO'
        # También reemplazar el valor del sueldo en la tabla si existe
        reemplazos['2.000.000'] = formatear_monto(sueldo_proporcional)
        
        # Adicionales - Solo si hay turnos de descansos
        if turnos_num > 0:
            # Buscar y reemplazar la sección de adicionales
            # Formato: "ADICIONALES | X TURNOS | $ YYY.YYY"
            adicionales_texto_turnos = f"{turnos_num} TURNOS"
            adicionales_texto_valor = formatear_monto(adicionales_valor)
            
            # Buscar patrones comunes en el documento
            reemplazos['4 TURNOS'] = adicionales_texto_turnos
            reemplazos['240.000'] = adicionales_texto_valor
            reemplazos['$ 240.000'] = f"$ {adicionales_texto_valor}"
        else:
            # Si no hay turnos, eliminar la fila de adicionales
            # Esto se manejará removiendo el texto o dejándolo vacío
            reemplazos['ADICIONALES'] = ''
            reemplazos['4 TURNOS'] = ''
            reemplazos['240.000'] = ''
            reemplazos['$ 240.000'] = ''
        
        # Reemplazar texto en el documento
        reemplazar_texto_en_documento(doc, reemplazos)
        
        # Si no hay turnos de descansos, intentar eliminar la fila de adicionales de la tabla
        if turnos_num == 0:
            # Buscar tablas y eliminar filas que contengan "ADICIONALES"
            for table in doc.tables:
                rows_to_remove = []
                for i, row in enumerate(table.rows):
                    row_text = ' '.join([cell.text for cell in row.cells])
                    if 'ADICIONALES' in row_text.upper() and 'SUELDO FIJO' not in row_text.upper() and 'BONO' not in row_text.upper():
                        rows_to_remove.append(i)
                
                # Eliminar filas en orden inverso para mantener índices
                for i in reversed(rows_to_remove):
                    # No podemos eliminar directamente, pero podemos limpiar el contenido
                    for cell in table.rows[i].cells:
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
