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
            pattern = re.escape(placeholder)
            matches = list(re.finditer(pattern, texto_nuevo, re.IGNORECASE))
            
            if matches:
                cambios_realizados = True
                # Reemplazar desde el final hacia el inicio para mantener índices
                for match in reversed(matches):
                    start, end = match.span()
                    texto_nuevo = texto_nuevo[:start] + str(valor) + texto_nuevo[end:]
        
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

@app.route('/generate-cuenta-cobro', methods=['POST'])
def generate_cuenta_cobro():
    """Genera una cuenta de cobro usando el template Word"""
    try:
        data = request.json
        
        # Obtener datos del formulario
        nombre = data.get('nombre', '').strip()
        cedula = data.get('cedula', '').strip()
        telefono = data.get('phone', '').strip() or data.get('telefono', '').strip() or data.get('phoneNumber', '').strip()
        print(f"📞 Teléfono recibido: '{telefono}'")  # Debug
        mes = data.get('mes', '').strip()
        año = data.get('año', '').strip()
        sueldo_fijo = data.get('sueldoFijo', '').strip()
        mes_completo = data.get('mesCompleto', True)
        dia_inicio = data.get('diaInicio', '1').strip()
        dia_fin = data.get('diaFin', '30').strip()
        dias_trabajados = data.get('diasTrabajados', '30').strip()
        
        # Si no es mes completo, calcular días trabajados desde día inicio hasta día fin
        if not mes_completo:
            try:
                dia_inicio_num = int(dia_inicio) if dia_inicio.isdigit() else 1
                dia_fin_num = int(dia_fin) if dia_fin.isdigit() else 30
                if dia_fin_num >= dia_inicio_num:
                    dias_trabajados = str((dia_fin_num - dia_inicio_num) + 1)
            except:
                pass
        bono_seguridad = data.get('bonoSeguridad', '').strip()
        turnos_descansos = data.get('turnosDescansos', '0').strip()
        paciente = data.get('paciente', '').strip()
        cuenta_bancaria = data.get('cuentaBancaria', '').strip()
        banco = data.get('banco', '').strip() or 'Bancolombia'
        
        # Calcular sueldo proporcional según días trabajados
        sueldo_proporcional = 0
        dias_num = int(dias_trabajados) if dias_trabajados.isdigit() else 30
        if dias_num < 1:
            dias_num = 30
        
        # Obtener el número de días del mes seleccionado
        try:
            mes_num = int(mes) if mes.isdigit() else 0
            año_num = int(año) if año.isdigit() else datetime.now().year
            if 1 <= mes_num <= 12:
                # Obtener el último día del mes (días del mes)
                from calendar import monthrange
                dias_del_mes = monthrange(año_num, mes_num)[1]  # monthrange devuelve (día de la semana, días del mes)
            else:
                dias_del_mes = 30  # Valor por defecto si el mes no es válido
        except:
            dias_del_mes = 30  # Valor por defecto en caso de error
        
        # Limitar días trabajados al máximo de días del mes
        if dias_num > dias_del_mes:
            dias_num = dias_del_mes
        
        try:
            if sueldo_fijo:
                sueldo_fijo_num = float(sueldo_fijo.replace('.', '').replace(',', '.'))
                # Calcular: (sueldo fijo / días del mes) * días trabajados
                # Primero dividir, luego multiplicar, y finalmente redondear
                valor_por_dia = sueldo_fijo_num / dias_del_mes
                sueldo_proporcional = valor_por_dia * dias_num
                # Redondear a número entero solo al final
                sueldo_proporcional = round(sueldo_proporcional)
        except:
            pass
        
        # Calcular adicionales (turnos * 60000)
        turnos_num = int(turnos_descansos) if turnos_descansos.isdigit() else 0
        valor_por_turno = 60000  # Cambiado a 60.000 por día
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
        # Buscar todas las variaciones posibles de las variables
        reemplazos = {}
        
        # Formatear bono seguridad (sin $ para evitar duplicaciones)
        bono_seguridad_formateado = ''
        bono_seguridad_sin_signo = ''
        if bono_seguridad:
            try:
                bono_num = float(bono_seguridad.replace('.', '').replace(',', '.'))
                bono_seguridad_formateado = formatear_monto(bono_num, incluir_signo=False)
                bono_seguridad_sin_signo = bono_seguridad_formateado
            except:
                bono_seguridad_formateado = bono_seguridad
                bono_seguridad_sin_signo = bono_seguridad
        
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
        
        # Paciente - múltiples variaciones
        paciente_valor = paciente.upper() if paciente else ''
        reemplazos['paciente1'] = paciente_valor
        reemplazos['PACIENTE1'] = paciente_valor
        reemplazos['{paciente1}'] = paciente_valor
        reemplazos['{{paciente1}}'] = paciente_valor
        reemplazos['[paciente1]'] = paciente_valor
        reemplazos['<<paciente1>>'] = paciente_valor
        
        # Sueldo proporcional - múltiples variaciones y valores de ejemplo
        # Reemplazar valores con $ y sin $ para evitar duplicaciones
        reemplazos['2.000.000'] = sueldo_proporcional_formateado
        reemplazos['$2.000.000'] = sueldo_proporcional_formateado  # Sin $ adicional
        reemplazos['$$2.000.000'] = sueldo_proporcional_formateado  # Limpiar múltiples $
        reemplazos['$$$2.000.000'] = sueldo_proporcional_formateado  # Limpiar múltiples $
        reemplazos['$ 2.000.000'] = sueldo_proporcional_formateado  # Sin $ adicional
        reemplazos['$ $ 2.000.000'] = sueldo_proporcional_formateado  # Limpiar múltiples $
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
        
        # Variable dia1 - día trabajado (del 1 al 30)
        reemplazos['dia1'] = str(dias_num)
        reemplazos['DIA1'] = str(dias_num)
        reemplazos['{dia1}'] = str(dias_num)
        reemplazos['{{dia1}}'] = str(dias_num)
        reemplazos['[dia1]'] = str(dias_num)
        reemplazos['<<dia1>>'] = str(dias_num)
        
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
            reemplazos['4 TURNOS'] = adicionales_texto_turnos
            reemplazos['TURNOS'] = adicionales_texto_turnos
            reemplazos['{turnos}'] = adicionales_texto_turnos
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
            reemplazos['ADICIONALES'] = adicionales_formateado
        else:
            # Si no hay turnos, dejar vacío
            reemplazos['4 4 TURNOS'] = ''  # Limpiar duplicación
            reemplazos['4 TURNOS'] = ''
            reemplazos['240.000'] = ''
            reemplazos['$240.000'] = ''
            reemplazos['$$240.000'] = ''
            reemplazos['$$$240.000'] = ''
            reemplazos['$ 240.000'] = ''
            reemplazos['ADICIONALES'] = ''
        
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
        
        # Log de reemplazos para debug
        print(f"🔍 Reemplazos a realizar: {len(reemplazos)} variables")
        for key, value in sorted(reemplazos.items()):
            if value:  # Solo mostrar los que tienen valor
                print(f"  - {key} -> {value}")
        
        # Reemplazar texto en el documento
        reemplazar_texto_en_documento(doc, reemplazos)
        print("✅ Reemplazos completados en el documento")
        
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
