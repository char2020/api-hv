from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
from docx import Document
from docx.shared import RGBColor, Pt, Inches
from docx.oxml import parse_xml
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH
import os
import re
import io
from datetime import datetime

app = Flask(__name__)
CORS(app)  # Permitir CORS para llamadas desde React

# Ruta a la plantilla Word
TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), 'templates', 'hv.docx')

@app.route('/', methods=['GET'])
def root():
    """Endpoint raíz para verificar que el servidor está funcionando"""
    return jsonify({
        "status": "ok", 
        "message": "API de Generación de Hojas de Vida funcionando",
        "endpoints": {
            "/health": "GET - Verificar estado del servidor",
            "/generate-word": "POST - Generar documento Word"
        }
    })

@app.route('/health', methods=['GET'])
def health():
    """Endpoint de salud para verificar que el servidor está funcionando"""
    return jsonify({"status": "ok", "message": "API funcionando correctamente"})

def reemplazar_en_parrafo(paragraph, reemplazos, es_encabezado=False, es_final=False):
    """Reemplaza variables en un párrafo preservando el formato"""
    texto_original = paragraph.text
    texto_nuevo = texto_original
    
    for variable, valor in reemplazos.items():
        if variable in texto_nuevo:
            texto_nuevo = texto_nuevo.replace(variable, valor)
    
    if texto_nuevo != texto_original:
        if paragraph.runs:
            # Preservar formato del primer run
            formato_original = None
            if paragraph.runs:
                primer_run = paragraph.runs[0]
                formato_original = {
                    'font_name': primer_run.font.name if primer_run.font.name else None,
                    'font_size': primer_run.font.size if primer_run.font.size else None,
                    'bold': primer_run.bold,
                    'italic': primer_run.italic,
                    'color': primer_run.font.color.rgb if primer_run.font.color and primer_run.font.color.rgb else None
                }
            paragraph.clear()
            nuevo_run = paragraph.add_run(texto_nuevo)
            if formato_original:
                if formato_original['font_name']:
                    nuevo_run.font.name = formato_original['font_name']
                if formato_original['font_size']:
                    nuevo_run.font.size = formato_original['font_size']
                if formato_original['color']:
                    nuevo_run.font.color.rgb = formato_original['color']
                nuevo_run.bold = formato_original['bold']
                nuevo_run.italic = formato_original['italic']
        else:
            paragraph.text = texto_nuevo

def reemplazar_en_encabezados(doc, reemplazos):
    """Reemplaza variables en encabezados y pies de página"""
    for section in doc.sections:
        # Encabezado
        header = section.header
        for paragraph in header.paragraphs:
            reemplazar_en_parrafo(paragraph, reemplazos, es_encabezado=True)
        # Pie de página
        footer = section.footer
        for paragraph in footer.paragraphs:
            reemplazar_en_parrafo(paragraph, reemplazos, es_encabezado=True)

@app.route('/generate-word', methods=['POST'])
def generate_word():
    """Genera un documento Word a partir de los datos recibidos"""
    try:
        data = request.json
        
        # Validar que existe la plantilla
        if not os.path.exists(TEMPLATE_PATH):
            return jsonify({"error": "Plantilla no encontrada"}), 404
        
        # Cargar la plantilla
        doc = Document(TEMPLATE_PATH)
        
        # Obtener datos básicos
        nombre = data.get('fullName', '').strip()
        direccion = data.get('address', '').strip()
        correo = data.get('email', '').strip()
        texto_perfil = data.get('profile', '').strip()
        
        # Obtener formaciones
        formaciones = data.get('formaciones', [])
        texto_formacion = ""
        if formaciones:
            primera_formacion = formaciones[0]
            tipo = primera_formacion.get('tipo', '')
            nombre_formacion = primera_formacion.get('nombre', '')
            texto_formacion = f"{tipo}: {nombre_formacion}" if tipo else nombre_formacion
        
        # Preparar reemplazos básicos (con variaciones de mayúsculas)
        reemplazos = {
            'nombre_01': nombre,
            'cedula': data.get('idNumber', '').strip(),
            'n_fecha': data.get('birthDate', '').strip(),
            'n_numer': data.get('phone', '').strip(),
            'dire_01': direccion,
            'Dire_01': direccion,
            'DIRE_01': direccion,
            'ciu_01': data.get('place', '').strip(),
            'est_01': data.get('estadoCivil', '').strip(),
            'exp_01': data.get('idIssuePlace', '').strip(),
            'exp_var': data.get('idIssuePlace', '').strip(),
            'perfil_01': texto_perfil,
            'Perfil_01': texto_perfil,
            'PERFIL_01': texto_perfil,
            'bac_01': data.get('highSchool', '').strip(),
            'cole_01': data.get('institution', '').strip(),
            'tec_01': texto_formacion,
        }
        
        # Agregar correo solo si tiene valor
        if correo:
            reemplazos['corr_01'] = correo
        
        # Agregar formaciones adicionales
        for i, form in enumerate(formaciones, start=1):
            tipo = form.get('tipo', '')
            nombre_formacion = form.get('nombre', '')
            texto = f"{tipo}: {nombre_formacion}" if tipo else nombre_formacion
            reemplazos[f'tec_{i:02d}'] = texto
            reemplazos[f'tec_0{i}'] = texto
            reemplazos[f'tec_{i}'] = texto
        
        # Obtener referencias
        referencias_familiares = data.get('referenciasFamiliares', [])
        referencias_personales = data.get('referenciasPersonales', [])
        
        # Agregar referencias familiares
        for i, ref in enumerate(referencias_familiares, start=1):
            nombre_ref = ref.get('nombre', '').strip()
            telefono_ref = ref.get('telefono', '').strip()
            reemplazos[f'Re_fam_{i:02d}'] = nombre_ref
            reemplazos[f'Re_fam_0{i}'] = nombre_ref
            reemplazos[f'Re_fam_{i}'] = nombre_ref
            reemplazos[f'Re_fam_1'] = nombre_ref if i == 1 else reemplazos.get('Re_fam_1', '')
            reemplazos[f'cel_f_{i:02d}'] = telefono_ref
            reemplazos[f'cel_f_0{i}'] = telefono_ref
            reemplazos[f'cel_f_{i}'] = telefono_ref
            reemplazos['cel_f_01'] = telefono_ref if i == 1 else reemplazos.get('cel_f_01', '')
            reemplazos['cel_f_1'] = telefono_ref if i == 1 else reemplazos.get('cel_f_1', '')
        
        # Agregar referencias personales
        for i, ref in enumerate(referencias_personales, start=1):
            nombre_ref = ref.get('nombre', '').strip()
            telefono_ref = ref.get('telefono', '').strip()
            reemplazos[f'Re_per_{i:02d}'] = nombre_ref
            reemplazos[f'Re_per_0{i}'] = nombre_ref
            reemplazos[f'Re_per_{i}'] = nombre_ref
            reemplazos[f'Re_per_1'] = nombre_ref if i == 1 else reemplazos.get('Re_per_1', '')
            reemplazos[f'cel_p_{i:02d}'] = telefono_ref
            reemplazos[f'cel_p_0{i}'] = telefono_ref
            reemplazos[f'cel_p_{i}'] = telefono_ref
            reemplazos['cel_p_01'] = telefono_ref if i == 1 else reemplazos.get('cel_p_01', '')
            reemplazos['cel_p_1'] = telefono_ref if i == 1 else reemplazos.get('cel_p_1', '')
        
        # Obtener experiencias laborales
        experiencias = data.get('experiencias', [])
        if experiencias:
            primera_exp = experiencias[0]
            reemplazos['local_01'] = primera_exp.get('empresa', '').strip()
            reemplazos['car_01'] = primera_exp.get('cargo', '').strip()
            fecha_inicio = primera_exp.get('fechaInicio', '').strip()
            fecha_fin = primera_exp.get('fechaFin', '').strip()
            tiempo = f"Desde {fecha_inicio} hasta {fecha_fin}" if fecha_inicio and fecha_fin else ""
            reemplazos['tiempo_01'] = tiempo
            
            # Agregar experiencias adicionales
            for i, exp in enumerate(experiencias[1:], start=2):
                reemplazos[f'local_{i:02d}'] = exp.get('empresa', '').strip()
                reemplazos[f'car_{i:02d}'] = exp.get('cargo', '').strip()
                fecha_inicio = exp.get('fechaInicio', '').strip()
                fecha_fin = exp.get('fechaFin', '').strip()
                tiempo = f"Desde {fecha_inicio} hasta {fecha_fin}" if fecha_inicio and fecha_fin else ""
                reemplazos[f'tiempo_{i:02d}'] = tiempo
        
        # Reemplazar en encabezados y pies de página primero
        reemplazar_en_encabezados(doc, reemplazos)
        
        # Reemplazar en párrafos del cuerpo
        total_paragraphs = len(doc.paragraphs)
        for idx, paragraph in enumerate(doc.paragraphs):
            es_final = (idx >= total_paragraphs - 3) and "nombre_01" in paragraph.text
            reemplazar_en_parrafo(paragraph, reemplazos, es_encabezado=False, es_final=es_final)
        
        # Reemplazar en tablas
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        es_final = "nombre_01" in paragraph.text
                        reemplazar_en_parrafo(paragraph, reemplazos, es_encabezado=False, es_final=es_final)
        
        # Reemplazo adicional específico para dire_01 (case-insensitive)
        if direccion:
            for paragraph in doc.paragraphs:
                texto = paragraph.text
                if "dire_01" in texto.lower() and direccion not in texto:
                    texto_nuevo = re.sub(r'dire_01', direccion, texto, flags=re.IGNORECASE)
                    if texto_nuevo != texto:
                        reemplazar_en_parrafo(paragraph, {'dire_01': direccion}, es_encabezado=False)
        
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

if __name__ == '__main__':
    # Crear directorio de templates si no existe
    os.makedirs(os.path.join(os.path.dirname(__file__), 'templates'), exist_ok=True)
    app.run(debug=True, host='0.0.0.0', port=5000)

