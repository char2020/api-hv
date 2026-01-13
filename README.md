# API de Generación de Hojas de Vida

API Flask para generar documentos Word (.docx) desde datos JSON.

## Instalación

```bash
pip install -r requirements.txt
```

## Uso Local

```bash
python app.py
```

La API estará disponible en `http://localhost:5000`

## Endpoints

### GET /health
Verifica que el servidor está funcionando.

### POST /generate-word
Genera un documento Word a partir de datos JSON.

**Body (JSON):**
```json
{
  "fullName": "Juan Pérez",
  "idNumber": "1234567890",
  "birthDate": "01/01/1990",
  "phone": "3001234567",
  "address": "Calle 123 #45-67",
  "place": "Bogotá",
  "estadoCivil": "Soltero",
  "email": "juan@example.com",
  "idIssuePlace": "Bogotá",
  "profile": "Perfil profesional...",
  "technicalEducation": "Técnico en...",
  "highSchool": "Bachiller Académico",
  "institution": "Colegio Ejemplo",
  "referenciasFamiliares": [
    {"nombre": "María Pérez", "telefono": "3001111111"}
  ],
  "referenciasPersonales": [
    {"nombre": "Carlos López", "telefono": "3002222222"}
  ],
  "experiencias": [
    {
      "empresa": "Empresa XYZ",
      "cargo": "Desarrollador",
      "fechaInicio": "01/2020",
      "fechaFin": "12/2023"
    }
  ]
}
```

**Response:** Archivo Word (.docx) para descargar

## Despliegue en Render

1. Conecta tu repositorio a Render
2. Selecciona "Web Service"
3. Configura:
   - **Build Command:** `pip install -r requirements.txt`
   - **Start Command:** `python app.py`
   - **Environment:** Python 3
4. Asegúrate de que el archivo `templates/hv.docx` esté en el repositorio

## Notas

- La plantilla Word debe estar en `templates/hv.docx`
- Las variables en la plantilla se reemplazan automáticamente
- El formato original de la plantilla se preserva

