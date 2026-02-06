# Configuración de Google Drive API

## Pasos para configurar la subida de anexos a Google Drive

### 1. Habilitar Google Drive API

1. Ve a [Google Cloud Console](https://console.cloud.google.com/)
2. Selecciona tu proyecto o crea uno nuevo
3. Ve a "APIs & Services" > "Library"
4. Busca "Google Drive API" y habilítala

### 2. Crear credenciales OAuth 2.0

1. Ve a "APIs & Services" > "Credentials"
2. Haz clic en "Create Credentials" > "OAuth client ID"
3. Si es la primera vez, configura la pantalla de consentimiento OAuth
4. Selecciona "Desktop app" como tipo de aplicación
5. Descarga el archivo JSON de credenciales

### 3. Obtener token de acceso

Ejecuta el siguiente script Python para obtener el token:

```python
from google_auth_oauthlib.flow import InstalledAppFlow
import json

SCOPES = ['https://www.googleapis.com/auth/drive.file']

# Ruta al archivo de credenciales descargado
CREDENTIALS_FILE = 'credentials.json'

flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
creds = flow.run_local_server(port=0)

# Guardar token como JSON para usar en variable de entorno
token_data = {
    'token': creds.token,
    'refresh_token': creds.refresh_token,
    'token_uri': creds.token_uri,
    'client_id': creds.client_id,
    'client_secret': creds.client_secret,
    'scopes': creds.scopes
}

print(json.dumps(token_data))
```

### 4. Configurar variables de entorno en Render

En tu servicio de Render, agrega las siguientes variables de entorno:

- `GOOGLE_DRIVE_CREDENTIALS`: El JSON del token obtenido en el paso 3 (como string)
- `GOOGLE_DRIVE_FOLDER_ID`: (Opcional) ID de la carpeta raíz donde se crearán las carpetas de clientes. Si no se proporciona, se creará en la raíz de Drive.

### 5. Obtener ID de carpeta raíz (opcional)

Si quieres que todas las carpetas de clientes se creen dentro de una carpeta específica:

1. Crea una carpeta en Google Drive
2. Abre la carpeta y copia el ID de la URL:
   - URL ejemplo: `https://drive.google.com/drive/folders/1a2b3c4d5e6f7g8h9i0j`
   - El ID es: `1a2b3c4d5e6f7g8h9i0j`
3. Configura este ID en la variable de entorno `GOOGLE_DRIVE_FOLDER_ID`

## Estructura de carpetas en Google Drive

Cuando un cliente sube anexos, se crea la siguiente estructura:

```
[Carpeta Raíz (opcional)]
└── NombreCliente_Cedula/
    ├── Cedula.pdf
    ├── Acta_Bachiller.pdf
    ├── Diploma_Bachiller.pdf
    └── ... (otros anexos)
```

## Notas importantes

- Los archivos se suben automáticamente cuando el cliente guarda el formulario
- Si falla la subida a Google Drive, el guardado del formulario continúa normalmente (no se bloquea)
- Los errores se registran en la consola del navegador y en los logs del servidor
