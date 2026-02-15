# Configuración de Google Drive API

## Pasos para configurar la subida de anexos a Google Drive

### 1. Habilitar Google Drive API

1. Ve a [Google Cloud Console](https://console.cloud.google.com/)
2. Selecciona tu proyecto o crea uno nuevo
3. Ve a "APIs & Services" > "Library"
4. Busca "Google Drive API" y habilítala

---

## Opción A: Usar tu cliente Web (recomendado si ya tienes "Aplicación web")

Sirve para que todo funcione desde la web sin descargar archivos de escritorio.

### 2A. Configurar el cliente Web en Google Cloud

1. Ve a [Credenciales](https://console.cloud.google.com/apis/credentials).
2. En **"IDs de clientes de OAuth 2.0"**, entra a tu cliente **"Web client (auto created by Google Service)"** (o tu aplicación web).
3. En **"URIs de redireccionamiento autorizados"** haz clic en **"Agregar URI"** y añade exactamente:
   - Para probar en local: `http://localhost:5000/oauth2callback`
   - Si tu API está en un servidor: `https://tu-dominio-de-api.com/oauth2callback`
4. Guarda los cambios.

### 3A. Poner ID y secreto del cliente en el servidor

En las variables de entorno de tu API (Render, tu PC, etc.) configura:

- **`GOOGLE_DRIVE_CLIENT_ID`**: el ID de cliente (ej: `572579001678-....apps.googleusercontent.com`).
- **`GOOGLE_DRIVE_CLIENT_SECRET`**: el secreto del cliente.  
  Si en la consola dice "Ya no se pueden ver ni descargar los secretos", haz clic en **"Agregar secreto"**, crea uno nuevo y copia el valor (solo se muestra una vez).

### 4A. Obtener el token desde el navegador (una sola vez)

1. Arranca tu API (por ejemplo `python app.py` en la carpeta `api` o despliega en Render).
2. En el navegador abre: **`http://localhost:5000/get-drive-token`** (o `https://tu-api.com/get-drive-token`).
3. Te redirigirá a Google. Inicia sesión con la cuenta de Google donde quieras guardar los archivos en Drive y acepta los permisos.
4. Google te redirigirá a `/oauth2callback` y la API guardará el token en **`token.json`** en la carpeta `api`.
5. Verás una página que dice "Token guardado". A partir de ahí la subida a Drive debería funcionar (prueba con `/test-upload-drive`).

En producción (Render, etc.) puedes copiar el contenido de `token.json` y guardarlo en la variable de entorno **`GOOGLE_DRIVE_CREDENTIALS`** para no depender del archivo.

---

## Opción B: Cliente "Aplicación de escritorio"

### 2B. Crear credenciales OAuth 2.0 (Desktop)

1. Ve a "APIs & Services" > "Credentials"
2. Haz clic en "Create Credentials" > "OAuth client ID"
3. Si es la primera vez, configura la pantalla de consentimiento OAuth
4. Selecciona **"Aplicación de escritorio"** (Desktop app)
5. Descarga el archivo JSON de credenciales

### 3B. Obtener token de acceso (script local)

Ejecuta el siguiente script Python para obtener el token:

```python
from google_auth_oauthlib.flow import InstalledAppFlow
import json

SCOPES = ['https://www.googleapis.com/auth/drive.file']

CREDENTIALS_FILE = 'credentials.json'

flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
creds = flow.run_local_server(port=0)

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

### 4. Configurar variables de entorno

En tu servicio (Render, tu PC, etc.) agrega:

- **Opción A (Web):** `GOOGLE_DRIVE_CLIENT_ID`, `GOOGLE_DRIVE_CLIENT_SECRET` y, tras obtener el token, `GOOGLE_DRIVE_CREDENTIALS` (JSON del token) o deja que use `token.json`.
- **Opción B (Desktop):** `GOOGLE_DRIVE_CREDENTIALS`: el JSON del token obtenido en el paso 3B.
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

## Probar que la subida funciona

Puedes usar el endpoint de prueba para subir un Word de prueba a Drive:

- **GET o POST** `https://tu-api.com/test-upload-drive`

Si todo está bien configurado, la respuesta será algo como:

```json
{
  "success": true,
  "message": "Archivo de prueba subido correctamente a Google Drive",
  "file_name": "Prueba_Cuenta_Cobro_20250212_143022.docx",
  "file_id": "...",
  "web_view_link": "https://drive.google.com/file/d/...",
  "folder_name": "Pruebas_Upload",
  "folder_id": "..."
}
```

El archivo se crea en la carpeta **Pruebas_Upload** (dentro de la carpeta raíz si configuraste `GOOGLE_DRIVE_FOLDER_ID`, o en la raíz de Drive si no). Si `success` es `false`, el cuerpo de la respuesta incluye `error` y `detail` para depurar (credenciales, carpeta, etc.).

## Notas importantes

- Los archivos se suben automáticamente cuando el cliente guarda el formulario
- Si falla la subida a Google Drive, el guardado del formulario continúa normalmente (no se bloquea)
- Los errores se registran en la consola del navegador y en los logs del servidor
