# Configuración de Google Drive API

## ✅ Opción más fácil: Service Account (sin OAuth, sin token, sin navegador)

**Ventaja:** No hay que autorizar en el navegador ni renovar token. Un solo JSON y listo.

**Si ya tienes la cuenta de servicio de Firebase** (`firebase-adminsdk-...@....iam.gserviceaccount.com`): usa ese mismo JSON. Solo asegúrate de tener **Google Drive API** habilitada en el mismo proyecto (paso 1).

### Pasos

1. **Google Cloud** (proyecto `generador-hojas-vida`) → [APIs & Services → Library](https://console.cloud.google.com/apis/library) → busca **Google Drive API** → **Enable**. (Si ya está habilitada, no hace falta.)
2. **Credenciales:** Si no tienes Service Account, [Credentials](https://console.cloud.google.com/apis/credentials) → **+ Create Credentials** → **Service account** → crear → Keys → **Add key** → **Create new key** → **JSON**. Si usas la de Firebase, el JSON ya lo tienes.
3. **En Render** → servicio **api-hv** → **Environment** → **Add Environment Variable**:
   - **Key:** `GOOGLE_DRIVE_SERVICE_ACCOUNT_JSON`
   - **Value:** pega **todo** el JSON (desde `{` hasta `}`, en una línea o con saltos, pero completo).
4. Guarda. Render redesplegará.
5. Comprueba: abre **https://api-hv.onrender.com/drive-status**. Debe salir `"service_account_configured": true` y `"drive_ok": true`.

**⚠️ Para que los archivos aparezcan en TU correo (tu Drive) y no pidan "solicitar acceso":**

Por defecto los archivos van al Drive de la **cuenta de servicio** (un robot). Para que caigan en **tu** Drive:

1. Entra a [Google Drive](https://drive.google.com) con **tu correo** (el que quieres usar para ver los documentos).
2. Crea una carpeta, por ejemplo **"Documentos ARL"** o **"Subidas App"**.
3. Clic derecho en la carpeta → **Compartir**.
4. En "Añadir personas y grupos" pega el email de la **cuenta de servicio** (el que está en el JSON, ej. `firebase-adminsdk-fbsvc@generador-hojas-vida.iam.gserviceaccount.com`).
5. Dale permiso **Editor** → quita la casilla "Notificar a las personas" si no quieres → **Compartir**.
6. Abre la carpeta y mira la URL: `https://drive.google.com/drive/folders/XXXXXXXX` → el **ID** es la parte `XXXXXXXX`.
7. En **Render** → tu servicio **api-hv** → **Environment** → añade variable:
   - **Key:** `GOOGLE_DRIVE_FOLDER_ID`
   - **Value:** ese ID (solo el ID, sin la URL).
8. Guarda. Render redesplegará.

A partir de ahí, todas las subidas (carpetas por cliente) se crearán **dentro** de esa carpeta en tu Drive. No tendrás que "solicitar acceso": la carpeta es tuya y la app escribe en ella gracias al permiso que le diste a la cuenta de servicio.

---

## ⚠️ Si usas OAuth y ves "No hay credenciales" o "No se pudo conectar con Google Drive"

**Causa:** En Render falta la variable `GOOGLE_DRIVE_CREDENTIALS` (el token). Si antes funcionaba, el token pudo haber expirado o perderse en un redeploy.

**Solución rápida (5 min):**

1. **En Google Cloud** → [Credenciales](https://console.cloud.google.com/apis/credentials) → tu cliente Web → **URIs de redireccionamiento autorizados** → **+ Agregar URI** → añade **exactamente** (copia y pega):
   ```
   http://localhost:5000/oauth2callback
   ```
   *(Si falla, añade también: `http://127.0.0.1:5000/oauth2callback`)*
   **Guarda** los cambios (botón azul abajo).
2. **En tu PC:** Ejecuta `obtener-token-drive.bat` (en la carpeta api) — abre la API y el navegador automáticamente
3. Autoriza con tu cuenta de Google (donde quieres guardar los archivos)
4. Haz clic en **"Copiar JSON"** en la página que aparece
5. **En Render** → tu servicio → Environment → añade variable `GOOGLE_DRIVE_CREDENTIALS` y pega el JSON
6. Guarda. Render redesplegará automáticamente.

---

## Ubicación de la API en este proyecto

| Dato | Valor |
|------|-------|
| **URL de la API** | `https://api-hv.onrender.com` |
| **Repositorio** | `char2020/api-hv` (GitHub) |
| **Despliegue** | Render (render.yaml) |
| **OAuth redirect** | `https://api-hv.onrender.com/oauth2callback` |

---

## Pasos para configurar la subida de anexos a Google Drive

### 1. Habilitar Google Drive API

1. Ve a [Google Cloud Console](https://console.cloud.google.com/)
2. Selecciona tu proyecto **generador-hojas-vida**
3. Ve a "APIs & Services" > "Library"
4. Busca "Google Drive API" y habilítala

---

## Opción A: Usar tu cliente Web (recomendado si ya tienes "Aplicación web")

Sirve para que todo funcione desde la web sin descargar archivos de escritorio.

### 2A. Configurar el cliente Web en Google Cloud

1. Ve a [Credenciales](https://console.cloud.google.com/apis/credentials).
2. En **"IDs de clientes de OAuth 2.0"**, entra a tu cliente **"Web client (auto created by Google Service)"**.
3. En **"URIs de redireccionamiento autorizados"** haz clic en **"+ Agregar URI"** y añade:
   - **Producción (Render):** `https://api-hv.onrender.com/oauth2callback`
   - **Local (pruebas):** `http://localhost:5000/oauth2callback`
4. Guarda los cambios.

### 3A. Poner ID y secreto del cliente en Render

1. Entra a [Render Dashboard](https://dashboard.render.com/) → tu servicio **api-hv** (o hv-generator-api).
2. Ve a **Environment** y agrega:
   - **`GOOGLE_DRIVE_CLIENT_ID`**: tu ID de cliente (de Google Cloud → Credenciales)
   - **`GOOGLE_DRIVE_CLIENT_SECRET`**: el secreto (si no lo ves, crea uno nuevo con "Agregar secreto" en Google Cloud).
3. Guarda y espera a que Render redespliegue.

### 4A. Obtener el token desde el navegador (una sola vez)

1. En Render (producción): abre **https://api-hv.onrender.com/get-drive-token**  
   *(Puede tardar 1–2 min si el servidor está dormido)*
2. En local: abre **http://localhost:5000/get-drive-token**
3. Te redirigirá a Google. Inicia sesión con la cuenta donde quieras guardar los archivos en Drive y acepta los permisos.
4. Google te redirigirá a `/oauth2callback` y la API guardará el token.
5. En **Render**: el archivo `token.json` no persiste. Debes copiar el JSON que aparece en pantalla y guardarlo en la variable de entorno **`GOOGLE_DRIVE_CREDENTIALS`** en Render → Environment.
6. En **local**: el token se guarda en `token.json` en la carpeta `api`; prueba con `/test-upload-drive`.

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

- **GET o POST** `https://api-hv.onrender.com/test-upload-drive`

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
