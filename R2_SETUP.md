# Cloudflare R2 como almacenamiento de anexos

Alternativa a Google Drive y Firebase Storage. R2 es S3-compatible y sin cargo por egress.

## 1. Crear bucket en R2

1. Entra a [Cloudflare Dashboard](https://dash.cloudflare.com) → **R2 object storage**.
2. Clic en **Create bucket**.
3. Nombre del bucket (ej. `anexos-hv`) y crea.

## 2. Crear API Token para R2

1. En R2 → **Manage R2 API Tokens** (o Account Details → API Tokens → Manage).
2. **Create API token**.
3. Permisos: **Object Read & Write**.
4. Especifica el bucket o “All buckets”.
5. Crea y **copia** el **Access Key ID** y **Secret Access Key** (solo se muestran una vez).

## 3. Variables en Render (api-hv)

En **Environment** del servicio **api-hv** añade:

| Variable | Valor |
|----------|--------|
| `R2_S3_ENDPOINT` | `https://c25cbe3e895e4152aa8daba74e9dd51d.r2.cloudflarestorage.com` |
| `R2_ACCESS_KEY_ID` | *(tu Access Key ID del token)* |
| `R2_SECRET_ACCESS_KEY` | *(tu Secret Access Key del token)* |
| `R2_BUCKET_NAME` | *(nombre del bucket, ej. `anexos-hv`)* |

Opcional: `API_PUBLIC_URL` = `https://api-hv.onrender.com` (para URLs de descarga).

## 4. Orden de uso

La API prueba en este orden: **Google Drive** → **Cloudflare R2** → **Firebase Storage**.  
Si Drive no está configurado pero R2 sí, las subidas irán a R2. Los archivos se guardan bajo `anexos/{NombreCliente}_{Documento}/` y la descarga se hace por la API (`/drive-download`).

## Cuenta / endpoint que estás usando

- **Account ID:** `c25cbe3e895e4152aa8daba74e9dd51d`
- **S3 API endpoint:** `https://c25cbe3e895e4152aa8daba74e9dd51d.r2.cloudflarestorage.com`
