#!/usr/bin/env python3
"""
Prueba de subida a la API upload-attachments.
Uso: python test_upload.py [URL_API]
Ejemplo: python test_upload.py
         python test_upload.py https://api-hv.onrender.com
"""
import json
import sys
import base64

try:
    import requests
except ImportError:
    print("Instala requests: pip install requests")
    sys.exit(1)

# PDF mínimo válido (~200 bytes)
MINI_PDF = b"""%PDF-1.4
1 0 obj<</Type/Catalog/Pages 2 0 R>>endobj
2 0 obj<</Type/Pages/Kids[3 0 R]/Count 1>>endobj
3 0 obj<</Type/Page/MediaBox[0 0 612 792]/Parent 2 0 R>>endobj
xref
0 4
0000000000 65535 f
0000000009 00000 n
0000000052 00000 n
0000000101 00000 n
trailer<</Size 4/Root 1 0 R>>
startxref
178
%%EOF"""

def main():
    api_url = (sys.argv[1] if len(sys.argv) > 1 else "https://api-hv.onrender.com").rstrip("/")
    data_url = "data:application/pdf;base64," + base64.b64encode(MINI_PDF).decode()
    payload = {
        "clientName": "Prueba Test",
        "clientId": "123456",
        "attachments": {
            "cedula": {"name": "test.pdf", "dataUrl": data_url}
        }
    }
    print(f"1. Comprobando estado de almacenamiento en {api_url}/storage-status ...")
    try:
        r = requests.get(f"{api_url}/storage-status", timeout=15)
        print(f"   Status: {r.status_code}")
        print(f"   Body: {json.dumps(r.json(), indent=2)}")
    except Exception as e:
        print(f"   Error: {e}")
    print()
    print(f"2. Enviando subida de prueba a {api_url}/upload-attachments ...")
    try:
        r = requests.post(
            f"{api_url}/upload-attachments",
            json=payload,
            headers={"Content-Type": "application/json"},
            timeout=60
        )
        print(f"   Status: {r.status_code}")
        try:
            data = r.json()
            print(f"   success: {data.get('success')}")
            print(f"   message: {data.get('message', '')}")
            if data.get('errors'):
                print(f"   errors: {data['errors']}")
            if data.get('uploaded_files'):
                print(f"   uploaded_files: {len(data['uploaded_files'])} archivo(s)")
            print(f"   Body completo: {json.dumps(data, indent=2, ensure_ascii=False)}")
        except Exception:
            print(f"   Body (texto): {r.text[:500]}")
    except requests.exceptions.Timeout:
        print("   Error: Timeout (la API en Render puede estar arrancando, espera 1 min y vuelve a intentar).")
    except Exception as e:
        print(f"   Error: {e}")

if __name__ == "__main__":
    main()
