#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Script de prueba para el endpoint de cuenta de cobro
Prueba la generación del Word con datos de ejemplo
"""

import requests
import json
import os
import sys

# Configurar encoding para Windows
if sys.platform == 'win32':
    import codecs
    sys.stdout = codecs.getwriter('utf-8')(sys.stdout.buffer, 'strict')
    sys.stderr = codecs.getwriter('utf-8')(sys.stderr.buffer, 'strict')

# URL local
API_URL = "http://localhost:5000"

# Datos de prueba
test_data = {
    "nombre": "JUAN PEREZ",
    "cedula": "1234567890",
    "mes": "1",  # Enero
    "año": "2026",
    "sueldoFijo": "2000000",
    "diasTrabajados": "25",
    "bonoSeguridad": "200000",
    "turnosDescansos": "2",
    "paciente": "MARIA GARCIA",
    "cuentaBancaria": "1234567890123456"
}

def test_generate_cuenta_cobro():
    """Prueba el endpoint de generación de cuenta de cobro"""
    print("[TEST] Probando endpoint de cuenta de cobro...")
    print(f"[ENVIO] Enviando datos: {json.dumps(test_data, indent=2, ensure_ascii=False)}")
    
    try:
        response = requests.post(
            f"{API_URL}/generate-cuenta-cobro",
            json=test_data,
            timeout=30
        )
        
        print(f"\n[STATUS] Status Code: {response.status_code}")
        print(f"[HEADERS] Content-Type: {response.headers.get('content-type', 'N/A')}")
        
        if response.status_code == 200:
            # Guardar el archivo
            filename = "test_cuenta_cobro.docx"
            with open(filename, 'wb') as f:
                f.write(response.content)
            print(f"[OK] Archivo generado exitosamente: {filename}")
            print(f"[SIZE] Tamaño del archivo: {len(response.content)} bytes")
            
            # Verificar que es un archivo Word válido
            if len(response.content) > 100:
                print(f"[VERIFY] Archivo parece válido (tamaño > 100 bytes)")
            else:
                print(f"[WARNING] Archivo muy pequeño, puede estar corrupto")
            
            return True
        else:
            print(f"[ERROR] Status Code: {response.status_code}")
            try:
                error_data = response.json()
                print(f"[ERROR] Error JSON: {json.dumps(error_data, indent=2, ensure_ascii=False)}")
            except:
                print(f"[ERROR] Error Text: {response.text[:500]}")
            return False
            
    except requests.exceptions.ConnectionError:
        print(f"[ERROR] No se pudo conectar a {API_URL}")
        print("[INFO] Asegúrate de que el servidor esté corriendo: python app.py")
        return False
    except Exception as e:
        print(f"[ERROR] Error inesperado: {str(e)}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    print("=" * 60)
    print("PRUEBA LOCAL - GENERACIÓN DE CUENTA DE COBRO")
    print("=" * 60)
    print()
    
    success = test_generate_cuenta_cobro()
    
    print()
    print("=" * 60)
    if success:
        print("[RESULTADO] PRUEBA EXITOSA")
        print("[INFO] El archivo Word se generó correctamente")
        print("[INFO] Revisa el archivo: test_cuenta_cobro.docx")
    else:
        print("[RESULTADO] PRUEBA FALLIDA")
        print("[INFO] Revisa los errores arriba")
    print("=" * 60)

