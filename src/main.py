#!/usr/bin/env python3
"""
Ejemplo de uso de OpenAI Excel Helper
======================================
Este script demuestra las diferentes formas de usar el módulo openai_excel_helper
"""

from openai_excel_helper import (
    OpenAIExcelProcessor,
    simple_excel_query,
    extract_structured_data
)
from config import get_api_key


def ejemplo_consulta_simple():
    """
    Ejemplo 1: Consulta simple y rápida sobre un archivo Excel
    """
    print("\n" + "="*80)
    print("EJEMPLO 1: Consulta Simple")
    print("="*80)
    
    # Obtener API key desde .env o variable de entorno
    API_KEY = get_api_key()
    if not API_KEY:
        print("❌ Error: No se encontró API_KEY. Configúrala en .env o como variable de entorno.")
        return
    
    excel_path = "datos.xlsx"
    
    # Hacer una pregunta simple sobre el archivo
    result = simple_excel_query(
        api_key=API_KEY,
        excel_path=excel_path,
        prompt="¿Cuántas filas y columnas tiene este archivo Excel? Describe brevemente su contenido."
    )
    
    if result["success"]:
        print(f"\nRespuesta: {result['response']}")
        print(f"Tokens usados: {result['total_tokens']}")
    else:
        print(f"Error: {result['error']}")


def ejemplo_multiples_consultas():
    """
    Ejemplo 2: Múltiples consultas manteniendo el contexto
    """
    print("\n" + "="*80)
    print("EJEMPLO 2: Múltiples Consultas con Contexto")
    print("="*80)
    
    # Obtener API key desde .env o variable de entorno
    API_KEY = get_api_key()
    if not API_KEY:
        print("❌ Error: No se encontró API_KEY. Configúrala en .env o como variable de entorno.")
        return
    
    excel_path = "datos.xlsx"
    
    # Crear un procesador
    processor = OpenAIExcelProcessor(API_KEY, model="gpt-4o")
    
    # Subir el archivo (solo una vez)
    file_id = processor.upload_excel_file(excel_path)
    print(f"Archivo subido con ID: {file_id}")
    
    # Primera consulta
    result1 = processor.query_excel(
        "Dame un resumen general del contenido del archivo"
    )
    if result1["success"]:
        print(f"\n1. {result1['response']}\n")
    
    # Segunda consulta (mantiene el contexto)
    result2 = processor.query_excel(
        "¿Cuáles son las columnas más importantes?"
    )
    if result2["success"]:
        print(f"\n2. {result2['response']}\n")
    
    # Tercera consulta (sigue manteniendo el contexto)
    result3 = processor.query_excel(
        "¿Hay algún dato faltante o inconsistente?"
    )
    if result3["success"]:
        print(f"\n3. {result3['response']}\n")
    
    # Ver el historial completo
    print("\n" + "-"*80)
    print("Historial de conversación:")
    print("-"*80)
    for msg in processor.get_conversation_history():
        print(f"\n{msg['role'].upper()}: {msg['content'][:100]}...")


def ejemplo_extraccion_estructurada():
    """
    Ejemplo 3: Extracción de datos estructurados según un schema
    """
    print("\n" + "="*80)
    print("EJEMPLO 3: Extracción de Datos Estructurados")
    print("="*80)
    
    # Obtener API key desde .env o variable de entorno
    API_KEY = get_api_key()
    if not API_KEY:
        print("❌ Error: No se encontró API_KEY. Configúrala en .env o como variable de entorno.")
        return
    
    excel_path = "datos.xlsx"
    
    # Definir el schema de datos que queremos extraer
    schema = {
        "type": "object",
        "required": ["productos", "total_ventas", "fecha_reporte"],
        "properties": {
            "productos": {
                "type": "array",
                "items": {
                    "type": "object",
                    "properties": {
                        "nombre": {"type": "string"},
                        "cantidad": {"type": "integer"},
                        "precio": {"type": "number"}
                    }
                }
            },
            "total_ventas": {
                "type": "number",
                "description": "Suma total de todas las ventas"
            },
            "fecha_reporte": {
                "type": "string",
                "description": "Fecha del reporte en formato YYYY-MM-DD"
            }
        }
    }
    
    instructions = """
Analiza este archivo Excel de ventas y extrae:
- Lista de productos con nombre, cantidad y precio
- Total de ventas
- Fecha del reporte (si está disponible)

Asegúrate de que el JSON cumpla con el schema proporcionado.
"""
    
    result = extract_structured_data(
        api_key=API_KEY,
        excel_path=excel_path,
        schema=schema,
        instructions=instructions
    )
    
    if result["success"]:
        import json
        print("\nDatos extraídos:")
        print(json.dumps(result["data"], indent=2, ensure_ascii=False))
    else:
        print(f"Error: {result['error']}")


def ejemplo_uso_similar_distri_hub():
    """
    Ejemplo 4: Uso similar al de cli_radicacion.py de Distri-Hub
    Procesamiento con reintentos y validación
    """
    print("\n" + "="*80)
    print("EJEMPLO 4: Uso Similar a Distri-Hub (con reintentos)")
    print("="*80)
    
    # Obtener API key desde .env o variable de entorno
    API_KEY = get_api_key()
    if not API_KEY:
        print("❌ Error: No se encontró API_KEY. Configúrala en .env o como variable de entorno.")
        return
    
    excel_path = "facturas.xlsx"
    
    processor = OpenAIExcelProcessor(API_KEY)
    
    # Definir schema para validación (similar al médico)
    schema = {
        "type": "object",
        "required": ["facturas_validas", "facturas_invalidas"],
        "properties": {
            "facturas_validas": {
                "type": "array",
                "items": {
                    "type": "object",
                    "properties": {
                        "numero_factura": {"type": "string"},
                        "nit_cliente": {"type": "string"},
                        "valor": {"type": "number"}
                    }
                }
            },
            "facturas_invalidas": {
                "type": "array",
                "items": {"type": "string"}
            }
        }
    }
    
    instructions = """
Analiza el archivo de facturas y clasifícalas en:
- Facturas válidas: tienen número, NIT y valor
- Facturas inválidas: les falta algún dato

Devuelve SOLO el JSON según el schema.
"""
    
    current_try = 0
    exit_tries = 3
    data = None
    
    while exit_tries > current_try:
        try:
            print(f"\nIntento {current_try + 1}/{exit_tries}...")
            
            result = extract_structured_data(
                api_key=API_KEY,
                excel_path=excel_path,
                schema=schema,
                instructions=instructions
            )
            
            if result["success"]:
                data = result["data"]
                
                # Validar que tenga al menos una factura válida
                if len(data.get("facturas_validas", [])) > 0:
                    print("✓ Extracción exitosa")
                    break
                else:
                    print("⚠ No se encontraron facturas válidas, reintentando...")
                    current_try += 1
                    continue
            else:
                print(f"✗ Error: {result['error']}")
                current_try += 1
                continue
                
        except Exception as e:
            print(f"✗ Error en el intento: {str(e)}")
            current_try += 1
    
    if data:
        import json
        print("\n" + "-"*80)
        print("RESULTADO FINAL:")
        print("-"*80)
        print(json.dumps(data, indent=2, ensure_ascii=False))
        print(f"\nFacturas válidas: {len(data['facturas_validas'])}")
        print(f"Facturas inválidas: {len(data['facturas_invalidas'])}")
    else:
        print("\n✗ No se pudo procesar el archivo después de varios intentos")


if __name__ == "__main__":
    print("""
╔═══════════════════════════════════════════════════════════════════════════════╗
║                    OpenAI Excel Helper - Ejemplos de Uso                     ║
╚═══════════════════════════════════════════════════════════════════════════════╝

Este archivo contiene varios ejemplos de cómo usar el módulo openai_excel_helper.

NOTA: Para usar estos ejemplos, necesitas:
1. Una API Key de OpenAI
2. Un archivo Excel para procesar
3. Instalar las dependencias: pip install -r requirements.txt

Para uso en producción, usa el CLI:
    python cli_excel_openai.py datos.xlsx --api-key tu-key --interactive

Descomenta la función que quieras probar:
""")
    
    # Descomenta el ejemplo que quieras probar:
    # ejemplo_consulta_simple()
    # ejemplo_multiples_consultas()
    # ejemplo_extraccion_estructurada()
    # ejemplo_uso_similar_distri_hub()
    
    print("\nPara ver todos los ejemplos, edita main.py y descomenta las funciones.")
