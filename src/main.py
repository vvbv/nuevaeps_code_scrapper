#!/usr/bin/env python3
"""
CLI para procesar archivos Excel con OpenAI
===========================================
"""

import argparse
import os
import sys
from openai_excel_helper import OpenAIExcelProcessor, simple_excel_query, extract_structured_data
from config import get_api_key


def interactive_mode(processor: OpenAIExcelProcessor):
    """
    Modo interactivo para hacer múltiples consultas sobre el archivo.
    """
    print("\n" + "="*80)
    print("MODO INTERACTIVO - Consultas sobre el archivo Excel")
    print("="*80)
    print("\nComandos especiales:")
    print("  - 'salir' o 'exit': Terminar la sesión")
    print("  - 'reset': Reiniciar el historial de conversación")
    print("  - 'historial': Ver el historial de conversación")
    print("\nEscribe tu consulta y presiona Enter:\n")
    
    while True:
        try:
            query = input(">>> ").strip()
            
            if not query:
                continue
            
            if query.lower() in ['salir', 'exit', 'quit']:
                print("\n¡Hasta luego!")
                break
            
            if query.lower() == 'reset':
                processor.reset_conversation()
                continue
            
            if query.lower() == 'historial':
                history = processor.get_conversation_history()
                print("\n" + "-"*80)
                print("HISTORIAL DE CONVERSACIÓN:")
                print("-"*80)
                for i, msg in enumerate(history, 1):
                    role = "TÚ" if msg["role"] == "user" else "ASISTENTE"
                    print(f"\n[{i}] {role}:")
                    print(msg["content"])
                print("-"*80 + "\n")
                continue
            
            # Hacer la consulta
            result = processor.query_excel(query)
            
            if result["success"]:
                print(f"\n{result['response']}")
                print(f"\n[Tokens usados: {result['total_tokens']}]")
            else:
                print(f"\n❌ Error: {result['error']}")
            
            print()  # Línea en blanco
            
        except KeyboardInterrupt:
            print("\n\n¡Hasta luego!")
            break
        except Exception as e:
            print(f"\n❌ Error: {e}\n")


def single_query_mode(excel_path: str, api_key: str, query: str, model: str):
    """
    Modo de consulta única.
    """
    print(f"\nConsultando: {query}")
    print("-" * 80)
    
    result = simple_excel_query(api_key, excel_path, query, model=model)
    
    if result["success"]:
        print(f"\n{result['response']}")
        print(f"\n[Tokens usados: {result['total_tokens']}]")
    else:
        print(f"\n❌ Error: {result['error']}")


def structured_extraction_mode(excel_path: str, api_key: str, model: str):
    """
    Modo de extracción estructurada - ejemplo con schema predefinido.
    """
    print("\nExtrayendo datos estructurados del archivo...")
    print("-" * 80)
    
    # Schema de ejemplo para extraer información básica
    schema = {
        "type": "object",
        "required": ["total_rows", "total_columns", "column_names"],
        "properties": {
            "total_rows": {
                "type": "integer",
                "description": "Número total de filas con datos (sin contar encabezados)"
            },
            "total_columns": {
                "type": "integer",
                "description": "Número total de columnas"
            },
            "column_names": {
                "type": "array",
                "items": {"type": "string"},
                "description": "Lista de nombres de columnas"
            },
            "summary": {
                "type": "string",
                "description": "Breve resumen del contenido del archivo"
            }
        }
    }
    
    instructions = """
Analiza el archivo Excel y extrae la siguiente información:
- Número total de filas con datos (excluyendo encabezados)
- Número total de columnas
- Nombres de todas las columnas
- Un breve resumen del tipo de datos que contiene
"""
    
    result = extract_structured_data(api_key, excel_path, schema, instructions, model)
    
    if result["success"]:
        import json
        print("\nDatos extraídos:")
        print(json.dumps(result["data"], indent=2, ensure_ascii=False))
    else:
        print(f"\n❌ Error: {result['error']}")


def main():
    parser = argparse.ArgumentParser(
        description="CLI para procesar archivos Excel con OpenAI",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Ejemplos de uso:
  %(prog)s datos.xlsx --interactive
  %(prog)s datos.xlsx --query "Resume el contenido"
  %(prog)s datos.xlsx --extract-structure
  %(prog)s datos.xlsx --query "¿Cuántos registros hay?" --model gpt-4o-mini
        """
    )
    
    parser.add_argument(
        "excel_file",
        type=str,
        help="Ruta al archivo Excel a procesar"
    )
    
    parser.add_argument(
        "--model",
        type=str,
        default="gpt-5",
        help="Modelo de OpenAI a utilizar (por defecto: gpt-4o)"
    )
    
    # Modos de operación (mutuamente excluyentes)
    mode_group = parser.add_mutually_exclusive_group()
    
    mode_group.add_argument(
        "--interactive",
        "-i",
        action="store_true",
        help="Modo interactivo para hacer múltiples consultas"
    )
    
    mode_group.add_argument(
        "--query",
        "-q",
        type=str,
        help="Hacer una consulta única sobre el archivo"
    )
    
    mode_group.add_argument(
        "--extract-structure",
        "-e",
        action="store_true",
        help="Extraer datos estructurados del archivo (ejemplo)"
    )
    
    args = parser.parse_args()
    
    # Validar que el archivo existe
    if not os.path.exists(args.excel_file):
        print(f"❌ Error: El archivo '{args.excel_file}' no existe.")
        sys.exit(1)
    
    # Validar que sea un archivo Excel
    if not args.excel_file.lower().endswith(('.xlsx', '.xls')):
        print(f"⚠️  Advertencia: El archivo no tiene extensión .xlsx o .xls")
    
    # Obtener API key (desde argumento, .env, o variable de entorno)
    api_key = get_api_key()
    
    if not api_key:
        print("❌ Error: Se requiere una API Key de OpenAI.")
        print("   Opciones:")
        print("   1. Configura la variable de entorno OPENAI_API_KEY o API_KEY")
        print("   2. Crea un archivo .env en src/ con: API_KEY=tu-key")
        sys.exit(1)
    
    print(f"\n📊 Archivo: {args.excel_file}")
    print(f"🤖 Modelo: {args.model}")
    print(f"🔑 API Key: {api_key[:20]}...{api_key[-4:]}")
    print()
    
    # Ejecutar según el modo seleccionado
    if args.interactive:
        # Modo interactivo
        processor = OpenAIExcelProcessor(api_key, args.model)
        try:
            processor.upload_excel_file(args.excel_file)
            interactive_mode(processor)
        except Exception as e:
            print(f"❌ Error: {e}")
            sys.exit(1)
    
    elif args.query:
        # Consulta única
        try:
            single_query_mode(args.excel_file, api_key, args.query, args.model)
        except Exception as e:
            print(f"❌ Error: {e}")
            sys.exit(1)
    
    elif args.extract_structure:
        # Extracción estructurada
        try:
            structured_extraction_mode(args.excel_file, api_key, args.model)
        except Exception as e:
            print(f"❌ Error: {e}")
            sys.exit(1)
    
    else:
        # Por defecto, mostrar ayuda si no se especifica modo
        print("⚠️  Debes especificar un modo de operación:")
        print("   --interactive     : Modo interactivo")
        print("   --query           : Consulta única")
        print("   --extract-structure : Extracción estructurada")
        print("\nUsa --help para más información")
        sys.exit(1)


if __name__ == "__main__":
    main()


if __name__ == "__main__":
    main()



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
    pass
