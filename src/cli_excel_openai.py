#!/usr/bin/env python3
"""
CLI para procesar archivos Excel con OpenAI
===========================================
Este script demuestra cómo usar el módulo openai_excel_helper para:
1. Cargar un archivo Excel
2. Hacer múltiples consultas sobre el contenido
3. Extraer datos estructurados

Uso:
    python cli_excel_openai.py <archivo.xlsx> --api-key <tu_api_key>
    python cli_excel_openai.py <archivo.xlsx> --api-key <tu_api_key> --interactive
    python cli_excel_openai.py <archivo.xlsx> --api-key <tu_api_key> --query "¿Cuántas filas tiene?"
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
  %(prog)s datos.xlsx --api-key sk-... --interactive
  %(prog)s datos.xlsx --api-key sk-... --query "Resume el contenido"
  %(prog)s datos.xlsx --api-key sk-... --extract-structure
  %(prog)s datos.xlsx --api-key sk-... --query "¿Cuántos registros hay?" --model gpt-4o-mini
        """
    )
    
    parser.add_argument(
        "excel_file",
        type=str,
        help="Ruta al archivo Excel a procesar"
    )
    
    parser.add_argument(
        "--api-key",
        type=str,
        default=None,
        help="API Key de OpenAI (opcional si está en .env o variable de entorno OPENAI_API_KEY)"
    )
    
    parser.add_argument(
        "--model",
        type=str,
        default="gpt-4o",
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
    api_key = get_api_key(args.api_key)
    
    if not api_key:
        print("❌ Error: Se requiere una API Key de OpenAI.")
        print("   Opciones:")
        print("   1. Usa --api-key <tu-key>")
        print("   2. Configura la variable de entorno OPENAI_API_KEY o API_KEY")
        print("   3. Crea un archivo .env en src/ con: API_KEY=tu-key")
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
