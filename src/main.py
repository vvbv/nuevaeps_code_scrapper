#!/usr/bin/env python3
"""
CLI para procesar archivos Excel con OpenAI
===========================================
"""

import argparse
import os
import sys
import json
from concurrent.futures import ThreadPoolExecutor, as_completed
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


def process_single_code(api_key: str, excel_path: str, model: str, line: str, idx: int, total_lines: int):
    """
    Procesa una sola línea del archivo de códigos.
    Función auxiliar para procesamiento paralelo.
    """
    line = line.strip()
    if not line:
        return None
    
    # Separar código original y nombre del producto (puede ser espacio o tab)
    if '\t' in line:
        parts = line.split('\t', 1)
    else:
        parts = line.split(' ', 1)
    
    if len(parts) < 2:
        print(f"[{idx}/{total_lines}] ⚠️  Línea inválida: {line}")
        return {
            "original_line": line,
            "original_code": parts[0] if parts else "",
            "product_name": "",
            "found_code": None,
            "error": "Formato inválido"
        }
    
    original_code = parts[0].strip()
    product_name = parts[1].strip()
    
    print(f"[{idx}/{total_lines}] 🔄 Procesando: {original_code} - {product_name[:50]}...")
    
    # Construir el query
    query = f"Código original: {original_code}\nBusca el código MD para: {product_name}"
    
    try:
        result = simple_excel_query(api_key, excel_path, query, model=model)
        
        if result["success"]:
            response = result["response"]
            print(f"[{idx}/{total_lines}] ✓ {original_code}: {response[:80]}...")
            
            return {
                "original_line": line,
                "original_code": original_code,
                "product_name": product_name,
                "found_code": response,
                "tokens_used": result.get("total_tokens", 0),
                "error": None
            }
        else:
            print(f"[{idx}/{total_lines}] ✗ {original_code}: {result['error']}")
            return {
                "original_line": line,
                "original_code": original_code,
                "product_name": product_name,
                "found_code": None,
                "error": result['error']
            }
    
    except Exception as e:
        print(f"[{idx}/{total_lines}] ✗ {original_code}: Excepción: {str(e)}")
        return {
            "original_line": line,
            "original_code": original_code,
            "product_name": product_name,
            "found_code": None,
            "error": str(e)
        }


def process_codes_to_solve(excel_path: str, api_key: str, txt_path: str, model: str, output_path: str = None, max_workers: int = 8):
    """
    Procesa un archivo de texto con códigos a resolver usando procesamiento paralelo en lotes.
    Lee el archivo donde cada línea tiene: CODIGO_ORIGINAL NOMBRE_PRODUCTO
    Busca el código MD correspondiente en el Excel.
    Los resultados se van guardando de forma acumulativa en result.json.
    
    Args:
        max_workers: Número de peticiones paralelas por lote (por defecto 8)
                    Límite TPM: 1,000,000 tokens/min
                    ~90,000 tokens por petición
                    8 peticiones simultáneas = ~720,000 tokens (margen de seguridad)
    """
    print(txt_path)
    if not os.path.exists(txt_path):
        print(f"❌ Error: El archivo '{txt_path}' no existe.")
        return
    
    # Definir la ruta del archivo de resultados
    if output_path is None:
        # Buscar la carpeta resources
        current_dir = os.path.dirname(os.path.abspath(__file__))
        parent_dir = os.path.dirname(current_dir)
        
        if os.path.exists(os.path.join(current_dir, 'resources')):
            resources_dir = os.path.join(current_dir, 'resources')
        elif os.path.exists(os.path.join(parent_dir, 'resources')):
            resources_dir = os.path.join(parent_dir, 'resources')
        else:
            resources_dir = os.path.join(parent_dir, 'resources')
            os.makedirs(resources_dir, exist_ok=True)
        
        output_path = os.path.join(resources_dir, 'result.json')
    
    # Cargar resultados existentes si el archivo existe
    existing_results = []
    processed_lines = set()  # Para trackear líneas ya procesadas exitosamente
    
    if os.path.exists(output_path):
        try:
            with open(output_path, 'r', encoding='utf-8') as f:
                existing_results = json.load(f)
            print(f"📄 Cargados {len(existing_results)} resultados previos de {output_path}")
            
            # Crear un set de líneas ya procesadas exitosamente (sin error)
            for result in existing_results:
                if result.get('error') is None:
                    processed_lines.add(result.get('original_line', '').strip())
            
            print(f"✅ {len(processed_lines)} códigos ya procesados exitosamente (se omitirán)")
        except Exception as e:
            print(f"⚠️  No se pudieron cargar resultados previos: {e}")
    
    print(f"\nProcesando códigos desde: {txt_path}")
    print(f"Usando Excel: {excel_path}")
    print(f"Guardando en: {output_path}")
    print(f"⚡ Procesamiento en lotes de {max_workers} peticiones simultáneas")
    print(f"💡 Límite TPM: 1,000,000 tokens (~{max_workers} peticiones × ~90k tokens)")
    print("-" * 80)
    
    # Leer el archivo de códigos y filtrar líneas vacías y ya procesadas
    with open(txt_path, 'r', encoding='utf-8') as f:
        all_lines = [(idx, line) for idx, line in enumerate(f.readlines(), 1) if line.strip()]
    
    # Filtrar las líneas ya procesadas exitosamente
    lines = [(idx, line) for idx, line in all_lines if line.strip() not in processed_lines]
    
    skipped_count = len(all_lines) - len(lines)
    if skipped_count > 0:
        print(f"⏭️  Omitiendo {skipped_count} códigos ya procesados exitosamente")
    
    if not lines:
        print("✅ No hay códigos nuevos para procesar. Todos ya fueron procesados exitosamente.")
        return
    
    results = existing_results.copy()
    total_lines = len(lines)
    processed_count = 0
    
    # Procesar en lotes
    for batch_start in range(0, len(lines), max_workers):
        batch_end = min(batch_start + max_workers, len(lines))
        batch = lines[batch_start:batch_end]
        batch_num = (batch_start // max_workers) + 1
        total_batches = (len(lines) + max_workers - 1) // max_workers
        
        print(f"\n{'='*80}")
        print(f"📦 LOTE {batch_num}/{total_batches} - Procesando {len(batch)} códigos...")
        print(f"{'='*80}")
        
        # Procesar el lote actual en paralelo
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            # Crear tareas para este lote
            futures = {
                executor.submit(process_single_code, api_key, excel_path, model, line, idx, total_lines): idx
                for idx, line in batch
            }
            
            # Esperar a que todas las tareas del lote se completen
            for future in as_completed(futures):
                try:
                    result_entry = future.result()
                    if result_entry:
                        results.append(result_entry)
                        processed_count += 1
                        
                        # Guardar después de cada resultado completado
                        try:
                            with open(output_path, 'w', encoding='utf-8') as f:
                                json.dump(results, f, indent=2, ensure_ascii=False)
                        except Exception as e:
                            print(f"⚠️  Error al guardar resultados: {e}")
                
                except Exception as e:
                    print(f"⚠️  Error procesando future: {e}")
        
        print(f"✅ Lote {batch_num}/{total_batches} completado. Total procesados: {processed_count}/{total_lines}")
    
    # Resumen
    print("\n" + "="*80)
    print("RESUMEN DE PROCESAMIENTO")
    print("="*80)
    print(f"Total de resultados en archivo: {len(results)}")
    print(f"Procesados en esta ejecución: {processed_count}")
    print(f"Exitosas (total): {sum(1 for r in results if r.get('error') is None)}")
    print(f"Con errores (total): {sum(1 for r in results if r.get('error') is not None)}")
    print(f"\n💾 Resultados guardados en: {output_path}")


def retry_failed_codes(excel_path: str, api_key: str, result_json_path: str, model: str, max_workers: int = 8):
    """
    Reintenta procesar los códigos que tuvieron error en result.json.
    Actualiza los registros con error y los marca como exitosos si se completan.
    """
    if not os.path.exists(result_json_path):
        print(f"❌ Error: El archivo '{result_json_path}' no existe.")
        return
    
    # Cargar resultados existentes
    try:
        with open(result_json_path, 'r', encoding='utf-8') as f:
            all_results = json.load(f)
    except Exception as e:
        print(f"❌ Error cargando {result_json_path}: {e}")
        return
    
    # Filtrar solo los que tienen error
    failed_results = [r for r in all_results if r.get('error') is not None]
    
    if not failed_results:
        print("✅ No hay registros con error para reintentar.")
        return
    
    print(f"\n🔄 Reintentando {len(failed_results)} códigos con error...")
    print(f"Usando Excel: {excel_path}")
    print(f"⚡ Procesamiento en lotes de {max_workers} peticiones simultáneas")
    print("-" * 80)
    
    retry_count = 0
    success_count = 0
    
    # Crear un diccionario para actualizar resultados por índice
    results_dict = {i: r for i, r in enumerate(all_results)}
    
    # Encontrar los índices de los elementos con error
    failed_indices = [i for i, r in enumerate(all_results) if r.get('error') is not None]
    
    # Procesar en lotes
    for batch_start in range(0, len(failed_indices), max_workers):
        batch_end = min(batch_start + max_workers, len(failed_indices))
        batch_indices = failed_indices[batch_start:batch_end]
        batch_num = (batch_start // max_workers) + 1
        total_batches = (len(failed_indices) + max_workers - 1) // max_workers
        
        print(f"\n{'='*80}")
        print(f"📦 LOTE {batch_num}/{total_batches} - Reintentando {len(batch_indices)} códigos...")
        print(f"{'='*80}")
        
        # Procesar el lote actual en paralelo
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            # Crear tareas para este lote
            futures = {}
            for idx in batch_indices:
                result = all_results[idx]
                original_code = result.get('original_code', '')
                product_name = result.get('product_name', '')
                line = result.get('original_line', f"{original_code}\t{product_name}")
                
                future = executor.submit(
                    process_single_code, 
                    api_key, 
                    excel_path, 
                    model, 
                    line, 
                    idx + 1,  # Mostrar índice base 1
                    len(all_results)
                )
                futures[future] = idx
            
            # Esperar a que todas las tareas del lote se completen
            for future in as_completed(futures):
                original_idx = futures[future]
                try:
                    result_entry = future.result()
                    if result_entry:
                        retry_count += 1
                        
                        # Si no tiene error, es exitoso
                        if result_entry.get('error') is None:
                            success_count += 1
                        
                        # Actualizar el resultado en la posición original
                        results_dict[original_idx] = result_entry
                        
                        # Guardar después de cada actualización
                        try:
                            updated_results = [results_dict[i] for i in range(len(results_dict))]
                            with open(result_json_path, 'w', encoding='utf-8') as f:
                                json.dump(updated_results, f, indent=2, ensure_ascii=False)
                        except Exception as e:
                            print(f"⚠️  Error al guardar resultados: {e}")
                
                except Exception as e:
                    print(f"⚠️  Error procesando future: {e}")
        
        print(f"✅ Lote {batch_num}/{total_batches} completado.")
    
    # Resumen
    print("\n" + "="*80)
    print("RESUMEN DE REINTENTOS")
    print("="*80)
    print(f"Códigos reintentados: {retry_count}")
    print(f"Exitosos en reintento: {success_count}")
    print(f"Aún con error: {retry_count - success_count}")
    
    # Estadísticas finales
    final_results = [results_dict[i] for i in range(len(results_dict))]
    print(f"\nEstadísticas finales:")
    print(f"  Total de registros: {len(final_results)}")
    print(f"  Exitosos: {sum(1 for r in final_results if r.get('error') is None)}")
    print(f"  Con errores: {sum(1 for r in final_results if r.get('error') is not None)}")
    print(f"\n💾 Resultados actualizados en: {result_json_path}")


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
    
    mode_group.add_argument(
        "--process-codes",
        "-p",
        action="store_true",
        help="Procesar archivo codes_to_solve.txt para obtener códigos MD"
    )
    
    mode_group.add_argument(
        "--retry-errors",
        "-r",
        action="store_true",
        help="Reintentar procesar los códigos que tuvieron error en result.json"
    )
    
    parser.add_argument(
        "--codes-file",
        type=str,
        default="resources/codes_to_solve.txt",
        help="Ruta al archivo de códigos a procesar (por defecto: resources/cods_to_solve.txt)"
    )
    
    parser.add_argument(
        "--result-file",
        type=str,
        default="resources/result.json",
        help="Ruta al archivo result.json para reintentar errores (por defecto: resources/result.json)"
    )
    
    parser.add_argument(
        "--output",
        "-o",
        type=str,
        help="Ruta del archivo de salida para resultados (opcional)"
    )
    
    parser.add_argument(
        "--max-workers",
        "-w",
        type=int,
        default=8,
        help="Número de peticiones paralelas por lote para --process-codes (por defecto: 8, límite TPM: 1M)"
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
    
    elif args.process_codes:
        # Procesar códigos desde archivo
        try:
            process_codes_to_solve(
                excel_path=args.excel_file,
                api_key=api_key,
                txt_path=args.codes_file,
                model=args.model,
                output_path=args.output,
                max_workers=args.max_workers
            )
        except Exception as e:
            print(f"❌ Error: {e}")
            sys.exit(1)
    
    elif args.retry_errors:
        # Reintentar códigos con error
        try:
            # Buscar el archivo result.json
            result_path = args.result_file
            if not os.path.isabs(result_path):
                # Buscar en resources
                current_dir = os.path.dirname(os.path.abspath(__file__))
                parent_dir = os.path.dirname(current_dir)
                
                # Intentar diferentes ubicaciones
                possible_paths = [
                    result_path,
                    os.path.join(current_dir, result_path),
                    os.path.join(parent_dir, result_path),
                    os.path.join(current_dir, 'resources', 'result.json'),
                    os.path.join(parent_dir, 'resources', 'result.json')
                ]
                
                result_path = None
                for path in possible_paths:
                    if os.path.exists(path):
                        result_path = path
                        break
                
                if not result_path:
                    print(f"❌ Error: No se encontró el archivo result.json")
                    print(f"   Ubicaciones buscadas:")
                    for path in possible_paths:
                        print(f"   - {path}")
                    sys.exit(1)
            
            retry_failed_codes(
                excel_path=args.excel_file,
                api_key=api_key,
                result_json_path=result_path,
                model=args.model,
                max_workers=args.max_workers
            )
        except Exception as e:
            print(f"❌ Error: {e}")
            sys.exit(1)
    
    else:
        # Por defecto, mostrar ayuda si no se especifica modo
        print("⚠️  Debes especificar un modo de operación:")
        print("   --interactive       : Modo interactivo")
        print("   --query             : Consulta única")
        print("   --extract-structure : Extracción estructurada")
        print("   --process-codes     : Procesar archivo de códigos a resolver")
        print("   --retry-errors      : Reintentar códigos con error en result.json")
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
