#!/usr/bin/env python3
"""
Ejemplo de integración con flujo similar a cli_radicacion.py
============================================================
Este ejemplo muestra cómo integrar el procesamiento de Excel
en un flujo de trabajo similar al de Distri-Hub.
"""

import os
import time
from typing import Dict, Any, List, Optional
from openai_excel_helper import OpenAIExcelProcessor, extract_structured_data
from config import get_api_key


class ExcelRadicationProcessor:
    """
    Procesador de archivos Excel para radicación.
    Similar al flujo de cli_radicacion.py pero para Excel.
    """
    
    def __init__(self, api_key: str, model: str = "gpt-4o"):
        self.api_key = api_key
        self.model = model
        self.processor = OpenAIExcelProcessor(api_key, model)
        
    def validate_excel_file(self, excel_path: str) -> bool:
        """Valida que el archivo Excel existe y es válido."""
        if not os.path.exists(excel_path):
            print(f"❌ Error: Archivo no encontrado: {excel_path}")
            return False
        
        if not excel_path.lower().endswith(('.xlsx', '.xls')):
            print(f"❌ Error: El archivo debe ser Excel (.xlsx o .xls)")
            return False
        
        return True
    
    def extract_invoices_from_excel(
        self, 
        excel_path: str, 
        max_retries: int = 3
    ) -> Optional[Dict[str, Any]]:
        """
        Extrae facturas de un archivo Excel con reintentos.
        Similar a la extracción de datos médicos en Distri-Hub.
        
        Args:
            excel_path: Ruta al archivo Excel
            max_retries: Número máximo de intentos
            
        Returns:
            Diccionario con las facturas extraídas o None si falla
        """
        
        # Schema para validar la respuesta (similar a MEDICAL_DATA_JSON_SCHEMA)
        schema = {
            "type": "object",
            "required": ["facturas", "resumen"],
            "properties": {
                "facturas": {
                    "type": "array",
                    "minItems": 1,
                    "items": {
                        "type": "object",
                        "required": ["numero_factura", "nit_cliente", "valor_total"],
                        "properties": {
                            "numero_factura": {
                                "type": "string",
                                "pattern": "^[A-Z0-9\\-]+$"
                            },
                            "nit_cliente": {
                                "type": "string",
                                "pattern": "^[0-9]{6,15}$"
                            },
                            "valor_total": {
                                "type": "number",
                                "minimum": 0
                            },
                            "fecha": {"type": "string"},
                            "items": {
                                "type": "array",
                                "items": {
                                    "type": "object",
                                    "properties": {
                                        "descripcion": {"type": "string"},
                                        "cantidad": {"type": "integer"},
                                        "valor_unitario": {"type": "number"}
                                    }
                                }
                            }
                        }
                    }
                },
                "resumen": {
                    "type": "object",
                    "properties": {
                        "total_facturas": {"type": "integer"},
                        "valor_total_general": {"type": "number"}
                    }
                }
            }
        }
        
        instructions = """
Analiza el archivo Excel que contiene facturas y extrae la siguiente información:

1. Lista de todas las facturas con:
   - Número de factura (debe ser alfanumérico, ej: FAC-001, 12345)
   - NIT del cliente (solo números, entre 6 y 15 dígitos)
   - Valor total de la factura
   - Fecha de la factura (si está disponible)
   - Items de la factura (descripción, cantidad, valor unitario)

2. Resumen con:
   - Total de facturas encontradas
   - Valor total general (suma de todas las facturas)

IMPORTANTE:
- Si un número de factura no es válido o un NIT tiene menos de 6 dígitos, usa valores por defecto
- Asegúrate de que cada factura tenga al menos: número, NIT y valor
- El JSON debe cumplir EXACTAMENTE con el schema proporcionado

Devuelve SOLO el JSON, sin explicaciones adicionales.
"""
        
        current_try = 0
        data = None
        
        print(f"\n{'='*80}")
        print(f"Procesando archivo: {excel_path}")
        print(f"{'='*80}\n")
        
        while max_retries > current_try:
            try:
                print(f"Intento {current_try + 1}/{max_retries}...")
                
                result = extract_structured_data(
                    api_key=self.api_key,
                    excel_path=excel_path,
                    schema=schema,
                    instructions=instructions,
                    model=self.model
                )
                
                if result["success"]:
                    data = result["data"]
                    
                    # Validar que haya al menos una factura válida
                    facturas = data.get("facturas", [])
                    if len(facturas) > 0:
                        print(f"✓ Extracción exitosa: {len(facturas)} facturas encontradas")
                        
                        # Validar NITs
                        invalid_nits = [
                            f["numero_factura"] 
                            for f in facturas 
                            if len(f.get("nit_cliente", "")) < 6
                        ]
                        
                        if invalid_nits:
                            print(f"⚠ Advertencia: {len(invalid_nits)} NITs inválidos, reintentando...")
                            current_try += 1
                            time.sleep(2)  # Esperar antes de reintentar
                            continue
                        
                        break
                    else:
                        print("⚠ No se encontraron facturas válidas, reintentando...")
                        current_try += 1
                        time.sleep(2)
                        continue
                else:
                    print(f"✗ Error en la extracción: {result.get('error', 'Desconocido')}")
                    current_try += 1
                    time.sleep(2)
                    continue
                    
            except Exception as e:
                print(f"✗ Error en el intento: {str(e)}")
                current_try += 1
                time.sleep(2)
        
        if not data:
            print(f"\n✗ No se pudo procesar el archivo después de {max_retries} intentos")
            # Valores por defecto (similar a Distri-Hub)
            data = {
                "facturas": [],
                "resumen": {
                    "total_facturas": 0,
                    "valor_total_general": 0
                }
            }
            print("⚠ Usando datos por defecto vacíos")
        
        return data
    
    def process_multiple_excel_files(
        self, 
        excel_files: List[str],
        output_dir: str = "output"
    ) -> Dict[str, Any]:
        """
        Procesa múltiples archivos Excel.
        Similar al loop en cli_radicacion.py
        
        Args:
            excel_files: Lista de rutas a archivos Excel
            output_dir: Directorio para guardar resultados
            
        Returns:
            Diccionario con estadísticas del procesamiento
        """
        
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
        
        total_processed = 0
        total_facturas = 0
        failed_files = []
        
        print(f"\n{'='*80}")
        print(f"Procesando {len(excel_files)} archivos Excel")
        print(f"{'='*80}\n")
        
        for idx, excel_file in enumerate(excel_files, 1):
            print(f"\n[{idx}/{len(excel_files)}] Procesando: {os.path.basename(excel_file)}")
            print("-" * 80)
            
            if not self.validate_excel_file(excel_file):
                failed_files.append(excel_file)
                continue
            
            try:
                # Extraer datos
                data = self.extract_invoices_from_excel(excel_file)
                
                if data and len(data.get("facturas", [])) > 0:
                    facturas = data["facturas"]
                    total_facturas += len(facturas)
                    
                    # Guardar resultado
                    import json
                    output_file = os.path.join(
                        output_dir, 
                        f"{os.path.splitext(os.path.basename(excel_file))[0]}_processed.json"
                    )
                    
                    with open(output_file, 'w', encoding='utf-8') as f:
                        json.dump(data, f, indent=2, ensure_ascii=False)
                    
                    print(f"✓ Guardado en: {output_file}")
                    print(f"  - Facturas: {len(facturas)}")
                    print(f"  - Valor total: ${data['resumen']['valor_total_general']:,.2f}")
                    
                    total_processed += 1
                else:
                    print(f"⚠ No se pudieron extraer datos del archivo")
                    failed_files.append(excel_file)
                
            except Exception as e:
                print(f"✗ Error procesando archivo: {str(e)}")
                failed_files.append(excel_file)
        
        # Resumen final
        print(f"\n{'='*80}")
        print("RESUMEN DEL PROCESAMIENTO")
        print(f"{'='*80}")
        print(f"Archivos procesados exitosamente: {total_processed}/{len(excel_files)}")
        print(f"Total de facturas extraídas: {total_facturas}")
        print(f"Archivos con errores: {len(failed_files)}")
        
        if failed_files:
            print("\nArchivos fallidos:")
            for f in failed_files:
                print(f"  - {os.path.basename(f)}")
        
        return {
            "total_files": len(excel_files),
            "processed": total_processed,
            "total_invoices": total_facturas,
            "failed_files": failed_files
        }


def main():
    """Ejemplo de uso del procesador."""
    
    import argparse
    
    parser = argparse.ArgumentParser(
        description="Procesador de archivos Excel para radicación"
    )
    parser.add_argument(
        "excel_files",
        nargs="+",
        help="Archivos Excel a procesar"
    )
    parser.add_argument(
        "--api-key",
        default=None,
        help="API Key de OpenAI (opcional si está en .env o variable de entorno)"
    )
    parser.add_argument(
        "--model",
        default="gpt-4o",
        help="Modelo de OpenAI (default: gpt-4o)"
    )
    parser.add_argument(
        "--output-dir",
        default="output",
        help="Directorio de salida (default: output)"
    )
    parser.add_argument(
        "--max-retries",
        type=int,
        default=3,
        help="Número máximo de reintentos (default: 3)"
    )
    
    args = parser.parse_args()
    
    # Obtener API key (desde argumento, .env, o variable de entorno)
    api_key = get_api_key(args.api_key)
    
    if not api_key:
        print("❌ Error: Se requiere una API Key de OpenAI.")
        print("   Opciones:")
        print("   1. Usa --api-key <tu-key>")
        print("   2. Configura la variable de entorno OPENAI_API_KEY o API_KEY")
        print("   3. Crea un archivo .env en src/ con: API_KEY=tu-key")
        import sys
        sys.exit(1)
    
    # Crear procesador
    processor = ExcelRadicationProcessor(api_key, args.model)
    
    # Procesar archivos
    results = processor.process_multiple_excel_files(
        args.excel_files,
        args.output_dir
    )
    
    # Guardar estadísticas
    import json
    stats_file = os.path.join(args.output_dir, "processing_stats.json")
    with open(stats_file, 'w', encoding='utf-8') as f:
        json.dump(results, f, indent=2, ensure_ascii=False)
    
    print(f"\nEstadísticas guardadas en: {stats_file}")


if __name__ == "__main__":
    main()
