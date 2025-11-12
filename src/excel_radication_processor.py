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
    
    def extract_medicine_codes_from_excel(
        self, 
        excel_path: str, 
        max_retries: int = 3
    ) -> Optional[Dict[str, Any]]:
        """
        Extrae códigos de medicamentos de un archivo Excel con reintentos.
        Busca el código MD más similar basado en principio activo y concentración.
        
        Args:
            excel_path: Ruta al archivo Excel
            max_retries: Número máximo de intentos
            
        Returns:
            Diccionario con los códigos extraídos o None si falla
        """
        
        # Schema para validar la respuesta
        schema = {
            "type": "object",
            "required": ["medicamentos"],
            "properties": {
                "medicamentos": {
                    "type": "array",
                    "minItems": 1,
                    "items": {
                        "type": "object",
                        "required": ["principio_activo", "concentracion", "codigo_md"],
                        "properties": {
                            "principio_activo": {
                                "type": "string"
                            },
                            "concentracion": {
                                "type": "string"
                            },
                            "codigo_md": {
                                "type": "string",
                                "pattern": "^MD[0-9]{6}$"
                            }
                        }
                    }
                }
            }
        }
        
        instructions = """
Analiza el archivo Excel que contiene medicamentos con principios activos y concentraciones.

Para cada medicamento, debes buscar el código MD más similar del rango MD000001 a MD999999.

IMPORTANTE - CONCENTRACIONES:
1. Las concentraciones pueden estar escritas de diferentes formas pero ser equivalentes:
   - "6.5mg" es equivalente a "6.500000mg"
   - "0.5g" es equivalente a "500mg"
   
2. Cuando hay concentraciones como suma (ej: "6.5mg + 2.5mg"):
   - NO significa buscar la suma (9mg)
   - Significa buscar un producto que tenga AMBAS concentraciones: 6.5mg Y 2.5mg
   - El producto debe tener exactamente esas dos concentraciones por separado

3. Presta ESPECIAL atención a:
   - Los números decimales (6.5 vs 6.500000 son iguales)
   - Las unidades de medida (mg, g, ml, etc.)
   - Los ceros insignificantes
   
4. Si hay múltiples principios activos separados por "+", cada uno tiene su concentración.

FORMATO DE RESPUESTA:
Devuelve SOLO un JSON con este formato exacto:
{
  "medicamentos": [
    {
      "principio_activo": "nombre del principio activo del excel",
      "concentracion": "concentración exacta del excel",
      "codigo_md": "MDxxxxxx"
    }
  ]
}

El código_md debe estar en el rango MD000001 a MD999999.
NO agregues explicaciones, SOLO el JSON.
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
                    
                    # Validar que haya al menos un medicamento válido
                    medicamentos = data.get("medicamentos", [])
                    if len(medicamentos) > 0:
                        print(f"✓ Extracción exitosa: {len(medicamentos)} medicamentos encontrados")
                        
                        # Validar códigos MD
                        invalid_codes = [
                            m["principio_activo"] 
                            for m in medicamentos 
                            if not (m.get("codigo_md", "").startswith("MD") and len(m.get("codigo_md", "")) == 8)
                        ]
                        
                        if invalid_codes:
                            print(f"⚠ Advertencia: {len(invalid_codes)} códigos MD inválidos, reintentando...")
                            current_try += 1
                            time.sleep(2)
                            continue
                        
                        break
                    else:
                        print("⚠ No se encontraron medicamentos válidos, reintentando...")
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
            # Valores por defecto
            data = {
                "medicamentos": []
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
        total_medicamentos = 0
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
                data = self.extract_medicine_codes_from_excel(excel_file)
                
                if data and len(data.get("medicamentos", [])) > 0:
                    medicamentos = data["medicamentos"]
                    total_medicamentos += len(medicamentos)
                    
                    # Guardar resultado
                    import json
                    output_file = os.path.join(
                        output_dir, 
                        f"{os.path.splitext(os.path.basename(excel_file))[0]}_processed.json"
                    )
                    
                    with open(output_file, 'w', encoding='utf-8') as f:
                        json.dump(data, f, indent=2, ensure_ascii=False)
                    
                    print(f"✓ Guardado en: {output_file}")
                    print(f"  - Medicamentos procesados: {len(medicamentos)}")
                    
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
        print(f"Total de medicamentos procesados: {total_medicamentos}")
        print(f"Archivos con errores: {len(failed_files)}")
        
        if failed_files:
            print("\nArchivos fallidos:")
            for f in failed_files:
                print(f"  - {os.path.basename(f)}")
        
        return {
            "total_files": len(excel_files),
            "processed": total_processed,
            "total_medicines": total_medicamentos,
            "failed_files": failed_files
        }


def main():
    """Ejemplo de uso del procesador."""
    
    import argparse
    
    parser = argparse.ArgumentParser(
        description="Procesador de archivos Excel para identificación de códigos MD de medicamentos"
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
