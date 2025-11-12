#!/usr/bin/env python3
"""
Script para convertir result.json a CSV
Extrae la información del found_code en columnas separadas
"""

import json
import csv
import sys
import os
import re


def extract_md_code(found_code_str):
    """
    Extrae el código MD del campo found_code.
    Puede estar en formato JSON o como texto plano.
    """
    if not found_code_str:
        return None
    
    try:
        # Intentar parsear como JSON
        data = json.loads(found_code_str)
        return data.get('codigo')
    except (json.JSONDecodeError, TypeError):
        # Si no es JSON, buscar patrón MD en el texto
        match = re.search(r'MD\d{6}', found_code_str)
        if match:
            return match.group(0)
    
    return None


def extract_description(found_code_str):
    """
    Extrae la descripción del campo found_code.
    Puede estar en formato JSON o como texto plano.
    """
    if not found_code_str:
        return None
    
    try:
        # Intentar parsear como JSON
        data = json.loads(found_code_str)
        return data.get('descripcion')
    except (json.JSONDecodeError, TypeError):
        # Si no es JSON, devolver el texto completo
        return found_code_str
    
    return None


def result_json_to_csv(json_path, csv_path=None):
    """
    Convierte result.json a CSV con columnas adicionales para found_code.
    
    Columnas del CSV:
    - original_code: Código original del producto
    - product_name: Nombre del producto
    - found_md_code: Código MD extraído
    - found_description: Descripción extraída
    - tokens_used: Tokens consumidos
    - error: Error si hubo alguno
    - original_line: Línea original del archivo
    """
    # Determinar ruta del CSV si no se proporciona
    if csv_path is None:
        csv_path = json_path.replace('.json', '.csv')
    
    # Leer el JSON
    try:
        with open(json_path, 'r', encoding='utf-8') as f:
            results = json.load(f)
    except Exception as e:
        print(f"❌ Error leyendo {json_path}: {e}")
        return False
    
    if not results:
        print("⚠️  El archivo JSON está vacío")
        return False
    
    print(f"📄 Leyendo {len(results)} registros de {json_path}")
    
    # Crear el CSV
    try:
        with open(csv_path, 'w', newline='', encoding='utf-8') as csvfile:
            fieldnames = [
                'original_code',
                'product_name',
                'found_md_code',
                'found_description',
                'tokens_used',
                'error',
                'original_line'
            ]
            
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            writer.writeheader()
            
            # Procesar cada resultado
            for result in results:
                found_code = result.get('found_code', '')
                
                row = {
                    'original_code': result.get('original_code', ''),
                    'product_name': result.get('product_name', ''),
                    'found_md_code': extract_md_code(found_code) or '',
                    'found_description': extract_description(found_code) or '',
                    'tokens_used': result.get('tokens_used', 0),
                    'error': result.get('error', ''),
                    'original_line': result.get('original_line', '')
                }
                
                writer.writerow(row)
        
        print(f"✅ CSV generado exitosamente: {csv_path}")
        print(f"📊 Total de registros: {len(results)}")
        
        # Estadísticas
        successful = sum(1 for r in results if not r.get('error'))
        with_error = sum(1 for r in results if r.get('error'))
        
        print(f"   - Exitosos: {successful}")
        print(f"   - Con error: {with_error}")
        
        return True
    
    except Exception as e:
        print(f"❌ Error escribiendo CSV: {e}")
        return False


def main():
    """
    Función principal del script
    """
    if len(sys.argv) < 2:
        # Buscar result.json en ubicaciones comunes
        current_dir = os.path.dirname(os.path.abspath(__file__))
        parent_dir = os.path.dirname(current_dir)
        
        possible_paths = [
            os.path.join(current_dir, 'resources', 'result.json'),
            os.path.join(parent_dir, 'resources', 'result.json'),
            os.path.join(current_dir, 'result.json'),
            os.path.join(parent_dir, 'result.json'),
        ]
        
        json_path = None
        for path in possible_paths:
            if os.path.exists(path):
                json_path = path
                break
        
        if not json_path:
            print("Uso: python3 result_to_csv.py [ruta_al_result.json] [ruta_salida.csv]")
            print("\nNo se encontró result.json en las ubicaciones predeterminadas:")
            for path in possible_paths:
                print(f"  - {path}")
            sys.exit(1)
    else:
        json_path = sys.argv[1]
    
    # Verificar que el archivo existe
    if not os.path.exists(json_path):
        print(f"❌ Error: El archivo '{json_path}' no existe")
        sys.exit(1)
    
    # Determinar ruta de salida
    csv_path = sys.argv[2] if len(sys.argv) > 2 else None
    
    # Convertir
    success = result_json_to_csv(json_path, csv_path)
    
    sys.exit(0 if success else 1)


if __name__ == "__main__":
    main()
