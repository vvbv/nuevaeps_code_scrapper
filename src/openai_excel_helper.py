"""
OpenAI Excel Helper
===================
Este módulo proporciona funciones para trabajar con archivos Excel usando OpenAI.
Permite subir archivos Excel y realizar múltiples consultas sobre el contenido del documento.

Basado en la implementación de openai_medical_helper.py de Distri-Hub.
"""

import openai
import os
import json
import base64
import requests
from typing import Optional, Dict, Any, List


class OpenAIExcelProcessor:
    """
    Clase para procesar archivos Excel con OpenAI.
    Permite mantener el contexto del archivo entre múltiples consultas.
    """
    
    def __init__(self, api_key: str, model: str = "gpt-5"):
        """
        Inicializa el procesador de Excel con OpenAI.
        
        Args:
            api_key: Clave API de OpenAI
            model: Modelo a utilizar (por defecto gpt-5 que soporta archivos)
        """
        self.api_key = api_key
        self.model = model
        self.file_path = None
        self.df = None
        self.num_rows = 0
        self.num_cols = 0
        self.columns = []
        self.conversation_messages = []
        
    def upload_excel_file(self, excel_path: str) -> str:
        """
        Carga un archivo Excel en memoria para su procesamiento.
        Convierte el contenido a formato que la IA puede procesar.
        
        Args:
            excel_path: Ruta al archivo Excel
            
        Returns:
            Confirmación de carga
            
        Raises:
            FileNotFoundError: Si el archivo no existe
            Exception: Si hay error en la carga
        """
        if not os.path.exists(excel_path):
            raise FileNotFoundError(f"Archivo Excel no encontrado: {excel_path}")
        
        print(f"Cargando archivo {excel_path}...")
        
        try:
            import pandas as pd
            
            # Leer el archivo Excel
            self.df = pd.read_excel(excel_path)
            self.file_path = excel_path
            
            # Obtener info básica
            self.num_rows = len(self.df)
            self.num_cols = len(self.df.columns)
            self.columns = list(self.df.columns)
            
            print(f"Archivo cargado exitosamente:")
            print(f"  - Filas: {self.num_rows}")
            print(f"  - Columnas: {self.num_cols}")
            print(f"  - Nombres: {', '.join(self.columns[:5])}{'...' if len(self.columns) > 5 else ''}")
            
            return f"loaded_{os.path.basename(excel_path)}"
            
        except Exception as e:
            print(f"Error al cargar archivo: {e}")
            raise
    
    def query_excel(self, prompt: str, temperature: float = 0.3) -> Dict[str, Any]:
        """
        Realiza una consulta sobre el archivo Excel previamente cargado.
        
        Args:
            prompt: Pregunta o instrucción sobre el archivo
            temperature: Temperatura para la generación (0-1)
            
        Returns:
            Diccionario con la respuesta
            
        Raises:
            ValueError: Si no hay archivo cargado
        """
        if not hasattr(self, 'df') or self.df is None:
            raise ValueError("No hay archivo Excel cargado. Usa upload_excel_file() primero.")
        
        try:
            # Preparar el contenido del Excel (limitar si es muy grande)
            if self.num_rows > 100:
                sample_df = self.df.head(100)
                csv_content = sample_df.to_csv(index=False)
                content_note = f"\nNOTA: El archivo tiene {self.num_rows} filas, pero solo se muestran las primeras 100 para análisis."
            else:
                csv_content = self.df.to_csv(index=False)
                content_note = ""
            
            # Preparar el contexto del archivo
            file_context = f"""Información del archivo Excel:
- Total de filas: {self.num_rows}
- Total de columnas: {self.num_cols}
- Columnas: {', '.join(self.columns)}

Contenido (formato CSV):
```
{csv_content}
```
{content_note}"""
            
            # Añadir el contexto del archivo solo en la primera consulta
            if len(self.conversation_messages) == 0:
                self.conversation_messages.append({
                    "role": "user",
                    "content": f"{file_context}\n\nArchivo cargado. Estoy listo para responder preguntas."
                })
                self.conversation_messages.append({
                    "role": "assistant",
                    "content": "Entendido. He analizado el archivo Excel. ¿Qué deseas saber?"
                })
            
            # Añadir el mensaje del usuario al historial
            self.conversation_messages.append({
                "role": "user",
                "content": prompt
            })
            
            # Hacer la consulta con el contexto del archivo
            client = openai.OpenAI(api_key=self.api_key)
            
            response = client.chat.completions.create(
                model=self.model,
                messages=[
                    {
                        "role": "system",
                        "content": "Eres un asistente experto en analizar archivos Excel. "
                                   "Respondes de manera precisa y estructurada basándote en los datos del archivo. "
                                   "El usuario te ha proporcionado el contenido completo del archivo."
                    }
                ] + self.conversation_messages,
                temperature=temperature
            )
            
            assistant_message = response.choices[0].message.content
            
            # Añadir la respuesta al historial
            self.conversation_messages.append({
                "role": "assistant",
                "content": assistant_message
            })
            
            result = {
                "success": True,
                "response": assistant_message,
                "model": self.model,
                "total_tokens": response.usage.total_tokens
            }
            
            return result
            
        except Exception as e:
            print(f"Error en la consulta: {e}")
            return {
                "success": False,
                "error": str(e)
            }
    
    def query_with_excel_content(
            self, excel_path: str, prompt: str, temperature: float = 1,
            aux_original_code: str = None
    ) -> Dict[str, Any]:
        """
        Procesa un archivo Excel directamente convirtiendo su contenido a texto.
        Útil para archivos pequeños o consultas únicas.
        
        Args:
            excel_path: Ruta al archivo Excel
            prompt: Pregunta o instrucción sobre el archivo
            temperature: Temperatura para la generación (0-1)
            
        Returns:
            Diccionario con la respuesta
        """
        if not os.path.exists(excel_path):
            raise FileNotFoundError(f"Archivo Excel no encontrado: {excel_path}")
        
        print(f"Procesando {excel_path} con OpenAI...")

        expected_response = {
            "type": "object",
            "properties": {
                "codigo": {
                    "type": "string",
                    "description": "Código alfanumérico identificador del elemento",
                    "pattern": "^[A-Za-z0-9]+$"
                },
                    "descripcion": {
                    "type": "string",
                    "description": "Descripción alfanumérica del elemento",
                    "pattern": "^[A-Za-z0-9\\s]+$"
                },
                    "original_code": {
                    "type": "string",
                    "description": "Código original inmutable",
                    "const": aux_original_code
                },
                    "original_file": {
                    "type": "string",
                    "description": "Nombre o ruta del archivo original inmutable",
                    "const": excel_path
                }
            },
            "required": ["codigo", "descripcion"]
            }

        
        try:
            # Leer el archivo Excel con pandas
            import pandas as pd
            
            # Leer el Excel
            df = pd.read_excel(excel_path)
            
            # Convertir a CSV string (más fácil de procesar por la IA)
            csv_content = df.to_csv(index=False)
            
            # Obtener info básica
            num_rows = len(df)
            num_cols = len(df.columns)
            columns = list(df.columns)
            
            # Limitar el contenido si es muy grande (primeras 100 filas)
            if num_rows > 5000:
                sample_df = df.head(5000)
                csv_content = sample_df.to_csv(index=False)
                content_note = f"\nNOTA: El archivo tiene {num_rows} filas, pero solo se muestran las primeras 100 para análisis."
            else:
                content_note = ""
            
            # Preparar el prompt con el contenido del Excel
            full_prompt = f"""Analiza el siguiente archivo Excel que tiene {num_rows} filas y {num_cols} columnas.
Columnas: {', '.join(columns)}

Contenido del archivo (formato CSV):
```
{csv_content}
```
{content_note}

INSTRUCCIONES IMPORTANTES:
- Cuando se te pida buscar un medicamento o principio activo, debes buscar en la COLUMNA D
- NO necesitas una coincidencia EXACTA, busca la descripción más aproximada o similar
- Considera variaciones en la escritura, concentraciones equivalentes, y sinónimos
- El CÓDIGO que debes retornar es el que se encuentra en la COLUMNA A de la fila que mejor coincida
- Busca en todas las filas del archivo
- Si hay múltiples coincidencias cercanas, elige la más similar considerando:
  * Principio activo
  * Concentración (ejemplo: 0.064g puede estar como 64mg, son equivalentes)
  * Forma farmacéutica

Pregunta del usuario: {prompt}

Responde basándote en los datos proporcionados. Si buscas un código, indica claramente el valor de la columna A.

cumpliendo la siguiente estructura de respuesta:

{json.dumps(expected_response, indent=2, ensure_ascii=False)}

"""
            
            # Usar la API de OpenAI
            client = openai.OpenAI(api_key=self.api_key)
            
            response = client.chat.completions.create(
                model=self.model,
                temperature=temperature,
                messages=[
                    {
                        "role": "system",
                        "content": "Eres un asistente experto en analizar archivos Excel. "
                                   "Respondes de manera precisa y estructurada basándote en los datos del archivo."
                    },
                    {
                        "role": "user",
                        "content": full_prompt
                    }
                ]
            )
            
            assistant_message = response.choices[0].message.content
            
            return {
                "success": True,
                "response": assistant_message,
                "model": self.model,
                "total_tokens": response.usage.total_tokens
            }
            
        except Exception as e:
            print(f"Error procesando {excel_path}: {e}")
            return {
                "success": False,
                "error": str(e)
            }
    
    def reset_conversation(self):
        """Reinicia el historial de conversación."""
        self.conversation_messages = []
        print("Historial de conversación reiniciado.")
    
    def get_conversation_history(self) -> List[Dict[str, str]]:
        """
        Obtiene el historial completo de la conversación.
        
        Returns:
            Lista de mensajes de la conversación
        """
        return self.conversation_messages


def simple_excel_query(api_key: str, excel_path: str, prompt: str, 
                       model: str = "gpt-5", temperature: float = 1) -> Dict[str, Any]:
    """
    Función simple para hacer una consulta única sobre un archivo Excel.
    
    Args:
        api_key: Clave API de OpenAI
        excel_path: Ruta al archivo Excel
        prompt: Pregunta o instrucción sobre el archivo
        model: Modelo a utilizar
        temperature: Temperatura para la generación (0-1)
        
    Returns:
        Diccionario con la respuesta
        
    Example:
        >>> result = simple_excel_query(
        ...     "sk-...",
        ...     "datos.xlsx",
        ...     "¿Cuántas filas tiene este archivo y cuáles son las columnas?"
        ... )
        >>> print(result["response"])
    """
    processor = OpenAIExcelProcessor(api_key, model)
    return processor.query_with_excel_content(excel_path, prompt, temperature)


def extract_structured_data(api_key: str, excel_path: str, schema: Dict[str, Any], 
                           instructions: str = "", model: str = "gpt-4o") -> Dict[str, Any]:
    """
    Extrae datos estructurados de un archivo Excel según un schema JSON.
    Similar a extract_medical_data pero genérico.
    
    Args:
        api_key: Clave API de OpenAI
        excel_path: Ruta al archivo Excel
        schema: Schema JSON que debe cumplir la respuesta
        instructions: Instrucciones adicionales para la extracción
        model: Modelo a utilizar
        
    Returns:
        Diccionario con los datos extraídos o error
        
    Example:
        >>> schema = {
        ...     "type": "object",
        ...     "properties": {
        ...         "total_rows": {"type": "integer"},
        ...         "columns": {"type": "array", "items": {"type": "string"}}
        ...     }
        ... }
        >>> result = extract_structured_data(
        ...     "sk-...",
        ...     "datos.xlsx",
        ...     schema,
        ...     "Extrae el número de filas y nombres de columnas"
        ... )
    """
    processor = OpenAIExcelProcessor(api_key, model)
    
    prompt = f"""
{instructions}

Debes responder ÚNICAMENTE con un JSON válido que cumpla con el siguiente schema:
{json.dumps(schema, indent=2, ensure_ascii=False)}

No incluyas explicaciones adicionales, solo el JSON.
"""
    
    result = processor.query_with_excel_content(excel_path, prompt, temperature=1)
    
    if not result["success"]:
        return result
    
    try:
        # Limpiar markdown si está presente
        raw_text = result["response"]
        if raw_text.startswith("```json"):
            raw_text = raw_text[7:]
        if raw_text.startswith("```"):
            raw_text = raw_text[3:]
        if raw_text.endswith("```"):
            raw_text = raw_text[:-3]
        raw_text = raw_text.strip()
        
        data = json.loads(raw_text)
        
        return {
            "success": True,
            "data": data,
            "model": model
        }
        
    except json.JSONDecodeError as e:
        return {
            "success": False,
            "error": f"La respuesta no es un JSON válido: {e}",
            "raw_response": result["response"]
        }
