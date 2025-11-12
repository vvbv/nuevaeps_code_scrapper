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
    
    def __init__(self, api_key: str, model: str = "gpt-4o"):
        """
        Inicializa el procesador de Excel con OpenAI.
        
        Args:
            api_key: Clave API de OpenAI
            model: Modelo a utilizar (por defecto gpt-4o que soporta archivos)
        """
        self.api_key = api_key
        self.model = model
        self.file_id = None
        self.file_path = None
        self.conversation_messages = []
        
    def upload_excel_file(self, excel_path: str) -> str:
        """
        Sube un archivo Excel a OpenAI para su procesamiento.
        
        Args:
            excel_path: Ruta al archivo Excel
            
        Returns:
            ID del archivo subido
            
        Raises:
            FileNotFoundError: Si el archivo no existe
            Exception: Si hay error en la subida
        """
        if not os.path.exists(excel_path):
            raise FileNotFoundError(f"Archivo Excel no encontrado: {excel_path}")
        
        print(f"Subiendo archivo {excel_path} a OpenAI...")
        
        try:
            # Usando la nueva API de OpenAI
            client = openai.OpenAI(api_key=self.api_key)
            
            with open(excel_path, "rb") as f:
                file_response = client.files.create(
                    file=f,
                    purpose='assistants'
                )
            
            self.file_id = file_response.id
            self.file_path = excel_path
            
            print(f"Archivo subido exitosamente. File ID: {self.file_id}")
            return self.file_id
            
        except Exception as e:
            print(f"Error al subir archivo: {e}")
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
        if not self.file_id:
            raise ValueError("No hay archivo Excel cargado. Usa upload_excel_file() primero.")
        
        print(f"Consultando sobre el archivo...")
        
        try:
            client = openai.OpenAI(api_key=self.api_key)
            
            # Añadir el mensaje del usuario al historial
            self.conversation_messages.append({
                "role": "user",
                "content": prompt
            })
            
            # Hacer la consulta con el contexto del archivo
            response = client.chat.completions.create(
                model=self.model,
                messages=[
                    {
                        "role": "system",
                        "content": "Eres un asistente experto en analizar archivos Excel. "
                                   "Respondes de manera precisa y estructurada basándote en los datos del archivo."
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
    
    def query_with_excel_content(self, excel_path: str, prompt: str, temperature: float = 0.3) -> Dict[str, Any]:
        """
        Procesa un archivo Excel directamente enviando su contenido en base64.
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
        
        try:
            # Leer el archivo y convertir a base64
            with open(excel_path, "rb") as f:
                excel_bytes = f.read()
            
            excel_base64 = base64.b64encode(excel_bytes).decode("utf-8")
            
            # Usar la API de responses (nueva versión)
            url = "https://api.openai.com/v1/chat/completions"
            headers = {
                "Content-Type": "application/json",
                "Authorization": f"Bearer {self.api_key}"
            }
            
            payload = {
                "model": self.model,
                "temperature": temperature,
                "messages": [
                    {
                        "role": "system",
                        "content": "Eres un asistente experto en analizar archivos Excel. "
                                   "Respondes de manera precisa y estructurada basándote en los datos del archivo."
                    },
                    {
                        "role": "user",
                        "content": [
                            {
                                "type": "text",
                                "text": prompt
                            }
                        ]
                    }
                ]
            }
            
            response = requests.post(url, headers=headers, json=payload)
            result = response.json()
            
            if "error" in result:
                return {
                    "success": False,
                    "error": result["error"]
                }
            
            assistant_message = result["choices"][0]["message"]["content"]
            
            return {
                "success": True,
                "response": assistant_message,
                "model": self.model,
                "total_tokens": result["usage"]["total_tokens"]
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
                       model: str = "gpt-4o", temperature: float = 0.3) -> Dict[str, Any]:
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
    
    result = processor.query_with_excel_content(excel_path, prompt, temperature=0)
    
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
