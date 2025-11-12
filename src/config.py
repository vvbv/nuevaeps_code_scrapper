"""
Configuration loader
====================
Carga la configuración desde variables de entorno y archivo .env
"""

import os
from pathlib import Path
from typing import Optional


def load_api_key() -> Optional[str]:
    """
    Carga la API key de OpenAI desde múltiples fuentes (en orden de prioridad):
    1. Variable de entorno OPENAI_API_KEY
    2. Variable de entorno API_KEY
    3. Archivo .env en el directorio src
    4. Archivo .env en el directorio raíz del proyecto
    
    Returns:
        API key si se encuentra, None en caso contrario
    """
    # Primero intentar desde variables de entorno
    api_key = os.getenv("OPENAI_API_KEY") or os.getenv("API_KEY")
    
    if api_key:
        return api_key
    
    # Intentar cargar desde archivo .env
    try:
        from dotenv import load_dotenv
        
        # Buscar .env en el directorio src
        src_dir = Path(__file__).parent
        env_file = src_dir / ".env"
        
        if env_file.exists():
            load_dotenv(env_file)
            api_key = os.getenv("OPENAI_API_KEY") or os.getenv("API_KEY")
            if api_key:
                return api_key
        
        # Buscar .env en el directorio raíz del proyecto
        root_dir = src_dir.parent
        env_file = root_dir / ".env"
        
        if env_file.exists():
            load_dotenv(env_file)
            api_key = os.getenv("OPENAI_API_KEY") or os.getenv("API_KEY")
            if api_key:
                return api_key
        
    except ImportError:
        # python-dotenv no está instalado, continuar sin él
        pass
    
    return None


def get_api_key(provided_key: Optional[str] = None) -> Optional[str]:
    """
    Obtiene la API key, priorizando la proporcionada explícitamente.
    
    Args:
        provided_key: API key proporcionada explícitamente (opcional)
        
    Returns:
        API key a usar, o None si no se encuentra
    """
    if provided_key:
        return provided_key
    
    return load_api_key()


def get_config(key: str, default: Optional[str] = None) -> Optional[str]:
    """
    Obtiene un valor de configuración desde variables de entorno o .env
    
    Args:
        key: Nombre de la variable
        default: Valor por defecto si no se encuentra
        
    Returns:
        Valor de la configuración o default
    """
    value = os.getenv(key)
    
    if value:
        return value
    
    # Intentar cargar desde .env
    try:
        from dotenv import load_dotenv
        
        src_dir = Path(__file__).parent
        env_file = src_dir / ".env"
        
        if env_file.exists():
            load_dotenv(env_file)
            value = os.getenv(key)
            if value:
                return value
        
        root_dir = src_dir.parent
        env_file = root_dir / ".env"
        
        if env_file.exists():
            load_dotenv(env_file)
            value = os.getenv(key)
            if value:
                return value
        
    except ImportError:
        pass
    
    return default
