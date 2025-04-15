import json
import os
from pathlib import Path
from typing import Dict, Any, Optional
from logs.logger_config import setup_logger
from llm_client import LLMClient, OpenAIClient
from triton_client import TritonClient


class LLMFactory:
    """Factory class for creating LLM clients based on configuration."""
    
    def __init__(self):
        self.logger = setup_logger('llm_factory')
        self.config = self._load_config()
        
    def _load_config(self) -> Dict[str, Any]:
        """
        Load configuration from config.json.
        
        Returns:
            Dictionary containing configuration settings
        """
        default_config = {
            "llm": {
                "provider": "openai",
                "use_triton": False
            }
        }
        
        config_path = Path("config.json")
        if not config_path.exists():
            self.logger.warning("Config file not found: config.json - using defaults")
            return default_config
            
        try:
            with open(config_path, 'r', encoding='utf-8') as file:
                config = json.load(file)
                self.logger.info("Loaded config from config.json")
                return config
        except Exception as e:
            self.logger.error(f"Error loading config: {str(e)}", exc_info=True)
            return default_config
    
    def create_client(self) -> LLMClient:
        """
        Create an LLM client based on configuration.
        
        Returns:
            An instance of a class implementing LLMClient
        """
        use_triton = self.config.get("llm", {}).get("use_triton", False)
        
        if use_triton:
            self.logger.info("Triton LLM selected in config")
            return TritonClient()
        else:
            self.logger.info("OpenAI LLM selected in config")
            return OpenAIClient()


def get_llm_client() -> LLMClient:
    """
    Convenience function to get an LLM client instance.
    
    Returns:
        An instance of a class implementing LLMClient
    """
    factory = LLMFactory()
    return factory.create_client() 