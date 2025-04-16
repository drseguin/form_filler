from abc import ABC, abstractmethod
import os
from pathlib import Path
from typing import Dict, Optional, Union
from AppLogger import logger
import streamlit as st


class LLMClient(ABC):
    """Abstract base class for LLM clients."""

    def __init__(self):
        self.logger = logger

    @abstractmethod
    def summarize(self, 
                  text: str, 
                  prompt: str, 
                  max_words: int = 100, 
                  temperature: float = 0.5) -> str:
        """
        Summarize the given text using an LLM.
        
        Args:
            text: The text to summarize
            prompt: Instructions for the summarization
            max_words: Maximum number of words in the summary
            temperature: Controls randomness (0.0-1.0)
            
        Returns:
            The summarized text
        """
        pass
    
    @abstractmethod
    def get_api_key(self) -> Optional[str]:
        """
        Get the API key for the LLM service.
        
        Returns:
            The API key if available, None otherwise
        """
        pass


class OpenAIClient(LLMClient):
    """OpenAI implementation of LLM client."""
    
    def __init__(self):
        super().__init__()
        self.logger.info("Initializing OpenAI client")
        self.api_key = self.get_api_key()
        
    def get_api_key(self) -> Optional[str]:
        """
        Get the OpenAI API key from session state or from .streamlit/secrets.toml
        
        Returns:
            The API key if available, None otherwise
        """
        # First try to get the API key from session state
        if 'openai_api_key' in st.session_state and st.session_state['openai_api_key']:
            self.logger.info("Using OpenAI API key from session state")
            return st.session_state['openai_api_key']
        
        # If not in session state, try to get from secrets.toml
        api_key = None
        secrets_path = Path(".streamlit/secrets.toml")
        
        if not secrets_path.exists():
            self.logger.warning("Secrets file not found: .streamlit/secrets.toml")
            return None
            
        try:
            # Parse toml file manually since toml module might not be available
            with open(secrets_path, 'r', encoding='utf-8') as file:
                for line in file:
                    if line.strip().startswith('openai_api_key'):
                        parts = line.strip().split('=', 1)
                        if len(parts) == 2:
                            api_key = parts[1].strip().strip('"\'')
                            if api_key:
                                self.logger.info("OpenAI API key found in secrets.toml")
                                # Also store in session state for future use
                                st.session_state['openai_api_key'] = api_key
                                return api_key
        except Exception as e:
            self.logger.error(f"Error reading secrets file: {str(e)}", exc_info=True)
            return None
            
        self.logger.warning("OpenAI API key not found in session state or secrets.toml")
        return None
    
    def summarize(self, 
                  text: str, 
                  prompt: str, 
                  max_words: int = 100, 
                  temperature: float = 0.5) -> str:
        """
        Summarize the given text using OpenAI.
        
        Args:
            text: The text to summarize
            prompt: Instructions for the summarization
            max_words: Maximum number of words in the summary
            temperature: Controls randomness (0.0-1.0)
            
        Returns:
            The summarized text
        """
        if not text.strip():
            return "[No text provided to summarize]"
            
        if not self.api_key:
            return "[OpenAI API key not found]"
            
        try:
            from openai import OpenAI
            
            # Create client
            client = OpenAI(api_key=self.api_key)
            
            # Create prompt with instructions
            full_prompt = f"{prompt}\n\nText to summarize (keep under {max_words} words):\n\n{text}"
            
            # Call OpenAI API
            response = client.chat.completions.create(
                model="gpt-4o",  # Use gpt-4o model
                messages=[
                    {"role": "user", "content": full_prompt}
                ],
                max_tokens=max_words * 2,  # Approximate token count
                temperature=temperature
            )
            
            summary = response.choices[0].message.content.strip()
            
            # Count words and warn if exceeded
            word_count = len(summary.split())
            if word_count > max_words:
                self.logger.warning(f"Summary exceeds word limit: {word_count} > {max_words}")
                # Truncate to the word limit
                summary_words = summary.split()
                summary = " ".join(summary_words[:max_words])
                summary += "..."
            
            return summary
        
        except ImportError:
            self.logger.error("OpenAI library not available")
            return "[Error: OpenAI library not available. Install with 'pip install openai>=1.0.0']"
        
        except Exception as e:
            self.logger.error(f"Error generating summary: {str(e)}", exc_info=True)
            return f"[Error generating summary: {str(e)}]" 