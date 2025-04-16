import logging
import os
from pathlib import Path

class AppLogger:
    _instances = {}

    @classmethod
    def get_logger(cls, logger_name):
        """Get or create a logger instance with the given name."""
        if logger_name in cls._instances:
            return cls._instances[logger_name]
        
        # Create a new logger instance
        logger = cls._setup_logger(logger_name)
        cls._instances[logger_name] = logger
        return logger

    @staticmethod
    def _setup_logger(logger_name):
        """Setup a logger with file and console handlers."""
        # Make sure the logs directory exists
        logs_dir = Path('logs')
        logs_dir.mkdir(exist_ok=True)
        
        # Configure the logger
        logger = logging.getLogger(logger_name)
        
        # Only configure if handlers aren't already set up
        if not logger.handlers:
            logger.setLevel(logging.INFO)
            
            # Create file handler
            log_file = logs_dir / f"{logger_name}.log"
            file_handler = logging.FileHandler(log_file)
            
            # Create console handler
            console_handler = logging.StreamHandler()
            
            # Create formatter and add it to the handlers
            formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
            file_handler.setFormatter(formatter)
            console_handler.setFormatter(formatter)
            
            # Add the handlers to the logger
            logger.addHandler(file_handler)
            logger.addHandler(console_handler)
        
        return logger

# Create a simplified interface for importing
logger = AppLogger.get_logger('app') 