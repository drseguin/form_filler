import os
import logging
from logs.logger_config import setup_logger

def ensure_templates_dir():
    """
    Ensure the templates directory exists.
    If it doesn't exist, create it.
    """
    logger = setup_logger('templates')
    templates_dir = 'templates'
    
    try:
        if not os.path.exists(templates_dir):
            os.makedirs(templates_dir)
            logger.info(f"Created templates directory: {templates_dir}")
        else:
            logger.info(f"Templates directory already exists: {templates_dir}")
    except Exception as e:
        logger.error(f"Error creating templates directory: {str(e)}")
        raise

if __name__ == "__main__":
    ensure_templates_dir() 