import os
import logging
from logs.logger_config import setup_logger

def ensure_json_dir():
    """
    Ensure the json directory exists.
    If it doesn't exist, create it.
    """
    logger = setup_logger('json')
    json_dir = 'json'
    
    try:
        if not os.path.exists(json_dir):
            os.makedirs(json_dir)
            logger.info(f"Created json directory: {json_dir}")
        else:
            logger.info(f"JSON directory already exists: {json_dir}")
    except Exception as e:
        logger.error(f"Error creating json directory: {str(e)}")
        raise

if __name__ == "__main__":
    ensure_json_dir() 