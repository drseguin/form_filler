#!/usr/bin/env python3
"""
A simple script to ensure the logs directory exists.
"""
import os
from pathlib import Path

def ensure_logs_directory():
    """Create the logs directory if it doesn't exist."""
    logs_dir = Path('logs')
    logs_dir.mkdir(exist_ok=True)
    print(f"Logs directory exists at: {logs_dir.absolute()}")
    
    # Create an empty __init__.py file to mark logs as a package
    init_file = logs_dir / "__init__.py"
    if not init_file.exists():
        init_file.touch()
        print(f"Created {init_file}")
    
    return logs_dir

if __name__ == "__main__":
    logs_path = ensure_logs_directory()
    print(f"Logs will be stored in: {logs_path}") 