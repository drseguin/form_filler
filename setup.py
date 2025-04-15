from pathlib import Path
import subprocess
import sys

def main():
    """Script to set up the Form Filler application dependencies."""
    print("Setting up Form Filler application...")
    
    # Install required packages if they're not already installed
    try:
        import spacy
        print("✓ spaCy is already installed")
    except ImportError:
        print("Installing spaCy...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "spacy"])
        import spacy
        print("✓ spaCy installed successfully")
    
    # Download spaCy model if not already downloaded
    try:
        spacy.load("en_core_web_sm")
        print("✓ spaCy English model is already downloaded")
    except OSError:
        print("Downloading spaCy English model...")
        subprocess.check_call([sys.executable, "-m", "spacy", "download", "en_core_web_sm"])
        print("✓ spaCy English model downloaded successfully")
    
    # Create necessary directories
    config_file = Path("config.json")
    if config_file.exists():
        import json
        with open(config_file, "r", encoding="utf-8") as f:
            config = json.load(f)
            
        # Create directories defined in config.json
        for key, path in config.get("paths", {}).items():
            directory = Path(path)
            if not directory.exists():
                directory.mkdir(parents=True, exist_ok=True)
                print(f"✓ Created directory: {directory}")
    
    print("\nSetup completed successfully!")

if __name__ == "__main__":
    main() 