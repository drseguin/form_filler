#!/usr/bin/env python
import sys
import os
import logging

# Add parent directory to path for imports
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from keyword_parser import keywordParser

# Set up logging
logging.basicConfig(level=logging.INFO, 
                    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
                    handlers=[logging.StreamHandler()])

logger = logging.getLogger("test_apostrophes")

# Different apostrophes to try
APOSTROPHES = [
    "'",    # ASCII apostrophe (U+0027)
    "'",    # Right single quotation mark (U+2019)
    "′",    # Prime (U+2032)
    "`",    # Backtick (U+0060)
    "´",    # Acute accent (U+00B4)
    "‵",    # Reversed prime (U+2035)
    "ʼ",    # Modifier letter apostrophe (U+02BC)
    "ʻ",    # Modifier letter turned comma (U+02BB)
    "ʿ",    # Modifier letter right half ring (U+02BF)
    "ʾ",    # Modifier letter left half ring (U+02BE)
]

def main():
    parser = keywordParser()
    template_path = "templates/1000-islands.docx"
    
    logger.info("Testing all apostrophe types for 'Millionaires' Row'")
    
    successful_versions = []
    keep_testing = True  # Change to True to test all apostrophe types, not just stop at first success
    
    # Try with different apostrophes
    for i, apostrophe in enumerate(APOSTROPHES):
        test_section = f"Millionaires{apostrophe} Row"
        logger.info(f"Test {i+1}: '{test_section}' (apostrophe: U+{ord(apostrophe):04X})")
        
        # Process the template keyword
        template_keyword = f"1000-islands.docx!section={test_section}"
        result = parser._process_template_keyword(template_keyword)
        
        # Check the result
        success = isinstance(result, dict) and "docx_template" in result
        status = "SUCCESS" if success else "FAILED"
        logger.info(f"  {status}: {test_section}")
        
        if success:
            successful_versions.append((apostrophe, f"U+{ord(apostrophe):04X}"))
            # Optionally stop after the first success
            if not keep_testing:
                break
    
    # Report results
    if successful_versions:
        logger.info("\nThe following apostrophe versions work successfully:")
        for apostrophe, code in successful_versions:
            logger.info(f"  - '{apostrophe}' ({code})")
        
        # Show working template keyword format
        working_section = f"Millionaires{successful_versions[0][0]} Row"
        logger.info(f"\nUse this format in your template: {{{{TEMPLATE!1000-islands.docx!section={working_section}}}}}")
    else:
        logger.info("\nNone of the tested apostrophe versions worked.")
        logger.info("Try using partial matching with: {{TEMPLATE!1000-islands.docx!section=Millionaires}}")

if __name__ == "__main__":
    main() 