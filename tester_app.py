import streamlit as st
import os
import pandas as pd
import tempfile
import json
import logging
from excel_manager import excelManager
from keyword_parser import keywordParser
from AppLogger import logger
from pathlib import Path

# Setup logger
# logger = setup_logger('tester_app')

def check_openai_api_key() -> bool:
    """
    Check if the OpenAI API key is set in the session state or in the .streamlit/secrets.toml file.
    
    Returns:
        bool: True if the API key is set, False otherwise
    """
    # First check if API key is in session state
    if 'openai_api_key' in st.session_state and st.session_state['openai_api_key']:
        logger.info("OpenAI API key found in session state")
        return True
        
    # If not in session state, check secrets.toml
    secrets_path = Path(".streamlit/secrets.toml")
    
    if not secrets_path.exists():
        logger.debug("Secrets file not found: .streamlit/secrets.toml")
        return False
        
    try:
        # Try to read the entire file content and log it for debugging (without the full key)
        with open(secrets_path, 'r', encoding='utf-8') as file:
            content = file.read()
            if "openai_api_key" in content:
                # Log a sanitized version of what we found (first 10 chars of key)
                key_start = content.find("openai_api_key")
                logger.info(f"Found openai_api_key entry in secrets.toml at position {key_start}")
            else:
                logger.warning("No openai_api_key entry found in secrets.toml content")
                
        # Now parse the file line by line to extract the key
        with open(secrets_path, 'r', encoding='utf-8') as file:
            for line in file:
                line = line.strip()
                if line.startswith('openai_api_key'):
                    parts = line.split('=', 1)
                    if len(parts) == 2:
                        # Handle all possible quote formats: "" or ''
                        api_key = parts[1].strip()
                        
                        # Remove quotes of any kind
                        if (api_key.startswith('"') and api_key.endswith('"')) or \
                           (api_key.startswith("'") and api_key.endswith("'")):
                            api_key = api_key[1:-1]
                        
                        if api_key:
                            # Store first and last 5 chars of the key for logging
                            key_preview = f"{api_key[:5]}...{api_key[-5:]}" if len(api_key) > 10 else "..."
                            logger.info(f"Found valid OpenAI API key in secrets.toml: {key_preview}")
                            
                            # Store in session state for future use
                            st.session_state['openai_api_key'] = api_key
                            st.session_state['api_key_valid'] = True  # Mark as valid since it came from secrets
                            return True
                        else:
                            logger.warning("OpenAI API key is empty in .streamlit/secrets.toml")
                            return False
        
        logger.warning("Could not find openai_api_key entry in secrets.toml")
        return False
        
    except Exception as e:
        logger.error(f"Error reading secrets file: {str(e)}", exc_info=True)
        return False

# Function to get the API key (for use in OpenAI client)
def get_openai_api_key() -> str:
    """
    Get the OpenAI API key from session state or return empty string if not set.
    
    Returns:
        str: The OpenAI API key or empty string
    """
    return st.session_state.get('openai_api_key', '')

# Function to save API key to secrets.toml
def save_openai_api_key(api_key: str) -> bool:
    """
    Save the OpenAI API key to the .streamlit/secrets.toml file.
    
    Args:
        api_key: The OpenAI API key to save
        
    Returns:
        bool: True if the API key was saved successfully, False otherwise
    """
    secrets_path = Path(".streamlit/secrets.toml")
    
    try:
        # Create .streamlit directory if it doesn't exist
        secrets_path.parent.mkdir(exist_ok=True)
        
        # Read existing content if file exists
        content = []
        if secrets_path.exists():
            with open(secrets_path, 'r', encoding='utf-8') as file:
                content = file.readlines()
        
        # Update or add the API key
        api_key_updated = False
        for i, line in enumerate(content):
            if line.strip().startswith('openai_api_key'):
                content[i] = f'openai_api_key = "{api_key}"\n'
                api_key_updated = True
                break
        
        if not api_key_updated:
            content.append(f'openai_api_key = "{api_key}"\n')
        
        # Write back to file
        with open(secrets_path, 'w', encoding='utf-8') as file:
            file.writelines(content)
        
        logger.info("OpenAI API key saved to .streamlit/secrets.toml")
        return True
    except Exception as e:
        logger.error(f"Error saving API key: {str(e)}", exc_info=True)
        return False

st.title("Excel Manager App")

logger.info("Tester application started")

# Initialize session state for API key validation
if 'api_key_valid' not in st.session_state:
    st.session_state['api_key_valid'] = False

# Load custom CSS
with open('style.css') as f:
    st.markdown(f'<style>{f.read()}</style>', unsafe_allow_html=True)

# Sidebar for file operations
with st.sidebar:
    st.header("File Operations")
    
    # Only show file operations if API key is valid
    if st.session_state.get('api_key_valid', False):
        # File upload
        uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx", "xls"])
        if uploaded_file is not None:
            # Save uploaded file to temp directory
            file_path = os.path.join(st.session_state.temp_dir, uploaded_file.name)
            with open(file_path, "wb") as f:
                f.write(uploaded_file.getbuffer())
            
            logger.info(f"Uploaded file saved to {file_path}")
            
            # Initialize ExcelManager with the uploaded file
            st.session_state.excel_manager = excelManager(file_path)
            st.session_state.keyword_parser = keywordParser(st.session_state.excel_manager)
            st.session_state.file_path = file_path
            st.sidebar.success(f"Loaded: {uploaded_file.name}")
            logger.info(f"Excel manager initialized with {uploaded_file.name}")

        # Create new file
        new_file_name = st.text_input("Or create a new file (name.xlsx):")
        if st.button("Create New File") and new_file_name:
            if not new_file_name.endswith(('.xlsx', '.xls')):
                new_file_name += '.xlsx'
            
            file_path = os.path.join(st.session_state.temp_dir, new_file_name)
            logger.info(f"Creating new Excel file at {file_path}")
            st.session_state.excel_manager = excelManager()
            st.session_state.excel_manager.create_workbook(file_path)
            st.session_state.keyword_parser = keywordParser(st.session_state.excel_manager)
            st.session_state.file_path = file_path
            st.sidebar.success(f"Created: {new_file_name}")
            logger.info(f"Created new Excel file: {new_file_name}")

        # Reset app
        if st.button("Reset"):
            logger.info("Resetting tester application")
            reset_app()
            st.sidebar.success("Reset complete")
            logger.info("Reset complete")

# Main content area - first check the API key
# Check for OpenAI API key before showing app content
api_key_set = check_openai_api_key()

if not api_key_set or not st.session_state.get('api_key_valid', False):
    st.header("OpenAI API Key Required")
    st.warning("An OpenAI API key is required to use AI features in this application.")
    st.info("Your key will be saved to .streamlit/secrets.toml so you only need to enter it once.")
    
    with st.form("api_key_form"):
        api_key = st.text_input("OpenAI API Key", type="password", 
                                help="Your key will be securely saved for future sessions")
        submitted = st.form_submit_button("Validate and Save API Key")
        
        if submitted:
            if api_key:
                # Store API key in session state
                st.session_state['openai_api_key'] = api_key
                
                # Validate the API key by trying to use it
                try:
                    from openai import OpenAI
                    
                    # Create client with the provided API key
                    client = OpenAI(api_key=api_key)
                    
                    # Simple validation request
                    response = client.chat.completions.create(
                        model="gpt-3.5-turbo",
                        messages=[{"role": "user", "content": "Hello"}],
                        max_tokens=5
                    )
                    
                    # If we get here, the API key is valid
                    st.session_state['api_key_valid'] = True
                    
                    # Save the API key to secrets.toml
                    if save_openai_api_key(api_key):
                        st.success("API key validated and saved to .streamlit/secrets.toml!")
                    else:
                        st.warning("API key validated but could not be saved to .streamlit/secrets.toml. It will be stored for this session only.")
                    
                    logger.info("API key validated successfully")
                    st.rerun()
                except Exception as e:
                    logger.error(f"Error validating API key: {str(e)}")
                    st.error(f"Invalid API key. Please check your key and try again. Error: {str(e)}")
                    st.session_state['api_key_valid'] = False
            else:
                st.error("Please enter a valid API key.")
    
    # Stop further execution until a valid API key is provided
    st.stop()

# Initialize session state (only if API key is valid)
if 'excel_manager' not in st.session_state:
    st.session_state.excel_manager = None
if 'keyword_parser' not in st.session_state:
    st.session_state.keyword_parser = None
if 'file_path' not in st.session_state:
    st.session_state.file_path = None
if 'temp_dir' not in st.session_state:
    st.session_state.temp_dir = tempfile.mkdtemp()

# Function to reset the app
def reset_app():
    st.session_state.excel_manager = None
    st.session_state.keyword_parser = None
    st.session_state.file_path = None

# Main content
if st.session_state.excel_manager is not None:
    st.subheader("Excel File Management")
    
    # Tabs for different operations
    tab1, tab2, tab3, tab4, tab5 = st.tabs(["Sheets", "Read", "Write", "Delete", "Keywords"])
    
    with tab1:
        st.subheader("Sheet Operations")
        
        # Count sheets
        if st.button("Count Sheets"):
            count = st.session_state.excel_manager.count_sheets()
            st.info(f"Number of sheets: {count}")
        
        # Get sheet names
        if st.button("Get Sheet Names"):
            names = st.session_state.excel_manager.get_sheet_names()
            st.info(f"Sheet names: {', '.join(names)}")
        
        # Create new sheet
        new_sheet_name = st.text_input("New sheet name:")
        if st.button("Create Sheet") and new_sheet_name:
            st.session_state.excel_manager.create_sheet(new_sheet_name)
            st.success(f"Created sheet: {new_sheet_name}")
            st.session_state.excel_manager.save()
    
    with tab2:
        st.subheader("Read Operations")
        
        # Select sheet
        if st.session_state.excel_manager:
            sheet_names = st.session_state.excel_manager.get_sheet_names()
            selected_sheet = st.selectbox("Select sheet", sheet_names)
            
            # Read cell (using cell reference)
            st.subheader("Read Cell")
            cell_reference = st.text_input("Cell Reference (e.g. A1, B5):", "A1")
            
            if st.button("Read Cell"):
                try:
                    value = st.session_state.excel_manager.read_cell(selected_sheet, cell_reference)
                    st.info(f"Cell value: {value}")
                except Exception as e:
                    st.error(f"Error reading cell: {str(e)}")
            
            # Read range
            st.subheader("Read Range")
            range_reference = st.text_input("Range Reference (e.g. A1:C5):", "A1:B5")
            
            if st.button("Read Range"):
                try:
                    values = st.session_state.excel_manager.read_range(selected_sheet, range_reference)
                    # Convert to pandas DataFrame for better display
                    df = pd.DataFrame(values)
                    st.dataframe(df)
                except Exception as e:
                    st.error(f"Error reading range: {str(e)}")
            
            # Read total (new functionality)
            st.subheader("Read Total")
            total_start_reference = st.text_input("Starting Cell (e.g. A1, F25):", "A1", key="total_start_ref")
            
            if st.button("Find Total"):
                try:
                    total_value = st.session_state.excel_manager.read_total(selected_sheet, total_start_reference)
                    if total_value is not None:
                        st.info(f"Total value: {total_value}")
                    else:
                        st.warning("No total value found in this column.")
                except Exception as e:
                    st.error(f"Error finding total: {str(e)}")
            
            # Read items (new functionality)
            st.subheader("Read Items")
            items_start_reference = st.text_input("Starting Cell (e.g. A1, F25):", "A1", key="items_start_ref")
            offset_value = st.number_input("Offset (rows to exclude from end):", min_value=0, value=0, key="offset_value")
            
            if st.button("Find Items"):
                try:
                    items = st.session_state.excel_manager.read_items(selected_sheet, items_start_reference, offset=offset_value)
                    if items:
                        st.info(f"Found {len(items)} items:")
                        # Display items as a dataframe for better formatting
                        df = pd.DataFrame({"Items": items})
                        st.dataframe(df)
                    else:
                        st.warning("No items found starting from this cell.")
                except Exception as e:
                    st.error(f"Error finding items: {str(e)}")
    
    with tab3:
        st.subheader("Write Operations")
        
        # Select sheet
        if st.session_state.excel_manager:
            sheet_names = st.session_state.excel_manager.get_sheet_names()
            selected_sheet = st.selectbox("Select sheet", sheet_names, key="write_sheet")
            
            # Write cell (using cell reference)
            st.subheader("Write Cell")
            cell_reference = st.text_input("Cell Reference (e.g. A1, B5):", "A1", key="write_cell_ref")
            write_value = st.text_input("Value:", key="write_value")
            
            if st.button("Write Cell"):
                try:
                    st.session_state.excel_manager.write_cell(selected_sheet, cell_reference, write_value)
                    st.success(f"Wrote '{write_value}' to cell {cell_reference}")
                    st.session_state.excel_manager.save()
                except Exception as e:
                    st.error(f"Error writing cell: {str(e)}")
            
            # Write range (using CSV input)
            st.subheader("Write Range")
            start_cell = st.text_input("Start Cell (e.g. A1):", "A1", key="range_start_cell")
            
            csv_data = st.text_area(
                "Enter CSV data (comma-separated values, one row per line):",
                "1,2,3\n4,5,6\n7,8,9"
            )
            
            if st.button("Write Range"):
                try:
                    # Parse CSV data
                    rows = []
                    for line in csv_data.strip().split("\n"):
                        values = line.split(",")
                        rows.append(values)
                    
                    st.session_state.excel_manager.write_range(selected_sheet, start_cell, rows)
                    st.success(f"Wrote data to range starting at {start_cell}")
                    st.session_state.excel_manager.save()
                except Exception as e:
                    st.error(f"Error writing range: {str(e)}")
    
    with tab4:
        st.subheader("Delete Operations")
        
        # Delete sheet
        if st.session_state.excel_manager:
            sheet_names = st.session_state.excel_manager.get_sheet_names()
            sheet_to_delete = st.selectbox("Select sheet to delete", sheet_names)
            
            if st.button("Delete Sheet") and len(sheet_names) > 1:
                st.session_state.excel_manager.delete_sheet(sheet_to_delete)
                st.success(f"Deleted sheet: {sheet_to_delete}")
                st.session_state.excel_manager.save()
            elif len(sheet_names) <= 1:
                st.error("Cannot delete the only sheet in the workbook.")
    
    with tab5:
        st.subheader("Keyword Parser")
        
        # Show help information
        if st.session_state.keyword_parser:
            with st.expander("Excel Keyword Reference Guide", expanded=False):
                st.markdown(st.session_state.keyword_parser.get_excel_keyword_help())
            with st.expander("Input Keyword Reference Guide", expanded=False):
                st.markdown(st.session_state.keyword_parser.get_input_keyword_help())
            with st.expander("Template Keyword Reference Guide", expanded=False):
                st.markdown(st.session_state.keyword_parser.get_template_keyword_help())
            with st.expander("JSON Keyword Reference Guide", expanded=False):
                st.markdown(st.session_state.keyword_parser.get_json_keyword_help())
            with st.expander("AI Keyword Reference Guide", expanded=False):
                st.markdown(st.session_state.keyword_parser.get_ai_keyword_help())
        
        # Input for keyword string
        st.subheader("Parse Keywords")
        keyword_input = st.text_area(
            "Enter text with keywords to parse:",
            "Hello, the value in cell A1 is {{XL:A1}}."
        )
        
        # Clear input cache option
        if st.button("Clear Input Cache"):
            if st.session_state.keyword_parser:
                st.session_state.keyword_parser.clear_input_cache()
                st.success("Input cache cleared")
        
        # Parse button
        if st.button("Parse Keywords"):
            if st.session_state.keyword_parser:
                try:
                    # Reset form state each time Parse is clicked
                    st.session_state.keyword_parser.reset_form_state()
                    
                    st.subheader("Result:")
                    result = st.session_state.keyword_parser.parse(keyword_input)
                    st.write(result)
                except Exception as e:
                    st.error(f"Error parsing keywords: {str(e)}")
            else:
                st.error("Keyword parser not initialized")
    
    # Download the file
    if st.session_state.file_path:
        with open(st.session_state.file_path, "rb") as file:
            file_name = os.path.basename(st.session_state.file_path)
            st.download_button(
                label="Download Excel file",
                data=file,
                file_name=file_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
else:
    st.info("Please upload an Excel file or create a new one to start.")