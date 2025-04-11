# main.py
import streamlit as st
import os
import re
import docx
import tempfile
import time
import logging
from excel_manager import excelManager # Assuming excel_manager.py is in the same directory
from keyword_parser import keywordParser # Assuming keyword_parser.py is in the same directory
from collections import Counter
from logs.logger_config import setup_logger

# Setup logger
logger = setup_logger('main')

def preprocess_word_doc(doc_path):
    """
    Analyze a Word document to determine what keywords it contains, using '!' separator.

    Args:
        doc_path: Path to the Word document

    Returns:
        Dictionary with keyword counts and whether Excel file is needed
    """
    logger.info(f"Preprocessing Word document: {doc_path}")
    doc = docx.Document(doc_path)
    pattern = r'{{(.*?)}}'

    keywords = {
        "excel": {"CELL": [], "LAST": [], "RANGE": [], "COLUMN": [], "OTHER": []},
        "input": {"text": [], "area": [], "date": [], "select": [], "check": []},
        "template": {"full": [], "section": [], "range": []},
        "json": [],
        "other": []
    }
    needs_excel = False
    total_keywords = 0

    def categorize_keyword(content):
        nonlocal needs_excel
        parts = content.split("!", 1) # Use '!' separator
        keyword_type = parts[0].strip().upper()

        if not keyword_type: return # Ignore empty keywords {{}}

        if keyword_type == "XL":
            needs_excel = True
            if len(parts) > 1:
                sub_parts = parts[1].split("!", 1)
                xl_subtype = sub_parts[0].strip().upper()
                if xl_subtype in keywords["excel"]:
                    keywords["excel"][xl_subtype].append(content)
                else:
                     # If subtype unknown, check if it looks like an old format/named range
                     if ':' not in parts[1] and '!' not in parts[1]: # Likely named range or old cell ref
                          keywords["excel"]["RANGE"].append(content) # Assume RANGE for named range
                     else:
                          keywords["excel"]["OTHER"].append(content) # Potentially old or invalid format
            else:
                 keywords["excel"]["OTHER"].append(content) # Invalid XL format {{XL}}

        elif keyword_type == "INPUT":
            if len(parts) > 1:
                input_parts = parts[1].split("!")
                input_type = input_parts[0].lower() if input_parts else "text"
                if input_type in keywords["input"]:
                    keywords["input"][input_type].append(content)
                else:
                    keywords["input"]["text"].append(content)
            else:
                 keywords["input"]["text"].append(content) # {{INPUT}} defaults to text

        elif keyword_type == "TEMPLATE":
            # Extract sections/ranges correctly
            if len(parts) <= 1:
                # No parameters, just the keyword type - should never happen for TEMPLATE but handle it
                logger.info(f"Categorizing bare TEMPLATE keyword without parameters")
                keywords["template"]["full"].append(content)
                return
                
            # Split the content after the TEMPLATE! prefix to analyze the parts
            template_parts = parts[1].split("!")
            template_path = template_parts[0]
            
            logger.info(f"Processing TEMPLATE keyword: '{content}' with path '{template_path}'")
            
            # If there's no second part with section=, it's a full template
            if len(template_parts) == 1:
                logger.info(f"Categorizing as FULL template: {content}")
                keywords["template"]["full"].append(content)
            # Check for section parameter
            elif len(template_parts) > 1 and "section=" in template_parts[1]:
                # Need to extract the section value to check for colon
                try:
                    section_param = template_parts[1].split("section=")[1].split("&")[0] if "section=" in template_parts[1] else ""
                    logger.info(f"Found section parameter: '{section_param}'")
                    
                    # Check if it's a range (contains ':') - {{TEMPLATE!filename.docx!section=Start:End}}
                    if ":" in section_param:
                        logger.info(f"Categorizing as RANGE template: {content}")
                        keywords["template"]["range"].append(content)
                    else:
                        # Just a single section - {{TEMPLATE!filename.docx!section=SectionName}}
                        logger.info(f"Categorizing as SECTION template: {content}")
                        keywords["template"]["section"].append(content)
                except Exception as e:
                    logger.error(f"Error parsing section parameter: {e}")
                    # Default to full if we can't parse the section
                    keywords["template"]["full"].append(content)
            else:
                # Any other template format defaults to full template
                logger.info(f"Categorizing as FULL template (default): {content}")
                keywords["template"]["full"].append(content)
        elif keyword_type == "JSON":
            keywords["json"].append(content)
        else:
             # If not a recognized type, check if it might be an Excel named range
             if '!' not in content and ':' not in content:
                  needs_excel = True
                  keywords["excel"]["RANGE"].append(content) # Treat as potential named range
             else:
                  keywords["other"].append(content)

    # Scan paragraphs
    for paragraph in doc.paragraphs:
        matches = list(re.finditer(pattern, paragraph.text))
        total_keywords += len(matches)
        for match in matches:
            categorize_keyword(match.group(1))

    # Scan tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    matches = list(re.finditer(pattern, paragraph.text))
                    total_keywords += len(matches)
                    for match in matches:
                        categorize_keyword(match.group(1))

    summary = {
        "total_keywords": total_keywords,
        "excel_counts": {k: len(v) for k, v in keywords["excel"].items()},
        "input_counts": {k: len(v) for k, v in keywords["input"].items()},
        "template_count": {k: len(v) for k, v in keywords["template"].items()},
        "template_total": sum(len(v) for v in keywords["template"].values()),
        "json_count": len(keywords["json"]),
        "other_count": len(keywords["other"]),
        "needs_excel": needs_excel,
        "keywords": keywords
    }
    
    # Debug log for template counts
    logger.info(f"Template summary: {summary['template_count']}")
    logger.info(f"Template total: {summary['template_total']}")
    for t_type, items in keywords["template"].items():
        if items:
            logger.info(f"Template {t_type} items: {items}")
    
    return summary


def process_word_doc(doc_path, excel_path=None, parser=None):
    """
    Process a Word document, replacing keywords with values using the provided parser.

    Args:
        doc_path: Path to the Word document
        excel_path: Path to the Excel spreadsheet (optional - manager passed via parser)
        parser: An initialized keywordParser instance

    Returns:
        Processed document object and a count of replaced keywords
    """
    logger.info(f"Starting document processing: {doc_path}")
    if excel_path:
        logger.info(f"Using Excel file: {excel_path}")
        
    if not parser:
        raise ValueError("KeywordParser instance is required.")

    doc = docx.Document(doc_path)
    parser.set_word_document(doc) # Ensure parser has the correct document object

    pattern = r'{{(.*?)}}'
    total_keywords_initial = 0

    # Count initial keywords
    elements_to_scan = list(doc.paragraphs)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                elements_to_scan.extend(cell.paragraphs)

    for paragraph in elements_to_scan:
        total_keywords_initial += len(re.findall(pattern, paragraph.text))

    logger.info(f"Found {total_keywords_initial} keywords in document")
    
    if total_keywords_initial == 0:
        st.warning("No keywords found in the document.")
        logger.warning("No keywords found in the document")
        return doc, 0

    progress_bar = st.progress(0)
    progress_text = st.empty()
    progress_text.text("Processing keywords...")

    # Process paragraph by paragraph, letting the parser handle replacements
    processed_keywords_count = 0
    elements_processed = 0
    total_elements = len(elements_to_scan)

    # Important: We'll set form_submitted to True BUT also ensure the input values
    # from the main form are properly transferred to the parser
    parser.form_submitted = True
    
    # Make sure all the input values from our main form are in the parser's input_values dict
    if 'input_values_main' in st.session_state and st.session_state.input_values_main:
        for content, value in st.session_state.input_values_main.items():
            # Store with the exact content format 
            keyword = f"{{{{{content}}}}}"
            parser.input_values[keyword] = value
            
            # Also store in alternate formats
            if content.startswith("INPUT!"):
                # Store without the INPUT! prefix
                non_prefix_content = content[6:]  # Remove "INPUT!"
                alt_keyword = f"{{{{{non_prefix_content}}}}}"
                parser.input_values[alt_keyword] = value
            else:
                # Store with the INPUT! prefix
                alt_keyword = f"{{{{INPUT!{content}}}}}"
                parser.input_values[alt_keyword] = value
    
    # Also check for any keywords directly in the form fields format
    form_keys = [k for k in st.session_state.keys() if k.startswith('input_field_')]
    if form_keys:
        for form_key in form_keys:
            if form_key.startswith('input_field_INPUT!'):
                content = form_key[12:]  # Remove 'input_field_' prefix
                value = st.session_state[form_key]
                keyword = f"{{{{{content}}}}}"
                parser.input_values[keyword] = value
    
    # Store input values for potential troubleshooting
    st.session_state['debug_input_values'] = parser.input_values.copy()

    for paragraph in elements_to_scan:
        original_text = paragraph.text
        keywords_in_para = len(re.findall(pattern, original_text))

        if keywords_in_para > 0:
            try:
                # parser.parse will handle replacements, including potential table creation
                parsed_result = parser.parse(original_text)

                # Check if we got a dict with a docx template
                if isinstance(parsed_result, dict) and "docx_template" in parsed_result:
                    try:
                        # Update the paragraph's text with any text content
                        paragraph.text = parsed_result["text"]
                        
                        # Get the template path from our result
                        template_path = parsed_result["docx_template"]
                        
                        # Load the template document with proper formatting
                        template_doc = docx.Document(template_path)
                        
                        # Insert the template document at the current paragraph location
                        paragraph_element = paragraph._element
                        paragraph_parent = paragraph_element.getparent()
                        paragraph_index = paragraph_parent.index(paragraph_element)
                        
                        # For each paragraph in the template, add it to the main document
                        for p in template_doc.paragraphs:
                            # Create a new paragraph in the main document
                            new_p = doc.add_paragraph()
                            # Copy over the paragraph's runs with their formatting
                            for run in p.runs:
                                new_run = new_p.add_run(run.text)
                                # Copy formatting from the original run
                                new_run.bold = run.bold
                                new_run.italic = run.italic
                                new_run.underline = run.underline
                                if run.font.size:
                                    new_run.font.size = run.font.size
                                if run.font.name:
                                    new_run.font.name = run.font.name
                                if run.font.color.rgb:
                                    new_run.font.color.rgb = run.font.color.rgb
                            
                            # Copy paragraph formatting
                            if p.style:
                                try:
                                    new_p.style = p.style.name
                                except:
                                    pass  # Style might not exist in target document
                            new_p.paragraph_format.alignment = p.paragraph_format.alignment
                            new_p.paragraph_format.left_indent = p.paragraph_format.left_indent
                            new_p.paragraph_format.right_indent = p.paragraph_format.right_indent
                            new_p.paragraph_format.space_before = p.paragraph_format.space_before
                            new_p.paragraph_format.space_after = p.paragraph_format.space_after
                            
                            # Position the new paragraph in the right place
                            new_p_element = new_p._element
                            # Remove it from wherever it was added automatically
                            new_p_element.getparent().remove(new_p_element)
                            # Insert at the right position
                            paragraph_index += 1
                            paragraph_parent.insert(paragraph_index, new_p_element)
                        
                        # For each table in the template, add it to the main document
                        for table in template_doc.tables:
                            # Create a new table in the main document with the same dimensions
                            rows = len(table.rows)
                            cols = len(table.columns)
                            new_table = doc.add_table(rows=rows, cols=cols)
                            
                            # Copy table style if it exists
                            if table.style:
                                try:
                                    new_table.style = table.style
                                except:
                                    pass  # Style might not exist in target document
                            
                            # Copy cell by cell
                            for i, row in enumerate(table.rows):
                                for j, cell in enumerate(row.cells):
                                    target_cell = new_table.cell(i, j)
                                    # Clear any default paragraph
                                    for p in target_cell.paragraphs:
                                        p._element.getparent().remove(p._element)
                                    
                                    # Copy each paragraph from source cell
                                    for p in cell.paragraphs:
                                        cell_p = target_cell.add_paragraph()
                                        for run in p.runs:
                                            cell_run = cell_p.add_run(run.text)
                                            cell_run.bold = run.bold
                                            cell_run.italic = run.italic
                                            cell_run.underline = run.underline
                                            if run.font.size:
                                                cell_run.font.size = run.font.size
                                            if run.font.name:
                                                cell_run.font.name = run.font.name
                                            if run.font.color.rgb:
                                                cell_run.font.color.rgb = run.font.color.rgb
                                            
                                        if p.style:
                                            try:
                                                cell_p.style = p.style.name
                                            except:
                                                pass
                                        cell_p.paragraph_format.alignment = p.paragraph_format.alignment
                            
                            # Position the new table in the right place
                            new_table_element = new_table._element
                            # Remove it from wherever it was added automatically
                            new_table_element.getparent().remove(new_table_element)
                            # Insert at the right position
                            paragraph_index += 1
                            paragraph_parent.insert(paragraph_index, new_table_element)
                            
                            # Add a blank paragraph after the table to prevent merging
                            blank_p = doc.add_paragraph()
                            blank_p_element = blank_p._element
                            blank_p_element.getparent().remove(blank_p_element)
                            paragraph_index += 1
                            paragraph_parent.insert(paragraph_index, blank_p_element)
                        
                        # Count as processed (we processed the entire keyword that produced the template)
                        processed_keywords_count += 1
                    except Exception as e:
                        st.error(f"Error inserting template document: {str(e)}")
                        logger.error(f"Error inserting template document: {str(e)}", exc_info=True)
                        # If template insertion fails, keep original text
                        paragraph.text = original_text

                # Check if we got a dict with a table object
                elif isinstance(parsed_result, dict) and "table" in parsed_result:
                    try:
                        # Update the paragraph's text with any text content
                        paragraph.text = parsed_result["text"]
                        
                        # Get the table object from our result
                        table = parsed_result["table"]
                        
                        # Insert the table at the current paragraph location
                        # We need to get the paragraph parent and insert after this paragraph
                        paragraph_element = paragraph._element
                        paragraph_parent = paragraph_element.getparent()
                        
                        # Insert the table after the current paragraph
                        paragraph_parent.insert(paragraph_element.getparent().index(paragraph_element) + 1, 
                                              table._element)
                        
                        # Add a paragraph after the table to prevent tables from merging
                        p = doc.add_paragraph()
                        p._element.getparent().insert(paragraph_element.getparent().index(paragraph_element) + 2, p._element)
                        
                        # Count as processed (we processed the entire keyword that produced the table)
                        processed_keywords_count += 1
                    except Exception as e:
                        st.error(f"Error inserting table: {str(e)}")
                        # If table insertion fails, keep original text
                        paragraph.text = original_text
                # Handle the old-style placeholder format for backwards compatibility
                elif isinstance(parsed_result, str) and "[TABLE_INSERTED]" in parsed_result:
                    # Check if the keyword was the only content (strip spaces for check)
                    is_only_keyword = False
                    matches = list(re.finditer(pattern, original_text))
                    if len(matches) == 1 and matches[0].group(0).strip() == original_text.strip():
                         is_only_keyword = True

                    if is_only_keyword:
                         paragraph.text = "" # Clear paragraph if only table keyword was present
                    else:
                         # Remove placeholder but keep other text
                         paragraph.text = parsed_result.replace("[TABLE_INSERTED]", "").strip()
                         
                    # Count as processed
                    processed_keywords_count += 1
                elif parsed_result != original_text:
                    paragraph.text = parsed_result
                    
                    # Estimate progress - count keywords *remaining* after parse
                    keywords_remaining = len(re.findall(pattern, paragraph.text))
                    processed_in_step = keywords_in_para - keywords_remaining
                    processed_keywords_count += processed_in_step

            except Exception as e:
                st.error(f"Error processing content '{original_text[:50]}...': {str(e)}")
                # Keep original text on error

        elements_processed += 1
        progress = elements_processed / total_elements if total_elements > 0 else 0
        progress_bar.progress(progress)
        # Update text based on estimated keywords processed vs initial total
        progress_text.text(f"Processing: {processed_keywords_count}/{total_keywords_initial} keywords estimated...")


    progress_bar.progress(1.0)
    progress_text.text(f"Processing finished. Approximately {processed_keywords_count} keywords processed.")

    return doc, processed_keywords_count


def display_keyword_summary(summary):
    """Display analysis summary with updated Excel categories and template details."""
    st.write(f"Total keywords found: **{summary['total_keywords']}**")
    with st.expander("Document Analysis Summary"):
        col1, col2 = st.columns(2)

        with col1:
            st.markdown("**Excel Keywords (`XL!`)**")
            total_excel = sum(summary["excel_counts"].values())
            st.write(f"Total: {total_excel}")
            if summary["needs_excel"]:
                st.write("*Excel file required*")
            for subtype, count in summary["excel_counts"].items():
                 if count > 0: st.write(f"- {subtype}: {count}")

            st.markdown("**Input Keywords (`INPUT!`)**")
            total_inputs = sum(summary["input_counts"].values())
            st.write(f"Total: {total_inputs}")
            for input_type, count in summary["input_counts"].items():
                 if count > 0: st.write(f"- {input_type}: {count}")

        with col2:
            # Enhanced Template section with more details
            st.markdown("**Template Keywords (`TEMPLATE!`)**")
            st.write(f"Total: {summary['template_total']}")
            for template_type, count in summary["template_count"].items():
                if count > 0: 
                    # Just show the uppercase name and count
                    st.write(f"- {template_type.upper()}: {count}")
            
            st.markdown("**JSON Keywords (`JSON!`)**")
            st.write(f"Total: {summary['json_count']}")
            if summary['json_count'] > 0 and 'keywords' in summary and summary['keywords']['json']:
                st.caption("Examples:")
                for item in summary['keywords']['json'][:2]:  # Show first 2
                    st.caption(f"- `{{{{{item}}}}}`")
            
            st.markdown("**Other/Invalid**")
            st.write(f"Total: {summary['other_count']}")
            if summary['other_count'] > 0 and 'keywords' in summary and summary['keywords']['other']:
                st.caption("Examples:")
                for item in summary['keywords']['other'][:2]:  # Show first 2
                    st.caption(f"- `{{{{{item}}}}}`")


def main():
    # Load custom CSS
    with open('style.css') as f:
        st.markdown(f'<style>{f.read()}</style>', unsafe_allow_html=True)
    
    logger.info("Application started")
    
    # Initialize parser instance for help text display
    if 'keyword_parser_instance_for_help' not in st.session_state:
        st.session_state.keyword_parser_instance_for_help = keywordParser()
    
    # --- State Management ---
    # Initialize state variables if they don't exist
    default_state = {
        'current_step': 1,  # Track the current wizard step
        'doc_uploaded': False, 'doc_path': None, 'analysis_summary': None,
        'excel_uploaded': False, 'excel_path': None, 'excel_manager_instance': None,
        'keyword_parser_instance': None, 'form_submitted_main': False, 'input_values_main': {},
        'processing_started': False, 'processed_doc_path': None, 'processed_count': 0
    }
    for key, value in default_state.items():
        if key not in st.session_state:
            st.session_state[key] = value
    
    # Sidebar with keyword reference guide and reset button
    with st.sidebar:
        # Load and display the Form Filler logo
        st.image("assets/images/form_filler_logo.png", width=250)
        
        st.subheader("Navigation")
        
        # Add visual step indicator
        st.markdown("""
        <div class="step-indicator">
            <div class="step {0}">1</div>
            <div class="step-line"></div>
            <div class="step {1}">2</div>
            <div class="step-line"></div>
            <div class="step {2}">3</div>
            <div class="step-line"></div>
            <div class="step {3}">4</div>
            <div class="step-line"></div>
            <div class="step {4}">5</div>
        </div>
        """.format(
            "active" if st.session_state.current_step == 1 else "",
            "active" if st.session_state.current_step == 2 else "",
            "active" if st.session_state.current_step == 3 else "",
            "active" if st.session_state.current_step == 4 else "",
            "active" if st.session_state.current_step == 5 else ""
        ), unsafe_allow_html=True)
        
        # Step indicator
        st.write("Current step: ", st.session_state.current_step)
        
        # Add Navigation buttons
        if st.session_state.current_step > 1:
            if st.button("← Previous Step"):
                st.session_state.current_step -= 1
                st.rerun()
        
        # Only show Next button if it makes sense for the current step
        can_proceed = False
        if st.session_state.current_step == 1 and st.session_state.doc_uploaded:
            can_proceed = True
        elif st.session_state.current_step == 2:
            needs_excel = st.session_state.analysis_summary and st.session_state.analysis_summary["needs_excel"]
            can_proceed = (not needs_excel) or st.session_state.excel_uploaded
        elif st.session_state.current_step == 3:
            has_inputs = st.session_state.analysis_summary and sum(st.session_state.analysis_summary['input_counts'].values()) > 0
            can_proceed = (not has_inputs) or st.session_state.form_submitted_main
        
        if can_proceed and st.session_state.current_step < 5:
            if st.button("Next Step →"):
                st.session_state.current_step += 1
                st.rerun()

        st.subheader("Keyword Reference Guides")
        # Keyword reference guide in expandable section
        with st.expander("Excel Keyword Reference Guide", expanded=False):
            st.markdown(st.session_state.keyword_parser_instance_for_help.get_excel_keyword_help())
        with st.expander("Input Keyword Reference Guide", expanded=False):
            st.markdown(st.session_state.keyword_parser_instance_for_help.get_input_keyword_help())
        with st.expander("Template Keyword Reference Guide", expanded=False):
            st.markdown(st.session_state.keyword_parser_instance_for_help.get_template_keyword_help())
        with st.expander("JSON Keyword Reference Guide", expanded=False):
            st.markdown(st.session_state.keyword_parser_instance_for_help.get_json_keyword_help())
        with st.expander("AI Keyword Reference Guide", expanded=False):
            st.markdown(st.session_state.keyword_parser_instance_for_help.get_ai_keyword_help())

        # Reset button
        if st.button("Reset Application"):
            logger.info("Resetting application state")
            # Clean up temp files
            if st.session_state.doc_path and os.path.exists(st.session_state.doc_path): 
                os.unlink(st.session_state.doc_path)
                logger.info(f"Removed temporary document: {st.session_state.doc_path}")
            if st.session_state.excel_path and os.path.exists(st.session_state.excel_path): 
                os.unlink(st.session_state.excel_path)
                logger.info(f"Removed temporary Excel file: {st.session_state.excel_path}")
            if st.session_state.processed_doc_path and os.path.exists(st.session_state.processed_doc_path): 
                os.unlink(st.session_state.processed_doc_path)
                logger.info(f"Removed processed document: {st.session_state.processed_doc_path}")
            # Close Excel Manager if open
            if st.session_state.excel_manager_instance: st.session_state.excel_manager_instance.close()
            # Reset state variables
            for key in default_state:
                st.session_state[key] = default_state[key]
            st.rerun()
    
    # Main content area - only show the current step
    # --- Step 1: Upload Word Document ---
    if st.session_state.current_step == 1:
        st.header("Step 1: Upload Document")
        st.write("Upload a Word document containing keywords that you want to process. The document should include keywords in double curly braces like `{{keyword}}`.")
        
        doc_file = st.file_uploader("Upload Word Document (.docx)", type=["docx"], key="main_doc_uploader")

        if doc_file and not st.session_state.doc_uploaded:
            # Reset relevant states for new upload
            st.session_state.update({k: v for k, v in default_state.items() 
                                   if k not in ['keyword_parser_instance_for_help', 'current_step']})
            st.session_state.current_step = 1  # Stay on step 1
            
            # Save new doc
            with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as tmp_doc:
                tmp_doc.write(doc_file.getvalue())
                st.session_state.doc_path = tmp_doc.name
            st.session_state.doc_uploaded = True
            st.rerun()
        
    
    # --- Step 2: Analysis & Excel Upload (if needed) ---
    elif st.session_state.current_step == 2:
        st.header("Step 2: Document Analysis & Required Files")
        st.write("This step analyzes your document to identify keywords and determines if additional files (like Excel) are needed.")
        
        # First run analysis if needed
        if not st.session_state.analysis_summary:
            with st.spinner("Analyzing document..."):
                try:
                    summary = preprocess_word_doc(st.session_state.doc_path)
                    st.session_state.analysis_summary = summary
                    st.rerun()
                except Exception as e:
                    st.error(f"Analysis failed: {e}")
                    st.session_state.doc_uploaded = False  # Allow re-upload
                    st.session_state.current_step = 1  # Go back to step 1
                    st.rerun()
        
        # Display analysis results
        if st.session_state.analysis_summary:
            display_keyword_summary(st.session_state.analysis_summary)
            needs_excel = st.session_state.analysis_summary["needs_excel"]
            
            # Only show Excel uploader if needed based on analysis
            if needs_excel:
                st.write("Based on the analysis, an Excel file is required.")
                excel_file = st.file_uploader("Upload Required Excel Spreadsheet (.xlsx)", 
                                            type=["xlsx"], key="main_excel_uploader")
                
                if excel_file and not st.session_state.excel_uploaded:
                    # Save new excel file
                    if st.session_state.excel_path and os.path.exists(st.session_state.excel_path): 
                        os.unlink(st.session_state.excel_path)  # Clean old temp excel
                    
                    with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_excel:
                        tmp_excel.write(excel_file.getvalue())
                        st.session_state.excel_path = tmp_excel.name
                    
                    st.session_state.excel_uploaded = True
                    
                    # Reset excel manager instance as file changed
                    if st.session_state.excel_manager_instance: 
                        st.session_state.excel_manager_instance.close()
                    st.session_state.excel_manager_instance = None
                    st.rerun()
                
            else:
                st.success("No Excel file required. You can proceed to the next step.")
            
            # Initialize Managers
            if needs_excel and st.session_state.excel_path and not st.session_state.excel_manager_instance:
                try:
                    with st.spinner("Loading Excel data..."):
                        st.session_state.excel_manager_instance = excelManager(st.session_state.excel_path)
#                        st.success("Excel data loaded successfully!")
                except Exception as e:
                    st.error(f"Failed to load Excel file: {e}")
                    st.session_state.excel_uploaded = False  # Reset upload status
                    st.rerun()
            
            # Always ensure parser instance exists, update if excel manager changes
            current_excel_manager = st.session_state.excel_manager_instance if needs_excel else None
            if not st.session_state.keyword_parser_instance or st.session_state.keyword_parser_instance.excel_manager != current_excel_manager:
                st.session_state.keyword_parser_instance = keywordParser(current_excel_manager)
    
    # --- Step 3: User Input Form (if needed) ---
    elif st.session_state.current_step == 3:
        st.header("Step 3: Provide Input Values")
        st.write("Fill in values for the input fields found in your document. These values will replace the corresponding keywords during processing.")
        
        # Check if inputs are needed
        has_inputs = st.session_state.analysis_summary and sum(st.session_state.analysis_summary['input_counts'].values()) > 0
        
        if not has_inputs:
            st.success("No user inputs required.")
            #st.info("Click 'Next Step →' in the sidebar to continue to processing.")
        else:
            st.write("Please provide values for the keywords found in your document.")
            
            if not st.session_state.form_submitted_main:
                parser = st.session_state.keyword_parser_instance
                
                with st.form(key="main_input_form"):
                    # Use analysis summary to find all input keywords
                    all_input_keywords = [item for sublist in st.session_state.analysis_summary['keywords']['input'].values() 
                                         for item in sublist]
                    unique_input_contents = sorted(list(set(all_input_keywords)))  # Get unique input definitions
                    
                    # Store fields in local state for this form
                    temp_inputs = {}
                    
                    # Disable the parser's internal form handling to prevent duplication
                    parser.form_submitted = True
                    
                    for content in unique_input_contents:
                        # Create field using parser's helper function
                        field_key = f"input_field_{content}"
                        temp_inputs[content] = parser._create_input_field(content)
                    
                    submitted = st.form_submit_button("Submit Inputs")
                    if submitted:
                        # Store values in session state
                        for content in unique_input_contents:
                            field_key = f"input_field_{content}"
                            # Extract field values from session state
                            if field_key in st.session_state:
                                field_value = st.session_state[field_key]
                                # Ensure we're storing the content exactly as it appears in the document
                                st.session_state.input_values_main[content] = field_value
                        
                        # Update the parser's internal values - ensure we use the full keyword format
                        for content, value in st.session_state.input_values_main.items():
                            # Store using the full format with INPUT!
                            keyword = f"{{{{{content}}}}}"
                            parser.input_values[keyword] = value
                            
                            # Also store in alternate formats to maximize chances of matching
                            if content.startswith("INPUT!"):
                                # Also store without the INPUT! prefix
                                non_prefix_content = content[6:]  # Remove "INPUT!"
                                alt_keyword = f"{{{{{non_prefix_content}}}}}"
                                parser.input_values[alt_keyword] = value
                            else:
                                # Also store with the INPUT! prefix
                                alt_keyword = f"{{{{INPUT!{content}}}}}"
                                parser.input_values[alt_keyword] = value
                        
                        st.session_state.form_submitted_main = True
                        logger.info("Form inputs submitted")
                        st.rerun()
            else:
                st.success("Input values submitted successfully!")
                #st.info("Click 'Next Step →' in the sidebar to continue to processing.")
    
    # --- Step 4: Process Document ---
    elif st.session_state.current_step == 4:
        st.header("Step 4: Process Document")
        st.write("Now the system will replace all keywords in your document with their corresponding values from Excel, user inputs, and templates.")
        
        # Determine if ready to process
        needs_excel = st.session_state.analysis_summary and st.session_state.analysis_summary["needs_excel"]
        has_inputs = st.session_state.analysis_summary and sum(st.session_state.analysis_summary['input_counts'].values()) > 0
        
        ready_to_process = st.session_state.doc_uploaded and \
                          (not needs_excel or st.session_state.excel_uploaded) and \
                          (not has_inputs or st.session_state.form_submitted_main)
        
        process_button_disabled = not ready_to_process or st.session_state.processing_started
        
        if not process_button_disabled:
            if st.button("Process Document Now", key="main_process_btn"):
                st.session_state.processing_started = True
                st.session_state.processed_doc_path = None  # Clear previous
                
                with st.spinner("Processing document... This may take a moment."):
                    try:
                        # Ensure parser has the submitted inputs
                        parser = st.session_state.keyword_parser_instance
                        
                        # Force the parser to use our input values, not its internal form
                        parser.form_submitted = True
                        
                        # Process the document
                        progress_bar = st.progress(0)
                        progress_text = st.empty()
                        progress_text.text("Processing keywords...")
                        
                        processed_doc, count = process_word_doc(
                            st.session_state.doc_path,
                            st.session_state.excel_path,
                            parser=parser
                        )
                        
                        if processed_doc:
                            tmp_folder = "tmp"
                            if not os.path.exists(tmp_folder): os.makedirs(tmp_folder)
                            # Use original filename for output
                            base_name = os.path.basename(st.session_state.doc_path)
                            output_filename = f"processed_{base_name}" if not base_name.startswith("tmp") else "processed_document.docx"
                            output_path = os.path.join(tmp_folder, output_filename)
                            
                            processed_doc.save(output_path)
                            st.session_state.processed_doc_path = output_path
                            st.session_state.processed_count = count
                            st.success(f"Processing Complete! Approximately {count} keywords processed.")
                            logger.info(f"Document processed successfully. Saved to {output_path}. {count} keywords processed.")
                            
                            # Automatically move to the download step
                            st.session_state.current_step = 5
                            st.rerun()
                        else:
                            st.warning("Processing did not return a document.")
                    
                    except Exception as e:
                        st.error(f"Error during processing: {e}")
                        logger.error(f"Processing error: {str(e)}", exc_info=True)
                    finally:
                        # Close excel manager instance if it exists
                        if st.session_state.excel_manager_instance:
                            st.session_state.excel_manager_instance.close()
                            st.session_state.excel_manager_instance = None
                        st.session_state.processing_started = False  # Reset processing flag
        else:
            if not ready_to_process:
                st.warning("Please complete the previous steps before processing.")
            elif st.session_state.processing_started:
                st.info("Processing is currently in progress...")
    
    # --- Step 5: Download ---
    elif st.session_state.current_step == 5:
        st.header("Step 5: Download Result")
        st.write("Your document has been processed successfully! You can now download the final document with all keywords replaced.")
        
        if st.session_state.processed_doc_path:
            st.success(f"Document processed successfully! {st.session_state.processed_count} keywords were replaced.")
            st.write("Your document is ready to download.")
            
            try:
                with open(st.session_state.processed_doc_path, "rb") as fp:
                    st.download_button(
                        label="📥 Download Processed Document",
                        data=fp,
                        file_name=os.path.basename(st.session_state.processed_doc_path),
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
                    
                    st.write("You can also:")
                    if st.button("Start Over with a New Document"):
                        # Reset to initial state but keep the help parser
                        if st.session_state.doc_path and os.path.exists(st.session_state.doc_path): 
                            os.unlink(st.session_state.doc_path)
                        if st.session_state.excel_path and os.path.exists(st.session_state.excel_path): 
                            os.unlink(st.session_state.excel_path)
                        if st.session_state.processed_doc_path and os.path.exists(st.session_state.processed_doc_path):
                            os.unlink(st.session_state.processed_doc_path)
                        # Reset state except keyword_parser_instance_for_help
                        parser_for_help = st.session_state.keyword_parser_instance_for_help
                        for key in default_state:
                            st.session_state[key] = default_state[key]
                        st.session_state.keyword_parser_instance_for_help = parser_for_help
                        st.rerun()
            except FileNotFoundError:
                st.error("Processed file not found. Please try processing again.")
                st.session_state.processed_doc_path = None  # Reset path
                st.session_state.current_step = 4  # Go back to processing step
                st.rerun()
        else:
            st.error("No processed document available. Please go back to the processing step.")
            if st.button("Return to Processing Step"):
                st.session_state.current_step = 4
                st.rerun()


if __name__ == "__main__":
    main()