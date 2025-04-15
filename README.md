# FORM FILLER

<p align="center">
  <img src="assets/images/form_filler_logo.png" alt="FORM Filler Logo" width="200"/>
</p>
<p align="center">
  <img src="assets/images/main_screen.png" alt="Main Screen" width="800"/>
</p>

## Overview

FORM Filler is a comprehensive document processing system that dynamically replaces keywords in Word documents with content from multiple sources. The system supports five key keyword types:

1. **User Input Keywords (`{{INPUT!...}}`)**: Create interactive form fields for gathering user input.
2. **Excel Keywords (`{{XL!...}}`)**: Extract data from specific cells, ranges, or columns in Excel spreadsheets.
3. **Template Keywords (`{{TEMPLATE!...}}`)**: Import content from other document templates.
4. **JSON Keywords (`{{JSON!...}}`)**: Pull data from JSON files using JSONPath expressions.
5. **AI Keywords (`{{AI!...}}`)**: Generate AI-powered summaries of documents or sections.

This powerful templating system enables dynamic document generation with data-driven content and interactive user inputs, using `!` as the primary separator within keywords.

The application provides a user-friendly Streamlit interface that guides users through a 5-step workflow:

1. **Document Upload**: Upload a Word document containing keywords.
2. **Analysis & File Uploads**: The system analyzes the keywords, identifies what data sources are needed, and prompts for Excel upload if necessary.
3. **User Input**: If the document contains `INPUT` keywords, the app generates a form to collect user inputs.
4. **Processing**: The keywords are processed and replaced with actual data.
5. **Download**: The processed document is available for download.

## Keywords

All keywords use double curly braces (`{{}}`) as delimiters and the exclamation mark (`!`) as the separator between keyword components. Any file specified in a keyword but not found in the appropriate folder will prompt the user to upload it. The uploaded file will be stored in the appropriate folder, so files only need to be uploaded once.

### User Input Keywords (`{{INPUT!...}}`)

These keywords create interactive input fields in the Streamlit application, allowing users to provide custom data for the document.

If User Input keywords are detected in the uploaded document, the user will be prompted for input value(s) in Step 3.

| Keyword Pattern | Description | Example |
| :-------------- | :---------- | :------ |
| `{{INPUT!TEXT!label!default_value}}` | Text input with label and default value | `{{INPUT!TEXT!Your Name!John Doe}}` |
| `{{INPUT!AREA!label!default_value!height}}` | Multi-line text input with label, default value, and height | `{{INPUT!AREA!Comments!!200}}` |
| `{{INPUT!DATE!label!default_date!format}}` | Date input with label, default date, and format | `{{INPUT!DATE!Select Date!today!YYYY/MM/DD}}` |
| `{{INPUT!SELECT!label!option1,option2,option3}}` | Dropdown selection with label and options | `{{INPUT!SELECT!Choose Color!Red,Green,Blue}}` |
| `{{INPUT!CHECK!label!default_state}}` | Checkbox input with label and default state | `{{INPUT!CHECK!Agree to Terms!false}}` |

### Excel Keywords (`{{XL!...}}`)

These keywords extract data from Excel spreadsheets, supporting single cells, ranges, and formatted tables.

If Excel keywords are detected in the uploaded document, the user will be prompted to upload the required Excel file(s) in Step 2.

| Keyword Pattern | Description | Example |
| :-------------- | :---------- | :------ |
| `{{XL!excel_file.xlsx!CELL!Cell}}` | Cell value | `{{XL!budget.xlsx!CELL!A1}}` |
| `{{XL!excel_file.xlsx!CELL!Sheet!Cell}}` | Cell value from specific sheet | `{{XL!budget.xlsx!CELL!Summary!B5}}` |
| `{{XL!excel_file.xlsx!LAST!Cell}}` | Last non-empty value in column | `{{XL!budget.xlsx!LAST!A1}}` |
| `{{XL!excel_file.xlsx!LAST!Sheet!Cell}}` | Last non-empty value in sheet column | `{{XL!budget.xlsx!LAST!Summary!A1}}` |
| `{{XL!excel_file.xlsx!LAST!Sheet!Cell!Title}}` | Last non-empty value in titled column | `{{XL!budget.xlsx!LAST!Summary!A1!Total}}` |
| `{{XL!excel_file.xlsx!RANGE!StartCell:EndCell}}` | Range of cells as formatted table | `{{XL!budget.xlsx!RANGE!A1:G13}}` |
| `{{XL!excel_file.xlsx!RANGE!Sheet!StartCell:EndCell}}` | Range from specific sheet as formatted table | `{{XL!budget.xlsx!RANGE!Summary!A1:G13}}` |
| `{{XL!excel_file.xlsx!COLUMN!Sheet!Cell1,Cell2,Cell3}}` | Multiple columns by cell references as table | `{{XL!budget.xlsx!COLUMN!Support!C4,E4,J4}}` |
| `{{XL!excel_file.xlsx!COLUMN!Sheet!Title1,Title2,Title3!Row}}` | Multiple columns by titles as table | `{{XL!sales.xlsx!COLUMN!Distribution Plan!Unit,DHTC,Total!4}}` |

### Template Keywords (`{{TEMPLATE!...}}`)

These keywords import content from other document templates, allowing for modular document construction.

If Template keywords are detected in the uploaded document, the application will look for the specified template file(s) in the `templates` folder.

| Keyword Pattern | Description | Example |
| :-------------- | :---------- | :------ |
| `{{TEMPLATE!filename.docx}}` | Include full document | `{{TEMPLATE!disclaimer.txt}}` |
| `{{TEMPLATE!filename.docx!section=heading}}` | Include section with heading | `{{TEMPLATE!report.docx!section=conclusion}}` |
| `{{TEMPLATE!filename.docx!section=heading!title=false}}` | Include section without heading | `{{TEMPLATE!report.docx!section=conclusion!title=false}}` |
| `{{TEMPLATE!filename.docx!section=heading_start:heading_end}}` | Include multiple sections with headings | `{{TEMPLATE!report.docx!section=intro:conclusion}}` |
| `{{TEMPLATE!filename.docx!section=heading_start:heading_end&title=false}}` | Include multiple sections without headings | `{{TEMPLATE!report.docx!section=intro:conclusion&title=false}}` |

### JSON Keywords (`{{JSON!...}}`)

These keywords extract data from JSON files using JSONPath expressions, with additional formatting options.

If JSON keywords are detected in the uploaded document, the application will look for the specified JSON file(s) in the `json` folder or at the specified path.

| Keyword Pattern | Description | Example |
| :-------------- | :---------- | :------ |
| `{{JSON!!filename.json}}` | Full JSON content | `{{JSON!!config.json}}` |
| `{{JSON!!filename.json!$.}}` | Full JSON content (alternative) | `{{JSON!!settings.json!$.}}` |
| `{{JSON!filename.json!$.key}}` | Value at specific JSON path | `{{JSON!launch.json!$.configurations}}` |
| `{{JSON!filename.json!$.key!SUM}}` | Sum numeric values in path | `{{JSON!sales.json!$.monthly_totals!SUM}}` |
| `{{JSON!filename.json!$.key!JOIN(, )}}` | Join values with separator | `{{JSON!users.json!$.names!JOIN(, )}}` |
| `{{JSON!filename.json!$.key!BOOL(Yes/No)}}` | Transform boolean values | `{{JSON!status.json!$.system_active!BOOL(Online/Offline)}}` |

### AI Keywords (`{{AI!...}}`)

These keywords generate AI-powered summaries of document content with intelligent formatting using spaCy and OpenAI.

If AI keywords are detected in the uploaded document, the application will look for the specified document(s) in the `ai` folder or at the specified path.

| Keyword Pattern | Description | Example |
| :-------------- | :---------- | :------ |
| `{{AI!source-doc.docx!prompt_file.txt!words=100}}` | Summarize entire document | `{{AI!report.docx!Summarize this report!words=150}}` |
| `{{AI!source-doc.docx!prompt_file.txt!section=section header&words=100}}` | Summarize specific section | `{{AI!contract.docx!Create a summary!section=Legal Terms&words=75}}` |
| `{{AI!source-doc.docx!prompt_file.txt!section=Attractions:Unique Experiences&words=100}}` | Summarize content range | `{{AI!travel-guide.docx!concise highlights!section=History:Culture&words=100}}` |

AI summaries are intelligently formatted with spaCy natural language processing to improve readability:
- Automatic paragraph breaks based on content structure
- Proper formatting of sentences and sections
- Recognition of bullet points and other structural elements
- Configuration options in config.json for customizing the formatting

## Setup and Installation

1. Install the required Python packages:
```bash
pip install -r requirements.txt
```

2. Set up spaCy and download the required model:
```bash
python setup.py
```

3. Configure API keys in `.streamlit/secrets.toml`:
```toml
openai_api_key = "your-api-key-here"
```

4. Start the application:
```bash
streamlit run main.py
```

## Configuration

The application can be configured through `config.json`. Key configuration options include:

```json
{
  "llm": {
    "provider": "openai",
    "use_triton": false,
    "settings": {
      "openai": {
        "model": "gpt-4o",
        "temperature": 0.5
      }
    }
  },
  "spacy": {
    "enabled": true,
    "model": "en_core_web_sm",
    "format_entities": true,
    "paragraph_breaks": true,
    "entity_styles": {
      "PERSON": {"bold": true},
      "ORG": {"bold": true, "underline": true},
      "DATE": {"italic": true},
      "MONEY": {"bold": true},
      "PLACE": {"underline": true}
    }
  },
  "paths": {
    "templates": "templates",
    "json": "json",
    "ai": "ai"
  }
}
```

This configuration file lets you customize:
- LLM provider and settings
- spaCy language model and formatting options
- Directory paths for templates, JSON files, and AI sources

## Developer

**David Seguin** is the creator and lead developer of FORM Filler.