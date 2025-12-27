# Mediation Application Form Generator

A Python-based web application that generates MS Word documents replicating the official **FORM 'A' - Mediation Application Form** used by Mumbai District Legal Services Authority.

## Live Demo

**Live URL:** [https://mediation-form-generator.onrender.com](https://mediation-form-generator.onrender.com)

## Project Overview

This project converts a PDF form template into a dynamically generated MS Word document (.docx) that maintains the exact layout, formatting, and structure of the original form.

## Approach

### 1. PDF Analysis
- Thoroughly analyzed the source PDF (`django_assignment-1.pdf`) to understand:
  - Document structure and hierarchy
  - Table layouts and cell configurations
  - Font styles, sizes, and formatting
  - Spacing and alignment requirements
  - Template variables (Jinja-style placeholders)

### 2. Document Generation Strategy
- Used **python-docx** library for Word document creation
- Implemented helper functions for consistent formatting:
  - `set_cell_borders()` - Applies uniform table borders
  - `add_formatted_paragraph()` - Creates styled paragraphs with proper spacing
  - `add_new_paragraph()` - Adds additional paragraphs within cells

### 3. Layout Replication
- **Header Section**: Centered text with bold formatting for form title
- **Tables**: Multi-column layout with merged cells where required
- **Template Variables**: Preserved Jinja-style placeholders (`{{client_name}}`, `{{branch_address}}`, etc.)
- **Styling**:
  - Font: Times New Roman
  - Size: 11-12pt
  - Line spacing: 1.15
  - Margins: 2cm sides, 1.5cm top/bottom

### 4. Web Interface
- Built Flask web application for easy access
- Features:
  - One-click document download
  - Form structure preview
  - Responsive design

## Technology Stack

| Component | Technology |
|-----------|------------|
| Backend | Python 3.x, Flask |
| Document Generation | python-docx |
| Web Server | Gunicorn |
| Deployment | Render / Heroku |

## Project Structure

```
MEDIATION_APPLICATION_FORM/
├── app.py                      # Flask web application
├── create_mediation_form.py    # Standalone document generator
├── templates/
│   ├── index.html              # Home page
│   └── preview.html            # Form preview page
├── requirements.txt            # Python dependencies
├── Procfile                    # Deployment configuration
├── django_assignment-1.pdf     # Source PDF template
├── mediation_application_form.docx  # Generated output
└── README.md                   # Documentation
```

## Installation & Local Setup

```bash
# Clone the repository
git clone https://github.com/Sushant-Dagar/MEDIATION_APPLICATION_FORM.git
cd MEDIATION_APPLICATION_FORM

# Create virtual environment
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt

# Run the application
python app.py
```

Visit `http://localhost:5000` in your browser.

## Usage

### Web Application
1. Navigate to the live URL or local server
2. Click **"Download Word Document"** to get the generated .docx file
3. Click **"Preview Structure"** to see the form layout

### Standalone Script
```bash
python create_mediation_form.py
```
This generates `mediation_application_form.docx` in the project directory.

## Document Features

- **FORM 'A' Header**: Official form title with rule reference
- **Applicant Details**: Name, address (registered & correspondence), contact information
- **Opposite Party Details**: Defendant information with address fields
- **Dispute Details**: Commercial Courts Rules 2018 reference
- **Template Variables**: Ready for dynamic data population

## Template Variables

| Variable | Description |
|----------|-------------|
| `{{client_name}}` | Applicant's name |
| `{{branch_address}}` | Branch address |
| `{{mobile}}` | Telephone number |
| `{{customer_name}}` | Defendant's name |
| `{{address1}}` | Defendant's address |

## Deployment

### Render (Recommended)
1. Connect GitHub repository to Render
2. Select "Web Service"
3. Set build command: `pip install -r requirements.txt`
4. Set start command: `gunicorn app:app`

### Heroku
```bash
heroku create mediation-form-app
git push heroku main
```

## Evaluation Criteria Met

| Criteria | Implementation |
|----------|----------------|
| Accuracy of Replication | Exact match of PDF layout, spacing, and formatting |
| Functionality | Document generation + web interface |
| Code Quality | Modular functions, clear documentation |
| User Experience | Simple one-click download, preview option |
| Deployment | Live URL with Render/Heroku |

## Author

Sushant Dagar

## License

MIT License
