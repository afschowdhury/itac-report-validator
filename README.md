# ITAC Report Validator

A beautiful web application for comparing and validating ITAC (Industrial Assessment Center) reports between DOCX and Excel formats. The application extracts data from both file types and highlights mismatches with detailed comparison views.

## Features

- 🎨 **Beautiful Web Interface** - Modern, responsive design with Bootstrap 5
- 📄 **DOCX Report Processing** - Extracts general information, energy usage, and assessment recommendations
- 📊 **Excel Template Analysis** - Processes IAC Assessment Template Excel files
- 🔍 **Smart Comparison** - Compares values with configurable tolerance for numeric data
- 🎯 **Visual Highlighting** - Color-coded highlights for matches, mismatches, and missing values
- 📱 **Responsive Design** - Works on desktop, tablet, and mobile devices
- 💾 **Drag & Drop Upload** - Easy file upload with drag and drop support

## Screenshots

### Upload Interface
The main upload page provides an intuitive interface for selecting DOCX and Excel files:
- Drag and drop file upload areas
- File validation and size checking
- Visual feedback for file selection

### Comparison Results
The comparison view shows detailed analysis with:
- Summary statistics cards
- Color-coded field comparisons
- Expandable sections for detailed data
- Original document HTML rendering

## Installation

1. **Clone the repository**
   ```bash
   git clone <your-repo-url>
   cd itac-report-validator
   ```

2. **Install dependencies**
   ```bash
   pip install -r requirements.txt
   ```

3. **Run the application**
   ```bash
   python app.py
   ```

4. **Open in browser**
   - Navigate to `http://localhost:5000`

## Usage

### Step 1: Upload Files
1. Open the web application in your browser
2. Upload your DOCX report file (IAC assessment document)
3. Upload your Excel file (IAC Assessment Template)
4. Click "Compare Reports"

### Step 2: Review Results
The comparison page shows:

#### Summary Statistics
- Number of matched fields
- Number of mismatched fields
- Total energy types compared
- Overall cost match status

#### General Information Comparison
- Field-by-field comparison of general info data
- Color-coded highlighting:
  - 🟢 **Green**: Perfect match
  - 🔴 **Red**: Mismatch detected
  - 🟡 **Yellow**: Missing value in one source

#### Energy Usage Comparison
- Energy type comparisons with cost and usage data
- Total cost validation
- Unit-aware comparisons (kWh, MMBtu, etc.)

#### Document Sections
- View original HTML extractions from DOCX
- Collapsible sections for detailed inspection

## File Requirements

### DOCX Files
- Must be IAC assessment reports in DOCX format
- Should contain standard sections:
  - General Information
  - Annual Energy Usages and Costs
  - Carbon Footprint
  - Assessment Recommendations (AR sections)

### Excel Files
- Must be IAC Assessment Template format (.xlsx)
- Should contain standard sheets:
  - General Info
  - Energy-Waste Info
  - Recommendation Info

## API Usage

The application also provides a JSON API endpoint:

```bash
curl -X POST http://localhost:5000/api/compare \
  -F "docx_file=@report.docx" \
  -F "excel_file=@template.xlsx"
```

Returns JSON with comparison results for programmatic access.

## Configuration

### Tolerance Settings
Numeric comparisons use a default tolerance of 1% (0.01). This can be adjusted in the `compare_values()` function in `app.py`.

### File Size Limits
- Maximum file size: 16MB per file
- Supported formats: `.docx` and `.xlsx` only

## Technical Details

### Architecture
- **Backend**: Flask web framework
- **Frontend**: Bootstrap 5, Font Awesome icons
- **Document Processing**: python-docx library
- **Excel Processing**: openpyxl and pandas
- **HTML Parsing**: BeautifulSoup4

### Extraction Logic
1. **DOCX Processing** (`document_extractor.py`):
   - Extracts structured data from Word documents
   - Converts to both HTML and JSON formats
   - Handles tables, paragraphs, and formatting

2. **Excel Processing** (`excel_keyinfo_extractor.py`):
   - Processes IAC Assessment Template sheets
   - Extracts key-value pairs and structured data
   - Standardizes field names for comparison

3. **Comparison Engine** (`app.py`):
   - Normalizes data types between sources
   - Handles numeric tolerance comparisons
   - Provides detailed mismatch analysis

### Highlighting Rules
- **Perfect Match**: Values are identical or within tolerance
- **Numeric Mismatch**: Values differ beyond tolerance threshold
- **Text Mismatch**: String values don't match (case-insensitive)
- **Missing Value**: Value present in one source but not the other
- **Type Mismatch**: Different data types that can't be compared

## Development

### Project Structure
```
itac-report-validator/
├── app.py                      # Main Flask application
├── document_extractor.py       # DOCX processing logic
├── excel_keyinfo_extractor.py  # Excel processing logic
├── templates/                  # HTML templates
│   ├── base.html              # Base template
│   ├── index.html             # Upload page
│   └── comparison.html        # Results page
├── static/                     # Static assets
│   ├── css/main.css           # Custom styles
│   └── js/main.js             # JavaScript functionality
├── uploads/                    # Temporary file storage
└── requirements.txt           # Python dependencies
```

### Adding Features
1. **New Comparison Types**: Extend the `compare_values()` function
2. **Additional File Formats**: Add new extractors following the existing pattern
3. **Enhanced Visualizations**: Modify the comparison templates
4. **API Enhancements**: Extend the `/api/compare` endpoint

## Troubleshooting

### Common Issues

1. **File Upload Errors**
   - Check file format (.docx, .xlsx only)
   - Verify file size is under 16MB
   - Ensure files are not corrupted

2. **Extraction Errors**
   - Verify document structure matches IAC format
   - Check for missing required sections
   - Review error messages in browser console

3. **Comparison Issues**
   - Ensure both files contain comparable data
   - Check for consistent field naming
   - Verify numeric formats are valid

### Debug Mode
Run the application in debug mode for detailed error messages:
```bash
export FLASK_DEBUG=1
python app.py
```

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Test thoroughly
5. Submit a pull request

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Support

For issues, questions, or feature requests, please:
1. Check the troubleshooting section
2. Search existing issues
3. Create a new issue with detailed information

---

Built with ❤️ for the Industrial Assessment Center community.
