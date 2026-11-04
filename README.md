# File-Converter
A simple and quick file converter for everyone's need.

## Features
Convert various file formats to PDF:
- **Word documents**: .docx, .doc
- **Excel spreadsheets**: .xlsx, .xls
- **PowerPoint presentations**: .pptx, .ppt
- **Images**: .png, .jpg, .jpeg, .gif, .bmp, .tiff
- **HTML files**: .html, .htm

Available in both **GUI** and **command-line** modes!

## Installation

1. Install Python 3.7 or higher
2. Install required dependencies:
```bash
pip install -r requirements.txt
```

**Note**: For HTML to PDF conversion, you also need to install [wkhtmltopdf](https://wkhtmltopdf.org/) and add it to your system PATH.

## Usage

### GUI Mode (Recommended)
Simply run the application without arguments to launch the graphical interface:
```bash
python file_converter.py
```

The GUI features:
- File browser for easy file selection
- Auto-generate output filename option
- Progress indicator during conversion
- Status messages and error handling
- Support for all file formats

### Command Line Mode
```bash
# Basic usage (output file will have same name with .pdf extension)
python file_converter.py document.docx

# Specify custom output path
python file_converter.py image.png output.pdf

# Convert Excel
python file_converter.py spreadsheet.xlsx

# Convert PowerPoint
python file_converter.py presentation.pptx

# Convert HTML
python file_converter.py webpage.html
```

### As a Library
```python
from file_converter import FileConverter

converter = FileConverter()

# Convert a file
converter.convert('document.docx', 'output.pdf')

# Or let it auto-generate the output filename
converter.convert('image.png')
```

## Requirements
- Python 3.7+ (tkinter is included with Python)
- docx2pdf (for Word conversion)
- Pillow (for image conversion)
- pdfkit + wkhtmltopdf (for HTML conversion)
- pywin32 (for Excel/PowerPoint conversion on Windows)

## Notes
- Excel and PowerPoint conversion require Windows and Microsoft Office installed
- HTML conversion requires wkhtmltopdf to be installed separately
- The GUI uses tkinter, which is included with standard Python installations
