# File Converter

A Windows-friendly desktop and command-line application for converting documents, spreadsheets, presentations, images, and HTML files to PDF.

## Supported formats

| Category | Formats |
|---|---|
| Word | `.docx`, `.doc` |
| Excel | `.xlsx`, `.xls` |
| PowerPoint | `.pptx`, `.ppt` |
| Images | `.png`, `.jpg`, `.jpeg`, `.gif`, `.bmp`, `.tiff` |
| HTML | `.html`, `.htm` |

## Requirements

- Python 3.7 or later
- Packages listed in `requirements.txt`
- Microsoft Office on Windows for Excel and PowerPoint conversion
- Microsoft Word for Word conversion with `docx2pdf`
- `wkhtmltopdf` installed and available on `PATH` for HTML conversion
- Tkinter for the graphical interface (usually included with Python)

Install the Python dependencies:

```powershell
python -m pip install -r requirements.txt
```

## Run from source

Launch the graphical interface:

```powershell
python file_converter.py
```

The GUI supports selecting multiple files, clearing the selection, and converting files to PDF next to their originals.

Convert a file from the command line:

```powershell
python file_converter.py document.docx
python file_converter.py image.png output.pdf
python file_converter.py --help
```

## Build a Windows application

The repository includes `app_icon.ico`. The GUI applies this icon to its window and taskbar entry. To also give the executable and its desktop shortcuts the same icon, build a standalone application with PyInstaller:

```powershell
python -m pip install pyinstaller
pyinstaller --noconfirm --onefile --windowed --name FileConverter --icon app_icon.ico --add-data "app_icon.ico;." file_converter.py
```

The executable is created at `dist\FileConverter.exe`. Create a desktop shortcut to that executable, or pin it to the taskbar; Windows will use the embedded application icon. Keep `app_icon.ico` alongside `file_converter.py` when running from source.

## Use as a library

```python
from file_converter import FileConverter

converter = FileConverter()
pdf_path = converter.convert("document.docx")
```

An output path can be provided as the second argument:

```python
pdf_path = converter.convert("image.png", "output/result.pdf")
```
