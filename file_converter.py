#!/usr/bin/env python3
"""
File Converter to PDF
Converts various file formats to PDF: Word, Excel, PowerPoint, Images, HTML
"""

import os
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pathlib import Path
from typing import Optional


class FileConverter:
    """Main file converter class supporting multiple formats"""

    def __init__(self):
        self.supported_formats = {
            'word': ['.docx', '.doc'],
            'excel': ['.xlsx', '.xls'],
            'powerpoint': ['.pptx', '.ppt'],
            'image': ['.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tiff'],
            'html': ['.html', '.htm']
        }

    def convert(self, input_path: str, output_path: Optional[str] = None) -> str:
        """
        Convert a file to PDF based on its extension

        Args:
            input_path: Path to the input file
            output_path: Optional path for the output PDF (default: same name with .pdf)

        Returns:
            Path to the generated PDF file
        """
        input_file = Path(input_path)

        if not input_file.exists():
            raise FileNotFoundError(f"Input file not found: {input_path}")

        if output_path is None:
            output_path = str(input_file.with_suffix('.pdf'))
        else:
            output_path = str(Path(output_path))

        ext = input_file.suffix.lower()

        # Determine conversion type
        if ext in self.supported_formats['word']:
            return self._convert_word_to_pdf(input_path, output_path)
        elif ext in self.supported_formats['excel']:
            return self._convert_excel_to_pdf(input_path, output_path)
        elif ext in self.supported_formats['powerpoint']:
            return self._convert_powerpoint_to_pdf(input_path, output_path)
        elif ext in self.supported_formats['image']:
            return self._convert_image_to_pdf(input_path, output_path)
        elif ext in self.supported_formats['html']:
            return self._convert_html_to_pdf(input_path, output_path)
        else:
            raise ValueError(f"Unsupported file format: {ext}")

    def _convert_word_to_pdf(self, input_path: str, output_path: str) -> str:
        """Convert Word document to PDF"""
        try:
            from docx2pdf import convert
            convert(input_path, output_path)
            return output_path
        except ImportError:
            raise ImportError(
                "docx2pdf library not installed. "
                "Install it with: pip install docx2pdf"
            )
        except Exception as e:
            raise RuntimeError(f"Failed to convert Word to PDF: {e}")

    def _convert_excel_to_pdf(self, input_path: str, output_path: str) -> str:
        """Convert Excel spreadsheet to PDF"""
        try:
            import win32com.client
            import pythoncom

            pythoncom.CoInitialize()

            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False

            workbook = excel.Workbooks.Open(os.path.abspath(input_path))
            workbook.ExportAsFixedFormat(0, os.path.abspath(output_path))
            workbook.Close(False)
            excel.Quit()

            pythoncom.CoUninitialize()
            return output_path
        except ImportError:
            raise ImportError(
                "pywin32 library not installed. "
                "Install it with: pip install pywin32"
            )
        except Exception as e:
            raise RuntimeError(f"Failed to convert Excel to PDF: {e}")

    def _convert_powerpoint_to_pdf(self, input_path: str, output_path: str) -> str:
        """Convert PowerPoint presentation to PDF"""
        try:
            import win32com.client
            import pythoncom

            pythoncom.CoInitialize()

            powerpoint = win32com.client.Dispatch("PowerPoint.Application")
            powerpoint.Visible = False

            presentation = powerpoint.Presentations.Open(os.path.abspath(input_path))
            presentation.ExportAsFixedFormat(0, os.path.abspath(output_path))
            presentation.Close()
            powerpoint.Quit()

            pythoncom.CoUninitialize()
            return output_path
        except ImportError:
            raise ImportError(
                "pywin32 library not installed. "
                "Install it with: pip install pywin32"
            )
        except Exception as e:
            raise RuntimeError(f"Failed to convert PowerPoint to PDF: {e}")

    def _convert_image_to_pdf(self, input_path: str, output_path: str) -> str:
        """Convert image to PDF"""
        try:
            from PIL import Image

            image = Image.open(input_path)

            # Convert RGBA to RGB if necessary
            if image.mode == 'RGBA':
                rgb_image = Image.new('RGB', image.size, (255, 255, 255))
                rgb_image.paste(image, mask=image.split()[3])
                image = rgb_image
            elif image.mode != 'RGB':
                image = image.convert('RGB')

            image.save(output_path, 'PDF', resolution=100.0)
            return output_path
        except ImportError:
            raise ImportError(
                "Pillow library not installed. "
                "Install it with: pip install Pillow"
            )
        except Exception as e:
            raise RuntimeError(f"Failed to convert image to PDF: {e}")

    def _convert_html_to_pdf(self, input_path: str, output_path: str) -> str:
        """Convert HTML file to PDF"""
        try:
            import pdfkit

            options = {
                'encoding': 'UTF-8',
                'quiet': ''
            }
            pdfkit.from_file(input_path, output_path, options=options)
            return output_path
        except ImportError:
            raise ImportError(
                "pdfkit library not installed. "
                "Install it with: pip install pdfkit"
            )
        except Exception as e:
            raise RuntimeError(f"Failed to convert HTML to PDF: {e}")


class FileConverterGUI:
    """GUI application for file conversion"""

    def __init__(self, root):
        self.root = root
        self.root.title("File Converter")
        self.root.geometry("600x400")
        self.root.resizable(True, True)

        self.converter = FileConverter()
        self.input_file = None
        self.output_file = None

        self.create_widgets()

    def create_widgets(self):
        """Create GUI widgets"""
        # Main frame
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)

        # Title
        title_label = ttk.Label(
            main_frame,
            text="File Converter to PDF",
            font=('Helvetica', 16, 'bold')
        )
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))

        # Input file section
        ttk.Label(main_frame, text="Input File:").grid(row=1, column=0, sticky=tk.W, pady=5)

        self.input_entry = ttk.Entry(main_frame, width=50)
        self.input_entry.grid(row=1, column=1, sticky=(tk.W, tk.E), padx=5, pady=5)

        browse_input_btn = ttk.Button(
            main_frame,
            text="Browse...",
            command=self.browse_input_file
        )
        browse_input_btn.grid(row=1, column=2, padx=5, pady=5)

        # Output file section
        ttk.Label(main_frame, text="Output File:").grid(row=2, column=0, sticky=tk.W, pady=5)

        self.output_entry = ttk.Entry(main_frame, width=50)
        self.output_entry.grid(row=2, column=1, sticky=(tk.W, tk.E), padx=5, pady=5)

        browse_output_btn = ttk.Button(
            main_frame,
            text="Browse...",
            command=self.browse_output_file
        )
        browse_output_btn.grid(row=2, column=2, padx=5, pady=5)

        # Auto-generate output checkbox
        self.auto_output_var = tk.BooleanVar(value=True)
        auto_output_check = ttk.Checkbutton(
            main_frame,
            text="Auto-generate output filename",
            variable=self.auto_output_var,
            command=self.toggle_output_entry
        )
        auto_output_check.grid(row=3, column=0, columnspan=3, sticky=tk.W, pady=5)

        # Convert button
        convert_btn = ttk.Button(
            main_frame,
            text="Convert to PDF",
            command=self.convert_file
        )
        convert_btn.grid(row=4, column=0, columnspan=3, pady=20)

        # Progress bar
        self.progress = ttk.Progressbar(
            main_frame,
            mode='indeterminate',
            length=400
        )
        self.progress.grid(row=5, column=0, columnspan=3, pady=10)

        # Status label
        self.status_label = ttk.Label(main_frame, text="Ready", foreground="gray")
        self.status_label.grid(row=6, column=0, columnspan=3, pady=5)

        # Supported formats info
        info_text = (
            "Supported formats: Word (.docx, .doc), Excel (.xlsx, .xls), "
            "PowerPoint (.pptx, .ppt), Images (.png, .jpg, .jpeg, .gif, .bmp, .tiff), "
            "HTML (.html, .htm)"
        )
        info_label = ttk.Label(main_frame, text=info_text, foreground="blue", wraplength=550)
        info_label.grid(row=7, column=0, columnspan=3, pady=20)

        # Initially disable output entry if auto-generate is checked
        self.toggle_output_entry()

    def toggle_output_entry(self):
        """Enable/disable output entry based on checkbox"""
        if self.auto_output_var.get():
            self.output_entry.config(state='disabled')
        else:
            self.output_entry.config(state='normal')

    def browse_input_file(self):
        """Browse for input file"""
        filetypes = [
            ('All supported files', '*.docx;*.doc;*.xlsx;*.xls;*.pptx;*.ppt;*.png;*.jpg;*.jpeg;*.gif;*.bmp;*.tiff;*.html;*.htm'),
            ('Word documents', '*.docx;*.doc'),
            ('Excel spreadsheets', '*.xlsx;*.xls'),
            ('PowerPoint presentations', '*.pptx;*.ppt'),
            ('Images', '*.png;*.jpg;*.jpeg;*.gif;*.bmp;*.tiff'),
            ('HTML files', '*.html;*.htm'),
            ('All files', '*.*')
        ]

        filename = filedialog.askopenfilename(
            title='Select file to convert',
            filetypes=filetypes
        )

        if filename:
            self.input_file = filename
            self.input_entry.delete(0, tk.END)
            self.input_entry.insert(0, filename)

            # Auto-generate output filename
            if self.auto_output_var.get():
                input_path = Path(filename)
                auto_output = str(input_path.with_suffix('.pdf'))
                self.output_entry.delete(0, tk.END)
                self.output_entry.insert(0, auto_output)
                self.output_file = auto_output

    def browse_output_file(self):
        """Browse for output file"""
        filename = filedialog.asksaveasfilename(
            title='Save PDF as',
            defaultextension='.pdf',
            filetypes=[('PDF files', '*.pdf')]
        )

        if filename:
            self.output_file = filename
            self.output_entry.delete(0, tk.END)
            self.output_entry.insert(0, filename)
            self.auto_output_var.set(False)
            self.toggle_output_entry()

    def convert_file(self):
        """Convert the selected file"""
        if not self.input_file:
            messagebox.showerror("Error", "Please select an input file")
            return

        if not self.auto_output_var.get() and not self.output_file:
            messagebox.showerror("Error", "Please select an output file")
            return

        # Disable button during conversion
        self.root.config(cursor='wait')
        self.progress.start(10)
        self.status_label.config(text="Converting...", foreground="orange")

        # Run conversion in a separate thread to keep GUI responsive
        thread = threading.Thread(target=self._convert_thread)
        thread.start()

    def _convert_thread(self):
        """Run conversion in a separate thread"""
        try:
            output = self.output_file if not self.auto_output_var.get() else None
            result = self.converter.convert(self.input_file, output)

            # Update GUI from main thread
            self.root.after(0, self._conversion_success, result)
        except Exception as e:
            self.root.after(0, self._conversion_error, str(e))

    def _conversion_success(self, result):
        """Handle successful conversion"""
        self.progress.stop()
        self.root.config(cursor='')
        self.status_label.config(text="Conversion successful!", foreground="green")
        messagebox.showinfo("Success", f"File converted successfully!\n\nOutput: {result}")

    def _conversion_error(self, error_message):
        """Handle conversion error"""
        self.progress.stop()
        self.root.config(cursor='')
        self.status_label.config(text="Conversion failed", foreground="red")
        messagebox.showerror("Error", f"Conversion failed:\n\n{error_message}")


def main():
    """Main entry point - GUI or CLI"""
    # If command-line arguments provided, use CLI mode
    if len(sys.argv) >= 2:
        if len(sys.argv) < 2:
            print("File Converter to PDF")
            print("=" * 40)
            print("\nUsage:")
            print("  python file_converter.py <input_file> [output_file]")
            print("\nSupported formats:")
            print("  - Word: .docx, .doc")
            print("  - Excel: .xlsx, .xls")
            print("  - PowerPoint: .pptx, .ppt")
            print("  - Images: .png, .jpg, .jpeg, .gif, .bmp, .tiff")
            print("  - HTML: .html, .htm")
            print("\nExample:")
            print("  python file_converter.py document.docx")
            print("  python file_converter.py image.png output.pdf")
            sys.exit(1)

        input_file = sys.argv[1]
        output_file = sys.argv[2] if len(sys.argv) > 2 else None

        converter = FileConverter()

        try:
            result = converter.convert(input_file, output_file)
            print(f"✓ Successfully converted to: {result}")
        except Exception as e:
            print(f"✗ Error: {e}", file=sys.stderr)
            sys.exit(1)
    else:
        # Launch GUI
        root = tk.Tk()
        app = FileConverterGUI(root)
        root.mainloop()


if __name__ == "__main__":
    main()
