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
        self.root.geometry("420x240")
        self.root.resizable(False, False)

        # Configure modern style
        self.style = ttk.Style()
        self.style.theme_use('clam')

        # Custom colors
        self.bg_color = '#f0f0f0'
        self.accent_color = '#0078d4'
        self.button_bg = '#0078d4'
        self.button_fg = 'white'
        self.text_color = '#333333'

        self.root.configure(bg=self.bg_color)

        self.converter = FileConverter()
        self.input_files = []

        self.create_widgets()

    def create_widgets(self):
        """Create GUI widgets"""
        # Main frame with modern styling
        main_frame = tk.Frame(self.root, bg=self.bg_color, padx=25, pady=20)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Title
        title_label = tk.Label(
            main_frame,
            text="File Converter",
            font=('Segoe UI', 16, 'bold'),
            bg=self.bg_color,
            fg=self.text_color
        )
        title_label.pack(pady=(0, 15))

        # Input file section
        input_frame = tk.Frame(main_frame, bg=self.bg_color)
        input_frame.pack(fill=tk.X, pady=(0, 12))

        # File listbox with scrollbar
        listbox_frame = tk.Frame(input_frame, bg=self.bg_color)
        listbox_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 10))

        scrollbar = tk.Scrollbar(listbox_frame, bg=self.bg_color)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.file_listbox = tk.Listbox(
            listbox_frame,
            font=('Segoe UI', 9),
            relief=tk.FLAT,
            bd=1,
            highlightthickness=1,
            highlightbackground=self.accent_color,
            highlightcolor=self.accent_color,
            yscrollcommand=scrollbar.set,
            height=3
        )
        self.file_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.file_listbox.yview)

        self.file_listbox.insert(0, "No files selected")
        self.file_listbox.config(fg='gray')

        button_frame = tk.Frame(input_frame, bg=self.bg_color)
        button_frame.pack(side=tk.RIGHT)

        browse_btn = tk.Button(
            button_frame,
            text="Add Files",
            font=('Segoe UI', 9, 'bold'),
            bg=self.button_bg,
            fg=self.button_fg,
            relief=tk.FLAT,
            cursor='hand2',
            command=self.browse_input_file,
            padx=12,
            pady=6,
            activebackground='#005a9e',
            activeforeground='white'
        )
        browse_btn.pack(pady=(0, 5))

        clear_btn = tk.Button(
            button_frame,
            text="Clear",
            font=('Segoe UI', 9, 'bold'),
            bg='#666666',
            fg='white',
            relief=tk.FLAT,
            cursor='hand2',
            command=self.clear_files,
            padx=12,
            pady=6,
            activebackground='#555555',
            activeforeground='white'
        )
        clear_btn.pack()

        # Convert button
        convert_btn = tk.Button(
            main_frame,
            text="Convert to PDF",
            font=('Segoe UI', 10, 'bold'),
            bg=self.button_bg,
            fg=self.button_fg,
            relief=tk.FLAT,
            cursor='hand2',
            command=self.convert_file,
            padx=25,
            pady=10,
            activebackground='#005a9e',
            activeforeground='white'
        )
        convert_btn.pack(fill=tk.X, pady=(0, 12))

        # Progress bar (hidden initially)
        self.progress = ttk.Progressbar(
            main_frame,
            mode='indeterminate',
            style='Custom.Horizontal.TProgressbar'
        )

        # Status label
        self.status_label = tk.Label(
            main_frame,
            text="Ready - 0 files selected",
            font=('Segoe UI', 8),
            bg=self.bg_color,
            fg='gray'
        )
        self.status_label.pack(pady=(0, 5))

        # Configure progress bar style
        self.style.configure(
            'Custom.Horizontal.TProgressbar',
            background=self.accent_color,
            troughcolor='#e0e0e0',
            borderwidth=0,
            thickness=8
        )

    def browse_input_file(self):
        """Browse for input files (multiple selection)"""
        filetypes = [
            ('All supported files', '*.docx;*.doc;*.xlsx;*.xls;*.pptx;*.ppt;*.png;*.jpg;*.jpeg;*.gif;*.bmp;*.tiff;*.html;*.htm'),
            ('Word documents', '*.docx;*.doc'),
            ('Excel spreadsheets', '*.xlsx;*.xls'),
            ('PowerPoint presentations', '*.pptx;*.ppt'),
            ('Images', '*.png;*.jpg;*.jpeg;*.gif;*.bmp;*.tiff'),
            ('HTML files', '*.html;*.htm'),
            ('All files', '*.*')
        ]

        filenames = filedialog.askopenfilenames(
            title='Select files to convert',
            filetypes=filetypes
        )

        if filenames:
            # Add new files to the list
            for filename in filenames:
                if filename not in self.input_files:
                    self.input_files.append(filename)

            self._update_file_list()

    def clear_files(self):
        """Clear all selected files"""
        self.input_files = []
        self._update_file_list()

    def _update_file_list(self):
        """Update the file listbox display"""
        self.file_listbox.delete(0, tk.END)

        if not self.input_files:
            self.file_listbox.insert(0, "No files selected")
            self.file_listbox.config(fg='gray')
            self.status_label.config(text="Ready - 0 files selected")
        else:
            self.file_listbox.config(fg=self.text_color)
            for filepath in self.input_files:
                filename = Path(filepath).name
                self.file_listbox.insert(tk.END, filename)
            self.status_label.config(text=f"Ready - {len(self.input_files)} file(s) selected")

    def convert_file(self):
        """Convert the selected files"""
        if not self.input_files:
            messagebox.showerror("Error", "Please select at least one file")
            return

        # Show progress bar and disable button during conversion
        self.progress.pack(fill=tk.X, pady=(0, 10))
        self.root.config(cursor='wait')
        self.progress.start(10)
        self.status_label.config(text="Converting...", foreground="orange")

        # Run conversion in a separate thread to keep GUI responsive
        thread = threading.Thread(target=self._convert_thread)
        thread.start()

    def _convert_thread(self):
        """Run conversion in a separate thread"""
        success_count = 0
        error_count = 0
        errors = []

        try:
            for i, filepath in enumerate(self.input_files):
                try:
                    # Auto-generate output path in same directory as input file
                    result = self.converter.convert(filepath)
                    success_count += 1
                except Exception as e:
                    error_count += 1
                    errors.append(f"{Path(filepath).name}: {str(e)}")

            # Update GUI from main thread
            self.root.after(0, self._conversion_success, success_count, error_count, errors)
        except Exception as e:
            self.root.after(0, self._conversion_error, str(e))

    def _conversion_success(self, success_count, error_count, errors):
        """Handle successful conversion"""
        self.progress.stop()
        self.progress.pack_forget()
        self.root.config(cursor='')

        if error_count == 0:
            self.status_label.config(text=f"Success! {success_count} file(s) converted", foreground="green")
            messagebox.showinfo("Success", f"Successfully converted {success_count} file(s) to PDF!")
        else:
            self.status_label.config(text=f"Partial success: {success_count} OK, {error_count} failed", foreground="orange")
            error_message = f"Converted {success_count} file(s) successfully.\n\nFailed files:\n" + "\n".join(errors)
            messagebox.showwarning("Partial Success", error_message)

    def _conversion_error(self, error_message):
        """Handle conversion error"""
        self.progress.stop()
        self.progress.pack_forget()
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
