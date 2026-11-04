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
        if input_file.is_dir():
            raise ValueError(f"Input path is a directory, not a file: {input_path}")

        if output_path is None:
            output_path = str(input_file.with_suffix('.pdf'))
        else:
            output_path = str(Path(output_path))

        output_file = Path(output_path)
        if output_file.parent and not output_file.parent.exists():
            output_file.parent.mkdir(parents=True, exist_ok=True)

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
        self.root.geometry("300x420")
        self.root.resizable(False, False)

        app_dir = Path(getattr(sys, '_MEIPASS', Path(__file__).resolve().parent))
        icon_path = app_dir / 'app_icon.ico'
        if sys.platform == 'win32':
            if not icon_path.is_file():
                raise FileNotFoundError(f"Application icon not found: {icon_path}")
            self.root.iconbitmap(str(icon_path))
            self.root.iconbitmap(default=str(icon_path))

        # Configure modern style
        self.style = ttk.Style()
        self.style.theme_use('clam')

        # Compact, modern color scheme
        self.bg_color = '#f3f4f6'
        self.card_bg = '#ffffff'
        self.accent_color = '#2563eb'
        self.button_bg = '#2563eb'
        self.button_hover = '#1d4ed8'
        self.button_fg = 'white'
        self.text_color = '#111827'
        self.text_secondary = '#6b7280'
        self.border_color = '#dfe3ea'

        self.root.configure(bg=self.bg_color)

        self.converter = FileConverter()
        self.input_files = []

        self.create_widgets()

    def create_widgets(self):
        """Create compact GUI widgets with a clean, polished layout"""
        header_frame = tk.Frame(self.root, bg=self.bg_color, padx=22, pady=18)
        header_frame.pack(fill=tk.X)

        title_label = tk.Label(
            header_frame,
            text="File Converter",
            font=('Segoe UI', 18, 'bold'),
            bg=self.bg_color,
            fg=self.text_color
        )
        title_label.pack(anchor=tk.W)

        subtitle_label = tk.Label(
            header_frame,
            text="Convert files to PDF format",
            font=('Segoe UI', 9),
            bg=self.bg_color,
            fg=self.text_secondary
        )
        subtitle_label.pack(anchor=tk.W, pady=(2, 0))

        content_frame = tk.Frame(self.root, bg=self.card_bg, padx=18, pady=16)
        content_frame.pack(fill=tk.BOTH, expand=True, padx=18, pady=(0, 18))

        files_label = tk.Label(
            content_frame,
            text="Selected Files",
            font=('Segoe UI', 10, 'bold'),
            bg=self.card_bg,
            fg=self.text_color
        )
        files_label.pack(anchor=tk.W, pady=(0, 8))

        listbox_container = tk.Frame(content_frame, bg=self.card_bg)
        listbox_container.pack(fill=tk.X, pady=(0, 10))

        scrollbar = tk.Scrollbar(listbox_container, bg=self.card_bg)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.file_listbox = tk.Listbox(
            listbox_container,
            font=('Segoe UI', 9),
            relief=tk.FLAT,
            bd=0,
            highlightthickness=1,
            highlightbackground=self.border_color,
            highlightcolor=self.accent_color,
            yscrollcommand=scrollbar.set,
            height=6,
            width=46,
            bg='white',
            selectbackground=self.accent_color,
            selectforeground='white',
            activestyle='none'
        )
        self.file_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.file_listbox.yview)

        self.file_listbox.insert(0, "No files selected")
        self.file_listbox.selection_clear(0, tk.END)
        self.file_listbox.config(fg=self.text_secondary)

        button_row = tk.Frame(content_frame, bg=self.card_bg)
        button_row.pack(fill=tk.X, pady=(0, 10))

        add_btn = tk.Button(
            button_row,
            text="Add Files",
            font=('Segoe UI', 9, 'bold'),
            bg=self.button_bg,
            fg=self.button_fg,
            relief=tk.FLAT,
            cursor='hand2',
            command=self.browse_input_file,
            padx=18,
            pady=7,
            activebackground=self.button_hover,
            activeforeground='white'
        )
        add_btn.pack(side=tk.LEFT, padx=(0, 8))

        clear_btn = tk.Button(
            button_row,
            text="Clear",
            font=('Segoe UI', 9, 'bold'),
            bg='#eef2ff',
            fg=self.text_color,
            relief=tk.FLAT,
            cursor='hand2',
            command=self.clear_files,
            padx=16,
            pady=7,
            activebackground='#e0e7ff',
            activeforeground=self.text_color,
            highlightthickness=1,
            highlightbackground=self.border_color
        )
        clear_btn.pack(side=tk.LEFT)

        convert_btn = tk.Button(
            content_frame,
            text="Convert to PDF",
            font=('Segoe UI', 10, 'bold'),
            bg=self.button_bg,
            fg=self.button_fg,
            relief=tk.FLAT,
            cursor='hand2',
            command=self.convert_file,
            padx=20,
            pady=10,
            activebackground=self.button_hover,
            activeforeground='white'
        )
        convert_btn.pack(fill=tk.X, pady=(0, 10))

        self.progress = ttk.Progressbar(
            content_frame,
            mode='indeterminate',
            style='Custom.Horizontal.TProgressbar'
        )

        self.status_label = tk.Label(
            content_frame,
            text="Ready - 0 files selected",
            font=('Segoe UI', 9),
            bg=self.card_bg,
            fg=self.text_secondary
        )
        self.status_label.pack()

        # Configure progress bar style
        self.style.configure(
            'Custom.Horizontal.TProgressbar',
            background=self.accent_color,
            troughcolor=self.border_color,
            borderwidth=0,
            thickness=6
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
            self.file_listbox.selection_clear(0, tk.END)
            self.file_listbox.config(fg=self.text_secondary)
            self.status_label.config(text="Ready - 0 files selected")
        else:
            self.file_listbox.config(fg=self.text_color)
            for filepath in self.input_files:
                filename = Path(filepath).name
                self.file_listbox.insert(tk.END, filename)
            self.file_listbox.selection_clear(0, tk.END)
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
        self.status_label.config(text="Converting...", foreground=self.accent_color)

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
            self.status_label.config(text=f"Success! {success_count} file(s) converted", foreground="#10b981")
            messagebox.showinfo("Success", f"Successfully converted {success_count} file(s) to PDF!")
        else:
            self.status_label.config(text=f"Partial success: {success_count} OK, {error_count} failed", foreground="#f59e0b")
            error_message = f"Converted {success_count} file(s) successfully.\n\nFailed files:\n" + "\n".join(errors)
            messagebox.showwarning("Partial Success", error_message)

    def _conversion_error(self, error_message):
        """Handle conversion error"""
        self.progress.stop()
        self.progress.pack_forget()
        self.root.config(cursor='')
        self.status_label.config(text="Conversion failed", foreground="#ef4444")
        messagebox.showerror("Error", f"Conversion failed:\n\n{error_message}")


def main():
    """Main entry point - GUI or CLI"""
    if len(sys.argv) > 1:
        if sys.argv[1] in {'-h', '--help'}:
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
            return

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
        if sys.platform == 'win32':
            import ctypes

            app_id = 'FileConverter.PDF'
            set_app_id = ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID
            set_app_id.argtypes = [ctypes.c_wchar_p]
            set_app_id.restype = ctypes.c_long
            result = set_app_id(app_id)
            if result != 0:
                raise OSError(result, "Could not set the Windows taskbar application ID")

        root = tk.Tk()
        app = FileConverterGUI(root)
        root.mainloop()


if __name__ == "__main__":
    main()
