#!/usr/bin/env python3
"""
File Converter to PDF
Converts various file formats to PDF: Word, Excel, PowerPoint, Images, HTML
"""

import os
import sys
import threading
import gc
from queue import Empty, SimpleQueue
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pathlib import Path
from typing import Optional


class FileConverter:
    """Main file converter class supporting multiple formats"""

    def __init__(self):
        self.supported_formats = {
            'word': ['.docx'],
            'excel': ['.xlsx', '.xls'],
            'powerpoint': ['.pptx', '.ppt'],
            'image': ['.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tiff'],
            'html': ['.html', '.htm']
        }
        self._excel_app = None
        self._pythoncom = None
        self.progress_callback = None

    @staticmethod
    def _error_detail(error: Exception) -> str:
        """Return useful text even when an exception has an empty message."""
        detail = str(error).strip()
        return detail if detail else repr(error)

    def start_excel_session(self):
        """Start an Excel COM session for a batch of conversions."""
        try:
            import win32com.client
            import pythoncom
        except ImportError:
            raise ImportError(
                "pywin32 library not installed. "
                "Install it with: pip install pywin32"
            )

        pythoncom.CoInitialize()
        try:
            excel = win32com.client.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            excel.EnableEvents = False
            excel.ScreenUpdating = False
            self._excel_app = excel
            self._pythoncom = pythoncom
        except Exception:
            pythoncom.CoUninitialize()
            raise

    def close_excel_session(self):
        """Close the batch Excel COM session and release its COM references."""
        excel = self._excel_app
        pythoncom = self._pythoncom
        self._excel_app = None
        self._pythoncom = None

        try:
            if excel is not None:
                excel.Quit()
        finally:
            excel = None
            gc.collect()
            if pythoncom is not None:
                pythoncom.CoUninitialize()

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
        """Convert a .docx document to PDF."""
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
            raise RuntimeError(
                f"Failed to convert Word to PDF: {self._error_detail(e)}"
            ) from e

    def _convert_excel_to_pdf(self, input_path: str, output_path: str) -> str:
        """Convert Excel spreadsheet to PDF"""
        try:
            import win32com.client
            import pythoncom

            owns_session = self._excel_app is None
            if owns_session:
                pythoncom.CoInitialize()
                excel = None
            else:
                excel = self._excel_app
            workbook = None
            try:
                if owns_session:
                    excel = win32com.client.DispatchEx("Excel.Application")
                    excel.Visible = False
                    excel.DisplayAlerts = False
                    excel.EnableEvents = False
                    excel.ScreenUpdating = False

                workbook = excel.Workbooks.Open(
                    os.path.abspath(input_path),
                    UpdateLinks=0,
                    ReadOnly=True,
                    IgnoreReadOnlyRecommended=True
                )
                workbook.ExportAsFixedFormat(0, os.path.abspath(output_path))
                if self.progress_callback is not None:
                    self.progress_callback(
                        f"PDF created: {Path(output_path).name}; closing workbook..."
                    )
                return output_path
            finally:
                try:
                    if workbook is not None:
                        workbook.Close(False)
                finally:
                    workbook = None
                    gc.collect()
                    if owns_session:
                        try:
                            if excel is not None:
                                excel.Quit()
                        finally:
                            excel = None
                            gc.collect()
                            pythoncom.CoUninitialize()
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
        self._conversion_queue = SimpleQueue()
        self._conversion_in_progress = False
        self.converter.progress_callback = self._queue_conversion_status

        self.create_widgets()
        self.root.after(100, self._poll_conversion_queue)

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
            ('All supported files', '*.docx;*.xlsx;*.xls;*.pptx;*.ppt;*.png;*.jpg;*.jpeg;*.gif;*.bmp;*.tiff;*.html;*.htm'),
            ('Word documents', '*.docx'),
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
        if self._conversion_in_progress:
            return
        if not self.input_files:
            messagebox.showerror("Error", "Please select at least one file")
            return

        self._conversion_in_progress = True
        # Show progress bar and disable button during conversion
        self.progress.pack(fill=tk.X, pady=(0, 10))
        self.root.config(cursor='wait')
        self.progress.start(10)
        self.status_label.config(text="Converting...", foreground=self.accent_color)

        # Run conversion in a separate thread to keep GUI responsive
        thread = threading.Thread(target=self._convert_thread)
        thread.start()

    def _queue_conversion_status(self, status):
        """Queue a status update from the conversion worker."""
        self._conversion_queue.put(("status", status))

    def _poll_conversion_queue(self):
        """Apply worker updates on the Tk main thread."""
        try:
            while True:
                event = self._conversion_queue.get_nowait()
                event_type = event[0]
                if event_type == "status":
                    self.status_label.config(
                        text=event[1],
                        foreground=self.accent_color
                    )
                elif event_type == "success":
                    self._conversion_success(*event[1:])
                elif event_type == "error":
                    self._conversion_error(event[1])
        except Empty:
            pass
        self.root.after(100, self._poll_conversion_queue)

    def _convert_thread(self):
        """Run conversion in a separate thread"""
        success_count = 0
        error_count = 0
        errors = []
        warnings = []
        input_files = self.input_files.copy()
        excel_extensions = set(self.converter.supported_formats['excel'])
        has_excel_files = any(
            Path(filepath).suffix.lower() in excel_extensions
            for filepath in input_files
        )
        excel_session_started = False
        excel_session_error = None

        try:
            if has_excel_files:
                try:
                    self.converter.start_excel_session()
                    excel_session_started = True
                except Exception as e:
                    excel_session_error = self.converter._error_detail(e)

            for i, filepath in enumerate(input_files, start=1):
                self._queue_conversion_status(
                    f"Converting {i}/{len(input_files)}: {Path(filepath).name}"
                )
                try:
                    if (
                        Path(filepath).suffix.lower() in excel_extensions
                        and excel_session_error is not None
                    ):
                        raise RuntimeError(excel_session_error)
                    # Auto-generate output path in same directory as input file
                    self.converter.convert(filepath)
                    success_count += 1
                except Exception as e:
                    error_count += 1
                    errors.append(
                        f"{Path(filepath).name}: {self.converter._error_detail(e)}"
                    )

            if excel_session_started:
                self._queue_conversion_status("PDF created; finalizing Excel...")
                try:
                    self.converter.close_excel_session()
                except Exception as e:
                    error_count += 1
                    errors.append(
                        f"Excel cleanup: {self.converter._error_detail(e)}"
                    )

            # Update GUI from main thread
            self._conversion_queue.put(
                ("success", success_count, error_count, errors, warnings)
            )
        except Exception as e:
            if excel_session_started:
                try:
                    self.converter.close_excel_session()
                except Exception as cleanup_error:
                    e = RuntimeError(
                        f"{self.converter._error_detail(e)}; Excel cleanup also failed: "
                        f"{self.converter._error_detail(cleanup_error)}"
                    )
            self._conversion_queue.put(
                ("error", self.converter._error_detail(e))
            )

    def _conversion_success(self, success_count, error_count, errors, warnings=None):
        """Handle successful conversion"""
        warnings = warnings or []
        self._conversion_in_progress = False
        self.progress.stop()
        self.progress.pack_forget()
        self.root.config(cursor='')

        if error_count == 0 and not warnings:
            self.status_label.config(text=f"Success! {success_count} file(s) converted", foreground="#10b981")
            messagebox.showinfo("Success", f"Successfully converted {success_count} file(s) to PDF!")
        elif error_count == 0:
            self.status_label.config(
                text=f"Converted {success_count} file(s) with a cleanup warning",
                foreground="#f59e0b"
            )
            messagebox.showwarning(
                "Conversion succeeded with warning",
                f"Successfully converted {success_count} file(s) to PDF.\n\n"
                "Word did not shut down cleanly. The PDF is ready, but Word may "
                "remain running:\n" + "\n".join(warnings)
            )
        else:
            self.status_label.config(text=f"Partial success: {success_count} OK, {error_count} failed", foreground="#f59e0b")
            error_message = f"Converted {success_count} file(s) successfully.\n\nFailed files:\n" + "\n".join(errors)
            if warnings:
                error_message += "\n\nWarnings:\n" + "\n".join(warnings)
            messagebox.showwarning("Partial Success", error_message)

    def _conversion_error(self, error_message):
        """Handle conversion error"""
        self._conversion_in_progress = False
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
            print("  - Word: .docx")
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
        root = tk.Tk()
        app = FileConverterGUI(root)
        root.mainloop()


if __name__ == "__main__":
    main()
