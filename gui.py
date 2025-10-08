from ttkbootstrap.constants import *
import requests
import ttkbootstrap as ttk
import pandas as pd
from tkinter import filedialog, StringVar
from ttkbootstrap.dialogs import Messagebox
from automation import extract_po_data
from automation_product_order import process_excel_files
from automation_technical_report import extract_technical_report_data
import os
from datetime import datetime


class PDFExtractor:
    def __init__(self):
        self.file_path = ""

    def on_entry_click(self, event, entry_widget, default_text):
        if entry_widget.get() == default_text:
            entry_widget.delete(0, 'end')
            entry_widget.config(foreground='black')

    def on_focus_out(self, event, entry_widget, default_text):
        if not entry_widget.get():
            entry_widget.insert(0, default_text)
            entry_widget.config(foreground='grey')

    def apply_placeholder(self, entry_widget, default_text):
        entry_widget.insert(0, default_text)
        entry_widget.config(foreground='grey')
        entry_widget.bind('<FocusIn>', lambda event: self.on_entry_click(event, entry_widget, default_text))
        entry_widget.bind('<FocusOut>', lambda event: self.on_focus_out(event, entry_widget, default_text))

    def browse_files(self, entry_widget, file_type="pdf"):
        if file_type == "excel":
            file_paths = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        else:
            file_paths = filedialog.askopenfilenames(filetypes=[("PDF files", "*.pdf")])

        if file_paths:
            entry_widget.delete(0, 'end')
            if isinstance(file_paths, tuple):  # Multiple files
                entry_widget.insert(0, "; ".join(file_paths))
            else:  # Single file
                entry_widget.insert(0, file_paths)
            entry_widget.config(foreground='black')

    def browse_excel(self, entry):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            entry.delete(0, 'end')
            entry.insert(0, file_path)
            entry.config(foreground='black')

    def process_excel_files(self):
        try:
            po_file = self.entry_po.get()
            order_form_file = self.entry_order.get()

            if (po_file == 'Select PO File' or po_file == '') or \
               (order_form_file == 'Select Order Form File' or order_form_file == ''):
                Messagebox.show_error('Error', 'Please select both PO and Order Form files')
                return

            output, labels_and_quantities = process_excel_files(po_file, order_form_file)
            Messagebox.show_info("Success", f"File processed successfully and saved at:\n{output}")
        except Exception as e:
            Messagebox.show_error("Processing Error", f"An error occurred:\n{str(e)}")

    def start_processing(self):
        try:
            pdf_files = self.file_entry_pdf.get()

            if pdf_files == 'Select PDF Files' or pdf_files == '':
                Messagebox.show_error('Error', 'Please select PDF files')
                return

            pdf_files = [f.strip() for f in pdf_files.split(";") if f.strip()]
            all_data = []

            try:
                result_window = ttk.Toplevel()
                result_window.title('Extracted Data')
                result_window.geometry("600x600")

                canvas = ttk.Canvas(result_window)
                scrollbar = ttk.Scrollbar(result_window, orient="vertical", command=canvas.yview)
                scroll_frame = ttk.Frame(canvas)

                scroll_frame.bind(
                    "<Configure>",
                    lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
                )

                canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
                canvas.configure(yscrollcommand=scrollbar.set)
                canvas.pack(side="left", fill="both", expand=True)
                scrollbar.pack(side="right", fill="y")

                for pdf_file in pdf_files:
                    try:
                        data = extract_po_data(pdf_file)
                        all_data.append(data)

                        file_label = ttk.Label(
                            scroll_frame,
                            text=f"📄 {pdf_file}",
                            font=('Helvetica', 14, 'bold'),
                            wraplength=560
                        )
                        file_label.pack(pady=(10, 5), anchor="w")

                        for key, value in data.items():
                            if value:
                                label = ttk.Label(
                                    scroll_frame,
                                    text=f"{key}: {value}",
                                    wraplength=560
                                )
                                label.pack(pady=2, anchor="w")

                        separator = ttk.Label(scroll_frame, text="-" * 80)
                        separator.pack(pady=8, anchor="w")

                    except Exception as e:
                        error_label = ttk.Label(
                            scroll_frame,
                            text=f"Error processing {pdf_file}: {str(e)}",
                            foreground="red",
                            wraplength=560
                        )
                        error_label.pack(pady=5, anchor="w")

                if all_data:
                    df = pd.DataFrame(all_data)
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    output_dir = os.path.dirname(pdf_files[0])
                    excel_path = os.path.join(output_dir, f"purchase_order_{timestamp}.xlsx")
                    df.to_excel(excel_path, index=False)

                    success_label = ttk.Label(
                        scroll_frame,
                        text=f"\n✅ Data saved to Excel file:\n{excel_path}",
                        bootstyle="success",
                        font=('Helvetica', 12, 'bold'),
                        wraplength=560
                    )
                    success_label.pack(pady=10, anchor="w")

            except Exception as e:
                Messagebox.show_error("Processing Error", f"An error occurred:\n{str(e)}")
        except Exception as e:
            Messagebox.show_error("Processing Error", f"An error occurred:\n{str(e)}")

    def setup_po_extractor_tab(self, tab):
        """Setup the Purchase Order Extractor tab"""
        # Frame for selecting files
        file_frame = ttk.Frame(tab)
        file_frame.pack(fill='x', padx=20, pady=(20, 10))

        # PDF Files selection
        label_filepath = ttk.Label(
            file_frame,
            text='PDF Files:',
            font=('Helvetica', 11)
        )
        label_filepath.pack(side='left', padx=(0, 10))

        self.file_entry_pdf = ttk.Entry(file_frame, bootstyle='primary', width=50)
        self.apply_placeholder(self.file_entry_pdf, 'Select PDF Files')
        self.file_entry_pdf.pack(side='left', expand=True, fill='x', padx=(0, 10))

        browse_button = ttk.Button(
            file_frame,
            width=10,
            text='Browse',
            bootstyle=SUCCESS,
            command=lambda: self.browse_files(self.file_entry_pdf),
            style='success.Outline.TButton'
        )
        browse_button.pack(side='right')

        # Frame for action buttons
        button_frame = ttk.Frame(tab)
        button_frame.pack(fill='x', padx=20, pady=10)

        self.po_execute_button = ttk.Button(
            button_frame,
            width=46,
            text='Process PDFs',
            bootstyle=PRIMARY,
            command=self.start_processing,
            style='primary.TButton'
        )
        self.po_execute_button.pack(fill='x')

    def setup_product_order_tab(self, tab):
        """Setup the Product Order Automation tab"""
        # Entry for PO File
        po_frame = ttk.Frame(tab)
        po_frame.pack(fill='x', padx=20, pady=(20, 5))

        self.entry_po = ttk.Entry(po_frame, bootstyle='primary', width=50)
        self.apply_placeholder(self.entry_po, 'Select PO File')
        self.entry_po.pack(side='left', expand=True, fill='x', padx=(0, 10))

        browse_po_button = ttk.Button(
            po_frame,
            width=10,
            text='Browse',
            bootstyle=SUCCESS,
            command=lambda: self.browse_excel(self.entry_po),
            style='success.Outline.TButton'
        )
        browse_po_button.pack(side='right')

        # Entry for Order Form File
        order_frame = ttk.Frame(tab)
        order_frame.pack(fill='x', padx=20, pady=5)

        self.entry_order = ttk.Entry(order_frame, bootstyle='primary', width=50)
        self.apply_placeholder(self.entry_order, 'Select Order Form File')
        self.entry_order.pack(side='left', expand=True, fill='x', padx=(0, 10))

        browse_order_button = ttk.Button(
            order_frame,
            width=10,
            text='Browse',
            bootstyle=SUCCESS,
            command=lambda: self.browse_excel(self.entry_order),
            style='success.Outline.TButton'
        )
        browse_order_button.pack(side='right')

        # Frame for action buttons
        button_frame = ttk.Frame(tab)
        button_frame.pack(fill='x', padx=20, pady=10)

        self.product_execute_button = ttk.Button(
            button_frame,
            width=46,
            text='Process Files',
            bootstyle=PRIMARY,
            command=self.process_excel_files,
            style='primary.TButton'
        )
        self.product_execute_button.pack(fill='x')

    def setup_technical_report_tab(self, tab):
        """Setup the Technical Report tab"""
        # Frame for selecting PDF file
        file_frame = ttk.Frame(tab)
        file_frame.pack(fill='x', padx=20, pady=(20, 10))

        # Technical Report PDF selection
        label_filepath = ttk.Label(
            file_frame,
            text='Technical Report PDF(s):',
            font=('Helvetica', 11)
        )
        label_filepath.pack(side='left', padx=(0, 10))

        self.file_entry_tech = ttk.Entry(file_frame, bootstyle='primary', width=50)
        self.apply_placeholder(self.file_entry_tech, 'Select Technical Report PDF files')
        self.file_entry_tech.pack(side='left', expand=True, fill='x', padx=(0, 10))

        browse_button = ttk.Button(
            file_frame,
            width=10,
            text='Browse',
            bootstyle=SUCCESS,
            command=lambda: self.browse_files(self.file_entry_tech),
            style='success.Outline.TButton'
        )
        browse_button.pack(side='right')

        # Frame for optional Excel file
        excel_frame = ttk.Frame(tab)
        excel_frame.pack(fill='x', padx=20, pady=10)

        # Optional Excel file selection
        label_excel = ttk.Label(
            excel_frame,
            text='Existing Excel (Optional):',
            font=('Helvetica', 11)
        )
        label_excel.pack(side='left', padx=(0, 10))

        self.file_entry_tech_excel = ttk.Entry(excel_frame, bootstyle='primary', width=50)
        self.apply_placeholder(self.file_entry_tech_excel, 'Select existing Excel file (optional)')
        self.file_entry_tech_excel.pack(side='left', expand=True, fill='x', padx=(0, 10))

        browse_excel_button = ttk.Button(
            excel_frame,
            width=10,
            text='Browse',
            bootstyle=SUCCESS,
            command=lambda: self.browse_files(self.file_entry_tech_excel, "excel"),
            style='success.Outline.TButton'
        )
        browse_excel_button.pack(side='right')

        # Frame for action buttons
        button_frame = ttk.Frame(tab)
        button_frame.pack(fill='x', padx=20, pady=10)

        self.tech_execute_button = ttk.Button(
            button_frame,
            width=46,
            text='Process Technical Report',
            bootstyle=PRIMARY,
            command=self.process_technical_report,
            style='primary.TButton'
        )
        self.tech_execute_button.pack(fill='x')

    def process_technical_report(self):
        try:
            pdf_files_str = self.file_entry_tech.get()
            excel_file = self.file_entry_tech_excel.get()

            if pdf_files_str == 'Select Technical Report PDF files' or pdf_files_str == '':
                Messagebox.show_error('Error', 'Please select Technical Report PDF file(s)')
                return

            # Make Excel file mandatory
            if excel_file == 'Select existing Excel file (optional)' or not excel_file:
                Messagebox.show_error('Error', 'Please select an Excel file to add the data to')
                return

            # Check if Excel file exists
            if not os.path.exists(excel_file):
                Messagebox.show_error('Error', 'Selected Excel file does not exist')
                return

            # Split the file paths (they are joined by "; ")
            pdf_files = [f.strip() for f in pdf_files_str.split(";")]
            
            # Show progress in a new window
            result_window = ttk.Toplevel()
            result_window.title('Processing Technical Reports')
            result_window.geometry('800x600')

            # Create a scrollable frame for results
            main_frame = ttk.Frame(result_window)
            main_frame.pack(fill=BOTH, expand=True, padx=10, pady=10)

            canvas = ttk.Canvas(main_frame)
            scrollbar = ttk.Scrollbar(main_frame, orient=VERTICAL, command=canvas.yview)
            scroll_frame = ttk.Frame(canvas)

            scroll_frame.bind(
                "<Configure>",
                lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
            )

            canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
            canvas.configure(yscrollcommand=scrollbar.set)

            # Pack scrollbar components
            scrollbar.pack(side=RIGHT, fill=Y)
            canvas.pack(side=LEFT, fill=BOTH, expand=True)

            # Show processing status
            status_label = ttk.Label(
                scroll_frame,
                text="Processing Technical Reports...",
                font=('Helvetica', 14, 'bold'),
                bootstyle="primary"
            )
            status_label.pack(pady=10, anchor="w")

            try:
                # Process the files to get the data (don't save to new file)
                _, all_data = extract_technical_report_data(pdf_files, save_excel=False)

                try:
                    # Read existing Excel file
                    existing_df = pd.read_excel(excel_file)
                    
                    # Create DataFrame from new data
                    new_df = pd.DataFrame(all_data)
                    
                    # Simply append new data to existing data
                    combined_df = pd.concat([existing_df, new_df], ignore_index=True)
                    
                    # Save back to the same Excel file
                    combined_df.to_excel(excel_file, index=False)
                    
                    # Update status label with append info
                    status_label.config(
                        text=f"✅ Successfully processed {len(all_data)} Technical Report(s)\nData added to: {excel_file}",
                        bootstyle="success"
                    )
                except Exception as excel_error:
                    raise Exception(f"Error updating Excel file: {str(excel_error)}")

                # Show data for each processed file
                for data in all_data:
                    for key, value in data.items():
                        if value and key != 'File Name':
                            label = ttk.Label(
                                scroll_frame,
                                text=f"{key}: {value}",
                                wraplength=700
                            )
                            label.pack(pady=2, anchor="w")

                    separator = ttk.Label(scroll_frame, text="-" * 100)
                    separator.pack(pady=8)

                # Show Excel file path
                excel_label = ttk.Label(
                    scroll_frame,
                    text=f"\n💾 Data saved to:\n{excel_file}",
                    bootstyle="success",
                    font=('Helvetica', 12),
                    wraplength=700
                )
                excel_label.pack(pady=10)

            except Exception as e:
                # Update status label to show error
                status_label.config(
                    text="❌ Error processing files",
                    bootstyle="danger"
                )
                
                # Show error message
                error_label = ttk.Label(
                    scroll_frame,
                    text=f"Error: {str(e)}",
                    bootstyle="danger",
                    wraplength=700
                )
                error_label.pack(pady=10)

            # Create scrollable frame
            canvas = ttk.Canvas(result_window)
            scrollbar = ttk.Scrollbar(result_window, orient='vertical', command=canvas.yview)
            scroll_frame = ttk.Frame(canvas)

            scroll_frame.bind(
                '<Configure>',
                lambda e: canvas.configure(scrollregion=canvas.bbox('all'))
            )

            canvas.create_window((0, 0), window=scroll_frame, anchor='nw')
            canvas.configure(yscrollcommand=scrollbar.set)
            canvas.pack(side='left', fill='both', expand=True)
            scrollbar.pack(side='right', fill='y')

            success_label = ttk.Label(
                scroll_frame,
                text=f'\n✅ Data saved to Excel file:\n{excel_file}',
                bootstyle='success',
                font=('Helvetica', 12, 'bold'),
                wraplength=700
            )
            success_label.pack(pady=20)

        except Exception as e:
            Messagebox.show_error('Processing Error', f'An error occurred:\n{str(e)}')

    def gui_execute(self):
        root = ttk.Window(themename='flatly')
        root.title('Purchase Order Automation')

        style = ttk.Style()
        style.configure('success.Outline.TButton', font=('Helvetica', 11))
        style.configure('primary.TButton', font=('Helvetica', 13))

        # Title
        label = ttk.Label(
            root,
            text='Purchase Order Automation',
            font=('Helvetica', 18)
        )
        label.pack(pady=(20, 0))

        # Create notebook for tabs
        notebook = ttk.Notebook(root)
        notebook.pack(fill='both', expand=True, padx=20, pady=20)

        # PO Extractor tab
        po_tab = ttk.Frame(notebook)
        notebook.add(po_tab, text='Purchase Order')
        self.setup_po_extractor_tab(po_tab)

        # Product Order tab
        product_tab = ttk.Frame(notebook)
        notebook.add(product_tab, text='Product Order')
        self.setup_product_order_tab(product_tab)

        # Technical Report tab
        tech_tab = ttk.Frame(notebook)
        notebook.add(tech_tab, text='Technical Report')
        self.setup_technical_report_tab(tech_tab)

        root.geometry('700x600')

        try:
            cond_AT = requests.get("https://saim2481.pythonanywhere.com/ATactivation-desktop-response/")
            cond_AT.raise_for_status()
            cond_AT = cond_AT.text
        except requests.exceptions.RequestException:
            cond_AT = False
            Messagebox.show_error("Connection Error", "Please Check your internet Connection")
        except:
            cond_AT = False
            Messagebox.show_error("Something Went Wrong", "Unexpected Error")

        if cond_AT != "true":
            self.po_execute_button.configure(state=DISABLED)
            self.product_execute_button.configure(state=DISABLED)
            self.tech_execute_button.configure(state=DISABLED)

        root.mainloop()


if __name__ == "__main__":
    try:
        extractor = PDFExtractor()
        extractor.gui_execute()
    except Exception as e:
        print(f"Fatal error: {str(e)}")
        while True:
            pass