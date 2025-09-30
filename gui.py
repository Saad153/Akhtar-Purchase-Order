try:
    import tkinter
    import customtkinter
    from tkinter import messagebox, filedialog
    from automation import extract_po_data
    import requests

    # Keep the same theme and appearance
    customtkinter.set_appearance_mode("light")
    customtkinter.set_default_color_theme('blue')

    def start_processing():
        pdf_files = entry_pdf.get()

        if pdf_files == '':
            messagebox.showerror('Error', 'Please select one or more PDF files')
            return

        pdf_files = pdf_files.split(";")  # split multiple file paths

        try:
            # Create result window
            result_window = customtkinter.CTkToplevel()
            result_window.title('Extracted Data')
            result_window.geometry("500x600")

            # Scrollable frame
            scroll_frame = customtkinter.CTkScrollableFrame(result_window, width=480, height=580)
            scroll_frame.pack(pady=10, padx=10)

            for pdf_file in pdf_files:
                pdf_file = pdf_file.strip()
                if not pdf_file:
                    continue

                try:
                    data = extract_po_data(pdf_file)

                    # File title
                    file_label = customtkinter.CTkLabel(
                        scroll_frame,
                        text=f"📄 {pdf_file}",
                        font=('Helvetica', 14, 'bold'),
                        wraplength=460
                    )
                    file_label.pack(pady=(10, 5), anchor="w")

                    # Fields inside this file
                    for key, value in data.items():
                        label = customtkinter.CTkLabel(
                            scroll_frame,
                            text=f"{key}: {value}",
                            wraplength=460
                        )
                        label.pack(pady=2, anchor="w")

                    # Separator
                    separator = customtkinter.CTkLabel(scroll_frame, text="-" * 80)
                    separator.pack(pady=8, anchor="w")

                except Exception as e:
                    error_label = customtkinter.CTkLabel(
                        scroll_frame,
                        text=f"Error processing {pdf_file}: {str(e)}",
                        text_color="red",
                        wraplength=460
                    )
                    error_label.pack(pady=5, anchor="w")

        except Exception as e:
            messagebox.showerror("Processing Error", f"An error occurred:\n{str(e)}")

    def browse_files(entry_widget):
        file_paths = filedialog.askopenfilenames(filetypes=[("PDF files", "*.pdf")])
        if file_paths:
            entry_widget.delete(0, 'end')
            entry_widget.insert(0, "; ".join(file_paths))  # show all selected files

    # --- GUI Layout ---
    root = customtkinter.CTk()
    root.title('PDF PO Data Extractor')

    label = customtkinter.CTkLabel(
        master=root,
        text="PDF Purchase Order Extractor",
        font=('Helvetica', 20)
    )
    label.place(relx=0.5, rely=0.1, anchor=tkinter.N)

    # Entry for PDF Files
    entry_pdf = customtkinter.CTkEntry(
        master=root,
        width=300,
        height=25,
        placeholder_text="Select PDF Files"
    )
    entry_pdf.place(relx=0.38, rely=0.2, anchor=tkinter.N)

    browse_pdf_button = customtkinter.CTkButton(
        master=root,
        text="Browse",
        command=lambda: browse_files(entry_pdf),
        width=120,
        height=25,
        border_width=0,
        corner_radius=8
    )
    browse_pdf_button.place(relx=0.82, rely=0.2, anchor=tkinter.N)

    # Execute Button
    execute_button = customtkinter.CTkButton(
        master=root,
        text="Process PDFs",
        command=start_processing,
        width=430,
        height=25,
        border_width=0,
        corner_radius=8
    )
    execute_button.place(relx=0.51, rely=0.33, anchor=tkinter.N)

    root.geometry("500x600")

    # --- Connection Check ---
    try:
        cond_AT = requests.get("https://saim2481.pythonanywhere.com/ATactivation-desktop-response/")
        cond_AT.raise_for_status()
        cond_AT = cond_AT.text
    except requests.exceptions.RequestException:
        cond_AT = False
        messagebox.showerror("Connection Error", "Please Check your internet Connection")
    except:
        cond_AT = False
        messagebox.showerror("Something Went Wrong", "Unexpected Error")

    print(cond_AT)
    if cond_AT != "true":
        execute_button.configure(state=customtkinter.DISABLED)

    root.mainloop()
except Exception as e:
    print(e)
    while True:
        pass
