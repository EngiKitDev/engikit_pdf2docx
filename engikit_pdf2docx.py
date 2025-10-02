import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pdf2docx import Converter
import os
import threading
import time

def convert_pdf_to_docx(pdf_path, docx_path, start_page, end_page, log_callback):
    try:
        log_callback(f"Starting conversion: {os.path.basename(pdf_path)}")
        start_time = time.time()
        cv = Converter(pdf_path)

        log_callback(f"Converting pages from {start_page} to {end_page if end_page else 'last'}...")

        def progress(page_number, total_pages):
            log_callback(f"Converting page {page_number + 1} of {total_pages}")

        cv.convert(docx_path, start=start_page - 1, end=end_page, progress=progress)
        cv.close()

        duration = time.time() - start_time
        log_callback(f"✅ Conversion finished successfully in {duration:.2f} seconds.")
        messagebox.showinfo("Success", f"File has been saved as:\n{docx_path}")
    except Exception as e:
        log_callback(f"❌ Error: {str(e)}")
        messagebox.showerror("Error", f"An error occurred during conversion:\n{str(e)}")

def start_conversion():
    pdf_path = pdf_entry.get()
    output_folder = output_entry.get()
    same_folder = same_folder_var.get()
    all_pages = all_pages_var.get()
    start_page = start_page_entry.get()
    end_page = end_page_entry.get()

    if not pdf_path:
        messagebox.showwarning("Missing data", "Please select a PDF file.")
        return

    if same_folder:
        output_folder = os.path.dirname(pdf_path)

    if not output_folder:
        messagebox.showwarning("Missing data", "Please select an output folder.")
        return

    try:
        if all_pages:
            start = 1
            end = None
        else:
            start = int(start_page)
            end = int(end_page) if end_page else None
    except ValueError:
        messagebox.showerror("Error", "Page range must be an integer.")
        return

    pdf_filename = os.path.basename(pdf_path)
    docx_filename = os.path.splitext(pdf_filename)[0] + "_converted.docx"
    docx_path = os.path.join(output_folder, docx_filename)

    threading.Thread(target=convert_pdf_to_docx, args=(pdf_path, docx_path, start, end, log_message)).start()

def browse_pdf():
    file_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
    if file_path:
        pdf_entry.delete(0, tk.END)
        pdf_entry.insert(0, file_path)
        if same_folder_var.get():
            output_entry.delete(0, tk.END)
            output_entry.insert(0, os.path.dirname(file_path))

def browse_output_folder():
    folder_path = filedialog.askdirectory()
    if folder_path:
        output_entry.delete(0, tk.END)
        output_entry.insert(0, folder_path)

def toggle_same_folder():
    if same_folder_var.get():
        pdf_path = pdf_entry.get()
        if pdf_path:
            output_entry.delete(0, tk.END)
            output_entry.insert(0, os.path.dirname(pdf_path))

def toggle_all_pages():
    state = 'disabled' if all_pages_var.get() else 'normal'
    start_page_entry.configure(state=state)
    end_page_entry.configure(state=state)

def log_message(message):
    log_text.configure(state='normal')
    log_text.insert(tk.END, message + "\n")
    log_text.see(tk.END)
    log_text.configure(state='disabled')

# GUI setup
root = tk.Tk()
root.title("PDF ➜ DOCX Converter (1:1)")

frame = ttk.Frame(root, padding="10")
frame.grid(row=0, column=0, sticky="nsew")

ttk.Label(frame, text="PDF File:").grid(row=0, column=0, sticky="w")
pdf_entry = ttk.Entry(frame, width=60)
pdf_entry.grid(row=0, column=1, padx=5)
ttk.Button(frame, text="Browse...", command=browse_pdf).grid(row=0, column=2)

ttk.Label(frame, text="Output Folder:").grid(row=1, column=0, sticky="w")
output_entry = ttk.Entry(frame, width=60)
output_entry.grid(row=1, column=1, padx=5)
ttk.Button(frame, text="Browse...", command=browse_output_folder).grid(row=1, column=2)

same_folder_var = tk.BooleanVar(value=True)
same_folder_check = ttk.Checkbutton(frame, text="Use same folder as PDF", variable=same_folder_var, command=toggle_same_folder)
same_folder_check.grid(row=2, column=1, sticky="w", pady=(0, 10))

ttk.Label(frame, text="Page Range:").grid(row=3, column=0, columnspan=3, sticky="w")
all_pages_var = tk.BooleanVar(value=True)
all_pages_check = ttk.Checkbutton(frame, text="All pages", variable=all_pages_var, command=toggle_all_pages)
all_pages_check.grid(row=4, column=0, columnspan=3, sticky="w")

ttk.Label(frame, text="From page:").grid(row=5, column=0, sticky="e")
start_page_entry = ttk.Entry(frame, width=10, state='disabled')
start_page_entry.grid(row=5, column=1, sticky="w")
ttk.Label(frame, text="To page:").grid(row=5, column=1, padx=(100, 0), sticky="w")
end_page_entry = ttk.Entry(frame, width=10, state='disabled')
end_page_entry.grid(row=5, column=1, padx=(160, 0), sticky="w")

ttk.Button(frame, text="Start Conversion", command=start_conversion).grid(row=6, column=0, columnspan=3, pady=10)

ttk.Label(frame, text="Conversion Log:").grid(row=7, column=0, columnspan=3, sticky="w")
log_text = tk.Text(frame, height=10, width=80, state='disabled')
log_text.grid(row=8, column=0, columnspan=3, pady=(0, 10))

root.mainloop()
