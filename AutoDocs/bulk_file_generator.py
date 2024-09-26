import os
import tkinter as tk
from tkinter import filedialog
from PyPDF2 import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.colors import white
from reportlab.lib.units import inch
from io import BytesIO

def select_files():
    files_selected = filedialog.askopenfilenames(filetypes=[("PDF Files", "*.pdf")])
    files_list.set(root.tk.splitlist(files_selected))

def select_output_folder():
    folder_selected = filedialog.askdirectory()
    output_folder.set(folder_selected)

def run_merge():
    files = files_list.get()
    start_delim = start_delimiter.get()
    end_delim = end_delimiter.get()
    output_filename = output_file.get()
    output_dir = output_folder.get()
    if not files or not start_delim or not end_delim or not output_filename or not output_dir:
        print("Please fill all fields")
        return
    if not output_filename.lower().endswith('.pdf'):
        output_filename += '.pdf'
    output_path = os.path.join(output_dir, output_filename)
    merge_pdfs(files, start_delim, end_delim, output_path)

def merge_pdfs(files, start_delim, end_delim, output_path):
    pdf_writer = PdfWriter()
    # files is a tuple of file paths
    files = list(files)
    files.sort()  # Optional: sort files alphabetically
    for filepath in files:
        filename = os.path.basename(filepath)
        investor_code = filename.split('_')[0]
        full_code = f"{start_delim}{investor_code}{end_delim}"
        pdf_reader = PdfReader(filepath)
        num_pages = len(pdf_reader.pages)
        for page_num in range(num_pages):
            page = pdf_reader.pages[page_num]
            width = float(page.mediabox.width)
            height = float(page.mediabox.height)
            packet = BytesIO()
            # Create a new PDF with ReportLab
            can = canvas.Canvas(packet, pagesize=(width, height))
            can.setFillColor(white)
            # Place the text at top-right corner
            margin = inch * 0.5  # Half an inch margin
            text_width = can.stringWidth(full_code)
            x = width - text_width - margin
            y = height - margin
            can.drawString(x, y, full_code)
            can.save()
            # Move to the beginning of the StringIO buffer
            packet.seek(0)
            overlay_pdf = PdfReader(packet)
            overlay_page = overlay_pdf.pages[0]
            page.merge_page(overlay_page)
            pdf_writer.add_page(page)
    with open(output_path, 'wb') as out_file:
        pdf_writer.write(out_file)
    print(f"Merged PDF saved as {output_path}")

# GUI Setup
root = tk.Tk()
root.title("PDF Merger")

files_list = tk.Variable(value=())
start_delimiter = tk.StringVar()
end_delimiter = tk.StringVar()
output_file = tk.StringVar()
output_folder = tk.StringVar()


tk.Label(root, text="Select PDF Files:").grid(row=0, column=0, padx=5, pady=5, sticky='e')
tk.Entry(root, textvariable=files_list, width=40).grid(row=0, column=1, padx=5, pady=5)
tk.Button(root, text="Browse", command=select_files).grid(row=0, column=2, padx=5, pady=5)

tk.Label(root, text="Start Delimiter:").grid(row=1, column=0, padx=5, pady=5, sticky='e')
tk.Entry(root, textvariable=start_delimiter).grid(row=1, column=1, padx=5, pady=5)

tk.Label(root, text="End Delimiter:").grid(row=2, column=0, padx=5, pady=5, sticky='e')
tk.Entry(root, textvariable=end_delimiter).grid(row=2, column=1, padx=5, pady=5)

tk.Label(root, text="Output File Name:").grid(row=3, column=0, padx=5, pady=5, sticky='e')
tk.Entry(root, textvariable=output_file).grid(row=3, column=1, padx=5, pady=5)

tk.Label(root, text="Output Folder:").grid(row=4, column=0, padx=5, pady=5, sticky='e')
tk.Entry(root, textvariable=output_folder, width=40).grid(row=4, column=1, padx=5, pady=5)
tk.Button(root, text="Browse", command=select_output_folder).grid(row=4, column=2, padx=5, pady=5)

tk.Button(root, text="Merge PDFs", command=run_merge).grid(row=5, column=1, padx=5, pady=10)

root.mainloop()
