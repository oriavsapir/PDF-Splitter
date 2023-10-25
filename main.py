import tkinter as tk
from tkinter import filedialog

import PyPDF2
import ttkbootstrap as ttk


def toggle_entry_fields(*args):
    action = selected_action.get()
    if action == "Extract Range":
        start_label.grid()
        start_entry.grid()
        end_label.grid()
        end_entry.grid()
    else:
        start_label.grid_remove()
        start_entry.grid_remove()
        end_label.grid_remove()
        end_entry.grid_remove()


def update_visibility(*args):
    if selected_action.get() == 'Combine PDFs':
        input_pdf_label.grid_remove()
        input_pdf_entry.grid_remove()
        input_pdf_button.grid_remove()
    else:
        input_pdf_label.grid()
        input_pdf_entry.grid()
        input_pdf_button.grid()


def execute_action():
    action = selected_action.get()
    output_folder = output_folder_path.get()

    try:
        if action == 'Extract First Page':
            extract_page(0)
            status_label.config(text="Successfully extracted the first page.")
        elif action == 'Extract Last Page':
            extract_page(-1)
            status_label.config(text="Successfully extracted the last page.")
        elif action == 'Split Pages':
            split_pdf()
            status_label.config(text="Successfully split the PDF pages.")
        elif action == 'Combine PDFs':
            combine_pdfs(output_folder)
            status_label.config(text="Successfully combined the PDFs.")
        elif action == 'Extract Range':
            start = int(start_num.get())
            end = int(end_num.get())
            extract_range(start, end)
            status_label.config(text="Successfully extracted the page range.")
    except Exception as e:
        status_label.config(text=f"Error: {str(e)}")


def extract_page(page_num):
    file_path = input_pdf_path.get()
    output_folder = output_folder_path.get()

    with open(file_path, 'rb') as pdf_file:
        reader = PyPDF2.PdfReader(pdf_file)
        pdf_writer = PyPDF2.PdfWriter()

        if page_num == -1:
            page_num = len(reader.pages) - 1

        pdf_writer.add_page(reader.pages[page_num])
        output_file_path = f"{output_folder}/extracted_page_{page_num + 1}.pdf"

        with open(output_file_path, 'wb') as output_pdf:
            pdf_writer.write(output_pdf)

        print(f"Created: {output_file_path}")


def split_pdf():
    file_path = input_pdf_path.get()
    output_folder = output_folder_path.get()

    with open(file_path, 'rb') as pdf_file:
        reader = PyPDF2.PdfReader(pdf_file)
        total_pages = len(reader.pages)

        for page_num in range(total_pages):
            pdf_writer = PyPDF2.PdfWriter()
            pdf_writer.add_page(reader.pages[page_num])
            output_file_path = f"{output_folder}/page_{page_num + 1}.pdf"

            with open(output_file_path, 'wb') as output_pdf:
                pdf_writer.write(output_pdf)

            print(f"Created: {output_file_path}")


def extract_range(start, end):
    file_path = input_pdf_path.get()
    output_folder = output_folder_path.get()

    with open(file_path, 'rb') as pdf_file:
        reader = PyPDF2.PdfReader(pdf_file)
        pdf_writer = PyPDF2.PdfWriter()

        for i in range(start, end + 1):
            pdf_writer.add_page(reader.pages[i])

        output_file_path = f"{output_folder}/extracted_pages_{start}_to_{end}.pdf"

        with open(output_file_path, 'wb') as output_pdf:
            pdf_writer.write(output_pdf)

        print(f"Created: {output_file_path}")


def combine_pdfs(output_folder):
    files = filedialog.askopenfilenames(filetypes=[("PDF files", "*.pdf")])

    pdf_writer = PyPDF2.PdfWriter()

    for file_path in files:
        pdf_reader = PyPDF2.PdfReader(file_path)
        for page_num in range(len(pdf_reader.pages)):
            page = pdf_reader.pages[page_num]
            pdf_writer.add_page(page)

    output_file_path = f"{output_folder}/combined.pdf"

    with open(output_file_path, 'wb') as output_pdf:
        pdf_writer.write(output_pdf)

    print(f"Created: {output_file_path}")


def browse_pdf():
    file_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
    input_pdf_path.set(file_path)


def browse_folder():
    folder_path = filedialog.askdirectory()
    output_folder_path.set(folder_path)


# GUI
# Initialize the themed Tk instance
root = ttk.Window("Data Entry", "superhero")

root.title("PDF Operations")

frame = ttk.Frame(root, padding="20 20 20 20")  # Added padding
frame.grid(row=0, column=0)

input_pdf_path = tk.StringVar()
output_folder_path = tk.StringVar()
selected_action = tk.StringVar()
selected_action.set('Split Pages')  # Default value
start_num = tk.StringVar()
end_num = tk.StringVar()

selected_action.trace_add("write", toggle_entry_fields)
start_label = ttk.Label(frame, text="Start Page:")
start_entry = ttk.Entry(frame, textvariable=start_num, width=10)

end_label = ttk.Label(frame, text="End Page:")
end_entry = ttk.Entry(frame, textvariable=end_num, width=10)
start_label.grid(row=5, column=0, sticky=tk.W, padx=5, pady=5)
start_entry.grid(row=5, column=1, padx=5, pady=5)

end_label.grid(row=6, column=0, sticky=tk.W, padx=5, pady=5)
end_entry.grid(row=6, column=1, padx=5, pady=5)

# Create and place widgets
input_pdf_label = ttk.Label(frame, text="Input PDF Path:")
input_pdf_entry = ttk.Entry(frame, textvariable=input_pdf_path, width=40)
input_pdf_button = ttk.Button(frame, text="Browse", command=browse_pdf)

input_pdf_label.grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
input_pdf_entry.grid(row=0, column=1, padx=5, pady=5)
input_pdf_button.grid(row=0, column=2, padx=5, pady=5)

ttk.Label(frame, text="Output Folder Path:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
ttk.Entry(frame, textvariable=output_folder_path, width=40).grid(row=1, column=1, padx=5, pady=5)
ttk.Button(frame, text="Browse", command=browse_folder).grid(row=1, column=2, padx=5, pady=5)

ttk.Label(frame, text="Action:").grid(row=2, column=0, sticky=tk.W, padx=5, pady=5)
option_menu = ttk.OptionMenu(frame, selected_action, 'Choose actions', 'Split Pages', 'Extract First Page',
                             'Extract Last Page', 'Combine PDFs', 'Extract Range')
option_menu.grid(row=2, column=1, sticky=tk.W, padx=5, pady=5)

ttk.Button(frame, text="Execute", command=execute_action).grid(row=3, columnspan=3, padx=5, pady=5)
status_label = ttk.Label(frame, text="")
status_label.grid(row=4, columnspan=3, padx=5, pady=5)
# Listen for changes in the OptionMenu
selected_action.trace_add('write', update_visibility)

root.mainloop()
