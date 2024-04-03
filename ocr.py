from tkinter import *
import tkinter as tk  
from tkinter import ttk  
import pytesseract
import pandas as pd
import os
from tkinter import filedialog
from fpdf import FPDF
from docx import Document
import cv2
import io
import openpyxl  # Import openpyxl

# Global variables
output_directory = None
save_dialog = None

def browseFiles():
    filename = filedialog.askopenfilename(initialdir="/",
                                          title="Select a File",
                                          filetypes=(("Image files", "*.png;*.jpg;*.jpeg;*.gif"),
                                                     ("Excel files", "*.xlsx;*.xls"),
                                                     ("all files", "*.*"))) 

    if filename:
        _, file_extension = os.path.splitext(filename.lower())

        if file_extension in ['.png', '.jpg', '.jpeg', '.gif']:
            process_image_file(filename)
        elif file_extension in ['.xlsx', '.xls']:
            process_excel_file(filename)
        else:
            print(f"Unsupported file format: {file_extension}")

def process_image_file(filename):
    # Change label contents
    img = cv2.imread(filename)
    cv2.imshow("Input image", img)
    pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files (x86)\Tesseract-OCR\tesseract'
    s = pytesseract.image_to_string(img)

    # Insert The Fact.
    T.insert(tk.END, s)

    save_options_dialog(s)

def process_excel_file(filename):
    try:
        df = pd.read_excel(filename)
        text = df.to_string(index=False)
        T.insert(tk.END, text)
        
        save_options_dialog(text)

    except Exception as e:
        print(f"Error reading Excel file: {e}")

def save_options_dialog(text):
    global save_dialog  # Make save_dialog global
    save_dialog = tk.Toplevel(root)
    save_dialog.title("Save Options")

    # Label for save options
    save_label = tk.Label(save_dialog, text="Choose Save Format:")
    save_label.pack()

    # Radio buttons for save options
    save_format = tk.StringVar()
    save_format.set("text")  # Default to text format

    text_radio = tk.Radiobutton(save_dialog, text="Text File", variable=save_format, value="text")
    pdf_radio = tk.Radiobutton(save_dialog, text="PDF", variable=save_format, value="pdf")
    docx_radio = tk.Radiobutton(save_dialog, text="DOCX", variable=save_format, value="docx")
    excel_radio = tk.Radiobutton(save_dialog, text="Excel", variable=save_format, value="excel")

    text_radio.pack()
    pdf_radio.pack()
    docx_radio.pack()
    excel_radio.pack()

    # Button to choose output directory
    output_dir_button = tk.Button(save_dialog, text="Choose Output Directory", command=choose_output_directory)
    output_dir_button.pack()

    # Button to save
    save_button = tk.Button(save_dialog, text="Save", command=lambda: save_text(text, save_format.get()))
    save_button.pack()

def choose_output_directory():
    global output_directory
    output_directory = filedialog.askdirectory()

def save_text_to_excel(text, filename):
    try:
        df = pd.read_csv(io.StringIO(text), delim_whitespace=True)
        df.to_excel(filename, index=False, engine='openpyxl')  # Specify the engine as openpyxl
        print(f"\n\nTable data saved to {filename} in the chosen directory...!" if output_directory else
              f"\n\nTable data saved to {filename} in the current working directory...!")

    except Exception as e:
        print(f"Error saving Excel file: {e}")

def save_text(text, save_format):
    if save_format == "text":
        save_text_to_file(text, "Text.txt")
    elif save_format == "pdf":
        save_text_to_pdf(text, "Text.pdf")
    elif save_format == "docx":
        save_text_to_docx(text, "Text.docx")
    elif save_format == "excel":
        save_text_to_excel(text, "TableData.xlsx")

def save_text_to_file(text, filename):
    if output_directory:
        filepath = os.path.join(output_directory, filename)
    else:
        filepath = filename

    with open(filepath, "w") as f:
        f.write(text)
    
    print(f"\n\nText saved to {filename} in the chosen directory...!" if output_directory else
          f"\n\nText saved to {filename} in the current working directory...!")

    # Close save options dialog
    save_dialog.destroy()

def save_text_to_pdf(text, filename):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=12)
    pdf.multi_cell(0, 10, text)
    
    if output_directory:
        filepath = os.path.join(output_directory, filename)
    else:
        filepath = filename

    pdf.output(filepath)
    
    print(f"\n\nText saved to {filename} in the chosen directory...!" if output_directory else
          f"\n\nText saved to {filename} in the current working directory...!")

    # Close save options dialog
    save_dialog.destroy()

def save_text_to_docx(text, filename):
    doc = Document()
    doc.add_paragraph(text)

    if output_directory:
        filepath = os.path.join(output_directory, filename)
    else:
        filepath = filename

    doc.save(filepath)

    print(f"\n\nText saved to {filename} in the chosen directory...!" if output_directory else
          f"\n\nText saved to {filename} in the current working directory...!")

    # Close save options dialog
    save_dialog.destroy()

# Main Tkinter window
root = Tk()
root.title("Text OCR")

# specify size of window.
root.geometry("500x500")
root.config(background="black")
# Create text widget and specify size.
T = Text(root, height=20, width=52)

# Create label
l = Label(root, text="Text Converter")
l.config(font=("Courier", 14))

# Create button for next text.
b1 = Button(root, text="Browse", command=browseFiles)

# Create an Exit button.
b2 = Button(root, text="Exit", command=root.destroy)

l.pack()
T.pack()
b1.pack()
b2.pack()

tk.mainloop()
