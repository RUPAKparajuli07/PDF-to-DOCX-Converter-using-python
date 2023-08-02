import os
import PyPDF2
from docx import Document
import tkinter as tk
from tkinter import filedialog, messagebox

def pdf_to_docx(input_pdf_path, output_docx_path):
    doc = Document()
    
    with open(input_pdf_path, 'rb') as pdf_file:
        pdf_reader = PyPDF2.PdfFileReader(pdf_file)
        num_pages = pdf_reader.numPages

        for page_num in range(num_pages):
            page = pdf_reader.getPage(page_num)
            text = page.extractText()
            doc.add_paragraph(text)

    doc.save(output_docx_path)
    messagebox.showinfo("Conversion Successful", f"PDF converted to DOCX. Output saved to '{output_docx_path}'.")

def on_enter(event):
    convert_button["bg"] = "gray"

def on_leave(event):
    convert_button["bg"] = "black"

def convert_pdf_to_docx():
    input_pdf_file = filedialog.askopenfilename(title="Select PDF File", filetypes=[("PDF Files", "*.pdf")])

    if input_pdf_file:
        output_docx_file = filedialog.asksaveasfilename(title="Save DOCX File", defaultextension=".docx", filetypes=[("Word Documents", "*.docx")])

        if output_docx_file:
            try:
                pdf_to_docx(input_pdf_file, output_docx_file)
            except Exception as e:
                messagebox.showerror("Conversion Error", f"An error occurred during conversion: {str(e)}")

# Create the main application window
root = tk.Tk()
root.title("PDF to DOCX Converter")
root.geometry("800x400")
root.configure(bg="black")

# Customize the font size and appearance
font = ("Times New Roman", 40)
button_font_color = "red"

# Add a button to trigger the conversion process
convert_button = tk.Button(root, text="Convert PDF to DOCX", font=font, fg=button_font_color, bg="black", command=convert_pdf_to_docx, anchor='center')
convert_button.pack(pady=50)

# Center the button on the screen
root.update_idletasks()
x = (root.winfo_screenwidth() - root.winfo_reqwidth()) // 2
y = (root.winfo_screenheight() - root.winfo_reqheight()) // 2
root.geometry("+{}+{}".format(x, y))

# Add mouse hover effect
convert_button.bind("<Enter>", on_enter)
convert_button.bind("<Leave>", on_leave)

root.mainloop()
