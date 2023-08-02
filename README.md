# PDF to DOCX Converter Documentation

**Description**
This documentation outlines the usage and functionality of the PDF to DOCX Converter, a Python script that provides a graphical user interface (GUI) to convert a PDF file to a DOCX file. The script utilizes the `PyPDF2` library to read the PDF file and extract its text content. It then uses the `docx` library to create a new Word document and adds the extracted text content to it. The user can select the input PDF file to convert and specify the output DOCX file's name and location using a file dialog.

**Requirements**
To run the PDF to DOCX converter, ensure the following requirements are met:

1. Python 3.x installed on your system.
2. The `PyPDF2` and `python-docx` libraries installed. You can install them using `pip`:
   ```
   pip install PyPDF2 python-docx
   ```

**Usage**
To use the PDF to DOCX converter, follow these steps:

1. Import the required libraries:
   ```python
   import os
   import PyPDF2
   from docx import Document
   import tkinter as tk
   from tkinter import filedialog, messagebox
   ```

2. Define the `pdf_to_docx` function to convert the PDF file to a DOCX file:
   ```python
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
   ```

3. Create event handlers for mouse hover effect on the button:
   ```python
   def on_enter(event):
       convert_button["bg"] = "gray"

   def on_leave(event):
       convert_button["bg"] = "black"
   ```

4. Define the main conversion function `convert_pdf_to_docx()` which will be called when the "Convert PDF to DOCX" button is clicked:
   ```python
   def convert_pdf_to_docx():
       input_pdf_file = filedialog.askopenfilename(title="Select PDF File", filetypes=[("PDF Files", "*.pdf")])

       if input_pdf_file:
           output_docx_file = filedialog.asksaveasfilename(title="Save DOCX File", defaultextension=".docx", filetypes=[("Word Documents", "*.docx")])

           if output_docx_file:
               try:
                   pdf_to_docx(input_pdf_file, output_docx_file)
               except Exception as e:
                   messagebox.showerror("Conversion Error", f"An error occurred during conversion: {str(e)}")
   ```

5. Create the main application window using `tk.Tk()` and configure its appearance:
   ```python
   root = tk.Tk()
   root.title("PDF to DOCX Converter")
   root.geometry("800x400")
   root.configure(bg="black")
   ```

6. Customize the font size and appearance for the "Convert PDF to DOCX" button:
   ```python
   font = ("Times New Roman", 40)
   button_font_color = "red"
   ```

7. Add the "Convert PDF to DOCX" button to the application window:
   ```python
   convert_button = tk.Button(root, text="Convert PDF to DOCX", font=font, fg=button_font_color, bg="black", command=convert_pdf_to_docx, anchor='center')
   convert_button.pack(pady=50)
   ```

8. Center the button on the screen:
   ```python
   root.update_idletasks()
   x = (root.winfo_screenwidth() - root.winfo_reqwidth()) // 2
   y = (root.winfo_screenheight() - root.winfo_reqheight()) // 2
   root.geometry("+{}+{}".format(x, y))
   ```

9. Add mouse hover effect to the button using the previously defined event handlers:
   ```python
   convert_button.bind("<Enter>", on_enter)
   convert_button.bind("<Leave>", on_leave)
   ```

10. Start the main event loop to run the application:
    ```python
    root.mainloop()
    ```

**How to Run**
To use the PDF to DOCX converter, follow these steps:

1. Ensure you have Python 3.x and the required libraries installed.
2. Save the script to a file with a `.py` extension (e.g., `pdf_to_docx_converter.py`).
3. Run the script using the Python interpreter:
   ```
   python pdf_to_docx_converter.py
   ```
4. The application window will open, allowing you to select the input PDF file and specify the output DOCX file's name and location. After conversion, a success message will be displayed.

**Limitations**
- The script only extracts the text content from the PDF file. Complex formatting, images, and other non-text elements in the PDF will not be preserved in the resulting DOCX file.
- In case of password-protected or encrypted PDF files, the script may raise exceptions or produce incorrect results.
- The script may not handle extremely large PDF files efficiently due to the limitations of the `PyPDF2` library.

**Disclaimer**
This script is provided as-is without any warranties. Use it at your own risk. The author is not responsible for any data loss or damage resulting from the use of this script. Always keep backups of your important files before performing any conversions.
