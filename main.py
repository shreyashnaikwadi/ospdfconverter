import os
import PyPDF2
from tkinter import *
from tkinter import filedialog
from docx2pdf import convert as convert_to_pdf
from img2pdf import convert as image_to_pdf
from tkinter import messagebox

def open_file(filetypes):
    file_path = filedialog.askopenfilename(filetypes=filetypes)
    return file_path

def save_file(filetypes):
    file_path = filedialog.asksaveasfilename(defaultextension=filetypes[0][1], filetypes=filetypes)
    return file_path

def convert_pdf_to_word():
    input_file = open_file([("PDF Files", "*.pdf")])
    if input_file:
        output_file = save_file([("Word Files", "*.docx")])
        if output_file:
            # Code to convert PDF to Word using external library here
            messagebox.showinfo("Conversion Complete", "PDF to Word conversion is complete!")

def convert_word_to_pdf():
    input_file = open_file([("Word Files", "*.docx")])
    if input_file:
        output_file = save_file([("PDF Files", "*.pdf")])
        if output_file:
            # Code to convert Word to PDF using external library here
            convert_to_pdf(input_file, output_file)
            messagebox.showinfo("Conversion Complete", "Word to PDF conversion is complete!")

def convert_image_to_pdf():
    input_file = open_file([("Image Files", "*.png;*.jpg;*.jpeg")])
    if input_file:
        output_file = save_file([("PDF Files", "*.pdf")])
        if output_file:
            # Code to convert Image to PDF using external library here
            with open(output_file, "wb") as f:
                f.write(image_to_pdf([input_file]))
            messagebox.showinfo("Conversion Complete", "Image to PDF conversion is complete!")


# Create the main application window
root = Tk()
root.title("PDF Converter")
root.geometry("500x500")  # Set the window size to 500x500

# Set the background color of the main window
root.configure(bg='black')

# Add buttons for each functionality
convert_pdf_to_word_button = Button(root, text="PDF to Word", command=convert_pdf_to_word, padx=10, pady=5, bg='blue', fg='white')
convert_word_to_pdf_button = Button(root, text="Word to PDF", command=convert_word_to_pdf, padx=10, pady=5, bg='blue', fg='white')
convert_image_to_pdf_button = Button(root, text="Image to PDF", command=convert_image_to_pdf, padx=10, pady=5, bg='blue', fg='white')

convert_pdf_to_word_button.pack(pady=10)
convert_word_to_pdf_button.pack(pady=10)
convert_image_to_pdf_button.pack(pady=10)


root.mainloop()
