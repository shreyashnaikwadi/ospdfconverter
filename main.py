import os
import io
import PyPDF2
import pandas as pd
import comtypes.client 
from tkinter import *
from tkinter import filedialog
from docx2pdf import convert as convert_to_pdf
from img2pdf import convert as image_to_pdf
from tkinter import messagebox
import openpyxl
from pptx import Presentation
from PIL import Image, ImageTk
from reportlab.pdfgen import canvas
from PyPDF2 import PdfReader, PdfWriter
import tkinter as tk
from tkinter import filedialog, simpledialog, messagebox
from openpyxl import Workbook
from fpdf import FPDF 
from tkinter import filedialog, Tk
import comtypes.client
import shutil
import time
from pdf2image import convert_from_path
from pdf2image import convert_from_bytes
from PyPDF2 import PdfReader
from tkinter import filedialog, messagebox
from pdf2docx import Converter

def open_file(filetypes):
    file_path = filedialog.askopenfilename(filetypes=filetypes)
    return file_path

def save_file(filetypes):
    file_path = filedialog.asksaveasfilename(defaultextension=filetypes[0][1], filetypes=filetypes)
    return file_path


# CONVERT PDF TO WORD
def convert_pdf_to_word():
    input_pdf = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
    if input_pdf:
        output_word = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word Files", "*.docx")])
        if output_word:
            try:
                # Convert PDF to Word using pdf2docx library
                cv = Converter(input_pdf)
                cv.convert(output_word, start=0, end=None)
                cv.close()
                messagebox.showinfo("Conversion Complete", "PDF to Word conversion is complete!")
            except Exception as e:
                messagebox.showerror("Error", f"An error occurred: {e}")
# CONVERT WORD TO PDF 
def convert_word_to_pdf():
    input_file = open_file([("Word Files", "*.docx")])
    if input_file:
        output_file = save_file([("PDF Files", "*.pdf")])
        if output_file:
            # Code to convert Word to PDF using external library here
            convert_to_pdf(input_file, output_file)
            messagebox.showinfo("Conversion Complete", "Word to PDF conversion is complete!")
# CONVERT IMAGE TO PDF
def convert_image_to_pdf():
    input_file = open_file([("Image Files", "*.png;*.jpg;*.jpeg")])
    if input_file:
        output_file = save_file([("PDF Files", "*.pdf")])
        if output_file:
            # Code to convert Image to PDF using external library here
            with open(output_file, "wb") as f:
                f.write(image_to_pdf([input_file]))
            messagebox.showinfo("Conversion Complete", "Image to PDF conversion is complete!")
# CONVERT PDF TO IMAGE


def pdf_to_image():
    pdf_file = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
    if pdf_file:
        try:
            output_directory = filedialog.askdirectory(title="Select Output Directory")
            if output_directory:
                convert_pdf_to_images(pdf_file, output_directory)
                messagebox.showinfo("Conversion Complete", "PDF to image conversion successful!")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")

def convert_pdf_to_images(pdf_file, output_directory):
    images = convert_from_path(pdf_file)

    for i, image in enumerate(images):
        image_path = os.path.join(output_directory, f"page_{i + 1}.png")
        image.save(image_path, "PNG")



# CONVERT  EXCEL TO PDF
def convert_excel_to_pdf():
    excel_file = open_file([("Excel Files", "*.xlsx")])
    if excel_file:
        try:
            output_path = save_file([("PDF Files", "*.pdf")])
            if output_path:
                #  a new PDF using fpdf
                pdf = FPDF()
                pdf.add_page()

                # Set font (e.g., Arial) and font size
                pdf.set_font("Arial", size=12)

                # Open the Excel file and read its content
                xls = openpyxl.load_workbook(excel_file)
                sheet = xls.active
                max_rows = sheet.max_row
                max_cols = sheet.max_column

                # Loop through the cells and add them to the PDF
                for row in range(1, max_rows + 1):
                    for col in range(1, max_cols + 1):
                        cell_value = sheet.cell(row=row, column=col).value
                        pdf.cell(200, 10, str(cell_value), ln=True)

                # Save the PDF
                pdf.output(output_path)
                messagebox.showinfo("Excel to PDF", "Excel successfully converted to PDF!")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")
# CONVERT PDF TO EXCEL 
def convert_pdf_to_excel(pdf_path, excel_path):
    pdf_text = ""
    with open(pdf_path, 'rb') as pdf_file:
        pdf_reader = PyPDF2.PdfReader(pdf_file)
        for page in pdf_reader.pages:
            pdf_text += page.extract_text()

    data = [line.split('\t') for line in pdf_text.split('\n')]

    df = pd.DataFrame(data)
    df.to_excel(excel_path, index=False, header=False)

    messagebox.showinfo("PDF to Excel", "PDF successfully converted to Excel!")
# Function to handle the PDF to Excel button
def convert_pdf_to_excel_callback():
    pdf_file = open_file([("PDF Files", "*.pdf")])
    if pdf_file:
        try:
            output_path = save_file([("Excel Files", "*.xlsx")])
            if output_path:
                convert_pdf_to_excel(pdf_file, output_path)
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")
# PROTECT PDF 
def protect_pdf(input_pdf, output_pdf, password):
    pdf_reader = PdfReader(input_pdf)
    pdf_writer = PdfWriter()

    for page in pdf_reader.pages:
        pdf_writer.add_page(page)

    pdf_writer.encrypt(user_pwd=password, owner_pwd=None, use_128bit=True)

    with open(output_pdf, "wb") as output_file:
        pdf_writer.write(output_file)
def protect_pdf_callback():
    pdf_file = open_file([("PDF Files", "*.pdf")])
    if pdf_file:
        try:
            password = simpledialog.askstring("Password Protection", "Enter password:")
            if password:
                output_path = save_file([("PDF Files", "*.pdf")])
                if output_path:
                    protect_pdf(pdf_file, output_path, password)
                    messagebox.showinfo("PDF Protected", "PDF file protected with password successfully!")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")
# ORGANIZE PDF BY DATE
def organize_pdfs(input_directory, output_directory):
    pdf_files = [f for f in os.listdir(input_directory) if f.endswith(".pdf")]

    for pdf_file in pdf_files:
        pdf_path = os.path.join(input_directory, pdf_file)
        creation_time = os.path.getctime(pdf_path)

        year = str(time.localtime(creation_time).tm_year)
        month = str(time.localtime(creation_time).tm_mon).zfill(2)  # Zero padding
        destination_dir = os.path.join(output_directory, year, month)

        os.makedirs(destination_dir, exist_ok=True)
        shutil.move(pdf_path, os.path.join(destination_dir, pdf_file))

    messagebox.showinfo("PDF Organization", "PDF files organized by date successfully!")
def organize_pdfs_callback():
    input_directory = filedialog.askdirectory(title="Select Input Directory")
    if input_directory:
        output_directory = filedialog.askdirectory(title="Select Output Directory")
        if output_directory:
            organize_pdfs(input_directory, output_directory)











# Create the main application window
root = Tk()
root.title(" OS PDF CONVERTER")
# Get the screen width and height
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()

# Set the geometry to full screen
root.geometry(f"{screen_width}x{screen_height}")


# Load the image to be used as the background
background_image = Image.open("E:\\New folder\\pxfuel.jpg")
background_photo = ImageTk.PhotoImage(background_image)

# Create a canvas to put the background image
canvas = Canvas(root, width=1500, height=800)
canvas.pack()
canvas.create_image(0, 0, anchor=NW, image=background_photo)

#  a title label 
title_label = Label(root, text="OS PDF CONVERTER", font=("Times New Roman", 36, "italic"), bg='black', fg='white')
title_label.place(x=550, y=50)


# Create a frame to group the buttons with a white background
frame_width = 400
frame_height = 300
button_frame = Frame(root, bg='white', width=frame_width, height=frame_height)
button_frame.place(x=(root.winfo_screenwidth() - frame_width) // 2, y=(root.winfo_screenheight() - frame_height) // 2)
button_frame.place(x=600, y=180)



# buttons for each functionality 
button_font = ("Arial", 14, "italic")
button_bg = 'White'
button_fg = 'black'
button_width = 16
button_height = 4

convert_pdf_to_word_button = Button(button_frame, text="PDF to Word", command=convert_pdf_to_word, padx=10, pady=5,width=button_width, height=button_height, bg=button_bg, fg=button_fg,font=("Helvetica", 12))

convert_word_to_pdf_button = Button(button_frame, text="Word to PDF", command=convert_word_to_pdf, padx=10, pady=5,width=button_width, height=button_height, bg=button_bg, fg=button_fg,font=("Helvetica", 12))

convert_image_to_pdf_button = Button(button_frame, text="Image to PDF", command=convert_image_to_pdf, padx=10, pady=5,width=button_width, height=button_height, bg=button_bg, fg=button_fg,font=("Helvetica", 12))

convert_excel_to_pdf_button = Button(button_frame, text="Excel to PDF", command=convert_excel_to_pdf, padx=10, pady=5,width=button_width, height=button_height, bg=button_bg, fg=button_fg,font=("Helvetica", 12))

convert_pdf_to_image_button = Button(button_frame, text="PDF to Image", command=pdf_to_image, padx=10, pady=5,width=button_width, height=button_height, bg=button_bg, fg=button_fg, font=("Helvetica", 12))

convert_pdf_to_excel_button = Button(button_frame, text="PDF to Excel", command=convert_pdf_to_excel_callback, padx=10, pady=5,width=button_width, height=button_height, bg=button_bg, fg=button_fg, font=("Helvetica", 12))

protect_pdf_button = Button(button_frame, text="Protect PDF", command=protect_pdf_callback, padx=10, pady=5,width=button_width, height=button_height, bg=button_bg, fg=button_fg, font=("Helvetica", 12))

organize_pdfs_button = Button(button_frame, text="Organize PDFs", command=organize_pdfs_callback, padx=10, pady=5,width=button_width, height=button_height, bg=button_bg, fg=button_fg, font=("Helvetica", 12))



convert_pdf_to_word_button.grid(row=0, column=0, padx=10, pady=10)
convert_word_to_pdf_button.grid(row=0, column=1, padx=10, pady=10)
convert_image_to_pdf_button.grid(row=1, column=0, padx=10, pady=10)
convert_excel_to_pdf_button.grid(row=4, column=0, padx=10, pady=10)
convert_pdf_to_image_button.grid(row=1, column=1, padx=10, pady=10)
convert_pdf_to_excel_button.grid(row=4, column=1, padx=10, pady=10)
protect_pdf_button.grid(row=5, column=0, padx=10, pady=10)
organize_pdfs_button.grid(row=5, column=1, padx=10, pady=10)






root.mainloop()




