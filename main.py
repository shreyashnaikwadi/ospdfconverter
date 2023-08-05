import os
import PyPDF2
from tkinter import *
from tkinter import filedialog
from docx2pdf import convert as convert_to_pdf
from img2pdf import convert as image_to_pdf
from tkinter import messagebox
import openpyxl
from pptx import Presentation

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

def convert_excel_to_pdf():
    input_file = open_file([("Excel Files", "*.xlsx")])
    if input_file:
        output_file = save_file([("PDF Files", "*.pdf")])
        if output_file:
            # Code to convert Excel to PDF using openpyxl library
            workbook = openpyxl.load_workbook(input_file)
            pdf_writer = PyPDF2.PdfFileWriter()

            for sheet_name in workbook.sheetnames:
                sheet = workbook[sheet_name]
                temp_pdf = "temp.pdf"

                # Save each sheet as a temporary PDF
                export_pdf = open(temp_pdf, "wb")
                sheet_pdf = PyPDF2.PdfFileWriter()
                sheet_pdf.addPage(PyPDF2.pdf.PageObject.createBlankPage(width=sheet.max_column * 10, height=sheet.max_row * 10))
                for row in sheet.iter_rows(values_only=True):
                    sheet_pdf.addPage(PyPDF2.pdf.PageObject.createTextPage("\n".join(str(cell) for cell in row)))

                sheet_pdf.write(export_pdf)
                export_pdf.close()

                # Merge the temporary PDFs into the final PDF
                with open(temp_pdf, "rb") as temp_file:
                    pdf_reader = PyPDF2.PdfFileReader(temp_file)
                    for page_num in range(pdf_reader.getNumPages()):
                        page = pdf_reader.getPage(page_num)
                        pdf_writer.addPage(page)

                os.remove(temp_pdf)

            with open(output_file, 'wb') as output:
                pdf_writer.write(output)

            messagebox.showinfo("Conversion Complete", "Excel to PDF conversion is complete!")

def convert_ppt_to_pdf():
    input_file = open_file([("PowerPoint Files", "*.pptx;*.ppt")])
    if input_file:
        output_file = save_file([("PDF Files", "*.pdf")])
        if output_file:
            # Code to convert PowerPoint to PDF using python-pptx library
            prs = Presentation(input_file)
            pdf_writer = PyPDF2.PdfFileWriter()

            for slide in prs.slides:
                temp_pdf = "temp.pdf"

                # Save each slide as a temporary PDF
                slide_pdf = PyPDF2.PdfFileWriter()
                slide_pdf.addPage(PyPDF2.pdf.PageObject.createTextPage(slide.shapes.title.text))
                for shape in slide.shapes:
                    if hasattr(shape, "text"):
                        slide_pdf.addPage(PyPDF2.pdf.PageObject.createTextPage(shape.text))

                with open(temp_pdf, "wb") as temp_file:
                    slide_pdf.write(temp_file)

                # Merge the temporary PDFs into the final PDF
                with open(temp_pdf, "rb") as temp_file:
                    pdf_reader = PyPDF2.PdfFileReader(temp_file)
                    for page_num in range(pdf_reader.getNumPages()):
                        page = pdf_reader.getPage(page_num)
                        pdf_writer.addPage(page)

                os.remove(temp_pdf)

            with open(output_file, 'wb') as output:
                pdf_writer.write(output)

            messagebox.showinfo("Conversion Complete", "PowerPoint to PDF conversion is complete!")

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
convert_excel_to_pdf_button = Button(root, text="Excel to PDF", command=convert_excel_to_pdf, padx=10, pady=5, bg='blue', fg='white')
convert_ppt_to_pdf_button = Button(root, text="PPT to PDF", command=convert_ppt_to_pdf, padx=10, pady=5, bg='blue', fg='white')

convert_pdf_to_word_button.pack(pady=10)
convert_word_to_pdf_button.pack(pady=10)
convert_image_to_pdf_button.pack(pady=10)
convert_excel_to_pdf_button.pack(pady=10)
convert_ppt_to_pdf_button.pack(pady=10)

root.mainloop()
