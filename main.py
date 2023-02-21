import tkinter as tk
from tkinter import ttk
import os
import comtypes.client


# This function converts a Word file into a PDF file
def convert_word_to_pdf():
    word_pdf_format = 17

    docx_file_location = os.path.abspath(entry_docx_location.get())
    docx_file_name = entry_docx_filename.get()

    pdf_file_location = os.path.abspath(entry_pdf_location.get())
    pdf_file_name = entry_pdf_filename.get()

    # Completing the file paths
    docx_file_path = docx_file_location + '\\' + docx_file_name + '.docx'
    pdf_file_path = pdf_file_location + '\\' + pdf_file_name + '.pdf'

    word_document = comtypes.client.CreateObject('Word.Application')
    new_document = word_document.Documents.Open(docx_file_path)
    new_document.SaveAs(pdf_file_path, FileFormat=word_pdf_format)
    new_document.Close()
    word_document.Quit()


# GUI root configuration
root = tk.Tk()
root.title("Convert Word to PDF")
root.geometry("960x175")
root.resizable(width=False, height=False)


# Program headline configuration
label_headline = ttk.Label(root, text="Word to PDF")
label_headline.pack()


# This section configures the labels and entry fields for the Word (.docx) file
label_docx_location = ttk.Label(root, text="Location of the Word (.docx) file:")
label_docx_location.pack()
label_docx_location.place(x=5, y=20)

entry_docx_location = ttk.Entry(root, width=90)
entry_docx_location.pack()
entry_docx_location.place(x=205, y=20)

label_docx_filename = ttk.Label(root, text="Name of the Word (.docx) file:")
label_docx_filename.pack()
label_docx_filename.place(x=5, y=50)

entry_docx_filename = ttk.Entry(root, width=90)
entry_docx_filename.pack()
entry_docx_filename.place(x=205, y=50)


# This section configures the labels and entry fields for the properties of the PDF file
label_pdf_location = ttk.Label(root, text="Location of the new PDF file:")
label_pdf_location.pack()
label_pdf_location.place(x=5, y=80)

entry_pdf_location = ttk.Entry(root, width=90)
entry_pdf_location.pack()
entry_pdf_location.place(x=205, y=80)

label_pdf_filename = ttk.Label(root, text="Name of the new PDF file:")
label_pdf_filename.pack()
label_pdf_filename.place(x=5, y=110)

entry_pdf_filename = ttk.Entry(root, width=90)
entry_pdf_filename.pack()
entry_pdf_filename.place(x=205, y=110)


# Configuration of the button that starts file conversion
button_convert = ttk.Button(root, text="Convert", command=convert_word_to_pdf)
button_convert.pack()
button_convert.place(x=440, y=140)


# Execution of the mainloop from tkinter
root.mainloop()
