import tkinter as tk
from tkinter import ttk
import os
import comtypes.client


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


def main():
    # Execution of the mainloop from tkinter, which makes the GUI appear on screen
    root.mainloop()
    # When the button 'button_convert' is pressed on the GUI, then the function 'convert_word_to_pdf' will be executed


# GUI root configuration
root = tk.Tk()
root.title('Convert Word to PDF')
root.geometry('960x175')
root.resizable(width=False, height=False)


# Creating the different widgets for the GUI
label_headline = ttk.Label(root, text='Word to PDF')

button_convert = ttk.Button(root, text='Convert', command=convert_word_to_pdf)

label_docx_location = ttk.Label(root, text='Location of the Word (.docx) file:')
entry_docx_location = ttk.Entry(root, width=90)

label_docx_filename = ttk.Label(root, text='Name of the Word (.docx) file:')
entry_docx_filename = ttk.Entry(root, width=90)

label_pdf_location = ttk.Label(root, text='Location of the new PDF file:')
entry_pdf_location = ttk.Entry(root, width=90)

label_pdf_filename = ttk.Label(root, text='Name of the new PDF file:')
entry_pdf_filename = ttk.Entry(root, width=90)


# Placing the different widgets onto the GUI
label_headline.pack()

button_convert.pack()
button_convert.place(x=440, y=140)

label_docx_location.pack()
label_docx_location.place(x=5, y=20)

entry_docx_location.pack()
entry_docx_location.place(x=205, y=20)

label_docx_filename.pack()
label_docx_filename.place(x=5, y=50)

entry_docx_filename.pack()
entry_docx_filename.place(x=205, y=50)

label_pdf_location.pack()
label_pdf_location.place(x=5, y=80)

entry_pdf_location.pack()
entry_pdf_location.place(x=205, y=80)

label_pdf_filename.pack()
label_pdf_filename.place(x=5, y=110)

entry_pdf_filename.pack()
entry_pdf_filename.place(x=205, y=110)


if __name__ == "__main__":
    main()
