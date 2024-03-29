from tkinter import Tk, filedialog
from pdf2docx import Converter


def pdf_to_docx(pdf_path, docx_path):
    cv = Converter(pdf_path)
    cv.convert(docx_path, start=0, end=None)
    cv.close()


def select_pdf_file():
    root = Tk()
    root.withdraw()  # Hide the main window

    file_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
    return file_path


def select_output_file():
    root = Tk()
    root.withdraw()  # Hide the main window

    file_path = filedialog.asksaveasfilename(
        defaultextension=".docx", filetypes=[("DOCX files", "*.docx")])
    return file_path


def main():
    pdf_file = select_pdf_file()

    if not pdf_file:
        print("No PDF file selected.")
        return

    docx_file = select_output_file()

    if not docx_file:
        print("No output DOCX file selected.")
        return

    pdf_to_docx(pdf_file, docx_file)
    print("Conversion complete!")


if __name__ == "__main__":
    main()
