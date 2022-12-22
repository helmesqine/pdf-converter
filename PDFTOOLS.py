import glob
import os
import img2pdf
from tkinter import messagebox
from docx2pdf import convert
from PyPDF2 import PdfMerger
from pdf2image import convert_from_path
from pdf2docx import Converter


class TOOLS:
    def __init__(self):
        self.file = None
        self.A4 = (img2pdf.mm_to_pt(210), img2pdf.mm_to_pt(297))
        self.A3 = (img2pdf.mm_to_pt(297), img2pdf.mm_to_pt(420))
        self.A5 = (img2pdf.mm_to_pt(148), img2pdf.mm_to_pt(210))

    def image_to_pdf(self, files, out, opt, size):
        if out == '':
            messagebox.showerror('PDFCONVERTER', 'Please choose output directory!')
            return
        if not files:
            messagebox.showerror('PDFCONVERTER', 'Please choose files to convert!')
            return
        if str(opt) == 'Multi files' or len(files) == 1:
            for self.file in files:
                filename = os.path.basename(self.file)
                with open(f'{out}/{filename}.pdf', "wb") as f:
                    if size == 'Same as image':
                        f.write(img2pdf.convert(files))
                    elif size == 'A4':
                        layout_fun = img2pdf.get_layout_fun(self.A4)
                        f.write(img2pdf.convert(files, layout_fun=layout_fun))
                    elif size == 'A5':
                        layout_fun = img2pdf.get_layout_fun(self.A5)
                        f.write(img2pdf.convert(files, layout_fun=layout_fun))
                    elif size == 'A3':
                        layout_fun = img2pdf.get_layout_fun(self.A3)
                        f.write(img2pdf.convert(files, layout_fun=layout_fun))
        elif str(opt) == 'Single file':
            with open(out + '/' + "image_to_pdf_all_join.pdf", "wb") as f:
                if size == 'Same as image':
                    f.write(img2pdf.convert(files))
                elif size == 'A4':
                    layout_fun = img2pdf.get_layout_fun(self.A4)
                    f.write(img2pdf.convert(files, layout_fun=layout_fun))
                elif size == 'A5':
                    layout_fun = img2pdf.get_layout_fun(self.A5)
                    f.write(img2pdf.convert(files, layout_fun=layout_fun))
                elif size == 'A3':
                    layout_fun = img2pdf.get_layout_fun(self.A3)
                    f.write(img2pdf.convert(files, layout_fun=layout_fun))

    def word_to_pdf(self, files, out, opt):
        self.file = None
        if str(opt) == 'Multi files' or len(files) == 1:
            for doc in files:
                filename = os.path.basename(doc)
                ext = os.path.splitext(doc)
                convert(doc, f"{out}/{str(filename).replace(ext[1], '.pdf')}")
        if str(opt) == 'Single file':
            for doc in files:
                filename = os.path.basename(doc)
                convert(doc, f"temp/{filename}.pdf")

            pdfs = glob.glob('temp/*.pdf')
            merger = PdfMerger()
            for pdf in pdfs:
                merger.append(pdf)
            merger.write(f"{out}/word_to_pdf_all_join.pdf")
            merger.close()
            for pdf in pdfs:
                os.remove(pdf)

    def pdf_to_img(self, files, out):
        self.file = None
        for pdf in files:
            file_name, ext = os.path.splitext(pdf)
            filename = os.path.basename(file_name)
            images = convert_from_path(pdf)
            for i in range(len(images)):
                images[i].save(f'{out}/{filename}_page_{i + 1}.png', 'PNG')

    def pdf_to_doc(self, files, out):
        self.file = None
        for pdf in files:
            file_name, ext = os.path.splitext(pdf)
            filename = os.path.basename(file_name)
            cv_obj = Converter(pdf)
            cv_obj.convert(out + '/' + filename + '.docx')
            cv_obj.close()
