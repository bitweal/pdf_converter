from PyPDF2 import PdfMerger, PdfReader, PdfWriter
import fitz
import os
from pdf2docx import Converter


def merge_pdfs(pdf_list, output):
    merger = PdfMerger()
    for pdf in pdf_list:
        merger.append(pdf)
    merger.write(output)
    merger.close()


def split_pdf(input_pdf, start_page, end_page, output_pdf):
    reader = PdfReader(input_pdf)
    writer = PdfWriter()
    for page in range(start_page, end_page):
        writer.add_page(reader.pages[page])
    with open(output_pdf, 'wb') as output:
        writer.write(output)


def compress_pdf(input_pdf, output_pdf, dpi=100):
    doc = fitz.open(input_pdf)
    new_doc = fitz.open()

    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        rect = page.rect
        pix = page.get_pixmap(dpi=dpi)
        image_path = f"media/temp_page_{page_num}.jpg"
        pix.save(image_path)
        new_page = new_doc.new_page(width=rect.width, height=rect.height)
        new_page.insert_image(rect, filename=image_path)
        os.remove(image_path)

    new_doc.save(output_pdf, garbage=4, deflate=True)
    new_doc.close()
    doc.close()


def pdf_to_word(input_pdf, output_docx):
    cv = Converter(input_pdf)
    cv.convert(output_docx, start=0, end=None)
    cv.close()


if __name__ == '__main__':
    pdf_file = 'media/iasa-open_21_.pdf'
    #merge_pdfs([pdf_file for _ in range(100)], 'media/merge_pdfs.pdf')
    #split_pdf('media/merge_pdfs.pdf', 15, 42, 'media/split.pdf')
    #compress_pdf('media/merge_pdfs.pdf', 'media/compress_pdf.pdf', 100)
    pdf_to_word(pdf_file, 'media/pdf_to_word.docx')
