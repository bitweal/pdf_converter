from PyPDF2 import PdfMerger, PdfReader, PdfWriter


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


if __name__ == '__main__':
    pdf_file = 'media/iasa-open_21_.pdf'
    #merge_pdfs([pdf_file, pdf_file, pdf_file], 'media/merge_pdfs.pdf')
    #split_pdf('media/merge_pdfs.pdf', 15, 42, 'media/split.pdf')
