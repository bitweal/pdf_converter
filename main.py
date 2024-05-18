from PyPDF2 import PdfMerger


def merge_pdfs(pdf_list, output):
    merger = PdfMerger()
    for pdf in pdf_list:
        merger.append(pdf)
    merger.write(output)
    merger.close()


if __name__ == '__main__':
    pdf_file = 'media/iasa-open_21_.pdf'
    merge_pdfs([pdf_file, pdf_file, pdf_file], 'merge_pdfs.pdf')
