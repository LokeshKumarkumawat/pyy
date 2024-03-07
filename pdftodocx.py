from pdf2docx import Converter

def pdf_to_word(pdf_path, word_path):
    cv = Converter(pdf_path)
    cv.convert(word_path, start=0, end=None)
    cv.close()

# Example usage
pdf_file = 'input.pdf'
word_file = 'output.docx'

pdf_to_word(pdf_file, word_file)
