import pdf2docx
import os


def convert_pdf_to_word(pdf_path, word_path):
    """Convert PDF file to Word DOCX file"""
    
    # Convert PDF to Word
    cv = pdf2docx.Converter(pdf_path) 
    cv.convert(word_path)  
    
    # Save Word file
    cv.close() 
    
file_path = os.getcwd() + "/[PV3] Relatório de CLIMA - SEAP I - v1.pdf"
output_path = os.getcwd() + "/[PV3] Relatório de CLIMA - SEAP I - v1.docx"

convert_pdf_to_word(file_path, output_path)
