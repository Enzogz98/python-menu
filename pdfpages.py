import os
from PyPDF2 import PdfReader, PdfWriter

def split_pdf_pages(pdf_directory, output_directory):
    if not os.path.exists(output_directory):
        os.makedirs(output_directory)

    for filename in os.listdir(pdf_directory):
        if filename.lower().endswith('.pdf'):
            pdf_path = os.path.join(pdf_directory, filename)
            base_name = os.path.splitext(filename)[0]
            with open(pdf_path, 'rb') as file:
                pdf_reader = PdfReader(file)
                num_pages = len(pdf_reader.pages)
                for page_num in range(num_pages):
                    page = pdf_reader.pages[page_num]
                    output_page_path = os.path.join(output_directory, f"{base_name}_page_{page_num + 1}.pdf")
                    pdf_writer = PdfWriter()
                    pdf_writer.add_page(page)
                    with open(output_page_path, 'wb') as output_file:
                        pdf_writer.write(output_file)
                    print(f"Página {page_num + 1} del archivo {filename} guardada como {output_page_path}")

pdf_directory = 'C:\\Python312\\pdfs'
output_directory = 'C:\\Python312\\pdfsPages'

split_pdf_pages(pdf_directory, output_directory)

