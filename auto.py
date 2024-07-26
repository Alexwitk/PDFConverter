import os
import time
from docx2pdf import convert
import subprocess
import aspose.slides as slides
from datetime import datetime

# Whatever folder you want to manage downloads in
watched_folder = r''
chrome_path = r''
# Folder new files are saved to
output_folder = r''
seen_files = set(os.listdir(watched_folder))

# Converts docx or pptx files to pdf's and opens in chrome browser
def convert_to_pdf(file):
    file_extension = os.path.splitext(file)[-1].lower()
    file_path = os.path.join(watched_folder, file)
    
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    pdf_filename = os.path.basename(file_path).replace(file_extension, f'_{timestamp}.pdf')
    pdf_path = os.path.join(output_folder, pdf_filename)

    if file_extension == '.docx':
        try:
            # Convert .docx to PDF using docx2pdf
            convert(file_path, output_folder)

            pdf_filename = os.path.basename(file_path).replace('.docx', '.pdf')
            pdf_path = os.path.join(output_folder, pdf_filename)

            # Open the PDF file with Google Chrome
            subprocess.Popen([chrome_path, pdf_path], shell=True)
        except Exception as e:
            print(f"Error converting .docx to PDF and opening with Chrome: {e}")

    if file_extension == '.pptx':
        try:
            # Converts pptx file to pdf
            with slides.Presentation(file_path) as presentation:
                presentation.save(pdf_path, slides.export.SaveFormat.PDF)

            # Open the PDF file with Google Chrome
            subprocess.Popen([chrome_path, pdf_path], shell=True)
        except Exception as e:
            print(f"Error converting .pptx to PDF and opening with Chrome: {e}")

while True:
    # List all files in the watched folder
    files = set(os.listdir(watched_folder))
    # Find new files in the folder
    new_files = files - seen_files

    for file in new_files:
        file_extension = os.path.splitext(file)[-1].lower()

        if file_extension in ['.docx', '.pptx']:
            convert_to_pdf(file)

    seen_files = files
    # Timeout between checking files againW
    time.sleep(3)
