import os
import time
from docx2pdf import convert
import subprocess
import aspose.slides as slides

# Whatever folder you want to manage downloads in
watched_folder = r'C:\Users\Alex\Downloads'
seen_files = set(os.listdir(watched_folder))


# Converts docx or pptx files to pdf's and opens in chrome browser
def convert_to_pdf(file):
    file_extension = os.path.splitext(file)[-1].lower()
    file_path = os.path.join(watched_folder, file)
    output_folder = r'C:\Users\Alex\Desktop\New folder (6)'

    if file_extension == '.docx':
        try:
            # Convert .docx to PDF using docx2pdf
            convert(file_path, r'C:\Users\Alex\Desktop\New folder (6)')

            pdf_filename = os.path.basename(file_path).replace('.docx', '.pdf')
            pdf_path = os.path.join(output_folder, pdf_filename)

            chrome_path = r'C:\Program Files\Google\Chrome\Application\chrome.exe'
            # Open the PDF file with Google Chrome
            subprocess.Popen([chrome_path, pdf_path], shell=True)
        except Exception as e:
            print(f"Error converting .docx to PDF and opening with Chrome: {e}")

    if file_extension == '.pptx':
        try:
            # Converts pptx file to pdf
            with slides.Presentation(file_path) as presentation:
                presentation.save("presentation.pdf", slides.export.SaveFormat.PDF)

            chrome_path = r'C:\Program Files\Google\Chrome\Application\chrome.exe'
            # Open the PDF file with Google Chrome
            subprocess.Popen([chrome_path, r'C:\Users\Alex\PycharmProjects\pythonProject2\presentation.pdf'],
                             shell=True)
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
    # Timeout between checking files again
    time.sleep(3)
