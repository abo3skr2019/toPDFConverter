import os
import comtypes.client
from docx2pdf import convert as docx_convert
import shutil

def convert_ppt_to_pdf(ppt_path, pdf_path):
    powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
    powerpoint.Visible = 1
    deck = powerpoint.Presentations.Open(ppt_path)
    deck.SaveAs(pdf_path, 32)  # 32 is the formatType for pdf
    deck.Close()
    powerpoint.Quit()

def convert_docx_to_pdf(docx_path, pdf_path):
    docx_convert(docx_path, pdf_path)

def move_file(original_path, new_directory):
    if not os.path.exists(new_directory):
        os.makedirs(new_directory)
    shutil.move(original_path, os.path.join(new_directory, os.path.basename(original_path)))

def main():
    current_directory = os.getcwd()
    ppt_dir = os.path.join(current_directory, 'PPT')
    doc_dir = os.path.join(current_directory, 'DOC')
    
    for root, dirs, files in os.walk(current_directory):
        for file in files:
            file_path = os.path.join(root, file)
            if file.endswith('.pptx') or file.endswith('.ppt'):
                pdf_path = file_path.replace('.pptx', '.pdf').replace('.ppt', '.pdf')
                try:
                    convert_ppt_to_pdf(file_path, pdf_path)
                    move_file(file_path, ppt_dir)
                except Exception as e:
                    print(f"Failed to convert {file}: {e}")

            elif file.endswith('.docx') or file.endswith('.doc'):
                pdf_path = file_path.replace('.docx', '.pdf').replace('.doc', '.pdf')
                try:
                    convert_docx_to_pdf(file_path, pdf_path)
                    move_file(file_path, doc_dir)
                except Exception as e:
                    print(f"Failed to convert {file}: {e}")

if __name__ == "__main__":
    main()
