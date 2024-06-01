import os
import comtypes.client
from docx2pdf import convert as docx_convert
import shutil

def convert_ppt_to_pdf(ppt_path, pdf_path):
    try:
        print(f"Converting {ppt_path} to {pdf_path}")
        powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
        powerpoint.Visible = 1
        deck = powerpoint.Presentations.Open(ppt_path)
        deck.SaveAs(pdf_path, 32)  # 32 is the formatType for pdf
        deck.Close()
    except Exception as e:
        print(f"Failed to convert PowerPoint {ppt_path} to PDF: {e}")
    finally:
        powerpoint.Quit()

def convert_docx_to_pdf(docx_path, pdf_path):
    try:
        print(f"Converting {docx_path} to {pdf_path}")
        docx_convert(docx_path, pdf_path)
    except Exception as e:
        print(f"Failed to convert DOCX {docx_path} to PDF: {e}")

def move_file(original_path, new_directory):
    try:
        if not os.path.exists(new_directory):
            os.makedirs(new_directory)
        shutil.move(original_path, os.path.join(new_directory, os.path.basename(original_path)))
    except Exception as e:
        print(f"Failed to move file {original_path} to {new_directory}: {e}")

def main():
    current_directory = os.getcwd()
    ppt_dir = os.path.join(current_directory, 'PPT')
    doc_dir = os.path.join(current_directory, 'DOC')
    
    for root, dirs, files in os.walk(current_directory):
        for file in files:
            file_path = os.path.join(root, file)
            file_name, file_extension = os.path.splitext(file_path)
            pdf_path = file_name + '.pdf'

            if file_extension.lower() in ['.ppt', '.pptx']:
                try:
                    convert_ppt_to_pdf(file_path, pdf_path)
                    move_file(file_path, ppt_dir)
                except Exception as e:
                    print(f"Failed to process PowerPoint file {file_path}: {e}")

            elif file_extension.lower() in ['.doc', '.docx']:
                try:
                    convert_docx_to_pdf(file_path, pdf_path)
                    move_file(file_path, doc_dir)
                except Exception as e:
                    print(f"Failed to process DOCX file {file_path}: {e}")

if __name__ == "__main__":
    main()
