import os
import comtypes.client
from docx2pdf import convert as docx_convert
import shutil
from concurrent.futures import ThreadPoolExecutor, as_completed
import pythoncom

def convert_ppt_to_pdf(ppt_path, pdf_path):
    try:
        pythoncom.CoInitialize()  # Initialize COM library in this thread
        print(f"Converting {ppt_path} to {pdf_path}")
        powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
        powerpoint.Visible = 1  # Make PowerPoint visible
        powerpoint.DisplayAlerts = False  # Disable alerts
        deck = powerpoint.Presentations.Open(ppt_path)
        deck.SaveAs(pdf_path, 32)  # 32 is the formatType for pdf
        deck.Close()
        return True
    except Exception as e:
        print(f"Failed to convert PowerPoint {ppt_path} to PDF: {e}")
        return False
    finally:
        if 'powerpoint' in locals():
            powerpoint.Quit()
        pythoncom.CoUninitialize()  # Uninitialize COM library in this thread

def convert_docx_to_pdf(docx_path, pdf_path):
    try:
        print(f"Converting {docx_path} to {pdf_path}")
        docx_convert(docx_path, pdf_path)
        return True
    except Exception as e:
        print(f"Failed to convert DOCX {docx_path} to PDF: {e}")
        return False

def convert_doc_to_pdf(doc_path, pdf_path):
    try:
        pythoncom.CoInitialize()  # Initialize COM library in this thread
        print(f"Converting {doc_path} to {pdf_path}")
        word = comtypes.client.CreateObject("Word.Application")
        word.Visible = 0
        word.DisplayAlerts = 0
        doc = word.Documents.Open(doc_path)
        doc.SaveAs(pdf_path, FileFormat=17)  # 17 is the formatType for pdf
        doc.Close()
        return True
    except Exception as e:
        print(f"Failed to convert DOC {doc_path} to PDF: {e}")
        return False
    finally:
        if 'word' in locals():
            word.Quit()
        pythoncom.CoUninitialize()  # Uninitialize COM library in this thread

def move_file(original_path, new_directory):
    try:
        if not os.path.exists(new_directory):
            os.makedirs(new_directory)
        shutil.move(original_path, os.path.join(new_directory, os.path.basename(original_path)))
        return True
    except Exception as e:
        print(f"Failed to move file {original_path} to {new_directory}: {e}")
        return False

def process_files(file_list, ppt_dir, doc_dir):
    with ThreadPoolExecutor(max_workers=os.cpu_count()) as executor:
        futures = []
        for file_path in file_list:
            file_name, file_extension = os.path.splitext(file_path)
            pdf_path = file_name + '.pdf'
            
            if file_extension.lower() in ['.ppt', '.pptx']:
                futures.append(executor.submit(convert_and_move, file_path, pdf_path, ppt_dir, 'ppt'))
            elif file_extension.lower() in ['.doc', '.docx']:
                futures.append(executor.submit(convert_and_move, file_path, pdf_path, doc_dir, 'doc'))
        
        for future in as_completed(futures):
            future.result()  # Handle exceptions if needed

def convert_and_move(file_path, pdf_path, new_directory, file_type):
    if file_type == 'ppt':
        if convert_ppt_to_pdf(file_path, pdf_path):
            move_file(file_path, new_directory)
    elif file_type == 'doc':
        if file_path.lower().endswith('.doc'):
            if convert_doc_to_pdf(file_path, pdf_path):
                move_file(file_path, new_directory)
        elif file_path.lower().endswith('.docx'):
            if convert_docx_to_pdf(file_path, pdf_path):
                move_file(file_path, new_directory)

def main():
    current_directory = os.getcwd()
    ppt_dir = os.path.join(current_directory, 'PPT')
    doc_dir = os.path.join(current_directory, 'DOC')
    file_list = []

    for root, dirs, files in os.walk(current_directory):
        for file in files:
            file_path = os.path.join(root, file)
            file_list.append(file_path)

    process_files(file_list, ppt_dir, doc_dir)

if __name__ == "__main__":
    main()
