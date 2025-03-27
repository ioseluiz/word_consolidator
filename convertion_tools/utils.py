import os
import win32com.client as win32
import win32com.client.makepy
import fitz # PyMuPDF

from pdf2docx import parse

def extract_filename(filepath):
    """
    Extracts the filename from a given filepath.

    Args:
        filepath (str): The full path to the file.

    Returns:
        str: The filename, or None if the filepath is invalid or empty.
    """
    if not filepath:
        return None

    return os.path.basename(filepath)


def convert_docx_to_doc(docx_filepath, destination_folder):
    print(docx_filepath)
    filename = extract_filename(docx_filepath)
    name, ext = os.path.splitext(filename)
    word = win32.gencache.EnsureDispatch("Word.Application")
    word.Visible = False
    # Open the word document in docx format
    docx_filepath = os.path.abspath(docx_filepath)
    destination_path = os.path.abspath(os.path.join(destination_folder, f"{name}.doc"))
    word_docx = word.Documents.Open(docx_filepath)
    word_docx.SaveAs(destination_path, FileFormat=0)
    word_docx.Close()
    
def process_files(file_list, target_folder):
    for item in file_list:
        if item[-5:] == ".docx":
            print('######################')
            print(item)
            convert_docx_to_doc(item, target_folder)
            
def convert_files_to_pdf(file_list, target_folder):
    for item in file_list:
        if item[-4:] == ".doc":
            print("PDF Conversion\n")
            print(item)
            doc_to_pdf(item, target_folder)
            
def doc_to_pdf(doc_filepath, destination_folder):
    filename = extract_filename(doc_filepath)
    name, ext = os.path.splitext(filename)
    # Ensure absolute paths
    doc_filepath = os.path.abspath(doc_filepath)
    print(doc_filepath)
    destination_path = os.path.abspath(os.path.join(destination_folder, f"{name}.pdf"))
    word = win32.gencache.EnsureDispatch("Word.Application")
    word.Visible = False
    # Open the word document in doc format
    word_doc = word.Documents.Open(doc_filepath)
    word_doc.SaveAs(destination_path, FileFormat=17)
    word_doc.Close()
    
def merge_pdf_files(target_folder):
    try:
        pdf_files = [f for f in os.listdir(target_folder) if f.lower().endswith(".pdf")]
        
        if not pdf_files:
            print("No PDF files in the folder.")
        
        with fitz.open() as merged_pdf:
            for pdf_file in pdf_files:
                pdf_path = os.path.join(target_folder, pdf_file)
                with fitz.open(pdf_path) as pdf_document:
                    merged_pdf.insert_pdf(pdf_document)
                    
            merged_pdf.save(os.path.join(target_folder, "merged_file.pdf"))
            
    except Exception as e:
        print(f"Error during PDF merge: {e}")
        
def convert_pdf_to_word(target_folder):
    pdf_file = os.path.join(target_folder, "merged_file.pdf")
    pdf_file = os.path.abspath(pdf_file)
    word_file = os.path.join(target_folder, "merged_file.docx")
    word_file = os.path.abspath(word_file)
    win32com.client.makepy.GenerateFromTypeLibSpec('Acrobat')
    adobe = win32.DispatchEx('AcroExch.App')
    avDoc = win32.DispatchEx('AcroExch.AVDoc')
    avDoc.Open(pdf_file, pdf_file)
    pdDoc = avDoc.GetPDDoc()
    jObject = pdDoc.GetJSObject()
    jObject.SaveAs(word_file, "com.adobe.acrobat.docx")
    avDoc.Close(-1)
    
        
            
