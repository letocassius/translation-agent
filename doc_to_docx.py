import os, sys
import win32com.client

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

def convert_doc_to_docx(file_path, output_path):
    # Create COM object to represent Word Application
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False

    # Open the .doc file
    try:
        doc = word.Documents.Open(file_path)
        
        # Save as .docx format
        doc.SaveAs(output_path, 16)  # 16 denotes the wdFormatXMLDocument format

        # Close the document
        doc.Close()
        print(f"Converted: {file_path} to {output_path}")
    except Exception as e:
        print(f"Failed to convert: {file_path}. Error: {e}")
    finally:
        # Quit the Word application
        word.Quit()

def convert_folder_of_docs(folder_path):
    for filename in os.listdir(folder_path):
        if filename.endswith(".doc"):
            file_path = os.path.join(folder_path, filename)
            output_path = os.path.join(folder_path, f"{os.path.splitext(filename)[0]}.docx")
            convert_doc_to_docx(file_path, output_path)

if __name__ == "__main__":
    file_path = input("Enter the path to the Word document: ")
    convert_folder_of_docs(file_path)