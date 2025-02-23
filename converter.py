import os
import time
from glob import glob
from modules.docx2md_converter import Docx2MdConverter
from modules.doc2docx_converter import Doc2DocxConverter
from modules.pdf2docx_converter import Pdf2DocxConverter
from modules.pptx2pdf_converter import Pptx2PdfConverter



def convert_docx2md(input_directory: str, output_directory: str, split_output: bool = False) -> None:
    input_directory_abs = os.path.abspath(input_directory)
    done_directory_abs = f"{input_directory_abs}\\done"

    paths = glob(f"{input_directory_abs}\\*.docx", recursive=False)

    for path in paths:
        converter = Docx2MdConverter(path, split_output)
        converter.save(output_directory)
        converter.move_to_done(done_directory_abs)

    print("All docx files converted to markdown")
            

def convert_doc2docx(input_directory: str, output_directory: str) -> None:    
    while True:
        try:
            input_directory_abs = os.path.abspath(input_directory)
            output_directory_abs = os.path.abspath(output_directory)
            done_directory_abs = f"{input_directory_abs}\\done"

            paths = glob(f"{input_directory_abs}\\*.doc", recursive=False)

            for path in paths:
                converter = Doc2DocxConverter(path)
                converter.convert_and_save(output_directory_abs)
                converter.move_to_done(done_directory_abs)
        except Exception:
            print("Trying again to convert doc files to docx...")
            time.sleep(1)
        else:
            print("All doc files converted to docx")
            break


def convert_pdf2docx(input_directory: str, output_directory: str) -> None:
    input_directory_abs = os.path.abspath(input_directory)
    output_directory_abs = os.path.abspath(output_directory)
    done_directory_abs = f"{input_directory_abs}\\done"

    paths = glob(f"{input_directory_abs}\\*.pdf", recursive=False)

    for path in paths:
        converter = Pdf2DocxConverter(path)
        converter.convert_and_save(output_directory_abs)
        converter.move_to_done(done_directory_abs)
    
    print("All pdf files converted to docx")


def convert_pptx2pdf(input_directory: str, output_directory: str) -> None:
    input_directory_abs = os.path.abspath(input_directory)
    output_directory_abs = os.path.abspath(output_directory)
    done_directory_abs = f"{input_directory_abs}\\done"

    paths = glob(f"{input_directory_abs}\\*.pptx", recursive=False)

    for path in paths:
        converter = Pptx2PdfConverter(path)
        converter.convert_and_save(output_directory_abs)
        converter.move_to_done(done_directory_abs)
    
    print("All pptx files converted to pdf")


def create_folder_structure():
    os.makedirs("input", exist_ok=True)
    os.makedirs("input\\doc", exist_ok=True)
    os.makedirs("input\\pdf", exist_ok=True)
    os.makedirs("input\\pptx", exist_ok=True)
    os.makedirs("input\\docx", exist_ok=True)

    print("Folder structure created")



def main():
    create_folder_structure()
    convert_doc2docx(input_directory = "input\\doc", output_directory = "input\\docx")
    convert_pptx2pdf(input_directory = "input\\pptx", output_directory = "input\\pdf")
    convert_pdf2docx(input_directory = "input\\pdf", output_directory = "input\\docx")
    convert_docx2md(input_directory = "input\\docx", output_directory = "output", split_output = False)



if __name__ == "__main__":
    main()
