import os
import time
from glob import glob
from modules.docx2md_converter import Docx2MdConverter
from modules.doc2docx_converter import Doc2DocxConverter
from modules.pdf2docx_converter import Pdf2DocxConverter



def convert_docx2md(input_directory: str, output_directory: str) -> None:
    input_directory_abs = os.path.abspath(input_directory)
    done_directory_abs = f"{input_directory_abs}\\done"

    paths = glob(f"{input_directory_abs}\\*.docx", recursive=False)

    for path in paths:
        converter = Docx2MdConverter(path)
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



def main():
    convert_doc2docx(input_directory = "input\\doc", output_directory = "input\\docx")
    convert_pdf2docx(input_directory = "input\\pdf", output_directory = "input\\docx")
    convert_docx2md(input_directory = "input\\docx", output_directory = "output")



if __name__ == "__main__":
    main()
