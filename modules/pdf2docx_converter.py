import os
import re
from pdf2docx import Converter



class Pdf2DocxConverter:
    def __init__(self, filepath_abs: str) -> None:
        self.filepath_abs = filepath_abs
        self.filename = os.path.split(filepath_abs)[1]

    def convert_and_save(self, output_directory: str) -> None:
        docx_file_abs = f"{output_directory}\\{self.filename}"
        docx_file_abs = re.sub(r'\.\w+$', '.docx', docx_file_abs)

        cv = Converter(self.filepath_abs)
        cv.convert(docx_file_abs)
        cv.close()
    
    def move_to_done(self, done_directory: str) -> None:
        os.makedirs(done_directory, exist_ok=True)

        done_file_abs = f"{done_directory}\\{self.filename}"
        os.rename(self.filepath_abs, done_file_abs)



def main():
    ...



if __name__ == "__main__":
    main()