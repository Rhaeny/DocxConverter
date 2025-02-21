import os
import win32com.client as win32
from win32com.client import constants
import re



class Doc2DocxConverter():
    def __init__(self, filepath_abs: str) -> None:
        self.filepath_abs = filepath_abs
        self.filename = os.path.split(filepath_abs)[1]

    def convert_and_save(self, output_directory: str) -> None:
        word = win32.gencache.EnsureDispatch('Word.Application')
        doc = word.Documents.Open(self.filepath_abs)
        doc.Activate()

        new_file_abs = f"{output_directory}\\{self.filename}"
        new_file_abs = re.sub(r'\.\w+$', '.docx', new_file_abs)

        word.ActiveDocument.SaveAs(
            new_file_abs, FileFormat=constants.wdFormatXMLDocument
        )
        doc.Close(True)
        word.Quit()
    
    def move_to_done(self, done_directory: str) -> None:
        done_file_abs = f"{done_directory}\\{self.filename}"
        os.rename(self.filepath_abs, done_file_abs)



def main():
    ...



if __name__ == "__main__":
    main()