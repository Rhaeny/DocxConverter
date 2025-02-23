import os
import re
import aspose.slides as slides



class Pptx2PdfConverter:
    def __init__(self, filepath_abs: str) -> None:
        self.filepath_abs = filepath_abs
        self.filename = os.path.split(filepath_abs)[1]

    def convert_and_save(self, output_directory: str) -> None:
        pdf_file_abs = f"{output_directory}\\{self.filename}"
        pdf_file_abs = re.sub(r'\.\w+$', '.pdf', pdf_file_abs)

        with slides.Presentation(self.filepath_abs) as presentation:
            presentation.save(pdf_file_abs, slides.export.SaveFormat.PDF)
    
    def move_to_done(self, done_directory: str) -> None:
        os.makedirs(done_directory, exist_ok=True)

        done_file_abs = f"{done_directory}\\{self.filename}"
        os.rename(self.filepath_abs, done_file_abs)



def main():
    ...



if __name__ == "__main__":
    main()