import os
import mammoth
import uuid
from markdownify import markdownify
from pathlib import Path



class Docx2MdConverter:
    def __init__(self, filepath_abs: str, split_output: bool = False) -> None:
        self.filepath_abs = filepath_abs
        self.filename = Path(filepath_abs).stem
        self.split_output = split_output

        with open(filepath_abs, "rb") as docx_file:
            result = mammoth.convert_to_html(docx_file, convert_image=mammoth.images.img_element(self.convert_image))
            html = result.value
            self.markdown = markdownify(html)


    def save(self, output_directory: str) -> None:
        os.makedirs(output_directory, exist_ok=True)
        filename = self.filename.replace(" ", "_")

        if self.split_output:
            output_directory_split = f"{output_directory}\\{filename}"
            os.makedirs(output_directory_split, exist_ok=True)
            with open(output_directory_split + "/" + filename + ".md", 'w', encoding="utf-8") as file:
                file.write(self.markdown)
        else:
            with open(output_directory + "/" + filename + ".md", 'w', encoding="utf-8") as file:
                file.write(self.markdown)        


    def move_to_done(self, done_directory: str) -> None:
        os.makedirs(done_directory, exist_ok=True)

        done_file_abs = f"{done_directory}\\{self.filename}.docx"
        os.rename(self.filepath_abs, done_file_abs)


    def convert_image(self, image):
        filename = self.filename.replace(" ", "_")
        image_folder = f".attachments\\{filename}"
        if self.split_output:
            output_dir = f"output\\{filename}\\.attachments\\{filename}"
        else:
            output_dir = f"output\\.attachments\\{filename}"

        os.makedirs(output_dir, exist_ok=True)
        
        unique_name = uuid.uuid4()
        image_path = os.path.join(image_folder, f"image-{unique_name}.png")
        output_path = os.path.join(output_dir, f"image-{unique_name}.png")
        
        with image.open() as image_bytes:
            with open(output_path, 'wb') as f:
                f.write(image_bytes.read())
        
        return {
            "src": image_path
        }



def main():
    ...



if __name__ == "__main__":
    main()