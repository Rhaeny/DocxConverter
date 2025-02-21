import os
import mammoth
import uuid
from markdownify import markdownify
from pathlib import Path



class Docx2MdConverter:
    def __init__(self, filepath_abs: str) -> None:
        self.filepath_abs = filepath_abs
        self.filename = Path(filepath_abs).stem
        with open(filepath_abs, "rb") as docx_file:
            result = mammoth.convert_to_html(docx_file, convert_image=mammoth.images.img_element(convert_image))
            html = result.value
            self.markdown = markdownify(html)

    def save(self, output_directory: str) -> None:
        if not os.path.exists(output_directory):
            os.makedirs(output_directory)
        filename = f"{self.filename}.md"
        with open(output_directory + "/" + filename, 'w', encoding="utf-8") as file:
            file.write(self.markdown)

    def move_to_done(self, done_directory: str) -> None:
        done_file_abs = f"{done_directory}\\{self.filename}.docx"
        os.rename(self.filepath_abs, done_file_abs)



def convert_image(image):
    image_folder = ".attachments" # Needs to be set to root, once deployed to DevOps - Absolute path is needed
    output_dir = f"output\\{image_folder}"

    os.makedirs(output_dir, exist_ok=True)
    
    unique_name = uuid.uuid4()
    image_path = os.path.join(image_folder, f"image_{unique_name}.png")
    output_path = os.path.join(output_dir, f"image_{unique_name}.png")
    
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