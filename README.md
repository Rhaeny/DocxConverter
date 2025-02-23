<h1>Docx Converter</h1>

This project aims to unify several packages to present one tool for conversion of various formats into markdown language with Azure DevOps Wiki flavor. Supported formats are:
- doc
- docx
- pdf
- pptx

Packages used for this project are:
- mammoth - Convert Word documents from docx to simple and clean HTML and Markdown
- markdownify - Convert HTML to markdown
- pdf2docx - Open source Python library converting pdf to docx.
- aspose.slides - PPTX to DOCX conversion

When you first run converter.py using `python converter.py` command, an input folder structure will be created as follows:
- input
  - doc
  - docx
  - pdf
  - pptx

The program expects users to sort their files into these folders. Once sorted, running `python converter.py` command again will do following:
- Files from **input\\doc** folder with .doc extension will be saved as .docx into **input\\docx** folder using opened *Word application*. Every file will be moved to folder **input\\doc\\done** after getting processed.
- Files from **input\\pptx** folder with .pptx extension will be converted to .pdf and saved into **input\\pdf** folder using *aspose.slides* package. Every file will be moved to folder **input\\pptx\\done** after getting processed.
- Files from **input\\pdf** folder with .pdf extension will be converted to .docx and saved into **input\\docx** folder using *pdf2docx* package. Every file will be moved to folder **input\\docx\\done** after getting processed.
- Files from **input\\docx** folder with .docx extension will be converted to markdown syntax (.md) and saved into **output** folder using *mammoth* and *markdownify* packages. Every file will be moved to folder **input\\docx\\done** after getting processed.