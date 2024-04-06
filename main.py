from docx import Document
from pathlib import Path

# Prompt for the input path
md_file_path = input("Enter the path of the input MD file: ")

# Convert backslashes to forward slashes and remove double quotes
md_file = Path(md_file_path.strip('"')).resolve().as_posix()

# Create the output file path by replacing the extension of the input file with ".docx"
docx_file = Path(md_file).with_suffix('.docx')

# Create a new Word document
doc = Document()

# Read the content from the MD file and add paragraphs to the document
with open(md_file, 'r', encoding='utf-8') as f:
    for line in f:
        doc.add_paragraph(line.strip())

# Save the Word document
doc.save(str(docx_file))

print("Conversion completed successfully!")