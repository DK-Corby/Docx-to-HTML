import sys
from docx import Document

def convert_styles_to_html(docx_path, output_path):
    # Load the document
    doc = Document(docx_path)

    # Iterate through all paragraphs and runs
    for paragraph in doc.paragraphs:
        for run in paragraph.runs:
            if run.bold and run.italic:
                run.text = f"<b><i>{run.text}</i></b>"
                run.bold = False
                run.italic = False
            elif run.bold:
                run.text = f"<b>{run.text}</b>"
                run.bold = False
            elif run.italic:
                run.text = f"<i>{run.text}</i>"
                run.italic = False

    # Save the modified document
    doc.save(output_path)

if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage: python convert_to_html.py <input_docx_path> <output_docx_path>")
    else:
        input_docx_path = sys.argv[1]
        output_docx_path = sys.argv[2]
        convert_styles_to_html(input_docx_path, output_docx_path)

