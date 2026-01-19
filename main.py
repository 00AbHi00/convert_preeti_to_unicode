from docx import Document
import npttf2utf
from tkinter import *

LEGACY_FONTS = {
    "preeti",
    "himalayan tt",
    "kantipur",
    "pcs nepali",
}

def is_legacy_font(run):
    # Direct font
    if run.font.name and run.font.name.lower() in LEGACY_FONTS:
        return True

    # Inherited from style
    if run.style and run.style.font.name:
        return run.style.font.name.lower() in LEGACY_FONTS

    return False


def convert_runs(runs, mapper):
    for run in runs:
        if is_legacy_font(run) and run.text.strip():
            run.text = mapper.map_to_unicode(run.text, from_font="Preeti")
            run.font.name = "Mangal"  # Unicode Nepali font


def convert_docx_preserve_everything(input_docx, output_docx, map_file="map.json"):
    doc = Document(input_docx)
    mapper = npttf2utf.FontMapper(map_file)

    # Convert paragraphs
    for para in doc.paragraphs:
        convert_runs(para.runs, mapper)

    # Convert tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    convert_runs(para.runs, mapper)

    # Headers & footers (often missed!)
    for section in doc.sections:
        for para in section.header.paragraphs:
            convert_runs(para.runs, mapper)
        for para in section.footer.paragraphs:
            convert_runs(para.runs, mapper)

    doc.save(output_docx)

def openFile():
    return 'hi'

if __name__ == "__main__":
    window= Tk()
    button = Button(text='Open', command=openFile)
    button.pack()
    try:
        convert_docx_preserve_everything(
            input_docx="dummy.docx",
            output_docx="output_unicode.docx"
        )
        print("Successfully saved to output_unicode.docx")
    except:
       print("Something went wrong")