from docx import Document
from docx.enum.text import WD_COLOR_INDEX
import converter
import re

global replaced_runs
replaced_runs=[]

def contains_unicode(text):
    # Check if the text contains Unicode characters
    return bool(re.search(r'[^\x00-\x7F]', text))

def replace_and_highlight(doc_path, save_path):
    doc = Document(doc_path)
    
    # Create an instance of the Unicode class from converter.py
    unicode_converter = converter.Unicode()
    
    for paragraph in doc.paragraphs:
        for run in paragraph.runs:
            if contains_unicode(run.text):
                if contains_unicode(run.text) and run.font.name == "SolaimanLipi":
            # Convert the text from Unicode to Bijoy
                    converted_text = unicode_converter.convertUnicodeToBijoy(run.text)
                    if run.text != converted_text:
                        run.text = converted_text
                        run.font.highlight_color = WD_COLOR_INDEX.TURQUOISE  # BLUE highlight
                        replaced_runs.append(run)
    
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        if contains_unicode(run.text):
                            if contains_unicode(run.text) and run.font.name == "SolaimanLipi":
                        # Convert the text from Unicode to Bijoy
                                converted_text = unicode_converter.convertUnicodeToBijoy(run.text)
                                if run.text != converted_text:
                                    run.text = converted_text
                                    run.font.highlight_color = WD_COLOR_INDEX.TURQUOISE  # BLUE highlight
                                    replaced_runs.append(run)

    for run in replaced_runs:
        run.font.name = "SutonnyMJ"
    
    doc.save(save_path)