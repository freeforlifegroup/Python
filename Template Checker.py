from docx import Document
from docxtpl import DocxTemplate
from jinja2 import TemplateSyntaxError

# Path to the .docx file
template_path = './Source Documents/ProgressReports.Template.docx'

def print_template_context(char_position, window=50):
    doc = Document(template_path)
    full_text = []
    for para in doc.paragraphs:
        full_text.append(para.text)
    text = "\n".join(full_text)
    
    start = max(char_position - window, 0)
    end = min(char_position + window, len(text))
    context = text[start:end]
    
    print(f"Context around character {char_position}:")
    print(context)

print_template_context(17499)
