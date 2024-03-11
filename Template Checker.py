from docx import Document
from docxtpl import DocxTemplate
from jinja2 import TemplateSyntaxError

# Path to the .docx file
doc_path = "C:\\Users\\vaner\\OneDrive\\Free for Life\\Project\\Projects\\Source Documents\\ProgressReports.Template.docx"

# Open the .docx file
doc = DocxTemplate(doc_path)

# Create a dummy context
context = {}

# Try to render the template with the dummy context
try:
    doc.render(context)
except TemplateSyntaxError as e:
    print(f"Caught a Jinja2 syntax error: {e}")