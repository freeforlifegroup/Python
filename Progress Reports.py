import json
import os
import glob
from datetime import datetime
from dateutil.relativedelta import relativedelta
import pandas as pd
from docxtpl import DocxTemplate
import sys
import calendar
import comtypes.client
import subprocess
import time
import tempfile
import re
from openpyxl.utils import column_index_from_string
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string
import psutil
from jinja2.exceptions import TemplateSyntaxError
import requests
from PIL import Image
from io import BytesIO
from docx import Document
from docx.shared import Inches

def print_context_around_error(context, char_position, window=50):
    context_string = str(context)  # Convert context dictionary to string
    start = max(char_position - window, 0)
    end = min(char_position + window, len(context_string))
    print(f"Context around character {char_position}:")
    print(context_string[start:end])

csv_file_path = sys.argv[1]
date = datetime.strptime(sys.argv[2], "%Y-%m-%d")

# Define the path to your .docx template file
template_path = './Source Documents/ProgressReports.Template.docx'

# Load the configuration settings from the JSON file
with open('config.json') as f:
    config = json.load(f)

output_path = config['progress_reports_path']['output_path']

# Check command-line arguments for date and CSV file path
if len(sys.argv) != 3:
    print("Usage: python pr.py <csv_file_path> <date> with date in YYYY-MM-DD format.")
    sys.exit(1)    

# Read the CSV file into a DataFrame and parse 'report_date' as a datetime object
print("Reading CSV file into DataFrame...")
df = pd.read_csv(csv_file_path, parse_dates=['report_date'])

# Convert 'report_date' to string in 'YYYY-MM-DD' format
df['report_date'] = df['report_date'].dt.strftime('%Y-%m-%d')

# Print a subset of 'report_date' and the date variable
print("Subset of 'report_date':", df['report_date'].head())
print("Date variable:", date)

# Print DataFrame's summary information
print("DataFrame info:")
print(df.info())

# Convert 'date' to string in 'YYYY-MM-DD' format
date = date.strftime('%Y-%m-%d')

# Combine the 'case_manager_first_name' and 'case_manager_last_name' columns to create a 'parole_officer' column
print("Combining 'case_manager_first_name' and 'case_manager_last_name' columns...")
df['parole_officer'] = df['case_manager_first_name'] + ' ' + df['case_manager_last_name']

# Remove leading and trailing spaces from 'report_date'
df['report_date'] = df['report_date'].str.strip()

# Now print the unique dates in 'report_date' before filtering
print("Unique dates in 'report_date' before filtering:", df['report_date'].unique())

# Filter the DataFrame based on the date
print("Filtering DataFrame based on the date...")
df = df[df['report_date'] == date]

# Print unique dates in the 'report_date' before converting
print("Unique dates in 'report_date' before filtering:", df['report_date'].unique())

# Print unique dates in the 'report_date' column
print("Unique dates in 'report_date':", df['report_date'].unique())

# Print unique dates in 'report_date' after converting to string
print("Unique dates in 'report_date' after converting to string:", df['report_date'].unique())

print("DataFrame after filtering:")
print(df)

# Function to convert Word to PDF
def convert_to_pdf(input_file_path, output_file_path):
    # Create a Word application object
    word = comtypes.client.CreateObject('Word.Application')

    # Set the application to be invisible
    word.Visible = False

    # Open the Word document
    doc = word.Documents.Open(input_file_path)

    # Save the document as a PDF
    doc.SaveAs(output_file_path, FileFormat=17)  # 17 represents the PDF format in Word

    # Close the document and quit Word
    doc.Close()
    word.Quit()

# Iterate over each row in the DataFrame
for index, row in df.iterrows():
    
    # Create a session and set headers
    headers = {
        'Accept': '*/*',
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3'
    }

    # Use the session to make the request
    response = requests.get(row['image_url'], headers=headers)

    if response.status_code == 200:
        try:
            # If the response is successful, proceed
            img = Image.open (BytesIO(response.content))
            img.save('temp.jpg')
        except Exception as e:
            print(f"Error opening image for {row['first_name']} {row['last_name']}: {e}")
            continue
    else:
        print(f"Failed to fetch the image for {row['first_name']} {row['last_name']}. Status code: {response.status_code}")
        continue    

    # Prepare the context for the template with correct information
    context = {
        'first_name': row['first_name'],
        'last_name': row['last_name'],
        'report_date': row['report_date'],
        'dob': row['dob'],
        'age': row['age'],
        'case_manager_first_name': row['case_manager_first_name'],
        'case_manager_last_name': row['case_manager_last_name'],
        'orientation_date': row['orientation_date'],
        'required_sessions': row['required_sessions'],
        'attended': row['attended'],
        'absence_unexcused': row['absence_unexcused'],
        'client_note': row['client_note'],
        'AA': row['speaks_significantly_in_group'],
        'AB': row['respectful_to_group'],
        'AC': row['takes_responsibility_for_past'],
        'AD': row['disruptive_argumentitive'],
        'AE': row['humor_inappropriate'],
        'AF': row['blames_victim'],
        'AG': row['appears_drug_alcohol'],
        'AH': row['inappropriate_to_staff'],
        'image_path' : 'temp.jpg'
    }

    for key, value in context.items():
        if isinstance(value, str):
            context[key] = value.replace("\\", "\\\\")

        # Load the template document correctly
        print_context_around_error(context, 17499)

        # Before your try-except block, remove or comment out the incorrect initialization
        # print_context_around_error(context, 17499) -- This line seems correctly placed for debugging purposes

        doc = DocxTemplate(template_path)  # Correct initialization of DocxTemplate

        try:
            doc.render(context)
        except TemplateSyntaxError as e:
            print(f"Error rendering document for {row['first_name']} {row['last_name']}: {e}")
            continue  # Correctly skip to the next iteration upon encountering a TemplateSyntaxError
        except Exception as e:  # Catching other exceptions
            print(f"Error rendering document for {row['first_name']} {row['last_name']}: {e}")
            continue

        # Construct the filename for saving
        doc_filename = os.path.join(output_path, f"Progress_Report_{row['first_name']}_{row['last_name']}.docx")

        # Save the rendered document
        doc.save(doc_filename)

        # Now you can open it with Document
        doc = Document(doc_filename)

        for paragraph in doc.paragraphs:
            if 'IMAGE_PLACEHOLDER' in paragraph.text:
                paragraph.clear()
                run = paragraph.add_run()
                run.add_picture(context['image_path'], width=Inches(1.0))            
    
    # Save the document
    doc.save(doc_filename)

    # Convert the document to PDF
    pdf_filename = doc_filename.replace('.docx', '.pdf')
    convert_to_pdf(doc_filename, pdf_filename)

    # Optionally, delete the Word document if only the PDF is needed
    # os.remove(doc_filename)

    # Determine the parole officer's folder path
    parole_officer_folder = os.path.join(output_path, "Reports", row['parole_officer'])
    os.makedirs(parole_officer_folder, exist_ok=True)

    # Move the PDF to the parole officer's folder
    final_pdf_path = os.path.join(parole_officer_folder, os.path.basename(pdf_filename))
    os.rename(pdf_filename, final_pdf_path)

    print(f"Progress Report for {row['first_name']} {row['last_name']} saved to: {final_pdf_path}")

# Optional: Open the Reports directory in the file explorer at the end
# os.startfile(output_path) # This line is platform-dependent and works on Windows.
