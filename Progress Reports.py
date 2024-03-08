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

csv_file_path = sys.argv[1]
date = datetime.strptime(sys.argv[2], "%Y-%m-%d")

# Load the configuration settings from the JSON file
with open('config.json') as f:
    config = json.load(f)

# Check command-line arguments for date and CSV file path
if len(sys.argv) != 3:
    print("Usage: python pr.py <csv_file_path> <date> with date in YYYY-MM-DD format.")
    sys.exit(1)

# Get the paths from the config file
template_path = config['progress_reports_path']['template_path']
output_path = config['progress_reports_path']['output_path']

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

# Rest of your code...

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

# Ensure the Word to PDF conversion function is defined before the loop

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
        'speaks_significantly_in_group': row['speaks_significantly_in_group'],
        'respectful_to_group': row['respectful_to_group'],
        'takes_responsibility_for_past': row['takes_responsibility_for_past'],
        'disruptive_argumentitive': row['disruptive_argumentitive'],
        'humor_inappropriate': row['humor_inappropriate'],
        'blames_victim': row['blames_victim'],
        'appears_drug_alcohol': row['appears_drug_alcohol'],
        'inappropriate_to_staff': row['inappropriate_to_staff'],
    }

    # Correctly handle backslashes if needed (typically when dealing with file paths or specific string content)
    # for key, value in context.items():
    #     if isinstance(value, str):
    #         context[key] = value.replace("\\", "\\\\")

    # Load the template document correctly
    doc = DocxTemplate(template_path)

    # Render the template with the prepared context
    doc.render(context)

    # Construct the filename for saving
    doc_filename = os.path.join(output_path, f"Progress_Report_{row['first_name']}_{row['last_name']}.docx")

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
