import json
import sys
from datetime import datetime
import pandas as pd
from docx import Document
import os
import comtypes.client
import subprocess

# Load the configuration settings from the JSON file
with open(r'C:\Users\vaner\OneDrive\Desktop\Project\config.json') as f:
    config = json.load(f)

# Check command-line arguments for date and CSV file path
if len(sys.argv) != 3:
    print("Usage: python pr.py <csv_file_path> <date> with date in YYYY-MM-DD format.")
    sys.exit(1)

csv_file_path = sys.argv[1]
date = datetime.strptime(sys.argv[2], "%Y-%m-%d")

# Get the paths from the config file
template_path = config['progress_reports_path']['template_path']
output_path = config['progress_reports_path']['output_path']

# Read the CSV file
df = pd.read_csv(csv_file_path, parse_dates=['report_date'])

# Combine the 'case_manager_first_name' and 'case_manager_last_name' columns to create a 'parole_officer' column
df['parole_officer'] = df['case_manager_first_name'] + ' ' + df['case_manager_last_name']

# Filter the DataFrame based on the date
df = df[df['report_date'] == date]

# Iterate over each row in the DataFrame
for index, row in df.iterrows():
    
    # Open the template document
    doc = Document(config['progress_reports_path']['template_path'])

    # Add the progress report data to the document
    doc.add_paragraph(f"Progress Report for {row['first_name']} {row['last_name']}")
    doc.add_paragraph(f"Report Date: {row['report_date'].strftime('%Y-%m-%d')}")
    doc.add_paragraph(f"DOB: {row['dob']}")
    doc.add_paragraph(f"Age: {row['age']}")
    doc.add_paragraph(f"Case Manager: {row['case_manager_first_name']} {row['case_manager_last_name']}")
    doc.add_paragraph(f"Orientation Date: {row['orientation_date']}")
    doc.add_paragraph(f"Required Sessions: {row['required_sessions']}")
    doc.add_paragraph(f"Attended: {row['attended']}")
    doc.add_paragraph(f"Unexcused Absences: {row['absence_unexcused']}")
    doc.add_paragraph(f"Client Note: {row['client_note']}")
    doc.add_paragraph(f"Speaks Significantly in Group: {row['speaks_significantly_in_group']}")
    doc.add_paragraph(f"Respectful to Group: {row['respectful_to_group']}")
    doc.add_paragraph(f"Takes Responsibility for Past: {row['takes_responsibility_for_past']}")
    doc.add_paragraph(f"Disruptive/Argumentative: {row['disruptive_argumentitive']}")
    doc.add_paragraph(f"Inappropriate Humor: {row['humor_inappropriate']}")
    doc.add_paragraph(f"Blames Victim: {row['blames_victim']}")
    doc.add_paragraph(f"Appears Under Influence: {row['appears_drug_alcohol']}")
    doc.add_paragraph(f"Inappropriate to Staff: {row['inappropriate_to_staff']}")

    doc_filename = f"C:\\Users\\vaner\\OneDrive\\Desktop\\Project\\Progress Reports\\Progress_Report_{row['first_name']}_{row['last_name']}.docx"

    # Save the document
    doc.save(doc_filename)

    # Convert the document to a PDF
    pdf_filename = doc_filename.replace('.docx', '.pdf')
    word = comtypes.client.CreateObject('Word.Application')
    doc = word.Documents.Open(doc_filename)
    doc.SaveAs(pdf_filename, FileFormat=17)
    doc.Close()
    word.Quit()

    # Move the PDF to the corresponding parole officer's folder
    parole_officer_folder = f"C:\\Users\\vaner\\OneDrive\\Desktop\\Project\\Progress Reports\\Reports\\{row['parole_officer']}"
    os.makedirs(parole_officer_folder, exist_ok=True)
    os.rename(pdf_filename, os.path.join(parole_officer_folder, os.path.basename(pdf_filename)))

    print(f"Progress Report for {row['first_name']} {row['last_name']} saved to: {os.path.join(parole_officer_folder, os.path.basename(pdf_filename))}")

# Run the PDF Categorizer script
subprocess.run(["python", "PDF Categorizer.py"])

# Open the Reports directory in the file explorer
os.startfile(output_path)
