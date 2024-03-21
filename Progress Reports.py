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
from docxtpl import InlineImage, RichText
import shutil

def clear_directory(dir_path):
    if os.path.exists(dir_path):
        for filename in os.listdir(dir_path):
            file_path = os.path.join(dir_path, filename)
            if os.path.isfile(file_path) or os.path.islink(file_path):
                os.unlink(file_path)
            elif os.path.isdir(file_path):
                shutil.rmtree(file_path)
    else:
        print(f"The directory {dir_path} does not exist.")

csv_file_path = sys.argv[1]
date = datetime.strptime(sys.argv[2], "%Y-%m-%d")

template_path = './Source Documents/ProgressReports.Template.docx'


# Load the configuration settings from the JSON file
with open('config.json') as f:
    config = json.load(f)

output_path = config['progress_reports_path']['output_path']
temp_images_path = './Temporary Images'

os.makedirs(temp_images_path, exist_ok=True)

clear_directory(temp_images_path)

def print_context_around_error(context, char_position, window=50):
    context_string = str(context)  # Convert context dictionary to string
    start = max(char_position - window, 0)
    end = min(char_position + window, len(context_string))
    print(f"Context around character {char_position}:")
    print(context_string[start:end])

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
    
# Prepare the context for the template with correct information
    context = {
        'first_name': row['first_name'],
        'last_name': row['last_name'],
        'report_date': datetime.strptime(row['report_date'], '%Y-%m-%d').strftime('%m/%d/%Y'),
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
        'image_url' : '', # Initialize with an empty string, will be updated later
        'c1_header': row['c1_header'],
        'c2_header': row['c2_header'],
        'c3_header': row['c3_header'],
        'c4_header': row['c4_header'],
        'c1_11': '' if pd.isna(row['c1_11']) else row['c1_11'],
        'c1_12': '' if pd.isna(row['c1_12']) else row['c1_12'],
        'c1_13': '' if pd.isna(row['c1_13']) else row['c1_13'],
        'c1_14': '' if pd.isna(row['c1_14']) else row['c1_14'],
        'c1_15': '' if pd.isna(row['c1_15']) else row['c1_15'],
        'c1_16': '' if pd.isna(row['c1_16']) else row['c1_16'],
        'c1_17': '' if pd.isna(row['c1_17']) else row['c1_17'],
        'c1_21': '' if pd.isna(row['c1_21']) else row['c1_21'],
        'c1_22': '' if pd.isna(row['c1_22']) else row['c1_22'],
        'c1_23': '' if pd.isna(row['c1_23']) else row['c1_23'],
        'c1_24': '' if pd.isna(row['c1_24']) else row['c1_24'],
        'c1_25': '' if pd.isna(row['c1_25']) else row['c1_25'],
        'c1_26': '' if pd.isna(row['c1_26']) else row['c1_26'],
        'c1_27': '' if pd.isna(row['c1_27']) else row['c1_27'],
        'c1_31': '' if pd.isna(row['c1_31']) else row['c1_31'],
        'c1_32': '' if pd.isna(row['c1_32']) else row['c1_32'],
        'c1_33': '' if pd.isna(row['c1_33']) else row['c1_33'],
        'c1_34': '' if pd.isna(row['c1_34']) else row['c1_34'],
        'c1_35': '' if pd.isna(row['c1_35']) else row['c1_35'],
        'c1_36': '' if pd.isna(row['c1_36']) else row['c1_36'],
        'c1_37': '' if pd.isna(row['c1_37']) else row['c1_37'],
        'c1_41': '' if pd.isna(row['c1_41']) else row['c1_41'],
        'c1_42': '' if pd.isna(row['c1_42']) else row['c1_42'],
        'c1_43': '' if pd.isna(row['c1_43']) else row['c1_43'],
        'c1_44': '' if pd.isna(row['c1_44']) else row['c1_44'],
        'c1_45': '' if pd.isna(row['c1_45']) else row['c1_45'],
        'c1_46': '' if pd.isna(row['c1_46']) else row['c1_46'],
        'c1_47': '' if pd.isna(row['c1_47']) else row['c1_47'],
        'c2_11': '' if pd.isna(row['c2_11']) else row['c2_11'],
        'c2_12': '' if pd.isna(row['c2_12']) else row['c2_12'],
        'c2_13': '' if pd.isna(row['c2_13']) else row['c2_13'],
        'c2_14': '' if pd.isna(row['c2_14']) else row['c2_14'],
        'c2_15': '' if pd.isna(row['c2_15']) else row['c2_15'],
        'c2_16': '' if pd.isna(row['c2_16']) else row['c2_16'],
        'c2_17': '' if pd.isna(row['c2_17']) else row['c2_17'],
        'c2_21': '' if pd.isna(row['c2_21']) else row['c2_21'],
        'c2_22': '' if pd.isna(row['c2_22']) else row['c2_22'],
        'c2_23': '' if pd.isna(row['c2_23']) else row['c2_23'],
        'c2_24': '' if pd.isna(row['c2_24']) else row['c2_24'],
        'c2_25': '' if pd.isna(row['c2_25']) else row['c2_25'],
        'c2_26': '' if pd.isna(row['c2_26']) else row['c2_26'],
        'c2_27': '' if pd.isna(row['c2_27']) else row['c2_27'],
        'c2_31': '' if pd.isna(row['c2_31']) else row['c2_31'],
        'c2_32': '' if pd.isna(row['c2_32']) else row['c2_32'],
        'c2_33': '' if pd.isna(row['c2_33']) else row['c2_33'],
        'c2_34': '' if pd.isna(row['c2_34']) else row['c2_34'],
        'c2_35': '' if pd.isna(row['c2_35']) else row['c2_35'],
        'c2_36': '' if pd.isna(row['c2_36']) else row['c2_36'],
        'c2_37': '' if pd.isna(row['c2_37']) else row['c2_37'],
        'c2_41': '' if pd.isna(row['c2_41']) else row['c2_41'],
        'c2_42': '' if pd.isna(row['c2_42']) else row['c2_42'],
        'c2_43': '' if pd.isna(row['c2_43']) else row['c2_43'],
        'c2_44': '' if pd.isna(row['c2_44']) else row['c2_44'],
        'c2_45': '' if pd.isna(row['c2_45']) else row['c2_45'],
        'c2_46': '' if pd.isna(row['c2_46']) else row['c2_46'],
        'c2_47': '' if pd.isna(row['c2_47']) else row['c2_47'],
        'c3_11': '' if pd.isna(row['c3_11']) else row['c3_11'],
        'c3_12': '' if pd.isna(row['c3_12']) else row['c3_12'],
        'c3_13': '' if pd.isna(row['c3_13']) else row['c3_13'],
        'c3_14': '' if pd.isna(row['c3_14']) else row['c3_14'],
        'c3_15': '' if pd.isna(row['c3_15']) else row['c3_15'],
        'c3_16': '' if pd.isna(row['c3_16']) else row['c3_16'],
        'c3_17': '' if pd.isna(row['c3_17']) else row['c3_17'],
        'c3_21': '' if pd.isna(row['c3_21']) else row['c3_21'],
        'c3_22': '' if pd.isna(row['c3_22']) else row['c3_22'],
        'c3_23': '' if pd.isna(row['c3_23']) else row['c3_23'],
        'c3_24': '' if pd.isna(row['c3_24']) else row['c3_24'],
        'c3_25': '' if pd.isna(row['c3_25']) else row['c3_25'],
        'c3_26': '' if pd.isna(row['c3_26']) else row['c3_26'],
        'c3_27': '' if pd.isna(row['c3_27']) else row['c3_27'],
        'c3_31': '' if pd.isna(row['c3_31']) else row['c3_31'],
        'c3_32': '' if pd.isna(row['c3_32']) else row['c3_32'],
        'c3_33': '' if pd.isna(row['c3_33']) else row['c3_33'],
        'c3_34': '' if pd.isna(row['c3_34']) else row['c3_34'],
        'c3_35': '' if pd.isna(row['c3_35']) else row['c3_35'],
        'c3_36': '' if pd.isna(row['c3_36']) else row['c3_36'],
        'c3_37': '' if pd.isna(row['c3_37']) else row['c3_37'],
        'c3_41': '' if pd.isna(row['c3_41']) else row['c3_41'],
        'c3_42': '' if pd.isna(row['c3_42']) else row['c3_42'],
        'c3_43': '' if pd.isna(row['c3_43']) else row['c3_43'],
        'c3_44': '' if pd.isna(row['c3_44']) else row['c3_44'],
        'c3_45': '' if pd.isna(row['c3_45']) else row['c3_45'],
        'c3_46': '' if pd.isna(row['c3_46']) else row['c3_46'],
        'c3_47': '' if pd.isna(row['c3_47']) else row['c3_47'],
        'c4_11': '' if pd.isna(row['c4_11']) else row['c4_11'],
        'c4_12': '' if pd.isna(row['c4_12']) else row['c4_12'],
        'c4_13': '' if pd.isna(row['c4_13']) else row['c4_13'],
        'c4_14': '' if pd.isna(row['c4_14']) else row['c4_14'],
        'c4_15': '' if pd.isna(row['c4_15']) else row['c4_15'],
        'c4_16': '' if pd.isna(row['c4_16']) else row['c4_16'],
        'c4_17': '' if pd.isna(row['c4_17']) else row['c4_17'],
        'c4_21': '' if pd.isna(row['c4_21']) else row['c4_21'],
        'c4_22': '' if pd.isna(row['c4_22']) else row['c4_22'],
        'c4_23': '' if pd.isna(row['c4_23']) else row['c4_23'],
        'c4_24': '' if pd.isna(row['c4_24']) else row['c4_24'],
        'c4_25': '' if pd.isna(row['c4_25']) else row['c4_25'],
        'c4_26': '' if pd.isna(row['c4_26']) else row['c4_26'],
        'c4_27': '' if pd.isna(row['c4_27']) else row['c4_27'],
        'c4_31': '' if pd.isna(row['c4_31']) else row['c4_31'],
        'c4_32': '' if pd.isna(row['c4_32']) else row['c4_32'],
        'c4_33': '' if pd.isna(row['c4_33']) else row['c4_33'],
        'c4_34': '' if pd.isna(row['c4_34']) else row['c4_34'],
        'c4_35': '' if pd.isna(row['c4_35']) else row['c4_35'],
        'c4_36': '' if pd.isna(row['c4_36']) else row['c4_36'],
        'c4_37': '' if pd.isna(row['c4_37']) else row['c4_37'],
        'c4_41': '' if pd.isna(row['c4_41']) else row['c4_41'],
        'c4_42': '' if pd.isna(row['c4_42']) else row['c4_42'],
        'c4_43': '' if pd.isna(row['c4_43']) else row['c4_43'],
        'c4_44': '' if pd.isna(row['c4_44']) else row['c4_44'],
        'c4_45': '' if pd.isna(row['c4_45']) else row['c4_45'],
        'c4_46': '' if pd.isna(row['c4_46']) else row['c4_46'],
        'c4_47': '' if pd.isna(row['c4_47']) else row['c4_47'],
    }

    # Create a session and set headers
    headers = {
        'Accept': '*/*',
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3'
    }

    # Attempt to download and process the image
    response = requests.get(row['image_url'], headers=headers)
    if response.status_code == 200:
        try:
            img = Image.open(BytesIO(response.content))
            image_filename = f"{context['first_name']}_{context['last_name']}_temp.jpg"
            image_path = os.path.join(temp_images_path, image_filename)
            img.save(image_path)

            # Prepare the InlineImage object for DocxTemplate
            doc = DocxTemplate(template_path)
            inline_image = InlineImage(doc, image_path, width=Inches(2))
            context['image_url'] = inline_image  # Update context with InlineImage
        except Exception as e:
            print(f"Error processing image for {row['first_name']} {row['last_name']}: {e}")
            continue  # Skip to next iteration on error
    else:
        print(f"Failed to fetch image for {row['first_name']} {row['last_name']}. Status: {response.status_code}")
        continue  # Skip to next iteration on error

    # Render and save the document with the updated context
    try:
        doc.render(context)
        doc_filename = os.path.join(output_path, f"Progress_Report_{row['first_name']}_{row['last_name']}.docx")
        doc.save(doc_filename)
    except TemplateSyntaxError as e:
        print(f"Error rendering document for {row['first_name']} {row['last_name']}: {e}")
    except Exception as e:
        print(f"Unexpected error: {e}")

    # Convert the document to PDF
    pdf_filename = doc_filename.replace('.docx', '.pdf')
    convert_to_pdf(doc_filename, pdf_filename)

    os.remove(doc_filename)

    # Determine the parole officer's folder path
    parole_officer_folder = os.path.join(output_path, row['parole_officer'])
    os.makedirs(parole_officer_folder, exist_ok=True)

    # Move the PDF to the parole officer's folder
    final_pdf_path = os.path.join(parole_officer_folder, os.path.basename(pdf_filename))
    shutil.move(pdf_filename, final_pdf_path)

    print(f"Progress Report for {row['first_name']} {row['last_name']} saved to: {final_pdf_path}")

# Open the Reports directory in the file explorer at the end (Windows-specific)
os.startfile(output_path)