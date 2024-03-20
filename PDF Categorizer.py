import os
import re
import PyPDF2
import shutil
import sys

# The directory containing your PDFs
# Get it from command-line arguments
directory_path = 'C:\\Users\\vaner\\OneDrive\\Desktop\\3.19.24'

# Dictionary to hold the mapping of officer names to their respective PDFs
officer_to_pdfs = {}

# Function to extract text from a single PDF
def extract_text_from_pdf(pdf_path):
    with open(pdf_path, 'rb') as pdf_file_obj:
        pdf_reader = PyPDF2.PdfReader(pdf_file_obj)
        page_obj = pdf_reader.pages[0]
        return page_obj.extract_text()
    
# Function to extract officer's name using regular expressions
def extract_officer_name(text):
    patterns = [
        r"Officer: ([^\n\r,]*)",  # Adjusted pattern to stop at a comma or newline
        r"To: Officer ([^\n\r,]*)",
        r"To whom it may concern: Officer ([^\n\r,]*)",
        r"Dear\s+Officer\s+([^\n\r,]*)",  # \s+ matches one or more spaces
        r"To Officer ([^\n\r,]*)"
    ]
    combined_pattern = '|'.join(patterns)
    
    match = re.search(combined_pattern, text, re.IGNORECASE)
    if match:
        officer_name = next(group for group in match.groups() if group is not None)
        return officer_name.strip()

    # If no match was found with the regular expressions, check for "Case Manager"
    lines = text.split('\n')
    for i in range(1, len(lines)):
        if 'Case Manager' in lines[i]:
            return lines[i-1].strip()

    return None

# Iterate over each PDF in the directory
for filename in os.listdir(directory_path):
    if filename.endswith('.pdf'):
        pdf_path = os.path.join(directory_path, filename)
        text = extract_text_from_pdf(pdf_path)
        
        officer_name = extract_officer_name(text)
        
        if officer_name:
            if officer_name not in officer_to_pdfs:
                officer_to_pdfs[officer_name] = []
            officer_to_pdfs[officer_name].append(pdf_path)

# Create folders and move the PDFs to the corresponding officer's folder
for officer, pdfs in officer_to_pdfs.items():
    officer_folder = os.path.join(directory_path, officer)
    if not os.path.exists(officer_folder):
        os.makedirs(officer_folder)
    
    for pdf_path in pdfs:
        shutil.move(pdf_path, os.path.join(officer_folder, os.path.basename(pdf_path)))

# Function to merge folders with the same officer name
def merge_folders(directory_path):
    officer_folders = [folder for folder in os.listdir(directory_path) if os.path.isdir(os.path.join(directory_path, folder))]
    officer_names = set()
    for folder in officer_folders:
        officer_name = re.sub(r'\d+$', '', folder).strip()  # Remove any numbers at the end of the folder name
        if officer_name not in officer_names:
            officer_names.add(officer_name)
        else:
            main_folder = os.path.join(directory_path, officer_name)
            duplicate_folder = os.path.join(directory_path, folder)
            for filename in os.listdir(duplicate_folder):
                shutil.move(os.path.join(duplicate_folder, filename), main_folder)
            os.rmdir(duplicate_folder)

# Call the merge_folders function at the end of the script
merge_folders(directory_path)

print("PDFs have been categorized and moved to respective officer folders. Duplicate folders have been merged.")
