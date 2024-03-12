import pandas as pd
from datetime import datetime, timedelta

# Adjust these paths as necessary
excel_file_path = 'C:\\Users\\vaner\\OneDrive\\Desktop\\Unexcused Absences Snapshot\\BIPP.xlsx'
output_file_path = 'C:\\Users\\vaner\\OneDrive\\Desktop\\Unexcused Absences Snapshot\\BIPP Absences\\BIPP Absences.txt'

# Load the Excel file
df = pd.read_excel(excel_file_path, sheet_name='Sheet1')

# Prepare an output list
output = []

# Iterate over each row
for index, row in df.iterrows():
    name = row['Name']
    attendances = row['column_name_for_attendances'].dropna()
    absences = row['column_name_for_absences'].dropna()
    
    # Convert to string
    attendances = attendances.astype(str)
    absences = absences.astype(str)

    # Pad years to four digits
    attendances = attendances.str.replace(r'(\d{1,2}/\d{1,2}/)(\d{2,3})$', lambda x: x.group(1) + '0'*(4-len(x.group(2))) + x.group(2), regex=True)
    absences = absences.str.replace(r'(\d{1,2}/\d{1,2}/)(\d{2,3})$', lambda x: x.group(1) + '0'*(4-len(x.group(2))) + x.group(2), regex=True)

    # Filter out non-date values
    attendances = attendances[pd.to_datetime(attendances, format='%m/%d/%Y', errors='coerce').notna()]
    absences = absences[pd.to_datetime(absences, format='%m/%d/%Y', errors='coerce').notna()]

    # Convert to datetime
    attendance_dates = pd.to_datetime(attendances, format='%m/%d/%Y').tolist()
    absence_dates = pd.to_datetime(absences, format='%m/%d/%Y').tolist()

    # Your logic to check for gaps and match against absences goes here
    # This will need to be adapted based on your specific logic for checking gaps
    
    # For simplicity, let's say you've identified unaccounted gaps
    unaccounted_gaps = []  # Populate this list based on your gap analysis
    
    if unaccounted_gaps:
        output.append(f"{name}: {', '.join(unaccounted_gaps)}")

# Write the output to a text file
with open(output_file_path, 'w') as f:
    for line in output:
        f.write(f"{line}\n")

print('Done processing.')