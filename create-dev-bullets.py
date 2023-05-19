# Load modules
import csv
from dotenv import load_dotenv
import os

# Reset Error Log
with open("errorlog.txt", "w") as f:
    f.write("")
f.close()

#Load the .env file (must be in same directory as this file)
load_dotenv()

#Read in the data from the .env file
file_dir = os.getenv('dir')
import_addr = file_dir + '\\' + os.getenv('import_file')
skip_addr = file_dir + '\\' + os.getenv('skip_file')

# Creates a dictionary from the CSV data
with open(import_addr, newline='') as csv_file:
    import_data = []
    reader = csv.DictReader(csv_file)
    for row in reader:
        import_data.append(row)
csv_file.close()

# Gets a set of addresses to skip
skip_entries = set()
if os.path.exists(skip_addr):
    with open(skip_addr, newline='') as csv_file:
        reader = csv.reader(csv_file, delimiter=',')
        for row in reader:
            skip_entries.add(row[0])
    csv_file.close()

# Sets up the fields that will turn into a bulleted list
bullet_fields = ['Entry Feature 1', 'Entry Feature 2', 'Entry Feature 3', 'Entry Feature 4', 'Entry Feature 5', 'Entry Feature 6']

# Writes text files
for entry in import_data:

    # Reset text output
    text_output = ''

    # If the Word doc field hasn't been set to 'yes' in MZ
    # and the adddress isn't in the Skip CSV file
    if entry['Entry Street Address'] not in skip_entries:
        
        # For developments and townhome communities
        if entry['Entry Type'] == 'Development' or entry['Entry Type'] == 'Townhome Community':
            # Add bulleted list of features
            for field in bullet_fields:
                if len(entry[field]) > 0:
                    text_output += entry[field] + "\n"
            file_addr = file_dir + '\\' + entry['Entry Street Address'] + ' - Bullets.txt'
            with open(file_addr, 'w') as txt_file:
                txt_file.write(text_output)
            txt_file.close()

        # Save the document
        # document.save(file_dir + '\\' + entry['Entry Street Address'] + ' - Magazine Info.docx')