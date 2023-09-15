# Showcase Data Tool
This is a tool I created to help automate the process of turning registrations into a Word document for our in-house designer.
## Packages Used
* csv
* json
* python-docx
* re
## Project Goals
1. Store directory and file name information in a JSON file (to allow directories to be quickly updated for the next event).
2. Pull in a list of files to exclude (such as due to incomplete information from participant).
3. Write a Word document:
    * Document should exclude fields not relevent for entry type (e.g. a development should not have a spot for bedrooms/bathrooms).
    * For bulleted list of features, document should only include as many bullet items as there are features provided (number of features can vary between entries).
4. Save document with correct naming convention.