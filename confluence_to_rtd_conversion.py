# This is a python script that aims to convert the documentation from 
# confluence to read the docs

# Process of convesion
# 1. Export the confluenece page to word (.doc)
# 2. Convert file from .doc to .docx (needed for step 3)
# 3. Call pandoc to convert the .docx to .rst (pandoc oldfilename.docx -o newfilename.rst)
# 4. Add the new page to the correct location in the "docs" folder
# 5. Add a refrence to thenew file in the index.rst's toc
# 6. Git add > commit > push to allow read the docs to display the new repo content

# Link address to convert page to word https://aurrigo.atlassian.net/wiki/exportword?pageId=10780673

#pip install atlassian-python-api

#pip install pypiwin32

print("Running conversion script")

# Add script that uses confluence REST API to export the page's to word
# Currently this has to be done manually and the file needs to be placed in the exportedDocs directroy

import os.path
import win32com.client

baseDir = 'exportedDocs\\' # Starting directory for directory walk

word = win32com.client.Dispatch("Word.application")

for dir_path, dirs, files in os.walk(baseDir):
    for file_name in files:

        file_path = os.path.join(dir_path, file_name)
        file_name, file_extension = os.path.splitext(file_path)

        if "~$" not in file_name:
            if file_extension.lower() == '.doc': #
                docx_file = '{0}{1}'.format(file_path, 'x')

                if not os.path.isfile(docx_file): # Skip conversion where docx file already exists

                    file_path = os.path.abspath(file_path)
                    docx_file = os.path.abspath(docx_file)
                    try:
                        wordDoc = word.Documents.Open(file_path)
                        wordDoc.SaveAs2(docx_file, FileFormat = 16)
                        wordDoc.Close()
                        print("File has successfully been converted")
                    except Exception as e:
                        print('Failed to Convert: {0}'.format(file_path))
                        print(e)

print("Converting Docx file to reStructuredText")
# Find a way to obtain the name of the file for the pandoc command
# Call the new docx file the same as the confluence filename
# And delete the .doc file.
import subprocess
subprocess.run(["pandoc", "exportedDocs/Test+Documentation.docx", "-o", "exportedDocs/testDocument.rst"])
print("reStructedText file now created")

print("Moving file to Docs directory")
import shutil
shutil.move('exportedDocs/testDocument.rst', 'docs/pages/testDocument.rst')

print("Pushing to GitHub")

subprocess.run(["git", "add", "."])
subprocess.run(["git", "commit", "-m", "'commit from python script'"])
subprocess.run(["git", "push"])