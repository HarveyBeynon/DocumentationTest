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

import os.path
import win32com.client
import subprocess
import shutil
import os

# TODO Add script that uses confluence REST API to export the page's to word
# Currently this has to be done manually and the file needs to be placed in the exportedDocs dir

# Convert files in exportedDocs folder from .doc to .docx
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

# TODO find a way to automate the filenames from the exported filename
# TODO Create for loop to allow for multiple files
# Call pandoc to convert the .docx to .rst
subprocess.run(["pandoc", "exportedDocs/Test+Documentation.docx", "-o", "exportedDocs/testDocument.rst"])

# Move the new page to the correct location in the "docs" folder
shutil.move('exportedDocs/testDocument.rst', 'docs/pages/testDocument.rst')

# Delete All docx file
dir_name = "exportedDocs\\"
test = os.listdir(dir_name)
for item in test:
    if item.endswith(".docx"):
        os.remove(os.path.join(dir_name, item))

# TODO create a for loop to append the path name to the index.rst for multiple files
# Add the path to the file to the index.rst toc so that the new file will appear in the contents
with open('docs/index.rst', 'r', encoding='utf-8') as file:
    data = file.readlines() 
data[10] = "pages/testDocument\n"
with open('docs/index.rst', 'w', encoding='utf-8') as file:
    file.writelines(data)

# Commit and push to GitHub
subprocess.run(["git", "add", "."])
subprocess.run(["git", "commit", "-m", "'commit from python script'"])
subprocess.run(["git", "push"])
print("Convesion finished - new docs should now be viewable on Read the Docs")