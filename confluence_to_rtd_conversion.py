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

import os.path
import win32com.client
import subprocess
import shutil
import os

from dotenv import load_dotenv
load_dotenv()

API_KEY = os.getenv("API_KEY")

# subprocess.run(["curl", "-D-", "-X", "GET", "-H", "Authorization: Basic aGFydmV5aWJleW5vbkBnbWFpbC5jb206a0VuNUFQUEk5ZWhUbDVsZkhSalo1MTM1", "Content-Type: application/json", "-H", "https://harveybeynon.atlassian.net/wiki/exportword?pageId=491526", "--output", "exportedDocs/apitest.doc"])

subprocess.run(["curl", "-D-", "-X", "GET", "-H", f"Authorization: Basic {API_KEY}", "-H", "Content-Type: application/json", "https://harveybeynon.atlassian.net/wiki/exportword?pageId=491526", "--output", "exportedDocs/testAPI.doc"])

# curl -D- \
#    -X GET \
#    -H "Authorization: Basic aGFydmV5aWJleW5vbkBnbWFpbC5jb206a0VuNUFQUEk5ZWhUbDVsZkhSalo1MTM1" \
#    -H "Content-Type: application/json" \
#    "https://harveybeynon.atlassian.net/wiki/exportword?pageId=491526" \
#    --output "exportedDocs/test.doc"

baseDir = 'exportedDocs\\' # Starting directory for directory walk

# Convert exported .doc file to .docx for pandoc
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

                        # Call pandoc to convert the .docx file to .rst
                        subprocess.run(["pandoc", f"{file_name}.docx", "-o", f"{file_name}.rst"])

                    except Exception as e:
                        print('Failed to Convert: {0}'.format(file_path))
                        print(e)

dir_name = "exportedDocs\\"
test = os.listdir(dir_name)
for item in test:

    # Delete the .docx files in the exportedDocs dir
    if (item.endswith(".docx")):
        os.remove(os.path.join(dir_name, item))

    # Move all .rst files to the docs/pages dir
    if (item.endswith(".rst")):
        src_path = os.path.join("exportedDocs/", item)
        dst_path = os.path.join("docs/pages/", item)
        shutil.move(src_path, dst_path)

# # TODO currently this scripts is called from a git bash shell - May need to get the git shh and secret key
# # if this script were to run from a an API call.
# # Commit and push to GitHub
# subprocess.run(["git", "add", "."])
# subprocess.run(["git", "commit", "-m", "'commit from python script'"])
# subprocess.run(["git", "push"])
# print("Convesion finished - new docs should now be viewable on Read the Docs")