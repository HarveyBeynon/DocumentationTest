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

# TODO Cycle through each confluence page that is to be published to read the docs and get their ID nums.


# pageIds = ["229377", "851969", "491526"] # Hardcoding page Ids into the array
# for page in pageIds:
#     print(page)
#     subprocess.run(["curl", "-D-", "-X", "GET", "-H", f"Authorization: Basic {API_KEY}", "-H", "Content-Type: application/json", f"https://harveybeynon.atlassian.net/wiki/exportword?pageId={page}", "--output", f"exportedDocs/{page}.doc"]) # Need to find file names

pageIds = {"Lorem+Ipsum":"229377", "Lorem+Ipsum+Example":"851969", "API+Page":"491526"} # Hardcoding pageIds and Filenames
for key, value in pageIds.items():
    print(key, value)
    subprocess.run(["curl", "-D-", "-X", "GET", "-H", f"Authorization: Basic {API_KEY}", "-H", "Content-Type: application/json", f"https://harveybeynon.atlassian.net/wiki/exportword?pageId={value}", "--output", f"exportedDocs/{key}.doc"])

# Export Confluence file as .doc
# subprocess.run(["curl", "-D-", "-X", "GET", "-H", f"Authorization: Basic {API_KEY}", "-H", "Content-Type: application/json", "https://harveybeynon.atlassian.net/wiki/exportword?pageId=491526", "--output", "exportedDocs/testAPI.doc"])

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

# TODO currently this scripts is called from a git bash shell - May need to get the git shh and secret key
# if this script were to run from a an API call.
# Commit and push to GitHub
# subprocess.run(["git", "add", "."])
# subprocess.run(["git", "commit", "-m", "'commit from python script'"])
# subprocess.run(["git", "push"])
# print("Convesion finished - new docs should now be viewable on Read the Docs")