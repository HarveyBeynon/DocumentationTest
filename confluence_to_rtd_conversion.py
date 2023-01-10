import os.path
import win32com.client
import subprocess
import shutil
import os

from dotenv import load_dotenv
load_dotenv()

API_KEY = os.getenv("API_KEY")

# Delete all files in a given directory
def delete_files_in_dir(dir_name, extension):
    currentDir = os.listdir(dir_name)
    for item in currentDir:
        if (item.endswith(extension)):
            os.remove(os.path.join(dir_name, item))

# Move all file with extention from on dir to another
def move_files(src_dir, dst_dir, extension):
    test = os.listdir(src_dir)
    for item in test:
        if (item.endswith(extension)):
            src_path = os.path.join(src_dir, item)
            dst_path = os.path.join(dst_dir, item)
            shutil.move(src_path, dst_path)

delete_files_in_dir("exportedDocs/", ".doc")
delete_files_in_dir("docs/pages/", ".rst")

# TODO Cycle through each confluence page that is to be published to read the docs and get their ID nums.

# Exporting confluence page as .doc
# NOTE - Read the Docs lists the files alphabetically, hence it will be good paractise to number the files
pageIds = {"Lorem+Ipsum":"229377", "Lorem+Ipsum+Example":"851969", "API+Page":"491526"} # Hardcoding pageIds and Filenames - NOTE these file need to be in order.
pageNum = 1
for key, value in pageIds.items():
    subprocess.run(["curl", "-D-", "-X", "GET", "-H", f"Authorization: Basic {API_KEY}", "-H", "Content-Type: application/json", f"https://harveybeynon.atlassian.net/wiki/exportword?pageId={value}", "--output", f"exportedDocs/{pageNum}+{key}.doc"])
    pageNum += 1

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

delete_files_in_dir("exportedDocs/", ".docx")
move_files("exportedDocs/", "docs/pages/", ".rst")

# TODO currently this scripts is called from a git bash shell - May need to get the git shh and secret key
# if this script were to run from a an API call.
# Commit and push to GitHub
# subprocess.run(["git", "add", "."])
# subprocess.run(["git", "commit", "-m", "'commit from python script'"])
# subprocess.run(["git", "push"])
# print("Convesion finished - new docs should now be viewable on Read the Docs")