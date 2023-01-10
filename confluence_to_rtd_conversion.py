import os.path
import win32com.client
import subprocess
import shutil
import json
import os
import sys
import fileinput


from dotenv import load_dotenv
load_dotenv()

API_KEY = os.getenv("API_KEY")

# Delete all files in a given directory
# def delete_files_in_dir(dir_name, extension):
#     current_dir = os.listdir(dir_name)
#     for item in current_dir:
#         if (item.endswith(extension)):
#             os.remove(os.path.join(dir_name, item))

# # Move all file with extention from on dir to another
# def move_files(src_dir, dst_dir, extension):
#     test = os.listdir(src_dir)
#     for item in test:
#         if (item.endswith(extension)):
#             src_path = os.path.join(src_dir, item)
#             dst_path = os.path.join(dst_dir, item)
#             shutil.move(src_path, dst_path)

# delete_files_in_dir("exported_docs/", ".doc")
# delete_files_in_dir("docs/pages/", ".rst")
# delete_files_in_dir("API_response/", "space_content")

# # Access the space's content
# subprocess.run(["curl", "-D-", "-X", "GET", "-H", f"Authorization: Basic {API_KEY}", "-H", "Content-Type: application/json", f"https://harveybeynon.atlassian.net/wiki/rest/api/space/MFS/content?expand=children.page&type=page&limit=9999", "--output", "API_response/space_content.json"])

# # Storing values from the json file
# page_ids = []
# page_title = []
# with open("API_response/space_content.json", "r", encoding="utf-8") as f:
#     data = json.load(f)
    
#     results_data = data["page"]["results"]

#     for results in results_data:
#         for key, value in results.items():
#             if key == 'id':
#                 page_ids.append(value)
#             if key == 'title':
#                 # might need to remove whitespace
#                 page_title.append(value)

# # Creating dictionary from id and title array 
# page_info = {}
# for key in page_title:
#     for value in page_ids:
#         page_info[key] = value
#         page_ids.remove(value)
#         break

# # Sorting dictionary alphabetically
# sorted_page_info = {key : value for key, value in sorted(page_info.items())}
 
# # Printing resultant dictionary
# print("Resultant dictionary is : " + str(sorted_page_info))

# # Exporting confluence page as .doc
# # NOTE - Read the Docs lists the files alphabetically by file name
# # page_ids = {"Lorem+Ipsum":"229377", "Lorem+Ipsum+Example":"851969", "API+Page":"491526"} # Hardcoding pageIds and Filenames - NOTE these file need to be in order.
# page_num = 1
# for key, value in sorted_page_info.items():
#     subprocess.run(["curl", "-D-", "-X", "GET", "-H", f"Authorization: Basic {API_KEY}", "-H", "Content-Type: application/json", f"https://harveybeynon.atlassian.net/wiki/exportword?pageId={value}", "--output", f"exported_docs/{page_num}_{key}.doc"])
#     page_num += 1

# base_dir = 'exported_docs\\' # Starting directory for directory walk

# # Convert exported .doc file to .docx for pandoc
# word = win32com.client.Dispatch("Word.application")

# for dir_path, dirs, files in os.walk(base_dir):
#     for file_name in files:

#         file_path = os.path.join(dir_path, file_name)
#         file_name, file_extension = os.path.splitext(file_path)

#         if "~$" not in file_name:
#             if file_extension.lower() == '.doc': #
#                 docx_file = '{0}{1}'.format(file_path, 'x')

#                 if not os.path.isfile(docx_file): # Skip conversion where docx file already exists

#                     file_path = os.path.abspath(file_path)
#                     docx_file = os.path.abspath(docx_file)
#                     try:
#                         word_doc = word.Documents.Open(file_path)
#                         word_doc.SaveAs2(docx_file, FileFormat = 16)
#                         word_doc.Close()

#                         # Call pandoc to convert the .docx file to .rst
#                         subprocess.run(["pandoc", f"{file_name}.docx", "-o", f"{file_name}.rst"])

#                     except Exception as e:
#                         print('Failed to Convert: {0}'.format(file_path))
#                         print(e)

# #TODO - get the images from every file, save them to a new folder, change the images in a file to point to the correct image path

# delete_files_in_dir("exported_docs/", ".docx")
# move_files("exported_docs/", "docs/pages/", ".rst")

# Replace images in .rst
# with open('docs/pages/4_4 - Hello Craig.rst', 'r') as file:
#     filedata = file.read()
#     filedata = filedata.replace('image1.jpeg', '1.jpg')

# with open('docs/pages/4_4 - Hello Craig.rst', 'w') as file:
#     file.write(filedata)



def replace_image_line(page, image_line):
    text = ".. image::" # if any line contains this text, I want to modify the whole line.
    new_line = "\n"
    for file_name in os.walk('docs/pages/'):
        x = fileinput.input(page, inplace=1)
        for line in x:
            if text in line:
                line = image_line + new_line
            sys.stdout.write(line)

replace_image_line("docs/pages/1_1 - Lorem Ipsum.rst", ".. image:: media/1.jpg")
replace_image_line("docs/pages/2_2 - Lorem Ipsum Example.rst", ".. image:: media/2.gif")
replace_image_line("docs/pages/3_3 - API Page.rst", ".. image:: media/3.jpg")
replace_image_line("docs/pages/4_4 - Hello Craig.rst", ".. image:: media/4.jpg")

# x = fileinput.input(files="docs/pages/4_4 - Hello Craig.rst", inplace=1)
# for line in x:
#     if text in line:
#         line = new_text + newline
#     sys.stdout.write(line)

# TODO currently this scripts is called from a git bash shell - May need to get the git shh and secret key
# if this script were to run from a an API call.
# Commit and push to GitHub
# subprocess.run(["git", "add", "."])
# subprocess.run(["git", "commit", "-m", "'commit from python script'"])
# subprocess.run(["git", "push"])
# print("Convesion finished - new docs should now be viewable on Read the Docs")