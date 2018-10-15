#! Python3

# PeerReview.py -- Pulls MS Word comments from docx files and copies them into
# consolidated file

import os, shutil, tkinter as tk, glob, re, zipfile
from tkinter import filedialog
from pathlib import Path

#Start by having the user select the directory where the files are located
root = tk.Tk()
root.withdraw()

file_path = filedialog.askdirectory()

#Check the folder for docx files
data_folder = Path(file_path).glob('*.docx')
data_folder_count = len(list(data_folder))

if data_folder_count == 0:
    print("This directory doesn't have any docx files")
 
    
#Check for the folder that will contain the zip files created from docx files. 
#If it doesn't exist, create folder.
copied_folder = file_path + '/CopiedZip'
if data_folder_count > 0:
    if not os.path.exists(copied_folder):
        os.makedirs(copied_folder)   
        
#Copy the docx files to the CopiedZip folder and save them as a zip file 
#in that folder
name_pattern = re.compile(r"""^(.*?) #All the text before the extension
    (.docx)    #The extension
    """, re.VERBOSE)

for current_name in os.listdir(file_path):
    working_name = name_pattern.search(current_name)
    if working_name == None:
        continue
    doc_name = working_name.group(1)
    new_file_name = doc_name + '.zip'
    shutil.copy(file_path + '/' +current_name, copied_folder + '/' + new_file_name)
    
#Create folder where files will be extracted.
    
os.makedirs(copied_folder + '/Extracted')
extracted_folder = copied_folder + '/Extracted'

#Create folders based on zip file names and extract the documents and comments
#files from the zip files into the new folders.

for file in os.listdir(copied_folder):
    if '.zip' in file:
        os.makedirs(extracted_folder + '/' + os.path.splitext(file)[0])
        new_folder = extracted_folder + '/' + os.path.splitext(file)[0]
        extracted_file = zipfile.ZipFile(copied_folder + '/' + file)
        extracted_file.extract('word/document.xml', new_folder)
        extracted_file.extract('word/comments.xml', new_folder)

#for file in os.listdir(copied_folder):
#    print(file)
#    if '.zip' in file:
#        extracted_file = zipfile.ZipFile(copied_folder + '/' + file)
#        for name in extracted_file.namelist():
#            member = extracted_file.open(name)
#            print(member)
#            with open(os.path.basename(name), 'wb') as outfile:
##                print(outfile)
#                shutil.copyfileobj(member, outfile)

#for file in os.listdir(copied_folder):
#    if '.zip' in file:
#        extracted_file = zipfile.ZipFile(copied_folder + '/' + file)
#        extracted_file.extract('word/document.xml', extracted_folder)
#        os.rename(extracted_folder + '/word/' + 'document.xml', extracted_folder + '/word/' + 'document1.xml')

#for file in os.listdir(copied_folder):
#    if '.zip' in file:
#        file_path = os.path.join(copied_folder, file)
#        with zipfile.ZipFile(file_path) as zf:
#            for target_file in file_list:
#                if target_file in zf.namelist():
#                    target_name = (os.path.splitext(target_file)[0]+ '1' + ".xml")
#                    target_path = os.path.join(extracted_folder, target_name)
##                    print(target_name)
##                    print(target_path)
#                    with open(target_path, 'w') as f:
#                        f.write(zf.read(target_file))