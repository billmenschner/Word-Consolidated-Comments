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
newFolder = file_path + '\\CopiedZip'
if data_folder_count > 0:
    if not os.path.exists(newFolder):
        os.makedirs(newFolder)   
        
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
    shutil.copy(file_path + '\\' +current_name, file_path + '\\CopiedZip\\' + new_file_name)
    
#Extract contents of Zip files. Should I just extract the document and comment
#files? I could extract them all to a new folder, give them the name of the 
#file and whether it's a document or comment file, and then iterate through
#each of the files when adding the comments to the master file. 
    
os.makedirs(file_path + '\\CopiedZip\\Extracted')

for file in os.listdir(file_path +  '\\CopiedZip'):
    if file == 'Testing.zip':
        print('True')
        extracted_file = zipfile.ZipFile(file_path + '/CopiedZip/' + file)
        extracted_file.extract('word/document.xml', file_path + '/CopiedZip/')
    else:
        print('False')
