#! Python3

# PeerReview.py -- Pulls MS Word comments from docx files and copies them into
# consolidated file

import os, shutil, tkinter as tk, glob, re, zipfile
from tkinter import filedialog
from pathlib import Path
from lxml import etree

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
pulled_text = {}
pulled_comments = {}

for file in os.listdir(copied_folder):
    if '.zip' in file:
        os.makedirs(extracted_folder + '/' + os.path.splitext(file)[0])
        new_folder = extracted_folder + '/' + os.path.splitext(file)[0]
        extracted_file = zipfile.ZipFile(copied_folder + '/' + file)
        document = extracted_file.extract('word/document.xml', new_folder)
        comments = extracted_file.extract('word/comments.xml', new_folder)
        #Parse both the documents and comments documents and create the dictionary
        #that will hold all of the information taken from the documents.
        
        document_root = etree.parse(document)
        comments_root = etree.parse(comments)
        
        #Extract text from document based on Comment ID
        
        text_commentId = ''
        document_text = ''
        
        for element in document_root.iter():
            if element in document_root.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t'):
                document_text += element.text
            elif element in document_root.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}commentRangeStart'):
                text_commentId = element.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}id')
                document_text = ''
            elif element in document_root.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}commentRangeEnd'):
                pulled_text[text_commentId] = document_text
        
        #Extract Comment ID, comment, and author from comments.xml. Start running count
        #that gets increased with each comment tag.
        
        comment_number = 0
        commentId = ''
        comment_text = ''
        author = ''
        
        for element in comments_root.iter():
            if element in comments_root.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}comment'):
                if comment_number == 0:
                    commentId = element.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}id')
                    author = element.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}author')
                    comment_number += 1
                else:
                    pulled_comments[comment_number] = commentId, author, comment_text, pulled_text[commentId]
                    comment_number +=1
                    commentId = element.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}id')
                    author = element.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}author')
                    comment_text = ''
            elif element in comments_root.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t'):
                comment_text += element.text
        

#Testing copying and pasting info from files.
        


